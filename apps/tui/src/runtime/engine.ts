import { writeFile } from 'node:fs/promises'
import { type CellStyle, parseRange, toA1 } from '@ascend/core'
import type { CellValue, InputValue, Operation, StyleInput } from '@ascend/schema'
import { topLeftScalar } from '@ascend/schema'
import type {
	AscendWorkbook,
	WorkbookInfo,
	WorkbookLoadOptions,
	TraceResult as WorkbookTraceResult,
} from '@ascend/sdk'
import { indexToColumn, inferExportFormat, normalizeExportFormat, parseA1 } from '@ascend/sdk'
import { commandsForGroup, findCommand, listCommands } from '../commands/registry.ts'
import type { DialogId, FindReplaceInput } from '../dialogs/index.ts'
import { buildDialogOperations, findDialog } from '../dialogs/index.ts'
import { ClipboardController } from '../input/clipboard.ts'
import { buildTerminalCalibrationReport } from '../input/terminal-profile.ts'
import { displayCellValue, parseInputValue } from '../model/display-cell.ts'
import {
	backspace,
	createFormulaEditState,
	cycleReferenceMode,
	deleteForward,
	type FormulaEditState,
	insertText,
	moveEditCursor,
} from '../model/formula-edit.ts'
import { buildGridSemanticModel } from '../model/grid-semantics.ts'
import { createSelection, moveSelection, selectCell, selectionRef } from '../model/selection.ts'
import { createViewport, ensureSelectionVisible, resizeViewport } from '../model/viewport.ts'
import { detectTerminalCapabilities } from '../render/terminal-capabilities.ts'
import { commandPaletteResults } from '../ui/command-palette.ts'
import { renderWorkspace } from '../ui/workspace.ts'
import { CalcWorker } from '../workbook/calc-worker.ts'
import { EditJournal } from '../workbook/edit-journal.ts'
import { HydrationService } from '../workbook/hydration-service.ts'
import { protectedReviewReasons } from '../workbook/protected-review.ts'
import { RecentWorkbookStore } from '../workbook/recent-store.ts'
import { WorkbookSessionController } from '../workbook/session-controller.ts'
import { TelemetryBuffer } from './telemetry.ts'
import type {
	CellCoord,
	CommandDescriptor,
	ContextMenuState,
	DialogViewState,
	DispatchResult,
	FocusRegion,
	InputEvent,
	RenderFrame,
	RenderStats,
	SelectionState,
	TerminalSize,
	TraceOptions,
	TraceResult,
	TuiEngine,
	TuiMode,
	TuiStateSnapshot,
	ViewportState,
	WorkbookWorkspace,
} from './types.ts'
import {
	activeDocument,
	createWorkspace,
	setActiveDocument,
	showFileHub,
	updateActiveDocument,
	updateDocument,
} from './workspace.ts'

export interface WorkbookTuiEngineOptions {
	readonly path?: string
	readonly sheet?: string
	readonly size?: TerminalSize
	readonly persistState?: boolean
	readonly recentStorePath?: string
	readonly loadOptions?: WorkbookLoadOptions
}

export class WorkbookTuiEngine implements TuiEngine {
	private readonly session = new WorkbookSessionController()
	private readonly hydration = new HydrationService(this.session)
	private readonly journal = new EditJournal()
	private readonly calcWorker = new CalcWorker()
	private readonly telemetry = new TelemetryBuffer(2000)
	private readonly clipboard = new ClipboardController()
	private recentStore: RecentWorkbookStore | undefined
	private workspace: WorkbookWorkspace = createWorkspace()
	private mode: TuiMode = 'ready'
	private selection: SelectionState = createSelection()
	private viewport: ViewportState
	private editBuffer = ''
	private editCursor = 0
	private editTarget: CellCoord | undefined
	private pointReference: { readonly start: number; readonly end: number } | undefined
	private commandBuffer = ''
	private commandPaletteIndex = 0
	private activeDialog: DialogViewState | undefined
	private contextMenu: ContextMenuState | undefined
	private keyTipActive = false
	private inspectorLines: readonly string[] = []
	private showFormulas = false
	private message = ''
	private activeSheetIndex = 0

	private constructor(size: TerminalSize) {
		this.viewport = createViewport(size)
	}

	static async create(options: WorkbookTuiEngineOptions = {}): Promise<WorkbookTuiEngine> {
		const engine = new WorkbookTuiEngine(options.size ?? defaultTerminalSize())
		engine.recentStore = createRecentWorkbookStore(options)
		engine.refreshRecentWorkbooks()
		if (options.path) {
			const document = await engine.session.open(options.path, options.loadOptions)
			engine.recordRecentWorkbook(options.path)
			engine.workspace = setActiveDocument(engine.workspace, document)
			if (options.sheet) engine.selectSheetByName(options.sheet)
			engine.message = openedWorkbookMessage(document, options.loadOptions)
		}
		return engine
	}

	async dispatch(event: InputEvent): Promise<DispatchResult> {
		const started = performance.now()
		let result: DispatchResult
		if (event.kind === 'resize') {
			this.viewport = resizeViewport(this.viewport, event.size)
			result = { handled: true, shouldRender: true }
		} else if (event.kind === 'command') {
			result = await this.executeCommand(event.command)
		} else if (event.kind === 'mouse') {
			result = this.dispatchMouse(event)
		} else if (this.activeDialog) {
			result = this.dispatchDialogKey(event)
		} else if (this.contextMenu) {
			result = await this.dispatchContextMenuKey(event)
		} else if (this.keyTipActive) {
			result = this.dispatchKeyTipKey(event)
		} else if (this.mode === 'command') {
			result = await this.dispatchCommandKey(event)
		} else if (this.mode === 'editing' || this.mode === 'entering' || this.mode === 'point') {
			result = this.dispatchEditKey(event)
		} else if (this.workspace.fileHub.visible) {
			result = await this.dispatchFileHubKey(event)
		} else {
			result = await this.dispatchReadyKey(event)
		}
		this.telemetry.record({
			inputToFrameMs: performance.now() - started,
			rss: process.memoryUsage().rss,
			heapUsed: process.memoryUsage().heapUsed,
		})
		return result
	}

	render(size: TerminalSize): RenderFrame {
		const started = performance.now()
		this.viewport = resizeViewport(this.viewport, size)
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		const sheetName = this.sheetName()
		const hydrateStart = performance.now()
		const data = this.hasWorkbook()
			? this.hydration.readViewport(sheetName, this.viewport)
			: undefined
		const sheetInfo = this.hasWorkbook() ? this.session.inspectSheet(sheetName) : undefined
		const hydrateMs = performance.now() - hydrateStart
		const frame = renderWorkspace({
			size,
			workspace: this.workspace,
			sheetNames: this.sheetNames(),
			activeSheetIndex: this.activeSheetIndex,
			sheetName,
			mode: this.mode,
			selection: this.selection,
			viewport: this.viewport,
			data,
			editBuffer: this.editBuffer,
			formulaBarContent:
				this.mode === 'editing' || this.mode === 'entering' || this.mode === 'point'
					? this.editBuffer
					: this.hasWorkbook()
						? this.activeCellEditText()
						: '',
			editCursor: this.editCursor,
			commandPalette: { query: this.commandBuffer, selectedIndex: this.commandPaletteIndex },
			activeDialog: this.activeDialog,
			contextMenu: this.contextMenu,
			inspectorLines: this.inspectorLines,
			showFormulas: this.showFormulas,
			gridSemantics: buildGridSemanticModel({ sheet: sheetInfo, viewport: this.viewport }),
			message: this.message,
			dirty: this.session.isDirty(),
			perfSummary: this.perfSummary(),
		})
		this.telemetry.record({
			layoutMs: performance.now() - started,
			hydrateMs,
			bytesWritten: frame.stats.bytes,
			dirtyCells: frame.stats.dirtyCells,
			dirtyRows: frame.stats.dirtyRows,
			tileCacheHitRate: this.hydration.cacheHitRate(),
		})
		return frame
	}

	recordRendererStats(stats: RenderStats): void {
		this.telemetry.record({
			diffMs: stats.frameDiffMs,
			encodeMs: stats.encodeMs,
			ptyWriteMs: stats.writeMs,
			bytesWritten: stats.bytesOut,
			dirtyCells: stats.changedCells,
			fps: stats.fps,
			droppedFrames: stats.droppedFrames,
		})
	}

	async runHeadless(
		trace: readonly InputEvent[],
		options: TraceOptions = { size: defaultTerminalSize() },
	): Promise<TraceResult> {
		const frames: RenderFrame[] = []
		for (const event of trace) {
			await this.dispatch(event)
			const shouldCapture = options.includeFrames ?? true
			if (shouldCapture) frames.push(this.render(options.size))
		}
		return {
			state: this.state(),
			frames,
			telemetry: this.telemetry.all(),
		}
	}

	state(): TuiStateSnapshot {
		return {
			workspace: this.workspace,
			mode: this.mode,
			sheetName: this.sheetName(),
			selection: this.selection,
			viewport: this.viewport,
			editBuffer: this.editBuffer,
			editCursor: this.editCursor,
			commandBuffer: this.commandBuffer,
			commandPalette: { query: this.commandBuffer, selectedIndex: this.commandPaletteIndex },
			...(this.activeDialog ? { activeDialog: this.activeDialog } : {}),
			...(this.contextMenu ? { contextMenu: this.contextMenu } : {}),
			inspectorLines: this.inspectorLines,
			showFormulas: this.showFormulas,
			dirty: this.session.isDirty(),
			message: this.message,
			telemetry: this.telemetry.all(),
		}
	}

	private async dispatchReadyKey(
		event: Extract<InputEvent, { kind: 'key' }>,
	): Promise<DispatchResult> {
		switch (event.key) {
			case 'Escape':
				if (this.workspace.focusedRegion === 'inspector' || this.inspectorLines.length > 0) {
					this.inspectorLines = []
					this.workspace = { ...this.workspace, focusedRegion: 'grid' }
					this.message = 'Ready'
					return { handled: true, shouldRender: true, message: this.message }
				}
				return { handled: false, shouldRender: false }
			case 'Ctrl+C':
			case 'Cmd+C':
				return this.copySelection()
			case 'Ctrl+V':
			case 'Cmd+V':
				return this.pasteFromClipboard()
			case 'Ctrl+X':
			case 'Cmd+X':
				return this.cutSelection()
			case 'Ctrl+S':
			case 'Cmd+S':
				return this.executeCommand('save')
			case 'Ctrl+Shift+S':
				return this.showFileHubSection('export', 'Export: run :export <path>')
			case 'Ctrl+Z':
			case 'Cmd+Z':
				return this.undo()
			case 'Ctrl+Y':
			case 'Cmd+Y':
				return this.redo()
			case 'Ctrl+O':
			case 'Cmd+O':
				this.workspace = showFileHub(this.workspace, true)
				this.message = 'File hub'
				return { handled: true, shouldRender: true }
			case 'Ctrl+1':
				return this.executeRegisteredCommand('format')
			case 'Ctrl+Alt+V':
				return this.executeRegisteredCommand('paste values')
			case 'Ctrl+F':
			case 'Cmd+F':
			case 'Ctrl+H':
			case 'Shift+F5':
				return this.executeRegisteredCommand('find')
			case '/':
				return this.executeRegisteredCommand('find')
			case 'Ctrl+T':
				return this.executeRegisteredCommand('table create')
			case 'Ctrl+Shift+L':
				return this.executeRegisteredCommand('filter')
			case 'Shift+F2':
				return this.executeRegisteredCommand('comment')
			case 'Ctrl+PageUp':
				return this.selectSheetByOffset(-1)
			case 'Ctrl+PageDown':
				return this.selectSheetByOffset(1)
			case 'F9':
				this.calcWorker.next()
				this.message = 'Recalculation queued'
				return { handled: true, shouldRender: true }
			case 'F12':
				return this.showFileHubSection('saveAs', 'Save As: run :save-as <path>')
			case 'Alt+F1':
				return this.executeRegisteredCommand('chart')
			case 'Ctrl+`':
				return this.executeRegisteredCommand('show formulas')
			case 'Alt+=':
				return this.executeRegisteredCommand('autosum')
			case 'Ctrl+P':
			case ':':
				this.mode = 'command'
				this.commandBuffer = ''
				this.commandPaletteIndex = 0
				return { handled: true, shouldRender: true }
			case 'Ctrl+G':
			case 'F5':
				this.mode = 'command'
				this.commandBuffer = 'goto '
				this.commandPaletteIndex = 0
				this.message = 'Go To'
				return { handled: true, shouldRender: true, message: this.message }
			case '?':
			case 'F1':
				return this.showHelp()
			case 'F10':
				return this.showKeyTips()
			case 'Shift+F10':
			case 'Menu':
				return this.openContextMenu()
			case 'F2':
				return this.beginEdit('')
			case 'Enter':
				return this.move(1, 0)
			case 'Delete':
				return this.applyOperations('home.clear', [
					{
						op: 'clearRange',
						sheet: this.sheetName(),
						range: selectionRef(this.selection, indexToColumn),
						what: 'all',
					},
				])
			case 'ArrowUp':
				return this.move(-1, 0)
			case 'ArrowDown':
				return this.move(1, 0)
			case 'ArrowLeft':
				return this.move(0, -1)
			case 'ArrowRight':
			case 'Tab':
				return this.move(0, 1)
			case 'Ctrl+ArrowUp':
			case 'Cmd+ArrowUp':
				return this.moveToSheetBoundary('up')
			case 'Ctrl+ArrowDown':
			case 'Cmd+ArrowDown':
				return this.moveToSheetBoundary('down')
			case 'Ctrl+ArrowLeft':
			case 'Cmd+ArrowLeft':
				return this.moveToSheetBoundary('left')
			case 'Ctrl+ArrowRight':
			case 'Cmd+ArrowRight':
				return this.moveToSheetBoundary('right')
			case 'Ctrl+Shift+ArrowUp':
			case 'Cmd+Shift+ArrowUp':
				return this.moveToSheetBoundary('up', true)
			case 'Ctrl+Shift+ArrowDown':
			case 'Cmd+Shift+ArrowDown':
				return this.moveToSheetBoundary('down', true)
			case 'Ctrl+Shift+ArrowLeft':
			case 'Cmd+Shift+ArrowLeft':
				return this.moveToSheetBoundary('left', true)
			case 'Ctrl+Shift+ArrowRight':
			case 'Cmd+Shift+ArrowRight':
				return this.moveToSheetBoundary('right', true)
			case 'Shift+ArrowUp':
				return this.move(-1, 0, true)
			case 'Shift+ArrowDown':
				return this.move(1, 0, true)
			case 'Shift+ArrowLeft':
				return this.move(0, -1, true)
			case 'Shift+ArrowRight':
				return this.move(0, 1, true)
			case 'PageUp':
				return this.move(-this.viewport.visibleRows, 0)
			case 'PageDown':
				return this.move(this.viewport.visibleRows, 0)
			case 'Shift+PageUp':
				return this.move(-this.viewport.visibleRows, 0, true)
			case 'Shift+PageDown':
				return this.move(this.viewport.visibleRows, 0, true)
			case 'Alt+PageUp':
				return this.move(0, -this.viewport.visibleCols)
			case 'Alt+PageDown':
				return this.move(0, this.viewport.visibleCols)
			case 'Alt+ArrowLeft':
				return this.selectSheetByOffset(-1)
			case 'Alt+ArrowRight':
				return this.selectSheetByOffset(1)
			case 'Shift+Enter':
				return this.move(-1, 0)
			case 'Shift+Tab':
				return this.move(0, -1)
			case 'Ctrl+A':
			case 'Cmd+A':
				return this.selectCurrentRegionOrSheet()
			case 'Ctrl+Space':
				return this.selectCurrentColumn()
			case 'Shift+Space':
				return this.selectCurrentRow()
			case 'Ctrl+D':
			case 'Cmd+D':
				return this.fillSelection('down')
			case 'Ctrl+R':
			case 'Cmd+R':
				return this.fillSelection('right')
			case 'Alt+Down':
			case 'Alt+ArrowDown':
				return this.executeRegisteredCommand('filter')
			case 'Ctrl+Home':
				this.selection = selectCell(0, 0)
				this.viewport = ensureSelectionVisible(this.viewport, this.selection)
				this.message = selectionRef(this.selection, indexToColumn)
				return { handled: true, shouldRender: true, message: this.message }
			case 'Ctrl+End':
				return this.moveToUsedRangeEnd()
			case 'Ctrl+Shift+Home':
				this.selection = { ...this.selection, active: { row: 0, col: 0 }, kind: 'range' }
				this.viewport = ensureSelectionVisible(this.viewport, this.selection)
				this.message = selectionRef(this.selection, indexToColumn)
				return { handled: true, shouldRender: true, message: this.message }
			case 'Ctrl+Shift+End':
				return this.moveToUsedRangeEnd(true)
			case 'Home':
				this.selection = selectCell(this.selection.active.row, 0)
				this.viewport = ensureSelectionVisible(this.viewport, this.selection)
				return { handled: true, shouldRender: true }
			default:
				if (event.key === 'text' && (event.text === ':' || event.text === '/')) {
					if (event.text === '/') return this.executeRegisteredCommand('find')
					this.mode = 'command'
					this.commandBuffer = ''
					this.commandPaletteIndex = 0
					return { handled: true, shouldRender: true }
				}
				if (event.key === 'text' && event.text && isGridPasteText(event.text)) {
					return this.pasteGridText(event.text)
				}
				if (event.key === 'text' && event.text) return this.beginEdit(event.text)
				return { handled: false, shouldRender: false }
		}
	}

	private async dispatchFileHubKey(
		event: Extract<InputEvent, { kind: 'key' }>,
	): Promise<DispatchResult> {
		switch (event.key) {
			case 'Escape':
				if (this.workspace.documents.length === 0) {
					this.message = 'Open or create a workbook first.'
					return { handled: true, shouldRender: true, message: this.message }
				}
				this.workspace = showFileHub(this.workspace, false)
				this.message = 'Ready'
				return { handled: true, shouldRender: true, message: this.message }
			case 'ArrowLeft':
				return this.cycleFileHubSection(-1)
			case 'ArrowRight':
			case 'Tab':
				return this.cycleFileHubSection(1)
			case 'ArrowUp':
				return this.moveFileHubSelection(-1)
			case 'ArrowDown':
				return this.moveFileHubSelection(1)
			case 'Enter':
				return this.activateFileHubEntry()
			case 'o':
			case 'O':
				this.mode = 'command'
				this.commandBuffer = 'open '
				this.commandPaletteIndex = 0
				this.message = 'Open workbook path'
				return { handled: true, shouldRender: true, message: this.message }
			case 'n':
			case 'N':
				return this.executeCommand('new')
			case 's':
			case 'S':
				return this.showFileHubSection('saveAs', 'Save As: run :save-as <path>')
			case 'e':
			case 'E':
				return this.showFileHubSection('export', 'Export: run :export <path>')
			case 'i':
			case 'I':
				return this.showFileHubSection('info', 'Workbook info')
			case '?':
			case 'F1':
				return this.showHelp()
			default:
				return { handled: false, shouldRender: false }
		}
	}

	private dispatchEditKey(event: Extract<InputEvent, { kind: 'key' }>): DispatchResult {
		switch (event.key) {
			case 'Escape':
				if (this.editTarget) {
					this.selection = selectCell(this.editTarget.row, this.editTarget.col)
					this.viewport = ensureSelectionVisible(this.viewport, this.selection)
				}
				this.mode = 'ready'
				this.editBuffer = ''
				this.editCursor = 0
				this.editTarget = undefined
				this.pointReference = undefined
				this.message = 'Ready'
				return { handled: true, shouldRender: true }
			case 'Backspace':
				return this.updateEditState(backspace)
			case 'Delete':
				return this.updateEditState(deleteForward)
			case 'ArrowLeft':
				if (this.shouldUseFormulaPointMode()) return this.updateFormulaPointReference(0, -1)
				return this.updateEditState((state) => moveEditCursor(state, -1))
			case 'ArrowRight':
				if (this.shouldUseFormulaPointMode()) return this.updateFormulaPointReference(0, 1)
				return this.updateEditState((state) => moveEditCursor(state, 1))
			case 'ArrowUp':
				if (this.shouldUseFormulaPointMode()) return this.updateFormulaPointReference(-1, 0)
				return { handled: false, shouldRender: false }
			case 'ArrowDown':
				if (this.shouldUseFormulaPointMode()) return this.updateFormulaPointReference(1, 0)
				return { handled: false, shouldRender: false }
			case 'Home':
				return this.updateEditState((state) => moveEditCursor(state, 'home'))
			case 'End':
				return this.updateEditState((state) => moveEditCursor(state, 'end'))
			case 'F4':
				return this.updateEditState(cycleReferenceMode)
			case 'Enter':
				return this.commitEdit(1, 0)
			case 'Shift+Enter':
				return this.commitEdit(-1, 0)
			case 'Tab':
				return this.commitEdit(0, 1)
			case 'Shift+Tab':
				return this.commitEdit(0, -1)
			default:
				if (event.key === 'text' && event.text) {
					this.pointReference = undefined
					this.mode = this.editBuffer.startsWith('=') ? 'editing' : 'entering'
					return this.updateEditState((state) => insertText(state, event.text ?? ''))
				}
				return { handled: false, shouldRender: false }
		}
	}

	private async dispatchCommandKey(
		event: Extract<InputEvent, { kind: 'key' }>,
	): Promise<DispatchResult> {
		switch (event.key) {
			case 'Escape':
				this.mode = 'ready'
				this.commandBuffer = ''
				this.commandPaletteIndex = 0
				return { handled: true, shouldRender: true }
			case 'ArrowUp':
				return this.moveCommandPaletteSelection(-1)
			case 'ArrowDown':
				return this.moveCommandPaletteSelection(1)
			case 'Backspace':
				this.commandBuffer = this.commandBuffer.slice(0, -1)
				this.commandPaletteIndex = 0
				return { handled: true, shouldRender: true }
			case 'Enter': {
				const command = commandPaletteExecution(this.commandBuffer, this.commandPaletteIndex)
				this.commandBuffer = ''
				this.commandPaletteIndex = 0
				this.mode = 'ready'
				return this.executeCommand(command)
			}
			default:
				if (event.key === 'text' && event.text) {
					this.commandBuffer += event.text
					this.commandPaletteIndex = 0
					return { handled: true, shouldRender: true }
				}
				return { handled: false, shouldRender: false }
		}
	}

	private moveCommandPaletteSelection(delta: number): DispatchResult {
		const results = commandPaletteResults(this.commandBuffer)
		if (results.length === 0) {
			this.commandPaletteIndex = 0
			return { handled: true, shouldRender: true }
		}
		this.commandPaletteIndex = (this.commandPaletteIndex + delta + results.length) % results.length
		return { handled: true, shouldRender: true }
	}

	private async dispatchContextMenuKey(
		event: Extract<InputEvent, { kind: 'key' }>,
	): Promise<DispatchResult> {
		const menu = this.contextMenu
		if (!menu) return { handled: false, shouldRender: false }
		switch (event.key) {
			case 'Escape':
				this.contextMenu = undefined
				this.workspace = { ...this.workspace, focusedRegion: 'grid' }
				this.message = 'Ready'
				return { handled: true, shouldRender: true, message: this.message }
			case 'ArrowUp':
			case 'Shift+Tab':
				return this.moveContextMenuSelection(-1)
			case 'ArrowDown':
			case 'Tab':
				return this.moveContextMenuSelection(1)
			case 'Enter':
				return this.activateContextMenuItem()
			default:
				if (event.key === 'text' && event.text) {
					return this.activateContextMenuMnemonic(event.text)
				}
				return { handled: false, shouldRender: false }
		}
	}

	private dispatchKeyTipKey(event: Extract<InputEvent, { kind: 'key' }>): DispatchResult {
		const key = event.key === 'text' ? event.text?.toUpperCase() : event.key.toUpperCase()
		switch (key) {
			case 'ESCAPE':
				this.keyTipActive = false
				this.inspectorLines = []
				this.workspace = { ...this.workspace, focusedRegion: 'grid' }
				this.message = 'Ready'
				return { handled: true, shouldRender: true, message: this.message }
			case 'F':
				this.keyTipActive = false
				return this.showFileHubSection('recent', 'File backstage')
			case 'H':
				return this.openKeyTipCommandSearch('home.')
			case 'I':
				return this.openKeyTipCommandSearch('insert.')
			case 'M':
				return this.openKeyTipCommandSearch('formulas.')
			case 'D':
				return this.openKeyTipCommandSearch('data.')
			case 'R':
				return this.openKeyTipCommandSearch('review.')
			case 'W':
				return this.openKeyTipCommandSearch('view.')
			default:
				return { handled: false, shouldRender: false }
		}
	}

	private openKeyTipCommandSearch(query: string): DispatchResult {
		this.keyTipActive = false
		this.inspectorLines = []
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, visible: false },
			focusedRegion: 'grid',
		}
		this.mode = 'command'
		this.commandBuffer = query
		this.commandPaletteIndex = 0
		this.message = `KeyTip ${query.slice(0, -1).toUpperCase()}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private moveContextMenuSelection(delta: number): DispatchResult {
		const menu = this.contextMenu
		if (!menu || menu.items.length === 0) return { handled: true, shouldRender: true }
		this.contextMenu = {
			...menu,
			selectedIndex: (menu.selectedIndex + delta + menu.items.length) % menu.items.length,
		}
		return { handled: true, shouldRender: true }
	}

	private async activateContextMenuItem(): Promise<DispatchResult> {
		const menu = this.contextMenu
		const item = menu?.items[menu.selectedIndex]
		if (!item) return { handled: true, shouldRender: true }
		this.contextMenu = undefined
		this.workspace = { ...this.workspace, focusedRegion: 'grid' }
		return this.executeRegisteredCommand(item.command)
	}

	private async activateContextMenuMnemonic(text: string): Promise<DispatchResult> {
		const menu = this.contextMenu
		if (!menu) return { handled: false, shouldRender: false }
		const key = text.trim().toLowerCase()
		if (!key) return { handled: false, shouldRender: false }
		const index = menu.items.findIndex((item) => item.title.toLowerCase().startsWith(key))
		if (index < 0) return { handled: false, shouldRender: false }
		this.contextMenu = { ...menu, selectedIndex: index }
		return this.activateContextMenuItem()
	}

	private dispatchMouse(event: Extract<InputEvent, { kind: 'mouse' }>): DispatchResult {
		if (event.action === 'wheel') {
			const delta = event.button === 64 ? -3 : 3
			this.viewport = { ...this.viewport, topRow: Math.max(0, this.viewport.topRow + delta) }
			return { handled: true, shouldRender: true }
		}
		if (event.action === 'press' && !this.workspace.fileHub.visible) {
			const gridRow = event.row - 5
			const gridCol = this.visibleColumnFromTerminalCol(event.col)
			if (gridRow >= 0 && gridCol !== null) {
				this.selection = selectCell(this.viewport.topRow + gridRow, gridCol)
				if (event.button === 2) return this.openContextMenu()
				return { handled: true, shouldRender: true }
			}
		}
		return { handled: false, shouldRender: false }
	}

	private visibleColumnFromTerminalCol(col: number): number | null {
		let start = 7
		for (let visibleCol = 0; visibleCol < this.viewport.visibleCols; visibleCol++) {
			const width = this.viewport.columnWidths[visibleCol] ?? 10
			const end = start + width
			if (col >= start && col <= end) return this.viewport.leftCol + visibleCol
			start = end + 1
		}
		return null
	}

	private dispatchDialogKey(event: Extract<InputEvent, { kind: 'key' }>): DispatchResult {
		const dialog = this.activeDialog
		if (!dialog) return { handled: false, shouldRender: false }
		switch (event.key) {
			case 'Escape':
				this.activeDialog = undefined
				this.workspace = { ...this.workspace, focusedRegion: 'grid' }
				this.message = 'Ready'
				return { handled: true, shouldRender: true }
			case 'ArrowUp':
			case 'Shift+Tab':
				return this.updateDialog({ activeField: previousDialogField(dialog) })
			case 'ArrowDown':
			case 'Tab':
				return this.updateDialog({ activeField: nextDialogField(dialog) })
			case 'ArrowLeft':
				return this.cycleDialogSelect(-1)
			case 'ArrowRight':
				return this.cycleDialogSelect(1)
			case 'Backspace':
				return this.editDialogField((value) => value.slice(0, -1))
			case 'Delete':
				return this.editDialogField(() => '')
			case ' ':
			case 'Space':
				return this.toggleDialogBoolean()
			case 'Enter':
				return this.applyActiveDialog()
			default:
				if (event.key === 'text' && event.text) {
					const field = dialog.fields[dialog.activeField]
					if (event.text === ' ' && field?.kind === 'boolean') return this.toggleDialogBoolean()
					if (event.text === ' ' && field?.kind === 'select') return this.cycleDialogSelect(1)
					return this.editDialogField((value) => value + event.text)
				}
				return { handled: false, shouldRender: false }
		}
	}

	private async executeCommand(raw: string): Promise<DispatchResult> {
		const command = raw.trim()
		if (!command) return { handled: true, shouldRender: true }
		const [name = '', ...rest] = command.split(/\s+/)
		const arg = rest.join(' ')
		switch (name.toLowerCase()) {
			case 'q!':
			case 'quit!':
				return { handled: true, shouldRender: false, shouldExit: true }
			case 'q':
			case 'quit':
				if (this.session.isDirty()) return this.unsavedChanges('quit', 'q!')
				return { handled: true, shouldRender: false, shouldExit: true }
			case 'wq': {
				const save = await this.executeCommand('save')
				return { ...save, shouldExit: save.handled && !this.session.isDirty() }
			}
			case 'new!':
				this.activateDocument(this.session.createEmpty(), 'Created Book1')
				return { handled: true, shouldRender: true, message: this.message }
			case 'new':
				if (this.session.isDirty()) return this.unsavedChanges('create a new workbook', 'new!')
				this.activateDocument(this.session.createEmpty(), 'Created Book1')
				return { handled: true, shouldRender: true, message: this.message }
			case 'open!':
				return this.openWorkbook(arg, { force: true })
			case 'open':
				return this.openWorkbook(arg)
			case 'file':
			case 'backstage':
				return this.showFileHubSection('recent', 'File backstage')
			case 'info':
				return this.showFileHubSection('info', 'Workbook info')
			case 'save':
			case 'w':
				return this.saveAsync()
			case 'save-as':
			case 'saveas':
				return this.saveAsAsync(arg)
			case 'save-copy':
			case 'save-copy-as':
				return this.saveCopyAsync(arg)
			case 'export':
				return this.exportAsync(arg)
			case 'sheet':
				return this.selectSheetByName(arg)
			case 'perf':
				this.message = this.perfSummary()
				return { handled: true, shouldRender: true }
			case 'status':
			case 'health':
				return this.showStatusInspector()
			case 'help':
			case '?':
				return this.showHelp()
			case 'keytips':
			case 'keys':
				return this.showKeyTips()
			case 'calibrate':
			case 'calibration':
				return this.showTerminalCalibration()
			case 'goto':
			case 'go':
				return this.gotoRef(arg)
			default:
				return this.executeRegisteredCommand(command)
		}
	}

	private async executeRegisteredCommand(query: string): Promise<DispatchResult> {
		const parsed = parseRegisteredCommand(query)
		if (parsed.error) {
			this.message = parsed.error
			return { handled: false, shouldRender: true, message: this.message }
		}
		const descriptor = findCommand(parsed.name)
		if (!descriptor) {
			this.message = `Unknown command: ${query}`
			return { handled: false, shouldRender: true, message: this.message }
		}
		if (!descriptor.contexts.includes(this.workspace.focusedRegion)) {
			this.message = `${descriptor.title} is not available while ${focusRegionLabel(this.workspace.focusedRegion)} is focused.`
			return { handled: false, shouldRender: true, message: this.message }
		}
		const native = await this.executeNativeRegisteredCommand(descriptor.id)
		if (native) return native
		if (descriptor.id === 'view.showFormulas') {
			this.showFormulas = !this.showFormulas
			this.message = this.showFormulas ? 'Showing formulas' : 'Showing values'
			return { handled: true, shouldRender: true, message: this.message }
		}
		if (descriptor.id === 'view.objects') {
			return this.showObjectInspector()
		}
		if (descriptor.id === 'formulas.tracePrecedents') {
			return this.showFormulaTrace(parsed.input)
		}
		if (descriptor.dialogId && parsed.input === undefined) {
			return this.openDialog(descriptor.dialogId)
		}
		if (descriptor.id === 'home.findReplace') {
			try {
				buildDialogOperations(
					'find-replace',
					{ sheet: this.sheetName(), selection: this.selection, input: parsed.input },
					parsed.input,
				)
			} catch (error) {
				this.message = error instanceof Error ? error.message : String(error)
				return { handled: false, shouldRender: true, message: this.message }
			}
			return this.executeFindReplace(parsed.input, descriptor.id)
		}
		let operations: readonly Operation[]
		try {
			operations = descriptor.toOperations(
				{ sheet: this.sheetName(), selection: this.selection, input: parsed.input },
				parsed.input,
			)
		} catch (error) {
			this.message = error instanceof Error ? error.message : String(error)
			return { handled: false, shouldRender: true, message: this.message }
		}
		if (operations.length === 0) {
			this.message = 'No changes.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		return this.applyOperations(descriptor.id, operations)
	}

	private async executeNativeRegisteredCommand(id: string): Promise<DispatchResult | null> {
		switch (id) {
			case 'file.open':
				this.workspace = showFileHub(this.workspace, true)
				this.message = 'File hub'
				return { handled: true, shouldRender: true, message: this.message }
			case 'file.save':
				return this.saveAsync()
			case 'file.saveAs':
				return this.showFileHubSection('saveAs', 'Save As: run :save-as <path>')
			case 'file.export':
				return this.showFileHubSection('export', 'Export: run :export <path>')
			case 'home.copy':
				return this.copySelection()
			case 'home.cut':
				return this.cutSelection()
			case 'home.paste':
				return this.pasteFromClipboard()
			case 'formulas.autosum':
				return this.autoSum()
			case 'formulas.recalculate':
				this.calcWorker.next()
				this.message = 'Recalculation queued'
				return { handled: true, shouldRender: true, message: this.message }
			case 'view.freeze':
				return this.freezeAtSelection()
			default:
				return null
		}
	}

	private showFileHubSection(
		section: WorkbookWorkspace['fileHub']['section'],
		message: string,
	): DispatchResult {
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, visible: true, section, selectedIndex: 0 },
			focusedRegion: 'fileHub',
		}
		this.message = message
		return { handled: true, shouldRender: true, message }
	}

	private openContextMenu(): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		this.contextMenu = {
			target: contextMenuTarget(this.selection),
			address: selectionRef(this.selection, indexToColumn),
			selectedIndex: 0,
			items: contextMenuItems(this.selection),
		}
		this.workspace = { ...this.workspace, focusedRegion: 'contextMenu' }
		this.mode = 'ready'
		this.commandBuffer = ''
		this.commandPaletteIndex = 0
		this.inspectorLines = []
		this.message = 'Context menu'
		return { handled: true, shouldRender: true, message: this.message }
	}

	private recordRecentWorkbook(path: string): void {
		try {
			this.recentStore?.record(path)
			this.refreshRecentWorkbooks()
		} catch {
			this.recentStore = undefined
		}
	}

	private refreshRecentWorkbooks(): void {
		const entries = this.recentStore?.entries() ?? this.workspace.fileHub.entries
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, entries },
		}
	}

	private cycleFileHubSection(delta: number): DispatchResult {
		const sections: readonly WorkbookWorkspace['fileHub']['section'][] = [
			'recent',
			'open',
			'new',
			'saveAs',
			'export',
			'recover',
			'info',
		]
		const current = sections.indexOf(this.workspace.fileHub.section)
		const next = sections[(current + delta + sections.length) % sections.length] ?? 'recent'
		return this.showFileHubSection(next, `File ${next === 'saveAs' ? 'Save As' : next}`)
	}

	private moveFileHubSelection(delta: number): DispatchResult {
		const count = fileHubEntryCount(this.workspace.fileHub)
		const selectedIndex =
			count <= 0 ? 0 : (this.workspace.fileHub.selectedIndex + delta + count) % count
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, selectedIndex },
		}
		return { handled: true, shouldRender: true }
	}

	private async activateFileHubEntry(): Promise<DispatchResult> {
		const section = this.workspace.fileHub.section
		const index = this.workspace.fileHub.selectedIndex
		const recent = this.workspace.fileHub.entries[index]
		if (section === 'recent' && recent?.path) {
			this.mode = 'command'
			this.commandBuffer = `open ${recent.path}`
			this.commandPaletteIndex = 0
			this.message = 'Open recent workbook'
			return { handled: true, shouldRender: true, message: this.message }
		}
		if (section === 'new') {
			return this.executeCommand(index === 1 ? 'new!' : 'new')
		}
		if (section === 'info') {
			return this.showObjectInspector()
		}
		const commandBySection: Partial<Record<WorkbookWorkspace['fileHub']['section'], string>> = {
			open: 'open ',
			saveAs: 'save-as ',
			export: 'export ',
			recover: 'recover',
			recent: 'open ',
		}
		const command = commandBySection[section] ?? 'open '
		this.mode = 'command'
		this.commandBuffer = command
		this.commandPaletteIndex = 0
		this.message = command.trim() ? `Command: ${command.trim()}` : 'Command'
		return { handled: true, shouldRender: true, message: this.message }
	}

	private openDialog(id: string): DispatchResult {
		const dialog = findDialog(id as DialogId)
		if (!dialog) {
			this.message = `Unknown dialog: ${id}`
			return { handled: false, shouldRender: true, message: this.message }
		}
		const defaultInput = dialog.defaultInput({
			sheet: this.sheetName(),
			selection: this.selection,
		})
		this.activeDialog = {
			id: dialog.id,
			title: dialog.title,
			activeField: 0,
			fields: dialog.fields.map((field) => ({
				name: field.name,
				label: field.label,
				kind: field.kind,
				required: field.required ?? false,
				options: field.options ?? [],
				value: initialDialogFieldValue(dialog.id, field.name, defaultInput),
			})),
		}
		this.mode = 'ready'
		this.commandBuffer = ''
		this.commandPaletteIndex = 0
		this.workspace = { ...this.workspace, focusedRegion: 'dialog' }
		this.message = `${dialog.title}: edit fields, press Enter to apply, or run command with JSON input.`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private updateDialog(update: Partial<DialogViewState>): DispatchResult {
		if (!this.activeDialog) return { handled: false, shouldRender: false }
		this.activeDialog = { ...this.activeDialog, ...update }
		return { handled: true, shouldRender: true }
	}

	private updateDialogFields(
		update: (fields: DialogViewState['fields']) => DialogViewState['fields'],
	): DispatchResult {
		if (!this.activeDialog) return { handled: false, shouldRender: false }
		this.activeDialog = { ...this.activeDialog, fields: update(this.activeDialog.fields) }
		return { handled: true, shouldRender: true }
	}

	private editDialogField(update: (value: string) => string): DispatchResult {
		const dialog = this.activeDialog
		if (!dialog) return { handled: false, shouldRender: false }
		const field = dialog.fields[dialog.activeField]
		if (!field || field.kind === 'boolean' || field.kind === 'select') {
			return { handled: false, shouldRender: false }
		}
		return this.updateDialogFields((fields) =>
			fields.map((entry, index) =>
				index === dialog.activeField ? { ...entry, value: update(entry.value) } : entry,
			),
		)
	}

	private toggleDialogBoolean(): DispatchResult {
		const dialog = this.activeDialog
		if (!dialog) return { handled: false, shouldRender: false }
		const field = dialog.fields[dialog.activeField]
		if (!field || field.kind !== 'boolean') return { handled: false, shouldRender: false }
		return this.updateDialogFields((fields) =>
			fields.map((entry, index) =>
				index === dialog.activeField
					? { ...entry, value: entry.value === 'true' ? 'false' : 'true' }
					: entry,
			),
		)
	}

	private cycleDialogSelect(delta: number): DispatchResult {
		const dialog = this.activeDialog
		if (!dialog) return { handled: false, shouldRender: false }
		const field = dialog.fields[dialog.activeField]
		if (!field || field.kind !== 'select' || field.options.length === 0) {
			return { handled: false, shouldRender: false }
		}
		const currentIndex = field.options.indexOf(field.value)
		const current = currentIndex >= 0 ? currentIndex : delta > 0 ? -1 : 0
		const next = (current + delta + field.options.length) % field.options.length
		return this.updateDialogFields((fields) =>
			fields.map((entry, index) =>
				index === dialog.activeField ? { ...entry, value: field.options[next] ?? '' } : entry,
			),
		)
	}

	private applyActiveDialog(): DispatchResult {
		const dialog = this.activeDialog
		if (!dialog) return { handled: false, shouldRender: false }
		try {
			const input = dialogInputFromValues(dialog)
			if (dialog.id === 'find-replace') {
				buildDialogOperations(
					'find-replace',
					{ sheet: this.sheetName(), selection: this.selection, input },
					input,
				)
				const result = this.executeFindReplace(input, `dialog.${dialog.id}`)
				if (result.handled) {
					this.activeDialog = undefined
					this.workspace = { ...this.workspace, focusedRegion: 'grid' }
				}
				return result
			}
			const operations = buildDialogOperations(
				dialog.id as DialogId,
				{
					sheet: this.sheetName(),
					selection: this.selection,
					input,
				},
				input,
			)
			if (operations.length === 0) {
				this.message = 'No changes.'
				return { handled: true, shouldRender: true, message: this.message }
			}
			this.activeDialog = undefined
			this.workspace = { ...this.workspace, focusedRegion: 'grid' }
			return this.applyOperations(`dialog.${dialog.id}`, operations)
		} catch (error) {
			this.message = error instanceof Error ? error.message : String(error)
			return { handled: false, shouldRender: true, message: this.message }
		}
	}

	private beginEdit(prefix: string): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true }
		}
		if (this.activeDocumentReadOnly()) {
			this.message = 'Workbook is read-only. Reopen with a full load before editing.'
			return { handled: false, shouldRender: true, message: this.message }
		}
		this.mode = prefix === '=' ? 'editing' : 'entering'
		this.editBuffer = prefix === '' ? this.activeCellEditText() : prefix
		this.editCursor = this.editBuffer.length
		this.editTarget = this.selection.active
		this.pointReference = undefined
		return { handled: true, shouldRender: true }
	}

	private updateEditState(update: (state: FormulaEditState) => FormulaEditState): DispatchResult {
		const next = update(createFormulaEditState(this.editBuffer, this.editCursor))
		this.editBuffer = next.buffer
		this.editCursor = next.cursor
		return { handled: true, shouldRender: true }
	}

	private commitEdit(deltaRow: number, deltaCol: number): DispatchResult {
		const input = this.editBuffer
		const target = this.editTarget ?? this.selection.active
		const ref = cellCoordRef(target)
		const operations: readonly Operation[] = input.startsWith('=')
			? [{ op: 'setFormula', sheet: this.sheetName(), ref, formula: input }]
			: [
					{ op: 'clearRange', sheet: this.sheetName(), range: ref, what: 'formulas' },
					{
						op: 'setCells',
						sheet: this.sheetName(),
						updates: [{ ref, value: parseInputValue(input) }],
					},
				]
		const result = this.applyOperations('home.edit', operations)
		if (!result.handled) return result
		this.mode = 'ready'
		this.editBuffer = ''
		this.editCursor = 0
		this.editTarget = undefined
		this.pointReference = undefined
		this.selection = selectCell(target.row, target.col)
		this.move(deltaRow, deltaCol)
		return result
	}

	private shouldUseFormulaPointMode(): boolean {
		if (!this.editBuffer.startsWith('=') || this.editCursor !== this.editBuffer.length) return false
		if (this.pointReference) return true
		return /[=(:,+\-*/^&\s]$/.test(this.editBuffer)
	}

	private updateFormulaPointReference(deltaRow: number, deltaCol: number): DispatchResult {
		this.selection = moveSelection(this.selection, deltaRow, deltaCol)
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		const ref = activeCellRef(this.selection)
		const replace = this.pointReference
		const start = replace?.start ?? this.editCursor
		const end = replace?.end ?? this.editCursor
		this.editBuffer = `${this.editBuffer.slice(0, start)}${ref}${this.editBuffer.slice(end)}`
		this.editCursor = start + ref.length
		this.pointReference = { start, end: this.editCursor }
		this.mode = 'point'
		this.message = `Point: ${ref}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private async copySelection(): Promise<DispatchResult> {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true }
		}
		const text = this.selectionToTsv()
		const result = await this.clipboard.writeText({ text, mode: 'copy' })
		const cellCount =
			text === '' ? 1 : text.split('\n').reduce((sum, row) => sum + row.split('\t').length, 0)
		this.message = result.system
			? `Copied ${cellCount} cell${cellCount === 1 ? '' : 's'} to clipboard.`
			: `Copied ${cellCount} cell${cellCount === 1 ? '' : 's'} internally; OSC52 fallback prepared.`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private async cutSelection(): Promise<DispatchResult> {
		const copy = await this.copySelection()
		if (!copy.handled || !this.hasWorkbook()) return copy
		const range = selectionRef(this.selection, indexToColumn)
		const clear = this.applyOperations('home.cut', [
			{ op: 'clearRange', sheet: this.sheetName(), range, what: 'all' },
		])
		if (clear.handled) this.message = `Cut ${range}`
		return { ...clear, message: this.message }
	}

	private autoSum(): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const target = this.selection.active
		const ref = cellCoordRef(target)
		const range = autoSumRange(target)
		if (!range) {
			this.message = 'AutoSum needs cells above or to the left.'
			return { handled: false, shouldRender: true, message: this.message }
		}
		return this.applyOperations(
			'formulas.autosum',
			[{ op: 'setFormula', sheet: this.sheetName(), ref, formula: `=SUM(${range})` }],
			{ message: `AutoSum ${range}` },
		)
	}

	private freezeAtSelection(): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const row = this.selection.active.row
		const col = this.selection.active.col
		const message =
			row === 0 && col === 0
				? 'Unfroze panes'
				: `Frozen panes above row ${row + 1} and left of ${indexToColumn(col)}`
		return this.applyOperations(
			'view.freeze',
			[{ op: 'freezePane', sheet: this.sheetName(), row, col }],
			{ message },
		)
	}

	private async pasteFromClipboard(): Promise<DispatchResult> {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true }
		}
		const clipboard = await this.clipboard.readText()
		if (clipboard.text === '') {
			this.message = 'Clipboard is empty.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const result = this.pasteGridText(clipboard.text)
		if (result.handled && clipboard.source === 'internal') {
			this.message = `${this.message} (internal clipboard)`
		}
		return result
	}

	private selectionToTsv(): string {
		const range = selectionRef(this.selection, indexToColumn)
		const data = this.session.requireWorkbook().readRangeCompact(this.sheetName(), range, {
			includeRefs: true,
		})
		const cells = new Map((data?.cells ?? []).map((cell) => [`${cell.row}:${cell.col}`, cell]))
		const startRow = Math.min(this.selection.anchor.row, this.selection.active.row)
		const endRow = Math.max(this.selection.anchor.row, this.selection.active.row)
		const startCol = Math.min(this.selection.anchor.col, this.selection.active.col)
		const endCol = Math.max(this.selection.anchor.col, this.selection.active.col)
		const rows: string[] = []
		for (let row = startRow; row <= endRow; row++) {
			const values: string[] = []
			for (let col = startCol; col <= endCol; col++) {
				const cell = cells.get(`${row}:${col}`)
				values.push(
					cell?.formula ? `=${cell.formula}` : displayCellValue(cell?.value ?? { kind: 'empty' }),
				)
			}
			rows.push(values.join('\t'))
		}
		return rows.join('\n')
	}

	private pasteGridText(text: string): DispatchResult {
		const rows = parseGridPaste(text)
		if (rows.length === 0) return { handled: true, shouldRender: false }
		const start = this.selection.active
		const width = Math.max(...rows.map((row) => row.length))
		const height = rows.length
		const range = `${indexToColumn(start.col)}${start.row + 1}:${indexToColumn(start.col + width - 1)}${start.row + height}`
		const valueUpdates: Extract<Operation, { op: 'setCells' }>['updates'][number][] = []
		const operations: Operation[] = []
		let pastedCells = 0
		for (let r = 0; r < rows.length; r++) {
			const row = rows[r] ?? []
			for (let c = 0; c < row.length; c++) {
				const raw = row[c] ?? ''
				const ref = `${indexToColumn(start.col + c)}${start.row + r + 1}`
				pastedCells += 1
				if (raw === '') {
					operations.push({ op: 'clearRange', sheet: this.sheetName(), range: ref, what: 'values' })
					operations.push({
						op: 'clearRange',
						sheet: this.sheetName(),
						range: ref,
						what: 'formulas',
					})
				} else if (raw.startsWith('=')) {
					operations.push({ op: 'setFormula', sheet: this.sheetName(), ref, formula: raw })
				} else {
					operations.push({
						op: 'clearRange',
						sheet: this.sheetName(),
						range: ref,
						what: 'formulas',
					})
					valueUpdates.push({ ref, value: parseInputValue(raw) })
				}
			}
		}
		if (valueUpdates.length > 0) {
			operations.push({ op: 'setCells', sheet: this.sheetName(), updates: valueUpdates })
		}
		const result = this.applyOperations('home.paste', operations)
		if (result.handled) {
			this.selection = {
				active: start,
				anchor: { row: start.row + height - 1, col: start.col + width - 1 },
				kind: height > 1 || width > 1 ? 'range' : 'cell',
			}
			this.viewport = ensureSelectionVisible(this.viewport, this.selection)
			this.message = `Pasted ${pastedCells} cells into ${range}`
		}
		return result
	}

	private executeFindReplace(input: unknown, commandId: string): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const request = normalizeFindReplaceInput(input, this.defaultFindRange())
		let matches: readonly FindReplaceMatch[]
		try {
			matches = this.findReplaceMatches(request)
		} catch (error) {
			this.message = error instanceof Error ? error.message : String(error)
			return { handled: false, shouldRender: true, message: this.message }
		}
		if (matches.length === 0) {
			this.message = `No matches for "${request.findText}".`
			return { handled: true, shouldRender: true, message: this.message }
		}
		const first = matches[0]
		if (!first) {
			this.message = `No matches for "${request.findText}".`
			return { handled: true, shouldRender: true, message: this.message }
		}
		const selectionBefore = this.selection
		this.selection = selectCell(first.row, first.col)
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		if (request.action === 'find') {
			this.message = `Found "${request.findText}" at ${first.ref}.`
			return { handled: true, shouldRender: true, message: this.message }
		}
		const replacementMatches = request.action === 'replaceAll' ? matches : [first]
		const operations = buildFindReplaceOperations(this.sheetName(), replacementMatches)
		if (operations.length === 0) {
			this.message = `No replaceable matches for "${request.findText}".`
			return { handled: true, shouldRender: true, message: this.message }
		}
		const result = this.applyOperations(commandId, operations, {
			selectionBefore,
			message: `Replaced ${replacementMatches.length} match${replacementMatches.length === 1 ? '' : 'es'} for "${request.findText}".`,
		})
		if (!result.handled) this.selection = selectionBefore
		return result
	}

	private defaultFindRange(): string {
		if (this.selection.kind === 'range') return selectionRef(this.selection, indexToColumn)
		const used = this.session.requireWorkbook().inspectSheet(this.sheetName())?.usedRange
		if (!used) return selectionRef(this.selection, indexToColumn)
		return `${toA1(used.start)}:${toA1(used.end)}`
	}

	private showObjectInspector(): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const info = this.session.inspect()
		this.inspectorLines = objectInspectorLines(info)
		this.workspace = { ...this.workspace, focusedRegion: 'inspector' }
		this.message = `Objects: ${info?.chartCount ?? 0} charts, ${info?.pivotTableCount ?? 0} pivots`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private showTerminalCalibration(): DispatchResult {
		const report = buildTerminalCalibrationReport(detectTerminalCapabilities())
		this.inspectorLines = report.lines
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, visible: false },
			focusedRegion: 'inspector',
		}
		this.message = `Terminal calibration: ${report.profile}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private showHelp(): DispatchResult {
		this.inspectorLines = helpInspectorLines()
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, visible: false },
			focusedRegion: 'inspector',
		}
		this.message = 'Help: Excel-style navigation and commands'
		return { handled: true, shouldRender: true, message: this.message }
	}

	private showStatusInspector(): DispatchResult {
		this.inspectorLines = statusInspectorLines({
			info: this.session.inspect(),
			document: activeDocument(this.workspace),
			sheetName: this.sheetName(),
			selection: selectionRef(this.selection, indexToColumn),
			dirty: this.session.isDirty(),
			perf: this.perfSummary(),
		})
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, visible: false },
			focusedRegion: 'inspector',
		}
		this.message = 'Workbook Health'
		return { handled: true, shouldRender: true, message: this.message }
	}

	private showKeyTips(): DispatchResult {
		this.inspectorLines = keyTipInspectorLines()
		this.keyTipActive = true
		this.workspace = {
			...this.workspace,
			fileHub: { ...this.workspace.fileHub, visible: false },
			focusedRegion: 'ribbon',
		}
		this.message = 'KeyTips: F File  H Home  I Insert  M Formulas  D Data  R Review  W View'
		return { handled: true, shouldRender: true, message: this.message }
	}

	private showFormulaTrace(input: unknown): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const options = normalizeFormulaTraceInput(input)
		if (options.error) {
			this.message = options.error
			return { handled: false, shouldRender: true, message: this.message }
		}
		const ref = activeCellRef(this.selection)
		const trace = this.session
			.requireWorkbook()
			.trace(
				`${this.sheetName()}!${ref}`,
				options.maxDepth !== undefined ? { maxDepth: options.maxDepth } : undefined,
			)
		if (!trace) {
			this.message = `Could not trace ${this.sheetName()}!${ref}.`
			return { handled: false, shouldRender: true, message: this.message }
		}
		this.inspectorLines = formulaTraceInspectorLines(trace)
		this.workspace = { ...this.workspace, focusedRegion: 'inspector' }
		this.message = `Trace: ${trace.precedents.length} precedents, ${trace.dependents.length} dependents`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private findReplaceMatches(input: FindReplaceInput): readonly FindReplaceMatch[] {
		const range = input.range || selectionRef(this.selection, indexToColumn)
		const data = this.session.requireWorkbook().readRangeCompact(this.sheetName(), range, {
			includeRefs: true,
		})
		const matches: FindReplaceMatch[] = []
		for (const cell of data?.cells ?? []) {
			const ref = cell.ref ?? `${indexToColumn(cell.col)}${cell.row + 1}`
			if (input.lookIn !== 'values' && cell.formula) {
				const formulaReplacement = replaceFormulaText(cell.formula, input)
				if (formulaReplacement !== null) {
					matches.push({
						ref,
						row: cell.row,
						col: cell.col,
						kind: 'formula',
						replacement: formulaReplacement,
					})
					continue
				}
			}
			if (input.lookIn !== 'formulas') {
				const valueText = displayCellValue(cell.value)
				const valueReplacement = replaceText(valueText, input)
				if (valueReplacement !== null) {
					matches.push({
						ref,
						row: cell.row,
						col: cell.col,
						kind: 'value',
						replacement: valueReplacement,
					})
				}
			}
		}
		return matches
	}

	private applyOperations(
		commandId: string,
		operations: readonly Operation[],
		options: { readonly selectionBefore?: SelectionState; readonly message?: string } = {},
	): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true }
		}
		if (this.activeDocumentReadOnly()) {
			this.message = 'Workbook is read-only. Reopen with a full load before editing.'
			return { handled: false, shouldRender: true, message: this.message }
		}
		const before = options.selectionBefore ?? this.selection
		const inverseOps = this.buildInverseOperations(operations)
		let result: ReturnType<WorkbookSessionController['applyAndRecalc']>
		try {
			result = this.session.applyAndRecalc(operations)
		} catch (error) {
			this.message = error instanceof Error ? error.message : String(error)
			return { handled: false, shouldRender: true, message: this.message }
		}
		const applyErrors = result?.apply.errors ?? []
		const recalcErrors = result?.recalc?.errors.map((issue) => issue.error) ?? []
		const errors = [...applyErrors, ...recalcErrors]
		if (errors.length > 0) {
			this.message = errors.map((error) => error.message).join('; ')
			return { handled: false, shouldRender: true, message: this.message }
		}
		this.hydration.invalidate()
		this.workspace = updateActiveDocument(this.workspace, {
			dirty: this.session.isDirty(),
			info: this.session.inspect(),
		})
		this.journal.append({
			commandId,
			selectionBefore: before,
			selectionAfter: this.selection,
			ops: operations,
			inverseOps,
			affectedCells: [],
			recalcDirtyRefs: [],
		})
		this.calcWorker.next()
		this.message =
			options.message ??
			`Committed ${operations.length} operation${operations.length === 1 ? '' : 's'}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private undo(): DispatchResult {
		const entry = this.journal.undo()
		if (!entry) {
			this.message = 'Nothing to undo'
			return { handled: true, shouldRender: true, message: this.message }
		}
		if (entry.inverseOps.length === 0) {
			this.journal.redo()
			this.message = `${entry.commandId} cannot be undone yet.`
			return { handled: true, shouldRender: true, message: this.message }
		}
		const result = this.applyWithoutJournal(entry.inverseOps)
		if (!result.handled) {
			this.journal.redo()
			return result
		}
		this.selection = entry.selectionBefore
		this.message = `Undid ${entry.commandId}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private redo(): DispatchResult {
		const entry = this.journal.redo()
		if (!entry) {
			this.message = 'Nothing to redo'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const result = this.applyWithoutJournal(entry.ops)
		if (!result.handled) {
			this.journal.undo()
			return result
		}
		this.selection = entry.selectionAfter
		this.message = `Redid ${entry.commandId}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private applyWithoutJournal(operations: readonly Operation[]): DispatchResult {
		let result: ReturnType<WorkbookSessionController['applyAndRecalc']>
		try {
			result = this.session.applyAndRecalc(operations)
		} catch (error) {
			this.message = error instanceof Error ? error.message : String(error)
			return { handled: false, shouldRender: true, message: this.message }
		}
		const applyErrors = result?.apply.errors ?? []
		const recalcErrors = result?.recalc?.errors.map((issue) => issue.error) ?? []
		const errors = [...applyErrors, ...recalcErrors]
		if (errors.length > 0) {
			this.message = errors.map((error) => error.message).join('; ')
			return { handled: false, shouldRender: true, message: this.message }
		}
		this.hydration.invalidate()
		this.workspace = updateActiveDocument(this.workspace, {
			dirty: this.session.isDirty(),
			info: this.session.inspect(),
		})
		this.calcWorker.next()
		return { handled: true, shouldRender: true }
	}

	private buildInverseOperations(operations: readonly Operation[]): readonly Operation[] {
		const rangesBySheet = new Map<string, Set<string>>()
		const inverse: Operation[] = []
		for (const operation of operations) {
			if (operation.op === 'setCells') {
				const ranges = rangesBySheet.get(operation.sheet) ?? new Set<string>()
				for (const update of operation.updates) ranges.add(update.ref)
				rangesBySheet.set(operation.sheet, ranges)
			} else if (operation.op === 'setFormula') {
				const ranges = rangesBySheet.get(operation.sheet) ?? new Set<string>()
				ranges.add(operation.ref)
				rangesBySheet.set(operation.sheet, ranges)
			} else if (operation.op === 'clearRange') {
				const ranges = rangesBySheet.get(operation.sheet) ?? new Set<string>()
				ranges.add(operation.range)
				rangesBySheet.set(operation.sheet, ranges)
			} else if (operation.op === 'copyRange' || operation.op === 'moveRange') {
				const ranges = rangesBySheet.get(operation.sheet) ?? new Set<string>()
				ranges.add(copyTargetRange(operation.source, operation.target))
				rangesBySheet.set(operation.sheet, ranges)
			} else if (operation.op === 'setStyle' || operation.op === 'setNumberFormat') {
				inverse.push(...this.inverseStyleOperations(operation.sheet, operation.range))
			} else if (operation.op === 'setComment') {
				inverse.push(this.inverseCommentOperation(operation.sheet, operation.ref))
			} else if (operation.op === 'deleteComment') {
				const comment = this.readComment(operation.sheet, operation.ref)
				if (comment) {
					inverse.push({
						op: 'setComment',
						sheet: operation.sheet,
						ref: operation.ref,
						text: comment.text,
						...(comment.author ? { author: comment.author } : {}),
					})
				}
			} else if (operation.op === 'createTable') {
				inverse.push({ op: 'deleteTable', table: operation.name })
			} else if (operation.op === 'setAutoFilter' || operation.op === 'clearAutoFilter') {
				inverse.push(this.inverseAutoFilterOperation(operation.sheet))
			} else if (operation.op === 'setPrintArea') {
				inverse.push(this.inversePrintAreaOperation(operation.sheet))
			} else if (operation.op === 'setPageSetup') {
				inverse.push(this.inversePageSetupOperation(operation.sheet))
			} else if (operation.op === 'setDataValidation') {
				inverse.push(...this.inverseDataValidationOperations(operation.sheet, operation.range))
			} else if (operation.op === 'deleteDataValidation') {
				inverse.push(...this.restoreDataValidationOperations(operation.sheet, operation.range))
			} else if (operation.op === 'setConditionalFormat') {
				inverse.push(...this.inverseConditionalFormatOperations(operation.sheet, operation.range))
			} else if (operation.op === 'deleteConditionalFormat') {
				inverse.push(...this.restoreConditionalFormatOperations(operation.sheet, operation.range))
			} else if (operation.op === 'freezePane') {
				inverse.push(this.inverseFreezePaneOperation(operation.sheet))
			}
		}
		for (const [sheet, ranges] of rangesBySheet) {
			for (const range of ranges) {
				const snapshot = this.session.requireWorkbook().readRangeCompact(sheet, range, {
					includeRefs: true,
				})
				const valueUpdates: Array<Extract<Operation, { op: 'setCells' }>['updates'][number]> = []
				if (!snapshot || snapshot.cells.length === 0) {
					inverse.push({ op: 'clearRange', sheet, range, what: 'all' })
					continue
				}
				for (const cell of snapshot.cells) {
					if (!cell.ref) continue
					const style = this.readCellStyle(sheet, cell.ref)
					const styleOp =
						style && !isDefaultStyle(style)
							? ({
									op: 'setStyle',
									sheet,
									range: cell.ref,
									style: styleToInput(style),
								} satisfies Operation)
							: undefined
					if (cell.formula) {
						inverse.push({ op: 'setFormula', sheet, ref: cell.ref, formula: cell.formula })
						if (styleOp) inverse.push(styleOp)
					} else if (cell.value.kind === 'empty') {
						if (styleOp) {
							inverse.push({ op: 'clearRange', sheet, range: cell.ref, what: 'values' })
							inverse.push({ op: 'clearRange', sheet, range: cell.ref, what: 'formulas' })
							inverse.push(styleOp)
						} else {
							inverse.push({ op: 'clearRange', sheet, range: cell.ref, what: 'all' })
						}
					} else {
						inverse.push({ op: 'clearRange', sheet, range: cell.ref, what: 'formulas' })
						valueUpdates.push({ ref: cell.ref, value: inputValueFromCell(cell.value) })
						if (styleOp) inverse.push(styleOp)
					}
				}
				if (valueUpdates.length > 0) inverse.push({ op: 'setCells', sheet, updates: valueUpdates })
			}
		}
		return inverse
	}

	private inverseStyleOperations(sheet: string, range: string): readonly Operation[] {
		const workbook = this.session.requireWorkbook()
		const existingCells = new Set<string>()
		const customStyleOps: Operation[] = []
		const absentCells: string[] = []
		const parsed = parseRange(range)
		const snapshot = workbook.readRangeCompact(sheet, range, { includeRefs: true })
		for (const cell of snapshot?.cells ?? []) {
			if (cell.ref) existingCells.add(cell.ref)
		}
		for (let row = parsed.start.row; row <= parsed.end.row; row++) {
			for (let col = parsed.start.col; col <= parsed.end.col; col++) {
				const ref = toA1({ row, col })
				if (!existingCells.has(ref)) {
					absentCells.push(ref)
					continue
				}
				const style = readCellStyle(workbook, sheet, ref)
				if (style && !isDefaultStyle(style)) {
					customStyleOps.push({
						op: 'setStyle',
						sheet,
						range: ref,
						style: styleToInput(style),
					})
				}
			}
		}
		return [
			{ op: 'clearRange', sheet, range, what: 'styles' },
			...absentCells.map((ref) => ({ op: 'clearRange', sheet, range: ref, what: 'all' }) as const),
			...customStyleOps,
		]
	}

	private inverseCommentOperation(
		sheet: string,
		ref: string,
	): Extract<Operation, { op: 'setComment' | 'deleteComment' }> {
		const comment = this.readComment(sheet, ref)
		if (!comment) return { op: 'deleteComment', sheet, ref }
		return {
			op: 'setComment',
			sheet,
			ref,
			text: comment.text,
			...(comment.author ? { author: comment.author } : {}),
		}
	}

	private readComment(
		sheet: string,
		ref: string,
	): { readonly text: string; readonly author?: string } | undefined {
		const handle = this.session.requireWorkbook().sheet(sheet)
		return handle?.comment(ref) ?? handle?.comment(ref.toUpperCase())
	}

	private readCellStyle(sheet: string, ref: string): CellStyle | undefined {
		return readCellStyle(this.session.requireWorkbook(), sheet, ref)
	}

	private inverseAutoFilterOperation(
		sheet: string,
	): Extract<Operation, { op: 'setAutoFilter' | 'clearAutoFilter' }> {
		const previous = this.session.requireWorkbook().inspectSheet(sheet)?.autoFilter
		return previous
			? { op: 'setAutoFilter', sheet, range: previous.ref }
			: { op: 'clearAutoFilter', sheet }
	}

	private inversePrintAreaOperation(
		sheet: string,
	): Extract<Operation, { op: 'setDefinedName' | 'deleteDefinedName' }> {
		const previous = this.session.requireWorkbook().definedName('_xlnm.Print_Area', sheet)
		return previous
			? { op: 'setDefinedName', name: '_xlnm.Print_Area', ref: previous.formula, scope: sheet }
			: { op: 'deleteDefinedName', name: '_xlnm.Print_Area', scope: sheet }
	}

	private inversePageSetupOperation(sheet: string): Extract<Operation, { op: 'setPageSetup' }> {
		const setup = this.session.requireWorkbook().inspectSheet(sheet)?.pageSetup
		return { op: 'setPageSetup', sheet, setup: pageSetupInput(setup) }
	}

	private inverseFreezePaneOperation(sheet: string): Extract<Operation, { op: 'freezePane' }> {
		const handle = this.session.requireWorkbook().sheet(sheet)
		return {
			op: 'freezePane',
			sheet,
			row: handle?.frozenRows ?? 0,
			col: handle?.frozenCols ?? 0,
		}
	}

	private inverseDataValidationOperations(sheet: string, range: string): readonly Operation[] {
		const restore = this.restoreDataValidationOperations(sheet, range)
		return restore.length > 0 ? restore : [{ op: 'deleteDataValidation', sheet, range }]
	}

	private restoreDataValidationOperations(sheet: string, range: string): readonly Operation[] {
		const previous = this.session
			.requireWorkbook()
			.inspectSheet(sheet)
			?.dataValidations?.find((validation) => validation.sqref === range)
		const rule = dataValidationRuleInput(previous)
		return rule ? [{ op: 'setDataValidation', sheet, range, rule }] : []
	}

	private inverseConditionalFormatOperations(sheet: string, range: string): readonly Operation[] {
		const restore = this.restoreConditionalFormatOperations(sheet, range)
		return restore.length > 0 ? restore : [{ op: 'deleteConditionalFormat', sheet, range }]
	}

	private restoreConditionalFormatOperations(
		sheet: string,
		range: string | undefined,
	): readonly Operation[] {
		if (!range) return []
		const previous = this.session
			.requireWorkbook()
			.inspectSheet(sheet)
			?.conditionalFormats?.find((format) => format.sqref === range)
		if (!previous) return []
		const rules = previous.rules.map(conditionalFormatRuleInput)
		if (rules.some((rule) => rule === null)) return []
		const operations: Operation[] = [{ op: 'deleteConditionalFormat', sheet, range }]
		for (let index = 0; index < rules.length; index++) {
			const rule = rules[index]
			if (!rule) continue
			operations.push({
				op: 'setConditionalFormat',
				sheet,
				range,
				rule,
				mode: index === 0 ? 'replace' : 'append',
			})
		}
		return operations
	}

	private move(row: number, col: number, extend = false): DispatchResult {
		this.selection = moveSelection(this.selection, row, col, extend)
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.message = selectionRef(this.selection, indexToColumn)
		return { handled: true, shouldRender: true }
	}

	private selectCurrentColumn(): DispatchResult {
		const col = this.selection.active.col
		this.selection = {
			active: { row: 0, col },
			anchor: { row: Math.max(0, this.usedRangeEnd().row), col },
			kind: 'column',
		}
		this.message = `${indexToColumn(col)}:${indexToColumn(col)}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private selectCurrentRow(): DispatchResult {
		const row = this.selection.active.row
		this.selection = {
			active: { row, col: 0 },
			anchor: { row, col: Math.max(0, this.usedRangeEnd().col) },
			kind: 'row',
		}
		this.message = `${row + 1}:${row + 1}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private selectCurrentRegionOrSheet(): DispatchResult {
		const used = this.usedRangeEnd()
		const isUsedRangeSelected =
			this.selection.kind === 'range' &&
			Math.min(this.selection.anchor.row, this.selection.active.row) === 0 &&
			Math.min(this.selection.anchor.col, this.selection.active.col) === 0 &&
			Math.max(this.selection.anchor.row, this.selection.active.row) >= used.row &&
			Math.max(this.selection.anchor.col, this.selection.active.col) >= used.col
		this.selection = isUsedRangeSelected
			? { active: { row: used.row, col: used.col }, anchor: { row: 0, col: 0 }, kind: 'sheet' }
			: { active: { row: used.row, col: used.col }, anchor: { row: 0, col: 0 }, kind: 'range' }
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.message = isUsedRangeSelected
			? `${this.sheetName()} selected`
			: selectionRef(this.selection, indexToColumn)
		return { handled: true, shouldRender: true, message: this.message }
	}

	private fillSelection(direction: 'down' | 'right'): DispatchResult {
		const bounds = selectionBounds(this.selection)
		if (direction === 'down') {
			if (bounds.endRow <= bounds.startRow) {
				this.message = 'Fill Down needs a multi-row selection.'
				return { handled: true, shouldRender: true, message: this.message }
			}
			const source = `${indexToColumn(bounds.startCol)}${bounds.startRow + 1}:${indexToColumn(bounds.endCol)}${bounds.startRow + 1}`
			const operations: Operation[] = []
			for (let row = bounds.startRow + 1; row <= bounds.endRow; row++) {
				operations.push({
					op: 'copyRange',
					sheet: this.sheetName(),
					source,
					target: `${indexToColumn(bounds.startCol)}${row + 1}`,
				})
			}
			return this.applyOperations('home.fillDown', operations, {
				message: `Filled down ${selectionRef(this.selection, indexToColumn)}`,
			})
		}
		if (bounds.endCol <= bounds.startCol) {
			this.message = 'Fill Right needs a multi-column selection.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const source = `${indexToColumn(bounds.startCol)}${bounds.startRow + 1}:${indexToColumn(bounds.startCol)}${bounds.endRow + 1}`
		const operations: Operation[] = []
		for (let col = bounds.startCol + 1; col <= bounds.endCol; col++) {
			operations.push({
				op: 'copyRange',
				sheet: this.sheetName(),
				source,
				target: `${indexToColumn(col)}${bounds.startRow + 1}`,
			})
		}
		return this.applyOperations('home.fillRight', operations, {
			message: `Filled right ${selectionRef(this.selection, indexToColumn)}`,
		})
	}

	private moveToSheetBoundary(
		direction: 'up' | 'down' | 'left' | 'right',
		extend = false,
	): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const used = this.session.requireWorkbook().inspectSheet(this.sheetName())?.usedRange
		const maxRow = Math.max(0, used?.end.row ?? this.selection.active.row)
		const maxCol = Math.max(0, used?.end.col ?? this.selection.active.col)
		const next = {
			row:
				direction === 'up'
					? 0
					: direction === 'down'
						? Math.max(this.selection.active.row, maxRow)
						: this.selection.active.row,
			col:
				direction === 'left'
					? 0
					: direction === 'right'
						? Math.max(this.selection.active.col, maxCol)
						: this.selection.active.col,
		}
		this.selection = extend
			? { active: next, anchor: this.selection.anchor, kind: 'range' }
			: selectCell(next.row, next.col)
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.message = selectionRef(this.selection, indexToColumn)
		return { handled: true, shouldRender: true, message: this.message }
	}

	private moveToUsedRangeEnd(extend = false): DispatchResult {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const used = this.session.requireWorkbook().inspectSheet(this.sheetName())?.usedRange
		const active = {
			row: Math.max(0, used?.end.row ?? this.selection.active.row),
			col: Math.max(0, used?.end.col ?? this.selection.active.col),
		}
		this.selection = extend
			? { ...this.selection, active, kind: 'range' }
			: selectCell(active.row, active.col)
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.message = selectionRef(this.selection, indexToColumn)
		return { handled: true, shouldRender: true, message: this.message }
	}

	private usedRangeEnd(): CellCoord {
		if (!this.hasWorkbook()) return this.selection.active
		const used = this.session.requireWorkbook().inspectSheet(this.sheetName())?.usedRange
		return {
			row: Math.max(0, used?.end.row ?? this.selection.active.row),
			col: Math.max(0, used?.end.col ?? this.selection.active.col),
		}
	}

	private gotoRef(ref: string): DispatchResult {
		try {
			const parsed = parseA1(ref)
			this.selection = selectCell(parsed.row, parsed.col)
			this.viewport = ensureSelectionVisible(this.viewport, this.selection)
			this.message = ref.toUpperCase()
			return { handled: true, shouldRender: true }
		} catch {
			this.message = `Invalid cell reference: ${ref}`
			return { handled: false, shouldRender: true }
		}
	}

	private async saveAsync(): Promise<DispatchResult> {
		const documentId = activeDocument(this.workspace)?.id
		try {
			const result = await this.session.save()
			if (documentId) {
				this.workspace = updateDocument(this.workspace, documentId, {
					dirty: result.current ? result.dirty : true,
					...(result.path ? { path: result.path } : {}),
				})
			} else {
				this.workspace = updateActiveDocument(this.workspace, {
					dirty: result.dirty,
					...(result.path ? { path: result.path } : {}),
				})
			}
			if (result.ok && result.path) this.recordRecentWorkbook(result.path)
			this.message = result.message
			return { handled: result.ok, shouldRender: true, message: result.message }
		} catch (error) {
			this.message = `Save failed: ${error instanceof Error ? error.message : String(error)}`
			return { handled: false, shouldRender: true, message: this.message }
		}
	}

	private async saveAsAsync(input: string): Promise<DispatchResult> {
		const request = parsePathRequest(input)
		if (request.error) {
			this.message = request.error
			return { handled: false, shouldRender: true, message: this.message }
		}
		try {
			const result = await this.session.save({ path: request.path })
			if (!result.current) {
				this.message = 'Save As finished for a stale workbook and was ignored.'
				return { handled: false, shouldRender: true, message: this.message }
			}
			this.workspace = updateActiveDocument(this.workspace, {
				dirty: result.dirty,
				path: result.path,
				name: result.path.split(/[\\/]/).pop() ?? result.path,
				info: this.session.inspect(),
			})
			if (result.ok && result.path) this.recordRecentWorkbook(result.path)
			this.message = result.message
			return { handled: result.ok, shouldRender: true, message: result.message }
		} catch (error) {
			this.message = `Save As failed: ${error instanceof Error ? error.message : String(error)}`
			return { handled: false, shouldRender: true, message: this.message }
		}
	}

	private async saveCopyAsync(input: string): Promise<DispatchResult> {
		const request = parsePathRequest(input)
		if (request.error) {
			this.message = request.error
			return { handled: false, shouldRender: true, message: this.message }
		}
		try {
			const result = await this.session.save({ path: request.path, saveCopy: true })
			this.message = result.ok
				? `Saved copy to ${request.path}. Continue editing ${activeDocument(this.workspace)?.name ?? 'current workbook'}.`
				: result.message
			return { handled: result.ok, shouldRender: true, message: this.message }
		} catch (error) {
			this.message = `Save Copy failed: ${error instanceof Error ? error.message : String(error)}`
			return { handled: false, shouldRender: true, message: this.message }
		}
	}

	private async exportAsync(input: string): Promise<DispatchResult> {
		if (!this.hasWorkbook()) {
			this.message = 'Open or create a workbook first.'
			return { handled: true, shouldRender: true, message: this.message }
		}
		const request = parseExportRequest(input)
		if (request.error) {
			this.message = request.error
			return { handled: false, shouldRender: true, message: this.message }
		}
		try {
			const workbook = this.session.requireWorkbook()
			if (request.format === 'json') {
				await writeFile(request.path, JSON.stringify(workbook.toJSON(), null, 2), 'utf-8')
			} else if (request.format === 'csv' || request.format === 'tsv') {
				const text = workbook.toCsv({
					...(request.sheet ? { sheet: request.sheet } : {}),
					...(request.format === 'tsv' ? { dialect: { delimiter: '\t' } } : {}),
				})
				await writeFile(request.path, text, 'utf-8')
			} else {
				await this.session.save({ path: request.path, saveCopy: true })
			}
			this.message = `Exported ${request.path} (${request.format})`
			return { handled: true, shouldRender: true, message: this.message }
		} catch (error) {
			this.message = `Export failed: ${error instanceof Error ? error.message : String(error)}`
			return { handled: false, shouldRender: true, message: this.message }
		}
	}

	private async openWorkbook(
		path: string,
		options: { readonly force?: boolean } = {},
	): Promise<DispatchResult> {
		const trimmed = path.trim()
		if (!trimmed) {
			this.message = 'Open requires a workbook path.'
			return { handled: false, shouldRender: true, message: this.message }
		}
		if (this.session.isDirty() && !options.force) {
			return this.unsavedChanges('open another workbook', `open! ${trimmed}`)
		}
		try {
			const document = await this.session.open(trimmed)
			this.recordRecentWorkbook(trimmed)
			this.activateDocument(
				document,
				document.protectedReview
					? `Protected review: ${document.info ? protectedReviewReasons(document.info).join(', ') : 'metadata unavailable'}`
					: `Opened ${document.name}`,
			)
			return { handled: true, shouldRender: true, message: this.message }
		} catch (error) {
			this.message = `Open failed: ${error instanceof Error ? error.message : String(error)}`
			return { handled: false, shouldRender: true, message: this.message }
		}
	}

	private activateDocument(
		document: ReturnType<WorkbookSessionController['createEmpty']>,
		message: string,
	): void {
		this.workspace = setActiveDocument(this.workspace, document)
		this.activeSheetIndex = 0
		this.selection = createSelection()
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.hydration.invalidate()
		this.journal.clear()
		this.mode = 'ready'
		this.editBuffer = ''
		this.editCursor = 0
		this.editTarget = undefined
		this.pointReference = undefined
		this.commandBuffer = ''
		this.commandPaletteIndex = 0
		this.inspectorLines = []
		this.activeDialog = undefined
		this.contextMenu = undefined
		this.keyTipActive = false
		this.workspace = { ...this.workspace, focusedRegion: 'grid' }
		this.message = message
	}

	private unsavedChanges(action: string, forceCommand: string): DispatchResult {
		this.message = `Unsaved changes. Use :wq to save or :${forceCommand} to ${action} without saving.`
		return { handled: false, shouldRender: true, message: this.message }
	}

	private selectSheetByName(sheetName: string): DispatchResult {
		const trimmed = sheetName.trim()
		if (!trimmed) {
			this.message = 'Sheet command requires a sheet name.'
			return { handled: false, shouldRender: true, message: this.message }
		}
		const index = this.sheetNames().findIndex(
			(name) => name.toLowerCase() === trimmed.toLowerCase(),
		)
		if (index < 0) {
			this.message = `Sheet not found: ${trimmed}`
			return { handled: false, shouldRender: true, message: this.message }
		}
		this.activeSheetIndex = index
		this.selection = createSelection()
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.message = `Sheet: ${this.sheetName()}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private selectSheetByOffset(delta: number): DispatchResult {
		const names = this.sheetNames()
		if (names.length === 0) return { handled: false, shouldRender: false }
		const next = (this.activeSheetIndex + delta + names.length) % names.length
		this.activeSheetIndex = next
		this.selection = createSelection()
		this.viewport = ensureSelectionVisible(this.viewport, this.selection)
		this.message = `Sheet: ${this.sheetName()}`
		return { handled: true, shouldRender: true, message: this.message }
	}

	private hasWorkbook(): boolean {
		return activeDocument(this.workspace) !== undefined
	}

	private activeDocumentReadOnly(): boolean {
		return activeDocument(this.workspace)?.readOnly === true
	}

	private sheetName(): string {
		return this.sheetNames()[this.activeSheetIndex] ?? 'Sheet1'
	}

	private sheetNames(): readonly string[] {
		return this.session.inspect()?.sheets.map((sheet) => sheet.name) ?? ['Sheet1']
	}

	private activeCellEditText(): string {
		const ref = activeCellRef(this.selection)
		const formula = this.session.formula(`${this.sheetName()}!${ref}`)
		if (formula?.formula) return `=${formula.formula}`
		const cell = this.session.cellCompact(this.sheetName(), ref)
		return cell ? displayCellValue(cell.value) : ''
	}

	private perfSummary(): string {
		const latest = this.telemetry.latest()
		if (!latest) return 'perf: awaiting samples'
		return `perf input=${formatMs(latest.inputToFrameMs)} layout=${formatMs(latest.layoutMs)} hydrate=${formatMs(latest.hydrateMs)} bytes=${latest.bytesWritten ?? 0}`
	}
}

function inputValueFromCell(value: CellValue): InputValue {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'empty':
			return null
		case 'number':
			return scalar.value
		case 'string':
			return scalar.value
		case 'boolean':
			return scalar.value
		case 'date':
			return scalar.serial
		case 'error':
			return scalar.value
		case 'richText':
			return scalar.runs.map((run) => run.text).join('')
	}
}

function objectInspectorLines(info: WorkbookInfo | null): readonly string[] {
	if (!info) return ['Object Inspector', 'No workbook metadata is available.']
	const lines = [
		'Object Inspector',
		`Charts ${info.chartCount}  Chart sheets ${info.chartSheetCount}  Pivots ${info.pivotTableCount}  Pivot caches ${info.pivotCacheCount}`,
	]
	if (info.charts.length === 0 && info.pivotTables.length === 0) {
		lines.push('No charts or pivot tables in this workbook.')
		return lines
	}
	for (const [index, chart] of info.charts.slice(0, 6).entries()) {
		const label = chart.title || chart.chartType || chart.partPath
		const sheet = chart.sheetName ? ` sheet=${chart.sheetName}` : ''
		lines.push(`Chart ${index + 1}: ${label}${sheet} series=${chart.series.length}`)
	}
	if (info.charts.length > 6) lines.push(`... ${info.charts.length - 6} more chart(s)`)
	for (const [index, pivot] of info.pivotTables.slice(0, 6).entries()) {
		const label = pivot.name || pivot.partPath
		const cache = pivot.cacheId !== undefined ? ` cache=${pivot.cacheId}` : ''
		const location = pivot.locationRef ? ` at=${pivot.locationRef}` : ''
		lines.push(`Pivot ${index + 1}: ${label} sheet=${pivot.sheetName}${location}${cache}`)
	}
	if (info.pivotTables.length > 6) lines.push(`... ${info.pivotTables.length - 6} more pivot(s)`)
	return lines
}

function helpInspectorLines(): readonly string[] {
	return [
		'Ascend TUI Help',
		'Excel landmarks: File backstage, ribbon tabs, name box, formula bar, grid, sheet tabs, status bar.',
		'Move: Arrows  Ctrl+Arrows or Mac Cmd+Arrows to data edge  Ctrl+PageUp/PageDown or Option+Arrows sheets.',
		'Edit: type to enter new content  F2 edits existing cell  = starts formula  F4 cycles references.',
		'Select: Shift+Arrows extends range. Ctrl+Space column, Shift+Space row, Ctrl/Cmd+A region then sheet.',
		'ASCII markers: [active] {selected} F filter A1/D1 sort c comment d validation ! invalid RO protected.',
		'File: Ctrl/Cmd+O open  Ctrl/Cmd+S save  F12 Save As  :save-copy <path> keeps editing original.',
		'Commands: Ctrl+P or : opens command search. Up/Down selects. / Find. F10 KeyTips. Esc returns to grid.',
	]
}

function keyTipInspectorLines(): readonly string[] {
	const groups: readonly CommandDescriptor['group'][] = [
		'file',
		'home',
		'insert',
		'formulas',
		'data',
		'review',
		'view',
	]
	const lines = [
		'KeyTips',
		'Press F File, H Home, I Insert, M Formulas, D Data, R Review, W View. Esc closes.',
		'Use Ctrl+P or : to run any command. Excel keys work when your terminal sends them.',
	]
	for (const group of groups) {
		const commands = commandsForGroup(group)
			.map((command) => {
				const key = command.excelKeys[0] ?? command.fallbackKeys[0] ?? command.id
				return `${command.title} (${key})`
			})
			.join('   ')
		lines.push(`${group.toUpperCase()}: ${commands}`)
	}
	return lines
}

function statusInspectorLines(input: {
	readonly info: WorkbookInfo | null
	readonly document: ReturnType<typeof activeDocument>
	readonly sheetName: string
	readonly selection: string
	readonly dirty: boolean
	readonly perf: string
}): readonly string[] {
	const info = input.info
	const sheet = info?.sheets.find((entry) => entry.name === input.sheetName)
	const warnings = workbookHealthWarnings(info, input.document)
	return [
		'Workbook Health',
		`State ${input.dirty ? 'Unsaved' : 'Saved'}  Path ${input.document?.path ?? '(unsaved workbook)'}`,
		`Selection ${input.sheetName}!${input.selection}`,
		'Calculation Auto Calc Done',
		`Workbook sheets=${info?.sheetCount ?? 0} cells=${formatNullableCount(info?.cellCount)} source=${info?.sourceFormat ?? 'unknown'}`,
		`Active sheet rows=${formatNullableCount(sheet?.rowCount)} cols=${formatNullableCount(sheet?.colCount)} cells=${formatNullableCount(sheet?.cellCount)}`,
		`Objects tables=${formatNullableCount(sheet?.tableCount)} charts=${info?.chartCount ?? 0} pivots=${info?.pivotTableCount ?? 0} slicers=${info?.slicerCount ?? 0}`,
		`Review comments=${formatNullableCount(info?.commentCount)} threaded=${formatNullableCount(info?.threadedCommentCount)} validation=${formatNullableCount(info?.dataValidationCount)} formats=${formatNullableCount(info?.conditionalFormatCount)}`,
		`Trust externalLinks=${info?.externalReferenceCount ?? 0} activeContent=${info?.activeContentCount ?? 0} protected=${info?.hasWorkbookProtection ? 'yes' : 'no'}`,
		`Telemetry ${input.perf}`,
		...(warnings.length > 0
			? ['Warnings', ...warnings.map((warning) => `  ${warning}`)]
			: ['Warnings none']),
	]
}

function workbookHealthWarnings(
	info: WorkbookInfo | null,
	document: ReturnType<typeof activeDocument>,
): readonly string[] {
	if (!info) return ['Workbook metadata is not hydrated yet.']
	const warnings: string[] = []
	if (document?.protectedReview) warnings.push('Review required before trusting workbook content.')
	if (document?.readOnly) warnings.push('Workbook is read-only.')
	if (info.externalReferenceCount > 0)
		warnings.push('External links are present; inspect before refresh.')
	if (info.activeContentCount > 0)
		warnings.push('Active content is preserved and blocked in the TUI.')
	if (info.pivotRefreshPlans.length > 0) warnings.push('Pivot tables may need refresh.')
	if (info.hasWorkbookProtection) warnings.push('Workbook structure protection is present.')
	return warnings
}

function openedWorkbookMessage(
	document: ReturnType<WorkbookSessionController['createEmpty']>,
	options?: WorkbookLoadOptions,
): string {
	if (document.readOnly) {
		const scope =
			options?.maxRows !== undefined
				? `first ${options.maxRows.toLocaleString()} rows`
				: 'partial workbook view'
		return `Opened ${document.name} (${scope}, read-only)`
	}
	if (document.protectedReview) {
		return `Protected review: ${document.info ? protectedReviewReasons(document.info).join(', ') : 'metadata unavailable'}`
	}
	return `Opened ${document.name}`
}

function formatNullableCount(value: number | null | undefined): string {
	return value === null || value === undefined ? 'unknown' : String(value)
}

function formulaTraceInspectorLines(trace: WorkbookTraceResult): readonly string[] {
	const lines = [
		`Trace ${trace.ref}`,
		`Value ${displayCellValue(trace.value)}`,
		trace.formula ? `Formula =${trace.formula}` : 'Formula (none)',
		`Precedents ${trace.precedents.length}`,
	]
	if (trace.precedents.length === 0) {
		lines.push('  (none)')
	} else {
		for (const node of trace.precedents.slice(0, 8)) lines.push(`  ${traceNodeLine(node)}`)
		if (trace.precedents.length > 8) lines.push(`  ... ${trace.precedents.length - 8} more`)
	}
	lines.push(`Dependents ${trace.dependents.length}`)
	if (trace.dependents.length === 0) {
		lines.push('  (none)')
	} else {
		for (const node of trace.dependents.slice(0, 8)) lines.push(`  ${traceNodeLine(node)}`)
		if (trace.dependents.length > 8) lines.push(`  ... ${trace.dependents.length - 8} more`)
	}
	return lines
}

function traceNodeLine(node: WorkbookTraceResult['precedents'][number]): string {
	const formula = node.formula ? ` formula=${node.formula}` : ''
	return `[${node.depth}] ${node.ref} value=${displayCellValue(node.value)}${formula}`
}

interface FindReplaceMatch {
	readonly ref: string
	readonly row: number
	readonly col: number
	readonly kind: 'formula' | 'value'
	readonly replacement: string
}

function normalizeFindReplaceInput(input: unknown, fallbackRange: string): FindReplaceInput {
	const record = isRecord(input) ? input : {}
	return {
		range: typeof record.range === 'string' && record.range ? record.range : fallbackRange,
		findText: typeof record.findText === 'string' ? record.findText : '',
		replaceText: typeof record.replaceText === 'string' ? record.replaceText : '',
		action: isFindReplaceAction(record.action) ? record.action : 'find',
		lookIn: isFindReplaceLookIn(record.lookIn) ? record.lookIn : 'values',
		matchCase: record.matchCase === true,
		matchEntireCell: record.matchEntireCell === true,
	}
}

function normalizeFormulaTraceInput(input: unknown): {
	readonly maxDepth?: number
	readonly error?: string
} {
	if (input === undefined) return {}
	if (!isRecord(input)) return { error: 'Trace options must be a JSON object.' }
	const value = input.maxDepth
	if (value === undefined) return {}
	if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
		return { error: 'Trace maxDepth must be a non-negative integer.' }
	}
	return { maxDepth: value }
}

function parsePathRequest(input: string): { readonly path: string; readonly error?: string } {
	const trimmed = input.trim()
	if (!trimmed) return { path: '', error: 'Path is required.' }
	if (!trimmed.startsWith('{')) return { path: trimmed }
	try {
		const payload = JSON.parse(trimmed)
		if (!isRecord(payload) || typeof payload.path !== 'string' || payload.path.trim() === '') {
			return { path: '', error: 'Path request must include a non-empty path.' }
		}
		return { path: payload.path }
	} catch (error) {
		return {
			path: '',
			error: `Invalid command JSON: ${error instanceof Error ? error.message : String(error)}`,
		}
	}
}

function parseExportRequest(input: string): {
	readonly path: string
	readonly format: 'csv' | 'tsv' | 'json' | 'xlsx' | 'xlsm'
	readonly sheet?: string
	readonly error?: string
} {
	const pathRequest = parsePathRequest(input)
	if (pathRequest.error) {
		return { path: '', format: 'xlsx', error: pathRequest.error }
	}
	let format = inferExportFormat(pathRequest.path)
	let sheet: string | undefined
	if (input.trim().startsWith('{')) {
		const payload = JSON.parse(input.trim())
		if (isRecord(payload)) {
			if (typeof payload.format === 'string') format = normalizeExportFormat(payload.format)
			if (typeof payload.sheet === 'string' && payload.sheet.trim()) sheet = payload.sheet
		}
	}
	if (!format) {
		return {
			path: pathRequest.path,
			format: 'xlsx',
			error: 'Invalid export format. Use one of: csv, tsv, json, xlsx, xlsm.',
		}
	}
	return {
		path: pathRequest.path,
		format,
		...(sheet ? { sheet } : {}),
	}
}

function buildFindReplaceOperations(
	sheet: string,
	matches: readonly FindReplaceMatch[],
): readonly Operation[] {
	const operations: Operation[] = []
	const valueUpdates: Array<Extract<Operation, { op: 'setCells' }>['updates'][number]> = []
	for (const match of matches) {
		if (match.kind === 'formula') {
			operations.push({ op: 'setFormula', sheet, ref: match.ref, formula: match.replacement })
		} else {
			operations.push({ op: 'clearRange', sheet, range: match.ref, what: 'formulas' })
			valueUpdates.push({ ref: match.ref, value: parseInputValue(match.replacement) })
		}
	}
	if (valueUpdates.length > 0) operations.push({ op: 'setCells', sheet, updates: valueUpdates })
	return operations
}

function pageSetupInput(setup: unknown): Extract<Operation, { op: 'setPageSetup' }>['setup'] {
	if (!isRecord(setup)) return {}
	const orientation = setup.orientation
	return {
		...(orientation === 'portrait' || orientation === 'landscape' ? { orientation } : {}),
		...(typeof setup.paperSize === 'number' ? { paperSize: setup.paperSize } : {}),
		...(typeof setup.scale === 'number' ? { scale: setup.scale } : {}),
		...(typeof setup.fitToWidth === 'number' ? { fitToWidth: setup.fitToWidth } : {}),
		...(typeof setup.fitToHeight === 'number' ? { fitToHeight: setup.fitToHeight } : {}),
	}
}

function dataValidationRuleInput(
	validation: unknown,
): Extract<Operation, { op: 'setDataValidation' }>['rule'] | null {
	if (!isRecord(validation) || !isDataValidationType(validation.type)) return null
	return {
		type: validation.type,
		...(typeof validation.formula1 === 'string' ? { formula1: validation.formula1 } : {}),
		...(typeof validation.formula2 === 'string' ? { formula2: validation.formula2 } : {}),
		...(isDataValidationOperator(validation.operator) ? { operator: validation.operator } : {}),
		...(typeof validation.allowBlank === 'boolean' ? { allowBlank: validation.allowBlank } : {}),
		...(typeof validation.showErrorMessage === 'boolean'
			? { showErrorMessage: validation.showErrorMessage }
			: {}),
		...(typeof validation.showDropDown === 'boolean'
			? { showDropDown: validation.showDropDown }
			: {}),
		...(typeof validation.errorTitle === 'string' ? { errorTitle: validation.errorTitle } : {}),
		...(typeof validation.error === 'string' ? { errorMessage: validation.error } : {}),
		...(typeof validation.errorStyle === 'string' ? { errorStyle: validation.errorStyle } : {}),
		...(typeof validation.imeMode === 'string' ? { imeMode: validation.imeMode } : {}),
		...(typeof validation.showInputMessage === 'boolean'
			? { showInputMessage: validation.showInputMessage }
			: {}),
		...(typeof validation.promptTitle === 'string' ? { promptTitle: validation.promptTitle } : {}),
		...(typeof validation.prompt === 'string' ? { prompt: validation.prompt } : {}),
	}
}

function conditionalFormatRuleInput(
	rule: unknown,
): Extract<Operation, { op: 'setConditionalFormat' }>['rule'] | null {
	if (!isRecord(rule) || !isConditionalFormatType(rule.type)) return null
	const formulas = Array.isArray(rule.formulas)
		? rule.formulas.filter((formula): formula is string => typeof formula === 'string')
		: []
	const colorScale = conditionalFormatColorScaleInput(rule.colorScale)
	const dataBar = conditionalFormatDataBarInput(rule.dataBar)
	const iconSet = conditionalFormatIconSetInput(rule.iconSet)
	return {
		type: rule.type,
		...(isConditionalFormatOperator(rule.operator) ? { operator: rule.operator } : {}),
		...(formulas[0] !== undefined ? { formula: formulas[0] } : {}),
		...(formulas[1] !== undefined ? { formula2: formulas[1] } : {}),
		...(typeof rule.priority === 'number' ? { priority: rule.priority } : {}),
		...(typeof rule.stopIfTrue === 'boolean' ? { stopIfTrue: rule.stopIfTrue } : {}),
		...(isRecord(rule.style) ? { style: rule.style } : {}),
		...(colorScale ? { colorScale } : {}),
		...(dataBar ? { dataBar } : {}),
		...(iconSet ? { iconSet } : {}),
	}
}

function conditionalFormatColorScaleInput(
	value: unknown,
): Extract<Operation, { op: 'setConditionalFormat' }>['rule']['colorScale'] | undefined {
	if (!isRecord(value) || !Array.isArray(value.cfvo) || !Array.isArray(value.colors)) {
		return undefined
	}
	return {
		cfvo: value.cfvo.filter(isRecord).map((entry) => ({ ...entry })),
		colors: value.colors.filter(isRecord).map((entry) => ({ ...entry })),
	}
}

function conditionalFormatDataBarInput(
	value: unknown,
): Extract<Operation, { op: 'setConditionalFormat' }>['rule']['dataBar'] | undefined {
	if (!isRecord(value) || !Array.isArray(value.cfvo)) return undefined
	const color = isRecord(value.color) ? { color: { ...value.color } } : {}
	return {
		...value,
		cfvo: value.cfvo.filter(isRecord).map((entry) => ({ ...entry })),
		...color,
	}
}

function conditionalFormatIconSetInput(
	value: unknown,
): Extract<Operation, { op: 'setConditionalFormat' }>['rule']['iconSet'] | undefined {
	if (!isRecord(value) || !Array.isArray(value.cfvo)) return undefined
	return {
		...value,
		cfvo: value.cfvo.filter(isRecord).map((entry) => ({ ...entry })),
	}
}

function isDataValidationType(
	value: unknown,
): value is Extract<Operation, { op: 'setDataValidation' }>['rule']['type'] {
	return (
		value === 'list' ||
		value === 'whole' ||
		value === 'decimal' ||
		value === 'date' ||
		value === 'time' ||
		value === 'textLength' ||
		value === 'custom'
	)
}

function isDataValidationOperator(
	value: unknown,
): value is NonNullable<Extract<Operation, { op: 'setDataValidation' }>['rule']['operator']> {
	return (
		value === 'between' ||
		value === 'notBetween' ||
		value === 'equal' ||
		value === 'notEqual' ||
		value === 'greaterThan' ||
		value === 'lessThan' ||
		value === 'greaterThanOrEqual' ||
		value === 'lessThanOrEqual'
	)
}

function isConditionalFormatType(
	value: unknown,
): value is Extract<Operation, { op: 'setConditionalFormat' }>['rule']['type'] {
	return (
		value === 'cellIs' ||
		value === 'expression' ||
		value === 'colorScale' ||
		value === 'dataBar' ||
		value === 'iconSet' ||
		value === 'top10' ||
		value === 'aboveAverage' ||
		value === 'duplicateValues' ||
		value === 'containsText'
	)
}

function isConditionalFormatOperator(
	value: unknown,
): value is NonNullable<Extract<Operation, { op: 'setConditionalFormat' }>['rule']['operator']> {
	return (
		value === 'greaterThan' ||
		value === 'lessThan' ||
		value === 'equal' ||
		value === 'between' ||
		value === 'greaterThanOrEqual' ||
		value === 'lessThanOrEqual' ||
		value === 'notEqual' ||
		value === 'notBetween'
	)
}

function replaceFormulaText(formula: string, input: FindReplaceInput): string | null {
	const rawReplacement = replaceText(formula, input)
	if (rawReplacement !== null) return rawReplacement
	const displayedReplacement = replaceText(`=${formula}`, input)
	return displayedReplacement
}

function replaceText(source: string, input: FindReplaceInput): string | null {
	const findText = input.findText
	if (!findText) return null
	if (input.matchEntireCell) {
		return equalsText(source, findText, input.matchCase) ? (input.replaceText ?? '') : null
	}
	const firstIndex = indexOfText(source, findText, input.matchCase, 0)
	if (firstIndex < 0) return null
	const replacement = input.replaceText ?? ''
	if (input.action !== 'replaceAll') {
		return `${source.slice(0, firstIndex)}${replacement}${source.slice(firstIndex + findText.length)}`
	}
	let out = ''
	let cursor = 0
	while (cursor <= source.length) {
		const index = indexOfText(source, findText, input.matchCase, cursor)
		if (index < 0) {
			out += source.slice(cursor)
			break
		}
		out += source.slice(cursor, index)
		out += replacement
		cursor = index + findText.length
	}
	return out
}

function equalsText(left: string, right: string, matchCase = false): boolean {
	return matchCase ? left === right : left.toLowerCase() === right.toLowerCase()
}

function indexOfText(source: string, needle: string, matchCase = false, offset = 0): number {
	if (matchCase) return source.indexOf(needle, offset)
	return source.toLowerCase().indexOf(needle.toLowerCase(), offset)
}

function isFindReplaceAction(value: unknown): value is FindReplaceInput['action'] {
	return value === 'find' || value === 'replace' || value === 'replaceAll'
}

function isFindReplaceLookIn(value: unknown): value is FindReplaceInput['lookIn'] {
	return value === 'values' || value === 'formulas' || value === 'both'
}

function previousDialogField(dialog: DialogViewState): number {
	if (dialog.fields.length === 0) return 0
	return (dialog.activeField - 1 + dialog.fields.length) % dialog.fields.length
}

function nextDialogField(dialog: DialogViewState): number {
	if (dialog.fields.length === 0) return 0
	return (dialog.activeField + 1) % dialog.fields.length
}

function initialDialogFieldValue(id: DialogId, name: string, input: unknown): string {
	const record = isRecord(input) ? input : {}
	switch (id) {
		case 'format-cells':
			return primitiveFieldValue(record[name])
		case 'paste-special':
			return primitiveFieldValue(record[name])
		case 'sort':
			return primitiveFieldValue(record[name])
		case 'filter':
			return primitiveFieldValue(record[name])
		case 'data-validation':
			if (name === 'formula1' && isRecord(record.rule)) {
				return primitiveFieldValue(record.rule.formula1)
			}
			return primitiveFieldValue(record[name])
		case 'conditional-formatting':
			if (name === 'formula' && isRecord(record.rule)) {
				return primitiveFieldValue(record.rule.formula)
			}
			return primitiveFieldValue(record[name])
		case 'create-table':
			return primitiveFieldValue(record[name])
		case 'comment':
			return primitiveFieldValue(record[name])
		case 'find-replace':
			return primitiveFieldValue(record[name])
		case 'chart-wizard':
			return primitiveFieldValue(record[name])
		case 'pivot-fields':
			return primitiveFieldValue(record[name])
		case 'print-preview':
			return primitiveFieldValue(record[name])
		default:
			return ''
	}
}

function dialogInputFromValues(dialog: DialogViewState): unknown {
	const values = Object.fromEntries(dialog.fields.map((field) => [field.name, field.value]))
	switch (dialog.id) {
		case 'format-cells':
			return stripEmpty({
				numberFormat: values.numberFormat,
				...(values.bold ? { bold: parseDialogBoolean(values.bold) } : {}),
				...(values.italic ? { italic: parseDialogBoolean(values.italic) } : {}),
				horizontal: values.horizontal,
			})
		case 'paste-special':
			return stripEmpty({
				source: values.source,
				target: values.target,
				mode: values.mode || 'all',
			})
		case 'sort':
			return stripEmpty({
				range: values.range,
				column: values.column,
				...(values.descending ? { descending: parseDialogBoolean(values.descending) } : {}),
			})
		case 'create-table':
			return stripEmpty({
				ref: values.ref,
				name: values.name,
				...(values.hasHeaders ? { hasHeaders: parseDialogBoolean(values.hasHeaders) } : {}),
			})
		case 'filter':
			return stripEmpty({ range: values.range })
		case 'data-validation':
			return stripEmpty({
				range: values.range,
				rule: { type: 'list', formula1: values.formula1 },
			})
		case 'conditional-formatting':
			return stripEmpty({
				range: values.range,
				rule: { type: 'expression', formula: values.formula },
			})
		case 'comment':
			return stripEmpty({
				ref: values.ref,
				text: values.text,
				author: values.author,
			})
		case 'find-replace':
			return stripEmpty({
				range: values.range,
				findText: values.findText,
				replaceText: values.replaceText,
				action: values.action || 'find',
				lookIn: values.lookIn || 'values',
				...(values.matchCase ? { matchCase: parseDialogBoolean(values.matchCase) } : {}),
				...(values.matchEntireCell
					? { matchEntireCell: parseDialogBoolean(values.matchEntireCell) }
					: {}),
			})
		case 'chart-wizard':
			return stripEmpty({
				seriesIndex: parseOptionalDialogNumber(values.seriesIndex) ?? 0,
				sheet: values.sheet,
				partPath: values.partPath,
				chartIndex: parseOptionalDialogNumber(values.chartIndex),
				nameRef: values.nameRef,
				categoryRef: values.categoryRef,
				valueRef: values.valueRef,
			})
		case 'pivot-fields':
			return stripEmpty({
				cacheId: parseOptionalDialogNumber(values.cacheId),
				partPath: values.partPath,
				pivotTable: values.pivotTable,
				sourceSheet: values.sourceSheet,
				sourceRef: values.sourceRef,
				...(values.refreshOnLoad
					? { refreshOnLoad: parseDialogBoolean(values.refreshOnLoad) }
					: {}),
				...(values.enableRefresh
					? { enableRefresh: parseDialogBoolean(values.enableRefresh) }
					: {}),
				...(values.invalid ? { invalid: parseDialogBoolean(values.invalid) } : {}),
				...(values.saveData ? { saveData: parseDialogBoolean(values.saveData) } : {}),
			})
		case 'print-preview':
			return stripEmpty({
				range: values.range,
				orientation: values.orientation,
				scale: parseOptionalDialogNumber(values.scale),
				fitToWidth: parseOptionalDialogNumber(values.fitToWidth),
				fitToHeight: parseOptionalDialogNumber(values.fitToHeight),
			})
		default:
			return values
	}
}

function stripEmpty<T extends Record<string, unknown>>(input: T): Record<string, unknown> {
	const out: Record<string, unknown> = {}
	for (const [key, value] of Object.entries(input)) {
		if (value === '') continue
		out[key] = value
	}
	return out
}

function parseDialogBoolean(value: string): boolean {
	return value.trim().toLowerCase() === 'true' || value.trim() === '1'
}

function parseOptionalDialogNumber(value: string | undefined): number | undefined {
	if (!value) return undefined
	const parsed = Number(value)
	return Number.isFinite(parsed) ? parsed : undefined
}

function primitiveFieldValue(value: unknown): string {
	if (value === undefined || value === null) return ''
	if (typeof value === 'boolean') return value ? 'true' : 'false'
	return String(value)
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === 'object' && value !== null
}

function isDefaultStyle(style: CellStyle): boolean {
	return (
		style.font === undefined &&
		style.fill === undefined &&
		style.border === undefined &&
		style.alignment === undefined &&
		style.numberFormat === undefined &&
		style.protection === undefined
	)
}

function styleToInput(style: CellStyle): StyleInput {
	return {
		...(style.font
			? {
					font: {
						...(style.font.name !== undefined ? { name: style.font.name } : {}),
						...(style.font.size !== undefined ? { size: style.font.size } : {}),
						...(style.font.bold !== undefined ? { bold: style.font.bold } : {}),
						...(style.font.italic !== undefined ? { italic: style.font.italic } : {}),
						...(style.font.underline !== undefined ? { underline: style.font.underline } : {}),
						...(style.font.strikethrough !== undefined
							? { strikethrough: style.font.strikethrough }
							: {}),
						...(style.font.color ? { color: { ...style.font.color } } : {}),
					},
				}
			: {}),
		...(style.fill
			? {
					fill: {
						...(style.fill.pattern !== undefined ? { pattern: style.fill.pattern } : {}),
						...(style.fill.fgColor ? { fgColor: { ...style.fill.fgColor } } : {}),
						...(style.fill.bgColor ? { bgColor: { ...style.fill.bgColor } } : {}),
					},
				}
			: {}),
		...(style.border
			? {
					border: {
						...(style.border.top ? { top: borderEdgeToInput(style.border.top) } : {}),
						...(style.border.bottom ? { bottom: borderEdgeToInput(style.border.bottom) } : {}),
						...(style.border.left ? { left: borderEdgeToInput(style.border.left) } : {}),
						...(style.border.right ? { right: borderEdgeToInput(style.border.right) } : {}),
						...(style.border.diagonal
							? { diagonal: borderEdgeToInput(style.border.diagonal) }
							: {}),
						...(style.border.diagonalUp !== undefined
							? { diagonalUp: style.border.diagonalUp }
							: {}),
						...(style.border.diagonalDown !== undefined
							? { diagonalDown: style.border.diagonalDown }
							: {}),
					},
				}
			: {}),
		...(style.alignment
			? {
					alignment: {
						...(style.alignment.horizontal !== undefined
							? { horizontal: style.alignment.horizontal }
							: {}),
						...(style.alignment.vertical !== undefined
							? { vertical: style.alignment.vertical }
							: {}),
						...(style.alignment.wrapText !== undefined
							? { wrapText: style.alignment.wrapText }
							: {}),
						...(style.alignment.shrinkToFit !== undefined
							? { shrinkToFit: style.alignment.shrinkToFit }
							: {}),
						...(style.alignment.textRotation !== undefined
							? { textRotation: style.alignment.textRotation }
							: {}),
						...(style.alignment.indent !== undefined ? { indent: style.alignment.indent } : {}),
						...(style.alignment.readingOrder !== undefined
							? { readingOrder: style.alignment.readingOrder }
							: {}),
					},
				}
			: {}),
		...(style.numberFormat !== undefined ? { numberFormat: style.numberFormat } : {}),
		...(style.protection
			? {
					protection: {
						...(style.protection.locked !== undefined ? { locked: style.protection.locked } : {}),
						...(style.protection.hidden !== undefined ? { hidden: style.protection.hidden } : {}),
					},
				}
			: {}),
	}
}

function borderEdgeToInput(
	edge: NonNullable<CellStyle['border']>['top'],
): NonNullable<NonNullable<StyleInput['border']>['top']> {
	return {
		...(edge?.style !== undefined ? { style: edge.style } : {}),
		...(edge?.color ? { color: { ...edge.color } } : {}),
	}
}

function isGridPasteText(text: string): boolean {
	return text.includes('\t') || text.includes('\n') || text.includes('\r')
}

function parseGridPaste(text: string): readonly (readonly string[])[] {
	const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
	const trimmed = normalized.endsWith('\n') ? normalized.slice(0, -1) : normalized
	if (trimmed === '') return []
	return trimmed.split('\n').map((line) => line.split('\t'))
}

function defaultTerminalSize(): TerminalSize {
	return {
		rows: process.stdout.rows || 32,
		cols: process.stdout.columns || 120,
	}
}

function activeCellRef(selection: SelectionState): string {
	return cellCoordRef(selection.active)
}

function cellCoordRef(coord: CellCoord): string {
	return `${indexToColumn(coord.col)}${coord.row + 1}`
}

function selectionBounds(selection: SelectionState): {
	readonly startRow: number
	readonly endRow: number
	readonly startCol: number
	readonly endCol: number
} {
	return {
		startRow: Math.min(selection.anchor.row, selection.active.row),
		endRow: Math.max(selection.anchor.row, selection.active.row),
		startCol: Math.min(selection.anchor.col, selection.active.col),
		endCol: Math.max(selection.anchor.col, selection.active.col),
	}
}

function autoSumRange(target: CellCoord): string | null {
	if (target.row > 0) {
		const col = indexToColumn(target.col)
		return `${col}1:${col}${target.row}`
	}
	if (target.col > 0) {
		return `A${target.row + 1}:${indexToColumn(target.col - 1)}${target.row + 1}`
	}
	return null
}

function parseRegisteredCommand(query: string): {
	readonly name: string
	readonly input?: unknown
	readonly error?: string
} {
	const trimmed = query.trim()
	const jsonStart = findJsonStart(trimmed)
	if (jsonStart < 0) return { name: trimmed }
	const name = trimmed.slice(0, jsonStart).trim()
	const payload = trimmed.slice(jsonStart).trim()
	if (!name) return { name: trimmed }
	try {
		return { name, input: JSON.parse(payload) }
	} catch (error) {
		return {
			name,
			error: `Invalid command JSON: ${error instanceof Error ? error.message : String(error)}`,
		}
	}
}

function findJsonStart(input: string): number {
	const objectStart = input.indexOf('{')
	const arrayStart = input.indexOf('[')
	if (objectStart < 0) return arrayStart
	if (arrayStart < 0) return objectStart
	return Math.min(objectStart, arrayStart)
}

function formatMs(value: number | undefined): string {
	return value === undefined ? '--' : `${value.toFixed(1)}ms`
}

function focusRegionLabel(region: FocusRegion): string {
	switch (region) {
		case 'fileHub':
			return 'File hub'
		case 'formulaBar':
			return 'formula bar'
		case 'nameBox':
			return 'name box'
		case 'sheetTabs':
			return 'sheet tabs'
		case 'contextMenu':
			return 'context menu'
		case 'statusBar':
			return 'status bar'
		default:
			return region
	}
}

function contextMenuTarget(selection: SelectionState): ContextMenuState['target'] {
	switch (selection.kind) {
		case 'row':
			return 'row'
		case 'column':
			return 'column'
		case 'sheet':
			return 'sheet'
		case 'range':
			return 'range'
		default:
			return 'cell'
	}
}

function contextMenuItems(selection: SelectionState): ContextMenuState['items'] {
	const base: ContextMenuState['items'] = [
		{ id: 'copy', title: 'Copy', command: 'copy', shortcut: 'Ctrl+C' },
		{ id: 'cut', title: 'Cut', command: 'cut', shortcut: 'Ctrl+X' },
		{ id: 'paste', title: 'Paste', command: 'paste', shortcut: 'Ctrl+V' },
		{
			id: 'paste-special',
			title: 'Paste Special',
			command: 'paste values',
			shortcut: 'Ctrl+Alt+V',
		},
		{ id: 'format-cells', title: 'Format Cells', command: 'format', shortcut: 'Ctrl+1' },
		{ id: 'clear', title: 'Clear Contents', command: 'clear', shortcut: 'Delete' },
	]
	if (selection.kind === 'row' || selection.kind === 'column' || selection.kind === 'sheet') {
		return [
			...base,
			{ id: 'sort', title: 'Sort', command: 'sort' },
			{ id: 'filter', title: 'Filter', command: 'filter', shortcut: 'Ctrl+Shift+L' },
			{ id: 'freeze', title: 'Freeze Panes', command: 'freeze' },
		]
	}
	return [
		...base,
		{ id: 'insert-table', title: 'Create Table', command: 'table create', shortcut: 'Ctrl+T' },
		{ id: 'sort', title: 'Sort', command: 'sort' },
		{ id: 'filter', title: 'Filter', command: 'filter', shortcut: 'Ctrl+Shift+L' },
		{ id: 'comment', title: 'New Comment', command: 'comment', shortcut: 'Shift+F2' },
		{ id: 'trace', title: 'Trace Precedents', command: 'trace precedents' },
	]
}

function commandPaletteExecution(input: string, selectedIndex = 0): string {
	const query = input.trim()
	const selected = commandPaletteResults(query)[selectedIndex]
	if (selected && (query === '' || !findCommand(query))) {
		return selected.fallbackKeys[0]?.slice(1) ?? selected.id
	}
	if (!query || /\s/.test(query) || findCommand(query)) return input
	const lower = query.toLowerCase()
	const match =
		listCommands().find((command) => command.title.toLowerCase().startsWith(lower)) ??
		listCommands().find((command) => command.id.toLowerCase().includes(lower)) ??
		listCommands().find((command) => command.title.toLowerCase().includes(lower))
	return match?.fallbackKeys[0]?.slice(1) ?? input
}

function readCellStyle(
	workbook: AscendWorkbook,
	sheetName: string,
	ref: string,
): CellStyle | undefined {
	const parsed = parseRange(ref)
	const model = workbook.getWorkbookModel()
	const cell = model.getSheet(sheetName)?.cells.get(parsed.start.row, parsed.start.col)
	if (!cell) return undefined
	return model.styles.get(cell.styleId)
}

function copyTargetRange(source: string, target: string): string {
	const sourceRange = parseRange(source)
	const targetRange = parseRange(target)
	const sourceRows = sourceRange.end.row - sourceRange.start.row
	const sourceCols = sourceRange.end.col - sourceRange.start.col
	const targetRows = targetRange.end.row - targetRange.start.row
	const targetCols = targetRange.end.col - targetRange.start.col
	const end = {
		row: targetRange.end.row + (targetRows === 0 ? sourceRows : 0),
		col: targetRange.end.col + (targetCols === 0 ? sourceCols : 0),
	}
	return `${toA1(targetRange.start)}:${toA1(end)}`
}

function fileHubEntryCount(fileHub: WorkbookWorkspace['fileHub']): number {
	if (fileHub.section === 'recent') return fileHub.entries.length > 0 ? fileHub.entries.length : 3
	return 3
}

function createRecentWorkbookStore(
	options: WorkbookTuiEngineOptions,
): RecentWorkbookStore | undefined {
	if (!options.persistState && !options.recentStorePath) return undefined
	try {
		return new RecentWorkbookStore(options.recentStorePath)
	} catch {
		return undefined
	}
}
