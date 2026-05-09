import type { CellValue, InputValue, Operation } from '@ascend/schema'
import type { CompactCellInfo, WorkbookInfo } from '@ascend/sdk'

export type TuiMode = 'ready' | 'entering' | 'editing' | 'point' | 'command'

export type FocusRegion =
	| 'grid'
	| 'formulaBar'
	| 'nameBox'
	| 'sheetTabs'
	| 'ribbon'
	| 'statusBar'
	| 'fileHub'
	| 'dialog'
	| 'contextMenu'
	| 'inspector'

export interface TerminalSize {
	readonly rows: number
	readonly cols: number
}

export interface CellCoord {
	readonly row: number
	readonly col: number
}

export interface SelectionState {
	readonly active: CellCoord
	readonly anchor: CellCoord
	readonly kind: 'cell' | 'range' | 'row' | 'column' | 'sheet'
}

export interface ViewportState {
	readonly topRow: number
	readonly leftCol: number
	readonly visibleRows: number
	readonly visibleCols: number
	readonly columnWidths: readonly number[]
	readonly overscanRows: number
	readonly overscanCols: number
}

export interface FileHubState {
	readonly visible: boolean
	readonly section: 'recent' | 'open' | 'new' | 'saveAs' | 'export' | 'recover' | 'info'
	readonly query: string
	readonly selectedIndex: number
	readonly entries: readonly FileHubEntry[]
}

export interface FileHubEntry {
	readonly label: string
	readonly path?: string
	readonly detail?: string
	readonly pinned?: boolean
	readonly missing?: boolean
}

export interface OpenWorkbook {
	readonly id: string
	readonly path: string | null
	readonly name: string
	readonly info: WorkbookInfo | null
	readonly readOnly: boolean
	readonly protectedReview: boolean
	readonly dirty: boolean
}

export interface WorkbookWorkspace {
	readonly documents: readonly OpenWorkbook[]
	readonly activeWorkbookId: string | null
	readonly fileHub: FileHubState
	readonly focusedRegion: FocusRegion
}

export interface DialogViewState {
	readonly id: string
	readonly title: string
	readonly activeField: number
	readonly fields: readonly {
		readonly name: string
		readonly label: string
		readonly kind: string
		readonly required: boolean
		readonly options: readonly string[]
		readonly value: string
	}[]
}

export interface CommandPaletteState {
	readonly query: string
	readonly selectedIndex: number
}

export interface ContextMenuItem {
	readonly id: string
	readonly title: string
	readonly command: string
	readonly shortcut?: string
	readonly detail?: string
}

export interface ContextMenuState {
	readonly target: 'cell' | 'range' | 'row' | 'column' | 'sheet'
	readonly address: string
	readonly selectedIndex: number
	readonly items: readonly ContextMenuItem[]
}

export type GridSemanticFlag =
	| 'table'
	| 'tableHeader'
	| 'tableTotal'
	| 'filterAvailable'
	| 'filterActive'
	| 'sortAsc'
	| 'sortDesc'
	| 'comment'
	| 'validationDropdown'
	| 'validationInvalid'
	| 'conditionalFormat'
	| 'hyperlink'
	| 'protected'
	| 'formulaError'

export interface GridSemanticCell {
	readonly ref: string
	readonly row: number
	readonly col: number
	readonly flags: readonly GridSemanticFlag[]
	readonly errorText?: string
}

export interface GridSemanticModel {
	readonly cells: ReadonlyMap<string, GridSemanticCell>
	readonly frozenRows: number
	readonly frozenCols: number
	readonly protected: boolean
	readonly activeFilterRanges: readonly string[]
}

export type InputEvent =
	| { readonly kind: 'key'; readonly key: string; readonly text?: string; readonly raw?: string }
	| {
			readonly kind: 'mouse'
			readonly action: 'press' | 'drag' | 'release' | 'wheel'
			readonly row: number
			readonly col: number
			readonly button?: number
	  }
	| { readonly kind: 'resize'; readonly size: TerminalSize }
	| { readonly kind: 'command'; readonly command: string }

export interface DispatchResult {
	readonly handled: boolean
	readonly shouldRender: boolean
	readonly shouldExit?: boolean
	readonly message?: string
}

export interface TraceResult {
	readonly state: TuiStateSnapshot
	readonly frames: readonly RenderFrame[]
	readonly telemetry: readonly TelemetrySample[]
}

export interface TraceOptions {
	readonly size: TerminalSize
	readonly includeFrames?: boolean
}

export interface TuiEngine {
	dispatch(event: InputEvent): DispatchResult | Promise<DispatchResult>
	render(size: TerminalSize): RenderFrame
	runHeadless(trace: readonly InputEvent[], options: TraceOptions): Promise<TraceResult>
	state(): TuiStateSnapshot
}

export interface RenderCell {
	readonly text: string
	readonly style?: string
}

export interface RenderFrame {
	readonly size: TerminalSize
	readonly lines: readonly string[]
	readonly cursor?: { readonly row: number; readonly col: number; readonly visible: boolean }
	readonly stats: FrameStats
}

export interface FrameStats {
	readonly fullFrameCells: number
	readonly dirtyCells: number
	readonly dirtyRows: number
	readonly bytes: number
}

export interface RenderPatch {
	readonly lines: readonly { readonly row: number; readonly text: string }[]
	readonly fullRedraw: boolean
	readonly bytes: number
}

export interface TerminalCapabilities {
	readonly isTty: boolean
	readonly color: 'truecolor' | '256' | '16' | 'none'
	readonly unicode: boolean
	readonly mouse: boolean
	readonly bracketedPaste: boolean
	readonly hyperlinks: boolean
	readonly graphics: 'kitty' | 'iterm' | 'sixel' | 'unicode' | 'off'
	readonly keyboardProtocol: 'kitty' | 'csi-u' | 'legacy'
	readonly profile: 'modern' | 'legacy' | 'mac-terminal' | 'windows-terminal' | 'ssh-tmux-limited'
}

export interface TerminalRenderer {
	init(capabilities: TerminalCapabilities): Promise<void>
	draw(frame: RenderFrame): Promise<RenderStats>
	shutdown(): Promise<void>
}

export interface RenderStats {
	readonly frameBuildMs: number
	readonly frameDiffMs: number
	readonly encodeMs: number
	readonly writeMs: number
	readonly changedCells: number
	readonly bytesOut: number
	readonly droppedFrames: number
	readonly fps: number
}

export interface TelemetrySample {
	readonly timestamp: number
	readonly inputToFrameMs?: number
	readonly layoutMs?: number
	readonly hydrateMs?: number
	readonly formatMs?: number
	readonly diffMs?: number
	readonly encodeMs?: number
	readonly ptyWriteMs?: number
	readonly bytesWritten?: number
	readonly dirtyCells?: number
	readonly dirtyRows?: number
	readonly tileCacheHitRate?: number
	readonly hydratedCells?: number
	readonly rss?: number
	readonly heapUsed?: number
	readonly recalcGenerationLag?: number
	readonly fps?: number
	readonly droppedFrames?: number
}

export interface CommandContext {
	readonly sheet: string
	readonly selection: SelectionState
	readonly input?: unknown
}

export interface CommandDescriptor {
	readonly id: string
	readonly title: string
	readonly group: 'file' | 'home' | 'insert' | 'formulas' | 'data' | 'review' | 'view'
	readonly excelKeys: readonly string[]
	readonly fallbackKeys: readonly string[]
	readonly contexts: readonly FocusRegion[]
	readonly dialogId?: string
	toOperations(ctx: CommandContext, input: unknown): readonly Operation[]
}

export interface CommandIntent {
	readonly id: string
	readonly input?: unknown
}

export interface CommandPreview {
	readonly intent: CommandIntent
	readonly operations: readonly Operation[]
	readonly warnings: readonly string[]
}

export interface JournalEntry {
	readonly id: string
	readonly generation: number
	readonly commandId: string
	readonly selectionBefore: SelectionState
	readonly selectionAfter: SelectionState
	readonly ops: readonly Operation[]
	readonly inverseOps: readonly Operation[]
	readonly preimageRanges: readonly string[]
	readonly affectedCells: readonly string[]
	readonly recalcDirtyRefs: readonly string[]
	readonly timestamp: number
}

export interface TuiStateSnapshot {
	readonly workspace: WorkbookWorkspace
	readonly mode: TuiMode
	readonly sheetName: string
	readonly selection: SelectionState
	readonly viewport: ViewportState
	readonly editBuffer: string
	readonly editCursor: number
	readonly commandBuffer: string
	readonly commandPalette: CommandPaletteState
	readonly activeDialog?: DialogViewState
	readonly contextMenu?: ContextMenuState
	readonly inspectorLines: readonly string[]
	readonly showFormulas: boolean
	readonly dirty: boolean
	readonly message: string
	readonly telemetry: readonly TelemetrySample[]
}

export interface ViewCell {
	readonly ref: string
	readonly value: CellValue
	readonly display: string
	readonly source?: CompactCellInfo
}

export type ParsedInputValue = InputValue
