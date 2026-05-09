import type { CellValue } from '@ascend/schema'
import type { CompactRangeInfo } from '@ascend/sdk'
import { summarizeValues } from '../model/status-summary.ts'
import { fitAnsi } from '../render/ansi-text.ts'
import { createFrame } from '../render/frame.ts'
import type {
	CommandPaletteState,
	ContextMenuState,
	DialogViewState,
	GridSemanticModel,
	RenderFrame,
	SelectionState,
	TerminalSize,
	TuiMode,
	ViewportState,
	WorkbookWorkspace,
} from '../runtime/types.ts'
import { renderCommandPalette } from './command-palette.ts'
import { renderContextMenu } from './context-menu.ts'
import { renderDialog } from './dialog.ts'
import { renderFileHub } from './file-hub.ts'
import { formulaBarScrollOffset, renderFormulaBar } from './formula-bar.ts'
import { renderGrid } from './grid.ts'
import { renderInspector } from './inspector.ts'
import { renderRibbon } from './ribbon.ts'
import { renderSheetTabs } from './sheet-tabs.ts'
import { renderStatusBar } from './status-bar.ts'
import { renderWorkbookTabs } from './workbook-tabs.ts'

export interface WorkspaceRenderInput {
	readonly size: TerminalSize
	readonly workspace: WorkbookWorkspace
	readonly sheetNames: readonly string[]
	readonly activeSheetIndex: number
	readonly sheetName: string
	readonly mode: TuiMode
	readonly selection: SelectionState
	readonly viewport: ViewportState
	readonly data: CompactRangeInfo | undefined
	readonly editBuffer: string
	readonly formulaBarContent: string
	readonly editCursor: number
	readonly commandPalette: CommandPaletteState
	readonly activeDialog: DialogViewState | undefined
	readonly contextMenu: ContextMenuState | undefined
	readonly inspectorLines: readonly string[]
	readonly showFormulas: boolean
	readonly gridSemantics?: GridSemanticModel
	readonly message: string
	readonly dirty: boolean
	readonly perfSummary: string
}

export function renderWorkspace(input: WorkspaceRenderInput): RenderFrame {
	const lines: string[] = []
	lines.push(
		renderWorkbookTabs(
			input.workspace.documents,
			input.workspace.activeWorkbookId,
			input.size.cols,
		),
	)
	lines.push(renderRibbon(input.size.cols))
	lines.push(
		renderFormulaBar({
			width: input.size.cols,
			address: selectionAddress(input.selection),
			mode: input.mode,
			content: input.formulaBarContent,
			cursor: input.editCursor,
		}),
	)

	const bodyHeight = Math.max(1, input.size.rows - lines.length - 2)
	if (input.workspace.fileHub.visible) {
		lines.push(...renderFileHub(input.workspace.fileHub, input.size.cols, bodyHeight))
	} else {
		lines.push(
			...renderGrid({
				viewport: input.viewport,
				selection: input.selection,
				data: input.data,
				width: input.size.cols,
				height: bodyHeight,
				editBuffer: input.editBuffer,
				editing: input.mode === 'editing' || input.mode === 'entering' || input.mode === 'point',
				showFormulas: input.showFormulas,
				...(input.gridSemantics ? { semantics: input.gridSemantics } : {}),
			}),
		)
	}

	lines.push(renderSheetTabs(input.sheetNames, input.activeSheetIndex, input.size.cols))
	lines.push(
		renderStatusBar({
			width: input.size.cols,
			mode: input.mode,
			sheet: input.sheetName,
			selection: input.selection,
			message: input.message || input.perfSummary,
			...optionalStatusSummary(input),
			dirty: input.dirty,
		}),
	)
	if (input.mode === 'command') {
		overlayBottom(
			lines,
			input.size.rows,
			renderCommandPalette(input.commandPalette, input.size.cols, 6),
		)
	}
	if (input.activeDialog) {
		overlayBottom(
			lines,
			input.size.rows - 1,
			renderDialog(input.activeDialog, input.size.cols, Math.max(8, bodyHeight)),
		)
	}
	if (input.contextMenu) {
		overlayBottom(
			lines,
			input.size.rows - 1,
			renderContextMenu(input.contextMenu, input.size.cols, Math.min(14, bodyHeight)),
		)
	}
	if (!input.workspace.fileHub.visible && input.inspectorLines.length > 0) {
		overlayBottom(
			lines,
			input.size.rows - 1,
			renderInspector(input.size.cols, input.inspectorLines),
		)
	}
	if (!input.workspace.fileHub.visible && input.message.includes('Protected review')) {
		overlayBottom(lines, input.size.rows - 1, renderInspector(input.size.cols))
	}
	return createFrame(input.size, normalizeLines(lines, input.size), cursorFor(input))
}

function optionalStatusSummary(input: WorkspaceRenderInput): { readonly summary?: string } {
	const summary = statusSummary(input)
	return summary ? { summary } : {}
}

function statusSummary(input: WorkspaceRenderInput): string | undefined {
	const values = selectedValues(input)
	const summary = summarizeValues(values)
	if (summary.count === 0) return undefined
	const parts = [`Count ${summary.count}`]
	if (summary.numericCount > 0) {
		if (summary.average !== null) parts.push(`Average ${formatStatusNumber(summary.average)}`)
		parts.push(`Numerical Count ${summary.numericCount}`)
		parts.push(`Sum ${formatStatusNumber(summary.sum)}`)
		if (summary.min !== null) parts.push(`Min ${formatStatusNumber(summary.min)}`)
		if (summary.max !== null) parts.push(`Max ${formatStatusNumber(summary.max)}`)
	}
	return parts.join('  ')
}

function selectedValues(input: WorkspaceRenderInput): readonly CellValue[] {
	const startRow = Math.min(input.selection.anchor.row, input.selection.active.row)
	const endRow = Math.max(input.selection.anchor.row, input.selection.active.row)
	const startCol = Math.min(input.selection.anchor.col, input.selection.active.col)
	const endCol = Math.max(input.selection.anchor.col, input.selection.active.col)
	const values: CellValue[] = []
	for (const cell of input.data?.cells ?? []) {
		if (cell.row >= startRow && cell.row <= endRow && cell.col >= startCol && cell.col <= endCol) {
			values.push(cell.value)
		}
	}
	return values
}

function formatStatusNumber(value: number): string {
	return Number.isInteger(value)
		? String(value)
		: value.toFixed(4).replace(/0+$/, '').replace(/\.$/, '')
}

function cellAddress(selection: SelectionState): string {
	return coordAddress(selection.active)
}

function selectionAddress(selection: SelectionState): string {
	const active = coordAddress(selection.active)
	const anchor = coordAddress(selection.anchor)
	return active === anchor ? active : `${anchor}:${active}`
}

function coordAddress(coord: { readonly row: number; readonly col: number }): string {
	let col = coord.col
	let label = ''
	do {
		label = String.fromCharCode(65 + (col % 26)) + label
		col = Math.floor(col / 26) - 1
	} while (col >= 0)
	return `${label}${coord.row + 1}`
}

function cursorFor(input: WorkspaceRenderInput): RenderFrame['cursor'] {
	if (
		input.workspace.fileHub.visible ||
		input.activeDialog ||
		input.inspectorLines.length > 0 ||
		input.workspace.focusedRegion !== 'grid'
	) {
		return { row: 1, col: 1, visible: false }
	}
	if (input.mode === 'editing' || input.mode === 'entering' || input.mode === 'point') {
		const formulaPrefixWidth = formulaBarPrefixWidth(input.selection)
		const scrollOffset = formulaBarScrollOffset({
			width: input.size.cols,
			address: selectionAddress(input.selection),
			mode: input.mode,
			content: input.formulaBarContent,
			cursor: input.editCursor,
		})
		return {
			row: 3,
			col: Math.min(input.size.cols, formulaPrefixWidth + input.editCursor - scrollOffset + 1),
			visible: true,
		}
	}
	return {
		row: Math.min(input.size.rows, 5 + input.selection.active.row - input.viewport.topRow),
		col: gridCellStartColumn(input.viewport, input.selection.active.col - input.viewport.leftCol),
		visible: input.mode !== 'command',
	}
}

function formulaBarPrefixWidth(selection: SelectionState): number {
	return ` ${cellAddress(selection).padEnd(10)} | fx* `.length
}

function gridCellStartColumn(viewport: ViewportState, visibleCol: number): number {
	let col = 8
	for (let index = 0; index < visibleCol; index++) {
		col += (viewport.columnWidths[index] ?? 10) + 1
	}
	return col
}

function overlayBottom(lines: string[], maxRows: number, overlay: readonly string[]): void {
	const start = Math.max(0, Math.min(lines.length, maxRows - overlay.length))
	for (let i = 0; i < overlay.length; i++) {
		lines[start + i] = overlay[i] ?? ''
	}
}

function normalizeLines(lines: readonly string[], size: TerminalSize): readonly string[] {
	const next = lines.slice(0, size.rows)
	while (next.length < size.rows) next.push('')
	return next.map((line) => fitAnsi(line, size.cols))
}
