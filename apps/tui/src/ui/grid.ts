import { EMPTY } from '@ascend/schema'
import type { CompactCellInfo, CompactRangeInfo } from '@ascend/sdk'
import { indexToColumn } from '@ascend/sdk'
import { displayCellValue } from '../model/display-cell.ts'
import { clipAnsi, fitAnsi, sanitizeTerminalText, visibleLength } from '../render/ansi-text.ts'
import { colorForKind, THEME } from '../render/styles.ts'
import type {
	GridSemanticCell,
	GridSemanticModel,
	SelectionState,
	ViewportState,
} from '../runtime/types.ts'

export function renderGrid(input: {
	readonly viewport: ViewportState
	readonly selection: SelectionState
	readonly data: CompactRangeInfo | undefined
	readonly width: number
	readonly height: number
	readonly editBuffer: string
	readonly editing: boolean
	readonly showFormulas: boolean
	readonly semantics?: GridSemanticModel
}): readonly string[] {
	const cellMap = new Map<string, CompactCellInfo>()
	for (const cell of input.data?.cells ?? []) {
		if (cell.ref) cellMap.set(cell.ref, cell)
	}
	const rowHeaderWidth = 6
	const lines: string[] = []
	const bounds = selectionBounds(input.selection)
	const headerParts = [`${THEME.header}${' '.repeat(rowHeaderWidth)}${THEME.reset}`]
	for (let c = 0; c < input.viewport.visibleCols; c++) {
		const col = input.viewport.leftCol + c
		const label = headerLabel(input, col)
		const activeCol = input.selection.active.col === col
		const selectedCol = isColumnSelected(input.selection, bounds, col)
		const headerStyle = activeCol
			? `${THEME.active}${THEME.bold}`
			: selectedCol
				? THEME.selection
				: THEME.header
		headerParts.push(
			`${THEME.grid}|${THEME.reset}${headerStyle}${fitAnsi(label, input.viewport.columnWidths[c] ?? 10)}${THEME.reset}`,
		)
	}
	lines.push(trimLine(headerParts.join(''), input.width))
	for (let r = 0; r < Math.min(input.height - 1, input.viewport.visibleRows); r++) {
		const row = input.viewport.topRow + r
		const activeRow = input.selection.active.row === row
		const selectedRow = isRowSelected(input.selection, bounds, row)
		const rowHeaderStyle = activeRow
			? `${THEME.active}${THEME.bold}`
			: selectedRow
				? THEME.selection
				: THEME.header
		const rowParts = [`${rowHeaderStyle}${String(row + 1).padStart(rowHeaderWidth)}${THEME.reset}`]
		for (let c = 0; c < input.viewport.visibleCols; c++) {
			const col = input.viewport.leftCol + c
			const width = input.viewport.columnWidths[c] ?? 10
			const ref = `${indexToColumn(col)}${row + 1}`
			const active = input.selection.active.row === row && input.selection.active.col === col
			const selected = isSelected(input.selection, row, col)
			const cell = cellMap.get(ref)
			const semantic = input.semantics?.cells.get(ref)
			const value = cell?.value ?? EMPTY
			const raw = sanitizeTerminalText(
				active && input.editing
					? input.editBuffer
					: input.showFormulas && cell?.formula
						? normalizeFormulaText(cell.formula)
						: displayCellValue(value),
			)
			const text = formatSemanticCell({
				text: raw,
				width,
				rightAlign: value.kind === 'number' || value.kind === 'date',
				active,
				selected,
				editing: input.editing,
				semantic,
			})
			const style = active
				? `${THEME.active}${THEME.bold}`
				: selected
					? THEME.selection
					: colorForKind(value.kind)
			rowParts.push(`${THEME.grid}|${THEME.reset}${style}${text}${THEME.reset}`)
		}
		lines.push(trimLine(rowParts.join(''), input.width))
	}
	return lines
}

function headerLabel(
	input: {
		readonly viewport: ViewportState
		readonly semantics?: GridSemanticModel
	},
	col: number,
): string {
	const base = indexToColumn(col)
	const flags = new Set<string>()
	const top = input.viewport.topRow
	for (let row = top; row < top + input.viewport.visibleRows; row++) {
		const ref = `${indexToColumn(col)}${row + 1}`
		const semantic = input.semantics?.cells.get(ref)
		if (!semantic) continue
		for (const flag of semantic.flags) flags.add(flag)
	}
	const suffix = [
		flags.has('filterActive') ? 'F' : flags.has('filterAvailable') ? 'v' : '',
		flags.has('sortAsc') ? 'A1' : flags.has('sortDesc') ? 'D1' : '',
	].filter(Boolean)
	return suffix.length === 0 ? base : `${base} ${suffix.join(' ')}`
}

function isSelected(selection: SelectionState, row: number, col: number): boolean {
	const bounds = selectionBounds(selection)
	if (selection.kind === 'sheet') return true
	if (selection.kind === 'row') return row >= bounds.startRow && row <= bounds.endRow
	if (selection.kind === 'column') return col >= bounds.startCol && col <= bounds.endCol
	return (
		row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol && col <= bounds.endCol
	)
}

function isRowSelected(
	selection: SelectionState,
	bounds: ReturnType<typeof selectionBounds>,
	row: number,
): boolean {
	return (
		selection.kind === 'sheet' ||
		selection.kind === 'row' ||
		(row >= bounds.startRow && row <= bounds.endRow)
	)
}

function isColumnSelected(
	selection: SelectionState,
	bounds: ReturnType<typeof selectionBounds>,
	col: number,
): boolean {
	return (
		selection.kind === 'sheet' ||
		selection.kind === 'column' ||
		(col >= bounds.startCol && col <= bounds.endCol)
	)
}

function selectionBounds(selection: SelectionState): {
	readonly startRow: number
	readonly endRow: number
	readonly startCol: number
	readonly endCol: number
} {
	const startRow = Math.min(selection.anchor.row, selection.active.row)
	const endRow = Math.max(selection.anchor.row, selection.active.row)
	const startCol = Math.min(selection.anchor.col, selection.active.col)
	const endCol = Math.max(selection.anchor.col, selection.active.col)
	return { startRow, endRow, startCol, endCol }
}

function normalizeFormulaText(formula: string): string {
	return formula.startsWith('=') ? formula : `=${formula}`
}

function formatSemanticCell(input: {
	readonly text: string
	readonly width: number
	readonly rightAlign: boolean
	readonly active: boolean
	readonly selected: boolean
	readonly editing: boolean
	readonly semantic: GridSemanticCell | undefined
}): string {
	const markers = semanticMarkers(input.semantic)
	const wrapper: readonly [string, string] = input.active
		? input.editing
			? ['', '']
			: ['[', ']']
		: input.selected
			? ['{', '}']
			: ['', '']
	const markerWidth = visibleLength(markers)
	const wrapperWidth = visibleLength(wrapper[0] + wrapper[1])
	const innerWidth = Math.max(0, input.width - markerWidth - wrapperWidth)
	const clipped =
		visibleLength(input.text) <= innerWidth
			? input.text
			: `${clipAnsi(input.text, Math.max(0, innerWidth - 1))}~`
	const inner = input.rightAlign ? fitAnsiLeft(clipped, innerWidth) : fitAnsi(clipped, innerWidth)
	return fitAnsi(`${wrapper[0]}${inner}${wrapper[1]}${markers}`, input.width)
}

function semanticMarkers(semantic: GridSemanticCell | undefined): string {
	if (!semantic || semantic.flags.length === 0) return ''
	const flags = semantic.flags
	const markers = [
		flags.includes('formulaError') ? 'ERR' : '',
		flags.includes('filterActive') ? 'F' : flags.includes('filterAvailable') ? 'v' : '',
		flags.includes('sortAsc') ? 'A1' : flags.includes('sortDesc') ? 'D1' : '',
		flags.includes('comment') ? 'c' : '',
		flags.includes('validationInvalid') ? '!' : flags.includes('validationDropdown') ? 'd' : '',
		flags.includes('conditionalFormat') ? '*' : '',
		flags.includes('hyperlink') ? '@' : '',
		flags.includes('protected') ? 'RO' : '',
		flags.includes('tableHeader')
			? 'H'
			: flags.includes('tableTotal')
				? 'T'
				: flags.includes('table')
					? ':'
					: '',
	].filter(Boolean)
	return markers.length === 0 ? '' : ` ${markers.join('')}`
}

function fitAnsiLeft(text: string, width: number): string {
	const clipped = clipAnsi(text, width)
	const padding = Math.max(0, width - visibleLength(clipped))
	return padding === 0 ? clipped : `${' '.repeat(padding)}${clipped}`
}

function trimLine(line: string, width: number): string {
	return fitAnsi(line, width)
}
