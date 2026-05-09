import type { SelectionState, TerminalSize, ViewportState } from '../runtime/types.ts'

export function createViewport(size: TerminalSize): ViewportState {
	const visibleRows = Math.max(1, size.rows - 6)
	const visibleCols = Math.max(1, Math.floor(Math.max(20, size.cols - 7) / 11))
	return {
		topRow: 0,
		leftCol: 0,
		visibleRows,
		visibleCols,
		columnWidths: Array.from({ length: visibleCols }, () => 10),
		overscanRows: visibleRows * 2,
		overscanCols: 4,
	}
}

export function resizeViewport(viewport: ViewportState, size: TerminalSize): ViewportState {
	const next = createViewport(size)
	return {
		...next,
		topRow: viewport.topRow,
		leftCol: viewport.leftCol,
	}
}

export function ensureSelectionVisible(
	viewport: ViewportState,
	selection: SelectionState,
): ViewportState {
	let topRow = viewport.topRow
	let leftCol = viewport.leftCol
	const row = selection.active.row
	const col = selection.active.col
	if (row < topRow) topRow = row
	if (row >= topRow + viewport.visibleRows) topRow = row - viewport.visibleRows + 1
	if (col < leftCol) leftCol = col
	if (col >= leftCol + viewport.visibleCols) leftCol = col - viewport.visibleCols + 1
	return topRow === viewport.topRow && leftCol === viewport.leftCol
		? viewport
		: { ...viewport, topRow, leftCol }
}

export function viewportRange(
	viewport: ViewportState,
	indexToColumn: (index: number) => string,
): string {
	const start = `${indexToColumn(viewport.leftCol)}${viewport.topRow + 1}`
	const end = `${indexToColumn(viewport.leftCol + viewport.visibleCols - 1)}${viewport.topRow + viewport.visibleRows}`
	return `${start}:${end}`
}
