import type { CellCoord, SelectionState } from '../runtime/types.ts'

export function createSelection(row = 0, col = 0): SelectionState {
	const active = clampCoord({ row, col })
	return { active, anchor: active, kind: 'cell' }
}

export function moveSelection(
	selection: SelectionState,
	deltaRow: number,
	deltaCol: number,
	extend = false,
): SelectionState {
	const active = clampCoord({
		row: selection.active.row + deltaRow,
		col: selection.active.col + deltaCol,
	})
	return {
		active,
		anchor: extend ? selection.anchor : active,
		kind: extend ? 'range' : 'cell',
	}
}

export function selectCell(row: number, col: number): SelectionState {
	return createSelection(row, col)
}

export function selectionRef(
	selection: SelectionState,
	indexToColumn: (index: number) => string,
): string {
	const startRow = Math.min(selection.anchor.row, selection.active.row)
	const endRow = Math.max(selection.anchor.row, selection.active.row)
	const startCol = Math.min(selection.anchor.col, selection.active.col)
	const endCol = Math.max(selection.anchor.col, selection.active.col)
	const start = `${indexToColumn(startCol)}${startRow + 1}`
	const end = `${indexToColumn(endCol)}${endRow + 1}`
	return start === end ? start : `${start}:${end}`
}

export function selectionDimensions(selection: SelectionState): { rows: number; cols: number } {
	return {
		rows: Math.abs(selection.active.row - selection.anchor.row) + 1,
		cols: Math.abs(selection.active.col - selection.anchor.col) + 1,
	}
}

function clampCoord(coord: CellCoord): CellCoord {
	return {
		row: Math.max(0, coord.row),
		col: Math.max(0, coord.col),
	}
}
