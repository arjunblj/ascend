import type { CellValue } from '@ascend/schema'
import type { StyleId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface Cell {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
}

const DEFAULT_STYLE_ID = 0 as StyleId

export class SparseGrid {
	private readonly data = new Map<number, StoredCell>()
	private _minRow = Number.POSITIVE_INFINITY
	private _maxRow = Number.NEGATIVE_INFINITY
	private _minCol = Number.POSITIVE_INFINITY
	private _maxCol = Number.NEGATIVE_INFINITY
	private _boundsDirty = false

	get(row: number, col: number): Cell | undefined {
		return materializeCell(this.data.get(packKey(row, col)))
	}

	set(row: number, col: number, cell: Cell): void {
		this.data.set(packKey(row, col), compactCell(cell))
		if (row < this._minRow) this._minRow = row
		if (row > this._maxRow) this._maxRow = row
		if (col < this._minCol) this._minCol = col
		if (col > this._maxCol) this._maxCol = col
	}

	delete(row: number, col: number): boolean {
		const deleted = this.data.delete(packKey(row, col))
		if (deleted) this._boundsDirty = true
		return deleted
	}

	getValue(row: number, col: number): CellValue | undefined {
		return readStoredValue(this.data.get(packKey(row, col)))
	}

	*getRange(range: RangeRef): Generator<readonly [number, number, Cell]> {
		for (const [key, stored] of this.data) {
			const [row, col] = unpackKey(key)
			if (
				row < range.start.row ||
				row > range.end.row ||
				col < range.start.col ||
				col > range.end.col
			) {
				continue
			}
			const cell = materializeCell(stored)
			if (cell) yield [row, col, cell] as const
		}
	}

	foldRange<T>(
		startRow: number,
		startCol: number,
		endRow: number,
		endCol: number,
		init: T,
		fn: (acc: T, value: CellValue, row: number, col: number) => T,
	): T {
		let acc = init
		for (const [key, stored] of this.data) {
			const [row, col] = unpackKey(key)
			if (row < startRow || row > endRow || col < startCol || col > endCol) continue
			const value = readStoredValue(stored)
			if (value) acc = fn(acc, value, row, col)
		}
		return acc
	}

	usedRange(): RangeRef | null {
		if (this.data.size === 0) return null
		if (this._boundsDirty) this._recomputeBounds()
		return {
			start: { row: this._minRow, col: this._minCol },
			end: { row: this._maxRow, col: this._maxCol },
		}
	}

	*iterate(): Generator<readonly [number, number, Cell]> {
		for (const [key, cell] of this.data) {
			const [row, col] = unpackKey(key)
			const materialized = materializeCell(cell)
			if (materialized) yield [row, col, materialized] as const
		}
	}

	*iterateRows(): Generator<readonly [number, readonly (readonly [number, Cell])[]]> {
		const rows = new Map<number, Array<readonly [number, Cell]>>()
		for (const [key, stored] of this.data) {
			const [row, col] = unpackKey(key)
			const cell = materializeCell(stored)
			if (!cell) continue
			const rowCells = rows.get(row)
			if (rowCells) rowCells.push([col, cell] as const)
			else rows.set(row, [[col, cell] as const])
		}
		const sortedRows = [...rows.entries()].sort((a, b) => a[0] - b[0])
		for (const [row, rowCells] of sortedRows) {
			const cells = rowCells.sort((a, b) => a[0] - b[0])
			yield [row, cells] as const
		}
	}

	cellCount(): number {
		return this.data.size
	}

	clear(): void {
		this.data.clear()
		this._minRow = Number.POSITIVE_INFINITY
		this._maxRow = Number.NEGATIVE_INFINITY
		this._minCol = Number.POSITIVE_INFINITY
		this._maxCol = Number.NEGATIVE_INFINITY
		this._boundsDirty = false
	}

	insertRows(at: number, count: number): void {
		this._rebuildAfterShift((row, col) => ({ row: row >= at ? row + count : row, col }), count > 0)
	}

	deleteRows(at: number, count: number): void {
		const deleteEnd = at + count
		this._rebuildAfterShift((row, col) => {
			if (row >= at && row < deleteEnd) return null
			return { row: row >= deleteEnd ? row - count : row, col }
		}, count > 0)
	}

	insertCols(at: number, count: number): void {
		this._rebuildAfterShift((row, col) => ({ row, col: col >= at ? col + count : col }), count > 0)
	}

	deleteCols(at: number, count: number): void {
		const deleteEnd = at + count
		this._rebuildAfterShift((row, col) => {
			if (col >= at && col < deleteEnd) return null
			return { row, col: col >= deleteEnd ? col - count : col }
		}, count > 0)
	}

	clone(): SparseGrid {
		const clone = new SparseGrid()
		for (const [key, cell] of this.data) {
			clone.data.set(key, cloneCell(cell))
		}
		clone._minRow = this._minRow
		clone._maxRow = this._maxRow
		clone._minCol = this._minCol
		clone._maxCol = this._maxCol
		clone._boundsDirty = this._boundsDirty
		return clone
	}

	private _recomputeBounds(): void {
		this._minRow = Number.POSITIVE_INFINITY
		this._maxRow = Number.NEGATIVE_INFINITY
		this._minCol = Number.POSITIVE_INFINITY
		this._maxCol = Number.NEGATIVE_INFINITY
		for (const key of this.data.keys()) {
			const [row, col] = unpackKey(key)
			if (row < this._minRow) this._minRow = row
			if (row > this._maxRow) this._maxRow = row
			if (col < this._minCol) this._minCol = col
			if (col > this._maxCol) this._maxCol = col
		}
		this._boundsDirty = false
	}

	private _rebuildAfterShift(
		mapCell: (row: number, col: number) => { row: number; col: number } | null,
		active: boolean,
	): void {
		if (!active || this.data.size === 0) return
		const nextData = new Map<number, StoredCell>()
		for (const [key, cell] of this.data) {
			const [row, col] = unpackKey(key)
			const next = mapCell(row, col)
			if (!next) continue
			nextData.set(packKey(next.row, next.col), cell)
		}
		this.data.clear()
		for (const [key, cell] of nextData) {
			this.data.set(key, cell)
		}
		this._boundsDirty = true
	}
}

type StoredCell = CellValue | StyledCell | FormulaCell

class StyledCell {
	constructor(
		readonly value: CellValue,
		readonly styleId: StyleId,
	) {}
}

class FormulaCell implements Cell {
	constructor(
		readonly value: CellValue,
		readonly formula: string,
		readonly styleId: StyleId,
	) {}
}

const PACK_FACTOR = 1 << 20

function packKey(row: number, col: number): number {
	return row * PACK_FACTOR + col
}

function unpackKey(key: number): readonly [number, number] {
	const col = key % PACK_FACTOR
	const row = (key - col) / PACK_FACTOR
	return [row, col] as const
}

function compactCell(cell: Cell): StoredCell {
	if (cell.formula === null && cell.styleId === DEFAULT_STYLE_ID) {
		return cell.value
	}
	if (cell.formula === null) {
		return new StyledCell(cell.value, cell.styleId)
	}
	return new FormulaCell(cell.value, cell.formula, cell.styleId)
}

function cloneCell(cell: StoredCell): StoredCell {
	if (cell instanceof FormulaCell) {
		return new FormulaCell(structuredClone(cell.value), cell.formula, cell.styleId)
	}
	if (cell instanceof StyledCell) {
		return new StyledCell(structuredClone(cell.value), cell.styleId)
	}
	return structuredClone(cell)
}

function materializeCell(cell: StoredCell | undefined): Cell | undefined {
	if (!cell) return undefined
	if (cell instanceof FormulaCell) return cell
	if (cell instanceof StyledCell) {
		return { value: cell.value, formula: null, styleId: cell.styleId }
	}
	return { value: cell, formula: null, styleId: DEFAULT_STYLE_ID }
}

function readStoredValue(cell: StoredCell | undefined): CellValue | undefined {
	if (!cell) return undefined
	if (cell instanceof FormulaCell || cell instanceof StyledCell) return cell.value
	return cell
}
