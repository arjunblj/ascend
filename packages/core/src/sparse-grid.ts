import type { CellValue } from '@ascend/schema'
import type { StyleId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface Cell {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
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

export class SparseGrid {
	private readonly data = new Map<number, Cell>()
	private _minRow = Number.POSITIVE_INFINITY
	private _maxRow = Number.NEGATIVE_INFINITY
	private _minCol = Number.POSITIVE_INFINITY
	private _maxCol = Number.NEGATIVE_INFINITY
	private _boundsDirty = false

	get(row: number, col: number): Cell | undefined {
		return this.data.get(packKey(row, col))
	}

	set(row: number, col: number, cell: Cell): void {
		this.data.set(packKey(row, col), cell)
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
		return this.data.get(packKey(row, col))?.value
	}

	*getRange(range: RangeRef): Generator<readonly [number, number, Cell]> {
		for (let r = range.start.row; r <= range.end.row; r++) {
			for (let c = range.start.col; c <= range.end.col; c++) {
				const cell = this.data.get(packKey(r, c))
				if (cell) yield [r, c, cell] as const
			}
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
		for (let r = startRow; r <= endRow; r++) {
			for (let c = startCol; c <= endCol; c++) {
				const cell = this.data.get(packKey(r, c))
				if (cell) acc = fn(acc, cell.value, r, c)
			}
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
			yield [row, col, cell] as const
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

	clone(): SparseGrid {
		const clone = new SparseGrid()
		for (const [key, cell] of this.data) {
			clone.data.set(key, structuredClone(cell))
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
}
