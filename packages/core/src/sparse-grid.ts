import type { CellValue } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import type { StyleId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface Cell {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
	readonly formulaInfo?: CellFormulaBinding
}

export interface SharedFormulaInfo {
	readonly kind: 'shared'
	readonly sharedIndex: string
	readonly isMaster: boolean
	readonly ref?: string
}

export interface ArrayFormulaInfo {
	readonly kind: 'array'
	readonly ref?: string
}

export interface SpillFormulaInfo {
	readonly kind: 'spill'
	readonly anchorRef: string
	readonly ref: string
	readonly isAnchor: boolean
}

export type CellFormulaBinding = SharedFormulaInfo | ArrayFormulaInfo | SpillFormulaInfo

const DEFAULT_STYLE_ID = 0 as StyleId

export class SparseGrid {
	private readonly data = new Map<number, StoredCell>()
	private readonly styledStringCache = new Map<string, StyledStringCell>()
	private readonly styledBooleanCache = new Map<string, StyledBooleanCell>()
	private readonly styledNumberCache = new Map<StyleId, Map<number, StyledNumberCell>>()
	private _isKeyOrderSorted = true
	private _lastInsertedKey = Number.NEGATIVE_INFINITY
	private _minRow = Number.POSITIVE_INFINITY
	private _maxRow = Number.NEGATIVE_INFINITY
	private _minCol = Number.POSITIVE_INFINITY
	private _maxCol = Number.NEGATIVE_INFINITY
	private _boundsDirty = false

	get(row: number, col: number): Cell | undefined {
		return materializeCell(this.data.get(packKey(row, col)))
	}

	set(row: number, col: number, cell: Cell): void {
		this.setResolved(row, col, cell.value, cell.formula, cell.styleId, cell.formulaInfo)
	}

	setResolved(
		row: number,
		col: number,
		value: CellValue,
		formula: string | null,
		styleId: StyleId,
		formulaInfo?: CellFormulaBinding,
	): void {
		const key = packKey(row, col)
		const isNew = !this.data.has(key)
		this.data.set(key, this.compactResolvedCell(value, formula, styleId, formulaInfo))
		if (isNew) {
			if (key < this._lastInsertedKey) this._isKeyOrderSorted = false
			else this._lastInsertedKey = key
		}
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

	has(row: number, col: number): boolean {
		return this.data.has(packKey(row, col))
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
		if (this._isKeyOrderSorted) {
			let currentRow = Number.NaN
			let rowCells: Array<readonly [number, Cell]> = []
			for (const [key, stored] of this.data) {
				const [row, col] = unpackKey(key)
				const cell = materializeCell(stored)
				if (!cell) continue
				if (!Number.isNaN(currentRow) && row !== currentRow) {
					yield [currentRow, rowCells] as const
					rowCells = []
				}
				currentRow = row
				rowCells.push([col, cell] as const)
			}
			if (!Number.isNaN(currentRow)) {
				yield [currentRow, rowCells] as const
			}
			return
		}
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

	*iterateRowsInRange(
		range: RangeRef,
	): Generator<readonly [number, readonly (readonly [number, Cell])[]]> {
		if (this._isKeyOrderSorted) {
			const minKey = packKey(range.start.row, 0)
			const maxKey = packKey(range.end.row, PACK_FACTOR - 1)
			let currentRow = Number.NaN
			let rowCells: Array<readonly [number, Cell]> = []
			for (const [key, stored] of this.data) {
				if (key < minKey) continue
				if (key > maxKey) break
				const [row, col] = unpackKey(key)
				if (col < range.start.col || col > range.end.col) continue
				const cell = materializeCell(stored)
				if (!cell) continue
				if (!Number.isNaN(currentRow) && row !== currentRow) {
					yield [currentRow, rowCells] as const
					rowCells = []
				}
				currentRow = row
				rowCells.push([col, cell] as const)
			}
			if (!Number.isNaN(currentRow) && rowCells.length > 0) {
				yield [currentRow, rowCells] as const
			}
			return
		}
		for (const [row, rowCells] of this.iterateRows()) {
			if (row < range.start.row) continue
			if (row > range.end.row) return
			const filtered = rowCells.filter(([col]) => col >= range.start.col && col <= range.end.col)
			if (filtered.length === 0) continue
			yield [row, filtered] as const
		}
	}

	cellCount(): number {
		return this.data.size
	}

	clear(): void {
		this.data.clear()
		this.styledStringCache.clear()
		this.styledBooleanCache.clear()
		this.styledNumberCache.clear()
		this._isKeyOrderSorted = true
		this._lastInsertedKey = Number.NEGATIVE_INFINITY
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
		clone.copyFrom(this)
		return clone
	}

	copyFrom(other: SparseGrid): void {
		this.data.clear()
		this.styledStringCache.clear()
		this.styledBooleanCache.clear()
		this.styledNumberCache.clear()
		for (const [key, cell] of other.data) {
			this.data.set(key, cloneCell(cell))
		}
		this._minRow = other._minRow
		this._maxRow = other._maxRow
		this._minCol = other._minCol
		this._maxCol = other._maxCol
		this._boundsDirty = other._boundsDirty
		this._isKeyOrderSorted = other._isKeyOrderSorted
		this._lastInsertedKey = other._lastInsertedKey
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
		let lastInsertedKey = Number.NEGATIVE_INFINITY
		for (const key of nextData.keys()) lastInsertedKey = key
		this._lastInsertedKey = lastInsertedKey
		this._boundsDirty = true
	}

	private compactResolvedCell(
		value: CellValue,
		formula: string | null,
		styleId: StyleId,
		formulaInfo?: CellFormulaBinding,
	): StoredCell {
		const compactValue = compactScalarValue(value)
		if (compactValue) {
			if (formula === null && styleId === DEFAULT_STYLE_ID) {
				switch (compactValue.kind) {
					case 'number':
					case 'string':
					case 'boolean':
						return compactValue.scalarValue as string | number | boolean
					case 'empty':
						return EMPTY
					default:
						return new StyledSpecialScalarCell(
							compactValue.kind,
							compactValue.scalarValue,
							DEFAULT_STYLE_ID,
						)
				}
			}
			if (formula === null) {
				switch (compactValue.kind) {
					case 'number': {
						const numericValue = compactValue.scalarValue as number
						let valuesByStyle = this.styledNumberCache.get(styleId)
						if (!valuesByStyle) {
							valuesByStyle = new Map()
							this.styledNumberCache.set(styleId, valuesByStyle)
						}
						let cached = valuesByStyle.get(numericValue)
						if (!cached) {
							cached = new StyledNumberCell(numericValue, styleId)
							valuesByStyle.set(numericValue, cached)
						}
						return cached
					}
					case 'string': {
						const text = compactValue.scalarValue as string
						const key = `${styleId}|${text}`
						let cached = this.styledStringCache.get(key)
						if (!cached) {
							cached = new StyledStringCell(text, styleId)
							this.styledStringCache.set(key, cached)
						}
						return cached
					}
					case 'boolean': {
						const bool = compactValue.scalarValue as boolean
						const key = `${styleId}|${bool ? 1 : 0}`
						let cached = this.styledBooleanCache.get(key)
						if (!cached) {
							cached = new StyledBooleanCell(bool, styleId)
							this.styledBooleanCache.set(key, cached)
						}
						return cached
					}
					default:
						return new StyledSpecialScalarCell(compactValue.kind, compactValue.scalarValue, styleId)
				}
			}
			return new FormulaScalarCell(
				compactValue.kind,
				compactValue.scalarValue,
				styleId,
				formula,
				formulaInfo,
			)
		}
		return new HeapCell(value, styleId, formula, formulaInfo)
	}
}

type StoredCell =
	| CellValue
	| string
	| number
	| boolean
	| StyledNumberCell
	| StyledStringCell
	| StyledBooleanCell
	| StyledSpecialScalarCell
	| FormulaScalarCell
	| HeapCell

class StyledNumberCell {
	constructor(
		readonly value: number,
		readonly styleId: StyleId,
	) {}
}

class StyledStringCell {
	constructor(
		readonly value: string,
		readonly styleId: StyleId,
	) {}
}

class StyledBooleanCell {
	constructor(
		readonly value: boolean,
		readonly styleId: StyleId,
	) {}
}

class StyledSpecialScalarCell {
	constructor(
		readonly valueKind: 'empty' | 'number' | 'string' | 'boolean' | 'error' | 'date',
		readonly scalarValue: number | string | boolean | null,
		readonly styleId: StyleId,
	) {}
}

class FormulaScalarCell {
	constructor(
		readonly valueKind: 'empty' | 'number' | 'string' | 'boolean' | 'error' | 'date',
		readonly scalarValue: number | string | boolean | null,
		readonly styleId: StyleId,
		readonly formula: string,
		readonly formulaInfo?: CellFormulaBinding,
	) {}
}

class HeapCell {
	constructor(
		readonly value: CellValue,
		readonly styleId: StyleId,
		readonly formula: string | null,
		readonly formulaInfo?: CellFormulaBinding,
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

function cloneCell(cell: StoredCell): StoredCell {
	if (
		cell instanceof StyledNumberCell ||
		cell instanceof StyledStringCell ||
		cell instanceof StyledBooleanCell ||
		cell instanceof StyledSpecialScalarCell ||
		cell instanceof FormulaScalarCell
	) {
		return cell
	}
	if (cell instanceof HeapCell) {
		return new HeapCell(structuredClone(cell.value), cell.styleId, cell.formula, cell.formulaInfo)
	}
	if (typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean') return cell
	if (cell === EMPTY) return cell
	return structuredClone(cell)
}

function materializeCell(cell: StoredCell | undefined): Cell | undefined {
	if (cell === undefined) return undefined
	if (cell instanceof StyledNumberCell) {
		return { value: numberValue(cell.value), formula: null, styleId: cell.styleId }
	}
	if (cell instanceof StyledStringCell) {
		return { value: stringValue(cell.value), formula: null, styleId: cell.styleId }
	}
	if (cell instanceof StyledBooleanCell) {
		return { value: booleanValue(cell.value), formula: null, styleId: cell.styleId }
	}
	if (cell instanceof StyledSpecialScalarCell) {
		return {
			value: materializeScalarValue(cell.valueKind, cell.scalarValue),
			formula: null,
			styleId: cell.styleId,
		}
	}
	if (cell instanceof FormulaScalarCell) {
		return {
			value: materializeScalarValue(cell.valueKind, cell.scalarValue),
			formula: cell.formula,
			styleId: cell.styleId,
			...(cell.formulaInfo ? { formulaInfo: cell.formulaInfo } : {}),
		}
	}
	if (cell instanceof HeapCell) {
		return {
			value: cell.value,
			formula: cell.formula,
			styleId: cell.styleId,
			...(cell.formulaInfo ? { formulaInfo: cell.formulaInfo } : {}),
		}
	}
	if (typeof cell === 'string')
		return { value: stringValue(cell), formula: null, styleId: DEFAULT_STYLE_ID }
	if (typeof cell === 'number')
		return { value: numberValue(cell), formula: null, styleId: DEFAULT_STYLE_ID }
	if (typeof cell === 'boolean') {
		return { value: booleanValue(cell), formula: null, styleId: DEFAULT_STYLE_ID }
	}
	return { value: cell, formula: null, styleId: DEFAULT_STYLE_ID }
}

function readStoredValue(cell: StoredCell | undefined): CellValue | undefined {
	if (cell === undefined) return undefined
	if (cell instanceof StyledNumberCell) return numberValue(cell.value)
	if (cell instanceof StyledStringCell) return stringValue(cell.value)
	if (cell instanceof StyledBooleanCell) return booleanValue(cell.value)
	if (cell instanceof StyledSpecialScalarCell || cell instanceof FormulaScalarCell) {
		return materializeScalarValue(cell.valueKind, cell.scalarValue)
	}
	if (cell instanceof HeapCell) return cell.value
	if (typeof cell === 'string') return stringValue(cell)
	if (typeof cell === 'number') return numberValue(cell)
	if (typeof cell === 'boolean') return booleanValue(cell)
	return cell
}

function compactScalarValue(value: CellValue): {
	kind: 'empty' | 'number' | 'string' | 'boolean' | 'error' | 'date'
	scalarValue: number | string | boolean | null
} | null {
	switch (value.kind) {
		case 'empty':
			return { kind: 'empty', scalarValue: null }
		case 'number':
			return { kind: 'number', scalarValue: value.value }
		case 'string':
			return { kind: 'string', scalarValue: value.value }
		case 'boolean':
			return { kind: 'boolean', scalarValue: value.value }
		case 'error':
			return { kind: 'error', scalarValue: value.value }
		case 'date':
			return { kind: 'date', scalarValue: value.serial }
		case 'array':
			return null
		default:
			return null
	}
}

function materializeScalarValue(
	kind: 'empty' | 'number' | 'string' | 'boolean' | 'error' | 'date',
	scalarValue: number | string | boolean | null,
): CellValue {
	switch (kind) {
		case 'empty':
			return EMPTY
		case 'number':
			return numberValue(scalarValue as number)
		case 'string':
			return stringValue(scalarValue as string)
		case 'boolean':
			return booleanValue(Boolean(scalarValue))
		case 'error':
			return errorValue(scalarValue as never)
		case 'date':
			return { kind: 'date', serial: scalarValue as number }
	}
}
