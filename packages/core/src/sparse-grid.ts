const YOUNG_POOL_LIMIT = 8_192
const OLD_POOL_LIMIT = 32_768
const PROMOTION_THRESHOLD = 3
const youngPool = new Map<string, string>()
const youngHits = new Map<string, number>()
const oldPool = new Map<string, string>()

function internString(s: string): string {
	const fromOld = oldPool.get(s)
	if (fromOld !== undefined) return fromOld

	const fromYoung = youngPool.get(s)
	if (fromYoung !== undefined) {
		const hits = (youngHits.get(s) ?? 0) + 1
		youngHits.set(s, hits)
		if (hits >= PROMOTION_THRESHOLD) {
			youngPool.delete(s)
			youngHits.delete(s)
			if (oldPool.size >= OLD_POOL_LIMIT) {
				const iter = oldPool.keys()
				const quarter = OLD_POOL_LIMIT >> 2
				for (let i = 0; i < quarter; i++) iter.next()
			}
			oldPool.set(s, s)
		}
		return fromYoung
	}

	if (youngPool.size >= YOUNG_POOL_LIMIT) {
		youngPool.clear()
		youngHits.clear()
	}
	youngPool.set(s, s)
	youngHits.set(s, 1)
	return s
}

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
	readonly masterRef?: string
	readonly ref?: string
}

export interface ArrayFormulaInfo {
	readonly kind: 'array'
	readonly ref?: string
}

export interface DynamicArrayFormulaInfo {
	readonly kind: 'dynamicArray'
	readonly metadataIndex: number
	readonly collapsed?: boolean
}

export interface SpillFormulaInfo {
	readonly kind: 'spill'
	readonly anchorRef: string
	readonly ref: string
	readonly isAnchor: boolean
}

export type CellFormulaBinding =
	| SharedFormulaInfo
	| ArrayFormulaInfo
	| DynamicArrayFormulaInfo
	| SpillFormulaInfo

const DEFAULT_STYLE_ID = 0 as StyleId

export class SparseGrid {
	private data = new Map<number, StoredCell>()
	private rowIndex = new Map<number, Map<number, StoredCell>>()
	private readonly styledStringCache = new Map<StyleId, Map<string, StyledStringCell>>()
	private readonly styledBooleanCache = new Map<StyleId, Map<string, StyledBooleanCell>>()
	private readonly styledNumberCache = new Map<StyleId, Map<number, StyledNumberCell>>()
	private _isKeyOrderSorted = true
	private _lastInsertedKey = Number.NEGATIVE_INFINITY
	private _minRow = Number.POSITIVE_INFINITY
	private _maxRow = Number.NEGATIVE_INFINITY
	private _minCol = Number.POSITIVE_INFINITY
	private _maxCol = Number.NEGATIVE_INFINITY
	private _boundsDirty = false
	private _hasMutableStorage = false
	private _shared = false
	private _denseArray: (CellValue | undefined)[] | null = null
	private _denseMinRow = 0
	private _denseMinCol = 0
	private _denseNumCols = 0
	private _densityChecked = false

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
		this.ensureWritable()
		const key = packKey(row, col)
		const isNew = !this.data.has(key)
		const stored = this.compactResolvedCell(value, formula, styleId, formulaInfo)
		this.data.set(key, stored)
		let rowMap = this.rowIndex.get(row)
		if (!rowMap) {
			rowMap = new Map()
			this.rowIndex.set(row, rowMap)
		}
		rowMap.set(col, stored)
		if (isMutableStoredCell(stored)) this._hasMutableStorage = true
		if (isNew) {
			if (key < this._lastInsertedKey) this._isKeyOrderSorted = false
			else this._lastInsertedKey = key
		}
		if (row < this._minRow) this._minRow = row
		if (row > this._maxRow) this._maxRow = row
		if (col < this._minCol) this._minCol = col
		if (col > this._maxCol) this._maxCol = col
		if (this._denseArray !== null) {
			const r = row - this._denseMinRow
			const c = col - this._denseMinCol
			if (r >= 0 && c >= 0 && c < this._denseNumCols) {
				const idx = r * this._denseNumCols + c
				if (idx < this._denseArray.length) {
					this._denseArray[idx] = value
				} else {
					this._denseArray = null
				}
			} else {
				this._denseArray = null
			}
		}
	}

	delete(row: number, col: number): boolean {
		this.ensureWritable()
		const deleted = this.data.delete(packKey(row, col))
		if (deleted) {
			const rowMap = this.rowIndex.get(row)
			if (rowMap) {
				rowMap.delete(col)
				if (rowMap.size === 0) this.rowIndex.delete(row)
			}
			if (
				row === this._minRow ||
				row === this._maxRow ||
				col === this._minCol ||
				col === this._maxCol
			) {
				this._boundsDirty = true
			}
			if (this._denseArray !== null) {
				const r = row - this._denseMinRow
				const c = col - this._denseMinCol
				if (r >= 0 && c >= 0 && c < this._denseNumCols) {
					const idx = r * this._denseNumCols + c
					if (idx < this._denseArray.length) {
						this._denseArray[idx] = undefined
					}
				}
			}
		}
		return deleted
	}

	getValue(row: number, col: number): CellValue | undefined {
		return readStoredValue(this.data.get(packKey(row, col)))
	}

	readValue(row: number, col: number): CellValue {
		let arr = this._denseArray
		if (arr === null && !this._densityChecked) {
			this._densityChecked = true
			this._maybeEnableDenseCache()
			arr = this._denseArray
		}
		if (arr !== null) {
			const r = row - this._denseMinRow
			const c = col - this._denseMinCol
			if (r >= 0 && c >= 0 && c < this._denseNumCols) {
				const idx = r * this._denseNumCols + c
				if (idx < arr.length) return arr[idx] ?? EMPTY
			}
			return EMPTY
		}
		return readStoredValue(this.data.get(packKey(row, col))) ?? EMPTY
	}

	has(row: number, col: number): boolean {
		return this.data.has(packKey(row, col))
	}

	*getRange(range: RangeRef): Generator<readonly [number, number, Cell]> {
		for (let r = range.start.row; r <= range.end.row; r++) {
			const rowMap = this.rowIndex.get(r)
			if (!rowMap) continue
			for (const [col, stored] of rowMap) {
				if (col < range.start.col || col > range.end.col) continue
				const cell = materializeCell(stored)
				if (cell) yield [r, col, cell] as const
			}
		}
	}

	forEachValueInRange(
		startRow: number,
		startCol: number,
		endRow: number,
		endCol: number,
		fn: (value: CellValue, row: number, col: number) => void,
	): void {
		if (startCol === endCol) {
			for (let r = startRow; r <= endRow; r++) {
				const rowMap = this.rowIndex.get(r)
				if (!rowMap) continue
				const stored = rowMap.get(startCol)
				if (!stored) continue
				const value = readStoredValue(stored)
				if (value) fn(value, r, startCol)
			}
			return
		}
		for (let r = startRow; r <= endRow; r++) {
			const rowMap = this.rowIndex.get(r)
			if (!rowMap) continue
			for (const [col, stored] of rowMap) {
				if (col < startCol || col > endCol) continue
				const value = readStoredValue(stored)
				if (value) fn(value, r, col)
			}
		}
	}

	forEachRow(fn: (row: number, cells: ReadonlyMap<number, CellValue>) => void): void {
		if (this._isKeyOrderSorted) {
			let currentRow = Number.NaN
			const rowCells = new Map<number, CellValue>()
			for (const [key, stored] of this.data) {
				const [row, col] = unpackKey(key)
				const value = readStoredValue(stored)
				if (!value) continue
				if (!Number.isNaN(currentRow) && row !== currentRow) {
					fn(currentRow, rowCells)
					rowCells.clear()
				}
				currentRow = row
				rowCells.set(col, value)
			}
			if (!Number.isNaN(currentRow)) {
				fn(currentRow, rowCells)
			}
			return
		}
		const rows = new Map<number, Map<number, CellValue>>()
		for (const [key, stored] of this.data) {
			const [row, col] = unpackKey(key)
			const value = readStoredValue(stored)
			if (!value) continue
			let rowCells = rows.get(row)
			if (!rowCells) {
				rowCells = new Map()
				rows.set(row, rowCells)
			}
			rowCells.set(col, value)
		}
		const sortedRows = [...rows.entries()].sort((a, b) => a[0] - b[0])
		for (const [row, rowCells] of sortedRows) {
			const sortedCells = new Map([...rowCells.entries()].sort((a, b) => a[0] - b[0]))
			fn(row, sortedCells)
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
			const rowMap = this.rowIndex.get(r)
			if (!rowMap) continue
			for (const [col, stored] of rowMap) {
				if (col < startCol || col > endCol) continue
				const value = readStoredValue(stored)
				if (value) acc = fn(acc, value, r, col)
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
		for (let r = range.start.row; r <= range.end.row; r++) {
			const rowMap = this.rowIndex.get(r)
			if (!rowMap) continue
			const rowCells: Array<readonly [number, Cell]> = []
			for (const [col, stored] of rowMap) {
				if (col < range.start.col || col > range.end.col) continue
				const cell = materializeCell(stored)
				if (cell) rowCells.push([col, cell] as const)
			}
			if (rowCells.length > 0) {
				rowCells.sort((a, b) => a[0] - b[0])
				yield [r, rowCells] as const
			}
		}
	}

	cellCount(): number {
		return this.data.size
	}

	clear(): void {
		if (this._shared) {
			this.data = new Map()
			this.rowIndex = new Map()
			this._shared = false
		} else {
			this.data.clear()
			this.rowIndex.clear()
		}
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
		this._hasMutableStorage = false
		this._denseArray = null
		this._densityChecked = false
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
		if (other._hasMutableStorage) {
			this.data = new Map(other.data)
			for (const [key, cell] of this.data) {
				const cloned = cloneCell(cell)
				if (cloned !== cell) this.data.set(key, cloned)
			}
			this._rebuildRowIndex()
			this._shared = false
		} else {
			this.data = other.data
			this.rowIndex = other.rowIndex
			this._shared = true
			other._shared = true
		}
		this.styledStringCache.clear()
		this.styledBooleanCache.clear()
		this.styledNumberCache.clear()
		this._minRow = other._minRow
		this._maxRow = other._maxRow
		this._minCol = other._minCol
		this._maxCol = other._maxCol
		this._boundsDirty = other._boundsDirty
		this._isKeyOrderSorted = other._isKeyOrderSorted
		this._lastInsertedKey = other._lastInsertedKey
		this._hasMutableStorage = other._hasMutableStorage
		this._denseArray = null
		this._densityChecked = false
	}

	private ensureWritable(): void {
		if (!this._shared) return
		this.data = new Map(this.data)
		if (this._hasMutableStorage) {
			for (const [key, cell] of this.data) {
				const cloned = cloneCell(cell)
				if (cloned !== cell) this.data.set(key, cloned)
			}
		}
		this._rebuildRowIndex()
		this._shared = false
		this._denseArray = null
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

	private _rebuildRowIndex(): void {
		this.rowIndex = new Map()
		for (const [key, cell] of this.data) {
			const [r, c] = unpackKey(key)
			let rowMap = this.rowIndex.get(r)
			if (!rowMap) {
				rowMap = new Map()
				this.rowIndex.set(r, rowMap)
			}
			rowMap.set(c, cell)
		}
	}

	private _rebuildAfterShift(
		mapCell: (row: number, col: number) => { row: number; col: number } | null,
		active: boolean,
	): void {
		if (!active || this.data.size === 0) return
		const cloneMutable = this._shared && this._hasMutableStorage
		const nextData = new Map<number, StoredCell>()
		for (const [key, cell] of this.data) {
			const [row, col] = unpackKey(key)
			const next = mapCell(row, col)
			if (!next) continue
			nextData.set(packKey(next.row, next.col), cloneMutable ? cloneCell(cell) : cell)
		}
		this.data = nextData
		this._shared = false
		this._rebuildRowIndex()
		let lastInsertedKey = Number.NEGATIVE_INFINITY
		for (const key of nextData.keys()) lastInsertedKey = key
		this._lastInsertedKey = lastInsertedKey
		this._boundsDirty = true
		this._denseArray = null
		this._densityChecked = false
	}

	private _maybeEnableDenseCache(): void {
		if (this.data.size < 2) return
		if (this._boundsDirty) this._recomputeBounds()
		const numRows = this._maxRow - this._minRow + 1
		const numCols = this._maxCol - this._minCol + 1
		const gridSize = numRows * numCols
		if (gridSize <= 0 || gridSize > 1_000_000) return
		if (this.data.size / gridSize <= 0.5) return
		this._denseMinRow = this._minRow
		this._denseMinCol = this._minCol
		this._denseNumCols = numCols
		const arr = new Array<CellValue | undefined>(gridSize)
		for (const [key, stored] of this.data) {
			const [r, c] = unpackKey(key)
			arr[(r - this._minRow) * numCols + (c - this._minCol)] = readStoredValue(stored)
		}
		this._denseArray = arr
	}

	private compactResolvedCell(
		value: CellValue,
		formula: string | null,
		styleId: StyleId,
		formulaInfo?: CellFormulaBinding,
	): StoredCell {
		const compactValue = compactScalarValue(value)
		if (compactValue) {
			if (compactValue.kind === 'string') {
				compactValue.scalarValue = internString(compactValue.scalarValue as string) as string
			}
			if (formula === null && formulaInfo === undefined && styleId === DEFAULT_STYLE_ID) {
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
			if (formula === null && formulaInfo === undefined) {
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
						let byText = this.styledStringCache.get(styleId)
						if (!byText) {
							byText = new Map()
							this.styledStringCache.set(styleId, byText)
						}
						let cached = byText.get(text)
						if (!cached) {
							cached = new StyledStringCell(text, styleId)
							byText.set(text, cached)
						}
						return cached
					}
					case 'boolean': {
						const bool = compactValue.scalarValue as boolean
						const textKey = bool ? '1' : '0'
						let byText = this.styledBooleanCache.get(styleId)
						if (!byText) {
							byText = new Map()
							this.styledBooleanCache.set(styleId, byText)
						}
						let cached = byText.get(textKey)
						if (!cached) {
							cached = new StyledBooleanCell(bool, styleId)
							byText.set(textKey, cached)
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
		readonly formula: string | null,
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

function cloneCellValue(value: CellValue): CellValue {
	if (value.kind === 'array') return structuredClone(value)
	if (value.kind === 'richText') return { kind: 'richText', runs: [...value.runs] }
	return value
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
		return new HeapCell(cloneCellValue(cell.value), cell.styleId, cell.formula, cell.formulaInfo)
	}
	if (typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean') return cell
	if (cell === EMPTY) return cell
	return cloneCellValue(cell) as StoredCell
}

function isMutableStoredCell(cell: StoredCell): boolean {
	if (cell instanceof HeapCell) {
		return cell.value.kind === 'array' || cell.value.kind === 'richText'
	}
	return (
		typeof cell === 'object' &&
		cell !== null &&
		'kind' in cell &&
		(cell.kind === 'array' || cell.kind === 'richText')
	)
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
