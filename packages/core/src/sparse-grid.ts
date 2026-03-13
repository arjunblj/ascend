import type { CellValue, ExcelError } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import type { StyleId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface Cell {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
	readonly formulaInfo?: CellFormulaBinding | undefined
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
const CHUNK_BITS = 6
const CHUNK_SIZE = 1 << CHUNK_BITS
const CHUNK_MASK = CHUNK_SIZE - 1
const CHUNK_AREA = CHUNK_SIZE * CHUNK_SIZE
const SPARSE_TO_DENSE_THRESHOLD = 96

enum SlotTag {
	Empty = 0,
	Number = 1,
	String = 2,
	Boolean = 3,
	Error = 4,
	Date = 5,
	Heap = 6,
}

type CellValueKind = CellValue['kind']

interface StoredSlot {
	readonly tag: SlotTag
	readonly styleId: StyleId
	readonly formula: string | null
	readonly formulaInfo: CellFormulaBinding | undefined
	readonly numberValue: number | undefined
	readonly stringId: number | undefined
	readonly heapValue: CellValue | undefined
}

interface GridChunk {
	readonly count: number
	has(localIndex: number): boolean
	getSlot(localIndex: number): StoredSlot | undefined
	getKind(localIndex: number): CellValueKind | undefined
	readValue(localIndex: number, stringTable: StringTable): CellValue | undefined
	readNumber(localIndex: number): number | null
	readString(localIndex: number, stringTable: StringTable): string | null
	readError(localIndex: number, stringTable: StringTable): ExcelError | null
	setSlot(localIndex: number, slot: StoredSlot): GridChunk
	delete(localIndex: number): boolean
	forEachRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		fn: (localCol: number, slot: StoredSlot) => void,
	): void
}

class StringTable {
	private readonly ids = new Map<string, number>()
	private readonly values = ['']

	intern(value: string): number {
		const existing = this.ids.get(value)
		if (existing !== undefined) return existing
		const id = this.values.length
		this.values.push(value)
		this.ids.set(value, id)
		return id
	}

	lookup(id: number): string {
		return this.values[id] ?? ''
	}
}

class SparseChunk implements GridChunk {
	private readonly slots = new Map<number, StoredSlot>()

	get count(): number {
		return this.slots.size
	}

	has(localIndex: number): boolean {
		return this.slots.has(localIndex)
	}

	getSlot(localIndex: number): StoredSlot | undefined {
		return this.slots.get(localIndex)
	}

	getKind(localIndex: number): CellValueKind | undefined {
		return kindFromSlot(this.slots.get(localIndex))
	}

	readValue(localIndex: number, stringTable: StringTable): CellValue | undefined {
		return readSlotValue(this.slots.get(localIndex), stringTable)
	}

	readNumber(localIndex: number): number | null {
		const slot = this.slots.get(localIndex)
		if (!slot) return null
		return slot.tag === SlotTag.Number || slot.tag === SlotTag.Boolean || slot.tag === SlotTag.Date
			? (slot.numberValue ?? 0)
			: null
	}

	readString(localIndex: number, stringTable: StringTable): string | null {
		const slot = this.slots.get(localIndex)
		if (!slot || slot.tag !== SlotTag.String) return null
		return stringTable.lookup(slot.stringId ?? 0)
	}

	readError(localIndex: number, stringTable: StringTable): ExcelError | null {
		const slot = this.slots.get(localIndex)
		if (!slot || slot.tag !== SlotTag.Error) return null
		return stringTable.lookup(slot.stringId ?? 0) as ExcelError
	}

	setSlot(localIndex: number, slot: StoredSlot): GridChunk {
		this.slots.set(localIndex, slot)
		if (this.slots.size < SPARSE_TO_DENSE_THRESHOLD) return this
		const dense = new DenseChunk()
		for (const [index, existing] of this.slots) {
			dense.setSlot(index, existing)
		}
		return dense
	}

	delete(localIndex: number): boolean {
		return this.slots.delete(localIndex)
	}

	forEachRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		fn: (localCol: number, slot: StoredSlot) => void,
	): void {
		const rowSlots: Array<readonly [number, StoredSlot]> = []
		for (const [index, slot] of this.slots) {
			const row = index >>> CHUNK_BITS
			if (row !== localRow) continue
			const col = index & CHUNK_MASK
			if (col < minLocalCol || col > maxLocalCol) continue
			rowSlots.push([col, slot] as const)
		}
		rowSlots.sort((a, b) => a[0] - b[0])
		for (const [col, slot] of rowSlots) fn(col, slot)
	}
}

class DenseChunk implements GridChunk {
	private readonly metaBuffer = new SharedArrayBuffer(CHUNK_AREA * 10)
	private readonly numberBuffer = new SharedArrayBuffer(CHUNK_AREA * 8)
	private readonly occupied = new Uint8Array(this.metaBuffer, 0, CHUNK_AREA)
	private readonly tags = new Uint8Array(this.metaBuffer, CHUNK_AREA, CHUNK_AREA)
	private readonly styleIds = new Uint32Array(this.metaBuffer, CHUNK_AREA * 2, CHUNK_AREA)
	private readonly stringIds = new Uint32Array(this.metaBuffer, CHUNK_AREA * 6, CHUNK_AREA)
	private readonly numbers = new Float64Array(this.numberBuffer)
	private readonly rowCounts = new Uint16Array(CHUNK_SIZE)
	private formulas: Array<string | null | undefined> | null = null
	private formulaInfos: Array<CellFormulaBinding | undefined> | null = null
	private heapValues: Array<CellValue | undefined> | null = null
	private _count = 0

	get count(): number {
		return this._count
	}

	has(localIndex: number): boolean {
		return this.occupied[localIndex] === 1
	}

	getSlot(localIndex: number): StoredSlot | undefined {
		if (!this.has(localIndex)) return undefined
		const tag = this.tags[localIndex] as SlotTag
		return {
			tag,
			styleId: this.styleIds[localIndex] as StyleId,
			formula: this.formulas?.[localIndex] ?? null,
			formulaInfo: this.formulaInfos?.[localIndex],
			numberValue:
				tag === SlotTag.Number || tag === SlotTag.Boolean || tag === SlotTag.Date
					? (this.numbers[localIndex] ?? 0)
					: undefined,
			stringId:
				tag === SlotTag.String || tag === SlotTag.Error
					? (this.stringIds[localIndex] ?? 0)
					: undefined,
			heapValue: tag === SlotTag.Heap ? (this.heapValues?.[localIndex] ?? EMPTY) : undefined,
		}
	}

	getKind(localIndex: number): CellValueKind | undefined {
		if (!this.has(localIndex)) return undefined
		const tag = (this.tags[localIndex] ?? SlotTag.Empty) as SlotTag
		return tag === SlotTag.Heap
			? (this.heapValues?.[localIndex]?.kind ?? 'array')
			: kindFromTag(tag)
	}

	readValue(localIndex: number, stringTable: StringTable): CellValue | undefined {
		if (!this.has(localIndex)) return undefined
		return readDenseValue(
			(this.tags[localIndex] ?? SlotTag.Empty) as SlotTag,
			this.numbers[localIndex] ?? 0,
			this.stringIds[localIndex] ?? 0,
			this.heapValues?.[localIndex],
			stringTable,
		)
	}

	readNumber(localIndex: number): number | null {
		if (!this.has(localIndex)) return null
		const tag = (this.tags[localIndex] ?? SlotTag.Empty) as SlotTag
		return tag === SlotTag.Number || tag === SlotTag.Boolean || tag === SlotTag.Date
			? (this.numbers[localIndex] ?? 0)
			: null
	}

	readString(localIndex: number, stringTable: StringTable): string | null {
		if (!this.has(localIndex) || this.tags[localIndex] !== SlotTag.String) return null
		return stringTable.lookup(this.stringIds[localIndex] ?? 0)
	}

	readError(localIndex: number, stringTable: StringTable): ExcelError | null {
		if (!this.has(localIndex) || this.tags[localIndex] !== SlotTag.Error) return null
		return stringTable.lookup(this.stringIds[localIndex] ?? 0) as ExcelError
	}

	setSlot(localIndex: number, slot: StoredSlot): GridChunk {
		if (!this.has(localIndex)) {
			this.occupied[localIndex] = 1
			const rowIndex = localIndex >>> CHUNK_BITS
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this._count++
		}
		this.tags[localIndex] = slot.tag
		this.styleIds[localIndex] = slot.styleId
		this.stringIds[localIndex] = slot.stringId ?? 0
		this.numbers[localIndex] = slot.numberValue ?? 0
		if (slot.formula !== null) {
			if (this.formulas === null) this.formulas = new Array(CHUNK_AREA)
			this.formulas[localIndex] = slot.formula
		} else if (this.formulas !== null) {
			this.formulas[localIndex] = null
		}
		if (slot.formulaInfo !== undefined) {
			if (this.formulaInfos === null) this.formulaInfos = new Array(CHUNK_AREA)
			this.formulaInfos[localIndex] = slot.formulaInfo
		} else if (this.formulaInfos !== null) {
			this.formulaInfos[localIndex] = undefined
		}
		if (slot.tag === SlotTag.Heap) {
			if (this.heapValues === null) this.heapValues = new Array(CHUNK_AREA)
			this.heapValues[localIndex] = slot.heapValue
		} else if (this.heapValues !== null) {
			this.heapValues[localIndex] = undefined
		}
		return this
	}

	delete(localIndex: number): boolean {
		if (!this.has(localIndex)) return false
		this.occupied[localIndex] = 0
		this.tags[localIndex] = SlotTag.Empty
		this.styleIds[localIndex] = DEFAULT_STYLE_ID
		this.stringIds[localIndex] = 0
		this.numbers[localIndex] = 0
		if (this.formulas !== null) this.formulas[localIndex] = undefined
		if (this.formulaInfos !== null) this.formulaInfos[localIndex] = undefined
		if (this.heapValues !== null) this.heapValues[localIndex] = undefined
		const rowIndex = localIndex >>> CHUNK_BITS
		this.rowCounts[rowIndex] = Math.max(0, (this.rowCounts[rowIndex] ?? 0) - 1)
		this._count--
		return true
	}

	forEachRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		fn: (localCol: number, slot: StoredSlot) => void,
	): void {
		if ((this.rowCounts[localRow] ?? 0) === 0) return
		const rowBase = localRow << CHUNK_BITS
		for (let localCol = minLocalCol; localCol <= maxLocalCol; localCol++) {
			const localIndex = rowBase + localCol
			if (!this.has(localIndex)) continue
			const slot = this.getSlot(localIndex)
			if (slot) fn(localCol, slot)
		}
	}
}

export class SparseGrid {
	private chunkRows = new Map<number, Map<number, GridChunk>>()
	private stringTable = new StringTable()
	private _cellCount = 0
	private _shared = false
	private _minRow = Number.POSITIVE_INFINITY
	private _maxRow = Number.NEGATIVE_INFINITY
	private _minCol = Number.POSITIVE_INFINITY
	private _maxCol = Number.NEGATIVE_INFINITY
	private _boundsDirty = false

	get(row: number, col: number): Cell | undefined {
		const slot = this._getSlot(row, col)
		return materializeCell(slot, this.stringTable)
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
		const slot = compactResolvedCell(value, formula, styleId, formulaInfo, this.stringTable)
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		let cols = this.chunkRows.get(chunkRow)
		if (!cols) {
			cols = new Map()
			this.chunkRows.set(chunkRow, cols)
		}
		let chunk = cols.get(chunkCol)
		if (!chunk) {
			chunk = new SparseChunk()
			cols.set(chunkCol, chunk)
		}
		const existed = chunk.has(localIndex)
		const nextChunk = chunk.setSlot(localIndex, slot)
		if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
		if (!existed) this._cellCount++
		this._trackBounds(row, col)
	}

	delete(row: number, col: number): boolean {
		this.ensureWritable()
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		const cols = this.chunkRows.get(chunkRow)
		const chunk = cols?.get(chunkCol)
		if (!chunk) return false
		const deleted = chunk.delete(localIndex)
		if (!deleted) return false
		this._cellCount--
		if (chunk.count === 0) {
			cols?.delete(chunkCol)
			if (cols && cols.size === 0) this.chunkRows.delete(chunkRow)
		}
		if (
			row === this._minRow ||
			row === this._maxRow ||
			col === this._minCol ||
			col === this._maxCol
		) {
			this._boundsDirty = true
		}
		return true
	}

	getValue(row: number, col: number): CellValue | undefined {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readValue(localIndex, this.stringTable)
	}

	readValue(row: number, col: number): CellValue {
		return this.getValue(row, col) ?? EMPTY
	}

	readKind(row: number, col: number): CellValueKind | undefined {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.getKind(localIndex)
	}

	readNumber(row: number, col: number): number | null {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readNumber(localIndex) ?? null
	}

	readString(row: number, col: number): string | null {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return (
			this.chunkRows.get(chunkRow)?.get(chunkCol)?.readString(localIndex, this.stringTable) ?? null
		)
	}

	readError(row: number, col: number): ExcelError | null {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return (
			this.chunkRows.get(chunkRow)?.get(chunkCol)?.readError(localIndex, this.stringTable) ?? null
		)
	}

	has(row: number, col: number): boolean {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.has(localIndex) ?? false
	}

	*getRange(range: RangeRef): Generator<readonly [number, number, Cell]> {
		for (const [row, entries] of this.iterateRowsInRange(range)) {
			for (const [col, cell] of entries) {
				yield [row, col, cell] as const
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
		for (const [row, entries] of this.iterateRowsInRange({
			start: { row: startRow, col: startCol },
			end: { row: endRow, col: endCol },
		})) {
			for (const [col, cell] of entries) {
				fn(cell.value, row, col)
			}
		}
	}

	forEachRow(fn: (row: number, cells: ReadonlyMap<number, CellValue>) => void): void {
		for (const [row, entries] of this.iterateRows()) {
			const rowCells = new Map<number, CellValue>()
			for (const [col, cell] of entries) rowCells.set(col, cell.value)
			fn(row, rowCells)
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
		this.forEachValueInRange(startRow, startCol, endRow, endCol, (value, row, col) => {
			acc = fn(acc, value, row, col)
		})
		return acc
	}

	usedRange(): RangeRef | null {
		if (this._cellCount === 0) return null
		if (this._boundsDirty) this._recomputeBounds()
		return {
			start: { row: this._minRow, col: this._minCol },
			end: { row: this._maxRow, col: this._maxCol },
		}
	}

	*iterate(): Generator<readonly [number, number, Cell]> {
		for (const [row, entries] of this.iterateRows()) {
			for (const [col, cell] of entries) {
				yield [row, col, cell] as const
			}
		}
	}

	*iterateRows(): Generator<readonly [number, readonly (readonly [number, Cell])[]]> {
		if (this._cellCount === 0) return
		const sortedChunkRows = [...this.chunkRows.keys()].sort((a, b) => a - b)
		for (const chunkRow of sortedChunkRows) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const sortedChunkCols = [...cols.keys()].sort((a, b) => a - b)
			for (let localRow = 0; localRow < CHUNK_SIZE; localRow++) {
				const row = (chunkRow << CHUNK_BITS) + localRow
				const rowCells: Array<readonly [number, Cell]> = []
				for (const chunkCol of sortedChunkCols) {
					const chunk = cols.get(chunkCol)
					if (!chunk) continue
					const baseCol = chunkCol << CHUNK_BITS
					chunk.forEachRow(localRow, 0, CHUNK_MASK, (localCol, slot) => {
						const cell = materializeCell(slot, this.stringTable)
						if (cell) rowCells.push([baseCol + localCol, cell] as const)
					})
				}
				if (rowCells.length > 0) yield [row, rowCells] as const
			}
		}
	}

	*iterateRowsInRange(
		range: RangeRef,
	): Generator<readonly [number, readonly (readonly [number, Cell])[]]> {
		if (this._cellCount === 0) return
		const startChunkRow = range.start.row >> CHUNK_BITS
		const endChunkRow = range.end.row >> CHUNK_BITS
		for (let chunkRow = startChunkRow; chunkRow <= endChunkRow; chunkRow++) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const sortedChunkCols = [...cols.keys()].sort((a, b) => a - b)
			const rowStart = chunkRow === startChunkRow ? range.start.row & CHUNK_MASK : 0
			const rowEnd = chunkRow === endChunkRow ? range.end.row & CHUNK_MASK : CHUNK_MASK
			for (let localRow = rowStart; localRow <= rowEnd; localRow++) {
				const row = (chunkRow << CHUNK_BITS) + localRow
				const rowCells: Array<readonly [number, Cell]> = []
				for (const chunkCol of sortedChunkCols) {
					const chunk = cols.get(chunkCol)
					if (!chunk) continue
					const chunkStartCol = chunkCol << CHUNK_BITS
					const chunkEndCol = chunkStartCol + CHUNK_MASK
					if (chunkEndCol < range.start.col || chunkStartCol > range.end.col) continue
					const minLocalCol =
						chunkCol === range.start.col >> CHUNK_BITS ? range.start.col & CHUNK_MASK : 0
					const maxLocalCol =
						chunkCol === range.end.col >> CHUNK_BITS ? range.end.col & CHUNK_MASK : CHUNK_MASK
					chunk.forEachRow(localRow, minLocalCol, maxLocalCol, (localCol, slot) => {
						const cell = materializeCell(slot, this.stringTable)
						if (cell) rowCells.push([chunkStartCol + localCol, cell] as const)
					})
				}
				if (rowCells.length > 0) yield [row, rowCells] as const
			}
		}
	}

	cellCount(): number {
		return this._cellCount
	}

	clear(): void {
		this.chunkRows = new Map()
		this._cellCount = 0
		this._shared = false
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
		if (other._hasMutableValues()) {
			const next = new SparseGrid()
			next.stringTable = other.stringTable
			for (const [row, col, cell] of other.iterate()) {
				next.setResolved(
					row,
					col,
					cloneCellValue(cell.value),
					cell.formula,
					cell.styleId,
					cell.formulaInfo,
				)
			}
			this.chunkRows = next.chunkRows
			this.stringTable = next.stringTable
			this._cellCount = next._cellCount
			this._minRow = next._minRow
			this._maxRow = next._maxRow
			this._minCol = next._minCol
			this._maxCol = next._maxCol
			this._boundsDirty = next._boundsDirty
			this._shared = false
			return
		}
		this.chunkRows = other.chunkRows
		this.stringTable = other.stringTable
		this._cellCount = other._cellCount
		this._minRow = other._minRow
		this._maxRow = other._maxRow
		this._minCol = other._minCol
		this._maxCol = other._maxCol
		this._boundsDirty = other._boundsDirty
		this._shared = true
		other._shared = true
	}

	private ensureWritable(): void {
		if (!this._shared) return
		const next = new SparseGrid()
		next.stringTable = this.stringTable
		for (const [row, col, cell] of this.iterate()) {
			next.setResolved(
				row,
				col,
				cloneCellValue(cell.value),
				cell.formula,
				cell.styleId,
				cell.formulaInfo,
			)
		}
		this.chunkRows = next.chunkRows
		this._cellCount = next._cellCount
		this._minRow = next._minRow
		this._maxRow = next._maxRow
		this._minCol = next._minCol
		this._maxCol = next._maxCol
		this._boundsDirty = next._boundsDirty
		this._shared = false
	}

	private _getSlot(row: number, col: number): StoredSlot | undefined {
		const [chunkRow, chunkCol, localIndex] = getChunkPosition(row, col)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.getSlot(localIndex)
	}

	private _trackBounds(row: number, col: number): void {
		if (row < this._minRow) this._minRow = row
		if (row > this._maxRow) this._maxRow = row
		if (col < this._minCol) this._minCol = col
		if (col > this._maxCol) this._maxCol = col
	}

	private _recomputeBounds(): void {
		this._minRow = Number.POSITIVE_INFINITY
		this._maxRow = Number.NEGATIVE_INFINITY
		this._minCol = Number.POSITIVE_INFINITY
		this._maxCol = Number.NEGATIVE_INFINITY
		for (const [row, entries] of this.iterateRows()) {
			if (row < this._minRow) this._minRow = row
			if (row > this._maxRow) this._maxRow = row
			const first = entries[0]
			const last = entries[entries.length - 1]
			if (first && first[0] < this._minCol) this._minCol = first[0]
			if (last && last[0] > this._maxCol) this._maxCol = last[0]
		}
		this._boundsDirty = false
	}

	private _rebuildAfterShift(
		mapCell: (row: number, col: number) => { row: number; col: number } | null,
		active: boolean,
	): void {
		if (!active || this._cellCount === 0) return
		const next = new SparseGrid()
		next.stringTable = this.stringTable
		for (const [row, col, cell] of this.iterate()) {
			const mapped = mapCell(row, col)
			if (!mapped) continue
			next.setResolved(
				mapped.row,
				mapped.col,
				cloneCellValue(cell.value),
				cell.formula,
				cell.styleId,
				cell.formulaInfo,
			)
		}
		this.chunkRows = next.chunkRows
		this._cellCount = next._cellCount
		this._minRow = next._minRow
		this._maxRow = next._maxRow
		this._minCol = next._minCol
		this._maxCol = next._maxCol
		this._boundsDirty = next._boundsDirty
		this._shared = false
	}

	private _hasMutableValues(): boolean {
		for (const [, , cell] of this.iterate()) {
			if (cell.value.kind === 'array' || cell.value.kind === 'richText') return true
		}
		return false
	}
}

function getChunkPosition(row: number, col: number): readonly [number, number, number] {
	const chunkRow = row >> CHUNK_BITS
	const chunkCol = col >> CHUNK_BITS
	const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
	return [chunkRow, chunkCol, localIndex] as const
}

function cloneCellValue(value: CellValue): CellValue {
	if (value.kind === 'array') return structuredClone(value)
	if (value.kind === 'richText') return { kind: 'richText', runs: [...value.runs] }
	return value
}

function compactResolvedCell(
	value: CellValue,
	formula: string | null,
	styleId: StyleId,
	formulaInfo: CellFormulaBinding | undefined,
	stringTable: StringTable,
): StoredSlot {
	switch (value.kind) {
		case 'empty':
			return {
				tag: SlotTag.Empty,
				styleId,
				formula,
				formulaInfo,
				numberValue: undefined,
				stringId: undefined,
				heapValue: undefined,
			}
		case 'number':
			return {
				tag: SlotTag.Number,
				styleId,
				formula,
				formulaInfo,
				numberValue: value.value,
				stringId: undefined,
				heapValue: undefined,
			}
		case 'string':
			return {
				tag: SlotTag.String,
				styleId,
				formula,
				formulaInfo,
				numberValue: undefined,
				stringId: stringTable.intern(value.value),
				heapValue: undefined,
			}
		case 'boolean':
			return {
				tag: SlotTag.Boolean,
				styleId,
				formula,
				formulaInfo,
				numberValue: value.value ? 1 : 0,
				stringId: undefined,
				heapValue: undefined,
			}
		case 'error':
			return {
				tag: SlotTag.Error,
				styleId,
				formula,
				formulaInfo,
				numberValue: undefined,
				stringId: stringTable.intern(value.value),
				heapValue: undefined,
			}
		case 'date':
			return {
				tag: SlotTag.Date,
				styleId,
				formula,
				formulaInfo,
				numberValue: value.serial,
				stringId: undefined,
				heapValue: undefined,
			}
		default:
			return {
				tag: SlotTag.Heap,
				styleId,
				formula,
				formulaInfo,
				numberValue: undefined,
				stringId: undefined,
				heapValue: value,
			}
	}
}

function materializeCell(slot: StoredSlot | undefined, stringTable: StringTable): Cell | undefined {
	if (!slot) return undefined
	return {
		value: readSlotValue(slot, stringTable) ?? EMPTY,
		formula: slot.formula,
		styleId: slot.styleId,
		formulaInfo: slot.formulaInfo,
	}
}

function readSlotValue(
	slot: StoredSlot | undefined,
	stringTable: StringTable,
): CellValue | undefined {
	if (!slot) return undefined
	switch (slot.tag) {
		case SlotTag.Empty:
			return EMPTY
		case SlotTag.Number:
			return numberValue(slot.numberValue ?? 0)
		case SlotTag.String:
			return stringValue(stringTable.lookup(slot.stringId ?? 0))
		case SlotTag.Boolean:
			return booleanValue((slot.numberValue ?? 0) !== 0)
		case SlotTag.Error:
			return errorValue(stringTable.lookup(slot.stringId ?? 0) as ExcelError)
		case SlotTag.Date:
			return { kind: 'date', serial: slot.numberValue ?? 0 }
		case SlotTag.Heap:
			return slot.heapValue
	}
}

function readDenseValue(
	tag: SlotTag,
	number: number,
	stringId: number,
	heapValue: CellValue | undefined,
	stringTable: StringTable,
): CellValue {
	switch (tag) {
		case SlotTag.Empty:
			return EMPTY
		case SlotTag.Number:
			return numberValue(number)
		case SlotTag.String:
			return stringValue(stringTable.lookup(stringId))
		case SlotTag.Boolean:
			return booleanValue(number !== 0)
		case SlotTag.Error:
			return errorValue(stringTable.lookup(stringId) as ExcelError)
		case SlotTag.Date:
			return { kind: 'date', serial: number }
		case SlotTag.Heap:
			return heapValue ?? EMPTY
	}
}

function kindFromSlot(slot: StoredSlot | undefined): CellValueKind | undefined {
	if (!slot) return undefined
	return slot.tag === SlotTag.Heap ? (slot.heapValue?.kind ?? 'array') : kindFromTag(slot.tag)
}

function kindFromTag(tag: SlotTag): CellValueKind {
	switch (tag) {
		case SlotTag.Empty:
			return 'empty'
		case SlotTag.Number:
			return 'number'
		case SlotTag.String:
			return 'string'
		case SlotTag.Boolean:
			return 'boolean'
		case SlotTag.Error:
			return 'error'
		case SlotTag.Date:
			return 'date'
		case SlotTag.Heap: {
			return 'array'
		}
	}
}
