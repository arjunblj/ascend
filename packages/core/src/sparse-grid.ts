import type { CellValue, ExcelError } from '@ascend/schema'
import {
	booleanValue,
	dateValue,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
} from '@ascend/schema'
import { DEFAULT_STYLE_ID, type StyleId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface Cell {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
	readonly formulaInfo?: CellFormulaBinding | undefined
}

export interface SparseGridStorageStats {
	readonly cellCount: number
	readonly chunkCount: number
	readonly denseChunkCount: number
	readonly sparseChunkCount: number
	readonly denseCellCount: number
	readonly sparseCellCount: number
	readonly denseCapacity: number
	readonly sparseCapacity: number
	readonly denseArrayBufferBytes: number
	readonly sparseArrayBufferBytes: number
	readonly styleArrayBufferBytes: number
	readonly totalArrayBufferBytes: number
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

export interface DataTableFormulaInfo {
	readonly kind: 'dataTable'
	readonly ref?: string
	readonly dt2D?: boolean
	readonly dtr?: boolean
	readonly r1?: string
	readonly r2?: string
	readonly del1?: boolean
	readonly del2?: boolean
}

export interface SpillFormulaInfo {
	readonly kind: 'spill'
	readonly anchorRef: string
	readonly ref: string
	readonly isAnchor: boolean
}

export interface BlockedSpillFormulaInfo {
	readonly kind: 'blockedSpill'
	readonly anchorRef: string
	readonly ref: string
	readonly reason?: 'occupied-cell' | 'sheet-edge'
	readonly blockingRefs: readonly string[]
}

export type CellFormulaBinding =
	| SharedFormulaInfo
	| ArrayFormulaInfo
	| DynamicArrayFormulaInfo
	| DataTableFormulaInfo
	| SpillFormulaInfo
	| BlockedSpillFormulaInfo

const CHUNK_BITS = resolveChunkBits()
const CHUNK_SIZE = 1 << CHUNK_BITS
export const SPARSE_GRID_CHUNK_SIZE = CHUNK_SIZE
const CHUNK_MASK = CHUNK_SIZE - 1
const CHUNK_AREA = CHUNK_SIZE * CHUNK_SIZE
const CHUNK_COL_FACTOR = (16_383 >> CHUNK_BITS) + 1
/** SparseChunk upgrades to DenseChunk when cell count reaches this. DenseChunk uses fixed arrays; SparseChunk uses a Map. */
export const SPARSE_TO_DENSE_THRESHOLD = Math.min(384, CHUNK_AREA)

const AUTO_DENSE_SAMPLE_CHUNKS = 4
const AUTO_DENSE_FILL_RATIO = 0.5
const SLOT_TAG_MASK = 0b111
const SLOT_OCCUPIED_BIT = 0b1000

type ChunkBuffer = ArrayBuffer

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
type SharedStringCellKind = 'string' | 'richText'

interface StoredSlot {
	readonly tag: SlotTag
	readonly styleId: StyleId
	readonly formula: string | null
	readonly formulaInfo: CellFormulaBinding | undefined
	readonly numberValue: number | undefined
	readonly stringId: number | undefined
	readonly heapValue: CellValue | undefined
}

const PLAIN_WRITE_EXISTED = 1 << 0
const PLAIN_WRITE_HAD_FORMULA = 1 << 1
const PLAIN_WRITE_HAD_FORMULA_INFO = 1 << 2
const PLAIN_WRITE_PREVIOUS_STRING = 1 << 3
const PLAIN_WRITE_PREVIOUS_RICH_TEXT = 1 << 4
const PLAIN_WRITE_PREVIOUS_ARRAY = 1 << 5

interface PlainNumberSpanWriteResult {
	inserted: number
	formulaCleared: number
	formulaInfoCleared: number
	previousString: number
	previousRichText: number
	previousArray: number
}

interface GridChunk {
	readonly count: number
	clone(): GridChunk
	storageStats(): ChunkStorageStats
	has(localIndex: number): boolean
	getSlot(localIndex: number): StoredSlot | undefined
	getKind(localIndex: number): CellValueKind | undefined
	readValue(localIndex: number, stringTable: StringTable): CellValue | undefined
	readNumber(localIndex: number): number | null
	readString(localIndex: number, stringTable: StringTable): string | null
	readError(localIndex: number, stringTable: StringTable): ExcelError | null
	readStyleId(localIndex: number): StyleId | undefined
	readFormula(localIndex: number): string | null | undefined
	readFormulaInfo(localIndex: number): CellFormulaBinding | undefined
	clearFormulaInfo(localIndex: number): void
	countFormulaCells(): number
	countFormulaInfoCells(): number
	countStringCells(): number
	countRichTextCells(): number
	countArrayCells(): number
	setStringResolved(
		localIndex: number,
		value: string,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
		stringTable: StringTable,
	): GridChunk
	setNumberResolved(
		localIndex: number,
		value: number,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
	): GridChunk
	setSlot(localIndex: number, slot: StoredSlot): GridChunk
	delete(localIndex: number): boolean
	forEachRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		fn: (localCol: number, slot: StoredSlot) => void,
	): void
	forEachValueInRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		stringTable: StringTable,
		fn: (localCol: number, value: CellValue) => void,
	): void
	forEachValueInRangeUnordered(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		stringTable: StringTable,
		fn: (value: CellValue, row: number, col: number) => void,
	): void
	forEachOccupiedInRangeUnordered(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		fn: (row: number, col: number) => void,
	): void
	aggregateNumericInRange(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		stringTable: StringTable,
		state: NumericRangeAggregateState,
	): void
}

interface ChunkStorageStats {
	readonly kind: 'dense' | 'sparse'
	readonly cellCount: number
	readonly capacity: number
	readonly arrayBufferBytes: number
	readonly styleArrayBufferBytes: number
}

export interface NumericRangeAggregate {
	readonly sum: number
	readonly count: number
	readonly min: number
	readonly max: number
	readonly error: ExcelError | null
}

interface NumericRangeAggregateState {
	sum: number
	count: number
	min: number
	max: number
	error: ExcelError | null
	errorRow: number
	errorCol: number
}

class StringTable {
	private readonly ids = new Map<string, number>()
	private readonly values = ['']
	private reverseLookupEnabled = true
	private lookupCount = 0
	private missCount = 0

	constructor() {
		for (const value of COMMON_INTERNED_STRINGS) {
			this.ids.set(value, this.values.length)
			this.values.push(value)
		}
	}

	intern(value: string): number {
		if (!this.reverseLookupEnabled || this.values.length > MAX_REVERSE_STRING_INTERN_IDS) {
			if (this.reverseLookupEnabled) this.disableReverseLookup()
			const id = this.values.length
			this.values.push(value)
			return id
		}
		this.lookupCount++
		const existing = this.ids.get(value)
		if (existing !== undefined) return existing
		this.missCount++
		if (
			this.values.length >= ADAPTIVE_REVERSE_STRING_INTERN_MIN_IDS &&
			this.missCount / this.lookupCount >= ADAPTIVE_REVERSE_STRING_INTERN_MISS_RATIO
		) {
			this.disableReverseLookup()
			const id = this.values.length
			this.values.push(value)
			return id
		}
		const id = this.values.length
		this.values.push(value)
		this.ids.set(value, id)
		return id
	}

	lookup(id: number): string {
		return this.values[id] ?? ''
	}

	private disableReverseLookup(): void {
		this.ids.clear()
		this.reverseLookupEnabled = false
	}
}

const MAX_REVERSE_STRING_INTERN_IDS = 65_536
const ADAPTIVE_REVERSE_STRING_INTERN_MIN_IDS = 1_024
const ADAPTIVE_REVERSE_STRING_INTERN_MISS_RATIO = 0.98

const COMMON_INTERNED_STRINGS = [
	'TRUE',
	'FALSE',
	'#NULL!',
	'#DIV/0!',
	'#VALUE!',
	'#REF!',
	'#NAME?',
	'#NUM!',
	'#N/A',
	'#SPILL!',
	'#CALC!',
] as const

class SparseChunk implements GridChunk {
	private readonly slots = new Map<number, StoredSlot>()
	private readonly rowCounts = new Uint16Array(CHUNK_SIZE)
	private readonly rowMaskLo = new Uint32Array(CHUNK_SIZE)
	private readonly rowMaskHi = new Uint32Array(CHUNK_SIZE)
	private readonly rowMinCol = new Int16Array(CHUNK_SIZE)
	private readonly rowMaxCol = new Int16Array(CHUNK_SIZE)

	constructor() {
		this.rowMinCol.fill(ROW_EMPTY_MIN)
		this.rowMaxCol.fill(ROW_EMPTY_MAX)
	}

	get count(): number {
		return this.slots.size
	}

	clone(): GridChunk {
		const next = new SparseChunk()
		for (const [index, slot] of this.slots) {
			next.slots.set(index, slot)
		}
		next.rowCounts.set(this.rowCounts)
		next.rowMaskLo.set(this.rowMaskLo)
		next.rowMaskHi.set(this.rowMaskHi)
		next.rowMinCol.set(this.rowMinCol)
		next.rowMaxCol.set(this.rowMaxCol)
		return next
	}

	storageStats(): ChunkStorageStats {
		return {
			kind: 'sparse',
			cellCount: this.slots.size,
			capacity: CHUNK_AREA,
			arrayBufferBytes:
				this.rowCounts.byteLength +
				this.rowMaskLo.byteLength +
				this.rowMaskHi.byteLength +
				this.rowMinCol.byteLength +
				this.rowMaxCol.byteLength,
			styleArrayBufferBytes: 0,
		}
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

	readStyleId(localIndex: number): StyleId | undefined {
		return this.slots.get(localIndex)?.styleId
	}

	readFormula(localIndex: number): string | null | undefined {
		return this.slots.get(localIndex)?.formula
	}

	readFormulaInfo(localIndex: number): CellFormulaBinding | undefined {
		return this.slots.get(localIndex)?.formulaInfo
	}

	clearFormulaInfo(localIndex: number): void {
		const slot = this.slots.get(localIndex)
		if (!slot || !slot.formulaInfo) return
		this.slots.set(localIndex, {
			tag: slot.tag,
			styleId: slot.styleId,
			formula: slot.formula,
			formulaInfo: undefined,
			numberValue: slot.numberValue,
			stringId: slot.stringId,
			heapValue: slot.heapValue,
		})
	}

	countFormulaCells(): number {
		let count = 0
		for (const slot of this.slots.values()) {
			if (slot.formula !== null || slot.formulaInfo !== undefined) count++
		}
		return count
	}

	countFormulaInfoCells(): number {
		let count = 0
		for (const slot of this.slots.values()) {
			if (slot.formulaInfo !== undefined) count++
		}
		return count
	}

	countStringCells(): number {
		let count = 0
		for (const slot of this.slots.values()) {
			if (slot.tag === SlotTag.String) count++
		}
		return count
	}

	countRichTextCells(): number {
		let count = 0
		for (const slot of this.slots.values()) {
			if (slot.heapValue?.kind === 'richText') count++
		}
		return count
	}

	countArrayCells(): number {
		let count = 0
		for (const slot of this.slots.values()) {
			if (slot.heapValue?.kind === 'array') count++
		}
		return count
	}

	setStringResolved(
		localIndex: number,
		value: string,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
		stringTable: StringTable,
	): GridChunk {
		return this.setSlot(
			localIndex,
			compactStringResolvedCell(value, formula, styleId, formulaInfo, stringTable),
		)
	}

	setNumberResolved(
		localIndex: number,
		value: number,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
	): GridChunk {
		return this.setSlot(localIndex, compactNumberResolvedCell(value, formula, styleId, formulaInfo))
	}

	setSlot(localIndex: number, slot: StoredSlot): GridChunk {
		if (!this.slots.has(localIndex)) {
			const rowIndex = localIndex >>> CHUNK_BITS
			const localCol = localIndex & CHUNK_MASK
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this.setRowMaskBit(rowIndex, localCol)
			this.updateRowBounds(rowIndex, localCol, true)
		}
		this.slots.set(localIndex, slot)
		if (this.slots.size < SPARSE_TO_DENSE_THRESHOLD) return this
		const dense = new DenseChunk()
		for (const [index, existing] of this.slots) {
			dense.setSlot(index, existing)
		}
		return dense
	}

	delete(localIndex: number): boolean {
		if (!this.slots.delete(localIndex)) return false
		const rowIndex = localIndex >>> CHUNK_BITS
		const localCol = localIndex & CHUNK_MASK
		this.rowCounts[rowIndex] = Math.max(0, (this.rowCounts[rowIndex] ?? 0) - 1)
		this.clearRowMaskBit(rowIndex, localCol)
		this.updateRowBounds(rowIndex, localCol, false)
		return true
	}

	private setRowMaskBit(rowIndex: number, localCol: number): void {
		if (localCol < 32) {
			this.rowMaskLo[rowIndex] = ((this.rowMaskLo[rowIndex] ?? 0) | ((1 << localCol) >>> 0)) >>> 0
		} else {
			const bit = localCol - 32
			this.rowMaskHi[rowIndex] = ((this.rowMaskHi[rowIndex] ?? 0) | ((1 << bit) >>> 0)) >>> 0
		}
	}

	private clearRowMaskBit(rowIndex: number, localCol: number): void {
		if (localCol < 32) {
			this.rowMaskLo[rowIndex] = ((this.rowMaskLo[rowIndex] ?? 0) & ~((1 << localCol) >>> 0)) >>> 0
		} else {
			const bit = localCol - 32
			this.rowMaskHi[rowIndex] = ((this.rowMaskHi[rowIndex] ?? 0) & ~((1 << bit) >>> 0)) >>> 0
		}
	}

	private updateRowBounds(rowIndex: number, localCol: number, isAdd: boolean): void {
		if (isAdd) {
			const count = this.rowCounts[rowIndex] ?? 0
			if (count === 1) {
				this.rowMinCol[rowIndex] = localCol
				this.rowMaxCol[rowIndex] = localCol
			} else {
				this.rowMinCol[rowIndex] = Math.min(this.rowMinCol[rowIndex] ?? ROW_EMPTY_MIN, localCol)
				this.rowMaxCol[rowIndex] = Math.max(this.rowMaxCol[rowIndex] ?? ROW_EMPTY_MAX, localCol)
			}
			return
		}
		const count = this.rowCounts[rowIndex] ?? 0
		if (count === 0) {
			this.rowMinCol[rowIndex] = ROW_EMPTY_MIN
			this.rowMaxCol[rowIndex] = ROW_EMPTY_MAX
		} else if (
			localCol === (this.rowMinCol[rowIndex] ?? ROW_EMPTY_MIN) ||
			localCol === (this.rowMaxCol[rowIndex] ?? ROW_EMPTY_MAX)
		) {
			this.recomputeRowBounds(rowIndex)
		}
	}

	private recomputeRowBounds(rowIndex: number): void {
		const lo = this.rowMaskLo[rowIndex] ?? 0
		const hi = this.rowMaskHi[rowIndex] ?? 0
		if (lo !== 0) {
			this.rowMinCol[rowIndex] = trailingZeroCount32(lo)
			this.rowMaxCol[rowIndex] = hi !== 0 ? 32 + (31 - Math.clz32(hi)) : 31 - Math.clz32(lo)
			return
		}
		if (hi !== 0) {
			this.rowMinCol[rowIndex] = 32 + trailingZeroCount32(hi)
			this.rowMaxCol[rowIndex] = 32 + (31 - Math.clz32(hi))
			return
		}
		this.rowMinCol[rowIndex] = ROW_EMPTY_MIN
		this.rowMaxCol[rowIndex] = ROW_EMPTY_MAX
	}

	forEachRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		fn: (localCol: number, slot: StoredSlot) => void,
	): void {
		if ((this.rowCounts[localRow] ?? 0) === 0) return
		const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
		const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
		if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) return
		const startCol = Math.max(minLocalCol, rowMin)
		const endCol = Math.min(maxLocalCol, rowMax)
		const rowBase = localRow << CHUNK_BITS
		this.forEachOccupiedLocalCol(localRow, startCol, endCol, (localCol) => {
			const slot = this.slots.get(rowBase | localCol)
			if (slot) fn(localCol, slot)
		})
	}

	forEachValueInRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		stringTable: StringTable,
		fn: (localCol: number, value: CellValue) => void,
	): void {
		if ((this.rowCounts[localRow] ?? 0) === 0) return
		const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
		const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
		if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) return
		const startCol = Math.max(minLocalCol, rowMin)
		const endCol = Math.min(maxLocalCol, rowMax)
		const rowBase = localRow << CHUNK_BITS
		this.forEachOccupiedLocalCol(localRow, startCol, endCol, (localCol) => {
			const slot = this.slots.get(rowBase | localCol)
			if (slot) {
				const value = readSlotValue(slot, stringTable)
				if (value !== undefined) fn(localCol, value)
			}
		})
	}

	forEachValueInRangeUnordered(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		stringTable: StringTable,
		fn: (value: CellValue, row: number, col: number) => void,
	): void {
		for (const [localIndex, slot] of this.slots) {
			const localRow = localIndex >>> CHUNK_BITS
			if (localRow < localRowStart || localRow > localRowEnd) continue
			const localCol = localIndex & CHUNK_MASK
			if (localCol < minLocalCol || localCol > maxLocalCol) continue
			const value = readSlotValue(slot, stringTable)
			if (value !== undefined) fn(value, baseRow + localRow, baseCol + localCol)
		}
	}

	forEachOccupiedInRangeUnordered(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		fn: (row: number, col: number) => void,
	): void {
		for (const localIndex of this.slots.keys()) {
			const localRow = localIndex >>> CHUNK_BITS
			if (localRow < localRowStart || localRow > localRowEnd) continue
			const localCol = localIndex & CHUNK_MASK
			if (localCol < minLocalCol || localCol > maxLocalCol) continue
			fn(baseRow + localRow, baseCol + localCol)
		}
	}

	aggregateNumericInRange(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		stringTable: StringTable,
		state: NumericRangeAggregateState,
	): void {
		for (const [localIndex, slot] of this.slots) {
			const localRow = localIndex >>> CHUNK_BITS
			if (localRow < localRowStart || localRow > localRowEnd) continue
			const localCol = localIndex & CHUNK_MASK
			if (localCol < minLocalCol || localCol > maxLocalCol) continue
			aggregateNumericSlot(slot, stringTable, state, baseRow + localRow, baseCol + localCol)
		}
	}

	private forEachOccupiedLocalCol(
		localRow: number,
		startCol: number,
		endCol: number,
		fn: (localCol: number) => void,
	): void {
		const loEnd = Math.min(endCol, 31)
		if (startCol <= loEnd) {
			let mask = ((this.rowMaskLo[localRow] ?? 0) & bitRangeMask32(startCol, loEnd)) >>> 0
			while (mask !== 0) {
				const localCol = trailingZeroCount32(mask)
				mask = (mask & (mask - 1)) >>> 0
				fn(localCol)
			}
		}
		if (endCol < 32) return
		const hiStart = Math.max(startCol - 32, 0)
		const hiEnd = Math.min(endCol - 32, 31)
		let mask = ((this.rowMaskHi[localRow] ?? 0) & bitRangeMask32(hiStart, hiEnd)) >>> 0
		while (mask !== 0) {
			const localCol = 32 + trailingZeroCount32(mask)
			mask = (mask & (mask - 1)) >>> 0
			fn(localCol)
		}
	}
}

const ROW_EMPTY_MIN = CHUNK_SIZE
const ROW_EMPTY_MAX = -1

function bitRangeMask32(startBit: number, endBit: number): number {
	const lower = startBit <= 0 ? 0xffffffff : (0xffffffff << startBit) >>> 0
	const upper = endBit >= 31 ? 0xffffffff : 0xffffffff >>> (31 - endBit)
	return (lower & upper) >>> 0
}

function trailingZeroCount32(value: number): number {
	return 31 - Math.clz32(value & -value)
}

class DenseChunk implements GridChunk {
	private readonly metaBuffer = createChunkBuffer(CHUNK_AREA * 5)
	private readonly numberBuffer = createChunkBuffer(CHUNK_AREA * 8)
	private readonly slotMeta = new Uint8Array(this.metaBuffer, 0, CHUNK_AREA)
	private readonly stringIds = new Uint32Array(this.metaBuffer, CHUNK_AREA, CHUNK_AREA)
	private readonly numbers = new Float64Array(this.numberBuffer)
	private readonly rowCounts = new Uint16Array(CHUNK_SIZE)
	private readonly rowMinCol = new Int8Array(CHUNK_SIZE)
	private readonly rowMaxCol = new Int8Array(CHUNK_SIZE)
	private styleIds: Uint32Array | null = null
	private formulas: Array<string | null | undefined> | null = null
	private formulaInfos: Array<CellFormulaBinding | undefined> | null = null
	private heapValues: Array<CellValue | undefined> | null = null
	private _count = 0
	private _reusableSlot: StoredSlot = {
		tag: SlotTag.Empty,
		styleId: 0 as StyleId,
		formula: null,
		formulaInfo: undefined,
		numberValue: undefined,
		stringId: undefined,
		heapValue: undefined,
	}

	get count(): number {
		return this._count
	}

	clone(): GridChunk {
		const next = new DenseChunk()
		new Uint8Array(next.metaBuffer).set(new Uint8Array(this.metaBuffer))
		new Uint8Array(next.numberBuffer).set(new Uint8Array(this.numberBuffer))
		next.rowCounts.set(this.rowCounts)
		next.rowMinCol.set(this.rowMinCol)
		next.rowMaxCol.set(this.rowMaxCol)
		next.styleIds = this.styleIds ? new Uint32Array(this.styleIds) : null
		next._count = this._count
		next.formulas = this.formulas ? this.formulas.slice() : null
		next.formulaInfos = this.formulaInfos ? this.formulaInfos.slice() : null
		next.heapValues = this.heapValues ? this.heapValues.slice() : null
		return next
	}

	storageStats(): ChunkStorageStats {
		const styleArrayBufferBytes = this.styleIds?.byteLength ?? 0
		return {
			kind: 'dense',
			cellCount: this._count,
			capacity: CHUNK_AREA,
			arrayBufferBytes:
				this.metaBuffer.byteLength +
				this.numberBuffer.byteLength +
				this.rowCounts.byteLength +
				this.rowMinCol.byteLength +
				this.rowMaxCol.byteLength +
				styleArrayBufferBytes,
			styleArrayBufferBytes,
		}
	}

	has(localIndex: number): boolean {
		return ((this.slotMeta[localIndex] ?? 0) & SLOT_OCCUPIED_BIT) !== 0
	}

	getSlot(localIndex: number): StoredSlot | undefined {
		if (!this.has(localIndex)) return undefined
		const tag = this.readTag(localIndex)
		const slot = this._reusableSlot as unknown as Record<string, unknown>
		slot.tag = tag
		slot.styleId = (this.styleIds?.[localIndex] ?? DEFAULT_STYLE_ID) as StyleId
		slot.formula = this.formulas?.[localIndex] ?? null
		slot.formulaInfo = this.formulaInfos?.[localIndex]
		slot.numberValue =
			tag === SlotTag.Number || tag === SlotTag.Boolean || tag === SlotTag.Date
				? (this.numbers[localIndex] ?? 0)
				: undefined
		slot.stringId =
			tag === SlotTag.String || tag === SlotTag.Error
				? (this.stringIds[localIndex] ?? 0)
				: undefined
		slot.heapValue = tag === SlotTag.Heap ? (this.heapValues?.[localIndex] ?? EMPTY) : undefined
		return this._reusableSlot
	}

	getKind(localIndex: number): CellValueKind | undefined {
		if (!this.has(localIndex)) return undefined
		const tag = this.readTag(localIndex)
		return tag === SlotTag.Heap
			? (this.heapValues?.[localIndex]?.kind ?? 'array')
			: kindFromTag(tag)
	}

	readValue(localIndex: number, stringTable: StringTable): CellValue | undefined {
		if (!this.has(localIndex)) return undefined
		return readDenseValue(
			this.readTag(localIndex),
			this.numbers[localIndex] ?? 0,
			this.stringIds[localIndex] ?? 0,
			this.heapValues?.[localIndex],
			stringTable,
		)
	}

	readNumber(localIndex: number): number | null {
		if (!this.has(localIndex)) return null
		const tag = this.readTag(localIndex)
		return tag === SlotTag.Number || tag === SlotTag.Boolean || tag === SlotTag.Date
			? (this.numbers[localIndex] ?? 0)
			: null
	}

	readString(localIndex: number, stringTable: StringTable): string | null {
		if (!this.has(localIndex) || this.readTag(localIndex) !== SlotTag.String) return null
		return stringTable.lookup(this.stringIds[localIndex] ?? 0)
	}

	readError(localIndex: number, stringTable: StringTable): ExcelError | null {
		if (!this.has(localIndex) || this.readTag(localIndex) !== SlotTag.Error) return null
		return stringTable.lookup(this.stringIds[localIndex] ?? 0) as ExcelError
	}

	readStyleId(localIndex: number): StyleId | undefined {
		return this.has(localIndex)
			? ((this.styleIds?.[localIndex] ?? DEFAULT_STYLE_ID) as StyleId)
			: undefined
	}

	readFormula(localIndex: number): string | null | undefined {
		if (!this.has(localIndex)) return undefined
		return this.formulas?.[localIndex] ?? null
	}

	readFormulaInfo(localIndex: number): CellFormulaBinding | undefined {
		if (!this.has(localIndex)) return undefined
		return this.formulaInfos?.[localIndex]
	}

	clearFormulaInfo(localIndex: number): void {
		if (this.formulaInfos !== null) this.formulaInfos[localIndex] = undefined
	}

	countFormulaCells(): number {
		if (this.formulas === null && this.formulaInfos === null) return 0
		let count = 0
		for (let localIndex = 0; localIndex < CHUNK_AREA; localIndex++) {
			if (!this.has(localIndex)) continue
			if (
				(this.formulas?.[localIndex] ?? null) !== null ||
				this.formulaInfos?.[localIndex] !== undefined
			) {
				count++
			}
		}
		return count
	}

	countFormulaInfoCells(): number {
		if (this.formulaInfos === null) return 0
		let count = 0
		for (let localIndex = 0; localIndex < CHUNK_AREA; localIndex++) {
			if (this.has(localIndex) && this.formulaInfos[localIndex] !== undefined) count++
		}
		return count
	}

	countStringCells(): number {
		let count = 0
		for (let localIndex = 0; localIndex < CHUNK_AREA; localIndex++) {
			if (this.has(localIndex) && this.readTag(localIndex) === SlotTag.String) count++
		}
		return count
	}

	countRichTextCells(): number {
		if (this.heapValues === null) return 0
		let count = 0
		for (let localIndex = 0; localIndex < CHUNK_AREA; localIndex++) {
			if (this.has(localIndex) && this.heapValues[localIndex]?.kind === 'richText') count++
		}
		return count
	}

	countArrayCells(): number {
		if (this.heapValues === null) return 0
		let count = 0
		for (let localIndex = 0; localIndex < CHUNK_AREA; localIndex++) {
			if (this.has(localIndex) && this.heapValues[localIndex]?.kind === 'array') count++
		}
		return count
	}

	setResolved(
		localIndex: number,
		value: CellValue,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
		stringTable: StringTable,
	): void {
		if (!this.has(localIndex)) {
			const rowIndex = localIndex >>> CHUNK_BITS
			const localCol = localIndex & CHUNK_MASK
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this._count++
			this.updateRowBounds(rowIndex, localCol, true)
		}
		switch (value.kind) {
			case 'empty':
				this.writeResolved(localIndex, SlotTag.Empty, styleId, 0, 0, formula, formulaInfo)
				return
			case 'number':
				this.writeResolved(
					localIndex,
					SlotTag.Number,
					styleId,
					value.value,
					0,
					formula,
					formulaInfo,
				)
				return
			case 'string':
				this.writeResolved(
					localIndex,
					SlotTag.String,
					styleId,
					0,
					stringTable.intern(value.value),
					formula,
					formulaInfo,
				)
				return
			case 'boolean':
				this.writeResolved(
					localIndex,
					SlotTag.Boolean,
					styleId,
					value.value ? 1 : 0,
					0,
					formula,
					formulaInfo,
				)
				return
			case 'error':
				this.writeResolved(
					localIndex,
					SlotTag.Error,
					styleId,
					0,
					stringTable.intern(value.value),
					formula,
					formulaInfo,
				)
				return
			case 'date':
				this.writeResolved(localIndex, SlotTag.Date, styleId, value.serial, 0, formula, formulaInfo)
				return
			default:
				this.writeResolved(localIndex, SlotTag.Heap, styleId, 0, 0, formula, formulaInfo, value)
				return
		}
	}

	setStringResolved(
		localIndex: number,
		value: string,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
		stringTable: StringTable,
	): GridChunk {
		if (!this.has(localIndex)) {
			const rowIndex = localIndex >>> CHUNK_BITS
			const localCol = localIndex & CHUNK_MASK
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this._count++
			this.updateRowBounds(rowIndex, localCol, true)
		}
		this.writeResolved(
			localIndex,
			SlotTag.String,
			styleId,
			0,
			stringTable.intern(value),
			formula,
			formulaInfo,
		)
		return this
	}

	setNumberResolved(
		localIndex: number,
		value: number,
		formula: string | null,
		styleId: StyleId,
		formulaInfo: CellFormulaBinding | undefined,
	): GridChunk {
		if (!this.has(localIndex)) {
			const rowIndex = localIndex >>> CHUNK_BITS
			const localCol = localIndex & CHUNK_MASK
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this._count++
			this.updateRowBounds(rowIndex, localCol, true)
		}
		this.writeResolved(localIndex, SlotTag.Number, styleId, value, 0, formula, formulaInfo)
		return this
	}

	writePlainNumber(localIndex: number, value: number): number {
		const previous = this.preparePlainWrite(localIndex)
		this.slotMeta[localIndex] = SLOT_OCCUPIED_BIT | SlotTag.Number
		this.numbers[localIndex] = value
		if (this.styleIds !== null) this.styleIds[localIndex] = DEFAULT_STYLE_ID
		if (this.formulas !== null) this.formulas[localIndex] = null
		if (this.formulaInfos !== null) this.formulaInfos[localIndex] = undefined
		if (this.heapValues !== null) this.heapValues[localIndex] = undefined
		return previous
	}

	writePlainNumberSpan(
		localRow: number,
		startLocalCol: number,
		values: readonly number[],
		valueOffset: number,
		count: number,
	): PlainNumberSpanWriteResult {
		const endLocalCol = startLocalCol + count - 1
		const rowBase = localRow << CHUNK_BITS
		const existingRowCount = this.rowCounts[localRow] ?? 0
		if (existingRowCount === 0) {
			for (let offset = 0; offset < count; offset++) {
				const localIndex = rowBase | (startLocalCol + offset)
				this.slotMeta[localIndex] = SLOT_OCCUPIED_BIT | SlotTag.Number
				this.numbers[localIndex] = values[valueOffset + offset] ?? 0
				if (this.styleIds !== null) this.styleIds[localIndex] = DEFAULT_STYLE_ID
				if (this.formulas !== null) this.formulas[localIndex] = null
				if (this.formulaInfos !== null) this.formulaInfos[localIndex] = undefined
				if (this.heapValues !== null) this.heapValues[localIndex] = undefined
			}
			this.rowCounts[localRow] = count
			this.rowMinCol[localRow] = startLocalCol
			this.rowMaxCol[localRow] = endLocalCol
			this._count += count
			return {
				inserted: count,
				formulaCleared: 0,
				formulaInfoCleared: 0,
				previousString: 0,
				previousRichText: 0,
				previousArray: 0,
			}
		}

		const result: PlainNumberSpanWriteResult = {
			inserted: 0,
			formulaCleared: 0,
			formulaInfoCleared: 0,
			previousString: 0,
			previousRichText: 0,
			previousArray: 0,
		}
		for (let offset = 0; offset < count; offset++) {
			const flags = this.writePlainNumber(
				rowBase | (startLocalCol + offset),
				values[valueOffset + offset] ?? 0,
			)
			if ((flags & PLAIN_WRITE_EXISTED) === 0) result.inserted++
			if ((flags & PLAIN_WRITE_HAD_FORMULA) !== 0) result.formulaCleared++
			if ((flags & PLAIN_WRITE_HAD_FORMULA_INFO) !== 0) result.formulaInfoCleared++
			if ((flags & PLAIN_WRITE_PREVIOUS_STRING) !== 0) result.previousString++
			if ((flags & PLAIN_WRITE_PREVIOUS_RICH_TEXT) !== 0) result.previousRichText++
			if ((flags & PLAIN_WRITE_PREVIOUS_ARRAY) !== 0) result.previousArray++
		}
		return result
	}

	writePlainString(localIndex: number, value: string, stringTable: StringTable): number {
		const previous = this.preparePlainWrite(localIndex)
		this.writeResolved(
			localIndex,
			SlotTag.String,
			DEFAULT_STYLE_ID,
			0,
			stringTable.intern(value),
			null,
			undefined,
		)
		return previous
	}

	setSlot(localIndex: number, slot: StoredSlot): GridChunk {
		if (!this.has(localIndex)) {
			const rowIndex = localIndex >>> CHUNK_BITS
			const localCol = localIndex & CHUNK_MASK
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this._count++
			this.updateRowBounds(rowIndex, localCol, true)
		}
		this.writeResolved(
			localIndex,
			slot.tag,
			slot.styleId,
			slot.numberValue ?? 0,
			slot.stringId ?? 0,
			slot.formula,
			slot.formulaInfo,
			slot.heapValue,
		)
		return this
	}

	private preparePlainWrite(localIndex: number): number {
		const existed = this.has(localIndex)
		if (!existed) {
			const rowIndex = localIndex >>> CHUNK_BITS
			const localCol = localIndex & CHUNK_MASK
			this.rowCounts[rowIndex] = (this.rowCounts[rowIndex] ?? 0) + 1
			this._count++
			this.updateRowBounds(rowIndex, localCol, true)
			return 0
		}
		const tag = this.readTag(localIndex)
		const formulaInfo = this.formulaInfos?.[localIndex]
		let flags = PLAIN_WRITE_EXISTED
		if ((this.formulas?.[localIndex] ?? null) !== null || formulaInfo !== undefined) {
			flags |= PLAIN_WRITE_HAD_FORMULA
		}
		if (formulaInfo !== undefined) flags |= PLAIN_WRITE_HAD_FORMULA_INFO
		if (tag === SlotTag.String) {
			flags |= PLAIN_WRITE_PREVIOUS_STRING
		} else if (tag === SlotTag.Heap && this.heapValues?.[localIndex]?.kind === 'richText') {
			flags |= PLAIN_WRITE_PREVIOUS_RICH_TEXT
		} else if (tag === SlotTag.Heap && this.heapValues?.[localIndex]?.kind === 'array') {
			flags |= PLAIN_WRITE_PREVIOUS_ARRAY
		}
		return flags
	}

	delete(localIndex: number): boolean {
		if (!this.has(localIndex)) return false
		const rowIndex = localIndex >>> CHUNK_BITS
		const localCol = localIndex & CHUNK_MASK
		this.slotMeta[localIndex] = SlotTag.Empty
		if (this.styleIds !== null) this.styleIds[localIndex] = DEFAULT_STYLE_ID
		this.stringIds[localIndex] = 0
		this.numbers[localIndex] = 0
		if (this.formulas !== null) this.formulas[localIndex] = undefined
		if (this.formulaInfos !== null) this.formulaInfos[localIndex] = undefined
		if (this.heapValues !== null) this.heapValues[localIndex] = undefined
		this.rowCounts[rowIndex] = Math.max(0, (this.rowCounts[rowIndex] ?? 0) - 1)
		this._count--
		this.updateRowBounds(rowIndex, localCol, false)
		return true
	}

	private updateRowBounds(rowIndex: number, localCol: number, isAdd: boolean): void {
		if (isAdd) {
			const count = this.rowCounts[rowIndex] ?? 0
			if (count === 1) {
				this.rowMinCol[rowIndex] = localCol
				this.rowMaxCol[rowIndex] = localCol
			} else {
				this.rowMinCol[rowIndex] = Math.min(this.rowMinCol[rowIndex] ?? ROW_EMPTY_MIN, localCol)
				this.rowMaxCol[rowIndex] = Math.max(this.rowMaxCol[rowIndex] ?? ROW_EMPTY_MAX, localCol)
			}
		} else {
			const count = this.rowCounts[rowIndex] ?? 0
			if (count === 0) {
				this.rowMinCol[rowIndex] = ROW_EMPTY_MIN
				this.rowMaxCol[rowIndex] = ROW_EMPTY_MAX
			} else if (
				localCol === (this.rowMinCol[rowIndex] ?? ROW_EMPTY_MIN) ||
				localCol === (this.rowMaxCol[rowIndex] ?? ROW_EMPTY_MAX)
			) {
				this.recomputeRowBounds(rowIndex)
			}
		}
	}

	private recomputeRowBounds(rowIndex: number): void {
		const rowBase = rowIndex << CHUNK_BITS
		let minC = ROW_EMPTY_MIN
		let maxC = ROW_EMPTY_MAX
		for (let localCol = 0; localCol <= CHUNK_MASK; localCol++) {
			if (this.has(rowBase + localCol)) {
				minC = Math.min(minC, localCol)
				maxC = Math.max(maxC, localCol)
			}
		}
		this.rowMinCol[rowIndex] = minC
		this.rowMaxCol[rowIndex] = maxC
	}

	private readTag(localIndex: number): SlotTag {
		return ((this.slotMeta[localIndex] ?? 0) & SLOT_TAG_MASK) as SlotTag
	}

	private writeResolved(
		localIndex: number,
		tag: SlotTag,
		styleId: StyleId,
		numberValue: number,
		stringId: number,
		formula: string | null,
		formulaInfo: CellFormulaBinding | undefined,
		heapValue?: CellValue,
	): void {
		this.slotMeta[localIndex] = SLOT_OCCUPIED_BIT | tag
		if (styleId !== DEFAULT_STYLE_ID || this.styleIds !== null) {
			if (this.styleIds === null) this.styleIds = new Uint32Array(CHUNK_AREA)
			this.styleIds[localIndex] = styleId
		}
		this.stringIds[localIndex] = stringId
		this.numbers[localIndex] = numberValue
		if (formula !== null) {
			if (this.formulas === null) this.formulas = new Array(CHUNK_AREA)
			this.formulas[localIndex] = formula
		} else if (this.formulas !== null) {
			this.formulas[localIndex] = null
		}
		if (formulaInfo !== undefined) {
			if (this.formulaInfos === null) this.formulaInfos = new Array(CHUNK_AREA)
			this.formulaInfos[localIndex] = formulaInfo
		} else if (this.formulaInfos !== null) {
			this.formulaInfos[localIndex] = undefined
		}
		if (tag === SlotTag.Heap) {
			if (this.heapValues === null) this.heapValues = new Array(CHUNK_AREA)
			this.heapValues[localIndex] = heapValue
		} else if (this.heapValues !== null) {
			this.heapValues[localIndex] = undefined
		}
	}

	forEachRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		fn: (localCol: number, slot: StoredSlot) => void,
	): void {
		if ((this.rowCounts[localRow] ?? 0) === 0) return
		const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
		const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
		if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) return
		const startCol = Math.max(minLocalCol, rowMin)
		const endCol = Math.min(maxLocalCol, rowMax)
		const rowBase = localRow << CHUNK_BITS
		const slot = this._reusableSlot as unknown as Record<string, unknown>
		for (let localCol = startCol; localCol <= endCol; localCol++) {
			const localIndex = rowBase + localCol
			if (!this.has(localIndex)) continue
			const tag = this.readTag(localIndex)
			slot.tag = tag
			slot.styleId = (this.styleIds?.[localIndex] ?? DEFAULT_STYLE_ID) as StyleId
			slot.formula = this.formulas?.[localIndex] ?? null
			slot.formulaInfo = this.formulaInfos?.[localIndex]
			slot.numberValue =
				tag === SlotTag.Number || tag === SlotTag.Boolean || tag === SlotTag.Date
					? (this.numbers[localIndex] ?? 0)
					: undefined
			slot.stringId =
				tag === SlotTag.String || tag === SlotTag.Error
					? (this.stringIds[localIndex] ?? 0)
					: undefined
			slot.heapValue = tag === SlotTag.Heap ? (this.heapValues?.[localIndex] ?? EMPTY) : undefined
			fn(localCol, this._reusableSlot)
		}
	}

	forEachValueInRow(
		localRow: number,
		minLocalCol: number,
		maxLocalCol: number,
		stringTable: StringTable,
		fn: (localCol: number, value: CellValue) => void,
	): void {
		if ((this.rowCounts[localRow] ?? 0) === 0) return
		const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
		const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
		if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) return
		const startCol = Math.max(minLocalCol, rowMin)
		const endCol = Math.min(maxLocalCol, rowMax)
		const rowBase = localRow << CHUNK_BITS
		for (let localCol = startCol; localCol <= endCol; localCol++) {
			const localIndex = rowBase + localCol
			if (!this.has(localIndex)) continue
			const value = readDenseValue(
				this.readTag(localIndex),
				this.numbers[localIndex] ?? 0,
				this.stringIds[localIndex] ?? 0,
				this.heapValues?.[localIndex],
				stringTable,
			)
			if (value !== undefined) fn(localCol, value)
		}
	}

	forEachValueInRangeUnordered(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		stringTable: StringTable,
		fn: (value: CellValue, row: number, col: number) => void,
	): void {
		for (let localRow = localRowStart; localRow <= localRowEnd; localRow++) {
			if ((this.rowCounts[localRow] ?? 0) === 0) continue
			const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
			const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
			if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) continue
			const startCol = Math.max(minLocalCol, rowMin)
			const endCol = Math.min(maxLocalCol, rowMax)
			const rowBase = localRow << CHUNK_BITS
			for (let localCol = startCol; localCol <= endCol; localCol++) {
				const localIndex = rowBase + localCol
				if (!this.has(localIndex)) continue
				const value = readDenseValue(
					this.readTag(localIndex),
					this.numbers[localIndex] ?? 0,
					this.stringIds[localIndex] ?? 0,
					this.heapValues?.[localIndex],
					stringTable,
				)
				if (value !== undefined) fn(value, baseRow + localRow, baseCol + localCol)
			}
		}
	}

	forEachOccupiedInRangeUnordered(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		fn: (row: number, col: number) => void,
	): void {
		for (let localRow = localRowStart; localRow <= localRowEnd; localRow++) {
			if ((this.rowCounts[localRow] ?? 0) === 0) continue
			const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
			const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
			if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) continue
			const startCol = Math.max(minLocalCol, rowMin)
			const endCol = Math.min(maxLocalCol, rowMax)
			const rowBase = localRow << CHUNK_BITS
			for (let localCol = startCol; localCol <= endCol; localCol++) {
				if (this.has(rowBase + localCol)) fn(baseRow + localRow, baseCol + localCol)
			}
		}
	}

	aggregateNumericInRange(
		localRowStart: number,
		localRowEnd: number,
		minLocalCol: number,
		maxLocalCol: number,
		baseRow: number,
		baseCol: number,
		stringTable: StringTable,
		state: NumericRangeAggregateState,
	): void {
		for (let localRow = localRowStart; localRow <= localRowEnd; localRow++) {
			if ((this.rowCounts[localRow] ?? 0) === 0) continue
			const rowMin = this.rowMinCol[localRow] ?? ROW_EMPTY_MIN
			const rowMax = this.rowMaxCol[localRow] ?? ROW_EMPTY_MAX
			if (rowMin > rowMax || rowMax < minLocalCol || rowMin > maxLocalCol) continue
			const startCol = Math.max(minLocalCol, rowMin)
			const endCol = Math.min(maxLocalCol, rowMax)
			const rowBase = localRow << CHUNK_BITS
			for (let localCol = startCol; localCol <= endCol; localCol++) {
				const localIndex = rowBase + localCol
				if (!this.has(localIndex)) continue
				aggregateNumericTag(
					this.readTag(localIndex),
					this.numbers[localIndex] ?? 0,
					this.stringIds[localIndex] ?? 0,
					stringTable,
					state,
					baseRow + localRow,
					baseCol + localCol,
				)
			}
		}
	}
}

export type ExpectedDensity = 'sparse' | 'dense' | 'auto'

export class SparseGrid {
	private chunkRows = new Map<number, Map<number, GridChunk>>()
	private _sortedChunkRows: readonly number[] | null = null
	private readonly _sortedChunkCols = new Map<number, readonly number[]>()
	private _sharedChunks: Set<number> | null = null
	private stringTable = new StringTable()
	private _cellCount = 0
	private _formulaCellCount = 0
	private _formulaInfoCellCount = 0
	private _stringCellCount = 0
	private _richTextCellCount = 0
	private _arrayCellCount = 0
	private _shared = false
	private _minRow = Number.POSITIVE_INFINITY
	private _maxRow = Number.NEGATIVE_INFINITY
	private _minCol = Number.POSITIVE_INFINITY
	private _maxCol = Number.NEGATIVE_INFINITY
	private _boundsDirty = false
	private _expectedDensity: ExpectedDensity = 'sparse'
	private _autoChunkCount = 0
	private _autoTotalCells = 0
	private _lastWriteChunkRow = -1
	private _lastWriteChunkCol = -1
	private _lastWriteCols: Map<number, GridChunk> | null = null
	private _lastWriteChunk: GridChunk | null = null

	setExpectedDensity(density: ExpectedDensity): void {
		this._expectedDensity = density
	}

	/** Returns the chunk implementation type at (row,col) for testing. */
	getChunkKindAt(row: number, col: number): 'sparse' | 'dense' | undefined {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const chunk = this.chunkRows.get(chunkRow)?.get(chunkCol)
		if (!chunk) return undefined
		return chunk instanceof DenseChunk ? 'dense' : 'sparse'
	}

	storageStats(): SparseGridStorageStats {
		const stats = {
			cellCount: this._cellCount,
			chunkCount: 0,
			denseChunkCount: 0,
			sparseChunkCount: 0,
			denseCellCount: 0,
			sparseCellCount: 0,
			denseCapacity: 0,
			sparseCapacity: 0,
			denseArrayBufferBytes: 0,
			sparseArrayBufferBytes: 0,
			styleArrayBufferBytes: 0,
			totalArrayBufferBytes: 0,
		}
		for (const cols of this.chunkRows.values()) {
			for (const chunk of cols.values()) {
				const chunkStats = chunk.storageStats()
				stats.chunkCount += 1
				stats.styleArrayBufferBytes += chunkStats.styleArrayBufferBytes
				stats.totalArrayBufferBytes += chunkStats.arrayBufferBytes
				if (chunkStats.kind === 'dense') {
					stats.denseChunkCount += 1
					stats.denseCellCount += chunkStats.cellCount
					stats.denseCapacity += chunkStats.capacity
					stats.denseArrayBufferBytes += chunkStats.arrayBufferBytes
				} else {
					stats.sparseChunkCount += 1
					stats.sparseCellCount += chunkStats.cellCount
					stats.sparseCapacity += chunkStats.capacity
					stats.sparseArrayBufferBytes += chunkStats.arrayBufferBytes
				}
			}
		}
		return stats
	}

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
		if (
			value.kind === 'number' &&
			formula === null &&
			styleId === DEFAULT_STYLE_ID &&
			formulaInfo === undefined
		) {
			this.setPlainNumber(row, col, value.value)
			return
		}
		this.ensureWritable()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		let chunk = this._writableChunkDirect(chunkRow, chunkCol)
		const cols = this._lastWriteCols as Map<number, GridChunk>
		chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		const existed = chunk.has(localIndex)
		const oldSlot = existed ? chunk.getSlot(localIndex) : undefined
		const hadFormula = existed && slotHasFormula(oldSlot)
		const hadFormulaInfo = oldSlot?.formulaInfo !== undefined
		const previousSharedStringKind = sharedStringKindFromSlot(oldSlot)
		const previousArray = arrayCellFromSlot(oldSlot)
		const hasFormula = formula !== null || formulaInfo !== undefined
		if (chunk instanceof DenseChunk) {
			chunk.setResolved(localIndex, value, formula, styleId, formulaInfo, this.stringTable)
		} else {
			const slot = compactResolvedCell(value, formula, styleId, formulaInfo, this.stringTable)
			const nextChunk = chunk.setSlot(localIndex, slot)
			if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
			chunk = nextChunk
		}
		this._rememberWriteChunk(chunkRow, chunkCol, cols, chunk)
		if (!existed) {
			this._cellCount++
			if (this._expectedDensity === 'auto') this._autoTotalCells++
		}
		if (hadFormula !== hasFormula) this._formulaCellCount += hasFormula ? 1 : -1
		if (hadFormulaInfo !== (formulaInfo !== undefined)) {
			this._formulaInfoCellCount += formulaInfo !== undefined ? 1 : -1
		}
		this.updateSharedStringCounts(previousSharedStringKind, sharedStringKindFromValue(value))
		this.updateArrayCounts(previousArray, value.kind === 'array')
		this._trackBounds(row, col)
	}

	setStringResolved(
		row: number,
		col: number,
		value: string,
		formula: string | null,
		styleId: StyleId,
		formulaInfo?: CellFormulaBinding,
	): void {
		this.ensureWritable()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		let chunk = this._writableChunkDirect(chunkRow, chunkCol)
		const cols = this._lastWriteCols as Map<number, GridChunk>
		chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		const existed = chunk.has(localIndex)
		const oldSlot = existed ? chunk.getSlot(localIndex) : undefined
		const hadFormula = existed && slotHasFormula(oldSlot)
		const hadFormulaInfo = oldSlot?.formulaInfo !== undefined
		const previousSharedStringKind = sharedStringKindFromSlot(oldSlot)
		const previousArray = arrayCellFromSlot(oldSlot)
		const hasFormula = formula !== null || formulaInfo !== undefined
		const nextChunk = chunk.setStringResolved(
			localIndex,
			value,
			formula,
			styleId,
			formulaInfo,
			this.stringTable,
		)
		if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
		this._rememberWriteChunk(chunkRow, chunkCol, cols, nextChunk)
		if (!existed) {
			this._cellCount++
			if (this._expectedDensity === 'auto') this._autoTotalCells++
		}
		if (hadFormula !== hasFormula) this._formulaCellCount += hasFormula ? 1 : -1
		if (hadFormulaInfo !== (formulaInfo !== undefined)) {
			this._formulaInfoCellCount += formulaInfo !== undefined ? 1 : -1
		}
		this.updateSharedStringCounts(previousSharedStringKind, 'string')
		this.updateArrayCounts(previousArray, false)
		this._trackBounds(row, col)
	}

	setPlainString(row: number, col: number, value: string): void {
		this.ensureWritable()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		let chunk = this._writableChunkDirect(chunkRow, chunkCol)
		const cols = this._lastWriteCols as Map<number, GridChunk>
		chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		const previousArray = arrayCellFromSlot(chunk.getSlot(localIndex))
		if (chunk instanceof DenseChunk) {
			const write = chunk.writePlainString(localIndex, value, this.stringTable)
			this._rememberWriteChunk(chunkRow, chunkCol, cols, chunk)
			if ((write & PLAIN_WRITE_EXISTED) === 0) {
				this._cellCount++
				if (this._expectedDensity === 'auto') this._autoTotalCells++
			} else {
				if ((write & PLAIN_WRITE_HAD_FORMULA) !== 0) this._formulaCellCount--
				if ((write & PLAIN_WRITE_HAD_FORMULA_INFO) !== 0) this._formulaInfoCellCount--
			}
			this.updateSharedStringCounts(previousSharedStringKindFromPlainWrite(write), 'string')
			this.updateArrayCounts((write & PLAIN_WRITE_PREVIOUS_ARRAY) !== 0, false)
			this._trackBounds(row, col)
			return
		}
		const existed = chunk.has(localIndex)
		let hadFormula = false
		let hadFormulaInfo = false
		if (existed) {
			const oldSlot = chunk.getSlot(localIndex)
			hadFormula = slotHasFormula(oldSlot)
			hadFormulaInfo = oldSlot?.formulaInfo !== undefined
			this.updateSharedStringCounts(sharedStringKindFromSlot(oldSlot), 'string')
		} else {
			this.updateSharedStringCounts(undefined, 'string')
		}
		const nextChunk = chunk.setStringResolved(
			localIndex,
			value,
			null,
			DEFAULT_STYLE_ID,
			undefined,
			this.stringTable,
		)
		if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
		this._rememberWriteChunk(chunkRow, chunkCol, cols, nextChunk)
		if (!existed) {
			this._cellCount++
			if (this._expectedDensity === 'auto') this._autoTotalCells++
		} else {
			if (hadFormula) this._formulaCellCount--
			if (hadFormulaInfo) this._formulaInfoCellCount--
		}
		this.updateArrayCounts(previousArray, false)
		this._trackBounds(row, col)
	}

	setNumberResolved(
		row: number,
		col: number,
		value: number,
		formula: string | null,
		styleId: StyleId,
		formulaInfo?: CellFormulaBinding,
	): void {
		this.ensureWritable()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		let chunk = this._writableChunkDirect(chunkRow, chunkCol)
		const cols = this._lastWriteCols as Map<number, GridChunk>
		chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		const existed = chunk.has(localIndex)
		const oldSlot = existed ? chunk.getSlot(localIndex) : undefined
		const hadFormula = existed && slotHasFormula(oldSlot)
		const hadFormulaInfo = oldSlot?.formulaInfo !== undefined
		const previousSharedStringKind = sharedStringKindFromSlot(oldSlot)
		const previousArray = arrayCellFromSlot(oldSlot)
		const hasFormula = formula !== null || formulaInfo !== undefined
		const nextChunk = chunk.setNumberResolved(localIndex, value, formula, styleId, formulaInfo)
		if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
		this._rememberWriteChunk(chunkRow, chunkCol, cols, nextChunk)
		if (!existed) {
			this._cellCount++
			if (this._expectedDensity === 'auto') this._autoTotalCells++
		}
		if (hadFormula !== hasFormula) this._formulaCellCount += hasFormula ? 1 : -1
		if (hadFormulaInfo !== (formulaInfo !== undefined)) {
			this._formulaInfoCellCount += formulaInfo !== undefined ? 1 : -1
		}
		this.updateSharedStringCounts(previousSharedStringKind, undefined)
		this.updateArrayCounts(previousArray, false)
		this._trackBounds(row, col)
	}

	setPlainNumber(row: number, col: number, value: number): void {
		this.ensureWritable()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		let chunk = this._writableChunkDirect(chunkRow, chunkCol)
		const cols = this._lastWriteCols as Map<number, GridChunk>
		chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		const previousArray = arrayCellFromSlot(chunk.getSlot(localIndex))
		if (chunk instanceof DenseChunk) {
			const write = chunk.writePlainNumber(localIndex, value)
			this._rememberWriteChunk(chunkRow, chunkCol, cols, chunk)
			if ((write & PLAIN_WRITE_EXISTED) === 0) {
				this._cellCount++
				if (this._expectedDensity === 'auto') this._autoTotalCells++
			} else {
				if ((write & PLAIN_WRITE_HAD_FORMULA) !== 0) this._formulaCellCount--
				if ((write & PLAIN_WRITE_HAD_FORMULA_INFO) !== 0) this._formulaInfoCellCount--
			}
			this.updateSharedStringCounts(previousSharedStringKindFromPlainWrite(write), undefined)
			this.updateArrayCounts((write & PLAIN_WRITE_PREVIOUS_ARRAY) !== 0, false)
			this._trackBounds(row, col)
			return
		}
		const existed = chunk.has(localIndex)
		let hadFormula = false
		let hadFormulaInfo = false
		if (existed) {
			const oldSlot = chunk.getSlot(localIndex)
			hadFormula = slotHasFormula(oldSlot)
			hadFormulaInfo = oldSlot?.formulaInfo !== undefined
			this.updateSharedStringCounts(sharedStringKindFromSlot(oldSlot), undefined)
		}
		const nextChunk = chunk.setNumberResolved(localIndex, value, null, DEFAULT_STYLE_ID, undefined)
		if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
		this._rememberWriteChunk(chunkRow, chunkCol, cols, nextChunk)
		if (!existed) {
			this._cellCount++
			if (this._expectedDensity === 'auto') this._autoTotalCells++
		} else {
			if (hadFormula) this._formulaCellCount--
			if (hadFormulaInfo) this._formulaInfoCellCount--
		}
		this.updateArrayCounts(previousArray, false)
		this._trackBounds(row, col)
	}

	setPlainNumberSpan(
		row: number,
		startCol: number,
		values: readonly number[],
		valueOffset = 0,
		count = values.length - valueOffset,
	): void {
		if (count <= 0) return
		this.ensureWritable()
		let remaining = count
		let col = startCol
		let offset = valueOffset
		while (remaining > 0) {
			const chunkRow = row >> CHUNK_BITS
			const chunkCol = col >> CHUNK_BITS
			const localRow = row & CHUNK_MASK
			const startLocalCol = col & CHUNK_MASK
			const segmentCount = Math.min(remaining, CHUNK_SIZE - startLocalCol)
			let chunk = this._writableChunkDirect(chunkRow, chunkCol)
			const cols = this._lastWriteCols as Map<number, GridChunk>
			chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
			if (!(chunk instanceof DenseChunk)) {
				for (let index = 0; index < segmentCount; index++) {
					this.setPlainNumber(row, col + index, values[offset + index] ?? 0)
				}
				col += segmentCount
				offset += segmentCount
				remaining -= segmentCount
				continue
			}
			const result = chunk.writePlainNumberSpan(
				localRow,
				startLocalCol,
				values,
				offset,
				segmentCount,
			)
			this._rememberWriteChunk(chunkRow, chunkCol, cols, chunk)
			this._cellCount += result.inserted
			if (this._expectedDensity === 'auto') this._autoTotalCells += result.inserted
			this._formulaCellCount -= result.formulaCleared
			this._formulaInfoCellCount -= result.formulaInfoCleared
			this._stringCellCount -= result.previousString
			this._richTextCellCount -= result.previousRichText
			this._arrayCellCount -= result.previousArray
			this._trackBounds(row, col)
			this._trackBounds(row, col + segmentCount - 1)
			col += segmentCount
			offset += segmentCount
			remaining -= segmentCount
		}
	}

	delete(row: number, col: number): boolean {
		this.ensureWritable()
		this._clearWriteCache()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		const cols = this.chunkRows.get(chunkRow)
		let chunk = cols?.get(chunkCol)
		if (!chunk) return false
		chunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		const oldSlot = chunk.getSlot(localIndex)
		const hadFormula = slotHasFormula(oldSlot)
		const hadFormulaInfo = chunk.readFormulaInfo(localIndex) !== undefined
		const previousSharedStringKind = sharedStringKindFromSlot(oldSlot)
		const previousArray = arrayCellFromSlot(oldSlot)
		const deleted = chunk.delete(localIndex)
		if (!deleted) return false
		this._cellCount--
		if (hadFormula) this._formulaCellCount--
		if (hadFormulaInfo) this._formulaInfoCellCount--
		this.updateSharedStringCounts(previousSharedStringKind, undefined)
		this.updateArrayCounts(previousArray, false)
		if (chunk.count === 0) {
			cols?.delete(chunkCol)
			this.invalidateChunkOrder(chunkRow)
			if (cols && cols.size === 0) {
				this.chunkRows.delete(chunkRow)
				this.invalidateChunkOrder()
			}
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
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readValue(localIndex, this.stringTable)
	}

	readValue(row: number, col: number): CellValue {
		return this.getValue(row, col) ?? EMPTY
	}

	readKind(row: number, col: number): CellValueKind | undefined {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.getKind(localIndex)
	}

	readNumber(row: number, col: number): number | null {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readNumber(localIndex) ?? null
	}

	readString(row: number, col: number): string | null {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return (
			this.chunkRows.get(chunkRow)?.get(chunkCol)?.readString(localIndex, this.stringTable) ?? null
		)
	}

	readError(row: number, col: number): ExcelError | null {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return (
			this.chunkRows.get(chunkRow)?.get(chunkCol)?.readError(localIndex, this.stringTable) ?? null
		)
	}

	readStyleId(row: number, col: number): StyleId | undefined {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readStyleId(localIndex)
	}

	readFormula(row: number, col: number): string | null | undefined {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readFormula(localIndex)
	}

	readFormulaInfo(row: number, col: number): CellFormulaBinding | undefined {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		return this.chunkRows.get(chunkRow)?.get(chunkCol)?.readFormulaInfo(localIndex)
	}

	clearFormulaInfo(row: number, col: number): void {
		this.ensureWritable()
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		const cols = this.chunkRows.get(chunkRow)
		const chunk = cols?.get(chunkCol)
		if (!chunk || !chunk.readFormulaInfo(localIndex)) return
		const writableChunk = this.ensureChunkWritable(chunkRow, chunkCol, cols, chunk)
		writableChunk.clearFormulaInfo(localIndex)
		if (writableChunk.readFormula(localIndex) === null) this._formulaCellCount--
		this._formulaInfoCellCount--
	}

	has(row: number, col: number): boolean {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
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
		if (this._cellCount === 0) return
		const startChunkRow = startRow >> CHUNK_BITS
		const endChunkRow = endRow >> CHUNK_BITS
		const startChunkCol = startCol >> CHUNK_BITS
		const endChunkCol = endCol >> CHUNK_BITS

		for (let chunkRow = startChunkRow; chunkRow <= endChunkRow; chunkRow++) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const useSparseChunkCols = cols.size * 2 < endChunkCol - startChunkCol + 1
			const sortedChunkCols = useSparseChunkCols ? this.getSortedChunkCols(chunkRow, cols) : null
			const localRowStart = chunkRow === startChunkRow ? startRow & CHUNK_MASK : 0
			const localRowEnd = chunkRow === endChunkRow ? endRow & CHUNK_MASK : CHUNK_MASK

			if (sortedChunkCols) {
				for (let localRow = localRowStart; localRow <= localRowEnd; localRow++) {
					const row = (chunkRow << CHUNK_BITS) + localRow
					for (const chunkCol of sortedChunkCols) {
						if (chunkCol < startChunkCol || chunkCol > endChunkCol) continue
						const chunk = cols.get(chunkCol)
						if (!chunk) continue
						const localColStart = chunkCol === startChunkCol ? startCol & CHUNK_MASK : 0
						const localColEnd = chunkCol === endChunkCol ? endCol & CHUNK_MASK : CHUNK_MASK
						const baseCol = chunkCol << CHUNK_BITS
						chunk.forEachValueInRow(
							localRow,
							localColStart,
							localColEnd,
							this.stringTable,
							(localCol, value) => fn(value, row, baseCol + localCol),
						)
					}
				}
				continue
			}

			for (let localRow = localRowStart; localRow <= localRowEnd; localRow++) {
				const row = (chunkRow << CHUNK_BITS) + localRow
				for (let chunkCol = startChunkCol; chunkCol <= endChunkCol; chunkCol++) {
					const chunk = cols.get(chunkCol)
					if (!chunk) continue
					const localColStart = chunkCol === startChunkCol ? startCol & CHUNK_MASK : 0
					const localColEnd = chunkCol === endChunkCol ? endCol & CHUNK_MASK : CHUNK_MASK
					const baseCol = chunkCol << CHUNK_BITS
					chunk.forEachValueInRow(
						localRow,
						localColStart,
						localColEnd,
						this.stringTable,
						(localCol, value) => fn(value, row, baseCol + localCol),
					)
				}
			}
		}
	}

	forEachValueInRangeUnordered(
		startRow: number,
		startCol: number,
		endRow: number,
		endCol: number,
		fn: (value: CellValue, row: number, col: number) => void,
	): void {
		if (this._cellCount === 0) return
		const startChunkRow = startRow >> CHUNK_BITS
		const endChunkRow = endRow >> CHUNK_BITS
		const startChunkCol = startCol >> CHUNK_BITS
		const endChunkCol = endCol >> CHUNK_BITS

		for (const [chunkRow, cols] of this.chunkRows) {
			if (chunkRow < startChunkRow || chunkRow > endChunkRow) continue
			const localRowStart = chunkRow === startChunkRow ? startRow & CHUNK_MASK : 0
			const localRowEnd = chunkRow === endChunkRow ? endRow & CHUNK_MASK : CHUNK_MASK
			for (const [chunkCol, chunk] of cols) {
				if (chunkCol < startChunkCol || chunkCol > endChunkCol) continue
				const localColStart = chunkCol === startChunkCol ? startCol & CHUNK_MASK : 0
				const localColEnd = chunkCol === endChunkCol ? endCol & CHUNK_MASK : CHUNK_MASK
				chunk.forEachValueInRangeUnordered(
					localRowStart,
					localRowEnd,
					localColStart,
					localColEnd,
					chunkRow << CHUNK_BITS,
					chunkCol << CHUNK_BITS,
					this.stringTable,
					fn,
				)
			}
		}
	}

	forEachOccupiedInRangeUnordered(
		startRow: number,
		startCol: number,
		endRow: number,
		endCol: number,
		fn: (row: number, col: number) => void,
	): void {
		if (this._cellCount === 0) return
		const startChunkRow = startRow >> CHUNK_BITS
		const endChunkRow = endRow >> CHUNK_BITS
		const startChunkCol = startCol >> CHUNK_BITS
		const endChunkCol = endCol >> CHUNK_BITS

		for (const [chunkRow, cols] of this.chunkRows) {
			if (chunkRow < startChunkRow || chunkRow > endChunkRow) continue
			const localRowStart = chunkRow === startChunkRow ? startRow & CHUNK_MASK : 0
			const localRowEnd = chunkRow === endChunkRow ? endRow & CHUNK_MASK : CHUNK_MASK
			for (const [chunkCol, chunk] of cols) {
				if (chunkCol < startChunkCol || chunkCol > endChunkCol) continue
				const localColStart = chunkCol === startChunkCol ? startCol & CHUNK_MASK : 0
				const localColEnd = chunkCol === endChunkCol ? endCol & CHUNK_MASK : CHUNK_MASK
				chunk.forEachOccupiedInRangeUnordered(
					localRowStart,
					localRowEnd,
					localColStart,
					localColEnd,
					chunkRow << CHUNK_BITS,
					chunkCol << CHUNK_BITS,
					fn,
				)
			}
		}
	}

	aggregateNumericInRange(
		startRow: number,
		startCol: number,
		endRow: number,
		endCol: number,
	): NumericRangeAggregate {
		const state: NumericRangeAggregateState = {
			sum: 0,
			count: 0,
			min: Number.POSITIVE_INFINITY,
			max: Number.NEGATIVE_INFINITY,
			error: null,
			errorRow: Number.POSITIVE_INFINITY,
			errorCol: Number.POSITIVE_INFINITY,
		}
		if (this._cellCount === 0) return numericRangeAggregateResult(state)
		const startChunkRow = startRow >> CHUNK_BITS
		const endChunkRow = endRow >> CHUNK_BITS
		const startChunkCol = startCol >> CHUNK_BITS
		const endChunkCol = endCol >> CHUNK_BITS

		for (let chunkRow = startChunkRow; chunkRow <= endChunkRow; chunkRow++) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const localRowStart = chunkRow === startChunkRow ? startRow & CHUNK_MASK : 0
			const localRowEnd = chunkRow === endChunkRow ? endRow & CHUNK_MASK : CHUNK_MASK
			const sortedChunkCols = this.getSortedChunkCols(chunkRow, cols)
			for (const chunkCol of sortedChunkCols) {
				if (chunkCol < startChunkCol || chunkCol > endChunkCol) continue
				const chunk = cols.get(chunkCol)
				if (!chunk) continue
				const localColStart = chunkCol === startChunkCol ? startCol & CHUNK_MASK : 0
				const localColEnd = chunkCol === endChunkCol ? endCol & CHUNK_MASK : CHUNK_MASK
				chunk.aggregateNumericInRange(
					localRowStart,
					localRowEnd,
					localColStart,
					localColEnd,
					chunkRow << CHUNK_BITS,
					chunkCol << CHUNK_BITS,
					this.stringTable,
					state,
				)
			}
		}
		return numericRangeAggregateResult(state)
	}

	forEachRow(fn: (row: number, cells: ReadonlyMap<number, CellValue>) => void): void {
		if (this._cellCount === 0) return
		const rowCells = new Map<number, CellValue>()
		for (const [row, entries] of this.iterateRows()) {
			rowCells.clear()
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
		const sortedChunkRows = this.getSortedChunkRows()
		for (const chunkRow of sortedChunkRows) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const sortedChunkCols = this.getSortedChunkCols(chunkRow, cols)
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
			const sortedChunkCols = this.getSortedChunkCols(chunkRow, cols)
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

	forEachCellInRange(range: RangeRef, fn: (row: number, col: number, cell: Cell) => void): void {
		if (this._cellCount === 0) return
		const startChunkRow = range.start.row >> CHUNK_BITS
		const endChunkRow = range.end.row >> CHUNK_BITS
		for (let chunkRow = startChunkRow; chunkRow <= endChunkRow; chunkRow++) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const sortedChunkCols = this.getSortedChunkCols(chunkRow, cols)
			const rowStart = chunkRow === startChunkRow ? range.start.row & CHUNK_MASK : 0
			const rowEnd = chunkRow === endChunkRow ? range.end.row & CHUNK_MASK : CHUNK_MASK
			for (let localRow = rowStart; localRow <= rowEnd; localRow++) {
				const row = (chunkRow << CHUNK_BITS) + localRow
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
						if (cell) fn(row, chunkStartCol + localCol, cell)
					})
				}
			}
		}
	}

	forEachCellContentInRange(
		range: RangeRef,
		fn: (
			row: number,
			col: number,
			value: CellValue,
			formula: string | null,
			formulaInfo: CellFormulaBinding | undefined,
		) => void,
	): void {
		if (this._cellCount === 0) return
		const startChunkRow = range.start.row >> CHUNK_BITS
		const endChunkRow = range.end.row >> CHUNK_BITS
		for (let chunkRow = startChunkRow; chunkRow <= endChunkRow; chunkRow++) {
			const cols = this.chunkRows.get(chunkRow)
			if (!cols) continue
			const sortedChunkCols = this.getSortedChunkCols(chunkRow, cols)
			const rowStart = chunkRow === startChunkRow ? range.start.row & CHUNK_MASK : 0
			const rowEnd = chunkRow === endChunkRow ? range.end.row & CHUNK_MASK : CHUNK_MASK
			for (let localRow = rowStart; localRow <= rowEnd; localRow++) {
				const row = (chunkRow << CHUNK_BITS) + localRow
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
						fn(
							row,
							chunkStartCol + localCol,
							readSlotValue(slot, this.stringTable) ?? EMPTY,
							slot.formula,
							slot.formulaInfo,
						)
					})
				}
			}
		}
	}

	cellCount(): number {
		return this._cellCount
	}

	formulaCellCount(): number {
		return this._formulaCellCount
	}

	formulaInfoCellCount(): number {
		return this._formulaInfoCellCount
	}

	stringCellCount(): number {
		return this._stringCellCount
	}

	richTextCellCount(): number {
		return this._richTextCellCount
	}

	arrayCellCount(): number {
		return this._arrayCellCount
	}

	clear(): void {
		this.chunkRows = new Map()
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		this._sharedChunks = null
		this._cellCount = 0
		this._formulaCellCount = 0
		this._formulaInfoCellCount = 0
		this._stringCellCount = 0
		this._richTextCellCount = 0
		this._arrayCellCount = 0
		this._shared = false
		this._minRow = Number.POSITIVE_INFINITY
		this._maxRow = Number.NEGATIVE_INFINITY
		this._minCol = Number.POSITIVE_INFINITY
		this._maxCol = Number.NEGATIVE_INFINITY
		this._boundsDirty = false
		this._autoChunkCount = 0
		this._autoTotalCells = 0
	}

	insertRows(at: number, count: number): void {
		if (count === 0 || this._cellCount === 0) return
		this.ensureWritable()
		const atChunkRow = at >> CHUNK_BITS
		const atLocalRow = at & CHUNK_MASK
		const countChunks = count >> CHUNK_BITS
		if (atLocalRow === 0 && (count & CHUNK_MASK) === 0 && countChunks > 0) {
			this._shiftChunkRows(atChunkRow, countChunks)
			return
		}
		this._rebuildAfterRowInsert(at, count)
	}

	deleteRows(at: number, count: number): void {
		if (count === 0 || this._cellCount === 0) return
		this.ensureWritable()
		const atChunkRow = at >> CHUNK_BITS
		const atLocalRow = at & CHUNK_MASK
		const countChunks = count >> CHUNK_BITS
		if (atLocalRow === 0 && (count & CHUNK_MASK) === 0 && countChunks > 0) {
			this._shiftChunkRowsForDelete(atChunkRow, countChunks)
			return
		}
		const deleteEnd = at + count
		this._rebuildAfterRowDelete(at, deleteEnd, count)
	}

	insertCols(at: number, count: number): void {
		if (count === 0 || this._cellCount === 0) return
		this.ensureWritable()
		const atChunkCol = at >> CHUNK_BITS
		const atLocalCol = at & CHUNK_MASK
		const countChunks = count >> CHUNK_BITS
		if (atLocalCol === 0 && (count & CHUNK_MASK) === 0 && countChunks > 0) {
			this._shiftChunkCols(atChunkCol, countChunks)
			return
		}
		this._rebuildAfterColInsert(at, count)
	}

	deleteCols(at: number, count: number): void {
		if (count === 0 || this._cellCount === 0) return
		this.ensureWritable()
		const atChunkCol = at >> CHUNK_BITS
		const atLocalCol = at & CHUNK_MASK
		const countChunks = count >> CHUNK_BITS
		if (atLocalCol === 0 && (count & CHUNK_MASK) === 0 && countChunks > 0) {
			this._shiftChunkColsForDelete(atChunkCol, countChunks)
			return
		}
		const deleteEnd = at + count
		this._rebuildAfterColDelete(at, deleteEnd, count)
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
			this._clearWriteCache()
			this._sortedChunkRows = null
			this._sortedChunkCols.clear()
			this._sharedChunks = null
			this.stringTable = next.stringTable
			this._cellCount = next._cellCount
			this._formulaCellCount = next._formulaCellCount
			this._formulaInfoCellCount = next._formulaInfoCellCount
			this._stringCellCount = next._stringCellCount
			this._richTextCellCount = next._richTextCellCount
			this._arrayCellCount = next._arrayCellCount
			this._minRow = next._minRow
			this._maxRow = next._maxRow
			this._minCol = next._minCol
			this._maxCol = next._maxCol
			this._boundsDirty = next._boundsDirty
			this._shared = false
			return
		}
		this.chunkRows = other.chunkRows
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		this._sharedChunks = null
		this.stringTable = other.stringTable
		this._cellCount = other._cellCount
		this._formulaCellCount = other._formulaCellCount
		this._formulaInfoCellCount = other._formulaInfoCellCount
		this._stringCellCount = other._stringCellCount
		this._richTextCellCount = other._richTextCellCount
		this._arrayCellCount = other._arrayCellCount
		this._minRow = other._minRow
		this._maxRow = other._maxRow
		this._minCol = other._minCol
		this._maxCol = other._maxCol
		this._boundsDirty = other._boundsDirty
		this._shared = true
		other._shared = true
	}

	private _createNewChunk(): GridChunk {
		if (this._expectedDensity === 'auto') {
			if (this._autoChunkCount >= AUTO_DENSE_SAMPLE_CHUNKS) {
				const fillRatio = this._autoTotalCells / (AUTO_DENSE_SAMPLE_CHUNKS * CHUNK_AREA)
				if (fillRatio > AUTO_DENSE_FILL_RATIO) {
					this._expectedDensity = 'dense'
					this._autoChunkCount++
					return new DenseChunk()
				}
			}
			this._autoChunkCount++
		}
		if (this._expectedDensity === 'dense') return new DenseChunk()
		return new SparseChunk()
	}

	private _writableChunkDirect(chunkRow: number, chunkCol: number): GridChunk {
		if (
			this._lastWriteChunkRow === chunkRow &&
			this._lastWriteChunkCol === chunkCol &&
			this._lastWriteCols !== null &&
			this._lastWriteChunk !== null
		) {
			return this._lastWriteChunk
		}
		let cols = this.chunkRows.get(chunkRow)
		if (!cols) {
			cols = new Map()
			this.chunkRows.set(chunkRow, cols)
			this.invalidateChunkOrder()
		}
		let chunk = cols.get(chunkCol)
		if (!chunk) {
			chunk = this._createNewChunk()
			cols.set(chunkCol, chunk)
			this.invalidateChunkOrder(chunkRow)
		}
		this._rememberWriteChunk(chunkRow, chunkCol, cols, chunk)
		return chunk
	}

	private _rememberWriteChunk(
		chunkRow: number,
		chunkCol: number,
		cols: Map<number, GridChunk>,
		chunk: GridChunk,
	): void {
		this._lastWriteChunkRow = chunkRow
		this._lastWriteChunkCol = chunkCol
		this._lastWriteCols = cols
		this._lastWriteChunk = chunk
	}

	private _clearWriteCache(): void {
		this._lastWriteChunkRow = -1
		this._lastWriteChunkCol = -1
		this._lastWriteCols = null
		this._lastWriteChunk = null
	}

	private ensureWritable(): void {
		if (!this._shared) return
		const nextChunkRows = new Map<number, Map<number, GridChunk>>()
		const sharedChunks = new Set<number>()
		for (const [chunkRow, cols] of this.chunkRows) {
			const nextCols = new Map<number, GridChunk>()
			for (const [chunkCol, chunk] of cols) {
				nextCols.set(chunkCol, chunk)
				sharedChunks.add(sharedChunkKey(chunkRow, chunkCol))
			}
			nextChunkRows.set(chunkRow, nextCols)
		}
		this.chunkRows = nextChunkRows
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		this._sharedChunks = sharedChunks
		this._shared = false
	}

	private _getSlot(row: number, col: number): StoredSlot | undefined {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
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

	private _shiftChunkRows(atChunkRow: number, countChunks: number): void {
		const next = new Map<number, Map<number, GridChunk>>()
		for (const [chunkRow, cols] of this.chunkRows) {
			if (chunkRow < atChunkRow) {
				next.set(chunkRow, cols)
			} else {
				next.set(chunkRow + countChunks, cols)
			}
		}
		this.chunkRows = next
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		if (this._minRow >= atChunkRow << CHUNK_BITS) this._minRow += countChunks << CHUNK_BITS
		if (this._maxRow >= atChunkRow << CHUNK_BITS) this._maxRow += countChunks << CHUNK_BITS
	}

	private _shiftChunkRowsForDelete(atChunkRow: number, countChunks: number): void {
		const next = new Map<number, Map<number, GridChunk>>()
		const deleteEndChunk = atChunkRow + countChunks
		let removedCount = 0
		let removedFormulaCount = 0
		let removedFormulaInfoCount = 0
		let removedStringCount = 0
		let removedRichTextCount = 0
		let removedArrayCount = 0
		for (const [chunkRow, cols] of this.chunkRows) {
			if (chunkRow < atChunkRow) {
				next.set(chunkRow, cols)
			} else if (chunkRow >= deleteEndChunk) {
				next.set(chunkRow - countChunks, cols)
			} else {
				for (const chunk of cols.values()) {
					removedCount += chunk.count
					removedFormulaCount += chunk.countFormulaCells()
					removedFormulaInfoCount += chunk.countFormulaInfoCells()
					removedStringCount += chunk.countStringCells()
					removedRichTextCount += chunk.countRichTextCells()
					removedArrayCount += chunk.countArrayCells()
				}
			}
		}
		this.chunkRows = next
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		this._cellCount -= removedCount
		this._formulaCellCount -= removedFormulaCount
		this._formulaInfoCellCount -= removedFormulaInfoCount
		this._stringCellCount -= removedStringCount
		this._richTextCellCount -= removedRichTextCount
		this._arrayCellCount -= removedArrayCount
		this._boundsDirty = true
	}

	private _shiftChunkCols(atChunkCol: number, countChunks: number): void {
		const next = new Map<number, Map<number, GridChunk>>()
		for (const [chunkRow, cols] of this.chunkRows) {
			const nextCols = new Map<number, GridChunk>()
			for (const [chunkCol, chunk] of cols) {
				if (chunkCol < atChunkCol) {
					nextCols.set(chunkCol, chunk)
				} else {
					nextCols.set(chunkCol + countChunks, chunk)
				}
			}
			next.set(chunkRow, nextCols)
		}
		this.chunkRows = next
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		const boundary = atChunkCol << CHUNK_BITS
		if (this._minCol >= boundary) this._minCol += countChunks << CHUNK_BITS
		if (this._maxCol >= boundary) this._maxCol += countChunks << CHUNK_BITS
	}

	private _shiftChunkColsForDelete(atChunkCol: number, countChunks: number): void {
		const next = new Map<number, Map<number, GridChunk>>()
		const deleteEndChunk = atChunkCol + countChunks
		let removedCount = 0
		let removedFormulaCount = 0
		let removedFormulaInfoCount = 0
		let removedStringCount = 0
		let removedRichTextCount = 0
		let removedArrayCount = 0
		for (const [chunkRow, cols] of this.chunkRows) {
			const nextCols = new Map<number, GridChunk>()
			for (const [chunkCol, chunk] of cols) {
				if (chunkCol < atChunkCol) {
					nextCols.set(chunkCol, chunk)
				} else if (chunkCol >= deleteEndChunk) {
					nextCols.set(chunkCol - countChunks, chunk)
				} else {
					removedCount += chunk.count
					removedFormulaCount += chunk.countFormulaCells()
					removedFormulaInfoCount += chunk.countFormulaInfoCells()
					removedStringCount += chunk.countStringCells()
					removedRichTextCount += chunk.countRichTextCells()
					removedArrayCount += chunk.countArrayCells()
				}
			}
			next.set(chunkRow, nextCols)
		}
		this.chunkRows = next
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		this._cellCount -= removedCount
		this._formulaCellCount -= removedFormulaCount
		this._formulaInfoCellCount -= removedFormulaInfoCount
		this._stringCellCount -= removedStringCount
		this._richTextCellCount -= removedRichTextCount
		this._arrayCellCount -= removedArrayCount
		this._boundsDirty = true
	}

	private _rebuildAfterRowInsert(at: number, count: number): void {
		const next = new SparseGrid()
		next.stringTable = this.stringTable
		next._expectedDensity = this._expectedDensity
		this._forEachSlot((row, col, slot) => {
			next._setShiftedSlot(row >= at ? row + count : row, col, slot)
		})
		this._replaceWithRebuilt(next)
	}

	private _rebuildAfterRowDelete(at: number, deleteEnd: number, count: number): void {
		const next = new SparseGrid()
		next.stringTable = this.stringTable
		next._expectedDensity = this._expectedDensity
		this._forEachSlot((row, col, slot) => {
			if (row >= at && row < deleteEnd) return
			next._setShiftedSlot(row >= deleteEnd ? row - count : row, col, slot)
		})
		this._replaceWithRebuilt(next)
	}

	private _rebuildAfterColInsert(at: number, count: number): void {
		const next = new SparseGrid()
		next.stringTable = this.stringTable
		next._expectedDensity = this._expectedDensity
		this._forEachSlot((row, col, slot) => {
			next._setShiftedSlot(row, col >= at ? col + count : col, slot)
		})
		this._replaceWithRebuilt(next)
	}

	private _rebuildAfterColDelete(at: number, deleteEnd: number, count: number): void {
		const next = new SparseGrid()
		next.stringTable = this.stringTable
		next._expectedDensity = this._expectedDensity
		this._forEachSlot((row, col, slot) => {
			if (col >= at && col < deleteEnd) return
			next._setShiftedSlot(row, col >= deleteEnd ? col - count : col, slot)
		})
		this._replaceWithRebuilt(next)
	}

	private _forEachSlot(fn: (row: number, col: number, slot: StoredSlot) => void): void {
		for (const [chunkRow, cols] of this.chunkRows) {
			const baseRow = chunkRow << CHUNK_BITS
			for (const [chunkCol, chunk] of cols) {
				const baseCol = chunkCol << CHUNK_BITS
				for (let localRow = 0; localRow < CHUNK_SIZE; localRow++) {
					const row = baseRow + localRow
					chunk.forEachRow(localRow, 0, CHUNK_MASK, (localCol, slot) => {
						fn(row, baseCol + localCol, slot)
					})
				}
			}
		}
	}

	private _setShiftedSlot(row: number, col: number, slot: StoredSlot): void {
		const chunkRow = row >> CHUNK_BITS
		const chunkCol = col >> CHUNK_BITS
		const localIndex = ((row & CHUNK_MASK) << CHUNK_BITS) | (col & CHUNK_MASK)
		let chunk = this._writableChunkDirect(chunkRow, chunkCol)
		const cols = this._lastWriteCols as Map<number, GridChunk>
		const nextChunk = chunk.setSlot(localIndex, slotForShiftTarget(slot, chunk))
		if (nextChunk !== chunk) cols.set(chunkCol, nextChunk)
		chunk = nextChunk
		this._rememberWriteChunk(chunkRow, chunkCol, cols, chunk)
		this._cellCount++
		if (this._expectedDensity === 'auto') this._autoTotalCells++
		if (slotHasFormula(slot)) this._formulaCellCount++
		if (slot.formulaInfo !== undefined) this._formulaInfoCellCount++
		this.updateSharedStringCounts(undefined, sharedStringKindFromSlot(slot))
		this.updateArrayCounts(false, arrayCellFromSlot(slot))
		this._trackBounds(row, col)
	}

	private _replaceWithRebuilt(next: SparseGrid): void {
		this.chunkRows = next.chunkRows
		this._clearWriteCache()
		this._sortedChunkRows = null
		this._sortedChunkCols.clear()
		this._sharedChunks = null
		this._cellCount = next._cellCount
		this._formulaCellCount = next._formulaCellCount
		this._formulaInfoCellCount = next._formulaInfoCellCount
		this._stringCellCount = next._stringCellCount
		this._richTextCellCount = next._richTextCellCount
		this._arrayCellCount = next._arrayCellCount
		this._minRow = next._minRow
		this._maxRow = next._maxRow
		this._minCol = next._minCol
		this._maxCol = next._maxCol
		this._boundsDirty = next._boundsDirty
		this._shared = false
		this._expectedDensity = next._expectedDensity
		this._autoChunkCount = next._autoChunkCount
		this._autoTotalCells = next._autoTotalCells
	}

	private invalidateChunkOrder(chunkRow?: number): void {
		this._sortedChunkRows = null
		if (chunkRow === undefined) {
			this._sortedChunkCols.clear()
			return
		}
		this._sortedChunkCols.delete(chunkRow)
	}

	private getSortedChunkRows(): readonly number[] {
		if (this._sortedChunkRows) return this._sortedChunkRows
		const sorted = [...this.chunkRows.keys()].sort((a, b) => a - b)
		this._sortedChunkRows = sorted
		return sorted
	}

	private getSortedChunkCols(
		chunkRow: number,
		cols: ReadonlyMap<number, GridChunk>,
	): readonly number[] {
		const cached = this._sortedChunkCols.get(chunkRow)
		if (cached) return cached
		const sorted = [...cols.keys()].sort((a, b) => a - b)
		this._sortedChunkCols.set(chunkRow, sorted)
		return sorted
	}

	private ensureChunkWritable(
		chunkRow: number,
		chunkCol: number,
		cols: Map<number, GridChunk> | undefined,
		chunk: GridChunk,
	): GridChunk {
		if (!cols || !this._sharedChunks) return chunk
		const key = sharedChunkKey(chunkRow, chunkCol)
		if (!this._sharedChunks.has(key)) return chunk
		const next = chunk.clone()
		cols.set(chunkCol, next)
		this._sharedChunks.delete(key)
		return next
	}

	private updateSharedStringCounts(
		previous: SharedStringCellKind | undefined,
		next: SharedStringCellKind | undefined,
	): void {
		if (previous === next) return
		if (previous === 'string') this._stringCellCount--
		if (previous === 'richText') this._richTextCellCount--
		if (next === 'string') this._stringCellCount++
		if (next === 'richText') this._richTextCellCount++
	}

	private updateArrayCounts(previous: boolean, next: boolean): void {
		if (previous === next) return
		this._arrayCellCount += next ? 1 : -1
	}

	private _hasMutableValues(): boolean {
		return this._arrayCellCount > 0 || this._richTextCellCount > 0
	}
}

function sharedChunkKey(chunkRow: number, chunkCol: number): number {
	return chunkRow * CHUNK_COL_FACTOR + chunkCol
}

function resolveChunkBits(): number {
	const env = (globalThis as { process?: { env?: Record<string, string | undefined> } }).process
		?.env
	const raw = env?.ASCEND_CHUNK_BITS
	if (!raw) return 5
	const parsed = Number.parseInt(raw, 10)
	return parsed === 4 || parsed === 5 || parsed === 6 ? parsed : 6
}

function createChunkBuffer(byteLength: number): ChunkBuffer {
	return new ArrayBuffer(byteLength)
}

function cloneCellValue(value: CellValue): CellValue {
	if (value.kind === 'array') return structuredClone(value)
	if (value.kind === 'richText') return { kind: 'richText', runs: [...value.runs] }
	return value
}

function cloneStoredSlotForShift(slot: StoredSlot): StoredSlot {
	return {
		tag: slot.tag,
		styleId: slot.styleId,
		formula: slot.formula,
		formulaInfo: slot.formulaInfo,
		numberValue: slot.numberValue,
		stringId: slot.stringId,
		heapValue: slot.heapValue ? cloneCellValue(slot.heapValue) : undefined,
	}
}

function slotForShiftTarget(slot: StoredSlot, target: GridChunk): StoredSlot {
	if (target instanceof DenseChunk && slot.heapValue === undefined) return slot
	return cloneStoredSlotForShift(slot)
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

function compactStringResolvedCell(
	value: string,
	formula: string | null,
	styleId: StyleId,
	formulaInfo: CellFormulaBinding | undefined,
	stringTable: StringTable,
): StoredSlot {
	return {
		tag: SlotTag.String,
		styleId,
		formula,
		formulaInfo,
		numberValue: undefined,
		stringId: stringTable.intern(value),
		heapValue: undefined,
	}
}

function compactNumberResolvedCell(
	value: number,
	formula: string | null,
	styleId: StyleId,
	formulaInfo: CellFormulaBinding | undefined,
): StoredSlot {
	return {
		tag: SlotTag.Number,
		styleId,
		formula,
		formulaInfo,
		numberValue: value,
		stringId: undefined,
		heapValue: undefined,
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

function slotHasFormula(slot: StoredSlot | undefined): boolean {
	return slot !== undefined && (slot.formula !== null || slot.formulaInfo !== undefined)
}

function sharedStringKindFromValue(value: CellValue): SharedStringCellKind | undefined {
	if (value.kind === 'string') return 'string'
	if (value.kind === 'richText') return 'richText'
	return undefined
}

function sharedStringKindFromSlot(slot: StoredSlot | undefined): SharedStringCellKind | undefined {
	if (!slot) return undefined
	if (slot.tag === SlotTag.String) return 'string'
	if (slot.heapValue?.kind === 'richText') return 'richText'
	return undefined
}

function arrayCellFromSlot(slot: StoredSlot | undefined): boolean {
	return slot?.heapValue?.kind === 'array'
}

function previousSharedStringKindFromPlainWrite(flags: number): SharedStringCellKind | undefined {
	if ((flags & PLAIN_WRITE_PREVIOUS_STRING) !== 0) return 'string'
	if ((flags & PLAIN_WRITE_PREVIOUS_RICH_TEXT) !== 0) return 'richText'
	return undefined
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
			return dateValue(slot.numberValue ?? 0)
		case SlotTag.Heap:
			return slot.heapValue
	}
}

function aggregateNumericSlot(
	slot: StoredSlot,
	stringTable: StringTable,
	state: NumericRangeAggregateState,
	row = Number.POSITIVE_INFINITY,
	col = Number.POSITIVE_INFINITY,
): void {
	aggregateNumericTag(
		slot.tag,
		slot.numberValue ?? 0,
		slot.stringId ?? 0,
		stringTable,
		state,
		row,
		col,
	)
}

function aggregateNumericTag(
	tag: SlotTag,
	number: number,
	stringId: number,
	stringTable: StringTable,
	state: NumericRangeAggregateState,
	row = Number.POSITIVE_INFINITY,
	col = Number.POSITIVE_INFINITY,
): void {
	if (tag === SlotTag.Number || tag === SlotTag.Date) {
		state.sum += number
		state.count += 1
		if (number < state.min) state.min = number
		if (number > state.max) state.max = number
		return
	}
	if (
		tag === SlotTag.Error &&
		(row < state.errorRow || (row === state.errorRow && col < state.errorCol))
	) {
		state.error = stringTable.lookup(stringId) as ExcelError
		state.errorRow = row
		state.errorCol = col
	}
}

function numericRangeAggregateResult(state: NumericRangeAggregateState): NumericRangeAggregate {
	if (state.error) {
		return {
			sum: 0,
			count: 0,
			min: Number.POSITIVE_INFINITY,
			max: Number.NEGATIVE_INFINITY,
			error: state.error,
		}
	}
	return {
		sum: state.sum,
		count: state.count,
		min: state.min,
		max: state.max,
		error: state.error,
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
			return dateValue(number)
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
