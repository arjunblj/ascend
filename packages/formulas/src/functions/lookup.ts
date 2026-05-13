import type { CellValue, ScalarCellValue } from '@ascend/schema'
import {
	arrayValue,
	booleanValue,
	EMPTY,
	errorValue,
	numberValue,
	topLeftScalar,
} from '@ascend/schema'
import type { FunctionDef } from './registry.ts'
import {
	cellOf,
	compareValues,
	type EvalArg,
	type ExactLookupHit,
	exactLookupHitFirst,
	exactLookupHitLast,
	type FunctionEvalContext,
	getRange,
	hasWildcardPatternSyntax,
	numArg,
	packExactLookupHit,
	rangeShape,
	valuesEqual,
	wildcardMatch,
} from './registry.ts'

function v(data: readonly CellValue[], i: number): CellValue {
	return data[i] ?? EMPTY
}

function isBlankLookupValue(value: CellValue): boolean {
	const scalar = topLeftScalar(value)
	return scalar.kind === 'empty' || (scalar.kind === 'string' && scalar.value === '')
}

function nearestNonBlankIndex(
	data: readonly CellValue[],
	lo: number,
	hi: number,
	mid: number,
): number {
	if (!isBlankLookupValue(v(data, mid))) return mid
	for (let offset = 1; mid - offset >= lo || mid + offset <= hi; offset++) {
		const left = mid - offset
		if (left >= lo && !isBlankLookupValue(v(data, left))) return left
		const right = mid + offset
		if (right <= hi && !isBlankLookupValue(v(data, right))) return right
	}
	return -1
}

function approximateMatch(lookup: CellValue, data: readonly CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		const probe = nearestNonBlankIndex(data, lo, hi, mid)
		if (probe < 0) break
		if (compareValues(lookup, v(data, probe)) >= 0) {
			result = probe
			lo = probe + 1
		} else {
			hi = probe - 1
		}
	}
	return result
}

function reverseApproximateMatch(lookup: CellValue, data: readonly CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		const probe = nearestNonBlankIndex(data, lo, hi, mid)
		if (probe < 0) break
		if (compareValues(v(data, probe), lookup) >= 0) {
			result = probe
			lo = probe + 1
		} else {
			hi = probe - 1
		}
	}
	return result
}

function descendingNextSmallerMatch(lookup: CellValue, data: readonly CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		const probe = nearestNonBlankIndex(data, lo, hi, mid)
		if (probe < 0) break
		if (compareValues(v(data, probe), lookup) <= 0) {
			result = probe
			hi = probe - 1
		} else {
			lo = probe + 1
		}
	}
	return result
}

function exactMatch(lookup: CellValue, data: readonly CellValue[]): number {
	for (let i = 0; i < data.length; i++) {
		if (valuesEqual(lookup, v(data, i))) return i
	}
	return -1
}

function exactLookupValueKey(value: CellValue): string | null {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'number':
			return `v:${String(scalar.value === 0 ? 0 : scalar.value)}`
		case 'date':
			return `v:${String(scalar.serial === 0 ? 0 : scalar.serial)}`
		case 'string':
			return `s:${scalar.value.toLowerCase()}`
		case 'boolean':
			return scalar.value ? 'b:1' : 'b:0'
		case 'empty':
			return 'e'
		default:
			return null
	}
}

function lookupRangeCacheKey(
	arg: EvalArg | undefined,
	orientation: 'row' | 'column',
): string | null {
	const ref = arg?.ref
	if (!ref || ref.kind !== 'range') return null
	if (arg?.areas && arg.areas.length > 1) return null
	if (orientation === 'column') {
		return `column:${ref.sheetIndex}:${ref.col}:${ref.row}:${ref.endRow ?? ref.row}`
	}
	return `row:${ref.sheetIndex}:${ref.row}:${ref.col}:${ref.endCol ?? ref.col}`
}

function getLookupVector(
	arg: EvalArg | undefined,
	matrix: readonly (readonly CellValue[])[],
	orientation: 'row' | 'column',
	ctx?: FunctionEvalContext,
): readonly CellValue[] {
	const cacheKey = lookupRangeCacheKey(arg, orientation)
	if (cacheKey && ctx?.lookupVectorCache) {
		const cached = ctx.lookupVectorCache.get(cacheKey)
		if (cached) return cached
		const built = orientation === 'row' ? [...(matrix[0] ?? [])] : extractColumn(matrix, 0)
		ctx.lookupVectorCache.set(cacheKey, built)
		return built
	}
	return orientation === 'row' ? [...(matrix[0] ?? [])] : extractColumn(matrix, 0)
}

function getArgVector(
	arg: EvalArg | undefined,
	ctx?: FunctionEvalContext,
): { values: readonly CellValue[]; orientation: 'row' | 'column' } | null {
	const result = getArgVectorWithIndex(arg, ctx, false)
	return result ? { values: result.values, orientation: result.orientation } : null
}

function getArgVectorWithIndex(
	arg: EvalArg | undefined,
	ctx?: FunctionEvalContext,
	buildIndex = false,
): {
	values: readonly CellValue[]
	orientation: 'row' | 'column'
	exactIndex?: ReadonlyMap<string, ExactLookupHit>
} | null {
	if (!arg || (arg.areas && arg.areas.length > 1)) return null
	const shape = rangeShape(arg)
	const orientation = shape.rows === 1 ? 'row' : shape.cols === 1 ? 'column' : null
	if (!orientation) return null
	const cacheKey = lookupRangeCacheKey(arg, orientation)
	if (cacheKey && ctx?.lookupVectorCache) {
		const cached = ctx.lookupVectorCache.get(cacheKey)
		if (cached) {
			let exactIndex: ReadonlyMap<string, ExactLookupHit> | undefined
			if (buildIndex && cacheKey && ctx.exactLookupCache) {
				exactIndex = ctx.exactLookupCache.get(cacheKey)
				if (!exactIndex) {
					exactIndex = buildExactLookupIndex(cached)
					ctx.exactLookupCache.set(cacheKey, exactIndex as Map<string, ExactLookupHit>)
				}
			}
			return {
				values: cached,
				orientation,
				...(buildIndex && exactIndex ? { exactIndex } : {}),
			}
		}
	}
	let values: readonly CellValue[]
	let exactIndex: Map<string, ExactLookupHit> | undefined
	if (arg.forEachValue) {
		const next: CellValue[] = []
		const index =
			buildIndex && cacheKey && ctx?.exactLookupCache
				? new Map<string, ExactLookupHit>()
				: undefined
		arg.forEachValue((value) => {
			const i = next.length
			next.push(value)
			if (index) {
				const key = exactLookupValueKey(value)
				if (key !== null) {
					const existing = index.get(key)
					if (existing !== undefined) {
						index.set(key, packExactLookupHit(exactLookupHitFirst(existing), i))
					} else {
						index.set(key, packExactLookupHit(i, i))
					}
				}
			}
		})
		values = next
		exactIndex = index
	} else if (arg.kind === 'range' && arg.values) {
		values = orientation === 'row' ? [...(arg.values[0] ?? [])] : extractColumn(arg.values, 0)
		exactIndex =
			buildIndex && cacheKey && ctx?.exactLookupCache ? buildExactLookupIndex(values) : undefined
	} else if (arg.value.kind === 'array') {
		values =
			orientation === 'row' ? [...(arg.value.rows[0] ?? [])] : extractColumn(arg.value.rows, 0)
		exactIndex =
			buildIndex && cacheKey && ctx?.exactLookupCache ? buildExactLookupIndex(values) : undefined
	} else {
		values = [arg.value]
		exactIndex =
			buildIndex && cacheKey && ctx?.exactLookupCache ? buildExactLookupIndex(values) : undefined
	}
	if (cacheKey && ctx?.lookupVectorCache) ctx.lookupVectorCache.set(cacheKey, values)
	if (cacheKey && exactIndex && ctx?.exactLookupCache)
		ctx.exactLookupCache.set(cacheKey, exactIndex)
	return {
		values,
		orientation,
		...(buildIndex && exactIndex ? { exactIndex } : {}),
	}
}

function buildExactLookupIndex(data: readonly CellValue[]): Map<string, ExactLookupHit> {
	const index = new Map<string, ExactLookupHit>()
	for (let i = 0; i < data.length; i++) {
		const key = exactLookupValueKey(v(data, i))
		if (key === null) continue
		const existing = index.get(key)
		if (existing !== undefined) index.set(key, packExactLookupHit(exactLookupHitFirst(existing), i))
		else index.set(key, packExactLookupHit(i, i))
	}
	return index
}

function getExactLookupIndex(
	arg: EvalArg | undefined,
	data: readonly CellValue[],
	orientation: 'row' | 'column',
	ctx?: FunctionEvalContext,
): ReadonlyMap<string, ExactLookupHit> {
	const cacheKey = lookupRangeCacheKey(arg, orientation)
	if (!cacheKey || !ctx?.exactLookupCache) return buildExactLookupIndex(data)
	const cached = ctx.exactLookupCache.get(cacheKey)
	if (cached) return cached
	const built = buildExactLookupIndex(data)
	ctx.exactLookupCache.set(cacheKey, built)
	return built
}

function indexedExactMatch(
	lookup: CellValue,
	index: ReadonlyMap<string, ExactLookupHit>,
	fromEnd = false,
): number {
	const key = exactLookupValueKey(lookup)
	if (key === null) return -1
	const hit = index.get(key)
	return hit !== undefined ? (fromEnd ? exactLookupHitLast(hit) : exactLookupHitFirst(hit)) : -1
}

function nextLargerMatch(lookup: CellValue, data: readonly CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		const probe = nearestNonBlankIndex(data, lo, hi, mid)
		if (probe < 0) break
		if (compareValues(v(data, probe), lookup) >= 0) {
			result = probe
			hi = probe - 1
		} else {
			lo = probe + 1
		}
	}
	return result
}

function binaryExactSearch(
	lookup: CellValue,
	data: readonly CellValue[],
	ascending: boolean,
): number {
	let lo = 0
	let hi = data.length - 1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		const midVal = v(data, mid)
		if (valuesEqual(lookup, midVal)) return mid
		const cmp = compareValues(lookup, midVal)
		if (ascending ? cmp > 0 : cmp < 0) lo = mid + 1
		else hi = mid - 1
	}
	return -1
}

function findInArray(
	lookup: CellValue,
	data: readonly CellValue[],
	matchMode: number,
	searchMode: number,
): number {
	if (matchMode === 2) {
		if (lookup.kind !== 'string') return -1
		const start = searchMode === -1 ? data.length - 1 : 0
		const step = searchMode === -1 ? -1 : 1
		for (let i = start; i >= 0 && i < data.length; i += step) {
			const cell = v(data, i)
			if (cell.kind === 'string' && wildcardMatch(lookup.value, cell.value)) return i
		}
		return -1
	}
	if (matchMode === 0) {
		if (Math.abs(searchMode) === 2) return binaryExactSearch(lookup, data, searchMode > 0)
		const start = searchMode === -1 ? data.length - 1 : 0
		const step = searchMode === -1 ? -1 : 1
		for (let i = start; i >= 0 && i < data.length; i += step) {
			if (valuesEqual(lookup, v(data, i))) return i
		}
		return -1
	}
	if (matchMode === -1)
		return searchMode === -2
			? descendingNextSmallerMatch(lookup, data)
			: approximateMatch(lookup, data)
	if (matchMode === 1)
		return searchMode === -2 ? reverseApproximateMatch(lookup, data) : nextLargerMatch(lookup, data)
	return -1
}

function isValidXlookupMatchMode(mode: number): boolean {
	return mode === -1 || mode === 0 || mode === 1 || mode === 2
}

function isValidXlookupSearchMode(mode: number): boolean {
	return mode === -2 || mode === -1 || mode === 1 || mode === 2
}

function resolveApproximate(v: CellValue): boolean {
	if (v.kind === 'boolean') return v.value
	if (v.kind === 'number') return v.value !== 0
	if (v.kind === 'empty') return true
	return true
}

function extractColumn(table: readonly (readonly CellValue[])[], col: number): CellValue[] {
	return table.map((r) => r[col] ?? EMPTY)
}

function lookupInputMatrix(
	arg: EvalArg | undefined,
): readonly (readonly CellValue[])[] | undefined {
	if (!arg) return undefined
	if (arg.kind === 'range' && arg.values) return arg.values
	if (arg.value.kind === 'array') return arg.value.rows
	return undefined
}

function scalarOrLookupArray(
	arg: EvalArg | undefined,
	resolve: (lookup: CellValue) => CellValue,
): CellValue {
	const matrix = lookupInputMatrix(arg)
	if (!matrix) return resolve(cellOf(arg))
	const rows: ScalarCellValue[][] = []
	for (const row of matrix) {
		const next: ScalarCellValue[] = []
		for (const value of row) {
			const resolved = resolve(topLeftScalar(value))
			if (resolved.kind === 'array') return errorValue('#VALUE!')
			next.push(topLeftScalar(resolved))
		}
		rows.push(next)
	}
	if (rows.length === 1 && (rows[0]?.length ?? 0) === 1) return rows[0]?.[0] ?? EMPTY
	return arrayValue(rows)
}

function scalarInputError(arg: EvalArg | undefined): CellValue | null {
	if (lookupInputMatrix(arg)) return null
	const value = arg?.value ?? EMPTY
	return value.kind === 'error' ? value : null
}

// --- Implementations ---

function vlookup(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const lookupError = scalarInputError(args[0])
	if (lookupError) return lookupError
	const tableError = scalarInputError(args[1])
	if (tableError) return tableError
	const table = getRange(args[1])
	const col = numArg(args[2])
	if (typeof col !== 'number') return col
	const colInt = Math.floor(col)
	if (table.length === 0) return errorValue('#REF!')

	const firstCol = getLookupVector(args[1], table, 'column', ctx)
	const approx = resolveApproximate(args.length > 3 ? cellOf(args[3]) : EMPTY)
	if (!approx && (colInt < 1 || colInt > (table[0]?.length ?? 0))) return errorValue('#REF!')
	const exactIndex = getExactLookupIndex(args[1], firstCol, 'column', ctx)
	return scalarOrLookupArray(args[0], (lookup) => {
		if (lookup.kind === 'error') return lookup
		const exactIdx =
			approx && isBlankLookupValue(lookup) ? -1 : indexedExactMatch(lookup, exactIndex)
		const idx = exactIdx >= 0 ? exactIdx : approx ? approximateMatch(lookup, firstCol) : -1
		if (idx < 0) return errorValue('#N/A')
		if (colInt < 1 || colInt > (table[0]?.length ?? 0)) return errorValue('#REF!')
		return table[idx]?.[colInt - 1] ?? EMPTY
	})
}

function hlookup(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const lookupError = scalarInputError(args[0])
	if (lookupError) return lookupError
	const tableError = scalarInputError(args[1])
	if (tableError) return tableError
	const table = getRange(args[1])
	const row = numArg(args[2])
	if (typeof row !== 'number') return row
	const rowInt = Math.floor(row)

	const firstRow = getLookupVector(args[1], table, 'row', ctx)
	if (firstRow.length === 0) return errorValue('#REF!')

	const approx = resolveApproximate(args.length > 3 ? cellOf(args[3]) : EMPTY)
	if (!approx && (rowInt < 1 || rowInt > table.length)) return errorValue('#REF!')
	const exactIndex = getExactLookupIndex(args[1], firstRow, 'row', ctx)
	return scalarOrLookupArray(args[0], (lookup) => {
		if (lookup.kind === 'error') return lookup
		const exactIdx =
			approx && isBlankLookupValue(lookup) ? -1 : indexedExactMatch(lookup, exactIndex)
		const idx = exactIdx >= 0 ? exactIdx : approx ? approximateMatch(lookup, firstRow) : -1
		if (idx < 0) return errorValue('#N/A')
		if (rowInt < 1 || rowInt > table.length) return errorValue('#REF!')
		return table[rowInt - 1]?.[idx] ?? EMPTY
	})
}

function indexFn(args: EvalArg[]): CellValue {
	const sourceError = scalarInputError(args[0])
	if (sourceError) return sourceError
	const array = getRange(args[0])
	const rowNum = numArg(args[1])
	if (typeof rowNum !== 'number') return rowNum
	const row = Math.floor(rowNum)

	if (args.length > 2) {
		const colNum = numArg(args[2])
		if (typeof colNum !== 'number') return colNum
		const col = Math.floor(colNum)
		if (row === 0 && col === 0) return errorValue('#VALUE!')
		if (row === 0) {
			if (col < 1 || col > (array[0]?.length ?? 0)) return errorValue('#REF!')
			return arrayValue(array.map((arrayRow) => [topLeftScalar(arrayRow[col - 1] ?? EMPTY)]))
		}
		if (col === 0) {
			if (row < 1 || row > array.length) return errorValue('#REF!')
			return arrayValue([(array[row - 1] ?? []).map((cell) => topLeftScalar(cell))])
		}
		if (row < 1 || row > array.length) return errorValue('#REF!')
		if (col < 1 || col > (array[0]?.length ?? 0)) return errorValue('#REF!')
		return array[row - 1]?.[col - 1] ?? EMPTY
	}

	if (array.length === 1) {
		const firstRow = array[0] ?? []
		if (row < 1 || row > firstRow.length) return errorValue('#REF!')
		return firstRow[row - 1] ?? EMPTY
	}
	if (row < 1 || row > array.length) return errorValue('#REF!')
	return array[row - 1]?.[0] ?? EMPTY
}

function matchFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const array = getRange(args[1])
	const matchType = args.length > 2 ? numArg(args[2]) : 1
	if (typeof matchType !== 'number') return matchType

	const flat = getLookupVector(args[1], array, array.length === 1 ? 'row' : 'column', ctx)
	const exactIndex =
		matchType === 0
			? getExactLookupIndex(args[1], flat, array.length === 1 ? 'row' : 'column', ctx)
			: null

	return scalarOrLookupArray(args[0], (lookup) => matchScalar(lookup, flat, matchType, exactIndex))
}

function matchScalar(
	lookup: CellValue,
	flat: readonly CellValue[],
	matchType: number,
	exactIndex?: ReadonlyMap<string, ExactLookupHit> | null,
): CellValue {
	if (lookup.kind === 'error') return lookup
	if (matchType === 0) {
		if (exactIndex && (lookup.kind !== 'string' || !hasWildcardPatternSyntax(lookup.value))) {
			const idx = indexedExactMatch(lookup, exactIndex)
			if (idx >= 0) return numberValue(idx + 1)
		}
		for (let i = 0; i < flat.length; i++) {
			const cell = v(flat, i)
			if (lookup.kind === 'string' && cell.kind === 'string') {
				if (wildcardMatch(lookup.value, cell.value)) return numberValue(i + 1)
			} else if (valuesEqual(lookup, cell)) {
				return numberValue(i + 1)
			}
		}
		return errorValue('#N/A')
	}
	if (matchType === 1) {
		const idx = approximateMatch(lookup, flat)
		return idx >= 0 ? numberValue(idx + 1) : errorValue('#N/A')
	}
	if (matchType === -1) {
		const idx = reverseApproximateMatch(lookup, flat)
		return idx >= 0 ? numberValue(idx + 1) : errorValue('#N/A')
	}
	return errorValue('#N/A')
}

function xlookup(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const ifNotFound = args.length > 3 ? cellOf(args[3]) : null
	const matchMode = args.length > 4 ? numArg(args[4]) : 0
	if (typeof matchMode !== 'number') return matchMode
	const searchMode = args.length > 5 ? numArg(args[5]) : 1
	if (typeof searchMode !== 'number') return searchMode
	if (!isValidXlookupMatchMode(matchMode) || !isValidXlookupSearchMode(searchMode)) {
		return errorValue('#VALUE!')
	}

	const needExactIndex = matchMode === 0 && Math.abs(searchMode) !== 2
	const lookupVector = getArgVectorWithIndex(args[1], ctx, needExactIndex)
	const returnVector = getArgVector(args[2], ctx)
	if (lookupVector && returnVector && lookupVector.orientation === returnVector.orientation) {
		const exactIndex = needExactIndex ? (lookupVector.exactIndex ?? null) : null
		return scalarOrLookupArray(args[0], (lookup) =>
			xlookupVectorScalar(
				lookup,
				lookupVector.values,
				returnVector.values,
				ifNotFound,
				matchMode,
				searchMode,
				exactIndex,
			),
		)
	}

	const lookupArray = getRange(args[1])
	const returnArray = getRange(args[2])

	let lookupFlat: readonly CellValue[]
	let vertical: boolean
	if (lookupArray.length === 1) {
		lookupFlat = getLookupVector(args[1], lookupArray, 'row', ctx)
		vertical = false
	} else {
		lookupFlat = getLookupVector(args[1], lookupArray, 'column', ctx)
		vertical = true
	}
	const exactIndex =
		matchMode === 0 && Math.abs(searchMode) !== 2
			? getExactLookupIndex(args[1], lookupFlat, vertical ? 'column' : 'row', ctx)
			: null

	return scalarOrLookupArray(args[0], (lookup) =>
		xlookupScalar(
			lookup,
			lookupFlat,
			returnArray,
			ifNotFound,
			matchMode,
			searchMode,
			vertical,
			exactIndex,
		),
	)
}

function xlookupVectorScalar(
	lookup: CellValue,
	lookupFlat: readonly CellValue[],
	returnFlat: readonly CellValue[],
	ifNotFound: CellValue | null,
	matchMode: number,
	searchMode: number,
	exactIndex?: ReadonlyMap<string, ExactLookupHit> | null,
): CellValue {
	if (lookup.kind === 'error') return lookup
	const idx =
		matchMode === 0 && Math.abs(searchMode) !== 2 && exactIndex
			? indexedExactMatch(lookup, exactIndex, searchMode === -1)
			: findInArray(lookup, lookupFlat, matchMode, searchMode)
	return idx < 0 ? (ifNotFound ?? errorValue('#N/A')) : (returnFlat[idx] ?? EMPTY)
}

function xlookupScalar(
	lookup: CellValue,
	lookupFlat: readonly CellValue[],
	returnArray: readonly (readonly CellValue[])[],
	ifNotFound: CellValue | null,
	matchMode: number,
	searchMode: number,
	vertical: boolean,
	exactIndex?: ReadonlyMap<string, ExactLookupHit> | null,
): CellValue {
	if (lookup.kind === 'error') return lookup
	const idx =
		matchMode === 0 && Math.abs(searchMode) !== 2 && exactIndex
			? indexedExactMatch(lookup, exactIndex, searchMode === -1)
			: findInArray(lookup, lookupFlat, matchMode, searchMode)
	if (idx < 0) return ifNotFound ?? errorValue('#N/A')
	if (vertical) {
		const row = returnArray[idx] ?? []
		if (row.length <= 1) return row[0] ?? EMPTY
		return arrayValue([row.map((cell) => topLeftScalar(cell))])
	}
	const rows: ScalarCellValue[][] = []
	for (const returnRow of returnArray) {
		rows.push([topLeftScalar(returnRow[idx] ?? EMPTY)])
	}
	return rows.length <= 1 ? (rows[0]?.[0] ?? EMPTY) : arrayValue(rows)
}

function xmatch(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const matchMode = args.length > 2 ? numArg(args[2]) : 0
	if (typeof matchMode !== 'number') return matchMode
	const searchMode = args.length > 3 ? numArg(args[3]) : 1
	if (typeof searchMode !== 'number') return searchMode
	if (!isValidXlookupMatchMode(matchMode) || !isValidXlookupSearchMode(searchMode)) {
		return errorValue('#VALUE!')
	}

	const lookupVector = getArgVector(args[1], ctx)
	if (lookupVector) {
		const exactIndex =
			matchMode === 0 && Math.abs(searchMode) !== 2
				? getExactLookupIndex(args[1], lookupVector.values, lookupVector.orientation, ctx)
				: null
		return scalarOrLookupArray(args[0], (lookup) => {
			if (lookup.kind === 'error') return lookup
			const idx =
				matchMode === 0 && Math.abs(searchMode) !== 2 && exactIndex
					? indexedExactMatch(lookup, exactIndex, searchMode === -1)
					: findInArray(lookup, lookupVector.values, matchMode, searchMode)
			return idx < 0 ? errorValue('#N/A') : numberValue(idx + 1)
		})
	}

	const lookupArray = getRange(args[1])
	const flat = getLookupVector(
		args[1],
		lookupArray,
		lookupArray.length === 1 ? 'row' : 'column',
		ctx,
	)
	const exactIndex =
		matchMode === 0 && Math.abs(searchMode) !== 2
			? getExactLookupIndex(args[1], flat, lookupArray.length === 1 ? 'row' : 'column', ctx)
			: null

	return scalarOrLookupArray(args[0], (lookup) => {
		if (lookup.kind === 'error') return lookup
		const idx =
			matchMode === 0 && Math.abs(searchMode) !== 2 && exactIndex
				? indexedExactMatch(lookup, exactIndex, searchMode === -1)
				: findInArray(lookup, flat, matchMode, searchMode)
		return idx < 0 ? errorValue('#N/A') : numberValue(idx + 1)
	})
}

function choose(args: EvalArg[]): CellValue {
	const idx = numArg(args[0])
	if (typeof idx !== 'number') return idx
	const n = Math.floor(idx)
	if (n < 1 || n >= args.length) return errorValue('#VALUE!')
	return cellOf(args[n])
}

function rowsFn(args: EvalArg[]): CellValue {
	const value = args[0]?.value
	if (value?.kind === 'error' && !args[0]?.ref && args[0]?.kind !== 'range') return value
	return numberValue(rangeShape(args[0]).rows)
}

function columnsFn(args: EvalArg[]): CellValue {
	const value = args[0]?.value
	if (value?.kind === 'error' && !args[0]?.ref && args[0]?.kind !== 'range') return value
	return numberValue(rangeShape(args[0]).cols)
}

function lookupFn(args: EvalArg[]): CellValue {
	const lookupArray = getRange(args[1])
	const resultArray = args.length > 2 ? getRange(args[2]) : lookupArray
	const lookupFlat: readonly CellValue[] =
		lookupArray.length === 1 ? [...(lookupArray[0] ?? [])] : extractColumn(lookupArray, 0)
	const resultFlat: readonly CellValue[] =
		resultArray.length === 1 ? [...(resultArray[0] ?? [])] : extractColumn(resultArray, 0)
	return scalarOrLookupArray(args[0], (lookup) => {
		if (lookup.kind === 'error') return lookup
		const exactIdx = exactMatch(lookup, lookupFlat)
		const idx = exactIdx >= 0 ? exactIdx : approximateMatch(lookup, lookupFlat)
		return idx < 0 ? errorValue('#N/A') : (resultFlat[idx] ?? EMPTY)
	})
}

function address(args: EvalArg[]): CellValue {
	const rowNum = numArg(args[0])
	if (typeof rowNum !== 'number') return rowNum
	const colNum = numArg(args[1])
	if (typeof colNum !== 'number') return colNum
	const absNum = args.length > 2 ? numArg(args[2]) : 1
	if (typeof absNum !== 'number') return absNum
	const useA1 = args.length > 3 ? cellOf(args[3]) : booleanValue(true)
	if (useA1.kind === 'error') return useA1
	const sheetText = args.length > 4 ? cellOf(args[4]) : null
	if (sheetText?.kind === 'error') return sheetText

	const row = Math.trunc(rowNum)
	const col = Math.trunc(colNum)
	if (row < 1 || col < 1) return errorValue('#VALUE!')

	const absoluteMode = Math.trunc(absNum)
	const absRow = absoluteMode === 1 || absoluteMode === 2
	const absCol = absoluteMode === 1 || absoluteMode === 3

	let ref: string
	if (useA1.kind === 'boolean' ? useA1.value : true) {
		const colLabel = toColumnLabel(col - 1)
		ref = `${absCol ? '$' : ''}${colLabel}${absRow ? '$' : ''}${row}`
	} else {
		ref = `R${absRow ? row : `[${row}]`}C${absCol ? col : `[${col}]`}`
	}

	if (sheetText && sheetText.kind === 'string' && sheetText.value.length > 0) {
		return { kind: 'string', value: `${formatSheetPrefix(sheetText.value)}!${ref}` }
	}
	return { kind: 'string', value: ref }
}

function formatSheetPrefix(sheetName: string): string {
	if (sheetName.startsWith('[') && /^(?:\[[^\]]+\])?(?:[A-Za-z_][A-Za-z0-9_.]*)$/.test(sheetName)) {
		return sheetName
	}
	if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) return `'${sheetName}'`
	if (/^(?:[A-Za-z_]|\[|\])(?:[A-Za-z0-9_.]|\[|\])*$/.test(sheetName)) return sheetName
	return `'${sheetName.replaceAll("'", "''")}'`
}

function toColumnLabel(colIndex: number): string {
	let n = colIndex + 1
	let label = ''
	while (n > 0) {
		const rem = (n - 1) % 26
		label = String.fromCharCode(65 + rem) + label
		n = Math.floor((n - 1) / 26)
	}
	return label
}

function formulaText(args: EvalArg[]): CellValue {
	const arg = args[0]
	if (!arg?.ref) {
		if (arg?.value.kind === 'error') return arg.value
		return errorValue('#N/A')
	}
	const formula = arg.formulaAtOffset?.(0, 0)
	return formula ? { kind: 'string', value: `=${formula}` } : errorValue('#N/A')
}

function areasFn(args: EvalArg[]): CellValue {
	const arg = args[0]
	if (!arg) return errorValue('#VALUE!')
	if (arg.areas?.length) return numberValue(arg.areas.length)
	if (arg.ref) return numberValue(1)
	if (arg.kind === 'range') return numberValue(1)
	return errorValue('#VALUE!')
}

function rowFn(args: EvalArg[]): CellValue {
	const ref = args[0]?.ref
	if (!ref) return errorValue('#VALUE!')
	const endRow = ref.endRow ?? ref.row
	if (endRow === ref.row) return numberValue(ref.row + 1)
	const rows: ScalarCellValue[][] = []
	for (let row = ref.row; row <= endRow; row++) rows.push([topLeftScalar(numberValue(row + 1))])
	return arrayValue(rows)
}

function columnFn(args: EvalArg[]): CellValue {
	const ref = args[0]?.ref
	if (!ref) return errorValue('#VALUE!')
	const endCol = ref.endCol ?? ref.col
	if (endCol === ref.col) return numberValue(ref.col + 1)
	const row: ScalarCellValue[] = []
	for (let col = ref.col; col <= endCol; col++) row.push(topLeftScalar(numberValue(col + 1)))
	return arrayValue([row])
}

export const lookupFunctions: FunctionDef[] = [
	{ name: 'VLOOKUP', minArgs: 3, maxArgs: 4, evaluate: vlookup },
	{ name: 'HLOOKUP', minArgs: 3, maxArgs: 4, evaluate: hlookup },
	{ name: 'INDEX', minArgs: 2, maxArgs: 3, evaluate: indexFn },
	{ name: 'MATCH', minArgs: 2, maxArgs: 3, evaluate: matchFn },
	{ name: 'XLOOKUP', minArgs: 3, maxArgs: 6, evaluate: xlookup },
	{ name: 'XMATCH', minArgs: 2, maxArgs: 4, evaluate: xmatch },
	{ name: 'CHOOSE', minArgs: 2, maxArgs: 255, evaluate: choose },
	{ name: 'LOOKUP', minArgs: 2, maxArgs: 3, evaluate: lookupFn },
	{ name: 'ADDRESS', minArgs: 2, maxArgs: 5, evaluate: address },
	{ name: 'ROWS', minArgs: 1, maxArgs: 1, evaluate: rowsFn },
	{ name: 'COLUMNS', minArgs: 1, maxArgs: 1, evaluate: columnsFn },
	{
		name: 'ROW',
		minArgs: 0,
		maxArgs: 1,
		evaluate: rowFn,
	},
	{
		name: 'COLUMN',
		minArgs: 0,
		maxArgs: 1,
		evaluate: columnFn,
	},
	{ name: 'FORMULATEXT', minArgs: 1, maxArgs: 1, volatile: false, evaluate: formulaText },
	{ name: 'AREAS', minArgs: 1, maxArgs: 1, evaluate: areasFn },
	// Stub: actual evaluation is bypassed by the engine evaluator (codegen)
	{ name: 'INDIRECT', minArgs: 1, maxArgs: 2, volatile: true, evaluate: () => errorValue('#REF!') },
	// Stub: actual evaluation is bypassed by the engine evaluator (codegen)
	{ name: 'OFFSET', minArgs: 3, maxArgs: 5, volatile: true, evaluate: () => errorValue('#REF!') },
]
