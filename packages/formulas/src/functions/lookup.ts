import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, numberValue } from '@ascend/schema'
import {
	cellOf,
	compareValues,
	type EvalArg,
	getRange,
	numArg,
	registerFunction,
	valuesEqual,
	wildcardMatch,
} from './registry.ts'

function v(data: CellValue[], i: number): CellValue {
	return data[i] ?? EMPTY
}

function approximateMatch(lookup: CellValue, data: CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		if (compareValues(lookup, v(data, mid)) >= 0) {
			result = mid
			lo = mid + 1
		} else {
			hi = mid - 1
		}
	}
	return result
}

function reverseApproximateMatch(lookup: CellValue, data: CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		if (compareValues(v(data, mid), lookup) >= 0) {
			result = mid
			lo = mid + 1
		} else {
			hi = mid - 1
		}
	}
	return result
}

function exactMatch(lookup: CellValue, data: CellValue[]): number {
	for (let i = 0; i < data.length; i++) {
		if (valuesEqual(lookup, v(data, i))) return i
	}
	return -1
}

function nextLargerMatch(lookup: CellValue, data: CellValue[]): number {
	let lo = 0
	let hi = data.length - 1
	let result = -1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		if (compareValues(v(data, mid), lookup) >= 0) {
			result = mid
			hi = mid - 1
		} else {
			lo = mid + 1
		}
	}
	return result
}

function binaryExactSearch(lookup: CellValue, data: CellValue[], ascending: boolean): number {
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
	data: CellValue[],
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
	if (matchMode === -1) return approximateMatch(lookup, data)
	if (matchMode === 1) return nextLargerMatch(lookup, data)
	return -1
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

// --- Implementations ---

function vlookup(args: EvalArg[]): CellValue {
	const lookup = cellOf(args[0])
	if (lookup.kind === 'error') return lookup
	const table = getRange(args[1])
	const col = numArg(args[2])
	if (typeof col !== 'number') return col
	const colInt = Math.floor(col)
	if (colInt < 1 || table.length === 0 || colInt > (table[0]?.length ?? 0))
		return errorValue('#REF!')

	const firstCol = extractColumn(table, 0)
	const approx = resolveApproximate(args.length > 3 ? cellOf(args[3]) : EMPTY)
	const idx = approx ? approximateMatch(lookup, firstCol) : exactMatch(lookup, firstCol)
	return idx < 0 ? errorValue('#N/A') : (table[idx]?.[colInt - 1] ?? EMPTY)
}

function hlookup(args: EvalArg[]): CellValue {
	const lookup = cellOf(args[0])
	if (lookup.kind === 'error') return lookup
	const table = getRange(args[1])
	const row = numArg(args[2])
	if (typeof row !== 'number') return row
	const rowInt = Math.floor(row)

	const firstRow = [...(table[0] ?? [])]
	if (rowInt < 1 || rowInt > table.length || firstRow.length === 0) return errorValue('#REF!')

	const approx = resolveApproximate(args.length > 3 ? cellOf(args[3]) : EMPTY)
	const idx = approx ? approximateMatch(lookup, firstRow) : exactMatch(lookup, firstRow)
	return idx < 0 ? errorValue('#N/A') : (table[rowInt - 1]?.[idx] ?? EMPTY)
}

function indexFn(args: EvalArg[]): CellValue {
	const array = getRange(args[0])
	const rowNum = numArg(args[1])
	if (typeof rowNum !== 'number') return rowNum
	const row = Math.floor(rowNum)

	if (args.length > 2) {
		const colNum = numArg(args[2])
		if (typeof colNum !== 'number') return colNum
		const col = Math.floor(colNum)
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

function matchFn(args: EvalArg[]): CellValue {
	const lookup = cellOf(args[0])
	if (lookup.kind === 'error') return lookup
	const array = getRange(args[1])
	const matchType = args.length > 2 ? numArg(args[2]) : 1
	if (typeof matchType !== 'number') return matchType

	const flat: CellValue[] = array.length === 1 ? [...(array[0] ?? [])] : extractColumn(array, 0)

	if (matchType === 0) {
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

function xlookup(args: EvalArg[]): CellValue {
	const lookup = cellOf(args[0])
	if (lookup.kind === 'error') return lookup
	const lookupArray = getRange(args[1])
	const returnArray = getRange(args[2])
	const ifNotFound = args.length > 3 ? cellOf(args[3]) : null
	const matchMode = args.length > 4 ? numArg(args[4]) : 0
	if (typeof matchMode !== 'number') return matchMode
	const searchMode = args.length > 5 ? numArg(args[5]) : 1
	if (typeof searchMode !== 'number') return searchMode

	let lookupFlat: CellValue[]
	let vertical: boolean
	if (lookupArray.length === 1) {
		lookupFlat = [...(lookupArray[0] ?? [])]
		vertical = false
	} else {
		lookupFlat = extractColumn(lookupArray, 0)
		vertical = true
	}

	const idx = findInArray(lookup, lookupFlat, matchMode, searchMode)
	if (idx < 0) return ifNotFound ?? errorValue('#N/A')
	return vertical ? (returnArray[idx]?.[0] ?? EMPTY) : (returnArray[0]?.[idx] ?? EMPTY)
}

function xmatch(args: EvalArg[]): CellValue {
	const lookup = cellOf(args[0])
	if (lookup.kind === 'error') return lookup
	const lookupArray = getRange(args[1])
	const matchMode = args.length > 2 ? numArg(args[2]) : 0
	if (typeof matchMode !== 'number') return matchMode
	const searchMode = args.length > 3 ? numArg(args[3]) : 1
	if (typeof searchMode !== 'number') return searchMode

	const flat: CellValue[] =
		lookupArray.length === 1 ? [...(lookupArray[0] ?? [])] : extractColumn(lookupArray, 0)

	const idx = findInArray(lookup, flat, matchMode, searchMode)
	return idx < 0 ? errorValue('#N/A') : numberValue(idx + 1)
}

function choose(args: EvalArg[]): CellValue {
	const idx = numArg(args[0])
	if (typeof idx !== 'number') return idx
	const n = Math.floor(idx)
	if (n < 1 || n >= args.length) return errorValue('#VALUE!')
	return cellOf(args[n])
}

function rowsFn(args: EvalArg[]): CellValue {
	return numberValue(getRange(args[0]).length)
}

function columnsFn(args: EvalArg[]): CellValue {
	return numberValue(getRange(args[0])[0]?.length ?? 0)
}

// --- Registration ---

registerFunction({ name: 'VLOOKUP', minArgs: 3, maxArgs: 4, evaluate: vlookup })
registerFunction({ name: 'HLOOKUP', minArgs: 3, maxArgs: 4, evaluate: hlookup })
registerFunction({ name: 'INDEX', minArgs: 2, maxArgs: 3, evaluate: indexFn })
registerFunction({ name: 'MATCH', minArgs: 2, maxArgs: 3, evaluate: matchFn })
registerFunction({
	name: 'XLOOKUP',
	minArgs: 3,
	maxArgs: 6,
	evaluate: xlookup,
})
registerFunction({ name: 'XMATCH', minArgs: 2, maxArgs: 4, evaluate: xmatch })
registerFunction({ name: 'CHOOSE', minArgs: 2, maxArgs: 255, evaluate: choose })
registerFunction({ name: 'ROWS', minArgs: 1, maxArgs: 1, evaluate: rowsFn })
registerFunction({
	name: 'COLUMNS',
	minArgs: 1,
	maxArgs: 1,
	evaluate: columnsFn,
})

registerFunction({
	name: 'ROW',
	minArgs: 0,
	maxArgs: 1,
	evaluate: () => errorValue('#REF!'),
})
registerFunction({
	name: 'COLUMN',
	minArgs: 0,
	maxArgs: 1,
	evaluate: () => errorValue('#REF!'),
})
registerFunction({
	name: 'INDIRECT',
	minArgs: 1,
	maxArgs: 2,
	evaluate: () => errorValue('#REF!'),
})
registerFunction({
	name: 'OFFSET',
	minArgs: 3,
	maxArgs: 5,
	evaluate: () => errorValue('#REF!'),
})
