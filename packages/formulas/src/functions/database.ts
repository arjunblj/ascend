import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, isEmpty, numberValue, topLeftScalar } from '@ascend/schema'
import { numericVal } from './math/helpers.ts'
import type { EvalArg, FunctionDef } from './registry.ts'
import { iterAreaRows, wildcardMatch } from './registry.ts'

function isBlankLike(value: CellValue): boolean {
	if (isEmpty(value)) return true
	if (value.kind === 'string') return value.value === ''
	return false
}

function parseCriteria(criteria: CellValue): (v: CellValue) => boolean {
	if (criteria.kind === 'number') {
		const t = criteria.value
		return (v) => {
			const n = numericVal(v)
			return n !== null && n === t
		}
	}
	if (criteria.kind === 'boolean') {
		const t = criteria.value
		return (v) => v.kind === 'boolean' && v.value === t
	}
	if (criteria.kind !== 'string') return () => false

	const s = criteria.value
	let op = ''
	let rest = s
	for (const prefix of ['>=', '<=', '<>', '>', '<', '=']) {
		if (s.startsWith(prefix)) {
			op = prefix
			rest = s.slice(prefix.length)
			break
		}
	}

	const numRest = Number(rest)
	const isNumeric = rest.trim() !== '' && !Number.isNaN(numRest)
	const isBlankCriterion = rest === ''
	const hasWildcards = /(^|[^~])[*?]/.test(rest)

	if (!op) {
		const lower = s.toLowerCase()
		return (v) => {
			if (lower === '') return isBlankLike(v)
			if (v.kind === 'string') {
				return hasWildcards ? wildcardMatch(s, v.value) : v.value.toLowerCase() === lower
			}
			if (isNumeric && v.kind === 'number') return v.value === numRest
			if (v.kind === 'boolean') return v.value === (lower === 'true')
			return false
		}
	}

	if (isNumeric) {
		const cmp = (val: number): boolean => {
			switch (op) {
				case '>=':
					return val >= numRest
				case '<=':
					return val <= numRest
				case '>':
					return val > numRest
				case '<':
					return val < numRest
				case '<>':
					return val !== numRest
				case '=':
					return val === numRest
				default:
					return false
			}
		}
		return (v) => {
			const n = numericVal(v)
			return n !== null && cmp(n)
		}
	}

	const lower = rest.toLowerCase()
	return (v) => {
		switch (op) {
			case '=':
				if (isBlankCriterion) return isBlankLike(v)
				if (v.kind !== 'string') return false
				return hasWildcards ? wildcardMatch(rest, v.value) : v.value.toLowerCase() === lower
			case '<>':
				if (isBlankCriterion) return !isBlankLike(v)
				if (v.kind !== 'string') return false
				return hasWildcards ? !wildcardMatch(rest, v.value) : v.value.toLowerCase() !== lower
			default:
				return false
		}
	}
}

export interface DatabaseFilterResult {
	values: CellValue[]
	error?: CellValue
}

export function databaseFilter(args: EvalArg[]): DatabaseFilterResult {
	const dbRange = iterAreaRows(args[0])
	const fieldArg = args[1]
	const criteriaRange = iterAreaRows(args[2])

	if (dbRange.length < 2) return { values: [] }
	const headers = dbRange[0] ?? []
	const dataRows = dbRange.slice(1)

	let colIndex: number
	const fieldVal = fieldArg?.value ?? EMPTY
	const fieldScalar = topLeftScalar(fieldVal)
	if (fieldScalar.kind === 'number') {
		colIndex = Math.trunc(fieldScalar.value) - 1
		if (colIndex < 0 || colIndex >= headers.length)
			return { values: [], error: errorValue('#VALUE!') }
	} else if (fieldScalar.kind === 'string') {
		const name = fieldScalar.value.toLowerCase()
		const idx = headers.findIndex((h) => {
			const s = topLeftScalar(h)
			return s.kind === 'string' && s.value.toLowerCase() === name
		})
		if (idx < 0) return { values: [], error: errorValue('#VALUE!') }
		colIndex = idx
	} else {
		return { values: [], error: errorValue('#VALUE!') }
	}

	const criteriaHeaders = criteriaRange[0] ?? []
	const criteriaRows = criteriaRange.slice(1)
	const criteriaColToDbCol: number[] = []
	for (let c = 0; c < criteriaHeaders.length; c++) {
		const ch = topLeftScalar(criteriaHeaders[c] ?? EMPTY)
		const hStr = ch.kind === 'string' ? ch.value.toLowerCase() : ''
		const dbIdx = headers.findIndex((h) => {
			const s = topLeftScalar(h)
			return s.kind === 'string' && s.value.toLowerCase() === hStr
		})
		criteriaColToDbCol.push(dbIdx)
	}

	const values: CellValue[] = []
	for (let r = 0; r < dataRows.length; r++) {
		const dataRow = dataRows[r] ?? []
		let matchesAny = false
		for (const criteriaRow of criteriaRows) {
			let matchesAll = true
			for (let c = 0; c < criteriaColToDbCol.length; c++) {
				const dbCol = criteriaColToDbCol[c] ?? -1
				if (dbCol < 0) continue
				const criterion = criteriaRow[c] ?? EMPTY
				const dataCell = dataRow[dbCol] ?? EMPTY
				const match = parseCriteria(criterion)
				if (!match(dataCell)) {
					matchesAll = false
					break
				}
			}
			if (matchesAll) {
				matchesAny = true
				break
			}
		}
		if (matchesAny) {
			const cell = dataRow[colIndex] ?? EMPTY
			if (cell.kind === 'error') return { values: [], error: cell }
			values.push(cell)
		}
	}
	return { values }
}

function dsum(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let sum = 0
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) sum += n
	}
	return numberValue(sum)
}

function daverage(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let sum = 0
	let count = 0
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) {
			sum += n
			count++
		}
	}
	return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
}

function dcount(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let count = 0
	for (const v of values) {
		if (v.kind === 'number' || v.kind === 'date') count++
	}
	return numberValue(count)
}

function dcounta(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let count = 0
	for (const v of values) {
		if (!isEmpty(v)) count++
	}
	return numberValue(count)
}

function dmax(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let max = Number.NEGATIVE_INFINITY
	let found = false
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) {
			max = Math.max(max, n)
			found = true
		}
	}
	return numberValue(found ? max : 0)
}

function dmin(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let min = Number.POSITIVE_INFINITY
	let found = false
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) {
			min = Math.min(min, n)
			found = true
		}
	}
	return numberValue(found ? min : 0)
}

function dproduct(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	let product = 1
	let found = false
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) {
			product *= n
			found = true
		}
	}
	return numberValue(found ? product : 0)
}

function dget(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	if (values.length === 0) return errorValue('#VALUE!')
	if (values.length > 1) return errorValue('#VALUE!')
	return values[0] ?? EMPTY
}

function dstdev(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	const nums: number[] = []
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) nums.push(n)
	}
	const divisor = nums.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')
	const mean = nums.reduce((a, b) => a + b, 0) / nums.length
	const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(Math.sqrt(sumSq / divisor))
}

function dstdevp(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	const nums: number[] = []
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) nums.push(n)
	}
	if (nums.length === 0) return errorValue('#DIV/0!')
	const mean = nums.reduce((a, b) => a + b, 0) / nums.length
	const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(Math.sqrt(sumSq / nums.length))
}

function dvar(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	const nums: number[] = []
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) nums.push(n)
	}
	const divisor = nums.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')
	const mean = nums.reduce((a, b) => a + b, 0) / nums.length
	const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(sumSq / divisor)
}

function dvarp(args: EvalArg[]): CellValue {
	const { values, error } = databaseFilter(args)
	if (error) return error
	const nums: number[] = []
	for (const v of values) {
		const n = numericVal(v)
		if (n !== null) nums.push(n)
	}
	if (nums.length === 0) return errorValue('#DIV/0!')
	const mean = nums.reduce((a, b) => a + b, 0) / nums.length
	const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(sumSq / nums.length)
}

export const databaseFunctions: FunctionDef[] = [
	{ name: 'DSUM', minArgs: 3, maxArgs: 3, evaluate: dsum },
	{ name: 'DAVERAGE', minArgs: 3, maxArgs: 3, evaluate: daverage },
	{ name: 'DCOUNT', minArgs: 3, maxArgs: 3, evaluate: dcount },
	{ name: 'DCOUNTA', minArgs: 3, maxArgs: 3, evaluate: dcounta },
	{ name: 'DMAX', minArgs: 3, maxArgs: 3, evaluate: dmax },
	{ name: 'DMIN', minArgs: 3, maxArgs: 3, evaluate: dmin },
	{ name: 'DPRODUCT', minArgs: 3, maxArgs: 3, evaluate: dproduct },
	{ name: 'DGET', minArgs: 3, maxArgs: 3, evaluate: dget },
	{ name: 'DSTDEV', minArgs: 3, maxArgs: 3, evaluate: dstdev },
	{ name: 'DSTDEVP', minArgs: 3, maxArgs: 3, evaluate: dstdevp },
	{ name: 'DVAR', minArgs: 3, maxArgs: 3, evaluate: dvar },
	{ name: 'DVARP', minArgs: 3, maxArgs: 3, evaluate: dvarp },
]
