import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, isEmpty, isError, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef, FunctionEvalContext } from './index.ts'

function fn(
	name: string,
	minArgs: number,
	maxArgs: number,
	evaluate: (args: EvalArg[], ctx?: FunctionEvalContext) => CellValue,
	volatile = false,
): FunctionDef {
	return { name, minArgs, maxArgs, volatile, evaluate }
}
function seededRandom(ctx?: FunctionEvalContext): number {
	const seed = ctx?.randomSeed ?? 42
	const row = ctx?.row ?? 0
	const col = ctx?.col ?? 0
	const sheet = ctx?.sheetIndex ?? 0
	let state =
		(seed ^ ((sheet + 1) * 0x9e3779b1) ^ ((row + 1) * 0x85ebca6b) ^ ((col + 1) * 0xc2b2ae35)) >>> 0
	state ^= state >>> 16
	state = Math.imul(state, 0x7feb352d) >>> 0
	state ^= state >>> 15
	state = Math.imul(state, 0x846ca68b) >>> 0
	state ^= state >>> 16
	return state / 0x1_0000_0000
}

function toNum(v: CellValue): number | CellValue {
	switch (v.kind) {
		case 'empty':
			return 0
		case 'number':
			return v.value
		case 'string': {
			if (v.value.trim() === '') return 0
			const n = Number(v.value)
			return Number.isNaN(n) ? errorValue('#VALUE!') : n
		}
		case 'boolean':
			return v.value ? 1 : 0
		case 'error':
			return v
		case 'date':
			return v.serial
		case 'richText':
			return errorValue('#VALUE!')
	}
}

function numArg(arg: EvalArg | undefined): number | CellValue {
	return toNum(arg?.value ?? EMPTY)
}

function getRange(arg: EvalArg | undefined): readonly (readonly CellValue[])[] {
	if (arg?.kind === 'range' && arg.values) return arg.values
	return [[arg?.value ?? EMPTY]]
}

function numericVal(cell: CellValue): number | null {
	if (cell.kind === 'number') return cell.value
	if (cell.kind === 'date') return cell.serial
	return null
}

// ---------------------------------------------------------------------------
// Criteria matching for SUMIF / COUNTIF / etc.
// Supports: ">=10", "<=5", "<>0", ">3", "<3", "=text", exact match
// ---------------------------------------------------------------------------

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

	if (!op) {
		const lower = s.toLowerCase()
		return (v) => {
			if (v.kind === 'string') return v.value.toLowerCase() === lower
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
		if (v.kind !== 'string') return false
		const vl = v.value.toLowerCase()
		switch (op) {
			case '=':
				return vl === lower
			case '<>':
				return vl !== lower
			default:
				return false
		}
	}
}

// ---------------------------------------------------------------------------
// Conditional-aggregate helper used by SUMIFS, AVERAGEIFS, MINIFS, MAXIFS
// ---------------------------------------------------------------------------

interface CriteriaPair {
	range: readonly (readonly CellValue[])[]
	match: (v: CellValue) => boolean
}

function buildPairs(args: EvalArg[], startIdx: number): CriteriaPair[] {
	const pairs: CriteriaPair[] = []
	for (let i = startIdx; i + 1 < args.length; i += 2) {
		pairs.push({
			range: getRange(args[i]),
			match: parseCriteria(args[i + 1]?.value ?? EMPTY),
		})
	}
	return pairs
}

function allPairsMatch(pairs: CriteriaPair[], r: number, c: number): boolean {
	for (const p of pairs) {
		if (!p.match(p.range[r]?.[c] ?? EMPTY)) return false
	}
	return true
}

// ---------------------------------------------------------------------------
// Function definitions
// ---------------------------------------------------------------------------

export const mathFunctions: FunctionDef[] = [
	fn('SUM', 1, 255, (args) => {
		let sum = 0
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						const n = numericVal(cell)
						if (n !== null) sum += n
					}
				}
			} else {
				const n = toNum(arg.value ?? EMPTY)
				if (typeof n !== 'number') return n
				sum += n
			}
		}
		return numberValue(sum)
	}),

	fn('SUMPRODUCT', 1, 255, (args) => {
		const ranges = args.map(getRange)
		const rows = ranges[0]?.length ?? 0
		const cols = ranges[0]?.[0]?.length ?? 0
		let total = 0
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				let product = 1
				for (const range of ranges) {
					const cell = range[r]?.[c] ?? EMPTY
					if (isError(cell)) return cell
					const n = numericVal(cell)
					product *= n ?? (cell.kind === 'boolean' ? (cell.value ? 1 : 0) : 0)
				}
				total += product
			}
		}
		return numberValue(total)
	}),

	fn('AVERAGE', 1, 255, (args) => {
		let sum = 0
		let count = 0
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						const n = numericVal(cell)
						if (n !== null) {
							sum += n
							count++
						}
					}
				}
			} else {
				const n = toNum(arg.value ?? EMPTY)
				if (typeof n !== 'number') return n
				sum += n
				count++
			}
		}
		return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
	}),

	fn('COUNT', 1, 255, (args) => {
		let count = 0
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (cell.kind === 'number' || cell.kind === 'date') count++
					}
				}
			} else {
				const v = arg.value ?? EMPTY
				if (v.kind === 'number' || v.kind === 'date' || v.kind === 'boolean') count++
			}
		}
		return numberValue(count)
	}),

	fn('COUNTA', 1, 255, (args) => {
		let count = 0
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (!isEmpty(cell)) count++
					}
				}
			} else {
				if (!isEmpty(arg.value ?? EMPTY)) count++
			}
		}
		return numberValue(count)
	}),

	fn('COUNTBLANK', 1, 1, (args) => {
		let count = 0
		for (const row of getRange(args[0])) {
			for (const cell of row) {
				if (isEmpty(cell) || (cell.kind === 'string' && cell.value === '')) count++
			}
		}
		return numberValue(count)
	}),

	fn('MIN', 1, 255, (args) => {
		let min = Number.POSITIVE_INFINITY
		let found = false
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						const n = numericVal(cell)
						if (n !== null) {
							min = Math.min(min, n)
							found = true
						}
					}
				}
			} else {
				const n = toNum(arg.value ?? EMPTY)
				if (typeof n !== 'number') return n
				min = Math.min(min, n)
				found = true
			}
		}
		return numberValue(found ? min : 0)
	}),

	fn('MAX', 1, 255, (args) => {
		let max = Number.NEGATIVE_INFINITY
		let found = false
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						const n = numericVal(cell)
						if (n !== null) {
							max = Math.max(max, n)
							found = true
						}
					}
				}
			} else {
				const n = toNum(arg.value ?? EMPTY)
				if (typeof n !== 'number') return n
				max = Math.max(max, n)
				found = true
			}
		}
		return numberValue(found ? max : 0)
	}),

	fn('ABS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.abs(n))
	}),

	fn('ROUND', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue((Math.sign(n) * Math.round(Math.abs(n) * factor)) / factor)
	}),

	fn('ROUNDUP', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue((Math.sign(n) * Math.ceil(Math.abs(n) * factor)) / factor)
	}),

	fn('ROUNDDOWN', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue((Math.sign(n) * Math.floor(Math.abs(n) * factor)) / factor)
	}),

	fn('INT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.floor(n))
	}),

	fn('MOD', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		if (d === 0) return errorValue('#DIV/0!')
		return numberValue(n - d * Math.floor(n / d))
	}),

	fn('POWER', 2, 2, (args) => {
		const base = numArg(args[0])
		if (typeof base !== 'number') return base
		const exp = numArg(args[1])
		if (typeof exp !== 'number') return exp
		const result = base ** exp
		if (!Number.isFinite(result)) return errorValue('#NUM!')
		return numberValue(result)
	}),

	fn('SQRT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < 0) return errorValue('#NUM!')
		return numberValue(Math.sqrt(n))
	}),

	fn('LOG', 1, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= 0) return errorValue('#NUM!')
		const base = args.length >= 2 ? numArg(args[1]) : 10
		if (typeof base !== 'number') return base
		if (base <= 0 || base === 1) return errorValue('#NUM!')
		return numberValue(Math.log(n) / Math.log(base))
	}),

	fn('LOG10', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= 0) return errorValue('#NUM!')
		return numberValue(Math.log10(n))
	}),

	fn('LN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= 0) return errorValue('#NUM!')
		return numberValue(Math.log(n))
	}),

	fn('EXP', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.exp(n))
	}),

	fn('CEILING', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const sig = numArg(args[1])
		if (typeof sig !== 'number') return sig
		if (sig === 0) return numberValue(0)
		if (n > 0 && sig < 0) return errorValue('#NUM!')
		return numberValue(Math.ceil(n / sig) * sig)
	}),

	fn('FLOOR', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const sig = numArg(args[1])
		if (typeof sig !== 'number') return sig
		if (sig === 0) return errorValue('#DIV/0!')
		if (n > 0 && sig < 0) return errorValue('#NUM!')
		return numberValue(Math.floor(n / sig) * sig)
	}),

	fn('SIGN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.sign(n))
	}),

	fn('RAND', 0, 0, (_args, ctx) => numberValue(seededRandom(ctx)), true),

	fn(
		'RANDBETWEEN',
		2,
		2,
		(args, ctx) => {
			const lo = numArg(args[0])
			if (typeof lo !== 'number') return lo
			const hi = numArg(args[1])
			if (typeof hi !== 'number') return hi
			const bottom = Math.ceil(lo)
			const top = Math.floor(hi)
			if (bottom > top) return errorValue('#NUM!')
			return numberValue(Math.floor(seededRandom(ctx) * (top - bottom + 1)) + bottom)
		},
		true,
	),

	fn('PI', 0, 0, () => numberValue(Math.PI)),

	fn('PRODUCT', 1, 255, (args) => {
		let product = 1
		let found = false
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						const n = numericVal(cell)
						if (n !== null) {
							product *= n
							found = true
						}
					}
				}
			} else {
				const n = toNum(arg.value ?? EMPTY)
				if (typeof n !== 'number') return n
				product *= n
				found = true
			}
		}
		return numberValue(found ? product : 0)
	}),

	fn('SUBTOTAL', 2, 255, subtotalFn),
]

function subtotalNums(data: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of data) {
		if (arg.kind === 'range' && arg.values) {
			for (const row of arg.values) {
				for (const cell of row) {
					if (isError(cell)) return cell
					const n = numericVal(cell)
					if (n !== null) nums.push(n)
				}
			}
		} else {
			const n = toNum(arg.value ?? EMPTY)
			if (typeof n !== 'number') return n
			nums.push(n)
		}
	}
	return nums
}

function subtotalDelegated(code: number, data: EvalArg[]): CellValue {
	return subtotalFn([{ value: numberValue(code) }, ...data])
}

function subtotalFn(args: EvalArg[]): CellValue {
	const fnNum = numArg(args[0])
	if (typeof fnNum !== 'number') return fnNum
	const code = Math.trunc(fnNum)
	const data = args.slice(1)
	switch (code) {
		case 1: {
			let sum = 0
			let count = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) {
								sum += n
								count++
							}
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					sum += n
					count++
				}
			}
			return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
		}
		case 2: {
			let count = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (cell.kind === 'number' || cell.kind === 'date') count++
						}
					}
				} else {
					const v = arg.value ?? EMPTY
					if (v.kind === 'number' || v.kind === 'date' || v.kind === 'boolean') count++
				}
			}
			return numberValue(count)
		}
		case 3: {
			let count = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (!isEmpty(cell)) count++
						}
					}
				} else if (!isEmpty(arg.value ?? EMPTY)) {
					count++
				}
			}
			return numberValue(count)
		}
		case 4: {
			let max = Number.NEGATIVE_INFINITY
			let found = false
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) {
								max = Math.max(max, n)
								found = true
							}
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					max = Math.max(max, n)
					found = true
				}
			}
			return numberValue(found ? max : 0)
		}
		case 5: {
			let min = Number.POSITIVE_INFINITY
			let found = false
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) {
								min = Math.min(min, n)
								found = true
							}
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					min = Math.min(min, n)
					found = true
				}
			}
			return numberValue(found ? min : 0)
		}
		case 6:
		case 106: {
			let product = 1
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) product *= n
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					product *= n
				}
			}
			return numberValue(product)
		}
		case 7:
		case 107: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			if (nums.length < 2) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
			return numberValue(Math.sqrt(sumSq / (nums.length - 1)))
		}
		case 8:
		case 108: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			if (nums.length < 2) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
			return numberValue(sumSq / (nums.length - 1))
		}
		case 9:
		case 109: {
			let sum = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) sum += n
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					sum += n
				}
			}
			return numberValue(sum)
		}
		case 101:
			return subtotalDelegated(1, data)
		case 102:
			return subtotalDelegated(2, data)
		case 103:
			return subtotalDelegated(3, data)
		case 104:
			return subtotalDelegated(4, data)
		case 105:
			return subtotalDelegated(5, data)
		case 110:
		case 111:
			return errorValue('#VALUE!')
		default:
			return errorValue('#VALUE!')
	}
}

const conditionalFunctions: FunctionDef[] = [
	fn('SUMIF', 2, 3, (args) => {
		const range = getRange(args[0])
		const match = parseCriteria(args[1]?.value ?? EMPTY)
		const sumRange = args.length >= 3 ? getRange(args[2]) : range
		let sum = 0
		for (let r = 0; r < range.length; r++) {
			for (let c = 0; c < (range[r]?.length ?? 0); c++) {
				if (match(range[r]?.[c] ?? EMPTY)) {
					const n = numericVal(sumRange[r]?.[c] ?? EMPTY)
					if (n !== null) sum += n
				}
			}
		}
		return numberValue(sum)
	}),

	fn('SUMIFS', 3, 255, (args) => {
		const sumRange = getRange(args[0])
		const pairs = buildPairs(args, 1)
		let sum = 0
		for (let r = 0; r < sumRange.length; r++) {
			for (let c = 0; c < (sumRange[r]?.length ?? 0); c++) {
				if (allPairsMatch(pairs, r, c)) {
					const n = numericVal(sumRange[r]?.[c] ?? EMPTY)
					if (n !== null) sum += n
				}
			}
		}
		return numberValue(sum)
	}),

	fn('COUNTIF', 2, 2, (args) => {
		const range = getRange(args[0])
		const match = parseCriteria(args[1]?.value ?? EMPTY)
		let count = 0
		for (const row of range) {
			for (const cell of row) {
				if (match(cell)) count++
			}
		}
		return numberValue(count)
	}),

	fn('COUNTIFS', 2, 255, (args) => {
		const pairs = buildPairs(args, 0)
		const first = pairs[0]?.range
		if (!first) return numberValue(0)
		let count = 0
		for (let r = 0; r < first.length; r++) {
			for (let c = 0; c < (first[r]?.length ?? 0); c++) {
				if (allPairsMatch(pairs, r, c)) count++
			}
		}
		return numberValue(count)
	}),

	fn('AVERAGEIF', 2, 3, (args) => {
		const range = getRange(args[0])
		const match = parseCriteria(args[1]?.value ?? EMPTY)
		const avgRange = args.length >= 3 ? getRange(args[2]) : range
		let sum = 0
		let count = 0
		for (let r = 0; r < range.length; r++) {
			for (let c = 0; c < (range[r]?.length ?? 0); c++) {
				if (match(range[r]?.[c] ?? EMPTY)) {
					const n = numericVal(avgRange[r]?.[c] ?? EMPTY)
					if (n !== null) {
						sum += n
						count++
					}
				}
			}
		}
		return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
	}),

	fn('AVERAGEIFS', 3, 255, (args) => {
		const avgRange = getRange(args[0])
		const pairs = buildPairs(args, 1)
		let sum = 0
		let count = 0
		for (let r = 0; r < avgRange.length; r++) {
			for (let c = 0; c < (avgRange[r]?.length ?? 0); c++) {
				if (allPairsMatch(pairs, r, c)) {
					const n = numericVal(avgRange[r]?.[c] ?? EMPTY)
					if (n !== null) {
						sum += n
						count++
					}
				}
			}
		}
		return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
	}),

	fn('MINIFS', 3, 255, (args) => {
		const minRange = getRange(args[0])
		const pairs = buildPairs(args, 1)
		let min = Number.POSITIVE_INFINITY
		let found = false
		for (let r = 0; r < minRange.length; r++) {
			for (let c = 0; c < (minRange[r]?.length ?? 0); c++) {
				if (allPairsMatch(pairs, r, c)) {
					const n = numericVal(minRange[r]?.[c] ?? EMPTY)
					if (n !== null) {
						min = Math.min(min, n)
						found = true
					}
				}
			}
		}
		return numberValue(found ? min : 0)
	}),

	fn('MAXIFS', 3, 255, (args) => {
		const maxRange = getRange(args[0])
		const pairs = buildPairs(args, 1)
		let max = Number.NEGATIVE_INFINITY
		let found = false
		for (let r = 0; r < maxRange.length; r++) {
			for (let c = 0; c < (maxRange[r]?.length ?? 0); c++) {
				if (allPairsMatch(pairs, r, c)) {
					const n = numericVal(maxRange[r]?.[c] ?? EMPTY)
					if (n !== null) {
						max = Math.max(max, n)
						found = true
					}
				}
			}
		}
		return numberValue(found ? max : 0)
	}),
]

mathFunctions.push(...conditionalFunctions)
