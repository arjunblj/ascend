import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, isEmpty, isError, numberValue, topLeftScalar } from '@ascend/schema'
import type { EvalArg, FunctionDef, FunctionEvalContext } from './index.ts'
import { iterAreaRows, wildcardMatch } from './registry.ts'

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
	return (state >>> 0) / 0x1_0000_0000
}

function toNum(v: CellValue): number | CellValue {
	v = topLeftScalar(v)
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
	return iterAreaRows(arg)
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

const CRITERIA_CACHE_MAX = 1024
const criteriaPredicateCache = new Map<string, (v: CellValue) => boolean>()

function criteriaCacheKey(criteria: CellValue): string {
	return JSON.stringify(criteria)
}

function parseCriteria(criteria: CellValue): (v: CellValue) => boolean {
	const key = criteriaCacheKey(criteria)
	const cached = criteriaPredicateCache.get(key)
	if (cached) return cached

	const predicate = parseCriteriaImpl(criteria)
	if (criteriaPredicateCache.size >= CRITERIA_CACHE_MAX) criteriaPredicateCache.clear()
	criteriaPredicateCache.set(key, predicate)
	return predicate
}

function parseCriteriaImpl(criteria: CellValue): (v: CellValue) => boolean {
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

function isBlankLike(value: CellValue): boolean {
	if (isEmpty(value)) return true
	if (value.kind === 'string') return value.value === ''
	return false
}

// ---------------------------------------------------------------------------
// Conditional-aggregate helper used by SUMIFS, AVERAGEIFS, MINIFS, MAXIFS
// ---------------------------------------------------------------------------

interface CriteriaPair {
	range: readonly (readonly CellValue[])[]
	match: (v: CellValue) => boolean
}

function sameShape(
	left: readonly (readonly CellValue[])[],
	right: readonly (readonly CellValue[])[],
): boolean {
	if (left.length !== right.length) return false
	for (let row = 0; row < left.length; row++) {
		if ((left[row]?.length ?? 0) !== (right[row]?.length ?? 0)) return false
	}
	return true
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
			if (arg.forEachValue) {
				let err: CellValue | undefined
				arg.forEachValue((cell) => {
					if (err) return
					if (isError(cell)) {
						err = cell
						return
					}
					const n = numericVal(cell)
					if (n !== null) sum += n
				})
				if (err) return err
			} else if (arg.kind === 'range' && arg.values) {
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
		for (let i = 1; i < ranges.length; i++) {
			if (!sameShape(ranges[0] ?? [], ranges[i] ?? [])) return errorValue('#VALUE!')
		}
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
			if (arg.forEachValue) {
				let err: CellValue | undefined
				arg.forEachValue((cell) => {
					if (err) return
					if (isError(cell)) {
						err = cell
						return
					}
					const n = numericVal(cell)
					if (n !== null) {
						sum += n
						count++
					}
				})
				if (err) return err
			} else if (arg.kind === 'range' && arg.values) {
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
			if (arg.forEachValue) {
				arg.forEachValue((cell) => {
					if (cell.kind === 'number' || cell.kind === 'date') count++
				})
			} else if (arg.kind === 'range' && arg.values) {
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
			if (arg.forEachValue) {
				arg.forEachValue((cell) => {
					if (!isEmpty(cell)) count++
				})
			} else if (arg.kind === 'range' && arg.values) {
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
		const arg = args[0]
		if (arg?.forEachValue) {
			arg.forEachValue((cell) => {
				if (isEmpty(cell) || (cell.kind === 'string' && cell.value === '')) count++
			})
		} else {
			for (const row of getRange(arg)) {
				for (const cell of row) {
					if (isEmpty(cell) || (cell.kind === 'string' && cell.value === '')) count++
				}
			}
		}
		return numberValue(count)
	}),

	fn('MIN', 1, 255, (args) => {
		let min = Number.POSITIVE_INFINITY
		let found = false
		for (const arg of args) {
			if (arg.forEachValue) {
				let err: CellValue | undefined
				arg.forEachValue((cell) => {
					if (err) return
					if (isError(cell)) {
						err = cell
						return
					}
					const n = numericVal(cell)
					if (n !== null) {
						min = Math.min(min, n)
						found = true
					}
				})
				if (err) return err
			} else if (arg.kind === 'range' && arg.values) {
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
			if (arg.forEachValue) {
				let err: CellValue | undefined
				arg.forEachValue((cell) => {
					if (err) return
					if (isError(cell)) {
						err = cell
						return
					}
					const n = numericVal(cell)
					if (n !== null) {
						max = Math.max(max, n)
						found = true
					}
				})
				if (err) return err
			} else if (arg.kind === 'range' && arg.values) {
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
			if (arg.forEachValue) {
				let err: CellValue | undefined
				arg.forEachValue((cell) => {
					if (err) return
					if (isError(cell)) {
						err = cell
						return
					}
					const n = numericVal(cell)
					if (n !== null) {
						product *= n
						found = true
					}
				})
				if (err) return err
			} else if (arg.kind === 'range' && arg.values) {
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

	fn('AGGREGATE', 3, 253, aggregateFn),

	fn('TRUNC', 1, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = args.length >= 2 ? numArg(args[1]) : 0
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue(Math.trunc(n * factor) / factor)
	}),

	fn('MROUND', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const m = numArg(args[1])
		if (typeof m !== 'number') return m
		if (n > 0 && m < 0) return errorValue('#NUM!')
		if (n < 0 && m > 0) return errorValue('#NUM!')
		if (m === 0) return numberValue(0)
		return numberValue(Math.round(n / m) * m)
	}),

	fn('GCD', 1, 255, (args) => {
		const nums: number[] = []
		for (const arg of args) {
			const n = numArg(arg)
			if (typeof n !== 'number') return n
			nums.push(Math.trunc(Math.abs(n)))
		}
		if (nums.length === 0) return numberValue(0)
		let g = nums[0] ?? 0
		for (let i = 1; i < nums.length; i++) {
			let a = g
			let b = nums[i] ?? 0
			while (b !== 0) {
				const t = b
				b = a % b
				a = t
			}
			g = a
		}
		return numberValue(g)
	}),

	fn('LCM', 1, 255, (args) => {
		const nums: number[] = []
		for (const arg of args) {
			const n = numArg(arg)
			if (typeof n !== 'number') return n
			nums.push(Math.trunc(Math.abs(n)))
		}
		if (nums.length === 0) return numberValue(0)
		const gcd = (a: number, b: number): number => (b === 0 ? a : gcd(b, a % b))
		let l = nums[0] ?? 0
		for (let i = 1; i < nums.length; i++) {
			const n = nums[i] ?? 0
			if (l === 0 || n === 0) {
				l = 0
				break
			}
			l = Math.abs(l * n) / gcd(l, n)
		}
		return numberValue(l)
	}),

	fn('FACT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = Math.trunc(n)
		if (k < 0) return errorValue('#NUM!')
		let f = 1
		for (let i = 2; i <= k; i++) f *= i
		return numberValue(f)
	}),

	fn('FACTDOUBLE', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = Math.trunc(n)
		if (k < 0) return errorValue('#NUM!')
		let f = 1
		for (let i = k; i > 0; i -= 2) f *= i
		return numberValue(f)
	}),

	fn('COMBIN', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = numArg(args[1])
		if (typeof k !== 'number') return k
		const ni = Math.trunc(n)
		const ki = Math.trunc(k)
		if (ni < 0 || ki < 0 || ki > ni) return errorValue('#NUM!')
		if (ki === 0 || ki === ni) return numberValue(1)
		let c = 1
		for (let i = 0; i < ki; i++) c = (c * (ni - i)) / (i + 1)
		return numberValue(Math.round(c))
	}),

	fn('PERMUT', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = numArg(args[1])
		if (typeof k !== 'number') return k
		const ni = Math.trunc(n)
		const ki = Math.trunc(k)
		if (ni < 0 || ki < 0 || ki > ni) return errorValue('#NUM!')
		let p = 1
		for (let i = 0; i < ki; i++) p *= ni - i
		return numberValue(p)
	}),

	fn('ODD', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const x = Math.ceil(Math.abs(n))
		const odd = x % 2 === 0 ? x + 1 : x
		return numberValue(n >= 0 ? odd : -odd)
	}),

	fn('EVEN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const x = Math.ceil(Math.abs(n))
		const even = x % 2 === 1 ? x + 1 : x
		return numberValue(n >= 0 ? even : -even)
	}),

	fn('QUOTIENT', 2, 2, (args) => {
		const num = numArg(args[0])
		if (typeof num !== 'number') return num
		const den = numArg(args[1])
		if (typeof den !== 'number') return den
		if (den === 0) return errorValue('#DIV/0!')
		return numberValue(Math.trunc(num / den))
	}),

	fn('DEGREES', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue((n * 180) / Math.PI)
	}),

	fn('RADIANS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue((n * Math.PI) / 180)
	}),

	fn('SIN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.sin(n))
	}),

	fn('COS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.cos(n))
	}),

	fn('TAN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.tan(n))
	}),

	fn('ASIN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < -1 || n > 1) return errorValue('#NUM!')
		return numberValue(Math.asin(n))
	}),

	fn('ACOS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < -1 || n > 1) return errorValue('#NUM!')
		return numberValue(Math.acos(n))
	}),

	fn('ATAN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.atan(n))
	}),

	fn('ATAN2', 2, 2, (args) => {
		const x = numArg(args[0])
		if (typeof x !== 'number') return x
		const y = numArg(args[1])
		if (typeof y !== 'number') return y
		if (x === 0 && y === 0) return errorValue('#DIV/0!')
		return numberValue(Math.atan2(y, x))
	}),

	fn('SINH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.sinh(n))
	}),

	fn('COSH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.cosh(n))
	}),

	fn('TANH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.tanh(n))
	}),

	fn('ASINH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.asinh(n))
	}),

	fn('ACOSH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < 1) return errorValue('#NUM!')
		return numberValue(Math.acosh(n))
	}),

	fn('ATANH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= -1 || n >= 1) return errorValue('#NUM!')
		return numberValue(Math.atanh(n))
	}),
]

function aggregateCollectNumbers(args: EvalArg[], ignoreErrors: boolean): number[] | CellValue {
	const nums: number[] = []
	for (const arg of args) {
		if (arg.forEachValue) {
			let err: CellValue | undefined
			arg.forEachValue((cell) => {
				if (err && !ignoreErrors) return
				if (cell.kind === 'error') {
					if (!ignoreErrors) err = cell
					return
				}
				const n = numericVal(cell)
				if (n !== null) nums.push(n)
			})
			if (err) return err
		} else if (arg.kind === 'range' && arg.values) {
			for (const row of arg.values) {
				for (const cell of row) {
					if (cell.kind === 'error') {
						if (!ignoreErrors) return cell
						continue
					}
					const n = numericVal(cell)
					if (n !== null) nums.push(n)
				}
			}
		} else {
			const v = arg.value ?? EMPTY
			if (v.kind === 'error') {
				if (!ignoreErrors) return v
				continue
			}
			const n = toNum(v)
			if (typeof n === 'number') nums.push(n)
		}
	}
	return nums
}

function aggregateFn(args: EvalArg[]): CellValue {
	const fnNum = numArg(args[0])
	if (typeof fnNum !== 'number') return fnNum
	const code = Math.trunc(fnNum)
	const opt = numArg(args[1])
	if (typeof opt !== 'number') return opt
	const options = Math.trunc(opt)
	const ignoreErrors = options === 2 || options === 3 || options === 6 || options === 7

	if (code >= 1 && code <= 11) {
		const data = args.slice(2)
		const numsOrErr = aggregateCollectNumbers(data, ignoreErrors)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		switch (code) {
			case 1: {
				if (numsOrErr.length === 0) return errorValue('#DIV/0!')
				const sum = numsOrErr.reduce((a, b) => a + b, 0)
				return numberValue(sum / numsOrErr.length)
			}
			case 2: {
				return numberValue(numsOrErr.length)
			}
			case 3: {
				const dataAll = args.slice(2)
				let count = 0
				for (const arg of dataAll) {
					if (arg.forEachValue) {
						let err: CellValue | undefined
						arg.forEachValue((cell) => {
							if (cell.kind === 'error') {
								if (!ignoreErrors) err = cell
								return
							}
							if (!isEmpty(cell)) count++
						})
						if (err) return err
					} else if (arg.kind === 'range' && arg.values) {
						for (const row of arg.values) {
							for (const cell of row) {
								if (cell.kind === 'error') {
									if (!ignoreErrors) return cell
									continue
								}
								if (!isEmpty(cell)) count++
							}
						}
					} else if (!isEmpty(arg.value ?? EMPTY)) count++
				}
				return numberValue(count)
			}
			case 4:
				return numberValue(numsOrErr.length === 0 ? 0 : Math.max(...numsOrErr))
			case 5:
				return numberValue(numsOrErr.length === 0 ? 0 : Math.min(...numsOrErr))
			case 6:
				return numberValue(numsOrErr.length === 0 ? 0 : numsOrErr.reduce((a, b) => a * b, 1))
			case 7: {
				if (numsOrErr.length < 2) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
				return numberValue(Math.sqrt(sumSq / (numsOrErr.length - 1)))
			}
			case 8: {
				if (numsOrErr.length < 2) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
				return numberValue(sumSq / (numsOrErr.length - 1))
			}
			case 9:
				return numberValue(numsOrErr.reduce((a, b) => a + b, 0))
			case 10: {
				if (numsOrErr.length < 2) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				return numberValue(
					numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / (numsOrErr.length - 1),
				)
			}
			case 11: {
				if (numsOrErr.length === 0) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				return numberValue(
					numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / numsOrErr.length,
				)
			}
			default:
				return errorValue('#VALUE!')
		}
	}

	if (code >= 12 && code <= 13) {
		const data = args.slice(2)
		const numsOrErr = aggregateCollectNumbers(data, ignoreErrors)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		if (numsOrErr.length === 0) return errorValue('#NUM!')
		if (code === 12) {
			numsOrErr.sort((a, b) => a - b)
			const mid = Math.floor(numsOrErr.length / 2)
			return numberValue(
				numsOrErr.length % 2 === 0
					? ((numsOrErr[mid - 1] ?? 0) + (numsOrErr[mid] ?? 0)) / 2
					: (numsOrErr[mid] ?? 0),
			)
		}
		const freq = new Map<number, number>()
		let maxCount = 0
		let modeVal = 0
		for (const n of numsOrErr) {
			const c = (freq.get(n) ?? 0) + 1
			freq.set(n, c)
			if (c > maxCount) {
				maxCount = c
				modeVal = n
			}
		}
		return maxCount < 2 ? errorValue('#N/A') : numberValue(modeVal)
	}

	if (code >= 14 && code <= 19) {
		if (args.length < 4) return errorValue('#VALUE!')
		const kArg = numArg(args[2])
		if (typeof kArg !== 'number') return kArg
		const k = Math.trunc(kArg)
		const data = args.slice(3)
		const numsOrErr = aggregateCollectNumbers(data, ignoreErrors)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		if (numsOrErr.length === 0) return errorValue('#NUM!')

		if (code === 14) {
			if (k < 1 || k > numsOrErr.length) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => b - a)
			return numberValue(numsOrErr[k - 1] ?? 0)
		}
		if (code === 15) {
			if (k < 1 || k > numsOrErr.length) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			return numberValue(numsOrErr[k - 1] ?? 0)
		}
		if (code === 16) {
			if (k < 0 || k > 1) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const n = numsOrErr.length
			const x = k * (n - 1)
			const i = Math.floor(x)
			const frac = x - i
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
		if (code === 17) {
			if (k < 0 || k > 4) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const n = numsOrErr.length
			if (k === 0) return numberValue(numsOrErr[0] ?? 0)
			if (k === 4) return numberValue(numsOrErr[n - 1] ?? 0)
			const q = (k / 4) * (n - 1)
			const i = Math.floor(q)
			const frac = q - i
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
		if (code === 18) {
			const n = numsOrErr.length
			if (k <= 0 || k >= 1) return errorValue('#NUM!')
			if (k < 1 / (n + 1) || k > n / (n + 1)) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const x = k * (n + 1) - 1
			const i = Math.floor(x)
			const frac = x - i
			if (i < 0) return numberValue(numsOrErr[0] ?? 0)
			if (i + 1 >= n) return numberValue(numsOrErr[n - 1] ?? 0)
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
		if (code === 19) {
			if (k < 1 || k > 3) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const n = numsOrErr.length
			const q = (k / 4) * (n + 1) - 1
			const i = Math.floor(q)
			const frac = q - i
			if (i < 0) return numberValue(numsOrErr[0] ?? 0)
			if (i + 1 >= n) return numberValue(numsOrErr[n - 1] ?? 0)
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
	}

	return errorValue('#VALUE!')
}

function subtotalNums(data: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of data) {
		if (arg.forEachValue) {
			let err: CellValue | undefined
			arg.forEachValue((cell) => {
				if (err) return
				if (isError(cell)) {
					err = cell
					return
				}
				const n = numericVal(cell)
				if (n !== null) nums.push(n)
			})
			if (err) return err
		} else if (arg.kind === 'range' && arg.values) {
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
		case 10:
		case 110: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			const divisor = nums.length - 1
			if (divisor < 1) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			return numberValue(nums.reduce((acc, v) => acc + (v - mean) ** 2, 0) / divisor)
		}
		case 11:
		case 111: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			if (nums.length === 0) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			return numberValue(nums.reduce((acc, v) => acc + (v - mean) ** 2, 0) / nums.length)
		}
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
		if (pairs.some((pair) => !sameShape(sumRange, pair.range))) return errorValue('#VALUE!')
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
		if (pairs.some((pair) => !sameShape(first, pair.range))) return errorValue('#VALUE!')
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
		if (pairs.some((pair) => !sameShape(avgRange, pair.range))) return errorValue('#VALUE!')
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
		if (pairs.some((pair) => !sameShape(minRange, pair.range))) return errorValue('#VALUE!')
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
		if (pairs.some((pair) => !sameShape(maxRange, pair.range))) return errorValue('#VALUE!')
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
