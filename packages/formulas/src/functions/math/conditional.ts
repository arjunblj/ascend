import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, isEmpty, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef, FunctionEvalContext } from '../index.ts'
import { wildcardMatch } from '../registry.ts'
import { fn, getRange, numericVal, sameShape } from './helpers.ts'

const CRITERIA_CACHE_MAX = 1024
const criteriaPredicateCache = new Map<string, (v: CellValue) => boolean>()

let criteriaMatchCache = new WeakMap<readonly (readonly CellValue[])[], Map<string, Uint8Array>>()

export function clearCriteriaMatchCache(): void {
	criteriaMatchCache = new WeakMap()
}

function getMatchBitmap(range: readonly (readonly CellValue[])[], criteria: CellValue): Uint8Array {
	let inner = criteriaMatchCache.get(range)
	if (!inner) {
		inner = new Map()
		criteriaMatchCache.set(range, inner)
	}
	const ck = criteriaCacheKey(criteria)
	const cached = inner.get(ck)
	if (cached) return cached

	const match = parseCriteria(criteria)
	const rows = range.length
	const cols = range[0]?.length ?? 1
	const bitmap = new Uint8Array(rows * cols)
	for (let r = 0; r < rows; r++) {
		const row = range[r]
		if (!row) continue
		for (let c = 0; c < cols; c++) {
			if (match(row[c] ?? EMPTY)) bitmap[r * cols + c] = 1
		}
	}
	inner.set(ck, bitmap)
	return bitmap
}

function criteriaCacheKey(criteria: CellValue): string {
	return JSON.stringify(criteria)
}

function rangeCacheKey(arg: EvalArg | undefined): string | null {
	const ref = arg?.ref
	if (!ref || ref.kind !== 'range') return null
	return `${ref.sheetIndex}:${ref.row}:${ref.col}:${ref.endRow ?? ref.row}:${ref.endCol ?? ref.col}`
}

function singleCriteriaCacheKey(name: string, args: EvalArg[]): string | null {
	const criteriaRange = rangeCacheKey(args[0])
	if (!criteriaRange) return null
	const sumRange = args.length >= 3 ? rangeCacheKey(args[2]) : criteriaRange
	if (!sumRange) return null
	return `COND:${name}:${criteriaRange}:${criteriaCacheKey(args[1]?.value ?? EMPTY)}:${sumRange}`
}

function ifsCacheKey(name: string, args: EvalArg[], firstCriteriaRangeIdx: number): string | null {
	const parts = [`COND:${name}`]
	if (firstCriteriaRangeIdx > 0) {
		const targetRange = rangeCacheKey(args[0])
		if (!targetRange) return null
		parts.push(targetRange)
	}
	for (let i = firstCriteriaRangeIdx; i + 1 < args.length; i += 2) {
		const criteriaRange = rangeCacheKey(args[i])
		if (!criteriaRange) return null
		parts.push(criteriaRange, criteriaCacheKey(args[i + 1]?.value ?? EMPTY))
	}
	return parts.join(':')
}

function withCachedConditionalAggregate(
	key: string | null,
	ctx: FunctionEvalContext | undefined,
	compute: () => CellValue,
): CellValue {
	const cache = ctx?.aggregateRangeCache
	if (!key || !cache) return compute()
	const cached = cache.get(key)
	if (cached) return cached
	const value = compute()
	cache.set(key, value)
	return value
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

interface CriteriaPair {
	range: readonly (readonly CellValue[])[]
	bitmap: Uint8Array
	cols: number
}

function buildPairs(args: EvalArg[], startIdx: number): CriteriaPair[] {
	const pairs: CriteriaPair[] = []
	for (let i = startIdx; i + 1 < args.length; i += 2) {
		const range = getRange(args[i])
		const cols = range[0]?.length ?? 1
		pairs.push({
			range,
			bitmap: getMatchBitmap(range, args[i + 1]?.value ?? EMPTY),
			cols,
		})
	}
	return pairs
}

function allPairsMatch(pairs: CriteriaPair[], r: number, c: number): boolean {
	for (const p of pairs) {
		if (!p.bitmap[r * p.cols + c]) return false
	}
	return true
}

function firstScalarError(args: readonly (EvalArg | undefined)[]): CellValue | null {
	for (const arg of args) {
		const value = arg?.value ?? EMPTY
		if (value.kind === 'error') return value
	}
	return null
}

export const conditionalFunctions: FunctionDef[] = [
	fn('SUMIF', 2, 3, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(singleCriteriaCacheKey('SUMIF', args), ctx, () => {
			const range = getRange(args[0])
			const bitmap = getMatchBitmap(range, args[1]?.value ?? EMPTY)
			const cols = range[0]?.length ?? 1
			const sumRange = args.length >= 3 ? getRange(args[2]) : range
			let sum = 0
			for (let r = 0; r < range.length; r++) {
				for (let c = 0; c < (range[r]?.length ?? 0); c++) {
					if (bitmap[r * cols + c]) {
						const n = numericVal(sumRange[r]?.[c] ?? EMPTY)
						if (n !== null) sum += n
					}
				}
			}
			return numberValue(sum)
		})
	}),

	fn('SUMIFS', 3, 255, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(ifsCacheKey('SUMIFS', args, 1), ctx, () => {
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
		})
	}),

	fn('COUNTIF', 2, 2, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(ifsCacheKey('COUNTIF', args, 0), ctx, () => {
			const range = getRange(args[0])
			const bitmap = getMatchBitmap(range, args[1]?.value ?? EMPTY)
			const cols = range[0]?.length ?? 1
			let count = 0
			for (let r = 0; r < range.length; r++) {
				for (let c = 0; c < (range[r]?.length ?? 0); c++) {
					if (bitmap[r * cols + c]) count++
				}
			}
			return numberValue(count)
		})
	}),

	fn('COUNTIFS', 2, 255, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(ifsCacheKey('COUNTIFS', args, 0), ctx, () => {
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
		})
	}),

	fn('AVERAGEIF', 2, 3, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(singleCriteriaCacheKey('AVERAGEIF', args), ctx, () => {
			const range = getRange(args[0])
			const bitmap = getMatchBitmap(range, args[1]?.value ?? EMPTY)
			const cols = range[0]?.length ?? 1
			const avgRange = args.length >= 3 ? getRange(args[2]) : range
			let sum = 0
			let count = 0
			for (let r = 0; r < range.length; r++) {
				for (let c = 0; c < (range[r]?.length ?? 0); c++) {
					if (bitmap[r * cols + c]) {
						const n = numericVal(avgRange[r]?.[c] ?? EMPTY)
						if (n !== null) {
							sum += n
							count++
						}
					}
				}
			}
			return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
		})
	}),

	fn('AVERAGEIFS', 3, 255, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(ifsCacheKey('AVERAGEIFS', args, 1), ctx, () => {
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
		})
	}),

	fn('MINIFS', 3, 255, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(ifsCacheKey('MINIFS', args, 1), ctx, () => {
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
		})
	}),

	fn('MAXIFS', 3, 255, (args, ctx) => {
		const directError = firstScalarError(args)
		if (directError) return directError
		return withCachedConditionalAggregate(ifsCacheKey('MAXIFS', args, 1), ctx, () => {
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
		})
	}),
]
