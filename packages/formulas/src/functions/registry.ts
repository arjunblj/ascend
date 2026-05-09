import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, topLeftScalar } from '@ascend/schema'

export interface EvalRef {
	readonly kind: 'cell' | 'range'
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly endRow?: number
	readonly endCol?: number
}

export interface EvalArea {
	readonly ref: EvalRef
	readonly values: readonly (readonly CellValue[])[]
	readonly topLeft?: CellValue
	/** Iterate only occupied cells (sparse). Use for SUM, AVERAGE, etc. */
	readonly forEachValue?: (fn: (value: CellValue) => void) => void
	/** Iterate all cells in range including empty. Use for COUNTBLANK. */
	readonly forEachCellInRange?: (fn: (value: CellValue) => void) => void
}

export interface EvalArg {
	readonly value: CellValue
	readonly kind?: 'range'
	readonly values?: readonly (readonly CellValue[])[]
	readonly ref?: EvalRef
	readonly areas?: readonly EvalArea[]
	readonly shapeRows?: number
	readonly shapeCols?: number
	/** Iterate only occupied cells (sparse). Use for SUM, AVERAGE, etc. */
	readonly forEachValue?: (fn: (value: CellValue) => void) => void
	/** Iterate all cells in range including empty. Use for COUNTBLANK. */
	readonly forEachCellInRange?: (fn: (value: CellValue) => void) => void
}

export type FnArg = EvalArg

export interface ExactLookupHit {
	readonly first: number
	readonly last: number
}

export type ExactLookupCache = Map<string, ReadonlyMap<string, ExactLookupHit>>
export type LookupVectorCache = Map<string, readonly CellValue[]>
export type AggregateRangeCache = Map<string, CellValue>
export type NumericVectorCache = Map<string, number[]>

export interface FunctionEvalContext {
	readonly now: Date
	readonly today: Date
	readonly randomSeed: number
	readonly locale: string
	readonly dateSystem: '1900' | '1904'
	readonly sheetIndex?: number
	readonly row?: number
	readonly col?: number
	readonly exactLookupCache: ExactLookupCache | undefined
	readonly lookupVectorCache: LookupVectorCache | undefined
	readonly aggregateRangeCache: AggregateRangeCache | undefined
	readonly numericVectorCache: NumericVectorCache | undefined
}

export interface FunctionDef {
	readonly name: string
	readonly minArgs: number
	readonly maxArgs: number
	readonly evaluate: (args: EvalArg[], ctx?: FunctionEvalContext) => CellValue
	readonly volatile?: boolean
}

export const functionRegistry = new Map<string, FunctionDef>()

export function registerFunction(def: FunctionDef): void {
	functionRegistry.set(def.name.toUpperCase(), def)
}

export function getRange(arg: EvalArg | undefined): readonly (readonly CellValue[])[] {
	if (arg?.areas?.length) {
		if (arg.areas.length === 1) return arg.areas[0]?.values ?? [[EMPTY]]
		return [[errorValue('#VALUE!')]]
	}
	if (arg?.kind === 'range' && arg.values) return arg.values
	if (arg?.value.kind === 'array') return arg.value.rows
	return [[arg?.value ?? EMPTY]]
}

export function rangeShape(arg: EvalArg | undefined): { rows: number; cols: number } {
	if (arg?.shapeRows !== undefined && arg?.shapeCols !== undefined) {
		return { rows: arg.shapeRows, cols: arg.shapeCols }
	}
	if (arg?.ref) {
		return {
			rows: (arg.ref.endRow ?? arg.ref.row) - arg.ref.row + 1,
			cols: (arg.ref.endCol ?? arg.ref.col) - arg.ref.col + 1,
		}
	}
	const range = getRange(arg)
	return { rows: range.length, cols: range[0]?.length ?? 0 }
}

export function iterAreaRows(arg: EvalArg | undefined): readonly (readonly CellValue[])[] {
	if (!arg) return []
	if (arg.areas?.length) {
		if (arg.areas.length === 1) return arg.areas[0]?.values ?? []
		return []
	}
	if (arg.kind === 'range' && arg.values) return arg.values
	if (arg.value.kind === 'array') return arg.value.rows
	return [[arg.value]]
}

export function toNumber(v: CellValue): number | null {
	const scalar = topLeftScalar(v)
	if (scalar.kind !== v.kind) return toNumber(scalar)
	switch (v.kind) {
		case 'number':
			return v.value
		case 'date':
			return v.serial
		case 'boolean':
			return v.value ? 1 : 0
		case 'empty':
			return 0
		case 'string': {
			const s = v.value.trim()
			if (s === '') return 0
			const n = Number(s)
			return Number.isNaN(n) ? null : n
		}
		default:
			return null
	}
}

export function numArg(arg: EvalArg | undefined): number | CellValue {
	const v = arg?.value ?? EMPTY
	if (v.kind === 'error') return v
	const n = toNumber(v)
	return n === null ? errorValue('#VALUE!') : n
}

export function cellOf(arg: EvalArg | undefined): CellValue {
	const value = arg?.value ?? EMPTY
	return topLeftScalar(value)
}

export function flattenArgs(args: EvalArg[]): CellValue[] {
	const result: CellValue[] = []
	for (const arg of args) {
		if (arg.forEachValue) {
			arg.forEachValue((cell) => result.push(cell))
			continue
		}
		if (arg.areas?.length) {
			for (const area of arg.areas) {
				if (area.forEachValue) {
					area.forEachValue((cell) => result.push(cell))
				} else {
					for (const row of area.values) {
						for (const cell of row) result.push(cell)
					}
				}
			}
			continue
		}
		if (arg.kind === 'range' && arg.values) {
			for (const row of arg.values) {
				for (const cell of row) result.push(cell)
			}
		} else {
			if (arg.value.kind === 'array') {
				for (const row of arg.value.rows) {
					for (const cell of row) result.push(cell)
				}
			} else {
				result.push(arg.value)
			}
		}
	}
	return result
}

export function collectNumbers(args: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of args) {
		if (arg.forEachValue) {
			let err: CellValue | undefined
			arg.forEachValue((cell) => {
				if (err) return
				if (cell.kind === 'error') {
					err = cell
					return
				}
				if (cell.kind === 'number') nums.push(cell.value)
				else if (cell.kind === 'date') nums.push(cell.serial)
			})
			if (err) return err
			continue
		}
		if (arg.areas?.length) {
			for (const area of arg.areas) {
				if (area.forEachValue) {
					let err: CellValue | undefined
					area.forEachValue((cell) => {
						if (err) return
						if (cell.kind === 'error') {
							err = cell
							return
						}
						if (cell.kind === 'number') nums.push(cell.value)
						else if (cell.kind === 'date') nums.push(cell.serial)
					})
					if (err) return err
				} else {
					for (const row of area.values) {
						for (const cell of row) {
							if (cell.kind === 'error') return cell
							if (cell.kind === 'number') nums.push(cell.value)
							else if (cell.kind === 'date') nums.push(cell.serial)
						}
					}
				}
			}
			continue
		}
		if (arg.kind === 'range' && arg.values) {
			for (const row of arg.values) {
				for (const cell of row) {
					if (cell.kind === 'error') return cell
					if (cell.kind === 'number') nums.push(cell.value)
					else if (cell.kind === 'date') nums.push(cell.serial)
				}
			}
		} else {
			const scalar = topLeftScalar(arg.value)
			if (scalar.kind === 'error') return scalar
			const n = toNumber(scalar)
			if (n !== null) nums.push(n)
		}
	}
	return nums
}

export function compareValues(a: CellValue, b: CellValue): number {
	a = topLeftScalar(a)
	b = topLeftScalar(b)
	const an = numericOf(a)
	const bn = numericOf(b)
	if (an !== null && bn !== null) return an - bn
	if (a.kind === 'string' && b.kind === 'string') {
		const al = a.value.toLowerCase()
		const bl = b.value.toLowerCase()
		return al < bl ? -1 : al > bl ? 1 : 0
	}
	if (a.kind === 'empty') return b.kind === 'empty' ? 0 : -1
	if (b.kind === 'empty') return 1
	return typeRank(a) - typeRank(b)
}

export function valuesEqual(a: CellValue, b: CellValue): boolean {
	a = topLeftScalar(a)
	b = topLeftScalar(b)
	const an = numericOf(a)
	const bn = numericOf(b)
	if (an !== null && bn !== null) return an === bn
	if (a.kind === 'string' && b.kind === 'string')
		return a.value.toLowerCase() === b.value.toLowerCase()
	if (a.kind === 'boolean' && b.kind === 'boolean') return a.value === b.value
	if (a.kind === 'error' && b.kind === 'error') return a.value === b.value
	return a.kind === 'empty' && b.kind === 'empty'
}

const wildcardCache = new Map<string, RegExp>()

function hasWildcardChars(pattern: string): boolean {
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i]
		if (ch === '~') {
			i++
			continue
		}
		if (ch === '*' || ch === '?') return true
	}
	return false
}

export function wildcardMatch(pattern: string, text: string): boolean {
	const p = pattern.toLowerCase()
	if (!hasWildcardChars(p)) return text.toLowerCase() === p
	let compiled = wildcardCache.get(p)
	if (!compiled) {
		const parts: string[] = ['^']
		let i = 0
		while (i < p.length) {
			const ch = p[i] ?? ''
			if (ch === '~' && i + 1 < p.length) {
				i++
				parts.push(escapeRe(p[i] ?? ''))
			} else if (ch === '*') {
				parts.push('.*')
			} else if (ch === '?') {
				parts.push('.')
			} else {
				parts.push(escapeRe(ch))
			}
			i++
		}
		parts.push('$')
		compiled = new RegExp(parts.join(''))
		if (wildcardCache.size > 1024) wildcardCache.clear()
		wildcardCache.set(p, compiled)
	}
	return compiled.test(text.toLowerCase())
}

const RE_SPECIAL = new Set(['.', '+', '*', '^', '$', '{', '}', '(', ')', '|', '[', ']', '\\'])

function escapeRe(ch: string): string {
	return RE_SPECIAL.has(ch) ? `\\${ch}` : ch
}

function numericOf(v: CellValue): number | null {
	v = topLeftScalar(v)
	if (v.kind === 'number') return v.value
	if (v.kind === 'date') return v.serial
	return null
}

function typeRank(v: CellValue): number {
	v = topLeftScalar(v)
	switch (v.kind) {
		case 'empty':
			return 0
		case 'number':
		case 'date':
			return 1
		case 'string':
			return 2
		case 'boolean':
			return 3
		default:
			return 4
	}
}
