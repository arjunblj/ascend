import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef } from './registry.ts'
import { compareValues, getRange, numArg, toNumber } from './registry.ts'

function num(arg: EvalArg | undefined): number | CellValue {
	return numArg(arg)
}

function scalarOrSpill(rows: readonly (readonly CellValue[])[]): CellValue {
	const first = rows[0]
	if (!first || first.length === 0) return EMPTY
	const rowCount = rows.length
	const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0)
	if (rowCount > 1 || colCount > 1) return errorValue('#SPILL!')
	return first[0] ?? EMPTY
}

export const dynamicFunctions: FunctionDef[] = [
	{
		name: 'SORT',
		minArgs: 1,
		maxArgs: 4,
		evaluate(args) {
			const data = getRange(args[0])
			if (data.length === 0) return EMPTY
			const sortIndex = args[1] ? num(args[1]) : 1
			if (typeof sortIndex !== 'number') return sortIndex
			const sortOrder = args[2] ? num(args[2]) : 1
			if (typeof sortOrder !== 'number') return sortOrder

			const col = Math.round(sortIndex) - 1
			const rows = data.map((r) => [...r])
			rows.sort((a, b) => {
				const av = a[col] ?? EMPTY
				const bv = b[col] ?? EMPTY
				return compareValues(av, bv) * (sortOrder === -1 ? -1 : 1)
			})
			return scalarOrSpill(rows)
		},
	},
	{
		name: 'SORTBY',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			if (data.length === 0) return EMPTY
			const byRange = getRange(args[1]).map((r) => r[0] ?? EMPTY)
			const order = args[2] ? num(args[2]) : 1
			if (typeof order !== 'number') return order

			const indices = Array.from({ length: data.length }, (_, i) => i)
			indices.sort((a, b) => {
				const av = byRange[a] ?? EMPTY
				const bv = byRange[b] ?? EMPTY
				return compareValues(av, bv) * (order === -1 ? -1 : 1)
			})
			const rows = indices.map((index) => data[index] ?? [])
			return scalarOrSpill(rows)
		},
	},
	{
		name: 'FILTER',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			const include = getRange(args[1]).map((r) => {
				const v = r[0]
				if (!v) return false
				if (v.kind === 'boolean') return v.value
				if (v.kind === 'number') return v.value !== 0
				return false
			})
			const ifEmpty = args[2]?.value

			const filtered = data.filter((_, i) => include[i])
			if (filtered.length === 0) {
				return ifEmpty ?? errorValue('#CALC!')
			}
			return scalarOrSpill(filtered)
		},
	},
	{
		name: 'UNIQUE',
		minArgs: 1,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			if (data.length === 0) return EMPTY
			const byCol = args[1] ? toNumber(args[1].value) === 1 : false
			const exactlyOnce = args[2] ? toNumber(args[2].value) === 1 : false

			if (byCol) {
				return scalarOrSpill(data)
			}

			const seen = new Set<string>()
			const counts = new Map<string, number>()
			const unique: (readonly CellValue[])[] = []

			for (const row of data) {
				const key = row
					.map(
						(c) =>
							`${c.kind}:${c.kind === 'number' ? c.value : c.kind === 'string' ? c.value : ''}`,
					)
					.join('|')
				counts.set(key, (counts.get(key) ?? 0) + 1)
				if (!seen.has(key)) {
					seen.add(key)
					unique.push(row)
				}
			}

			if (exactlyOnce) {
				const once = unique.filter((row) => {
					const key = row
						.map(
							(c) =>
								`${c.kind}:${c.kind === 'number' ? c.value : c.kind === 'string' ? c.value : ''}`,
						)
						.join('|')
					return counts.get(key) === 1
				})
				if (once.length === 0) return errorValue('#CALC!')
				return scalarOrSpill(once)
			}

			return scalarOrSpill(unique)
		},
	},
	{
		name: 'SEQUENCE',
		minArgs: 1,
		maxArgs: 4,
		evaluate(args) {
			const rows = num(args[0])
			if (typeof rows !== 'number') return rows
			const cols = args[1] ? num(args[1]) : 1
			if (typeof cols !== 'number') return cols
			const start = args[2] ? num(args[2]) : 1
			if (typeof start !== 'number') return start
			const step = args[3] ? num(args[3]) : 1
			if (typeof step !== 'number') return step
			const r = Math.trunc(rows)
			const c = Math.trunc(cols)
			if (r <= 0 || c <= 0) return errorValue('#CALC!')
			if (r > 1 || c > 1) return errorValue('#SPILL!')
			return numberValue(start)
		},
	},
	{
		name: 'RANDARRAY',
		minArgs: 0,
		maxArgs: 5,
		volatile: true,
		evaluate(args, ctx) {
			const seed = ctx?.randomSeed ?? Math.random()
			const min = args[2] ? num(args[2]) : 0
			if (typeof min !== 'number') return min
			const max = args[3] ? num(args[3]) : 1
			if (typeof max !== 'number') return max
			const whole = args[4] ? toNumber(args[4].value) === 1 : false
			const val = min + seed * (max - min)
			return numberValue(whole ? Math.floor(val) : val)
		},
	},
	{
		name: 'LET',
		minArgs: 3,
		maxArgs: 255,
		evaluate(args) {
			if (args.length < 3 || args.length % 2 === 0) return errorValue('#VALUE!')
			const last = args[args.length - 1]
			return last?.value ?? EMPTY
		},
	},
	{
		name: 'TRANSPOSE',
		minArgs: 1,
		maxArgs: 1,
		evaluate(args) {
			const data = getRange(args[0])
			if (data.length === 0) return EMPTY
			return scalarOrSpill(data)
		},
	},
	{
		name: 'TOCOL',
		minArgs: 1,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'TOROW',
		minArgs: 1,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'WRAPCOLS',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'WRAPROWS',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'HSTACK',
		minArgs: 1,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'VSTACK',
		minArgs: 1,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'TAKE',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'DROP',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			const rows = args[1] ? num(args[1]) : 0
			if (typeof rows !== 'number') return rows
			if (Math.abs(rows) >= data.length) return errorValue('#CALC!')
			const dropped = rows > 0 ? data.slice(rows) : data.slice(0, data.length + rows)
			return scalarOrSpill(dropped)
		},
	},
	{
		name: 'EXPAND',
		minArgs: 2,
		maxArgs: 4,
		evaluate(args) {
			const data = getRange(args[0])
			return scalarOrSpill(data)
		},
	},
	{
		name: 'CHOOSECOLS',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			const col = args[1] ? num(args[1]) : 1
			if (typeof col !== 'number') return col
			const idx = Math.round(col) - 1
			const first = data[0]
			if (!first || idx < 0 || idx >= first.length) return errorValue('#VALUE!')
			return scalarOrSpill(data.map((row) => [row[idx] ?? EMPTY]))
		},
	},
	{
		name: 'CHOOSEROWS',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			const row = args[1] ? num(args[1]) : 1
			if (typeof row !== 'number') return row
			const idx = Math.round(row) - 1
			if (idx < 0 || idx >= data.length) return errorValue('#VALUE!')
			const r = data[idx]
			return scalarOrSpill(r ? [r] : [])
		},
	},
]
