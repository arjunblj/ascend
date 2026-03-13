import type { CellValue, ScalarCellValue } from '@ascend/schema'
import { arrayValue, EMPTY, errorValue, numberValue, topLeftScalar } from '@ascend/schema'
import type { EvalArg, FunctionDef } from './registry.ts'
import { compareValues, getRange, numArg, toNumber } from './registry.ts'

function num(arg: EvalArg | undefined): number | CellValue {
	return numArg(arg)
}

function scalarOrArray(rows: readonly (readonly CellValue[])[]): CellValue {
	const first = rows[0]
	if (!first || first.length === 0) return EMPTY
	const rowCount = rows.length
	const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0)
	if (rowCount === 1 && colCount === 1) return first[0] ?? EMPTY
	return arrayValue(rows.map((row) => row.map((cell) => topLeftScalar(cell))))
}

function flattenByColumn(data: readonly (readonly CellValue[])[]): ScalarCellValue[][] {
	const rows: ScalarCellValue[][] = []
	for (const row of data) {
		for (const cell of row) rows.push([topLeftScalar(cell)])
	}
	return rows
}

function flattenByRow(data: readonly (readonly CellValue[])[]): ScalarCellValue[][] {
	const flat: ScalarCellValue[] = []
	for (const row of data) {
		for (const cell of row) flat.push(topLeftScalar(cell))
	}
	return [flat]
}

function flattenScalars(
	data: readonly (readonly CellValue[])[],
	scanByColumn: boolean,
): ScalarCellValue[] {
	if (!scanByColumn) return flattenByRow(data)[0] ?? []
	const flat: ScalarCellValue[] = []
	const rowCount = data.length
	const colCount = data.reduce((max, row) => Math.max(max, row.length), 0)
	for (let col = 0; col < colCount; col++) {
		for (let row = 0; row < rowCount; row++) {
			flat.push(topLeftScalar(data[row]?.[col] ?? EMPTY))
		}
	}
	return flat
}

function transposeRows(data: readonly (readonly CellValue[])[]): ScalarCellValue[][] {
	const rowCount = data.length
	const colCount = data.reduce((max, row) => Math.max(max, row.length), 0)
	const rows: ScalarCellValue[][] = []
	for (let col = 0; col < colCount; col++) {
		const row: ScalarCellValue[] = []
		for (let sourceRow = 0; sourceRow < rowCount; sourceRow++) {
			row.push(topLeftScalar(data[sourceRow]?.[col] ?? EMPTY))
		}
		rows.push(row)
	}
	return rows
}

function chooseCols(
	data: readonly (readonly CellValue[])[],
	indices: readonly number[],
): ScalarCellValue[][] | CellValue {
	const first = data[0] ?? []
	for (const idx of indices) {
		if (idx < 0 || idx >= first.length) return errorValue('#VALUE!')
	}
	return data.map((row) => indices.map((idx) => topLeftScalar(row[idx] ?? EMPTY)))
}

function chooseRows(
	data: readonly (readonly CellValue[])[],
	indices: readonly number[],
): ScalarCellValue[][] | CellValue {
	for (const idx of indices) {
		if (idx < 0 || idx >= data.length) return errorValue('#VALUE!')
	}
	return indices.map((idx) => (data[idx] ?? []).map((cell) => topLeftScalar(cell)))
}

function padRows(
	rows: readonly (readonly ScalarCellValue[])[],
	targetCols: number,
	fill: ScalarCellValue,
): ScalarCellValue[][] {
	return rows.map((row) => {
		const next = [...row]
		while (next.length < targetCols) next.push(fill)
		return next
	})
}

function readIgnoreMode(arg: EvalArg | undefined): number | CellValue {
	if (!arg) return 0
	const value = toNumber(arg.value)
	if (value === null || !Number.isInteger(value) || value < 0 || value > 3) {
		return errorValue('#VALUE!')
	}
	return value
}

function readScanByColumn(arg: EvalArg | undefined): boolean | CellValue {
	if (!arg) return false
	const value = toNumber(arg.value)
	if (value === null) return errorValue('#VALUE!')
	return value !== 0
}

function shouldIgnoreValue(value: ScalarCellValue, ignoreMode: number): boolean {
	if ((ignoreMode === 1 || ignoreMode === 3) && value.kind === 'empty') return true
	if ((ignoreMode === 2 || ignoreMode === 3) && value.kind === 'error') return true
	return false
}

function fingerprintCell(value: CellValue): string {
	switch (value.kind) {
		case 'number':
			return `n:${value.value}`
		case 'string':
			return `s:${value.value}`
		case 'boolean':
			return `b:${value.value}`
		case 'date':
			return `d:${value.serial}`
		case 'error':
			return `e:${value.value}`
		case 'richText':
			return `t:${value.runs.map((r) => r.text).join('\x01')}`
		case 'array':
			return `a:${value.rows.map((row) => row.map((c) => fingerprintCell(c)).join('\x00')).join('\x01')}`
		case 'empty':
			return ''
	}
}

function fingerprintVector(values: readonly CellValue[]): string {
	return values.map((v) => fingerprintCell(v)).join('\x00')
}

function columnVectors(data: readonly (readonly CellValue[])[]): CellValue[][] {
	const colCount = data.reduce((max, row) => Math.max(max, row.length), 0)
	const columns: CellValue[][] = []
	for (let col = 0; col < colCount; col++) {
		const values: CellValue[] = []
		for (const row of data) values.push(row[col] ?? EMPTY)
		columns.push(values)
	}
	return columns
}

function rangeLooksLikeReference(arg: EvalArg | undefined): boolean {
	if (!arg) return false
	return arg.kind === 'range' || arg.value.kind === 'array'
}

function resolveSelectionIndex(index: number, size: number): number | CellValue {
	if (index === 0) return errorValue('#VALUE!')
	const normalized = index < 0 ? size + index : index - 1
	return normalized < 0 || normalized >= size ? errorValue('#VALUE!') : normalized
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
			const byCol = !!args[3] && !!toNumber(args[3].value)

			if (byCol) {
				const cols = transposeRows(data)
				const row = Math.round(sortIndex) - 1
				cols.sort((a, b) => {
					const av = a[row] ?? EMPTY
					const bv = b[row] ?? EMPTY
					return compareValues(av, bv) * (sortOrder === -1 ? -1 : 1)
				})
				return scalarOrArray(transposeRows(cols))
			}

			const col = Math.round(sortIndex) - 1
			const rows = data.map((r) => [...r])
			rows.sort((a, b) => {
				const av = a[col] ?? EMPTY
				const bv = b[col] ?? EMPTY
				return compareValues(av, bv) * (sortOrder === -1 ? -1 : 1)
			})
			return scalarOrArray(rows)
		},
	},
	{
		name: 'SORTBY',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			if (data.length === 0) return EMPTY
			const sortKeys: Array<{ values: CellValue[]; order: 1 | -1 }> = []
			let index = 1
			while (index < args.length) {
				const byArg = args[index]
				if (!byArg) break
				const values = getRange(byArg).map((row) => row[0] ?? EMPTY)
				let order: 1 | -1 = 1
				const next = args[index + 1]
				if (next && !rangeLooksLikeReference(next)) {
					const parsedOrder = num(next)
					if (typeof parsedOrder !== 'number') return parsedOrder
					order = parsedOrder === -1 ? -1 : 1
					index += 2
				} else {
					index += 1
				}
				sortKeys.push({ values, order })
			}
			if (sortKeys.length === 0) return scalarOrArray(data)

			const indices = Array.from({ length: data.length }, (_, i) => i)
			indices.sort((a, b) => {
				for (const key of sortKeys) {
					const av = key.values[a] ?? EMPTY
					const bv = key.values[b] ?? EMPTY
					const cmp = compareValues(av, bv)
					if (cmp !== 0) return cmp * key.order
				}
				return 0
			})
			const rows = indices.map((index) => data[index] ?? [])
			return scalarOrArray(rows)
		},
	},
	{
		name: 'FILTER',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			const includeRange = getRange(args[1])
			const ifEmpty = args[2]?.value
			const rowCount = data.length
			const colCount = data.reduce((max, row) => Math.max(max, row.length), 0)
			const includeRows = includeRange.length
			const includeCols = includeRange.reduce((max, row) => Math.max(max, row.length), 0)

			const coerceInclude = (value: CellValue): boolean | CellValue => {
				if (value.kind === 'error') return value
				if (value.kind === 'boolean') return value.value
				if (value.kind === 'number') return value.value !== 0
				return false
			}

			if (includeRows === rowCount && includeCols === 1) {
				const filtered: (readonly CellValue[])[] = []
				for (let row = 0; row < rowCount; row++) {
					const includeValue = coerceInclude(includeRange[row]?.[0] ?? EMPTY)
					if (typeof includeValue !== 'boolean') return includeValue
					if (includeValue) filtered.push(data[row] ?? [])
				}
				if (filtered.length === 0) return ifEmpty ?? errorValue('#CALC!')
				return scalarOrArray(filtered)
			}
			if (includeRows === 1 && includeCols === colCount) {
				const selectedCols: number[] = []
				for (let col = 0; col < colCount; col++) {
					const includeValue = coerceInclude(includeRange[0]?.[col] ?? EMPTY)
					if (typeof includeValue !== 'boolean') return includeValue
					if (includeValue) selectedCols.push(col)
				}
				if (selectedCols.length === 0) return ifEmpty ?? errorValue('#CALC!')
				return scalarOrArray(
					data.map((row) => selectedCols.map((col) => topLeftScalar(row[col] ?? EMPTY))),
				)
			}
			return errorValue('#VALUE!')
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
				const seen = new Set<string>()
				const counts = new Map<string, number>()
				const columns = columnVectors(data)
				const unique: CellValue[][] = []
				for (const column of columns) {
					const key = fingerprintVector(column)
					counts.set(key, (counts.get(key) ?? 0) + 1)
					if (!seen.has(key)) {
						seen.add(key)
						unique.push(column)
					}
				}
				const filtered = exactlyOnce
					? unique.filter((column) => counts.get(fingerprintVector(column)) === 1)
					: unique
				if (filtered.length === 0) return errorValue('#CALC!')
				return scalarOrArray(transposeRows(filtered))
			}

			const seen = new Set<string>()
			const counts = new Map<string, number>()
			const unique: (readonly CellValue[])[] = []

			for (const row of data) {
				const key = fingerprintVector(row)
				counts.set(key, (counts.get(key) ?? 0) + 1)
				if (!seen.has(key)) {
					seen.add(key)
					unique.push(row)
				}
			}

			if (exactlyOnce) {
				const once = unique.filter((row) => {
					const key = fingerprintVector(row)
					return counts.get(key) === 1
				})
				if (once.length === 0) return errorValue('#CALC!')
				return scalarOrArray(once)
			}

			return scalarOrArray(unique)
		},
	},
	{
		name: 'SEQUENCE',
		minArgs: 1,
		maxArgs: 4,
		evaluate(args) {
			const rowCountArg = num(args[0])
			if (typeof rowCountArg !== 'number') return rowCountArg
			const colCountArg = args[1] ? num(args[1]) : 1
			if (typeof colCountArg !== 'number') return colCountArg
			const start = args[2] ? num(args[2]) : 1
			if (typeof start !== 'number') return start
			const step = args[3] ? num(args[3]) : 1
			if (typeof step !== 'number') return step
			const r = Math.trunc(rowCountArg)
			const c = Math.trunc(colCountArg)
			if (r <= 0 || c <= 0) return errorValue('#CALC!')
			const resultRows: ScalarCellValue[][] = []
			let current = start
			for (let row = 0; row < r; row++) {
				const values: ScalarCellValue[] = []
				for (let col = 0; col < c; col++) {
					values.push(numberValue(current) as ScalarCellValue)
					current += step
				}
				resultRows.push(values)
			}
			return scalarOrArray(resultRows)
		},
	},
	{
		name: 'RANDARRAY',
		minArgs: 0,
		maxArgs: 5,
		volatile: true,
		evaluate(args, ctx) {
			const seed = ctx?.randomSeed ?? Math.random()
			const rowCount = args[0] ? num(args[0]) : 1
			if (typeof rowCount !== 'number') return rowCount
			const colCount = args[1] ? num(args[1]) : 1
			if (typeof colCount !== 'number') return colCount
			const min = args[2] ? num(args[2]) : 0
			if (typeof min !== 'number') return min
			const max = args[3] ? num(args[3]) : 1
			if (typeof max !== 'number') return max
			const whole = args[4] ? toNumber(args[4].value) === 1 : false
			const rows = Math.max(1, Math.trunc(rowCount))
			const cols = Math.max(1, Math.trunc(colCount))
			const result: ScalarCellValue[][] = []
			for (let row = 0; row < rows; row++) {
				const rowValues: ScalarCellValue[] = []
				for (let col = 0; col < cols; col++) {
					const mixedSeed = (seed + row * 0.131 + col * 0.071) % 1
					const val = min + mixedSeed * (max - min)
					rowValues.push(numberValue(whole ? Math.floor(val) : val) as ScalarCellValue)
				}
				result.push(rowValues)
			}
			return scalarOrArray(result)
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
			return scalarOrArray(transposeRows(data))
		},
	},
	{
		name: 'TOCOL',
		minArgs: 1,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			const ignoreMode = readIgnoreMode(args[1])
			if (typeof ignoreMode !== 'number') return ignoreMode
			const scanByColumn = readScanByColumn(args[2])
			if (typeof scanByColumn !== 'boolean') return scanByColumn
			const flat = flattenScalars(data, scanByColumn).filter(
				(value) => !shouldIgnoreValue(value, ignoreMode),
			)
			if (flat.length === 0) return errorValue('#CALC!')
			return scalarOrArray(flat.map((value) => [value]))
		},
	},
	{
		name: 'TOROW',
		minArgs: 1,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			const ignoreMode = readIgnoreMode(args[1])
			if (typeof ignoreMode !== 'number') return ignoreMode
			const scanByColumn = readScanByColumn(args[2])
			if (typeof scanByColumn !== 'boolean') return scanByColumn
			const flat = flattenScalars(data, scanByColumn).filter(
				(value) => !shouldIgnoreValue(value, ignoreMode),
			)
			if (flat.length === 0) return errorValue('#CALC!')
			return scalarOrArray([flat])
		},
	},
	{
		name: 'WRAPCOLS',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = flattenByColumn(getRange(args[0]))
			const wrap = num(args[1])
			if (typeof wrap !== 'number') return wrap
			const fill = topLeftScalar(args[2]?.value ?? errorValue('#N/A'))
			if (wrap <= 0) return errorValue('#VALUE!')
			const rows: ScalarCellValue[][] = []
			for (let index = 0; index < data.length; index += wrap) {
				const chunk = data
					.slice(index, index + wrap)
					.map((entry) => topLeftScalar(entry[0] ?? EMPTY))
				while (chunk.length < wrap) chunk.push(fill)
				rows.push(chunk)
			}
			return scalarOrArray(transposeRows(rows))
		},
	},
	{
		name: 'WRAPROWS',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = flattenByRow(getRange(args[0]))[0] ?? []
			const wrap = num(args[1])
			if (typeof wrap !== 'number') return wrap
			const fill = topLeftScalar(args[2]?.value ?? errorValue('#N/A'))
			if (wrap <= 0) return errorValue('#VALUE!')
			const rows: ScalarCellValue[][] = []
			for (let index = 0; index < data.length; index += wrap) {
				const chunk = data.slice(index, index + wrap)
				while (chunk.length < wrap) chunk.push(fill)
				rows.push(chunk)
			}
			return scalarOrArray(rows)
		},
	},
	{
		name: 'HSTACK',
		minArgs: 1,
		maxArgs: 255,
		evaluate(args) {
			const blocks = args.map((arg) => getRange(arg))
			const rowCount = blocks.reduce((max, block) => Math.max(max, block.length), 0)
			const totalCols = blocks.reduce(
				(sum, block) => sum + block.reduce((max, row) => Math.max(max, row.length), 0),
				0,
			)
			const fill = errorValue('#N/A') as ScalarCellValue
			const rows: ScalarCellValue[][] = []
			for (let row = 0; row < rowCount; row++) {
				const rowValues: ScalarCellValue[] = []
				for (const block of blocks) {
					const blockCols = block.reduce((max, entries) => Math.max(max, entries.length), 0)
					const source = block[row]
					for (let col = 0; col < blockCols; col++) {
						rowValues.push(topLeftScalar(source?.[col] ?? fill))
					}
				}
				rows.push(rowValues)
			}
			return scalarOrArray(padRows(rows, totalCols, fill))
		},
	},
	{
		name: 'VSTACK',
		minArgs: 1,
		maxArgs: 255,
		evaluate(args) {
			const fill = errorValue('#N/A') as ScalarCellValue
			const rows: ScalarCellValue[][] = []
			for (const arg of args) {
				for (const row of getRange(arg)) {
					rows.push(row.map((cell) => topLeftScalar(cell)))
				}
			}
			const width = rows.reduce((max, row) => Math.max(max, row.length), 0)
			return scalarOrArray(padRows(rows, width, fill))
		},
	},
	{
		name: 'TAKE',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const data = getRange(args[0])
			const rowCountArg = args[1] ? num(args[1]) : 0
			if (typeof rowCountArg !== 'number') return rowCountArg
			const colCountArg = args[2] ? num(args[2]) : undefined
			if (colCountArg !== undefined && typeof colCountArg !== 'number') return colCountArg
			if (rowCountArg === 0 || colCountArg === 0) return errorValue('#CALC!')
			let rows =
				rowCountArg >= 0
					? data.slice(0, rowCountArg)
					: data.slice(Math.max(0, data.length + rowCountArg))
			if (colCountArg !== undefined) {
				rows = rows.map((row) =>
					colCountArg >= 0
						? row.slice(0, colCountArg)
						: row.slice(Math.max(0, row.length + colCountArg)),
				)
			}
			if (rows.length === 0 || rows.every((row) => row.length === 0)) return errorValue('#CALC!')
			return scalarOrArray(rows)
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
			if (rows === 0) return errorValue('#CALC!')
			if (Math.abs(rows) >= data.length) return errorValue('#CALC!')
			const dropped = rows > 0 ? data.slice(rows) : data.slice(0, data.length + rows)
			let result = dropped
			const cols = args[2] ? num(args[2]) : 0
			if (typeof cols !== 'number') return cols
			if (cols === 0 && args[2] !== undefined) return errorValue('#CALC!')
			if (cols !== 0) {
				result = dropped.map((row) =>
					cols > 0 ? row.slice(cols) : row.slice(0, Math.max(0, row.length + cols)),
				)
			}
			if (result.length === 0 || result.every((row) => row.length === 0))
				return errorValue('#CALC!')
			return scalarOrArray(result)
		},
	},
	{
		name: 'EXPAND',
		minArgs: 2,
		maxArgs: 4,
		evaluate(args) {
			const data = getRange(args[0]).map((row) => row.map((cell) => topLeftScalar(cell)))
			const targetRows = num(args[1])
			if (typeof targetRows !== 'number') return targetRows
			const targetCols = args[2]
				? num(args[2])
				: data.reduce((max, row) => Math.max(max, row.length), 0)
			if (typeof targetCols !== 'number') return targetCols
			const fill = topLeftScalar(args[3]?.value ?? errorValue('#N/A'))
			const rows: ScalarCellValue[][] = []
			for (let row = 0; row < targetRows; row++) {
				const source = data[row] ?? []
				const next: ScalarCellValue[] = []
				for (let col = 0; col < targetCols; col++) {
					next.push(source[col] ?? fill)
				}
				rows.push(next)
			}
			return scalarOrArray(rows)
		},
	},
	{
		name: 'CHOOSECOLS',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			const indices: number[] = []
			for (let i = 1; i < args.length; i++) {
				const col = num(args[i])
				if (typeof col !== 'number') return col
				const resolved = resolveSelectionIndex(Math.round(col), (data[0] ?? []).length)
				if (typeof resolved !== 'number') return resolved
				indices.push(resolved)
			}
			const chosen = chooseCols(data, indices)
			return 'kind' in chosen ? chosen : scalarOrArray(chosen)
		},
	},
	{
		name: 'CHOOSEROWS',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const data = getRange(args[0])
			const indices: number[] = []
			for (let i = 1; i < args.length; i++) {
				const row = num(args[i])
				if (typeof row !== 'number') return row
				const resolved = resolveSelectionIndex(Math.round(row), data.length)
				if (typeof resolved !== 'number') return resolved
				indices.push(resolved)
			}
			const chosen = chooseRows(data, indices)
			return 'kind' in chosen ? chosen : scalarOrArray(chosen)
		},
	},
	{
		name: 'LAMBDA',
		minArgs: 2,
		maxArgs: 255,
		evaluate() {
			return errorValue('#CALC!')
		},
	},
	{
		name: 'MAP',
		minArgs: 2,
		maxArgs: 2,
		evaluate(args) {
			return args[0]?.value ?? EMPTY
		},
	},
	{
		name: 'REDUCE',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			return args[0]?.value ?? EMPTY
		},
	},
	{
		name: 'SCAN',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			return args[0]?.value ?? EMPTY
		},
	},
]
