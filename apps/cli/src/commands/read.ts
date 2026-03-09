import type { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { formatCellValue, table } from '../output/pretty.ts'
import { openWorkbookWithProgress } from '../progress.ts'

export const usage = `Usage: ascend read <file> <selector> [flags]

  Read cell values from a range, table, or defined name.

Arguments:
  <file>              Path to the workbook file
  <selector>          Cell range (A1:B2), table (table:Name), or defined name (name:Name)

Flags:
  --sheet <name>      Sheet name when selector is a plain range
  --row-offset <N>    Start reading from row offset N
  --row-limit <N>     Limit number of rows returned
  --json              Output as JSON
`

export async function readCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const selectorArg = args[1]
	if (!file || !selectorArg) {
		console.error('Usage: ascend read <file> <selector> [--sheet <name>]')
		return 1
	}

	const requestedSheet = flags.get('sheet')
	const selector = parseSelector(selectorArg, requestedSheet)
	const { workbook: wb } = await openWorkbookWithProgress(file, inferOpenOptions(selector))

	const rowOffset = parseOptionalInt(flags.get('row-offset'))
	const rowLimit = parseOptionalInt(flags.get('row-limit'))

	switch (selector.kind) {
		case 'table': {
			const handle = wb.table(selector.name)
			if (!handle) {
				console.error(`Table "${selector.name}" not found`)
				return 1
			}
			const rows = handle.rows(rowLimit !== undefined ? { limit: rowLimit } : undefined)
			if (flags.has('json')) {
				console.log(
					jsonOut({
						kind: 'table',
						name: handle.name,
						columns: handle.columns,
						rowCount: handle.rowCount,
						rows,
					}),
				)
				return 0
			}
			const grid = rows.map((row) =>
				handle.columns.map((column) => formatCellValue(row[column] ?? { kind: 'empty' })),
			)
			console.log(table([...handle.columns], grid))
			return 0
		}
		case 'name': {
			const handle = wb.definedName(selector.name, selector.sheetName)
			if (!handle) {
				console.error(`Defined name "${selector.name}" not found`)
				return 1
			}
			const resolvedRange = parseNamedRangeFormula(handle.formula)
			if (flags.has('json')) {
				console.log(
					jsonOut({
						kind: 'name',
						name: handle.name,
						formula: handle.formula,
						scope: handle.scope,
						sheet: handle.sheet,
						resolvedRange,
					}),
				)
				return 0
			}
			console.log(`${handle.name}: ${handle.formula}`)
			if (!resolvedRange) return 0
			return readRangeLike(wb, resolvedRange.sheet, resolvedRange.range, rowOffset, rowLimit)
		}
		case 'range':
			return readRangeLike(
				wb,
				resolveSheetName(wb, selector.sheet),
				selector.range,
				rowOffset,
				rowLimit,
				flags.has('json'),
			)
	}
}

function inferOpenOptions(selector: ReadSelector): { sheets?: readonly string[] } | undefined {
	if (selector.kind === 'range' && selector.sheet) {
		return { sheets: [selector.sheet] }
	}
	if (selector.kind === 'name' && selector.sheetName) {
		return { sheets: [selector.sheetName] }
	}
	return undefined
}

function readRangeLike(
	wb: AscendWorkbook,
	sheetName: string | undefined,
	range: string,
	rowOffset: number | undefined,
	rowLimit: number | undefined,
	asJson = false,
): number {
	if (!sheetName) {
		console.error(
			wb.sheets.length === 0
				? 'No sheets in workbook'
				: 'Multiple sheets available; specify a sheet explicitly',
		)
		return 1
	}

	const sheet = wb.sheet(sheetName)
	if (!sheet) {
		console.error(`Sheet "${sheetName}" not found`)
		return 1
	}

	const info = sheet.readWindow(range, {
		...(rowOffset !== undefined ? { rowOffset } : {}),
		...(rowLimit !== undefined ? { rowLimit } : {}),
	})

	if (asJson) {
		console.log(jsonOut(info))
		return 0
	}

	const grid: string[][] = []
	const cellMap = new Map(info.cells.map((cell) => [`${cell.row}:${cell.col}`, cell] as const))
	for (let r = 0; r < info.rowCount; r++) {
		const row: string[] = []
		for (let c = 0; c < info.colCount; c++) {
			const cell = cellMap.get(`${info.ref.start.row + r}:${info.ref.start.col + c}`)
			row.push(cell ? formatCellValue(cell.value) : '')
		}
		grid.push(row)
	}

	const headers = Array.from({ length: info.colCount }, (_, i) =>
		columnLabel(info.ref.start.col + i),
	)
	console.log(table(headers, grid))
	if (info.hasMore) {
		console.log(`\nMore rows available. Re-run with --row-offset ${info.nextRowOffset}.`)
	}
	return 0
}

type ReadSelector =
	| { readonly kind: 'range'; readonly range: string; readonly sheet?: string }
	| { readonly kind: 'table'; readonly name: string }
	| { readonly kind: 'name'; readonly name: string; readonly sheetName?: string }

function parseSelector(selector: string, requestedSheet: string | undefined): ReadSelector {
	if (selector.startsWith('table:')) {
		return { kind: 'table', name: selector.slice('table:'.length) }
	}
	if (selector.startsWith('name:')) {
		const raw = selector.slice('name:'.length)
		const split = splitSheetQualifier(raw)
		return split
			? { kind: 'name', name: split.value, sheetName: split.sheet }
			: { kind: 'name', name: raw }
	}
	const split = splitSheetQualifier(selector)
	if (split) {
		return { kind: 'range', sheet: split.sheet, range: split.value }
	}
	return requestedSheet
		? { kind: 'range', sheet: requestedSheet, range: selector }
		: { kind: 'range', range: selector }
}

function splitSheetQualifier(input: string): { sheet: string; value: string } | undefined {
	const bang = input.lastIndexOf('!')
	if (bang === -1) return undefined
	const sheet = input.slice(0, bang).replace(/^'|'$/g, '')
	const value = input.slice(bang + 1)
	if (!sheet || !value) return undefined
	return { sheet, value }
}

function resolveSheetName(
	wb: AscendWorkbook,
	explicitSheet: string | undefined,
): string | undefined {
	if (explicitSheet) return explicitSheet
	return wb.sheets.length === 1 ? wb.sheets[0] : undefined
}

function parseNamedRangeFormula(formula: string): { sheet: string; range: string } | undefined {
	const match = /^'?([^']+)'?!([A-Za-z]+\d+(?::[A-Za-z]+\d+)?)$/.exec(formula)
	if (!match) return undefined
	const sheet = match[1]
	const range = match[2]
	if (!sheet || !range) return undefined
	return { sheet, range }
}

function parseOptionalInt(value: string | undefined): number | undefined {
	if (value === undefined || value === '') return undefined
	const parsed = Number.parseInt(value, 10)
	return Number.isNaN(parsed) ? undefined : parsed
}

function columnLabel(col: number): string {
	let n = col
	let label = ''
	while (n >= 0) {
		label = String.fromCharCode(65 + (n % 26)) + label
		n = Math.floor(n / 26) - 1
	}
	return label
}
