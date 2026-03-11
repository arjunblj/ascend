import type { CompactCellInfo, WorkbookDocument } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { formatCellValue, table } from '../output/pretty.ts'
import { openWorkbookDocumentWithProgress } from '../progress.ts'

export const usage = `Usage: ascend read <file> <selector> [flags]

  Read cell values from a range, table, or defined name.

Arguments:
  <file>              Path to the workbook file
  <selector>          Cell range (A1:B2), table (table:Name), or defined name (name:Name)

Flags:
  --sheet <name>      Sheet name when selector is a plain range
  --mode <mode>       Load mode: values or full
  --row-offset <N>    Start reading from row offset N
  --row-limit <N>     Limit number of rows returned (default: 1000 for pretty output)
  --display           Render display strings for user-facing output
  --json              Output as JSON
`

export async function readCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const selectorArg = args[1]
	if (!file || !selectorArg) {
		cliError('Usage: ascend read <file> <selector> [--sheet <name>]', flags)
		return 1
	}

	const requestedSheet = flags.get('sheet')
	const explicitMode = parseReadMode(flags.get('mode'))
	if (flags.has('mode') && explicitMode === null) {
		cliError('Invalid --mode. Use one of: values, full', flags)
		return 1
	}
	const selector = parseSelector(selectorArg, requestedSheet)
	const { document: wb } = await openWorkbookDocumentWithProgress(
		file,
		inferOpenOptions(selector, explicitMode ?? 'values'),
	)

	const rowOffset = parseOptionalInt(flags.get('row-offset'))
	if (flags.has('row-offset') && rowOffset == null) {
		cliError('Invalid --row-offset. Use a non-negative integer.', flags)
		return 1
	}
	const rowLimit = parseOptionalInt(flags.get('row-limit'))
	if (flags.has('row-limit') && (rowLimit == null || rowLimit < 1)) {
		cliError('Invalid --row-limit. Use a positive integer.', flags)
		return 1
	}
	const validatedRowOffset = rowOffset ?? undefined
	const prettyDefaultLimit = 1000
	const validatedRowLimit = rowLimit ?? (flags.has('json') ? undefined : prettyDefaultLimit)
	const display = flags.has('display')

	switch (selector.kind) {
		case 'table': {
			const handle = wb.table(selector.name)
			if (!handle) {
				cliError(`Table "${selector.name}" not found`, flags)
				return 1
			}
			const page = handle.readRows({
				...(validatedRowOffset !== undefined ? { offset: validatedRowOffset } : {}),
				...(validatedRowLimit !== undefined ? { limit: validatedRowLimit } : {}),
			})
			if (flags.has('json')) {
				console.log(
					jsonOut({
						kind: 'table',
						name: handle.name,
						columns: handle.columns,
						rowCount: handle.rowCount,
						hasHeaders: handle.hasHeaders,
						hasTotals: handle.hasTotals,
						headerRow: handle.headerRow(),
						totalsRow: handle.totalsRow(),
						sortState: handle.sortState,
						autoFilter: handle.autoFilter,
						page,
						rows: page.rows.map((row) => row.values),
					}),
				)
				return 0
			}
			const grid = page.rows.map((row) =>
				handle.columns.map((column) =>
					formatCellValue(row.values[column] ?? { kind: 'empty' }, { display }),
				),
			)
			console.log(table([...handle.columns], grid))
			if (page.hasMore) {
				console.log(`\nMore rows available. Re-run with --row-offset ${page.nextRowOffset}.`)
			}
			return 0
		}
		case 'name': {
			const handle = wb.definedName(selector.name, selector.sheetName)
			if (!handle) {
				cliError(`Defined name "${selector.name}" not found`, flags)
				return 1
			}
			const resolvedRange = resolveNamedRead(handle)
			if (flags.has('json')) {
				console.log(
					jsonOut({
						kind: 'name',
						name: handle.name,
						formula: handle.formula,
						normalizedFormula: handle.normalizedFormula,
						scope: handle.scope,
						sheet: handle.sheet,
						references: handle.references,
						functions: handle.functions,
						volatile: handle.volatile,
						parseError: handle.parseError,
						resolutionKind: describeNamedReadResolution(handle),
						resolvedRange,
					}),
				)
				return 0
			}
			console.log(`${handle.name}: ${handle.formula}`)
			console.log(`  normalized: ${handle.normalizedFormula}`)
			console.log(`  scope: ${handle.scope}${handle.sheet ? ` (${handle.sheet})` : ''}`)
			if (handle.references.length > 0) {
				console.log(`  refs: ${handle.references.map((reference) => reference.text).join(', ')}`)
			}
			if (handle.functions.length > 0) {
				console.log(`  functions: ${handle.functions.join(', ')}`)
			}
			if (handle.parseError) {
				console.log(`  parse-error: ${handle.parseError}`)
			}
			if (!resolvedRange) return 0
			return readRangeLike(
				wb,
				resolvedRange.sheet,
				resolvedRange.range,
				validatedRowOffset,
				validatedRowLimit,
				flags.has('json'),
				display,
				flags,
			)
		}
		case 'range':
			return readRangeLike(
				wb,
				resolveSheetName(wb, selector.sheet),
				selector.range,
				validatedRowOffset,
				validatedRowLimit,
				flags.has('json'),
				display,
				flags,
			)
	}
}

function inferOpenOptions(
	selector: ReadSelector,
	mode: 'values' | 'full',
): { mode: 'values' | 'full'; sheets?: readonly string[] } {
	if (selector.kind === 'range' && selector.sheet) {
		return { mode, sheets: [selector.sheet] }
	}
	if (selector.kind === 'name' && selector.sheetName) {
		return { mode, sheets: [selector.sheetName] }
	}
	return { mode }
}

function readRangeLike(
	wb: WorkbookDocument,
	sheetName: string | undefined,
	range: string,
	rowOffset: number | undefined,
	rowLimit: number | undefined,
	asJson = false,
	display = false,
	flags: Map<string, string> = new Map(),
): number {
	if (!sheetName) {
		cliError(
			wb.sheets.length === 0
				? 'No sheets in workbook'
				: 'Multiple sheets available; specify a sheet explicitly',
			flags,
		)
		return 1
	}

	const sheet = wb.sheet(sheetName)
	if (!sheet) {
		cliError(`Sheet "${sheetName}" not found`, flags)
		return 1
	}

	if (asJson) {
		const info = sheet.readWindow(normalizeReadableRange(range), {
			...(rowOffset !== undefined ? { rowOffset } : {}),
			...(rowLimit !== undefined ? { rowLimit } : {}),
		})
		console.log(jsonOut(info))
		return 0
	}

	const info = sheet.readWindowCompact(normalizeReadableRange(range), {
		...(rowOffset !== undefined ? { rowOffset } : {}),
		...(rowLimit !== undefined ? { rowLimit } : {}),
		includeRefs: false,
	})
	const grid = buildWindowGrid(
		info.cells,
		info.rowCount,
		info.colCount,
		info.ref.start.row,
		info.ref.start.col,
		display,
	)

	const headers = Array.from({ length: info.colCount }, (_, i) =>
		columnLabel(info.ref.start.col + i),
	)
	console.log(table(headers, grid))
	if (info.hasMore) {
		console.log(`\nMore rows available. Re-run with --row-offset ${info.nextRowOffset}.`)
	}
	return 0
}

function buildWindowGrid(
	cells: readonly CompactCellInfo[],
	rowCount: number,
	colCount: number,
	startRow: number,
	startCol: number,
	display: boolean,
): string[][] {
	const grid: string[][] = []
	let index = 0
	for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
		const rowIndex = startRow + rowOffset
		const row: string[] = []
		for (let colOffset = 0; colOffset < colCount; colOffset++) {
			const colIndex = startCol + colOffset
			const cell = cells[index]
			if (cell && cell.row === rowIndex && cell.col === colIndex) {
				row.push(formatCellValue(cell.value, { display }))
				index += 1
			} else {
				row.push('')
			}
		}
		grid.push(row)
	}
	return grid
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
	wb: WorkbookDocument,
	explicitSheet: string | undefined,
): string | undefined {
	if (explicitSheet) return explicitSheet
	return wb.sheets.length === 1 ? wb.sheets[0] : undefined
}

function parseNamedRangeFormula(formula: string): { sheet: string; range: string } | undefined {
	const match = /^'?([^']+)'?!(.+)$/.exec(formula)
	if (!match) return undefined
	const sheet = match[1]
	const range = match[2]
	if (!sheet || !range || !isSimpleRangeReference(range)) return undefined
	return { sheet, range }
}

function resolveNamedRead(name: {
	formula: string
	references: readonly { kind: string; text: string; scope?: { kind: string; sheet?: string } }[]
	sheet?: string
}): { sheet: string; range: string } | undefined {
	const first = name.references[0]
	if (
		first &&
		(first.kind === 'cell' ||
			first.kind === 'range' ||
			first.kind === 'wholeRow' ||
			first.kind === 'wholeColumn')
	) {
		const bang = first.text.lastIndexOf('!')
		if (bang !== -1) {
			return {
				sheet: first.text.slice(0, bang).replace(/^'|'$/g, ''),
				range: first.text.slice(bang + 1),
			}
		}
		if (first.scope?.kind === 'sheet' && first.scope.sheet) {
			return { sheet: first.scope.sheet, range: first.text }
		}
		if (name.sheet) {
			return { sheet: name.sheet, range: first.text }
		}
	}
	return parseNamedRangeFormula(name.formula)
}

function describeNamedReadResolution(name: {
	references: readonly { kind: string }[]
	parseError?: string
}): string {
	if (name.parseError) return 'parse-error'
	return name.references[0]?.kind ?? 'unresolved'
}

function isSimpleRangeReference(value: string): boolean {
	return (
		/^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(value) ||
		/^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(value) ||
		/^\$?\d+:\$?\d+$/.test(value)
	)
}

function normalizeReadableRange(value: string): string {
	return value.replaceAll('$', '')
}

function parseOptionalInt(value: string | undefined): number | null | undefined {
	if (value === undefined || value === '') return undefined
	const parsed = Number.parseInt(value, 10)
	if (Number.isNaN(parsed)) return null
	return parsed
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

function parseReadMode(mode: string | undefined): 'values' | 'full' | undefined | null {
	if (mode === undefined || mode === '') return undefined
	switch (mode) {
		case 'values':
		case 'full':
			return mode
		default:
			return null
	}
}
