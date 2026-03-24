import { AscendException, ascendError, type CellValue, type Operation } from '@ascend/schema'
import {
	Ascend,
	type CompactRangeWindowInfo,
	ensureOutputExtension,
	escapeDelimitedCell,
	formatDisplayCellValue,
	inferExportFormat,
	normalizeExportFormat,
	type RangeRowsInfo,
	toA1Ref,
	WorkbookDocument,
} from '@ascend/sdk'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'
import { parseOperations } from './operation-schema.ts'
import { errorResponse, okResponse } from './response.ts'

export function createServer(): McpServer {
	const server = new McpServer({
		name: 'ascend',
		version: '0.0.0',
	})

	server.tool(
		'ascend.inspect',
		'Inspect workbook or sheet metadata',
		{
			file: z.string().describe('Path to workbook file'),
			sheet: z.string().optional().describe('Sheet name to inspect'),
		},
		async ({ file, sheet }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'values', sheets: [sheet] } : { mode: 'metadata-only' },
				)
				if (sheet) {
					const data = wb.inspectSheet(sheet)
					if (!data) {
						return errorResponse(
							sheetNotFoundError(sheet, await loadAvailableSheets(file, wb.sheets)),
						)
					}
					return {
						...okResponse(data, `Inspected sheet "${sheet}"`),
					}
				}
				const data = wb.inspect()
				return okResponse(data, `Inspected workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.list_sheets',
		'List all sheet names and tables in a workbook (lightweight alternative to inspect)',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
				const info = wb.inspect()
				const sheets = info.sheets.map((s) => ({
					name: s.name,
					rows: s.rowCount,
					cols: s.colCount,
					tableCount: s.tableCount,
				}))
				return okResponse({ sheets }, `${sheets.length} sheet(s) in "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.read',
		'Read cell values from a range',
		{
			file: z.string().describe('Path to workbook file'),
			range: z.string().describe('Cell range (e.g. "A1:C10")'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			rowOffset: z.number().int().nonnegative().optional().describe('Row offset within the range'),
			rowLimit: z.number().int().positive().optional().describe('Maximum rows to return'),
			format: z
				.enum(['cells', 'rows', 'objects', 'tsv', 'compact'])
				.optional()
				.describe(
					'Read format: cell records, row arrays, object rows, TSV text, or sparse compact cells',
				),
			display: z
				.boolean()
				.optional()
				.describe('Return display strings instead of raw typed values'),
			headers: z
				.array(z.string())
				.optional()
				.describe('Explicit headers for object mode; defaults to first-row headers'),
		},
		async ({ file, range, sheet, rowOffset, rowLimit, format, display, headers }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'values', sheets: [sheet] } : { mode: 'values' },
				)
				const sheetName = sheet ?? wb.sheets[0]
				if (!sheetName) {
					return errorResponse('No sheets in workbook')
				}
				const handle = wb.sheet(sheetName)
				if (!handle) {
					return errorResponse(
						sheetNotFoundError(sheetName, await loadAvailableSheets(file, wb.sheets)),
					)
				}
				const readOpts = {
					...(rowOffset !== undefined ? { rowOffset } : {}),
					...(rowLimit !== undefined ? { rowLimit } : {}),
				}
				const mode = format ?? 'cells'
				const info =
					mode === 'compact'
						? buildCompactReadResult(
								handle.readWindowCompact(range, {
									...readOpts,
									includeRefs: false,
									omitEmpty: true,
									flatValues: true,
								}),
							)
						: mode === 'tsv'
							? buildTsvReadResult(handle.readRows(range, readOpts))
							: mode === 'rows'
								? handle.readRows(range, readOpts)
								: mode === 'objects'
									? handle.readObjects(range, {
											...readOpts,
											headers: headers && headers.length > 0 ? headers : 'first-row',
										})
									: handle.readWindow(range, readOpts)
				return okResponse(
					mode === 'tsv' || mode === 'compact'
						? info
						: display
							? displayReadResult(mode, info)
							: info,
					`Read range ${range} from "${sheetName}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.find',
		'Find cells by value or formula text',
		{
			file: z.string().describe('Path to workbook file'),
			query: z.string().min(1).describe('Text to search for'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			in: z
				.enum(['value', 'formula', 'both'])
				.optional()
				.describe('Search values, formulas, or both'),
			caseSensitive: z.boolean().optional().describe('Match case exactly'),
			limit: z.number().int().positive().max(500).optional().describe('Maximum matches to return'),
		},
		async ({ file, query, sheet, in: searchIn, caseSensitive, limit }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'formula', sheets: [sheet] } : { mode: 'formula' },
				)
				const sheetName = sheet ?? wb.sheets[0]
				if (!sheetName) return errorResponse('No sheets in workbook')
				const handle = wb.sheet(sheetName)
				if (!handle) {
					return errorResponse(
						sheetNotFoundError(sheetName, await loadAvailableSheets(file, wb.sheets)),
					)
				}
				const usedRange = handle.usedRange()
				if (!usedRange) {
					return okResponse(
						{ sheet: sheetName, query, in: searchIn ?? 'both', matches: [], truncated: false },
						`No populated cells to search in "${sheetName}"`,
					)
				}
				const cells = handle.rangeCompact(rangeRefToString(usedRange), {
					includeRefs: true,
				}).cells
				const searchMode = searchIn ?? 'both'
				const maxMatches = limit ?? 100
				const matches: Array<{
					ref: string
					row: number
					col: number
					value: string
					formula: string | null
					matchedOn: 'value' | 'formula'
				}> = []
				for (const cell of cells) {
					const ref = cell.ref
					if (!ref) continue
					const valueText = formatDisplayCellValue(cell.value)
					const formulaText = cell.formula ?? ''
					const matchValue =
						(searchMode === 'value' || searchMode === 'both') &&
						includesQuery(valueText, query, caseSensitive ?? false)
					const matchFormula =
						(searchMode === 'formula' || searchMode === 'both') &&
						cell.formula !== null &&
						includesQuery(formulaText, query, caseSensitive ?? false)
					if (!matchValue && !matchFormula) continue
					matches.push({
						ref,
						row: cell.row,
						col: cell.col,
						value: valueText,
						formula: cell.formula,
						matchedOn: matchFormula ? 'formula' : 'value',
					})
					if (matches.length >= maxMatches) break
				}
				return okResponse(
					{
						sheet: sheetName,
						query,
						in: searchMode,
						caseSensitive: caseSensitive ?? false,
						limit: maxMatches,
						truncated: matches.length >= maxMatches && cells.length > matches.length,
						matches,
					},
					`Found ${matches.length} matching cells in "${sheetName}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.agent_view',
		'Read a compressed semantic summary for a worksheet range',
		{
			file: z.string().describe('Path to workbook file'),
			range: z.string().describe('Cell range (e.g. "A1:Z200")'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			rowChunkSize: z.number().int().positive().optional().describe('Rows per streamed chunk'),
			sampleRowLimit: z.number().int().positive().optional().describe('Maximum sample rows'),
			sampleValueLimit: z
				.number()
				.int()
				.positive()
				.optional()
				.describe('Maximum sample values per column'),
		},
		async ({ file, range, sheet, rowChunkSize, sampleRowLimit, sampleValueLimit }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'formula', sheets: [sheet] } : { mode: 'formula' },
				)
				const sheetName = sheet ?? wb.sheets[0]
				if (!sheetName) return errorResponse('No sheets in workbook')
				const view = wb.agentView(sheetName, range, {
					...(rowChunkSize !== undefined ? { rowChunkSize } : {}),
					...(sampleRowLimit !== undefined ? { sampleRowLimit } : {}),
					...(sampleValueLimit !== undefined ? { sampleValueLimit } : {}),
				})
				if (!view) {
					return errorResponse(
						sheetNotFoundError(sheetName, await loadAvailableSheets(file, wb.sheets)),
					)
				}
				return okResponse(view, `Generated agent view for ${range} on "${sheetName}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.preview',
		'Preview operations without saving the workbook',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z.array(z.record(z.string(), z.unknown())).describe('Operations to preview'),
		},
		async ({ file, ops }) => {
			const parsed = parseOperations(ops)
			if (!parsed.ok) {
				return errorResponse(ascendError('VALIDATION_ERROR', parsed.error))
			}
			try {
				const wb = await Ascend.open(file)
				const result = wb.preview(parsed.value as readonly Operation[])
				if (result.errors.length > 0) {
					const first = result.errors[0]
					return errorResponse(
						first
							? {
									...first,
									details: { ...(first.details ?? {}), preview: result },
								}
							: ascendError('VALIDATION_ERROR', 'Preview failed', { details: { preview: result } }),
					)
				}
				return okResponse(result, `Previewed ${ops.length} operation(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.write',
		'Apply operations to a workbook',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z.array(z.record(z.string(), z.unknown())).describe('Operations to apply'),
		},
		async ({ file, ops }) => {
			const parsed = parseOperations(ops)
			if (!parsed.ok) {
				return errorResponse(ascendError('VALIDATION_ERROR', parsed.error))
			}
			try {
				const wb = await Ascend.open(file)
				const result = wb.apply(parsed.value as readonly Operation[])
				if (result.errors.length > 0) {
					const first = result.errors[0]
					return errorResponse(
						first ?? ascendError('VALIDATION_ERROR', 'Failed to apply operations'),
					)
				}
				if (result.recalcRequired) {
					const recalc = wb.recalc()
					if (recalc.errors.length > 0) {
						const first = recalc.errors[0]
						return errorResponse(
							first
								? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
										refs: [first.ref],
										details: { evalError: first.error },
									})
								: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
						)
					}
				}
				await wb.save(file)
				return okResponse(result, `Applied ${ops.length} operation(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.calc',
		'Recalculate all formulas in a workbook',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await Ascend.open(file)
				const result = wb.recalc()
				if (result.errors.length > 0) {
					const first = result.errors[0]
					return errorResponse(
						first
							? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
									refs: [first.ref],
									details: { evalError: first.error },
								})
							: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
					)
				}
				await wb.save(file)
				return okResponse(result, `Recalculated workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.list_operations',
		'List all available spreadsheet operations with their parameters and JSON Schema for LLM tool use',
		{},
		async () => {
			const ops = Ascend.listOperations()
			const schemas = Ascend.getOperationsSchema()
			return okResponse({ operations: ops, schemas }, `${ops.length} operations available`)
		},
	)

	server.tool(
		'ascend.check',
		'Run structural checks on a workbook',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'formula' })
				const result = wb.check()
				if (!result.valid) {
					const summary = `${result.issues.length} issue(s) found`
					return errorResponse(
						ascendError('VALIDATION_ERROR', summary, {
							details: { check: result },
						}),
					)
				}
				return okResponse(result, `Checked workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.lint',
		'Lint formulas for common issues',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'formula' })
				const result = wb.lint()
				return okResponse(result, `Linted workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.trace',
		'Trace cell dependencies (precedents and dependents)',
		{
			file: z.string().describe('Path to workbook file'),
			cell: z.string().describe('Cell reference (e.g. "Sheet1!A1" or "A1")'),
		},
		async ({ file, cell }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'formula' })
				const result = wb.trace(cell)
				if (!result) {
					return errorResponse(`Cannot trace "${cell}"`)
				}
				return okResponse(result, `Traced cell "${cell}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.diff',
		'Compute semantic diff between two workbooks',
		{
			fileA: z.string().describe('Path to first workbook'),
			fileB: z.string().describe('Path to second workbook'),
		},
		async ({ fileA, fileB }) => {
			try {
				const [a, b] = await Promise.all([Ascend.open(fileA), Ascend.open(fileB)])
				const result = a.diff(b)
				return okResponse(result, `Diffed "${fileA}" against "${fileB}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.export',
		'Export workbook to another format',
		{
			file: z.string().describe('Path to source workbook'),
			output: z.string().describe('Output file path'),
			format: z.string().optional().describe('Output format (inferred from extension if omitted)'),
		},
		async ({ file, output, format }) => {
			try {
				const wb = await Ascend.open(file)
				const normalized = format ? normalizeExportFormat(format) : inferExportFormat(output)
				if (!normalized) return errorResponse(`Unsupported format: ${format ?? output}`)
				const target = ensureOutputExtension(output, normalized)
				if (normalized === 'json') {
					await Bun.write(target, JSON.stringify(wb.toJSON(), null, 2))
				} else if (normalized === 'csv' || normalized === 'tsv') {
					const text = wb.toCsv(normalized === 'tsv' ? { dialect: { delimiter: '\t' } } : undefined)
					await Bun.write(target, text)
				} else {
					await wb.save(target)
				}
				return okResponse({ exported: target }, `Exported workbook to "${target}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	return server
}

function displayReadResult(mode: 'cells' | 'rows' | 'objects', info: unknown): unknown {
	if (mode === 'rows') {
		const rowsInfo = info as { rows: readonly (readonly CellValue[])[] }
		return {
			...rowsInfo,
			rows: rowsInfo.rows.map((row) => row.map((cell) => formatDisplayCellValue(cell))),
		}
	}
	if (mode === 'objects') {
		const objectInfo = info as {
			headers: readonly string[]
			rows: readonly Readonly<Record<string, CellValue>>[]
		}
		return {
			...objectInfo,
			rows: objectInfo.rows.map((row) =>
				Object.fromEntries(
					Object.entries(row).map(([key, value]) => [key, formatDisplayCellValue(value)]),
				),
			),
		}
	}
	const cellInfo = info as { cells: readonly { ref: string; value: CellValue }[] }
	return {
		...cellInfo,
		cells: cellInfo.cells.map((cell) => ({
			...cell,
			value: formatDisplayCellValue(cell.value),
		})),
	}
}

function buildTsvReadResult(info: RangeRowsInfo) {
	const rows = info.rows.map((row) =>
		row.map((cell) => escapeDelimitedCell(formatDisplayCellValue(cell), '\t')),
	)
	return {
		requestedRef: info.requestedRef,
		ref: info.ref,
		rowCount: info.rowCount,
		colCount: info.colCount,
		rowOffset: info.rowOffset,
		rowLimit: info.rowLimit,
		hasMore: info.hasMore,
		...(info.nextRowOffset !== undefined ? { nextRowOffset: info.nextRowOffset } : {}),
		format: 'tsv' as const,
		tsv: rows.map((row) => row.join('\t')).join('\n'),
	}
}

function buildCompactReadResult(info: CompactRangeWindowInfo) {
	return {
		requestedRef: info.requestedRef,
		ref: info.ref,
		rowCount: info.rowCount,
		colCount: info.colCount,
		rowOffset: info.rowOffset,
		rowLimit: info.rowLimit,
		hasMore: info.hasMore,
		...(info.nextRowOffset !== undefined ? { nextRowOffset: info.nextRowOffset } : {}),
		...(info.changeToken !== undefined ? { changeToken: info.changeToken } : {}),
		format: 'compact' as const,
		cells: info.cells.map((cell) => [
			cell.row - info.ref.start.row,
			cell.col - info.ref.start.col,
			cell.value as unknown,
			...(cell.formula ? [cell.formula] : []),
		]),
	}
}

async function loadAvailableSheets(
	file: string,
	fallbackSheets: readonly string[],
): Promise<readonly string[]> {
	if (fallbackSheets.length > 0) return fallbackSheets
	try {
		const workbook = await WorkbookDocument.open(file, { mode: 'metadata-only' })
		return workbook.sheets
	} catch {
		return fallbackSheets
	}
}

function sheetNotFoundError(sheetName: string, availableSheets: readonly string[]) {
	return ascendError('SHEET_NOT_FOUND', `Sheet "${sheetName}" not found`, {
		details: { availableSheets },
		suggestedFix:
			availableSheets.length > 0
				? `Use one of the available sheets: ${availableSheets.join(', ')}`
				: 'Inspect the workbook first to list available sheets.',
	})
}

function includesQuery(value: string, query: string, caseSensitive: boolean): boolean {
	if (caseSensitive) return value.includes(query)
	return value.toLowerCase().includes(query.toLowerCase())
}

function rangeRefToString(ref: {
	start: { row: number; col: number }
	end: { row: number; col: number }
}): string {
	return `${toA1Ref(ref.start.row, ref.start.col)}:${toA1Ref(ref.end.row, ref.end.col)}`
}

if (import.meta.main) {
	const server = createServer()
	const transport = new StdioServerTransport()
	server.connect(transport).catch((err) => {
		process.stderr.write(`${err}\n`)
		process.exit(1)
	})
}
