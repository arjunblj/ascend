import {
	AscendException,
	ascendError,
	type CellValue,
	machineFailure,
	type Operation,
} from '@ascend/schema'
import { Ascend, WorkbookDocument } from '@ascend/sdk'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'
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
						return errorResponse(`Sheet "${sheet}" not found`)
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
		'ascend.read',
		'Read cell values from a range',
		{
			file: z.string().describe('Path to workbook file'),
			range: z.string().describe('Cell range (e.g. "A1:C10")'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			rowOffset: z.number().int().nonnegative().optional().describe('Row offset within the range'),
			rowLimit: z.number().int().positive().optional().describe('Maximum rows to return'),
			format: z
				.enum(['cells', 'rows', 'objects'])
				.optional()
				.describe('Read format: cell records, row arrays, or object rows'),
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
					return errorResponse(`Sheet "${sheetName}" not found`)
				}
				const readOpts = {
					...(rowOffset !== undefined ? { rowOffset } : {}),
					...(rowLimit !== undefined ? { rowLimit } : {}),
				}
				const mode = format ?? 'cells'
				const info =
					mode === 'rows'
						? handle.readRows(range, readOpts)
						: mode === 'objects'
							? handle.readObjects(range, {
									...readOpts,
									headers: headers && headers.length > 0 ? headers : 'first-row',
								})
							: handle.readWindow(range, readOpts)
				return okResponse(
					display ? displayReadResult(mode, info) : info,
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
				if (!view) return errorResponse(`Sheet "${sheetName}" not found`)
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
			try {
				const wb = await Ascend.open(file)
				const result = wb.preview(ops as unknown as readonly Operation[])
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
			try {
				const wb = await Ascend.open(file)
				const result = wb.apply(ops as unknown as readonly Operation[])
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
					return {
						...errorResponse(
							ascendError('VALIDATION_ERROR', summary, {
								details: { issues: result.issues },
							}),
						),
						structuredContent: machineFailure(
							ascendError('VALIDATION_ERROR', summary, {
								details: { check: result },
							}),
						) as unknown as Record<string, unknown>,
					}
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

function normalizeExportFormat(format: string): 'csv' | 'tsv' | 'json' | 'xlsx' | 'xlsm' | null {
	switch (format.toLowerCase()) {
		case 'csv':
		case 'tsv':
		case 'json':
		case 'xlsx':
		case 'xlsm':
			return format.toLowerCase() as 'csv' | 'tsv' | 'json' | 'xlsx' | 'xlsm'
		default:
			return null
	}
}

function inferExportFormat(path: string): 'csv' | 'tsv' | 'json' | 'xlsx' | 'xlsm' | null {
	const ext = path.split('.').pop()?.toLowerCase() ?? ''
	return normalizeExportFormat(ext)
}

function ensureOutputExtension(
	output: string,
	format: 'csv' | 'tsv' | 'json' | 'xlsx' | 'xlsm',
): string {
	return output.endsWith(`.${format}`) ? output : `${output.replace(/\.[^.]+$/, '')}.${format}`
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

function formatDisplayCellValue(value: CellValue): string {
	switch (value.kind) {
		case 'empty':
			return ''
		case 'number':
			return String(value.value)
		case 'string':
			return value.value
		case 'boolean':
			return value.value ? 'TRUE' : 'FALSE'
		case 'error':
			return value.value
		case 'date': {
			const parts = serialToDateParts(Math.floor(value.serial))
			if (!parts) return `[date:${value.serial}]`
			return `${String(parts.year).padStart(4, '0')}-${String(parts.month).padStart(2, '0')}-${String(parts.day).padStart(2, '0')}`
		}
		case 'richText':
			return value.runs.map((run) => run.text).join('')
	}
	return ''
}

function serialToDateParts(
	serial: number,
	dateSystem: '1900' | '1904' = '1900',
): { year: number; month: number; day: number } | null {
	const msPerDay = 86_400_000
	const makeUtc = (year: number, month: number, day: number) => {
		const date = new Date(Date.UTC(year, month - 1, day))
		if (year >= 0 && year < 100) date.setUTCFullYear(year)
		return date
	}
	if (dateSystem === '1904') {
		if (serial < 0) return null
		const epoch = makeUtc(1904, 1, 1).getTime()
		const date = new Date(epoch + serial * msPerDay)
		return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
	}
	if (serial < 1) return null
	if (serial === 60) return { year: 1900, month: 2, day: 29 }
	const epoch = makeUtc(1900, 1, 1).getTime()
	const days = serial < 60 ? serial - 1 : serial - 2
	const date = new Date(epoch + days * msPerDay)
	return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
}

if (import.meta.main) {
	const server = createServer()
	const transport = new StdioServerTransport()
	server.connect(transport).catch((err) => {
		process.stderr.write(`${err}\n`)
		process.exit(1)
	})
}
