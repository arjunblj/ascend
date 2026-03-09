import type { Operation } from '@ascend/schema'
import { Ascend, WorkbookSession } from '@ascend/sdk'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'

function okResponse<T>(data: T, summary: string) {
	return {
		content: [{ type: 'text' as const, text: summary }],
		structuredContent: data as Record<string, unknown>,
	}
}

function errorResponse(message: string) {
	return {
		content: [{ type: 'text' as const, text: message }],
		structuredContent: { error: message },
		isError: true,
	}
}

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
			const wb = await WorkbookSession.open(
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
		},
		async ({ file, range, sheet, rowOffset, rowLimit }) => {
			const wb = await WorkbookSession.open(
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
			const info = handle.readWindow(range, {
				...(rowOffset !== undefined ? { rowOffset } : {}),
				...(rowLimit !== undefined ? { rowLimit } : {}),
			})
			return okResponse(info, `Read range ${range} from "${sheetName}"`)
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
			const wb = await Ascend.open(file)
			const result = wb.apply(ops as unknown as readonly Operation[])
			if (result.errors.length === 0) {
				await wb.save(file)
			}
			return {
				...okResponse(result, `Applied ${ops.length} operation(s)`),
				isError: result.errors.length > 0,
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
			const wb = await Ascend.open(file)
			const result = wb.recalc()
			await wb.save(file)
			return okResponse(result, `Recalculated workbook "${file}"`)
		},
	)

	server.tool(
		'ascend.check',
		'Run structural checks on a workbook',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			const wb = await WorkbookSession.open(file, { mode: 'formula' })
			const result = wb.check()
			return { ...okResponse(result, `Checked workbook "${file}"`), isError: !result.valid }
		},
	)

	server.tool(
		'ascend.lint',
		'Lint formulas for common issues',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			const wb = await WorkbookSession.open(file, { mode: 'formula' })
			const result = wb.lint()
			return okResponse(result, `Linted workbook "${file}"`)
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
			const wb = await WorkbookSession.open(file, { mode: 'formula' })
			const result = wb.trace(cell)
			if (!result) {
				return errorResponse(`Cannot trace "${cell}"`)
			}
			return okResponse(result, `Traced cell "${cell}"`)
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
			const [a, b] = await Promise.all([Ascend.open(fileA), Ascend.open(fileB)])
			const result = a.diff(b)
			return okResponse(result, `Diffed "${fileA}" against "${fileB}"`)
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
			const wb = await Ascend.open(file)
			const target = format ? `${output.replace(/\.[^.]+$/, '')}.${format}` : output
			await wb.save(target)
			return okResponse({ exported: target }, `Exported workbook to "${target}"`)
		},
	)

	return server
}

if (import.meta.main) {
	const server = createServer()
	const transport = new StdioServerTransport()
	server.connect(transport).catch((err) => {
		process.stderr.write(`${err}\n`)
		process.exit(1)
	})
}
