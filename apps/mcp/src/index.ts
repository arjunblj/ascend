import type { Operation } from '@ascend/schema'
import { Ascend } from '@ascend/sdk'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'

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
			const wb = await Ascend.open(file, sheet ? { sheets: [sheet] } : { mode: 'metadata-only' })
			if (sheet) {
				const handle = wb.sheet(sheet)
				if (!handle) {
					return {
						content: [
							{ type: 'text', text: JSON.stringify({ error: `Sheet "${sheet}" not found` }) },
						],
						isError: true,
					}
				}
				return {
					content: [
						{
							type: 'text',
							text: JSON.stringify({
								name: handle.name,
								rowCount: handle.rowCount,
								colCount: handle.colCount,
								usedRange: handle.usedRange(),
							}),
						},
					],
				}
			}
			return { content: [{ type: 'text', text: JSON.stringify(wb.inspect()) }] }
		},
	)

	server.tool(
		'ascend.read',
		'Read cell values from a range',
		{
			file: z.string().describe('Path to workbook file'),
			range: z.string().describe('Cell range (e.g. "A1:C10")'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
		},
		async ({ file, range, sheet }) => {
			const wb = await Ascend.open(file)
			const sheetName = sheet ?? wb.sheets[0]
			if (!sheetName) {
				return {
					content: [{ type: 'text', text: JSON.stringify({ error: 'No sheets in workbook' }) }],
					isError: true,
				}
			}
			const handle = wb.sheet(sheetName)
			if (!handle) {
				return {
					content: [
						{ type: 'text', text: JSON.stringify({ error: `Sheet "${sheetName}" not found` }) },
					],
					isError: true,
				}
			}
			const info = handle.range(range)
			return { content: [{ type: 'text', text: JSON.stringify(info) }] }
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
				content: [{ type: 'text', text: JSON.stringify(result) }],
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
			return { content: [{ type: 'text', text: JSON.stringify(result) }] }
		},
	)

	server.tool(
		'ascend.check',
		'Run structural checks on a workbook',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			const wb = await Ascend.open(file)
			const result = wb.check()
			return { content: [{ type: 'text', text: JSON.stringify(result) }], isError: !result.valid }
		},
	)

	server.tool(
		'ascend.lint',
		'Lint formulas for common issues',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			const wb = await Ascend.open(file)
			const result = wb.lint()
			return { content: [{ type: 'text', text: JSON.stringify(result) }] }
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
			const wb = await Ascend.open(file)
			const result = wb.trace(cell)
			if (!result) {
				return {
					content: [{ type: 'text', text: JSON.stringify({ error: `Cannot trace "${cell}"` }) }],
					isError: true,
				}
			}
			return { content: [{ type: 'text', text: JSON.stringify(result) }] }
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
			return { content: [{ type: 'text', text: JSON.stringify(result) }] }
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
			return { content: [{ type: 'text', text: JSON.stringify({ exported: target }) }] }
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
