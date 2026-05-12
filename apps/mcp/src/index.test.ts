import { afterAll, describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook, parseOperations } from '@ascend/sdk'
import { makeXlsx } from '../../../packages/io-xlsx/test/helpers.ts'
import { createServer } from './index.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-mcp-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const PIVOT_FIXTURE = join(
	import.meta.dir,
	'../../../fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
)

afterAll(async () => {
	await unlink(TEMP_FILE).catch(() => {})
})

describe('MCP server', () => {
	test('createServer returns a McpServer instance', () => {
		const server = createServer()
		expect(server).toBeDefined()
		expect(server.server).toBeDefined()
	})

	test('all ascend tools are registered', () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing private internals for test verification
		const registered = (server as any)._registeredTools as Record<string, unknown>
		const names = Object.keys(registered)

		expect(names).toContain('ascend.inspect')
		expect(names).toContain('ascend.active_content')
		expect(names).toContain('ascend.search_docs')
		expect(names).toContain('ascend.search_examples')
		expect(names).toContain('ascend.read')
		expect(names).toContain('ascend.read_table')
		expect(names).toContain('ascend.find')
		expect(names).toContain('ascend.agent_view')
		expect(names).toContain('ascend.preview')
		expect(names).toContain('ascend.write')
		expect(names).toContain('ascend.calc')
		expect(names).toContain('ascend.eval')
		expect(names).toContain('ascend.check')
		expect(names).toContain('ascend.list_operations')
		expect(names).toContain('ascend.lint')
		expect(names).toContain('ascend.trace')
		expect(names).toContain('ascend.diff')
		expect(names).toContain('ascend.export')
		expect(names).toContain('ascend.list_sheets')
		expect(names).toContain('ascend.package_graph')
		expect(names).toContain('ascend.visuals')
		expect(names).toContain('ascend.pivots')
		expect(names).toContain('ascend.capabilities')
		expect(names).toContain('ascend.plan')
		expect(names).toContain('ascend.commit')
		expect(names).toContain('ascend.repair_plan')
		expect(names.length).toBe(26)
	})

	test('agent resources and prompts are registered', () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing private MCP registry internals for test verification
		const resources = (server as any)._registeredResources as Record<string, unknown>
		// biome-ignore lint/suspicious/noExplicitAny: accessing private MCP registry internals for test verification
		const prompts = (server as any)._registeredPrompts as Record<string, unknown>

		expect(Object.keys(resources)).toEqual([
			'ascend://llms.txt',
			'ascend://llms-full.txt',
			'ascend://docs/agent-api.md',
			'ascend://capabilities',
			'ascend://operations',
			'ascend://agent-workflow',
		])
		expect(Object.keys(prompts)).toContain('ascend.agent_workflow')
	})

	test('agent resources return canonical workflow context', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing private MCP registry internals for behavior testing
		const resources = (server as any)._registeredResources as Record<
			string,
			{ readCallback: (uri: URL) => Promise<{ contents: Array<{ text?: string }> }> }
		>

		const capabilities = await resources['ascend://capabilities']?.readCallback(
			new URL('ascend://capabilities'),
		)
		const llms = await resources['ascend://llms.txt']?.readCallback(new URL('ascend://llms.txt'))
		const llmsFull = await resources['ascend://llms-full.txt']?.readCallback(
			new URL('ascend://llms-full.txt'),
		)
		const agentApi = await resources['ascend://docs/agent-api.md']?.readCallback(
			new URL('ascend://docs/agent-api.md'),
		)
		const operations = await resources['ascend://operations']?.readCallback(
			new URL('ascend://operations'),
		)
		const workflow = await resources['ascend://agent-workflow']?.readCallback(
			new URL('ascend://agent-workflow'),
		)

		expect(llms?.contents[0]?.text).toContain('ascend.search_docs')
		expect(llmsFull?.contents[0]?.text).toContain('Ascend Full Agent Context')
		expect(agentApi?.contents[0]?.text).toContain('Ascend Agent API')
		expect(capabilities?.contents[0]?.text).toContain('"capabilities"')
		expect(operations?.contents[0]?.text).toContain('"schemas"')
		expect(workflow?.contents[0]?.text).toContain('ascend.plan')
	})

	test('documentation search tools return local docs and examples', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const docsHandler = (server as any)._registeredTools['ascend.search_docs'].handler as (args: {
			query: string
			limit?: number
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { results?: Array<{ path?: string }> } }
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const examplesHandler = (server as any)._registeredTools['ascend.search_examples']
			.handler as (args: { query: string; limit?: number }) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { results?: Array<{ kind?: string; path?: string }> }
			}
		}>

		const docs = await docsHandler({ query: 'plan commit allowLoss', limit: 3 })
		expect(docs.structuredContent?.ok).toBe(true)
		expect(
			docs.structuredContent?.data?.results?.some((result) => result.path?.includes('llms')),
		).toBe(true)

		const examples = await examplesHandler({ query: 'mcp setup cursor', limit: 3 })
		expect(examples.structuredContent?.ok).toBe(true)
		expect(
			examples.structuredContent?.data?.results?.every((result) => result.kind === 'example'),
		).toBe(true)
		expect(
			examples.structuredContent?.data?.results?.some(
				(result) => result.path === 'examples/mcp-setup.md',
			),
		).toBe(true)
	})

	test('agent workflow prompt includes safe plan and commit guidance', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing private MCP registry internals for behavior testing
		const prompts = (server as any)._registeredPrompts as Record<
			string,
			{
				callback: (args: { file?: string; task?: string }) => Promise<{
					messages: Array<{ content: { text?: string } }>
				}>
			}
		>

		const prompt = await prompts['ascend.agent_workflow']?.callback({
			file: 'book.xlsx',
			task: 'update the forecast',
		})
		const text = prompt?.messages[0]?.content.text ?? ''
		expect(text).toContain('Workbook: book.xlsx')
		expect(text).toContain('ascend.plan')
		expect(text).toContain('ascend.commit')
		expect(text).toContain('ascend.active_content')
		expect(text).toContain('allowLoss')
	})

	test('ascend.active_content exposes focused active-content provenance for agents', async () => {
		const activeFile = join(
			tmpdir(),
			`ascend-mcp-active-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
		)
		await Bun.write(activeFile, signedMacroWorkbook())
		try {
			const server = createServer()
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const handler = (server as any)._registeredTools['ascend.active_content'].handler as (args: {
				file: string
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						activeContentCount?: number
						activeContent?: Array<{
							kind?: string
							partPath?: string
							sourcePartPath?: string
							sourceRelationshipId?: string
							relType?: string
						}>
						compatibilityFeatures?: Array<{ feature?: string; locations?: readonly string[] }>
					}
				}
			}>

			const result = await handler({ file: activeFile })
			const data = result.structuredContent?.data

			expect(result.structuredContent?.ok).toBe(true)
			expect(data?.activeContentCount).toBe(2)
			expect(data?.activeContent).toContainEqual(
				expect.objectContaining({
					kind: 'vbaProject',
					partPath: 'xl/vbaProject.bin',
					sourceRelationshipId: 'rIdVba',
				}),
			)
			expect(data?.activeContent).toContainEqual(
				expect.objectContaining({
					kind: 'vbaSignature',
					partPath: 'xl/vbaProjectSignature.bin',
					sourcePartPath: 'xl/vbaProject.bin',
					sourceRelationshipId: 'rIdVbaSignature',
					relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature',
				}),
			)
			expect(data?.compatibilityFeatures).toContainEqual(
				expect.objectContaining({
					feature: 'preservedMacro',
					locations: ['xl/vbaProject.bin', 'xl/vbaProjectSignature.bin'],
				}),
			)
		} finally {
			await unlink(activeFile).catch(() => {})
		}
	})

	test('ascend.package_graph exposes package identity for agents', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.package_graph'].handler as (args: {
			file: string
		}) => Promise<{
			structuredContent?: {
				data?: {
					parts?: Array<{
						path?: string
						featureFamily?: string
						ownerScope?: string
						sourceRelationshipId?: string
					}>
					relationships?: Array<{
						relationshipPartPath?: string
						id?: string
						resolvedTarget?: string
					}>
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE })
		expect(result.structuredContent?.data?.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/workbook.xml',
				featureFamily: 'workbook',
				ownerScope: 'workbook',
				sourceRelationshipId: 'rId1',
			}),
		)
		expect(result.structuredContent?.data?.relationships).toContainEqual(
			expect.objectContaining({
				relationshipPartPath: '_rels/.rels',
				id: 'rId1',
				resolvedTarget: 'xl/workbook.xml',
			}),
		)
	})

	test('ascend.write recalculates before saving when needed', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: '=A1*2' },
		])
		await wb.recalc()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<unknown>
		await handler({
			file: TEMP_FILE,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
		})

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 10 })
	})

	test('ascend.write returns an error response when recalculation fails', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<{ isError?: boolean; structuredContent?: { error?: { message?: string } } }>
		const result = await handler({
			file: TEMP_FILE,
			ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }],
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.message).toContain('Circular reference detected')
	})

	test('ascend.calc returns an error response when recalculation fails', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.calc'].handler as (args: {
			file: string
		}) => Promise<{ isError?: boolean; structuredContent?: { error?: { message?: string } } }>
		const result = await handler({ file: TEMP_FILE })

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.message).toContain('Circular reference detected')
	})

	test('ascend.preview returns an error response when recalculation fails', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: { message?: string; details?: { preview?: { errors?: unknown[] } } }
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }],
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.ok).toBe(false)
		expect(result.structuredContent?.error?.message).toContain('Circular reference detected')
		expect(result.structuredContent?.error?.details?.preview?.errors?.length).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')).toBeUndefined()
	})

	test('ascend.export writes JSON/TSV and rejects unsupported formats', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.export'].handler as (args: {
			file: string
			output: string
			format?: string
		}) => Promise<{ isError?: boolean; structuredContent?: Record<string, unknown> }>

		const jsonPath = `${TEMP_FILE}.json`
		await handler({ file: TEMP_FILE, output: jsonPath, format: 'json' })
		expect(await Bun.file(jsonPath).text()).toContain('"sheets"')
		await unlink(jsonPath).catch(() => {})

		const tsvPath = `${TEMP_FILE}.tsv`
		await handler({ file: TEMP_FILE, output: tsvPath, format: 'tsv' })
		expect(await Bun.file(tsvPath).text()).toContain('Name\tScore')
		await unlink(tsvPath).catch(() => {})

		const bad = await handler({ file: TEMP_FILE, output: `${TEMP_FILE}.weird`, format: 'weird' })
		expect(bad.isError).toBe(true)
	})

	test('ascend.read can return TSV output', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'tsv'
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { format?: string; tsv?: string } }
		}>

		const result = await handler({ file: TEMP_FILE, range: 'A1:B2', format: 'tsv' })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.format).toBe('tsv')
		expect(result.structuredContent?.data?.tsv).toBe('Name\tScore\nAlice\t10')
	})

	test('ascend.read prunes TSV columns by letter and range position', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'C1', value: 'Notes' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
					{ ref: 'C2', value: 'ok' },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'tsv'
			cols?: string[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { format?: string; tsv?: string; colCount?: number; selectedColumns?: unknown[] }
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:C2',
			format: 'tsv',
			cols: ['A', '3'],
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.format).toBe('tsv')
		expect(result.structuredContent?.data?.colCount).toBe(2)
		expect(result.structuredContent?.data?.tsv).toBe('Name\tNotes\nAlice\tok')
		expect(result.structuredContent?.data?.selectedColumns).toEqual([
			{ position: 1, col: 0, letter: 'A' },
			{ position: 3, col: 2, letter: 'C' },
		])
	})

	test('ascend.read prunes object columns by header', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'C1', value: 'Notes' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
					{ ref: 'C2', value: 'ok' },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'objects'
			cols?: string[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { headers?: string[]; rows?: Array<Record<string, unknown>>; colCount?: number }
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:C2',
			format: 'objects',
			cols: ['Score'],
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.colCount).toBe(1)
		expect(result.structuredContent?.data?.headers).toEqual(['Score'])
		expect(result.structuredContent?.data?.rows).toEqual([{ Score: { kind: 'number', value: 10 } }])
	})

	test('ascend.read can return sparse compact output', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'C2', value: 10 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'compact'
			changedSince?: string
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { format?: string; cells?: unknown[]; changeToken?: string }
			}
		}>

		const result = await handler({ file: TEMP_FILE, range: 'A1:C2', format: 'compact' })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.format).toBe('compact')
		expect(result.structuredContent?.data?.cells).toEqual([
			[0, 0, 'Name'],
			[1, 2, 10],
		])
		expect(result.structuredContent?.data?.changeToken).toBeDefined()

		const unchanged = await handler({
			file: TEMP_FILE,
			range: 'A1:C2',
			format: 'compact',
			changedSince: result.structuredContent?.data?.changeToken,
		})
		expect(unchanged.structuredContent?.ok).toBe(true)
		expect(unchanged.structuredContent?.data?.cells).toEqual([])
		expect(unchanged.structuredContent?.data?.changeToken).toBeDefined()
	})

	test('ascend.read prunes compact cells by column', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'C2', value: 10 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'compact'
			cols?: string[]
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { cells?: unknown[]; colCount?: number } }
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:C2',
			format: 'compact',
			cols: ['C'],
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.colCount).toBe(1)
		expect(result.structuredContent?.data?.cells).toEqual([[1, 2, 10]])
	})

	test('ascend.read_table reads table rows by name', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
					{ ref: 'A3', value: 'Bob' },
					{ ref: 'B3', value: 20 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Scores', hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read_table'].handler as (args: {
			file: string
			table: string
			rowLimit?: number
			display?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { name?: string; columns?: string[]; rows?: Array<Record<string, string>> }
			}
		}>

		const result = await handler({ file: TEMP_FILE, table: 'Scores', rowLimit: 1, display: true })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.name).toBe('Scores')
		expect(result.structuredContent?.data?.columns).toEqual(['Name', 'Score'])
		expect(result.structuredContent?.data?.rows).toEqual([{ Name: 'Alice', Score: '10' }])
	})

	test('ascend.eval evaluates formulas without writing scratch cells', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 7 },
					{ ref: 'A2', value: 5 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.eval'].handler as (args: {
			file: string
			formula: string
			display?: boolean
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { value?: unknown; display?: string } }
		}>

		const result = await handler({ file: TEMP_FILE, formula: '=SUM(A1:A2)', display: true })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.value).toEqual({ kind: 'number', value: 12 })
		expect(result.structuredContent?.data?.display).toBe('12')
	})

	test('ascend.find can search values and formulas', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Alice' },
					{ ref: 'A2', value: 'Bob' },
					{ ref: 'B1', value: 10 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=SUM(B1:B1)' },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.find'].handler as (args: {
			file: string
			query: string
			in?: 'value' | 'formula' | 'both'
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { matches?: Array<{ ref: string; matchedOn: 'value' | 'formula' }> }
			}
		}>

		const valueResult = await handler({ file: TEMP_FILE, query: 'Alice', in: 'value' })
		expect(valueResult.structuredContent?.ok).toBe(true)
		expect(valueResult.structuredContent?.data?.matches?.[0]?.ref).toBe('A1')
		expect(valueResult.structuredContent?.data?.matches?.[0]?.matchedOn).toBe('value')

		const formulaResult = await handler({ file: TEMP_FILE, query: 'SUM', in: 'formula' })
		expect(formulaResult.structuredContent?.ok).toBe(true)
		expect(formulaResult.structuredContent?.data?.matches?.[0]?.ref).toBe('C1')
		expect(formulaResult.structuredContent?.data?.matches?.[0]?.matchedOn).toBe('formula')
	})

	test('ascend.write rejects invalid operations with validation error', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<{ isError?: boolean; structuredContent?: { error?: { code?: string } } }>

		const badOps = [
			{ op: 'setCells', sheet: 'Sheet1' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1' },
			{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: '2' },
			{ op: 'unknownOp', sheet: 'Sheet1' },
		]
		for (const ops of badOps) {
			const result = await handler({ file: TEMP_FILE, ops: [ops] })
			expect(result.isError).toBe(true)
			expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		}
	})

	test('MCP operation schema accepts capability extension operations', () => {
		const result = parseOperations([
			{
				op: 'setWorkbookProtection',
				protection: { lockStructure: true },
			},
			{ op: 'deleteTable', table: 'Sales' },
			{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
			{ op: 'resizeTable', table: 'Revenue', ref: 'A1:D20' },
			{
				op: 'replaceImage',
				sheet: 'Sheet1',
				targetPath: 'xl/media/image1.png',
				contentBase64: 'iVBORw0KGgo=',
				contentType: 'image/png',
			},
			{
				op: 'setPivotCache',
				pivotTable: 'PivotTable1',
				sourceSheet: 'RawData',
				sourceRef: 'A1:E200',
				refreshOnLoad: true,
				invalid: true,
			},
		])

		expect(result.ok).toBe(true)
	})

	test('ascend.preview rejects invalid operations with validation error', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<{ isError?: boolean; structuredContent?: { error?: { code?: string } } }>

		const result = await handler({
			file: TEMP_FILE,
			ops: [{ op: 'setCells', sheet: 'Sheet1' }],
		})
		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
	})

	test('sheet-not-found MCP errors include suggested fixes', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			sheet: string
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: { code?: string; suggestedFix?: string; details?: { availableSheets?: string[] } }
			}
		}>

		const result = await handler({ file: TEMP_FILE, range: 'A1:A1', sheet: 'Missing' })
		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('SHEET_NOT_FOUND')
		expect(result.structuredContent?.error?.suggestedFix).toContain('Sheet1')
		expect(result.structuredContent?.error?.details?.availableSheets).toContain('Sheet1')
	})

	test('ascend.check returns result on clean workbook', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.check'].handler as (args: {
			file: string
		}) => Promise<{ content?: Array<{ text?: string }> }>

		const result = await handler({ file: TEMP_FILE })
		const allText = (result.content ?? []).map((c) => c.text ?? '').join('\n')
		expect(allText.length).toBeGreaterThan(0)
	})

	test('ascend.check preserves structured issue metadata for agent repair', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'renameSheet', sheet: 'Sheet1', newName: 'SummaryData' }])
		wb.apply([{ op: 'setFormula', sheet: 'SummaryData', ref: 'A1', formula: '=Summary!B1' }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.check'].handler as (args: {
			file: string
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					details?: {
						check?: {
							issues?: Array<{
								rule?: string
								ref?: string
								refs?: string[]
								suggestedFix?: string
							}>
						}
					}
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE })
		expect(result.isError).toBe(true)
		const issue = result.structuredContent?.error?.details?.check?.issues?.find(
			(entry) => entry.rule === 'broken-refs',
		)
		expect(issue?.ref).toBe('SummaryData!A1')
		expect(issue?.refs).toEqual(['SummaryData!A1'])
		expect(issue?.suggestedFix).toContain('SummaryData')
	})

	test('ascend.lint returns result', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.lint'].handler as (args: {
			file: string
		}) => Promise<{ content?: Array<{ text?: string }> }>

		const result = await handler({ file: TEMP_FILE })
		expect(result.content).toBeDefined()
	})

	test('ascend.list_operations returns operations', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.list_operations'].handler as (
			args: Record<string, never>,
		) => Promise<{ content?: Array<{ text?: string }> }>

		const result = await handler({})
		const allText = (result.content ?? []).map((c) => c.text ?? '').join('\n')
		expect(allText).toContain('operations')
	})

	test('ascend.inspect returns workbook info', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'TestSheet' }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.inspect'].handler as (args: {
			file: string
		}) => Promise<{
			content?: Array<{ text?: string }>
			structuredContent?: { data?: { sheets?: Array<{ name: string }> } }
		}>

		const result = await handler({ file: TEMP_FILE })
		const names = (result.structuredContent?.data?.sheets ?? []).map((s) => s.name)
		expect(names).toContain('TestSheet')
	})

	test('sheet-not-found errors tell agents to modify the request', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Data' }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.inspect'].handler as (args: {
			file: string
			sheet?: string
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					code?: string
					retryStrategy?: string
					details?: { availableSheets?: string[] }
					suggestedFix?: string
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE, sheet: 'Missing' })
		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('SHEET_NOT_FOUND')
		expect(result.structuredContent?.error?.retryStrategy).toBe('modified')
		expect(result.structuredContent?.error?.details?.availableSheets).toContain('Data')
		expect(result.structuredContent?.error?.suggestedFix).toContain('Data')
	})

	test('ascend.list_sheets returns sheet names', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Alpha' }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.list_sheets'].handler as (args: {
			file: string
		}) => Promise<{
			content?: Array<{ text?: string }>
			structuredContent?: { data?: { sheets?: Array<{ name: string }> } }
		}>

		const result = await handler({ file: TEMP_FILE })
		const names = (result.structuredContent?.data?.sheets ?? []).map((s) => s.name)
		expect(names).toContain('Alpha')
	})

	test('ascend.visuals returns visual inventory', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.visuals'].handler as (args: {
			file: string
		}) => Promise<{
			content?: Array<{ text?: string }>
			structuredContent?: {
				data?: {
					load?: { mode?: string }
					sheetImageCount?: number
					sheets?: Array<{ sheet: string; imageCount: number }>
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE })
		expect(result.structuredContent?.data?.load?.mode).toBe('full')
		expect(result.structuredContent?.data?.sheetImageCount).toBe(0)
		expect(result.structuredContent?.data?.sheets?.[0]).toMatchObject({
			sheet: 'Sheet1',
			imageCount: 0,
		})
	})

	test('ascend.pivots returns audits and materialization operations', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.pivots'].handler as (args: {
			file: string
			pivotTable?: string
		}) => Promise<{
			structuredContent?: {
				data?: {
					pivotTables?: Array<{ name?: string }>
					pivotOutputAudits?: Array<{ pivotTable?: string; status?: string }>
					pivotOutputMaterializePlan?: {
						ops: unknown[]
						plannedCellCount: number
						unsupported: unknown[]
					}
				}
			}
		}>

		const result = await handler({ file: PIVOT_FIXTURE, pivotTable: 'PivotTable1' })
		expect(result.structuredContent?.data?.pivotTables?.[0]?.name).toBe('PivotTable1')
		expect(result.structuredContent?.data?.pivotOutputAudits?.[0]).toMatchObject({
			pivotTable: 'PivotTable1',
			status: 'passed',
		})
		expect(result.structuredContent?.data?.pivotOutputMaterializePlan).toEqual({
			ops: [],
			plannedCellCount: 0,
			unsupported: [],
		})
	})

	test('ascend.commit requires approval for destructive operations', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Scratch' }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			file: string
			ops: unknown[]
			inPlace?: boolean
			approvals?: string[] | string
		}) => Promise<{
			isError?: boolean
			structuredContent?: { data?: { approvals?: Array<{ id: string }> } }
		}>
		const ops = [{ op: 'deleteSheet', sheet: 'Scratch' }]

		const blocked = await handler({ file: TEMP_FILE, ops, inPlace: true })
		expect(blocked.isError).toBe(true)

		const committed = await handler({
			file: TEMP_FILE,
			ops,
			inPlace: true,
			approvals: ['op:0:deletesheet'],
		})
		expect(committed.isError).not.toBe(true)
		expect(committed.structuredContent?.data?.approvals?.[0]?.id).toBe('op:0:deletesheet')
	})

	test('file-not-found returns structured error', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.inspect'].handler as (args: {
			file: string
		}) => Promise<{ isError?: boolean }>

		const result = await handler({ file: '/nonexistent/path/to/file.xlsx' })
		expect(result.isError).toBe(true)
	})
})

function signedMacroWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`,
		'xl/_rels/vbaProject.bin.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/vbaProject.bin': 'macro-bytes',
		'xl/vbaProjectSignature.bin': 'signature-bytes',
	})
}
