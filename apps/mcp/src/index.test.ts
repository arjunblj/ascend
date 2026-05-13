import { afterAll, describe, expect, test } from 'bun:test'
import { Buffer } from 'node:buffer'
import { createHash } from 'node:crypto'
import { unlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook, parseOperations } from '@ascend/sdk'
import { createZip, encode } from '../../../packages/io-xlsx/src/writer/zip.ts'
import { makeXlsx } from '../../../packages/io-xlsx/test/helpers.ts'
import { createServer } from './index.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-mcp-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const TEMP_MACRO_FILE = join(
	tmpdir(),
	`ascend-mcp-macro-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)
const TEMP_MACRO_OUTPUT = join(
	tmpdir(),
	`ascend-mcp-macro-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)
const PIVOT_FIXTURE = join(
	import.meta.dir,
	'../../../fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
)

afterAll(async () => {
	await unlink(TEMP_FILE).catch(() => {})
	await unlink(TEMP_MACRO_FILE).catch(() => {})
	await unlink(TEMP_MACRO_OUTPUT).catch(() => {})
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
		expect(names).toContain('ascend.dump')
		expect(names).toContain('ascend.template_merge')
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
		expect(names).toContain('ascend.raw_part')
		expect(names).toContain('ascend.visuals')
		expect(names).toContain('ascend.pivots')
		expect(names).toContain('ascend.capabilities')
		expect(names).toContain('ascend.plan')
		expect(names).toContain('ascend.commit')
		expect(names).toContain('ascend.repair_plan')
		expect(names.length).toBe(29)
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
					locations: ['xl/vbaProject.bin'],
				}),
			)
			expect(data?.compatibilityFeatures).toContainEqual(
				expect.objectContaining({
					feature: 'preservedSignature',
					locations: ['xl/vbaProjectSignature.bin'],
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

	test('ascend.raw_part exposes bounded package text and metadata', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.raw_part'].handler as (args: {
			file: string
			partPath: string
			encoding?: 'text' | 'base64' | 'none'
			maxBytes?: number
			caseInsensitive?: boolean
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				data?: {
					partPath?: string
					featureFamily?: string
					text?: string
					base64?: string
					origin?: string
					semantics?: string
					encoding?: string
					previewByteLength?: number
					truncated?: boolean
					sha256?: string
					caseInsensitiveFallback?: boolean
					load?: { mode?: string; isPartial?: boolean }
				}
				error?: { code?: string; details?: { validPath?: boolean } }
			}
		}>

		const result = await handler({ file: TEMP_FILE, partPath: 'xl/workbook.xml', maxBytes: 64 })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.partPath).toBe('xl/workbook.xml')
		expect(result.structuredContent?.data?.origin).toBe('source')
		expect(result.structuredContent?.data?.load?.mode).toBe('metadata-only')
		expect(result.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(result.structuredContent?.data?.semantics).toBe('raw-package-bytes')
		expect(result.structuredContent?.data?.featureFamily).toBe('workbook')
		expect(result.structuredContent?.data?.text).toContain('<?xml')
		expect(result.structuredContent?.data?.previewByteLength).toBe(64)
		expect(result.structuredContent?.data?.truncated).toBe(true)
		expect(result.structuredContent?.data?.sha256).toMatch(/^[a-f0-9]{64}$/)

		const base64 = await handler({
			file: TEMP_FILE,
			partPath: '/xl/workbook.xml',
			encoding: 'base64',
			maxBytes: 12,
		})
		expect(base64.structuredContent?.ok).toBe(true)
		expect(base64.structuredContent?.data?.encoding).toBe('base64')
		expect(base64.structuredContent?.data?.base64).toBeDefined()
		expect(base64.structuredContent?.data?.text).toBeUndefined()

		const fallback = await handler({
			file: TEMP_FILE,
			partPath: 'XL/WORKBOOK.XML',
			caseInsensitive: true,
			maxBytes: 8,
		})
		expect(fallback.structuredContent?.ok).toBe(true)
		expect(fallback.structuredContent?.data?.partPath).toBe('xl/workbook.xml')
		expect(fallback.structuredContent?.data?.caseInsensitiveFallback).toBe(true)

		const missing = await handler({ file: TEMP_FILE, partPath: 'xl/missing.xml' })
		expect(missing.isError).toBe(true)
		expect(missing.structuredContent?.error?.code).toBe('FILE_NOT_FOUND')

		const invalid = await handler({ file: TEMP_FILE, partPath: 'xl//workbook.xml' })
		expect(invalid.isError).toBe(true)
		expect(invalid.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(invalid.structuredContent?.error?.details?.validPath).toBe(false)
	})

	test('ascend.raw_part returns binary base64 previews with full-byte metadata', async () => {
		const binaryBytes = new Uint8Array([0, 1, 2, 3, 4, 255])
		const binaryFile = `${TEMP_FILE}.raw-binary.xlsx`
		await writeFile(binaryFile, binaryRawPartWorkbook(binaryBytes))
		try {
			const server = createServer()
			// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
			const handler = (server as any)._registeredTools['ascend.raw_part'].handler as (args: {
				file: string
				partPath: string
				encoding?: 'text' | 'base64' | 'none'
				maxBytes?: number
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						encoding?: string
						base64?: string
						text?: string
						previewByteLength?: number
						truncated?: boolean
						sha256?: string
					}
				}
			}>

			const result = await handler({
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'base64',
				maxBytes: 3,
			})
			expect(result.structuredContent?.ok).toBe(true)
			expect(result.structuredContent?.data?.encoding).toBe('base64')
			expect(result.structuredContent?.data?.base64).toBe(
				Buffer.from(binaryBytes.subarray(0, 3)).toString('base64'),
			)
			expect(result.structuredContent?.data?.text).toBeUndefined()
			expect(result.structuredContent?.data?.previewByteLength).toBe(3)
			expect(result.structuredContent?.data?.truncated).toBe(true)
			expect(result.structuredContent?.data?.sha256).toBe(
				createHash('sha256').update(binaryBytes).digest('hex'),
			)

			const metadataOnly = await handler({
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'none',
				maxBytes: 3,
			})
			expect(metadataOnly.structuredContent?.ok).toBe(true)
			expect(metadataOnly.structuredContent?.data?.encoding).toBe('none')
			expect(metadataOnly.structuredContent?.data?.base64).toBeUndefined()
			expect(metadataOnly.structuredContent?.data?.text).toBeUndefined()
			expect(metadataOnly.structuredContent?.data?.previewByteLength).toBe(0)
			expect(metadataOnly.structuredContent?.data?.truncated).toBe(false)
			expect(metadataOnly.structuredContent?.data?.sha256).toBe(
				result.structuredContent?.data?.sha256,
			)
		} finally {
			await unlink(binaryFile).catch(() => {})
		}
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
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { journal?: { supported?: boolean; inverseOps?: unknown[] } }
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])

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

	test('ascend.preview can include mutation journal metadata', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { journal?: { supported?: boolean; inverseOps?: unknown[] } }
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'new' }] }],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] },
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'old' })
	})

	test('ascend.preview and ascend.write accept path-addressed mutations', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const preview = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					pathMutations?: { replayable?: boolean; ops?: unknown[] }
					journal?: { supported?: boolean; inverseOps?: unknown[] }
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { pathMutations?: { ops?: unknown[] } } }
		}>
		const ambiguousPreview = preview as (args: {
			file: string
			ops: unknown[]
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			isError?: boolean
			structuredContent?: { error?: { message?: string } }
		}>

		const previewResult = await preview({
			file: TEMP_FILE,
			journal: true,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' }],
		})
		expect(previewResult.structuredContent?.ok).toBe(true)
		expect(previewResult.structuredContent?.data?.pathMutations?.replayable).toBe(true)
		expect(previewResult.structuredContent?.data?.pathMutations?.ops).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'new' }] },
		])
		expect(previewResult.structuredContent?.data?.journal?.inverseOps).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] },
		])

		const ambiguous = await ambiguousPreview({
			file: TEMP_FILE,
			ops: [],
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' }],
		})
		expect(ambiguous.isError).toBe(true)
		expect(ambiguous.structuredContent?.error?.message).toBe(
			'Provide either ops or mutations, not both',
		)

		const writeResult = await write({
			file: TEMP_FILE,
			mutations: [
				{ path: '/sheets/Sheet1/cells/B1/formula', value: 'A1&"-ok"' },
				{ path: '/sheets/Sheet1/cells/C1/comment', value: { text: 'review', author: 'agent' } },
			],
		})
		expect(writeResult.structuredContent?.ok).toBe(true)
		expect(writeResult.structuredContent?.data?.pathMutations?.ops).toEqual([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1&"-ok"' },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'C1', text: 'review', author: 'agent' },
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'old' })
		expect(reopened.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1&"-ok"')
		expect(reopened.sheet('Sheet1')?.comment('C1')?.text).toBe('review')
	})

	test('ascend.preview, plan, write, and commit keep escaped path mutations canonical', async () => {
		const sheetName = "Q1.Forecast's Café Δ"
		const tableName = 'Sales.Δ'
		const columnName = 'Gross Profit/Δ~'
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: sheetName },
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: columnName },
					{ ref: 'A2', value: 'North' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: sheetName, ref: 'A1:B2', name: tableName, hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const preview = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			mutations: Array<{ path: string | string[]; value?: unknown }>
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { pathMutations?: { replayable?: boolean; ops?: unknown[] } }
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			mutations: Array<{ path: string | string[]; value?: unknown }>
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					approvals?: Array<{ id: string }>
					pathMutations?: { ops?: unknown[] }
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			mutations: Array<{ path: string | string[]; value?: unknown }>
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { pathMutations?: { ops?: unknown[] } }
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			file: string
			output: string
			mutations: Array<{ path: string | string[]; value?: unknown }>
			approvals?: string[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { pathMutations?: { ops?: unknown[] } }
			}
		}>

		const mutations = [
			{ path: `/sheets/${pointerSegment(sheetName)}/cells/A2/value`, value: 'pointer' },
			{ path: `sheets.${dotSegment(sheetName)}.cells.A3.value`, value: 'dot' },
			{ path: ['sheets', sheetName, 'cells', 'A4', 'value'], value: 'array' },
			{
				path: `tables.${dotSegment(tableName)}.columns.${dotSegment(columnName)}.formula`,
				value: 'SUM([Region])',
			},
			{ path: ['tables', tableName, 'columns', columnName, 'name'], value: 'Net_Δ' },
		]
		const canonicalOps = [
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A2', value: 'pointer' },
					{ ref: 'A3', value: 'dot' },
					{ ref: 'A4', value: 'array' },
				],
			},
			{
				op: 'setTableColumn',
				table: tableName,
				column: columnName,
				formula: 'SUM([Region])',
			},
			{ op: 'setTableColumn', table: tableName, column: columnName, newName: 'Net_Δ' },
		]

		const result = await preview({ file: TEMP_FILE, mutations })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.pathMutations?.replayable).toBe(true)
		expect(result.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)

		const planned = await plan({ file: TEMP_FILE, mutations })
		expect(planned.structuredContent?.ok).toBe(true)
		expect(planned.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
		const approvalIds =
			planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

		const writePath = `${TEMP_FILE}.escaped-write.xlsx`
		const commitInput = `${TEMP_FILE}.escaped-commit-input.xlsx`
		const commitOutput = `${TEMP_MACRO_OUTPUT}.escaped-commit-output.xlsx`
		try {
			await wb.save(writePath)
			const written = await write({ file: writePath, mutations })
			expect(written.structuredContent?.ok).toBe(true)
			expect(written.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const writeReopened = await AscendWorkbook.open(writePath)
			expect(writeReopened.sheet(sheetName)?.cell('A3')?.value).toEqual({
				kind: 'string',
				value: 'dot',
			})
			expect(writeReopened.sheet(sheetName)?.cell('A4')?.value).toEqual({
				kind: 'string',
				value: 'array',
			})

			await wb.save(commitInput)
			const committed = await commit({
				file: commitInput,
				output: commitOutput,
				mutations,
				approvals: approvalIds,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const commitReopened = await AscendWorkbook.open(commitOutput)
			expect(commitReopened.sheet(sheetName)?.cell('A3')?.value).toEqual({
				kind: 'string',
				value: 'dot',
			})
			expect(commitReopened.sheet(sheetName)?.cell('A4')?.value).toEqual({
				kind: 'string',
				value: 'array',
			})
		} finally {
			await unlink(writePath).catch(() => {})
			await unlink(commitInput).catch(() => {})
			await unlink(commitOutput).catch(() => {})
		}
	})

	test('ascend.preview defers path mutation renames after dependent edits', async () => {
		const sheetName = 'Q1.Forecast'
		const tableName = 'Sales.Δ'
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: sheetName },
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Revenue' },
					{ ref: 'A2', value: 'North' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: sheetName, ref: 'A1:B2', name: tableName, hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const preview = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { pathMutations?: { replayable?: boolean; ops?: unknown[] } }
			}
		}>

		const result = await preview({
			file: TEMP_FILE,
			mutations: [
				{ path: `/sheets/${pointerSegment(sheetName)}/name`, value: 'Summary' },
				{ path: `/sheets/${pointerSegment(sheetName)}/cells/C1/value`, value: 'safe order' },
				{ path: `/tables/${pointerSegment(tableName)}/name`, value: 'SalesData' },
				{
					path: `/tables/${pointerSegment(tableName)}/columns/Revenue/formula`,
					value: 'SUM([Revenue])',
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.pathMutations?.replayable).toBe(true)
		expect(result.structuredContent?.data?.pathMutations?.ops).toEqual([
			{ op: 'setCells', sheet: sheetName, updates: [{ ref: 'C1', value: 'safe order' }] },
			{
				op: 'setTableColumn',
				table: tableName,
				column: 'Revenue',
				formula: 'SUM([Revenue])',
			},
			{ op: 'renameSheet', sheet: sheetName, newName: 'Summary' },
			{ op: 'renameTable', table: tableName, newName: 'SalesData' },
		])
	})

	test('ascend.plan reports path mutation compiler errors as structured repair details', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					details?: {
						issueCount?: number
						supportedPathShapes?: readonly string[]
						issueDetails?: readonly { code?: string; path?: string }[]
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			mutations: [{ path: '/sheets/Missing/cells/A1/value', value: 1 }],
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.ok).toBe(false)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(result.structuredContent?.error?.details?.issueCount).toBe(1)
		expect(result.structuredContent?.error?.details?.issueDetails).toEqual([
			expect.objectContaining({
				code: 'sheet_not_found',
				path: '/sheets/Missing/cells/A1/value',
			}),
		])
		expect(result.structuredContent?.error?.details?.supportedPathShapes).toEqual(
			expect.arrayContaining([
				'/sheets/{sheet}/ranges/{A1:B2}/conditionalFormat',
				'/tables/{table}/columns/{nameOrIndex}/totalsRowLabel',
				'/sheets/{sheet}/names/{name}/ref',
			]),
		)
	})

	test('ascend.plan reports malformed path syntax as structured repair details', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					details?: {
						issueCount?: number
						issues?: readonly string[]
						issueDetails?: readonly { code?: string; path?: string }[]
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			mutations: [
				{ path: '/sheets//cells/A1/value', value: 1 },
				{ path: '/sheets/%E0%A4%A/cells/A1/value', value: 1 },
				{ path: '/sheets/Sheet1~2/cells/A1/value', value: 1 },
			],
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.ok).toBe(false)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(result.structuredContent?.error?.details?.issueCount).toBe(3)
		expect(result.structuredContent?.error?.details?.issues).toEqual([
			'Path segment 1 must not be empty.',
			'Invalid percent encoding in path segment "%E0%A4%A".',
			'Invalid JSON Pointer escape in path segment "Sheet1~2".',
		])
		expect(result.structuredContent?.error?.details?.issueDetails).toEqual([
			expect.objectContaining({ code: 'invalid_path', path: '/sheets//cells/A1/value' }),
			expect.objectContaining({ code: 'invalid_path', path: '/sheets/%E0%A4%A/cells/A1/value' }),
			expect.objectContaining({ code: 'invalid_path', path: '/sheets/Sheet1~2/cells/A1/value' }),
		])
	})

	test('ascend.plan can return compact bounded preview details', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
					{ ref: 'A3', value: 3 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			compact?: boolean
			maxChangedCells?: number
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					preview?: {
						changedCellCount?: number
						emittedChangedCellCount?: number
						changedCells?: unknown[]
						wouldSucceed?: boolean
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			compact: true,
			maxChangedCells: 1,
			ops: [
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 10 },
						{ ref: 'A2', value: 20 },
						{ ref: 'A3', value: 30 },
					],
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.preview?.wouldSucceed).toBe(true)
		expect(result.structuredContent?.data?.preview?.changedCellCount).toBe(3)
		expect(result.structuredContent?.data?.preview?.emittedChangedCellCount).toBe(1)
		expect(result.structuredContent?.data?.preview?.changedCells).toHaveLength(1)
	})

	test('ascend.commit accepts prepared plan handles', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${TEMP_FILE}.prepared-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				prepare?: boolean
				mutations: Array<{ path: string; value?: unknown }>
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string; expiresAt?: string; ttlMs?: number }
						pathMutations?: { ops?: unknown[] }
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: string[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: { pathMutations?: { ops?: unknown[] } }
				}
			}>
			const planned = await plan({
				file: TEMP_FILE,
				prepare: true,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 321 }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(planned.structuredContent?.data?.preparedPlan?.expiresAt).toBeString()
			expect(planned.structuredContent?.data?.preparedPlan?.ttlMs).toBeNumber()
			expect(planned.structuredContent?.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 321 }] },
			])

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: [],
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 321 }] },
			])
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 321 })

			const reused = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: `${output}.reuse.xlsx`,
				approvals: [],
			})
			expect(reused.structuredContent?.ok).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
			await unlink(`${output}.reuse.xlsx`).catch(() => {})
		}
	})

	test('prepared MCP plan handles are bounded', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${TEMP_FILE}.bounded-prepared-mcp.xlsx`
		const server = createServer({ preparedPlanMaxHandles: 1 })
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				prepare?: boolean
				mutations: Array<{ path: string; value?: unknown }>
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: { preparedPlan?: { id?: string; ttlMs?: number } }
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: string[]
			}) => Promise<{ structuredContent?: { ok?: boolean } }>

			const first = await plan({
				file: TEMP_FILE,
				prepare: true,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
			})
			const second = await plan({
				file: TEMP_FILE,
				prepare: true,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 2 }],
			})
			expect(first.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(second.structuredContent?.data?.preparedPlan?.id).toBeString()

			const evicted = await commit({
				planHandle: first.structuredContent?.data?.preparedPlan?.id,
				output: `${output}.evicted.xlsx`,
				approvals: [],
			})
			expect(evicted.structuredContent?.ok).toBe(false)

			const committed = await commit({
				planHandle: second.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: [],
			})
			expect(committed.structuredContent?.ok).toBe(true)
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 2 })
		} finally {
			await unlink(output).catch(() => {})
			await unlink(`${output}.evicted.xlsx`).catch(() => {})
		}
	})

	test('malformed path mutations block MCP preview, write, and commit consistently', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const tools = (server as any)._registeredTools as Record<
			string,
			{
				handler: (args: {
					file: string
					output?: string
					mutations: Array<{ path: string; value?: unknown }>
				}) => Promise<{
					isError?: boolean
					structuredContent?: {
						ok?: boolean
						error?: {
							code?: string
							details?: {
								issueCount?: number
								issues?: readonly string[]
								issueDetails?: readonly { code?: string; path?: string }[]
							}
						}
					}
				}>
			}
		>

		for (const toolName of ['ascend.preview', 'ascend.write', 'ascend.commit'] as const) {
			const result = await tools[toolName]?.handler({
				file: TEMP_FILE,
				output: `${TEMP_FILE}.out.xlsx`,
				mutations: [{ path: '/sheets//cells/A1/value', value: 'new' }],
			})

			expect(result?.isError).toBe(true)
			expect(result?.structuredContent?.ok).toBe(false)
			expect(result?.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(result?.structuredContent?.error?.details?.issueCount).toBe(1)
			expect(result?.structuredContent?.error?.details?.issues).toEqual([
				'Path segment 1 must not be empty.',
			])
			expect(result?.structuredContent?.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'invalid_path',
					path: '/sheets//cells/A1/value',
				}),
			])
		}

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})
	})

	test('non-replayable path mutation batches do not expose or apply partial MCP ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const tools = (server as any)._registeredTools as Record<
			string,
			{
				handler: (args: {
					file: string
					output?: string
					mutations: Array<{ path: string; value?: unknown }>
				}) => Promise<{
					isError?: boolean
					structuredContent?: {
						ok?: boolean
						error?: {
							code?: string
							details?: {
								issueCount?: number
								compiledOps?: readonly unknown[]
								issueDetails?: readonly { code?: string; path?: string }[]
							}
						}
					}
				}>
			}
		>

		for (const toolName of [
			'ascend.plan',
			'ascend.preview',
			'ascend.write',
			'ascend.commit',
		] as const) {
			const result = await tools[toolName]?.handler({
				file: TEMP_FILE,
				output: `${TEMP_FILE}.out.xlsx`,
				mutations: [
					{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' },
					{ path: '/sheets/Missing/cells/A1/value', value: 1 },
				],
			})

			expect(result?.isError).toBe(true)
			expect(result?.structuredContent?.ok).toBe(false)
			expect(result?.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(result?.structuredContent?.error?.details?.issueCount).toBe(1)
			expect(result?.structuredContent?.error?.details?.compiledOps).toEqual([])
			expect(result?.structuredContent?.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'sheet_not_found',
					path: '/sheets/Missing/cells/A1/value',
				}),
			])
		}

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})
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

	test('ascend.read returns compact first-window data with partial load metadata', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 20 }, (_, row) => [
					{ ref: `A${row + 1}`, value: row + 1 },
					{ ref: `B${row + 1}`, value: `row-${row + 1}` },
				]).flat(),
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'compact'
			rowLimit?: number
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					cells?: unknown[]
					rowCount?: number
					load?: {
						mode?: string
						isPartial?: boolean
						maxRows?: number
						partialReasons?: readonly string[]
						cellsHydrated?: boolean
						loadedSheets?: readonly string[]
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:B20',
			format: 'compact',
			rowLimit: 3,
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.rowCount).toBe(3)
		expect(result.structuredContent?.data?.cells).toHaveLength(6)
		expect(result.structuredContent?.data?.load?.mode).toBe('values')
		expect(result.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(result.structuredContent?.data?.load?.maxRows).toBe(3)
		expect(result.structuredContent?.data?.load?.partialReasons).toContain(
			'only the first 3 row(s) are hydrated per loaded sheet',
		)
		expect(result.structuredContent?.data?.load?.cellsHydrated).toBe(true)
		expect(result.structuredContent?.data?.load?.loadedSheets).toEqual(['Sheet1'])
	})

	test('ascend.trace returns structured partial-load diagnostics for capped formula views', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.trace'].handler as (args: {
			file: string
			cell: string
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					details?: {
						rule?: string
						load?: {
							maxRows?: number
							partialReasons?: readonly string[]
						}
					}
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE, cell: 'Sheet1!A1', maxRows: 1 })

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.ok).toBe(false)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(result.structuredContent?.error?.details?.rule).toBe('partial-dependency-analysis')
		expect(result.structuredContent?.error?.details?.load?.maxRows).toBe(1)
		expect(result.structuredContent?.error?.details?.load?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
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

	test('ascend.dump emits replayable operation batches', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'B1', value: 'label' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: 'A1*2' },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.dump'].handler as (args: {
			file: string
			sheet?: string
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					replayable?: boolean
					ops?: unknown[]
					formulaCount?: number
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE, sheet: 'Sheet1' })
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.replayable).toBe(true)
		expect(result.structuredContent?.data?.formulaCount).toBe(1)
		expect(result.structuredContent?.data?.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'B1', value: 'label' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: 'A1*2' },
		])
	})

	test('ascend.template_merge emits replayable operation batches and unresolved placeholders', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: '{{amount}}' },
					{ ref: 'A2', value: 'Missing {{client}}' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+{{tax}}' },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.template_merge'].handler as (args: {
			file: string
			sheet?: string
			data: Record<string, string | number | boolean | null>
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					replayable?: boolean
					ops?: unknown[]
					unresolved?: unknown[]
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			sheet: 'Sheet1',
			data: { amount: 10, tax: 2 },
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.replayable).toBe(false)
		expect(result.structuredContent?.data?.unresolved).toEqual([
			{
				sheet: 'Sheet1',
				ref: 'A2',
				source: 'value',
				placeholder: '{{client}}',
				key: 'client',
			},
		])
		expect(result.structuredContent?.data?.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 10 }],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+2' },
		])
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

	test('ascend.plan invalid ops return structured batch repair details', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					code?: string
					details?: {
						issueCount?: number
						issues?: readonly string[]
						issueDetails?: readonly {
							code?: string
							opIndex?: number
							path?: string
						}[]
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			ops: [
				{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: '2' },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: { nested: true } }] },
				{ op: 'missingOp', sheet: 'Sheet1' },
			],
		})
		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(result.structuredContent?.error?.details?.issueCount).toBe(3)
		expect(result.structuredContent?.error?.details?.issues).toEqual(
			expect.arrayContaining([
				'ops[0].count must be a positive integer',
				'ops[1].updates[0].value must be a scalar value or null',
				'ops[2].op "missingOp" is not supported',
			]),
		)
		expect(result.structuredContent?.error?.details?.issueDetails).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ code: 'invalid_type', opIndex: 0, path: 'ops[0].count' }),
				expect.objectContaining({ code: 'invalid_type', opIndex: 1 }),
				expect.objectContaining({ code: 'invalid_operation', opIndex: 2, path: 'ops[2].op' }),
			]),
		)
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
		await Bun.write(TEMP_FILE, threadedCommentMissingPersonsWorkbook())

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
								details?: { kind?: string }
							}>
						}
					}
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE })
		expect(result.isError).toBe(true)
		const issue = result.structuredContent?.error?.details?.check?.issues?.find(
			(entry) => entry.rule === 'threaded-comment-integrity',
		)
		expect(issue?.refs).toEqual(['Sheet1!A1'])
		expect(issue?.details?.kind).toBe('threaded-comment-unknown-person-id')
		expect(issue?.suggestedFix).toContain('persons part')
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

		const aliasBlocked = await handler({
			file: TEMP_FILE,
			ops,
			inPlace: true,
			approvals: ['deleteSheet'],
		})
		expect(aliasBlocked.isError).toBe(true)

		const committed = await handler({
			file: TEMP_FILE,
			ops,
			inPlace: true,
			approvals: ['op:0:deletesheet'],
		})
		expect(committed.isError).not.toBe(true)
		expect(committed.structuredContent?.data?.approvals?.[0]?.id).toBe('op:0:deletesheet')
	})

	test('ascend.commit requires exact approval ids for preserved lossy features', async () => {
		await Bun.write(TEMP_MACRO_FILE, signedMacroWorkbook())
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const planHandler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops: unknown[]
		}) => Promise<{ structuredContent?: { data?: { approvals?: Array<{ id: string }> } } }>
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const commitHandler = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			file: string
			ops: unknown[]
			output?: string
			approvals?: string[] | string
		}) => Promise<{
			isError?: boolean
			structuredContent?: { data?: { approvals?: Array<{ id: string }> } }
		}>
		const ops = [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 7 }] }]
		const planned = await planHandler({ file: TEMP_MACRO_FILE, ops })
		const approvalIds = planned.structuredContent?.data?.approvals?.map((approval) => approval.id)
		expect(approvalIds).toEqual(
			expect.arrayContaining([
				expect.stringMatching(/^loss:preservedmacro:preserved:/),
				expect.stringMatching(/^loss:preservedsignature:preserved:/),
			]),
		)

		const aliasBlocked = await commitHandler({
			file: TEMP_MACRO_FILE,
			ops,
			output: TEMP_MACRO_OUTPUT,
			approvals: ['preservedMacro', 'preservedSignature'],
		})
		expect(aliasBlocked.isError).toBe(true)

		const committed = await commitHandler({
			file: TEMP_MACRO_FILE,
			ops,
			output: TEMP_MACRO_OUTPUT,
			approvals: approvalIds ?? [],
		})
		expect(committed.isError).not.toBe(true)
		expect(committed.structuredContent?.data?.approvals?.map((approval) => approval.id)).toEqual(
			approvalIds,
		)
	})

	test('ascend.commit preserves path mutation canonical ops and exact approval ids', async () => {
		await Bun.write(TEMP_MACRO_FILE, signedMacroWorkbook())
		const output = `${TEMP_MACRO_OUTPUT}.path.xlsm`
		await unlink(output).catch(() => {})
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const planHandler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			structuredContent?: {
				data?: {
					approvals?: Array<{ id: string }>
					pathMutations?: { replayable?: boolean; ops?: unknown[] }
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const commitHandler = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
			output?: string
			approvals?: string[] | string
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				data?: {
					approvals?: Array<{ id: string }>
					pathMutations?: { ops?: unknown[] }
				}
				error?: { message?: string }
			}
		}>
		const mutations = [{ path: '/sheets/Data/cells/A1/value', value: 11 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }]

		const planned = await planHandler({ file: TEMP_MACRO_FILE, mutations })
		expect(planned.structuredContent?.data?.pathMutations?.replayable).toBe(true)
		expect(planned.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
		const approvalIds = planned.structuredContent?.data?.approvals?.map((approval) => approval.id)
		expect(approvalIds).toEqual(
			expect.arrayContaining([
				expect.stringMatching(/^loss:preservedmacro:preserved:/),
				expect.stringMatching(/^loss:preservedsignature:preserved:/),
			]),
		)

		const aliasBlocked = await commitHandler({
			file: TEMP_MACRO_FILE,
			mutations,
			output,
			approvals: ['preservedMacro', 'preservedSignature'],
		})
		expect(aliasBlocked.isError).toBe(true)
		expect(aliasBlocked.structuredContent?.error?.message).toBe('Commit requires explicit approval')

		const committed = await commitHandler({
			file: TEMP_MACRO_FILE,
			mutations,
			output,
			approvals: approvalIds ?? [],
		})
		expect(committed.isError).not.toBe(true)
		expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
		expect(committed.structuredContent?.data?.approvals?.map((approval) => approval.id)).toEqual(
			approvalIds,
		)
		const reopened = await AscendWorkbook.open(output)
		expect(reopened.sheet('Data')?.cell('A1')?.value).toEqual({ kind: 'number', value: 11 })
		await unlink(output).catch(() => {})
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

function pointerSegment(value: string): string {
	return encodeURIComponent(value.replace(/~/g, '~0').replace(/\//g, '~1'))
}

function dotSegment(value: string): string {
	return value.replace(/\\/g, '\\\\').replace(/\./g, '\\.')
}

function binaryRawPartWorkbook(binaryBytes: Uint8Array): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/media/image1.png': binaryBytes,
			}),
		),
	)
}

function threadedCommentMissingPersonsWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`,
		'xl/threadedComments/threadedComment1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Please review</text>
  </threadedComment>
</ThreadedComments>`,
	})
}
