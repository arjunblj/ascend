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
					binaryLike?: boolean
					textWarning?: string
					load?: { mode?: string; isPartial?: boolean }
				}
				error?: {
					code?: string
					details?: {
						found?: boolean
						validPath?: boolean
						semantics?: string
						rule?: string
						caseInsensitiveAmbiguous?: boolean
					}
				}
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
		expect(missing.structuredContent?.error?.details?.found).toBe(false)
		expect(missing.structuredContent?.error?.details?.validPath).toBe(true)
		expect(missing.structuredContent?.error?.details?.semantics).toBe('raw-package-bytes')

		const invalid = await handler({ file: TEMP_FILE, partPath: 'xl//workbook.xml' })
		expect(invalid.isError).toBe(true)
		expect(invalid.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(invalid.structuredContent?.error?.details?.validPath).toBe(false)

		const tooLargeMaxBytes = await handler({
			file: TEMP_FILE,
			partPath: 'xl/workbook.xml',
			maxBytes: 1024 * 1024 + 1,
		})
		expect(tooLargeMaxBytes.isError).toBe(true)
		expect(tooLargeMaxBytes.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(tooLargeMaxBytes.structuredContent?.error?.details?.rule).toContain('at most')
	})

	test('ascend.raw_part returns binary base64 previews with full-byte metadata', async () => {
		const binaryBytes = Uint8Array.from({ length: 70 * 1024 }, (_, index) => index % 251)
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
				caseInsensitive?: boolean
			}) => Promise<{
				isError?: boolean
				structuredContent?: {
					ok?: boolean
					data?: {
						encoding?: string
						base64?: string
						text?: string
						previewByteLength?: number
						truncated?: boolean
						sha256?: string
						binaryLike?: boolean
						textWarning?: string
					}
					error?: {
						code?: string
						details?: { caseInsensitiveAmbiguous?: boolean }
					}
				}
			}>

			const textPreview = await handler({
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'text',
				maxBytes: 6,
			})
			expect(textPreview.structuredContent?.ok).toBe(true)
			expect(textPreview.structuredContent?.data?.binaryLike).toBe(true)
			expect(textPreview.structuredContent?.data?.textWarning).toContain('Part appears binary')

			const defaultBounded = await handler({
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'base64',
			})
			expect(defaultBounded.structuredContent?.ok).toBe(true)
			expect(defaultBounded.structuredContent?.data?.previewByteLength).toBe(64 * 1024)
			expect(defaultBounded.structuredContent?.data?.truncated).toBe(true)
			expect(defaultBounded.structuredContent?.data?.sha256).toBe(
				createHash('sha256').update(binaryBytes).digest('hex'),
			)

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

			const ambiguous = await handler({
				file: binaryFile,
				partPath: 'Xl/Media/Case.Png',
				caseInsensitive: true,
			})
			expect(ambiguous.isError).toBe(true)
			expect(ambiguous.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(ambiguous.structuredContent?.error?.details?.caseInsensitiveAmbiguous).toBe(true)
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

	test('ascend.calc supports range-scoped recalc without clearing pending formulas outside the range', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 10 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'D1', formula: 'C1*2' },
		])
		wb.recalc()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 5 },
					{ ref: 'C1', value: 20 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.calc'].handler as (args: {
			file: string
			range?: string
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					changed?: readonly string[]
					dirtyRegions?: readonly unknown[]
					generations?: { formulas?: number }
				}
			}
		}>

		const ranged = await handler({ file: TEMP_FILE, range: 'Sheet1!B1:B1' })
		expect(ranged.structuredContent?.ok).toBe(true)
		expect(ranged.structuredContent?.data?.changed).toEqual(['Sheet1!B1'])
		expect(ranged.structuredContent?.data?.dirtyRegions).toEqual([
			{ sheet: 'Sheet1', range: 'B1:B1', refs: ['Sheet1!B1'] },
		])
		expect(ranged.structuredContent?.data?.generations?.formulas).toBeNumber()
		let reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 10 })
		expect(reopened.sheet('Sheet1')?.cell('D1')?.value).toEqual({ kind: 'number', value: 20 })

		const full = await handler({ file: TEMP_FILE })
		expect(full.structuredContent?.ok).toBe(true)
		expect(full.structuredContent?.data?.changed).toEqual(['Sheet1!D1'])
		reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('D1')?.value).toEqual({ kind: 'number', value: 40 })
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

	test('ops and path mutations are mutually exclusive across MCP edit tools', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const tools = (server as any)._registeredTools as Record<
			string,
			{
				handler: (args: {
					file: string
					output?: string
					ops?: unknown[]
					mutations?: Array<{ path: string; value?: unknown }>
				}) => Promise<{
					isError?: boolean
					structuredContent?: {
						ok?: boolean
						error?: { code?: string; message?: string }
					}
				}>
			}
		>

		for (const toolName of [
			'ascend.preview',
			'ascend.plan',
			'ascend.write',
			'ascend.commit',
		] as const) {
			const result = await tools[toolName]?.handler({
				file: TEMP_FILE,
				output: `${TEMP_FILE}.out.xlsx`,
				ops: [],
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' }],
			})

			expect(result?.isError).toBe(true)
			expect(result?.structuredContent?.ok).toBe(false)
			expect(result?.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(result?.structuredContent?.error?.message).toBe(
				'Provide either ops or mutations, not both',
			)
		}
	})

	test('ascend.preview, plan, write, and commit keep escaped path mutations canonical', async () => {
		const sheetName = "Q1.Forecast's Café Δ"
		const tableName = 'Sales.Δ'
		const tablePathName = tableName.toLowerCase()
		const columnName = "Gross.Profit / Δ~'s"
		const columnPathName = columnName.toLowerCase()
		const workbookName = 'Global.Rate_Δ'
		const scopedName = 'Local.Rate_Δ'
		const definedNameRef = `'${sheetName.replace(/'/g, "''")}'!$B$2`
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
				path: `tables.${dotSegment(tablePathName)}.columns.${dotSegment(columnPathName)}.formula`,
				value: 'SUM([Region])',
			},
			{ path: `/names/${pointerSegment(workbookName)}/ref`, value: definedNameRef },
			{
				path: `sheets.${dotSegment(sheetName)}.names.${dotSegment(scopedName)}.ref`,
				value: definedNameRef,
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
			{ op: 'setDefinedName', name: workbookName, ref: definedNameRef },
			{ op: 'setDefinedName', name: scopedName, scope: sheetName, ref: definedNameRef },
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
			expect(writeReopened.definedName(workbookName)?.formula).toBe(definedNameRef)
			expect(writeReopened.definedName(scopedName, sheetName)?.formula).toBe(definedNameRef)

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
			expect(commitReopened.definedName(workbookName)?.formula).toBe(definedNameRef)
			expect(commitReopened.definedName(scopedName, sheetName)?.formula).toBe(definedNameRef)
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
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					error?: {
						message?: string
						details?: { rule?: string; reason?: string; planHandle?: string }
					}
					data?: {
						pathMutations?: { ops?: unknown[] }
						apply?: { affectedCellCount?: number }
						timings?: {
							applyMs?: number
							writePlanSummaryMs?: number
							writePolicyCheckMs?: number
							toBytesMs?: number
							outputByteReadMs?: number
						}
						postWrite?: {
							valid?: boolean
							auditsPassed?: boolean
							expectedPackageGraphIssueCount?: number
							unresolvedPackageGraphIssueCount?: number
							reopened?: boolean
							timings?: { reopenMs?: number }
							check?: { valid?: boolean }
							packageGraphAudit?: { ok?: boolean }
						}
						trace?: { artifactCount?: number; artifacts?: unknown[] }
					}
				}
			}>
			const planned = await plan({
				file: TEMP_FILE,
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
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 321 }] },
			])
			expect(committed.structuredContent?.data?.apply?.affectedCellCount).toBe(1)
			expect(committed.structuredContent?.data?.timings?.applyMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.writePlanSummaryMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.writePolicyCheckMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.toBytesMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.outputByteReadMs).toBeNumber()
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.expectedPackageGraphIssueCount).toBe(0)
			expect(committed.structuredContent?.data?.postWrite?.unresolvedPackageGraphIssueCount).toBe(0)
			expect(committed.structuredContent?.data?.postWrite?.reopened).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.timings?.reopenMs).toBeNumber()
			expect(committed.structuredContent?.data?.postWrite?.check?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)
			expect(committed.structuredContent?.data?.trace?.artifactCount).toBeNumber()
			expect(committed.structuredContent?.data?.trace?.artifacts).toBeUndefined()
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 321 })

			const reused = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: `${output}.reuse.xlsx`,
				approvals: [],
			})
			expect(reused.structuredContent?.ok).toBe(false)
			expect(reused.structuredContent?.error?.message).toBe(
				'Prepared plan handle has already been used',
			)
			expect(reused.structuredContent?.error?.details).toMatchObject({
				rule: 'prepared-plan-handle-unavailable',
				reason: 'already-used',
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
			})
		} finally {
			await unlink(output).catch(() => {})
			await unlink(`${output}.reuse.xlsx`).catch(() => {})
		}
	})

	test('direct MCP path mutation commits preserve in-place backups and post-write truth', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'original' }] }])
		await wb.save(TEMP_FILE)
		const backup = `${TEMP_FILE}.direct-backup-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				file: string
				mutations: Array<{ path: string; value?: unknown }>
				inPlace?: boolean
				backup?: string
				approvals?: string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						output?: string
						backup?: string
						outputSha256?: string
						pathMutations?: { ops?: unknown[] }
						postWrite?: {
							valid?: boolean
							outputSha256?: string
							auditsPassed?: boolean
							reopened?: boolean
							check?: { valid?: boolean }
							packageGraphAudit?: { ok?: boolean }
						}
					}
				}
			}>

			const committed = await commit({
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'updated' }],
				inPlace: true,
				backup,
				approvals: [],
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.output).toBe(TEMP_FILE)
			expect(committed.structuredContent?.data?.backup).toBe(backup)
			expect(committed.structuredContent?.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'updated' }] },
			])
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.reopened).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.outputSha256).toBe(
				committed.structuredContent?.data?.outputSha256,
			)
			expect(committed.structuredContent?.data?.postWrite?.check?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

			const reopenedInput = await AscendWorkbook.open(TEMP_FILE)
			expect(reopenedInput.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'updated',
			})
			const reopenedBackup = await AscendWorkbook.open(backup)
			expect(reopenedBackup.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'original',
			})
		} finally {
			await unlink(backup).catch(() => {})
		}
	})

	test('prepared MCP path mutation handles preserve in-place backups and remain one-shot', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'original' }] }])
		await wb.save(TEMP_FILE)
		const backup = `${TEMP_FILE}.prepared-backup-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				mutations: Array<{ path: string; value?: unknown }>
			}) => Promise<{
				structuredContent?: { data?: { preparedPlan?: { id?: string } } }
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				inPlace?: boolean
				backup?: string
				approvals?: string[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					error?: { message?: string }
					data?: {
						output?: string
						outputSha256?: string
						backup?: string
						pathMutations?: { ops?: unknown[] }
						postWrite?: {
							valid?: boolean
							auditsPassed?: boolean
							reopened?: boolean
							outputSha256?: string
							check?: { valid?: boolean }
							packageGraphAudit?: { ok?: boolean }
						}
					}
				}
			}>
			const planned = await plan({
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'updated' }],
			})
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				inPlace: true,
				backup,
				approvals: [],
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.output).toBe(TEMP_FILE)
			expect(committed.structuredContent?.data?.backup).toBe(backup)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'updated' }] },
			])
			expect(committed.structuredContent?.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.reopened).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.outputSha256).toBe(
				committed.structuredContent?.data?.outputSha256,
			)
			expect(committed.structuredContent?.data?.postWrite?.check?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

			const reopenedInput = await AscendWorkbook.open(TEMP_FILE)
			expect(reopenedInput.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'updated',
			})
			const reopenedBackup = await AscendWorkbook.open(backup)
			expect(reopenedBackup.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'original',
			})

			const reused = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				inPlace: true,
				backup,
				approvals: [],
			})
			expect(reused.structuredContent?.ok).toBe(false)
			expect(reused.structuredContent?.error?.message).toBe(
				'Prepared plan handle has already been used',
			)
		} finally {
			await unlink(backup).catch(() => {})
		}
	})

	test('prepared MCP plan handles require exact destructive approval ids', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Scratch' }])
		await wb.save(TEMP_FILE)
		const output = `${TEMP_FILE}.prepared-approval-mcp.xlsx`
		const ops = [{ op: 'deleteSheet', sheet: 'Scratch' }]
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				ops: unknown[]
			}) => Promise<{
				structuredContent?: {
					data?: {
						preparedPlan?: { id?: string }
						approvals?: Array<{ id: string }>
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
					error?: { message?: string }
					data?: { approvals?: Array<{ id: string }> }
				}
			}>
			const planned = await plan({ file: TEMP_FILE, ops })
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			const approvalId = planned.structuredContent?.data?.approvals?.[0]?.id
			expect(approvalId).toBe('op:0:deletesheet')

			const aliasCommit = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: ['deleteSheet'],
			})
			expect(aliasCommit.structuredContent?.ok).toBe(false)
			expect(aliasCommit.structuredContent?.error?.message).toBe(
				'Commit requires explicit approval',
			)
			expect(await Bun.file(output).exists()).toBe(false)

			const retryPlan = await plan({ file: TEMP_FILE, ops })
			const exactCommit = await commit({
				planHandle: retryPlan.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: [approvalId ?? ''],
			})
			expect(exactCommit.structuredContent?.ok).toBe(true)
			expect(exactCommit.structuredContent?.data?.approvals?.[0]?.id).toBe(approvalId)
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheets).not.toContain('Scratch')
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('prepared MCP path mutation handles require exact preserved-loss approval ids', async () => {
		await Bun.write(TEMP_MACRO_FILE, signedMacroWorkbook())
		const output = `${TEMP_MACRO_OUTPUT}.prepared-path.xlsm`
		const mutations = [{ path: '/sheets/Data/cells/A1/value', value: 17 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 17 }] }]
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				mutations: Array<{ path: string; value?: unknown }>
			}) => Promise<{
				structuredContent?: {
					data?: {
						preparedPlan?: { id?: string }
						approvals?: Array<{ id: string }>
						pathMutations?: { replayable?: boolean; ops?: unknown[] }
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: string[]
			}) => Promise<{
				isError?: boolean
				structuredContent?: {
					ok?: boolean
					error?: { message?: string }
					data?: {
						approvals?: Array<{ id: string }>
						pathMutations?: { ops?: unknown[] }
						postWrite?: { valid?: boolean; reopened?: boolean }
					}
				}
			}>
			const planned = await plan({ file: TEMP_MACRO_FILE, mutations })
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(planned.structuredContent?.data?.pathMutations?.replayable).toBe(true)
			expect(planned.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual(
				expect.arrayContaining([
					expect.stringMatching(/^loss:preservedmacro:preserved:/),
					expect.stringMatching(/^loss:preservedsignature:preserved:/),
				]),
			)

			const aliasCommit = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: ['preservedMacro', 'preservedSignature'],
			})
			expect(aliasCommit.structuredContent?.ok).toBe(false)
			expect(aliasCommit.structuredContent?.error?.message).toBe(
				'Commit requires explicit approval',
			)
			expect(await Bun.file(output).exists()).toBe(false)

			const retryPlan = await plan({ file: TEMP_MACRO_FILE, mutations })
			expect(retryPlan.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(retryPlan.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const exactCommit = await commit({
				planHandle: retryPlan.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
			})
			expect(exactCommit.structuredContent?.ok).toBe(true)
			expect(exactCommit.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			expect(
				exactCommit.structuredContent?.data?.approvals?.map((approval) => approval.id),
			).toEqual(approvalIds)
			expect(exactCommit.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(exactCommit.structuredContent?.data?.postWrite?.reopened).toBe(true)
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Data')?.cell('A1')?.value).toEqual({ kind: 'number', value: 17 })
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('prepared MCP path mutation handles surface post-write audit failures as blocked output', async () => {
		await Bun.write(TEMP_FILE, preservedCustomWorkbook())
		const output = `${TEMP_FILE}.prepared-preserved-mcp.xlsx`
		const mutations = [{ path: '/sheets/Sheet1/cells/A1/value', value: 17 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 17 }] }]
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				mutations: Array<{ path: string; value?: unknown }>
			}) => Promise<{
				structuredContent?: {
					data?: {
						preparedPlan?: { id?: string }
						approvals?: Array<{ id: string }>
						pathMutations?: { ops?: unknown[] }
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: string[]
				compact?: boolean
			}) => Promise<{
				isError?: boolean
				structuredContent?: {
					ok?: boolean
					error?: { message?: string }
					data?: {
						outputSha256?: string
						approvals?: Array<{ id: string }>
						pathMutations?: { ops?: unknown[] }
						postWrite?: {
							valid?: boolean
							outputSha256?: string
							auditsPassed?: boolean
							expectedPackageGraphIssueCount?: number
							unresolvedPackageGraphIssueCount?: number
							packageGraphAudit?: { ok?: boolean; issueCount?: number }
						}
						modelOutput?: {
							blocked?: boolean
							nextActions?: readonly string[]
							counts?: { postWritePackageGraphIssues?: number }
						}
					}
				}
			}>
			const planned = await plan({ file: TEMP_FILE, mutations })
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(planned.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedother:preserved:/)])

			const aliasCommit = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: ['preservedOther'],
				compact: true,
			})
			expect(aliasCommit.isError).toBe(true)
			expect(aliasCommit.structuredContent?.ok).toBe(false)
			expect(aliasCommit.structuredContent?.error?.message).toBe(
				'Commit requires explicit approval',
			)
			expect(await Bun.file(output).exists()).toBe(false)

			const retryPlan = await plan({ file: TEMP_FILE, mutations })
			expect(retryPlan.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(retryPlan.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const exactCommit = await commit({
				planHandle: retryPlan.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(exactCommit.isError).not.toBe(true)
			expect(exactCommit.structuredContent?.ok).toBe(true)
			expect(exactCommit.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			expect(
				exactCommit.structuredContent?.data?.approvals?.map((approval) => approval.id),
			).toEqual(approvalIds)
			expect(exactCommit.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(exactCommit.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
			expect(exactCommit.structuredContent?.data?.postWrite?.outputSha256).toBe(
				exactCommit.structuredContent?.data?.outputSha256,
			)
			expect(exactCommit.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(false)
			expect(
				exactCommit.structuredContent?.data?.postWrite?.packageGraphAudit?.issueCount,
			).toBeGreaterThan(0)
			expect(exactCommit.structuredContent?.data?.postWrite?.expectedPackageGraphIssueCount).toBe(0)
			expect(
				exactCommit.structuredContent?.data?.postWrite?.unresolvedPackageGraphIssueCount,
			).toBeGreaterThan(0)
			expect(exactCommit.structuredContent?.data?.modelOutput?.blocked).toBe(true)
			expect(
				exactCommit.structuredContent?.data?.modelOutput?.counts?.postWritePackageGraphIssues,
			).toBeGreaterThan(0)
			expect(exactCommit.structuredContent?.data?.modelOutput?.nextActions?.join('\n')).toContain(
				'postWrite.packageGraphAudit.issues',
			)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('prepared MCP commits surface post-write formula lint failures as blocked output', async () => {
		const input = `${TEMP_FILE}.prepared-lint-source.xlsx`
		const output = `${TEMP_FILE}.prepared-lint-out.xlsx`
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const complexFormula = `=${Array.from({ length: 26 }, () => '1').join('+')}`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				ops: unknown[]
			}) => Promise<{
				structuredContent?: { data?: { preparedPlan?: { id?: string } } }
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				compact?: boolean
			}) => Promise<{
				isError?: boolean
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							valid?: boolean
							auditsPassed?: boolean
							lint?: { clean?: boolean; errorCount?: number }
							packageGraphAudit?: { ok?: boolean }
						}
						modelOutput?: {
							blocked?: boolean
							nextActions?: readonly string[]
							counts?: { postWriteLintFailures?: number }
						}
					}
				}
			}>
			const planned = await plan({
				file: input,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: complexFormula }],
			})
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(committed.isError).not.toBe(true)
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.lint?.clean).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.lint?.errorCount).toBeGreaterThan(0)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)
			expect(committed.structuredContent?.data?.modelOutput?.blocked).toBe(true)
			expect(
				committed.structuredContent?.data?.modelOutput?.counts?.postWriteLintFailures,
			).toBeGreaterThan(0)
			expect(committed.structuredContent?.data?.modelOutput?.nextActions?.join('\n')).toContain(
				'postWrite.lint.warnings',
			)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('direct MCP commits surface post-write formula lint failures as blocked output', async () => {
		const input = `${TEMP_FILE}.direct-lint-source.xlsx`
		const output = `${TEMP_FILE}.direct-lint-out.xlsx`
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const complexFormula = `=${Array.from({ length: 26 }, () => '1').join('+')}`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				file?: string
				ops?: unknown[]
				output?: string
				compact?: boolean
			}) => Promise<{
				isError?: boolean
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							valid?: boolean
							auditsPassed?: boolean
							lint?: { clean?: boolean; errorCount?: number }
							packageGraphAudit?: { ok?: boolean }
						}
						modelOutput?: {
							blocked?: boolean
							nextActions?: readonly string[]
							counts?: { postWriteLintFailures?: number }
						}
					}
				}
			}>
			const committed = await commit({
				file: input,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: complexFormula }],
				output,
				compact: true,
			})
			expect(committed.isError).not.toBe(true)
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.lint?.clean).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.lint?.errorCount).toBeGreaterThan(0)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)
			expect(committed.structuredContent?.data?.modelOutput?.blocked).toBe(true)
			expect(
				committed.structuredContent?.data?.modelOutput?.counts?.postWriteLintFailures,
			).toBeGreaterThan(0)
			expect(committed.structuredContent?.data?.modelOutput?.nextActions?.join('\n')).toContain(
				'postWrite.lint.warnings',
			)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('ascend.plan can opt out of default prepared handles', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const server = createServer()

		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			prepare?: boolean
			ops?: unknown[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					preparedPlan?: { id?: string }
					preview?: { wouldSucceed?: boolean }
				}
			}
		}>

		const planned = await plan({
			file: TEMP_FILE,
			prepare: false,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 12 }] }],
		})

		expect(planned.structuredContent?.ok).toBe(true)
		expect(planned.structuredContent?.data?.preparedPlan).toBeUndefined()
		expect(planned.structuredContent?.data?.preview?.wouldSucceed).toBe(true)
	})

	test('prepared MCP path mutation handles reject stale input before writing output', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${TEMP_FILE}.prepared-stale-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				mutations: Array<{ path: string; value?: unknown }>
			}) => Promise<{
				structuredContent?: { data?: { preparedPlan?: { id?: string } } }
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: string[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					error?: {
						code?: string
						message?: string
						details?: {
							expected?: string
							actual?: string
							planDigest?: string
						}
					}
				}
			}>
			const planned = await plan({
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 123 }],
			})
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()

			const changed = AscendWorkbook.create()
			changed.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 9 }] }])
			await changed.save(TEMP_FILE)

			const stale = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: [],
			})
			expect(stale.structuredContent?.ok).toBe(false)
			expect(stale.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(stale.structuredContent?.error?.message).toBe(
				'Input workbook changed after agent plan was prepared',
			)
			expect(stale.structuredContent?.error?.details?.expected).toMatch(/^[a-f0-9]{64}$/)
			expect(stale.structuredContent?.error?.details?.actual).toMatch(/^[a-f0-9]{64}$/)
			expect(stale.structuredContent?.error?.details?.actual).not.toBe(
				stale.structuredContent?.error?.details?.expected,
			)
			expect(stale.structuredContent?.error?.details?.planDigest).toMatch(/^[a-f0-9]{64}$/)
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
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
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					error?: {
						message?: string
						details?: { rule?: string; reason?: string; planHandle?: string }
					}
				}
			}>

			const first = await plan({
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
			})
			const second = await plan({
				file: TEMP_FILE,
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
			expect(evicted.structuredContent?.error?.message).toBe('Prepared plan handle was evicted')
			expect(evicted.structuredContent?.error?.details).toMatchObject({
				rule: 'prepared-plan-handle-unavailable',
				reason: 'evicted',
				planHandle: first.structuredContent?.data?.preparedPlan?.id,
			})

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

	test('prepared MCP plan handles expire with structured reason metadata', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		let now = 1_000
		const server = createServer({ preparedPlanTtlMs: 10, now: () => now })

		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			mutations: Array<{ path: string; value?: unknown }>
		}) => Promise<{
			structuredContent?: { data?: { preparedPlan?: { id?: string; ttlMs?: number } } }
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			planHandle?: string
			output?: string
			approvals?: string[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				error?: {
					message?: string
					details?: { rule?: string; reason?: string; planHandle?: string }
				}
			}
		}>

		const planned = await plan({
			file: TEMP_FILE,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 9 }],
		})
		expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
		expect(planned.structuredContent?.data?.preparedPlan?.ttlMs).toBe(10)

		now += 11
		const expired = await commit({
			planHandle: planned.structuredContent?.data?.preparedPlan?.id,
			output: `${TEMP_FILE}.expired-prepared-mcp.xlsx`,
			approvals: [],
		})
		expect(expired.structuredContent?.ok).toBe(false)
		expect(expired.structuredContent?.error?.message).toBe('Prepared plan handle expired')
		expect(expired.structuredContent?.error?.details).toMatchObject({
			rule: 'prepared-plan-handle-unavailable',
			reason: 'expired',
			planHandle: planned.structuredContent?.data?.preparedPlan?.id,
		})
		await unlink(`${TEMP_FILE}.expired-prepared-mcp.xlsx`).catch(() => {})
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
					prepare?: boolean
					mutations: Array<{ path: string; value?: unknown }>
				}) => Promise<{
					isError?: boolean
					structuredContent?: {
						ok?: boolean
						data?: { preparedPlan?: { id?: string } }
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

	test('invalid path mutation shapes return structured MCP repair details consistently', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const tools = (server as any)._registeredTools as Record<
			string,
			{
				handler: (args: {
					file: string
					output?: string
					mutations: readonly unknown[]
				}) => Promise<{
					isError?: boolean
					structuredContent?: {
						ok?: boolean
						error?: {
							code?: string
							details?: {
								issueCount?: number
								issues?: readonly string[]
								issueDetails?: readonly {
									code?: string
									mutationIndex?: number
									path?: string
								}[]
							}
						}
					}
				}>
			}
		>

		for (const toolName of [
			'ascend.preview',
			'ascend.plan',
			'ascend.write',
			'ascend.commit',
		] as const) {
			const result = await tools[toolName]?.handler({
				file: TEMP_FILE,
				output: `${TEMP_FILE}.out.xlsx`,
				mutations: [{ path: 123, value: 'new' }],
			})

			expect(result?.isError).toBe(true)
			expect(result?.structuredContent?.ok).toBe(false)
			expect(result?.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(result?.structuredContent?.error?.details?.issueCount).toBe(1)
			expect(result?.structuredContent?.error?.details?.issues).toEqual([
				'mutations[0]: Mutation path must be a string or string array.',
			])
			expect(result?.structuredContent?.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'invalid_path_mutation',
					mutationIndex: 0,
					path: 'mutations[0]',
				}),
			])
		}
	})

	test('non-replayable path mutation batches do not expose or apply partial MCP ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'old' },
					{ ref: 'B1', value: 'Amount' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])
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
					{ path: '/sheets/Sheet1/name', value: 'Bad/Name' },
					{ path: '/tables/Sales/name', value: 'Bad Name' },
				],
			})

			expect(result?.isError).toBe(true)
			expect(result?.structuredContent?.ok).toBe(false)
			expect(result?.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
			expect(result?.structuredContent?.error?.details?.issueCount).toBe(2)
			expect(result?.structuredContent?.error?.details?.compiledOps).toEqual([])
			expect(result?.structuredContent?.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'invalid_value',
					path: '/sheets/Sheet1/name',
				}),
				expect.objectContaining({
					code: 'invalid_value',
					path: '/tables/Sales/name',
				}),
			])
		}

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})
		expect(reopened.sheets).toContain('Sheet1')
		expect(reopened.table('Sales')?.name).toBe('Sales')

		const prepared = await tools['ascend.plan']?.handler({
			file: TEMP_FILE,
			prepare: true,
			mutations: [
				{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' },
				{ path: '/sheets/Sheet1/name', value: 'Bad/Name' },
			],
		})
		expect(prepared?.isError).toBe(true)
		expect(prepared?.structuredContent?.ok).toBe(false)
		expect(prepared?.structuredContent?.data?.preparedPlan).toBeUndefined()
		expect(prepared?.structuredContent?.error?.details?.compiledOps).toEqual([])
		expect(prepared?.structuredContent?.error?.details?.issueDetails).toEqual([
			expect.objectContaining({
				code: 'invalid_value',
				path: '/sheets/Sheet1/name',
			}),
		])
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

	test('ascend.read compact changedSince returns a fresh window after source changes', async () => {
		const original = AscendWorkbook.create()
		original.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await original.save(TEMP_FILE)

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
				data?: { cells?: unknown[]; changeToken?: string }
			}
		}>

		const first = await handler({ file: TEMP_FILE, range: 'A1:A1', format: 'compact' })
		expect(first.structuredContent?.ok).toBe(true)
		expect(first.structuredContent?.data?.cells).toEqual([[0, 0, 'old']])
		expect(first.structuredContent?.data?.changeToken).toBeDefined()

		const changed = AscendWorkbook.create()
		changed.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'new' }] }])
		await changed.save(TEMP_FILE)

		const afterChange = await handler({
			file: TEMP_FILE,
			range: 'A1:A1',
			format: 'compact',
			changedSince: first.structuredContent?.data?.changeToken,
		})
		expect(afterChange.structuredContent?.ok).toBe(true)
		expect(afterChange.structuredContent?.data?.cells).toEqual([[0, 0, 'new']])
		expect(afterChange.structuredContent?.data?.changeToken).toBeDefined()
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
					snapshot?: {
						token?: string
						generations?: {
							workbook?: number
							sheetMetadata?: number
							formulas?: number
							styles?: number
						}
						load?: {
							mode?: string
							isPartial?: boolean
							maxRows?: number
						}
					}
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
		expect(result.structuredContent?.data?.snapshot?.token).toContain('partial')
		expect(result.structuredContent?.data?.snapshot?.generations).toEqual({
			workbook: 0,
			sheetMetadata: 0,
			formulas: 0,
			styles: 0,
		})
		expect(result.structuredContent?.data?.snapshot?.load).toMatchObject({
			mode: 'values',
			isPartial: true,
			maxRows: 3,
		})
		expect(result.structuredContent?.data?.load?.mode).toBe('values')
		expect(result.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(result.structuredContent?.data?.load?.maxRows).toBe(3)
		expect(result.structuredContent?.data?.load?.partialReasons).toContain(
			'only the first 3 row(s) are hydrated per loaded sheet',
		)
		expect(result.structuredContent?.data?.load?.cellsHydrated).toBe(true)
		expect(result.structuredContent?.data?.load?.loadedSheets).toEqual(['Sheet1'])
	})

	test('ascend.read preview defaults compact reads to a bounded first window', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 520 }, (_, row) => [
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
			preview?: boolean
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
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:B520',
			format: 'compact',
			preview: true,
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.rowCount).toBe(500)
		expect(result.structuredContent?.data?.cells).toHaveLength(1000)
		expect(result.structuredContent?.data?.load?.mode).toBe('values')
		expect(result.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(result.structuredContent?.data?.load?.maxRows).toBe(500)
		expect(result.structuredContent?.data?.load?.partialReasons).toContain(
			'only the first 500 row(s) are hydrated per loaded sheet',
		)
	})

	test('ascend.read compact reads default to a bounded first window', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 520 }, (_, row) => [
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
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:B520',
			format: 'compact',
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.rowCount).toBe(500)
		expect(result.structuredContent?.data?.cells).toHaveLength(1000)
		expect(result.structuredContent?.data?.load?.mode).toBe('values')
		expect(result.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(result.structuredContent?.data?.load?.maxRows).toBe(500)
		expect(result.structuredContent?.data?.load?.partialReasons).toContain(
			'only the first 500 row(s) are hydrated per loaded sheet',
		)
	})

	test('ascend.agent_view exposes partial-load metadata for sheet-scoped capped formula views', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Data' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
			{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'hidden' }] },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.agent_view'].handler as (args: {
			file: string
			sheet?: string
			range: string
			maxRows?: number
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					load?: {
						isPartial?: boolean
						maxRows?: number
						partialReasons?: readonly string[]
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			sheet: 'Sheet1',
			range: 'A1:A3',
			maxRows: 1,
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(result.structuredContent?.data?.load?.maxRows).toBe(1)
		expect(result.structuredContent?.data?.load?.partialReasons).toContain(
			'only selected sheets are loaded',
		)
		expect(result.structuredContent?.data?.load?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
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

	test('ascend.check and ascend.lint expose partial-load metadata for capped formula views', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const checkHandler = (server as any)._registeredTools['ascend.check'].handler as (args: {
			file: string
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					details?: {
						check?: {
							valid?: boolean
							issues?: readonly { readonly rule?: string }[]
							load?: {
								isPartial?: boolean
								maxRows?: number
								partialReasons?: readonly string[]
							}
						}
					}
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const lintHandler = (server as any)._registeredTools['ascend.lint'].handler as (args: {
			file: string
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				data?: {
					clean?: boolean
					warnings?: readonly { readonly rule?: string }[]
					load?: {
						isPartial?: boolean
						maxRows?: number
						partialReasons?: readonly string[]
					}
				}
			}
		}>

		const check = await checkHandler({ file: TEMP_FILE, maxRows: 1 })

		expect(check.isError).toBe(true)
		expect(check.structuredContent?.ok).toBe(false)
		expect(check.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(check.structuredContent?.error?.details?.check?.valid).toBe(false)
		expect(check.structuredContent?.error?.details?.check?.issues?.[0]?.rule).toBe(
			'partial-dependency-analysis',
		)
		expect(check.structuredContent?.error?.details?.check?.load?.isPartial).toBe(true)
		expect(check.structuredContent?.error?.details?.check?.load?.maxRows).toBe(1)
		expect(check.structuredContent?.error?.details?.check?.load?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)

		const lint = await lintHandler({ file: TEMP_FILE, maxRows: 1 })

		expect(lint.isError).toBeFalsy()
		expect(lint.structuredContent?.ok).toBe(true)
		expect(lint.structuredContent?.data?.clean).toBe(false)
		expect(lint.structuredContent?.data?.warnings?.[0]?.rule).toBe('partial-dependency-analysis')
		expect(lint.structuredContent?.data?.load?.isPartial).toBe(true)
		expect(lint.structuredContent?.data?.load?.maxRows).toBe(1)
		expect(lint.structuredContent?.data?.load?.partialReasons).toContain(
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
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops?: unknown[]
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { preparedPlan?: { id?: string } } }
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			planHandle?: string
			output?: string
			compact?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					outputSha256?: string
					postWrite?: {
						valid?: boolean
						auditsPassed?: boolean
						reopened?: boolean
						outputSha256?: string
						check?: { valid?: boolean }
						packageGraphAudit?: { ok?: boolean }
					}
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

		const replayInput = `${TEMP_FILE}.dump-replay-input.xlsx`
		const replayOutput = `${TEMP_FILE}.dump-replay-output.xlsx`
		try {
			await AscendWorkbook.create().save(replayInput)
			const planned = await plan({ file: replayInput, ops: result.structuredContent?.data?.ops })
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: replayOutput,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.reopened).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.outputSha256).toBe(
				committed.structuredContent?.data?.outputSha256,
			)
			expect(committed.structuredContent?.data?.postWrite?.check?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

			const replayed = await AscendWorkbook.open(replayOutput)
			expect(replayed.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'number',
				value: 10,
			})
			expect(replayed.sheet('Sheet1')?.cell('B1')?.value).toEqual({
				kind: 'string',
				value: 'label',
			})
			expect(replayed.sheet('Sheet1')?.cell('B2')?.formula).toBe('A1*2')
		} finally {
			await unlink(replayInput).catch(() => {})
			await unlink(replayOutput).catch(() => {})
		}
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
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops?: unknown[]
		}) => Promise<{
			structuredContent?: { ok?: boolean; data?: { preparedPlan?: { id?: string } } }
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			planHandle?: string
			output?: string
			compact?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					outputSha256?: string
					postWrite?: {
						valid?: boolean
						auditsPassed?: boolean
						reopened?: boolean
						outputSha256?: string
						check?: { valid?: boolean }
						packageGraphAudit?: { ok?: boolean }
					}
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

		const replayable = await handler({
			file: TEMP_FILE,
			sheet: 'Sheet1',
			data: { amount: 10, tax: 2, client: 'Acme' },
		})
		expect(replayable.structuredContent?.ok).toBe(true)
		expect(replayable.structuredContent?.data?.replayable).toBe(true)
		expect(replayable.structuredContent?.data?.unresolved).toEqual([])

		const replayOutput = `${TEMP_FILE}.template-replay-output.xlsx`
		try {
			const planned = await plan({ file: TEMP_FILE, ops: replayable.structuredContent?.data?.ops })
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: replayOutput,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.reopened).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.outputSha256).toBe(
				committed.structuredContent?.data?.outputSha256,
			)
			expect(committed.structuredContent?.data?.postWrite?.check?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

			const merged = await AscendWorkbook.open(replayOutput)
			expect(merged.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 10 })
			expect(merged.sheet('Sheet1')?.cell('A2')?.value).toEqual({
				kind: 'string',
				value: 'Missing Acme',
			})
			expect(merged.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1+2')
		} finally {
			await unlink(replayOutput).catch(() => {})
		}
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

	test('ascend.plan rejects capped load options instead of silently producing full plans', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops?: unknown[]
			mutations?: Array<{ path: string; value?: unknown }>
			prepare?: boolean
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				data?: { preparedPlan?: { id?: string } }
				error?: {
					code?: string
					details?: {
						unsupportedLoadOptions?: readonly string[]
						requiredLoad?: { mode?: string; allSheets?: boolean; maxRows?: null | number }
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			maxRows: 1,
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(result.structuredContent?.error?.details?.unsupportedLoadOptions).toEqual(['maxRows'])
		expect(result.structuredContent?.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})

		const mutationPlan = await handler({
			file: TEMP_FILE,
			prepare: true,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
			maxRows: 1,
		})
		expect(mutationPlan.isError).toBe(true)
		expect(mutationPlan.structuredContent?.data?.preparedPlan).toBeUndefined()
		expect(mutationPlan.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(mutationPlan.structuredContent?.error?.details?.unsupportedLoadOptions).toEqual([
			'maxRows',
		])
	})

	test('ascend.commit rejects capped load options instead of silently producing full commits', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			file: string
			ops?: unknown[]
			mutations?: Array<{ path: string; value?: unknown }>
			output: string
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					code?: string
					details?: {
						unsupportedLoadOptions?: readonly string[]
						requiredLoad?: { mode?: string; allSheets?: boolean; maxRows?: null | number }
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			output: `${TEMP_FILE}.partial-commit.xlsx`,
			maxRows: 1,
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(result.structuredContent?.error?.details?.unsupportedLoadOptions).toEqual(['maxRows'])
		expect(result.structuredContent?.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})

		const mutationOutput = `${TEMP_FILE}.partial-mutation-commit.xlsx`
		const mutationCommit = await handler({
			file: TEMP_FILE,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
			output: mutationOutput,
			maxRows: 1,
		})
		expect(mutationCommit.isError).toBe(true)
		expect(mutationCommit.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(mutationCommit.structuredContent?.error?.details?.unsupportedLoadOptions).toEqual([
			'maxRows',
		])
		expect(await Bun.file(mutationOutput).exists()).toBe(false)
		await unlink(mutationOutput).catch(() => {})
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

function preservedCustomWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/custom/custom1.xml" ContentType="application/custom+xml"/>
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
		'xl/custom/custom1.xml': '<custom>preserve me</custom>',
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
				'xl/media/case.png': new Uint8Array([1]),
				'XL/MEDIA/CASE.PNG': new Uint8Array([2]),
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
