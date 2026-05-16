import { afterAll, describe, expect, test } from 'bun:test'
import { Buffer } from 'node:buffer'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { unlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
	AscendWorkbook,
	MUTATION_JOURNAL_ISSUE_SCHEMA,
	MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
	parseOperations,
} from '@ascend/sdk'
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
const TEMP_PACKAGE_ACTION_OUTPUT = join(
	tmpdir(),
	`ascend-mcp-package-actions-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const PIVOT_FIXTURE = join(
	import.meta.dir,
	'../../../fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
)
const CHARTSHEET_FIXTURE = join(import.meta.dir, '../../../fixtures/xlsx/exceljs/chart-sheet.xlsx')
const ENCRYPTED_FIXTURE = join(
	import.meta.dir,
	'../../../fixtures/xlsx/calamine/pass_protected.xlsx',
)
const JOURNAL_V1_FIXTURE = JSON.parse(
	readFileSync(
		join(import.meta.dir, '../../../fixtures/journal/mutation-journal-v1.json'),
		'utf-8',
	),
) as {
	readonly scenario: {
		readonly ops: readonly Record<string, unknown>[]
		readonly journal: {
			readonly schemaVersion: number
			readonly schemaId: string
			readonly supported: boolean
			readonly exact: boolean
			readonly inverseOpCount: number
			readonly issueCount: number
			readonly issues: readonly unknown[]
		}
	}
}

function compactJournal(journal: {
	readonly schemaVersion?: number
	readonly schemaId?: string
	readonly supported?: boolean
	readonly exact?: boolean
	readonly inverseOps?: readonly unknown[]
	readonly issues?: readonly unknown[]
}): typeof JOURNAL_V1_FIXTURE.scenario.journal {
	const { schemaVersion, schemaId, supported, exact, inverseOps, issues } = journal
	if (
		schemaVersion === undefined ||
		schemaId === undefined ||
		supported === undefined ||
		exact === undefined ||
		inverseOps === undefined ||
		issues === undefined
	) {
		throw new Error('journal is missing required v1 fields')
	}
	return {
		schemaVersion,
		schemaId,
		supported,
		exact,
		inverseOpCount: inverseOps.length,
		issueCount: issues.length,
		issues,
	}
}

afterAll(async () => {
	await unlink(TEMP_FILE).catch(() => {})
	await unlink(TEMP_MACRO_FILE).catch(() => {})
	await unlink(TEMP_MACRO_OUTPUT).catch(() => {})
	await unlink(TEMP_PACKAGE_ACTION_OUTPUT).catch(() => {})
})

describe('MCP server', () => {
	test('ascend.plan reports missing workbook files with structured retry guidance', async () => {
		const missing = join(tmpdir(), `ascend-mcp-missing-${Date.now()}.xlsx`)
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					retryable?: boolean
					retryStrategy?: string
					details?: { file?: string }
					suggestedFix?: string
				}
			}
		}>

		const result = await plan({
			file: missing,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent).toMatchObject({
			ok: false,
			error: {
				code: 'FILE_NOT_FOUND',
				retryable: true,
				retryStrategy: 'modified',
				details: { file: missing },
				suggestedFix: expect.stringContaining('existing workbook path'),
			},
		})
	})

	test('ascend.commit reports missing workbook files without creating output artifacts', async () => {
		const missing = join(tmpdir(), `ascend-mcp-missing-commit-${Date.now()}.xlsx`)
		const output = `${missing}.out.xlsx`
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			file: string
			output: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					retryable?: boolean
					retryStrategy?: string
					details?: { file?: string }
					suggestedFix?: string
				}
			}
		}>

		try {
			const result = await commit({
				file: missing,
				output,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			})

			expect(result.isError).toBe(true)
			expect(result.structuredContent).toMatchObject({
				ok: false,
				error: {
					code: 'FILE_NOT_FOUND',
					retryable: true,
					retryStrategy: 'modified',
					details: { file: missing },
					suggestedFix: expect.stringContaining('existing workbook path'),
				},
			})
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('ascend.commit rejects missing workbook references with structured retry guidance', async () => {
		const output = join(tmpdir(), `ascend-mcp-missing-reference-${Date.now()}.xlsx`)
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
			output?: string
			ops?: readonly Record<string, unknown>[]
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: { required?: string[] }
					suggestedFix?: string
				}
			}
		}>

		try {
			const result = await commit({
				output,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			})

			expect(result.isError).toBe(true)
			expect(result.structuredContent).toMatchObject({
				ok: false,
				error: {
					code: 'VALIDATION_ERROR',
					message: 'Missing or invalid commit workbook reference',
					retryable: true,
					retryStrategy: 'modified',
					details: { required: ['file or planHandle'] },
					suggestedFix: expect.stringContaining('Pass either file with ops/mutations'),
				},
			})
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('ascend.export reports unsupported formats with structured retry guidance', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = join(tmpdir(), `ascend-mcp-unsupported-export-${Date.now()}.bad`)
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const exportTool = (server as any)._registeredTools['ascend.export'].handler as (args: {
			file: string
			output: string
			format?: string
		}) => Promise<{
			isError?: boolean
			content?: Array<{ type?: string; text?: string }>
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: {
						field?: string
						received?: string
						allowedFormats?: readonly string[]
						workflow?: readonly string[]
					}
					suggestedFix?: string
				}
			}
		}>

		try {
			const result = await exportTool({ file: TEMP_FILE, output, format: 'bad' })

			expect(result.isError).toBe(true)
			expect(result.content?.[0]?.text).toBe('Unsupported export format: bad')
			expect(result.structuredContent).toMatchObject({
				ok: false,
				error: {
					code: 'VALIDATION_ERROR',
					message: 'Unsupported export format: bad',
					retryable: true,
					retryStrategy: 'modified',
					details: {
						field: 'format',
						received: 'bad',
						allowedFormats: ['csv', 'tsv', 'json', 'xlsx', 'xlsm'],
						workflow: ['reopen', 'verify', 'export'],
					},
					suggestedFix: 'Use format csv, tsv, json, xlsx, or xlsm.',
				},
			})
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('ascend.plan accepts encrypted workbook passwords and commit fails closed before decrypted export', async () => {
		const input = join(
			tmpdir(),
			`ascend-mcp-encrypted-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.out.xlsx`
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			password?: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { preparedPlan?: { id?: string } }
			}
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
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: {
						sourceWasEncrypted?: boolean
						reEncryptionSupported?: boolean
						requestedExport?: string
					}
				}
			}
		}>
		try {
			await Bun.write(input, readFileSync(ENCRYPTED_FIXTURE))
			const planned = await plan({
				file: input,
				password: '123',
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'Z10', value: 'blocked' }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(JSON.stringify(planned)).not.toContain('"123"')

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(committed.isError).toBe(true)
			expect(committed.structuredContent?.error).toMatchObject({
				code: 'EXPORT_ERROR',
				retryable: false,
				details: {
					sourceWasEncrypted: true,
					reEncryptionSupported: false,
					requestedExport: 'xlsx',
				},
			})
			expect(committed.structuredContent?.error?.message).toContain(
				'Cannot export an edited encrypted workbook without re-encryption support',
			)
			expect(JSON.stringify(committed)).not.toContain('"123"')
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('ascend.plan accepts encrypted workbook passwords for path mutations and commit fails closed without source drift', async () => {
		const input = join(
			tmpdir(),
			`ascend-mcp-encrypted-path-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.out.xlsx`
		const sourceBytes = readFileSync(ENCRYPTED_FIXTURE)
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			password?: string
			mutations: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					pathMutations?: { replayable?: boolean; ops?: unknown[] }
					preparedPlan?: { id?: string }
				}
			}
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
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: {
						sourceWasEncrypted?: boolean
						reEncryptionSupported?: boolean
						requestedExport?: string
					}
				}
			}
		}>
		try {
			await Bun.write(input, sourceBytes)
			const planned = await plan({
				file: input,
				password: '123',
				mutations: [{ path: '/sheets/Sheet1/cells/Z10/value', value: 'blocked' }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.pathMutations).toMatchObject({
				replayable: true,
				ops: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [{ ref: 'Z10', value: 'blocked' }],
					},
				],
			})
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(JSON.stringify(planned)).not.toContain('"123"')

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(committed.isError).toBe(true)
			expect(committed.structuredContent?.error).toMatchObject({
				code: 'EXPORT_ERROR',
				retryable: false,
				details: {
					sourceWasEncrypted: true,
					reEncryptionSupported: false,
					requestedExport: 'xlsx',
				},
			})
			expect(committed.structuredContent?.error?.message).toContain(
				'Cannot export an edited encrypted workbook without re-encryption support',
			)
			expect(JSON.stringify(committed)).not.toContain('"123"')
			expect(await Bun.file(output).exists()).toBe(false)
			expect(Buffer.from(readFileSync(input)).equals(sourceBytes)).toBe(true)
			await expect(
				AscendWorkbook.open(new Uint8Array(readFileSync(input)), { password: '123' }),
			).resolves.toBeDefined()
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

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

		expect(names).toContain('ascend.agent_workflow')
		expect(names).toContain('ascend.open_plan')
		expect(names).toContain('ascend.inspect')
		expect(names).toContain('ascend.trust_report')
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
		expect(names).toContain('ascend.formula_assist')
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
		expect(names.length).toBe(33)
	})

	test('ascend.agent_workflow exposes machine-readable safe edit guidance', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing private MCP registry internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.agent_workflow'].handler as (
			args: Record<string, never>,
		) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					workflow?: unknown[]
					tools?: Record<string, string>
					examples?: Record<string, string>
					packageInstallExampleContext?: {
						workdir?: string
						requires?: string[]
						proofOutput?: string[]
					}
					exampleContext?: {
						workdir?: string
						requires?: string[]
						proofCommand?: string
					}
					resources?: string[]
					preparedHandles?: { scope?: string; oneShot?: boolean }
				}
			}
		}>
		const result = await handler({})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.tools).toMatchObject({
			operations: 'ascend.list_operations',
			openPlan: 'ascend.open_plan',
			plan: 'ascend.plan',
			commit: 'ascend.commit',
			check: 'ascend.check',
			lint: 'ascend.lint',
		})
		expect(result.structuredContent?.data?.resources).toContain('ascend://agent-workflow')
		expect(result.structuredContent?.data?.examples).toMatchObject({
			installedCliSafeEdit: 'ascend example-safe-edit <file.xlsx> <out.xlsx>',
			installedSdkSafeEdit: 'node_modules/.bin/ascend-sdk-safe-edit <file.xlsx> <out.xlsx>',
			sdkSafeEdit: 'bun run example:safe-edit <file.xlsx> <out.xlsx>',
			apiSafeEdit: 'bun run example:safe-edit:http <file.xlsx> <out.xlsx>',
			mcpSafeEdit: 'bun run example:safe-edit:mcp <file.xlsx> <out.xlsx>',
		})
		expect(result.structuredContent?.data?.packageInstallExampleContext).toMatchObject({
			workdir: 'consumer-project',
			requires: expect.arrayContaining([
				'@ascend/cli installed for ascend example-safe-edit',
				'@ascend/sdk installed for node_modules/.bin/ascend-sdk-safe-edit',
			]),
			proofOutput: expect.arrayContaining([
				'proofBundle.safeToUse',
				'proofBundle.whatChanged',
				'proofBundle.whySafe',
				'postWrite.dataConnections',
				'postWrite.formulaState',
				'postWrite.visuals',
			]),
		})
		expect(result.structuredContent?.data?.exampleContext).toMatchObject({
			workdir: 'repository-root',
			requires: expect.arrayContaining(['source checkout', 'bun install']),
			proofCommand: 'bun test examples/root-scripts.test.ts',
		})
		expect(result.structuredContent?.data?.preparedHandles).toMatchObject({
			scope: 'process-local',
			oneShot: true,
		})
		expect(result.structuredContent?.data?.workflow).toContainEqual(
			expect.objectContaining({
				step: 'trust-preflight',
				tool: 'ascend.trust_report',
				proof: expect.arrayContaining(['trust', 'posture', 'findingCount']),
			}),
		)
		expect(result.structuredContent?.data?.workflow).toContainEqual(
			expect.objectContaining({
				step: 'plan',
				tool: 'ascend.plan',
				proof: expect.arrayContaining(['inputSha256', 'planDigest', 'preparedPlan']),
			}),
		)
		expect(result.structuredContent?.data?.workflow).toContainEqual(
			expect.objectContaining({
				step: 'reopen-verify',
				tools: expect.arrayContaining(['ascend.check', 'ascend.lint', 'ascend.diff']),
			}),
		)
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
		expect(workflow?.contents[0]?.text).toContain('ascend.open_plan')
		expect(workflow?.contents[0]?.text).toContain('before hydrating unknown XLSX/XLSM')
		expect(workflow?.contents[0]?.text).toContain('ascend.plan')
		expect(workflow?.contents[0]?.text).toContain('ascend.trust_report')
		expect(workflow?.contents[0]?.text).toContain('planHandle')
		expect(workflow?.contents[0]?.text).toContain('must not echo it')
		expect(workflow?.contents[0]?.text).toContain('formula_assist')
		expect(workflow?.contents[0]?.text).toContain('node_modules/.bin/ascend-sdk-safe-edit')
		expect(workflow?.contents[0]?.text).toContain('proofBundle.safeToUse')
		expect(workflow?.contents[0]?.text).toContain('postWrite.dataConnections')
		expect(workflow?.contents[0]?.text).toContain('postWrite.formulaState')
		expect(workflow?.contents[0]?.text).toContain('postWrite.visuals')
		expect(workflow?.contents[0]?.text).toContain('bun run example:safe-edit:mcp')
		expect(workflow?.contents[0]?.text).toContain('bun test examples/root-scripts.test.ts')
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

		const docs = await docsHandler({ query: 'llms plan commit allowLoss', limit: 3 })
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
		expect(text).toContain('ascend.open_plan')
		expect(text).toContain('ascend.plan')
		expect(text).toContain('ascend.commit')
		expect(text).toContain('ascend.trust_report')
		expect(text).toContain('ascend.active_content')
		expect(text).toContain('ascend.formula_assist')
		expect(text).toContain('planHandle')
		expect(text).toContain('allowLoss')
		expect(text).toContain('never echo it')
	})

	test('ascend.formula_assist exposes formula IDE helpers for agents', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.formula_assist'].handler as (args: {
			formula: string
			cursor?: number
			prefix?: string
			completionLimit?: number
			functionName?: string
			reference?: string
			replaceReferenceAtCursor?: boolean
			cycleReference?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					diagnostics?: { parseOk?: boolean }
					activeReference?: { text?: string; kind?: string }
					completions?: Array<{ name?: string }>
					signature?: { name?: string }
					signatureHelp?: { signature?: { name?: string } }
					cycle?: { formula?: string; changed?: boolean }
					insertion?: { formula?: string; replaced?: { text?: string } }
					renameTarget?: { ok?: boolean; reason?: string; role?: { role?: string; text?: string } }
				}
			}
		}>

		const result = await handler({
			formula: '=SUM(A1:B2',
			cursor: 8,
			prefix: 'SU',
			completionLimit: 3,
			functionName: 'SUM',
			reference: 'C1',
			replaceReferenceAtCursor: true,
			cycleReference: true,
		})
		const data = result.structuredContent?.data

		expect(result.structuredContent?.ok).toBe(true)
		expect(data?.diagnostics?.parseOk).toBe(false)
		expect(data?.activeReference).toMatchObject({ text: 'A1:B2', kind: 'range' })
		expect(data?.completions?.some((completion) => completion.name === 'SUM')).toBe(true)
		expect(data?.signature?.name).toBe('SUM')
		expect(data?.signatureHelp?.signature?.name).toBe('SUM')
		expect(data?.cycle).toMatchObject({ formula: '=SUM(A1:$B$2', changed: true })
		expect(data?.insertion).toMatchObject({ formula: '=SUM(C1', replaced: { text: 'A1:B2' } })

		const refusal = await handler({
			formula: '=Budget+Sales[Amount]',
			cursor: 10,
		})
		expect(refusal.structuredContent?.data?.renameTarget).toMatchObject({
			ok: false,
			reason: 'workbook-context-required',
			role: { role: 'table-name-use', text: 'Sales' },
		})
	})

	test('ascend.plan and ascend.commit can include package action proof evidence', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const tools = (server as any)._registeredTools as Record<
			string,
			{
				handler: (args: Record<string, unknown>) => Promise<{
					structuredContent?: { ok?: boolean; data?: Record<string, unknown> }
				}>
			}
		>
		const ops = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'mcp' }] }]

		const plan = await tools['ascend.plan']?.handler({
			file: TEMP_FILE,
			ops,
			includePackageActions: true,
			prepare: false,
		})
		expect(plan?.structuredContent?.ok).toBe(true)
		expect((plan?.structuredContent?.data?.packageActions as { kind?: string })?.kind).toBe(
			'ascend-package-action-proof',
		)

		const commit = await tools['ascend.commit']?.handler({
			file: TEMP_FILE,
			ops,
			output: TEMP_PACKAGE_ACTION_OUTPUT,
			includePackageActions: true,
		})
		expect(commit?.structuredContent?.ok).toBe(true)
		const packageActions = commit?.structuredContent?.data?.packageActions as
			| {
					kind?: string
					byAction?: { regenerate?: number }
					coverage?: {
						sourceByteDigestCount?: number
						outputByteDigestCount?: number
						matchingByteDigestCount?: number
						mismatchedByteDigestCount?: number
					}
					actions?: { outputSha256?: string }[]
			  }
			| undefined
		expect(packageActions?.kind).toBe('ascend-package-action-proof')
		expect(packageActions?.byAction?.regenerate).toBeGreaterThan(0)
		expect(packageActions?.coverage?.sourceByteDigestCount).toBeGreaterThan(0)
		expect(packageActions?.coverage?.outputByteDigestCount).toBeGreaterThan(0)
		expect(
			(packageActions?.coverage?.matchingByteDigestCount ?? 0) +
				(packageActions?.coverage?.mismatchedByteDigestCount ?? 0),
		).toBeGreaterThan(0)
		expect(packageActions?.actions?.some((action) => action.outputSha256 !== undefined)).toBe(true)
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

	test('ascend.trust_report exposes untrusted workbook boundaries for agents', async () => {
		const trustFile = join(
			tmpdir(),
			`ascend-mcp-trust-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
		)
		await Bun.write(trustFile, signedMacroWorkbook())
		try {
			const server = createServer()
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const handler = (server as any)._registeredTools['ascend.trust_report'].handler as (args: {
				file: string
				maxFindings?: number
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						trust?: string
						posture?: string
						includedInAgentContext?: { activeContent?: boolean; hiddenSheets?: boolean }
						executionPolicy?: { macros?: string; externalLinks?: string }
						findings?: Array<{ code?: string; nextAction?: string }>
						nextActions?: readonly string[]
					}
				}
			}>

			const result = await handler({ file: trustFile, maxFindings: 10 })
			const data = result.structuredContent?.data

			expect(result.structuredContent?.ok).toBe(true)
			expect(data?.trust).toBe('untrusted')
			expect(data?.posture).toBe('safe-parser-preserver')
			expect(data?.includedInAgentContext).toMatchObject({
				activeContent: false,
				hiddenSheets: false,
			})
			expect(data?.executionPolicy).toMatchObject({
				macros: 'preserve-only',
				externalLinks: 'do-not-refresh',
			})
			expect(data?.findings).toContainEqual(
				expect.objectContaining({ code: 'workbook.vbaProject' }),
			)
			expect(data?.nextActions).toContain(
				'Use visible workbook data as the default agent context; opt into hidden sheets, comments, names, and metadata only when the task requires them.',
			)
		} finally {
			await unlink(trustFile).catch(() => {})
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

	test('ascend.open_plan recommends values mode for read intent', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.open_plan'].handler as (args: {
			file: string
			intent?: 'read-values'
		}) => Promise<{
			structuredContent?: {
				data?: {
					intent?: string
					recommendedLoadOptions?: { mode?: string; richMetadata?: boolean }
					reviewBeforeHydration?: boolean
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE, intent: 'read-values' })

		expect(result.structuredContent?.data?.intent).toBe('read-values')
		expect(result.structuredContent?.data?.recommendedLoadOptions).toEqual({ mode: 'values' })
		expect(result.structuredContent?.data?.reviewBeforeHydration).toBe(false)
	})

	test('ascend.open_plan routes macro workbooks to metadata review before edit planning', async () => {
		await Bun.write(TEMP_MACRO_FILE, signedMacroWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.open_plan'].handler as (args: {
			file: string
		}) => Promise<{
			structuredContent?: {
				data?: {
					recommendedLoadOptions?: { mode?: string; richMetadata?: boolean }
					reviewBeforeHydration?: boolean
					riskFeatures?: Array<{ featureFamily?: string }>
				}
			}
		}>

		const result = await handler({ file: TEMP_MACRO_FILE })

		expect(result.structuredContent?.data?.recommendedLoadOptions).toEqual({
			mode: 'metadata-only',
		})
		expect(result.structuredContent?.data?.reviewBeforeHydration).toBe(true)
		expect(result.structuredContent?.data?.riskFeatures).toContainEqual(
			expect.objectContaining({ featureFamily: 'preservedMacro' }),
		)
	})

	test('ascend.open_plan accepts encrypted workbook passwords without echoing them', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.open_plan'].handler as (args: {
			file: string
			password?: string
		}) => Promise<{
			structuredContent?: {
				data?: {
					recommendedLoadOptions?: { mode?: string; richMetadata?: boolean }
					partCount?: number
				}
			}
		}>

		const result = await handler({ file: ENCRYPTED_FIXTURE, password: '123' })
		const serialized = JSON.stringify(result)

		expect(result.structuredContent?.data?.recommendedLoadOptions).toEqual({ mode: 'full' })
		expect(result.structuredContent?.data?.partCount).toBeGreaterThan(0)
		expect(serialized).not.toContain('"123"')
	})

	test('ascend.plan accepts encrypted workbook passwords without echoing them', async () => {
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const handler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			password?: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					preview?: { wouldSucceed?: boolean }
					preparedPlan?: { id?: string }
				}
			}
		}>

		const result = await handler({
			file: ENCRYPTED_FIXTURE,
			password: '123',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'planned' }] }],
		})
		const serialized = JSON.stringify(result)

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.preview?.wouldSucceed).toBe(true)
		expect(result.structuredContent?.data?.preparedPlan?.id).toBeString()
		expect(serialized).not.toContain('"123"')
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
				data?: {
					journal?: {
						schemaVersion?: number
						schemaId?: string
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.schemaVersion).toBe(
			MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
		)
		expect(result.structuredContent?.data?.journal?.schemaId).toBe(
			MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
		)
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

	test('ascend.preview and ascend.write return exact empty journals for no-op requests', async () => {
		const wb = AscendWorkbook.create()
		const file = join(
			tmpdir(),
			`ascend-mcp-noop-journal-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		await wb.save(file)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const preview = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			ops?: unknown[]
			mutations?: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { journal?: Record<string, unknown>; pathMutations?: Record<string, unknown> }
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops?: unknown[]
			mutations?: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { journal?: Record<string, unknown>; pathMutations?: Record<string, unknown> }
			}
		}>
		const expectedJournal = {
			schemaVersion: MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
			schemaId: MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
			supported: true,
			exact: true,
			entries: [],
			inverseOps: [],
			issues: [],
			undoPolicy: {
				undoable: true,
				exact: true,
				riskLevel: 'none',
				reason: 'exact',
				userMessage: 'Undo available.',
			},
		}

		try {
			const previewed = await preview({ file, journal: true, ops: [] })
			const written = await write({ file, journal: true, ops: [] })
			const previewedMutations = await preview({ file, journal: true, mutations: [] })
			const writtenMutations = await write({ file, journal: true, mutations: [] })

			expect(previewed.structuredContent?.ok).toBe(true)
			expect(previewed.structuredContent?.data?.journal).toEqual(expectedJournal)
			expect(written.structuredContent?.ok).toBe(true)
			expect(written.structuredContent?.data?.journal).toEqual(expectedJournal)
			expect(previewedMutations.structuredContent?.ok).toBe(true)
			expect(previewedMutations.structuredContent?.data?.journal).toEqual(expectedJournal)
			expect(previewedMutations.structuredContent?.data?.pathMutations).toMatchObject({
				mutationCount: 0,
				issueCount: 0,
				issues: [],
				replayable: true,
			})
			expect(writtenMutations.structuredContent?.ok).toBe(true)
			expect(writtenMutations.structuredContent?.data?.journal).toEqual(expectedJournal)
			expect(writtenMutations.structuredContent?.data?.pathMutations).toMatchObject({
				mutationCount: 0,
				issueCount: 0,
				issues: [],
				replayable: true,
			})
		} finally {
			await unlink(file).catch(() => {})
		}
	})

	test('ascend.preview preserves lossy journal issue metadata', async () => {
		const wb = AscendWorkbook.create()
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
				data?: {
					journal?: {
						schemaVersion?: number
						schemaId?: string
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'groupRows', sheet: 'Sheet1', from: 1, to: 2, collapsed: true },
				{
					op: 'groupCols',
					sheet: 'Sheet1',
					from: 0,
					to: 1,
					collapsed: true,
					summaryRight: false,
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.schemaVersion).toBe(
			MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
		)
		expect(result.structuredContent?.data?.journal?.schemaId).toBe(
			MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
		)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([])
		expect(result.structuredContent?.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped rows for Sheet1 cannot be restored with public operations',
				surface: 'row-layout',
				reason: 'row-layout-created',
				refs: [
					'Sheet1!2',
					'Sheet1!3',
					'Sheet1!4',
					'sheet:Sheet1:outlinePr:summaryBelow',
					'sheet:Sheet1:sheetFormatPr:outlineLevelRow',
				],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped columns for Sheet1 cannot be restored with public operations',
				surface: 'column-layout',
				reason: 'column-layout-created',
				refs: [
					'Sheet1!A',
					'Sheet1!B',
					'sheet:Sheet1:outlinePr:summaryRight',
					'sheet:Sheet1:sheetFormatPr:outlineLevelCol',
				],
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.getWorkbookModel().getSheet('Sheet1')?.rowDefs.size).toBe(0)
		expect(reopened.getWorkbookModel().getSheet('Sheet1')?.colDefs).toEqual([])
	})

	test('ascend.preview and ascend.write preserve the public journal v1 golden issue payload', async () => {
		const file = join(
			tmpdir(),
			`ascend-mcp-journal-v1-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const writeFilePath = `${file}.write.xlsx`
		const previewWorkbook = AscendWorkbook.create()
		const writeWorkbook = AscendWorkbook.create()
		await previewWorkbook.save(file)
		await writeWorkbook.save(writeFilePath)
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const preview = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { journal?: Parameters<typeof compactJournal>[0] }
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { journal?: Parameters<typeof compactJournal>[0] }
			}
		}>

		try {
			const previewed = await preview({
				file,
				journal: true,
				ops: JOURNAL_V1_FIXTURE.scenario.ops,
			})
			const written = await write({
				file: writeFilePath,
				journal: true,
				ops: JOURNAL_V1_FIXTURE.scenario.ops,
			})

			expect(previewed.structuredContent?.ok).toBe(true)
			expect(compactJournal(previewed.structuredContent?.data?.journal ?? {})).toEqual(
				JOURNAL_V1_FIXTURE.scenario.journal,
			)
			expect(written.structuredContent?.ok).toBe(true)
			expect(compactJournal(written.structuredContent?.data?.journal ?? {})).toEqual(
				JOURNAL_V1_FIXTURE.scenario.journal,
			)
		} finally {
			await unlink(file).catch(() => {})
			await unlink(writeFilePath).catch(() => {})
		}
	})

	test('ascend.preview exposes unsupported journal status with partial inverse ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])
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
				data?: {
					journal?: {
						schemaVersion?: number
						schemaId?: string
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'appendRows', table: 'Sales', rows: [['East', 2]] },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 'audit' }] },
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(false)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([
			{ op: 'clearRange', sheet: 'Sheet1', range: 'D1', what: 'all' },
		])
		expect(result.structuredContent?.data?.journal?.issues).toContainEqual({
			code: 'UNSUPPORTED_OPERATION',
			message: 'No reversible journal support for appendRows',
			reason: 'operation-unsupported',
			surface: 'tables',
		})
	})

	test('ascend.write preserves lossy journal issue metadata while saving', async () => {
		const wb = AscendWorkbook.create()
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'groupRows', sheet: 'Sheet1', from: 1, to: 2, collapsed: true },
				{
					op: 'groupCols',
					sheet: 'Sheet1',
					from: 0,
					to: 1,
					collapsed: true,
					summaryRight: false,
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.schemaVersion).toBe(
			MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
		)
		expect(result.structuredContent?.data?.journal?.schemaId).toBe(
			MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
		)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([])
		expect(result.structuredContent?.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped rows for Sheet1 cannot be restored with public operations',
				surface: 'row-layout',
				reason: 'row-layout-created',
				refs: [
					'Sheet1!2',
					'Sheet1!3',
					'Sheet1!4',
					'sheet:Sheet1:outlinePr:summaryBelow',
					'sheet:Sheet1:sheetFormatPr:outlineLevelRow',
				],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped columns for Sheet1 cannot be restored with public operations',
				surface: 'column-layout',
				reason: 'column-layout-created',
				refs: [
					'Sheet1!A',
					'Sheet1!B',
					'sheet:Sheet1:outlinePr:summaryRight',
					'sheet:Sheet1:sheetFormatPr:outlineLevelCol',
				],
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		const sheet = reopened.getWorkbookModel().getSheet('Sheet1')
		expect(sheet?.rowDefs.get(1)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(sheet?.rowDefs.get(2)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(sheet?.rowDefs.get(3)).toEqual({ collapsed: true })
		expect(sheet?.colDefs).toContainEqual({ min: 0, max: 0, hidden: true, outlineLevel: 1 })
		expect(sheet?.colDefs).toContainEqual({ min: 1, max: 1, hidden: true, outlineLevel: 1 })
	})

	test('ascend.write exposes unsupported journal status with partial inverse ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'appendRows', table: 'Sales', rows: [['East', 2]] },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 'audit' }] },
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(false)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([
			{ op: 'clearRange', sheet: 'Sheet1', range: 'D1', what: 'all' },
		])
		expect(result.structuredContent?.data?.journal?.issues).toContainEqual({
			code: 'UNSUPPORTED_OPERATION',
			message: 'No reversible journal support for appendRows',
			reason: 'operation-unsupported',
			surface: 'tables',
		})
		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('D1')?.value).toEqual({
			kind: 'string',
			value: 'audit',
		})
	})

	test('ascend.preview marks new theme additions lossy with public inverse metadata', async () => {
		const wb = AscendWorkbook.create()
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setTheme',
					themeName: 'New Brand',
					themeColors: [{ slot: 'accent1', rgb: '123456' }],
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([])
		expect(result.structuredContent?.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Theme metadata field themeName cannot be removed with public operations',
				surface: 'package-parts',
				reason: 'package-part-preservation',
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Theme color slot accent1 cannot be removed with public operations',
				surface: 'package-parts',
				reason: 'package-part-preservation',
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.inspect().themeSummary.hasThemePart).toBe(false)
	})

	test('ascend.preview marks saved defined-name edits lossy for package-part proof', async () => {
		const wb = AscendWorkbook.create()
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$B$1' }],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([
			{ op: 'deleteDefinedName', name: 'Budget' },
		])
		expect(result.structuredContent?.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message:
					'setDefinedName changes saved package state that public inverse operations cannot restore byte-for-byte',
				surface: 'package-parts',
				reason: 'package-part-preservation',
				refs: ['name:Budget'],
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.definedName('Budget')).toBeUndefined()
	})

	test('ascend.write lossy journal inverse ops restore saved workbook truth after reopen', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Status' },
					{ ref: 'A2', value: 'Open' },
					{ ref: 'A3', value: 'Closed' },
					{ ref: 'B1', value: 2 },
					{ ref: 'C1', value: 3.14 },
					{ ref: 'J1', value: 'Product' },
					{ ref: 'K1', value: 'Qty' },
					{ ref: 'J2', value: 'Widget' },
					{ ref: 'K2', value: 5 },
					{ ref: 'J3', value: 'Bolt' },
					{ ref: 'K3', value: 6 },
					{ ref: 'L1', value: 'currency-style-anchor' },
					{ ref: 'L2', value: 'decimal-style-anchor' },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B1:B1', style: { numberFormat: '0.00' } },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.0' },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'L1:L1', style: { numberFormat: '$0.00' } },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'L2:L2', format: '0.000' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'F1:F3',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'G1:G3',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '5', priority: 1 },
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'J1:K3', name: 'Sales', hasHeaders: true },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:A3', column: 0, values: ['Open'] },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://example.com' },
			{ op: 'setWorkbookView', index: 0, view: { activeTab: 0, firstSheet: 0, tabRatio: 600 } },
			{
				op: 'setWorkbookProtection',
				protection: { lockStructure: true, workbookPassword: 'ABCD' },
			},
			{
				op: 'setSheetProtection',
				sheet: 'Sheet1',
				password: 'ABCD',
				options: { formatCells: false, autoFilter: true },
			},
			{ op: 'setTabColor', sheet: 'Sheet1', color: 'FF0000' },
			{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 },
			{
				op: 'setDocumentProperties',
				mode: 'replace',
				properties: {
					core: { title: 'Before' },
					app: { company: 'Ascend' },
					custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
				},
			},
		])
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 7 }] },
				{ op: 'setStyle', sheet: 'Sheet1', range: 'B1:B1', style: { numberFormat: '$0.00' } },
				{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.000' },
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'F1:F3',
					rule: { type: 'whole', operator: 'greaterThan', formula1: '10' },
				},
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'G1:G3',
					rule: { type: 'cellIs', operator: 'greaterThan', formula: '7', priority: 1 },
				},
				{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:A3', column: 0, values: ['Closed'] },
				{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
				{ op: 'setTableColumn', table: 'Revenue', column: 'Qty', newName: 'Units' },
				{ op: 'setTableStyle', table: 'Revenue', styleName: null },
				{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://changed.example' },
				{
					op: 'setWorkbookView',
					index: 0,
					mode: 'replace',
					view: { activeTab: 0, firstSheet: 0, tabRatio: 720 },
				},
				{
					op: 'setWorkbookProtection',
					protection: { lockWindows: true, workbookPassword: 'DCBA' },
				},
				{
					op: 'setSheetProtection',
					sheet: 'Sheet1',
					password: 'DCBA',
					options: { insertRows: true, deleteRows: false },
				},
				{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' },
				{ op: 'freezePane', sheet: 'Sheet1', row: 2, col: 0 },
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.issues).toContainEqual(
			expect.objectContaining({
				surface: 'package-parts',
				reason: 'package-part-preservation',
			}),
		)
		const inverse = parseOperations(result.structuredContent?.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		expect(changed.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 7 })
		expect(changed.cellStyle('Sheet1!B1')?.numberFormat).toBe('$0.00')
		expect(changed.cellStyle('Sheet1!C1')?.numberFormat).toBe('0.000')
		expect(changed.sheet('Sheet1')?.dataValidations[0]).toMatchObject({
			sqref: 'F1:F3',
			formula1: '10',
		})
		expect(changed.sheet('Sheet1')?.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['7'])
		expect(changed.sheet('Sheet1')?.autoFilter).toMatchObject({
			ref: 'A1:A3',
			columns: [{ colId: 0, kind: 'filters', values: ['Closed'] }],
		})
		expect(changed.table('Sales')).toBeUndefined()
		expect(changed.table('Revenue')?.columns).toEqual(['Product', 'Units'])
		expect(changed.table('Revenue')?.columnDefs[1]?.formula).toBeUndefined()
		expect(changed.table('Revenue')?.styleInfo).toBeUndefined()
		expect(changed.inspectSheet('Sheet1')?.hyperlinks?.[0]?.target).toBe('https://changed.example')
		expect(changed.workbookViews()[0]).toMatchObject({ activeTab: 0, firstSheet: 0, tabRatio: 720 })
		expect(changed.getWorkbookModel().workbookProtection).toMatchObject({
			lockWindows: true,
			workbookPassword: 'DCBA',
		})
		expect(changed.sheet('Sheet1')?.protection).toMatchObject({
			password: 'DCBA',
			insertRows: true,
			deleteRows: false,
		})
		expect(changed.sheet('Sheet1')?.tabColor).toEqual({ rgb: '00FF00' })
		expect(changed.sheet('Sheet1')?.frozenRows).toBe(2)
		expect(changed.sheet('Sheet1')?.frozenCols).toBe(0)
		expect(changed.inspect().documentProperties).toMatchObject({
			core: { title: 'Before' },
			app: { company: 'Ascend' },
			custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
		})

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		expect(restored.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(restored.cellStyle('Sheet1!B1')?.numberFormat).toBe('0.00')
		expect(restored.cellStyle('Sheet1!C1')?.numberFormat).toBe('0.0')
		expect(restored.sheet('Sheet1')?.dataValidations[0]).toMatchObject({
			sqref: 'F1:F3',
			formula1: '0',
			allowBlank: true,
		})
		expect(restored.sheet('Sheet1')?.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['5'])
		expect(restored.sheet('Sheet1')?.autoFilter).toMatchObject({
			ref: 'A1:A3',
			columns: [{ colId: 0, kind: 'filters', values: ['Open'] }],
		})
		expect(restored.table('Revenue')).toBeUndefined()
		expect(restored.table('Sales')?.columns).toEqual(['Product', 'Qty'])
		expect(restored.table('Sales')?.columnDefs[1]?.formula).toBeUndefined()
		expect(restored.table('Sales')?.styleInfo).toEqual({ name: 'TableStyleMedium2' })
		expect(restored.inspectSheet('Sheet1')?.hyperlinks?.[0]?.target).toBe('https://example.com')
		expect(restored.workbookViews()[0]).toMatchObject({
			activeTab: 0,
			firstSheet: 0,
			tabRatio: 600,
		})
		expect(restored.getWorkbookModel().workbookProtection).toMatchObject({
			lockStructure: true,
			workbookPassword: 'ABCD',
		})
		expect(restored.sheet('Sheet1')?.protection).toMatchObject({
			password: 'ABCD',
			formatCells: false,
			autoFilter: true,
		})
		expect(restored.sheet('Sheet1')?.tabColor).toEqual({ rgb: 'FF0000' })
		expect(restored.sheet('Sheet1')?.frozenRows).toBe(1)
		expect(restored.sheet('Sheet1')?.frozenCols).toBe(1)
		expect(restored.inspect().documentProperties).toMatchObject({
			core: { title: 'Before' },
			app: { company: 'Ascend' },
			custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
		})
		expect(restored.check().valid).toBe(true)
	})

	test('ascend.write exact theme journal inverse ops restore saved theme truth after reopen', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setTheme',
				themeName: 'Office',
				colorSchemeName: 'Office Colors',
				majorFontLatin: 'Aptos Display',
				minorFontLatin: 'Aptos',
				themeColors: [
					{ slot: 'accent1', rgb: '4F81BD' },
					{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
				],
			},
		])
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setTheme',
					themeName: 'Brand',
					colorSchemeName: 'Brand Colors',
					majorFontLatin: 'Inter Display',
					minorFontLatin: 'Inter',
					themeColors: [
						{ slot: 'accent1', rgb: '0F6CBD' },
						{ slot: 'lt1', systemColor: 'windowText', lastColor: '000000' },
					],
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(true)
		expect(result.structuredContent?.data?.journal?.issues).toEqual([])
		const inverse = parseOperations(result.structuredContent?.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact theme journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		expect(changed.inspect().themeSummary).toMatchObject({
			hasThemePart: true,
			name: 'Brand',
			colorSchemeName: 'Brand Colors',
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		})
		expect(changed.inspect().themeSummary.colors.find((color) => color.slot === 'accent1')).toEqual(
			{
				slot: 'accent1',
				rgb: '0F6CBD',
			},
		)
		expect(changed.inspect().themeSummary.colors.find((color) => color.slot === 'lt1')).toEqual({
			slot: 'lt1',
			systemColor: 'windowText',
			lastColor: '000000',
		})

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		expect(restored.inspect().themeSummary).toMatchObject({
			hasThemePart: true,
			name: 'Office',
			colorSchemeName: 'Office Colors',
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		})
		expect(
			restored.inspect().themeSummary.colors.find((color) => color.slot === 'accent1'),
		).toEqual({
			slot: 'accent1',
			rgb: '4F81BD',
		})
		expect(restored.inspect().themeSummary.colors.find((color) => color.slot === 'lt1')).toEqual({
			slot: 'lt1',
			systemColor: 'window',
			lastColor: 'FFFFFF',
		})
		expect(restored.check().valid).toBe(true)
	})

	test('ascend.write exact chart journal inverse ops restore saved chart truth after reopen', async () => {
		await Bun.write(TEMP_FILE, Bun.file(CHARTSHEET_FIXTURE))

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$6',
					valueRef: 'Sheet1!$B$2:$B$6',
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(true)
		expect(result.structuredContent?.data?.journal?.issues).toEqual([])
		const inverse = parseOperations(result.structuredContent?.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact chart journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		const changedChart = changed.getWorkbookModel().chartParts[0]
		expect(changedChart).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Chart1',
			chartType: 'barChart',
			title: 'Wildlife Population',
		})
		expect(changedChart?.series).toHaveLength(3)
		expect(changedChart?.series[0]).toMatchObject({
			nameRef: 'Sheet1!$B$1',
			nameText: 'Bears',
			categoryRef: 'Sheet1!$A$2:$A$6',
			valueRef: 'Sheet1!$B$2:$B$6',
		})
		expect(changedChart?.series[1]).toMatchObject({
			nameRef: 'Sheet1!$C$1',
			nameText: 'Dolphins',
			categoryRef: 'Sheet1!$A$2:$A$7',
			valueRef: 'Sheet1!$C$2:$C$7',
		})

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		const restoredChart = restored.getWorkbookModel().chartParts[0]
		expect(restoredChart).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Chart1',
			chartType: 'barChart',
			title: 'Wildlife Population',
		})
		expect(restoredChart?.series).toHaveLength(3)
		expect(restoredChart?.series[0]).toMatchObject({
			nameRef: 'Sheet1!$B$1',
			nameText: 'Bears',
			categoryRef: 'Sheet1!$A$2:$A$7',
			valueRef: 'Sheet1!$B$2:$B$7',
		})
		expect(restoredChart?.series[1]).toMatchObject({
			nameRef: 'Sheet1!$C$1',
			nameText: 'Dolphins',
			categoryRef: 'Sheet1!$A$2:$A$7',
			valueRef: 'Sheet1!$C$2:$C$7',
		})
		expect(restored.check().valid).toBe(true)
	})

	test('ascend.write exact pivot journal inverse ops restore saved pivot cache truth after reopen', async () => {
		await Bun.write(TEMP_FILE, Bun.file(PIVOT_FIXTURE))

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setPivotCache',
					pivotTable: 'PivotTable1',
					sourceSheet: 'Sheet1',
					sourceRef: 'A1:K4',
				},
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(true)
		expect(result.structuredContent?.data?.journal?.issues).toEqual([])
		const inverse = parseOperations(result.structuredContent?.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact pivot journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		const changedCache = changed.getWorkbookModel().pivotCaches[0]
		expect(changedCache).toMatchObject({
			cacheId: 37,
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			recordCount: 4,
			sourceSheet: 'Sheet1',
			sourceRef: 'A1:K4',
		})
		expect(changedCache?.fields).toHaveLength(11)
		expect(changedCache?.records?.parsedCount).toBe(4)

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		const restoredCache = restored.getWorkbookModel().pivotCaches[0]
		expect(restoredCache).toMatchObject({
			cacheId: 37,
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			recordCount: 4,
			sourceSheet: 'Sheet1',
			sourceRef: 'A1:K5',
		})
		expect(restoredCache?.fields).toHaveLength(11)
		expect(restoredCache?.records?.parsedCount).toBe(4)
		expect(restored.check().valid).toBe(true)
	})

	test('ascend.preview marks pivot cache public rollback gaps as lossy', async () => {
		await Bun.write(TEMP_FILE, Bun.file(PIVOT_FIXTURE))

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.preview'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
						issues?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setPivotCache', pivotTable: 'PivotTable1', refreshOnLoad: true }],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.inverseOps).toEqual([])
		expect(result.structuredContent?.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Pivot cache selector cannot be restored exactly',
				surface: 'pivot-caches',
				reason: 'pivot-cache-unsettable',
			},
		])
		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.getWorkbookModel().pivotCaches[0]?.refreshOnLoad).toBeUndefined()
	})

	test('ascend.write lossy journal inverse ops restore recalculated workbook truth after reopen', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'D1', value: 'Item' },
					{ ref: 'E1', value: 'Qty' },
					{ ref: 'F1', value: 'Calc' },
					{ ref: 'D2', value: 'A' },
					{ ref: 'E2', value: 5 },
					{ ref: 'D3', value: 'B' },
					{ ref: 'E3', value: 6 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*3' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'D1:F3', name: 'CalcTable', hasHeaders: true },
			{ op: 'setTableColumn', table: 'CalcTable', column: 'Calc', formula: '[@Qty]*2' },
		])
		expect(wb.recalc().errors).toEqual([])
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
				data?: {
					journal?: {
						supported?: boolean
						exact?: boolean
						inverseOps?: unknown[]
					}
				}
			}
		}>
		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
				{ op: 'setTableColumn', table: 'CalcTable', column: 'Calc', formula: '[@Qty]*3' },
			],
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.journal?.supported).toBe(true)
		expect(result.structuredContent?.data?.journal?.exact).toBe(false)
		expect(result.structuredContent?.data?.journal?.issues).toContainEqual(
			expect.objectContaining({
				surface: 'package-parts',
				reason: 'package-part-preservation',
			}),
		)
		const inverse = parseOperations(result.structuredContent?.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		expect(changed.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 5 })
		expect(changed.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 15 })
		expect(changed.table('CalcTable')?.columnDefs[2]?.formula).toBe('[@Qty]*3')
		expect(changed.sheet('Sheet1')?.cell('F2')?.value).toEqual({ kind: 'number', value: 15 })
		expect(changed.sheet('Sheet1')?.cell('F3')?.value).toEqual({ kind: 'number', value: 18 })

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		expect(changed.recalc().errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		expect(restored.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(restored.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 6 })
		expect(restored.table('CalcTable')?.columnDefs[2]?.formula).toBe('[@Qty]*2')
		expect(restored.sheet('Sheet1')?.cell('F2')?.value).toEqual({ kind: 'number', value: 10 })
		expect(restored.sheet('Sheet1')?.cell('F3')?.value).toEqual({ kind: 'number', value: 12 })
		expect(restored.check().valid).toBe(true)
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
					preparedPlan?: { id?: string }
					preview?: {
						changedCellCount?: number
						emittedChangedCellCount?: number
						changedCells?: unknown[]
						changedRanges?: readonly { readonly sheet?: string; readonly range?: string }[]
						wouldSucceed?: boolean
						journalSummary?: {
							supported?: boolean
							exact?: boolean
							inverseOpCount?: number
							issueCount?: number
							issues?: unknown[]
						}
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
		expect(result.structuredContent?.data?.preview?.changedRanges).toEqual([
			{ sheet: 'Sheet1', range: 'A1:A3' },
		])
	})

	test('compact prepared MCP plan and commit preserve journal v1 issue compatibility', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${TEMP_FILE}.compact-journal-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
						preview?: {
							journalSummary?: {
								schemaVersion?: number
								schemaId?: string
								supported?: boolean
								exact?: boolean
								inverseOpCount?: number
								issueCount?: number
								issues?: unknown[]
							}
						}
					}
				}
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
						apply?: {
							journalSummary?: {
								schemaVersion?: number
								schemaId?: string
								supported?: boolean
								exact?: boolean
								inverseOpCount?: number
								issueCount?: number
								issues?: unknown[]
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: TEMP_FILE,
				compact: true,
				ops: JOURNAL_V1_FIXTURE.scenario.ops,
			})

			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()
			expect(planned.structuredContent?.data?.preview?.journalSummary).toEqual(
				JOURNAL_V1_FIXTURE.scenario.journal,
			)

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.apply?.journalSummary).toEqual(
				JOURNAL_V1_FIXTURE.scenario.journal,
			)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commits expose bounded affected refs and ranges', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${TEMP_FILE}.compact-affected-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				file: string
				output?: string
				compact?: boolean
				maxAffectedCells?: number
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						apply?: {
							affectedCellCount?: number
							emittedAffectedCellCount?: number
							affectedCellRefs?: readonly string[]
							affectedRanges?: readonly { readonly sheet?: string; readonly range?: string }[]
						}
						postWrite?: { auditsPassed?: boolean }
					}
				}
			}>

			const committed = await commit({
				file: TEMP_FILE,
				output,
				compact: true,
				maxAffectedCells: 2,
				ops: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 1 },
							{ ref: 'A2', value: 2 },
							{ ref: 'A3', value: 3 },
						],
					},
				],
			})

			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.apply?.affectedCellCount).toBe(3)
			expect(committed.structuredContent?.data?.apply?.emittedAffectedCellCount).toBe(2)
			expect(committed.structuredContent?.data?.apply?.affectedCellRefs).toEqual(['A1', 'A2'])
			expect(committed.structuredContent?.data?.apply?.affectedRanges).toEqual([
				{ sheet: 'Sheet1', range: 'A1:A3' },
			])
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP plan and commit expose preservation mode summaries', async () => {
		await Bun.write(TEMP_FILE, preservedCustomWorkbook())
		const output = `${TEMP_FILE}.compact-preservation-modes-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
						approvals?: readonly { readonly id: string }[]
						writePolicy?: {
							summary?: {
								preservationModes?: {
									generatedWithOpaquePayloads?: number
									reviewRequiredParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							opaquePayloads?: {
								generatedWithOpaquePayloads?: number
								x14ConditionalFormatExtensionPayloads?: number
								x14DataValidationExtensionPayloads?: number
								worksheetParts?: readonly string[]
								preservationMode?: string
								verification?: string
							}
						}
						writePolicy?: {
							summary?: {
								preservationModes?: {
									generatedWithOpaquePayloads?: number
									reviewRequiredParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: TEMP_FILE,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(
				planned.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				reviewRequiredParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedother:preserved:/)])

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(
				committed.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				reviewRequiredParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP plan and commit expose generated opaque preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-opaque-preservation-modes-input.xlsx`
		await writeOpaqueX14Workbook(input)
		const output = `${TEMP_FILE}.compact-opaque-preservation-modes-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
						writePolicy?: {
							summary?: {
								preservationModes?: {
									generatedWithOpaquePayloads?: number
									reviewRequiredParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
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
						postWrite?: { auditsPassed?: boolean }
						writePolicy?: {
							summary?: {
								preservationModes?: {
									generatedWithOpaquePayloads?: number
									reviewRequiredParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(
				planned.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				generatedWithOpaquePayloads: 2,
				reviewRequiredParts: 0,
				lossyApprovalRequiredFeatures: 0,
			})

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(
				committed.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				generatedWithOpaquePayloads: 2,
				reviewRequiredParts: 0,
				lossyApprovalRequiredFeatures: 0,
			})
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.opaquePayloads).toMatchObject({
				generatedWithOpaquePayloads: 2,
				x14ConditionalFormatExtensionPayloads: 1,
				x14DataValidationExtensionPayloads: 1,
				worksheetParts: ['xl/worksheets/sheet1.xml'],
				preservationMode: 'generated-with-opaque-payload',
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP plan and commit expose inspect-only preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-inspect-only-preservation-modes-input.xlsx`
		await Bun.write(input, inspectOnlyWorkbook())
		const output = `${TEMP_FILE}.compact-inspect-only-preservation-modes-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
						approvals?: readonly { readonly id: string }[]
						writePolicy?: {
							diagnostics?: readonly {
								readonly code?: string
								readonly featureFamily?: string
								readonly preservationMode?: string
								readonly packageParts?: readonly {
									readonly partPath?: string
									readonly preservationPolicy?: string
									readonly preservationMode?: string
								}[]
							}[]
							summary?: {
								preservationModes?: {
									inspectOnlyParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: { auditsPassed?: boolean }
						writePolicy?: {
							summary?: {
								preservationModes?: {
									inspectOnlyParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(
				planned.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				inspectOnlyParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(planned.structuredContent?.data?.writePolicy?.diagnostics).toContainEqual(
				expect.objectContaining({
					code: 'approval-required-feature',
					featureFamily: 'preservedPowerQuery',
					preservationMode: 'inspect-only',
					packageParts: [
						expect.objectContaining({
							partPath: 'xl/customData/item1.data',
							preservationPolicy: 'inspect-only',
							preservationMode: 'inspect-only',
						}),
					],
				}),
			)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedpowerquery:preserved:/)])

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(
				committed.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				inspectOnlyParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened legacy comment summary', async () => {
		const input = `${TEMP_FILE}.compact-comment-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-comment-summary-mcp.xlsx`
		const workbook = AscendWorkbook.create()
		workbook.getWorkbookModel().sheets[0]?.comments.set('B2', {
			text: 'Review this',
			author: 'Ada',
		})
		await workbook.save(input)
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							comments?: {
								legacyCommentLocations?: number
								threadedCommentLocations?: number
								legacyDrawingLocations?: number
								locations?: readonly string[]
								threadedCommentPartPaths?: readonly string[]
								verification?: string
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.comments).toMatchObject({
				legacyCommentLocations: 1,
				threadedCommentLocations: 0,
				legacyDrawingLocations: 1,
				locations: ['Sheet1!B2'],
				threadedCommentPartPaths: [],
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened table summary', async () => {
		const input = `${TEMP_FILE}.compact-table-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-table-summary-mcp.xlsx`
		const workbook = AscendWorkbook.create()
		expect(
			workbook.apply([
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 'Qty' },
						{ ref: 'B1', value: 'Price' },
						{ ref: 'A2', value: 2 },
						{ ref: 'B2', value: 5 },
						{ ref: 'A3', value: 3 },
						{ ref: 'B3', value: 7 },
					],
				},
				{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Sales', hasHeaders: true },
			]).errors,
		).toHaveLength(0)
		await workbook.save(input)
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							tables?: {
								tableLocations?: number
								queryTableLocations?: number
								tableAutoFilterLocations?: number
								tableNames?: readonly string[]
								locations?: readonly string[]
								tablePartPaths?: readonly string[]
								queryTablePartPaths?: readonly string[]
								preservationMode?: string
								verification?: string
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.tables).toMatchObject({
				tableLocations: 1,
				queryTableLocations: 0,
				tableAutoFilterLocations: 1,
				tableNames: ['Sales'],
				locations: ['Sheet1!A1:B3'],
				tablePartPaths: ['xl/tables/table1.xml'],
				queryTablePartPaths: [],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened defined name summary', async () => {
		const input = `${TEMP_FILE}.compact-defined-name-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-defined-name-summary-mcp.xlsx`
		const workbook = AscendWorkbook.create()
		const sheet = workbook.getWorkbookModel().sheets[0]
		workbook.getWorkbookModel().definedNames.set('GlobalRate', 'Sheet1!$A$1')
		if (sheet) {
			workbook
				.getWorkbookModel()
				.definedNames.set(
					'LocalRate',
					'Sheet1!$B$1',
					{ kind: 'sheet', sheetId: sheet.id },
					{ hidden: true },
				)
		}
		await workbook.save(input)
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
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
						postWrite?: {
							auditsPassed?: boolean
							definedNames?: {
								total?: number
								workbookScoped?: number
								sheetScoped?: number
								hidden?: number
								names?: Array<{
									name?: string
									formula?: string
									scope?: string
									sheet?: string
									hidden?: boolean
								}>
								verification?: string
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.definedNames).toMatchObject({
				total: 2,
				workbookScoped: 1,
				sheetScoped: 1,
				hidden: 1,
				names: [
					{ name: 'GlobalRate', formula: 'Sheet1!$A$1', scope: 'workbook' },
					{
						name: 'LocalRate',
						formula: 'Sheet1!$B$1',
						scope: 'sheet',
						sheet: 'Sheet1',
						hidden: true,
					},
				],
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened workbook and sheet security summary', async () => {
		const input = `${TEMP_FILE}.compact-security-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-security-summary-mcp.xlsx`
		const workbook = AscendWorkbook.create()
		const model = workbook.getWorkbookModel()
		model.workbookProtection = {
			lockStructure: true,
			lockWindows: true,
			workbookPassword: 'ABCD',
		}
		const sheet = model.getSheet('Sheet1')
		if (!sheet) throw new Error('Expected Sheet1')
		sheet.protection = {
			sheet: true,
			password: 'DCBA',
			autoFilter: true,
			sort: true,
		}
		sheet.protectedRanges = [{ name: 'Editable', sqref: 'C:C', password: '1234' }]
		await workbook.save(input)
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
					}
				}
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
						postWrite?: {
							auditsPassed?: boolean
							security?: {
								workbookProtected?: boolean
								workbookLocks?: readonly string[]
								workbookPasswordProtected?: boolean
								workbookRevisionPasswordProtected?: boolean
								protectedSheets?: number
								protectedSheetNames?: readonly string[]
								sheetPasswordProtected?: number
								sheetStrongHashProtected?: number
								protectedRanges?: number
								protectedRangeLocations?: readonly string[]
								passwordHashVerification?: string
								preservationMode?: string
								verification?: string
								sheets?: Array<{
									sheetName?: string
									protected?: boolean
									passwordProtected?: boolean
									allowedActions?: readonly string[]
									protectedRanges?: number
									protectedRangeLocations?: readonly string[]
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.security).toMatchObject({
				workbookProtected: true,
				workbookLocks: ['lockStructure', 'lockWindows'],
				workbookPasswordProtected: true,
				workbookRevisionPasswordProtected: false,
				protectedSheets: 1,
				protectedSheetNames: ['Sheet1'],
				sheetPasswordProtected: 1,
				sheetStrongHashProtected: 0,
				protectedRanges: 1,
				protectedRangeLocations: ['Sheet1!C:C'],
				passwordHashVerification: 'reported-not-validated',
				preservationMode: 'generated',
				verification: 'reopened-output',
				sheets: [
					expect.objectContaining({
						sheetName: 'Sheet1',
						protected: true,
						passwordProtected: true,
						allowedActions: ['sort', 'autoFilter'],
						protectedRanges: 1,
						protectedRangeLocations: ['Sheet1!C:C'],
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened external reference binding summary', async () => {
		const input = `${TEMP_FILE}.compact-external-reference-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-external-reference-summary-mcp.xlsx`
		await Bun.write(input, externalLinkBoundWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							externalReferences?: {
								total?: number
								boundByExternalBookRelId?: number
								fallbackPathRelationships?: number
								missingPathRelationships?: number
								partPaths?: readonly string[]
								targets?: readonly string[]
								parts?: Array<{
									partPath?: string
									relId?: string
									externalBookRelId?: string
									linkRelId?: string
									linkBindingStatus?: string
									target?: string
									targetMode?: string
								}>
								preservationMode?: string
								verification?: string
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.externalReferences).toMatchObject({
				total: 1,
				boundByExternalBookRelId: 1,
				fallbackPathRelationships: 0,
				missingPathRelationships: 0,
				partPaths: ['xl/externalLinks/externalLink1.xml'],
				targets: ['../sources/source.xlsx'],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				parts: [
					expect.objectContaining({
						partPath: 'xl/externalLinks/externalLink1.xml',
						relId: 'rIdExternal',
						externalBookRelId: 'rIdExt',
						linkRelId: 'rIdExt',
						linkBindingStatus: 'externalBookRelId',
						target: '../sources/source.xlsx',
						targetMode: 'External',
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP plan exposes workbook-qualified 3D external references as sheet-span risk', async () => {
		const input = `${TEMP_FILE}.compact-external-3d-plan-input.xlsx`
		await Bun.write(input, externalLinkBoundWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						writePolicy?: {
							diagnostics?: readonly {
								code?: string
								details?: {
									relatedOperations?: readonly unknown[]
									externalLinks?: readonly unknown[]
								}
							}[]
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [
					{
						op: 'setFormula',
						sheet: 'Sheet1',
						ref: 'B2',
						formula: '=SUM([1]FY26:FY28!B2:B10)',
					},
				],
			})

			expect(planned.structuredContent?.ok).toBe(true)
			const dependency = planned.structuredContent?.data?.writePolicy?.diagnostics?.find(
				(diagnostic) => diagnostic.code === 'external-link-dependency',
			)
			expect(dependency?.details?.relatedOperations).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						operationIndex: 0,
						op: 'setFormula',
						sourceKind: 'cellFormula',
						sourceRef: 'Sheet1!B2',
						formula: '=SUM([1]FY26:FY28!B2:B10)',
						workbook: '1',
						sheetSpan: { startSheet: 'FY26', endSheet: 'FY28' },
						references: ["'[1]FY26:FY28'!B2:B10"],
					}),
				]),
			)
			expect(dependency?.details?.externalLinks).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						workbook: '1',
						sheetSpans: [{ startSheet: 'FY26', endSheet: 'FY28' }],
					}),
				]),
			)
		} finally {
			await unlink(input).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened analytics refresh summary', async () => {
		const input = `${TEMP_FILE}.compact-analytics-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-analytics-summary-mcp.xlsx`
		await Bun.write(input, analyticsRefreshWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							unresolvedPackageGraphIssueCount?: number
							analytics?: {
								pivotCaches?: number
								pivotTables?: number
								slicerCaches?: number
								slicers?: number
								timelineCaches?: number
								timelines?: number
								partPaths?: readonly string[]
								requiresExternalRefresh?: boolean
								preservationMode?: string
								verification?: string
								pivotCacheDetails?: Array<{
									partPath?: string
									cacheId?: number
									sourceSheet?: string
									sourceRef?: string
									outputState?: string
									requiresExternalRefresh?: boolean
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'PivotSheet', updates: [{ ref: 'A1', value: 'ok' }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.unresolvedPackageGraphIssueCount).toBe(0)
			expect(committed.structuredContent?.data?.postWrite?.analytics).toMatchObject({
				pivotCaches: 1,
				pivotTables: 1,
				slicerCaches: 1,
				slicers: 1,
				timelineCaches: 1,
				timelines: 1,
				partPaths: expect.arrayContaining([
					'xl/pivotCache/pivotCacheDefinition1.xml',
					'xl/pivotTables/pivotTable1.xml',
					'xl/slicerCaches/slicerCache1.xml',
					'xl/timelineCaches/timelineCache1.xml',
				]),
				requiresExternalRefresh: true,
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				pivotCacheDetails: [
					expect.objectContaining({
						partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
						cacheId: 34,
						sourceSheet: 'Raw',
						sourceRef: 'A1:B3',
						outputState: 'refresh-on-open',
						requiresExternalRefresh: true,
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened active content summary', async () => {
		const input = `${TEMP_FILE}.compact-active-content-input.xlsm`
		const output = `${TEMP_FILE}.compact-active-content-mcp.xlsm`
		await Bun.write(input, signedMacroWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							activeContent?: {
								total?: number
								vbaProjects?: number
								activeXControls?: number
								formControls?: number
								macroSheets?: number
								vbaSignatures?: number
								digitalSignatures?: number
								customUi?: number
								unknownActiveContent?: number
								partPaths?: readonly string[]
								executionPolicy?: string
								preservationMode?: string
								verification?: string
								entries?: Array<{
									kind?: string
									partPath?: string
									contentType?: string
									anchor?: string
									opaque?: boolean
									executionPolicy?: string
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.activeContent).toMatchObject({
				total: 2,
				vbaProjects: 1,
				activeXControls: 0,
				vbaSignatures: 1,
				digitalSignatures: 0,
				partPaths: ['xl/vbaProject.bin', 'xl/vbaProjectSignature.bin'],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: expect.arrayContaining([
					expect.objectContaining({
						kind: 'vbaProject',
						partPath: 'xl/vbaProject.bin',
						contentType: 'application/vnd.ms-office.vbaProject',
						anchor: 'workbook',
						opaque: true,
						executionPolicy: 'blocked',
					}),
					expect.objectContaining({
						kind: 'vbaSignature',
						partPath: 'xl/vbaProjectSignature.bin',
						invalidationPolicy: 'invalidatedByPackageEdit',
						resigningPolicy: 'notSupported',
					}),
				]),
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened control active content summary', async () => {
		const input = `${TEMP_FILE}.compact-control-active-content-input.xlsx`
		const output = `${TEMP_FILE}.compact-control-active-content-mcp.xlsx`
		await Bun.write(input, controlWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							activeContent?: {
								total?: number
								vbaProjects?: number
								activeXControls?: number
								formControls?: number
								macroSheets?: number
								vbaSignatures?: number
								digitalSignatures?: number
								partPaths?: readonly string[]
								executionPolicy?: string
								preservationMode?: string
								verification?: string
								entries?: Array<{
									kind?: string
									partPath?: string
									contentType?: string
									anchor?: string
									sheetName?: string
									sourceRelationshipId?: string
									relType?: string
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.activeContent).toMatchObject({
				total: 3,
				vbaProjects: 0,
				activeXControls: 1,
				formControls: 1,
				macroSheets: 0,
				vbaSignatures: 0,
				digitalSignatures: 0,
				partPaths: [
					'xl/activeX/activeX1.xml',
					'xl/activeX/activeX1.bin',
					'xl/ctrlProps/ctrlProp1.xml',
				],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: expect.arrayContaining([
					expect.objectContaining({
						kind: 'activeX',
						partPath: 'xl/activeX/activeX1.xml',
						contentType: 'application/vnd.ms-office.activeX+xml',
						anchor: 'sheet',
						sheetName: 'Data',
						sourceRelationshipId: 'rIdActiveX',
						relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControl',
					}),
					expect.objectContaining({
						kind: 'activeX',
						partPath: 'xl/activeX/activeX1.bin',
						contentType: 'application/vnd.ms-office.activeX',
						sourcePartPath: 'xl/activeX/activeX1.xml',
						sourceRelationshipId: 'rId1',
						relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary',
					}),
					expect.objectContaining({
						kind: 'formControl',
						partPath: 'xl/ctrlProps/ctrlProp1.xml',
						contentType: 'application/vnd.ms-excel.controlproperties+xml',
						anchor: 'sheet',
						sheetName: 'Data',
						sourceRelationshipId: 'rIdCtrl',
						relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
					}),
				]),
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened macro sheet active content summary', async () => {
		const input = `${TEMP_FILE}.compact-macro-sheet-input.xlsm`
		const output = `${TEMP_FILE}.compact-macro-sheet-mcp.xlsm`
		await Bun.write(input, macroSheetWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							activeContent?: {
								total?: number
								vbaProjects?: number
								activeXControls?: number
								formControls?: number
								macroSheets?: number
								vbaSignatures?: number
								digitalSignatures?: number
								partPaths?: readonly string[]
								executionPolicy?: string
								preservationMode?: string
								verification?: string
								entries?: Array<{
									kind?: string
									partPath?: string
									contentType?: string
									anchor?: string
									sheetName?: string
									sourceRelationshipId?: string
									relType?: string
									opaque?: boolean
									executionPolicy?: string
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.activeContent).toMatchObject({
				total: 1,
				vbaProjects: 0,
				activeXControls: 0,
				formControls: 0,
				macroSheets: 1,
				vbaSignatures: 0,
				digitalSignatures: 0,
				partPaths: ['xl/macrosheets/sheet1.xml'],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: [
					expect.objectContaining({
						kind: 'macroSheet',
						partPath: 'xl/macrosheets/sheet1.xml',
						contentType: 'application/vnd.ms-excel.macrosheet+xml',
						anchor: 'sheet',
						sheetName: 'Macro1',
						sourceRelationshipId: 'rIdMacro',
						relType: 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
						opaque: true,
						executionPolicy: 'blocked',
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened custom UI active content summary', async () => {
		const input = `${TEMP_FILE}.compact-custom-ui-input.xlsm`
		const output = `${TEMP_FILE}.compact-custom-ui-mcp.xlsm`
		await Bun.write(input, customUiWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							activeContent?: {
								total?: number
								vbaProjects?: number
								activeXControls?: number
								formControls?: number
								macroSheets?: number
								vbaSignatures?: number
								digitalSignatures?: number
								customUi?: number
								unknownActiveContent?: number
								partPaths?: readonly string[]
								executionPolicy?: string
								preservationMode?: string
								verification?: string
								entries?: Array<{
									kind?: string
									partPath?: string
									contentType?: string
									anchor?: string
									executionPolicy?: string
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.activeContent).toMatchObject({
				total: 1,
				vbaProjects: 0,
				activeXControls: 0,
				formControls: 0,
				macroSheets: 0,
				vbaSignatures: 0,
				digitalSignatures: 0,
				customUi: 1,
				unknownActiveContent: 0,
				partPaths: ['customUI/customUI2.xml'],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: [
					expect.objectContaining({
						kind: 'customUi',
						partPath: 'customUI/customUI2.xml',
						contentType: 'application/vnd.ms-office.customUI+xml',
						anchor: 'workbook',
						executionPolicy: 'blocked',
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened visual summary', async () => {
		const input = `${TEMP_FILE}.compact-visual-summary-input.xlsx`
		const output = `${TEMP_FILE}.compact-visual-summary-mcp.xlsx`
		await Bun.write(input, visualWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							visuals?: {
								sheetsWithVisuals?: number
								images?: number
								drawingObjects?: number
								drawingMlObjects?: number
								vmlObjects?: number
								chartParts?: number
								chartSheets?: number
								drawingPartPaths?: readonly string[]
								mediaPartPaths?: readonly string[]
								chartPartPaths?: readonly string[]
								vmlPartPaths?: readonly string[]
								preservationMode?: string
								verification?: string
								sheets?: Array<{
									sheetName?: string
									hasDrawingMl?: boolean
									hasVml?: boolean
									imageCount?: number
									drawingPartPaths?: readonly string[]
									mediaPartPaths?: readonly string[]
								}>
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.visuals).toMatchObject({
				sheetsWithVisuals: 1,
				images: 1,
				drawingObjects: 0,
				drawingMlObjects: 0,
				vmlObjects: 0,
				chartParts: 0,
				chartSheets: 0,
				drawingPartPaths: ['xl/drawings/drawing1.xml'],
				mediaPartPaths: ['xl/media/image1.png'],
				chartPartPaths: [],
				vmlPartPaths: [],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				sheets: [
					expect.objectContaining({
						sheetName: 'Sheet1',
						hasDrawingMl: true,
						hasVml: false,
						imageCount: 1,
						drawingPartPaths: ['xl/drawings/drawing1.xml'],
						mediaPartPaths: ['xl/media/image1.png'],
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP commit exposes reopened chartsheet visual summary', async () => {
		const input = CHARTSHEET_FIXTURE
		const output = `${TEMP_FILE}.compact-chartsheet-summary-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
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
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: {
							auditsPassed?: boolean
							visuals?: {
								sheetsWithVisuals?: number
								images?: number
								chartParts?: number
								chartSheets?: number
								drawingPartPaths?: readonly string[]
								mediaPartPaths?: readonly string[]
								chartPartPaths?: readonly string[]
								vmlPartPaths?: readonly string[]
								preservationMode?: string
								verification?: string
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.visuals).toMatchObject({
				sheetsWithVisuals: 0,
				images: 0,
				chartParts: 1,
				chartSheets: 1,
				drawingPartPaths: [],
				mediaPartPaths: [],
				chartPartPaths: ['xl/charts/chart1.xml'],
				vmlPartPaths: [],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
			})
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP plan and commit expose invalidated signature preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-signature-invalidation-input.xlsx`
		await Bun.write(input, signedPackageWorkbook())
		const output = `${TEMP_FILE}.compact-signature-invalidation-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
						approvals?: readonly { readonly id: string }[]
						writePolicy?: {
							summary?: {
								preservationModes?: {
									invalidatedOnEditParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				planHandle?: string
				output?: string
				approvals?: readonly string[]
				compact?: boolean
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						postWrite?: { auditsPassed?: boolean }
						writePolicy?: {
							summary?: {
								preservationModes?: {
									invalidatedOnEditParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(
				planned.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				invalidatedOnEditParts: 2,
				lossyApprovalRequiredFeatures: 1,
			})
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedsignature:preserved:/)])

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(
				committed.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				invalidatedOnEditParts: 2,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact MCP plan and commit expose discarded calc-chain preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-calc-chain-discard-input.xlsx`
		await Bun.write(input, calcChainWorkbook())
		const output = `${TEMP_FILE}.compact-calc-chain-discard-mcp.xlsx`
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: readonly Record<string, unknown>[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						preparedPlan?: { id?: string }
						writePolicy?: {
							summary?: {
								calcChainPolicy?: string
								preservationModes?: {
									discardedForRecalcParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
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
						postWrite?: { auditsPassed?: boolean }
						writePolicy?: {
							summary?: {
								calcChainPolicy?: string
								preservationModes?: {
									discardedForRecalcParts?: number
									lossyApprovalRequiredFeatures?: number
								}
							}
						}
					}
				}
			}>

			const planned = await plan({
				file: input,
				compact: true,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1+A1' }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.writePolicy?.summary?.calcChainPolicy).toBe(
				'discarded-for-formula-topology',
			)
			expect(
				planned.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				discardedForRecalcParts: 1,
				lossyApprovalRequiredFeatures: 0,
			})

			const committed = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.writePolicy?.summary?.calcChainPolicy).toBe(
				'discarded-for-formula-topology',
			)
			expect(
				committed.structuredContent?.data?.writePolicy?.summary?.preservationModes,
			).toMatchObject({
				discardedForRecalcParts: 1,
				lossyApprovalRequiredFeatures: 0,
			})
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
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
						inputSha256?: string
						planDigest?: string
						operationCount?: number
						preparedPlan?: {
							id?: string
							file?: string
							inputSha256?: string
							planDigest?: string
							operationCount?: number
							expiresAt?: string
							ttlMs?: number
						}
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
				includeProofBundle?: boolean
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
						proofBundle?: {
							safeToUse?: boolean
							whatChanged?: Array<{ ref: string }>
							whySafe?: Array<{ gate: string; ok: boolean }>
							evidence?: {
								outputSha256?: string
								postWriteValid?: boolean
								auditsPassed?: boolean
								reopened?: boolean
								checkValid?: boolean
								lintClean?: boolean
							}
						}
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
			expect(planned.structuredContent?.data?.preparedPlan?.file).toBe(TEMP_FILE)
			expect(planned.structuredContent?.data?.preparedPlan?.inputSha256).toBe(
				planned.structuredContent?.data?.inputSha256,
			)
			expect(planned.structuredContent?.data?.preparedPlan?.planDigest).toBe(
				planned.structuredContent?.data?.planDigest,
			)
			expect(planned.structuredContent?.data?.preparedPlan?.operationCount).toBe(
				planned.structuredContent?.data?.operationCount,
			)
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
				includeProofBundle: true,
			})
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 321 }] },
			])
			expect(committed.structuredContent?.data?.apply?.affectedCellCount).toBe(1)
			expect(committed.structuredContent?.data?.proofBundle).toMatchObject({
				safeToUse: true,
				evidence: {
					postWriteValid: true,
					auditsPassed: true,
					reopened: true,
					checkValid: true,
					lintClean: true,
				},
			})
			expect(committed.structuredContent?.data?.proofBundle?.whatChanged).toEqual([{ ref: 'A1' }])
			expect(
				committed.structuredContent?.data?.proofBundle?.whySafe?.map((gate) => [
					gate.gate,
					gate.ok,
				]),
			).toEqual([
				['input-guard', true],
				['approval', true],
				['write-policy', true],
				['commit', true],
				['reopen-verify', true],
				['package-graph', true],
			])
			expect(committed.structuredContent?.data?.timings?.applyMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.writePlanSummaryMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.writePolicyCheckMs).toBeNumber()
			expect(committed.structuredContent?.data?.timings?.writePolicyCheckMs).toBe(0)
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

	test('ascend.commit prepared plan handles report failed writes and remain retryable', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const blockedOutput = join(
			tmpdir(),
			`ascend-mcp-prepared-missing-parent-${Date.now()}`,
			'out.xlsx',
		)
		const retryOutput = `${TEMP_FILE}.prepared-failed-write-retry-mcp.xlsx`
		const ops = [{ op: 'addSheet' as const, name: 'PreparedRetryMcp' }]
		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const plan = (server as any)._registeredTools['ascend.plan'].handler as (args: {
			file: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				data?: { preparedPlan?: { id?: string } }
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
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: { output?: string; operation?: string }
				}
				data?: { postWrite?: { valid?: boolean; auditsPassed?: boolean } }
			}
		}>

		try {
			const planned = await plan({ file: TEMP_FILE, ops })
			expect(planned.structuredContent?.data?.preparedPlan?.id).toBeString()

			const blocked = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: blockedOutput,
				approvals: [],
				compact: true,
			})
			expect(blocked.isError).toBe(true)
			expect(blocked.structuredContent).toMatchObject({
				ok: false,
				error: {
					code: 'EXPORT_ERROR',
					retryable: true,
					retryStrategy: 'modified',
					details: {
						output: blockedOutput,
						operation: 'atomic-workbook-write',
					},
				},
			})
			expect(await Bun.file(blockedOutput).exists()).toBe(false)

			const retried = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: retryOutput,
				approvals: [],
				compact: true,
			})
			expect(retried.structuredContent?.ok).toBe(true)
			expect(retried.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(retried.structuredContent?.data?.postWrite?.auditsPassed).toBe(true)
			const reopened = await AscendWorkbook.open(retryOutput)
			expect(reopened.sheets).toContain('PreparedRetryMcp')

			const reused = await commit({
				planHandle: planned.structuredContent?.data?.preparedPlan?.id,
				output: `${retryOutput}.reuse.xlsx`,
				approvals: [],
			})
			expect(reused.structuredContent?.error?.message).toBe(
				'Prepared plan handle has already been used',
			)
		} finally {
			await unlink(retryOutput).catch(() => {})
			await unlink(`${retryOutput}.reuse.xlsx`).catch(() => {})
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
							packageGraphAudit?: {
								ok?: boolean
								issueCount?: number
								emittedIssueCount?: number
								issues?: Array<{
									code?: string
									partPath?: string
									preservationPolicy?: string
									preservationMode?: string
								}>
							}
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
			expect(
				exactCommit.structuredContent?.data?.postWrite?.packageGraphAudit?.emittedIssueCount,
			).toBeGreaterThan(0)
			expect(exactCommit.structuredContent?.data?.postWrite?.packageGraphAudit?.issues).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						code: 'package_feature_classification',
						partPath: 'xl/custom/custom1.xml',
						preservationPolicy: 'unknown-review-required',
						preservationMode: 'review-required',
					}),
				]),
			)
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

	test('direct MCP path mutation commits surface post-write audit failures as blocked output', async () => {
		await Bun.write(TEMP_FILE, preservedCustomWorkbook())
		const output = `${TEMP_FILE}.direct-preserved-mcp.xlsx`
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
						approvals?: Array<{ id: string }>
						pathMutations?: { ops?: unknown[] }
					}
				}
			}>
			// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
			const commit = (server as any)._registeredTools['ascend.commit'].handler as (args: {
				file?: string
				mutations?: Array<{ path: string; value?: unknown }>
				output?: string
				approvals?: string[]
				compact?: boolean
			}) => Promise<{
				isError?: boolean
				structuredContent?: {
					ok?: boolean
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
			expect(planned.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			const approvalIds =
				planned.structuredContent?.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedother:preserved:/)])

			const committed = await commit({
				file: TEMP_FILE,
				mutations,
				output,
				approvals: approvalIds,
			})
			expect(committed.isError).not.toBe(true)
			expect(committed.structuredContent?.ok).toBe(true)
			expect(committed.structuredContent?.data?.pathMutations?.ops).toEqual(canonicalOps)
			expect(committed.structuredContent?.data?.approvals?.map((approval) => approval.id)).toEqual(
				approvalIds,
			)
			expect(committed.structuredContent?.data?.postWrite?.valid).toBe(true)
			expect(committed.structuredContent?.data?.postWrite?.auditsPassed).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.outputSha256).toBe(
				committed.structuredContent?.data?.outputSha256,
			)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.ok).toBe(false)
			expect(committed.structuredContent?.data?.postWrite?.packageGraphAudit?.issues).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						code: 'package_feature_classification',
						partPath: 'xl/custom/custom1.xml',
						preservationPolicy: 'unknown-review-required',
						preservationMode: 'review-required',
					}),
				]),
			)
			expect(committed.structuredContent?.data?.postWrite?.expectedPackageGraphIssueCount).toBe(0)
			expect(
				committed.structuredContent?.data?.postWrite?.unresolvedPackageGraphIssueCount,
			).toBeGreaterThan(0)
			expect(committed.structuredContent?.data?.modelOutput?.blocked).toBe(true)
			expect(
				committed.structuredContent?.data?.modelOutput?.counts?.postWritePackageGraphIssues,
			).toBeGreaterThan(0)
			expect(committed.structuredContent?.data?.modelOutput?.nextActions?.join('\n')).toContain(
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

	test('ascend.read exposes array and shared formula binding metadata', async () => {
		await Bun.write(TEMP_FILE, sharedFormulaWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			sheet?: string
			range: string
			format: 'cells'
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					cells?: Array<{
						readonly ref?: string
						readonly formula?: string
						readonly formulaBinding?: unknown
					}>
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:D2',
			format: 'cells',
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(
			result.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['A1', 'SUM(B1:B2)', { kind: 'array', ref: 'A1:A2' }],
			['B1', 'A1*2', { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' }],
			['C1', 'SUM(Sales[[Revenue]:[Quantity]])', null],
			['D1', 'BudgetTotal*2', null],
			['B2', 'A2*2', { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' }],
		])

		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const inspectHandler = (server as any)._registeredTools['ascend.inspect'].handler as (args: {
			file: string
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					definedNameDetails?: readonly {
						readonly name?: string
						readonly formula?: string
						readonly normalizedFormula?: string
						readonly scope?: string
						readonly refs?: readonly string[]
					}[]
				}
			}
		}>
		const inspect = await inspectHandler({ file: TEMP_FILE })
		expect(inspect.structuredContent?.ok).toBe(true)
		expect(inspect.structuredContent?.data?.definedNameDetails).toContainEqual({
			name: 'BudgetTotal',
			formula: 'Calc!$A$1:$A$2',
			normalizedFormula: 'Calc!$A$1:$A$2',
			scope: 'workbook',
			references: [
				{ kind: 'range', text: 'Calc!$A$1:$A$2', scope: { kind: 'sheet', sheet: 'Calc' } },
			],
			refs: ['Calc!$A$1:$A$2'],
			functions: [],
			volatile: false,
		})
	})

	test('ascend.read exposes dynamic-array formulas and binding metadata', async () => {
		await Bun.write(TEMP_FILE, dynamicArrayWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			sheet?: string
			range: string
			format: 'cells'
			display?: boolean
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					cells?: Array<{
						readonly ref?: string
						readonly value?: string
						readonly formula?: string
						readonly formulaBinding?: unknown
					}>
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:C1',
			format: 'cells',
			display: true,
		})
		expect(result.structuredContent?.ok).toBe(true)
		expect(
			result.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.value,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['A1', '1', 'SEQUENCE(3)', { kind: 'dynamicArray', metadataIndex: 1, collapsed: false }],
			['B1', '6', 'SUM(A1#)', null],
			['C1', '1', '@A1', null],
		])
	})

	test('ascend.write blocks structural edits against imported formula bindings without changing read truth', async () => {
		await Bun.write(TEMP_FILE, sharedOnlyFormulaWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					code?: string
					message?: string
					refs?: readonly string[]
					suggestedFix?: string
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const read = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			sheet?: string
			range: string
			format: 'cells'
			display?: boolean
		}) => Promise<{
			structuredContent?: {
				data?: {
					cells?: Array<{
						readonly ref?: string
						readonly value?: string
						readonly formula?: string
						readonly formulaBinding?: unknown
					}>
				}
			}
		}>

		const sharedWrite = await write({
			file: TEMP_FILE,
			ops: [{ op: 'insertRows', sheet: 'Calc', at: 0, count: 1 }],
		})
		expect(sharedWrite.isError).toBe(true)
		expect(sharedWrite.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(sharedWrite.structuredContent?.error?.message).toContain(
			'Calc!B1 contains imported shared formula metadata',
		)
		expect(sharedWrite.structuredContent?.error?.refs).toEqual(['Calc!B1'])
		expect(sharedWrite.structuredContent?.error?.suggestedFix).toContain('Materialize or rewrite')
		const sharedRead = await read({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'B1:B2',
			format: 'cells',
		})
		expect(
			sharedRead.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['B1', 'A1*2', { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' }],
			['B2', 'A2*2', { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' }],
		])

		await Bun.write(TEMP_FILE, dynamicArrayWorkbook())
		const dynamicWrite = await write({
			file: TEMP_FILE,
			ops: [{ op: 'insertCols', sheet: 'Calc', at: 0, count: 1 }],
		})
		expect(dynamicWrite.isError).toBe(true)
		expect(dynamicWrite.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(dynamicWrite.structuredContent?.error?.message).toContain(
			'Calc!A1 contains imported dynamicArray formula metadata',
		)
		expect(dynamicWrite.structuredContent?.error?.refs).toEqual(['Calc!A1'])
		expect(dynamicWrite.structuredContent?.error?.suggestedFix).toContain('Materialize or rewrite')
		const dynamicRead = await read({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:C1',
			format: 'cells',
			display: true,
		})
		expect(
			dynamicRead.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.value,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['A1', '1', 'SEQUENCE(3)', { kind: 'dynamicArray', metadataIndex: 1, collapsed: false }],
			['B1', '6', 'SUM(A1#)', null],
			['C1', '1', '@A1', null],
		])
	})

	test('ascend.write can explicitly rewrite imported formula bindings without stale metadata', async () => {
		await Bun.write(TEMP_FILE, sharedOnlyFormulaWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const read = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			sheet?: string
			range: string
			format: 'cells'
			display?: boolean
		}) => Promise<{
			structuredContent?: {
				data?: {
					cells?: Array<{
						readonly ref?: string
						readonly value?: string
						readonly formula?: string
						readonly formulaBinding?: unknown
					}>
				}
			}
		}>

		const sharedWrite = await write({
			file: TEMP_FILE,
			ops: [
				{ op: 'setFormula', sheet: 'Calc', ref: 'B1', formula: '1+1' },
				{ op: 'setFormula', sheet: 'Calc', ref: 'B2', formula: '2+2' },
			],
		})
		expect(sharedWrite.structuredContent?.ok).toBe(true)
		const sharedRead = await read({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'B1:B2',
			format: 'cells',
			display: true,
		})
		expect(
			sharedRead.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.value,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['B1', '2', '1+1', null],
			['B2', '4', '2+2', null],
		])

		await Bun.write(TEMP_FILE, dynamicArrayWorkbook())
		const dynamicWrite = await write({
			file: TEMP_FILE,
			ops: [
				{ op: 'setFormula', sheet: 'Calc', ref: 'A1', formula: '1+1' },
				{ op: 'setFormula', sheet: 'Calc', ref: 'B1', formula: 'A1+1' },
				{ op: 'setFormula', sheet: 'Calc', ref: 'C1', formula: 'A1+2' },
			],
		})
		expect(dynamicWrite.structuredContent?.ok).toBe(true)
		const dynamicRead = await read({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:C1',
			format: 'cells',
			display: true,
		})
		expect(
			dynamicRead.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.value,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['A1', '2', '1+1', null],
			['B1', '3', 'A1+1', null],
			['C1', '4', 'A1+2', null],
		])
	})

	test('ascend.write reports materialized shared formula group when rewriting one member', async () => {
		await Bun.write(TEMP_FILE, sharedOnlyFormulaWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const write = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: readonly Record<string, unknown>[]
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					affectedCells?: readonly string[]
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const read = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			sheet?: string
			range: string
			format: 'cells'
		}) => Promise<{
			structuredContent?: {
				data?: {
					cells?: Array<{
						readonly ref?: string
						readonly formula?: string
						readonly formulaBinding?: unknown
					}>
				}
			}
		}>

		const writeResult = await write({
			file: TEMP_FILE,
			ops: [{ op: 'setFormula', sheet: 'Calc', ref: 'B2', formula: '2+2' }],
		})
		expect(writeResult.structuredContent?.ok).toBe(true)
		expect(writeResult.structuredContent?.data?.affectedCells).toEqual(['B1', 'B2'])

		const readResult = await read({
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'B1:B2',
			format: 'cells',
		})
		expect(
			readResult.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.formula,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['B1', 'A1*2', null],
			['B2', '2+2', null],
		])
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

	test('ascend.read compact changedSince invalidates when the requested window changes', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
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
			rowLimit?: number
			changedSince?: string
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: { cells?: unknown[]; changeToken?: string; changeInvalidation?: unknown }
			}
		}>

		const first = await handler({
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'compact',
			rowLimit: 1,
		})
		expect(first.structuredContent?.ok).toBe(true)
		expect(first.structuredContent?.data?.cells).toEqual([[0, 0, 1]])
		expect(first.structuredContent?.data?.changeToken).toBeDefined()

		const widened = await handler({
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'compact',
			rowLimit: 2,
			changedSince: first.structuredContent?.data?.changeToken,
		})
		expect(widened.structuredContent?.ok).toBe(true)
		expect(widened.structuredContent?.data?.cells).toEqual([
			[0, 0, 1],
			[1, 0, 2],
		])
		expect(widened.structuredContent?.data?.changeInvalidation).toEqual({
			baseToken: first.structuredContent?.data?.changeToken,
			changeToken: widened.structuredContent?.data?.changeToken,
			reason: 'base-snapshot-missing',
			requiredAction: 'use-returned-window',
		})

		const invalid = await handler({
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'compact',
			rowLimit: 2,
			changedSince: 'not-a-token',
		})
		expect(invalid.structuredContent?.ok).toBe(true)
		expect(invalid.structuredContent?.data?.cells).toEqual([
			[0, 0, 1],
			[1, 0, 2],
		])
		expect(invalid.structuredContent?.data?.changeInvalidation).toEqual({
			baseToken: 'not-a-token',
			changeToken: invalid.structuredContent?.data?.changeToken,
			reason: 'base-token-invalid',
			requiredAction: 'use-returned-window',
		})
	})

	test('ascend.read compact changedSince invalidates when selected columns change', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'left' },
					{ ref: 'C1', value: 'right' },
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
			changedSince?: string
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					cells?: unknown[]
					changeToken?: string
					changeInvalidation?: unknown
					selectedColumns?: readonly unknown[]
				}
			}
		}>

		const first = await handler({
			file: TEMP_FILE,
			range: 'A1:C1',
			format: 'compact',
			cols: ['A'],
		})
		expect(first.structuredContent?.ok).toBe(true)
		expect(first.structuredContent?.data?.cells).toEqual([[0, 0, 'left']])
		expect(first.structuredContent?.data?.changeToken).toBeDefined()

		const changedProjection = await handler({
			file: TEMP_FILE,
			range: 'A1:C1',
			format: 'compact',
			cols: ['C'],
			changedSince: first.structuredContent?.data?.changeToken,
		})
		expect(changedProjection.structuredContent?.ok).toBe(true)
		expect(changedProjection.structuredContent?.data?.cells).toEqual([[0, 2, 'right']])
		expect(changedProjection.structuredContent?.data?.changeInvalidation).toEqual({
			baseToken: first.structuredContent?.data?.changeToken,
			changeToken: changedProjection.structuredContent?.data?.changeToken,
			reason: 'base-snapshot-missing',
			requiredAction: 'use-returned-window',
		})
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
				data?: { cells?: unknown[]; changeToken?: string; changeInvalidation?: unknown }
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
		expect(afterChange.structuredContent?.data?.changeInvalidation).toEqual({
			baseToken: first.structuredContent?.data?.changeToken,
			changeToken: afterChange.structuredContent?.data?.changeToken,
			reason: 'base-snapshot-missing',
			requiredAction: 'use-returned-window',
		})
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

	test('ascend.agent_view exposes token budget metadata', async () => {
		const wb = AscendWorkbook.create()
		const updates = []
		for (let row = 1; row <= 20; row++) {
			for (let col = 0; col < 4; col++) {
				updates.push({
					ref: `${String.fromCharCode(65 + col)}${row}`,
					value: row === 1 ? `Header ${col + 1}` : `r${row}-c${col + 1}`,
				})
			}
		}
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.agent_view'].handler as (args: {
			file: string
			range: string
			maxApproxTokens?: number
		}) => Promise<{
			structuredContent?: {
				ok?: boolean
				data?: {
					rowCount?: number
					colCount?: number
					budget?: {
						requestedApproxTokens?: number
						truncated?: boolean
						omittedSampleRows?: number
						omittedColumnSampleValues?: number
						omittedEvidence?: {
							sampleRows?: { count?: number }
						}
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			range: 'A1:D20',
			maxApproxTokens: 384,
		})

		expect(result.structuredContent?.ok).toBe(true)
		expect(result.structuredContent?.data?.budget?.requestedApproxTokens).toBe(384)
		expect(result.structuredContent?.data?.budget?.truncated).toBe(true)
		expect(result.structuredContent?.data?.rowCount).toBe(20)
		expect(result.structuredContent?.data?.colCount).toBe(4)
		expect(
			(result.structuredContent?.data?.budget?.omittedSampleRows ?? 0) +
				(result.structuredContent?.data?.budget?.omittedColumnSampleValues ?? 0),
		).toBeGreaterThan(0)
		expect(result.structuredContent?.data?.budget?.omittedEvidence?.sampleRows?.count).toBe(
			result.structuredContent?.data?.budget?.omittedSampleRows,
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

	test('ascend.trace reports missing cells with structured retry guidance', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] }])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.trace'].handler as (args: {
			file: string
			cell: string
		}) => Promise<{
			isError?: boolean
			content?: Array<{ text?: string }>
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: {
						cell?: string
						workflow?: readonly string[]
						load?: { mode?: string }
					}
					suggestedFix?: string
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE, cell: 'Missing!A1' })

		expect(result.isError).toBe(true)
		expect(result.content?.[0]?.text).toBe('Trace cell not found')
		expect(result.structuredContent).toMatchObject({
			ok: false,
			error: {
				code: 'VALIDATION_ERROR',
				message: 'Trace cell not found',
				retryable: true,
				retryStrategy: 'modified',
				details: {
					cell: 'Missing!A1',
					workflow: ['inspect', 'read', 'trace'],
					load: { mode: 'formula' },
				},
			},
		})
		expect(result.structuredContent?.error?.suggestedFix).toContain('ascend.inspect')
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

	test('ascend.read_table reports missing tables with structured retry guidance', async () => {
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
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Scores', hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.read_table'].handler as (args: {
			file: string
			table: string
		}) => Promise<{
			isError?: boolean
			content?: Array<{ text?: string }>
			structuredContent?: {
				ok?: boolean
				error?: {
					code?: string
					message?: string
					retryable?: boolean
					retryStrategy?: string
					details?: {
						table?: string
						availableTables?: readonly string[]
						workflow?: readonly string[]
					}
					suggestedFix?: string
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE, table: 'Missing' })

		expect(result.isError).toBe(true)
		expect(result.content?.[0]?.text).toBe('Table "Missing" not found')
		expect(result.structuredContent).toMatchObject({
			ok: false,
			error: {
				code: 'TABLE_NOT_FOUND',
				message: 'Table "Missing" not found',
				retryable: true,
				retryStrategy: 'modified',
				details: {
					table: 'Missing',
					availableTables: ['Scores'],
					workflow: ['inspect', 'read_table'],
				},
			},
		})
		expect(result.structuredContent?.error?.suggestedFix).toContain('Scores')
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

	test('ascend.dump and ascend.template_merge reject capped load options instead of emitting replay ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: '{{name}}' }],
			},
		])
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const dumpHandler = (server as any)._registeredTools['ascend.dump'].handler as (args: {
			file: string
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				data?: { ops?: unknown[] }
				error?: {
					code?: string
					details?: {
						unsupportedLoadOptions?: readonly string[]
						requiredLoad?: { mode?: string; allSheets?: boolean; maxRows?: null | number }
					}
				}
			}
		}>
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const mergeHandler = (server as any)._registeredTools['ascend.template_merge']
			.handler as (args: {
			file: string
			data: Record<string, string | number | boolean | null>
			maxRows?: number
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				data?: { ops?: unknown[] }
				error?: {
					code?: string
					details?: {
						unsupportedLoadOptions?: readonly string[]
						requiredLoad?: { mode?: string; allSheets?: boolean; maxRows?: null | number }
					}
				}
			}
		}>

		const dump = await dumpHandler({ file: TEMP_FILE, maxRows: 1 })
		expect(dump.isError).toBe(true)
		expect(dump.structuredContent?.data?.ops).toBeUndefined()
		expect(dump.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(dump.structuredContent?.error?.details?.unsupportedLoadOptions).toEqual(['maxRows'])
		expect(dump.structuredContent?.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})

		const merge = await mergeHandler({ file: TEMP_FILE, data: { name: 'Acme' }, maxRows: 1 })
		expect(merge.isError).toBe(true)
		expect(merge.structuredContent?.data?.ops).toBeUndefined()
		expect(merge.structuredContent?.error?.code).toBe('VALIDATION_ERROR')
		expect(merge.structuredContent?.error?.details?.unsupportedLoadOptions).toEqual(['maxRows'])
		expect(merge.structuredContent?.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})
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

	test('ascend.write errors preserve structured journal build failures', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const handler = (server as any)._registeredTools['ascend.write'].handler as (args: {
			file: string
			ops: unknown[]
			journal?: boolean
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					code?: string
					details?: {
						apply?: {
							journal?: {
								supported?: boolean
								exact?: boolean
								inverseOps?: unknown[]
								issues?: unknown[]
								undoPolicy?: { reason?: string; riskLevel?: string }
							}
						}
					}
				}
			}
		}>

		const result = await handler({
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'clearRange', sheet: 'Sheet1', range: 'A1:', what: 'all' }],
		})

		expect(result.isError).toBe(true)
		expect(result.structuredContent?.error?.code).toBe('INVALID_RANGE')
		expect(result.structuredContent?.error?.details?.apply?.journal).toMatchObject({
			supported: false,
			exact: false,
			inverseOps: [],
			issues: [
				{
					code: 'JOURNAL_UNAVAILABLE',
					surface: 'package-parts',
					reason: 'journal-unavailable',
				},
			],
			undoPolicy: {
				reason: 'unavailable',
				riskLevel: 'high',
			},
		})
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

	test('ascend.check exposes blocked spill diagnostics for public agent repair', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SEQUENCE(3)' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 'blocker' }] },
		])
		wb.recalc()
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
								details?: unknown
							}>
						}
					}
				}
			}
		}>

		const result = await handler({ file: TEMP_FILE })
		expect(result.isError).toBe(true)
		const issue = result.structuredContent?.error?.details?.check?.issues?.find(
			(entry) => entry.rule === 'spill-diagnostics',
		)
		expect(issue?.ref).toBe('Sheet1!A1')
		expect(issue?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(issue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
	})

	test('ascend.check refreshes stale imported spill caches for public agent repair', async () => {
		await Bun.write(TEMP_FILE, staleSpillCacheWorkbook())

		const server = createServer()
		// biome-ignore lint/suspicious/noExplicitAny: using MCP registration internals for behavior testing
		const readHandler = (server as any)._registeredTools['ascend.read'].handler as (args: {
			file: string
			range: string
			format: 'cells'
		}) => Promise<{
			structuredContent?: {
				data?: {
					cells?: Array<{
						readonly ref?: string
						readonly formula?: string
						readonly formulaBinding?: unknown
					}>
				}
			}
		}>
		const read = await readHandler({ file: TEMP_FILE, range: 'A1:A2', format: 'cells' })
		expect(
			read.structuredContent?.data?.cells?.map((cell) => [
				cell.ref,
				cell.formula ?? null,
				cell.formulaBinding ?? null,
			]),
		).toEqual([
			['A1', 'SEQUENCE(3)', null],
			['A2', null, null],
		])

		// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
		const checkHandler = (server as any)._registeredTools['ascend.check'].handler as (args: {
			file: string
		}) => Promise<{
			isError?: boolean
			structuredContent?: {
				error?: {
					details?: {
						check?: {
							issues?: Array<{
								rule?: string
								refs?: string[]
								details?: unknown
							}>
						}
					}
				}
			}
		}>

		const result = await checkHandler({ file: TEMP_FILE })
		expect(result.isError).toBe(true)
		const issue = result.structuredContent?.error?.details?.check?.issues?.find(
			(entry) => entry.rule === 'spill-diagnostics',
		)
		expect(issue?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(issue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
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
		}) => Promise<{
			structuredContent?: {
				data?: {
					approvals?: Array<{ id: string }>
					writePolicy?: {
						diagnostics?: Array<{
							code?: string
							featureFamily?: string
							preservationMode?: string
						}>
					}
				}
			}
		}>
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
		expect(planned.structuredContent?.data?.writePolicy?.diagnostics).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					code: 'active-content-preserved',
					featureFamily: 'preservedMacro',
					preservationMode: 'preserve-exact',
				}),
				expect.objectContaining({
					code: 'approval-required-feature',
					featureFamily: 'preservedSignature',
					preservationMode: 'invalidated-on-edit',
				}),
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

	test('ascend.plan warns for embedded object and vendor security sidecars', async () => {
		const input = `${TEMP_FILE}.embedding-vendor-security.xlsx`
		await Bun.write(input, embeddingVendorSecurityWorkbook())
		const server = createServer()
		try {
			// biome-ignore lint/suspicious/noExplicitAny: accessing internals for test
			const planHandler = (server as any)._registeredTools['ascend.plan'].handler as (args: {
				file: string
				compact?: boolean
				ops: unknown[]
			}) => Promise<{
				structuredContent?: {
					ok?: boolean
					data?: {
						writePolicy?: {
							diagnostics?: Array<{
								code?: string
								featureFamily?: string
								preservationMode?: string
							}>
						}
					}
				}
			}>

			const planned = await planHandler({
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 7 }] }],
			})
			expect(planned.structuredContent?.ok).toBe(true)
			expect(planned.structuredContent?.data?.writePolicy?.diagnostics).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						code: 'active-content-preserved',
						featureFamily: 'preservedEmbedding',
						preservationMode: 'preserve-exact',
					}),
					expect.objectContaining({
						code: 'active-content-preserved',
						featureFamily: 'preservedVendorSecurity',
						preservationMode: 'preserve-exact',
					}),
				]),
			)
		} finally {
			await unlink(input).catch(() => {})
		}
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

function sharedFormulaWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
  <sheets><sheet name="Calc" sheetId="1" r:id="rIdSheet"/></sheets>
  <definedNames><definedName name="BudgetTotal">Calc!$A$1:$A$2</definedName></definedNames>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>3</v></c>
      <c r="B1"><f t="shared" si="0">A1*2</f><v>6</v></c>
      <c r="C1"><f>SUM(Sales[[Revenue]:[Quantity]])</f><v>14</v></c>
      <c r="D1"><f>BudgetTotal*2</f><v>6</v></c>
    </row>
    <row r="2">
      <c r="B2"><f t="shared" si="0"/><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`,
	})
}

function sharedOnlyFormulaWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
  <sheets><sheet name="Calc" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="B1"><f t="shared" si="0">A1*2</f><v>6</v></c></row>
    <row r="2"><c r="B2"><f t="shared" si="0"/><v>8</v></c></row>
  </sheetData>
</worksheet>`,
	})
}

function dynamicArrayWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMetadata" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/metadata.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c>
      <c r="B1"><f>SUM(_xlfn.ANCHORARRAY(A1))</f><v>6</v></c>
      <c r="C1"><f>_xlfn.SINGLE(A1)</f><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`,
	})
}

function staleSpillCacheWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="e"><f>_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>blocker</t></is></c></row>
  </sheetData>
</worksheet>`,
	})
}

function controlWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
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
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdActiveX" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rIdCtrl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
		'xl/activeX/activeX1.xml': `<?xml version="1.0"?><ax:ocx ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
		'xl/activeX/_rels/activeX1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
</Relationships>`,
		'xl/activeX/activeX1.bin': 'active-binary',
		'xl/ctrlProps/ctrlProp1.xml': `<?xml version="1.0"?><formControlPr macro="Module1.Run" fmlaLink="$A$1" fmlaRange="$A$2:$A$4"/>`,
	})
}

function macroSheetWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMacro" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rIdData"/>
    <sheet name="Macro1" sheetId="2" r:id="rIdMacro" state="hidden"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/macrosheets/sheet1.xml': `<?xml version="1.0"?>
<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><f>RUN("Task")</f><v>0</v></c></row></sheetData>
</xm:macrosheet>`,
		'xl/macrosheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
	})
}

function embeddingVendorSecurityWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="xen" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rDellEncryptedDoc" Type="http://schemas.dell.com/ddp/2016/relationships/xenFile" Target="ddp/ddpfile.xen"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <oleObjects><oleObject progId="Package" r:id="rIdOle" shapeId="1025"/></oleObjects>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>`,
		'xl/embeddings/oleObject1.bin': 'embedded-payload',
		'ddp/ddpfile.xen': 'opaque-vendor-security',
	})
}

function customUiWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/customUI/customUI2.xml" ContentType="application/vnd.ms-office.customUI+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdCustomUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI2.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'customUI/customUI2.xml': `<?xml version="1.0"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon.OnLoad" loadImage="Ribbon.LoadImage">
  <ribbon><tabs><tab id="tabAscend" label="Ascend">
    <group id="grpActions" label="Actions">
      <button id="runReport" label="Run" onAction="Module1.RunReport" getEnabled="Ribbon.CanRun"/>
    </group>
  </tab></tabs></ribbon>
</customUI>`,
		'customUI/_rels/customUI2.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../xl/media/image1.png"/>
</Relationships>`,
		'xl/media/image1.png': 'image-bytes',
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

function inspectOnlyWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/customData/item1.data" ContentType="application/vnd.ms-excel.customData"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdPowerQuery" Type="http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup" Target="customData/item1.data"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/customData/item1.data': 'power-query-mashup-bytes',
	})
}

async function writeOpaqueX14Workbook(path: string): Promise<void> {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().sheets[0]
	if (!sheet) throw new Error('expected default sheet')
	sheet.x14ConditionalFormats.push({
		index: 0,
		sqref: 'A1:A5',
		type: 'dataBar',
		priority: 4,
		formulas: [],
		preservedRuleAttributes: { 'xr:uid': '{CF-UID}' },
		preservedRuleChildXml: [
			'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
		],
	})
	sheet.x14DataValidations.push({
		index: 0,
		sqref: 'C2:C5',
		type: 'list',
		operator: 'between',
		formula1: '$A$1:$A$4',
		preservedAttributes: { 'xr:uid': '{DV-UID}' },
		preservedChildXml: ['<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>'],
	})
	await wb.save(path)
}

function signedPackageWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>
  <Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>
</Relationships>`,
		'_xmlsignatures/_rels/origin.sigs.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>
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
		'_xmlsignatures/origin.sigs': '',
		'_xmlsignatures/sig1.xml':
			'<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>',
	})
}

function calcChainWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>
  </sheetData>
</worksheet>`,
		'xl/calcChain.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="B1" i="1"/>
</calcChain>`,
	})
}

function externalLinkBoundWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/externalLinks/externalLink1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rIdExt"/>
</externalLink>`,
		'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../sources/source.xlsx" TargetMode="External"/>
</Relationships>`,
	})
}

function analyticsRefreshWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdSheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rIdTimelineCache" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="timelineCaches/timelineCache1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches><pivotCache cacheId="34" r:id="rIdPivotCache"/></pivotCaches>
  <sheets>
    <sheet name="PivotSheet" sheetId="1" r:id="rIdSheet1"/>
    <sheet name="Raw" sheetId="2" r:id="rIdSheet2"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/sheet2.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>West</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>East</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rIdTimeline" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`,
		'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <location ref="A3:C8" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="1"><pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField></pivotFields>
  <pageFields count="1"><pageField fld="0" item="0" name="Region"/></pageFields>
</pivotTableDefinition>`,
		'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rIdRecords" recordCount="2" refreshOnLoad="1" enableRefresh="1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Raw"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Region" databaseField="1"><sharedItems count="2"><s v="West"/><s v="East"/></sharedItems></cacheField>
    <cacheField name="Sales" databaseField="1"><sharedItems containsNumber="1" count="2"><n v="10"/><n v="20"/></sharedItems></cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
		'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
		'xl/pivotCache/pivotCacheRecords1.xml': `<?xml version="1.0"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <r><x v="0"/><n v="10"/></r>
  <r><x v="1"/><n v="20"/></r>
</pivotCacheRecords>`,
		'xl/slicerCaches/slicerCache1.xml': `<?xml version="1.0"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region">
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="34"><items count="2"><i x="0" s="1"/><i x="1"/></items></tabular></data>
</slicerCacheDefinition>`,
		'xl/slicerCaches/_rels/slicerCache1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlicerUi" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
</Relationships>`,
		'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><slicer name="Region" cache="Slicer_Region" caption="Region"/></slicers>`,
		'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <data><tabular pivotCacheId="34"/></data>
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <state filterId="7" filterPivotName="PivotTable1" filterType="dateRange" filterTabId="2" pivotCacheId="34" singleRangeFilterState="1">
    <selection startDate="2023-01-01T00:00:00" endDate="2023-12-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`,
		'xl/timelineCaches/_rels/timelineCache1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTimelineUi" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`,
		'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><timeline name="Order_Date" cache="Timeline_Order_Date" caption="Order Date"/></timelines>`,
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

function visualWorkbook(): Uint8Array {
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
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`),
				'xl/drawings/_rels/drawing1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDrawing1"/>
</worksheet>`),
				'xl/drawings/drawing1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
</xdr:wsDr>`),
				'xl/media/image1.png': encode('png-bytes'),
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
