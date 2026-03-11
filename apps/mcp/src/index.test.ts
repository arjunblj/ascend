import { afterAll, describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { createServer } from './index.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-mcp-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
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
		expect(names).toContain('ascend.read')
		expect(names).toContain('ascend.preview')
		expect(names).toContain('ascend.write')
		expect(names).toContain('ascend.calc')
		expect(names).toContain('ascend.check')
		expect(names).toContain('ascend.lint')
		expect(names).toContain('ascend.trace')
		expect(names).toContain('ascend.diff')
		expect(names).toContain('ascend.export')
		expect(names.length).toBe(10)
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
})
