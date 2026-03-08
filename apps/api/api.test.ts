import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { createServer } from './src/server.ts'

let server: ReturnType<typeof createServer>
let tempDir: string

beforeAll(() => {
	server = createServer({ port: 0 })
	tempDir = mkdtempSync(join(tmpdir(), 'ascend-api-test-'))
})

afterAll(() => {
	rmSync(tempDir, { recursive: true, force: true })
})

describe('API', () => {
	test('health check returns 200', async () => {
		const res = await fetch(`http://localhost:${server.port}/health`)
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body).toEqual({ status: 'ok' })
	})

	test('inspect returns workbook info', async () => {
		const tempFile = join(tempDir, 'test.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/inspect`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.sheetCount).toBe(1)
		expect(body.sheets).toHaveLength(1)
		expect(body.sheets[0].name).toBe('Sheet1')
		expect(body.cellCount).toBe(0)
	})

	test('unknown route returns 404', async () => {
		const res = await fetch(`http://localhost:${server.port}/unknown`)
		expect(res.status).toBe(404)
		const body = await res.json()
		expect(body.error).toBe('Not Found')
	})
})
