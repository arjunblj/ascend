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
		const internal = wb as unknown as {
			wb: {
				workbookProtection: unknown
				sheets: Array<{
					comments: Map<string, { text: string }>
					conditionalFormats: Array<Record<string, unknown>>
					dataValidations: Array<Record<string, unknown>>
					imageRefs: Array<Record<string, unknown>>
				}>
			}
		}
		internal.wb.workbookProtection = { lockStructure: true }
		internal.wb.sheets[0]?.comments.set('A1', { text: 'note' })
		internal.wb.sheets[0]?.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'cellIs', formulas: ['1'] }],
		})
		internal.wb.sheets[0]?.dataValidations.push({ sqref: 'B1', type: 'list', formula1: '"A,B"' })
		internal.wb.sheets[0]?.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: 'xl/media/image1.png',
		})
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/inspect`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.sheetCount).toBe(1)
		expect(body.loadedSheetCount).toBe(1)
		expect(body.commentCount).toBe(1)
		expect(body.conditionalFormatCount).toBe(1)
		expect(body.dataValidationCount).toBe(1)
		expect(body.imageCount).toBe(0)
		expect(body.hasWorkbookProtection).toBe(true)
		expect(body.sheets).toHaveLength(1)
		expect(body.sheets[0].name).toBe('Sheet1')
		expect(body.sheets[0].commentCount).toBe(1)
		expect(body.sheets[0].conditionalFormatCount).toBe(1)
		expect(body.sheets[0].dataValidationCount).toBe(1)
		expect(body.sheets[0].imageCount).toBe(0)
		expect(body.cellCount).toBe(0)
		expect(body.load.mode).toBe('full')
	})

	test('unknown route returns 404', async () => {
		const res = await fetch(`http://localhost:${server.port}/unknown`)
		expect(res.status).toBe(404)
		const body = await res.json()
		expect(body.error).toBe('Not Found')
	})
})
