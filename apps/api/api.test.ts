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
		expect(body.commentCount).toBeNull()
		expect(body.conditionalFormatCount).toBeNull()
		expect(body.dataValidationCount).toBeNull()
		expect(body.imageCount).toBeNull()
		expect(body.pivotTableCount).toBe(0)
		expect(body.pivotCacheCount).toBe(0)
		expect(body.slicerCount).toBe(0)
		expect(body.slicerCacheCount).toBe(0)
		expect(body.hasWorkbookProtection).toBe(true)
		expect(body.sheets).toHaveLength(1)
		expect(body.sheets[0].name).toBe('Sheet1')
		expect(body.sheets[0].commentCount).toBeNull()
		expect(body.sheets[0].conditionalFormatCount).toBeNull()
		expect(body.sheets[0].dataValidationCount).toBeNull()
		expect(body.sheets[0].imageCount).toBeNull()
		expect(body.cellCount).toBeNull()
		expect(body.load.mode).toBe('metadata-only')
	})

	test('inspect can return a values-loaded sheet summary', async () => {
		const tempFile = join(tempDir, 'sheet-inspect.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 42 }] }])
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/inspect`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, sheet: 'Sheet1' }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.name).toBe('Sheet1')
		expect(body.cellDataLoaded).toBe(true)
		expect(body.cellCount).toBe(1)
		expect(body.rowCount).toBe(2)
		expect(body.colCount).toBe(2)
	})

	test('unknown route returns 404', async () => {
		const res = await fetch(`http://localhost:${server.port}/unknown`)
		expect(res.status).toBe(404)
		const body = await res.json()
		expect(body.error).toBe('Not Found')
	})
})
