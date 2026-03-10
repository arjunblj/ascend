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
		expect(body).toEqual({
			formatVersion: 1,
			ok: true,
			data: { status: 'ok' },
		})
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
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.sheetCount).toBe(1)
		expect(body.data.loadedSheetCount).toBe(1)
		expect(body.data.commentCount).toBeNull()
		expect(body.data.conditionalFormatCount).toBeNull()
		expect(body.data.dataValidationCount).toBeNull()
		expect(body.data.imageCount).toBeNull()
		expect(body.data.pivotTableCount).toBe(0)
		expect(body.data.pivotCacheCount).toBe(0)
		expect(body.data.slicerCount).toBe(0)
		expect(body.data.slicerCacheCount).toBe(0)
		expect(body.data.hasWorkbookProtection).toBe(true)
		expect(body.data.sheets).toHaveLength(1)
		expect(body.data.sheets[0].name).toBe('Sheet1')
		expect(body.data.sheets[0].commentCount).toBeNull()
		expect(body.data.sheets[0].conditionalFormatCount).toBeNull()
		expect(body.data.sheets[0].dataValidationCount).toBeNull()
		expect(body.data.sheets[0].imageCount).toBeNull()
		expect(body.data.cellCount).toBeNull()
		expect(body.data.load.mode).toBe('metadata-only')
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
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.name).toBe('Sheet1')
		expect(body.data.cellDataLoaded).toBe(true)
		expect(body.data.cellCount).toBe(1)
		expect(body.data.rowCount).toBe(2)
		expect(body.data.colCount).toBe(2)
	})

	test('read returns versioned machine envelope', async () => {
		const tempFile = join(tempDir, 'read.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] }])
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/read`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, sheet: 'Sheet1', range: 'A1:A1' }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.cells).toHaveLength(1)
		expect(body.data.cells[0].ref).toBe('A1')
	})

	test('read honors pagination and display options', async () => {
		const tempFile = join(tempDir, 'read-paged.xlsx')
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
		])
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/read`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				sheet: 'Sheet1',
				range: 'A1:B3',
				format: 'objects',
				headers: ['Name', 'Score'],
				rowOffset: 1,
				rowLimit: 1,
				display: true,
			}),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.rowOffset).toBe(1)
		expect(body.data.rowLimit).toBe(1)
		expect(body.data.rows).toEqual([{ Name: 'Alice', Score: '10' }])
	})

	test('read errors return machine failure envelope', async () => {
		const res = await fetch(`http://localhost:${server.port}/read`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ range: 'A1:A1' }),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(false)
		expect(body.error.message).toBe('Missing or invalid file')
	})

	test('preview returns a diff without mutating the workbook on disk', async () => {
		const tempFile = join(tempDir, 'preview.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.recalc()
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/preview`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
			}),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.cellChanges.length).toBeGreaterThan(0)
		expect(body.data.writePlan.totalParts).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(tempFile)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 4 })
	})

	test('write returns a failure envelope on operation errors', async () => {
		const tempFile = join(tempDir, 'write-error.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/write`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] }],
			}),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Sheet')
	})

	test('write returns a failure envelope when recalculation reports errors', async () => {
		const tempFile = join(tempDir, 'write-recalc-error.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/write`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }],
			}),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Circular reference detected')
	})

	test('calc returns a failure envelope when recalculation reports errors', async () => {
		const tempFile = join(tempDir, 'calc-recalc-error.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }])
		await wb.save(tempFile)

		const res = await fetch(`http://localhost:${server.port}/calc`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Circular reference detected')
	})

	test('unknown route returns 404', async () => {
		const res = await fetch(`http://localhost:${server.port}/unknown`)
		expect(res.status).toBe(404)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(false)
		expect(body.error.message).toBe('Not Found')
	})

	test('export returns TSV bytes and rejects unsupported formats', async () => {
		const tempFile = join(tempDir, 'export.xlsx')
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
		await wb.save(tempFile)

		const tsvRes = await fetch(`http://localhost:${server.port}/export`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, format: 'tsv' }),
		})
		expect(tsvRes.status).toBe(200)
		expect(tsvRes.headers.get('Content-Type')).toContain('text/tab-separated-values')
		expect(await tsvRes.text()).toContain('Name\tScore')

		const badRes = await fetch(`http://localhost:${server.port}/export`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, format: 'weird' }),
		})
		expect(badRes.status).toBe(400)
		const badBody = await badRes.json()
		expect(badBody.ok).toBe(false)
		expect(badBody.error.message).toContain('Unsupported format')
	})
})
