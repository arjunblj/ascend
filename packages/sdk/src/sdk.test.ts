import { describe, expect, test } from 'bun:test'
import { Ascend, AscendWorkbook } from './index.ts'

describe('AscendWorkbook', () => {
	test('Ascend is an alias for AscendWorkbook', () => {
		expect(Ascend).toBe(AscendWorkbook)
	})

	test('create returns an empty workbook with one sheet', () => {
		const wb = AscendWorkbook.create()
		expect(wb.sheets).toEqual(['Sheet1'])
		const info = wb.inspect()
		expect(info.sheetCount).toBe(1)
		expect(info.cellCount).toBe(0)
		expect(info.sourceFormat).toBe('ascend')
	})

	test('inspect returns correct sheet info', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'hello' }] }])
		const info = wb.inspect()
		expect(info.cellCount).toBe(1)
		expect(info.sheets[0]?.cellCount).toBe(1)
		expect(info.sheets[0]?.name).toBe('Sheet1')
	})

	test('sheet handle reads cells', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 42 },
					{ ref: 'B2', value: 'text' },
				],
			},
		])
		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()

		const a1 = handle?.cell('A1')
		expect(a1).toBeDefined()
		expect(a1?.value).toEqual({ kind: 'number', value: 42 })

		const b2 = handle?.cell('B2')
		expect(b2).toBeDefined()
		expect(b2?.value).toEqual({ kind: 'string', value: 'text' })

		expect(handle?.cell('C3')).toBeUndefined()
	})

	test('sheet handle returns range info', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
					{ ref: 'B1', value: 3 },
				],
			},
		])
		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		const range = handle?.range('A1:B2')
		expect(range.rowCount).toBe(2)
		expect(range.colCount).toBe(2)
		expect(range.cells.length).toBe(3)
	})

	test('sheet handle usedRange', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 1 },
					{ ref: 'D4', value: 2 },
				],
			},
		])
		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		const used = handle?.usedRange()
		expect(used).toBeDefined()
		expect(used?.start).toEqual({ row: 1, col: 1 })
		expect(used?.end).toEqual({ row: 3, col: 3 })
	})

	test('sheet returns undefined for nonexistent sheet', () => {
		const wb = AscendWorkbook.create()
		expect(wb.sheet('Nope')).toBeUndefined()
	})

	test('apply operations modifies workbook', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 100 }] },
		])
		expect(result.errors).toHaveLength(0)
		expect(result.affectedCells).toContain('A1')
		expect(result.sheetsModified).toContain('Sheet1')

		const cell = wb.sheet('Sheet1')?.cell('A1')
		expect(cell?.value).toEqual({ kind: 'number', value: 100 })
	})

	test('apply returns error for invalid sheet', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
		])
		expect(result.errors).toHaveLength(1)
		expect(result.errors[0]?.code).toBe('SHEET_NOT_FOUND')
	})

	test('recalc updates formula values', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'A2', value: 20 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A3', formula: 'A1+A2' },
		])
		const result = wb.recalc()
		expect(result.changed.length).toBeGreaterThanOrEqual(1)

		const a3 = wb.sheet('Sheet1')?.cell('A3')
		expect(a3).toBeDefined()
		expect(a3?.value).toEqual({ kind: 'number', value: 30 })
	})

	test('save to bytes and re-open roundtrips', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'roundtrip' },
					{ ref: 'B1', value: 42 },
				],
			},
		])

		const bytes = wb.toBytes()
		expect(bytes).toBeInstanceOf(Uint8Array)
		expect(bytes.length).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(bytes)
		const cell = reopened.sheet('Sheet1')?.cell('A1')
		expect(cell).toBeDefined()
		expect(cell?.value).toEqual({ kind: 'string', value: 'roundtrip' })

		const b1 = reopened.sheet('Sheet1')?.cell('B1')
		expect(b1).toBeDefined()
		expect(b1?.value).toEqual({ kind: 'number', value: 42 })
	})

	test('metadata-only open preserves workbook structure without parsing cells', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'summary' }] },
			{ op: 'addSheet', name: 'Archive' },
		])
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { mode: 'metadata-only' })
		expect(reopened.sheets).toEqual(['Sheet1', 'Archive'])
		expect(reopened.inspect().cellCount).toBe(0)
		expect(reopened.sheet('Sheet1')?.cell('A1')).toBeUndefined()
	})

	test('selective open parses only requested sheets', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'main' }] },
			{ op: 'addSheet', name: 'Archive' },
			{ op: 'setCells', sheet: 'Archive', updates: [{ ref: 'A1', value: 'extra' }] },
		])
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { sheets: ['Archive'] })
		expect(reopened.sheets).toEqual(['Archive'])
		expect(reopened.sheet('Archive')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'extra' })
	})

	test('CSV import creates workbook', () => {
		const csv = 'Name,Age\nAlice,30\nBob,25'
		const wb = AscendWorkbook.fromCsv(csv)
		expect(wb.sheets).toEqual(['Sheet1'])

		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		expect(handle?.cell('A1')?.value).toEqual({ kind: 'string', value: 'Name' })
		expect(handle?.cell('B2')?.value).toEqual({ kind: 'number', value: 30 })
	})

	test('CSV export works', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'x' },
					{ ref: 'B1', value: 1 },
					{ ref: 'A2', value: 'y' },
					{ ref: 'B2', value: 2 },
				],
			},
		])
		const csv = wb.toCsv()
		expect(csv).toContain('x')
		expect(csv).toContain('1')
		expect(csv).toContain('y')
		expect(csv).toContain('2')
	})

	test('preview shows changes without modifying workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'before' }] }])

		const preview = wb.preview([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'after' }] },
		])
		expect(preview.errors).toHaveLength(0)
		expect(preview.cellChanges.length).toBeGreaterThanOrEqual(1)

		const cell = wb.sheet('Sheet1')?.cell('A1')
		expect(cell?.value).toEqual({ kind: 'string', value: 'before' })
	})

	test('check returns clean for valid workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		const result = wb.check()
		expect(result.valid).toBe(true)
		expect(result.issues).toHaveLength(0)
	})

	test('lint detects formula parse errors', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=INVALID((' }])
		const result = wb.lint()
		expect(result.clean).toBe(false)
		expect(result.warnings.length).toBeGreaterThanOrEqual(1)
	})

	test('diff detects changes between workbooks', () => {
		const wb1 = AscendWorkbook.create()
		wb1.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const wb2 = AscendWorkbook.create()
		wb2.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] }])

		const diff = wb1.diff(wb2)
		expect(diff.sheets.length).toBeGreaterThanOrEqual(1)
	})

	test('snapshot captures workbook state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] }])
		const snap = wb.snapshot()
		expect(snap.sheets).toHaveLength(1)
		expect(snap.sheets[0]?.name).toBe('Sheet1')
		expect(snap.sheets[0]?.cells.A1?.value).toEqual({ kind: 'number', value: 99 })
		expect(snap.timestamp).toBeGreaterThan(0)
	})

	test('toJSON returns serializable structure', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'json' }] }])
		const json = wb.toJSON()
		const parsed = JSON.parse(JSON.stringify(json))
		expect(parsed.sheets).toBeDefined()
		expect(parsed.sheets[0].name).toBe('Sheet1')
	})

	test('report returns compatibility info', () => {
		const wb = AscendWorkbook.create()
		expect(wb.report.status).toBe('clean')
		expect(wb.report.sourceFormat).toBe('ascend')
	})

	test('names returns defined name keys', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'MyRange', ref: 'Sheet1!A1:B10' }])
		expect(wb.names).toContain('MyRange')
	})

	test('add and delete sheets', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Extra' }])
		expect(wb.sheets).toEqual(['Sheet1', 'Extra'])

		wb.apply([{ op: 'deleteSheet', sheet: 'Extra' }])
		expect(wb.sheets).toEqual(['Sheet1'])
	})
})
