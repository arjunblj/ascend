import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('AscendWorkbook e2e integration', () => {
	test('full load-recalc-save cycle', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'A2', value: 20 },
					{ ref: 'A3', value: 30 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=SUM(A1:A3)' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: '=A1*A2' },
		])
		wb.recalc()

		const b1Before = wb.sheet('Sheet1')?.cell('B1')?.value
		const b2Before = wb.sheet('Sheet1')?.cell('B2')?.value
		expect(b1Before).toEqual({ kind: 'number', value: 60 })
		expect(b2Before).toEqual({ kind: 'number', value: 200 })

		const bytes1 = wb.toBytes()
		const loaded = await AscendWorkbook.open(bytes1)
		loaded.recalc()

		expect(loaded.sheet('Sheet1')?.cell('B1')?.value).toEqual(b1Before)
		expect(loaded.sheet('Sheet1')?.cell('B2')?.value).toEqual(b2Before)

		const bytes2 = loaded.toBytes()
		const loaded2 = await AscendWorkbook.open(bytes2)
		loaded2.recalc()

		expect(loaded2.sheet('Sheet1')?.cell('B1')?.value).toEqual(b1Before)
		expect(loaded2.sheet('Sheet1')?.cell('B2')?.value).toEqual(b2Before)
	})

	test('CSV round-trip with values', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 85 },
					{ ref: 'A3', value: 'Bob' },
					{ ref: 'B3', value: 92 },
				],
			},
		])

		const csv = wb.toCsv()
		const fromCsv = AscendWorkbook.fromCsv(csv)

		expect(fromCsv.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'Name' })
		expect(fromCsv.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'string', value: 'Score' })
		expect(fromCsv.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'string', value: 'Alice' })
		expect(fromCsv.sheet('Sheet1')?.cell('B2')?.value).toEqual({ kind: 'number', value: 85 })
		expect(fromCsv.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'string', value: 'Bob' })
		expect(fromCsv.sheet('Sheet1')?.cell('B3')?.value).toEqual({ kind: 'number', value: 92 })
	})

	test('multi-sheet workflow', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Data' },
			{ op: 'addSheet', name: 'Summary' },
		])
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Data',
				updates: [
					{ ref: 'A1', value: 5 },
					{ ref: 'A2', value: 10 },
					{ ref: 'A3', value: 15 },
				],
			},
			{ op: 'setFormula', sheet: 'Summary', ref: 'A1', formula: '=SUM(Data!A1:A3)' },
			{ op: 'setFormula', sheet: 'Summary', ref: 'A2', formula: '=Data!A1*2' },
		])

		wb.recalc()
		expect(wb.sheet('Summary')?.cell('A1')?.value).toEqual({ kind: 'number', value: 30 })
		expect(wb.sheet('Summary')?.cell('A2')?.value).toEqual({ kind: 'number', value: 10 })

		wb.apply([{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 100 }] }])
		wb.recalc()

		expect(wb.sheet('Summary')?.cell('A1')?.value).toEqual({ kind: 'number', value: 125 })
		expect(wb.sheet('Summary')?.cell('A2')?.value).toEqual({ kind: 'number', value: 200 })
	})

	test('operation batch workflow', () => {
		const wb = AscendWorkbook.create()
		const batchResult = wb.batch([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 2 }] },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A3', value: 3 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A4', formula: '=SUM(A1:A3)' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1+10' },
		])
		expect(batchResult.errors).toHaveLength(0)

		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 1 })
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 2 })
		expect(wb.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'number', value: 3 })
		expect(wb.sheet('Sheet1')?.cell('A4')?.formula).toBe('SUM(A1:A3)')
		expect(wb.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1+10')

		wb.recalc()
		expect(wb.sheet('Sheet1')?.cell('A4')?.value).toEqual({ kind: 'number', value: 6 })
		expect(wb.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 11 })
	})
})
