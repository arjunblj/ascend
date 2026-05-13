import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('range aggregates', () => {
	test('summarizes sparse selections without materializing empty cells', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'C3', value: 20 },
					{ ref: 'E5', value: true },
					{ ref: 'G7', value: 'label' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'I9', formula: '1/0' },
		])
		expect(wb.recalc().errors).toEqual([])

		const aggregate = wb.aggregateRange('Sheet1', 'A1:J1000000')

		expect(aggregate).toEqual({
			ref: {
				start: { row: 0, col: 0 },
				end: { row: 999999, col: 9 },
			},
			totalCells: 10_000_000,
			populatedCells: 5,
			blankCount: 9_999_995,
			numericCount: 2,
			textCount: 1,
			booleanCount: 1,
			errorCount: 1,
			sum: 30,
			average: 15,
			min: 10,
			max: 20,
			firstError: '#DIV/0!',
		})
	})

	test('counts formatted dates as numeric serials for spreadsheet status bars', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: new Date(Date.UTC(2020, 0, 2)) },
					{ ref: 'A2', value: 2 },
				],
			},
		])

		const aggregate = wb.sheet('Sheet1')?.aggregateRange('A1:A2')

		expect(aggregate?.numericCount).toBe(2)
		expect(aggregate?.sum).toBeGreaterThan(2)
		expect(aggregate?.min).toBe(2)
	})
})
