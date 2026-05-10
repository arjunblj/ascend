import { describe, expect, test } from 'bun:test'
import { createSheet } from './sheet.ts'

describe('AutoFilter metadata', () => {
	test('sheet clone preserves autoFilter metadata', () => {
		const sheet = createSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					values: ['Cash', 'Debt'],
				},
			],
			sortState: {
				ref: 'A1:B5',
				conditions: [{ ref: 'B2:B5', descending: true }],
			},
		}
		sheet.sortState = {
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1' }],
		}

		const clone = sheet.clone()
		expect(clone.autoFilter).toEqual(sheet.autoFilter)
		expect(clone.sortState).toEqual(sheet.sortState)
	})

	test('ensureWritable deep-clones filter and top-level sort metadata', () => {
		const sheet = createSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					values: ['Cash', 'Debt'],
					dateGroupItems: [{ year: 2026, month: 3 }],
				},
			],
			sortState: {
				ref: 'A1:B5',
				conditions: [{ ref: 'B2:B5', descending: true }],
			},
		}
		sheet.sortState = {
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1', descending: true }],
		}

		const clone = sheet.clone()
		clone.ensureWritable()
		const filter = clone.autoFilter
		if (!filter) throw new Error('expected autoFilter')
		;(filter.columns[0] as { values?: string[] }).values = ['Equity']
		;(filter.sortState?.conditions[0] as { descending?: boolean }).descending = false
		;(clone.sortState?.conditions[0] as { descending?: boolean }).descending = false

		expect(sheet.autoFilter?.columns[0]?.values).toEqual(['Cash', 'Debt'])
		expect(sheet.autoFilter?.sortState?.conditions[0]?.descending).toBe(true)
		expect(sheet.sortState?.conditions[0]?.descending).toBe(true)
	})
})
