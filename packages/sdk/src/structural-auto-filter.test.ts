import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('autoFilter structural range moves', () => {
	test('moveRange rewrites sheet autoFilter metadata through SDK save and reopen', async () => {
		const wb = AscendWorkbook.create()
		const applied = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Owner' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 'Ada' },
					{ ref: 'C2', value: 10 },
					{ ref: 'A3', value: 'East' },
					{ ref: 'B3', value: 'Grace' },
					{ ref: 'C3', value: 5 },
					{ ref: 'A4', value: 'West' },
					{ ref: 'B4', value: 'Linus' },
					{ ref: 'C4', value: 20 },
				],
			},
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C4',
				column: 0,
				values: ['West'],
				sortRef: 'A2:C4',
				sortBy: 'C2:C4',
				descending: true,
			},
			{ op: 'moveRange', sheet: 'Sheet1', source: 'A1:C4', target: 'E2' },
		])
		expect(applied.errors).toEqual([])

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.inspectSheet('Sheet1')?.autoFilter).toMatchObject({
			ref: 'E2:G5',
			columns: [{ colId: 0, kind: 'filters', values: ['West'] }],
			sortState: {
				ref: 'E3:G5',
				conditions: [{ ref: 'G3:G5', descending: true }],
			},
		})
	})
})
