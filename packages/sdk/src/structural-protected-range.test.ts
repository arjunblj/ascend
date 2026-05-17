import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('protected range structural range moves', () => {
	test('moveRange rewrites protected ranges through SDK save and reopen', async () => {
		const wb = AscendWorkbook.create()
		const applied = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 'Editable' },
					{ ref: 'B3', value: 10 },
					{ ref: 'B4', value: 20 },
				],
			},
			{
				op: 'setProtectedRange',
				sheet: 'Sheet1',
				name: 'EditableBudget',
				sqref: 'B2:B4',
				passwordPlaintext: 'password',
			},
			{ op: 'moveRange', sheet: 'Sheet1', source: 'B2:B4', target: 'F6' },
		])
		expect(applied.errors).toEqual([])

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.getWorkbookModel().sheets[0]?.protectedRanges).toEqual([
			{ name: 'EditableBudget', sqref: 'F6:F8', password: '83AF' },
		])
	})
})
