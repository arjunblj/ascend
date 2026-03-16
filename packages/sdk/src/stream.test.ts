import { describe, expect, test } from 'bun:test'
import { streamWorkbookRows } from './stream.ts'
import { AscendWorkbook } from './workbook.ts'

function createWorkbookBytes(): Uint8Array {
	const workbook = AscendWorkbook.create()
	workbook.apply([
		{
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 1 },
				{ ref: 'B1', value: 'alpha' },
				{ ref: 'A3', value: 3 },
				{ ref: 'B3', value: 9 },
			],
		},
	])
	return workbook.toBytes()
}

describe('streamWorkbookRows', () => {
	test('streams row payloads in values mode', async () => {
		const iter = await streamWorkbookRows(createWorkbookBytes(), { mode: 'values' })
		const rows: Array<{ row: number; colCount: number }> = []
		for await (const row of iter) {
			rows.push({ row: row.row, colCount: row.cells.length })
		}
		expect(rows).toEqual([
			{ row: 0, colCount: 2 },
			{ row: 2, colCount: 2 },
		])
	})

	test('supports selecting a sheet by name', async () => {
		const workbook = AscendWorkbook.create()
		workbook.addSheet('B')
		workbook.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 10 }],
			},
			{
				op: 'setCells',
				sheet: 'B',
				updates: [{ ref: 'A1', value: 20 }],
			},
		])

		const iter = await streamWorkbookRows(workbook.toBytes(), { sheet: 'B' })
		const rows: number[] = []
		for await (const row of iter) {
			const first = row.cells[0]
			if (first?.[1].value.kind === 'number') rows.push(first[1].value.value)
		}
		expect(rows).toEqual([20])
	})
})
