import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { readXlsx, writeXlsx } from '@ascend/io-xlsx'
import { EMPTY, numberValue } from '@ascend/schema'
import { defaultCalcContext, recalculate } from './index.ts'

const sid = 0 as StyleId

describe('stress: large workbooks', () => {
	test(
		'100K cells (1000x100) write to XLSX and read back, verify cell count',
		async () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const rows = 1000
			const cols = 100
			for (let r = 0; r < rows; r++) {
				for (let c = 0; c < cols; c++) {
					sheet.cells.set(r, c, {
						value: numberValue(r * cols + c + 1),
						formula: null,
						styleId: sid,
					})
				}
			}
			const written = writeXlsx(wb)
			if (!written.ok) throw new Error(written.error.message)
			const read = readXlsx(written.value)
			if (!read.ok) throw new Error(read.error.message)
			const readSheet = read.value.workbook.sheets[0]
			if (!readSheet) throw new Error('No sheet in read result')
			let count = 0
			for (const [, rowCells] of readSheet.cells.iterateRows()) {
				count += rowCells.length
			}
			expect(count).toBe(rows * cols)
		},
		{ timeout: 60_000 },
	)

	test(
		'10K formula cells recalculate and verify results',
		() => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const n = 10_000
			for (let r = 0; r < n; r++) {
				sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
				sheet.cells.set(r, 1, {
					value: EMPTY,
					formula: `A${r + 1}*2`,
					styleId: sid,
				})
			}
			const result = recalculate(wb, defaultCalcContext())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
			expect(sheet.cells.get(n - 1, 1)?.value).toEqual(numberValue(n * 2))
		},
		{ timeout: 30_000 },
	)

	test(
		'10K SUM formulas recalc completes in reasonable time',
		() => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const n = 10_000
			for (let r = 0; r < n; r++) {
				sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
				sheet.cells.set(r, 1, {
					value: EMPTY,
					formula: `SUM(A1:A${r + 1})`,
					styleId: sid,
				})
			}
			const start = performance.now()
			const result = recalculate(wb, defaultCalcContext())
			const duration = performance.now() - start
			expect(result.errors).toEqual([])
			const expectedSum = (n * (n + 1)) / 2
			expect(sheet.cells.get(n - 1, 1)?.value).toEqual(numberValue(expectedSum))
			expect(duration).toBeLessThan(10_000)
		},
		{ timeout: 15_000 },
	)
})
