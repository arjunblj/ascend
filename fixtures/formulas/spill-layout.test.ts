import { describe, expect, test } from 'bun:test'
import { createWorkbook, type StyleId } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import type { CellValue } from '@ascend/schema'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'

const sid = 0 as StyleId

function grid(
	wb: ReturnType<typeof createWorkbook>,
	sheetName: string,
	startRow: number,
	startCol: number,
	rows: number,
	cols: number,
): CellValue[][] {
	const sheet = wb.getSheet(sheetName)
	if (!sheet) throw new Error(`Sheet ${sheetName} not found`)
	const result: CellValue[][] = []
	for (let r = 0; r < rows; r++) {
		const row: CellValue[] = []
		for (let c = 0; c < cols; c++) {
			row.push(sheet.cells.readValue(startRow + r, startCol + c))
		}
		result.push(row)
	}
	return result
}

describe('full spill layout', () => {
	test('SEQUENCE(5) spills 5 cells vertically', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(5)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		const values = grid(wb, 'Sheet1', 0, 0, 5, 1).map((r) => r[0])
		expect(values).toEqual([
			numberValue(1),
			numberValue(2),
			numberValue(3),
			numberValue(4),
			numberValue(5),
		])
	})

	test('SEQUENCE(2,3) spills 2×3 grid', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(2,3)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		const values = grid(wb, 'Sheet1', 0, 0, 2, 3)
		expect(values).toEqual([
			[numberValue(1), numberValue(2), numberValue(3)],
			[numberValue(4), numberValue(5), numberValue(6)],
		])
	})

	test('SORT spills sorted column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'SORT(A1:A3)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		const values = grid(wb, 'Sheet1', 0, 2, 3, 1).map((r) => r[0])
		expect(values).toEqual([numberValue(10), numberValue(20), numberValue(30)])
	})

	test('UNIQUE removes duplicates and spills', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('c'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'UNIQUE(A1:A4)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		const values = grid(wb, 'Sheet1', 0, 2, 3, 1).map((r) => r[0])
		expect(values).toEqual([stringValue('a'), stringValue('b'), stringValue('c')])
	})

	test('FILTER with explicit boolean array returns matching values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'FILTER(A1:A3,{FALSE;FALSE;TRUE})',
			styleId: sid,
		})
		recalculate(wb, defaultCalcContext())
		expect(sheet.cells.readValue(0, 2)).toEqual(numberValue(3))
	})

	test('spill blocked by existing cell produces #SPILL!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: sid })
		recalculate(wb, defaultCalcContext())
		const anchor = sheet.cells.readValue(0, 0)
		expect(anchor.kind).toBe('error')
		if (anchor.kind === 'error') expect(anchor.value).toBe('#SPILL!')
	})
})
