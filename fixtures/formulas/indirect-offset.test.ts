import { describe, expect, test } from 'bun:test'
import { createWorkbook, type StyleId } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'

const sid = 0 as StyleId

describe('INDIRECT/OFFSET edge cases', () => {
	test('INDIRECT basic A1 reference', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('A1'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'INDIRECT(B1)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		expect(sheet.cells.readValue(0, 2)).toEqual(numberValue(42))
	})

	test('INDIRECT cross-sheet reference', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Data')
		s2.cells.set(0, 0, { value: numberValue(99), formula: null, styleId: sid })
		s1.cells.set(0, 0, { value: EMPTY, formula: 'INDIRECT("Data!A1")', styleId: sid })
		recalculate(wb, defaultCalcContext())
		expect(s1.cells.readValue(0, 0)).toEqual(numberValue(99))
	})

	test('INDIRECT with R1C1 style (FALSE)', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(7), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'INDIRECT("R1C1", FALSE)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		expect(sheet.cells.readValue(1, 0)).toEqual(numberValue(7))
	})

	test('INDIRECT with invalid ref returns #REF!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'INDIRECT("ZZZZZ99999")', styleId: sid })
		recalculate(wb, defaultCalcContext())
		const v = sheet.cells.readValue(0, 0)
		expect(v.kind).toBe('error')
	})

	test('OFFSET basic', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'OFFSET(A1,1,1)', styleId: sid })
		recalculate(wb, defaultCalcContext())
		expect(sheet.cells.readValue(2, 2)).toEqual(numberValue(20))
	})

	test('OFFSET with height/width as SUM range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'SUM(OFFSET(A1,0,0,3,1))', styleId: sid })
		recalculate(wb, defaultCalcContext())
		expect(sheet.cells.readValue(3, 0)).toEqual(numberValue(6))
	})

	test('nested INDIRECT inside SUM', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(15), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'SUM(INDIRECT("A1:B1"))', styleId: sid })
		recalculate(wb, defaultCalcContext())
		expect(sheet.cells.readValue(1, 0)).toEqual(numberValue(20))
	})
})
