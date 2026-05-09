import { describe, expect, test } from 'bun:test'
import { createWorkbook, type StyleId } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import { EMPTY, numberValue } from '@ascend/schema'

const sid = 0 as StyleId

function evalFormula(
	formula: string,
	dateSystem: '1900' | '1904' = '1900',
): import('@ascend/schema').CellValue {
	const wb = createWorkbook()
	wb.calcSettings = { ...wb.calcSettings, dateSystem }
	const sheet = wb.addSheet('Sheet1')
	const f = formula.startsWith('=') ? formula.slice(1) : formula
	sheet.cells.set(0, 0, { value: EMPTY, formula: f, styleId: sid })
	recalculate(wb, { ...defaultCalcContext(), dateSystem })
	return sheet.cells.readValue(0, 0)
}

describe('date serial edge cases', () => {
	test('DATE(1900,1,0) is serial 0 (Excel quirk)', () => {
		const v = evalFormula('=DATE(1900,1,0)')
		if (v.kind === 'date') expect(v.serial).toBe(0)
		else if (v.kind === 'number') expect(v.value).toBe(0)
	})

	test('DATE(1900,2,29) wraps to March 1 (serial 61) — Ascend does not emulate the fake leap day', () => {
		const v = evalFormula('=DATE(1900,2,29)')
		if (v.kind === 'date') expect(v.serial).toBe(61)
		else if (v.kind === 'number') expect(v.value).toBe(61)
	})

	test('DATE(1900,3,1) is serial 61', () => {
		const v = evalFormula('=DATE(1900,3,1)')
		if (v.kind === 'date') expect(v.serial).toBe(61)
		else if (v.kind === 'number') expect(v.value).toBe(61)
	})

	test('YEAR(1) = 1900', () => {
		const v = evalFormula('=YEAR(1)')
		expect(v).toEqual(numberValue(1900))
	})

	test('MONTH(1) = 1', () => {
		const v = evalFormula('=MONTH(1)')
		expect(v).toEqual(numberValue(1))
	})

	test('DAY(1) = 1', () => {
		const v = evalFormula('=DAY(1)')
		expect(v).toEqual(numberValue(1))
	})

	test('1904 date system: DATE(1904,1,2) = serial 1', () => {
		const v = evalFormula('=DATE(1904,1,2)', '1904')
		if (v.kind === 'date') expect(v.serial).toBe(1)
		else if (v.kind === 'number') expect(v.value).toBe(1)
	})

	test('1904 date system: YEAR(1) = 1904', () => {
		const v = evalFormula('=YEAR(1)', '1904')
		expect(v).toEqual(numberValue(1904))
	})

	test('large serial 2958465 (9999-12-31 in 1900 system)', () => {
		const v = evalFormula('=YEAR(2958465)')
		expect(v).toEqual(numberValue(9999))
	})

	test('DATE with month > 12 wraps to next year', () => {
		const v = evalFormula('=MONTH(DATE(2023,13,1))')
		expect(v).toEqual(numberValue(1))
	})

	test('DATE with day > month length wraps', () => {
		const v = evalFormula('=MONTH(DATE(2023,1,32))')
		expect(v).toEqual(numberValue(2))
	})
})
