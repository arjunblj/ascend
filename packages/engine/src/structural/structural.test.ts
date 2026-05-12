import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook, parseRange } from '@ascend/core'
import { parseFormula } from '@ascend/formulas'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import { rewriteFormulaAstForShift, rewriteFormulaTextForShift } from './formula-rewrite.ts'
import { shiftSqref } from './ref-shift.ts'
import { sortSheetRange } from './sort-range.ts'

const sid = 0 as StyleId

describe('structural helpers', () => {
	test('rewriteFormulaTextForShift updates row refs on the target sheet', () => {
		expect(rewriteFormulaTextForShift('SUM(A1:A2)', 'Sheet1', 'Sheet1', 'row', 1, 1)).toBe(
			'SUM(A1:A3)',
		)
	})

	test('rewriteFormulaAstForShift preserves unaffected range refs without allocation', () => {
		const parsed = parseFormula('SUM(A1:A2)')
		expect(parsed.ok).toBe(true)
		if (!parsed.ok) return

		const rewritten = rewriteFormulaAstForShift(parsed.value, 'Sheet1', 'Sheet1', 'row', 10, 1)

		expect(rewritten).toBe(parsed.value)
	})

	test('shiftSqref shifts multipart row ranges', () => {
		expect(shiftSqref('A1:B2 C3:C4', 'row', 1, 2)).toBe('A1:B4 C5:C6')
	})

	test('sortSheetRange requires at least one sort key', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })

		const result = sortSheetRange(wb, sheet, parseRange('A1:A2'), [])
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect(result.error.code).toBe('VALIDATION_ERROR')
		expect(result.error.message).toContain('at least one sort key')
	})

	test('sortSheetRange rejects unknown header names', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Debt'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: sid })

		const result = sortSheetRange(wb, sheet, parseRange('A1:B3'), [{ column: 'Missing' }])
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect(result.error.code).toBe('VALIDATION_ERROR')
		expect(result.error.message).toContain('Unknown sort column')
	})

	test('sortSheetRange sorts stable ascending by a header name', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Debt'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('Equity'), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: numberValue(20), formula: null, styleId: sid })

		const result = sortSheetRange(wb, sheet, parseRange('A1:B4'), [{ column: 'Value' }])
		expect(result.ok).toBe(true)
		if (!result.ok) return
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('Debt'))
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('Cash'))
		expect(sheet.cells.get(3, 0)?.value).toEqual(stringValue('Equity'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('Name'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('Value'))
		expect(sheet.cells.get(10, 10)?.value ?? EMPTY).toEqual(EMPTY)
	})
})
