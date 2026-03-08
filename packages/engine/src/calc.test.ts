import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { dateToSerial } from '@ascend/formulas'
import { EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { recalculate } from './calc.ts'
import type { CalcContext } from './calc-context.ts'
import { defaultCalcContext } from './calc-context.ts'

const sid = 0 as StyleId

function makeCtx(overrides?: Partial<CalcContext>): CalcContext {
	return { ...defaultCalcContext(), ...overrides }
}

describe('recalculate', () => {
	test('simple SUM formula', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		const cell = sheet.cells.get(2, 0)
		expect(cell?.value).toEqual(numberValue(3))
	})

	test('chain of dependent formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1*2', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+5', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(25))
	})

	test('recalculation after value change reports changed cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+10', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(11))

		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(15))
		expect(result.changed.length).toBeGreaterThan(0)
	})

	test('multi-sheet reference', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Data')
		const s2 = wb.addSheet('Summary')
		s1.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: EMPTY, formula: 'Data!A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(s2.cells.get(0, 0)?.value).toEqual(numberValue(42))
	})

	test('deterministic NOW via CalcContext', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'TODAY()', styleId: sid })

		const fixedDate = new Date(2025, 0, 15)
		const ctx = makeCtx({ today: fixedDate, now: fixedDate })
		recalculate(wb, ctx)

		const cell = sheet.cells.get(0, 0)
		const expectedSerial = dateToSerial(2025, 1, 15)
		expect(cell?.value).toEqual(numberValue(expectedSerial))
	})

	test('error propagation through formula chain', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1/B1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#DIV/0!'))
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('circular reference produces error', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
	})

	test('string concatenation formula', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue(' World'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1&B1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('Hello World'))
	})

	test('formula parse error is reported', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '=INVALID(((', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors.length).toBeGreaterThan(0)
		expect(result.errors[0]?.error.code).toBe('FORMULA_PARSE_ERROR')
	})

	test('duration is tracked', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const result = recalculate(wb, makeCtx())
		expect(result.duration).toBeGreaterThanOrEqual(0)
	})

	test('arithmetic operators evaluate correctly', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1-B1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A1*B1', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'A1^2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(100))
	})
})
