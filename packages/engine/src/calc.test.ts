import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, createWorkbook } from '@ascend/core'
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

	test('sheet-scoped defined name shadows workbook-scoped name', () => {
		const wb = createWorkbook()
		const sheet1 = wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		wb.definedNames.set('Rate', '0.1')
		wb.definedNames.set('Rate', '0.2', { kind: 'sheet', sheetId: sheet2.id })
		sheet1.cells.set(0, 0, { value: EMPTY, formula: 'Rate*100', styleId: sid })
		sheet2.cells.set(0, 0, { value: EMPTY, formula: 'Rate*100', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet1.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(sheet2.cells.get(0, 0)?.value).toEqual(numberValue(20))
	})

	test('qualified local defined names can be referenced from another sheet', () => {
		const wb = createWorkbook()
		const sheet1 = wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		const calc = wb.addSheet('Calc')
		wb.definedNames.set('Budget', '10', { kind: 'sheet', sheetId: sheet1.id })
		wb.definedNames.set('Budget', '20', { kind: 'sheet', sheetId: sheet2.id })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'Sheet1!Budget+Sheet2!Budget', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(calc.cells.get(0, 0)?.value).toEqual(numberValue(30))
	})

	test('structured references can sum a table column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Player'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Mina'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Noah'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(12), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Scores',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Player' }, { name: 'Score' }],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Scores[Score])', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(22))
	})

	test('current-row structured references resolve within a table body', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: EMPTY, formula: '[@Qty]*[@Price]', styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: '[@Qty]*[@Price]', styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
			columns: [{ name: 'Qty' }, { name: 'Price' }, { name: 'Total' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(20))
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
