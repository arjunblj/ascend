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

	test('dirty refs limit recalculation scope to affected dependency subgraph', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'C1+1', styleId: sid })

		recalculate(wb, makeCtx())
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(11))
		expect(result.changed).toContain('Sheet1!B1')
		expect(result.changed).not.toContain('Sheet1!D1')
	})

	test('dirty recalc only reports parse errors inside the affected subgraph', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SUM((', styleId: sid })

		recalculate(wb, makeCtx())
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
		expect(result.errors).toEqual([])
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

	test('whole-column references can aggregate used cells in a column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'SUM(A:A)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(10))
	})

	test('dirty recalc updates formulas using whole-column references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'SUM(A:A)', styleId: sid })

		recalculate(wb, makeCtx())
		sheet.cells.set(2, 0, { value: numberValue(10), formula: null, styleId: sid })

		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A3'] })
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(14))
	})

	test('whole-row references can aggregate used cells in a row', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: EMPTY, formula: 'SUM(2:2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(10))
	})

	test('ROWS and COLUMNS treat whole-dimension refs as full-sheet ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'ROWS(A:A)', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'COLUMNS(1:1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1_048_576))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(16_384))
	})

	test('bare contiguous ranges spill under array semantics', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1:A3', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(3))
	})

	test('implicit intersection uses the formula row for single-column ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: '@A:A', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
	})

	test('implicit intersection on whole-column refs can return empty cells outside usedRange', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(9, 1, { value: EMPTY, formula: '@A:A', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(9, 1)?.value).toEqual(EMPTY)
	})

	test('implicit intersection uses the formula column for single-row ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: '@1:1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
	})

	test('SUM supports union references inside a parenthesized single argument', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SUM((A1:A2,C1:C2))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
	})

	test('SUM supports intersection references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'SUM(A:B 2:2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(10))
	})

	test('COUNTA, COUNTBLANK, and PRODUCT support union references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'COUNTA((A1:A2,C1:C2))', styleId: sid })
		sheet.cells.set(1, 4, { value: EMPTY, formula: 'COUNTBLANK((A1:A2,C1:C2))', styleId: sid })
		sheet.cells.set(2, 4, { value: EMPTY, formula: 'PRODUCT((A1:A2,C1:C2))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(24))
	})

	test('SUMPRODUCT returns #VALUE! for mismatched range shapes', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUMPRODUCT(A1:A2,B1:B1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('SUMIFS returns #VALUE! for mismatched criteria shapes', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUMIFS(A1:A2,B1:B1,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#VALUE!'))
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

	test('COUNTIF supports wildcard criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('North'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Northeast'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('South'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('Northwest'), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'COUNTIF(A1:A4,"North*")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(3))
	})

	test('COUNTIF blank and nonblank criteria match Excel semantics', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue(''), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('value'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'COUNTIF(A1:A3,"")', styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'COUNTIF(A1:A3,"<>")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(1))
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

	test('RAND is deterministic per cell and seed', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'RAND()', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'RAND()', styleId: sid })

		const ctx = makeCtx({ randomSeed: 7 })
		recalculate(wb, ctx)
		const a1 = sheet.cells.get(0, 0)?.value
		const a2 = sheet.cells.get(1, 0)?.value
		recalculate(wb, ctx)
		expect(sheet.cells.get(0, 0)?.value).toEqual(a1)
		expect(sheet.cells.get(1, 0)?.value).toEqual(a2)
		expect(a1).not.toEqual(a2)
	})

	test('wrong function arity returns #VALUE!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'ABS()', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'ABS(1,2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('exponentiation precedence matches spreadsheet semantics', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '2^3^2', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: '-2^2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(512))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(-4))
	})

	test('INDIRECT resolves A1-style ranges inside aggregate functions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUM(INDIRECT("A1:A2"))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(5))
	})

	test('dirty recalc re-evaluates INDIRECT formulas through the volatile path', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'INDIRECT("A1")*2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(4))

		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
	})

	test('INDIRECT resolves R1C1-style references when requested', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(11), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'INDIRECT("R[1]C[-1]",FALSE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(11))
	})

	test('OFFSET returns shifted ranges to aggregators', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1,1,0,2,1))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(11))
	})

	test('dirty recalc re-evaluates OFFSET formulas through the volatile path', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1,1,0,1,1))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(5))

		sheet.cells.set(1, 0, { value: numberValue(9), formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A2'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(9))
	})

	test('SEQUENCE spills vertically into neighboring cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))
	})

	test('blocked spill returns #SPILL! in the anchor cell', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('blocker'))
	})

	test('dirty recalc clears stale spill cells when a spill shrinks', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 4; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
			sheet.cells.set(r, 1, {
				value: { kind: 'boolean', value: true },
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A4,B1:B4,0)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(4))

		sheet.cells.set(2, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })

		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!B3:B4'] })
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 2)).toBeUndefined()
		expect(sheet.cells.get(3, 2)).toBeUndefined()
	})

	test('spill operator resolves a spilled range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1#)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
	})

	test('dynamic array functions spill sorted results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SORT(A1:A3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(30))
	})

	test('TRANSPOSE spills values across columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'TRANSPOSE(A1:A2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(5))
	})

	test('TAKE spills the requested leading rows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'TAKE(A1:A3,2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
	})

	test('DROP spills the remaining rows after dropping the requested prefix', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'DROP(A1:A3,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(3))
	})

	test('CHOOSECOLS can spill multiple selected columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'CHOOSECOLS(A1:C1,1,3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(3))
	})

	test('HSTACK spills joined ranges across columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'HSTACK(A1:A2,B1:B2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(20))
	})

	test('SORTBY supports multiple sort keys', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'SORTBY(A1:B3,A1:A3,1,B1:B3,-1)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('a'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue('b'))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 4)?.value).toEqual(stringValue('a'))
	})

	test('UNIQUE can compare columns when by_col is true', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'UNIQUE(A1:C2,TRUE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(4))
	})

	test('TOCOL supports ignore blanks and scan-by-column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'TOCOL(A1:B2,1,TRUE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(3))
	})

	test('TOCOL returns #CALC! when ignore rules remove every value', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'TOCOL(A1:B2,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#CALC!'))
	})

	test('TOROW supports ignore errors and scan-by-column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: errorValue('#N/A'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'TOROW(A1:B2,2,TRUE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(3))
	})

	test('FILTER returns #VALUE! for include shape mismatch and propagates include errors', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'FILTER(A1:B2,{TRUE})', styleId: sid })
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'FILTER(A1:A2,{#N/A;TRUE})', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(errorValue('#N/A'))
	})

	test('TAKE and DROP return #CALC! for zero extents', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'TAKE(A1:A2,0)', styleId: sid })
		sheet.cells.set(1, 2, { value: EMPTY, formula: 'DROP(A1:A2,0)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#CALC!'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(errorValue('#CALC!'))
	})

	test('CHOOSECOLS and CHOOSEROWS support negative indices from the end', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'CHOOSECOLS(A1:C2,-1,-2)', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'CHOOSEROWS(A1:C2,-1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(6))
	})

	test('XMATCH spills results for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'XMATCH({20;30},A1:A3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(3))
	})

	test('MATCH spills results for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('North'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('South'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('West'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'MATCH({"South";"West"},A1:A3,0)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(3))
	})

	test('XLOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'XLOOKUP({2;3},A1:A3,B1:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Three'))
	})

	test('VLOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'VLOOKUP({2;3},A1:B3,2,FALSE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Three'))
	})

	test('HLOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'HLOOKUP({2,3},A1:C2,2,FALSE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(0, 5)?.value).toEqual(stringValue('Three'))
	})

	test('LOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'LOOKUP({2;3},A1:A3,B1:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Three'))
	})

	test('TEXTBEFORE and TEXTAFTER support modern text slicing formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTBEFORE("alpha-beta-gamma","-",2)',
			styleId: sid,
		})
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'TEXTAFTER("alpha-beta-gamma","-",2)',
			styleId: sid,
		})
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('alpha-beta'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('gamma'))
	})

	test('TEXTBEFORE and TEXTAFTER support case-insensitive matching and if_not_found', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTBEFORE("Alpha-beta","BETA",1,1,0,"missing")',
			styleId: sid,
		})
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'TEXTAFTER("Alpha","-")', styleId: sid })
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('Alpha-'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#N/A'))
	})

	test('TEXTSPLIT spills rows and columns from delimited text', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTSPLIT("a,b;c,d",",",";")',
			styleId: sid,
		})
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('a'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('b'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('c'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('d'))
	})

	test('TEXTSPLIT supports ignore_empty and pad_with', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTSPLIT("a,,c",",",,TRUE,0,"pad")',
			styleId: sid,
		})
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('a'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('c'))
		expect(sheet.cells.get(0, 2)).toBeUndefined()
	})

	test('external workbook references currently evaluate to #REF!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '[Book.xlsx]Sheet1!A1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('3D sheet-span references aggregate across contiguous sheets', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const s3 = wb.addSheet('Sheet3')
		const calc = wb.addSheet('Calc')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		s3.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'SUM(Sheet1:Sheet3!A1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(calc.cells.get(0, 0)?.value).toEqual(numberValue(6))
	})

	test('bare 3D sheet-span references remain unsupported in scalar contexts', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const calc = wb.addSheet('Calc')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'Sheet1:Sheet2!A1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(calc.cells.get(0, 0)?.value).toEqual(errorValue('#VALUE!'))
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

describe('iterative calculation', () => {
	test('simple circular ref converges with iterative calc enabled', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '0.5*B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: '0.5*A1+1', styleId: sid })

		const ctx = makeCtx({
			iterativeCalc: {
				enabled: true,
				maxIterations: 100,
				maxChange: 0.001,
			},
		})
		recalculate(wb, ctx)

		const a1 = sheet.cells.get(0, 0)?.value
		const b1 = sheet.cells.get(0, 1)?.value
		expect(a1?.kind).toBe('number')
		expect(b1?.kind).toBe('number')
		if (a1?.kind === 'number' && b1?.kind === 'number') {
			expect(a1.value).toBeGreaterThanOrEqual(0)
			expect(a1.value).toBeLessThanOrEqual(1)
			expect(b1.value).toBeGreaterThanOrEqual(1)
			expect(b1.value).toBeLessThanOrEqual(2)
		}
	})

	test('max iterations respected', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1+1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const ctx = makeCtx({
			iterativeCalc: {
				enabled: true,
				maxIterations: 5,
				maxChange: 0.001,
			},
		})
		recalculate(wb, ctx)

		const a1 = sheet.cells.get(0, 0)?.value
		const b1 = sheet.cells.get(0, 1)?.value
		expect(a1?.kind).toBe('number')
		expect(b1?.kind).toBe('number')
		if (a1?.kind === 'number' && b1?.kind === 'number') {
			expect(a1.value).toBeLessThanOrEqual(6)
			expect(b1.value).toBeLessThanOrEqual(6)
		}
	})

	test('without iterative calc, circular refs produce #REF!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#REF!'))
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
	})
})

describe('shared formula evaluation', () => {
	test('shared formulas A1:A10 = B1*2 shifted per row evaluate correctly', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 10; r++) {
			sheet.cells.set(r, 1, { value: numberValue(r + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		for (let r = 1; r < 10; r++) {
			sheet.cells.set(r, 0, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})
		}

		recalculate(wb, makeCtx())
		for (let r = 0; r < 10; r++) {
			const expected = (r + 1) * 2
			expect(sheet.cells.get(r, 0)?.value).toEqual(numberValue(expected))
		}
	})

	test('recalc after changing source cell updates shared formula results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(6))

		sheet.cells.set(0, 1, { value: numberValue(7), formula: null, styleId: sid })
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(14))
	})
})

describe('large range correctness', () => {
	test('SUM over 10K cells verifies exact result', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		let expectedSum = 0
		for (let i = 0; i < 10_000; i++) {
			const val = i + 1
			expectedSum += val
			sheet.cells.set(Math.floor(i / 100), i % 100, {
				value: numberValue(val),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(101, 0, {
			value: EMPTY,
			formula: 'SUM(A1:CV100)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(101, 0)?.value).toEqual(numberValue(expectedSum))
	})

	test('AVERAGE over 10K cells verifies correct average', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		let sum = 0
		for (let i = 0; i < 10_000; i++) {
			const val = i + 1
			sum += val
			sheet.cells.set(Math.floor(i / 100), i % 100, {
				value: numberValue(val),
				formula: null,
				styleId: sid,
			})
		}
		const expectedAvg = sum / 10_000
		sheet.cells.set(101, 0, {
			value: EMPTY,
			formula: 'AVERAGE(A1:CV100)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		const result = sheet.cells.get(101, 0)?.value
		expect(result?.kind).toBe('number')
		if (result?.kind === 'number') {
			expect(Math.abs(result.value - expectedAvg)).toBeLessThan(0.0001)
		}
	})

	test('cross-sheet SUM - 3 sheets, SUM(Sheet1!A1:A100) on Sheet2', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		let expectedSum = 0
		for (let r = 0; r < 100; r++) {
			const val = r + 1
			expectedSum += val
			s1.cells.set(r, 0, { value: numberValue(val), formula: null, styleId: sid })
		}
		s2.cells.set(0, 0, {
			value: EMPTY,
			formula: 'SUM(Sheet1!A1:A100)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(s2.cells.get(0, 0)?.value).toEqual(numberValue(expectedSum))
	})
})
