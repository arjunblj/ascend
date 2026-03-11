import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { parseFormula } from '@ascend/formulas'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { recalculate } from './calc.ts'
import { defaultCalcContext } from './calc-context.ts'
import type { EvalContext } from './evaluator.ts'
import { evaluate } from './evaluator.ts'

const sid = 0 as StyleId

function makeEvalCtx(
	wb: ReturnType<typeof createWorkbook>,
	sheetIndex: number,
	row: number,
	col: number,
): EvalContext {
	return {
		workbook: wb,
		calcContext: defaultCalcContext(),
		sheetIndex,
		row,
		col,
	}
}

function evalFormula(
	formula: string,
	wb: ReturnType<typeof createWorkbook>,
	sheetIndex: number,
	row: number,
	col: number,
) {
	const parsed = parseFormula(formula)
	if (!parsed.ok) throw new Error(`Parse failed: ${formula}`)
	const ctx = makeEvalCtx(wb, sheetIndex, row, col)
	return evaluate(parsed.value, ctx)
}

describe('evaluator', () => {
	describe('binary operations', () => {
		test('Number + Number', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(8))
		})

		test('Number - Number', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1-B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(7))
		})

		test('Number * Number', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(7), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1*B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(28))
		})

		test('Number / Number', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(15), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1/B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(5))
		})

		test('Division by zero produces #DIV/0!', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1/B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#DIV/0!'))
		})

		test('Exponentiation ^', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1^B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(1024))
		})

		test('String concatenation with &', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: stringValue(' World'), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1&B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('Hello World'))
		})

		test('Error propagation - left error', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1/B1', styleId: sid })
			sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#DIV/0!'))
			expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#DIV/0!'))
		})

		test('Error propagation - right error', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: errorValue('#N/A'), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#N/A'))
		})

		test('Type coercion - string "5" + number 3', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: stringValue('5'), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(8))
		})
	})

	describe('comparison edge cases', () => {
		test('String comparison is case-insensitive', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: stringValue('ABC'), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: stringValue('abc'), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1=B1', styleId: sid })
			sheet.cells.set(2, 0, { value: EMPTY, formula: 'A1<>B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(booleanValue(true))
			expect(sheet.cells.get(2, 0)?.value).toEqual(booleanValue(false))
		})

		test('Boolean vs number comparison', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: booleanValue(true), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1=B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(booleanValue(true))
		})

		test('Empty vs number comparison', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: EMPTY, formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1=B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(booleanValue(true))
		})

		test('Comparison operators < > <= >=', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1<B1', styleId: sid })
			sheet.cells.set(2, 0, { value: EMPTY, formula: 'A1>B1', styleId: sid })
			sheet.cells.set(3, 0, { value: EMPTY, formula: 'A1<=B1', styleId: sid })
			sheet.cells.set(4, 0, { value: EMPTY, formula: 'A1>=B1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(booleanValue(true))
			expect(sheet.cells.get(2, 0)?.value).toEqual(booleanValue(false))
			expect(sheet.cells.get(3, 0)?.value).toEqual(booleanValue(true))
			expect(sheet.cells.get(4, 0)?.value).toEqual(booleanValue(false))
		})
	})

	describe('unary operations', () => {
		test('Unary +', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(-7), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: '+A1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(-7))
		})

		test('Unary -', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: '-A1', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(-42))
		})

		test('Unary %', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(50), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1%', styleId: sid })

			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(0.5))
		})
	})

	describe('direct evaluate calls', () => {
		test('evaluate binary via workbook context', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(25), formula: null, styleId: sid })

			const result = evalFormula('A1/B1', wb, 0, 0, 2)
			expect(result).toEqual(numberValue(4))
		})

		test('evaluate division by zero returns #DIV/0!', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })

			const result = evalFormula('A1/B1', wb, 0, 0, 2)
			expect(result).toEqual(errorValue('#DIV/0!'))
		})

		test('evaluate concatenation', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: stringValue('x'), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: stringValue('y'), formula: null, styleId: sid })

			const result = evalFormula('A1&B1', wb, 0, 0, 2)
			expect(result).toEqual(stringValue('xy'))
		})
	})
})
