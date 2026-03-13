import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { parseFormula } from '@ascend/formulas'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { defaultCalcContext } from './calc-context.ts'
import { clearCodegenCache, codegenFormula } from './codegen.ts'
import type { EvalContext } from './evaluator.ts'

const sid = 0 as StyleId

function makeCtx(
	wb: ReturnType<typeof createWorkbook>,
	sheetIndex = 0,
	row = 0,
	col = 0,
): EvalContext {
	return { workbook: wb, calcContext: defaultCalcContext(), sheetIndex, row, col }
}

function codegenEval(formula: string, wb: ReturnType<typeof createWorkbook>, row = 0, col = 0) {
	const parsed = parseFormula(formula)
	if (!parsed.ok) throw new Error(`Parse failed: ${formula}`)
	const fn = codegenFormula(formula, parsed.value)
	if (!fn) return null
	return fn(makeCtx(wb, 0, row, col))
}

describe('codegen', () => {
	test('arithmetic: A1+B1*2', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(5), formula: null, styleId: sid })

		const result = codegenEval('A1+B1*2', wb)
		expect(result).toEqual(numberValue(20))
	})

	test('subtraction: A1-B1', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })

		expect(codegenEval('A1-B1', wb)).toEqual(numberValue(7))
	})

	test('division by zero returns #DIV/0!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })

		expect(codegenEval('A1/B1', wb)).toEqual(errorValue('#DIV/0!'))
	})

	test('power: A1^2', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })

		expect(codegenEval('A1^2', wb)).toEqual(numberValue(9))
	})

	test('unary negation: -A1', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })

		expect(codegenEval('-A1', wb)).toEqual(numberValue(-42))
	})

	test('cell ref to empty returns EMPTY', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const parsed = parseFormula('A1')
		if (!parsed.ok) throw new Error('parse failed')
		const fn = codegenFormula('A1-test-empty', parsed.value)
		if (!fn) throw new Error('codegen returned null')
		const result = fn(makeCtx(wb))
		expect(result).toEqual(EMPTY)
	})

	test('comparisons', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: sid })

		expect(codegenEval('A1>B1', wb)).toEqual(booleanValue(false))
		expect(codegenEval('A1<B1', wb)).toEqual(booleanValue(true))
		expect(codegenEval('A1=B1', wb)).toEqual(booleanValue(false))
		expect(codegenEval('A1<>B1', wb)).toEqual(booleanValue(true))
		expect(codegenEval('A1>=10', wb)).toEqual(booleanValue(true))
		expect(codegenEval('A1<=9', wb)).toEqual(booleanValue(false))
	})

	test('IF function', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })

		expect(codegenEval('IF(A1>5,1,0)', wb)).toEqual(numberValue(1))
		expect(codegenEval('IF(A1>20,1,0)', wb)).toEqual(numberValue(0))
	})

	test('IF without else branch', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: booleanValue(false), formula: null, styleId: sid })

		expect(codegenEval('IF(A1,1)', wb)).toEqual(booleanValue(false))
	})

	test('IFERROR function', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })

		expect(codegenEval('IFERROR(A1/B1,0)', wb)).toEqual(numberValue(0))
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		clearCodegenCache()
		expect(codegenEval('IFERROR(A1/B1,0)', wb)).toEqual(numberValue(5))
	})

	test('IFNA function', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: errorValue('#N/A'), formula: null, styleId: sid })

		expect(codegenEval('IFNA(A1,99)', wb)).toEqual(numberValue(99))
		sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
		clearCodegenCache()
		expect(codegenEval('IFNA(A1,99)', wb)).toEqual(numberValue(42))
	})

	test('concatenation with &', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue(' World'), formula: null, styleId: sid })

		expect(codegenEval('A1&B1', wb)).toEqual(stringValue('Hello World'))
	})

	test('cross-sheet ref', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s2.cells.set(0, 0, { value: numberValue(77), formula: null, styleId: sid })

		expect(codegenEval('Sheet2!A1+1', wb)).toEqual(numberValue(78))
	})

	test('cross-sheet ref to nonexistent sheet returns #REF!', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		expect(codegenEval('NoSheet!A1+1', wb)).toEqual(errorValue('#REF!'))
	})

	test('error propagation in arithmetic', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: errorValue('#VALUE!'), formula: null, styleId: sid })

		expect(codegenEval('A1+1', wb)).toEqual(errorValue('#VALUE!'))
	})

	test('returns null for unsupported formulas (range functions)', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const parsed = parseFormula('SUM(A1:A10)')
		if (!parsed.ok) throw new Error('parse failed')
		const fn = codegenFormula('SUM(A1:A10)', parsed.value)
		expect(fn).toBeNull()
	})

	test('caching returns same function for same formula text', () => {
		clearCodegenCache()
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const parsed = parseFormula('A1+B1')
		if (!parsed.ok) throw new Error('parse failed')

		const fn1 = codegenFormula('A1+B1', parsed.value)
		const fn2 = codegenFormula('A1+B1', parsed.value)
		expect(fn1).toBe(fn2)
	})

	test('percentage operator', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(50), formula: null, styleId: sid })

		expect(codegenEval('A1%', wb)).toEqual(numberValue(0.5))
	})

	test('string literal', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		expect(codegenEval('"hello"', wb)).toEqual(stringValue('hello'))
	})

	test('boolean literal', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		expect(codegenEval('TRUE', wb)).toEqual(booleanValue(true))
	})

	test('nested IF', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(15), formula: null, styleId: sid })

		expect(codegenEval('IF(A1>20,"big",IF(A1>10,"med","small"))', wb)).toEqual(stringValue('med'))
	})
})
