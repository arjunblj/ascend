import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { parseFormula } from '@ascend/formulas'
import { numberValue } from '@ascend/schema'
import { defaultCalcContext } from './calc-context.ts'
import { compileFormula, evaluateCompiled } from './compiled-eval.ts'

const sid = 0 as StyleId

function makeWorkbook() {
	const wb = createWorkbook()
	wb.addSheet('Sheet1')
	return wb
}

describe('compiled numeric fast path', () => {
	test('IF with comparison condition and numeric branches stays on numeric fast path', () => {
		const wb = makeWorkbook()
		const sheet = wb.sheets[0]
		if (!sheet) throw new Error('missing sheet')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })

		const parsed = parseFormula('IF(A1>0,A1*2,0)')
		if (!parsed.ok) throw new Error('parse failed')
		const compiled = compileFormula(parsed.value)
		if (!compiled) throw new Error('compile failed')

		expect(compiled.numericOnly).toBe(true)
		expect(
			evaluateCompiled(compiled, {
				workbook: wb,
				calcContext: defaultCalcContext(),
				sheetIndex: 0,
				row: 0,
				col: 1,
			}),
		).toEqual(numberValue(10))
	})

	test('top-level comparisons do not claim numeric fast path', () => {
		const parsed = parseFormula('A1>0')
		if (!parsed.ok) throw new Error('parse failed')
		const compiled = compileFormula(parsed.value)
		if (!compiled) throw new Error('compile failed')

		expect(compiled.numericOnly).toBe(false)
	})

	test('numeric fast path supports AND conditions inside IF', () => {
		const wb = makeWorkbook()
		const sheet = wb.sheets[0]
		if (!sheet) throw new Error('missing sheet')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })

		const parsed = parseFormula('IF(AND(A1>0,B1>0),A1+B1,0)')
		if (!parsed.ok) throw new Error('parse failed')
		const compiled = compileFormula(parsed.value)
		if (!compiled) throw new Error('compile failed')

		expect(compiled.numericOnly).toBe(true)
		expect(
			evaluateCompiled(compiled, {
				workbook: wb,
				calcContext: defaultCalcContext(),
				sheetIndex: 0,
				row: 0,
				col: 2,
			}),
		).toEqual(numberValue(8))
	})

	test('repeated same-cell arithmetic compiles and evaluates correctly', () => {
		const wb = makeWorkbook()
		const sheet = wb.sheets[0]
		if (!sheet) throw new Error('missing sheet')
		sheet.cells.set(0, 0, { value: numberValue(7), formula: null, styleId: sid })

		const parsed = parseFormula('A1+A1')
		if (!parsed.ok) throw new Error('parse failed')
		const compiled = compileFormula(parsed.value)
		if (!compiled) throw new Error('compile failed')

		expect(
			evaluateCompiled(compiled, {
				workbook: wb,
				calcContext: defaultCalcContext(),
				sheetIndex: 0,
				row: 0,
				col: 1,
			}),
		).toEqual(numberValue(14))
	})

	test('constant IF conditions fold during compilation without changing results', () => {
		const wb = makeWorkbook()
		const sheet = wb.sheets[0]
		if (!sheet) throw new Error('missing sheet')
		sheet.cells.set(0, 0, { value: numberValue(11), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(99), formula: null, styleId: sid })

		const parsed = parseFormula('IF(TRUE,A1*2,B1*3)')
		if (!parsed.ok) throw new Error('parse failed')
		const compiled = compileFormula(parsed.value)
		if (!compiled) throw new Error('compile failed')

		expect(
			evaluateCompiled(compiled, {
				workbook: wb,
				calcContext: defaultCalcContext(),
				sheetIndex: 0,
				row: 0,
				col: 2,
			}),
		).toEqual(numberValue(22))
	})
})
