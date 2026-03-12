import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook, parseA1 } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import type { CellValue } from '@ascend/schema'
import { booleanValue, EMPTY, numberValue, stringValue } from '@ascend/schema'

const sid = 0 as StyleId

function evalFormula(
	formula: string,
	cells: Record<string, string | number | boolean> = {},
): CellValue {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	for (const [a1, raw] of Object.entries(cells)) {
		const { row, col } = parseA1(a1)
		const value =
			typeof raw === 'number'
				? numberValue(raw)
				: typeof raw === 'boolean'
					? booleanValue(raw)
					: stringValue(String(raw))
		sheet.cells.set(row, col, { value, formula: null, styleId: sid })
	}
	const f = formula.startsWith('=') ? formula.slice(1) : formula
	sheet.cells.set(99, 0, { value: EMPTY, formula: f, styleId: sid })
	recalculate(wb, defaultCalcContext())
	return sheet.cells.get(99, 0)?.value ?? EMPTY
}

function expectNum(result: CellValue, expected: number, tolerance = 1e-6): void {
	expect(result.kind).toBe('number')
	if (result.kind === 'number') {
		expect(Math.abs(result.value - expected)).toBeLessThanOrEqual(tolerance)
	}
}

describe('Excel conformance', () => {
	describe('financial', () => {
		test('PMT zero interest', () => {
			expectNum(evalFormula('PMT(0, 10, -1000)'), 100)
		})

		test('PMT with interest', () => {
			expectNum(evalFormula('PMT(0.05/12, 360, -100000)'), 536.82, 0.01)
		})

		test('FV zero interest', () => {
			expectNum(evalFormula('FV(0, 10, -100)'), 1000)
		})

		test('FV with interest', () => {
			expectNum(evalFormula('FV(0.06/12, 120, -200)'), 32775.87, 1.0)
		})

		test('PV annuity', () => {
			expectNum(evalFormula('PV(0.08, 5, -100)'), 399.27, 0.01)
		})

		test('NPV positive cash flows', () => {
			expectNum(evalFormula('NPV(0.1, A1, A2, A3)', { A1: 100, A2: 200, A3: 300 }), 481.59, 0.01)
		})

		test('IRR simple return', () => {
			expectNum(evalFormula('IRR(A1:A2)', { A1: -100, A2: 110 }), 0.1, 0.001)
		})
	})

	describe('statistical', () => {
		test('MEDIAN odd count', () => {
			expectNum(evalFormula('MEDIAN(A1:A5)', { A1: 3, A2: 1, A3: 5, A4: 2, A5: 4 }), 3)
		})

		test('MEDIAN even count', () => {
			expectNum(evalFormula('MEDIAN(A1:A4)', { A1: 1, A2: 2, A3: 3, A4: 4 }), 2.5)
		})

		test('STDEV sample', () => {
			expectNum(
				evalFormula('STDEV(A1:A8)', {
					A1: 2,
					A2: 4,
					A3: 4,
					A4: 4,
					A5: 5,
					A6: 5,
					A7: 7,
					A8: 9,
				}),
				2.138,
				0.001,
			)
		})

		test('CORREL perfect positive', () => {
			expectNum(
				evalFormula('CORREL(A1:A3, B1:B3)', {
					A1: 1,
					A2: 2,
					A3: 3,
					B1: 2,
					B2: 4,
					B3: 6,
				}),
				1.0,
			)
		})

		test('NORM.DIST standard normal CDF at zero', () => {
			expectNum(evalFormula('NORM.DIST(0, 0, 1, TRUE)'), 0.5, 0.0001)
		})

		test('LARGE second largest', () => {
			expectNum(evalFormula('LARGE(A1:A5, 2)', { A1: 3, A2: 1, A3: 5, A4: 2, A5: 4 }), 4)
		})
	})

	describe('text', () => {
		test('TEXT number format', () => {
			expect(evalFormula('TEXT(1234.5, "0.00")')).toEqual(stringValue('1234.50'))
		})

		test('TEXT percentage format', () => {
			expect(evalFormula('TEXT(0.75, "0%")')).toEqual(stringValue('75%'))
		})

		test('UPPER', () => {
			expect(evalFormula('UPPER("hello world")')).toEqual(stringValue('HELLO WORLD'))
		})

		test('CONCATENATE', () => {
			expect(evalFormula('CONCATENATE(A1, " ", A2)', { A1: 'Hello', A2: 'World' })).toEqual(
				stringValue('Hello World'),
			)
		})

		test('SUBSTITUTE', () => {
			expect(evalFormula('SUBSTITUTE("Hello World", "World", "Earth")')).toEqual(
				stringValue('Hello Earth'),
			)
		})
	})

	describe('date', () => {
		test('DATEDIF years', () => {
			expectNum(evalFormula('DATEDIF(DATE(2020,1,15), DATE(2021,3,20), "Y")'), 1)
		})

		test('DATEDIF months', () => {
			expectNum(evalFormula('DATEDIF(DATE(2020,1,15), DATE(2021,3,20), "M")'), 14)
		})

		test('YEARFRAC full year basis 0', () => {
			expectNum(evalFormula('YEARFRAC(DATE(2020,1,1), DATE(2021,1,1), 0)'), 1.0, 0.01)
		})

		test('YEAR of EDATE', () => {
			expectNum(evalFormula('YEAR(EDATE(DATE(2020,3,15), 2))'), 2020)
		})

		test('MONTH of EDATE', () => {
			expectNum(evalFormula('MONTH(EDATE(DATE(2020,3,15), 2))'), 5)
		})
	})

	describe('lookup', () => {
		test('VLOOKUP exact match', () => {
			const cells = {
				A1: 'Apple',
				B1: 10,
				A2: 'Banana',
				B2: 20,
				A3: 'Cherry',
				B3: 30,
			}
			expectNum(evalFormula('VLOOKUP("Banana", A1:B3, 2, FALSE)', cells), 20)
		})

		test('VLOOKUP approximate match', () => {
			const cells = {
				A1: 10,
				B1: 'low',
				A2: 20,
				B2: 'medium',
				A3: 30,
				B3: 'high',
			}
			expect(evalFormula('VLOOKUP(25, A1:B3, 2, TRUE)', cells)).toEqual(stringValue('medium'))
		})

		test('XLOOKUP exact match', () => {
			const cells = {
				A1: 'Apple',
				B1: 10,
				A2: 'Banana',
				B2: 20,
				A3: 'Cherry',
				B3: 30,
			}
			expectNum(evalFormula('XLOOKUP("Cherry", A1:A3, B1:B3)', cells), 30)
		})

		test('XLOOKUP not found with default', () => {
			const cells = { A1: 'Apple', B1: 10, A2: 'Banana', B2: 20 }
			expect(evalFormula('XLOOKUP("Durian", A1:A2, B1:B2, "Not found")', cells)).toEqual(
				stringValue('Not found'),
			)
		})

		test('INDEX MATCH combo', () => {
			const cells = {
				A1: 'Apple',
				B1: 10,
				A2: 'Banana',
				B2: 20,
				A3: 'Cherry',
				B3: 30,
			}
			expectNum(evalFormula('INDEX(B1:B3, MATCH("Banana", A1:A3, 0))', cells), 20)
		})

		test('MATCH exact', () => {
			const cells = { A1: 'Apple', A2: 'Banana', A3: 'Cherry' }
			expectNum(evalFormula('MATCH("Cherry", A1:A3, 0)', cells), 3)
		})
	})

	describe('dynamic arrays', () => {
		test('SEQUENCE generates sequence', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(4)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 0)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(1, 0)?.value ?? EMPTY, 2)
			expectNum(sheet.cells.get(2, 0)?.value ?? EMPTY, 3)
			expectNum(sheet.cells.get(3, 0)?.value ?? EMPTY, 4)
		})

		test('SORT ascending', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(3, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'SORT(A1:A4)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(3, 2)?.value ?? EMPTY, 4)
		})

		test('UNIQUE removes duplicates', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(3, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(4, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'UNIQUE(A1:A5)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(1, 2)?.value ?? EMPTY, 2)
			expectNum(sheet.cells.get(2, 2)?.value ?? EMPTY, 3)
		})

		test('FILTER with include mask', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(1, 1, { value: numberValue(0), formula: null, styleId: sid })
			sheet.cells.set(2, 1, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(0, 3, { value: EMPTY, formula: 'FILTER(A1:A3, B1:B3)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 3)?.value ?? EMPTY, 10)
			expectNum(sheet.cells.get(1, 3)?.value ?? EMPTY, 30)
		})
	})

	describe('logical', () => {
		test('nested IF', () => {
			expect(evalFormula('IF(A1>10, IF(A1>20, "high", "medium"), "low")', { A1: 15 })).toEqual(
				stringValue('medium'),
			)
		})

		test('IFS multi-condition', () => {
			expect(evalFormula('IFS(A1>20, "high", A1>10, "medium", TRUE, "low")', { A1: 15 })).toEqual(
				stringValue('medium'),
			)
		})

		test('SWITCH value matching', () => {
			expect(evalFormula('SWITCH(A1, 1, "one", 2, "two", "other")', { A1: 2 })).toEqual(
				stringValue('two'),
			)
		})

		test('SWITCH default', () => {
			expect(evalFormula('SWITCH(A1, 1, "one", 2, "two", "other")', { A1: 5 })).toEqual(
				stringValue('other'),
			)
		})

		test('AND all true', () => {
			expect(evalFormula('AND(A1>0, A2>0, A3>0)', { A1: 5, A2: 3, A3: 1 })).toEqual(
				booleanValue(true),
			)
		})

		test('AND with false', () => {
			expect(evalFormula('AND(A1>0, A2>0)', { A1: 5, A2: -1 })).toEqual(booleanValue(false))
		})

		test('OR mixed', () => {
			expect(evalFormula('OR(A1>10, A2>10)', { A1: 5, A2: 15 })).toEqual(booleanValue(true))
		})

		test('IFERROR catches #N/A', () => {
			const cells = { A1: 'Apple', B1: 10 }
			expect(
				evalFormula('IFERROR(VLOOKUP("missing", A1:B1, 2, FALSE), "not found")', cells),
			).toEqual(stringValue('not found'))
		})

		test('IFERROR passes through non-error', () => {
			const cells = { A1: 'Apple', B1: 10 }
			expectNum(evalFormula('IFERROR(VLOOKUP("Apple", A1:B1, 2, FALSE), 0)', cells), 10)
		})
	})
})
