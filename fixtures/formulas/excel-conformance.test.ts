import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook, parseA1 } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import type { CellValue } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'

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

		test('AND with numeric coercion', () => {
			expect(evalFormula('AND(1, 2, 3)')).toEqual(booleanValue(true))
		})

		test('AND with zero is false', () => {
			expect(evalFormula('AND(1, 0, 1)')).toEqual(booleanValue(false))
		})

		test('OR with all false numerics', () => {
			expect(evalFormula('OR(0, 0, 0)')).toEqual(booleanValue(false))
		})

		test('OR with single truthy', () => {
			expect(evalFormula('OR(0, 0, 1)')).toEqual(booleanValue(true))
		})

		test('XOR odd number of true', () => {
			expect(evalFormula('XOR(TRUE, FALSE, TRUE)')).toEqual(booleanValue(false))
		})

		test('XOR single true', () => {
			expect(evalFormula('XOR(TRUE, FALSE, FALSE)')).toEqual(booleanValue(true))
		})

		test('IFNA catches only #N/A', () => {
			const result = evalFormula('IFNA(A1/0, "caught")', { A1: 1 })
			expect(result).toEqual(errorValue('#DIV/0!'))
		})

		test('IFNA catches #N/A from MATCH', () => {
			expect(evalFormula('IFNA(MATCH("z", A1:A3, 0), -1)', { A1: 'a', A2: 'b', A3: 'c' })).toEqual(
				numberValue(-1),
			)
		})

		test('triple nested IF', () => {
			expect(
				evalFormula('IF(A1>30,"high",IF(A1>20,"med-high",IF(A1>10,"medium","low")))', { A1: 25 }),
			).toEqual(stringValue('med-high'))
		})
	})

	describe('math edge cases', () => {
		test('ROUND -0.5', () => {
			expectNum(evalFormula('ROUND(-0.5, 0)'), 0)
		})

		test('ROUND 0.5 rounds away from zero', () => {
			expectNum(evalFormula('ROUND(0.5, 0)'), 1)
		})

		test('ROUND 2.5 rounds to 3', () => {
			expectNum(evalFormula('ROUND(2.5, 0)'), 3)
		})

		test('ROUND negative places', () => {
			expectNum(evalFormula('ROUND(1234, -2)'), 1200)
		})

		test('ROUND large decimal places', () => {
			expectNum(evalFormula('ROUND(1.23456789, 5)'), 1.23457, 1e-6)
		})

		test('ROUNDUP -2.1 rounds away from zero', () => {
			expectNum(evalFormula('ROUNDUP(-2.1, 0)'), -3)
		})

		test('ROUNDDOWN -2.9 towards zero', () => {
			expectNum(evalFormula('ROUNDDOWN(-2.9, 0)'), -2)
		})

		test('POWER negative base even exponent', () => {
			expectNum(evalFormula('POWER(-2, 4)'), 16)
		})

		test('POWER negative base odd exponent', () => {
			expectNum(evalFormula('POWER(-2, 3)'), -8)
		})

		test('POWER fractional exponent of negative base returns #NUM!', () => {
			expect(evalFormula('POWER(-4, 0.5)').kind).toBe('error')
		})

		test('LOG of 0 returns #NUM!', () => {
			expect(evalFormula('LOG(0)').kind).toBe('error')
		})

		test('LOG of negative returns #NUM!', () => {
			expect(evalFormula('LOG(-1)').kind).toBe('error')
		})

		test('LOG base 2', () => {
			expectNum(evalFormula('LOG(8, 2)'), 3)
		})

		test('LN of E', () => {
			expectNum(evalFormula('LN(EXP(1))'), 1, 1e-10)
		})

		test('MOD with negative dividend (Excel convention)', () => {
			expectNum(evalFormula('MOD(-7, 3)'), 2)
		})

		test('MOD with negative divisor', () => {
			expectNum(evalFormula('MOD(7, -3)'), -2)
		})

		test('SQRT of 0', () => {
			expectNum(evalFormula('SQRT(0)'), 0)
		})

		test('SQRT of negative returns #NUM!', () => {
			expect(evalFormula('SQRT(-1)').kind).toBe('error')
		})

		test('ABS of 0', () => {
			expectNum(evalFormula('ABS(0)'), 0)
		})

		test('INT floors negative numbers', () => {
			expectNum(evalFormula('INT(-3.7)'), -4)
		})

		test('TRUNC truncates negative numbers', () => {
			expectNum(evalFormula('TRUNC(-3.7)'), -3)
		})

		test('0^0 equals 1', () => {
			expectNum(evalFormula('0^0'), 1)
		})

		test('division by zero returns #DIV/0!', () => {
			expect(evalFormula('1/0')).toEqual(errorValue('#DIV/0!'))
		})

		test('SUMPRODUCT with multiplication', () => {
			expectNum(
				evalFormula('SUMPRODUCT(A1:A3, B1:B3)', {
					A1: 1,
					A2: 2,
					A3: 3,
					B1: 4,
					B2: 5,
					B3: 6,
				}),
				32,
			)
		})

		test('GCD of multiple values', () => {
			expectNum(evalFormula('GCD(12, 18, 24)'), 6)
		})

		test('LCM of multiple values', () => {
			expectNum(evalFormula('LCM(4, 6, 10)'), 60)
		})

		test('SIGN negative', () => {
			expectNum(evalFormula('SIGN(-42)'), -1)
		})

		test('SIGN zero', () => {
			expectNum(evalFormula('SIGN(0)'), 0)
		})

		test('FACT of 0', () => {
			expectNum(evalFormula('FACT(0)'), 1)
		})

		test('COMBIN(10, 3)', () => {
			expectNum(evalFormula('COMBIN(10, 3)'), 120)
		})
	})

	describe('lookup edge cases', () => {
		test('VLOOKUP case-insensitive exact match', () => {
			const cells = { A1: 'Apple', B1: 10, A2: 'Banana', B2: 20, A3: 'Cherry', B3: 30 }
			expectNum(evalFormula('VLOOKUP("banana", A1:B3, 2, FALSE)', cells), 20)
		})

		test('VLOOKUP returns last match in approximate mode', () => {
			const cells = { A1: 1, B1: 'one', A2: 5, B2: 'five', A3: 10, B3: 'ten' }
			expect(evalFormula('VLOOKUP(7, A1:B3, 2, TRUE)', cells)).toEqual(stringValue('five'))
		})

		test('VLOOKUP col_index out of range returns #REF!', () => {
			expect(evalFormula('VLOOKUP("a", A1:B1, 3, FALSE)', { A1: 'a', B1: 1 }).kind).toBe('error')
		})

		test('VLOOKUP not found returns #N/A', () => {
			expect(evalFormula('VLOOKUP("z", A1:B1, 2, FALSE)', { A1: 'a', B1: 1 })).toEqual(
				errorValue('#N/A'),
			)
		})

		test('HLOOKUP exact match', () => {
			const cells = { A1: 'x', B1: 'y', C1: 'z', A2: 10, B2: 20, C2: 30 }
			expectNum(evalFormula('HLOOKUP("y", A1:C2, 2, FALSE)', cells), 20)
		})

		test('XLOOKUP with -1 search mode (reverse)', () => {
			const cells = {
				A1: 'Apple',
				B1: 10,
				A2: 'Banana',
				B2: 20,
				A3: 'Cherry',
				B3: 30,
			}
			expectNum(evalFormula('XLOOKUP("Cherry", A1:A3, B1:B3, , 0, -1)', cells), 30)
		})

		test('XLOOKUP next larger match mode', () => {
			const cells = { A1: 10, B1: 'ten', A2: 20, B2: 'twenty', A3: 30, B3: 'thirty' }
			expect(evalFormula('XLOOKUP(15, A1:A3, B1:B3, "none", 1)', cells)).toEqual(
				stringValue('twenty'),
			)
		})

		test('XLOOKUP next smaller match mode', () => {
			const cells = { A1: 10, B1: 'ten', A2: 20, B2: 'twenty', A3: 30, B3: 'thirty' }
			expect(evalFormula('XLOOKUP(25, A1:A3, B1:B3, "none", -1)', cells)).toEqual(
				stringValue('twenty'),
			)
		})

		test('INDEX single cell', () => {
			const cells = { A1: 10, A2: 20, A3: 30 }
			expectNum(evalFormula('INDEX(A1:A3, 2)', cells), 20)
		})

		test('INDEX 2D array', () => {
			const cells = { A1: 1, B1: 2, C1: 3, A2: 4, B2: 5, C2: 6 }
			expectNum(evalFormula('INDEX(A1:C2, 2, 3)', cells), 6)
		})

		test('MATCH approximate ascending', () => {
			const cells = { A1: 10, A2: 20, A3: 30, A4: 40 }
			expectNum(evalFormula('MATCH(25, A1:A4, 1)', cells), 2)
		})

		test('MATCH approximate descending', () => {
			const cells = { A1: 40, A2: 30, A3: 20, A4: 10 }
			expectNum(evalFormula('MATCH(25, A1:A4, -1)', cells), 2)
		})

		test('XMATCH exact match', () => {
			const cells = { A1: 'a', A2: 'b', A3: 'c' }
			expectNum(evalFormula('XMATCH("b", A1:A3, 0)', cells), 2)
		})

		test('CHOOSE selects nth value', () => {
			expect(evalFormula('CHOOSE(3, "a", "b", "c", "d")')).toEqual(stringValue('c'))
		})

		test('ADDRESS absolute', () => {
			expect(evalFormula('ADDRESS(1, 1, 1)')).toEqual(stringValue('$A$1'))
		})

		test('ADDRESS relative', () => {
			expect(evalFormula('ADDRESS(3, 2, 4)')).toEqual(stringValue('B3'))
		})

		test('ROWS of range', () => {
			expectNum(evalFormula('ROWS(A1:A10)'), 10)
		})

		test('COLUMNS of range', () => {
			expectNum(evalFormula('COLUMNS(A1:E1)'), 5)
		})
	})

	describe('date edge cases', () => {
		test('DATE across year boundary Dec to Jan', () => {
			const result = evalFormula('MONTH(DATE(2023, 12, 31) + 1)')
			expectNum(result, 1)
		})

		test('DATE with month overflow', () => {
			expectNum(evalFormula('MONTH(DATE(2023, 13, 1))'), 1)
		})

		test('DATE with month overflow year', () => {
			expectNum(evalFormula('YEAR(DATE(2023, 13, 1))'), 2024)
		})

		test('DATE with negative month wraps back', () => {
			expectNum(evalFormula('MONTH(DATE(2023, -1, 1))'), 11)
		})

		test('DATEDIF days', () => {
			expectNum(evalFormula('DATEDIF(DATE(2020,1,1), DATE(2020,1,31), "D")'), 30)
		})

		test('DATEDIF MD (days ignoring months/years)', () => {
			expectNum(evalFormula('DATEDIF(DATE(2020,1,15), DATE(2021,3,20), "MD")'), 5)
		})

		test('NETWORKDAYS simple', () => {
			expectNum(evalFormula('NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,5))'), 5)
		})

		test('NETWORKDAYS spanning weekend', () => {
			expectNum(evalFormula('NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,8))'), 6)
		})

		test('YEARFRAC basis 0 (US 30/360)', () => {
			expectNum(evalFormula('YEARFRAC(DATE(2020,1,1), DATE(2020,7,1), 0)'), 0.5, 0.01)
		})

		test('YEARFRAC basis 1 (actual/actual)', () => {
			expectNum(evalFormula('YEARFRAC(DATE(2020,1,1), DATE(2021,1,1), 1)'), 1.0, 0.01)
		})

		test('YEARFRAC basis 3 (actual/365)', () => {
			expectNum(evalFormula('YEARFRAC(DATE(2020,1,1), DATE(2020,7,1), 3)'), 182 / 365, 0.01)
		})

		test('EOMONTH January gives end of March with +2', () => {
			expectNum(evalFormula('DAY(EOMONTH(DATE(2024,1,15), 2))'), 31)
		})

		test('EOMONTH February leap year', () => {
			expectNum(evalFormula('DAY(EOMONTH(DATE(2024,1,15), 1))'), 29)
		})

		test('EOMONTH February non-leap year', () => {
			expectNum(evalFormula('DAY(EOMONTH(DATE(2023,1,15), 1))'), 28)
		})

		test('WEEKDAY Sunday=1 default', () => {
			expectNum(evalFormula('WEEKDAY(DATE(2024,1,7))'), 1)
		})

		test('WEEKDAY Monday=1 return_type 2', () => {
			expectNum(evalFormula('WEEKDAY(DATE(2024,1,8), 2)'), 1)
		})

		test('DAYS between dates', () => {
			expectNum(evalFormula('DAYS(DATE(2024,3,1), DATE(2024,1,1))'), 60)
		})

		test('WORKDAY skips weekends', () => {
			expectNum(evalFormula('DAY(WORKDAY(DATE(2024,1,5), 1))'), 8)
		})
	})

	describe('text edge cases', () => {
		test('SUBSTITUTE with instance_num replaces only nth occurrence', () => {
			expect(evalFormula('SUBSTITUTE("a-b-c-d", "-", ".", 2)')).toEqual(stringValue('a-b.c-d'))
		})

		test('SUBSTITUTE all occurrences (no instance_num)', () => {
			expect(evalFormula('SUBSTITUTE("aaa", "a", "b")')).toEqual(stringValue('bbb'))
		})

		test('SUBSTITUTE not found leaves unchanged', () => {
			expect(evalFormula('SUBSTITUTE("hello", "xyz", "abc")')).toEqual(stringValue('hello'))
		})

		test('MID with start beyond length returns empty', () => {
			expect(evalFormula('MID("abc", 10, 5)')).toEqual(stringValue(''))
		})

		test('MID with 0 length returns empty', () => {
			expect(evalFormula('MID("abc", 1, 0)')).toEqual(stringValue(''))
		})

		test('TRIM collapses internal spaces', () => {
			expect(evalFormula('TRIM("  hello   world  ")')).toEqual(stringValue('hello world'))
		})

		test('LEN of empty string', () => {
			expectNum(evalFormula('LEN("")'), 0)
		})

		test('FIND case-sensitive', () => {
			expectNum(evalFormula('FIND("B", "aBcBd")'), 2)
		})

		test('SEARCH case-insensitive', () => {
			expectNum(evalFormula('SEARCH("b", "aBcBd")'), 2)
		})

		test('FIND not found returns #VALUE!', () => {
			expect(evalFormula('FIND("z", "hello")').kind).toBe('error')
		})

		test('REPLACE in middle', () => {
			expect(evalFormula('REPLACE("Hello World", 7, 5, "Earth")')).toEqual(
				stringValue('Hello Earth'),
			)
		})

		test('EXACT case sensitive comparison', () => {
			expect(evalFormula('EXACT("Hello", "hello")')).toEqual(booleanValue(false))
		})

		test('EXACT identical strings', () => {
			expect(evalFormula('EXACT("Hello", "Hello")')).toEqual(booleanValue(true))
		})

		test('PROPER capitalizes each word', () => {
			expect(evalFormula('PROPER("hello world")')).toEqual(stringValue('Hello World'))
		})

		test('REPT repeats string', () => {
			expect(evalFormula('REPT("ab", 3)')).toEqual(stringValue('ababab'))
		})

		test('REPT zero times returns empty', () => {
			expect(evalFormula('REPT("ab", 0)')).toEqual(stringValue(''))
		})

		test('CODE returns ASCII value', () => {
			expectNum(evalFormula('CODE("A")'), 65)
		})

		test('CHAR from ASCII value', () => {
			expect(evalFormula('CHAR(65)')).toEqual(stringValue('A'))
		})

		test('VALUE parses number string', () => {
			expectNum(evalFormula('VALUE("123.45")'), 123.45)
		})

		test('TEXTJOIN with delimiter', () => {
			expect(evalFormula('TEXTJOIN("-", TRUE, A1, A2, A3)', { A1: 'a', A2: 'b', A3: 'c' })).toEqual(
				stringValue('a-b-c'),
			)
		})

		test('TEXTJOIN ignores empty when TRUE', () => {
			expect(evalFormula('TEXTJOIN(",", TRUE, "a", "", "c")')).toEqual(stringValue('a,c'))
		})

		test('LEFT defaults to 1 char', () => {
			expect(evalFormula('LEFT("hello")')).toEqual(stringValue('h'))
		})

		test('RIGHT defaults to 1 char', () => {
			expect(evalFormula('RIGHT("hello")')).toEqual(stringValue('o'))
		})

		test('CLEAN removes non-printable chars', () => {
			expect(evalFormula('CLEAN(CHAR(7)&"hello"&CHAR(10))')).toEqual(stringValue('hello'))
		})

		test('LOWER', () => {
			expect(evalFormula('LOWER("HELLO WORLD")')).toEqual(stringValue('hello world'))
		})
	})

	describe('statistical edge cases', () => {
		test('PERCENTILE at 0 is minimum', () => {
			expectNum(evalFormula('PERCENTILE(A1:A5, 0)', { A1: 10, A2: 20, A3: 30, A4: 40, A5: 50 }), 10)
		})

		test('PERCENTILE at 1 is maximum', () => {
			expectNum(evalFormula('PERCENTILE(A1:A5, 1)', { A1: 10, A2: 20, A3: 30, A4: 40, A5: 50 }), 50)
		})

		test('PERCENTILE at 0.25 is Q1', () => {
			expectNum(evalFormula('PERCENTILE(A1:A5, 0.25)', { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5 }), 2)
		})

		test('STDEV with two identical values returns 0', () => {
			expectNum(evalFormula('STDEV(A1:A2)', { A1: 5, A2: 5 }), 0)
		})

		test('VAR of single value is not computable (N-1=0)', () => {
			expect(evalFormula('VAR(A1)', { A1: 5 }).kind).toBe('error')
		})

		test('VAR.P of single value is 0', () => {
			expectNum(evalFormula('VAR.P(A1)', { A1: 5 }), 0)
		})

		test('CORREL perfect negative', () => {
			expectNum(
				evalFormula('CORREL(A1:A3, B1:B3)', {
					A1: 1,
					A2: 2,
					A3: 3,
					B1: 6,
					B2: 4,
					B3: 2,
				}),
				-1.0,
			)
		})

		test('MODE returns most frequent', () => {
			expectNum(evalFormula('MODE(A1:A7)', { A1: 1, A2: 2, A3: 2, A4: 3, A5: 3, A6: 3, A7: 4 }), 3)
		})

		test('RANK ascending', () => {
			expectNum(evalFormula('RANK(A2, A1:A5, 1)', { A1: 10, A2: 30, A3: 20, A4: 50, A5: 40 }), 3)
		})

		test('RANK descending (default)', () => {
			expectNum(evalFormula('RANK(A2, A1:A5)', { A1: 10, A2: 30, A3: 20, A4: 50, A5: 40 }), 3)
		})

		test('SMALL second smallest', () => {
			expectNum(evalFormula('SMALL(A1:A5, 2)', { A1: 5, A2: 3, A3: 1, A4: 4, A5: 2 }), 2)
		})

		test('GEOMEAN', () => {
			expectNum(evalFormula('GEOMEAN(A1:A3)', { A1: 8, A2: 12, A3: 18 }), 12, 0.01)
		})

		test('HARMEAN', () => {
			expectNum(evalFormula('HARMEAN(A1:A3)', { A1: 1, A2: 2, A3: 4 }), 12 / 7, 0.01)
		})

		test('QUARTILE Q1', () => {
			expectNum(evalFormula('QUARTILE(A1:A5, 1)', { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5 }), 2)
		})

		test('QUARTILE Q3', () => {
			expectNum(evalFormula('QUARTILE(A1:A5, 3)', { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5 }), 4)
		})

		test('NORM.S.DIST standard normal CDF at 0', () => {
			expectNum(evalFormula('NORM.S.DIST(0, TRUE)'), 0.5, 0.0001)
		})

		test('NORM.INV at 0.975 gives ~1.96', () => {
			expectNum(evalFormula('NORM.INV(0.975, 0, 1)'), 1.96, 0.01)
		})

		test('AVERAGEA with boolean', () => {
			expectNum(evalFormula('AVERAGEA(A1:A3)', { A1: 1, A2: true, A3: false }), 2 / 3, 0.01)
		})

		test('FORECAST.LINEAR extrapolation', () => {
			expectNum(
				evalFormula('FORECAST(6, A1:A5, B1:B5)', {
					A1: 2,
					A2: 4,
					A3: 6,
					A4: 8,
					A5: 10,
					B1: 1,
					B2: 2,
					B3: 3,
					B4: 4,
					B5: 5,
				}),
				12,
				0.01,
			)
		})

		test('SLOPE of perfect linear data', () => {
			expectNum(
				evalFormula('SLOPE(A1:A4, B1:B4)', {
					A1: 2,
					A2: 4,
					A3: 6,
					A4: 8,
					B1: 1,
					B2: 2,
					B3: 3,
					B4: 4,
				}),
				2,
			)
		})

		test('INTERCEPT of perfect linear data', () => {
			expectNum(
				evalFormula('INTERCEPT(A1:A4, B1:B4)', {
					A1: 2,
					A2: 4,
					A3: 6,
					A4: 8,
					B1: 1,
					B2: 2,
					B3: 3,
					B4: 4,
				}),
				0,
				0.001,
			)
		})
	})

	describe('financial edge cases', () => {
		test('NPV with mixed positive and negative flows', () => {
			expectNum(
				evalFormula('NPV(0.1, A1, A2, A3, A4)', { A1: -500, A2: 200, A3: 200, A4: 200 }),
				-2.63,
				1.0,
			)
		})

		test('IRR with multiple sign changes', () => {
			expectNum(evalFormula('IRR(A1:A4)', { A1: -100, A2: 50, A3: 50, A4: 50 }), 0.2339, 0.01)
		})

		test('IRR break-even', () => {
			expectNum(evalFormula('IRR(A1:A3)', { A1: -200, A2: 100, A3: 100 }), 0.0, 0.01)
		})

		test('PMT with type=1 (beginning of period)', () => {
			expectNum(evalFormula('PMT(0.05/12, 60, -10000, 0, 1)'), 188.71, 2.0)
		})

		test('NPER zero interest', () => {
			expectNum(evalFormula('NPER(0, -100, 1000)'), 10)
		})

		test('SLN straight-line depreciation', () => {
			expectNum(evalFormula('SLN(30000, 5000, 10)'), 2500)
		})

		test('SYD first year depreciation', () => {
			expectNum(evalFormula('SYD(30000, 5000, 10, 1)'), 4545.45, 0.01)
		})

		test('IPMT interest portion first period', () => {
			expectNum(evalFormula('IPMT(0.1/12, 1, 12, -10000)'), 83.33, 0.01)
		})

		test('PPMT principal portion first period', () => {
			const ipmt = evalFormula('IPMT(0.1/12, 1, 12, -10000)')
			const pmt = evalFormula('PMT(0.1/12, 12, -10000)')
			const ppmt = evalFormula('PPMT(0.1/12, 1, 12, -10000)')
			if (ipmt.kind === 'number' && pmt.kind === 'number' && ppmt.kind === 'number') {
				expect(Math.abs(ppmt.value - (pmt.value - ipmt.value))).toBeLessThan(0.01)
			}
		})

		test('EFFECT annual effective rate', () => {
			expectNum(evalFormula('EFFECT(0.10, 4)'), 0.10381, 0.001)
		})

		test('NOMINAL from effective', () => {
			expectNum(evalFormula('NOMINAL(0.10381, 4)'), 0.1, 0.001)
		})

		test('FV with payments and type=1', () => {
			expectNum(evalFormula('FV(0.05/12, 60, -200, 0, 1)'), 13644.37, 200.0)
		})

		test('PV with future value target', () => {
			expectNum(evalFormula('PV(0.08, 10, 0, -10000)'), 4631.93, 0.1)
		})
	})

	describe('dynamic array edge cases', () => {
		test('SORT preserves order of equal elements (stability)', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(3, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'SORT(A1:A4)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(1, 2)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(2, 2)?.value ?? EMPTY, 2)
			expectNum(sheet.cells.get(3, 2)?.value ?? EMPTY, 2)
		})

		test('SORT descending', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'SORT(A1:A3, 1, -1)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 3)
			expectNum(sheet.cells.get(1, 2)?.value ?? EMPTY, 2)
			expectNum(sheet.cells.get(2, 2)?.value ?? EMPTY, 1)
		})

		test('UNIQUE preserves first-seen order', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(3, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(4, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'UNIQUE(A1:A5)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 3)
			expectNum(sheet.cells.get(1, 2)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(2, 2)?.value ?? EMPTY, 2)
		})

		test('UNIQUE with strings', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: stringValue('a'), formula: null, styleId: sid })
			sheet.cells.set(3, 0, { value: stringValue('c'), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'UNIQUE(A1:A4)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('a'))
			expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('b'))
			expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('c'))
		})

		test('FILTER with no matches returns #CALC!', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: booleanValue(false), formula: null, styleId: sid })
			sheet.cells.set(1, 1, { value: booleanValue(false), formula: null, styleId: sid })
			sheet.cells.set(2, 1, { value: booleanValue(false), formula: null, styleId: sid })
			sheet.cells.set(0, 3, { value: EMPTY, formula: 'FILTER(A1:A3, B1:B3)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			const result = sheet.cells.get(0, 3)?.value ?? EMPTY
			expect(result.kind).toBe('error')
		})

		test('SEQUENCE with start and step', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3, 1, 10, 5)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 0)?.value ?? EMPTY, 10)
			expectNum(sheet.cells.get(1, 0)?.value ?? EMPTY, 15)
			expectNum(sheet.cells.get(2, 0)?.value ?? EMPTY, 20)
		})

		test('SEQUENCE 2D grid', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(2, 3)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 0)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(0, 1)?.value ?? EMPTY, 2)
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 3)
			expectNum(sheet.cells.get(1, 0)?.value ?? EMPTY, 4)
			expectNum(sheet.cells.get(1, 1)?.value ?? EMPTY, 5)
			expectNum(sheet.cells.get(1, 2)?.value ?? EMPTY, 6)
		})

		test('TRANSPOSE column to row', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'TRANSPOSE(A1:A3)', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(sheet.cells.get(0, 2)?.value ?? EMPTY, 1)
			expectNum(sheet.cells.get(0, 3)?.value ?? EMPTY, 2)
			expectNum(sheet.cells.get(0, 4)?.value ?? EMPTY, 3)
		})
	})

	describe('engineering', () => {
		test('BIN2DEC basic', () => {
			expectNum(evalFormula('BIN2DEC("1010")'), 10)
		})

		test('DEC2HEX basic', () => {
			expect(evalFormula('DEC2HEX(255)')).toEqual(stringValue('FF'))
		})

		test('HEX2DEC basic', () => {
			expectNum(evalFormula('HEX2DEC("FF")'), 255)
		})

		test('DELTA equal values', () => {
			expectNum(evalFormula('DELTA(5, 5)'), 1)
		})

		test('DELTA different values', () => {
			expectNum(evalFormula('DELTA(5, 6)'), 0)
		})

		test('COMPLEX real and imaginary', () => {
			expect(evalFormula('COMPLEX(3, 4)')).toEqual(stringValue('3+4i'))
		})

		test('IMABS of 3+4i = 5', () => {
			expectNum(evalFormula('IMABS("3+4i")'), 5)
		})

		test('ERF at 1', () => {
			expectNum(evalFormula('ERF(1)'), 0.8427, 0.001)
		})

		test('BITAND', () => {
			expectNum(evalFormula('BITAND(12, 10)'), 8)
		})

		test('BITOR', () => {
			expectNum(evalFormula('BITOR(12, 10)'), 14)
		})

		test('BITXOR', () => {
			expectNum(evalFormula('BITXOR(12, 10)'), 6)
		})
	})

	describe('info functions', () => {
		test('ISBLANK on empty cell', () => {
			expect(evalFormula('ISBLANK(A1)')).toEqual(booleanValue(true))
		})

		test('ISBLANK on non-empty cell', () => {
			expect(evalFormula('ISBLANK(A1)', { A1: 1 })).toEqual(booleanValue(false))
		})

		test('ISNUMBER on number', () => {
			expect(evalFormula('ISNUMBER(A1)', { A1: 42 })).toEqual(booleanValue(true))
		})

		test('ISNUMBER on text', () => {
			expect(evalFormula('ISNUMBER(A1)', { A1: 'hello' })).toEqual(booleanValue(false))
		})

		test('ISTEXT on text', () => {
			expect(evalFormula('ISTEXT(A1)', { A1: 'hello' })).toEqual(booleanValue(true))
		})

		test('ISTEXT on number', () => {
			expect(evalFormula('ISTEXT(A1)', { A1: 42 })).toEqual(booleanValue(false))
		})

		test('ISLOGICAL on boolean', () => {
			expect(evalFormula('ISLOGICAL(A1)', { A1: true })).toEqual(booleanValue(true))
		})

		test('ISLOGICAL on number', () => {
			expect(evalFormula('ISLOGICAL(A1)', { A1: 1 })).toEqual(booleanValue(false))
		})

		test('ISEVEN', () => {
			expect(evalFormula('ISEVEN(4)')).toEqual(booleanValue(true))
		})

		test('ISODD', () => {
			expect(evalFormula('ISODD(3)')).toEqual(booleanValue(true))
		})

		test('N of boolean TRUE is 1', () => {
			expectNum(evalFormula('N(TRUE)'), 1)
		})

		test('N of boolean FALSE is 0', () => {
			expectNum(evalFormula('N(FALSE)'), 0)
		})

		test('TYPE of number is 1', () => {
			expectNum(evalFormula('TYPE(42)'), 1)
		})

		test('TYPE of text is 2', () => {
			expectNum(evalFormula('TYPE("hello")'), 2)
		})

		test('TYPE of boolean is 4', () => {
			expectNum(evalFormula('TYPE(TRUE)'), 4)
		})

		test('TYPE of error is 16', () => {
			expectNum(evalFormula('TYPE(1/0)'), 16)
		})
	})

	describe('cross-sheet references', () => {
		test('SUM across two sheets', () => {
			const wb = createWorkbook()
			const s1 = wb.addSheet('Sheet1')
			const s2 = wb.addSheet('Sheet2')
			s1.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
			s1.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
			s2.cells.set(0, 0, { value: numberValue(30), formula: null, styleId: sid })
			s2.cells.set(1, 0, { value: numberValue(40), formula: null, styleId: sid })
			s1.cells.set(5, 0, { value: EMPTY, formula: 'SUM(A1:A2)+Sheet2!A1+Sheet2!A2', styleId: sid })
			recalculate(wb, defaultCalcContext())
			expectNum(s1.cells.get(5, 0)?.value ?? EMPTY, 100)
		})

		test('VLOOKUP with cross-sheet table', () => {
			const wb = createWorkbook()
			const s1 = wb.addSheet('Sheet1')
			const s2 = wb.addSheet('Sheet2')
			s2.cells.set(0, 0, { value: stringValue('Apple'), formula: null, styleId: sid })
			s2.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
			s2.cells.set(1, 0, { value: stringValue('Banana'), formula: null, styleId: sid })
			s2.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
			s1.cells.set(0, 0, { value: stringValue('Banana'), formula: null, styleId: sid })
			s1.cells.set(0, 1, {
				value: EMPTY,
				formula: 'VLOOKUP(A1, Sheet2!A1:B2, 2, FALSE)',
				styleId: sid,
			})
			recalculate(wb, defaultCalcContext())
			expectNum(s1.cells.get(0, 1)?.value ?? EMPTY, 20)
		})
	})
})
