import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'

const S0 = 0 as StyleId

function makeWorkbook() {
	const wb = createWorkbook()
	wb.addSheet('Sheet1')
	return wb
}

function setNum(wb: ReturnType<typeof createWorkbook>, row: number, col: number, val: number) {
	wb.sheets[0]?.cells.set(row, col, { value: numberValue(val), formula: null, styleId: S0 })
}

function setStr(wb: ReturnType<typeof createWorkbook>, row: number, col: number, val: string) {
	wb.sheets[0]?.cells.set(row, col, { value: stringValue(val), formula: null, styleId: S0 })
}

function setBool(wb: ReturnType<typeof createWorkbook>, row: number, col: number, val: boolean) {
	wb.sheets[0]?.cells.set(row, col, { value: booleanValue(val), formula: null, styleId: S0 })
}

function setFormula(wb: ReturnType<typeof createWorkbook>, row: number, col: number, f: string) {
	wb.sheets[0]?.cells.set(row, col, { value: EMPTY, formula: f, styleId: S0 })
}

function getResult(wb: ReturnType<typeof createWorkbook>, row: number, col: number) {
	return wb.sheets[0]?.cells.get(row, col)?.value
}

function recalc(wb: ReturnType<typeof createWorkbook>) {
	return recalculate(wb, defaultCalcContext())
}

describe('formula functions', () => {
	describe('math functions', () => {
		test('SUM with mixed types - sums numbers and dates, ignores text and boolean', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setStr(wb, 1, 0, '5')
			setBool(wb, 2, 0, true)
			setNum(wb, 3, 0, 3)
			setFormula(wb, 4, 0, 'SUM(A1:A4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(13))
		})

		test('SUM with error in range propagates error', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			wb.sheets[0]?.cells.set(1, 0, {
				value: errorValue('#N/A'),
				formula: null,
				styleId: S0,
			})
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'SUM(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#N/A'))
		})

		test('SUM with range argument uses forEachValue path', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'SUM(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(6))
		})

		test('AVERAGE with empty range returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'AVERAGE(A2:A5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('COUNT vs COUNTA vs COUNTBLANK differences', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setStr(wb, 1, 0, 'hello')
			setNum(wb, 2, 0, 3)
			setBool(wb, 3, 0, true)
			// A5 is empty
			setFormula(wb, 5, 0, 'COUNT(A1:A5)')
			setFormula(wb, 6, 0, 'COUNTA(A1:A5)')
			setFormula(wb, 7, 0, 'COUNTBLANK(A1:A5)')
			recalc(wb)
			expect(getResult(wb, 5, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 6, 0)).toEqual(numberValue(4))
			expect(getResult(wb, 7, 0)).toEqual(numberValue(1))
		})

		test('SUMIF with wildcard pattern', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'North')
			setStr(wb, 1, 0, 'Northeast')
			setStr(wb, 2, 0, 'South')
			setStr(wb, 3, 0, 'Northwest')
			setNum(wb, 0, 1, 10)
			setNum(wb, 1, 1, 20)
			setNum(wb, 2, 1, 30)
			setNum(wb, 3, 1, 40)
			setFormula(wb, 4, 0, 'SUMIF(A1:A4,"North*",B1:B4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(70))
		})

		test('COUNTIF with wildcard pattern', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'apple')
			setStr(wb, 1, 0, 'banana')
			setStr(wb, 2, 0, 'apricot')
			setFormula(wb, 3, 0, 'COUNTIF(A1:A3,"ap*")')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
		})

		test('SUMPRODUCT with mismatched dimensions returns #VALUE!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 4)
			setNum(wb, 0, 2, 5)
			setFormula(wb, 2, 0, 'SUMPRODUCT(A1:B2,C1:C2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#VALUE!'))
		})

		test('MOD with negative divisor', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 7)
			setNum(wb, 0, 1, -3)
			setFormula(wb, 0, 2, 'MOD(A1,B1)')
			recalc(wb)
			expect(getResult(wb, 0, 2)).toEqual(numberValue(-2))
		})

		test('ROUND with negative digits rounds to left of decimal', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1234.56)
			setFormula(wb, 0, 1, 'ROUND(A1,-2)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1200))
		})

		test('ROUNDUP with negative digits', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1234.12)
			setFormula(wb, 0, 1, 'ROUNDUP(A1,-2)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1300))
		})

		test('ROUNDDOWN with negative digits', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1234.99)
			setFormula(wb, 0, 1, 'ROUNDDOWN(A1,-2)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1200))
		})
	})

	describe('text functions', () => {
		test('CONCATENATE vs CONCAT - CONCATENATE takes multiple args', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'Hello')
			setStr(wb, 1, 0, ' ')
			setStr(wb, 2, 0, 'World')
			setFormula(wb, 3, 0, 'CONCATENATE(A1,A2,A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('Hello World'))
		})

		test('CONCAT flattens range', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'Q')
			setNum(wb, 1, 0, 2)
			setStr(wb, 2, 0, '-report')
			setFormula(wb, 3, 0, 'CONCAT(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('Q2-report'))
		})

		test('TEXTJOIN with ignore_empty skips blanks', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'red')
			setStr(wb, 1, 0, '')
			setStr(wb, 2, 0, '')
			setStr(wb, 3, 0, 'blue')
			setFormula(wb, 4, 0, 'TEXTJOIN(" / ",TRUE,A1:A4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(stringValue('red / blue'))
		})

		test('TEXTJOIN with ignore_empty false includes blanks', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 1, 0, '')
			setStr(wb, 2, 0, 'b')
			setFormula(wb, 3, 0, 'TEXTJOIN("-",FALSE,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('a--b'))
		})

		test('FIND is case-sensitive', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'P')
			setStr(wb, 0, 1, 'Apple')
			setFormula(wb, 0, 2, 'FIND(A1,B1)')
			recalc(wb)
			expect(getResult(wb, 0, 2)).toEqual(errorValue('#VALUE!'))
		})

		test('SEARCH is case-insensitive', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'P')
			setStr(wb, 0, 1, 'Apple')
			setFormula(wb, 0, 2, 'SEARCH(A1,B1)')
			recalc(wb)
			expect(getResult(wb, 0, 2)).toEqual(numberValue(2))
		})

		test('SUBSTITUTE with instance_num replaces only Nth occurrence', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SUBSTITUTE("2026-Q1-Q1","Q1","Q2",2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('2026-Q1-Q2'))
		})

		test('TEXT formatting', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.125)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.0%")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('12.5%'))
		})
	})

	describe('lookup functions', () => {
		test('VLOOKUP exact match', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'apple')
			setStr(wb, 1, 0, 'banana')
			setStr(wb, 2, 0, 'cherry')
			setNum(wb, 0, 1, 10)
			setNum(wb, 1, 1, 20)
			setNum(wb, 2, 1, 30)
			setFormula(wb, 3, 0, 'VLOOKUP("banana",A1:B3,2,FALSE)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(20))
		})

		test('VLOOKUP exact match returns the first duplicate hit', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'banana')
			setStr(wb, 1, 0, 'banana')
			setNum(wb, 0, 1, 20)
			setNum(wb, 1, 1, 30)
			setFormula(wb, 2, 0, 'VLOOKUP("banana",A1:B2,2,FALSE)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(20))
		})

		test('VLOOKUP approximate match', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setNum(wb, 2, 0, 30)
			setStr(wb, 0, 1, 'small')
			setStr(wb, 1, 1, 'medium')
			setStr(wb, 2, 1, 'large')
			setFormula(wb, 3, 0, 'VLOOKUP(25,A1:B3,2,TRUE)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('medium'))
		})

		test('INDEX with row 0 returns entire column (spills, anchor has first value)', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 2, 0, 'INDEX(A1:B2,0,2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(4))
		})

		test('INDEX with col 0 returns entire row (spills, anchor has first value)', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 2, 0, 'INDEX(A1:B2,2,0)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(4))
		})

		test('MATCH with 0 (exact)', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 1, 0, 'b')
			setStr(wb, 2, 0, 'c')
			setFormula(wb, 3, 0, 'MATCH("b",A1:A3,0)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
		})

		test('MATCH with 1 (ascending approximate)', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setNum(wb, 2, 0, 30)
			setFormula(wb, 3, 0, 'MATCH(25,A1:A3,1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
		})

		test('MATCH with -1 (descending approximate)', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 30)
			setNum(wb, 1, 0, 20)
			setNum(wb, 2, 0, 10)
			setFormula(wb, 3, 0, 'MATCH(25,A1:A3,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
		})

		test('XLOOKUP with search_mode and match_mode', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'apple')
			setStr(wb, 1, 0, 'banana')
			setStr(wb, 2, 0, 'banana')
			setNum(wb, 0, 1, 10)
			setNum(wb, 1, 1, 20)
			setNum(wb, 2, 1, 30)
			setFormula(wb, 3, 0, 'XLOOKUP("banana",A1:A3,B1:B3,"missing",0,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(30))
		})

		test('XMATCH exact reverse search returns the last duplicate hit', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'banana')
			setStr(wb, 1, 0, 'apple')
			setStr(wb, 2, 0, 'banana')
			setFormula(wb, 3, 0, 'XMATCH("banana",A1:A3,0,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(3))
		})
	})

	describe('logical functions', () => {
		test('IF short-circuits the untaken branch', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IF(FALSE,1/0,7)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(7))
		})

		test('IFERROR short-circuits the fallback branch when not needed', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IFERROR(42,1/0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(42))
		})

		test('IFNA short-circuits the fallback branch when not needed', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IFNA(42,1/0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(42))
		})

		test('IFS with no match returns #N/A', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 50)
			setFormula(wb, 0, 1, 'IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(errorValue('#N/A'))
		})

		test('SWITCH with default value', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SWITCH("z","A",1,"B",2,"C",3,99)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(99))
		})

		test('XOR with range', () => {
			const wb = makeWorkbook()
			setBool(wb, 0, 0, true)
			setBool(wb, 1, 0, false)
			setBool(wb, 2, 0, true)
			setFormula(wb, 3, 0, 'XOR(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(booleanValue(false))
		})
	})

	describe('stats functions', () => {
		test('AVERAGEA includes text as 0 and booleans as 0/1', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setStr(wb, 1, 0, 'text')
			setBool(wb, 2, 0, true)
			setNum(wb, 3, 0, 5)
			setFormula(wb, 4, 0, 'AVERAGEA(A1:A4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(4))
		})

		test('MAXA treats TRUE as 1 and text as 0', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -5)
			setStr(wb, 1, 0, 'text')
			setBool(wb, 2, 0, true)
			setFormula(wb, 3, 0, 'MAXA(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
		})

		test('MINA treats FALSE as 0 and text as 0', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 5)
			setBool(wb, 1, 0, false)
			setStr(wb, 2, 0, 'text')
			setFormula(wb, 3, 0, 'MINA(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(0))
		})

		test('RANK.EQ returns same as RANK', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 3)
			setNum(wb, 1, 0, 1)
			setNum(wb, 2, 0, 5)
			setFormula(wb, 3, 0, 'RANK.EQ(3,A1:A3)')
			setFormula(wb, 4, 0, 'RANK(3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(getResult(wb, 4, 0))
		})

		test('RANK.AVG returns average for tied values', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 3)
			setNum(wb, 1, 0, 3)
			setNum(wb, 2, 0, 5)
			setNum(wb, 3, 0, 1)
			setFormula(wb, 4, 0, 'RANK.AVG(3,A1:A4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2.5))
		})

		test('GEOMEAN of positive values', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 8)
			setFormula(wb, 2, 0, 'GEOMEAN(A1:A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(4))
		})

		test('GEOMEAN with non-positive value returns #NUM!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 0)
			setFormula(wb, 2, 0, 'GEOMEAN(A1:A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#NUM!'))
		})

		test('HARMEAN of positive values', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 4)
			setFormula(wb, 2, 0, 'HARMEAN(A1:A2)')
			recalc(wb)
			const r = getResult(wb, 2, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(8 / 3)
		})

		test('TRIMMEAN trims 40% and averages interior', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 3, 0, 4)
			setNum(wb, 4, 0, 5)
			setFormula(wb, 5, 0, 'TRIMMEAN(A1:A5,0.4)')
			recalc(wb)
			expect(getResult(wb, 5, 0)).toEqual(numberValue(3))
		})

		test('PERCENTRANK.INC returns inclusive percentile rank', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 3, 0, 4)
			setNum(wb, 4, 0, 5)
			setFormula(wb, 5, 0, 'PERCENTRANK.INC(A1:A5,3)')
			recalc(wb)
			expect(getResult(wb, 5, 0)).toEqual(numberValue(0.5))
		})

		test('PERCENTRANK.EXC returns exclusive percentile rank', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 3, 0, 4)
			setNum(wb, 4, 0, 5)
			setFormula(wb, 5, 0, 'PERCENTRANK.EXC(A1:A5,3)')
			recalc(wb)
			expect(getResult(wb, 5, 0)).toEqual(numberValue(0.5))
		})
	})

	describe('additional math functions', () => {
		test('SUMSQ sums squares of values', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 3)
			setNum(wb, 1, 0, 4)
			setFormula(wb, 2, 0, 'SUMSQ(A1:A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(25))
		})

		test('ROMAN converts Arabic to Roman numeral', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ROMAN(499)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('CDXCIX'))
		})

		test('ARABIC converts Roman to Arabic number', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ARABIC("CDXCIX")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(499))
		})

		test('BASE converts to target radix with padding', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BASE(255,16,4)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('00FF'))
		})

		test('DECIMAL converts from target radix', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DECIMAL("FF",16)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(255))
		})
	})

	describe('additional rounding functions', () => {
		test('CEILING.MATH rounds up to nearest multiple', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CEILING.MATH(6.3,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(8))
		})

		test('CEILING.MATH with negative and mode rounds away from zero', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CEILING.MATH(-4.1,2,1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(-6))
		})

		test('CEILING.PRECISE rounds toward positive infinity', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CEILING.PRECISE(-4.1,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(-4))
		})

		test('FLOOR.MATH rounds down to nearest multiple', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FLOOR.MATH(6.7,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(6))
		})

		test('FLOOR.PRECISE rounds toward negative infinity', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FLOOR.PRECISE(-4.1,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(-6))
		})

		test('ISO.CEILING same as CEILING.PRECISE', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ISO.CEILING(-4.1,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(-4))
		})
	})

	describe('logical constants', () => {
		test('TRUE() returns TRUE', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'TRUE()')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(booleanValue(true))
		})

		test('FALSE() returns FALSE', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FALSE()')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(booleanValue(false))
		})
	})

	describe('text format functions', () => {
		test('FIXED with decimals and commas', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FIXED(1234567.89,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('1,234,567.89'))
		})

		test('FIXED with no commas flag', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FIXED(1234.56,2,TRUE)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('1234.56'))
		})

		test('DOLLAR formats as currency', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DOLLAR(1234.56)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('$1,234.56'))
		})

		test('VALUETOTEXT returns string representation', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 42)
			setFormula(wb, 1, 0, 'VALUETOTEXT(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(stringValue('42'))
		})
	})

	describe('date functions - DAYS', () => {
		test('DAYS computes difference between dates', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATE(2021,1,1)')
			setFormula(wb, 0, 1, 'DATE(2022,1,1)')
			setFormula(wb, 0, 2, 'DAYS(B1,A1)')
			recalc(wb)
			expect(getResult(wb, 0, 2)).toEqual(numberValue(365))
		})
	})

	describe('database aggregate functions', () => {
		test('DSTDEV computes sample standard deviation', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'Name')
			setStr(wb, 0, 1, 'Score')
			setStr(wb, 1, 0, 'A')
			setNum(wb, 1, 1, 10)
			setStr(wb, 2, 0, 'A')
			setNum(wb, 2, 1, 20)
			setStr(wb, 3, 0, 'A')
			setNum(wb, 3, 1, 30)
			setStr(wb, 5, 0, 'Name')
			setStr(wb, 6, 0, 'A')
			setFormula(wb, 7, 0, 'DSTDEV(A1:B4,2,A6:A7)')
			recalc(wb)
			expect(getResult(wb, 7, 0)).toEqual(numberValue(10))
		})

		test('DVARP computes population variance', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'Name')
			setStr(wb, 0, 1, 'Score')
			setStr(wb, 1, 0, 'A')
			setNum(wb, 1, 1, 10)
			setStr(wb, 2, 0, 'A')
			setNum(wb, 2, 1, 20)
			setStr(wb, 3, 0, 'A')
			setNum(wb, 3, 1, 30)
			setStr(wb, 5, 0, 'Name')
			setStr(wb, 6, 0, 'A')
			setFormula(wb, 7, 0, 'DVARP(A1:B4,2,A6:A7)')
			recalc(wb)
			const r = getResult(wb, 7, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(200 / 3)
		})
	})

	describe('dynamic array functions', () => {
		test('SORT descending', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 30)
			setNum(wb, 1, 0, 10)
			setNum(wb, 2, 0, 20)
			setFormula(wb, 3, 0, 'SORT(A1:A3,1,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(30))
		})

		test('FILTER with no matches returns #CALC! when no fallback', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setBool(wb, 0, 1, false)
			setBool(wb, 1, 1, false)
			setFormula(wb, 2, 0, 'FILTER(A1:A2,B1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#CALC!'))
		})

		test('UNIQUE with by_col', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 0, 1, 'b')
			setStr(wb, 0, 2, 'a')
			setFormula(wb, 1, 0, 'UNIQUE(A1:C1,TRUE)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(stringValue('a'))
		})

		test('SEQUENCE with step', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SEQUENCE(4,1,0,5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
		})
	})

	describe('regression functions', () => {
		test('FORECAST.LINEAR with perfect linear data y=2x', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 3, 0, 'FORECAST.LINEAR(4,B1:B3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(8))
		})

		test('FORECAST.LINEAR with noisy data', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 5)
			setNum(wb, 3, 1, 4)
			setNum(wb, 4, 1, 5)
			setFormula(wb, 5, 0, 'FORECAST.LINEAR(6,B1:B5,A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(5.8)
		})

		test('FORECAST is alias for FORECAST.LINEAR', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 3, 0, 'FORECAST(4,B1:B3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(8))
		})

		test('SLOPE of perfect linear data y=2x', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 3, 0, 'SLOPE(B1:B3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
		})

		test('SLOPE with noisy data', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 5)
			setNum(wb, 3, 1, 4)
			setNum(wb, 4, 1, 5)
			setFormula(wb, 5, 0, 'SLOPE(B1:B5,A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.6)
		})

		test('INTERCEPT of y=2x+1', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 5)
			setNum(wb, 2, 1, 7)
			setFormula(wb, 3, 0, 'INTERCEPT(B1:B3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
		})

		test('INTERCEPT with noisy data', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 5)
			setNum(wb, 3, 1, 4)
			setNum(wb, 4, 1, 5)
			setFormula(wb, 5, 0, 'INTERCEPT(B1:B5,A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(2.2)
		})

		test('RSQ of perfectly correlated data is 1', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 3, 0, 'RSQ(B1:B3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
		})

		test('RSQ with noisy data', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 5)
			setNum(wb, 3, 1, 4)
			setNum(wb, 4, 1, 5)
			setFormula(wb, 5, 0, 'RSQ(B1:B5,A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.6)
		})

		test('CORREL of perfectly correlated data is 1', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 3, 0, 'CORREL(A1:A3,B1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
		})

		test('CORREL with negative correlation', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 6)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 2)
			setFormula(wb, 3, 0, 'CORREL(A1:A3,B1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(-1))
		})

		test('PEARSON is alias for CORREL', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 3, 0, 'PEARSON(A1:A3,B1:B3)')
			setFormula(wb, 3, 1, 'CORREL(A1:A3,B1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(getResult(wb, 3, 1))
		})

		test('STEYX with noisy data', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 5)
			setNum(wb, 3, 1, 4)
			setNum(wb, 4, 1, 5)
			setFormula(wb, 5, 0, 'STEYX(B1:B5,A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.8944, 3)
		})

		test('STEYX requires at least 3 data points', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 5)
			setFormula(wb, 2, 0, 'STEYX(B1:B2,A1:A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('COVARIANCE.P computes population covariance', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 5)
			setNum(wb, 2, 1, 7)
			setFormula(wb, 3, 0, 'COVARIANCE.P(A1:A3,B1:B3)')
			recalc(wb)
			const r = getResult(wb, 3, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(4 / 3)
		})

		test('COVARIANCE.S computes sample covariance', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 5)
			setNum(wb, 2, 1, 7)
			setFormula(wb, 3, 0, 'COVARIANCE.S(A1:A3,B1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
		})

		test('COVARIANCE.S with single pair returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 5)
			setNum(wb, 0, 1, 10)
			setFormula(wb, 1, 0, 'COVARIANCE.S(A1:A1,B1:B1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#DIV/0!'))
		})
	})

	describe('deviation and shape functions', () => {
		test('AVEDEV computes average absolute deviation', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 4)
			setNum(wb, 2, 0, 6)
			setNum(wb, 3, 0, 8)
			setFormula(wb, 4, 0, 'AVEDEV(A1:A4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
		})

		test('AVEDEV with scalar argument', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'AVEDEV(5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
		})

		test('DEVSQ computes sum of squared deviations', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 4)
			setNum(wb, 2, 0, 6)
			setFormula(wb, 3, 0, 'DEVSQ(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(8))
		})

		test('DEVSQ of identical values is 0', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 5)
			setNum(wb, 1, 0, 5)
			setNum(wb, 2, 0, 5)
			setFormula(wb, 3, 0, 'DEVSQ(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(0))
		})

		test('KURT of uniform-spaced data', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setFormula(wb, 5, 0, 'KURT(A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(-1.2)
		})

		test('KURT with fewer than 4 values returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'KURT(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('SKEW of symmetric data is zero', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) setNum(wb, i, 0, i + 1)
			setFormula(wb, 5, 0, 'SKEW(A1:A5)')
			recalc(wb)
			const r = getResult(wb, 5, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0)
		})

		test('SKEW with fewer than 3 values returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setFormula(wb, 2, 0, 'SKEW(A1:A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#DIV/0!'))
		})
	})

	describe('frequency and multimode functions', () => {
		test('FREQUENCY distributes data into bins', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 3)
			setNum(wb, 2, 0, 5)
			setNum(wb, 3, 0, 7)
			setNum(wb, 4, 0, 9)
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 6)
			setFormula(wb, 5, 0, 'FREQUENCY(A1:A5,B1:B2)')
			recalc(wb)
			expect(getResult(wb, 5, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 6, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 7, 0)).toEqual(numberValue(2))
		})

		test('FREQUENCY with all data above bins', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setNum(wb, 0, 1, 5)
			setFormula(wb, 2, 0, 'FREQUENCY(A1:A2,B1:B1)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
		})

		test('MODE.MULT returns multiple modes as array', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 2)
			setNum(wb, 3, 0, 3)
			setNum(wb, 4, 0, 3)
			setNum(wb, 5, 0, 4)
			setFormula(wb, 6, 0, 'MODE.MULT(A1:A6)')
			recalc(wb)
			expect(getResult(wb, 6, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 7, 0)).toEqual(numberValue(3))
		})

		test('MODE.MULT with no repeats returns #N/A', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'MODE.MULT(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#N/A'))
		})
	})

	describe('normal distribution functions', () => {
		test('NORM.S.DIST CDF at z=0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.S.DIST(0,TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5, 7)
		})

		test('NORM.S.DIST PDF at z=0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.S.DIST(0,FALSE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.39894, 4)
		})

		test('NORM.S.DIST CDF at z=1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.S.DIST(1,TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.8413, 3)
		})

		test('NORM.S.INV at p=0.5 is 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.S.INV(0.5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0, 5)
		})

		test('NORM.S.INV at p=0.975 is approximately 1.96', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.S.INV(0.975)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1.96, 1)
		})

		test('NORM.S.INV out of range returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.S.INV(0)')
			setFormula(wb, 0, 1, 'NORM.S.INV(1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
			expect(getResult(wb, 0, 1)).toEqual(errorValue('#NUM!'))
		})

		test('NORM.DIST CDF at mean equals 0.5', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.DIST(10,10,2,TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5, 7)
		})

		test('NORM.DIST with negative stdev returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.DIST(0,0,-1,TRUE)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('NORM.DIST PDF matches standard formula', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.DIST(0,0,1,FALSE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.39894, 4)
		})

		test('NORM.INV at p=0.5 returns mean', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.INV(0.5,10,2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(10, 5)
		})

		test('NORM.INV out of range returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.INV(1.5,10,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('NORM.INV roundtrips with NORM.DIST', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NORM.DIST(NORM.INV(0.95,100,15),100,15,TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.95, 2)
		})
	})
})
