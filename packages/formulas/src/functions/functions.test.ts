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
	})

	describe('logical functions', () => {
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
})
