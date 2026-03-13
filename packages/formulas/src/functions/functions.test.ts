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

		test('TEXT with #,##0.00 format', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1234567.891)
			setFormula(wb, 0, 1, 'TEXT(A1,"#,##0.00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1,234,567.89'))
		})

		test('TEXT with 0 digit placeholder', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 5)
			setFormula(wb, 0, 1, 'TEXT(A1,"000")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('005'))
		})

		test('TEXT with # digit placeholder trims trailing zeros', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.##")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1.5'))
		})

		test('TEXT with $#,##0.00 currency', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 9876.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"$#,##0.00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('$9,876.50'))
		})

		test('TEXT with 0.00E+00 scientific', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 12345)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.00E+00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1.23E+04'))
		})

		test('TEXT with quoted literal text', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 42)
			setStr(wb, 0, 2, '0" units"')
			setFormula(wb, 0, 1, 'TEXT(A1,C1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('42 units'))
		})

		test('TEXT with date format m/d/yyyy', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 45658)
			setFormula(wb, 0, 1, 'TEXT(A1,"m/d/yyyy")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1/1/2025'))
		})

		test('TEXT with date format yyyy-mm-dd', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 45658)
			setFormula(wb, 0, 1, 'TEXT(A1,"yyyy-mm-dd")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('2025-01-01'))
		})

		test('TEXT with 0.00% format', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.7523)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.00%")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('75.23%'))
		})

		test('TEXT with negative number', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -1234.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"#,##0.00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('-1,234.50'))
		})

		test('TEXT with @ text placeholder', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 99)
			setFormula(wb, 0, 1, 'TEXT(A1,"@")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('99'))
		})

		test('TEXT # placeholder gives empty for zero', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0)
			setFormula(wb, 0, 1, 'TEXT(A1,"#.##")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue(''))
		})

		test('TEXT with time format hh:mm:ss', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 45658.75)
			setFormula(wb, 0, 1, 'TEXT(A1,"hh:mm:ss")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('18:00:00'))
		})

		test('TEXT with AM/PM time format', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 45658.75)
			setFormula(wb, 0, 1, 'TEXT(A1,"h:mm AM/PM")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('6:00 PM'))
		})

		test('TEXT scientific E-00 only shows minus for negative exp', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.005)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.00E-00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('5.00E-03'))
		})

		test('TEXT uses explicit negative and zero sections', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -12.3)
			setNum(wb, 1, 0, 0)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.0;(0.0);zero")')
			setFormula(wb, 1, 1, 'TEXT(A2,"0.0;(0.0);zero")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('(12.3)'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('zero'))
		})

		test('TEXT honors conditional sections', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1500)
			setNum(wb, 1, 0, 12)
			setFormula(wb, 0, 1, 'TEXT(A1,"[>=1000]0.0,""K"";0")')
			setFormula(wb, 1, 1, 'TEXT(A2,"[>=1000]0.0,""K"";0")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1.5K'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('12'))
		})

		test('TEXT supports scaling commas', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 12200000)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.0,,""M""")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('12.2M'))
		})

		test('TEXT supports multiple percent characters', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.0123)
			setFormula(wb, 0, 1, 'TEXT(A1,"0.0%%")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('123.0%%'))
		})

		test('TEXT supports currency locale markers', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 42)
			setFormula(wb, 0, 1, 'TEXT(A1,"[$USD]0.00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('USD42.00'))
		})

		test('TEXT supports elapsed hours tokens', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"[h]:mm")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('36:00'))
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
		test('FILTER keeps rows where include is TRUE', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setNum(wb, 2, 0, 30)
			setBool(wb, 0, 1, true)
			setBool(wb, 1, 1, false)
			setBool(wb, 2, 1, true)
			setFormula(wb, 3, 0, 'FILTER(A1:A3,B1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(10))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(30))
		})

		test('FILTER column mode keeps matching columns', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setNum(wb, 1, 0, 4)
			setNum(wb, 1, 1, 5)
			setNum(wb, 1, 2, 6)
			setBool(wb, 2, 0, true)
			setBool(wb, 2, 1, false)
			setBool(wb, 2, 2, true)
			setFormula(wb, 3, 0, 'FILTER(A1:C2,A3:C3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(4))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(6))
		})

		test('FILTER with if_empty returns fallback when no matches', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setBool(wb, 0, 1, false)
			setFormula(wb, 1, 0, 'FILTER(A1:A1,B1:B1,"none")')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(stringValue('none'))
		})

		test('FILTER with no matches and no fallback returns #CALC!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setBool(wb, 0, 1, false)
			setBool(wb, 1, 1, false)
			setFormula(wb, 2, 0, 'FILTER(A1:A2,B1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#CALC!'))
		})

		test('SORT ascending by default', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 30)
			setNum(wb, 1, 0, 10)
			setNum(wb, 2, 0, 20)
			setFormula(wb, 3, 0, 'SORT(A1:A3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(10))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(20))
			expect(getResult(wb, 5, 0)).toEqual(numberValue(30))
		})

		test('SORT descending', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 30)
			setNum(wb, 1, 0, 10)
			setNum(wb, 2, 0, 20)
			setFormula(wb, 3, 0, 'SORT(A1:A3,1,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(30))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(20))
			expect(getResult(wb, 5, 0)).toEqual(numberValue(10))
		})

		test('SORT by second column', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'c')
			setStr(wb, 1, 0, 'a')
			setStr(wb, 2, 0, 'b')
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 1)
			setNum(wb, 2, 1, 2)
			setFormula(wb, 3, 0, 'SORT(A1:B3,2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('a'))
			expect(getResult(wb, 4, 0)).toEqual(stringValue('b'))
			expect(getResult(wb, 5, 0)).toEqual(stringValue('c'))
		})

		test('SORT with by_col sorts columns', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 3)
			setNum(wb, 0, 1, 1)
			setNum(wb, 0, 2, 2)
			setNum(wb, 1, 0, 30)
			setNum(wb, 1, 1, 10)
			setNum(wb, 1, 2, 20)
			setFormula(wb, 2, 0, 'SORT(A1:C2,1,1,TRUE)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 2)).toEqual(numberValue(3))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(10))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(20))
			expect(getResult(wb, 3, 2)).toEqual(numberValue(30))
		})

		test('SORTBY sorts by external array', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'c')
			setStr(wb, 1, 0, 'a')
			setStr(wb, 2, 0, 'b')
			setNum(wb, 0, 1, 3)
			setNum(wb, 1, 1, 1)
			setNum(wb, 2, 1, 2)
			setFormula(wb, 3, 0, 'SORTBY(A1:A3,B1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('a'))
			expect(getResult(wb, 4, 0)).toEqual(stringValue('b'))
			expect(getResult(wb, 5, 0)).toEqual(stringValue('c'))
		})

		test('SORTBY with descending order', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 1, 0, 'b')
			setStr(wb, 2, 0, 'c')
			setNum(wb, 0, 1, 1)
			setNum(wb, 1, 1, 2)
			setNum(wb, 2, 1, 3)
			setFormula(wb, 3, 0, 'SORTBY(A1:A3,B1:B3,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('c'))
			expect(getResult(wb, 4, 0)).toEqual(stringValue('b'))
			expect(getResult(wb, 5, 0)).toEqual(stringValue('a'))
		})

		test('SORTBY with two sort keys', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'x')
			setStr(wb, 1, 0, 'y')
			setStr(wb, 2, 0, 'z')
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 1)
			setNum(wb, 2, 1, 2)
			setNum(wb, 0, 2, 20)
			setNum(wb, 1, 2, 10)
			setNum(wb, 2, 2, 10)
			setFormula(wb, 3, 0, 'SORTBY(A1:A3,B1:B3,1,C1:C3,1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('y'))
			expect(getResult(wb, 4, 0)).toEqual(stringValue('z'))
			expect(getResult(wb, 5, 0)).toEqual(stringValue('x'))
		})

		test('UNIQUE removes duplicate rows', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 1, 0, 'b')
			setStr(wb, 2, 0, 'a')
			setStr(wb, 3, 0, 'c')
			setFormula(wb, 4, 0, 'UNIQUE(A1:A4)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(stringValue('a'))
			expect(getResult(wb, 5, 0)).toEqual(stringValue('b'))
			expect(getResult(wb, 6, 0)).toEqual(stringValue('c'))
		})

		test('UNIQUE with exactly_once returns non-duplicated values', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 1, 0, 'b')
			setStr(wb, 2, 0, 'a')
			setStr(wb, 3, 0, 'c')
			setFormula(wb, 4, 0, 'UNIQUE(A1:A4,FALSE,TRUE)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(stringValue('b'))
			expect(getResult(wb, 5, 0)).toEqual(stringValue('c'))
		})

		test('UNIQUE with by_col removes duplicate columns', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 0, 1, 'b')
			setStr(wb, 0, 2, 'a')
			setFormula(wb, 1, 0, 'UNIQUE(A1:C1,TRUE)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(stringValue('a'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('b'))
		})

		test('SEQUENCE generates column of numbers', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SEQUENCE(3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(3))
		})

		test('SEQUENCE with rows and columns', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SEQUENCE(2,3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(3))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(4))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(5))
			expect(getResult(wb, 1, 2)).toEqual(numberValue(6))
		})

		test('SEQUENCE with custom start and step', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SEQUENCE(4,1,0,5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(5))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(10))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(15))
		})

		test('SEQUENCE with zero rows returns #CALC!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SEQUENCE(0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#CALC!'))
		})

		test('CHOOSECOLS selects specific columns', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setNum(wb, 1, 0, 4)
			setNum(wb, 1, 1, 5)
			setNum(wb, 1, 2, 6)
			setFormula(wb, 2, 0, 'CHOOSECOLS(A1:C2,1,3)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(4))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(6))
		})

		test('CHOOSECOLS with negative index selects from end', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setFormula(wb, 1, 0, 'CHOOSECOLS(A1:C1,-1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(3))
		})

		test('CHOOSECOLS with zero index returns #VALUE!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setFormula(wb, 1, 0, 'CHOOSECOLS(A1:A1,0)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#VALUE!'))
		})

		test('CHOOSEROWS selects specific rows', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setNum(wb, 2, 0, 30)
			setFormula(wb, 3, 0, 'CHOOSEROWS(A1:A3,1,3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(10))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(30))
		})

		test('CHOOSEROWS with negative index selects from end', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setNum(wb, 1, 0, 20)
			setNum(wb, 2, 0, 30)
			setFormula(wb, 3, 0, 'CHOOSEROWS(A1:A3,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(30))
		})

		test('CHOOSEROWS with zero index returns #VALUE!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 10)
			setFormula(wb, 1, 0, 'CHOOSEROWS(A1:A1,0)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#VALUE!'))
		})

		test('TAKE returns first N rows', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'TAKE(A1:A3,2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
		})

		test('TAKE with negative rows returns from end', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'TAKE(A1:A3,-2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(3))
		})

		test('TAKE with rows and columns', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setNum(wb, 1, 0, 4)
			setNum(wb, 1, 1, 5)
			setNum(wb, 1, 2, 6)
			setFormula(wb, 2, 0, 'TAKE(A1:C2,1,2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(2))
		})

		test('TAKE with zero returns #CALC!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setFormula(wb, 1, 0, 'TAKE(A1:A1,0)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#CALC!'))
		})

		test('DROP removes first N rows', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'DROP(A1:A3,1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(3))
		})

		test('DROP with negative rows removes from end', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 3, 0, 'DROP(A1:A3,-1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
		})

		test('DROP all rows returns #CALC!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setFormula(wb, 2, 0, 'DROP(A1:A2,2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#CALC!'))
		})

		test('HSTACK joins arrays horizontally', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 3, 3)
			setNum(wb, 1, 3, 4)
			setFormula(wb, 3, 0, 'HSTACK(A1:A2,D1:D2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(4))
		})

		test('HSTACK pads shorter arrays with #N/A', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 3, 3)
			setFormula(wb, 3, 0, 'HSTACK(A1:A2,D1:D1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 1)).toEqual(errorValue('#N/A'))
		})

		test('VSTACK joins arrays vertically', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 3, 0, 'VSTACK(A1:B1,A2:B2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(4))
		})

		test('VSTACK pads narrower arrays with #N/A', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setFormula(wb, 3, 0, 'VSTACK(A1:B1,A2:A2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 1)).toEqual(errorValue('#N/A'))
		})

		test('TOCOL flattens array to single column', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 2, 0, 'TOCOL(A1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 5, 0)).toEqual(numberValue(4))
		})

		test('TOCOL with scan by column', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 2, 0, 'TOCOL(A1:B2,0,TRUE)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 5, 0)).toEqual(numberValue(4))
		})

		test('TOROW flattens array to single row', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 2, 0, 'TOROW(A1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 2)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 3)).toEqual(numberValue(4))
		})

		test('WRAPCOLS wraps vector into columns', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setNum(wb, 0, 3, 4)
			setFormula(wb, 1, 0, 'WRAPCOLS(A1:D1,2)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(4))
		})

		test('WRAPCOLS pads incomplete columns with #N/A', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setFormula(wb, 1, 0, 'WRAPCOLS(A1:C1,2)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 1)).toEqual(errorValue('#N/A'))
		})

		test('WRAPROWS wraps vector into rows', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setNum(wb, 0, 3, 4)
			setFormula(wb, 1, 0, 'WRAPROWS(A1:D1,2)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(4))
		})

		test('WRAPROWS pads incomplete rows with #N/A', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 0, 2, 3)
			setFormula(wb, 1, 0, 'WRAPROWS(A1:C1,2)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 1)).toEqual(errorValue('#N/A'))
		})

		test('EXPAND pads array to target dimensions', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setFormula(wb, 1, 0, 'EXPAND(A1:B1,2,3)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 1, 2)).toEqual(errorValue('#N/A'))
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#N/A'))
			expect(getResult(wb, 2, 1)).toEqual(errorValue('#N/A'))
			expect(getResult(wb, 2, 2)).toEqual(errorValue('#N/A'))
		})

		test('EXPAND with custom pad value', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setFormula(wb, 1, 0, 'EXPAND(A1,2,2,0)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(0))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(0))
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

	describe('financial bond/depreciation functions', () => {
		test('DB first year with 7 months', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DB(1000,100,6,1,7)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(186.083, 2)
		})

		test('DB second year', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DB(1000,100,6,2,7)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(259.639, 2)
		})

		test('DB with zero cost returns 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DB(0,100,6,1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
		})

		test('VDB first period matches DDB', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'VDB(2400,300,10,0,1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(480))
		})

		test('VDB fractional period', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'VDB(2400,300,10,0,0.5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(240))
		})

		test('VDB with no_switch TRUE', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'VDB(2400,300,10,0,1,2,TRUE)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(480))
		})

		test('MIRR basic calculation', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -100)
			setNum(wb, 1, 0, 50)
			setNum(wb, 2, 0, 40)
			setNum(wb, 3, 0, 30)
			setFormula(wb, 4, 0, 'MIRR(A1:A4,0.10,0.12)')
			recalc(wb)
			const r = getResult(wb, 4, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.112, 2)
		})

		test('MIRR with all positive returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 100)
			setNum(wb, 1, 0, 200)
			setFormula(wb, 2, 0, 'MIRR(A1:A2,0.10,0.12)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('ISPMT computes interest for specific period', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ISPMT(0.1,1,3,8000000)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(-533333.33, 0)
		})

		test('ISPMT with zero nper returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ISPMT(0.1,1,0,1000)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('CUMIPMT sums interest over periods', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CUMIPMT(0.1,4,1000,1,1,0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(-100, 0)
		})

		test('CUMIPMT with invalid args returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CUMIPMT(-0.1,4,1000,1,1,0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('CUMPRINC sums principal over periods', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CUMPRINC(0.1,4,1000,1,4,0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(-1000, 0)
		})

		test('CUMPRINC with start > end returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CUMPRINC(0.1,4,1000,3,1,0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('EFFECT computes effective annual rate', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'EFFECT(0.10,4)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.10381, 4)
		})

		test('EFFECT with invalid rate returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'EFFECT(-0.1,4)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('NOMINAL roundtrips with EFFECT', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NOMINAL(EFFECT(0.10,4),4)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.1, 5)
		})

		test('NOMINAL computes nominal rate', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'NOMINAL(0.10,4)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.09645, 4)
		})

		test('PDURATION computes periods to reach target', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PDURATION(0.10,1000,2000)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(7.2725, 3)
		})

		test('PDURATION with non-positive rate returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PDURATION(0,1000,2000)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('RRI computes equivalent interest rate', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'RRI(10,1000,2000)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.07177, 4)
		})

		test('RRI with zero pv returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'RRI(10,0,2000)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('FVSCHEDULE with variable rates', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 1, 0.05)
			setNum(wb, 1, 1, 0.1)
			setNum(wb, 2, 1, 0.15)
			setFormula(wb, 0, 0, 'FVSCHEDULE(100,B1:B3)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(132.825, 2)
		})

		test('FVSCHEDULE with single rate', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 1, 0.1)
			setFormula(wb, 0, 0, 'FVSCHEDULE(1000,B1:B1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1100))
		})

		test('DISC computes discount rate', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DISC(DATE(2024,1,1),DATE(2024,7,1),97,100,2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.05934, 3)
		})

		test('DISC with invalid basis returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DISC(DATE(2024,1,1),DATE(2024,7,1),97,100,5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('INTRATE computes interest rate', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'INTRATE(DATE(2024,1,1),DATE(2024,7,1),97,100,2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.06118, 3)
		})

		test('INTRATE with settlement >= maturity returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'INTRATE(DATE(2024,7,1),DATE(2024,1,1),97,100)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})
	})

	describe('engineering functions - CONVERT', () => {
		test('CONVERT(lbm, kg) - Excel example 1 lb = 0.4535924 kg', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(1,"lbm","kg")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.4535924, 5)
		})
		test('CONVERT(F, C) - Excel example 68 F = 20 C', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(68,"F","C")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(20, 5)
		})
		test('CONVERT(C, F) - Excel example 6 C = 42.8 F', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(6,"C","F")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(42.8, 4)
		})
		test('CONVERT(tsp, tbs) - Excel example 6 tsp = 2 tbs', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(6,"tsp","tbs")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(2, 5)
		})
		test('CONVERT(gal, l) - Excel example 6 gal ≈ 22.72 L', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(6,"gal","l")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(22.71741274, 2)
		})
		test('CONVERT(mi, km) - Excel example 6 mi ≈ 9.656 km', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(6,"mi","km")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(9.656064, 4)
		})
		test('CONVERT(in, ft) - 6 in = 0.5 ft', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(6,"in","ft")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5, 5)
		})
		test('CONVERT(ft, m) - 100 ft^2 to m^2 - Excel example', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(CONVERT(100,"ft","m"),"ft","m")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(9.290304, 4)
		})
		test('CONVERT cross-category returns #N/A', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(2.5,"ft","sec")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#N/A'))
		})
		test('CONVERT invalid unit returns #N/A', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(1,"lbm","xyz")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#N/A'))
		})
		test('CONVERT with SI prefix - km to mi', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(6,"km","mi")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(3.728227153, 4)
		})
		test('CONVERT bit to byte', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONVERT(8,"bit","byte")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1, 5)
		})
	})

	describe('engineering functions - ERF/ERFC', () => {
		test('ERF(0) is 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERF(0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0, 5)
		})

		test('ERF(1) is approximately 0.8427', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERF(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.8427, 3)
		})

		test('ERF with two arguments computes difference', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERF(0,1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.8427, 3)
		})

		test('ERF.PRECISE(1) matches ERF(1)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERF.PRECISE(1)')
			setFormula(wb, 0, 1, 'ERF(1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(getResult(wb, 0, 1))
		})

		test('ERFC(0) is 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERFC(0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1, 5)
		})

		test('ERFC(1) is approximately 0.1573', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERFC(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.1573, 3)
		})

		test('ERFC.PRECISE(1) matches ERFC(1)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERFC.PRECISE(1)')
			setFormula(wb, 0, 1, 'ERFC(1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(getResult(wb, 0, 1))
		})

		test('ERF + ERFC = 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERF(0.5)+ERFC(0.5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1, 7)
		})
	})

	describe('engineering functions - complex numbers', () => {
		test('COMPLEX creates complex number string', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COMPLEX(3,4)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('3+4i'))
		})

		test('COMPLEX with j suffix', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COMPLEX(3,4,"j")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('3+4j'))
		})

		test('COMPLEX with zero imaginary', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COMPLEX(5,0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('5'))
		})

		test('IMREAL extracts real part', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMREAL("3+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(3))
		})

		test('IMREAL of pure imaginary is 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMREAL("4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
		})

		test('IMAGINARY extracts imaginary part', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMAGINARY("3+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(4))
		})

		test('IMAGINARY of pure real is 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMAGINARY("3")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
		})

		test('IMABS of 3+4i is 5', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMABS("3+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(5))
		})

		test('IMABS of real number', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMABS("-5")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(5))
		})

		test('IMARGUMENT of 1+i is pi/4', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMARGUMENT("1+i")')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(Math.PI / 4, 7)
		})

		test('IMARGUMENT of 0 returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMARGUMENT("0")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('IMCONJUGATE flips imaginary sign', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMCONJUGATE("3+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('3-4i'))
		})

		test('IMCONJUGATE of real number', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMCONJUGATE("5")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('5'))
		})

		test('IMSUM adds complex numbers', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSUM("3+4i","1+2i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('4+6i'))
		})

		test('IMSUM with three arguments', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSUM("1+i","2+2i","3+3i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('6+6i'))
		})

		test('IMSUB subtracts complex numbers', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSUB("3+4i","1+2i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('2+2i'))
		})

		test('IMSUB resulting in real', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSUB("3+4i","1+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('2'))
		})

		test('IMPRODUCT multiplies complex numbers', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMPRODUCT("1+2i","3+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('-5+10i'))
		})

		test('IMPRODUCT with real scaling', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMPRODUCT("3+4i","2")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('6+8i'))
		})

		test('IMDIV divides complex numbers', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMDIV("-5+10i","1+2i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('3+4i'))
		})

		test('IMDIV by zero returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMDIV("3+4i","0")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('IMPOWER squares complex number', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMPOWER("3+4i",2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('-7+24i'))
		})

		test('IMPOWER with i squared is -1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMPOWER("i",2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('-1'))
		})

		test('IMSQRT of -1 is i', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSQRT("-1")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('i'))
		})

		test('IMSQRT of 4 is 2', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSQRT("4")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('2'))
		})

		test('IMEXP of 0 is 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMEXP("0")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('1'))
		})

		test('IMEXP of i*pi is -1 (Euler identity)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMEXP(COMPLEX(0,PI()))')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('-1'))
		})

		test('IMLN of 1 is 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMLN("1")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('0'))
		})

		test('IMLN of 0 returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMLN("0")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('IMSIN of 0 is 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMSIN("0")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('0'))
		})

		test('IMSIN of real argument matches SIN', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMREAL(IMSIN(COMPLEX(1,0)))')
			setFormula(wb, 0, 1, 'SIN(1)')
			recalc(wb)
			const a = getResult(wb, 0, 0)
			const b = getResult(wb, 0, 1)
			expect(a?.kind).toBe('number')
			expect(b?.kind).toBe('number')
			if (a?.kind === 'number' && b?.kind === 'number') expect(a.value).toBeCloseTo(b.value, 7)
		})

		test('IMCOS of 0 is 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMCOS("0")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('1'))
		})

		test('IMCOS of real argument matches COS', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMREAL(IMCOS(COMPLEX(1,0)))')
			setFormula(wb, 0, 1, 'COS(1)')
			recalc(wb)
			const a = getResult(wb, 0, 0)
			const b = getResult(wb, 0, 1)
			expect(a?.kind).toBe('number')
			expect(b?.kind).toBe('number')
			if (a?.kind === 'number' && b?.kind === 'number') expect(a.value).toBeCloseTo(b.value, 7)
		})
	})

	describe('extended trig functions', () => {
		test('COT(1) = cos(1)/sin(1)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COT(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(Math.cos(1) / Math.sin(1), 10)
		})

		test('COT(0) returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COT(0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('CSC(PI()/2) = 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CSC(PI()/2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1, 10)
		})

		test('CSC(0) returns #DIV/0!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CSC(0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('SEC(0) = 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SEC(0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})

		test('SECH(0) = 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SECH(0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})

		test('COTH(1) = cosh(1)/sinh(1)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COTH(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(Math.cosh(1) / Math.sinh(1), 10)
		})

		test('CSCH(1) = 1/sinh(1)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CSCH(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1 / Math.sinh(1), 10)
		})

		test('ACOT(1) = PI()/4', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ACOT(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(Math.PI / 4, 10)
		})

		test('ACOTH(2) returns valid result', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ACOTH(2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5 * Math.log(3), 10)
		})

		test('ACOTH(0.5) returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ACOTH(0.5)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('large input returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COT(2^27)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})
	})

	describe('extended math functions', () => {
		test('SQRTPI(1) = sqrt(PI)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SQRTPI(1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(Math.sqrt(Math.PI), 10)
		})

		test('SQRTPI(-1) returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SQRTPI(-1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
		})

		test('COMBINA(4,3) = 20', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COMBINA(4,3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(20))
		})

		test('COMBINA(0,0) = 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COMBINA(0,0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})

		test('PERMUTATIONA(3,2) = 9', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PERMUTATIONA(3,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(9))
		})

		test('MULTINOMIAL(2,3,4) = 1260', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'MULTINOMIAL(2,3,4)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1260))
		})

		test('SERIESSUM computes power series', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, -1 / 6)
			setFormula(wb, 2, 0, 'SERIESSUM(PI()/4,1,2,A1:A2)')
			recalc(wb)
			const r = getResult(wb, 2, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') {
				const x = Math.PI / 4
				const expected = 1 * x + (-1 / 6) * x ** 3
				expect(r.value).toBeCloseTo(expected, 10)
			}
		})

		test('SUMX2MY2 computes sum of x²-y²', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 0, 1, 1)
			setNum(wb, 1, 1, 2)
			setFormula(wb, 2, 0, 'SUMX2MY2(A1:A2,B1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(4 - 1 + (9 - 4)))
		})

		test('SUMX2PY2 computes sum of x²+y²', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 0, 1, 1)
			setNum(wb, 1, 1, 2)
			setFormula(wb, 2, 0, 'SUMX2PY2(A1:A2,B1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(4 + 1 + (9 + 4)))
		})

		test('SUMXMY2 computes sum of (x-y)²', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 0, 1, 1)
			setNum(wb, 1, 1, 2)
			setFormula(wb, 2, 0, 'SUMXMY2(A1:A2,B1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1 + 1))
		})
	})

	describe('matrix functions', () => {
		test('MMULT multiplies 2x2 matrices', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setNum(wb, 0, 2, 5)
			setNum(wb, 0, 3, 6)
			setNum(wb, 1, 2, 7)
			setNum(wb, 1, 3, 8)
			setFormula(wb, 3, 0, 'MMULT(A1:B2,C1:D2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(19))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(22))
			expect(getResult(wb, 4, 0)).toEqual(numberValue(43))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(50))
		})

		test('MMULT incompatible dimensions returns #VALUE!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setNum(wb, 0, 2, 5)
			setFormula(wb, 3, 0, 'MMULT(A1:B2,C1:C1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#VALUE!'))
		})

		test('MDETERM of 2x2 matrix', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 2, 0, 'MDETERM(A1:B2)')
			recalc(wb)
			const r = getResult(wb, 2, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(-2, 10)
		})

		test('MINVERSE of 2x2 matrix', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 4)
			setNum(wb, 0, 1, 7)
			setNum(wb, 1, 0, 2)
			setNum(wb, 1, 1, 6)
			setFormula(wb, 3, 0, 'MINVERSE(A1:B2)')
			recalc(wb)
			const v00 = getResult(wb, 3, 0)
			const v01 = getResult(wb, 3, 1)
			expect(v00?.kind).toBe('number')
			if (v00?.kind === 'number') expect(v00.value).toBeCloseTo(0.6, 10)
			if (v01?.kind === 'number') expect(v01.value).toBeCloseTo(-0.7, 10)
		})

		test('MUNIT(3) returns 3x3 identity via spill', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'MUNIT(3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(0))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(1))
			expect(getResult(wb, 2, 2)).toEqual(numberValue(1))
		})

		test('MUNIT(0) returns #VALUE!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'MUNIT(0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#VALUE!'))
		})
	})

	describe('LAMBDA helpers', () => {
		test('BYROW applies LAMBDA to each row', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 0, 5)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 0, 3, 'BYROW(A1:B3,LAMBDA(r,SUM(r)))')
			recalc(wb)
			expect(getResult(wb, 0, 3)).toEqual(numberValue(3))
			expect(getResult(wb, 1, 3)).toEqual(numberValue(7))
			expect(getResult(wb, 2, 3)).toEqual(numberValue(11))
		})

		test('BYCOL applies LAMBDA to each column', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 3, 0, 'BYCOL(A1:B2,LAMBDA(c,SUM(c)))')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(4))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(6))
		})

		test('MAKEARRAY generates array with LAMBDA', () => {
			const wb = makeWorkbook()
			setFormula(wb, 5, 0, 'MAKEARRAY(3,3,LAMBDA(r,c,r*c))')
			recalc(wb)
			expect(getResult(wb, 5, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 5, 2)).toEqual(numberValue(3))
			expect(getResult(wb, 6, 1)).toEqual(numberValue(4))
			expect(getResult(wb, 7, 2)).toEqual(numberValue(9))
		})

		test('MAKEARRAY(0,1,...) returns #VALUE!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'MAKEARRAY(0,1,LAMBDA(r,c,r))')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#VALUE!'))
		})

		test('BYROW with non-scalar lambda returns #CALC!', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setFormula(wb, 2, 0, 'BYROW(A1:B1,LAMBDA(r,r))')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#CALC!'))
		})
	})

	describe('statistical distributions', () => {
		test('GAMMALN(5) = ln(4!) = ln(24)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'GAMMALN(5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(Math.log(24), 8)
		})

		test('GAMMA(5) = 4! = 24', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'GAMMA(5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(24, 8)
		})

		test('T.DIST(1, 10, TRUE) CDF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'T.DIST(1, 10, TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.8295534, 4)
		})

		test('T.INV(0.5, 10) = 0', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'T.INV(0.5, 10)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0, 5)
		})

		test('CHISQ.DIST(3, 5, TRUE) CDF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CHISQ.DIST(3, 5, TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.30001, 3)
		})

		test('BINOM.DIST(3, 10, 0.5, FALSE) PMF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BINOM.DIST(3, 10, 0.5, FALSE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.11719, 4)
		})

		test('POISSON.DIST(3, 5, FALSE) PMF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'POISSON.DIST(3, 5, FALSE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.14037, 4)
		})

		test('EXPON.DIST(1, 2, TRUE) CDF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'EXPON.DIST(1, 2, TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1 - Math.exp(-2), 10)
		})

		test('WEIBULL.DIST(1, 2, 3, TRUE) CDF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'WEIBULL.DIST(1, 2, 3, TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1 - Math.exp(-((1 / 3) ** 2)), 10)
		})

		test('BETA.DIST(0.5, 2, 3, TRUE)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BETA.DIST(0.5, 2, 3, TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.6875, 4)
		})

		test('LOGNORM.DIST CDF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'LOGNORM.DIST(4, 3, 2, TRUE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') {
				expect(r.value).toBeGreaterThan(0)
				expect(r.value).toBeLessThan(1)
			}
		})

		test('FISHER(0.5)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FISHER(0.5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5 * Math.log(3), 10)
		})

		test('FISHERINV(FISHER(0.5)) = 0.5', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FISHERINV(FISHER(0.5))')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5, 10)
		})

		test('STANDARDIZE(42, 40, 1.5)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'STANDARDIZE(42, 40, 1.5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(4 / 3, 10)
		})

		test('CONFIDENCE.NORM(0.05, 1, 100)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'CONFIDENCE.NORM(0.05, 1, 100)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.196, 2)
		})

		test('LINEST returns slope and intercept', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 2, 0, 3)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 1, 4)
			setNum(wb, 2, 1, 6)
			setFormula(wb, 4, 0, 'LINEST(B1:B3,A1:A3)')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(0))
		})
	})

	describe('info functions - SHEET/SHEETS', () => {
		test('SHEETS() returns sheet count', () => {
			const wb = makeWorkbook()
			wb.addSheet('Sheet2')
			setFormula(wb, 0, 0, 'SHEETS()')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(2))
		})

		test('SHEET() returns current sheet number', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SHEET()')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})

		test('SHEET("Sheet2") resolves sheet name to index', () => {
			const wb = makeWorkbook()
			wb.addSheet('Sheet2')
			setFormula(wb, 0, 0, 'SHEET("Sheet2")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(2))
		})

		test('SHEET(Sheet2!A1) resolves reference sheet index', () => {
			const wb = makeWorkbook()
			wb.addSheet('Sheet2')
			wb.sheets[1]?.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
			setFormula(wb, 0, 0, 'SHEET(Sheet2!A1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(2))
		})

		test('SHEETS(Sheet2!A1) returns 1 for single-sheet ref', () => {
			const wb = makeWorkbook()
			wb.addSheet('Sheet2')
			wb.sheets[1]?.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
			setFormula(wb, 0, 0, 'SHEETS(Sheet2!A1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})
	})

	describe('info functions - CELL', () => {
		test('CELL("address", B3) returns absolute address', () => {
			const wb = makeWorkbook()
			setNum(wb, 2, 1, 99)
			setFormula(wb, 0, 0, 'CELL("address",B3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('$B$3'))
		})

		test('CELL("row", B3) and CELL("col", B3) return coordinates', () => {
			const wb = makeWorkbook()
			setNum(wb, 2, 1, 99)
			setFormula(wb, 0, 0, 'CELL("row",B3)')
			setFormula(wb, 0, 1, 'CELL("col",B3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(3))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(2))
		})

		test('CELL("contents", ref) returns cell contents', () => {
			const wb = makeWorkbook()
			setStr(wb, 1, 0, 'hello')
			setFormula(wb, 0, 0, 'CELL("contents",A2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('hello'))
		})

		test('CELL("type", ref) classifies blank, label, and value', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'hello')
			setNum(wb, 1, 0, 42)
			setFormula(wb, 2, 0, 'CELL("type",A4)')
			setFormula(wb, 2, 1, 'CELL("type",A1)')
			setFormula(wb, 2, 2, 'CELL("type",A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(stringValue('b'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('l'))
			expect(getResult(wb, 2, 2)).toEqual(stringValue('v'))
		})
	})

	describe('ISFORMULA', () => {
		test('ISFORMULA returns TRUE for formula cells', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, '1+1')
			setFormula(wb, 1, 0, 'ISFORMULA(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(booleanValue(true))
		})

		test('ISFORMULA returns FALSE for value cells', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 42)
			setFormula(wb, 1, 0, 'ISFORMULA(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(booleanValue(false))
		})

		test('ISFORMULA returns FALSE for empty cells', () => {
			const wb = makeWorkbook()
			setFormula(wb, 1, 0, 'ISFORMULA(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(booleanValue(false))
		})

		test('ISFORMULA returns TRUE for error-producing formulas', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, '1/0')
			setFormula(wb, 1, 0, 'ISFORMULA(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(booleanValue(true))
		})
	})
})
