import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import {
	booleanValue,
	dateValue,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
} from '@ascend/schema'

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

function recalc(
	wb: ReturnType<typeof createWorkbook>,
	overrides?: Parameters<typeof defaultCalcContext>[0],
) {
	return recalculate(wb, defaultCalcContext(overrides))
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

		test('aggregates ignore nonnumeric referenced scalar cells but coerce typed literals', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '5')
			setBool(wb, 1, 0, true)
			setNum(wb, 2, 0, 3)
			setFormula(wb, 0, 1, 'SUM(A1:A3)')
			setFormula(wb, 1, 1, 'SUM(A1,A2,A3)')
			setFormula(wb, 2, 1, 'SUM("5",TRUE,A3)')
			setFormula(wb, 3, 1, 'MAX(A1,A2,A3)')
			setFormula(wb, 4, 1, 'PRODUCT(A1,A2,A3)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(9))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(3))
		})

		test('aggregates still propagate referenced errors', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: S0 })
			setFormula(wb, 1, 0, 'SUM(A1)')
			setFormula(wb, 2, 0, 'MAX(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#DIV/0!'))
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#DIV/0!'))
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

		test('COUNT counts numeric text and booleans supplied directly', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '1')
			setBool(wb, 1, 0, true)
			setFormula(wb, 2, 0, 'COUNT("1", TRUE, "nope", "")')
			setFormula(wb, 3, 0, 'COUNT(A1:A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(2))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(0))
		})

		test('COUNT and COUNTA handle direct error values like Excel', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COUNT(2, "A", "", #REF!, #DIV/0!)')
			setFormula(wb, 1, 0, 'COUNTA(2, "A", "", #REF!, #DIV/0!)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(5))
		})

		test('COUNT and COUNTA handle referenced error cells like Excel', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: S0 })
			setFormula(wb, 1, 0, 'COUNT(A1)')
			setFormula(wb, 2, 0, 'COUNTA(A1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
		})

		test('scalar numeric contexts reject empty text while aggregates keep direct empty text as zero', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, '""')
			setFormula(wb, 1, 0, 'COS(A1)')
			setFormula(wb, 2, 0, 'A1+1')
			setFormula(wb, 3, 0, 'SUM("")')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(0))
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

		test('SUMIF uses third argument as top-left of criteria-shaped sum range', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'x')
			setStr(wb, 1, 0, 'y')
			setStr(wb, 2, 0, 'x')
			setNum(wb, 0, 1, 10)
			setNum(wb, 1, 1, 20)
			setNum(wb, 2, 1, 30)
			setFormula(wb, 3, 0, 'SUMIF(A1:A3,"x",B1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(40))
		})

		test('AVERAGEIF uses third argument as top-left of criteria-shaped average range', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'x')
			setStr(wb, 1, 0, 'y')
			setStr(wb, 2, 0, 'x')
			setNum(wb, 0, 1, 10)
			setNum(wb, 1, 1, 20)
			setNum(wb, 2, 1, 30)
			setFormula(wb, 3, 0, 'AVERAGEIF(A1:A3,"x",B1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(20))
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

		test('COUNTIF escaped wildcards match literal characters', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '*')
			setStr(wb, 1, 0, '?')
			setStr(wb, 2, 0, '~')
			setStr(wb, 3, 0, '~*')
			setFormula(wb, 4, 0, 'COUNTIF(A1:A4,"~*")')
			setFormula(wb, 5, 0, 'COUNTIF(A1:A4,"~?")')
			setFormula(wb, 6, 0, 'COUNTIF(A1:A4,"~~")')
			recalc(wb)
			expect(getResult(wb, 4, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 5, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 6, 0)).toEqual(numberValue(1))
		})

		test('COUNTIF uses error criteria as match values', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: S0 })
			setNum(wb, 1, 0, 1)
			setFormula(wb, 2, 0, 'COUNTIF(A1:A2,1/0)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
		})

		test('COUNTIF criteria cache keeps error and string criteria distinct', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: S0 })
			setStr(wb, 1, 0, '#DIV/0!')
			wb.sheets[0]?.cells.set(0, 1, { value: errorValue('#DIV/0!'), formula: null, styleId: S0 })
			setStr(wb, 1, 1, '#DIV/0!')
			setFormula(wb, 2, 0, 'COUNTIF(A1:A2,B1)')
			setFormula(wb, 3, 0, 'COUNTIF(A1:A2,B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(1))
		})

		test('COUNTIF implicitly intersects range criteria by formula column', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 1, 1)
			setNum(wb, 0, 2, 2)
			setNum(wb, 0, 3, 3)
			setNum(wb, 0, 5, 2)
			setNum(wb, 0, 6, 2)
			setNum(wb, 0, 7, 3)
			setFormula(wb, 2, 2, 'COUNTIF(F1:H1,B1:D1)')
			recalc(wb)
			expect(getResult(wb, 2, 2)).toEqual(numberValue(2))
		})

		test('COUNTIF and SUMIF use #VALUE! criteria when range criteria cannot intersect', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, { value: errorValue('#VALUE!'), formula: null, styleId: S0 })
			setNum(wb, 1, 0, 2)
			wb.sheets[0]?.cells.set(2, 0, { value: errorValue('#VALUE!'), formula: null, styleId: S0 })
			setNum(wb, 0, 3, 5)
			setNum(wb, 1, 3, 6)
			setNum(wb, 2, 3, 7)
			setFormula(wb, 4, 4, 'COUNTIF(A1:A3,B1:C1)')
			setFormula(wb, 5, 4, 'SUMIF(A1:A3,B1:C1,D1)')
			recalc(wb)
			expect(getResult(wb, 4, 4)).toEqual(numberValue(2))
			expect(getResult(wb, 5, 4)).toEqual(numberValue(12))
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

		test('SUMPRODUCT returns #VALUE! for text in range operands', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '4')
			setNum(wb, 0, 1, 2)
			setFormula(wb, 1, 0, 'SUMPRODUCT(A1:A1,B1:B1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#VALUE!'))
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

		test('RIGHT with zero characters returns empty string', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'RIGHT("world",0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue(''))
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

		test('byte-oriented text compatibility functions mirror text semantics for ASCII', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'foo bar baz')
			setFormula(wb, 0, 1, 'LENB(A1)')
			setFormula(wb, 1, 1, 'LEFTB(A1,3)')
			setFormula(wb, 2, 1, 'RIGHTB(A1,3)')
			setFormula(wb, 3, 1, 'MIDB(A1,9,3)')
			setFormula(wb, 4, 1, 'FINDB("bar",A1)')
			setFormula(wb, 5, 1, 'SEARCHB("BAR",A1)')
			setFormula(wb, 6, 1, 'REPLACEB(A1,5,3,"FOO")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(11))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('foo'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('baz'))
			expect(getResult(wb, 3, 1)).toEqual(stringValue('baz'))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(5))
			expect(getResult(wb, 5, 1)).toEqual(numberValue(5))
			expect(getResult(wb, 6, 1)).toEqual(stringValue('foo FOO baz'))
		})

		test('ASC, PHONETIC, and legacy CHAR/CODE match Excel compatibility edges', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'ＡＢＣ　123')
			setFormula(wb, 0, 1, 'ASC(A1)')
			setFormula(wb, 1, 1, 'PHONETIC(A1)')
			setFormula(wb, 2, 1, 'CHAR(240)')
			setFormula(wb, 3, 1, 'CODE(CHAR(240))')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('ABC 123'))
			expect(getResult(wb, 1, 1)).toEqual(errorValue('#N/A'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('\uf8ff'))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(240))
		})

		test('SUBSTITUTE with instance_num replaces only Nth occurrence', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SUBSTITUTE("2026-Q1-Q1","Q1","Q2",2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('2026-Q1-Q2'))
		})

		test('SUBSTITUTE with empty old text returns original text', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'SUBSTITUTE("testblankmatch","","!",1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('testblankmatch'))
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

		test('TEXT supports locale year token e in date formats', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 28313)
			setFormula(wb, 0, 1, 'TEXT(A1,"ee-mm-dd")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1977-07-07'))
		})

		test('TEXT repeatedly recognizes long month date formats', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 44623)
			setFormula(wb, 0, 1, 'TEXT(A1,"mmmm dd, yyyy")')
			setFormula(wb, 1, 1, 'TEXT(A1,"mmmm dd, yyyy")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('March 03, 2022'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('March 03, 2022'))
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

		test('TEXT preserves text and booleans with text sections', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'jello')
			setBool(wb, 1, 0, true)
			setBool(wb, 2, 0, false)
			setFormula(wb, 0, 1, 'TEXT(A1,";;;@")')
			setFormula(wb, 1, 1, 'TEXT(A2,";;;@")')
			setFormula(wb, 2, 1, 'TEXT(A3,";;;@")')
			setFormula(wb, 0, 2, 'TEXT(A1,";;;""hi""")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('jello'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('TRUE'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('FALSE'))
			expect(getResult(wb, 0, 2)).toEqual(stringValue('hi'))
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

		test('TEXT supports one-year date tokens, uppercase tokens, and fractional seconds', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, {
				value: dateValue(17816.607951388887),
				formula: null,
				styleId: S0,
			})
			wb.sheets[0]?.cells.set(1, 0, {
				value: dateValue(36191.1702084375),
				formula: null,
				styleId: S0,
			})
			setFormula(wb, 0, 1, 'TEXT(A1,"d-m-y")')
			setFormula(wb, 0, 2, 'TEXT(A1,"DDD-MMM-YYY")')
			setFormula(wb, 1, 1, 'TEXT(A2,"h:m:s.00")')
			recalc(wb, { dateSystem: '1904' })
			expect(getResult(wb, 0, 1)).toEqual(stringValue('11-10-52'))
			expect(getResult(wb, 0, 2)).toEqual(stringValue('Sat-Oct-1952'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('4:5:6.01'))
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

		test('TEXT matches Excel scientific placeholder edges from POI fixtures', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 123456.789)
			setNum(wb, 1, 0, 0.0000123456789)
			setNum(wb, 2, 0, -0.0000123456789)
			setNum(wb, 3, 0, 1.2345678e-142)
			setFormula(wb, 0, 1, 'TEXT(A1,"|#,e-#|")')
			setFormula(wb, 0, 2, 'TEXT(A1,"|#.#|E+|#|")')
			setFormula(wb, 1, 1, 'TEXT(A2,"|##.####|e+|#|")')
			setFormula(wb, 2, 1, 'TEXT(A3,"|#|e+|#|")')
			setFormula(wb, 3, 1, 'TEXT(A4,"|#.#|e+|00000|")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('|1e5|'))
			expect(getResult(wb, 0, 2)).toEqual(stringValue('|1.2|E|+5|'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('|12.3457|e|-6|'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('-|1|e|-5|'))
			expect(getResult(wb, 3, 1)).toEqual(stringValue('|1.2|e|-00142|'))
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
			setNum(wb, 2, 0, 10)
			setNum(wb, 3, 0, 1)
			setNum(wb, 4, 0, -11)
			setFormula(wb, 0, 1, 'TEXT(A1,"[>=1000]0.0,""K"";0")')
			setFormula(wb, 1, 1, 'TEXT(A2,"[>=1000]0.0,""K"";0")')
			setFormula(wb, 2, 1, 'TEXT(A3,"[<10]#"" Wow""")')
			setFormula(wb, 3, 1, 'TEXT(A4,"[< 10]#"" Wow""")')
			setFormula(wb, 4, 1, 'TEXT(A5,"[<-10]#"" Wow""")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1.5K'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('12'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('10'))
			expect(getResult(wb, 3, 1)).toEqual(stringValue('1 Wow'))
			expect(getResult(wb, 4, 1)).toEqual(stringValue('-11 Wow'))
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

		test('TEXT General matches Excel cached-value formatting edges', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 110000000000)
			setNum(wb, 1, 0, 1.123e-10)
			setBool(wb, 2, 0, true)
			setNum(wb, 3, 0, 0)
			setFormula(wb, 0, 1, 'TEXT(A1,"General")')
			setFormula(wb, 1, 1, 'TEXT(A2,"General")')
			setFormula(wb, 2, 1, 'TEXT(A3,"General")')
			setFormula(wb, 3, 1, 'TEXT(A4,"")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1.1E+11'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('1.123E-10'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('TRUE'))
			expect(getResult(wb, 3, 1)).toEqual(stringValue(''))
		})

		test('TEXT supports currency locale markers', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 42)
			setFormula(wb, 0, 1, 'TEXT(A1,"[$USD]0.00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('USD42.00'))
		})

		test('TEXT ignores locale id markers while preserving date formatting', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 44927)
			setFormula(wb, 0, 1, 'TEXT(A1,"[$-409]m/d/yy")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1/1/23'))
		})

		test('TEXT supports elapsed hours tokens', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"[h]:mm")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('36:00'))
		})

		test('TEXT supports elapsed seconds and mixed time tokens', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 314159 / 100000)
			setFormula(wb, 0, 1, 'TEXT(A1,"[ss].000")')
			setFormula(wb, 0, 2, 'TEXT(A1,"s:m"" @ hour ""[hh]")')
			setFormula(wb, 0, 3, 'TEXT(A1,"[s]"" [yes, ""ss""] seconds""")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('271433.376'))
			expect(getResult(wb, 0, 2)).toEqual(stringValue('53:23 @ hour 75'))
			expect(getResult(wb, 0, 3)).toEqual(stringValue('271433 [yes, 271433] seconds'))
		})

		test('TEXT(1234.5, "#,##0.00") = "1,234.50"', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1234.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"#,##0.00")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1,234.50'))
		})

		test('TEXT(0.75, "0%") = "75%"', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.75)
			setFormula(wb, 0, 1, 'TEXT(A1,"0%")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('75%'))
		})

		test('TEXT(1234, "$#,##0") = "$1,234"', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1234)
			setFormula(wb, 0, 1, 'TEXT(A1,"$#,##0")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('$1,234'))
		})

		test('TEXT formats common fraction placeholders', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.5)
			setFormula(wb, 0, 1, 'TEXT(A1,"# ?/?")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('1/2'))
		})

		test('TEXT(44941, "yyyy-mm-dd") date format', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 44941)
			setFormula(wb, 0, 1, 'TEXT(A1,"yyyy-mm-dd")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('2023-01-15'))
		})

		test('TEXT(-5, "[Red]0") strips color codes', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -5)
			setFormula(wb, 0, 1, 'TEXT(A1,"[Red]0")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('-5'))
		})

		test('TEXT with conditional color format [Red]#,##0;[Blue]-#,##0', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 100)
			setNum(wb, 1, 0, -100)
			setFormula(wb, 0, 1, 'TEXT(A1,"[Red]#,##0;[Blue]-#,##0")')
			setFormula(wb, 1, 1, 'TEXT(A2,"[Red]#,##0;[Blue]-#,##0")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('100'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('-100'))
		})

		test('TEXT with custom text "Total: "0.00', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 42.5)
			setStr(wb, 0, 2, '"Total: "0.00')
			setFormula(wb, 0, 1, 'TEXT(A1,C1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('Total: 42.50'))
		})

		test('TEXT with multiple sections #,##0;-#,##0;"zero";@', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 100)
			setNum(wb, 1, 0, -100)
			setNum(wb, 2, 0, 0)
			setFormula(wb, 0, 1, 'TEXT(A1,"#,##0;-#,##0;""zero"";@")')
			setFormula(wb, 1, 1, 'TEXT(A2,"#,##0;-#,##0;""zero"";@")')
			setFormula(wb, 2, 1, 'TEXT(A3,"#,##0;-#,##0;""zero"";@")')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('100'))
			expect(getResult(wb, 1, 1)).toEqual(stringValue('-100'))
			expect(getResult(wb, 2, 1)).toEqual(stringValue('zero'))
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

		test('lookup functions preserve scalar input error identity', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, {
				value: errorValue('#DIV/0!'),
				formula: null,
				styleId: S0,
			})
			setFormula(wb, 1, 0, 'VLOOKUP(A1,B1:C1,2,FALSE)')
			setFormula(wb, 2, 0, 'HLOOKUP(A1,B1:C2,2,FALSE)')
			setFormula(wb, 3, 0, 'INDEX(A1,1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#DIV/0!'))
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#DIV/0!'))
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#DIV/0!'))
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

		test('VLOOKUP and HLOOKUP approximate mode prefer exact hits before approximate fallback', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'zeta')
			setStr(wb, 1, 0, 'target')
			setStr(wb, 2, 0, 'alpha')
			setStr(wb, 0, 1, 'wrong')
			setStr(wb, 1, 1, 'exact')
			setStr(wb, 2, 1, 'fallback')
			setFormula(wb, 3, 0, 'VLOOKUP("target",A1:B3,2,TRUE)')

			setStr(wb, 5, 0, 'zeta')
			setStr(wb, 5, 1, 'target')
			setStr(wb, 5, 2, 'alpha')
			setStr(wb, 6, 0, 'wrong')
			setStr(wb, 6, 1, 'exact')
			setStr(wb, 6, 2, 'fallback')
			setFormula(wb, 7, 0, 'HLOOKUP("target",A6:C7,2,TRUE)')

			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('exact'))
			expect(getResult(wb, 7, 0)).toEqual(stringValue('exact'))
		})

		test('VLOOKUP and HLOOKUP approximate blank lookup miss before return-index errors', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'VLOOKUP(B1,B1,B1,3)')
			setFormula(wb, 1, 0, 'HLOOKUP(B1,B1,B1,3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#N/A'))
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#N/A'))
		})

		test('LOOKUP prefers exact matches before approximate fallback', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 534)
			setNum(wb, 2, 0, 2)
			setStr(wb, 0, 1, 'low')
			setStr(wb, 1, 1, 'exact')
			setStr(wb, 2, 1, 'fallback')
			setFormula(wb, 3, 0, 'LOOKUP(534,A1:A3,B1:B3)')
			setNum(wb, 5, 0, 10)
			setNum(wb, 6, 0, 20)
			setNum(wb, 7, 0, 30)
			setStr(wb, 5, 1, 'small')
			setStr(wb, 6, 1, 'medium')
			setStr(wb, 7, 1, 'large')
			setFormula(wb, 4, 0, 'LOOKUP(25,A6:A8,B6:B8)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(stringValue('exact'))
			expect(getResult(wb, 4, 0)).toEqual(stringValue('medium'))
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

		test('MATCH with ascending approximate skips interleaved blank header cells', () => {
			const wb = makeWorkbook()
			wb.sheets[0]?.cells.set(0, 0, { value: dateValue(44562), formula: null, styleId: S0 })
			wb.sheets[0]?.cells.set(0, 1, { value: EMPTY, formula: null, styleId: S0 })
			wb.sheets[0]?.cells.set(0, 2, { value: dateValue(44593), formula: null, styleId: S0 })
			wb.sheets[0]?.cells.set(0, 3, { value: EMPTY, formula: null, styleId: S0 })
			setFormula(wb, 1, 0, 'MATCH(44568,A1:D1,1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
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

		test('exact lookup treats date serials and numbers consistently', () => {
			const wb = makeWorkbook()
			wb.getSheet('Sheet1')?.cells.set(0, 0, {
				value: dateValue(45000),
				formula: null,
				styleId: S0,
			})
			setNum(wb, 0, 1, 123)
			setFormula(wb, 1, 0, 'VLOOKUP(45000,A1:B1,2,FALSE)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(123))
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

		test('AVERAGEA includes referenced scalar text and booleans', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'TRUE')
			setBool(wb, 1, 0, true)
			setNum(wb, 2, 0, 2)
			setFormula(wb, 0, 1, 'AVERAGEA(A1,2)')
			setFormula(wb, 1, 1, 'AVERAGEA(A2,2)')
			setFormula(wb, 2, 1, 'STDEVA(A1,A2,A3)')
			setFormula(wb, 3, 1, 'VARA(A1,A2,A3)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(1.5))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(1))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(1))
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

		test('LARGE and SMALL return ranked values from an unsorted range', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 130; i++) setNum(wb, i, 0, ((i * 37) % 130) + 1)
			setFormula(wb, 0, 1, 'LARGE(A1:A130,1)')
			setFormula(wb, 1, 1, 'LARGE(A1:A130,64)')
			setFormula(wb, 2, 1, 'LARGE(A1:A130,65)')
			setFormula(wb, 3, 1, 'SMALL(A1:A130,1)')
			setFormula(wb, 4, 1, 'SMALL(A1:A130,64)')
			setFormula(wb, 5, 1, 'SMALL(A1:A130,65)')
			setFormula(wb, 6, 1, 'LARGE({5,1,9,3},1)')
			setFormula(wb, 7, 1, 'SMALL({5,1,9,3},1)')
			setFormula(wb, 8, 1, 'LARGE({5,1,9,3},2)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(130))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(67))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(66))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(1))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(64))
			expect(getResult(wb, 5, 1)).toEqual(numberValue(65))
			expect(getResult(wb, 6, 1)).toEqual(numberValue(9))
			expect(getResult(wb, 7, 1)).toEqual(numberValue(1))
			expect(getResult(wb, 8, 1)).toEqual(numberValue(5))
		})

		test('AGGREGATE rank and percentile functions read array before k', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 10; i++) setNum(wb, i, 0, i + 1)
			setFormula(wb, 10, 0, 'SUBTOTAL(9,A1:A9)')
			setFormula(wb, 0, 1, 'AGGREGATE(14,0,A1:A10,2)')
			setFormula(wb, 1, 1, 'AGGREGATE(15,0,A1:A10,2)')
			setFormula(wb, 2, 1, 'AGGREGATE(16,0,A1:A10,0.4)')
			setFormula(wb, 3, 1, 'AGGREGATE(18,0,A1:A10,0.4)')
			setFormula(wb, 4, 1, 'AGGREGATE(1,0,A1:A11)')
			recalc(wb)

			expect(getResult(wb, 0, 1)).toEqual(numberValue(9))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(4.6))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(4.4))
			expect(getResult(wb, 4, 1)).toEqual(numberValue(5.5))
		})

		test('MODE and AGGREGATE mode return the first value among tied modes', () => {
			const wb = makeWorkbook()
			for (const [row, value] of [77, 28, 28, 77].entries()) setNum(wb, row, 0, value)
			setFormula(wb, 0, 1, 'MODE.SNGL(A1:A4)')
			setFormula(wb, 1, 1, 'AGGREGATE(13,0,A1:A4)')
			recalc(wb)

			expect(getResult(wb, 0, 1)).toEqual(numberValue(77))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(77))
		})

		test('AGGREGATE maps function 8 to STDEV.P', () => {
			const wb = makeWorkbook()
			for (const [row, value] of [2, 4, 6].entries()) setNum(wb, row, 0, value)
			setFormula(wb, 0, 1, 'AGGREGATE(8,0,A1:A3)')
			recalc(wb)

			const result = getResult(wb, 0, 1)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBeCloseTo(Math.sqrt(8 / 3), 12)
			}
		})

		test('STDEV and VAR ignore nonnumeric referenced scalars but coerce typed literals', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '10')
			setBool(wb, 1, 0, true)
			setNum(wb, 2, 0, 1)
			setNum(wb, 3, 0, 2)
			setNum(wb, 4, 0, 3)
			setFormula(wb, 0, 1, 'STDEV(A1,A2,A3:A5)')
			setFormula(wb, 1, 1, 'VAR(A1,A2,A3:A5)')
			setFormula(wb, 2, 1, 'VARP(A1,A2,A3:A5)')
			setFormula(wb, 3, 1, 'STDEV("10",TRUE,A3:A5)')
			setFormula(wb, 4, 1, 'VAR("10",TRUE,A3:A5)')
			setFormula(wb, 5, 1, 'VAR("nope",1,2)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(1))
			const varp = getResult(wb, 2, 1)
			expect(varp?.kind).toBe('number')
			if (varp?.kind === 'number') expect(varp.value).toBeCloseTo(2 / 3)
			const directStdev = getResult(wb, 3, 1)
			expect(directStdev?.kind).toBe('number')
			if (directStdev?.kind === 'number') expect(directStdev.value).toBeCloseTo(Math.sqrt(14.3))
			const directVar = getResult(wb, 4, 1)
			expect(directVar?.kind).toBe('number')
			if (directVar?.kind === 'number') expect(directVar.value).toBeCloseTo(14.3)
			expect(getResult(wb, 5, 1)).toEqual(errorValue('#VALUE!'))
		})

		test('VAR and VARP stay stable for large-magnitude mixed ranges', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 1.001)
			setNum(wb, 1, 0, 9_999_999_999)
			setNum(wb, 1, 1, 2.3456)
			setNum(wb, 2, 0, 0.00001)
			setStr(wb, 2, 1, 'alphanum')
			setFormula(wb, 3, 0, 'VAR(A1:B3)')
			setFormula(wb, 3, 1, 'VARP(A1:B3)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(19_999_999_991_653_392_000))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(15_999_999_993_322_713_000))
		})

		test('MODE errors on direct text but ignores text in references', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setStr(wb, 1, 0, '1')
			setNum(wb, 2, 0, 2)
			setNum(wb, 3, 0, 2)
			setFormula(wb, 0, 1, 'MODE(A1:A4)')
			setFormula(wb, 1, 1, 'MODE(A1,"1",2,3)')
			setFormula(wb, 2, 1, 'MODE.MULT(A1,"1",2,3)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(2))
			expect(getResult(wb, 1, 1)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 2, 1)).toEqual(errorValue('#VALUE!'))
		})

		test('MODE errors on direct blank cell arguments', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setFormula(wb, 1, 0, 'MODE(A1,2,B1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#VALUE!'))
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

		test('SUMSQ ignores nonnumeric referenced scalar cells but coerces typed literals', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '3')
			setNum(wb, 1, 0, 4)
			setFormula(wb, 2, 0, 'SUMSQ(A1,A2)')
			setFormula(wb, 3, 0, 'SUMSQ("3",A2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(16))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(25))
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

		test('ECMA.CEILING aliases legacy CEILING behavior', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ECMA.CEILING(1.2,2)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(2))
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

		test('FLOOR returns zero for zero number and zero significance', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FLOOR(0,0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0))
		})

		test('FACT returns #NUM! when the factorial overflows', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FACT(171)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#NUM!'))
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

		test('DOLLAR formats negative currency with parentheses', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DOLLAR(-1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('($1.00)'))
		})

		test('DOLLAR accepts currency-formatted text', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, '$1.00')
			setFormula(wb, 0, 1, 'DOLLAR(A1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(stringValue('$1.00'))
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

		test('WORKDAY.INTL treats omitted weekend argument as Saturday/Sunday', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'WORKDAY.INTL(DATE(2014,4,1),24,,)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(41764))
		})
	})

	describe('date functions - 1900 leap year bug', () => {
		test('DATEVALUE("1900-01-01") → serial 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATEVALUE("1900-01-01")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})

		test('DATEVALUE("1900-02-28") → serial 59', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATEVALUE("1900-02-28")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(59))
		})

		test('DATEVALUE("1900-03-01") → serial 61 (skips phantom serial 60)', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATEVALUE("1900-03-01")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(61))
		})

		test('DATEVALUE applies Excel two-digit year cutoff', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATEVALUE("1/1/29")')
			setFormula(wb, 0, 1, 'DATEVALUE("1/1/30")')
			setFormula(wb, 0, 2, 'DATEVALUE("7/5/98")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(47119))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(10959))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(35981))
		})

		test('DATEVALUE parses English month-name dates', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATEVALUE("June 9, 1969")')
			setFormula(wb, 0, 1, 'DATEVALUE("Jun 9 1969")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(25363))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(25363))
		})

		test('DATE(1900,1,1) → serial 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATE(1900,1,1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
		})

		test('DATE(1900,2,29) → serial 60 phantom leap day', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATE(1900,2,29)')
			setFormula(wb, 0, 1, 'DATE(1900,1,60)')
			setFormula(wb, 0, 2, 'DATE(1900,3,0)')
			setFormula(wb, 0, 3, 'DATEVALUE("1900-02-29")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(60))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(60))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(60))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(60))
		})

		test('DATE(2024,1,1) → serial 45292', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATE(2024,1,1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(45292))
		})

		test('YEAR/MONTH/DAY round-trip through serial 60 (phantom Feb 29, 1900)', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 60)
			setFormula(wb, 0, 1, 'YEAR(A1)')
			setFormula(wb, 0, 2, 'MONTH(A1)')
			setFormula(wb, 0, 3, 'DAY(A1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1900))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(2))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(29))
		})

		test('YEAR/MONTH/DAY for serial 61 = Mar 1, 1900', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 61)
			setFormula(wb, 0, 1, 'YEAR(A1)')
			setFormula(wb, 0, 2, 'MONTH(A1)')
			setFormula(wb, 0, 3, 'DAY(A1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1900))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(3))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(1))
		})

		test('YEAR/MONTH/DAY for 1900-system serial zero follows Excel date-part behavior', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0)
			setFormula(wb, 0, 1, 'YEAR(A1)')
			setFormula(wb, 0, 2, 'MONTH(A1)')
			setFormula(wb, 0, 3, 'DAY(A1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1900))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(1))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(0))
		})

		test('WEEKDAY for dates before March 1, 1900', () => {
			const wb = makeWorkbook()
			// Serial 1 = Jan 1, 1900 → Sunday (1) in Excel (historically Monday, but Excel says Sunday)
			setFormula(wb, 0, 0, 'WEEKDAY(1,1)')
			// Serial 59 = Feb 28, 1900 → Tuesday (3)
			setFormula(wb, 0, 1, 'WEEKDAY(59,1)')
			// Serial 60 = phantom Feb 29, 1900 → Wednesday (4)
			setFormula(wb, 0, 2, 'WEEKDAY(60,1)')
			// Serial 61 = Mar 1, 1900 → Thursday (5)
			setFormula(wb, 0, 3, 'WEEKDAY(61,1)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(3))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(4))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(5))
		})

		test('serial numbers are contiguous across the phantom date gap', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DATE(1900,2,28)')
			setFormula(wb, 0, 1, 'DATE(1900,3,1)')
			setFormula(wb, 0, 2, 'DATE(1900,3,1)-DATE(1900,2,28)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(59))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(61))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(2))
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

		test('database criteria rows ignore blank criteria cells', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'Team')
			setStr(wb, 0, 1, 'Score')
			setStr(wb, 1, 0, 'A')
			setNum(wb, 1, 1, 10)
			setStr(wb, 2, 0, 'B')
			setNum(wb, 2, 1, 20)
			setStr(wb, 3, 0, 'B')
			setNum(wb, 3, 1, 30)
			setStr(wb, 5, 0, 'Team')
			setStr(wb, 5, 1, 'Score')
			setStr(wb, 6, 0, 'B')
			setFormula(wb, 7, 0, 'DSUM(A1:B4,2,A6:B7)')
			recalc(wb)
			expect(getResult(wb, 7, 0)).toEqual(numberValue(50))
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

		test('SORT rejects invalid sort indexes and orders', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 30)
			setNum(wb, 1, 0, 10)
			setNum(wb, 2, 0, 20)
			setFormula(wb, 3, 0, 'SORT(A1:A3,0)')
			setFormula(wb, 4, 0, 'SORT(A1:A3,2)')
			setFormula(wb, 5, 0, 'SORT(A1:A3,1,0)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 4, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 5, 0)).toEqual(errorValue('#VALUE!'))
		})

		test('SORT by_col rejects indexes outside the row axis', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 3)
			setNum(wb, 0, 1, 1)
			setNum(wb, 0, 2, 2)
			setNum(wb, 1, 0, 30)
			setNum(wb, 1, 1, 10)
			setNum(wb, 1, 2, 20)
			setFormula(wb, 2, 0, 'SORT(A1:C2,3,1,TRUE)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#VALUE!'))
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

		test('SORTBY rejects mismatched sort arrays and invalid orders', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'a')
			setStr(wb, 1, 0, 'b')
			setStr(wb, 2, 0, 'c')
			setNum(wb, 0, 1, 1)
			setNum(wb, 1, 1, 2)
			setNum(wb, 2, 1, 3)
			setFormula(wb, 3, 0, 'SORTBY(A1:A3,B1:B2)')
			setFormula(wb, 4, 0, 'SORTBY(A1:A3,B1:B3,0)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 4, 0)).toEqual(errorValue('#VALUE!'))
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

		test('RANDARRAY with seed produces reproducible output', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'RANDARRAY(2,2,0,1,0,42)')
			recalc(wb)
			const a = getResult(wb, 0, 0)
			const b = getResult(wb, 0, 1)
			const c = getResult(wb, 1, 0)
			const d = getResult(wb, 1, 1)
			expect(a?.kind).toBe('number')
			expect(b?.kind).toBe('number')
			expect(c?.kind).toBe('number')
			expect(d?.kind).toBe('number')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(a)
			expect(getResult(wb, 0, 1)).toEqual(b)
			expect(getResult(wb, 1, 0)).toEqual(c)
			expect(getResult(wb, 1, 1)).toEqual(d)
		})

		test('RANDARRAY with invalid seed returns #VALUE!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'RANDARRAY(1,1,0,1,0,"x")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(errorValue('#VALUE!'))
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

		test('EXPAND rejects dimensions that cannot contain the source array', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 2)
			setNum(wb, 1, 0, 3)
			setNum(wb, 1, 1, 4)
			setFormula(wb, 3, 0, 'EXPAND(A1:B2,1,2)')
			setFormula(wb, 4, 0, 'EXPAND(A1:B2,2,1)')
			setFormula(wb, 5, 0, 'EXPAND(A1:B2,0,2)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 4, 0)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 5, 0)).toEqual(errorValue('#VALUE!'))
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

		test('SLOPE propagates scalar name errors before paired range shape checks', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setFormula(wb, 2, 0, 'SLOPE(A1:A2,missing_name)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#NAME?'))
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

		test('regression functions stay stable with large offsets', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 5; i++) {
				const x = 1_000_000_000_000 + i + 1
				setNum(wb, i, 0, x)
				setNum(wb, i, 1, 3 * x - 7)
			}
			setFormula(wb, 5, 0, 'SLOPE(B1:B5,A1:A5)')
			setFormula(wb, 5, 1, 'INTERCEPT(B1:B5,A1:A5)')
			setFormula(wb, 5, 2, 'FORECAST.LINEAR(1000000000006,B1:B5,A1:A5)')
			setFormula(wb, 6, 0, 'CORREL(B1:B5,A1:A5)')
			setFormula(wb, 6, 1, 'PEARSON(B1:B5,A1:A5)')
			setFormula(wb, 6, 2, 'RSQ(B1:B5,A1:A5)')
			setFormula(wb, 7, 0, 'STEYX(B1:B5,A1:A5)')
			recalc(wb)

			const slope = getResult(wb, 5, 0)
			const intercept = getResult(wb, 5, 1)
			const forecast = getResult(wb, 5, 2)
			const correl = getResult(wb, 6, 0)
			const pearson = getResult(wb, 6, 1)
			const rsq = getResult(wb, 6, 2)
			const steyx = getResult(wb, 7, 0)
			expect(slope?.kind).toBe('number')
			expect(intercept?.kind).toBe('number')
			expect(forecast?.kind).toBe('number')
			expect(correl?.kind).toBe('number')
			expect(pearson?.kind).toBe('number')
			expect(rsq?.kind).toBe('number')
			expect(steyx?.kind).toBe('number')
			if (slope?.kind === 'number') expect(slope.value).toBeCloseTo(3)
			if (intercept?.kind === 'number') expect(intercept.value).toBeCloseTo(-7)
			if (forecast?.kind === 'number') expect(forecast.value).toBe(3_000_000_000_011)
			if (correl?.kind === 'number') expect(correl.value).toBeCloseTo(1)
			if (pearson?.kind === 'number') expect(pearson.value).toBeCloseTo(1)
			if (rsq?.kind === 'number') expect(rsq.value).toBeCloseTo(1)
			if (steyx?.kind === 'number') expect(steyx.value).toBeCloseTo(0)
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

		test('normal and lognormal functions match Excel cached-value precision points', () => {
			const wb = makeWorkbook()
			const sample = [1, 2, 3, 4, 5, 4]
			for (let i = 0; i < sample.length; i++) setNum(wb, i, 0, sample[i] as number)
			setFormula(wb, 0, 1, 'NORM.DIST(0,-1,3,TRUE)')
			setFormula(wb, 1, 1, 'NORM.S.DIST(3,TRUE)')
			setFormula(wb, 2, 1, 'NORM.S.INV(0.9986501019683699)')
			setFormula(wb, 3, 1, 'NORMSDIST(2)')
			setFormula(wb, 4, 1, 'LOGNORM.DIST(1,2,5,TRUE)')
			setFormula(wb, 5, 1, 'LOGNORM.INV(0.34457825838967576,2,5)')
			setFormula(wb, 6, 1, 'Z.TEST(A1:A6,3)')
			setFormula(wb, 7, 1, 'ZTEST(A1:A6,3,5)')
			recalc(wb)
			const expected = [
				0.6305586598182364, 0.9986501019683699, 3, 0.9772498680518208, 0.34457825838967576, 1,
				0.39075564749935676, 0.46746265582105867,
			]
			for (let row = 0; row < expected.length; row++) {
				const r = getResult(wb, row, 1)
				expect(r?.kind).toBe('number')
				if (r?.kind === 'number')
					expect(Math.abs(r.value - (expected[row] as number))).toBeLessThan(1e-12)
			}
			expect(getResult(wb, 2, 1)).toEqual(numberValue(3))
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

		test('XNPV and XIRR validate Excel date and sign requirements', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -100)
			setNum(wb, 1, 0, 110)
			setNum(wb, 0, 1, 44927.9)
			setNum(wb, 1, 1, 45292.2)
			setFormula(wb, 0, 2, 'XIRR(A1:A2,B1:B2)')
			setFormula(wb, 1, 2, 'XNPV(0.1,A1:A2,B1:B2)')

			setNum(wb, 3, 0, 100)
			setNum(wb, 4, 0, 110)
			setNum(wb, 3, 1, 44927)
			setNum(wb, 4, 1, 45292)
			setFormula(wb, 3, 2, 'XIRR(A4:A5,B4:B5)')
			setFormula(wb, 4, 2, 'XNPV(0.1,A4:A5,B4:B5)')

			setNum(wb, 6, 0, -100)
			setNum(wb, 7, 0, 110)
			setNum(wb, 6, 1, 44927)
			setNum(wb, 7, 1, 44926)
			setFormula(wb, 6, 2, 'XIRR(A7:A8,B7:B8)')
			setFormula(wb, 7, 2, 'XNPV(0.1,A7:A8,B7:B8)')

			recalc(wb)
			const xirr = getResult(wb, 0, 2)
			const xnpv = getResult(wb, 1, 2)
			expect(xirr?.kind).toBe('number')
			expect(xnpv?.kind).toBe('number')
			if (xirr?.kind === 'number') expect(xirr.value).toBeCloseTo(0.1, 12)
			if (xnpv?.kind === 'number') expect(xnpv.value).toBeCloseTo(0, 12)
			expect(getResult(wb, 3, 2)).toEqual(errorValue('#NUM!'))
			expect(getResult(wb, 4, 2)).toEqual(errorValue('#NUM!'))
			expect(getResult(wb, 6, 2)).toEqual(errorValue('#NUM!'))
			expect(getResult(wb, 7, 2)).toEqual(errorValue('#NUM!'))
		})

		test('XIRR falls back to a bracketed solve when Newton misses a negative root', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, -6572.746741492301)
			setNum(wb, 1, 0, 261.1244022846222)
			setNum(wb, 2, 0, 817.0722676441073)
			setNum(wb, 0, 1, 40000)
			setNum(wb, 1, 1, 42179)
			setNum(wb, 2, 1, 43257)
			setFormula(wb, 0, 2, 'XIRR(A1:A3,B1:B3)')
			setFormula(wb, 1, 2, 'XNPV(C1,A1:A3,B1:B3)')
			recalc(wb)
			const xirr = getResult(wb, 0, 2)
			const residual = getResult(wb, 1, 2)
			expect(xirr?.kind).toBe('number')
			expect(residual?.kind).toBe('number')
			if (xirr?.kind === 'number') expect(xirr.value).toBeCloseTo(-0.1944067716, 9)
			if (residual?.kind === 'number') expect(residual.value).toBeCloseTo(0, 6)
		})

		test('FV preserves Excel-compatible precision under cancellation', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'FV(2.95,13,13000,-17406.78521481564,1)')
			setFormula(wb, 1, 0, 'FV(2.95,13,13000,-4406.785442944962,0)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(333891.23010253906))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(333891.2300109863))
		})

		test('PMT and PPMT use stable finite annuity factors for short integer terms', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PMT(0.1/12,3,8000)')
			setFormula(wb, 1, 0, 'PPMT(0.1/12,3,3,8000)')
			recalc(wb)
			const pmtResult = getResult(wb, 0, 0)
			const ppmtResult = getResult(wb, 1, 0)
			expect(pmtResult?.kind).toBe('number')
			expect(ppmtResult?.kind).toBe('number')
			if (pmtResult?.kind === 'number')
				expect(pmtResult.value.toPrecision(15)).toBe('-2711.23405492681')
			if (ppmtResult?.kind === 'number')
				expect(ppmtResult.value.toPrecision(15)).toBe('-2688.82716191088')
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

		test('DISC and INTRATE use basis-aware US 30/360 February end-of-month days', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DISC(DATE(2024,2,29),DATE(2024,3,31),97,100,0)')
			setFormula(wb, 0, 1, 'INTRATE(DATE(2024,2,29),DATE(2024,3,31),97,100,0)')
			recalc(wb)
			const disc = getResult(wb, 0, 0)
			const intrate = getResult(wb, 0, 1)
			expect(disc?.kind).toBe('number')
			expect(intrate?.kind).toBe('number')
			if (disc?.kind === 'number') expect(disc.value).toBeCloseTo(0.36, 12)
			if (intrate?.kind === 'number') expect(intrate.value).toBeCloseTo(0.3711340206185567, 12)
		})

		test('DISC uses split-year Actual/Actual basis over leap years', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DISC(DATE(2023,7,1),DATE(2025,7,1),97,100,1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.015, 12)
		})

		test('coupon date helpers match expected semiannual schedule', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COUPPCD(DATE(2024,3,15),DATE(2024,12,31),2,0)')
			setFormula(wb, 0, 1, 'COUPNCD(DATE(2024,3,15),DATE(2024,12,31),2,0)')
			setFormula(wb, 0, 2, 'COUPDAYBS(DATE(2024,3,15),DATE(2024,12,31),2,0)')
			setFormula(wb, 0, 3, 'COUPDAYS(DATE(2024,3,15),DATE(2024,12,31),2,0)')
			setFormula(wb, 0, 4, 'COUPDAYSNC(DATE(2024,3,15),DATE(2024,12,31),2,0)')
			setFormula(wb, 0, 5, 'COUPNUM(DATE(2024,3,15),DATE(2024,12,31),2,0)')
			setFormula(wb, 1, 0, 'COUPPCD(DATE(2024,3,15),DATE(2024,12,31),2,0)=DATE(2023,12,31)')
			setFormula(wb, 1, 1, 'COUPNCD(DATE(2024,3,15),DATE(2024,12,31),2,0)=DATE(2024,6,30)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(booleanValue(true))
			expect(getResult(wb, 1, 1)).toEqual(booleanValue(true))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(75))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(180))
			expect(getResult(wb, 0, 4)).toEqual(numberValue(105))
			expect(getResult(wb, 0, 5)).toEqual(numberValue(2))
		})

		test('US 30/360 coupon day count handles February end-of-month', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'COUPPCD(DATE(2024,3,31),DATE(2024,8,31),2,0)')
			setFormula(wb, 0, 1, 'COUPDAYBS(DATE(2024,3,31),DATE(2024,8,31),2,0)')
			setFormula(wb, 0, 2, 'COUPDAYS(DATE(2024,3,31),DATE(2024,8,31),2,0)')
			setFormula(wb, 0, 3, 'COUPDAYSNC(DATE(2024,3,31),DATE(2024,8,31),2,0)')
			setFormula(wb, 1, 0, 'COUPPCD(DATE(2024,3,31),DATE(2024,8,31),2,0)=DATE(2024,2,29)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(booleanValue(true))
			expect(getResult(wb, 0, 1)).toEqual(numberValue(30))
			expect(getResult(wb, 0, 2)).toEqual(numberValue(180))
			expect(getResult(wb, 0, 3)).toEqual(numberValue(150))
		})

		test('PRICE and YIELD are approximately inverse', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PRICE(DATE(2024,3,15),DATE(2029,12,31),0.08,0.06,100,2,0)')
			setFormula(wb, 0, 1, 'YIELD(DATE(2024,3,15),DATE(2029,12,31),0.08,A1,100,2,0)')
			recalc(wb)
			const y = getResult(wb, 0, 1)
			expect(y?.kind).toBe('number')
			if (y?.kind === 'number') expect(y.value).toBeCloseTo(0.06, 4)
		})

		test('PRICE uses compound fractional discounting for single-coupon bonds', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PRICE(DATE(2024,3,31),DATE(2024,8,31),0.06,0.05,100,2,0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(100.4022080906, 9)
		})

		test('PRICEDISC and YIELDDISC are approximately inverse', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PRICEDISC(DATE(2024,1,1),DATE(2024,7,1),0.06,100,2)')
			setFormula(wb, 0, 1, 'YIELDDISC(DATE(2024,1,1),DATE(2024,7,1),A1,100,2)')
			recalc(wb)
			const y = getResult(wb, 0, 1)
			expect(y?.kind).toBe('number')
			if (y?.kind === 'number') expect(y.value).toBeCloseTo(0.06188, 4)
		})

		test('PRICEMAT and YIELDMAT are approximately inverse', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'PRICEMAT(DATE(2024,1,1),DATE(2024,7,1),DATE(2023,7,1),0.05,0.06,2)')
			setFormula(wb, 0, 1, 'YIELDMAT(DATE(2024,1,1),DATE(2024,7,1),DATE(2023,7,1),0.05,A1,2)')
			recalc(wb)
			const y = getResult(wb, 0, 1)
			expect(y?.kind).toBe('number')
			if (y?.kind === 'number') expect(y.value).toBeCloseTo(0.06, 4)
		})

		test('DURATION and MDURATION maintain modified-duration identity', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'DURATION(DATE(2024,3,15),DATE(2029,12,31),0.08,0.06,2,0)')
			setFormula(wb, 0, 1, 'MDURATION(DATE(2024,3,15),DATE(2029,12,31),0.08,0.06,2,0)')
			recalc(wb)
			const d = getResult(wb, 0, 0)
			const md = getResult(wb, 0, 1)
			expect(d?.kind).toBe('number')
			expect(md?.kind).toBe('number')
			if (d?.kind === 'number' && md?.kind === 'number') {
				expect(md.value).toBeCloseTo(d.value / 1.03, 4)
			}
		})

		test('ACCRINTM computes accrued interest at maturity', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ACCRINTM(DATE(2024,1,1),DATE(2024,7,1),0.1,1000,2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(50.56, 2)
		})

		test('RECEIVED computes amount at maturity', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'RECEIVED(DATE(2024,1,1),DATE(2024,7,1),970,0.06,2)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeGreaterThan(970)
		})

		test('TBILL functions compute known values', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'TBILLPRICE(DATE(2024,1,1),DATE(2024,6,29),0.05)')
			setFormula(wb, 0, 1, 'TBILLYIELD(DATE(2024,1,1),DATE(2024,6,29),A1)')
			setFormula(wb, 0, 2, 'TBILLEQ(DATE(2024,1,1),DATE(2024,6,29),0.05)')
			recalc(wb)
			const price = getResult(wb, 0, 0)
			const y = getResult(wb, 0, 1)
			const eq = getResult(wb, 0, 2)
			expect(price?.kind).toBe('number')
			expect(y?.kind).toBe('number')
			expect(eq?.kind).toBe('number')
			if (price?.kind === 'number') expect(price.value).toBeCloseTo(97.5, 3)
			if (y?.kind === 'number') expect(y.value).toBeCloseTo(0.05128, 4)
			if (eq?.kind === 'number') expect(eq.value).toBeCloseTo(0.05199, 4)
		})

		test('ACCRINT computes positive accrued interest before first coupon', () => {
			const wb = makeWorkbook()
			setFormula(
				wb,
				0,
				0,
				'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,3,15),0.08,1000,2,0,TRUE)',
			)
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeGreaterThan(0)
		})

		test('ODDFPRICE and ODDFYIELD are approximately inverse', () => {
			const wb = makeWorkbook()
			setFormula(
				wb,
				0,
				0,
				'ODDFPRICE(DATE(2024,3,15),DATE(2029,12,31),DATE(2023,10,1),DATE(2024,7,1),0.08,0.06,100,2,0)',
			)
			setFormula(
				wb,
				0,
				1,
				'ODDFYIELD(DATE(2024,3,15),DATE(2029,12,31),DATE(2023,10,1),DATE(2024,7,1),0.08,A1,100,2,0)',
			)
			recalc(wb)
			const y = getResult(wb, 0, 1)
			expect(y?.kind).toBe('number')
			if (y?.kind === 'number') expect(y.value).toBeCloseTo(0.06, 4)
		})

		test('ODDFPRICE matches Excel short odd-first coupon example', () => {
			const wb = makeWorkbook()
			setFormula(
				wb,
				0,
				0,
				'ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)',
			)
			recalc(wb)
			const result = getResult(wb, 0, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') expect(result.value).toBeCloseTo(113.597717474079, 12)
		})

		test('ODDLPRICE and ODDLYIELD are approximately inverse', () => {
			const wb = makeWorkbook()
			setFormula(
				wb,
				0,
				0,
				'ODDLPRICE(DATE(2024,10,1),DATE(2024,12,31),DATE(2024,7,1),0.08,0.06,100,2,0)',
			)
			setFormula(
				wb,
				0,
				1,
				'ODDLYIELD(DATE(2024,10,1),DATE(2024,12,31),DATE(2024,7,1),0.08,A1,100,2,0)',
			)
			recalc(wb)
			const y = getResult(wb, 0, 1)
			expect(y?.kind).toBe('number')
			if (y?.kind === 'number') expect(y.value).toBeCloseTo(0.06, 4)
		})

		test('AMORLINC computes first and later depreciation amounts', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'AMORLINC(1000,DATE(2024,1,1),DATE(2024,7,1),100,0,0.1,0)')
			setFormula(wb, 0, 1, 'AMORLINC(1000,DATE(2024,1,1),DATE(2024,7,1),100,1,0.1,0)')
			recalc(wb)
			const first = getResult(wb, 0, 0)
			const later = getResult(wb, 0, 1)
			expect(first?.kind).toBe('number')
			expect(later?.kind).toBe('number')
			if (first?.kind === 'number' && later?.kind === 'number') {
				expect(first.value).toBeGreaterThan(0)
				expect(later.value).toBeCloseTo(100, 0)
			}
		})

		test('AMORDEGRC computes declining depreciation and rejects basis 2', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'AMORDEGRC(1000,DATE(2024,1,1),DATE(2024,7,1),100,0,0.2,0)')
			setFormula(wb, 0, 1, 'AMORDEGRC(1000,DATE(2024,1,1),DATE(2024,7,1),100,0,0.2,2)')
			recalc(wb)
			const valid = getResult(wb, 0, 0)
			expect(valid?.kind).toBe('number')
			if (valid?.kind === 'number') expect(valid.value).toBeGreaterThan(0)
			expect(getResult(wb, 0, 1)).toEqual(errorValue('#NUM!'))
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

		test('ERF.PRECISE and ERFC.PRECISE match Excel cached tails', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'ERF.PRECISE(2)')
			setFormula(wb, 1, 0, 'ERFC.PRECISE(2)')
			recalc(wb)
			const erf = getResult(wb, 0, 0)
			const erfc = getResult(wb, 1, 0)
			expect(erf?.kind).toBe('number')
			expect(erfc?.kind).toBe('number')
			if (erf?.kind === 'number') {
				expect(erf.value).toBeCloseTo(0.995322265018953, 14)
			}
			if (erfc?.kind === 'number') {
				expect(erfc.value).toBeCloseTo(0.00467773498104726, 14)
			}
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

	describe('engineering functions - Bessel', () => {
		test('BESSELJ(1.5, 1) is approximately 0.5579', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELJ(1.5, 1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.5579365079, 5)
		})

		test('BESSELJ(0, 0) is 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELJ(0, 0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1, 10)
		})

		test('BESSELI(1.5, 1) is approximately 0.9817', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELI(1.5, 1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.9816664289, 5)
		})

		test('BESSELI(0, 0) is 1', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELI(0, 0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(1, 10)
		})

		test('BESSELY(2.5, 1) is approximately 0.1459', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELY(2.5, 1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.1459181379, 4)
		})

		test('BESSELY integer-order recurrence matches Excel cached oracle', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELY(2, 3)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(Math.abs(r.value - -1.12778376512206)).toBeLessThan(2e-8)
		})

		test('BESSELK(2.5, 1) is approximately 0.0739', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELK(2.5, 1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.0738908163, 4)
		})

		test('integer-order Bessel functions match Excel cached oracle for stress inputs', () => {
			const wb = makeWorkbook()
			const formulas = ['BESSELI(2,3)', 'BESSELJ(2,3)', 'BESSELK(2,3)', 'BESSELY(2,3)']
			const expected = [0.21273995970273565, 0.12894324997562717, 0.6473854, -1.1277837651220644]
			for (let row = 0; row < formulas.length; row++) {
				setFormula(wb, row, 0, formulas[row] as string)
			}
			recalc(wb)
			for (let row = 0; row < expected.length; row++) {
				const r = getResult(wb, row, 0)
				expect(r?.kind).toBe('number')
				if (r?.kind === 'number')
					expect(Math.abs(r.value - (expected[row] as number))).toBeLessThan(1e-15)
			}
		})

		test('BESSELY(0, 0) returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELY(0, 0)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('error')
			if (r?.kind === 'error') expect(r.value).toBe('#NUM!')
		})

		test('BESSELK(0, 1) returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELK(0, 1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('error')
			if (r?.kind === 'error') expect(r.value).toBe('#NUM!')
		})

		test('BESSELJ(1, -1) returns #NUM!', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELJ(1, -1)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('error')
			if (r?.kind === 'error') expect(r.value).toBe('#NUM!')
		})

		test('BESSELJ truncates non-integer order like Excel', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELJ(1, 1.5)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.4400506, 6)
		})

		test('BESSELJ remains stable for small x and larger order', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BESSELJ(0.1, 10)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(Math.abs(r.value)).toBeLessThan(1e-12)
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

		test('IMDIV uses Excel-compatible stable complex division', () => {
			const wb = makeWorkbook()
			setFormula(
				wb,
				0,
				0,
				'IMDIV("1.85021985907055+1.41787163074572j","0.556971676153418+0.426821890855467j")',
			)
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('3.32192809488737-8.28846662730513E-15j'))
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

		test('complex functions use Excel-style 15 significant digit strings', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'IMLN("2+4i")')
			setFormula(wb, 1, 0, 'IMEXP("2+4i")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('1.497866136777+1.10714871779409i'))
			expect(getResult(wb, 1, 0)).toEqual(stringValue('-4.82980938326939-5.59205609364098i'))
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
		test('TAN matches Excel cached precision for POI stress cases', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 5)
			setFormula(wb, 0, 1, 'TAN(A1)')
			setFormula(wb, 1, 1, 'TAN(A2)')
			setFormula(wb, 2, 1, 'TAN(B1)')
			recalc(wb)
			expect(getResult(wb, 0, 1)).toEqual(numberValue(1.5574077246549023))
			expect(getResult(wb, 1, 1)).toEqual(numberValue(-3.380515006246586))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(74.68593339876537))
		})

		test('TANH matches Excel cached precision for nested POI cases', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 2)
			setFormula(wb, 1, 0, 'TANH(A1)')
			setFormula(wb, 2, 0, 'TANH(A2)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(numberValue(0.964027580075817))
			expect(getResult(wb, 2, 0)).toEqual(numberValue(0.7460679984455995))
		})

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

		test('SUMX2PY2 pairs flattened vectors and ignores nonnumeric pairs', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 0.00001)
			setNum(wb, 1, 0, 1.1)
			setNum(wb, 2, 0, 1.00001)
			setNum(wb, 0, 1, -1)
			setNum(wb, 0, 2, -1.00001)
			setFormula(wb, 3, 0, 'SUMX2PY2(A1:A3,B1:D1)')
			recalc(wb)
			expect(getResult(wb, 3, 0)).toEqual(numberValue(3.2100200002))
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

		test('SUMX pair functions return #DIV/0! when no numeric pairs remain', () => {
			const wb = makeWorkbook()
			setStr(wb, 0, 0, 'x')
			setStr(wb, 0, 1, 'y')
			setFormula(wb, 1, 0, 'SUMX2MY2(A1,B1)')
			recalc(wb)
			expect(getResult(wb, 1, 0)).toEqual(errorValue('#DIV/0!'))
		})

		test('SUMX pair functions return #N/A for different flattened lengths', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 1, 3)
			setFormula(wb, 2, 0, 'SUMXMY2(A1:A2,B1)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(errorValue('#N/A'))
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

		test('matrix functions reject blank matrix entries', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 3, 'MDETERM(A1:B2)')
			setFormula(wb, 1, 3, 'MINVERSE(A1:B2)')
			setFormula(wb, 2, 3, 'MMULT(A1:B2,A1:B2)')
			recalc(wb)
			expect(getResult(wb, 0, 3)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 1, 3)).toEqual(errorValue('#VALUE!'))
			expect(getResult(wb, 2, 3)).toEqual(errorValue('#VALUE!'))
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

		test('TRANSPOSE materializes blank source cells as zero', () => {
			const wb = makeWorkbook()
			setFormula(wb, 2, 0, 'TRANSPOSE(A1:B2)')
			recalc(wb)
			expect(getResult(wb, 2, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 2, 1)).toEqual(numberValue(0))
			expect(getResult(wb, 3, 0)).toEqual(numberValue(0))
			expect(getResult(wb, 3, 1)).toEqual(numberValue(0))
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
			setFormula(wb, 1, 0, 'GAMMALN.PRECISE(5)')
			recalc(wb)
			for (let row = 0; row < 2; row++) {
				const r = getResult(wb, row, 0)
				expect(r?.kind).toBe('number')
				if (r?.kind === 'number') expect(r.value.toPrecision(15)).toBe('3.17805383034795')
			}
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

		test('F inverse functions preserve Excel cached distribution round-trips', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'F.INV(F.DIST(1,2,3,TRUE),2,3)')
			setFormula(wb, 1, 0, 'F.INV.RT(F.DIST.RT(1,2,3),2,3)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(1))
			expect(getResult(wb, 1, 0)).toEqual(numberValue(1))
		})

		test('F distribution functions truncate degrees of freedom', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'F.DIST(2,6.6,8,TRUE)')
			setFormula(wb, 1, 0, 'F.DIST(2,6.6,8,FALSE)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0.8208))
			const density = getResult(wb, 1, 0)
			expect(density?.kind).toBe('number')
			if (density?.kind === 'number') {
				expect(density.value).toBeCloseTo(0.165888, 14)
			}
		})

		test('BETA.INV resolves roots close to the lower bound', () => {
			const wb = makeWorkbook()
			setFormula(
				wb,
				0,
				0,
				'BETA.INV(0.13333333333333333,0.13333333333333333,1.3333333333333333,1,6)',
			)
			recalc(wb)
			const result = getResult(wb, 0, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBeCloseTo(1.00000090592338, 13)
			}
		})

		test('BETA.INV defaults omitted bounds to zero and one', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BETA.INV(0.6875,2,3)')
			recalc(wb)
			const result = getResult(wb, 0, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBeCloseTo(0.5, 13)
			}
		})

		test('right-tail distributions preserve tiny nonzero probabilities', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'F.DIST.RT(1E6,5,10)')
			setFormula(wb, 1, 0, 'CHISQ.DIST.RT(100,5)')
			setFormula(wb, 2, 0, 'CHISQ.DIST.RT(1000,5)')
			recalc(wb)
			const fTail = getResult(wb, 0, 0)
			const chiTail = getResult(wb, 1, 0)
			const farChiTail = getResult(wb, 2, 0)
			expect(fTail?.kind).toBe('number')
			expect(chiTail?.kind).toBe('number')
			expect(farChiTail?.kind).toBe('number')
			if (fTail?.kind === 'number') expect(fTail.value.toPrecision(15)).toBe('3.75370307872001e-28')
			if (chiTail?.kind === 'number')
				expect(chiTail.value.toPrecision(15)).toBe('5.28514836094322e-20')
			if (farChiTail?.kind === 'number')
				expect(farChiTail.value.toPrecision(15)).toBe('6.01007768792085e-214')
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

		test('finite discrete distributions match Excel cached precision', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BINOM.DIST(5,10,0.3,TRUE)')
			setFormula(wb, 1, 0, 'HYPGEOMDIST(2,3,6,8)')
			setFormula(wb, 2, 0, 'HYPGEOM.DIST(1,2,7,9,TRUE)')
			recalc(wb)
			const binom = getResult(wb, 0, 0)
			const hypgeom = getResult(wb, 1, 0)
			const cumulativeHypgeom = getResult(wb, 2, 0)
			expect(binom?.kind).toBe('number')
			expect(hypgeom?.kind).toBe('number')
			expect(cumulativeHypgeom?.kind).toBe('number')
			if (binom?.kind === 'number') expect(binom.value.toPrecision(15)).toBe('0.952651012600000')
			if (hypgeom?.kind === 'number')
				expect(hypgeom.value.toPrecision(15)).toBe('0.535714285714286')
			if (cumulativeHypgeom?.kind === 'number')
				expect(cumulativeHypgeom.value.toPrecision(15)).toBe('0.416666666666666')
		})

		test('POISSON.DIST(3, 5, FALSE) PMF', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'POISSON.DIST(3, 5, FALSE)')
			recalc(wb)
			const r = getResult(wb, 0, 0)
			expect(r?.kind).toBe('number')
			if (r?.kind === 'number') expect(r.value).toBeCloseTo(0.14037, 4)
		})

		test('POISSON.DIST small PMF matches Excel cached precision', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'POISSON.DIST(2, 5, FALSE)')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(numberValue(0.08422433748856833))
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

		test('BETADIST legacy alias is cumulative without a cumulative flag', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'BETADIST(0.5, 2, 2)')
			setFormula(wb, 0, 1, 'BETADIST(0.5, 2, 2, 0, 1)')
			recalc(wb)
			const first = getResult(wb, 0, 0)
			const second = getResult(wb, 0, 1)
			expect(first?.kind).toBe('number')
			expect(second?.kind).toBe('number')
			if (first?.kind === 'number') expect(first.value).toBeCloseTo(0.5, 12)
			if (second?.kind === 'number') expect(second.value).toBeCloseTo(0.5, 12)
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

		test('PROB sums probabilities over an inclusive numeric interval', () => {
			const wb = makeWorkbook()
			for (const [index, value] of [1, 2, 3, 4].entries()) {
				setNum(wb, index, 0, value)
			}
			for (const [index, value] of [0.1, 0.2, 0.3, 0.4].entries()) {
				setNum(wb, index, 1, value)
			}
			setFormula(wb, 0, 2, 'PROB(A1:A4,B1:B4,2,3)')
			recalc(wb)
			expect(getResult(wb, 0, 2)).toEqual(numberValue(0.5))
		})

		test('PROB validates range shape and total probability', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 1, 0, 2)
			setNum(wb, 0, 1, 0.2)
			setNum(wb, 1, 1, 0.2)
			setFormula(wb, 0, 2, 'PROB(A1:A2,B1:B2,1)')
			setFormula(wb, 1, 2, 'PROB(A1:A2,B1:C2,1)')
			recalc(wb)
			expect(getResult(wb, 0, 2)).toEqual(errorValue('#NUM!'))
			expect(getResult(wb, 1, 2)).toEqual(errorValue('#N/A'))
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

		test('legacy statistical test aliases match dotted function names', () => {
			const wb = makeWorkbook()
			const left = [10, 12, 9, 11, 13, 8]
			const right = [9, 11, 10, 10, 12, 7]
			for (let i = 0; i < left.length; i++) {
				setNum(wb, i, 0, left[i] as number)
				setNum(wb, i, 1, right[i] as number)
			}
			setFormula(wb, 7, 0, 'T.TEST(A1:A6,B1:B6,2,1)')
			setFormula(wb, 7, 1, 'TTEST(A1:A6,B1:B6,2,1)')
			setFormula(wb, 8, 0, 'F.TEST(A1:A6,B1:B6)')
			setFormula(wb, 8, 1, 'FTEST(A1:A6,B1:B6)')
			setFormula(wb, 9, 0, 'CHISQ.TEST(A1:B2,A3:B4)')
			setFormula(wb, 9, 1, 'CHITEST(A1:B2,A3:B4)')
			setFormula(wb, 10, 0, 'Z.TEST(A1:A6,10,2)')
			setFormula(wb, 10, 1, 'ZTEST(A1:A6,10,2)')
			recalc(wb)
			expect(getResult(wb, 7, 1)).toEqual(getResult(wb, 7, 0))
			expect(getResult(wb, 8, 1)).toEqual(getResult(wb, 8, 0))
			expect(getResult(wb, 9, 1)).toEqual(getResult(wb, 9, 0))
			expect(getResult(wb, 10, 1)).toEqual(getResult(wb, 10, 0))
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

		test('INFO returns deterministic workbook-environment metadata', () => {
			const wb = makeWorkbook()
			setFormula(wb, 0, 0, 'INFO("release")')
			setFormula(wb, 0, 1, 'INFO("system")')
			setFormula(wb, 0, 2, 'INFO("missing")')
			recalc(wb)
			expect(getResult(wb, 0, 0)).toEqual(stringValue('14.3'))
			expect(getResult(wb, 0, 1)).toEqual(stringValue('mac'))
			expect(getResult(wb, 0, 2)).toEqual(errorValue('#VALUE!'))
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

	describe('FORECAST.ETS family', () => {
		test('FORECAST.ETS with linear trend close to FORECAST.LINEAR', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 10; i++) {
				setNum(wb, i, 0, i + 1)
				setNum(wb, i, 1, 2 * (i + 1))
			}
			setFormula(wb, 10, 0, 'FORECAST.ETS(11,B1:B10,A1:A10)')
			setFormula(wb, 10, 1, 'FORECAST.LINEAR(11,B1:B10,A1:A10)')
			recalc(wb)
			const ets = getResult(wb, 10, 0)
			const linear = getResult(wb, 10, 1)
			expect(ets?.kind).toBe('number')
			expect(linear?.kind).toBe('number')
			if (ets?.kind === 'number' && linear?.kind === 'number') {
				expect(ets.value).toBeCloseTo(linear.value, 0)
			}
		})

		test('FORECAST.ETS.SEASONALITY with non-seasonal data returns 1', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 10; i++) {
				setNum(wb, i, 0, 2 * (i + 1) + 1)
				setNum(wb, i, 1, i + 1)
			}
			setFormula(wb, 10, 0, 'FORECAST.ETS.SEASONALITY(A1:A10,B1:B10)')
			recalc(wb)
			const result = getResult(wb, 10, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBe(1)
			}
		})

		test('FORECAST.ETS with numeric timeline', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 8; i++) {
				setNum(wb, i, 0, (i + 1) * 10)
				setNum(wb, i, 1, 100 + 5 * (i + 1))
			}
			setFormula(wb, 8, 0, 'FORECAST.ETS(90,B1:B8,A1:A8)')
			recalc(wb)
			const result = getResult(wb, 8, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBeCloseTo(145, 0)
			}
		})

		test('FORECAST.ETS.STAT returns step size', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 6; i++) {
				setNum(wb, i, 0, 10 + 3 * i)
				setNum(wb, i, 1, (i + 1) * 5)
			}
			setFormula(wb, 6, 0, 'FORECAST.ETS.STAT(A1:A6,B1:B6,8)')
			recalc(wb)
			const result = getResult(wb, 6, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBe(5)
			}
		})

		test('FORECAST.ETS.CONFINT returns positive width', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 10; i++) {
				setNum(wb, i, 0, i + 1)
				setNum(wb, i, 1, 2 * (i + 1) + (i % 2 === 0 ? 0.5 : -0.5))
			}
			setFormula(wb, 10, 0, 'FORECAST.ETS.CONFINT(11,B1:B10,A1:A10,0.95)')
			recalc(wb)
			const result = getResult(wb, 10, 0)
			expect(result?.kind).toBe('number')
			if (result?.kind === 'number') {
				expect(result.value).toBeGreaterThan(0)
			}
		})

		test('FORECAST.ETS rejects target inside historical timeline', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 10; i++) {
				setNum(wb, i, 0, i + 1)
				setNum(wb, i, 1, 2 * (i + 1))
			}
			setFormula(wb, 10, 0, 'FORECAST.ETS(5,B1:B10,A1:A10)')
			recalc(wb)
			const result = getResult(wb, 10, 0)
			expect(result?.kind).toBe('error')
			if (result?.kind === 'error') expect(result.value).toBe('#NUM!')
		})

		test('FORECAST.ETS rejects target that does not continue the step size', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 8; i++) {
				setNum(wb, i, 0, (i + 1) * 10)
				setNum(wb, i, 1, 100 + 5 * (i + 1))
			}
			setFormula(wb, 8, 0, 'FORECAST.ETS(85,B1:B8,A1:A8)')
			recalc(wb)
			const result = getResult(wb, 8, 0)
			expect(result?.kind).toBe('error')
			if (result?.kind === 'error') expect(result.value).toBe('#NUM!')
		})

		test('FORECAST.ETS rejects fractional seasonality values', () => {
			const wb = makeWorkbook()
			for (let i = 0; i < 8; i++) {
				setNum(wb, i, 0, i + 1)
				setNum(wb, i, 1, 100 + i)
			}
			setFormula(wb, 8, 0, 'FORECAST.ETS(9,B1:B8,A1:A8,1.5)')
			recalc(wb)
			const result = getResult(wb, 8, 0)
			expect(result?.kind).toBe('error')
			if (result?.kind === 'error') expect(result.value).toBe('#NUM!')
		})

		test('FORECAST.ETS with too few data points returns error', () => {
			const wb = makeWorkbook()
			setNum(wb, 0, 0, 1)
			setNum(wb, 0, 1, 10)
			setNum(wb, 1, 0, 2)
			setNum(wb, 1, 1, 20)
			setFormula(wb, 2, 0, 'FORECAST.ETS(3,B1:B2,A1:A2)')
			recalc(wb)
			const result = getResult(wb, 2, 0)
			expect(result?.kind).toBe('error')
		})
	})
})
