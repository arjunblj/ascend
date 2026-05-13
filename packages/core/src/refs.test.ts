import { describe, expect, test } from 'bun:test'
import {
	columnToIndex,
	expandRange,
	indexToColumn,
	normalizeRange,
	parseA1,
	parseA1Safe,
	parseRange,
	parseSqref,
	rangeIntersection,
	rangeIntersects,
	sqrefIntersects,
	toA1,
	toRangeString,
} from './refs.ts'

describe('columnToIndex', () => {
	test('single letters', () => {
		expect(columnToIndex('A')).toBe(0)
		expect(columnToIndex('B')).toBe(1)
		expect(columnToIndex('Z')).toBe(25)
	})

	test('double letters', () => {
		expect(columnToIndex('AA')).toBe(26)
		expect(columnToIndex('AZ')).toBe(51)
		expect(columnToIndex('BA')).toBe(52)
	})

	test('XFD is last Excel column (16383)', () => {
		expect(columnToIndex('XFD')).toBe(16383)
	})
})

describe('indexToColumn', () => {
	test('single letters', () => {
		expect(indexToColumn(0)).toBe('A')
		expect(indexToColumn(1)).toBe('B')
		expect(indexToColumn(25)).toBe('Z')
	})

	test('double letters', () => {
		expect(indexToColumn(26)).toBe('AA')
		expect(indexToColumn(51)).toBe('AZ')
		expect(indexToColumn(52)).toBe('BA')
	})

	test('XFD roundtrip', () => {
		expect(indexToColumn(16383)).toBe('XFD')
	})

	test('roundtrip all single-letter columns', () => {
		for (let i = 0; i < 26; i++) {
			expect(columnToIndex(indexToColumn(i))).toBe(i)
		}
	})
})

describe('parseA1 / toA1', () => {
	test('A1 is origin', () => {
		expect(parseA1('A1')).toEqual({ row: 0, col: 0 })
	})

	test('B3 maps to row 2, col 1', () => {
		expect(parseA1('B3')).toEqual({ row: 2, col: 1 })
	})

	test('max Excel cell XFD1048576', () => {
		expect(parseA1('XFD1048576')).toEqual({ row: 1048575, col: 16383 })
	})

	test('case insensitive', () => {
		expect(parseA1('b3')).toEqual({ row: 2, col: 1 })
	})

	test('roundtrip', () => {
		const cases = ['A1', 'B3', 'Z100', 'AA1', 'XFD1048576']
		for (const ref of cases) {
			expect(toA1(parseA1(ref))).toBe(ref)
		}
	})

	test('throws on invalid ref', () => {
		expect(() => parseA1('')).toThrow()
		expect(() => parseA1('123')).toThrow()
		expect(() => parseA1('A')).toThrow()
	})

	test('parseA1Safe returns null on invalid ref', () => {
		expect(parseA1Safe(undefined)).toBeNull()
		expect(parseA1Safe('')).toBeNull()
		expect(parseA1Safe('123')).toBeNull()
		expect(parseA1Safe('A1')).toEqual({ row: 0, col: 0 })
	})
})

describe('parseRange / toRangeString', () => {
	test('basic range', () => {
		const range = parseRange('A1:C10')
		expect(range.start).toEqual({ row: 0, col: 0 })
		expect(range.end).toEqual({ row: 9, col: 2 })
		expect(range.sheet).toBeUndefined()
	})

	test('range with sheet prefix', () => {
		const range = parseRange('Sheet1!A1:C10')
		expect(range.sheet).toBe('Sheet1')
		expect(range.start).toEqual({ row: 0, col: 0 })
		expect(range.end).toEqual({ row: 9, col: 2 })
	})

	test('quoted sheet name', () => {
		const range = parseRange("'My Sheet'!A1:B2")
		expect(range.sheet).toBe('My Sheet')
	})

	test('single cell range', () => {
		const range = parseRange('D5')
		expect(range.start).toEqual({ row: 4, col: 3 })
		expect(range.end).toEqual({ row: 4, col: 3 })
	})

	test('roundtrip', () => {
		expect(toRangeString(parseRange('A1:C10'))).toBe('A1:C10')
		expect(toRangeString(parseRange('Sheet1!A1:C10'))).toBe('Sheet1!A1:C10')
	})
})

describe('expandRange', () => {
	test('single cell', () => {
		const cells = expandRange(parseRange('B2'))
		expect(cells).toEqual([{ row: 1, col: 1 }])
	})

	test('2x3 range', () => {
		const cells = expandRange(parseRange('A1:C2'))
		expect(cells).toEqual([
			{ row: 0, col: 0 },
			{ row: 0, col: 1 },
			{ row: 0, col: 2 },
			{ row: 1, col: 0 },
			{ row: 1, col: 1 },
			{ row: 1, col: 2 },
		])
	})

	test('cell count matches dimensions', () => {
		const cells = expandRange(parseRange('A1:D5'))
		expect(cells).toHaveLength(4 * 5)
	})
})

describe('range algebra', () => {
	test('normalizes inverted ranges', () => {
		expect(normalizeRange({ start: { row: 9, col: 3 }, end: { row: 1, col: 0 } })).toEqual({
			start: { row: 1, col: 0 },
			end: { row: 9, col: 3 },
		})
	})

	test('detects range intersections and respects explicit sheet names', () => {
		expect(rangeIntersects(parseRange('A1:C3'), parseRange('C3:D4'))).toBe(true)
		expect(rangeIntersects(parseRange('A1:B2'), parseRange('C3:D4'))).toBe(false)
		expect(rangeIntersects(parseRange('Sheet1!A1:B2'), parseRange('Sheet2!A1:B2'))).toBe(false)
	})

	test('returns intersection ranges', () => {
		expect(rangeIntersection(parseRange('A1:C3'), parseRange('B2:D4'))).toEqual({
			start: { row: 1, col: 1 },
			end: { row: 2, col: 2 },
		})
		expect(rangeIntersection(parseRange('A1:A1'), parseRange('B1:B1'))).toBeNull()
	})

	test('parses and intersects sqref ranges', () => {
		expect(parseSqref('A1:A2 C1:D1')).toEqual([parseRange('A1:A2'), parseRange('C1:D1')])
		expect(sqrefIntersects('A1:A2 C1:D1', parseRange('B1:B2'))).toBe(false)
		expect(sqrefIntersects('A1:A2 C1:D1', parseRange('B1:C2'))).toBe(true)
	})
})
