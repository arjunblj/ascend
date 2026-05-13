import { describe, expect, test } from 'bun:test'
import { RangeIndex, rangeMaskOffsets } from './range-index.ts'
import { parseRange, parseSqref, type RangeRef, rangeIntersects } from './refs.ts'

function exhaustiveSqrefValues<T>(
	values: readonly T[],
	query: RangeRef,
	sqrefForValue: (value: T, index: number) => string,
): readonly T[] {
	return values.filter((value, index) =>
		parseSqref(sqrefForValue(value, index)).some((range) => rangeIntersects(range, query)),
	)
}

describe('RangeIndex', () => {
	test('returns intersecting range values once for sqref-backed metadata', () => {
		const values = [
			{ id: 'validation-a', sqref: 'A1:A3 C1:C3' },
			{ id: 'validation-b', sqref: 'E1:E3' },
		]
		const index = RangeIndex.fromSqrefs(values, (value) => value.sqref)

		expect(index.intersectingValues(parseRange('B1:C2')).map((value) => value.id)).toEqual([
			'validation-a',
		])
		expect(index.intersectingValues(parseRange('D1:D3'))).toEqual([])
	})

	test('matches exhaustive sqref scans for absolute whole-axis and sheet-qualified ranges', () => {
		const values = [
			{ id: 'absolute', sqref: "'Data Sheet'!$A$1:$B$3 'Data Sheet'!$D$4:$D$4" },
			{ id: 'whole-row', sqref: '2:2' },
			{ id: 'whole-column', sqref: '$E:$F' },
			{ id: 'other-sheet', sqref: 'Other!B2:C3' },
		]
		const query = parseRange("'Data Sheet'!B2:E4")
		const index = RangeIndex.fromSqrefs(values, (value) => value.sqref)
		const expected = exhaustiveSqrefValues(values, query, (value) => value.sqref)

		expect(index.intersectingValues(query).map((value) => value.id)).toEqual(
			expected.map((value) => value.id),
		)
	})

	test('deduplicates multi-range metadata by source entry not primitive value identity', () => {
		const index = RangeIndex.fromRanges(['same', 'same'], (_value, index) =>
			index === 0 ? parseRange('A1:A1') : parseRange('B1:B1'),
		)

		expect(index.intersectingValues(parseRange('A1:B1'))).toEqual(['same', 'same'])
	})

	test('checks single-cell membership with row-ordered early exits', () => {
		const index = RangeIndex.fromRanges(
			[parseRange('A1:B2'), parseRange('D10:D20')],
			(range) => range,
		)

		expect(index.containsCell(0, 0)).toBe(true)
		expect(index.containsCell(2, 0)).toBe(false)
		expect(index.containsCell(15, 3)).toBe(true)
	})

	test('builds viewport-relative coverage masks', () => {
		const mask = rangeMaskOffsets([parseRange('B2:C3'), parseRange('D4:D4')], parseRange('A1:D4'))

		expect([...mask].sort((a, b) => a - b)).toEqual([5, 6, 9, 10, 15])
	})

	test('builds masks for absolute whole-row whole-column and sheet-qualified ranges', () => {
		const mask = rangeMaskOffsets(
			[parseRange('$B:$B'), parseRange('3:3'), parseRange("'Sheet 1'!$D$4:$D$4")],
			parseRange("'Sheet 1'!B2:D4"),
		)

		expect([...mask].sort((a, b) => a - b)).toEqual([0, 3, 4, 5, 6, 8])
	})
})
