import { describe, expect, test } from 'bun:test'
import { RangeIndex, rangeMaskOffsets } from './range-index.ts'
import {
	normalizeRange,
	parseRange,
	parseSqref,
	type RangeRef,
	rangeIntersects,
	toA1,
	toRangeString,
} from './refs.ts'

function exhaustiveSqrefValues<T>(
	values: readonly T[],
	query: RangeRef,
	sqrefForValue: (value: T, index: number) => string,
): readonly T[] {
	return values.filter((value, index) =>
		parseSqref(sqrefForValue(value, index)).some((range) => rangeIntersects(range, query)),
	)
}

function exhaustiveSqrefEntries<T>(
	values: readonly T[],
	query: RangeRef,
	sqrefForValue: (value: T, index: number) => string,
): readonly { range: RangeRef; value: T; sourceIndex: number }[] {
	return values
		.flatMap((value, sourceIndex) =>
			parseSqref(sqrefForValue(value, sourceIndex)).map((range) => ({
				range: normalizeRange(range),
				value,
				sourceIndex,
			})),
		)
		.filter((entry) => rangeIntersects(entry.range, query))
		.sort(
			(a, b) =>
				a.sourceIndex - b.sourceIndex ||
				a.range.start.row - b.range.start.row ||
				a.range.start.col - b.range.start.col,
		)
}

function exhaustiveMaskOffsets(
	ranges: readonly RangeRef[],
	viewport: RangeRef,
): ReadonlySet<number> {
	const normalizedViewport = normalizeRange(viewport)
	const width = normalizedViewport.end.col - normalizedViewport.start.col + 1
	const mask = new Set<number>()
	for (let row = normalizedViewport.start.row; row <= normalizedViewport.end.row; row++) {
		for (let col = normalizedViewport.start.col; col <= normalizedViewport.end.col; col++) {
			const cell = { start: { row, col }, end: { row, col }, sheet: normalizedViewport.sheet }
			if (ranges.some((range) => rangeIntersects(range, cell))) {
				mask.add((row - normalizedViewport.start.row) * width + col - normalizedViewport.start.col)
			}
		}
	}
	return mask
}

function entryKeys<T extends { readonly id: string }>(
	entries: readonly { range: RangeRef; value: T; sourceIndex: number }[],
): readonly string[] {
	return entries.map(
		(entry) =>
			`${entry.sourceIndex}:${entry.value.id}:${toRangeString(normalizeRange(entry.range))}`,
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

	test('matches exhaustive entry scans for adversarial metadata sqrefs', () => {
		const values = [
			{
				id: 'escaped-sheet',
				sqref: "'Bob''s.Data/Δ'!$A$1:$C$3 'Bob''s.Data/Δ'!$XFD$1048576:$XFD$1048576",
			},
			{ id: 'overlap-union', sqref: 'B2:D4 C3:E5 B2:D4' },
			{ id: 'whole-row', sqref: '4:4' },
			{ id: 'whole-column', sqref: '$C:$C' },
			{ id: 'inverted', sqref: 'F10:D8' },
			{ id: 'other-sheet', sqref: "'Other Sheet'!A1:Z99" },
		]
		const index = RangeIndex.fromSqrefs(values, (value) => value.sqref)
		const queries = [
			parseRange("'Bob''s.Data/Δ'!A1:E5"),
			parseRange("'Bob''s.Data/Δ'!XFD1048576:XFD1048576"),
			parseRange("'Other Sheet'!B2:C3"),
			parseRange('C4:C4'),
			parseRange('$A$1048576:$XFD$1048576'),
			parseRange('D8:F10'),
		]

		for (const query of queries) {
			expect(entryKeys(index.intersectingEntries(query))).toEqual(
				entryKeys(exhaustiveSqrefEntries(values, query, (value) => value.sqref)),
			)
			expect(index.intersectingValues(query).map((value) => value.id)).toEqual(
				exhaustiveSqrefValues(values, query, (value) => value.sqref).map((value) => value.id),
			)
		}
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

	test('viewport masks match exhaustive cell scans for unions overlaps and boundary refs', () => {
		const viewport = parseRange("'Sheet 1'!B2:E5")
		const ranges = [
			parseRange("'Sheet 1'!$B$2:$C$3"),
			parseRange('3:4'),
			parseRange('$E:$E'),
			parseRange("'Other Sheet'!B2:E5"),
			parseRange('E5:B5'),
		]

		expect([...rangeMaskOffsets(ranges, viewport)].sort((a, b) => a - b)).toEqual(
			[...exhaustiveMaskOffsets(ranges, viewport)].sort((a, b) => a - b),
		)

		const boundaryViewport = parseRange('XFC1048575:XFD1048576')
		const boundaryRanges = [parseRange('XFD:XFD'), parseRange('1048576:1048576')]
		const boundaryRefs = [...rangeMaskOffsets(boundaryRanges, boundaryViewport)]
			.sort((a, b) => a - b)
			.map((offset) => {
				const width =
					normalizeRange(boundaryViewport).end.col - normalizeRange(boundaryViewport).start.col + 1
				const row = normalizeRange(boundaryViewport).start.row + Math.floor(offset / width)
				const col = normalizeRange(boundaryViewport).start.col + (offset % width)
				return toA1({ row, col })
			})
		expect(boundaryRefs).toEqual(['XFD1048575', 'XFC1048576', 'XFD1048576'])
	})
})
