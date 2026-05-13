import type { RangeRef } from './refs.ts'
import { normalizeRange, parseSqref, rangeIntersection, rangeIntersects } from './refs.ts'

export interface RangeIndexEntry<T> {
	readonly range: RangeRef
	readonly value: T
	readonly sourceIndex: number
}

export class RangeIndex<T> {
	private readonly entries: readonly RangeIndexEntry<T>[]

	constructor(entries: readonly RangeIndexEntry<T>[]) {
		this.entries = [...entries]
			.map((entry) => ({ ...entry, range: normalizeRange(entry.range) }))
			.sort(
				(a, b) => a.range.start.row - b.range.start.row || a.range.start.col - b.range.start.col,
			)
	}

	static fromRanges<T>(
		values: readonly T[],
		rangeForValue: (value: T, index: number) => RangeRef,
	): RangeIndex<T> {
		return new RangeIndex(
			values.map((value, index) => ({
				range: rangeForValue(value, index),
				value,
				sourceIndex: index,
			})),
		)
	}

	static fromSqrefs<T>(
		values: readonly T[],
		sqrefForValue: (value: T, index: number) => string,
	): RangeIndex<T> {
		const entries: RangeIndexEntry<T>[] = []
		values.forEach((value, sourceIndex) => {
			for (const range of parseSqref(sqrefForValue(value, sourceIndex))) {
				entries.push({ range, value, sourceIndex })
			}
		})
		return new RangeIndex(entries)
	}

	intersectingEntries(range: RangeRef): readonly RangeIndexEntry<T>[] {
		const query = normalizeRange(range)
		const matches: RangeIndexEntry<T>[] = []
		for (const entry of this.entries) {
			if (entry.range.start.row > query.end.row) break
			if (rangeIntersects(entry.range, query)) matches.push(entry)
		}
		return matches.sort(
			(a, b) =>
				a.sourceIndex - b.sourceIndex ||
				a.range.start.row - b.range.start.row ||
				a.range.start.col - b.range.start.col,
		)
	}

	intersectingValues(range: RangeRef): readonly T[] {
		const values: T[] = []
		const seen = new Set<number>()
		for (const entry of this.intersectingEntries(range)) {
			if (seen.has(entry.sourceIndex)) continue
			seen.add(entry.sourceIndex)
			values.push(entry.value)
		}
		return values
	}

	containsCell(row: number, col: number): boolean {
		for (const entry of this.entries) {
			if (entry.range.start.row > row) break
			if (
				row >= entry.range.start.row &&
				row <= entry.range.end.row &&
				col >= entry.range.start.col &&
				col <= entry.range.end.col
			) {
				return true
			}
		}
		return false
	}
}

export function rangeMaskOffsets(
	ranges: readonly RangeRef[],
	viewport: RangeRef,
): ReadonlySet<number> {
	const mask = new Set<number>()
	const normalizedViewport = normalizeRange(viewport)
	const width = normalizedViewport.end.col - normalizedViewport.start.col + 1
	for (const range of ranges) {
		const intersection = rangeIntersection(range, normalizedViewport)
		if (!intersection) continue
		for (let row = intersection.start.row; row <= intersection.end.row; row++) {
			const rowOffset = (row - normalizedViewport.start.row) * width
			for (let col = intersection.start.col; col <= intersection.end.col; col++) {
				mask.add(rowOffset + col - normalizedViewport.start.col)
			}
		}
	}
	return mask
}
