import type { RangeRef } from '@ascend/core'
import { parseA1, parseRange, toA1 } from '@ascend/core'

export function shiftSqref(
	sqref: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): string | null {
	const shifted = sqref
		.split(/\s+/)
		.map((part) => shiftA1RangeOrCell(part, axis, at, delta))
		.filter((part): part is string => part !== null && part.length > 0)
	return shifted.length > 0 ? shifted.join(' ') : null
}

export function expandSqrefRows(sqref: string, count: number): string {
	return sqref
		.split(/\s+/)
		.map((part) => {
			try {
				const range = parseRange(part)
				const start = toA1(range.start)
				const end = toA1({ row: range.end.row + count, col: range.end.col })
				return start === end ? start : `${start}:${end}`
			} catch {
				return part
			}
		})
		.join(' ')
}

export function shiftA1RangeOrCell(
	input: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): string | null {
	if (!input) return null
	try {
		const range = parseRange(input)
		const shifted = shiftRangeRef(range, axis, at, delta)
		if (!shifted) return null
		const start = toA1(shifted.start)
		const end = toA1(shifted.end)
		return start === end ? start : `${start}:${end}`
	} catch {
		return shiftA1Ref(input, axis, at, delta)
	}
}

export function shiftRangeRef(
	range: RangeRef,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): RangeRef | null {
	const startIndex = axis === 'row' ? range.start.row : range.start.col
	const endIndex = axis === 'row' ? range.end.row : range.end.col
	const shifted = shiftRangeBounds(startIndex, endIndex, at, delta)
	if (!shifted) return null
	if (axis === 'row') {
		return {
			start: { ...range.start, row: shifted.start },
			end: { ...range.end, row: shifted.end },
		}
	}
	return {
		start: { ...range.start, col: shifted.start },
		end: { ...range.end, col: shifted.end },
	}
}

export function shiftRangeBounds(
	start: number,
	end: number,
	at: number,
	delta: number,
): { start: number; end: number } | null {
	if (delta > 0) {
		return {
			start: start >= at ? start + delta : start,
			end: end >= at ? end + delta : end,
		}
	}
	const count = Math.abs(delta)
	const deleteEnd = at + count
	if (end < at) return { start, end }
	if (start >= deleteEnd) {
		return { start: start + delta, end: end + delta }
	}
	if (start >= at && end < deleteEnd) return null
	const nextStart = start >= at ? at : start
	const nextEnd = end >= deleteEnd ? end + delta : at - 1
	return nextEnd >= nextStart ? { start: nextStart, end: nextEnd } : null
}

export function shiftA1Ref(
	ref: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): string | null {
	try {
		const parsed = parseA1(ref)
		const next =
			axis === 'row' ? shiftIndex(parsed.row, at, delta) : shiftIndex(parsed.col, at, delta)
		if (next === null) return null
		return axis === 'row'
			? toA1({ row: next, col: parsed.col })
			: toA1({ row: parsed.row, col: next })
	} catch {
		return ref
	}
}

export function shiftIndex(index: number, at: number, delta: number): number | null {
	if (delta > 0) return index >= at ? index + delta : index
	const count = Math.abs(delta)
	const deleteEnd = at + count
	if (index >= at && index < deleteEnd) return null
	return index >= deleteEnd ? index + delta : index
}
