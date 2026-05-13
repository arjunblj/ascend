import type { RangeRef, Sheet, Table, TableId } from '@ascend/core'
import { shiftRangeRef } from './structural/ref-shift.ts'

export interface TableRangeOverlap {
	readonly left: Table
	readonly leftRef: RangeRef
	readonly right: Table
	readonly rightRef: RangeRef
}

export interface QueryTableColumnShiftBlocker {
	readonly table: Table
	readonly currentRef: RangeRef
	readonly shiftedRef: RangeRef
}

export function tableRangesOverlap(left: RangeRef, right: RangeRef): boolean {
	return (
		Math.min(left.start.row, left.end.row) <= Math.max(right.start.row, right.end.row) &&
		Math.min(right.start.row, right.end.row) <= Math.max(left.start.row, left.end.row) &&
		Math.min(left.start.col, left.end.col) <= Math.max(right.start.col, right.end.col) &&
		Math.min(right.start.col, right.end.col) <= Math.max(left.start.col, left.end.col)
	)
}

export function findOverlappingTable(
	sheet: Pick<Sheet, 'tables'>,
	ref: RangeRef,
	exceptTableId?: TableId,
): Table | null {
	return (
		sheet.tables.find(
			(table) => table.id !== exceptTableId && tableRangesOverlap(table.ref, ref),
		) ?? null
	)
}

export function findTableRangeOverlaps(sheet: Pick<Sheet, 'tables'>): TableRangeOverlap[] {
	const overlaps: TableRangeOverlap[] = []
	for (let i = 0; i < sheet.tables.length; i++) {
		const left = sheet.tables[i]
		if (!left) continue
		for (let j = i + 1; j < sheet.tables.length; j++) {
			const right = sheet.tables[j]
			if (!right || !tableRangesOverlap(left.ref, right.ref)) continue
			overlaps.push({ left, leftRef: left.ref, right, rightRef: right.ref })
		}
	}
	return overlaps
}

export function findShiftedTableRangeOverlap(
	sheet: Pick<Sheet, 'tables'>,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): TableRangeOverlap | null {
	const shiftedTables: Array<{ readonly table: Table; readonly ref: RangeRef }> = []
	for (const table of sheet.tables) {
		const ref = shiftRangeRef(table.ref, axis, at, delta)
		if (ref) shiftedTables.push({ table, ref })
	}
	for (let i = 0; i < shiftedTables.length; i++) {
		const left = shiftedTables[i]
		if (!left) continue
		for (let j = i + 1; j < shiftedTables.length; j++) {
			const right = shiftedTables[j]
			if (!right || !tableRangesOverlap(left.ref, right.ref)) continue
			return {
				left: left.table,
				leftRef: left.ref,
				right: right.table,
				rightRef: right.ref,
			}
		}
	}
	return null
}

export function findQueryTableColumnShiftBlocker(
	sheet: Pick<Sheet, 'tables'>,
	at: number,
	delta: number,
): QueryTableColumnShiftBlocker | null {
	for (const table of sheet.tables) {
		if (!table.queryTable) continue
		const shiftedRef = shiftRangeRef(table.ref, 'col', at, delta)
		if (!shiftedRef) continue
		if (rangeWidth(shiftedRef) !== rangeWidth(table.ref)) {
			return { table, currentRef: table.ref, shiftedRef }
		}
	}
	return null
}

function rangeWidth(ref: RangeRef): number {
	return ref.end.col - ref.start.col + 1
}
