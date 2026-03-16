import {
	type CellRef,
	parseA1,
	parseRange,
	type RangeRef,
	toA1,
	toRangeString,
	type Workbook,
} from '@ascend/core'

export interface CellSelectorObject {
	readonly sheet?: string
	readonly cell: CellRef
}

export interface RangeSelectorObject {
	readonly sheet?: string
	readonly range: RangeRef
}

export type CellSelector = string | CellSelectorObject
export type RangeSelector = string | RangeSelectorObject

export function normalizeCellSelector(
	selector: CellSelector,
	workbook: Workbook,
): { sheetName: string; ref: string; cacheKey: string } {
	if (typeof selector === 'string') {
		const { sheetName, ref } = parseFullRef(selector, workbook)
		return { sheetName, ref, cacheKey: `${sheetName}!${ref}` }
	}
	assertCellRef(selector.cell)
	const sheetName = selector.sheet ?? defaultSheetName(workbook)
	const ref = toA1(selector.cell)
	return { sheetName, ref, cacheKey: `${sheetName}!${ref}` }
}

export function normalizeRangeSelector(
	selector: RangeSelector,
	workbook: Workbook,
): { sheetName: string; ref: string } {
	if (typeof selector === 'string') {
		return parseFullRef(selector, workbook)
	}
	assertRangeRef(selector.range)
	const sheetName = selector.sheet ?? selector.range.sheet ?? defaultSheetName(workbook)
	const range: RangeRef = {
		start: selector.range.start,
		end: selector.range.end,
	}
	return { sheetName, ref: toRangeString(range) }
}

export function parseLocalCellSelector(selector: CellSelector): { ref: string; cell: CellRef } {
	if (typeof selector === 'string') {
		return { ref: selector, cell: parseA1(selector) }
	}
	assertCellRef(selector.cell)
	return { ref: toA1(selector.cell), cell: selector.cell }
}

export function parseLocalRangeSelector(selector: RangeSelector): { ref: RangeRef; text: string } {
	if (typeof selector === 'string') {
		const ref = parseRange(selector)
		return { ref, text: selector }
	}
	assertRangeRef(selector.range)
	return { ref: selector.range, text: toRangeString(selector.range) }
}

export function parseFullRef(
	cellRef: string,
	workbook: Workbook,
): { sheetName: string; ref: string } {
	const bang = cellRef.indexOf('!')
	if (bang !== -1) {
		const sheetName = cellRef.substring(0, bang).replace(/^'|'$/g, '')
		return { sheetName, ref: cellRef.substring(bang + 1) }
	}
	return { sheetName: defaultSheetName(workbook), ref: cellRef }
}

function defaultSheetName(workbook: Workbook): string {
	return workbook.sheets[0]?.name ?? 'Sheet1'
}

function assertCellRef(cell: CellRef): void {
	assertIndex(cell.row, 'row')
	assertIndex(cell.col, 'col')
}

function assertRangeRef(range: RangeRef): void {
	assertCellRef(range.start)
	assertCellRef(range.end)
	if (range.end.row < range.start.row || range.end.col < range.start.col) {
		throw new Error('Invalid range selector: end must not precede start')
	}
}

function assertIndex(value: number, label: string): void {
	if (!Number.isInteger(value) || value < 0) {
		throw new Error(`Invalid ${label} index: ${value}`)
	}
}
