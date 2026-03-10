import type { Cell, RangeRef, Sheet, Workbook } from '@ascend/core'
import { columnToIndex, parseA1, parseRange, toA1 } from '@ascend/core'
import { compareValues } from '@ascend/formulas'
import type { CellValue, Result, SortSpec } from '@ascend/schema'
import { ascendError, EMPTY, err, ok } from '@ascend/schema'

export function sortSheetRange(
	_workbook: Workbook,
	sheet: Sheet,
	range: RangeRef,
	specs: readonly SortSpec[],
): Result<boolean> {
	if (sheet.merges.some((merge) => rangesOverlap(merge, range))) {
		return err(ascendError('VALIDATION_ERROR', 'sortRange does not support merged cells yet'))
	}

	const resolvedColumns = resolveSortColumns(sheet, range, specs)
	if (!resolvedColumns.ok) return resolvedColumns
	const { columns, headerRow } = resolvedColumns.value
	const dataStartRow = headerRow ? range.start.row + 1 : range.start.row
	if (dataStartRow > range.end.row) return ok(false)

	const rows = captureSortedRows(sheet, range, dataStartRow)
	rows.sort((left, right) => compareSortRows(left, right, columns, range.start.col))
	rewriteSortedRows(sheet, range, dataStartRow, rows)
	return ok(true)
}

function resolveSortColumns(
	sheet: Sheet,
	range: RangeRef,
	specs: readonly SortSpec[],
): Result<{ columns: readonly { col: number; descending: boolean }[]; headerRow: boolean }> {
	if (specs.length === 0) {
		return err(ascendError('VALIDATION_ERROR', 'sortRange requires at least one sort key'))
	}

	const headerMap = new Map<string, number>()
	for (let col = range.start.col; col <= range.end.col; col++) {
		const value = sheet.cells.get(range.start.row, col)?.value
		if (value?.kind === 'string' && value.value.trim() !== '') {
			headerMap.set(value.value.trim().toLowerCase(), col)
		}
	}

	let headerRow = false
	const columns: Array<{ col: number; descending: boolean }> = []
	for (const spec of specs) {
		if (typeof spec.column === 'number') {
			const col = range.start.col + Math.trunc(spec.column) - 1
			if (col < range.start.col || col > range.end.col) {
				return err(
					ascendError(
						'VALIDATION_ERROR',
						`Sort column ${spec.column} is outside ${opRange(range)}`,
					),
				)
			}
			columns.push({ col, descending: spec.descending ?? false })
			continue
		}

		const trimmed = spec.column.trim()
		const matched = headerMap.get(trimmed.toLowerCase())
		if (matched !== undefined) {
			headerRow = true
			columns.push({ col: matched, descending: spec.descending ?? false })
			continue
		}

		if (/^[A-Za-z]{1,3}$/.test(trimmed)) {
			const col = columnToIndex(trimmed.toUpperCase())
			if (col < range.start.col || col > range.end.col) {
				return err(
					ascendError('VALIDATION_ERROR', `Sort column ${trimmed} is outside ${opRange(range)}`),
				)
			}
			columns.push({ col, descending: spec.descending ?? false })
			continue
		}

		return err(ascendError('VALIDATION_ERROR', `Unknown sort column "${spec.column}"`))
	}

	return ok({ columns, headerRow })
}

interface SortRow {
	readonly originalIndex: number
	readonly cells: Array<{ colOffset: number; cell: Cell }>
	readonly comments: Array<{
		colOffset: number
		ref: string
		comment: Sheet['comments'] extends Map<string, infer T> ? T : never
	}>
	readonly hyperlinks: Array<{
		colOffset: number
		ref: string
		hyperlink: Sheet['hyperlinks'] extends Map<string, infer T> ? T : never
	}>
	readonly dataValidations: Array<RowScopedSqrefEntry<Sheet['dataValidations'][number]>>
	readonly conditionalFormats: Array<RowScopedSqrefEntry<Sheet['conditionalFormats'][number]>>
	readonly ignoredErrors: Array<RowScopedSqrefEntry<Sheet['ignoredErrors'][number]>>
	readonly rowHeight: number | undefined
}

interface RowScopedSqrefEntry<T extends { sqref: string }> {
	readonly startColOffset: number
	readonly endColOffset: number
	readonly entry: T
}

function captureSortedRows(sheet: Sheet, range: RangeRef, startRow: number): SortRow[] {
	const rows: SortRow[] = []
	for (let row = startRow; row <= range.end.row; row++) {
		const cells: Array<{ colOffset: number; cell: Cell }> = []
		for (let col = range.start.col; col <= range.end.col; col++) {
			const cell = sheet.cells.get(row, col)
			if (cell) cells.push({ colOffset: col - range.start.col, cell })
		}

		const comments: SortRow['comments'] = []
		for (const [ref, comment] of sheet.comments) {
			const pos = parseA1(ref)
			if (pos.row === row && pos.col >= range.start.col && pos.col <= range.end.col) {
				comments.push({ colOffset: pos.col - range.start.col, ref, comment })
			}
		}

		const hyperlinks: SortRow['hyperlinks'] = []
		for (const [ref, hyperlink] of sheet.hyperlinks) {
			const pos = parseA1(ref)
			if (pos.row === row && pos.col >= range.start.col && pos.col <= range.end.col) {
				hyperlinks.push({ colOffset: pos.col - range.start.col, ref, hyperlink })
			}
		}

		rows.push({
			originalIndex: row - startRow,
			cells,
			comments,
			hyperlinks,
			dataValidations: captureRowScopedSqrefEntries(sheet.dataValidations, range, row),
			conditionalFormats: captureRowScopedSqrefEntries(sheet.conditionalFormats, range, row),
			ignoredErrors: captureRowScopedSqrefEntries(sheet.ignoredErrors, range, row),
			rowHeight: sheet.rowHeights.get(row),
		})
	}
	return rows
}

function compareSortRows(
	left: SortRow,
	right: SortRow,
	columns: readonly { col: number; descending: boolean }[],
	startCol: number,
): number {
	for (const column of columns) {
		const leftValue = cellValueAt(left, column.col - startCol)
		const rightValue = cellValueAt(right, column.col - startCol)
		const result = compareValues(leftValue, rightValue)
		if (result !== 0) return column.descending ? -result : result
	}
	return left.originalIndex - right.originalIndex
}

function cellValueAt(row: SortRow, colOffset: number): CellValue {
	const entry = row.cells.find((cell) => cell.colOffset === colOffset)
	return entry?.cell.value ?? EMPTY
}

function rewriteSortedRows(
	sheet: Sheet,
	range: RangeRef,
	startRow: number,
	rows: readonly SortRow[],
): void {
	replaceArrayContents(
		sheet.dataValidations,
		sheet.dataValidations.filter((entry) => !isSortableRowScopedSqrefEntry(entry, range)),
	)
	replaceArrayContents(
		sheet.conditionalFormats,
		sheet.conditionalFormats.filter((entry) => !isSortableRowScopedSqrefEntry(entry, range)),
	)
	replaceArrayContents(
		sheet.ignoredErrors,
		sheet.ignoredErrors.filter((entry) => !isSortableRowScopedSqrefEntry(entry, range)),
	)
	for (let row = startRow; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			sheet.cells.delete(row, col)
			const ref = toA1({ row, col })
			sheet.comments.delete(ref)
			sheet.hyperlinks.delete(ref)
		}
		sheet.rowHeights.delete(row)
	}

	rows.forEach((rowData, index) => {
		const targetRow = startRow + index
		for (const entry of rowData.cells) {
			sheet.cells.set(targetRow, range.start.col + entry.colOffset, entry.cell)
		}
		for (const entry of rowData.comments) {
			sheet.comments.set(
				toA1({ row: targetRow, col: range.start.col + entry.colOffset }),
				entry.comment,
			)
		}
		for (const entry of rowData.hyperlinks) {
			sheet.hyperlinks.set(
				toA1({ row: targetRow, col: range.start.col + entry.colOffset }),
				entry.hyperlink,
			)
		}
		for (const entry of rowData.dataValidations) {
			sheet.dataValidations.push(rewriteRowScopedSqrefEntry(entry, targetRow, range.start.col))
		}
		for (const entry of rowData.conditionalFormats) {
			sheet.conditionalFormats.push(rewriteRowScopedSqrefEntry(entry, targetRow, range.start.col))
		}
		for (const entry of rowData.ignoredErrors) {
			sheet.ignoredErrors.push(rewriteRowScopedSqrefEntry(entry, targetRow, range.start.col))
		}
		if (rowData.rowHeight !== undefined) {
			sheet.rowHeights.set(targetRow, rowData.rowHeight)
		}
	})
}

function captureRowScopedSqrefEntries<T extends { sqref: string }>(
	entries: readonly T[],
	range: RangeRef,
	row: number,
): Array<RowScopedSqrefEntry<T>> {
	const captured: Array<RowScopedSqrefEntry<T>> = []
	for (const entry of entries) {
		const parsed = parseRowScopedSqref(entry.sqref, range)
		if (!parsed || parsed.row !== row) continue
		captured.push({
			startColOffset: parsed.startCol - range.start.col,
			endColOffset: parsed.endCol - range.start.col,
			entry,
		})
	}
	return captured
}

function isSortableRowScopedSqrefEntry(entry: { sqref: string }, range: RangeRef): boolean {
	return parseRowScopedSqref(entry.sqref, range) !== null
}

function parseRowScopedSqref(
	sqref: string,
	range: RangeRef,
): { row: number; startCol: number; endCol: number } | null {
	if (sqref.includes(' ')) return null
	try {
		const parsed = parseRange(sqref)
		if (parsed.start.row !== parsed.end.row) return null
		if (parsed.start.row < range.start.row || parsed.start.row > range.end.row) return null
		if (parsed.start.col < range.start.col || parsed.end.col > range.end.col) return null
		return {
			row: parsed.start.row,
			startCol: parsed.start.col,
			endCol: parsed.end.col,
		}
	} catch {
		return null
	}
}

function rewriteRowScopedSqrefEntry<T extends { sqref: string }>(
	entry: RowScopedSqrefEntry<T>,
	targetRow: number,
	startCol: number,
): T {
	const start = toA1({ row: targetRow, col: startCol + entry.startColOffset })
	const end = toA1({ row: targetRow, col: startCol + entry.endColOffset })
	return {
		...entry.entry,
		sqref: start === end ? start : `${start}:${end}`,
	}
}

function replaceArrayContents<T>(target: T[], next: readonly T[]): void {
	target.splice(0, target.length, ...next)
}

function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
	return !(
		a.end.row < b.start.row ||
		a.start.row > b.end.row ||
		a.end.col < b.start.col ||
		a.start.col > b.end.col
	)
}

function opRange(range: RangeRef): string {
	return `${toA1(range.start)}:${toA1(range.end)}`
}
