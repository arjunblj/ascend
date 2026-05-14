import type {
	Cell,
	RangeRef,
	Sheet,
	SheetComment,
	SheetConditionalFormatValueObject,
	Workbook,
} from '@ascend/core'
import { columnToIndex, parseA1, parseRange, toA1 } from '@ascend/core'
import { cachedParseFormula, compareValues, printFormulaWithOffset } from '@ascend/formulas'
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
		const value = sheet.cells.readValue(range.start.row, col)
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
	readonly threadedComments: Array<{
		colOffset: number
		comment: Sheet['threadedComments'][number]
	}>
	readonly hyperlinks: Array<{
		colOffset: number
		ref: string
		hyperlink: Sheet['hyperlinks'] extends Map<string, infer T> ? T : never
	}>
	readonly dataValidations: Array<RowScopedSqrefEntry<Sheet['dataValidations'][number]>>
	readonly conditionalFormats: Array<RowScopedSqrefEntry<Sheet['conditionalFormats'][number]>>
	readonly x14DataValidations: Array<RowScopedSqrefEntry<Sheet['x14DataValidations'][number]>>
	readonly x14ConditionalFormats: Array<RowScopedSqrefEntry<Sheet['x14ConditionalFormats'][number]>>
	readonly ignoredErrors: Array<RowScopedSqrefEntry<Sheet['ignoredErrors'][number]>>
	readonly rowHeight: number | undefined
	readonly rowDef: RowDef | undefined
}

type RowDef = NonNullable<ReturnType<Sheet['rowDefs']['get']>>

interface RowScopedSqrefEntry<T extends { sqref: string }> {
	readonly row: number
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

		const threadedComments: SortRow['threadedComments'] = []
		for (const comment of sheet.threadedComments) {
			const pos = parseA1(comment.ref)
			if (pos.row === row && pos.col >= range.start.col && pos.col <= range.end.col) {
				threadedComments.push({ colOffset: pos.col - range.start.col, comment })
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
			threadedComments,
			hyperlinks,
			dataValidations: captureRowScopedSqrefEntries(sheet.dataValidations, range, row),
			conditionalFormats: captureRowScopedSqrefEntries(sheet.conditionalFormats, range, row),
			x14DataValidations: captureRowScopedSqrefEntries(
				sheet.x14DataValidations.filter((entry) => !entry.deleted),
				range,
				row,
			),
			x14ConditionalFormats: captureRowScopedSqrefEntries(
				sheet.x14ConditionalFormats.filter((entry) => !entry.deleted),
				range,
				row,
			),
			ignoredErrors: captureRowScopedSqrefEntries(sheet.ignoredErrors, range, row),
			rowHeight: sheet.rowHeights.get(row),
			rowDef: cloneRowDef(sheet.rowDefs.get(row)),
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
	replaceArrayContents(
		sheet.x14DataValidations,
		sheet.x14DataValidations.filter(
			(entry) => entry.deleted || !isSortableRowScopedSqrefEntry(entry, range),
		),
	)
	replaceArrayContents(
		sheet.x14ConditionalFormats,
		sheet.x14ConditionalFormats.filter(
			(entry) => entry.deleted || !isSortableRowScopedSqrefEntry(entry, range),
		),
	)
	for (let row = startRow; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			sheet.cells.delete(row, col)
			const ref = toA1({ row, col })
			sheet.comments.delete(ref)
			sheet.hyperlinks.delete(ref)
		}
		sheet.rowHeights.delete(row)
		sheet.rowDefs.delete(row)
	}
	sheet.threadedComments = sheet.threadedComments.filter((comment) => {
		const pos = parseA1(comment.ref)
		return (
			pos.row < startRow ||
			pos.row > range.end.row ||
			pos.col < range.start.col ||
			pos.col > range.end.col
		)
	})

	rows.forEach((rowData, index) => {
		const targetRow = startRow + index
		const rowDelta = targetRow - (startRow + rowData.originalIndex)
		for (const entry of rowData.cells) {
			sheet.cells.set(
				targetRow,
				range.start.col + entry.colOffset,
				rewriteSortedCell(entry.cell, rowDelta, 0),
			)
		}
		for (const entry of rowData.comments) {
			const target = { row: targetRow, col: range.start.col + entry.colOffset }
			const original = parseA1(entry.ref)
			sheet.comments.set(
				toA1(target),
				retargetLegacyCommentDrawing(entry.comment, target, target.row - original.row),
			)
		}
		for (const entry of rowData.threadedComments) {
			sheet.threadedComments.push({
				...entry.comment,
				ref: toA1({ row: targetRow, col: range.start.col + entry.colOffset }),
			})
		}
		for (const entry of rowData.hyperlinks) {
			sheet.hyperlinks.set(
				toA1({ row: targetRow, col: range.start.col + entry.colOffset }),
				entry.hyperlink,
			)
		}
		for (const entry of rowData.dataValidations) {
			sheet.dataValidations.push(
				rewriteRowScopedDataValidationEntry(entry, targetRow, range.start.col),
			)
		}
		for (const entry of rowData.conditionalFormats) {
			sheet.conditionalFormats.push(
				rewriteRowScopedConditionalFormatEntry(entry, targetRow, range.start.col),
			)
		}
		for (const entry of rowData.x14DataValidations) {
			sheet.x14DataValidations.push(
				rewriteRowScopedDataValidationEntry(entry, targetRow, range.start.col),
			)
		}
		for (const entry of rowData.x14ConditionalFormats) {
			sheet.x14ConditionalFormats.push(
				rewriteRowScopedX14ConditionalFormatEntry(entry, targetRow, range.start.col),
			)
		}
		for (const entry of rowData.ignoredErrors) {
			sheet.ignoredErrors.push(rewriteRowScopedSqrefEntry(entry, targetRow, range.start.col))
		}
		if (rowData.rowHeight !== undefined) {
			sheet.rowHeights.set(targetRow, rowData.rowHeight)
		}
		if (rowData.rowDef !== undefined) {
			sheet.rowDefs.set(targetRow, rowData.rowDef)
		}
	})
}

function rewriteSortedCell(cell: Cell, rowDelta: number, colDelta: number): Cell {
	if (!cell.formula || (rowDelta === 0 && colDelta === 0)) return cell
	return {
		...cell,
		formula: translateMetadataFormula(cell.formula, rowDelta, colDelta),
	}
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
			row: parsed.row,
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

function retargetLegacyCommentDrawing(
	comment: SheetComment,
	target: { readonly row: number; readonly col: number },
	rowDelta: number,
): SheetComment {
	const drawing = comment.legacyDrawing
	if (!drawing) return comment
	return {
		...comment,
		legacyDrawing: {
			...drawing,
			...(drawing.row !== undefined ? { row: target.row } : {}),
			...(drawing.column !== undefined ? { column: target.col } : {}),
			...(drawing.anchor ? { anchor: translateLegacyCommentAnchor(drawing.anchor, rowDelta) } : {}),
		},
	}
}

function translateLegacyCommentAnchor(
	anchor: NonNullable<NonNullable<SheetComment['legacyDrawing']>['anchor']>,
	rowDelta: number,
): NonNullable<NonNullable<SheetComment['legacyDrawing']>['anchor']> {
	const next = [...anchor] as [number, number, number, number, number, number, number, number]
	next[2] = Math.max(0, next[2] + rowDelta)
	next[6] = Math.max(next[2], next[6] + rowDelta)
	return next
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

function rewriteRowScopedDataValidationEntry<
	T extends { sqref: string; formula1?: string; formula2?: string },
>(entry: RowScopedSqrefEntry<T>, targetRow: number, startCol: number): T {
	const rewritten = rewriteRowScopedSqrefEntry(entry, targetRow, startCol)
	const rowDelta = targetRow - entry.row
	return {
		...rewritten,
		...(rewritten.formula1 !== undefined
			? { formula1: translateMetadataFormula(rewritten.formula1, rowDelta, 0) }
			: {}),
		...(rewritten.formula2 !== undefined
			? { formula2: translateMetadataFormula(rewritten.formula2, rowDelta, 0) }
			: {}),
	}
}

function rewriteRowScopedConditionalFormatEntry(
	entry: RowScopedSqrefEntry<Sheet['conditionalFormats'][number]>,
	targetRow: number,
	startCol: number,
): Sheet['conditionalFormats'][number] {
	const rewritten = rewriteRowScopedSqrefEntry(entry, targetRow, startCol)
	const rowDelta = targetRow - entry.row
	return {
		...rewritten,
		rules: rewritten.rules.map((rule) => ({
			...rule,
			formulas: rule.formulas.map((formula) => translateMetadataFormula(formula, rowDelta, 0)),
			...(rule.colorScale
				? {
						colorScale: {
							...rule.colorScale,
							cfvo: rule.colorScale.cfvo.map((value) =>
								translateConditionalFormatValueObject(value, rowDelta, 0),
							),
							colors: rule.colorScale.colors.map((color) => ({ ...color })),
						},
					}
				: {}),
			...(rule.dataBar
				? {
						dataBar: {
							...rule.dataBar,
							cfvo: rule.dataBar.cfvo.map((value) =>
								translateConditionalFormatValueObject(value, rowDelta, 0),
							),
							...(rule.dataBar.color ? { color: { ...rule.dataBar.color } } : {}),
						},
					}
				: {}),
			...(rule.iconSet
				? {
						iconSet: {
							...rule.iconSet,
							cfvo: rule.iconSet.cfvo.map((value) =>
								translateConditionalFormatValueObject(value, rowDelta, 0),
							),
						},
					}
				: {}),
		})),
	}
}

function rewriteRowScopedX14ConditionalFormatEntry(
	entry: RowScopedSqrefEntry<Sheet['x14ConditionalFormats'][number]>,
	targetRow: number,
	startCol: number,
): Sheet['x14ConditionalFormats'][number] {
	const rewritten = rewriteRowScopedSqrefEntry(entry, targetRow, startCol)
	const rowDelta = targetRow - entry.row
	return {
		...rewritten,
		formulas: rewritten.formulas.map((formula) => translateMetadataFormula(formula, rowDelta, 0)),
		...(rewritten.colorScale
			? {
					colorScale: {
						...rewritten.colorScale,
						cfvo: rewritten.colorScale.cfvo.map((value) =>
							translateConditionalFormatValueObject(value, rowDelta, 0),
						),
						colors: rewritten.colorScale.colors.map((color) => ({ ...color })),
					},
				}
			: {}),
		...(rewritten.dataBar
			? {
					dataBar: {
						...rewritten.dataBar,
						cfvo: rewritten.dataBar.cfvo.map((value) =>
							translateConditionalFormatValueObject(value, rowDelta, 0),
						),
						...(rewritten.dataBar.fillColor
							? { fillColor: { ...rewritten.dataBar.fillColor } }
							: {}),
						...(rewritten.dataBar.borderColor
							? { borderColor: { ...rewritten.dataBar.borderColor } }
							: {}),
						...(rewritten.dataBar.negativeFillColor
							? { negativeFillColor: { ...rewritten.dataBar.negativeFillColor } }
							: {}),
						...(rewritten.dataBar.negativeBorderColor
							? { negativeBorderColor: { ...rewritten.dataBar.negativeBorderColor } }
							: {}),
						...(rewritten.dataBar.axisColor
							? { axisColor: { ...rewritten.dataBar.axisColor } }
							: {}),
					},
				}
			: {}),
		...(rewritten.iconSet
			? {
					iconSet: {
						...rewritten.iconSet,
						cfvo: rewritten.iconSet.cfvo.map((value) =>
							translateConditionalFormatValueObject(value, rowDelta, 0),
						),
						...(rewritten.iconSet.icons
							? { icons: rewritten.iconSet.icons.map((icon) => ({ ...icon })) }
							: {}),
					},
				}
			: {}),
	}
}

function translateConditionalFormatValueObject<T extends SheetConditionalFormatValueObject>(
	entry: T,
	rowDelta: number,
	colDelta: number,
): T {
	if (entry.value === undefined) return entry
	return {
		...entry,
		value: translateMetadataFormula(entry.value, rowDelta, colDelta),
	}
}

function translateMetadataFormula(formula: string, rowDelta: number, colDelta: number): string {
	const parsed = cachedParseFormula(formula)
	return parsed.ok ? printFormulaWithOffset(parsed.value, rowDelta, colDelta) : formula
}

function replaceArrayContents<T>(target: T[], next: readonly T[]): void {
	target.splice(0, target.length, ...next)
}

function cloneRowDef(def: RowDef | undefined): RowDef | undefined {
	return def ? { ...def } : undefined
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
