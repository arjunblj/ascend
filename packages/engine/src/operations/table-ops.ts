import type { AutoFilter, Sheet, Table, TableColumn, TableStyleInfo, Workbook } from '@ascend/core'
import { createTableId, toA1 } from '@ascend/core'
import { normalizeFormulaInput } from '@ascend/formulas'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, EMPTY, err, ok, stringValue } from '@ascend/schema'
import {
	rewriteFormulaTextForTableColumnRename,
	rewriteTableColumnInDefinedNames,
	rewriteTableColumnInFormulas,
	rewriteTableNameInDefinedNames,
	rewriteTableNameInFormulas,
} from '../structural/formula-rewrite.ts'
import { expandSqrefRows } from '../structural/ref-shift.ts'
import { sortSheetRange } from '../structural/sort-range.ts'
import type { PatchResult } from './helpers.ts'
import {
	buildTableColumns,
	cellWithExisting,
	clearFormulaMetadata,
	clearFormulaMetadataForSheet,
	createLegacyArrayFormulaIndex,
	DEFAULT_SID,
	findTable,
	getSheet,
	inputToCellValue,
	legacyArrayFormulaEditError,
	patch,
	safeParseRange,
} from './helpers.ts'

export function handleCreateTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'createTable' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	const rangeResult = safeParseRange(op.ref)
	if (!rangeResult.ok) return rangeResult
	if (sheet.tables.some((table) => table.name === op.name)) {
		return err(
			ascendError('NAME_CONFLICT', `Table "${op.name}" already exists`, {
				suggestedFix:
					'Choose a different table name or use a range that does not overlap with existing tables',
			}),
		)
	}

	const ref = rangeResult.value
	const width = ref.end.col - ref.start.col + 1
	const columns = buildTableColumns(sheet, ref, width, op.hasHeaders)
	sheet.tables.push({
		id: createTableId(),
		name: op.name,
		sheetId: sheet.id,
		ref,
		columns,
		hasHeaders: op.hasHeaders,
		hasTotals: false,
	})
	if (op.hasHeaders) {
		sheet.autoFilter = {
			ref: op.ref,
			columns: [],
		}
	}
	return ok(patch([], [op.sheet]))
}

export function handleAppendRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'appendRows' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	const { table, sheet } = located
	if (op.rows.length === 0) return ok(patch([], [sheet.name], false))

	const width = table.columns.length
	const affected: string[] = []
	const originalEndRow = table.ref.end.row
	if (table.hasTotals) {
		sheet.cells.insertRows(originalEndRow, op.rows.length)
	}
	let nextRow = table.hasTotals ? originalEndRow : originalEndRow + 1
	for (const row of op.rows) {
		if (row.length > width) {
			return err(
				ascendError(
					'VALIDATION_ERROR',
					`Appended row has ${row.length} values but table "${table.name}" expects ${width}`,
				),
			)
		}
		for (let colOffset = 0; colOffset < width; colOffset++) {
			const col = table.ref.start.col + colOffset
			const ref = toA1({ row: nextRow, col })
			const provided = row[colOffset]
			const existing = sheet.cells.get(nextRow, col)
			if (provided !== undefined) {
				sheet.cells.set(
					nextRow,
					col,
					cellWithExisting(
						inputToCellValue(provided, workbook.calcSettings.dateSystem),
						existing?.formula ?? null,
						existing?.styleId ?? DEFAULT_SID,
						existing?.formulaInfo,
					),
				)
			} else {
				const formula = table.columns[colOffset]?.formula
				sheet.cells.set(
					nextRow,
					col,
					cellWithExisting(
						existing?.value ?? EMPTY,
						formula ?? null,
						existing?.styleId ?? DEFAULT_SID,
					),
				)
			}
			affected.push(ref)
		}
		nextRow++
	}

	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		const rowDelta = op.rows.length
		sheet.tables.splice(tableIndex, 1, {
			...table,
			ref: {
				start: table.ref.start,
				end: { row: originalEndRow + rowDelta, col: table.ref.end.col },
			},
			...(table.autoFilter
				? {
						autoFilter: {
							...table.autoFilter,
							ref: expandSqrefRows(table.autoFilter.ref, rowDelta),
							...(table.autoFilter.sortState
								? {
										sortState: {
											...table.autoFilter.sortState,
											ref: expandSqrefRows(table.autoFilter.sortState.ref, rowDelta),
										},
									}
								: {}),
						},
					}
				: {}),
			...(table.sortState
				? {
						sortState: {
							...table.sortState,
							ref: expandSqrefRows(table.sortState.ref, rowDelta),
						},
					}
				: {}),
		})
	}
	return ok(patch(affected, [sheet.name], true))
}

export function handleSortRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'sortRange' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const range = rangeResult.value
	const legacyArrayImpact = createLegacyArrayFormulaIndex(sheet).findIntersection(range)
	if (legacyArrayImpact) {
		return err(legacyArrayFormulaEditError(legacyArrayImpact.targetRef, legacyArrayImpact.ref))
	}
	clearFormulaMetadataForSheet(sheet)
	const sorted = sortSheetRange(workbook, sheet, range, op.by)
	if (!sorted.ok) return sorted
	return ok(patch([], [op.sheet], sorted.value))
}

export function handleDeleteTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteTable' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	const { table, sheet } = located
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	if (idx >= 0) sheet.tables.splice(idx, 1)
	if (sheet.autoFilter?.ref) {
		const af = safeParseRange(sheet.autoFilter.ref)
		if (
			af.ok &&
			af.value.start.row === table.ref.start.row &&
			af.value.start.col === table.ref.start.col &&
			af.value.end.row === table.ref.end.row &&
			af.value.end.col === table.ref.end.col
		) {
			sheet.autoFilter = null
		}
	}
	clearFormulaMetadata(workbook)
	return ok(patch([], [sheet.name], true))
}

export function handleRenameTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameTable' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	const { table, sheet } = located
	if (sheet.tables.some((t) => t.name === op.newName && t.id !== table.id)) {
		return err(ascendError('NAME_CONFLICT', `Table "${op.newName}" already exists`))
	}
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	if (idx >= 0) {
		sheet.tables[idx] = { ...table, name: op.newName }
	}
	clearFormulaMetadata(workbook)
	rewriteTableNameInFormulas(workbook, table.name, op.newName)
	rewriteTableNameInDefinedNames(workbook, table.name, op.newName)
	return ok(patch([], [sheet.name], true))
}

export function handleResizeTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'resizeTable' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	const { table, sheet } = located
	const rangeResult = safeParseRange(op.ref)
	if (!rangeResult.ok) return rangeResult
	const ref = rangeResult.value
	const width = ref.end.col - ref.start.col + 1
	const columns =
		width === table.columns.length
			? table.columns
			: buildTableColumns(sheet, ref, width, table.hasHeaders)
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	if (idx >= 0) {
		sheet.tables[idx] = {
			...table,
			ref,
			columns,
			...(table.autoFilter ? { autoFilter: resizeTableAutoFilter(table.autoFilter, op.ref) } : {}),
			...(table.sortState ? { sortState: { ...table.sortState, ref: op.ref } } : {}),
		}
	}
	clearFormulaMetadata(workbook)
	return ok(patch([], [sheet.name], true))
}

export function handleSetTableColumn(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableColumn' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	const { table, sheet } = located
	const columnIndex = resolveTableColumnIndex(table.columns, op.column)
	if (columnIndex < 0) {
		return err(
			ascendError(
				'NAME_NOT_FOUND',
				`Column "${String(op.column)}" not found in table "${op.table}"`,
				{
					suggestedFix: `Available columns: ${table.columns.map((column) => column.name).join(', ')}`,
				},
			),
		)
	}

	const column = table.columns[columnIndex]
	if (!column) return err(ascendError('NAME_NOT_FOUND', `Column "${String(op.column)}" not found`))
	const newNameLower = op.newName?.toLowerCase()
	if (
		newNameLower !== undefined &&
		table.columns.some(
			(candidate, index) => index !== columnIndex && candidate.name.toLowerCase() === newNameLower,
		)
	) {
		return err(
			ascendError('NAME_CONFLICT', `Column "${op.newName}" already exists in table "${op.table}"`),
		)
	}
	const nextColumn = updateTableColumn(column, op)
	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		sheet.tables[tableIndex] = {
			...table,
			columns: table.columns.map((candidate, index) =>
				index === columnIndex ? nextColumn : candidate,
			),
		}
	}

	const affected: string[] = []
	let recalcRequired = op.formula !== undefined || op.newName !== undefined
	if (op.newName !== undefined) {
		if (table.hasHeaders) {
			const headerCell = sheet.cells.get(table.ref.start.row, table.ref.start.col + columnIndex)
			sheet.cells.set(table.ref.start.row, table.ref.start.col + columnIndex, {
				value: stringValue(op.newName),
				formula: null,
				styleId: headerCell?.styleId ?? DEFAULT_SID,
			})
			affected.push(toA1({ row: table.ref.start.row, col: table.ref.start.col + columnIndex }))
		}
		clearFormulaMetadata(workbook)
		rewriteTableColumnInFormulas(workbook, table, column.name, op.newName)
		rewriteTableColumnInDefinedNames(workbook, table.name, column.name, op.newName)
	}
	const totalsResult = materializeTotalsRowCell(workbook, sheet, table, columnIndex, nextColumn, op)
	if (totalsResult) {
		affected.push(totalsResult.ref)
		recalcRequired ||= totalsResult.recalcRequired
	}
	if (op.formula !== undefined) {
		let formula = op.formula === null ? null : normalizeFormulaInput(op.formula)
		if (formula !== null && op.newName !== undefined) {
			formula =
				rewriteFormulaTextForTableColumnRename(
					formula,
					table.name,
					column.name,
					op.newName,
					true,
				) ?? formula
		}
		const col = table.ref.start.col + columnIndex
		const startRow = table.ref.start.row + (table.hasHeaders ? 1 : 0)
		const endRow = table.ref.end.row - (table.hasTotals ? 1 : 0)
		for (let row = startRow; row <= endRow; row++) {
			const existing = sheet.cells.get(row, col)
			sheet.cells.set(
				row,
				col,
				cellWithExisting(existing?.value ?? EMPTY, formula, existing?.styleId ?? DEFAULT_SID),
			)
			affected.push(toA1({ row, col }))
		}
		clearFormulaMetadata(workbook)
	}

	return ok(patch(affected, [sheet.name], recalcRequired))
}

export function handleSetTableStyle(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableStyle' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	if (
		op.styleName === undefined &&
		op.showFirstColumn === undefined &&
		op.showLastColumn === undefined &&
		op.showRowStripes === undefined &&
		op.showColumnStripes === undefined
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'setTableStyle requires at least one style field', {
				suggestedFix:
					'Provide styleName, showFirstColumn, showLastColumn, showRowStripes, or showColumnStripes.',
			}),
		)
	}

	const { table, sheet } = located
	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		const nextStyle = updateTableStyle(table.tableStyleInfo, op)
		const { tableStyleInfo: _tableStyleInfo, ...tableWithoutStyle } = table
		sheet.tables[tableIndex] = nextStyle
			? { ...tableWithoutStyle, tableStyleInfo: nextStyle }
			: tableWithoutStyle
	}
	return ok(patch([], [sheet.name], false))
}

function resolveTableColumnIndex(columns: readonly TableColumn[], column: string | number): number {
	if (typeof column === 'number') {
		return Number.isInteger(column) && column >= 0 && column < columns.length ? column : -1
	}
	return columns.findIndex((candidate) => candidate.name.toLowerCase() === column.toLowerCase())
}

function updateTableColumn(
	column: TableColumn,
	op: Extract<Operation, { op: 'setTableColumn' }>,
): TableColumn {
	let next: TableColumn = { ...column }
	if (op.newName !== undefined) next = { ...next, name: op.newName }
	if (op.formula !== undefined) {
		const { formula: _formula, formulaIsArray: _formulaIsArray, ...rest } = next
		next = op.formula === null ? rest : { ...rest, formula: normalizeFormulaInput(op.formula) }
	}
	if (op.totalsRowFunction !== undefined) {
		const { totalsRowFunction: _totalsRowFunction, ...rest } = next
		next =
			op.totalsRowFunction === null ? rest : { ...rest, totalsRowFunction: op.totalsRowFunction }
	}
	if (op.totalsRowFormula !== undefined) {
		const { totalsRowFormula: _totalsRowFormula, ...rest } = next
		next =
			op.totalsRowFormula === null
				? rest
				: { ...rest, totalsRowFormula: normalizeFormulaInput(op.totalsRowFormula) }
	}
	if (op.totalsRowLabel !== undefined) {
		const { totalsRowLabel: _totalsRowLabel, ...rest } = next
		next = op.totalsRowLabel === null ? rest : { ...rest, totalsRowLabel: op.totalsRowLabel }
	}
	return next
}

function materializeTotalsRowCell(
	workbook: Workbook,
	sheet: Sheet,
	table: Table,
	columnIndex: number,
	column: TableColumn,
	op: Extract<Operation, { op: 'setTableColumn' }>,
): { readonly ref: string; readonly recalcRequired: boolean } | null {
	if (!table.hasTotals) return null
	const touchedTotals =
		op.totalsRowFunction !== undefined ||
		op.totalsRowFormula !== undefined ||
		op.totalsRowLabel !== undefined
	if (!touchedTotals) return null

	const row = table.ref.end.row
	const col = table.ref.start.col + columnIndex
	const ref = toA1({ row, col })
	const existing = sheet.cells.get(row, col)
	if (column.totalsRowFormula) {
		sheet.cells.set(row, col, {
			value: existing?.value ?? EMPTY,
			formula: column.totalsRowFormula,
			styleId: existing?.styleId ?? DEFAULT_SID,
		})
		clearFormulaMetadata(workbook)
		return { ref, recalcRequired: true }
	}
	const subtotalFormula = subtotalFormulaFor(column.totalsRowFunction, table.name, column.name)
	if (subtotalFormula) {
		sheet.cells.set(row, col, {
			value: existing?.value ?? EMPTY,
			formula: subtotalFormula,
			styleId: existing?.styleId ?? DEFAULT_SID,
		})
		clearFormulaMetadata(workbook)
		return { ref, recalcRequired: true }
	}
	if (column.totalsRowLabel) {
		sheet.cells.set(row, col, {
			value: stringValue(column.totalsRowLabel),
			formula: null,
			styleId: existing?.styleId ?? DEFAULT_SID,
		})
		return { ref, recalcRequired: false }
	}
	sheet.cells.delete(row, col)
	return { ref, recalcRequired: false }
}

function subtotalFormulaFor(
	totalsRowFunction: string | undefined,
	tableName: string,
	columnName: string,
): string | null {
	const code = TOTALS_ROW_SUBTOTAL_CODES.get(totalsRowFunction?.toLowerCase() ?? '')
	if (code === undefined) return null
	return `SUBTOTAL(${code},${tableName}[${escapeStructuredRefColumn(columnName)}])`
}

const TOTALS_ROW_SUBTOTAL_CODES = new Map([
	['average', 101],
	['count', 103],
	['countnums', 102],
	['max', 104],
	['min', 105],
	['sum', 109],
	['stddev', 107],
	['var', 110],
])

function escapeStructuredRefColumn(name: string): string {
	return name.replace(/([#@[\]'])/g, "'$1")
}

function updateTableStyle(
	style: TableStyleInfo | undefined,
	op: Extract<Operation, { op: 'setTableStyle' }>,
): TableStyleInfo | undefined {
	const next: TableStyleInfo = {
		...(style ?? {}),
		...(op.styleName === undefined || op.styleName === null ? {} : { name: op.styleName }),
		...(op.showFirstColumn === undefined ? {} : { showFirstColumn: op.showFirstColumn }),
		...(op.showLastColumn === undefined ? {} : { showLastColumn: op.showLastColumn }),
		...(op.showRowStripes === undefined ? {} : { showRowStripes: op.showRowStripes }),
		...(op.showColumnStripes === undefined ? {} : { showColumnStripes: op.showColumnStripes }),
	}
	if (op.styleName === null) {
		const { name: _name, ...withoutName } = next
		return Object.keys(withoutName).length > 0 ? withoutName : undefined
	}
	return Object.keys(next).length > 0 ? next : undefined
}

function resizeTableAutoFilter(autoFilter: AutoFilter, ref: string): AutoFilter {
	return {
		...autoFilter,
		ref,
		...(autoFilter.sortState ? { sortState: { ...autoFilter.sortState, ref } } : {}),
	}
}
