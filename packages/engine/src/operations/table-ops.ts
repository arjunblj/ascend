import type {
	AutoFilter,
	RangeRef,
	Sheet,
	SortCondition,
	SortState,
	Table,
	TableColumn,
	TableId,
	TableStyleInfo,
	Workbook,
} from '@ascend/core'
import { createTableId, parseA1Safe, parseRange, toA1 } from '@ascend/core'
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
import { sortSheetRange } from '../structural/sort-range.ts'
import {
	collectDroppedTableColumnsForResize,
	type DeletedTableColumnReference,
	findDeletedTableColumnReference,
} from '../structural/table-field-guards.ts'
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
	const nameError = validateTableName(op.name)
	if (nameError) return err(nameError)
	if (findTableNameCollision(workbook, op.name)) {
		return err(
			ascendError('NAME_CONFLICT', `Table "${op.name}" already exists`, {
				suggestedFix:
					'Choose a workbook-unique table name so structured references remain unambiguous.',
			}),
		)
	}

	const ref = rangeResult.value
	const overlappingTable = findOverlappingTable(sheet, ref)
	if (overlappingTable) {
		return err(tableRangeOverlapError('create', op.name, op.ref, overlappingTable))
	}
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
	const rowDelta = op.rows.length
	const nextTableRef = {
		start: table.ref.start,
		end: { row: originalEndRow + rowDelta, col: table.ref.end.col },
	}
	const overlappingTable = findOverlappingTable(sheet, nextTableRef, table.id)
	if (overlappingTable) {
		return err(
			tableRangeOverlapError('append', table.name, rangeToA1(nextTableRef), overlappingTable),
		)
	}
	if (table.hasTotals) {
		const shiftedTable = findTableShiftedByTotalsAppend(sheet, table, originalEndRow)
		if (shiftedTable) return err(tableAppendTotalsShiftError(table, shiftedTable))
	}
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
		const autoFilter = table.autoFilter
			? resizeTableAutoFilter(table.autoFilter, table.ref, nextTableRef)
			: undefined
		const sortState = resizeTableSortState(table.sortState, table.ref, nextTableRef)
		sheet.tables.splice(tableIndex, 1, {
			...table,
			ref: {
				start: table.ref.start,
				end: { row: originalEndRow + rowDelta, col: table.ref.end.col },
			},
			...(autoFilter ? { autoFilter } : {}),
			...(sortState ? { sortState } : {}),
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
	const nameError = validateTableName(op.newName)
	if (nameError) return err(nameError)
	if (findTableNameCollision(workbook, op.newName, table.id)) {
		return err(
			ascendError('NAME_CONFLICT', `Table "${op.newName}" already exists`, {
				suggestedFix:
					'Choose a workbook-unique table name so structured references remain unambiguous.',
			}),
		)
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
	const overlappingTable = findOverlappingTable(sheet, ref, table.id)
	if (overlappingTable) {
		return err(tableRangeOverlapError('resize', table.name, op.ref, overlappingTable))
	}
	const droppedColumnBlocker = findDeletedTableColumnReference(
		workbook,
		collectDroppedTableColumnsForResize(table, ref),
		{ skipDeletedTableColumnFormulas: true },
	)
	if (droppedColumnBlocker) {
		return err(tableResizeDroppedColumnReferenceError(droppedColumnBlocker))
	}
	const columns = buildResizedTableColumns(sheet, table, ref)
	const autoFilter = table.autoFilter
		? resizeTableAutoFilter(table.autoFilter, table.ref, ref)
		: undefined
	const sortState = resizeTableSortState(table.sortState, table.ref, ref)
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	if (idx >= 0) {
		const { autoFilter: _autoFilter, sortState: _sortState, ...tableWithoutFilterState } = table
		sheet.tables[idx] = {
			...tableWithoutFilterState,
			ref,
			columns,
			...(autoFilter ? { autoFilter } : {}),
			...(sortState ? { sortState } : {}),
		}
	}
	clearFormulaMetadata(workbook)
	return ok(patch([], [sheet.name], true))
}

function validateTableName(name: string): ReturnType<typeof ascendError> | null {
	const suggestedFix =
		'Use a table name that starts with a letter, underscore, or backslash; uses only letters, numbers, periods, and underscores after that; is not C, R, A1-style, or R1C1-style; and is 255 characters or fewer.'
	if (name.length === 0) {
		return ascendError('VALIDATION_ERROR', 'Table name cannot be empty', { suggestedFix })
	}
	if (name.length > 255) {
		return ascendError('VALIDATION_ERROR', `Table name "${name}" exceeds 255 characters`, {
			suggestedFix,
		})
	}
	if (/^[cr]$/i.test(name)) {
		return ascendError('VALIDATION_ERROR', `Table name "${name}" is reserved`, { suggestedFix })
	}
	if (isA1StyleReference(name) || /^R\d+C\d+$/i.test(name)) {
		return ascendError('VALIDATION_ERROR', `Table name "${name}" cannot be a cell reference`, {
			suggestedFix,
		})
	}
	if (!/^[\p{L}_\\][\p{L}\p{N}._]*$/u.test(name)) {
		return ascendError('VALIDATION_ERROR', `Table name "${name}" contains invalid characters`, {
			suggestedFix,
		})
	}
	return null
}

function isA1StyleReference(name: string): boolean {
	const parsed = parseA1Safe(name)
	if (!parsed) return false
	return parsed.col >= 0 && parsed.col <= 16383 && parsed.row >= 0 && parsed.row <= 1048575
}

function findTableNameCollision(
	workbook: Workbook,
	name: string,
	exceptTableId?: TableId,
): Table | null {
	const lowerName = name.toLowerCase()
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.id === exceptTableId) continue
			if (table.name.toLowerCase() === lowerName) return table
		}
	}
	return null
}

function findOverlappingTable(
	sheet: Sheet,
	ref: Table['ref'],
	exceptTableId?: TableId,
): Table | null {
	return (
		sheet.tables.find(
			(table) => table.id !== exceptTableId && tableRangesOverlap(table.ref, ref),
		) ?? null
	)
}

function findTableShiftedByTotalsAppend(
	sheet: Sheet,
	table: Table,
	insertAt: number,
): Table | null {
	return (
		sheet.tables.find(
			(candidate) => candidate.id !== table.id && candidate.ref.start.row >= insertAt,
		) ?? null
	)
}

function tableRangesOverlap(a: Table['ref'], b: Table['ref']): boolean {
	return (
		a.start.row <= b.end.row &&
		a.end.row >= b.start.row &&
		a.start.col <= b.end.col &&
		a.end.col >= b.start.col
	)
}

function tableRangeOverlapError(
	operation: 'append' | 'create' | 'resize',
	tableName: string,
	ref: string,
	overlappingTable: Table,
): ReturnType<typeof ascendError> {
	const action =
		operation === 'append'
			? `append rows to table "${tableName}" through ${ref}`
			: `${operation} table "${tableName}" at ${ref}`
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${action} because it overlaps table "${overlappingTable.name}" at ${rangeToA1(overlappingTable.ref)}`,
		{
			refs: [ref, rangeToA1(overlappingTable.ref)],
			suggestedFix:
				'Choose a non-overlapping range or resize/delete the existing table before changing table ownership.',
		},
	)
}

function tableAppendTotalsShiftError(
	table: Table,
	shiftedTable: Table,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot append rows to table "${table.name}" because inserting before its totals row would shift table "${shiftedTable.name}" at ${rangeToA1(shiftedTable.ref)}`,
		{
			refs: [rangeToA1(table.ref), rangeToA1(shiftedTable.ref)],
			suggestedFix:
				'Move the following table or remove the totals row before appending rows that require worksheet row insertion.',
		},
	)
}

function tableResizeDroppedColumnReferenceError(
	blocker: DeletedTableColumnReference,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot resize table because ${blocker.sourceRef} ${blocker.sourceKind} references table column ${blocker.tableName}[${blocker.columnName}]`,
		{
			refs: [blocker.sourceRef],
			suggestedFix:
				'Rewrite or remove structured references to the table field before resizing the table to exclude it.',
		},
	)
}

function buildResizedTableColumns(
	sheet: Sheet,
	table: Table,
	ref: Table['ref'],
): readonly TableColumn[] {
	const width = ref.end.col - ref.start.col + 1
	const inferred = buildTableColumns(sheet, ref, width, table.hasHeaders)
	const shiftedByOffset = new Map<number, TableColumn>()
	for (const [index, column] of table.columns.entries()) {
		const oldColumn = table.ref.start.col + index
		if (oldColumn < ref.start.col || oldColumn > ref.end.col) continue
		shiftedByOffset.set(oldColumn - ref.start.col, { ...column })
	}
	const hasExplicitIds = table.columns.some((column) => column.id !== undefined)
	let nextId = Math.max(0, ...table.columns.map((column) => column.id ?? 0)) + 1
	const usedNames = new Set<string>()
	const columns: TableColumn[] = []
	for (let offset = 0; offset < width; offset++) {
		const shifted = shiftedByOffset.get(offset)
		if (shifted) {
			columns.push(shifted)
			usedNames.add(shifted.name.toLowerCase())
			continue
		}
		const name = nextUniqueTableColumnName(
			inferred[offset]?.name ?? `Column${offset + 1}`,
			usedNames,
		)
		columns.push({
			name,
			...(hasExplicitIds ? { id: nextId++ } : {}),
		})
	}
	return columns
}

function rangeToA1(range: Table['ref']): string {
	return `${toA1(range.start)}:${toA1(range.end)}`
}

function nextUniqueTableColumnName(base: string, usedNames: Set<string>): string {
	let candidate = base
	let suffix = 2
	while (usedNames.has(candidate.toLowerCase())) {
		candidate = `${base}_${suffix}`
		suffix++
	}
	usedNames.add(candidate.toLowerCase())
	return candidate
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

function resizeTableAutoFilter(
	autoFilter: AutoFilter,
	oldRef: RangeRef,
	nextRef: RangeRef,
): AutoFilter {
	const sortState = resizeTableSortState(autoFilter.sortState, oldRef, nextRef)
	const { sortState: _sortState, ref: _ref, columns: _columns, ...rest } = autoFilter
	return {
		...rest,
		ref: rangeToA1(nextRef),
		columns: remapTableFilterColumns(autoFilter.columns, oldRef, nextRef),
		...(sortState ? { sortState } : {}),
	}
}

function remapTableFilterColumns(
	columns: AutoFilter['columns'],
	oldRef: RangeRef,
	nextRef: RangeRef,
): AutoFilter['columns'] {
	return columns
		.map((column) => {
			const absoluteCol = oldRef.start.col + column.colId
			if (absoluteCol < nextRef.start.col || absoluteCol > nextRef.end.col) return null
			return { ...column, colId: absoluteCol - nextRef.start.col }
		})
		.filter((column): column is AutoFilter['columns'][number] => column !== null)
		.sort((left, right) => left.colId - right.colId)
}

function resizeTableSortState(
	sortState: SortState | null | undefined,
	oldRef: RangeRef,
	nextRef: RangeRef,
): SortState | undefined {
	if (!sortState) return undefined
	const conditions = sortState.conditions
		.map((condition) => resizeTableSortCondition(condition, oldRef, nextRef))
		.filter((condition): condition is SortCondition => condition !== null)
	if (conditions.length === 0) return undefined
	return { ...sortState, ref: rangeToA1(nextRef), conditions }
}

function resizeTableSortCondition(
	condition: SortCondition,
	oldRef: RangeRef,
	nextRef: RangeRef,
): SortCondition | null {
	try {
		const conditionRef = parseRange(condition.ref)
		if (!tableRangesOverlap(conditionRef, nextRef)) return null
		const resizedRef = resizeTableSortConditionRef(conditionRef, oldRef, nextRef)
		return { ...condition, ref: rangeToA1(resizedRef) }
	} catch {
		return condition
	}
}

function resizeTableSortConditionRef(
	conditionRef: RangeRef,
	oldRef: RangeRef,
	nextRef: RangeRef,
): RangeRef {
	if (conditionRef.end.row !== oldRef.end.row) return conditionRef
	const endRow = nextRef.end.row
	if (endRow < conditionRef.start.row) return conditionRef
	return {
		start: conditionRef.start,
		end: { ...conditionRef.end, row: endRow },
	}
}
