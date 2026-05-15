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
import { createTableId, parseA1, parseRange, toA1 } from '@ascend/core'
import { normalizeFormulaInput } from '@ascend/formulas'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, EMPTY, err, ok, stringValue, validateExcelTableName } from '@ascend/schema'
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
	collectTableColumnsForDelete,
	type DeletedTableColumnReference,
	findDeletedTableColumnReference,
} from '../structural/table-field-guards.ts'
import { findOverlappingTable, tableRangesOverlap } from '../table-topology.ts'
import type { PatchResult } from './helpers.ts'
import {
	buildTableColumns,
	cellWithExisting,
	collectFormulaBindingGroupRefsForRefs,
	createLegacyArrayFormulaIndex,
	DEFAULT_SID,
	getSheet,
	inputToCellValue,
	legacyArrayFormulaEditError,
	materializeFormulaBindingGroupsForRangeEdit,
	materializeFormulaBindingGroupsForRefs,
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
	const nameError = tableNameError(op.name)
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
	sheet.ensureWritable()
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
	return ok(patch([], [op.sheet], true))
}

export function handleAppendRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'appendRows' }>,
): Result<PatchResult> {
	const located = resolveUniqueTable(workbook, op.table, 'appendRows')
	if (!located.ok) return located
	const { table, sheet } = located.value
	if (op.rows.length === 0) return ok(patch([], [sheet.name], false))

	const width = table.columns.length
	for (const row of op.rows) {
		if (row.length > width) {
			return err(
				ascendError(
					'VALIDATION_ERROR',
					`Appended row has ${row.length} values but table "${table.name}" expects ${width}`,
				),
			)
		}
	}
	const affected = new Set<string>()
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
	const appendStartRow = table.hasTotals ? originalEndRow : originalEndRow + 1
	const appendRange = {
		start: { row: appendStartRow, col: table.ref.start.col },
		end: { row: appendStartRow + op.rows.length - 1, col: table.ref.end.col },
	}
	const legacyArrayImpact = createLegacyArrayFormulaIndex(sheet).findIntersection(appendRange)
	if (legacyArrayImpact) {
		return err(legacyArrayFormulaEditError(legacyArrayImpact.targetRef, legacyArrayImpact.ref))
	}
	if (table.hasTotals) {
		sheet.cells.insertRows(originalEndRow, op.rows.length)
	}
	for (const ref of materializeFormulaBindingGroupsForRangeEdit(workbook, sheet, appendRange)) {
		affected.add(ref)
	}
	let nextRow = appendStartRow
	for (const row of op.rows) {
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
						null,
						existing?.styleId ?? DEFAULT_SID,
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
			affected.add(ref)
		}
		nextRow++
	}

	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		sheet.ensureWritable()
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
	return ok(patch([...affected], [sheet.name], true))
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
	const affected = new Set<string>()
	const sheetsModified = new Set([sheet.name])
	const sortedDataRange = sortRangeDataRange(sheet, range, op.by)
	if (sortedDataRange) {
		for (const ref of materializeFormulaBindingGroupsForRangeEdit(
			workbook,
			sheet,
			sortedDataRange,
		)) {
			affected.add(ref)
		}
	}
	const sortedAffectedCells = sortRangeAffectedCells(sheet, range, op.by)
	const sorted = sortSheetRange(workbook, sheet, range, op.by)
	if (!sorted.ok) return sorted
	if (sorted.value) {
		for (const ref of sortedAffectedCells) affected.add(ref)
	}
	return ok(patch([...affected], [...sheetsModified], sorted.value))
}

export function handleDeleteTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteTable' }>,
): Result<PatchResult> {
	const located = resolveUniqueTable(workbook, op.table, 'deleteTable')
	if (!located.ok) return located
	const { table, sheet } = located.value
	const deletedColumnBlocker = findDeletedTableColumnReference(
		workbook,
		collectTableColumnsForDelete(table),
		{ skipDeletedTableColumnFormulas: true },
	)
	if (deletedColumnBlocker) {
		return err(tableDeleteReferenceError(deletedColumnBlocker))
	}
	const affected = new Set<string>()
	const sheetsModified = new Set([sheet.name])
	materializeWorkbookFormulaBindings(workbook, sheet.name, affected, sheetsModified)
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	sheet.ensureWritable()
	if (idx >= 0) sheet.tables.splice(idx, 1)
	if (table.queryTable?.partPath) {
		for (let i = workbook.connectionParts.length - 1; i >= 0; i--) {
			const part = workbook.connectionParts[i]
			if (part?.kind === 'queryTable' && part.partPath === table.queryTable.partPath) {
				workbook.connectionParts.splice(i, 1)
			}
		}
	}
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
	return ok(patch([...affected], [...sheetsModified], true))
}

export function handleRenameTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameTable' }>,
): Result<PatchResult> {
	const located = resolveUniqueTable(workbook, op.table, 'renameTable')
	if (!located.ok) return located
	const { table, sheet } = located.value
	const nameError = tableNameError(op.newName)
	if (nameError) return err(nameError)
	if (findTableNameCollision(workbook, op.newName, table.id)) {
		return err(
			ascendError('NAME_CONFLICT', `Table "${op.newName}" already exists`, {
				suggestedFix:
					'Choose a workbook-unique table name so structured references remain unambiguous.',
			}),
		)
	}
	const affected = new Set<string>()
	const sheetsModified = new Set([sheet.name])
	materializeWorkbookFormulaBindings(workbook, sheet.name, affected, sheetsModified)
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	if (idx >= 0) {
		sheet.ensureWritable()
		sheet.tables[idx] = {
			...table,
			name: op.newName,
			...(table.nameAttribute !== undefined && table.nameAttribute !== null
				? { nameAttribute: op.newName }
				: {}),
		}
	}
	for (const rewritten of rewriteTableNameInFormulas(workbook, table.name, op.newName)) {
		affected.add(
			rewritten.sheetName === sheet.name
				? rewritten.ref
				: `${rewritten.sheetName}!${rewritten.ref}`,
		)
		const rewrittenSheet = workbook.getSheet(rewritten.sheetName)
		if (rewrittenSheet) {
			for (const groupRef of collectFormulaBindingGroupRefsForRefs(workbook, rewrittenSheet, [
				parseA1(rewritten.ref),
			])) {
				affected.add(
					rewritten.sheetName === sheet.name ? groupRef : `${rewritten.sheetName}!${groupRef}`,
				)
			}
		}
		sheetsModified.add(rewritten.sheetName)
	}
	rewriteTableNameInDefinedNames(workbook, table.name, op.newName)
	return ok(patch([...affected], [...sheetsModified], true))
}

export function handleResizeTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'resizeTable' }>,
): Result<PatchResult> {
	const located = resolveUniqueTable(workbook, op.table, 'resizeTable')
	if (!located.ok) return located
	const { table, sheet } = located.value
	const rangeResult = safeParseRange(op.ref)
	if (!rangeResult.ok) return rangeResult
	const ref = rangeResult.value
	const overlappingTable = findOverlappingTable(sheet, ref, table.id)
	if (overlappingTable) {
		return err(tableRangeOverlapError('resize', table.name, op.ref, overlappingTable))
	}
	if (table.queryTable && tableColumnsChanged(table.ref, ref)) {
		return err(queryTableColumnResizeError(table, op.ref))
	}
	const droppedColumnBlocker = findDeletedTableColumnReference(
		workbook,
		collectDroppedTableColumnsForResize(table, ref),
		{ skipDeletedTableColumnFormulas: true },
	)
	if (droppedColumnBlocker) {
		return err(tableResizeDroppedColumnReferenceError(droppedColumnBlocker))
	}
	const affected = new Set<string>()
	const sheetsModified = new Set([sheet.name])
	materializeWorkbookFormulaBindings(workbook, sheet.name, affected, sheetsModified)
	const columns = buildResizedTableColumns(sheet, table, ref)
	const autoFilter = table.autoFilter
		? resizeTableAutoFilter(table.autoFilter, table.ref, ref)
		: undefined
	const sortState = resizeTableSortState(table.sortState, table.ref, ref)
	const idx = sheet.tables.findIndex((t) => t.id === table.id)
	if (idx >= 0) {
		sheet.ensureWritable()
		const { autoFilter: _autoFilter, sortState: _sortState, ...tableWithoutFilterState } = table
		sheet.tables[idx] = {
			...tableWithoutFilterState,
			ref,
			columns,
			...(autoFilter ? { autoFilter } : {}),
			...(sortState ? { sortState } : {}),
		}
	}
	return ok(patch([...affected], [...sheetsModified], true))
}

function tableColumnsChanged(before: RangeRef, after: RangeRef): boolean {
	return before.start.col !== after.start.col || before.end.col !== after.end.col
}

function resolveUniqueTable(
	workbook: Workbook,
	name: string,
	operation: Extract<
		Operation['op'],
		| 'appendRows'
		| 'deleteTable'
		| 'renameTable'
		| 'resizeTable'
		| 'setTableColumn'
		| 'setTableStyle'
	>,
): Result<{ readonly table: Table; readonly sheet: Sheet }> {
	const lowerName = name.toLowerCase()
	const matches: { table: Table; sheet: Sheet }[] = []
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.name.toLowerCase() === lowerName) matches.push({ table, sheet })
		}
	}
	if (matches.length === 0) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${name}" not found`))
	}
	if (matches.length > 1) {
		return err(duplicateTableNameOperationError(name, operation, matches))
	}
	const match = matches[0]
	if (!match) return err(ascendError('NAME_NOT_FOUND', `Table "${name}" not found`))
	return ok(match)
}

function duplicateTableNameOperationError(
	tableName: string,
	operation: string,
	matches: readonly { readonly table: Table; readonly sheet: Sheet }[],
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${operation} table "${tableName}" because ${matches.length} table parts use that name, making structured references ambiguous`,
		{
			refs: matches.map(({ table }) => rangeToA1(table.ref)),
			suggestedFix:
				'Repair duplicate imported table names first so the target table and structured-reference rewrites are unambiguous.',
			details: {
				kind: 'duplicate-table-name-operation',
				operation,
				tableName,
				matches: matches.map(({ table, sheet }) => ({
					sheetName: sheet.name,
					ref: rangeToA1(table.ref),
					...(table.partPath ? { partPath: table.partPath } : {}),
				})),
			},
		},
	)
}

function tableNameError(name: string): ReturnType<typeof ascendError> | null {
	const validation = validateExcelTableName(name)
	if (!validation) return null
	return ascendError('VALIDATION_ERROR', validation.message, {
		suggestedFix: validation.suggestedFix,
	})
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
			details: {
				kind: 'overlapping-table-ranges',
				operation,
				tableName,
				ref,
				overlappingTable: {
					tableName: overlappingTable.name,
					ref: rangeToA1(overlappingTable.ref),
					...(overlappingTable.partPath ? { partPath: overlappingTable.partPath } : {}),
				},
			},
		},
	)
}

function queryTableColumnResizeError(table: Table, ref: string): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot resize queryTable-backed table "${table.name}" to ${ref} because changing columns would leave queryTable field bindings ambiguous`,
		{
			refs: [rangeToA1(table.ref), ref],
			details: {
				kind: 'query-table-column-resize',
				tableName: table.name,
				currentRef: rangeToA1(table.ref),
				requestedRef: ref,
				queryTablePartPath: table.queryTable?.partPath,
			},
			suggestedFix:
				'Resize only the row span, or remove and rebuild the queryTable sidecar with matching queryTableField bindings before changing table columns.',
		},
	)
}

function queryTableColumnRenameError(
	table: Table,
	column: Table['columns'][number],
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot rename queryTable-backed column "${column.name}" in table "${table.name}" because the queryTable field binding cannot be rewritten safely`,
		{
			refs: [rangeToA1(table.ref), table.queryTable?.partPath ?? table.name],
			details: {
				kind: 'query-table-column-rename',
				tableName: table.name,
				columnName: column.name,
				queryTableFieldId: column.queryTableFieldId,
				queryTablePartPath: table.queryTable?.partPath,
			},
			suggestedFix:
				'Edit calculated-column or totals metadata only, or rebuild the queryTable sidecar with matching field bindings before renaming query-backed table columns.',
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
			details: {
				kind: 'table-totals-append-would-shift-table',
				tableName: table.name,
				ref: rangeToA1(table.ref),
				shiftedTable: {
					tableName: shiftedTable.name,
					ref: rangeToA1(shiftedTable.ref),
					...(shiftedTable.partPath ? { partPath: shiftedTable.partPath } : {}),
				},
			},
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

function tableDeleteReferenceError(
	blocker: DeletedTableColumnReference,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot delete table because ${blocker.sourceRef} ${blocker.sourceKind} references table column ${blocker.tableName}[${blocker.columnName}]`,
		{
			refs: [blocker.sourceRef],
			suggestedFix:
				'Rewrite or remove structured references to the table before deleting the table metadata.',
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

function sortRangeAffectedCells(
	sheet: Sheet,
	range: RangeRef,
	specs: Extract<Operation, { op: 'sortRange' }>['by'],
): string[] {
	const dataRange = sortRangeDataRange(sheet, range, specs)
	const refs: string[] = []
	if (!dataRange) return refs
	for (let row = dataRange.start.row; row <= dataRange.end.row; row++) {
		for (let col = dataRange.start.col; col <= dataRange.end.col; col++) {
			refs.push(toA1({ row, col }))
		}
	}
	return refs
}

function sortRangeDataRange(
	sheet: Sheet,
	range: RangeRef,
	specs: Extract<Operation, { op: 'sortRange' }>['by'],
): RangeRef | null {
	const startRow = sortRangeHasHeaderRow(sheet, range, specs)
		? range.start.row + 1
		: range.start.row
	if (startRow > range.end.row) return null
	return {
		start: { row: startRow, col: range.start.col },
		end: { ...range.end },
	}
}

function sortRangeHasHeaderRow(
	sheet: Sheet,
	range: RangeRef,
	specs: Extract<Operation, { op: 'sortRange' }>['by'],
): boolean {
	const headerMap = new Set<string>()
	for (let col = range.start.col; col <= range.end.col; col++) {
		const value = sheet.cells.readValue(range.start.row, col)
		if (value?.kind === 'string' && value.value.trim() !== '') {
			headerMap.add(value.value.trim().toLowerCase())
		}
	}
	return specs.some(
		(spec) => typeof spec.column === 'string' && headerMap.has(spec.column.trim().toLowerCase()),
	)
}

function materializeWorkbookFormulaBindings(
	workbook: Workbook,
	primarySheetName: string,
	affected: Set<string>,
	sheetsModified: Set<string>,
): void {
	for (const formulaSheet of workbook.sheets) {
		materializeSheetFormulaBindings(
			workbook,
			formulaSheet,
			primarySheetName,
			affected,
			sheetsModified,
		)
	}
}

function materializeSheetFormulaBindings(
	workbook: Workbook,
	formulaSheet: Sheet,
	primarySheetName: string,
	affected: Set<string>,
	sheetsModified: Set<string>,
): void {
	if (formulaSheet.cells.formulaInfoCellCount() === 0) return
	const refs: Array<{ readonly row: number; readonly col: number }> = []
	for (const [row, col, cell] of formulaSheet.cells.iterate()) {
		if (cell.formulaInfo?.kind === 'shared') refs.push({ row, col })
	}
	for (const ref of materializeFormulaBindingGroupsForRefs(workbook, formulaSheet, refs)) {
		affected.add(formulaSheet.name === primarySheetName ? ref : `${formulaSheet.name}!${ref}`)
		sheetsModified.add(formulaSheet.name)
	}
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
	const located = resolveUniqueTable(workbook, op.table, 'setTableColumn')
	if (!located.ok) return located
	const { table, sheet } = located.value
	const inputError = validateTableColumnInput(op)
	if (inputError) return err(inputError)
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
	if (op.newName !== undefined && table.queryTable) {
		return err(queryTableColumnRenameError(table, column))
	}
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
	const formulaTargetRange =
		op.formula === undefined ? null : tableBodyColumnRange(table, columnIndex)
	if (formulaTargetRange) {
		const legacyArrayImpact =
			createLegacyArrayFormulaIndex(sheet).findIntersection(formulaTargetRange)
		if (legacyArrayImpact) {
			return err(legacyArrayFormulaEditError(legacyArrayImpact.targetRef, legacyArrayImpact.ref))
		}
	}
	const affected = new Set<string>()
	const sheetsModified = new Set([sheet.name])
	if (op.newName !== undefined) {
		materializeWorkbookFormulaBindings(workbook, sheet.name, affected, sheetsModified)
	}
	const totalsTarget =
		table.hasTotals &&
		(op.totalsRowFunction !== undefined ||
			op.totalsRowFormula !== undefined ||
			op.totalsRowLabel !== undefined)
			? { row: table.ref.end.row, col: table.ref.start.col + columnIndex }
			: null
	if (totalsTarget) {
		const blocked = createLegacyArrayFormulaIndex(sheet).findCell(
			totalsTarget.row,
			totalsTarget.col,
		)
		if (blocked) {
			return err(legacyArrayFormulaEditError(toA1(totalsTarget), blocked.ref))
		}
		for (const ref of materializeFormulaBindingGroupsForRefs(workbook, sheet, [totalsTarget])) {
			affected.add(ref)
		}
	}
	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		sheet.ensureWritable()
		sheet.tables[tableIndex] = {
			...table,
			columns: table.columns.map((candidate, index) =>
				index === columnIndex ? nextColumn : candidate,
			),
		}
	}
	let recalcRequired = op.formula !== undefined || op.newName !== undefined
	if (op.newName !== undefined) {
		if (table.hasHeaders) {
			const headerCell = sheet.cells.get(table.ref.start.row, table.ref.start.col + columnIndex)
			sheet.cells.set(table.ref.start.row, table.ref.start.col + columnIndex, {
				value: stringValue(op.newName),
				formula: null,
				styleId: headerCell?.styleId ?? DEFAULT_SID,
			})
			affected.add(toA1({ row: table.ref.start.row, col: table.ref.start.col + columnIndex }))
		}
		for (const rewritten of rewriteTableColumnInFormulas(
			workbook,
			table,
			column.name,
			op.newName,
		)) {
			affected.add(
				rewritten.sheetName === sheet.name
					? rewritten.ref
					: `${rewritten.sheetName}!${rewritten.ref}`,
			)
			const rewrittenSheet = workbook.getSheet(rewritten.sheetName)
			if (rewrittenSheet) {
				for (const groupRef of collectFormulaBindingGroupRefsForRefs(workbook, rewrittenSheet, [
					parseA1(rewritten.ref),
				])) {
					affected.add(
						rewritten.sheetName === sheet.name ? groupRef : `${rewritten.sheetName}!${groupRef}`,
					)
				}
			}
			sheetsModified.add(rewritten.sheetName)
		}
		rewriteTableColumnInDefinedNames(workbook, table.name, column.name, op.newName)
	}
	const totalsResult = materializeTotalsRowCell(sheet, table, columnIndex, nextColumn, op)
	if (totalsResult) {
		affected.add(totalsResult.ref)
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
		if (formulaTargetRange) {
			for (const ref of materializeFormulaBindingGroupsForRangeEdit(
				workbook,
				sheet,
				formulaTargetRange,
			)) {
				affected.add(ref)
			}
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
			affected.add(toA1({ row, col }))
		}
	}

	return ok(patch([...affected], [...sheetsModified], recalcRequired))
}

export function handleSetTableStyle(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableStyle' }>,
): Result<PatchResult> {
	const located = resolveUniqueTable(workbook, op.table, 'setTableStyle')
	if (!located.ok) return located
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
	const styleValidation = validateTableStyleInput(op)
	if (styleValidation) return err(styleValidation)

	const { table, sheet } = located.value
	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		sheet.ensureWritable()
		const nextStyle = updateTableStyle(table.tableStyleInfo, op)
		const { tableStyleInfo: _tableStyleInfo, ...tableWithoutStyle } = table
		sheet.tables[tableIndex] = nextStyle
			? { ...tableWithoutStyle, tableStyleInfo: nextStyle }
			: tableWithoutStyle
	}
	return ok(patch([], [sheet.name], false))
}

function validateTableStyleInput(op: Extract<Operation, { op: 'setTableStyle' }>) {
	if (op.styleName !== undefined && op.styleName !== null) {
		if (typeof op.styleName !== 'string' || op.styleName.trim() === '') {
			return ascendError(
				'VALIDATION_ERROR',
				'setTableStyle styleName must be a non-empty string or null',
				{
					suggestedFix: 'Use a table style name such as TableStyleMedium2, or null to clear it.',
				},
			)
		}
	}
	for (const field of [
		'showFirstColumn',
		'showLastColumn',
		'showRowStripes',
		'showColumnStripes',
	] as const) {
		const value = op[field]
		if (value !== undefined && typeof value !== 'boolean') {
			return ascendError('VALIDATION_ERROR', `setTableStyle ${field} must be boolean`, {
				suggestedFix: `Set ${field}=true or ${field}=false.`,
			})
		}
	}
	return null
}

function resolveTableColumnIndex(columns: readonly TableColumn[], column: string | number): number {
	if (typeof column === 'number') {
		return Number.isInteger(column) && column >= 0 && column < columns.length ? column : -1
	}
	if (typeof column !== 'string') return -1
	return columns.findIndex((candidate) => candidate.name.toLowerCase() === column.toLowerCase())
}

function validateTableColumnInput(op: Extract<Operation, { op: 'setTableColumn' }>) {
	if (
		(typeof op.column !== 'string' || op.column.trim() === '') &&
		(typeof op.column !== 'number' || !Number.isInteger(op.column) || op.column < 0)
	) {
		return ascendError(
			'VALIDATION_ERROR',
			'setTableColumn column must be a non-empty string or a zero-based column index',
			{ suggestedFix: 'Use the table column name or a zero-based integer column index.' },
		)
	}
	if (op.newName !== undefined) {
		if (typeof op.newName !== 'string' || op.newName.trim() === '') {
			return ascendError('VALIDATION_ERROR', 'setTableColumn newName must be a non-empty string', {
				suggestedFix: 'Use the target Excel table column header text.',
			})
		}
	}
	for (const field of ['formula', 'totalsRowFormula'] as const) {
		const value = op[field]
		if (value !== undefined && value !== null && typeof value !== 'string') {
			return ascendError('VALIDATION_ERROR', `setTableColumn ${field} must be a string or null`, {
				suggestedFix: `Use a formula string for ${field}, or null to clear it.`,
			})
		}
	}
	for (const field of ['totalsRowFunction', 'totalsRowLabel'] as const) {
		const value = op[field]
		if (value !== undefined && value !== null && typeof value !== 'string') {
			return ascendError('VALIDATION_ERROR', `setTableColumn ${field} must be a string or null`, {
				suggestedFix: `Use a string for ${field}, or null to clear it.`,
			})
		}
	}
	return null
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
		return { ref, recalcRequired: true }
	}
	const subtotalFormula = subtotalFormulaFor(column.totalsRowFunction, table.name, column.name)
	if (subtotalFormula) {
		sheet.cells.set(row, col, {
			value: existing?.value ?? EMPTY,
			formula: subtotalFormula,
			styleId: existing?.styleId ?? DEFAULT_SID,
		})
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

function tableBodyColumnRange(table: Table, columnIndex: number): RangeRef | null {
	const startRow = table.ref.start.row + (table.hasHeaders ? 1 : 0)
	const endRow = table.ref.end.row - (table.hasTotals ? 1 : 0)
	if (startRow > endRow) return null
	const col = table.ref.start.col + columnIndex
	return { start: { row: startRow, col }, end: { row: endRow, col } }
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
