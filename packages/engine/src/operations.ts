import type { Cell, CellStyle, RangeRef, Sheet, StyleId, Workbook } from '@ascend/core'
import { createTableId, parseA1, parseRange, toA1 } from '@ascend/core'
import type { FormulaCellRef, FormulaNode } from '@ascend/formulas'
import { cachedParseFormula, dateToSerial, printFormula, rewriteRefs } from '@ascend/formulas'
import type { CellValue, InputValue, Operation, Result } from '@ascend/schema'
import { ascendError, booleanValue, EMPTY, err, numberValue, ok, stringValue } from '@ascend/schema'
import { invalidateWorkbookAnalysis } from './analysis.ts'
import { invalidateSheetIndexCache } from './evaluator.ts'
import {
	rewriteDefinedNameFormulasForShift,
	rewriteSheetMetadataFormulasForRename,
	rewriteSheetNameInDefinedNames,
	rewriteSheetNameInFormulas,
	rewriteWorkbookFormulasForShift,
} from './structural/formula-rewrite.ts'
import { expandSqrefRows } from './structural/ref-shift.ts'
import { renameHyperlinkLocation, shiftSheetCellMetadata } from './structural/sheet-topology.ts'
import { sortSheetRange } from './structural/sort-range.ts'

export interface PatchResult {
	readonly affectedCells: string[]
	readonly sheetsModified: string[]
	readonly recalcRequired: boolean
}

const DEFAULT_SID = 0 as unknown as StyleId

function inputToCellValue(input: InputValue): CellValue {
	if (input === null) return EMPTY
	if (typeof input === 'number') return numberValue(input)
	if (typeof input === 'string') return stringValue(input)
	if (typeof input === 'boolean') return booleanValue(input)
	if (input instanceof Date) {
		return {
			kind: 'date',
			serial: dateToSerial(input.getFullYear(), input.getMonth() + 1, input.getDate()),
		}
	}
	return EMPTY
}

function getSheet(workbook: Workbook, name: string): Result<Sheet> {
	const sheet = workbook.getSheet(name)
	if (!sheet) {
		return err(ascendError('SHEET_NOT_FOUND', `Sheet "${name}" not found`))
	}
	return ok(sheet)
}

function patch(affected: string[], sheets: string[], recalc = false): PatchResult {
	return {
		affectedCells: affected,
		sheetsModified: [...new Set(sheets)],
		recalcRequired: recalc,
	}
}

function cell(value: CellValue, formula: string | null, styleId: StyleId): Cell {
	return { value, formula, styleId }
}

function cellWithExisting(
	value: CellValue,
	formula: string | null,
	styleId: StyleId,
	existing?: Cell,
): Cell {
	return {
		value,
		formula,
		styleId,
		...(formula !== null && existing?.formulaInfo ? { formulaInfo: existing.formulaInfo } : {}),
	}
}

function safeParseRange(range: string): Result<RangeRef> {
	try {
		return ok(parseRange(range))
	} catch {
		return err(ascendError('INVALID_RANGE', `Invalid range: ${range}`))
	}
}

// --- Operation handlers ---

function handleSetCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCells' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const affected: string[] = []
	for (const update of op.updates) {
		const ref = parseA1(update.ref)
		const value = inputToCellValue(update.value)
		const existing = sheet.cells.get(ref.row, ref.col)
		sheet.cells.set(
			ref.row,
			ref.col,
			cellWithExisting(
				value,
				existing?.formula ?? null,
				existing?.styleId ?? DEFAULT_SID,
				existing,
			),
		)
		affected.push(update.ref)
	}

	return ok(patch(affected, [op.sheet], true))
}

function handleSetFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setFormula' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const ref = parseA1(op.ref)
	const existing = sheet.cells.get(ref.row, ref.col)
	sheet.cells.set(
		ref.row,
		ref.col,
		cellWithExisting(
			existing?.value ?? EMPTY,
			normalizeFormulaInput(op.formula),
			existing?.styleId ?? DEFAULT_SID,
			undefined,
		),
	)

	return ok(patch([op.ref], [op.sheet], true))
}

function handleFillFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'fillFormula' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const baseFormula = normalizeFormulaInput(op.formula)
	const parsed = cachedParseFormula(baseFormula)
	if (!parsed.ok) {
		return err(ascendError('VALIDATION_ERROR', `Invalid formula: ${op.formula}`))
	}
	const range = rangeResult.value
	const anchor = range.start
	const affected: string[] = []

	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			const translated = translateFormula(parsed.value, row - anchor.row, col - anchor.col)
			const existing = sheet.cells.get(row, col)
			sheet.cells.set(
				row,
				col,
				cellWithExisting(existing?.value ?? EMPTY, translated, existing?.styleId ?? DEFAULT_SID),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, [op.sheet], true))
}

function handleInsertRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertRows' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	sheet.cells.insertRows(op.at, op.count)

	shiftMerges(sheet.merges, 'row', op.at, op.count)
	shiftSheetCellMetadata(sheet, 'row', op.at, op.count)
	clearFormulaMetadata(workbook)
	rewriteWorkbookFormulasForShift(workbook, op.sheet, 'row', op.at, op.count)
	rewriteDefinedNameFormulasForShift(workbook, op.sheet, 'row', op.at, op.count)

	return ok(patch([], [op.sheet], true))
}

function handleDeleteRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteRows' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	sheet.cells.deleteRows(op.at, op.count)

	shiftMerges(sheet.merges, 'row', op.at, -op.count)
	shiftSheetCellMetadata(sheet, 'row', op.at, -op.count)
	clearFormulaMetadata(workbook)
	rewriteWorkbookFormulasForShift(workbook, op.sheet, 'row', op.at, -op.count)
	rewriteDefinedNameFormulasForShift(workbook, op.sheet, 'row', op.at, -op.count)

	return ok(patch([], [op.sheet], true))
}

function handleInsertCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertCols' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	sheet.cells.insertCols(op.at, op.count)

	shiftMerges(sheet.merges, 'col', op.at, op.count)
	shiftSheetCellMetadata(sheet, 'col', op.at, op.count)
	clearFormulaMetadata(workbook)
	rewriteWorkbookFormulasForShift(workbook, op.sheet, 'col', op.at, op.count)
	rewriteDefinedNameFormulasForShift(workbook, op.sheet, 'col', op.at, op.count)

	return ok(patch([], [op.sheet], true))
}

function handleDeleteCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteCols' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	sheet.cells.deleteCols(op.at, op.count)

	shiftMerges(sheet.merges, 'col', op.at, -op.count)
	shiftSheetCellMetadata(sheet, 'col', op.at, -op.count)
	clearFormulaMetadata(workbook)
	rewriteWorkbookFormulasForShift(workbook, op.sheet, 'col', op.at, -op.count)
	rewriteDefinedNameFormulasForShift(workbook, op.sheet, 'col', op.at, -op.count)

	return ok(patch([], [op.sheet], true))
}

function shiftMerges(merges: RangeRef[], axis: 'row' | 'col', at: number, delta: number): void {
	const updated: RangeRef[] = []
	for (const m of merges) {
		const s = axis === 'row' ? m.start.row : m.start.col
		const e = axis === 'row' ? m.end.row : m.end.col

		if (delta < 0) {
			const deleteEnd = at - delta
			if (s >= at && e < deleteEnd) continue
		}

		const shift = (v: number): number => {
			if (delta > 0) return v >= at ? v + delta : v
			const deleteEnd = at - delta
			if (v >= deleteEnd) return v + delta
			if (v >= at) return at
			return v
		}

		if (axis === 'row') {
			updated.push({
				start: { row: shift(m.start.row), col: m.start.col },
				end: { row: shift(m.end.row), col: m.end.col },
			})
		} else {
			updated.push({
				start: { row: m.start.row, col: shift(m.start.col) },
				end: { row: m.end.row, col: shift(m.end.col) },
			})
		}
	}
	merges.length = 0
	merges.push(...updated)
}

function translateFormula(node: FormulaNode, rowDelta: number, colDelta: number): string {
	const rewritten = rewriteRefs(node, (ref: FormulaCellRef) => ({
		...ref,
		row: ref.rowAbsolute ? ref.row : ref.row + rowDelta,
		col: ref.colAbsolute ? ref.col : ref.col + colDelta,
	}))
	return printFormula(rewritten)
}

function normalizeFormulaInput(formula: string): string {
	return formula.startsWith('=') ? formula.slice(1) : formula
}

function handleAddSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'addSheet' }>,
): Result<PatchResult> {
	if (workbook.getSheet(op.name)) {
		return err(ascendError('NAME_CONFLICT', `Sheet "${op.name}" already exists`))
	}
	const sheet = workbook.addSheet(op.name)
	if (op.position !== undefined) {
		const idx = workbook.sheets.indexOf(sheet)
		workbook.sheets.splice(idx, 1)
		workbook.sheets.splice(op.position, 0, sheet)
		workbook.invalidateSheetCache()
	}
	return ok(patch([], [op.name]))
}

function handleCreateTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'createTable' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	const rangeResult = safeParseRange(op.ref)
	if (!rangeResult.ok) return rangeResult
	if (sheet.tables.some((table) => table.name === op.name)) {
		return err(ascendError('NAME_CONFLICT', `Table "${op.name}" already exists`))
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

function handleAppendRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'appendRows' }>,
): Result<PatchResult> {
	const located = findTable(workbook, op.table)
	if (!located) {
		return err(ascendError('NAME_NOT_FOUND', `Table "${op.table}" not found`))
	}
	const { table, sheet } = located
	if (table.hasTotals) {
		return err(
			ascendError('VALIDATION_ERROR', 'appendRows does not support tables with totals rows yet'),
		)
	}
	if (op.rows.length === 0) return ok(patch([], [sheet.name], false))

	const width = table.columns.length
	const affected: string[] = []
	let nextRow = table.ref.end.row + 1
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
						inputToCellValue(provided),
						existing?.formula ?? null,
						existing?.styleId ?? DEFAULT_SID,
						existing,
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
				end: { row: table.ref.end.row + rowDelta, col: table.ref.end.col },
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

function handleDeleteSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteSheet' }>,
): Result<PatchResult> {
	const targetSheet = workbook.getSheet(op.sheet)
	if (!targetSheet) {
		return err(ascendError('SHEET_NOT_FOUND', `Sheet "${op.sheet}" not found`))
	}
	const removedPivotNames = workbook.pivotTables
		.filter((entry) => entry.sheetName === op.sheet)
		.map((entry) => entry.name)
		.filter((name): name is string => Boolean(name))
	workbook.removeSheet(op.sheet)
	removeSheetScopedDefinedNames(workbook, targetSheet.id)
	removeWorkbookMetadataForDeletedSheet(workbook, op.sheet, removedPivotNames)
	return ok(patch([], [op.sheet]))
}

function handleRenameSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameSheet' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	if (workbook.getSheet(op.newName)) {
		return err(ascendError('NAME_CONFLICT', `Sheet "${op.newName}" already exists`))
	}

	const oldName = sheet.name
	sheet.name = op.newName
	workbook.invalidateSheetCache()
	clearFormulaMetadata(workbook)
	rewriteSheetNameInFormulas(workbook, oldName, op.newName)
	rewriteSheetNameInDefinedNames(workbook, oldName, op.newName)
	for (const workbookSheet of workbook.sheets) {
		rewriteSheetMetadataFormulasForRename(workbookSheet, oldName, op.newName)
		for (const [ref, hyperlink] of workbookSheet.hyperlinks) {
			const location = renameHyperlinkLocation(hyperlink.location, oldName, op.newName)
			if (location === hyperlink.location) continue
			workbookSheet.hyperlinks.set(ref, {
				...hyperlink,
				...(location !== undefined ? { location } : {}),
			})
		}
	}

	return ok(patch([], [op.newName]))
}

function handleMoveSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'moveSheet' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const idx = workbook.sheets.indexOf(sheet)
	workbook.sheets.splice(idx, 1)
	workbook.sheets.splice(op.position, 0, sheet)
	workbook.invalidateSheetCache()

	return ok(patch([], [op.sheet]))
}

function removeSheetScopedDefinedNames(workbook: Workbook, sheetId: string): void {
	const scopedEntries = workbook.definedNames
		.list()
		.filter((entry) => entry.scope.kind === 'sheet' && entry.scope.sheetId === sheetId)
	for (const entry of scopedEntries) {
		workbook.definedNames.delete(entry.name, entry.scope)
	}
}

function removeWorkbookMetadataForDeletedSheet(
	workbook: Workbook,
	sheetName: string,
	removedPivotNames: readonly string[],
): void {
	for (let index = workbook.pivotTables.length - 1; index >= 0; index--) {
		if (workbook.pivotTables[index]?.sheetName === sheetName) {
			workbook.pivotTables.splice(index, 1)
		}
	}
	for (let index = workbook.slicers.length - 1; index >= 0; index--) {
		const slicer = workbook.slicers[index]
		if (!slicer) continue
		if (
			removedPivotNames.some(
				(pivotName) => slicer.name === pivotName || slicer.cacheName === pivotName,
			)
		) {
			workbook.slicers.splice(index, 1)
		}
	}
	for (let index = workbook.slicerCaches.length - 1; index >= 0; index--) {
		const cache = workbook.slicerCaches[index]
		if (!cache) continue
		const remainingPivotNames = cache.pivotTableNames.filter(
			(name) => !removedPivotNames.includes(name),
		)
		if (remainingPivotNames.length === 0 && cache.pivotTableNames.length > 0) {
			workbook.slicerCaches.splice(index, 1)
		} else if (remainingPivotNames.length !== cache.pivotTableNames.length) {
			workbook.slicerCaches[index] = {
				...cache,
				pivotTableNames: remainingPivotNames,
			}
		}
	}
}

function buildTableColumns(
	sheet: Sheet,
	ref: RangeRef,
	width: number,
	hasHeaders: boolean,
): Array<{ name: string }> {
	const usedNames = new Set<string>()
	const columns: Array<{ name: string }> = []
	for (let colOffset = 0; colOffset < width; colOffset++) {
		let name = `Column${colOffset + 1}`
		if (hasHeaders) {
			const cellValue = sheet.cells.get(ref.start.row, ref.start.col + colOffset)?.value
			if (cellValue?.kind === 'string' && cellValue.value.trim() !== '') {
				name = cellValue.value.trim()
			}
		}
		let candidate = name
		let suffix = 2
		while (usedNames.has(candidate.toLowerCase())) {
			candidate = `${name}_${suffix}`
			suffix++
		}
		usedNames.add(candidate.toLowerCase())
		columns.push({ name: candidate })
	}
	return columns
}

function findTable(
	workbook: Workbook,
	name: string,
): { table: Sheet['tables'][number]; sheet: Sheet } | null {
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.name === name) return { table, sheet }
		}
	}
	return null
}

function handleMergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'mergeCells' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult

	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult

	sheetResult.value.merges.push(rangeResult.value)
	return ok(patch([], [op.sheet]))
}

function handleUnmergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'unmergeCells' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult

	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const r = rangeResult.value

	const merges = sheetResult.value.merges
	const idx = merges.findIndex(
		(m) =>
			m.start.row === r.start.row &&
			m.start.col === r.start.col &&
			m.end.row === r.end.row &&
			m.end.col === r.end.col,
	)
	if (idx >= 0) merges.splice(idx, 1)

	return ok(patch([], [op.sheet]))
}

function handleSetComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setComment' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result

	const comment = op.author !== undefined ? { text: op.text, author: op.author } : { text: op.text }
	result.value.comments.set(op.ref, comment)

	return ok(patch([op.ref], [op.sheet]))
}

function handleSetDefinedName(
	_workbook: Workbook,
	op: Extract<Operation, { op: 'setDefinedName' }>,
): Result<PatchResult> {
	if (op.scope) {
		const sheet = _workbook.getSheet(op.scope)
		if (!sheet) {
			return err(ascendError('SHEET_NOT_FOUND', `Sheet "${op.scope}" not found`))
		}
		_workbook.definedNames.set(op.name, op.ref, { kind: 'sheet', sheetId: sheet.id })
	} else {
		_workbook.definedNames.set(op.name, op.ref)
	}
	return ok(patch([], []))
}

function handleDeleteDefinedName(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDefinedName' }>,
): Result<PatchResult> {
	if (op.scope) {
		const sheet = workbook.getSheet(op.scope)
		if (!sheet) {
			return err(ascendError('SHEET_NOT_FOUND', `Sheet "${op.scope}" not found`))
		}
		if (!workbook.definedNames.delete(op.name, { kind: 'sheet', sheetId: sheet.id })) {
			return err(
				ascendError('NAME_NOT_FOUND', `Defined name "${op.name}" not found in scope "${op.scope}"`),
			)
		}
		return ok(patch([], []))
	}
	if (!workbook.definedNames.has(op.name)) {
		return err(ascendError('NAME_NOT_FOUND', `Defined name "${op.name}" not found`))
	}
	workbook.definedNames.delete(op.name)
	return ok(patch([], []))
}

function handleClearRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'clearRange' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value

	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const range = rangeResult.value

	const affected: string[] = []
	for (let r = range.start.row; r <= range.end.row; r++) {
		for (let c = range.start.col; c <= range.end.col; c++) {
			const existing = sheet.cells.get(r, c)
			if (!existing) continue

			const ref = toA1({ row: r, col: c })
			affected.push(ref)

			switch (op.what) {
				case 'all':
					sheet.cells.delete(r, c)
					break
				case 'values':
					sheet.cells.set(
						r,
						c,
						cellWithExisting(EMPTY, existing.formula, existing.styleId, existing),
					)
					break
				case 'formulas':
					sheet.cells.set(r, c, cell(existing.value, null, existing.styleId))
					break
				case 'styles':
					sheet.cells.set(r, c, cell(existing.value, existing.formula, DEFAULT_SID))
					break
			}
		}
	}

	return ok(patch(affected, [op.sheet], op.what !== 'styles'))
}

function handleFreezePane(
	workbook: Workbook,
	op: Extract<Operation, { op: 'freezePane' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.frozenRows = op.row
	result.value.frozenCols = op.col
	return ok(patch([], [op.sheet]))
}

function handleSetColWidth(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setColWidth' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.colWidths.set(op.col, op.width)
	return ok(patch([], [op.sheet]))
}

function handleSetRowHeight(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRowHeight' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.rowHeights.set(op.row, op.height)
	return ok(patch([], [op.sheet]))
}

function handleSortRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'sortRange' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const range = rangeResult.value
	clearFormulaMetadataForSheet(sheet)
	const sorted = sortSheetRange(workbook, sheet, range, op.by)
	if (!sorted.ok) return sorted
	return ok(patch([], [op.sheet], sorted.value))
}

function clearFormulaMetadataForSheet(sheet: Workbook['sheets'][number]): void {
	for (const [row, col, existing] of sheet.cells.iterate()) {
		if (!existing.formulaInfo) continue
		sheet.cells.set(row, col, {
			value: existing.value,
			formula: existing.formula,
			styleId: existing.styleId,
		})
	}
}

function clearFormulaMetadata(workbook: Workbook): void {
	for (const sheet of workbook.sheets) {
		clearFormulaMetadataForSheet(sheet)
	}
}

function handleSetHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setHyperlink' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.hyperlinks.set(op.ref, {
		target: op.url,
		...(op.display ? { display: op.display } : {}),
	})
	return ok(patch([op.ref], [op.sheet]))
}

function mergeStyleInput(current: CellStyle, input: CellStyle): CellStyle {
	return {
		...current,
		...(input.font && { font: { ...current.font, ...input.font } }),
		...(input.fill && { fill: { ...current.fill, ...input.fill } }),
		...(input.border && { border: { ...current.border, ...input.border } }),
		...(input.alignment && { alignment: { ...current.alignment, ...input.alignment } }),
		...(input.numberFormat !== undefined && { numberFormat: input.numberFormat }),
	}
}

function handleSetStyle(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setStyle' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const range = rangeResult.value

	const input = op.style as unknown as CellStyle
	const affected: string[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			const existing = sheet.cells.get(row, col)
			const currentStyle = workbook.styles.get(existing?.styleId ?? DEFAULT_SID) ?? {}
			const merged = mergeStyleInput(currentStyle, input)
			const styleId = workbook.styles.register(merged)
			sheet.cells.set(
				row,
				col,
				cellWithExisting(existing?.value ?? EMPTY, existing?.formula ?? null, styleId, existing),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, [op.sheet]))
}

function handleSetNumberFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setNumberFormat' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult

	const range = rangeResult.value
	const affected: string[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			const existing = sheet.cells.get(row, col)
			const currentStyle = workbook.styles.get(existing?.styleId ?? DEFAULT_SID) ?? {}
			const style: CellStyle = {
				...currentStyle,
				numberFormat: op.format,
			}
			const styleId = workbook.styles.register(style)
			if (
				workbook.preservedStyles &&
				existing?.styleId !== undefined &&
				styleId !== existing.styleId
			) {
				const baseStyleId =
					workbook.preservedStyles.baseStyleIdByStyleId?.[existing.styleId] ?? existing.styleId
				workbook.preservedStyles = {
					...workbook.preservedStyles,
					baseStyleIdByStyleId: {
						...(workbook.preservedStyles.baseStyleIdByStyleId ?? {}),
						[styleId]: baseStyleId,
					},
				}
			}
			sheet.cells.set(
				row,
				col,
				cellWithExisting(existing?.value ?? EMPTY, existing?.formula ?? null, styleId, existing),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, [op.sheet]))
}

function operationAffectsFormulas(op: Operation): boolean {
	switch (op.op) {
		case 'setComment':
		case 'setHyperlink':
		case 'setNumberFormat':
		case 'setStyle':
		case 'freezePane':
		case 'setColWidth':
		case 'setRowHeight':
		case 'mergeCells':
		case 'unmergeCells':
			return false
		case 'setCells':
			return false
		case 'createTable':
			return false
		case 'clearRange':
			return op.what === 'all' || op.what === 'formulas'
		default:
			return true
	}
}

// --- Public API ---

export function applyOperation(workbook: Workbook, op: Operation): Result<PatchResult> {
	if (operationAffectsFormulas(op)) {
		invalidateWorkbookAnalysis(workbook)
	}
	invalidateSheetIndexCache(workbook)
	switch (op.op) {
		case 'setCells':
			return handleSetCells(workbook, op)
		case 'setFormula':
			return handleSetFormula(workbook, op)
		case 'fillFormula':
			return handleFillFormula(workbook, op)
		case 'insertRows':
			return handleInsertRows(workbook, op)
		case 'deleteRows':
			return handleDeleteRows(workbook, op)
		case 'insertCols':
			return handleInsertCols(workbook, op)
		case 'deleteCols':
			return handleDeleteCols(workbook, op)
		case 'addSheet':
			return handleAddSheet(workbook, op)
		case 'createTable':
			return handleCreateTable(workbook, op)
		case 'appendRows':
			return handleAppendRows(workbook, op)
		case 'deleteSheet':
			return handleDeleteSheet(workbook, op)
		case 'renameSheet':
			return handleRenameSheet(workbook, op)
		case 'moveSheet':
			return handleMoveSheet(workbook, op)
		case 'mergeCells':
			return handleMergeCells(workbook, op)
		case 'unmergeCells':
			return handleUnmergeCells(workbook, op)
		case 'setComment':
			return handleSetComment(workbook, op)
		case 'setDefinedName':
			return handleSetDefinedName(workbook, op)
		case 'deleteDefinedName':
			return handleDeleteDefinedName(workbook, op)
		case 'clearRange':
			return handleClearRange(workbook, op)
		case 'freezePane':
			return handleFreezePane(workbook, op)
		case 'setColWidth':
			return handleSetColWidth(workbook, op)
		case 'setRowHeight':
			return handleSetRowHeight(workbook, op)
		case 'sortRange':
			return handleSortRange(workbook, op)
		case 'setHyperlink':
			return handleSetHyperlink(workbook, op)
		case 'setNumberFormat':
			return handleSetNumberFormat(workbook, op)
		case 'setStyle':
			return handleSetStyle(workbook, op)
		default:
			return assertUnreachable(op)
	}
}

export function applyOperations(
	workbook: Workbook,
	ops: readonly Operation[],
): Result<PatchResult> {
	const allAffected: string[] = []
	const allSheets: string[] = []
	let needsRecalc = false

	for (const op of ops) {
		const result = applyOperation(workbook, op)
		if (!result.ok) return result
		allAffected.push(...result.value.affectedCells)
		allSheets.push(...result.value.sheetsModified)
		if (result.value.recalcRequired) needsRecalc = true
	}

	return ok(patch(allAffected, allSheets, needsRecalc))
}

function assertUnreachable(value: never): Result<PatchResult> {
	return err(ascendError('VALIDATION_ERROR', `Unsupported operation: ${JSON.stringify(value)}`))
}
