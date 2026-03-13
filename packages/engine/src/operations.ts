import type { Cell, CellStyle, RangeRef, Sheet, StyleId, Workbook } from '@ascend/core'
import { createTableId, parseA1, parseRange, toA1 } from '@ascend/core'
import type { FormulaCellRef, FormulaNode } from '@ascend/formulas'
import {
	cachedParseFormula,
	dateToSerial,
	normalizeFormulaInput,
	printFormula,
	rewriteRefs,
} from '@ascend/formulas'
import type { CellValue, InputValue, Operation, Result } from '@ascend/schema'
import { ascendError, booleanValue, EMPTY, err, numberValue, ok, stringValue } from '@ascend/schema'
import { invalidateWorkbookAnalysis, patchWorkbookAnalysis } from './analysis.ts'
import { type CellKey, cellKey } from './dep-graph.ts'
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

function applyAxisShift(
	workbook: Workbook,
	sheetName: string,
	axis: 'row' | 'col',
	at: number,
	count: number,
	delta: number,
): Result<PatchResult> {
	const result = getSheet(workbook, sheetName)
	if (!result.ok) return result
	const sheet = result.value

	if (axis === 'row') {
		delta > 0 ? sheet.cells.insertRows(at, count) : sheet.cells.deleteRows(at, count)
	} else {
		delta > 0 ? sheet.cells.insertCols(at, count) : sheet.cells.deleteCols(at, count)
	}

	shiftMerges(sheet.merges, axis, at, delta)
	shiftSheetCellMetadata(sheet, axis, at, delta)
	clearFormulaMetadata(workbook)
	rewriteWorkbookFormulasForShift(workbook, sheetName, axis, at, delta)
	rewriteDefinedNameFormulasForShift(workbook, sheetName, axis, at, delta)

	return ok(patch([], [sheetName], true))
}

function handleInsertRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertRows' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'row', op.at, op.count, op.count)
}

function handleDeleteRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteRows' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'row', op.at, op.count, -op.count)
}

function handleInsertCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertCols' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'col', op.at, op.count, op.count)
}

function handleDeleteCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteCols' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'col', op.at, op.count, -op.count)
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

function collectRangeCells(
	sheet: Sheet,
	range: RangeRef,
): Array<{ row: number; col: number; cell: Cell | undefined }> {
	const cells: Array<{ row: number; col: number; cell: Cell | undefined }> = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			cells.push({ row, col, cell: sheet.cells.get(row, col) })
		}
	}
	return cells
}

function handleTransferRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'copyRange' | 'moveRange' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	const sourceResult = safeParseRange(op.source)
	if (!sourceResult.ok) return sourceResult
	const targetStart = parseA1(op.target)
	const source = sourceResult.value
	const rowDelta = targetStart.row - source.start.row
	const colDelta = targetStart.col - source.start.col
	const snapshot = collectRangeCells(sheet, source)
	const affected: string[] = []

	for (const entry of snapshot) {
		const targetRow = entry.row + rowDelta
		const targetCol = entry.col + colDelta
		const existingTarget = sheet.cells.get(targetRow, targetCol)
		if (!entry.cell) {
			sheet.cells.delete(targetRow, targetCol)
			affected.push(toA1({ row: targetRow, col: targetCol }))
			continue
		}

		let formula = entry.cell.formula
		if (formula) {
			const parsed = cachedParseFormula(formula)
			if (parsed.ok) formula = translateFormula(parsed.value, rowDelta, colDelta)
		}

		sheet.cells.set(
			targetRow,
			targetCol,
			cellWithExisting(
				entry.cell.value,
				formula,
				entry.cell.styleId,
				formula !== null ? undefined : existingTarget,
			),
		)
		affected.push(toA1({ row: targetRow, col: targetCol }))
	}

	if (op.op === 'moveRange') {
		for (const entry of snapshot) {
			sheet.cells.delete(entry.row, entry.col)
			affected.push(toA1({ row: entry.row, col: entry.col }))
		}
	}

	return ok(patch(affected, [op.sheet], true))
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

function resolveAffectedCellKeys(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setFormula' | 'fillFormula' }>,
): CellKey[] {
	const sheet = workbook.getSheet(op.sheet)
	if (!sheet) return []
	const sheetIndex = workbook.sheets.indexOf(sheet)
	if (sheetIndex < 0) return []
	if (op.op === 'setFormula') {
		const ref = parseA1(op.ref)
		return [cellKey(sheetIndex, ref.row, ref.col)]
	}
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return []
	const range = rangeResult.value
	const keys: CellKey[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			keys.push(cellKey(sheetIndex, row, col))
		}
	}
	return keys
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
		case 'setRichText':
			return false
		case 'createTable':
			return false
		case 'clearRange':
			return op.what === 'all' || op.what === 'formulas'
		default:
			return true
	}
}

function handleDeleteComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteComment' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.comments.delete(op.ref.toUpperCase())
	return ok(patch([`${op.sheet}!${op.ref}`], [op.sheet]))
}

function handleDeleteHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteHyperlink' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.hyperlinks.delete(op.ref.toUpperCase())
	return ok(patch([`${op.sheet}!${op.ref}`], [op.sheet]))
}

function handleSetDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDataValidation' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const existing = sheet.dataValidations.findIndex((dv) => dv.sqref === op.range)
	const dv: Record<string, unknown> = {
		sqref: op.range,
		type: op.rule.type,
		allowBlank: op.rule.allowBlank ?? true,
		showErrorMessage: op.rule.showErrorMessage ?? true,
	}
	if (op.rule.formula1 !== undefined) dv.formula1 = op.rule.formula1
	if (op.rule.formula2 !== undefined) dv.formula2 = op.rule.formula2
	if (op.rule.operator !== undefined) dv.operator = op.rule.operator
	if (op.rule.errorTitle !== undefined) dv.errorTitle = op.rule.errorTitle
	if (op.rule.errorMessage !== undefined) dv.error = op.rule.errorMessage
	if (op.rule.showInputMessage !== undefined) dv.showInputMessage = op.rule.showInputMessage
	if (op.rule.promptTitle !== undefined) dv.promptTitle = op.rule.promptTitle
	if (op.rule.prompt !== undefined) dv.prompt = op.rule.prompt
	if (existing >= 0) sheet.dataValidations[existing] = dv as never
	else sheet.dataValidations.push(dv as never)
	return ok(patch([], [op.sheet]))
}

function handleDeleteDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDataValidation' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.dataValidations = sheet.dataValidations.filter((dv) => dv.sqref !== op.range)
	return ok(patch([], [op.sheet]))
}

function handleSetAutoFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setAutoFilter' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.autoFilter = { ref: op.range, columns: [] }
	return ok(patch([], [op.sheet]))
}

function handleClearAutoFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'clearAutoFilter' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.autoFilter = null
	return ok(patch([], [op.sheet]))
}

function handleSetSheetProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setSheetProtection' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const prot: Record<string, unknown> = { sheet: true }
	if (op.password) prot.password = op.password
	if (op.options?.formatCells !== undefined) prot.formatCells = op.options.formatCells
	if (op.options?.formatColumns !== undefined) prot.formatColumns = op.options.formatColumns
	if (op.options?.formatRows !== undefined) prot.formatRows = op.options.formatRows
	if (op.options?.insertColumns !== undefined) prot.insertColumns = op.options.insertColumns
	if (op.options?.insertRows !== undefined) prot.insertRows = op.options.insertRows
	if (op.options?.deleteColumns !== undefined) prot.deleteColumns = op.options.deleteColumns
	if (op.options?.deleteRows !== undefined) prot.deleteRows = op.options.deleteRows
	if (op.options?.sort !== undefined) prot.sort = op.options.sort
	if (op.options?.autoFilter !== undefined) prot.autoFilter = op.options.autoFilter
	sheet.protection = prot as never
	return ok(patch([], [op.sheet]))
}

function handleSetTabColor(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTabColor' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.tabColor = { rgb: op.color }
	return ok(patch([], [op.sheet]))
}

function handleHideSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideSheet' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.state = (op.hidden ?? true) ? 'hidden' : 'visible'
	return ok(patch([], [op.sheet]))
}

function handleHideRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideRows' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const hidden = op.hidden ?? true
	for (let r = op.at; r < op.at + op.count; r++) {
		if (hidden) sheet.rowHeights.set(r, 0)
	}
	return ok(patch([], [op.sheet]))
}

function handleHideCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideCols' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const hidden = op.hidden ?? true
	for (let c = op.at; c < op.at + op.count; c++) {
		const idx = sheet.colDefs.findIndex((d) => d.min === c + 1 && d.max === c + 1)
		if (idx >= 0) {
			const existing = sheet.colDefs[idx]
			if (existing) sheet.colDefs[idx] = { ...existing, hidden }
		} else {
			sheet.colDefs.push({ min: c + 1, max: c + 1, hidden })
		}
	}
	return ok(patch([], [op.sheet]))
}

function handleCopySheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'copySheet' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const source = sheetResult.value
	const pos = op.position ?? workbook.sheets.length
	const newSheet = source.clone()
	newSheet.name = op.newName
	workbook.sheets.splice(pos, 0, newSheet)
	invalidateSheetIndexCache(workbook)
	return ok(patch([], [op.newName]))
}

function handleSetConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setConditionalFormat' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const existing = sheet.conditionalFormats.findIndex((cf) => cf.sqref === op.range)
	const rule: Record<string, unknown> = {
		type: op.rule.type,
		formulas: [op.rule.formula, op.rule.formula2].filter((f): f is string => f !== undefined),
	}
	if (op.rule.operator) rule.operator = op.rule.operator
	if (op.rule.priority !== undefined) rule.priority = op.rule.priority
	if (op.rule.stopIfTrue !== undefined) rule.stopIfTrue = op.rule.stopIfTrue
	if (op.rule.style) rule.style = op.rule.style
	const cf = { sqref: op.range, rules: [rule] }
	if (existing >= 0) sheet.conditionalFormats[existing] = cf as never
	else sheet.conditionalFormats.push(cf as never)
	return ok(patch([], [op.sheet]))
}

function handleDeleteConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteConditionalFormat' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.conditionalFormats = sheet.conditionalFormats.filter((cf) => cf.sqref !== op.range)
	return ok(patch([], [op.sheet]))
}

function handleSetPageSetup(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPageSetup' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const setup: Record<string, unknown> = {}
	if (op.setup.orientation) setup.orientation = op.setup.orientation
	if (op.setup.paperSize !== undefined) setup.paperSize = op.setup.paperSize
	if (op.setup.scale !== undefined) setup.scale = op.setup.scale
	if (op.setup.fitToWidth !== undefined) setup.fitToWidth = op.setup.fitToWidth
	if (op.setup.fitToHeight !== undefined) setup.fitToHeight = op.setup.fitToHeight
	sheet.pageSetup = setup as never
	if (op.setup.margins) {
		sheet.pageMargins = op.setup.margins as never
	}
	return ok(patch([], [op.sheet]))
}

function handleSetPrintArea(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPrintArea' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	workbook.definedNames.set('_xlnm.Print_Area', `'${op.sheet}'!${op.range}`, {
		kind: 'sheet',
		sheetId: sheetResult.value.id,
	})
	return ok(patch([], [op.sheet]))
}

function handleSetRichText(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRichText' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const pos = parseA1(op.ref)
	const richTextValue: CellValue = { kind: 'richText', runs: op.runs }
	sheet.cells.setResolved(pos.row, pos.col, richTextValue, null, DEFAULT_SID)
	return ok(patch([`${op.sheet}!${op.ref}`], [op.sheet]))
}

// --- Handler registry ---

type OperationHandler = (workbook: Workbook, op: never) => Result<PatchResult>

const handlers: Record<string, OperationHandler> = {
	setCells: handleSetCells as OperationHandler,
	setFormula: handleSetFormula as OperationHandler,
	fillFormula: handleFillFormula as OperationHandler,
	insertRows: handleInsertRows as OperationHandler,
	deleteRows: handleDeleteRows as OperationHandler,
	insertCols: handleInsertCols as OperationHandler,
	deleteCols: handleDeleteCols as OperationHandler,
	addSheet: handleAddSheet as OperationHandler,
	createTable: handleCreateTable as OperationHandler,
	appendRows: handleAppendRows as OperationHandler,
	deleteSheet: handleDeleteSheet as OperationHandler,
	renameSheet: handleRenameSheet as OperationHandler,
	moveSheet: handleMoveSheet as OperationHandler,
	mergeCells: handleMergeCells as OperationHandler,
	unmergeCells: handleUnmergeCells as OperationHandler,
	setComment: handleSetComment as OperationHandler,
	setDefinedName: handleSetDefinedName as OperationHandler,
	deleteDefinedName: handleDeleteDefinedName as OperationHandler,
	clearRange: handleClearRange as OperationHandler,
	freezePane: handleFreezePane as OperationHandler,
	setColWidth: handleSetColWidth as OperationHandler,
	setRowHeight: handleSetRowHeight as OperationHandler,
	sortRange: handleSortRange as OperationHandler,
	setHyperlink: handleSetHyperlink as OperationHandler,
	setNumberFormat: handleSetNumberFormat as OperationHandler,
	setStyle: handleSetStyle as OperationHandler,
	deleteComment: handleDeleteComment as OperationHandler,
	deleteHyperlink: handleDeleteHyperlink as OperationHandler,
	setDataValidation: handleSetDataValidation as OperationHandler,
	deleteDataValidation: handleDeleteDataValidation as OperationHandler,
	setAutoFilter: handleSetAutoFilter as OperationHandler,
	clearAutoFilter: handleClearAutoFilter as OperationHandler,
	setSheetProtection: handleSetSheetProtection as OperationHandler,
	setTabColor: handleSetTabColor as OperationHandler,
	hideSheet: handleHideSheet as OperationHandler,
	hideRows: handleHideRows as OperationHandler,
	hideCols: handleHideCols as OperationHandler,
	copySheet: handleCopySheet as OperationHandler,
	setConditionalFormat: handleSetConditionalFormat as OperationHandler,
	deleteConditionalFormat: handleDeleteConditionalFormat as OperationHandler,
	setPageSetup: handleSetPageSetup as OperationHandler,
	setPrintArea: handleSetPrintArea as OperationHandler,
	copyRange: handleTransferRange as OperationHandler,
	moveRange: handleTransferRange as OperationHandler,
	setRichText: handleSetRichText as OperationHandler,
}

// --- Public API ---

export function applyOperation(workbook: Workbook, op: Operation): Result<PatchResult> {
	const useIncrementalPatch = op.op === 'setFormula' || op.op === 'fillFormula'
	if (operationAffectsFormulas(op) && !useIncrementalPatch) {
		invalidateWorkbookAnalysis(workbook)
	}
	invalidateSheetIndexCache(workbook)

	const handler = handlers[op.op]
	if (!handler) {
		return err(ascendError('VALIDATION_ERROR', `Unknown operation: ${op.op}`))
	}
	const result = handler(workbook, op as never)

	if (useIncrementalPatch && result.ok) {
		const changedKeys = resolveAffectedCellKeys(
			workbook,
			op as Extract<Operation, { op: 'setFormula' | 'fillFormula' }>,
		)
		if (changedKeys.length > 0) {
			patchWorkbookAnalysis(workbook, changedKeys)
		}
	}
	return result
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
