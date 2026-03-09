import type { Cell, CellStyle, RangeRef, Sheet, StyleId, Workbook } from '@ascend/core'
import { columnToIndex, createTableId, parseA1, parseRange, toA1 } from '@ascend/core'
import type { FormulaCellRef, FormulaNode } from '@ascend/formulas'
import {
	compareValues,
	dateToSerial,
	parseFormula,
	printFormula,
	rewriteRefs,
} from '@ascend/formulas'
import type { CellValue, InputValue, Operation, Result, SortSpec } from '@ascend/schema'
import { ascendError, booleanValue, EMPTY, err, numberValue, ok, stringValue } from '@ascend/schema'

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

function safeParseRange(range: string): Result<RangeRef> {
	try {
		return ok(parseRange(range))
	} catch {
		return err(ascendError('INVALID_RANGE', `Invalid range: ${range}`))
	}
}

// --- Formula rewriting helpers ---

function rewriteNodeForShift(
	node: FormulaNode,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): FormulaNode {
	const onTarget = (refSheet: string | undefined) => (refSheet ?? formulaSheet) === targetSheet

	const shiftRef = (ref: FormulaCellRef, isOnTarget: boolean): FormulaCellRef => {
		if (!isOnTarget) return ref
		if (axis === 'row') {
			if (delta > 0) {
				if (ref.row >= at) return { ...ref, row: ref.row + delta }
			} else {
				const deleteEnd = at - delta
				if (ref.row >= deleteEnd) return { ...ref, row: ref.row + delta }
			}
		} else {
			if (delta > 0) {
				if (ref.col >= at) return { ...ref, col: ref.col + delta }
			} else {
				const deleteEnd = at - delta
				if (ref.col >= deleteEnd) return { ...ref, col: ref.col + delta }
			}
		}
		return ref
	}

	switch (node.type) {
		case 'cellRef': {
			const ref = shiftRef(node.ref, onTarget(node.sheet))
			return node.sheet !== undefined
				? { type: 'cellRef', ref, sheet: node.sheet }
				: { type: 'cellRef', ref }
		}
		case 'rangeRef': {
			const hit = onTarget(node.sheet)
			const start = shiftRef(node.start, hit)
			const end = shiftRef(node.end, hit)
			return node.sheet !== undefined
				? { type: 'rangeRef', start, end, sheet: node.sheet }
				: { type: 'rangeRef', start, end }
		}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteNodeForShift(node.left, targetSheet, formulaSheet, axis, at, delta),
				right: rewriteNodeForShift(node.right, targetSheet, formulaSheet, axis, at, delta),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteNodeForShift(node.operand, targetSheet, formulaSheet, axis, at, delta),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((a) =>
					rewriteNodeForShift(a, targetSheet, formulaSheet, axis, at, delta),
				),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((r) =>
					r.map((c) => rewriteNodeForShift(c, targetSheet, formulaSheet, axis, at, delta)),
				),
			}
		default:
			return node
	}
}

function rewriteSheetName(node: FormulaNode, oldName: string, newName: string): FormulaNode {
	switch (node.type) {
		case 'cellRef':
			return node.sheet === oldName ? { type: 'cellRef', ref: node.ref, sheet: newName } : node
		case 'rangeRef':
			return node.sheet === oldName
				? { type: 'rangeRef', start: node.start, end: node.end, sheet: newName }
				: node
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteSheetName(node.left, oldName, newName),
				right: rewriteSheetName(node.right, oldName, newName),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteSheetName(node.operand, oldName, newName),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((a) => rewriteSheetName(a, oldName, newName)),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((r) => r.map((c) => rewriteSheetName(c, oldName, newName))),
			}
		default:
			return node
	}
}

function rewriteAllFormulas(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, c] of sheet.cells.iterate()) {
			if (c.formula === null) continue
			const parsed = parseFormula(c.formula)
			if (!parsed.ok) continue
			const rewritten = rewriteNodeForShift(parsed.value, targetSheet, sheet.name, axis, at, delta)
			const newFormula = printFormula(rewritten)
			if (newFormula !== c.formula) {
				updates.push([row, col, cell(c.value, newFormula, c.styleId)])
			}
		}
		for (const [r, c, updated] of updates) {
			sheet.cells.set(r, c, updated)
		}
	}
}

function rewriteSheetNameInFormulas(workbook: Workbook, oldName: string, newName: string): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, c] of sheet.cells.iterate()) {
			if (c.formula === null) continue
			const parsed = parseFormula(c.formula)
			if (!parsed.ok) continue
			const rewritten = rewriteSheetName(parsed.value, oldName, newName)
			const newFormula = printFormula(rewritten)
			if (newFormula !== c.formula) {
				updates.push([row, col, cell(c.value, newFormula, c.styleId)])
			}
		}
		for (const [r, c, updated] of updates) {
			sheet.cells.set(r, c, updated)
		}
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
			cell(value, existing?.formula ?? null, existing?.styleId ?? DEFAULT_SID),
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
		cell(
			existing?.value ?? EMPTY,
			normalizeFormulaInput(op.formula),
			existing?.styleId ?? DEFAULT_SID,
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
	const parsed = parseFormula(baseFormula)
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
				cell(existing?.value ?? EMPTY, translated, existing?.styleId ?? DEFAULT_SID),
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

	const entries: [number, number, Cell][] = []
	for (const [row, col, c] of sheet.cells.iterate()) {
		entries.push([row, col, c])
	}
	sheet.cells.clear()
	for (const [row, col, c] of entries) {
		sheet.cells.set(row >= op.at ? row + op.count : row, col, c)
	}

	shiftMerges(sheet.merges, 'row', op.at, op.count)
	rewriteAllFormulas(workbook, op.sheet, 'row', op.at, op.count)

	return ok(patch([], [op.sheet], true))
}

function handleDeleteRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteRows' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const deleteEnd = op.at + op.count
	const entries: [number, number, Cell][] = []
	for (const [row, col, c] of sheet.cells.iterate()) {
		entries.push([row, col, c])
	}
	sheet.cells.clear()
	for (const [row, col, c] of entries) {
		if (row >= op.at && row < deleteEnd) continue
		sheet.cells.set(row >= deleteEnd ? row - op.count : row, col, c)
	}

	shiftMerges(sheet.merges, 'row', op.at, -op.count)
	rewriteAllFormulas(workbook, op.sheet, 'row', op.at, -op.count)

	return ok(patch([], [op.sheet], true))
}

function handleInsertCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertCols' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const entries: [number, number, Cell][] = []
	for (const [row, col, c] of sheet.cells.iterate()) {
		entries.push([row, col, c])
	}
	sheet.cells.clear()
	for (const [row, col, c] of entries) {
		sheet.cells.set(row, col >= op.at ? col + op.count : col, c)
	}

	shiftMerges(sheet.merges, 'col', op.at, op.count)
	rewriteAllFormulas(workbook, op.sheet, 'col', op.at, op.count)

	return ok(patch([], [op.sheet], true))
}

function handleDeleteCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteCols' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const deleteEnd = op.at + op.count
	const entries: [number, number, Cell][] = []
	for (const [row, col, c] of sheet.cells.iterate()) {
		entries.push([row, col, c])
	}
	sheet.cells.clear()
	for (const [row, col, c] of entries) {
		if (col >= op.at && col < deleteEnd) continue
		sheet.cells.set(row, col >= deleteEnd ? col - op.count : col, c)
	}

	shiftMerges(sheet.merges, 'col', op.at, -op.count)
	rewriteAllFormulas(workbook, op.sheet, 'col', op.at, -op.count)

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
					cell(
						inputToCellValue(provided),
						existing?.formula ?? null,
						existing?.styleId ?? DEFAULT_SID,
					),
				)
			} else {
				const formula = table.columns[colOffset]?.formula
				sheet.cells.set(
					nextRow,
					col,
					cell(existing?.value ?? EMPTY, formula ?? null, existing?.styleId ?? DEFAULT_SID),
				)
			}
			affected.push(ref)
		}
		nextRow++
	}

	const tableIndex = sheet.tables.findIndex((candidate) => candidate.id === table.id)
	if (tableIndex >= 0) {
		sheet.tables.splice(tableIndex, 1, {
			...table,
			ref: {
				start: table.ref.start,
				end: { row: table.ref.end.row + op.rows.length, col: table.ref.end.col },
			},
		})
	}
	return ok(patch(affected, [sheet.name], true))
}

function handleDeleteSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteSheet' }>,
): Result<PatchResult> {
	if (!workbook.getSheet(op.sheet)) {
		return err(ascendError('SHEET_NOT_FOUND', `Sheet "${op.sheet}" not found`))
	}
	workbook.removeSheet(op.sheet)
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
	rewriteSheetNameInFormulas(workbook, oldName, op.newName)

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

	return ok(patch([], [op.sheet]))
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
					sheet.cells.set(r, c, cell(EMPTY, existing.formula, existing.styleId))
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

	if (sheet.merges.some((merge) => rangesOverlap(merge, range))) {
		return err(ascendError('VALIDATION_ERROR', 'sortRange does not support merged cells yet'))
	}

	const resolvedColumns = resolveSortColumns(sheet, range, op.by)
	if (!resolvedColumns.ok) return resolvedColumns
	const { columns, headerRow } = resolvedColumns.value
	const dataStartRow = headerRow ? range.start.row + 1 : range.start.row
	if (dataStartRow > range.end.row) return ok(patch([], [op.sheet], false))

	const rows = captureSortedRows(sheet, range, dataStartRow)
	rows.sort((left, right) => compareSortRows(left, right, columns, range.start.col))
	rewriteSortedRows(sheet, range, dataStartRow, rows)
	return ok(patch([], [op.sheet], true))
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
			sheet.cells.set(row, col, cell(existing?.value ?? EMPTY, existing?.formula ?? null, styleId))
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, [op.sheet]))
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

function captureSortedRows(sheet: Sheet, range: RangeRef, startRow: number): SortRow[] {
	const rows: SortRow[] = []
	for (let row = startRow; row <= range.end.row; row++) {
		const cells: Array<{ colOffset: number; cell: Cell }> = []
		for (let col = range.start.col; col <= range.end.col; col++) {
			const cell = sheet.cells.get(row, col)
			if (cell) cells.push({ colOffset: col - range.start.col, cell })
		}

		const comments: Array<{
			colOffset: number
			ref: string
			comment: Sheet['comments'] extends Map<string, infer T> ? T : never
		}> = []
		for (const [ref, comment] of sheet.comments) {
			const pos = parseA1(ref)
			if (pos.row === row && pos.col >= range.start.col && pos.col <= range.end.col) {
				comments.push({ colOffset: pos.col - range.start.col, ref, comment })
			}
		}

		const hyperlinks: Array<{
			colOffset: number
			ref: string
			hyperlink: Sheet['hyperlinks'] extends Map<string, infer T> ? T : never
		}> = []
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
			rowHeight: sheet.rowHeights.get(row),
		})
	}
	return rows
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
	readonly rowHeight: number | undefined
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
		if (rowData.rowHeight !== undefined) {
			sheet.rowHeights.set(targetRow, rowData.rowHeight)
		}
	})
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

// --- Public API ---

export function applyOperation(workbook: Workbook, op: Operation): Result<PatchResult> {
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
