import type { Workbook } from '@ascend/core'
import { parseA1, toA1 } from '@ascend/core'
import { cachedParseFormula } from '@ascend/formulas'
import type { Operation, Result } from '@ascend/schema'
import { ok } from '@ascend/schema'
import {
	rewriteDefinedNameFormulasForShift,
	rewriteWorkbookFormulasForShift,
} from '../structural/formula-rewrite.ts'
import { shiftSheetCellMetadata } from '../structural/sheet-topology.ts'
import type { PatchResult } from './helpers.ts'
import {
	cellWithExisting,
	clearFormulaMetadata,
	collectRangeCells,
	getSheet,
	patch,
	safeParseRange,
	shiftMerges,
	translateFormula,
} from './helpers.ts'

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

export function handleInsertRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertRows' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'row', op.at, op.count, op.count)
}

export function handleDeleteRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteRows' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'row', op.at, op.count, -op.count)
}

export function handleInsertCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'insertCols' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'col', op.at, op.count, op.count)
}

export function handleDeleteCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteCols' }>,
): Result<PatchResult> {
	return applyAxisShift(workbook, op.sheet, 'col', op.at, op.count, -op.count)
}

export function handleTransferRange(
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
				formula !== null ? undefined : existingTarget?.formulaInfo,
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

export function handleMergeCells(
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

export function handleUnmergeCells(
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
