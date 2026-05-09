import type { RangeRef, Sheet, SheetDataValidation, Workbook } from '@ascend/core'
import { parseA1, parseRange, toA1 } from '@ascend/core'
import { cachedParseFormula } from '@ascend/formulas'
import type { Operation, PasteMode, Result } from '@ascend/schema'
import { EMPTY, ok } from '@ascend/schema'
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
	DEFAULT_SID,
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
	const mode = op.mode ?? 'all'
	const snapshot = collectRangeCells(sheet, source)
	const affected: string[] = []

	if (pasteCells(mode)) {
		for (const entry of snapshot) {
			const targetRow = entry.row + rowDelta
			const targetCol = entry.col + colDelta
			const existingTarget = sheet.cells.get(targetRow, targetCol)
			const ref = toA1({ row: targetRow, col: targetCol })

			if (!entry.cell && mode === 'all') {
				sheet.cells.delete(targetRow, targetCol)
				affected.push(ref)
				continue
			}

			const formula = translateCellFormula(entry.cell?.formula ?? null, rowDelta, colDelta)
			const targetValue = existingTarget?.value ?? EMPTY
			const targetFormula = existingTarget?.formula ?? null
			const targetStyle = existingTarget?.styleId ?? DEFAULT_SID

			if (mode === 'values') {
				sheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(entry.cell?.value ?? EMPTY, null, targetStyle),
				)
			} else if (mode === 'formulas') {
				sheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(entry.cell?.value ?? EMPTY, formula, targetStyle),
				)
			} else if (mode === 'formats' || mode === 'styles') {
				sheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(
						targetValue,
						targetFormula,
						entry.cell?.styleId ?? DEFAULT_SID,
						targetFormula !== null ? existingTarget?.formulaInfo : undefined,
					),
				)
			} else if (entry.cell) {
				sheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(entry.cell.value, formula, entry.cell.styleId),
				)
			} else {
				sheet.cells.delete(targetRow, targetCol)
			}
			affected.push(ref)
		}
	}

	copyTransferMetadata(sheet, source, rowDelta, colDelta, mode, op.op === 'moveRange')

	if (op.op === 'moveRange') {
		for (const entry of snapshot) {
			if (pasteCells(mode)) sheet.cells.delete(entry.row, entry.col)
			if (pasteCells(mode)) affected.push(toA1({ row: entry.row, col: entry.col }))
		}
	}

	return ok(patch(affected, [op.sheet], pasteRequiresRecalc(mode)))
}

function pasteCells(mode: PasteMode): boolean {
	return (
		mode === 'all' ||
		mode === 'values' ||
		mode === 'formulas' ||
		mode === 'formats' ||
		mode === 'styles'
	)
}

function pasteRequiresRecalc(mode: PasteMode): boolean {
	return mode === 'all' || mode === 'values' || mode === 'formulas'
}

function translateCellFormula(
	formula: string | null,
	rowDelta: number,
	colDelta: number,
): string | null {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	return parsed.ok ? translateFormula(parsed.value, rowDelta, colDelta) : formula
}

function copyTransferMetadata(
	sheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	mode: PasteMode,
	move: boolean,
): void {
	if (mode === 'all' || mode === 'comments') {
		copyCellMap(sheet.comments, source, rowDelta, colDelta, move)
	}
	if (mode === 'all' || mode === 'hyperlinks') {
		copyCellMap(sheet.hyperlinks, source, rowDelta, colDelta, move)
	}
	if (mode === 'all' || mode === 'validations') {
		copyDataValidations(sheet, source, rowDelta, colDelta, move)
	}
}

function copyCellMap<T extends object>(
	map: Map<string, T>,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const entries = [...map.entries()].filter(([ref]) => rangeContainsCell(source, parseA1(ref)))
	for (const [ref, value] of entries) {
		const pos = parseA1(ref)
		map.set(toA1({ row: pos.row + rowDelta, col: pos.col + colDelta }), { ...value })
	}
	if (move) {
		for (const [ref] of entries) map.delete(ref)
	}
}

function copyDataValidations(
	sheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const copied: SheetDataValidation[] = []
	const retained: SheetDataValidation[] = []

	for (const validation of sheet.dataValidations) {
		const ranges = parseSqref(validation.sqref)
		if (ranges.length === 0) {
			retained.push(validation)
			continue
		}

		const copiedRanges = ranges.filter((range) => rangeContainsRange(source, range))
		if (copiedRanges.length === 0) {
			retained.push(validation)
			continue
		}

		if (!move) retained.push(validation)
		else {
			const keptRanges = ranges.filter((range) => !rangeContainsRange(source, range))
			if (keptRanges.length > 0) {
				retained.push({ ...validation, sqref: rangesToSqref(keptRanges) })
			}
		}

		copied.push({
			...validation,
			sqref: rangesToSqref(copiedRanges.map((range) => shiftRange(range, rowDelta, colDelta))),
			...(validation.formula1
				? { formula1: translateMetadataFormula(validation.formula1, rowDelta, colDelta) }
				: {}),
			...(validation.formula2
				? { formula2: translateMetadataFormula(validation.formula2, rowDelta, colDelta) }
				: {}),
		})
	}

	sheet.dataValidations = [...retained, ...copied]
}

function parseSqref(sqref: string): RangeRef[] {
	const ranges: RangeRef[] = []
	for (const token of sqref.trim().split(/\s+/)) {
		if (!token) continue
		try {
			ranges.push(parseRange(token))
		} catch {
			return []
		}
	}
	return ranges
}

function rangeContainsCell(range: RangeRef, ref: { row: number; col: number }): boolean {
	return (
		ref.row >= range.start.row &&
		ref.row <= range.end.row &&
		ref.col >= range.start.col &&
		ref.col <= range.end.col
	)
}

function rangeContainsRange(outer: RangeRef, inner: RangeRef): boolean {
	return rangeContainsCell(outer, inner.start) && rangeContainsCell(outer, inner.end)
}

function shiftRange(range: RangeRef, rowDelta: number, colDelta: number): RangeRef {
	return {
		start: { row: range.start.row + rowDelta, col: range.start.col + colDelta },
		end: { row: range.end.row + rowDelta, col: range.end.col + colDelta },
	}
}

function rangesToSqref(ranges: readonly RangeRef[]): string {
	return ranges.map(rangeToA1).join(' ')
}

function rangeToA1(range: RangeRef): string {
	const start = toA1(range.start)
	const end = toA1(range.end)
	return start === end ? start : `${start}:${end}`
}

function translateMetadataFormula(formula: string, rowDelta: number, colDelta: number): string {
	const parsed = cachedParseFormula(formula)
	return parsed.ok ? translateFormula(parsed.value, rowDelta, colDelta) : formula
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
