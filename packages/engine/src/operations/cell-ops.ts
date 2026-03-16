import type { Workbook } from '@ascend/core'
import { parseA1, toA1 } from '@ascend/core'
import { cachedParseFormula, normalizeFormulaInput } from '@ascend/formulas'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, EMPTY, err, ok, richTextValue } from '@ascend/schema'
import { validateCellValue } from '../data-validation.ts'
import type { PatchResult } from './helpers.ts'
import {
	cell,
	cellWithExisting,
	DEFAULT_SID,
	getSheet,
	inputToCellValue,
	mergeStyleInput,
	patch,
	safeParseRange,
	translateFormula,
} from './helpers.ts'

export function handleSetCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCells' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const affected: string[] = []
	const warnings: ReturnType<typeof ascendError>[] = []
	for (const update of op.updates) {
		const ref = parseA1(update.ref)
		const value = inputToCellValue(update.value, workbook.calcSettings.dateSystem)
		const validation = validateCellValue(sheet, ref.row, ref.col, value, workbook)
		if (!validation.valid && validation.message) {
			warnings.push(ascendError('VALIDATION_ERROR', validation.message, { refs: [update.ref] }))
		}
		sheet.cells.set(
			ref.row,
			ref.col,
			cellWithExisting(
				value,
				sheet.cells.readFormula(ref.row, ref.col) ?? null,
				sheet.cells.readStyleId(ref.row, ref.col) ?? DEFAULT_SID,
				sheet.cells.readFormulaInfo(ref.row, ref.col),
			),
		)
		affected.push(update.ref)
	}

	return ok(patch(affected, [op.sheet], true, warnings.length > 0 ? warnings : undefined))
}

export function handleSetFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setFormula' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const ref = parseA1(op.ref)
	sheet.cells.set(
		ref.row,
		ref.col,
		cellWithExisting(
			sheet.cells.readValue(ref.row, ref.col),
			normalizeFormulaInput(op.formula),
			sheet.cells.readStyleId(ref.row, ref.col) ?? DEFAULT_SID,
		),
	)

	return ok(patch([op.ref], [op.sheet], true))
}

export function handleFillFormula(
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
			sheet.cells.set(
				row,
				col,
				cellWithExisting(
					sheet.cells.readValue(row, col),
					translated,
					sheet.cells.readStyleId(row, col) ?? DEFAULT_SID,
				),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, [op.sheet], true))
}

export function handleSetRichText(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRichText' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const pos = parseA1(op.ref)
	sheet.cells.setResolved(pos.row, pos.col, richTextValue(op.runs), null, DEFAULT_SID)
	return ok(patch([`${op.sheet}!${op.ref}`], [op.sheet]))
}

export function handleClearRange(
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
						cellWithExisting(EMPTY, existing.formula, existing.styleId, existing.formulaInfo),
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

export function handleSetStyle(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setStyle' }>,
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
			const existingStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_SID
			const currentStyle = workbook.styles.get(existingStyleId) ?? {}
			const merged = mergeStyleInput(currentStyle, op.style)
			const styleId = workbook.styles.register(merged)
			sheet.cells.set(
				row,
				col,
				cellWithExisting(
					sheet.cells.readValue(row, col),
					sheet.cells.readFormula(row, col) ?? null,
					styleId,
					sheet.cells.readFormulaInfo(row, col),
				),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, [op.sheet]))
}
