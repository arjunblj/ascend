import type { Workbook } from '@ascend/core'
import { parseA1, toA1 } from '@ascend/core'
import { cachedParseFormula, normalizeFormulaInput } from '@ascend/formulas'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, EMPTY, err, isEmpty, ok, richTextValue, valuesEqual } from '@ascend/schema'
import { validateCellValue } from '../data-validation.ts'
import type { PatchResult } from './helpers.ts'
import {
	cell,
	cellPreservingFormulaInfo,
	cellWithExisting,
	createLegacyArrayFormulaIndex,
	DEFAULT_SID,
	getSheet,
	inputToCellValue,
	legacyArrayFormulaEditError,
	materializeFormulaBindingGroupsForRangeEdit,
	materializeFormulaBindingGroupsForRefs,
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

	const warnings: ReturnType<typeof ascendError>[] = []
	const legacyArrayIndex = createLegacyArrayFormulaIndex(sheet)
	const prepared: Array<{
		readonly ref: ReturnType<typeof parseA1>
		readonly update: Extract<Operation, { op: 'setCells' }>['updates'][number]
		readonly value: ReturnType<typeof inputToCellValue>
	}> = []
	for (const update of op.updates) {
		const ref = parseA1(update.ref)
		const blocked = legacyArrayIndex.findCell(ref.row, ref.col)
		if (blocked) return err(legacyArrayFormulaEditError(update.ref, blocked.ref))
		const value = inputToCellValue(update.value, workbook.calcSettings.dateSystem)
		prepared.push({ ref, update, value })
	}

	const affected = materializeFormulaBindingGroupsForRefs(
		workbook,
		sheet,
		prepared.map(({ ref }) => ref),
	)
	for (const { ref, update, value } of prepared) {
		const existing = sheet.cells.get(ref.row, ref.col)
		if (!existing && isEmpty(value)) continue
		if (
			existing &&
			existing.formula === null &&
			existing.formulaInfo === undefined &&
			valuesEqual(existing.value, value)
		) {
			continue
		}
		const validation = validateCellValue(sheet, ref.row, ref.col, value, workbook)
		if (!validation.valid && validation.message) {
			warnings.push(ascendError('VALIDATION_ERROR', validation.message, { refs: [update.ref] }))
		}
		sheet.cells.set(
			ref.row,
			ref.col,
			cellWithExisting(value, null, existing?.styleId ?? DEFAULT_SID),
		)
		affected.add(update.ref)
	}

	return ok(
		patch(
			[...affected],
			affected.size > 0 ? [op.sheet] : [],
			affected.size > 0,
			warnings.length > 0 ? warnings : undefined,
		),
	)
}

export function handleSetFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setFormula' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const ref = parseA1(op.ref)
	const blocked = createLegacyArrayFormulaIndex(sheet).findCell(ref.row, ref.col)
	if (blocked) return err(legacyArrayFormulaEditError(op.ref, blocked.ref))
	const affected = materializeFormulaBindingGroupsForRefs(workbook, sheet, [ref])
	sheet.cells.set(
		ref.row,
		ref.col,
		cellWithExisting(
			sheet.cells.readValue(ref.row, ref.col),
			normalizeFormulaInput(op.formula),
			sheet.cells.readStyleId(ref.row, ref.col) ?? DEFAULT_SID,
		),
	)
	affected.add(op.ref)

	return ok(patch([...affected], [op.sheet], true))
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
	const blocked = createLegacyArrayFormulaIndex(sheet).findIntersection(range)
	if (blocked) return err(legacyArrayFormulaEditError(blocked.targetRef, blocked.ref))
	const affected = materializeFormulaBindingGroupsForRangeEdit(workbook, sheet, range)

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
			affected.add(toA1({ row, col }))
		}
	}

	return ok(patch([...affected], [op.sheet], true))
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
	const blocked = createLegacyArrayFormulaIndex(sheet).findCell(pos.row, pos.col)
	if (blocked) return err(legacyArrayFormulaEditError(op.ref, blocked.ref))
	const affected = new Set(
		[...materializeFormulaBindingGroupsForRefs(workbook, sheet, [pos])].map(
			(ref) => `${op.sheet}!${ref}`,
		),
	)
	sheet.cells.setResolved(pos.row, pos.col, richTextValue(op.runs), null, DEFAULT_SID)
	affected.add(`${op.sheet}!${op.ref}`)
	return ok(patch([...affected], [op.sheet], true))
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

	const blocked = createLegacyArrayFormulaIndex(sheet).findIntersection(range)
	if (blocked) return err(legacyArrayFormulaEditError(blocked.targetRef, blocked.ref))
	const affected =
		op.what === 'styles' || op.what === 'values'
			? new Set<string>()
			: materializeFormulaBindingGroupsForRangeEdit(workbook, sheet, range)
	for (let r = range.start.row; r <= range.end.row; r++) {
		for (let c = range.start.col; c <= range.end.col; c++) {
			const existing = sheet.cells.get(r, c)
			if (!existing) continue

			const ref = toA1({ row: r, col: c })
			affected.add(ref)

			switch (op.what) {
				case 'all':
					sheet.cells.delete(r, c)
					break
				case 'values':
					sheet.cells.set(
						r,
						c,
						cellPreservingFormulaInfo(
							EMPTY,
							existing.formula,
							existing.styleId,
							existing.formulaInfo,
						),
					)
					break
				case 'formulas':
					sheet.cells.set(r, c, cell(existing.value, null, existing.styleId))
					break
				case 'styles':
					sheet.cells.set(
						r,
						c,
						cellPreservingFormulaInfo(
							existing.value,
							existing.formula,
							DEFAULT_SID,
							existing.formulaInfo,
						),
					)
					break
			}
		}
	}

	return ok(patch([...affected], [op.sheet], op.what !== 'styles'))
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
			if (styleId === existingStyleId) continue
			const formulaInfo = sheet.cells.readFormulaInfo(row, col)
			sheet.cells.set(
				row,
				col,
				cellPreservingFormulaInfo(
					sheet.cells.readValue(row, col),
					sheet.cells.readFormula(row, col) ?? null,
					styleId,
					formulaInfo,
				),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, affected.length > 0 ? [op.sheet] : []))
}
