import type {
	CellStyle,
	SheetConditionalFormat,
	SheetConditionalFormatRule,
	SheetDataValidation,
	SheetPageSetup,
	Workbook,
} from '@ascend/core'
import { toA1 } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import type { PatchResult } from './helpers.ts'
import {
	cellWithExisting,
	DEFAULT_SID,
	getSheet,
	patch,
	safeParseRange,
	updateSheetOutlineLevels,
} from './helpers.ts'

export function handleSetNumberFormat(
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
			const existingStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_SID
			const currentStyle = workbook.styles.get(existingStyleId) ?? {}
			const style: CellStyle = {
				...currentStyle,
				numberFormat: op.format,
			}
			const styleId = workbook.styles.register(style)
			if (workbook.preservedStyles && styleId !== existingStyleId) {
				const baseStyleId =
					workbook.preservedStyles.baseStyleIdByStyleId?.[existingStyleId] ?? existingStyleId
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

export function handleSetConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setConditionalFormat' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const existing = sheet.conditionalFormats.findIndex((cf) => cf.sqref === op.range)
	const rule: SheetConditionalFormatRule = {
		type: op.rule.type,
		formulas: [op.rule.formula, op.rule.formula2].filter((f): f is string => f !== undefined),
		...(op.rule.operator ? { operator: op.rule.operator } : {}),
		...(op.rule.priority !== undefined ? { priority: op.rule.priority } : {}),
		...(op.rule.stopIfTrue !== undefined ? { stopIfTrue: op.rule.stopIfTrue } : {}),
		...(op.rule.style ? { style: op.rule.style } : {}),
		...(op.rule.colorScale
			? {
					colorScale: {
						cfvo: op.rule.colorScale.cfvo.map((entry) => ({ ...entry })),
						colors: op.rule.colorScale.colors.map((entry) => ({ ...entry })),
					},
				}
			: {}),
		...(op.rule.dataBar
			? {
					dataBar: {
						...op.rule.dataBar,
						cfvo: op.rule.dataBar.cfvo.map((entry) => ({ ...entry })),
						...(op.rule.dataBar.color ? { color: { ...op.rule.dataBar.color } } : {}),
					},
				}
			: {}),
		...(op.rule.iconSet
			? {
					iconSet: {
						...op.rule.iconSet,
						cfvo: op.rule.iconSet.cfvo.map((entry) => ({ ...entry })),
					},
				}
			: {}),
	}
	const cf: SheetConditionalFormat = { sqref: op.range, rules: [rule] }
	if (existing >= 0) sheet.conditionalFormats[existing] = cf
	else sheet.conditionalFormats.push(cf)
	return ok(patch([], [op.sheet]))
}

export function handleDeleteConditionalFormat(
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

export function handleSetDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDataValidation' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const existing = sheet.dataValidations.findIndex((dv) => dv.sqref === op.range)
	const dv: SheetDataValidation = {
		sqref: op.range,
		type: op.rule.type,
		allowBlank: op.rule.allowBlank ?? true,
		showErrorMessage: op.rule.showErrorMessage ?? true,
		...(op.rule.formula1 !== undefined ? { formula1: op.rule.formula1 } : {}),
		...(op.rule.formula2 !== undefined ? { formula2: op.rule.formula2 } : {}),
		...(op.rule.operator !== undefined ? { operator: op.rule.operator } : {}),
		...(op.rule.errorTitle !== undefined ? { errorTitle: op.rule.errorTitle } : {}),
		...(op.rule.errorMessage !== undefined ? { error: op.rule.errorMessage } : {}),
		...(op.rule.showInputMessage !== undefined
			? { showInputMessage: op.rule.showInputMessage }
			: {}),
		...(op.rule.promptTitle !== undefined ? { promptTitle: op.rule.promptTitle } : {}),
		...(op.rule.prompt !== undefined ? { prompt: op.rule.prompt } : {}),
	}
	if (existing >= 0) sheet.dataValidations[existing] = dv
	else sheet.dataValidations.push(dv)
	return ok(patch([], [op.sheet]))
}

export function handleDeleteDataValidation(
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

export function handleSetAutoFilter(
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

export function handleClearAutoFilter(
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

export function handleSetPageSetup(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPageSetup' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const setup: SheetPageSetup = {
		...(op.setup.orientation ? { orientation: op.setup.orientation } : {}),
		...(op.setup.paperSize !== undefined ? { paperSize: op.setup.paperSize } : {}),
		...(op.setup.scale !== undefined ? { scale: op.setup.scale } : {}),
		...(op.setup.fitToWidth !== undefined ? { fitToWidth: op.setup.fitToWidth } : {}),
		...(op.setup.fitToHeight !== undefined ? { fitToHeight: op.setup.fitToHeight } : {}),
	}
	sheet.pageSetup = setup
	if (op.setup.margins) {
		sheet.pageMargins = { ...op.setup.margins }
	}
	return ok(patch([], [op.sheet]))
}

export function handleSetPrintArea(
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

export function handleSetComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setComment' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result

	const comment = op.author !== undefined ? { text: op.text, author: op.author } : { text: op.text }
	result.value.comments.set(op.ref, comment)

	return ok(patch([op.ref], [op.sheet]))
}

export function handleDeleteComment(
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

export function handleSetHyperlink(
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

export function handleDeleteHyperlink(
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

export function handleSetDefinedName(
	_workbook: Workbook,
	op: Extract<Operation, { op: 'setDefinedName' }>,
): Result<PatchResult> {
	if (op.scope) {
		const sheet = _workbook.getSheet(op.scope)
		if (!sheet) {
			const available = _workbook.sheets.map((s) => s.name).join(', ')
			return err(
				ascendError('SHEET_NOT_FOUND', `Sheet "${op.scope}" not found`, {
					suggestedFix: available ? `Available sheets: ${available}` : 'Workbook has no sheets',
				}),
			)
		}
		_workbook.definedNames.set(op.name, op.ref, { kind: 'sheet', sheetId: sheet.id })
	} else {
		_workbook.definedNames.set(op.name, op.ref)
	}
	return ok(patch([], []))
}

export function handleDeleteDefinedName(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDefinedName' }>,
): Result<PatchResult> {
	if (op.scope) {
		const sheet = workbook.getSheet(op.scope)
		if (!sheet) {
			const available = workbook.sheets.map((s) => s.name).join(', ')
			return err(
				ascendError('SHEET_NOT_FOUND', `Sheet "${op.scope}" not found`, {
					suggestedFix: available ? `Available sheets: ${available}` : 'Workbook has no sheets',
				}),
			)
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

export function handleGroupRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'groupRows' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	if (op.from > op.to || op.from < 0) {
		return err(ascendError('VALIDATION_ERROR', 'Invalid row group range'))
	}
	const summaryBelow = op.summaryBelow ?? sheet.outlinePr?.summaryBelow ?? true
	sheet.outlinePr = { ...(sheet.outlinePr ?? {}), summaryBelow }
	for (let row = op.from; row <= op.to; row++) {
		const existing = sheet.rowDefs.get(row)
		sheet.rowDefs.set(row, {
			...existing,
			outlineLevel: Math.min(7, (existing?.outlineLevel ?? 0) + 1),
			...(op.collapsed !== undefined ? { hidden: op.collapsed } : {}),
		})
	}
	if (op.collapsed) {
		const boundaryRow = summaryBelow ? op.to + 1 : op.from - 1
		if (boundaryRow >= 0) {
			const existing = sheet.rowDefs.get(boundaryRow)
			sheet.rowDefs.set(boundaryRow, { ...existing, collapsed: true })
		}
	}
	updateSheetOutlineLevels(sheet)
	return ok(patch([], [op.sheet]))
}

export function handleGroupCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'groupCols' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	if (op.from > op.to || op.from < 0) {
		return err(ascendError('VALIDATION_ERROR', 'Invalid column group range'))
	}
	const summaryRight = op.summaryRight ?? sheet.outlinePr?.summaryRight ?? true
	sheet.outlinePr = { ...(sheet.outlinePr ?? {}), summaryRight }
	for (let col = op.from; col <= op.to; col++) {
		const idx = sheet.colDefs.findIndex((def) => def.min === col && def.max === col)
		const existing = idx >= 0 ? sheet.colDefs[idx] : undefined
		const next = {
			...(existing ?? { min: col, max: col }),
			outlineLevel: Math.min(7, (existing?.outlineLevel ?? 0) + 1),
			...(op.collapsed !== undefined ? { hidden: op.collapsed } : {}),
		}
		if (idx >= 0) sheet.colDefs[idx] = next
		else sheet.colDefs.push(next)
	}
	if (op.collapsed) {
		const boundaryCol = summaryRight ? op.to + 1 : op.from - 1
		if (boundaryCol >= 0) {
			const idx = sheet.colDefs.findIndex(
				(def) => def.min === boundaryCol && def.max === boundaryCol,
			)
			const existing = idx >= 0 ? sheet.colDefs[idx] : undefined
			const next = { ...(existing ?? { min: boundaryCol, max: boundaryCol }), collapsed: true }
			if (idx >= 0) sheet.colDefs[idx] = next
			else sheet.colDefs.push(next)
		}
	}
	updateSheetOutlineLevels(sheet)
	return ok(patch([], [op.sheet]))
}
