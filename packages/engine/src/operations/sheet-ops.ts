import {
	type CellFormulaBinding,
	createSheetId,
	parseA1,
	type Sheet,
	toA1,
	type Workbook,
} from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok, validateExcelWorksheetName } from '@ascend/schema'
import { invalidateSheetIndexCache } from '../evaluator.ts'
import {
	rewriteFormulaTextForRename,
	rewriteSheetMetadataFormulasForRename,
	rewriteSheetNameInDefinedNames,
	rewriteSheetNameInFormulas,
} from '../structural/formula-rewrite.ts'
import { renameHyperlinkLocation } from '../structural/sheet-topology.ts'
import type { PatchResult } from './helpers.ts'
import {
	collectFormulaBindingGroupRefsForRefs,
	getSheet,
	materializeFormulaBindingGroupsForRefs,
	patch,
} from './helpers.ts'

export function handleAddSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'addSheet' }>,
): Result<PatchResult> {
	const nameError = worksheetNameError(op.name)
	if (nameError) return err(nameError)
	if (op.position !== undefined) {
		const positionError = validateSheetPosition(
			'addSheet position',
			op.position,
			workbook.sheets.length,
		)
		if (positionError) return err(positionError)
	}
	if (workbook.getSheet(op.name)) {
		return err(
			ascendError('NAME_CONFLICT', `Sheet "${op.name}" already exists`, {
				suggestedFix: 'Choose a different sheet name, or delete the existing sheet first',
			}),
		)
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

export function handleDeleteSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteSheet' }>,
): Result<PatchResult> {
	const targetSheet = workbook.getSheet(op.sheet)
	if (!targetSheet) {
		const available = workbook.sheets.map((s) => s.name).join(', ')
		return err(
			ascendError('SHEET_NOT_FOUND', `Sheet "${op.sheet}" not found`, {
				suggestedFix: available ? `Available sheets: ${available}` : 'Workbook has no sheets',
			}),
		)
	}
	const removedPivotNames = workbook.pivotTables
		.filter((entry) => sameSheetName(entry.sheetName, op.sheet))
		.map((entry) => entry.name)
		.filter((name): name is string => Boolean(name))
	workbook.removeSheet(op.sheet)
	removeSheetScopedDefinedNames(workbook, targetSheet.id)
	removeWorkbookMetadataForDeletedSheet(workbook, op.sheet, removedPivotNames)
	removeChartsForDeletedSheet(workbook, op.sheet)
	return ok(patch([], [op.sheet]))
}

export function handleRenameSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameSheet' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	const nameError = worksheetNameError(op.newName)
	if (nameError) return err(nameError)
	if (workbook.getSheet(op.newName)) {
		return err(
			ascendError('NAME_CONFLICT', `Sheet "${op.newName}" already exists`, {
				suggestedFix: 'Choose a different name; a sheet with that name already exists',
			}),
		)
	}

	const oldName = sheet.name
	const affected = new Set<string>()
	const sheetsModified = new Set<string>([op.newName])
	materializeWorkbookFormulaBindingsForRename(workbook, sheet, op.newName, affected, sheetsModified)
	sheet.name = op.newName
	workbook.invalidateSheetCache()
	for (const rewritten of rewriteSheetNameInFormulas(workbook, oldName, op.newName)) {
		affected.add(`${rewritten.sheetName}!${rewritten.ref}`)
		const rewrittenSheet = workbook.getSheet(rewritten.sheetName)
		if (rewrittenSheet) {
			for (const groupRef of collectFormulaBindingGroupRefsForRefs(workbook, rewrittenSheet, [
				parseA1(rewritten.ref),
			])) {
				affected.add(`${rewritten.sheetName}!${groupRef}`)
			}
		}
		sheetsModified.add(rewritten.sheetName)
	}
	rewriteSheetNameInDefinedNames(workbook, oldName, op.newName)
	rewriteChartSheetReferencesForRename(workbook, oldName, op.newName)
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

	return ok(patch([...affected], [...sheetsModified]))
}

export function handleMoveSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'moveSheet' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	const positionError = validateSheetPosition(
		'moveSheet position',
		op.position,
		workbook.sheets.length - 1,
	)
	if (positionError) return err(positionError)

	const idx = workbook.sheets.indexOf(sheet)
	workbook.sheets.splice(idx, 1)
	workbook.sheets.splice(op.position, 0, sheet)
	workbook.invalidateSheetCache()

	return ok(patch([], [op.sheet]))
}

export function handleCopySheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'copySheet' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const source = sheetResult.value
	const nameError = worksheetNameError(op.newName)
	if (nameError) return err(nameError)
	if (workbook.getSheet(op.newName)) {
		return err(
			ascendError('NAME_CONFLICT', `Sheet "${op.newName}" already exists`, {
				suggestedFix: 'Choose a different name; a sheet with that name already exists',
			}),
		)
	}
	const pos = op.position ?? workbook.sheets.length
	const positionError = validateSheetPosition('copySheet position', pos, workbook.sheets.length)
	if (positionError) return err(positionError)
	const newSheet = source.clone()
	const mutableNewSheet = newSheet as { id: Sheet['id'] }
	mutableNewSheet.id = createSheetId()
	newSheet.name = op.newName
	newSheet.ensureWritable()
	newSheet.preservedXml = null
	retargetCopiedSheetFormulaBindings(newSheet, source.name, op.newName)
	rewriteSheetMetadataFormulasForRename(newSheet, source.name, op.newName)
	duplicateCopiedSheetDefinedNames(workbook, source, newSheet)
	retargetCopiedSheetDrawingParts(workbook, newSheet)
	retargetCopiedSheetImageTargets(workbook, newSheet)
	const chartPartPaths = cloneChartsForCopiedSheet(workbook, source.name, op.newName)
	retargetCopiedSheetChartRelationships(newSheet, chartPartPaths)
	workbook.sheets.splice(pos, 0, newSheet)
	workbook.invalidateSheetCache()
	invalidateSheetIndexCache(workbook)
	return ok(patch([], [op.newName]))
}

function duplicateCopiedSheetDefinedNames(workbook: Workbook, source: Sheet, copy: Sheet): void {
	const copiedScope = { kind: 'sheet' as const, sheetId: copy.id }
	for (const entry of workbook.definedNames.list()) {
		if (entry.scope.kind !== 'sheet' || entry.scope.sheetId !== source.id) continue
		workbook.definedNames.add(
			entry.name,
			rewriteFormulaTextForRename(entry.formula, source.name, copy.name) ?? entry.formula,
			copiedScope,
			{
				...(entry.hidden !== undefined ? { hidden: entry.hidden } : {}),
				...(entry.extraAttributes ? { extraAttributes: entry.extraAttributes } : {}),
			},
		)
	}
}

function worksheetNameError(name: string): ReturnType<typeof ascendError> | null {
	const validation = validateExcelWorksheetName(name)
	if (!validation) return null
	return ascendError('VALIDATION_ERROR', validation.message, {
		suggestedFix: validation.suggestedFix,
	})
}

export function handleHideSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideSheet' }>,
): Result<PatchResult> {
	const hiddenValidation = validateOptionalBoolean('hideSheet hidden', op.hidden)
	if (hiddenValidation) return err(hiddenValidation)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.state = (op.hidden ?? true) ? 'hidden' : 'visible'
	return ok(patch([], [op.sheet]))
}

export function handleSetTabColor(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTabColor' }>,
): Result<PatchResult> {
	if (!/^(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/.test(op.color)) {
		return err(ascendError('VALIDATION_ERROR', 'tab color must be a 6 or 8 digit RGB hex value'))
	}
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.tabColor = { rgb: op.color }
	return ok(patch([], [op.sheet]))
}

export function handleSetSheetProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setSheetProtection' }>,
): Result<PatchResult> {
	const validation = validateSheetProtectionInput(op)
	if (validation) return err(validation)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const prot: import('@ascend/core').SheetProtection = {
		sheet: true,
		...(op.password ? { password: op.password } : {}),
		...(op.options?.formatCells !== undefined ? { formatCells: op.options.formatCells } : {}),
		...(op.options?.formatColumns !== undefined ? { formatColumns: op.options.formatColumns } : {}),
		...(op.options?.formatRows !== undefined ? { formatRows: op.options.formatRows } : {}),
		...(op.options?.insertColumns !== undefined ? { insertColumns: op.options.insertColumns } : {}),
		...(op.options?.insertRows !== undefined ? { insertRows: op.options.insertRows } : {}),
		...(op.options?.deleteColumns !== undefined ? { deleteColumns: op.options.deleteColumns } : {}),
		...(op.options?.deleteRows !== undefined ? { deleteRows: op.options.deleteRows } : {}),
		...(op.options?.sort !== undefined ? { sort: op.options.sort } : {}),
		...(op.options?.autoFilter !== undefined ? { autoFilter: op.options.autoFilter } : {}),
	}
	sheet.protection = prot
	return ok(patch([], [op.sheet]))
}

function validateSheetProtectionInput(op: Extract<Operation, { op: 'setSheetProtection' }>) {
	if (op.password !== undefined && typeof op.password !== 'string') {
		return ascendError('VALIDATION_ERROR', 'setSheetProtection password must be a string', {
			suggestedFix: 'Use a string password value or omit password.',
		})
	}
	if (
		op.options !== undefined &&
		(!op.options || typeof op.options !== 'object' || Array.isArray(op.options))
	) {
		return ascendError('VALIDATION_ERROR', 'setSheetProtection options must be an object', {
			suggestedFix: 'Provide sheet protection options such as { "sort": true }.',
		})
	}
	for (const field of SHEET_PROTECTION_OPTION_FIELDS) {
		const value = op.options?.[field]
		if (value !== undefined && typeof value !== 'boolean') {
			return ascendError('VALIDATION_ERROR', `sheet protection ${field} must be boolean`, {
				suggestedFix: `Set ${field}=true or ${field}=false.`,
			})
		}
	}
	return null
}

const SHEET_PROTECTION_OPTION_FIELDS = [
	'formatCells',
	'formatColumns',
	'formatRows',
	'insertColumns',
	'insertRows',
	'deleteColumns',
	'deleteRows',
	'sort',
	'autoFilter',
] as const

export function handleFreezePane(
	workbook: Workbook,
	op: Extract<Operation, { op: 'freezePane' }>,
): Result<PatchResult> {
	const rowError = validateAxisIndex('row', op.row)
	if (rowError) return err(rowError)
	const colError = validateAxisIndex('column', op.col)
	if (colError) return err(colError)
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.frozenRows = op.row
	result.value.frozenCols = op.col
	return ok(patch([], [op.sheet]))
}

export function handleSetColWidth(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setColWidth' }>,
): Result<PatchResult> {
	const validation = validateAxisIndex('column', op.col)
	if (validation) return err(validation)
	if (!Number.isFinite(op.width) || op.width < 0) {
		return err(ascendError('VALIDATION_ERROR', 'column width must be a non-negative finite number'))
	}
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	sheet.colWidths.set(op.col, op.width)
	setColDefWidth(sheet, op.col, op.width)
	return ok(patch([], [op.sheet]))
}

export function handleSetRowHeight(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRowHeight' }>,
): Result<PatchResult> {
	const validation = validateAxisIndex('row', op.row)
	if (validation) return err(validation)
	if (!Number.isFinite(op.height) || op.height < 0) {
		return err(ascendError('VALIDATION_ERROR', 'row height must be a non-negative finite number'))
	}
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	sheet.rowHeights.set(op.row, op.height)
	const rowDef = sheet.rowDefs.get(op.row)
	if (rowDef?.customHeight === false) {
		sheet.rowDefs.set(op.row, { ...rowDef, customHeight: true })
	}
	return ok(patch([], [op.sheet]))
}

export function handleHideRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideRows' }>,
): Result<PatchResult> {
	const validation = validateAxisSpan('row', op.at, op.count)
	if (validation) return err(validation)
	const hiddenValidation = validateOptionalBoolean('hideRows hidden', op.hidden)
	if (hiddenValidation) return err(hiddenValidation)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const hidden = op.hidden ?? true
	for (let r = op.at; r < op.at + op.count; r++) {
		setRowHidden(sheet, r, hidden)
	}
	return ok(patch([], [op.sheet]))
}

function setRowHidden(sheet: Sheet, row: number, hidden: boolean): void {
	const existing = sheet.rowDefs.get(row)
	if (hidden) {
		sheet.rowDefs.set(row, { ...(existing ?? {}), hidden: true })
		return
	}
	if (!existing) return
	const { hidden: _hidden, ...next } = existing
	if (Object.keys(next).length === 0) sheet.rowDefs.delete(row)
	else sheet.rowDefs.set(row, next)
}

export function handleHideCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideCols' }>,
): Result<PatchResult> {
	const validation = validateAxisSpan('column', op.at, op.count)
	if (validation) return err(validation)
	const hiddenValidation = validateOptionalBoolean('hideCols hidden', op.hidden)
	if (hiddenValidation) return err(hiddenValidation)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const hidden = op.hidden ?? true
	for (let c = op.at; c < op.at + op.count; c++) {
		setColHidden(sheet, c, hidden)
	}
	return ok(patch([], [op.sheet]))
}

function setColHidden(sheet: Sheet, col: number, hidden: boolean): void {
	const idx = findColDefIndex(sheet, col)
	const existing = idx >= 0 ? sheet.colDefs[idx] : undefined
	if (existing?.hidden === hidden) return
	if (!hidden && existing?.hidden === undefined) return
	if (!existing) {
		if (hidden) {
			insertColDef(sheet, {
				min: col,
				max: col,
				...publicColWidthMetadata(sheet, col),
				hidden: true,
			})
		}
		return
	}

	const replacements: Sheet['colDefs'] = []
	if (existing.min < col) replacements.push({ ...existing, max: col - 1 })
	const target = {
		...existing,
		...(existing.width === undefined ? publicColWidthMetadata(sheet, col) : {}),
		min: col,
		max: col,
	}
	if (hidden) {
		replacements.push({ ...target, hidden: true })
	} else {
		const { hidden: _hidden, ...next } = target
		if (!isEmptyColDef(next)) replacements.push(next)
	}
	if (col < existing.max) replacements.push({ ...existing, min: col + 1 })
	sheet.colDefs.splice(idx, 1, ...replacements)
	mergeAdjacentColDefs(sheet)
}

function publicColWidthMetadata(sheet: Sheet, col: number): Partial<Sheet['colDefs'][number]> {
	const width = sheet.colWidths.get(col)
	return width === undefined ? {} : { width, customWidth: true }
}

function setColDefWidth(sheet: Sheet, col: number, width: number): void {
	if (sheet.colDefs.length === 0) return
	const idx = findColDefIndex(sheet, col)
	const existing = idx >= 0 ? sheet.colDefs[idx] : undefined
	if (!existing) {
		insertColDef(sheet, { min: col, max: col, width, customWidth: true })
		return
	}
	if (
		existing.min === col &&
		existing.max === col &&
		existing.width === width &&
		existing.customWidth === true
	) {
		return
	}
	const replacements: Sheet['colDefs'] = []
	if (existing.min < col) replacements.push({ ...existing, max: col - 1 })
	replacements.push({ ...existing, min: col, max: col, width, customWidth: true })
	if (col < existing.max) replacements.push({ ...existing, min: col + 1 })
	sheet.colDefs.splice(idx, 1, ...replacements)
	mergeAdjacentColDefs(sheet)
}

function findColDefIndex(sheet: Sheet, col: number): number {
	const exact = sheet.colDefs.findIndex((def) => def.min === col && def.max === col)
	if (exact >= 0) return exact
	return sheet.colDefs.findIndex((def) => def.min <= col && def.max >= col)
}

function insertColDef(sheet: Sheet, def: Sheet['colDefs'][number]): void {
	const idx = sheet.colDefs.findIndex((existing) => existing.min > def.min)
	if (idx >= 0) sheet.colDefs.splice(idx, 0, def)
	else sheet.colDefs.push(def)
	mergeAdjacentColDefs(sheet)
}

function mergeAdjacentColDefs(sheet: Sheet): void {
	for (let i = 1; i < sheet.colDefs.length; i++) {
		const previous = sheet.colDefs[i - 1]
		const current = sheet.colDefs[i]
		if (!previous || !current) continue
		if (previous.max + 1 !== current.min || !sameColDefMetadata(previous, current)) continue
		sheet.colDefs.splice(i - 1, 2, { ...previous, max: current.max })
		i--
	}
}

function sameColDefMetadata(
	left: Sheet['colDefs'][number],
	right: Sheet['colDefs'][number],
): boolean {
	const { min: _leftMin, max: _leftMax, ...leftMetadata } = left
	const { min: _rightMin, max: _rightMax, ...rightMetadata } = right
	return JSON.stringify(leftMetadata) === JSON.stringify(rightMetadata)
}

function isEmptyColDef(def: Sheet['colDefs'][number]): boolean {
	return Object.keys(def).every((key) => key === 'min' || key === 'max')
}

function validateAxisIndex(label: 'row' | 'column', index: number) {
	if (Number.isInteger(index) && index >= 0) return null
	return ascendError('VALIDATION_ERROR', `${label} index must be a non-negative integer`)
}

function validateAxisSpan(label: 'row' | 'column', at: number, count: number) {
	const indexError = validateAxisIndex(label, at)
	if (indexError) return indexError
	if (Number.isInteger(count) && count > 0) return null
	return ascendError('VALIDATION_ERROR', `${label} count must be a positive integer`)
}

function validateOptionalBoolean(label: string, value: boolean | undefined) {
	if (value === undefined || typeof value === 'boolean') return null
	return ascendError('VALIDATION_ERROR', `${label} must be boolean`, {
		suggestedFix: `Set ${label}=true or ${label}=false.`,
	})
}

function validateSheetPosition(label: string, position: number, maxInclusive: number) {
	if (Number.isInteger(position) && position >= 0 && position <= maxInclusive) return null
	return ascendError(
		'VALIDATION_ERROR',
		`${label} must be an integer between 0 and ${maxInclusive}`,
	)
}

export function handleSetWorkbookProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookProtection' }>,
): Result<PatchResult> {
	const validation = validateWorkbookProtectionInput(op.protection)
	if (validation) return err(validation)
	workbook.workbookProtection = { ...op.protection }
	return ok(patch([], []))
}

function validateWorkbookProtectionInput(
	protection: Extract<Operation, { op: 'setWorkbookProtection' }>['protection'],
) {
	if (!protection || typeof protection !== 'object' || Array.isArray(protection)) {
		return ascendError('VALIDATION_ERROR', 'setWorkbookProtection requires a protection object', {
			suggestedFix: 'Provide workbook protection metadata such as { "lockStructure": true }.',
		})
	}
	for (const field of ['lockStructure', 'lockWindows', 'lockRevision'] as const) {
		const value = protection[field]
		if (value !== undefined && typeof value !== 'boolean') {
			return ascendError('VALIDATION_ERROR', `workbook protection ${field} must be boolean`, {
				suggestedFix: `Set ${field}=true or ${field}=false.`,
			})
		}
	}
	for (const field of [
		'workbookPassword',
		'revisionsPassword',
		'workbookAlgorithmName',
		'workbookHashValue',
		'workbookSaltValue',
		'revisionsAlgorithmName',
		'revisionsHashValue',
		'revisionsSaltValue',
	] as const) {
		const value = protection[field]
		if (value !== undefined && typeof value !== 'string') {
			return ascendError('VALIDATION_ERROR', `workbook protection ${field} must be a string`, {
				suggestedFix: `Use a string value for ${field}.`,
			})
		}
	}
	for (const field of ['workbookSpinCount', 'revisionsSpinCount'] as const) {
		const value = protection[field]
		if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
			return ascendError(
				'VALIDATION_ERROR',
				`workbook protection ${field} must be a non-negative integer`,
				{
					suggestedFix: `Use a non-negative integer value for ${field}.`,
				},
			)
		}
	}
	return null
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
		const pivot = workbook.pivotTables[index]
		if (pivot && sameSheetName(pivot.sheetName, sheetName)) {
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

function removeChartsForDeletedSheet(workbook: Workbook, sheetName: string): void {
	for (let index = workbook.chartParts.length - 1; index >= 0; index--) {
		const chart = workbook.chartParts[index]
		if (chart?.sheetName && sameSheetName(chart.sheetName, sheetName)) {
			workbook.chartParts.splice(index, 1)
		}
	}
}

function materializeWorkbookFormulaBindingsForRename(
	workbook: Workbook,
	renamedSheet: Sheet,
	newSheetName: string,
	affected: Set<string>,
	sheetsModified: Set<string>,
): void {
	for (const formulaSheet of workbook.sheets) {
		if (formulaSheet.cells.formulaInfoCellCount() === 0) continue
		const refs: Array<{ readonly row: number; readonly col: number }> = []
		for (const [row, col, cell] of formulaSheet.cells.iterate()) {
			if (cell.formulaInfo?.kind === 'shared') refs.push({ row, col })
		}
		const sheetName = formulaSheet === renamedSheet ? newSheetName : formulaSheet.name
		for (const ref of materializeFormulaBindingGroupsForRefs(workbook, formulaSheet, refs)) {
			affected.add(`${sheetName}!${ref}`)
			sheetsModified.add(sheetName)
		}
		retargetFormulaBindingsForRename(
			formulaSheet,
			renamedSheet.name,
			newSheetName,
			sheetName,
			affected,
		)
	}
}

function retargetFormulaBindingsForRename(
	sheet: Sheet,
	oldSheetName: string,
	newSheetName: string,
	affectedSheetName: string,
	affected: Set<string>,
): void {
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (!cell.formulaInfo) continue
		const formulaInfo = retargetCopiedFormulaBinding(cell.formulaInfo, oldSheetName, newSheetName)
		if (formulaInfo === cell.formulaInfo) continue
		sheet.cells.set(row, col, { ...cell, formulaInfo })
		affected.add(`${affectedSheetName}!${toA1({ row, col })}`)
	}
}

function retargetCopiedSheetFormulaBindings(
	sheet: Sheet,
	sourceSheetName: string,
	newSheetName: string,
): void {
	for (const [row, col, cell] of sheet.cells.iterate()) {
		const formula =
			cell.formula === null
				? null
				: (rewriteFormulaTextForRename(cell.formula, sourceSheetName, newSheetName) ?? cell.formula)
		const formulaInfo = retargetCopiedFormulaBinding(
			cell.formulaInfo,
			sourceSheetName,
			newSheetName,
		)
		if (formula === cell.formula && formulaInfo === cell.formulaInfo) continue
		sheet.cells.set(row, col, { ...cell, formula, formulaInfo })
	}
}

function retargetCopiedFormulaBinding(
	binding: CellFormulaBinding | undefined,
	sourceSheetName: string,
	newSheetName: string,
): CellFormulaBinding | undefined {
	if (!binding) return binding
	const rewrite = (ref: string | undefined) =>
		rewriteFormulaTextForRename(ref, sourceSheetName, newSheetName) ?? ref
	switch (binding.kind) {
		case 'shared': {
			const masterRef = rewrite(binding.masterRef)
			const ref = rewrite(binding.ref)
			return masterRef === binding.masterRef && ref === binding.ref
				? binding
				: {
						...binding,
						...(masterRef !== undefined ? { masterRef } : {}),
						...(ref !== undefined ? { ref } : {}),
					}
		}
		case 'array': {
			const ref = rewrite(binding.ref)
			return ref === binding.ref ? binding : { ...binding, ...(ref !== undefined ? { ref } : {}) }
		}
		case 'dataTable': {
			const ref = rewrite(binding.ref)
			const r1 = rewrite(binding.r1)
			const r2 = rewrite(binding.r2)
			return ref === binding.ref && r1 === binding.r1 && r2 === binding.r2
				? binding
				: {
						...binding,
						...(ref !== undefined ? { ref } : {}),
						...(r1 !== undefined ? { r1 } : {}),
						...(r2 !== undefined ? { r2 } : {}),
					}
		}
		case 'spill':
		case 'blockedSpill': {
			const anchorRef = rewrite(binding.anchorRef) ?? binding.anchorRef
			const ref = rewrite(binding.ref) ?? binding.ref
			if (binding.kind === 'spill') {
				return anchorRef === binding.anchorRef && ref === binding.ref
					? binding
					: { ...binding, anchorRef, ref }
			}
			const blockingRefs = binding.blockingRefs.map(
				(blockingRef) => rewrite(blockingRef) ?? blockingRef,
			)
			const changed =
				anchorRef !== binding.anchorRef ||
				ref !== binding.ref ||
				blockingRefs.some((blockingRef, index) => blockingRef !== binding.blockingRefs[index])
			return changed ? { ...binding, anchorRef, ref, blockingRefs } : binding
		}
		case 'dynamicArray':
			return binding
	}
}

function rewriteChartSheetReferencesForRename(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	for (let index = 0; index < workbook.chartParts.length; index++) {
		const chart = workbook.chartParts[index]
		if (!chart) continue
		workbook.chartParts[index] = {
			...chart,
			...(chart.sheetName !== undefined && sameSheetName(chart.sheetName, oldName)
				? { sheetName: newName }
				: {}),
			series: chart.series.map((series) => ({
				...series,
				...(series.nameRef !== undefined
					? { nameRef: rewriteChartSeriesRef(series.nameRef, oldName, newName) }
					: {}),
				...(series.categoryRef !== undefined
					? { categoryRef: rewriteChartSeriesRef(series.categoryRef, oldName, newName) }
					: {}),
				...(series.valueRef !== undefined
					? { valueRef: rewriteChartSeriesRef(series.valueRef, oldName, newName) }
					: {}),
			})),
		}
	}
}

function cloneChartsForCopiedSheet(
	workbook: Workbook,
	sourceSheetName: string,
	newSheetName: string,
): Map<string, string> {
	const pathMap = new Map<string, string>()
	const copiedCharts = workbook.chartParts
		.filter(
			(chart) => chart.sheetName !== undefined && sameSheetName(chart.sheetName, sourceSheetName),
		)
		.map((chart) => {
			const partPath = nextChartPartPath(workbook, pathMap)
			pathMap.set(chart.partPath, partPath)
			return {
				...chart,
				partPath,
				sheetName: newSheetName,
				series: chart.series.map((series) => ({
					...series,
					...(series.nameRef !== undefined
						? {
								nameRef: rewriteChartSeriesRef(series.nameRef, sourceSheetName, newSheetName),
							}
						: {}),
					...(series.categoryRef !== undefined
						? {
								categoryRef: rewriteChartSeriesRef(
									series.categoryRef,
									sourceSheetName,
									newSheetName,
								),
							}
						: {}),
					...(series.valueRef !== undefined
						? {
								valueRef: rewriteChartSeriesRef(series.valueRef, sourceSheetName, newSheetName),
							}
						: {}),
				})),
			}
		})
	workbook.chartParts.push(...copiedCharts)
	return pathMap
}

function sameSheetName(left: string, right: string): boolean {
	return left.toLowerCase() === right.toLowerCase()
}

function rewriteChartSeriesRef(ref: string, oldName: string, newName: string): string {
	return rewriteFormulaTextForRename(ref, oldName, newName) ?? ref
}

function retargetCopiedSheetDrawingParts(
	workbook: Workbook,
	sheet: Workbook['sheets'][number],
): void {
	const drawingPartPaths = new Set<string>()
	for (const image of sheet.imageRefs) drawingPartPaths.add(image.drawingPartPath)
	for (const object of sheet.drawingObjectRefs) drawingPartPaths.add(object.drawingPartPath)
	if (drawingPartPaths.size === 0) return

	const partPathMap = new Map<string, string>()
	for (const partPath of drawingPartPaths) {
		partPathMap.set(partPath, nextDrawingPartPath(workbook, partPathMap))
	}
	for (let index = 0; index < sheet.imageRefs.length; index++) {
		const image = sheet.imageRefs[index]
		if (!image) continue
		sheet.imageRefs[index] = {
			...image,
			drawingPartPath: partPathMap.get(image.drawingPartPath) ?? image.drawingPartPath,
		}
	}
	for (let index = 0; index < sheet.drawingObjectRefs.length; index++) {
		const object = sheet.drawingObjectRefs[index]
		if (!object) continue
		sheet.drawingObjectRefs[index] = {
			...object,
			drawingPartPath: partPathMap.get(object.drawingPartPath) ?? object.drawingPartPath,
		}
	}
}

function retargetCopiedSheetImageTargets(
	workbook: Workbook,
	sheet: Workbook['sheets'][number],
): void {
	const targetPathMap = new Map<string, string>()
	for (let index = 0; index < sheet.imageRefs.length; index++) {
		const image = sheet.imageRefs[index]
		if (!image?.content || !image.contentType) continue
		let targetPath = targetPathMap.get(image.targetPath)
		if (!targetPath) {
			targetPath = nextImageTargetPath(
				workbook,
				targetPathMap,
				imageExtensionForContentType(image.contentType),
			)
			targetPathMap.set(image.targetPath, targetPath)
		}
		sheet.imageRefs[index] = { ...image, targetPath }
	}
}

function retargetCopiedSheetChartRelationships(
	sheet: Workbook['sheets'][number],
	chartPartPaths: ReadonlyMap<string, string>,
): void {
	if (chartPartPaths.size === 0) return
	for (let index = 0; index < sheet.drawingObjectRefs.length; index++) {
		const object = sheet.drawingObjectRefs[index]
		if (!object?.relationshipRefs) continue
		const relationshipRefs = object.relationshipRefs.map((relationship) => {
			const partPath = findRetargetedChartPath(relationship.target, chartPartPaths)
			return partPath
				? { ...relationship, target: chartRelationshipTarget(partPath) }
				: relationship
		})
		sheet.drawingObjectRefs[index] = { ...object, relationshipRefs }
	}
}

function findRetargetedChartPath(
	target: string,
	chartPartPaths: ReadonlyMap<string, string>,
): string | undefined {
	for (const [oldPath, newPath] of chartPartPaths) {
		if (chartRelationshipTargetMatches(target, oldPath)) return newPath
	}
	return undefined
}

function chartRelationshipTargetMatches(target: string, chartPartPath: string): boolean {
	const normalizedTarget = target.replace(/^\/+/, '')
	const normalizedPartPath = chartPartPath.replace(/^\/+/, '')
	if (normalizedTarget === normalizedPartPath) return true
	const fileName = normalizedPartPath.slice(normalizedPartPath.lastIndexOf('/') + 1)
	return (
		normalizedTarget === `../charts/${fileName}` ||
		normalizedTarget === `charts/${fileName}` ||
		normalizedTarget.endsWith(`/charts/${fileName}`)
	)
}

function chartRelationshipTarget(chartPartPath: string): string {
	const fileName = chartPartPath.slice(chartPartPath.lastIndexOf('/') + 1)
	return `../charts/${fileName}`
}

function nextChartPartPath(workbook: Workbook, pendingPaths: ReadonlyMap<string, string>): string {
	const used = new Set(workbook.chartParts.map((chart) => chart.partPath))
	for (const path of pendingPaths.values()) used.add(path)
	let index = used.size + 1
	while (used.has(`xl/charts/chart${index}.xml`)) index++
	return `xl/charts/chart${index}.xml`
}

function nextDrawingPartPath(
	workbook: Workbook,
	pendingPaths: ReadonlyMap<string, string>,
): string {
	const used = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const image of sheet.imageRefs) used.add(image.drawingPartPath)
		for (const object of sheet.drawingObjectRefs) used.add(object.drawingPartPath)
	}
	for (const path of pendingPaths.values()) used.add(path)
	let index = used.size + 1
	while (used.has(`xl/drawings/drawing${index}.xml`)) index++
	return `xl/drawings/drawing${index}.xml`
}

function nextImageTargetPath(
	workbook: Workbook,
	pendingPaths: ReadonlyMap<string, string>,
	extension: string,
): string {
	const used = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const image of sheet.imageRefs) used.add(image.targetPath)
	}
	for (const path of pendingPaths.values()) used.add(path)
	let index = used.size + 1
	while (used.has(`xl/media/image${index}.${extension}`)) index++
	return `xl/media/image${index}.${extension}`
}

function imageExtensionForContentType(contentType: string): string {
	switch (contentType) {
		case 'image/jpeg':
		case 'image/jpg':
			return 'jpg'
		case 'image/gif':
			return 'gif'
		case 'image/bmp':
			return 'bmp'
		case 'image/webp':
			return 'webp'
		default:
			return 'png'
	}
}
