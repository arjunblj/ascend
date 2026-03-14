import type { Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { invalidateSheetIndexCache } from '../evaluator.ts'
import {
	rewriteSheetMetadataFormulasForRename,
	rewriteSheetNameInDefinedNames,
	rewriteSheetNameInFormulas,
} from '../structural/formula-rewrite.ts'
import { renameHyperlinkLocation } from '../structural/sheet-topology.ts'
import type { PatchResult } from './helpers.ts'
import { clearFormulaMetadata, getSheet, patch } from './helpers.ts'

export function handleAddSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'addSheet' }>,
): Result<PatchResult> {
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
		.filter((entry) => entry.sheetName === op.sheet)
		.map((entry) => entry.name)
		.filter((name): name is string => Boolean(name))
	workbook.removeSheet(op.sheet)
	removeSheetScopedDefinedNames(workbook, targetSheet.id)
	removeWorkbookMetadataForDeletedSheet(workbook, op.sheet, removedPivotNames)
	return ok(patch([], [op.sheet]))
}

export function handleRenameSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameSheet' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value

	if (workbook.getSheet(op.newName)) {
		return err(
			ascendError('NAME_CONFLICT', `Sheet "${op.newName}" already exists`, {
				suggestedFix: 'Choose a different name; a sheet with that name already exists',
			}),
		)
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

export function handleMoveSheet(
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

export function handleCopySheet(
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

export function handleHideSheet(
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

export function handleSetTabColor(
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

export function handleSetSheetProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setSheetProtection' }>,
): Result<PatchResult> {
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

export function handleFreezePane(
	workbook: Workbook,
	op: Extract<Operation, { op: 'freezePane' }>,
): Result<PatchResult> {
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
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.colWidths.set(op.col, op.width)
	return ok(patch([], [op.sheet]))
}

export function handleSetRowHeight(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRowHeight' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	result.value.rowHeights.set(op.row, op.height)
	return ok(patch([], [op.sheet]))
}

export function handleHideRows(
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

export function handleHideCols(
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
