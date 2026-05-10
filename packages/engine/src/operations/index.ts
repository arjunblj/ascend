import { cloneActiveContentInfo, type Workbook } from '@ascend/core'
import type { AscendError, CellUpdate, Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import {
	invalidateWorkbookAnalysis,
	patchWorkbookAnalysis,
	shiftWorkbookAnalysisForAxis,
} from '../analysis.ts'
import { invalidateSheetIndexCache } from '../evaluator.ts'
import * as cellOps from './cell-ops.ts'
import * as connectionOps from './connection-ops.ts'
import * as formatOps from './format-ops.ts'
import {
	operationAffectsFormulas,
	patch,
	resolveAffectedCellKeys,
	resolvePatchResultCellKeys,
} from './helpers.ts'
import * as pivotOps from './pivot-ops.ts'
import * as sheetOps from './sheet-ops.ts'
import * as slicerOps from './slicer-ops.ts'
import * as structuralOps from './structural-ops.ts'
import * as tableOps from './table-ops.ts'
import * as visualOps from './visual-ops.ts'
import * as workbookOps from './workbook-ops.ts'

export type { PatchResult } from './helpers.ts'

export interface ApplyOperationsOptions {
	readonly collectAllErrors?: boolean
}

export interface ApplyOperationsErrors {
	readonly errors: readonly AscendError[]
}

type OperationHandler = (
	workbook: Workbook,
	op: never,
) => Result<import('./helpers.ts').PatchResult>

const handlers: Record<string, OperationHandler> = {
	setCells: cellOps.handleSetCells as OperationHandler,
	setFormula: cellOps.handleSetFormula as OperationHandler,
	fillFormula: cellOps.handleFillFormula as OperationHandler,
	setRichText: cellOps.handleSetRichText as OperationHandler,
	clearRange: cellOps.handleClearRange as OperationHandler,
	setStyle: cellOps.handleSetStyle as OperationHandler,
	insertRows: structuralOps.handleInsertRows as OperationHandler,
	deleteRows: structuralOps.handleDeleteRows as OperationHandler,
	insertCols: structuralOps.handleInsertCols as OperationHandler,
	deleteCols: structuralOps.handleDeleteCols as OperationHandler,
	copyRange: structuralOps.handleTransferRange as OperationHandler,
	moveRange: structuralOps.handleTransferRange as OperationHandler,
	mergeCells: structuralOps.handleMergeCells as OperationHandler,
	unmergeCells: structuralOps.handleUnmergeCells as OperationHandler,
	addSheet: sheetOps.handleAddSheet as OperationHandler,
	deleteSheet: sheetOps.handleDeleteSheet as OperationHandler,
	renameSheet: sheetOps.handleRenameSheet as OperationHandler,
	moveSheet: sheetOps.handleMoveSheet as OperationHandler,
	copySheet: sheetOps.handleCopySheet as OperationHandler,
	hideSheet: sheetOps.handleHideSheet as OperationHandler,
	setTabColor: sheetOps.handleSetTabColor as OperationHandler,
	setSheetProtection: sheetOps.handleSetSheetProtection as OperationHandler,
	freezePane: sheetOps.handleFreezePane as OperationHandler,
	setColWidth: sheetOps.handleSetColWidth as OperationHandler,
	setRowHeight: sheetOps.handleSetRowHeight as OperationHandler,
	hideRows: sheetOps.handleHideRows as OperationHandler,
	hideCols: sheetOps.handleHideCols as OperationHandler,
	createTable: tableOps.handleCreateTable as OperationHandler,
	appendRows: tableOps.handleAppendRows as OperationHandler,
	sortRange: tableOps.handleSortRange as OperationHandler,
	setNumberFormat: formatOps.handleSetNumberFormat as OperationHandler,
	setConditionalFormat: formatOps.handleSetConditionalFormat as OperationHandler,
	deleteConditionalFormat: formatOps.handleDeleteConditionalFormat as OperationHandler,
	setDataValidation: formatOps.handleSetDataValidation as OperationHandler,
	deleteDataValidation: formatOps.handleDeleteDataValidation as OperationHandler,
	setAutoFilter: formatOps.handleSetAutoFilter as OperationHandler,
	clearAutoFilter: formatOps.handleClearAutoFilter as OperationHandler,
	setPageSetup: formatOps.handleSetPageSetup as OperationHandler,
	setPrintArea: formatOps.handleSetPrintArea as OperationHandler,
	setComment: formatOps.handleSetComment as OperationHandler,
	setThreadedComment: formatOps.handleSetThreadedComment as OperationHandler,
	deleteComment: formatOps.handleDeleteComment as OperationHandler,
	setHyperlink: formatOps.handleSetHyperlink as OperationHandler,
	deleteHyperlink: formatOps.handleDeleteHyperlink as OperationHandler,
	setDefinedName: formatOps.handleSetDefinedName as OperationHandler,
	deleteDefinedName: formatOps.handleDeleteDefinedName as OperationHandler,
	groupRows: formatOps.handleGroupRows as OperationHandler,
	groupCols: formatOps.handleGroupCols as OperationHandler,
	setWorkbookProtection: sheetOps.handleSetWorkbookProtection as OperationHandler,
	setWorkbookProperties: workbookOps.handleSetWorkbookProperties as OperationHandler,
	setWorkbookView: workbookOps.handleSetWorkbookView as OperationHandler,
	setCalcSettings: workbookOps.handleSetCalcSettings as OperationHandler,
	setTheme: workbookOps.handleSetTheme as OperationHandler,
	deleteTable: tableOps.handleDeleteTable as OperationHandler,
	renameTable: tableOps.handleRenameTable as OperationHandler,
	resizeTable: tableOps.handleResizeTable as OperationHandler,
	setTableColumn: tableOps.handleSetTableColumn as OperationHandler,
	replaceImage: visualOps.handleReplaceImage as OperationHandler,
	setChartSeriesSource: visualOps.handleSetChartSeriesSource as OperationHandler,
	setPivotCache: pivotOps.handleSetPivotCache as OperationHandler,
	setPivotFieldItem: pivotOps.handleSetPivotFieldItem as OperationHandler,
	setSlicerCacheItem: slicerOps.handleSetSlicerCacheItem as OperationHandler,
	setConnectionRefresh: connectionOps.handleSetConnectionRefresh as OperationHandler,
	rewriteExternalLink: workbookOps.handleRewriteExternalLink as OperationHandler,
	insertImage: visualOps.handleInsertImage as OperationHandler,
	deleteImage: visualOps.handleDeleteImage as OperationHandler,
	setDrawingText: visualOps.handleSetDrawingText as OperationHandler,
}

export function applyOperation(
	workbook: Workbook,
	op: Operation,
): Result<import('./helpers.ts').PatchResult> {
	const warnings: AscendError[] = []
	const useIncrementalPatch =
		op.op === 'setFormula' ||
		op.op === 'fillFormula' ||
		op.op === 'copyRange' ||
		op.op === 'moveRange' ||
		op.op === 'insertRows' ||
		op.op === 'deleteRows' ||
		op.op === 'insertCols' ||
		op.op === 'deleteCols'
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
		switch (op.op) {
			case 'insertRows':
				shiftWorkbookAnalysisForAxis(workbook, op.sheet, 'row', op.at, op.count)
				break
			case 'deleteRows':
				shiftWorkbookAnalysisForAxis(workbook, op.sheet, 'row', op.at, -op.count)
				break
			case 'insertCols':
				shiftWorkbookAnalysisForAxis(workbook, op.sheet, 'col', op.at, op.count)
				break
			case 'deleteCols':
				shiftWorkbookAnalysisForAxis(workbook, op.sheet, 'col', op.at, -op.count)
				break
			case 'copyRange':
			case 'moveRange': {
				const changedKeys = resolvePatchResultCellKeys(
					workbook,
					op.sheet,
					result.value.affectedCells,
					warnings,
				)
				if (changedKeys.length > 0) patchWorkbookAnalysis(workbook, changedKeys)
				break
			}
			case 'setFormula':
			case 'fillFormula': {
				const changedKeys = resolveAffectedCellKeys(workbook, op)
				if (changedKeys.length > 0) patchWorkbookAnalysis(workbook, changedKeys)
				break
			}
		}
	}
	if (result.ok && warnings.length > 0) {
		return ok({
			...result.value,
			warnings:
				result.value.warnings && result.value.warnings.length > 0
					? [...result.value.warnings, ...warnings]
					: warnings,
		})
	}
	return result
}

function coalesceSetCells(ops: readonly Operation[]): Operation[] {
	const out: Operation[] = []
	let pending: { sheet: string; updates: CellUpdate[] } | null = null

	for (const op of ops) {
		if (op.op === 'setCells') {
			if (pending && pending.sheet === op.sheet) {
				pending.updates = [...pending.updates, ...op.updates]
			} else {
				if (pending) {
					out.push({ op: 'setCells', sheet: pending.sheet, updates: pending.updates })
				}
				pending = { sheet: op.sheet, updates: [...op.updates] }
			}
		} else {
			if (pending) {
				out.push({ op: 'setCells', sheet: pending.sheet, updates: pending.updates })
				pending = null
			}
			out.push(op)
		}
	}
	if (pending) {
		out.push({ op: 'setCells', sheet: pending.sheet, updates: pending.updates })
	}
	return out
}

export function applyOperations(
	workbook: Workbook,
	ops: readonly Operation[],
	options?: ApplyOperationsOptions,
): Result<import('./helpers.ts').PatchResult, AscendError | ApplyOperationsErrors> {
	const collectAllErrors = options?.collectAllErrors ?? false

	if (collectAllErrors) {
		const errors: AscendError[] = []
		for (const op of ops) {
			const clone = workbook.clone()
			const result = applyOperation(clone, op)
			if (!result.ok) errors.push(result.error)
		}
		if (errors.length > 0) return err({ errors })
	}

	const coalesced = coalesceSetCells(ops)

	const allAffected: string[] = []
	const allSheets: string[] = []
	const warnings: AscendError[] = []
	let needsRecalc = false

	for (const op of coalesced) {
		const result = applyOperation(workbook, op)
		if (!result.ok) return result
		allAffected.push(...result.value.affectedCells)
		allSheets.push(...result.value.sheetsModified)
		if (result.value.recalcRequired) needsRecalc = true
		if (result.value.warnings) warnings.push(...result.value.warnings)
	}

	return ok(patch(allAffected, allSheets, needsRecalc, warnings.length > 0 ? warnings : undefined))
}

function restoreWorkbookFromSnapshot(workbook: Workbook, snapshot: Workbook): void {
	workbook.sheets.splice(0, workbook.sheets.length, ...snapshot.sheets)
	workbook.invalidateSheetCache()
	workbook.definedNames.copyFrom(snapshot.definedNames)
	workbook.styles.copyFrom(snapshot.styles)
	workbook.differentialStyles.splice(0, workbook.differentialStyles.length)
	workbook.differentialStyles.push(...snapshot.differentialStyles)
	workbook.pivotCaches.splice(0, workbook.pivotCaches.length)
	workbook.pivotCaches.push(
		...snapshot.pivotCaches.map((e) => ({
			...e,
			fields: e.fields.map((field) => ({
				...field,
				...(field.sharedItems
					? { sharedItems: field.sharedItems.map((item) => ({ ...item })) }
					: {}),
			})),
		})),
	)
	workbook.pivotTables.splice(0, workbook.pivotTables.length)
	workbook.pivotTables.push(
		...snapshot.pivotTables.map((e) => ({
			...e,
			fields: e.fields.map((field) => ({
				...field,
				...(field.items ? { items: field.items.map((item) => ({ ...item })) } : {}),
			})),
			rowFields: e.rowFields.map((field) => ({ ...field })),
			columnFields: e.columnFields.map((field) => ({ ...field })),
			pageFields: e.pageFields.map((field) => ({ ...field })),
			dataFields: e.dataFields.map((field) => ({ ...field })),
		})),
	)
	workbook.slicerCaches.splice(0, workbook.slicerCaches.length)
	workbook.slicerCaches.push(
		...snapshot.slicerCaches.map((e) => ({
			...e,
			pivotTableNames: [...e.pivotTableNames],
			...(e.items ? { items: e.items.map((item) => ({ ...item })) } : {}),
		})),
	)
	workbook.slicers.splice(0, workbook.slicers.length)
	workbook.slicers.push(...snapshot.slicers.map((e) => ({ ...e })))
	workbook.timelineCaches.splice(0, workbook.timelineCaches.length)
	workbook.timelineCaches.push(
		...snapshot.timelineCaches.map((e) => ({ ...e, pivotTableNames: [...e.pivotTableNames] })),
	)
	workbook.timelines.splice(0, workbook.timelines.length)
	workbook.timelines.push(...snapshot.timelines.map((e) => ({ ...e })))
	workbook.chartParts.splice(0, workbook.chartParts.length)
	workbook.chartParts.push(
		...snapshot.chartParts.map((entry) => ({
			...entry,
			series: entry.series.map((series) => ({ ...series })),
		})),
	)
	workbook.activeContent.splice(0, workbook.activeContent.length)
	workbook.activeContent.push(...snapshot.activeContent.map(cloneActiveContentInfo))
	workbook.workbookViews.splice(0, workbook.workbookViews.length)
	workbook.workbookViews.push(...snapshot.workbookViews.map((v) => ({ ...v })))
	workbook.externalReferences.splice(0, workbook.externalReferences.length)
	workbook.externalReferences.push(...snapshot.externalReferences)
	workbook.externalReferenceDetails.splice(0, workbook.externalReferenceDetails.length)
	workbook.externalReferenceDetails.push(
		...snapshot.externalReferenceDetails.map((entry) => ({ ...entry })),
	)
	workbook.connectionParts.splice(0, workbook.connectionParts.length)
	workbook.connectionParts.push(...snapshot.connectionParts.map((entry) => ({ ...entry })))
	workbook.dataModelParts.splice(0, workbook.dataModelParts.length)
	workbook.dataModelParts.push(...snapshot.dataModelParts.map((entry) => ({ ...entry })))
	workbook.workbookProperties = { ...snapshot.workbookProperties }
	workbook.workbookProtection = snapshot.workbookProtection
		? { ...snapshot.workbookProtection }
		: null
	workbook.styleMetadata = { ...snapshot.styleMetadata }
	workbook.themeMetadata = { ...snapshot.themeMetadata }
	workbook.themeColors.splice(0, workbook.themeColors.length)
	workbook.themeColors.push(...snapshot.themeColors.map((color) => ({ ...color })))
	workbook.preservedStyles = snapshot.preservedStyles
		? {
				...snapshot.preservedStyles,
				xfByStyleId: { ...snapshot.preservedStyles.xfByStyleId },
				...(snapshot.preservedStyles.baseStyleIdByStyleId
					? { baseStyleIdByStyleId: { ...snapshot.preservedStyles.baseStyleIdByStyleId } }
					: {}),
			}
		: null
	workbook.preservedTheme = snapshot.preservedTheme ? { ...snapshot.preservedTheme } : null
	workbook.preservedSharedStrings = snapshot.preservedSharedStrings
		? { ...snapshot.preservedSharedStrings }
		: null
	workbook.preservedMetadata = snapshot.preservedMetadata ? { ...snapshot.preservedMetadata } : null
	workbook.preservedXml = snapshot.preservedXml ? { ...snapshot.preservedXml } : null
	workbook.calcSettings = {
		...snapshot.calcSettings,
		iterativeCalc: { ...snapshot.calcSettings.iterativeCalc },
	}
	invalidateWorkbookAnalysis(workbook)
	invalidateSheetIndexCache(workbook)
}

/**
 * Applies operations atomically: all succeed or all fail. On failure, the workbook
 * is rolled back to its state before any mutations.
 */
export function applyWithTransaction(
	workbook: Workbook,
	ops: readonly Operation[],
	options?: ApplyOperationsOptions,
): Result<import('./helpers.ts').PatchResult, AscendError | ApplyOperationsErrors> {
	const collectAllErrors = options?.collectAllErrors ?? false
	if (collectAllErrors) {
		const errors: AscendError[] = []
		const validationClone = workbook.clone()
		for (const op of coalesceSetCells(ops)) {
			const result = applyOperation(validationClone, op)
			if (!result.ok) errors.push(result.error)
		}
		if (errors.length > 0) return err({ errors })
	}

	const snapshot = workbook.clone()
	const coalesced = coalesceSetCells(ops)

	const allAffected: string[] = []
	const allSheets: string[] = []
	const warnings: AscendError[] = []
	let needsRecalc = false

	for (const op of coalesced) {
		const result = applyOperation(workbook, op)
		if (!result.ok) {
			restoreWorkbookFromSnapshot(workbook, snapshot)
			return result
		}
		allAffected.push(...result.value.affectedCells)
		allSheets.push(...result.value.sheetsModified)
		if (result.value.recalcRequired) needsRecalc = true
		if (result.value.warnings) warnings.push(...result.value.warnings)
	}

	return ok(patch(allAffected, allSheets, needsRecalc, warnings.length > 0 ? warnings : undefined))
}
