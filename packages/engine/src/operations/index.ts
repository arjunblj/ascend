import type { Workbook } from '@ascend/core'
import type { AscendError, Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import {
	invalidateWorkbookAnalysis,
	patchWorkbookAnalysis,
	shiftWorkbookAnalysisForAxis,
} from '../analysis.ts'
import { invalidateSheetIndexCache } from '../evaluator.ts'
import * as cellOps from './cell-ops.ts'
import * as formatOps from './format-ops.ts'
import {
	operationAffectsFormulas,
	patch,
	resolveAffectedCellKeys,
	resolvePatchResultCellKeys,
} from './helpers.ts'
import * as sheetOps from './sheet-ops.ts'
import * as structuralOps from './structural-ops.ts'
import * as tableOps from './table-ops.ts'

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
	deleteComment: formatOps.handleDeleteComment as OperationHandler,
	setHyperlink: formatOps.handleSetHyperlink as OperationHandler,
	deleteHyperlink: formatOps.handleDeleteHyperlink as OperationHandler,
	setDefinedName: formatOps.handleSetDefinedName as OperationHandler,
	deleteDefinedName: formatOps.handleDeleteDefinedName as OperationHandler,
	groupRows: formatOps.handleGroupRows as OperationHandler,
	groupCols: formatOps.handleGroupCols as OperationHandler,
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

	const allAffected: string[] = []
	const allSheets: string[] = []
	const warnings: AscendError[] = []
	let needsRecalc = false

	for (const op of ops) {
		const result = applyOperation(workbook, op)
		if (!result.ok) return result
		allAffected.push(...result.value.affectedCells)
		allSheets.push(...result.value.sheetsModified)
		if (result.value.recalcRequired) needsRecalc = true
		if (result.value.warnings) warnings.push(...result.value.warnings)
	}

	return ok(patch(allAffected, allSheets, needsRecalc, warnings.length > 0 ? warnings : undefined))
}
