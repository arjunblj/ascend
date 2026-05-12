import type {
	RangeRef,
	Sheet,
	SheetConditionalFormat,
	SheetConditionalFormatValueObject,
	SheetDataValidation,
	Workbook,
} from '@ascend/core'
import { parseA1, parseRange, toA1 } from '@ascend/core'
import { cachedParseFormula } from '@ascend/formulas'
import type { Operation, PasteMode, Result } from '@ascend/schema'
import { ascendError, EMPTY, err, ok } from '@ascend/schema'
import { resolveCellFormulaText } from '../analysis.ts'
import {
	retargetExplicitFormulaSheetRefsInRange,
	rewriteDefinedNameFormulasForMove,
	rewriteDefinedNameFormulasForShift,
	rewriteWorkbookFormulasForMove,
	rewriteWorkbookFormulasForShift,
	rewriteWorkbookMetadataFormulasForMove,
} from '../structural/formula-rewrite.ts'
import { shiftSheetCellMetadata } from '../structural/sheet-topology.ts'
import {
	collectDeletedTableColumns,
	type DeletedTableColumnReference,
	findDeletedTableColumnReference,
} from '../structural/table-field-guards.ts'
import type { PatchResult } from './helpers.ts'
import {
	cellWithExisting,
	clearFormulaMetadata,
	collectRangeCells,
	createLegacyArrayFormulaIndex,
	DEFAULT_SID,
	getSheet,
	legacyArrayFormulaEditError,
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
	const formulaBindingBlocker = findWorkbookFormulaBinding(workbook)
	if (formulaBindingBlocker) return err(formulaBindingStructuralEditError(formulaBindingBlocker))
	const deletedTableColumns =
		axis === 'col' && delta < 0 ? collectDeletedTableColumns(sheet, at, count) : []
	const deletedTableColumnBlocker = findDeletedTableColumnReference(workbook, deletedTableColumns, {
		skipDeletedCells: { sheet, startCol: at, endColExclusive: at + count },
		skipDeletedTableColumnFormulas: true,
	})
	if (deletedTableColumnBlocker) {
		return err(deletedTableColumnStructuralEditError(deletedTableColumnBlocker))
	}
	const tableBoundaryBlocker =
		axis === 'row' && delta < 0 ? findPartialTableBoundaryRowDelete(sheet, at, count) : null
	if (tableBoundaryBlocker) return err(tableBoundaryRowStructuralEditError(tableBoundaryBlocker))

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
	const sourceSheet = sheetResult.value
	const targetSheetName = op.targetSheet ?? op.sheet
	const targetSheetResult =
		targetSheetName === op.sheet ? sheetResult : getSheet(workbook, targetSheetName)
	if (!targetSheetResult.ok) return targetSheetResult
	const targetSheet = targetSheetResult.value
	const crossSheet = sourceSheet !== targetSheet
	const sheetIndex = workbook.sheets.indexOf(sourceSheet)
	const sourceResult = safeParseRange(op.source)
	if (!sourceResult.ok) return sourceResult
	const targetStart = parseA1(op.target)
	const source = sourceResult.value
	const rowDelta = targetStart.row - source.start.row
	const colDelta = targetStart.col - source.start.col
	const mode = op.mode ?? 'all'
	const affected: string[] = []

	if (pasteCells(mode)) {
		const targetLegacyArrayIndex = createLegacyArrayFormulaIndex(targetSheet)
		const targetRange = shiftRange(source, rowDelta, colDelta)
		const blockedTarget = targetLegacyArrayIndex.findIntersection(targetRange)
		if (blockedTarget) {
			return err(legacyArrayFormulaEditError(blockedTarget.targetRef, blockedTarget.ref))
		}
		if (op.op === 'moveRange') {
			const sourceLegacyArrayIndex = crossSheet
				? createLegacyArrayFormulaIndex(sourceSheet)
				: targetLegacyArrayIndex
			const blockedSource = sourceLegacyArrayIndex.findIntersection(source)
			if (blockedSource) {
				return err(legacyArrayFormulaEditError(blockedSource.targetRef, blockedSource.ref))
			}
		}
	}
	const mergePlan = pasteMerges(mode)
		? planMergeTransfer(sourceSheet, targetSheet, source, rowDelta, colDelta, op.op === 'moveRange')
		: ok<MergeTransferPlan>({
				targetRange: shiftRange(source, rowDelta, colDelta),
				sourceMerges: [],
				targetMerges: [],
				targetReplacedMerges: [],
				removeSourceMerges: false,
			})
	if (!mergePlan.ok) return mergePlan

	const snapshot = collectRangeCells(sourceSheet, source)

	if (pasteCells(mode)) {
		for (const entry of snapshot) {
			const targetRow = entry.row + rowDelta
			const targetCol = entry.col + colDelta
			const existingTarget = targetSheet.cells.get(targetRow, targetCol)
			const ref = toA1({ row: targetRow, col: targetCol })

			if (!entry.cell && mode === 'all') {
				targetSheet.cells.delete(targetRow, targetCol)
				affected.push(affectedRef(targetSheet, ref, crossSheet))
				continue
			}

			const sourceFormula = entry.cell
				? resolveCellFormulaText(workbook, sheetIndex, entry.row, entry.col, entry.cell)
				: null
			const formula = translateCellFormula(sourceFormula, rowDelta, colDelta)
			const targetValue = existingTarget?.value ?? EMPTY
			const targetFormula = existingTarget?.formula ?? null
			const targetStyle = existingTarget?.styleId ?? DEFAULT_SID

			if (mode === 'values') {
				targetSheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(entry.cell?.value ?? EMPTY, null, targetStyle),
				)
			} else if (mode === 'formulas') {
				targetSheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(entry.cell?.value ?? EMPTY, formula, targetStyle),
				)
			} else if (mode === 'formats' || mode === 'styles') {
				targetSheet.cells.set(
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
				targetSheet.cells.set(
					targetRow,
					targetCol,
					cellWithExisting(entry.cell.value, formula, entry.cell.styleId),
				)
			} else {
				targetSheet.cells.delete(targetRow, targetCol)
			}
			affected.push(affectedRef(targetSheet, ref, crossSheet))
		}
	}

	copyTransferMetadata(
		sourceSheet,
		targetSheet,
		source,
		rowDelta,
		colDelta,
		mode,
		op.op === 'moveRange',
	)
	applyMergeTransfer(sourceSheet, targetSheet, mergePlan.value)

	if (op.op === 'moveRange') {
		for (const entry of snapshot) {
			if (pasteCells(mode)) sourceSheet.cells.delete(entry.row, entry.col)
			if (pasteCells(mode)) {
				affected.push(
					affectedRef(sourceSheet, toA1({ row: entry.row, col: entry.col }), crossSheet),
				)
			}
		}
		if (pasteCells(mode)) {
			rewriteWorkbookFormulasForMove(
				workbook,
				sourceSheet.name,
				targetSheet.name,
				source,
				mergePlan.value.targetRange,
			)
			rewriteDefinedNameFormulasForMove(
				workbook,
				sourceSheet.name,
				targetSheet.name,
				source,
				mergePlan.value.targetRange,
			)
			rewriteWorkbookMetadataFormulasForMove(
				workbook,
				sourceSheet.name,
				targetSheet.name,
				source,
				mergePlan.value.targetRange,
			)
		}
	}

	return ok(patch(affected, [sourceSheet.name, targetSheet.name], pasteRequiresRecalc(mode)))
}

function affectedRef(sheet: Sheet, ref: string, qualify: boolean): string {
	return qualify ? `${sheet.name}!${ref}` : ref
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

function pasteMerges(mode: PasteMode): boolean {
	return mode === 'all' || mode === 'formats' || mode === 'styles'
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

function findWorkbookFormulaBinding(
	workbook: Workbook,
): { readonly kind: string; readonly sheetName: string; readonly ref: string } | null {
	for (const sheet of workbook.sheets) {
		if (sheet.cells.formulaInfoCellCount() === 0) continue
		for (const [row, col, cell] of sheet.cells.iterate()) {
			const binding = cell.formulaInfo
			if (!binding) continue
			return { kind: binding.kind, sheetName: sheet.name, ref: toA1({ row, col }) }
		}
	}
	return null
}

function formulaBindingStructuralEditError(blocker: {
	readonly kind: string
	readonly sheetName: string
	readonly ref: string
}): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot structurally edit rows or columns because ${blocker.sheetName}!${blocker.ref} contains imported ${blocker.kind} formula metadata`,
		{
			refs: [`${blocker.sheetName}!${blocker.ref}`],
			suggestedFix:
				'Materialize or rewrite the imported formula binding before applying row or column structural edits.',
		},
	)
}

function deletedTableColumnStructuralEditError(
	blocker: DeletedTableColumnReference,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot delete column because ${blocker.sourceRef} ${blocker.sourceKind} references table column ${blocker.tableName}[${blocker.columnName}]`,
		{
			refs: [blocker.sourceRef],
			suggestedFix:
				'Rewrite or remove structured references to the table field before deleting the column.',
		},
	)
}

interface TableBoundaryRowDelete {
	readonly tableName: string
	readonly boundary: 'header' | 'totals'
	readonly ref: string
}

function findPartialTableBoundaryRowDelete(
	sheet: Sheet,
	at: number,
	count: number,
): TableBoundaryRowDelete | null {
	const deleteEnd = at + count
	for (const table of sheet.tables) {
		if (table.ref.start.row >= at && table.ref.end.row < deleteEnd) continue
		if (table.hasHeaders && table.ref.start.row >= at && table.ref.start.row < deleteEnd) {
			return {
				tableName: table.name,
				boundary: 'header',
				ref: rangeToA1(table.ref),
			}
		}
		if (table.hasTotals && table.ref.end.row >= at && table.ref.end.row < deleteEnd) {
			return {
				tableName: table.name,
				boundary: 'totals',
				ref: rangeToA1(table.ref),
			}
		}
	}
	return null
}

function tableBoundaryRowStructuralEditError(
	blocker: TableBoundaryRowDelete,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot delete the ${blocker.boundary} row of table ${blocker.tableName} without deleting the whole table range`,
		{
			refs: [blocker.ref],
			suggestedFix:
				'Resize or delete the table explicitly before removing its header or totals row, or delete the full table range in one structural edit.',
		},
	)
}

function copyTransferMetadata(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	mode: PasteMode,
	move: boolean,
): void {
	if (mode === 'all' || mode === 'comments') {
		copyCellMap(sourceSheet.comments, targetSheet.comments, source, rowDelta, colDelta, move)
	}
	if (mode === 'all' || mode === 'hyperlinks') {
		copyCellMap(sourceSheet.hyperlinks, targetSheet.hyperlinks, source, rowDelta, colDelta, move)
	}
	if (mode === 'all' || mode === 'validations') {
		copyDataValidations(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
		copyX14DataValidations(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
	}
	if (mode === 'all' || mode === 'formats' || mode === 'styles') {
		copyConditionalFormats(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
		copyX14ConditionalFormats(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
	}
}

interface MergeTransferPlan {
	readonly targetRange: RangeRef
	readonly sourceMerges: readonly RangeRef[]
	readonly targetMerges: readonly RangeRef[]
	readonly targetReplacedMerges: readonly RangeRef[]
	readonly removeSourceMerges: boolean
}

function planMergeTransfer(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): Result<MergeTransferPlan> {
	const targetRange = shiftRange(source, rowDelta, colDelta)
	const sourceMerges = sourceSheet.merges.filter((merge) => rangesOverlap(merge, source))
	for (const merge of sourceMerges) {
		if (!rangeContainsRange(source, merge)) {
			return err(partialMergeTransferError(move ? 'move' : 'copy', source, merge))
		}
	}
	if (
		sourceMerges.length > 0 &&
		sourceSheet === targetSheet &&
		(rowDelta !== 0 || colDelta !== 0) &&
		rangesOverlap(source, targetRange)
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`Cannot ${move ? 'move' : 'copy'} merged cells onto an overlapping target range`,
				{
					refs: [rangeToA1(source), rangeToA1(targetRange)],
					suggestedFix:
						'Choose a non-overlapping target when copying or moving ranges that contain merged cells.',
				},
			),
		)
	}

	const targetReplacedMerges: RangeRef[] = []
	for (const merge of targetSheet.merges) {
		if (!rangesOverlap(merge, targetRange)) continue
		if (!rangeContainsRange(targetRange, merge)) {
			return err(partialTargetMergeError(targetRange, merge))
		}
		targetReplacedMerges.push(merge)
	}

	return ok({
		targetRange,
		sourceMerges,
		targetMerges: sourceMerges.map((merge) => shiftRange(merge, rowDelta, colDelta)),
		targetReplacedMerges,
		removeSourceMerges: move,
	})
}

function applyMergeTransfer(sourceSheet: Sheet, targetSheet: Sheet, plan: MergeTransferPlan): void {
	if (plan.sourceMerges.length === 0 && plan.targetReplacedMerges.length === 0) return
	const targetRemove = new Set(plan.targetReplacedMerges.map(rangeKey))
	if (plan.removeSourceMerges) {
		const sourceRemove = new Set(plan.sourceMerges.map(rangeKey))
		if (sourceSheet === targetSheet) {
			for (const key of sourceRemove) targetRemove.add(key)
		} else {
			sourceSheet.merges = sourceSheet.merges.filter((merge) => !sourceRemove.has(rangeKey(merge)))
		}
	}
	targetSheet.merges = targetSheet.merges.filter((merge) => !targetRemove.has(rangeKey(merge)))
	targetSheet.merges.push(...plan.targetMerges.map(cloneRange))
}

function partialMergeTransferError(
	operation: 'copy' | 'move',
	source: RangeRef,
	merge: RangeRef,
): ReturnType<typeof ascendError> {
	return ascendError('VALIDATION_ERROR', `Cannot ${operation} part of a merged range`, {
		refs: [rangeToA1(source), rangeToA1(merge)],
		suggestedFix:
			'Select the full merged range, unmerge it first, or use a paste mode that does not transfer merged-cell layout.',
	})
}

function partialTargetMergeError(
	target: RangeRef,
	merge: RangeRef,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		'Target range partially overlaps an existing merged range',
		{
			refs: [rangeToA1(target), rangeToA1(merge)],
			suggestedFix:
				'Choose a target that fully covers the existing merged range, unmerge it first, or paste without formats.',
		},
	)
}

function cloneRange(range: RangeRef): RangeRef {
	return {
		start: { ...range.start },
		end: { ...range.end },
	}
}

function rangeKey(range: RangeRef): string {
	return `${range.start.row}:${range.start.col}:${range.end.row}:${range.end.col}`
}

function copyCellMap<T extends object>(
	sourceMap: Map<string, T>,
	targetMap: Map<string, T>,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const entries = [...sourceMap.entries()].filter(([ref]) =>
		rangeContainsCell(source, parseA1(ref)),
	)
	for (const [ref, value] of entries) {
		const pos = parseA1(ref)
		targetMap.set(toA1({ row: pos.row + rowDelta, col: pos.col + colDelta }), { ...value })
	}
	if (move) {
		for (const [ref] of entries) sourceMap.delete(ref)
	}
}

function copyDataValidations(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const copied: SheetDataValidation[] = []
	const retained: SheetDataValidation[] = []
	const targetRange = shiftRange(source, rowDelta, colDelta)

	for (const validation of sourceSheet.dataValidations) {
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
				? {
						formula1: translateTransferMetadataFormula(
							validation.formula1,
							sourceSheet.name,
							targetSheet.name,
							targetRange,
							rowDelta,
							colDelta,
							move,
						),
					}
				: {}),
			...(validation.formula2
				? {
						formula2: translateTransferMetadataFormula(
							validation.formula2,
							sourceSheet.name,
							targetSheet.name,
							targetRange,
							rowDelta,
							colDelta,
							move,
						),
					}
				: {}),
		})
	}

	if (sourceSheet === targetSheet) {
		sourceSheet.dataValidations = [...retained, ...copied]
		return
	}
	if (move) sourceSheet.dataValidations = retained
	targetSheet.dataValidations = [...targetSheet.dataValidations, ...copied]
}

function copyConditionalFormats(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const copied: SheetConditionalFormat[] = []
	const retained: SheetConditionalFormat[] = []
	const targetRange = shiftRange(source, rowDelta, colDelta)

	for (const format of sourceSheet.conditionalFormats) {
		const ranges = parseSqref(format.sqref)
		if (ranges.length === 0) {
			retained.push(format)
			continue
		}

		const copiedRanges = ranges.filter((range) => rangeContainsRange(source, range))
		if (copiedRanges.length === 0) {
			retained.push(format)
			continue
		}

		if (!move) retained.push(format)
		else {
			const keptRanges = ranges.filter((range) => !rangeContainsRange(source, range))
			if (keptRanges.length > 0) {
				retained.push({ ...format, sqref: rangesToSqref(keptRanges) })
			}
		}

		copied.push({
			...format,
			sqref: rangesToSqref(copiedRanges.map((range) => shiftRange(range, rowDelta, colDelta))),
			rules: format.rules.map((rule) => ({
				...rule,
				formulas: rule.formulas.map((formula) =>
					translateTransferMetadataFormula(
						formula,
						sourceSheet.name,
						targetSheet.name,
						targetRange,
						rowDelta,
						colDelta,
						move,
					),
				),
				...(rule.colorScale
					? {
							colorScale: translateConditionalFormatColorScale(
								rule.colorScale,
								sourceSheet.name,
								targetSheet.name,
								targetRange,
								rowDelta,
								colDelta,
								move,
							),
						}
					: {}),
				...(rule.dataBar
					? {
							dataBar: translateConditionalFormatDataBar(
								rule.dataBar,
								sourceSheet.name,
								targetSheet.name,
								targetRange,
								rowDelta,
								colDelta,
								move,
							),
						}
					: {}),
				...(rule.iconSet
					? {
							iconSet: translateConditionalFormatIconSet(
								rule.iconSet,
								sourceSheet.name,
								targetSheet.name,
								targetRange,
								rowDelta,
								colDelta,
								move,
							),
						}
					: {}),
			})),
		})
	}

	if (sourceSheet === targetSheet) {
		sourceSheet.conditionalFormats = [...retained, ...copied]
		return
	}
	if (move) sourceSheet.conditionalFormats = retained
	targetSheet.conditionalFormats = [...targetSheet.conditionalFormats, ...copied]
}

function copyX14DataValidations(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const copied: Sheet['x14DataValidations'] = []
	const retained: Sheet['x14DataValidations'] = []
	const targetRange = shiftRange(source, rowDelta, colDelta)
	let nextIndex = nextX14Index(targetSheet.x14DataValidations)

	for (const validation of sourceSheet.x14DataValidations) {
		if (validation.deleted) {
			retained.push(validation)
			continue
		}
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

		const keptRanges = ranges.filter((range) => !rangeContainsRange(source, range))
		if (!move) retained.push(validation)
		else if (keptRanges.length > 0) {
			retained.push({ ...validation, sqref: rangesToSqref(keptRanges) })
		} else if (sourceSheet !== targetSheet) {
			retained.push({ ...validation, sqref: '', deleted: true })
		}

		copied.push({
			...validation,
			index:
				move && sourceSheet === targetSheet && keptRanges.length === 0
					? validation.index
					: nextIndex++,
			sqref: rangesToSqref(copiedRanges.map((range) => shiftRange(range, rowDelta, colDelta))),
			...(validation.formula1
				? {
						formula1: translateTransferMetadataFormula(
							validation.formula1,
							sourceSheet.name,
							targetSheet.name,
							targetRange,
							rowDelta,
							colDelta,
							move,
						),
					}
				: {}),
			...(validation.formula2
				? {
						formula2: translateTransferMetadataFormula(
							validation.formula2,
							sourceSheet.name,
							targetSheet.name,
							targetRange,
							rowDelta,
							colDelta,
							move,
						),
					}
				: {}),
		})
	}

	if (sourceSheet === targetSheet) {
		sourceSheet.x14DataValidations = [...retained, ...copied]
		return
	}
	if (move) sourceSheet.x14DataValidations = retained
	targetSheet.x14DataValidations = [...targetSheet.x14DataValidations, ...copied]
}

function copyX14ConditionalFormats(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const copied: Sheet['x14ConditionalFormats'] = []
	const retained: Sheet['x14ConditionalFormats'] = []
	const targetRange = shiftRange(source, rowDelta, colDelta)
	let nextIndex = nextX14Index(targetSheet.x14ConditionalFormats)

	for (const format of sourceSheet.x14ConditionalFormats) {
		if (format.deleted) {
			retained.push(format)
			continue
		}
		const ranges = parseSqref(format.sqref)
		if (ranges.length === 0) {
			retained.push(format)
			continue
		}

		const copiedRanges = ranges.filter((range) => rangeContainsRange(source, range))
		if (copiedRanges.length === 0) {
			retained.push(format)
			continue
		}

		const keptRanges = ranges.filter((range) => !rangeContainsRange(source, range))
		if (!move) retained.push(format)
		else if (keptRanges.length > 0) {
			retained.push({ ...format, sqref: rangesToSqref(keptRanges) })
		} else if (sourceSheet !== targetSheet) {
			retained.push({ ...format, sqref: '', deleted: true })
		}

		copied.push({
			...format,
			index:
				move && sourceSheet === targetSheet && keptRanges.length === 0 ? format.index : nextIndex++,
			sqref: rangesToSqref(copiedRanges.map((range) => shiftRange(range, rowDelta, colDelta))),
			formulas: format.formulas.map((formula) =>
				translateTransferMetadataFormula(
					formula,
					sourceSheet.name,
					targetSheet.name,
					targetRange,
					rowDelta,
					colDelta,
					move,
				),
			),
			...(format.dataBar
				? {
						dataBar: translateX14ConditionalFormatDataBar(
							format.dataBar,
							sourceSheet.name,
							targetSheet.name,
							targetRange,
							rowDelta,
							colDelta,
							move,
						),
					}
				: {}),
			...(format.iconSet
				? {
						iconSet: translateX14ConditionalFormatIconSet(
							format.iconSet,
							sourceSheet.name,
							targetSheet.name,
							targetRange,
							rowDelta,
							colDelta,
							move,
						),
					}
				: {}),
		})
	}

	if (sourceSheet === targetSheet) {
		sourceSheet.x14ConditionalFormats = [...retained, ...copied]
		return
	}
	if (move) sourceSheet.x14ConditionalFormats = retained
	targetSheet.x14ConditionalFormats = [...targetSheet.x14ConditionalFormats, ...copied]
}

function nextX14Index(entries: readonly { readonly index: number }[]): number {
	return entries.reduce((max, entry) => Math.max(max, entry.index + 1), 0)
}

function translateConditionalFormatColorScale(
	colorScale: NonNullable<SheetConditionalFormat['rules'][number]['colorScale']>,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): NonNullable<SheetConditionalFormat['rules'][number]['colorScale']> {
	return {
		cfvo: colorScale.cfvo.map((entry) =>
			translateConditionalFormatValueObject(
				entry,
				sourceSheetName,
				targetSheetName,
				targetRange,
				rowDelta,
				colDelta,
				move,
			),
		),
		colors: colorScale.colors.map((color) => ({ ...color })),
	}
}

function translateConditionalFormatDataBar(
	dataBar: NonNullable<SheetConditionalFormat['rules'][number]['dataBar']>,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): NonNullable<SheetConditionalFormat['rules'][number]['dataBar']> {
	return {
		...dataBar,
		cfvo: dataBar.cfvo.map((entry) =>
			translateConditionalFormatValueObject(
				entry,
				sourceSheetName,
				targetSheetName,
				targetRange,
				rowDelta,
				colDelta,
				move,
			),
		),
		...(dataBar.color ? { color: { ...dataBar.color } } : {}),
	}
}

function translateConditionalFormatIconSet(
	iconSet: NonNullable<SheetConditionalFormat['rules'][number]['iconSet']>,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): NonNullable<SheetConditionalFormat['rules'][number]['iconSet']> {
	return {
		...iconSet,
		cfvo: iconSet.cfvo.map((entry) =>
			translateConditionalFormatValueObject(
				entry,
				sourceSheetName,
				targetSheetName,
				targetRange,
				rowDelta,
				colDelta,
				move,
			),
		),
	}
}

function translateX14ConditionalFormatDataBar(
	dataBar: NonNullable<Sheet['x14ConditionalFormats'][number]['dataBar']>,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): NonNullable<Sheet['x14ConditionalFormats'][number]['dataBar']> {
	return {
		...dataBar,
		cfvo: dataBar.cfvo.map((entry) =>
			translateConditionalFormatValueObject(
				entry,
				sourceSheetName,
				targetSheetName,
				targetRange,
				rowDelta,
				colDelta,
				move,
			),
		),
		...(dataBar.borderColor ? { borderColor: { ...dataBar.borderColor } } : {}),
		...(dataBar.negativeFillColor ? { negativeFillColor: { ...dataBar.negativeFillColor } } : {}),
		...(dataBar.negativeBorderColor
			? { negativeBorderColor: { ...dataBar.negativeBorderColor } }
			: {}),
		...(dataBar.axisColor ? { axisColor: { ...dataBar.axisColor } } : {}),
	}
}

function translateX14ConditionalFormatIconSet(
	iconSet: NonNullable<Sheet['x14ConditionalFormats'][number]['iconSet']>,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): NonNullable<Sheet['x14ConditionalFormats'][number]['iconSet']> {
	return {
		...iconSet,
		cfvo: iconSet.cfvo.map((entry) =>
			translateConditionalFormatValueObject(
				entry,
				sourceSheetName,
				targetSheetName,
				targetRange,
				rowDelta,
				colDelta,
				move,
			),
		),
		...(iconSet.icons ? { icons: iconSet.icons.map((icon) => ({ ...icon })) } : {}),
	}
}

function translateConditionalFormatValueObject(
	entry: SheetConditionalFormatValueObject,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): SheetConditionalFormatValueObject {
	if (entry.type !== 'formula' || entry.value === undefined) return { ...entry }
	return {
		...entry,
		value: translateTransferMetadataFormula(
			entry.value,
			sourceSheetName,
			targetSheetName,
			targetRange,
			rowDelta,
			colDelta,
			move,
		),
	}
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

function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
	return !(
		a.end.row < b.start.row ||
		a.start.row > b.end.row ||
		a.end.col < b.start.col ||
		a.start.col > b.end.col
	)
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

function translateTransferMetadataFormula(
	formula: string,
	sourceSheetName: string,
	targetSheetName: string,
	targetRange: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): string {
	const translated = translateMetadataFormula(formula, rowDelta, colDelta)
	if (!move || sourceSheetName === targetSheetName) return translated
	return (
		retargetExplicitFormulaSheetRefsInRange(
			translated,
			sourceSheetName,
			targetSheetName,
			targetRange,
		) ?? translated
	)
}

export function handleMergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'mergeCells' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult

	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult

	const sheet = sheetResult.value
	const overlapping = sheet.merges.find((merge) => rangesOverlap(merge, rangeResult.value))
	if (overlapping) {
		return err(
			ascendError('VALIDATION_ERROR', 'Merged ranges cannot overlap', {
				refs: [op.range, rangeToA1(overlapping)],
				suggestedFix: 'Unmerge the existing range first or choose a non-overlapping range.',
			}),
		)
	}

	sheet.merges.push(rangeResult.value)
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
