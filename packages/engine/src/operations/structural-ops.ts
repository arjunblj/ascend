import type {
	AutoFilter,
	RangeRef,
	Sheet,
	SheetAdvancedFilterInfo,
	SheetAnchorMarker,
	SheetComment,
	SheetConditionalFormat,
	SheetConditionalFormatValueObject,
	SheetDataValidation,
	SheetImageAnchor,
	SheetProtectedRange,
	SortState,
	Workbook,
} from '@ascend/core'
import { parseA1, parseRange, toA1 } from '@ascend/core'
import { cachedParseFormula } from '@ascend/formulas'
import type { Operation, PasteMode, Result } from '@ascend/schema'
import { ascendError, EMPTY, err, ok } from '@ascend/schema'
import { resolveCellFormulaText } from '../analysis.ts'
import {
	findPartialFormulaMoveReference,
	formulaAstHasLocalStructuralReference,
	type PartialFormulaMoveReference,
	retargetExplicitFormulaSheetRefsInRange,
	rewriteDefinedNameFormulasForMove,
	rewriteDefinedNameFormulasForShift,
	rewriteWorkbookChartSourceRefsForMove,
	rewriteWorkbookChartSourceRefsForShift,
	rewriteWorkbookFormulasForMove,
	rewriteWorkbookFormulasForShift,
	rewriteWorkbookHyperlinkLocationsForMove,
	rewriteWorkbookMetadataFormulasForMove,
	rewriteWorkbookMetadataFormulasForShift,
} from '../structural/formula-rewrite.ts'
import { shiftSheetCellMetadata } from '../structural/sheet-topology.ts'
import {
	collectDeletedTableColumns,
	type DeletedTableColumnReference,
	findDeletedTableColumnReference,
} from '../structural/table-field-guards.ts'
import {
	findQueryTableColumnShiftBlocker,
	findShiftedTableRangeOverlap,
	type QueryTableColumnShiftBlocker,
	type TableRangeOverlap,
} from '../table-topology.ts'
import type { PatchResult } from './helpers.ts'
import {
	cellPreservingFormulaInfo,
	cellWithExisting,
	clearFormulaMetadata,
	collectFormulaBindingGroupRefsForRefs,
	collectRangeCells,
	createLegacyArrayFormulaIndex,
	DEFAULT_SID,
	getSheet,
	legacyArrayFormulaEditError,
	materializeFormulaBindingGroupsForRangeEdit,
	patch,
	safeParseA1,
	safeParseRange,
	shiftMerges,
	translateFormula,
} from './helpers.ts'

const EXCEL_MAX_ROWS = 1_048_576
const EXCEL_MAX_COLS = 16_384

function applyAxisShift(
	workbook: Workbook,
	sheetName: string,
	axis: 'row' | 'col',
	at: number,
	count: number,
	delta: number,
): Result<PatchResult> {
	const spanError = validateAxisSpan(axis, at, count)
	if (spanError) return err(spanError)
	const result = getSheet(workbook, sheetName)
	if (!result.ok) return result
	const sheet = result.value
	const protectionBlocker = validateStructuralProtection(sheet, axis, delta)
	if (protectionBlocker) return err(protectionBlocker)
	const insertOverflowBlocker =
		delta > 0 ? findInsertShiftedCellOutOfBounds(sheet, axis, at, count) : null
	if (insertOverflowBlocker) {
		return err(insertShiftedCellOutOfBoundsError(sheet, axis, at, count, insertOverflowBlocker))
	}
	const ambiguousDefinedName = findAmbiguousWorkbookDefinedNameStructuralReference(workbook)
	if (ambiguousDefinedName) {
		return err(ambiguousWorkbookDefinedNameStructuralEditError(ambiguousDefinedName))
	}
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
	const queryTableColumnBlocker =
		axis === 'col' ? findQueryTableColumnShiftBlocker(sheet, at, delta) : null
	if (queryTableColumnBlocker) {
		return err(queryTableColumnStructuralEditError(queryTableColumnBlocker))
	}
	const tableOverlapBlocker = findShiftedTableRangeOverlap(sheet, axis, at, delta)
	if (tableOverlapBlocker) return err(tableStructuralShiftOverlapError(tableOverlapBlocker))
	const deletedQueryTablePartPaths = collectStructurallyDeletedQueryTablePartPaths(
		sheet,
		axis,
		at,
		count,
		delta,
	)

	sheet.ensureWritable()
	if (axis === 'row') {
		delta > 0 ? sheet.cells.insertRows(at, count) : sheet.cells.deleteRows(at, count)
	} else {
		delta > 0 ? sheet.cells.insertCols(at, count) : sheet.cells.deleteCols(at, count)
	}

	shiftMerges(sheet.merges, axis, at, delta)
	shiftSheetCellMetadata(sheet, axis, at, delta)
	clearFormulaMetadata(workbook)
	const affected = new Set<string>()
	const sheetsModified = new Set<string>([sheetName])
	for (const rewritten of rewriteWorkbookFormulasForShift(workbook, sheetName, axis, at, delta)) {
		affected.add(affectedNamedRef(rewritten.sheetName, rewritten.ref, sheetName, false))
		sheetsModified.add(rewritten.sheetName)
	}
	rewriteDefinedNameFormulasForShift(workbook, sheetName, axis, at, delta)
	rewriteWorkbookMetadataFormulasForShift(workbook, sheetName, axis, at, delta)
	for (const chartSheetName of rewriteWorkbookChartSourceRefsForShift(
		workbook,
		sheetName,
		axis,
		at,
		delta,
	)) {
		sheetsModified.add(chartSheetName)
	}
	removeDeletedQueryTableConnectionParts(workbook, deletedQueryTablePartPaths)

	return ok(patch([...affected], [...sheetsModified], true))
}

function validateStructuralProtection(sheet: Sheet, axis: 'row' | 'col', delta: number) {
	const protection = sheet.protection
	if (!protection || protection.sheet === false) return null
	const operation =
		axis === 'row'
			? delta > 0
				? 'insertRows'
				: 'deleteRows'
			: delta > 0
				? 'insertColumns'
				: 'deleteColumns'
	const allowed = protection[operation] === true
	if (allowed) return null
	return ascendError(
		'PROTECTION_ERROR',
		`Cannot ${operation} on protected sheet "${sheet.name}" because sheet protection does not allow it`,
		{
			refs: [`sheet:${sheet.name}:protection:${operation}`],
			suggestedFix: `Unprotect the sheet or set sheet protection ${operation}=true before applying this structural edit.`,
			details: {
				kind: 'sheet-protection-structural-edit-blocked',
				sheetName: sheet.name,
				operation,
				allowed,
			},
		},
	)
}

function findAmbiguousWorkbookDefinedNameStructuralReference(workbook: Workbook) {
	for (const entry of workbook.definedNames.list()) {
		if (entry.scope.kind !== 'workbook') continue
		const parsed = cachedParseFormula(entry.formula)
		if (!parsed.ok) continue
		if (formulaAstHasLocalStructuralReference(parsed.value)) return entry
	}
	return null
}

function ambiguousWorkbookDefinedNameStructuralEditError(
	entry: NonNullable<ReturnType<typeof findAmbiguousWorkbookDefinedNameStructuralReference>>,
) {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot apply structural row or column edit while workbook-scoped defined name "${entry.name}" contains unqualified references`,
		{
			refs: [`definedName:${entry.name}`],
			suggestedFix:
				'Qualify the defined name formula with an explicit sheet name or convert it to a sheet-scoped name before applying structural edits.',
			details: {
				kind: 'ambiguous-workbook-defined-name-structural-edit',
				name: entry.name,
				formula: entry.formula,
				scope: entry.scope,
			},
		},
	)
}

function validateAxisSpan(axis: 'row' | 'col', at: number, count: number) {
	const label = axis === 'row' ? 'row' : 'column'
	const limit = axis === 'row' ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS
	const maxRef = axis === 'row' ? '1048576' : 'XFD'
	if (!Number.isInteger(at) || at < 0) {
		return ascendError('VALIDATION_ERROR', `${label} index must be a non-negative integer`)
	}
	if (!Number.isInteger(count) || count <= 0) {
		return ascendError('VALIDATION_ERROR', `${label} count must be a positive integer`)
	}
	if (at >= limit || count > limit - at) {
		return ascendError(
			'INVALID_RANGE',
			`Cannot structurally edit ${label}s starting at ${at + 1} with count ${count} because the span is outside Excel worksheet bounds`,
			{
				suggestedFix: `Choose a ${label} span within Excel's ${limit} ${label} limit ending no later than ${maxRef}.`,
				details: {
					kind: 'structural-axis-span-out-of-bounds',
					axis,
					at,
					count,
					limit,
				},
			},
		)
	}
	return null
}

interface InsertOverflowCell {
	readonly ref: string
	readonly shiftedRef: string
}

function findInsertShiftedCellOutOfBounds(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	count: number,
): InsertOverflowCell | null {
	const limit = axis === 'row' ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS
	const overflowStart = limit - count
	for (const [row, col] of sheet.cells.iterate()) {
		const coordinate = axis === 'row' ? row : col
		if (coordinate < at || coordinate < overflowStart) continue
		return {
			ref: toA1({ row, col }),
			shiftedRef:
				axis === 'row' ? toA1({ row: row + count, col }) : toA1({ row, col: col + count }),
		}
	}
	return null
}

function insertShiftedCellOutOfBoundsError(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	count: number,
	cell: InsertOverflowCell,
) {
	const label = axis === 'row' ? 'row' : 'column'
	const maxRef = axis === 'row' ? '1048576' : 'XFD'
	const qualifiedRef = `${formatSheetName(sheet.name)}!${cell.ref}`
	return ascendError(
		'INVALID_RANGE',
		`Cannot insert ${label}s because ${qualifiedRef} would shift outside Excel worksheet bounds`,
		{
			refs: [qualifiedRef],
			suggestedFix: `Clear or move populated cells that would shift past ${maxRef}, then retry the insert.`,
			details: {
				kind: 'structural-insert-shifts-cell-out-of-bounds',
				axis,
				at,
				count,
				ref: qualifiedRef,
				shiftedRef: cell.shiftedRef,
			},
		},
	)
}

function validateTransferRangeWithinGrid(range: RangeRef, role: 'source' | 'target') {
	if (
		range.start.row >= 0 &&
		range.end.row < EXCEL_MAX_ROWS &&
		range.start.col >= 0 &&
		range.end.col < EXCEL_MAX_COLS
	) {
		return null
	}
	const ref = rangeToA1(range)
	return ascendError(
		'INVALID_RANGE',
		`Cannot transfer ${role} range ${ref} because it is outside Excel worksheet bounds`,
		{
			refs: [ref],
			suggestedFix: `Choose a ${role} range within A1:XFD1048576.`,
			details: {
				kind: 'range-transfer-out-of-bounds',
				role,
				ref,
				maxRef: 'XFD1048576',
			},
		},
	)
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
	const targetStartResult = safeParseA1(op.target)
	if (!targetStartResult.ok) return targetStartResult
	const targetStart = targetStartResult.value
	const source = sourceResult.value
	const sourceBoundsError = validateTransferRangeWithinGrid(source, 'source')
	if (sourceBoundsError) return err(sourceBoundsError)
	const rowDelta = targetStart.row - source.start.row
	const colDelta = targetStart.col - source.start.col
	const targetRange = shiftRange(source, rowDelta, colDelta)
	const targetBoundsError = validateTransferRangeWithinGrid(targetRange, 'target')
	if (targetBoundsError) return err(targetBoundsError)
	const mode = op.mode ?? 'all'
	const affected = new Set<string>()
	const sheetsModified = new Set<string>([sourceSheet.name, targetSheet.name])

	if (overwritesTargetFormulas(mode) || (op.op === 'moveRange' && clearsSourceCellContent(mode))) {
		const targetLegacyArrayIndex = createLegacyArrayFormulaIndex(targetSheet)
		if (overwritesTargetFormulas(mode)) {
			const blockedTarget = targetLegacyArrayIndex.findIntersection(targetRange)
			if (blockedTarget) {
				return err(legacyArrayFormulaEditError(blockedTarget.targetRef, blockedTarget.ref))
			}
		}
		if (op.op === 'moveRange' && clearsSourceCellContent(mode)) {
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
	if (
		op.op === 'moveRange' &&
		sourceSheet === targetSheet &&
		rangesOverlap(source, mergePlan.value.targetRange)
	) {
		return err(overlappingMoveRangeError(source, mergePlan.value.targetRange))
	}
	const visualPlan = planVisualTransfer(
		sourceSheet,
		targetSheet,
		source,
		rowDelta,
		colDelta,
		mode,
		op.op,
	)
	if (!visualPlan.ok) return visualPlan
	const sheetFilterPlan = planSheetFilterTransfer(
		sourceSheet,
		targetSheet,
		source,
		rowDelta,
		colDelta,
		mode,
		op.op,
	)
	if (!sheetFilterPlan.ok) return sheetFilterPlan
	const protectedRangePlan = planProtectedRangeTransfer(
		sourceSheet,
		targetSheet,
		source,
		rowDelta,
		colDelta,
		mode,
		op.op,
	)
	if (!protectedRangePlan.ok) return protectedRangePlan
	const advancedFilterPlan = planAdvancedFilterTransfer(
		sourceSheet,
		targetSheet,
		source,
		rowDelta,
		colDelta,
		mode,
		op.op,
	)
	if (!advancedFilterPlan.ok) return advancedFilterPlan

	const partialMetadataBlocker = findPartialMetadataRangeTransfer(sourceSheet, source, mode)
	if (partialMetadataBlocker) {
		return err(
			partialMetadataRangeTransferError(
				op.op === 'moveRange' ? 'move' : 'copy',
				source,
				partialMetadataBlocker,
			),
		)
	}

	if (op.op === 'moveRange' && clearsSourceCellContent(mode)) {
		const skipCells = [{ sheetName: sourceSheet.name, range: source }]
		if (overwritesTargetFormulas(mode)) {
			skipCells.push({ sheetName: targetSheet.name, range: mergePlan.value.targetRange })
		}
		const partialFormulaBlocker = findPartialFormulaMoveReference(
			workbook,
			sourceSheet.name,
			source,
			{ skipCells },
		)
		if (partialFormulaBlocker) {
			return err(partialMoveFormulaReferenceError(partialFormulaBlocker, source))
		}
	}

	const snapshot = collectRangeCells(sourceSheet, source)

	if (pasteCells(mode)) {
		if (overwritesTargetFormulas(mode)) {
			for (const ref of materializeFormulaBindingGroupsForRangeEdit(
				workbook,
				targetSheet,
				mergePlan.value.targetRange,
			)) {
				affected.add(affectedRef(targetSheet, ref, crossSheet))
			}
		}
		if (op.op === 'moveRange' && clearsSourceCellContent(mode)) {
			for (const ref of materializeFormulaBindingGroupsForRangeEdit(
				workbook,
				sourceSheet,
				source,
			)) {
				affected.add(affectedRef(sourceSheet, ref, crossSheet))
			}
		}
		for (const entry of snapshot) {
			const targetRow = entry.row + rowDelta
			const targetCol = entry.col + colDelta
			const existingTarget = targetSheet.cells.get(targetRow, targetCol)
			const ref = toA1({ row: targetRow, col: targetCol })

			if (!entry.cell && mode === 'all') {
				targetSheet.cells.delete(targetRow, targetCol)
				affected.add(affectedRef(targetSheet, ref, crossSheet))
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
					cellPreservingFormulaInfo(
						targetValue,
						targetFormula,
						entry.cell?.styleId ?? DEFAULT_SID,
						existingTarget?.formulaInfo,
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
			affected.add(affectedRef(targetSheet, ref, crossSheet))
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
	applyVisualTransfer(visualPlan.value)
	applySheetFilterTransfer(sheetFilterPlan.value)
	applyProtectedRangeTransfer(protectedRangePlan.value)
	applyAdvancedFilterTransfer(advancedFilterPlan.value)

	if (op.op === 'moveRange') {
		for (const entry of snapshot) {
			if (pasteCells(mode)) {
				if (clearsSourceCellContent(mode)) {
					sourceSheet.cells.delete(entry.row, entry.col)
				} else if (entry.cell && (mode === 'formats' || mode === 'styles')) {
					sourceSheet.cells.set(
						entry.row,
						entry.col,
						cellPreservingFormulaInfo(
							entry.cell.value,
							entry.cell.formula,
							DEFAULT_SID,
							entry.cell.formulaInfo,
						),
					)
				}
				affected.add(affectedRef(sourceSheet, toA1({ row: entry.row, col: entry.col }), crossSheet))
			}
		}
		if (clearsSourceCellContent(mode)) {
			for (const rewritten of rewriteWorkbookFormulasForMove(
				workbook,
				sourceSheet.name,
				targetSheet.name,
				source,
				mergePlan.value.targetRange,
			)) {
				affected.add(
					affectedNamedRef(rewritten.sheetName, rewritten.ref, sourceSheet.name, crossSheet),
				)
				const rewrittenSheet = workbook.getSheet(rewritten.sheetName)
				const rewrittenRef = parseA1(rewritten.ref)
				if (rewrittenSheet && rewrittenRef) {
					for (const groupRef of collectFormulaBindingGroupRefsForRefs(workbook, rewrittenSheet, [
						rewrittenRef,
					])) {
						affected.add(
							affectedNamedRef(rewritten.sheetName, groupRef, sourceSheet.name, crossSheet),
						)
					}
				}
			}
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
			for (const chartSheetName of rewriteWorkbookChartSourceRefsForMove(
				workbook,
				sourceSheet.name,
				targetSheet.name,
				source,
				mergePlan.value.targetRange,
			)) {
				sheetsModified.add(chartSheetName)
			}
			for (const rewritten of rewriteWorkbookHyperlinkLocationsForMove(
				workbook,
				sourceSheet.name,
				targetSheet.name,
				source,
				mergePlan.value.targetRange,
			)) {
				affected.add(
					affectedNamedRef(rewritten.sheetName, rewritten.ref, sourceSheet.name, crossSheet),
				)
			}
		}
	}

	return ok(patch([...affected], [...sheetsModified], pasteRequiresRecalc(mode)))
}

function affectedRef(sheet: Sheet, ref: string, qualify: boolean): string {
	return qualify ? `${sheet.name}!${ref}` : ref
}

function affectedNamedRef(
	sheetName: string,
	ref: string,
	primarySheetName: string,
	qualifyPrimary: boolean,
): string {
	return qualifyPrimary || sheetName !== primarySheetName ? `${sheetName}!${ref}` : ref
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

function pasteValidations(mode: PasteMode): boolean {
	return mode === 'all' || mode === 'validations'
}

function pasteConditionalFormats(mode: PasteMode): boolean {
	return mode === 'all' || mode === 'formats' || mode === 'styles'
}

function pasteRequiresRecalc(mode: PasteMode): boolean {
	return mode === 'all' || mode === 'values' || mode === 'formulas'
}

function clearsSourceCellContent(mode: PasteMode): boolean {
	return mode === 'all' || mode === 'values' || mode === 'formulas'
}

function overwritesTargetFormulas(mode: PasteMode): boolean {
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

function tableStructuralShiftOverlapError(
	blocker: TableRangeOverlap,
): ReturnType<typeof ascendError> {
	const leftRef = rangeToA1(blocker.leftRef)
	const rightRef = rangeToA1(blocker.rightRef)
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot structurally edit rows or columns because shifted table ${blocker.left.name} at ${leftRef} would overlap table ${blocker.right.name} at ${rightRef}`,
		{
			refs: [leftRef, rightRef],
			suggestedFix:
				'Resolve the overlapping table ranges before applying row or column structural edits.',
			details: {
				kind: 'overlapping-table-ranges',
				left: {
					tableName: blocker.left.name,
					ref: leftRef,
					...(blocker.left.partPath ? { partPath: blocker.left.partPath } : {}),
				},
				right: {
					tableName: blocker.right.name,
					ref: rightRef,
					...(blocker.right.partPath ? { partPath: blocker.right.partPath } : {}),
				},
			},
		},
	)
}

function queryTableColumnStructuralEditError(
	blocker: QueryTableColumnShiftBlocker,
): ReturnType<typeof ascendError> {
	const currentRef = rangeToA1(blocker.currentRef)
	const shiftedRef = rangeToA1(blocker.shiftedRef)
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot structurally edit columns through queryTable-backed table "${blocker.table.name}" because changing columns would leave queryTable field bindings ambiguous`,
		{
			refs: [currentRef, shiftedRef],
			details: {
				kind: 'query-table-column-structural-edit',
				tableName: blocker.table.name,
				currentRef,
				shiftedRef,
				queryTablePartPath: blocker.table.queryTable?.partPath,
			},
			suggestedFix:
				'Insert or delete columns outside the table, resize only the row span, or remove and rebuild the queryTable sidecar with matching queryTableField bindings before changing table columns.',
		},
	)
}

function collectStructurallyDeletedQueryTablePartPaths(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	count: number,
	delta: number,
): string[] {
	if (delta >= 0) return []
	const deleteEnd = at + count
	const partPaths: string[] = []
	for (const table of sheet.tables) {
		if (!table.queryTable) continue
		const start = axis === 'row' ? table.ref.start.row : table.ref.start.col
		const end = axis === 'row' ? table.ref.end.row : table.ref.end.col
		if (start >= at && end < deleteEnd) partPaths.push(table.queryTable.partPath)
	}
	return partPaths
}

function removeDeletedQueryTableConnectionParts(
	workbook: Workbook,
	partPaths: readonly string[],
): void {
	if (partPaths.length === 0) return
	const deleted = new Set(partPaths)
	for (let index = workbook.connectionParts.length - 1; index >= 0; index--) {
		const connection = workbook.connectionParts[index]
		if (connection?.kind === 'queryTable' && deleted.has(connection.partPath)) {
			workbook.connectionParts.splice(index, 1)
		}
	}
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
		copyCommentMap(sourceSheet.comments, targetSheet.comments, source, rowDelta, colDelta, move)
		copyThreadedComments(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
	}
	if (mode === 'all' || mode === 'hyperlinks') {
		copyCellMap(sourceSheet.hyperlinks, targetSheet.hyperlinks, source, rowDelta, colDelta, move)
	}
	if (pasteValidations(mode)) {
		copyDataValidations(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
		copyX14DataValidations(sourceSheet, targetSheet, source, rowDelta, colDelta, move)
	}
	if (pasteConditionalFormats(mode)) {
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

interface VisualTransferPlan {
	readonly sheet: Sheet
	readonly entries: readonly VisualTransferEntry[]
}

interface SheetFilterTransferPlan {
	readonly sheet: Sheet
	readonly autoFilter?: AutoFilter
	readonly sortState?: SortState
}

interface ProtectedRangeTransferPlan {
	readonly sheet: Sheet
	readonly entries: readonly ProtectedRangeTransferEntry[]
}

interface ProtectedRangeTransferEntry {
	readonly index: number
	readonly range: SheetProtectedRange
}

interface AdvancedFilterTransferPlan {
	readonly sheet: Sheet
	readonly entries: readonly AdvancedFilterTransferEntry[]
}

interface AdvancedFilterTransferEntry {
	readonly index: number
	readonly filter: SheetAdvancedFilterInfo
}

interface PartialMetadataRangeTransfer {
	readonly detailKind:
		| 'partial-data-validation-range-transfer'
		| 'partial-x14-data-validation-range-transfer'
		| 'partial-conditional-format-range-transfer'
		| 'partial-x14-conditional-format-range-transfer'
	readonly metadataKind:
		| 'dataValidation'
		| 'x14DataValidation'
		| 'conditionalFormat'
		| 'x14ConditionalFormat'
	readonly index: number
	readonly ref: string
	readonly label: string
}

interface VisualTransferEntry {
	readonly kind: 'image' | 'drawingObject'
	readonly index: number
	readonly anchor: SheetImageAnchor
}

interface VisualAnchorIntersection extends VisualTransferEntry {
	readonly contained: boolean
	readonly ref: string
	readonly label: string
}

interface SheetFilterIntersection {
	readonly kind: 'autoFilter' | 'sortState'
	readonly contained: boolean
	readonly ref: string
	readonly label: string
}

interface ProtectedRangeIntersection {
	readonly range: SheetProtectedRange
	readonly index: number
	readonly ref: string
	readonly label: string
	readonly overlappingRanges: readonly RangeRef[]
	readonly partialRange?: RangeRef
}

interface AdvancedFilterIntersection {
	readonly filter: SheetAdvancedFilterInfo
	readonly index: number
	readonly contained: boolean
	readonly ref: string
	readonly label: string
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

function planVisualTransfer(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	mode: PasteMode,
	operation: 'copyRange' | 'moveRange',
): Result<VisualTransferPlan> {
	if (mode !== 'all') return ok({ sheet: sourceSheet, entries: [] })
	const intersecting = collectIntersectingVisualAnchors(sourceSheet, source)
	if (intersecting.length === 0) return ok({ sheet: sourceSheet, entries: [] })
	const first = intersecting[0]
	if (!first) return ok({ sheet: sourceSheet, entries: [] })
	if (operation === 'copyRange') {
		return err(unsupportedVisualTransferError('copy', source, first))
	}
	if (sourceSheet !== targetSheet) {
		return err(unsupportedVisualTransferError('move to another sheet', source, first))
	}

	const entries: VisualTransferEntry[] = []
	for (const entry of intersecting) {
		if (!entry.contained) return err(partialVisualTransferError(source, entry))
		entries.push({
			kind: entry.kind,
			index: entry.index,
			anchor: translateSheetImageAnchor(entry.anchor, rowDelta, colDelta),
		})
	}
	return ok({ sheet: sourceSheet, entries })
}

function applyVisualTransfer(plan: VisualTransferPlan): void {
	for (const entry of plan.entries) {
		if (entry.kind === 'image') {
			const image = plan.sheet.imageRefs[entry.index]
			if (image) plan.sheet.imageRefs[entry.index] = { ...image, anchor: entry.anchor }
		} else {
			const object = plan.sheet.drawingObjectRefs[entry.index]
			if (object) plan.sheet.drawingObjectRefs[entry.index] = { ...object, anchor: entry.anchor }
		}
	}
}

function planSheetFilterTransfer(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	mode: PasteMode,
	operation: 'copyRange' | 'moveRange',
): Result<SheetFilterTransferPlan> {
	if (mode !== 'all') return ok({ sheet: sourceSheet })
	const intersecting = collectIntersectingSheetFilters(sourceSheet, source)
	if (intersecting.length === 0) return ok({ sheet: sourceSheet })
	const first = intersecting[0]
	if (!first) return ok({ sheet: sourceSheet })
	if (operation === 'copyRange') {
		return err(unsupportedSheetFilterTransferError('copy', source, first))
	}
	if (sourceSheet !== targetSheet) {
		return err(unsupportedSheetFilterTransferError('move to another sheet', source, first))
	}

	let autoFilter: AutoFilter | undefined
	let sortState: SortState | undefined
	for (const entry of intersecting) {
		if (!entry.contained) return err(partialSheetFilterTransferError(source, entry))
		if (entry.kind === 'autoFilter') {
			if (!sourceSheet.autoFilter) continue
			const translated = translateAutoFilter(sourceSheet.autoFilter, rowDelta, colDelta, source, {
				label: 'sheet autoFilter',
				kind: 'sheet-filter',
			})
			if (!translated.ok) return translated
			autoFilter = translated.value
		} else {
			if (!sourceSheet.sortState) continue
			const translated = translateSortState(sourceSheet.sortState, rowDelta, colDelta, source, {
				label: 'sheet sortState',
				kind: 'sheet-sort-state',
			})
			if (!translated.ok) return translated
			sortState = translated.value
		}
	}
	return ok({
		sheet: sourceSheet,
		...(autoFilter ? { autoFilter } : {}),
		...(sortState ? { sortState } : {}),
	})
}

function applySheetFilterTransfer(plan: SheetFilterTransferPlan): void {
	if (plan.autoFilter) plan.sheet.autoFilter = plan.autoFilter
	if (plan.sortState) plan.sheet.sortState = plan.sortState
}

function collectIntersectingSheetFilters(
	sheet: Sheet,
	source: RangeRef,
): SheetFilterIntersection[] {
	const entries: SheetFilterIntersection[] = []
	const autoFilterRange = sheet.autoFilter ? parseOptionalRange(sheet.autoFilter.ref) : null
	if (autoFilterRange && rangesOverlap(autoFilterRange, source)) {
		entries.push({
			kind: 'autoFilter',
			contained: rangeContainsRange(source, autoFilterRange),
			ref: rangeToA1(autoFilterRange),
			label: 'sheet autoFilter',
		})
	}
	const sortStateRange = sheet.sortState ? parseOptionalRange(sheet.sortState.ref) : null
	if (sortStateRange && rangesOverlap(sortStateRange, source)) {
		entries.push({
			kind: 'sortState',
			contained: rangeContainsRange(source, sortStateRange),
			ref: rangeToA1(sortStateRange),
			label: 'sheet sortState',
		})
	}
	return entries
}

function planProtectedRangeTransfer(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	mode: PasteMode,
	operation: 'copyRange' | 'moveRange',
): Result<ProtectedRangeTransferPlan> {
	if (mode !== 'all') return ok({ sheet: sourceSheet, entries: [] })
	const intersecting = collectIntersectingProtectedRanges(sourceSheet, source)
	if (intersecting.length === 0) return ok({ sheet: sourceSheet, entries: [] })
	const first = intersecting[0]
	if (!first) return ok({ sheet: sourceSheet, entries: [] })
	if (operation === 'copyRange') {
		return err(unsupportedProtectedRangeTransferError('copy', source, first))
	}
	if (sourceSheet !== targetSheet) {
		return err(unsupportedProtectedRangeTransferError('move to another sheet', source, first))
	}

	const entries: ProtectedRangeTransferEntry[] = []
	for (const entry of intersecting) {
		if (entry.partialRange) return err(partialProtectedRangeTransferError(source, entry))
		const translated = translateProtectedRange(entry.range, source, rowDelta, colDelta)
		if (!translated.ok) return translated
		entries.push({ index: entry.index, range: translated.value })
	}
	return ok({ sheet: sourceSheet, entries })
}

function applyProtectedRangeTransfer(plan: ProtectedRangeTransferPlan): void {
	for (const entry of plan.entries) {
		if (plan.sheet.protectedRanges[entry.index]) {
			plan.sheet.protectedRanges[entry.index] = entry.range
		}
	}
}

function collectIntersectingProtectedRanges(
	sheet: Sheet,
	source: RangeRef,
): ProtectedRangeIntersection[] {
	const entries: ProtectedRangeIntersection[] = []
	for (const [index, range] of sheet.protectedRanges.entries()) {
		const refs = parseSqref(range.sqref)
		if (refs.length === 0) continue
		const overlappingRanges = refs.filter((ref) => rangesOverlap(ref, source))
		if (overlappingRanges.length === 0) continue
		const partialRange = overlappingRanges.find((ref) => !rangeContainsRange(source, ref))
		entries.push({
			range,
			index,
			ref: rangesToSqref(overlappingRanges),
			label: range.name ? `protected range "${range.name}"` : `protected range ${index + 1}`,
			overlappingRanges,
			...(partialRange ? { partialRange } : {}),
		})
	}
	return entries
}

function translateProtectedRange(
	entry: SheetProtectedRange,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
): Result<SheetProtectedRange> {
	const ranges = parseSqref(entry.sqref)
	if (ranges.length === 0) {
		return err(invalidProtectedRangeTransferError(entry.sqref, source))
	}
	const translated = ranges.map((range) =>
		rangeContainsRange(source, range) ? shiftRange(range, rowDelta, colDelta) : range,
	)
	return ok({ ...entry, sqref: rangesToSqref(translated) })
}

function planAdvancedFilterTransfer(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	mode: PasteMode,
	operation: 'copyRange' | 'moveRange',
): Result<AdvancedFilterTransferPlan> {
	if (mode !== 'all') return ok({ sheet: sourceSheet, entries: [] })
	const intersecting = collectIntersectingAdvancedFilters(sourceSheet, source)
	if (intersecting.length === 0) return ok({ sheet: sourceSheet, entries: [] })
	const first = intersecting[0]
	if (!first) return ok({ sheet: sourceSheet, entries: [] })
	if (operation === 'copyRange') {
		return err(unsupportedAdvancedFilterTransferError('copy', source, first))
	}
	if (sourceSheet !== targetSheet) {
		return err(unsupportedAdvancedFilterTransferError('move to another sheet', source, first))
	}

	const entries: AdvancedFilterTransferEntry[] = []
	for (const entry of intersecting) {
		if (!entry.contained) return err(partialAdvancedFilterTransferError(source, entry))
		const translated = translateAdvancedFilter(entry.filter, rowDelta, colDelta, source)
		if (!translated.ok) return translated
		entries.push({ index: entry.index, filter: translated.value })
	}
	return ok({ sheet: sourceSheet, entries })
}

function applyAdvancedFilterTransfer(plan: AdvancedFilterTransferPlan): void {
	for (const entry of plan.entries) {
		if (plan.sheet.advancedFilters[entry.index]) {
			plan.sheet.advancedFilters[entry.index] = entry.filter
		}
	}
}

function findPartialMetadataRangeTransfer(
	sheet: Sheet,
	source: RangeRef,
	mode: PasteMode,
): PartialMetadataRangeTransfer | null {
	if (pasteValidations(mode)) {
		const dataValidation = findPartialSqrefMetadataRange(
			sheet.dataValidations,
			source,
			'partial-data-validation-range-transfer',
			'dataValidation',
			'data validation',
		)
		if (dataValidation) return dataValidation

		const x14DataValidation = findPartialSqrefMetadataRange(
			sheet.x14DataValidations,
			source,
			'partial-x14-data-validation-range-transfer',
			'x14DataValidation',
			'x14 data validation',
		)
		if (x14DataValidation) return x14DataValidation
	}

	if (pasteConditionalFormats(mode)) {
		const conditionalFormat = findPartialSqrefMetadataRange(
			sheet.conditionalFormats,
			source,
			'partial-conditional-format-range-transfer',
			'conditionalFormat',
			'conditional format',
		)
		if (conditionalFormat) return conditionalFormat

		const x14ConditionalFormat = findPartialSqrefMetadataRange(
			sheet.x14ConditionalFormats,
			source,
			'partial-x14-conditional-format-range-transfer',
			'x14ConditionalFormat',
			'x14 conditional format',
		)
		if (x14ConditionalFormat) return x14ConditionalFormat
	}

	return null
}

function findPartialSqrefMetadataRange(
	entries: readonly { readonly sqref: string; readonly deleted?: boolean }[],
	source: RangeRef,
	detailKind: PartialMetadataRangeTransfer['detailKind'],
	metadataKind: PartialMetadataRangeTransfer['metadataKind'],
	label: string,
): PartialMetadataRangeTransfer | null {
	for (const [index, entry] of entries.entries()) {
		if (entry.deleted) continue
		for (const range of parseSqref(entry.sqref)) {
			if (!rangesOverlap(range, source) || rangeContainsRange(source, range)) continue
			return {
				detailKind,
				metadataKind,
				index,
				ref: rangeToA1(range),
				label: `${label} ${index + 1}`,
			}
		}
	}
	return null
}

function collectIntersectingAdvancedFilters(
	sheet: Sheet,
	source: RangeRef,
): AdvancedFilterIntersection[] {
	const entries: AdvancedFilterIntersection[] = []
	for (const [index, filter] of sheet.advancedFilters.entries()) {
		const range = advancedFilterRange(filter)
		if (!range || !rangesOverlap(range, source)) continue
		const ref = rangeToA1(range)
		entries.push({
			filter,
			index,
			contained: rangeContainsRange(source, range),
			ref,
			label: filter.viewName
				? `advanced filter "${filter.viewName}"`
				: `advanced filter ${index + 1}`,
		})
	}
	return entries
}

function advancedFilterRange(filter: SheetAdvancedFilterInfo): RangeRef | null {
	for (const ref of [filter.autoFilter?.ref, filter.ref]) {
		const range = ref ? parseOptionalRange(ref) : null
		if (range) return range
	}
	return null
}

function parseOptionalRange(ref: string): RangeRef | null {
	try {
		return parseRange(ref)
	} catch {
		return null
	}
}

function translateAdvancedFilter(
	filter: SheetAdvancedFilterInfo,
	rowDelta: number,
	colDelta: number,
	source: RangeRef,
): Result<SheetAdvancedFilterInfo> {
	const owner: RangeTransferOwner = {
		label: filter.viewName ? `advanced filter "${filter.viewName}"` : 'advanced filter metadata',
		kind: 'advanced-filter',
	}
	const ref = filter.ref
		? translateContainedRangeRef(filter.ref, rowDelta, colDelta, source, owner)
		: ok(undefined)
	if (!ref.ok) return ref
	const autoFilter = filter.autoFilter
		? translateAutoFilter(filter.autoFilter, rowDelta, colDelta, source, owner)
		: ok(undefined)
	if (!autoFilter.ok) return autoFilter
	const nextRef = autoFilter.value?.ref ?? ref.value
	return ok({
		...filter,
		...(nextRef !== undefined ? { ref: nextRef } : {}),
		...(autoFilter.value !== undefined ? { autoFilter: autoFilter.value } : {}),
		filterColumnCount: autoFilter.value
			? autoFilter.value.columns.length
			: filter.filterColumnCount,
		sortConditionCount: autoFilter.value
			? (autoFilter.value.sortState?.conditions.length ?? 0)
			: filter.sortConditionCount,
	})
}

function translateAutoFilter(
	autoFilter: AutoFilter,
	rowDelta: number,
	colDelta: number,
	source: RangeRef,
	owner: RangeTransferOwner,
): Result<AutoFilter> {
	const ref = translateContainedRangeRef(autoFilter.ref, rowDelta, colDelta, source, owner)
	if (!ref.ok) return ref
	const sortState = autoFilter.sortState
		? translateSortState(autoFilter.sortState, rowDelta, colDelta, source, owner)
		: ok(undefined)
	if (!sortState.ok) return sortState
	const { ref: _ref, columns: _columns, sortState: _sortState, ...rest } = autoFilter
	return ok({
		...rest,
		ref: ref.value,
		columns: autoFilter.columns.map((column) => ({ ...column })),
		...(sortState.value ? { sortState: sortState.value } : {}),
	})
}

function translateSortState(
	sortState: SortState,
	rowDelta: number,
	colDelta: number,
	source: RangeRef,
	owner: RangeTransferOwner,
): Result<SortState> {
	const ref = translateContainedRangeRef(sortState.ref, rowDelta, colDelta, source, owner)
	if (!ref.ok) return ref
	const conditions: SortState['conditions'][number][] = []
	for (const condition of sortState.conditions) {
		const conditionRef = translateContainedRangeRef(
			condition.ref,
			rowDelta,
			colDelta,
			source,
			owner,
		)
		if (!conditionRef.ok) return conditionRef
		conditions.push({ ...condition, ref: conditionRef.value })
	}
	return ok({ ...sortState, ref: ref.value, conditions })
}

interface RangeTransferOwner {
	readonly label: string
	readonly kind: 'advanced-filter' | 'sheet-filter' | 'sheet-sort-state'
}

function translateContainedRangeRef(
	ref: string,
	rowDelta: number,
	colDelta: number,
	source: RangeRef,
	owner: RangeTransferOwner,
): Result<string> {
	let range: RangeRef
	try {
		range = parseRange(ref)
	} catch {
		return err(
			ascendError('VALIDATION_ERROR', `Cannot move ${owner.label} with invalid range ${ref}`, {
				refs: [ref, rangeToA1(source)],
				suggestedFix:
					'Repair or remove the filter or sort metadata before moving the filtered range.',
				details: {
					kind: `invalid-${owner.kind}-range-transfer`,
					reference: ref,
					source: rangeToA1(source),
				},
			}),
		)
	}
	if (!rangeContainsRange(source, range)) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`Cannot move ${rangeToA1(source)} because ${owner.label} references ${ref} outside the moved cells`,
				{
					refs: [rangeToA1(source), ref],
					suggestedFix:
						'Move the full filter or sort range and its references together, or remove the metadata before moving the cells.',
					details: {
						kind: `${owner.kind}-reference-outside-transfer`,
						reference: ref,
						source: rangeToA1(source),
					},
				},
			),
		)
	}
	return ok(rangeToA1(shiftRange(range, rowDelta, colDelta)))
}

function collectIntersectingVisualAnchors(
	sheet: Sheet,
	source: RangeRef,
): VisualAnchorIntersection[] {
	const entries: VisualAnchorIntersection[] = []
	for (const [index, image] of sheet.imageRefs.entries()) {
		if (!image.anchor) continue
		const range = sheetImageAnchorRange(image.anchor)
		if (!range || !rangesOverlap(range, source)) continue
		entries.push({
			kind: 'image',
			index,
			anchor: image.anchor,
			contained: rangeContainsRange(source, range),
			ref: rangeToA1(range),
			label: image.name ? `image "${image.name}"` : `image ${index + 1}`,
		})
	}
	for (const [index, object] of sheet.drawingObjectRefs.entries()) {
		if (!object.anchor) continue
		const range = sheetImageAnchorRange(object.anchor)
		if (!range || !rangesOverlap(range, source)) continue
		entries.push({
			kind: 'drawingObject',
			index,
			anchor: object.anchor,
			contained: rangeContainsRange(source, range),
			ref: rangeToA1(range),
			label: object.name ? `drawing object "${object.name}"` : `drawing object ${index + 1}`,
		})
	}
	return entries
}

function sheetImageAnchorRange(anchor: SheetImageAnchor): RangeRef | null {
	switch (anchor.kind) {
		case 'absolute':
			return null
		case 'oneCell':
			return markerRange(anchor.from, anchor.from)
		case 'twoCell':
			return markerRange(anchor.from, anchor.to)
	}
}

function markerRange(from: SheetAnchorMarker, to: SheetAnchorMarker): RangeRef {
	return {
		start: {
			row: Math.min(from.row, to.row),
			col: Math.min(from.col, to.col),
		},
		end: {
			row: Math.max(from.row, to.row),
			col: Math.max(from.col, to.col),
		},
	}
}

function translateSheetImageAnchor(
	anchor: SheetImageAnchor,
	rowDelta: number,
	colDelta: number,
): SheetImageAnchor {
	switch (anchor.kind) {
		case 'absolute':
			return anchor
		case 'oneCell':
			return { ...anchor, from: translateAnchorMarker(anchor.from, rowDelta, colDelta) }
		case 'twoCell':
			return {
				...anchor,
				from: translateAnchorMarker(anchor.from, rowDelta, colDelta),
				to: translateAnchorMarker(anchor.to, rowDelta, colDelta),
			}
	}
}

function translateAnchorMarker(
	marker: SheetAnchorMarker,
	rowDelta: number,
	colDelta: number,
): SheetAnchorMarker {
	return {
		...marker,
		row: Math.max(0, marker.row + rowDelta),
		col: Math.max(0, marker.col + colDelta),
	}
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

function overlappingMoveRangeError(
	source: RangeRef,
	target: RangeRef,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		'Cannot move a range onto an overlapping target range on the same sheet',
		{
			refs: [rangeToA1(source), rangeToA1(target)],
			suggestedFix:
				'Choose a target range that does not overlap the moved cells, or copy the range first and clear the source explicitly after reviewing the result.',
			details: {
				kind: 'overlapping-move-range',
				source: rangeToA1(source),
				target: rangeToA1(target),
			},
		},
	)
}

function unsupportedVisualTransferError(
	action: 'copy' | 'move to another sheet',
	source: RangeRef,
	entry: VisualAnchorIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${action} ${rangeToA1(source)} because it contains ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				action === 'copy'
					? 'Copy the cell data without mode "all", or insert a new image/drawing object explicitly after copying the range.'
					: 'Move the range within the same sheet, or recreate the image/drawing object on the target sheet explicitly after moving the cells.',
			details: {
				kind: 'unsupported-visual-range-transfer',
				action,
				visualKind: entry.kind,
				visualRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function partialVisualTransferError(
	source: RangeRef,
	entry: VisualAnchorIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot move ${rangeToA1(source)} because it partially overlaps ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				'Move the full visual anchor range, move the cells without mode "all", or reposition the image/drawing object explicitly before moving the cells.',
			details: {
				kind: 'partial-visual-range-transfer',
				visualKind: entry.kind,
				visualRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function unsupportedSheetFilterTransferError(
	action: 'copy' | 'move to another sheet',
	source: RangeRef,
	entry: SheetFilterIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${action} ${rangeToA1(source)} because it contains ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				action === 'copy'
					? 'Copy the cells without mode "all", or recreate the filter or sort metadata explicitly after copying the range.'
					: 'Move the range within the same sheet, or recreate the filter or sort metadata explicitly on the target sheet after moving the cells.',
			details: {
				kind: 'unsupported-sheet-filter-range-transfer',
				action,
				filterKind: entry.kind,
				filterRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function partialSheetFilterTransferError(
	source: RangeRef,
	entry: SheetFilterIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot move ${rangeToA1(source)} because it partially overlaps ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				'Move the full filter or sort range, move the cells without mode "all", or remove the metadata before moving the cells.',
			details: {
				kind: 'partial-sheet-filter-range-transfer',
				filterKind: entry.kind,
				filterRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function unsupportedProtectedRangeTransferError(
	action: 'copy' | 'move to another sheet',
	source: RangeRef,
	entry: ProtectedRangeIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${action} ${rangeToA1(source)} because it contains ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				action === 'copy'
					? 'Copy the cells without mode "all", or recreate the protected range explicitly after copying the range.'
					: 'Move the range within the same sheet, or recreate the protected range explicitly on the target sheet after moving the cells.',
			details: {
				kind: 'unsupported-protected-range-transfer',
				action,
				protectedRangeIndex: entry.index,
				protectedRangeRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function partialProtectedRangeTransferError(
	source: RangeRef,
	entry: ProtectedRangeIntersection,
): ReturnType<typeof ascendError> {
	const partialRef = entry.partialRange ? rangeToA1(entry.partialRange) : entry.ref
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot move ${rangeToA1(source)} because it partially overlaps ${entry.label} at ${partialRef}`,
		{
			refs: [rangeToA1(source), partialRef],
			suggestedFix:
				'Move the full protected range, move the cells without mode "all", or remove the protected range before moving the cells.',
			details: {
				kind: 'partial-protected-range-transfer',
				protectedRangeIndex: entry.index,
				protectedRangeRef: partialRef,
				source: rangeToA1(source),
			},
		},
	)
}

function invalidProtectedRangeTransferError(
	ref: string,
	source: RangeRef,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot move protected range metadata with invalid sqref ${ref}`,
		{
			refs: [ref, rangeToA1(source)],
			suggestedFix: 'Repair or remove the protected range before moving the cells.',
			details: {
				kind: 'invalid-protected-range-transfer',
				protectedRangeRef: ref,
				source: rangeToA1(source),
			},
		},
	)
}

function unsupportedAdvancedFilterTransferError(
	action: 'copy' | 'move to another sheet',
	source: RangeRef,
	entry: AdvancedFilterIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${action} ${rangeToA1(source)} because it contains ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				action === 'copy'
					? 'Copy the cells without mode "all", or recreate the advanced filter explicitly after copying the range.'
					: 'Move the range within the same sheet, or recreate the advanced filter explicitly on the target sheet after moving the cells.',
			details: {
				kind: 'unsupported-advanced-filter-range-transfer',
				action,
				filterIndex: entry.index,
				filterRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function partialAdvancedFilterTransferError(
	source: RangeRef,
	entry: AdvancedFilterIntersection,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot move ${rangeToA1(source)} because it partially overlaps ${entry.label} at ${entry.ref}`,
		{
			refs: [rangeToA1(source), entry.ref],
			suggestedFix:
				'Move the full advanced filter range, move the cells without mode "all", or remove the advanced filter before moving the cells.',
			details: {
				kind: 'partial-advanced-filter-range-transfer',
				filterIndex: entry.index,
				filterRef: entry.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function partialMetadataRangeTransferError(
	action: 'copy' | 'move',
	source: RangeRef,
	blocker: PartialMetadataRangeTransfer,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot ${action} ${rangeToA1(source)} because ${blocker.label} at ${blocker.ref} partially overlaps the source range`,
		{
			refs: [rangeToA1(source), blocker.ref],
			suggestedFix:
				'Select the full validation or conditional-format range, use a paste mode that does not transfer that metadata, or remove the metadata before editing the cells.',
			details: {
				kind: blocker.detailKind,
				action,
				metadataKind: blocker.metadataKind,
				metadataIndex: blocker.index,
				metadataRef: blocker.ref,
				source: rangeToA1(source),
			},
		},
	)
}

function partialMoveFormulaReferenceError(
	blocker: PartialFormulaMoveReference,
	source: RangeRef,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot move ${rangeToA1(source)} because ${blocker.owner} contains a range reference that partially overlaps the moved cells`,
		{
			refs: [rangeToA1(source), blocker.owner, blocker.reference],
			suggestedFix:
				'Move the full referenced range, edit the reference first, or split it so it no longer spans cells that will be left behind.',
			details: {
				kind: 'partial-move-formula-reference',
				ownerKind: blocker.ownerKind,
				owner: blocker.owner,
				formula: blocker.formula,
				reference: blocker.reference,
				source: rangeToA1(source),
			},
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

function copyCommentMap(
	sourceMap: Map<string, SheetComment>,
	targetMap: Map<string, SheetComment>,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const entries = [...sourceMap.entries()].filter(([ref]) =>
		rangeContainsCell(source, parseA1(ref)),
	)
	const copied = entries.map(([ref, value]) => {
		const pos = parseA1(ref)
		const target = { row: pos.row + rowDelta, col: pos.col + colDelta }
		return [toA1(target), retargetLegacyCommentDrawing(value, target, rowDelta, colDelta)] as const
	})
	if (move) {
		for (const [ref] of entries) sourceMap.delete(ref)
	}
	for (const [ref, value] of copied) {
		targetMap.set(ref, value)
	}
}

function copyThreadedComments(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	source: RangeRef,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): void {
	const sourceEntries = sourceSheet.threadedComments
		.map((comment, index) => ({ comment, index }))
		.filter(({ comment }) => rangeContainsCell(source, parseA1(comment.ref)))
	if (sourceEntries.length === 0) return

	const sourceIndexes = new Set(sourceEntries.map((entry) => entry.index))
	const targetRange = shiftRange(source, rowDelta, colDelta)
	const targetIndexes = new Set(
		targetSheet.threadedComments
			.map((comment, index) => ({ comment, index }))
			.filter(({ comment }) => rangeContainsCell(targetRange, parseA1(comment.ref)))
			.map((entry) => entry.index),
	)
	const idMap = move
		? new Map<string, string>()
		: buildThreadedCommentIdMap(sourceSheet, targetSheet, sourceEntries)
	const copied = sourceEntries.map(({ comment }) =>
		retargetThreadedComment(comment, rowDelta, colDelta, move, idMap),
	)

	if (sourceSheet === targetSheet) {
		const removeIndexes = new Set([...targetIndexes, ...(move ? sourceIndexes : [])])
		sourceSheet.threadedComments = sourceSheet.threadedComments.filter(
			(_, index) => !removeIndexes.has(index),
		)
		sourceSheet.threadedComments.push(...copied)
		return
	}

	if (move) {
		sourceSheet.threadedComments = sourceSheet.threadedComments.filter(
			(_, index) => !sourceIndexes.has(index),
		)
	}
	targetSheet.threadedComments = targetSheet.threadedComments.filter(
		(_, index) => !targetIndexes.has(index),
	)
	targetSheet.threadedComments.push(...copied)
}

function buildThreadedCommentIdMap(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	entries: readonly { readonly comment: Sheet['threadedComments'][number] }[],
): Map<string, string> {
	const existingIds = new Set<string>()
	for (const sheet of sourceSheet === targetSheet ? [sourceSheet] : [sourceSheet, targetSheet]) {
		for (const comment of sheet.threadedComments) {
			if (comment.id) existingIds.add(comment.id)
		}
	}
	const idMap = new Map<string, string>()
	for (const { comment } of entries) {
		if (!comment.id || idMap.has(comment.id)) continue
		idMap.set(comment.id, nextThreadedCommentId(comment.id, existingIds))
	}
	return idMap
}

function nextThreadedCommentId(baseId: string, existingIds: Set<string>): string {
	const base = `${baseId}-copy`
	let candidate = base
	let suffix = 2
	while (existingIds.has(candidate)) candidate = `${base}-${suffix++}`
	existingIds.add(candidate)
	return candidate
}

function retargetThreadedComment(
	comment: Sheet['threadedComments'][number],
	rowDelta: number,
	colDelta: number,
	move: boolean,
	idMap: Map<string, string>,
): Sheet['threadedComments'][number] {
	const pos = parseA1(comment.ref)
	const ref = toA1({ row: pos.row + rowDelta, col: pos.col + colDelta })
	if (move) return { ...comment, ref }

	const { id, parentId, ...rest } = comment
	const nextId = id ? idMap.get(id) : undefined
	const nextParentId = parentId ? idMap.get(parentId) : undefined
	return {
		...rest,
		ref,
		...(nextId ? { id: nextId } : {}),
		...(nextParentId ? { parentId: nextParentId } : {}),
	}
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
	const copied = entries.map(([ref, value]) => {
		const pos = parseA1(ref)
		return [toA1({ row: pos.row + rowDelta, col: pos.col + colDelta }), { ...value }] as const
	})
	if (move) {
		for (const [ref] of entries) sourceMap.delete(ref)
	}
	for (const [ref, value] of copied) {
		targetMap.set(ref, value)
	}
}

function retargetLegacyCommentDrawing(
	comment: SheetComment,
	target: { readonly row: number; readonly col: number },
	rowDelta: number,
	colDelta: number,
): SheetComment {
	const drawing = comment.legacyDrawing
	if (!drawing) return { ...comment }
	return {
		...comment,
		legacyDrawing: {
			...drawing,
			...(drawing.row !== undefined ? { row: target.row } : {}),
			...(drawing.column !== undefined ? { column: target.col } : {}),
			...(drawing.anchor
				? { anchor: translateLegacyCommentAnchor(drawing.anchor, rowDelta, colDelta) }
				: {}),
		},
	}
}

function translateLegacyCommentAnchor(
	anchor: NonNullable<NonNullable<SheetComment['legacyDrawing']>['anchor']>,
	rowDelta: number,
	colDelta: number,
): NonNullable<NonNullable<SheetComment['legacyDrawing']>['anchor']> {
	const next = [...anchor] as [number, number, number, number, number, number, number, number]
	next[0] = Math.max(0, next[0] + colDelta)
	next[2] = Math.max(0, next[2] + rowDelta)
	next[4] = Math.max(next[0], next[4] + colDelta)
	next[6] = Math.max(next[2], next[6] + rowDelta)
	return next
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
	for (const token of splitSqrefTokens(sqref)) {
		if (!token) continue
		try {
			ranges.push(parseRange(token))
		} catch {
			return []
		}
	}
	return ranges
}

function splitSqrefTokens(sqref: string): string[] {
	const refs: string[] = []
	let current = ''
	let quoted = false
	for (let index = 0; index < sqref.length; index++) {
		const ch = sqref[index] ?? ''
		if (ch === "'") {
			current += ch
			if (quoted && sqref[index + 1] === "'") {
				current += "'"
				index++
			} else {
				quoted = !quoted
			}
		} else if (/\s/.test(ch) && !quoted) {
			if (current) {
				refs.push(current)
				current = ''
			}
		} else {
			current += ch
		}
	}
	if (current) refs.push(current)
	return refs
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
		...(range.sheet !== undefined ? { sheet: range.sheet } : {}),
	}
}

function rangesToSqref(ranges: readonly RangeRef[]): string {
	return ranges.map(rangeToA1).join(' ')
}

function rangeToA1(range: RangeRef): string {
	const start = toA1(range.start)
	const end = toA1(range.end)
	const body = start === end ? start : `${start}:${end}`
	return range.sheet !== undefined ? `${formatSheetName(range.sheet)}!${body}` : body
}

function formatSheetName(sheet: string): string {
	if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet)) return sheet
	return `'${sheet.replace(/'/g, "''")}'`
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

	sheet.ensureWritable()
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

	sheetResult.value.ensureWritable()
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
