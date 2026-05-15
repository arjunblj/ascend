import type {
	Cell,
	CellFormulaBinding,
	CellStyle,
	RangeRef,
	Sheet,
	StyleId,
	Workbook,
} from '@ascend/core'
import { DEFAULT_STYLE_ID, parseA1Safe, parseRange, toA1 } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { dateToSerial, printFormulaWithOffset } from '@ascend/formulas'
import type {
	AscendError,
	CellValue,
	InputValue,
	Operation,
	Result,
	StyleInput,
} from '@ascend/schema'
import { ascendError, booleanValue, EMPTY, err, numberValue, ok, stringValue } from '@ascend/schema'
import { resolveCellFormulaText } from '../analysis.ts'
import { type CellKey, cellKey } from '../dep-graph.ts'

export interface PatchResult {
	readonly affectedCells: string[]
	readonly sheetsModified: string[]
	readonly recalcRequired: boolean
	readonly warnings?: readonly AscendError[]
}

export const DEFAULT_SID = DEFAULT_STYLE_ID

export function inputToCellValue(
	input: InputValue,
	dateSystem: '1900' | '1904' = '1900',
): CellValue {
	if (input === null) return EMPTY
	if (typeof input === 'number') return numberValue(input)
	if (typeof input === 'string') return stringValue(input)
	if (typeof input === 'boolean') return booleanValue(input)
	if (input instanceof Date) {
		return {
			kind: 'date',
			serial: dateToSerial(input.getFullYear(), input.getMonth() + 1, input.getDate(), dateSystem),
		}
	}
	return EMPTY
}

export function getSheet(workbook: Workbook, name: string): Result<Sheet> {
	const sheet = workbook.getSheet(name)
	if (!sheet) {
		const available = workbook.sheets.map((s) => s.name).join(', ')
		return err(
			ascendError('SHEET_NOT_FOUND', `Sheet "${name}" not found`, {
				suggestedFix: available ? `Available sheets: ${available}` : 'Workbook has no sheets',
			}),
		)
	}
	return ok(sheet)
}

export function patch(
	affected: string[],
	sheets: string[],
	recalc = false,
	warnings?: readonly AscendError[],
): PatchResult {
	const result = {
		affectedCells: affected,
		sheetsModified: [...new Set(sheets)],
		recalcRequired: recalc,
	}
	return warnings && warnings.length > 0 ? { ...result, warnings } : result
}

export function cell(value: CellValue, formula: string | null, styleId: StyleId): Cell {
	return { value, formula, styleId }
}

export function cellWithExisting(
	value: CellValue,
	formula: string | null,
	styleId: StyleId,
	existingFormulaInfo?: Cell['formulaInfo'],
): Cell {
	return {
		value,
		formula,
		styleId,
		...(formula !== null && existingFormulaInfo ? { formulaInfo: existingFormulaInfo } : {}),
	}
}

export function cellPreservingFormulaInfo(
	value: CellValue | undefined,
	formula: string | null,
	styleId: StyleId,
	formulaInfo?: Cell['formulaInfo'],
): Cell {
	return {
		value: value ?? EMPTY,
		formula,
		styleId,
		...(formulaInfo ? { formulaInfo } : {}),
	}
}

export function safeParseRange(range: string): Result<RangeRef> {
	try {
		return ok(parseRange(range))
	} catch {
		return err(
			ascendError('INVALID_RANGE', `Invalid range: ${range}`, {
				suggestedFix: 'Expected format: A1 for single cell, or A1:B10 for range',
			}),
		)
	}
}

export interface LegacyArrayFormulaBlocker {
	readonly ref: string
}

export interface LegacyArrayFormulaIntersection {
	readonly ref: string
	readonly targetRef: string
}

export interface LegacyArrayFormulaIndex {
	readonly isEmpty: () => boolean
	readonly first: () => LegacyArrayFormulaIntersection | null
	readonly findCell: (row: number, col: number) => LegacyArrayFormulaBlocker | null
	readonly findIntersection: (range: RangeRef) => LegacyArrayFormulaIntersection | null
}

export function createLegacyArrayFormulaIndex(sheet: Sheet): LegacyArrayFormulaIndex {
	if (sheet.cells.formulaInfoCellCount() === 0) return EMPTY_LEGACY_ARRAY_INDEX
	const seen = new Set<string>()
	const ranges: Array<{ readonly ref: string; readonly range: RangeRef }> = []
	for (const [, , cell] of sheet.cells.iterate()) {
		const binding = cell.formulaInfo
		if (binding?.kind !== 'array' || !binding.ref || seen.has(binding.ref)) continue
		try {
			ranges.push({ ref: binding.ref, range: parseRange(binding.ref) })
			seen.add(binding.ref)
		} catch {
			seen.add(binding.ref)
		}
	}
	return {
		isEmpty() {
			return ranges.length === 0
		},
		first() {
			const entry = ranges[0]
			if (!entry) return null
			return { ref: entry.ref, targetRef: toA1(entry.range.start) }
		},
		findCell(row, col) {
			for (const entry of ranges) {
				if (
					row >= entry.range.start.row &&
					row <= entry.range.end.row &&
					col >= entry.range.start.col &&
					col <= entry.range.end.col
				) {
					return { ref: entry.ref }
				}
			}
			return null
		},
		findIntersection(range) {
			for (const entry of ranges) {
				const row = Math.max(entry.range.start.row, range.start.row)
				const col = Math.max(entry.range.start.col, range.start.col)
				if (row <= Math.min(entry.range.end.row, range.end.row)) {
					if (col <= Math.min(entry.range.end.col, range.end.col)) {
						return { ref: entry.ref, targetRef: toA1({ row, col }) }
					}
				}
			}
			return null
		},
	}
}

const EMPTY_LEGACY_ARRAY_INDEX: LegacyArrayFormulaIndex = {
	isEmpty: () => true,
	first: () => null,
	findCell: () => null,
	findIntersection: () => null,
}

export function legacyArrayFormulaEditError(
	targetRef: string,
	arrayRef: string,
): ReturnType<typeof ascendError> {
	return ascendError(
		'VALIDATION_ERROR',
		`Cannot edit ${targetRef} because it is part of legacy array formula ${arrayRef}`,
		{
			refs: [targetRef, arrayRef],
			suggestedFix:
				'Replace the full array formula range with an explicit array-formula operation before editing individual cells.',
		},
	)
}

export function translateFormula(node: FormulaNode, rowDelta: number, colDelta: number): string {
	return printFormulaWithOffset(node, rowDelta, colDelta)
}

export function collectRangeCells(
	sheet: Sheet,
	range: RangeRef,
): Array<{ row: number; col: number; cell: Cell | undefined }> {
	const cells: Array<{ row: number; col: number; cell: Cell | undefined }> = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			cells.push({ row, col, cell: sheet.cells.get(row, col) })
		}
	}
	return cells
}

export function materializeFormulaBindingGroupsForRefs(
	workbook: Workbook,
	sheet: Sheet,
	refs: Iterable<{ readonly row: number; readonly col: number }>,
	options: { readonly blockedSpillBlockers?: boolean } = {},
): Set<string> {
	const affected = new Set<string>()
	const refList = [...refs]
	for (const ref of refList) {
		const existing = sheet.cells.get(ref.row, ref.col)
		const binding = existing?.formulaInfo
		if (!binding) continue
		const materialized =
			binding.kind === 'shared'
				? materializeSharedFormulaGroup(workbook, sheet, binding)
				: isSpillGroupBinding(binding)
					? materializeSpillFormulaGroup(sheet, binding, ref)
					: new Set<string>()
		for (const entry of materialized) affected.add(entry)
	}
	for (const entry of materializeDataTableFormulaGroupsForRefs(sheet, refList)) {
		affected.add(entry)
	}
	if (options.blockedSpillBlockers !== false) {
		for (const entry of materializeBlockedSpillFormulaGroupsForRefs(sheet, refList)) {
			affected.add(entry)
		}
	}
	return affected
}

export function materializeFormulaBindingGroupsForRangeEdit(
	workbook: Workbook,
	sheet: Sheet,
	range: RangeRef,
): Set<string> {
	const refs: Array<{ readonly row: number; readonly col: number }> = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) refs.push({ row, col })
	}
	return materializeFormulaBindingGroupsForRefs(workbook, sheet, refs)
}

export function materializeFormulaBindingGroupsForFormulaClear(
	workbook: Workbook,
	sheet: Sheet,
	range: RangeRef,
): Set<string> {
	const refs: Array<{ readonly row: number; readonly col: number }> = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) refs.push({ row, col })
	}
	return materializeFormulaBindingGroupsForRefs(workbook, sheet, refs, {
		blockedSpillBlockers: false,
	})
}

export function materializeDataTableFormulaGroupsForRangeEdit(
	sheet: Sheet,
	range: RangeRef,
): Set<string> {
	const refs: Array<{ readonly row: number; readonly col: number }> = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) refs.push({ row, col })
	}
	return materializeDataTableFormulaGroupsForRefs(sheet, refs)
}

export function materializeBlockedSpillFormulaGroupsForRangeEdit(
	sheet: Sheet,
	range: RangeRef,
): Set<string> {
	const refs: Array<{ readonly row: number; readonly col: number }> = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) refs.push({ row, col })
	}
	return materializeBlockedSpillFormulaGroupsForRefs(sheet, refs)
}

function materializeSharedFormulaGroup(
	workbook: Workbook,
	sheet: Sheet,
	binding: Extract<CellFormulaBinding, { kind: 'shared' }>,
): Set<string> {
	const affected = new Set<string>()
	const sheetIndex = workbook.sheets.indexOf(sheet)
	if (sheetIndex < 0) return affected
	const materialized: Array<{
		readonly row: number
		readonly col: number
		readonly formula: string | null
	}> = []
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (!sameSharedFormulaGroup(binding, cell.formulaInfo)) continue
		materialized.push({
			row,
			col,
			formula: resolveCellFormulaText(workbook, sheetIndex, row, col, cell),
		})
	}
	for (const entry of materialized) {
		const current = sheet.cells.get(entry.row, entry.col)
		if (!current) continue
		sheet.cells.set(
			entry.row,
			entry.col,
			cellWithExisting(current.value, entry.formula, current.styleId ?? DEFAULT_SID),
		)
		affected.add(toA1({ row: entry.row, col: entry.col }))
	}
	return affected
}

function sameSharedFormulaGroup(
	binding: Extract<CellFormulaBinding, { kind: 'shared' }>,
	candidate: CellFormulaBinding | undefined,
): candidate is Extract<CellFormulaBinding, { kind: 'shared' }> {
	if (candidate?.kind !== 'shared') return false
	if (binding.sharedIndex !== undefined) return candidate.sharedIndex === binding.sharedIndex
	return candidate.masterRef === binding.masterRef
}

function materializeSpillFormulaGroup(
	sheet: Sheet,
	binding: Extract<CellFormulaBinding, { kind: 'dynamicArray' | 'spill' | 'blockedSpill' }>,
	anchor: { readonly row: number; readonly col: number },
): Set<string> {
	const affected = new Set<string>()
	const dynamicAnchorRef =
		binding.kind === 'dynamicArray' ? `${sheet.name}!${toA1(anchor)}` : undefined
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (!sameSpillFormulaGroup(binding, cell.formulaInfo, dynamicAnchorRef)) continue
		sheet.cells.set(row, col, cellWithExisting(cell.value, null, cell.styleId ?? DEFAULT_SID))
		affected.add(toA1({ row, col }))
	}
	return affected
}

function materializeDataTableFormulaGroupsForRefs(
	sheet: Sheet,
	refs: readonly { readonly row: number; readonly col: number }[],
): Set<string> {
	const affected = new Set<string>()
	if (refs.length === 0 || sheet.cells.formulaInfoCellCount() === 0) return affected
	for (const [row, col, cell] of sheet.cells.iterate()) {
		const binding = cell.formulaInfo
		if (binding?.kind !== 'dataTable') continue
		const tableRange = dataTableBindingRange(binding, row, col)
		if (!refs.some((ref) => rangeContainsCell(tableRange, ref))) continue
		sheet.cells.set(row, col, cellWithExisting(cell.value, null, cell.styleId ?? DEFAULT_SID))
		affected.add(toA1({ row, col }))
	}
	return affected
}

function materializeBlockedSpillFormulaGroupsForRefs(
	sheet: Sheet,
	refs: readonly { readonly row: number; readonly col: number }[],
): Set<string> {
	const affected = new Set<string>()
	if (refs.length === 0 || sheet.cells.formulaInfoCellCount() === 0) return affected
	for (const [row, col, cell] of sheet.cells.iterate()) {
		const binding = cell.formulaInfo
		if (binding?.kind !== 'blockedSpill') continue
		if (
			!refs.some(
				(ref) =>
					formulaBindingRangeContainsCell(binding.ref, sheet.name, ref) ||
					binding.blockingRefs.some((blockingRef) =>
						formulaBindingRangeContainsCell(blockingRef, sheet.name, ref),
					),
			)
		) {
			continue
		}
		sheet.cells.set(
			row,
			col,
			cellWithExisting(cell.value, cell.formula, cell.styleId ?? DEFAULT_SID),
		)
		affected.add(toA1({ row, col }))
	}
	return affected
}

function formulaBindingRangeContainsCell(
	refText: string,
	sheetName: string,
	ref: { readonly row: number; readonly col: number },
): boolean {
	try {
		const range = parseRange(refText)
		if (range.sheet !== undefined && range.sheet !== sheetName) return false
		return rangeContainsCell(range, ref)
	} catch {
		return false
	}
}

function dataTableBindingRange(
	binding: Extract<CellFormulaBinding, { kind: 'dataTable' }>,
	row: number,
	col: number,
): RangeRef {
	if (binding.ref) {
		try {
			return parseRange(binding.ref)
		} catch {
			// Fall through to the anchor cell when imported metadata has an invalid ref.
		}
	}
	return { start: { row, col }, end: { row, col } }
}

function rangeContainsCell(
	range: RangeRef,
	ref: { readonly row: number; readonly col: number },
): boolean {
	return (
		ref.row >= range.start.row &&
		ref.row <= range.end.row &&
		ref.col >= range.start.col &&
		ref.col <= range.end.col
	)
}

function isSpillGroupBinding(
	binding: CellFormulaBinding,
): binding is Extract<CellFormulaBinding, { kind: 'dynamicArray' | 'spill' | 'blockedSpill' }> {
	return (
		binding.kind === 'dynamicArray' || binding.kind === 'spill' || binding.kind === 'blockedSpill'
	)
}

function sameSpillFormulaGroup(
	binding: Extract<CellFormulaBinding, { kind: 'dynamicArray' | 'spill' | 'blockedSpill' }>,
	candidate: CellFormulaBinding | undefined,
	dynamicAnchorRef?: string,
): boolean {
	if (!candidate) return false
	if (binding.kind === 'dynamicArray') {
		if (candidate.kind === 'dynamicArray') return candidate.metadataIndex === binding.metadataIndex
		return (
			dynamicAnchorRef !== undefined &&
			(candidate.kind === 'spill' || candidate.kind === 'blockedSpill') &&
			candidate.anchorRef === dynamicAnchorRef
		)
	}
	if (candidate.kind !== 'spill' && candidate.kind !== 'blockedSpill') return false
	return candidate.anchorRef === binding.anchorRef
}

export function shiftMerges(
	merges: RangeRef[],
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	const updated: RangeRef[] = []
	for (const m of merges) {
		const s = axis === 'row' ? m.start.row : m.start.col
		const e = axis === 'row' ? m.end.row : m.end.col

		if (delta < 0) {
			const deleteEnd = at - delta
			if (s >= at && e < deleteEnd) continue
		}

		const shift = (v: number): number => {
			if (delta > 0) return v >= at ? v + delta : v
			const deleteEnd = at - delta
			if (v >= deleteEnd) return v + delta
			if (v >= at) return at
			return v
		}

		if (axis === 'row') {
			updated.push({
				start: { row: shift(m.start.row), col: m.start.col },
				end: { row: shift(m.end.row), col: m.end.col },
			})
		} else {
			updated.push({
				start: { row: m.start.row, col: shift(m.start.col) },
				end: { row: m.end.row, col: shift(m.end.col) },
			})
		}
	}
	merges.length = 0
	merges.push(...updated)
}

export function clearFormulaMetadataForSheet(sheet: Workbook['sheets'][number]): void {
	for (const [row, col, existing] of sheet.cells.iterate()) {
		if (!existing.formulaInfo) continue
		sheet.cells.clearFormulaInfo(row, col)
	}
}

export function clearFormulaMetadata(workbook: Workbook): void {
	for (const sheet of workbook.sheets) {
		clearFormulaMetadataForSheet(sheet)
	}
}

export function mergeStyleInput(current: CellStyle, input: StyleInput): CellStyle {
	return {
		...current,
		...(input.font && { font: { ...current.font, ...input.font } }),
		...(input.fill && { fill: { ...current.fill, ...input.fill } }),
		...(input.border && { border: { ...current.border, ...input.border } }),
		...(input.alignment && { alignment: { ...current.alignment, ...input.alignment } }),
		...(input.numberFormat !== undefined && { numberFormat: input.numberFormat }),
	}
}

export function resolvePatchResultCellKeys(
	workbook: Workbook,
	sheetName: string,
	affectedCells: readonly string[],
	warnings?: AscendError[],
): CellKey[] {
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) return []
	const sheetIndex = workbook.sheets.indexOf(sheet)
	if (sheetIndex < 0) return []
	const keys: CellKey[] = []
	for (const refText of affectedCells) {
		const ref = parseA1Safe(refText)
		if (!ref) {
			warnings?.push(
				ascendError(
					'INVALID_REF',
					`Failed to resolve affected cell reference "${refText}" in sheet "${sheetName}"`,
				),
			)
			continue
		}
		keys.push(cellKey(sheetIndex, ref.row, ref.col))
	}
	return keys
}

export function operationAffectsFormulas(op: Operation): boolean {
	switch (op.op) {
		case 'setComment':
		case 'setHyperlink':
		case 'setNumberFormat':
		case 'setStyle':
		case 'freezePane':
		case 'setColWidth':
		case 'setRowHeight':
		case 'mergeCells':
		case 'unmergeCells':
		case 'groupRows':
		case 'groupCols':
		case 'rewriteExternalLink':
		case 'insertImage':
		case 'deleteImage':
			return false
		case 'setWorkbookProtection':
			return false
		case 'createTable':
		case 'deleteTable':
		case 'renameTable':
		case 'resizeTable':
			return true
		case 'clearRange':
			return op.what !== 'styles'
		default:
			return true
	}
}

export function updateSheetOutlineLevels(sheet: Sheet): void {
	const rowLevel = Math.max(0, ...[...sheet.rowDefs.values()].map((def) => def.outlineLevel ?? 0))
	const colLevel = Math.max(0, ...sheet.colDefs.map((def) => def.outlineLevel ?? 0))
	sheet.sheetFormatPr = {
		...(sheet.sheetFormatPr ?? {}),
		...(rowLevel > 0 ? { outlineLevelRow: rowLevel } : {}),
		...(colLevel > 0 ? { outlineLevelCol: colLevel } : {}),
	}
}

export function buildTableColumns(
	sheet: Sheet,
	ref: RangeRef,
	width: number,
	hasHeaders: boolean,
): Array<{ name: string }> {
	const usedNames = new Set<string>()
	const columns: Array<{ name: string }> = []
	for (let colOffset = 0; colOffset < width; colOffset++) {
		let name = `Column${colOffset + 1}`
		if (hasHeaders) {
			const cellValue = sheet.cells.get(ref.start.row, ref.start.col + colOffset)?.value
			if (cellValue?.kind === 'string' && cellValue.value.trim() !== '') {
				name = cellValue.value.trim()
			}
		}
		let candidate = name
		let suffix = 2
		while (usedNames.has(candidate.toLowerCase())) {
			candidate = `${name}_${suffix}`
			suffix++
		}
		usedNames.add(candidate.toLowerCase())
		columns.push({ name: candidate })
	}
	return columns
}
