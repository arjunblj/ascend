import type {
	AutoFilter,
	CellStyle,
	FilterColumn,
	SheetAdvancedFilterInfo,
	SheetConditionalFormat,
	SheetConditionalFormatRule,
	SheetDataValidation,
	SheetPageSetup,
	SortState,
	Workbook,
} from '@ascend/core'
import { toA1 } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok, validateExcelDefinedName } from '@ascend/schema'
import type { PatchResult } from './helpers.ts'
import {
	cellPreservingFormulaInfo,
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
			if (styleId === existingStyleId) continue
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
				cellPreservingFormulaInfo(
					sheet.cells.readValue(row, col),
					sheet.cells.readFormula(row, col) ?? null,
					styleId,
					sheet.cells.readFormulaInfo(row, col),
				),
			)
			affected.push(toA1({ row, col }))
		}
	}

	return ok(patch(affected, affected.length > 0 ? [op.sheet] : []))
}

export function handleSetConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setConditionalFormat' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const existing = sheet.conditionalFormats.findIndex((cf) => cf.sqref === op.range)
	const rule = conditionalFormatRuleFromOperation(op.rule)
	if (existing >= 0 && op.mode === 'append') {
		const cf = sheet.conditionalFormats[existing]
		if (cf) sheet.conditionalFormats[existing] = { ...cf, rules: [...cf.rules, rule] }
	} else {
		const cf: SheetConditionalFormat = { sqref: op.range, rules: [rule] }
		if (existing >= 0) sheet.conditionalFormats[existing] = cf
		else sheet.conditionalFormats.push(cf)
	}
	if (op.reassignPriorities) reassignConditionalFormatPriorities(sheet.conditionalFormats)
	return ok(patch([], [op.sheet]))
}

export function handleDeleteConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteConditionalFormat' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	if (op.range !== undefined) {
		const rangeResult = safeParseRange(op.range)
		if (!rangeResult.ok) return rangeResult
	}
	const sheet = sheetResult.value
	sheet.ensureWritable()
	if (op.range === undefined && op.priority === undefined && op.ruleIndex === undefined) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'deleteConditionalFormat requires range, priority, or ruleIndex',
				{
					suggestedFix:
						'Provide a sqref range to delete a block, or priority/ruleIndex to remove a specific rule.',
				},
			),
		)
	}
	sheet.conditionalFormats = sheet.conditionalFormats
		.map((cf) => {
			if (op.range !== undefined && cf.sqref !== op.range) return cf
			if (op.priority === undefined && op.ruleIndex === undefined) return null
			const rules = cf.rules.filter((rule, index) => {
				if (op.ruleIndex !== undefined && index !== op.ruleIndex) return true
				if (op.priority !== undefined && rule.priority !== op.priority) return true
				return false
			})
			return rules.length > 0 ? { ...cf, rules } : null
		})
		.filter((cf): cf is SheetConditionalFormat => cf !== null)
	return ok(patch([], [op.sheet]))
}

function conditionalFormatRuleFromOperation(
	rule: Extract<Operation, { op: 'setConditionalFormat' }>['rule'],
): SheetConditionalFormatRule {
	return {
		type: rule.type,
		formulas: [rule.formula, rule.formula2].filter((f): f is string => f !== undefined),
		...(rule.operator ? { operator: rule.operator } : {}),
		...(rule.priority !== undefined ? { priority: rule.priority } : {}),
		...(rule.stopIfTrue !== undefined ? { stopIfTrue: rule.stopIfTrue } : {}),
		...(rule.style ? { style: rule.style } : {}),
		...(rule.colorScale
			? {
					colorScale: {
						cfvo: rule.colorScale.cfvo.map((entry) => ({ ...entry })),
						colors: rule.colorScale.colors.map((entry) => ({ ...entry })),
					},
				}
			: {}),
		...(rule.dataBar
			? {
					dataBar: {
						...rule.dataBar,
						cfvo: rule.dataBar.cfvo.map((entry) => ({ ...entry })),
						...(rule.dataBar.color ? { color: { ...rule.dataBar.color } } : {}),
					},
				}
			: {}),
		...(rule.iconSet
			? {
					iconSet: {
						...rule.iconSet,
						cfvo: rule.iconSet.cfvo.map((entry) => ({ ...entry })),
					},
				}
			: {}),
	}
}

function reassignConditionalFormatPriorities(conditionalFormats: SheetConditionalFormat[]): void {
	let priority = 1
	for (let cfIndex = 0; cfIndex < conditionalFormats.length; cfIndex++) {
		const cf = conditionalFormats[cfIndex]
		if (!cf) continue
		conditionalFormats[cfIndex] = {
			...cf,
			rules: cf.rules.map((rule) => ({ ...rule, priority: priority++ })),
		}
	}
}

export function handleSetDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDataValidation' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
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
		...(op.rule.errorStyle !== undefined ? { errorStyle: op.rule.errorStyle } : {}),
		...(op.rule.imeMode !== undefined ? { imeMode: op.rule.imeMode } : {}),
		...(op.rule.showDropDown !== undefined ? { showDropDown: op.rule.showDropDown } : {}),
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
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.dataValidations = sheet.dataValidations.filter((dv) => dv.sqref !== op.range)
	return ok(patch([], [op.sheet]))
}

export function handleSetAutoFilter(workbook: Workbook, op: SetAutoFilterOp): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	const validation = validateFilterCriteriaUpdate('setAutoFilter', op)
	if (validation) return err(validation)
	const sheet = sheetResult.value
	sheet.ensureWritable()
	sheet.autoFilter = applyAutoFilterUpdate(sheet.autoFilter ?? { ref: op.range, columns: [] }, op)
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

export function handleSetAdvancedFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setAdvancedFilter' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	if (!Number.isInteger(op.filterIndex) || op.filterIndex < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'setAdvancedFilter filterIndex must be non-negative', {
				suggestedFix: 'Use the zero-based filterIndex from inspectSheet().advancedFilters.',
			}),
		)
	}
	if (!hasAdvancedFilterUpdate(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setAdvancedFilter requires filter or sort metadata', {
				suggestedFix:
					'Provide range, column+values, sortRef, sortBy, or descending for the custom sheet view filter.',
			}),
		)
	}
	const validation = validateFilterCriteriaUpdate('setAdvancedFilter', op)
	if (validation) return err(validation)

	const current = sheet.advancedFilters[op.filterIndex]
	if (!current?.autoFilter) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching advanced filter found', {
				suggestedFix: 'Inspect sheet.advancedFilters and provide an existing filterIndex.',
			}),
		)
	}

	sheet.ensureWritable()
	const autoFilter = applyAdvancedAutoFilterUpdate(current.autoFilter, op)
	const updated: SheetAdvancedFilterInfo = {
		...current,
		ref: autoFilter.ref,
		autoFilter,
		filterColumnCount: autoFilter.columns.length,
		sortConditionCount: autoFilter.sortState?.conditions.length ?? 0,
	}
	sheet.advancedFilters.splice(op.filterIndex, 1, updated)
	return ok(patch([], [op.sheet], false))
}

export function handleSetPageSetup(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPageSetup' }>,
): Result<PatchResult> {
	const validationError = validatePageSetupInput(op.setup)
	if (validationError) return err(validationError)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const setup: SheetPageSetup = {
		...(sheet.pageSetup ?? {}),
		...(op.setup.orientation !== undefined ? { orientation: op.setup.orientation } : {}),
		...(op.setup.paperSize !== undefined ? { paperSize: op.setup.paperSize } : {}),
		...(op.setup.scale !== undefined ? { scale: op.setup.scale } : {}),
		...(op.setup.fitToWidth !== undefined ? { fitToWidth: op.setup.fitToWidth } : {}),
		...(op.setup.fitToHeight !== undefined ? { fitToHeight: op.setup.fitToHeight } : {}),
	}
	sheet.pageSetup = setup
	if (op.setup.margins) {
		sheet.pageMargins = { ...(sheet.pageMargins ?? {}), ...op.setup.margins }
	}
	return ok(patch([], [op.sheet]))
}

function validatePageSetupInput(
	setup: Extract<Operation, { op: 'setPageSetup' }>['setup'],
): ReturnType<typeof ascendError> | null {
	if (
		setup.orientation !== undefined &&
		setup.orientation !== 'portrait' &&
		setup.orientation !== 'landscape'
	) {
		return ascendError('VALIDATION_ERROR', 'page setup orientation must be portrait or landscape')
	}
	for (const [field, value] of [
		['paperSize', setup.paperSize],
		['fitToWidth', setup.fitToWidth],
		['fitToHeight', setup.fitToHeight],
	] as const) {
		if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
			return ascendError('VALIDATION_ERROR', `page setup ${field} must be a non-negative integer`)
		}
	}
	if (setup.scale !== undefined && (!Number.isInteger(setup.scale) || setup.scale <= 0)) {
		return ascendError('VALIDATION_ERROR', 'page setup scale must be a positive integer')
	}
	if (setup.margins) {
		for (const [field, value] of Object.entries(setup.margins)) {
			if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
				return ascendError('VALIDATION_ERROR', `page margin ${field} must be non-negative`)
			}
		}
	}
	return null
}

type SetAdvancedFilterOp = Extract<Operation, { op: 'setAdvancedFilter' }>
type SetAutoFilterOp = Extract<Operation, { op: 'setAutoFilter' }>
type FilterCriteriaUpdateOp = SetAutoFilterOp | SetAdvancedFilterOp

function hasAdvancedFilterUpdate(op: SetAdvancedFilterOp): boolean {
	return (
		op.range !== undefined ||
		op.values !== undefined ||
		op.sortRef !== undefined ||
		op.sortBy !== undefined ||
		op.descending !== undefined
	)
}

function validateFilterCriteriaUpdate(
	opName: 'setAutoFilter' | 'setAdvancedFilter',
	op: FilterCriteriaUpdateOp,
) {
	for (const field of ['range', 'sortRef', 'sortBy'] as const) {
		const value = op[field]
		if (value === undefined) continue
		const rangeResult = safeParseRange(value)
		if (!rangeResult.ok) {
			return ascendError('VALIDATION_ERROR', `${opName} ${field} must be a valid A1 range`, {
				suggestedFix: `Use an A1 range such as A1:C20 for ${field}.`,
			})
		}
	}
	if (op.values !== undefined && !Array.isArray(op.values)) {
		return ascendError('VALIDATION_ERROR', `${opName} values must be an array`)
	}
	if (op.values?.some((value) => typeof value !== 'string')) {
		return ascendError('VALIDATION_ERROR', `${opName} values must be strings`)
	}
	if (op.values !== undefined && op.column === undefined) {
		return ascendError('VALIDATION_ERROR', `${opName} values require column`, {
			suggestedFix: 'Set column to the zero-based autoFilter colId that should receive values.',
		})
	}
	if (op.column !== undefined && op.values === undefined) {
		return ascendError('VALIDATION_ERROR', `${opName} column requires values`, {
			suggestedFix: 'Provide values for the selected zero-based autoFilter colId.',
		})
	}
	if (op.column !== undefined && (!Number.isInteger(op.column) || op.column < 0)) {
		return ascendError('VALIDATION_ERROR', `${opName} column must be non-negative`, {
			suggestedFix: 'Use the zero-based colId from inspectSheet().autoFilter.columns.',
		})
	}
	if (op.descending !== undefined && typeof op.descending !== 'boolean') {
		return ascendError('VALIDATION_ERROR', `${opName} descending must be boolean`, {
			suggestedFix: 'Set descending=true or descending=false.',
		})
	}
	return null
}

function applyAutoFilterUpdate(autoFilter: AutoFilter, op: FilterCriteriaUpdateOp): AutoFilter {
	const ref = op.range ?? autoFilter.ref
	const columns =
		op.column === undefined
			? autoFilter.columns.map(cloneFilterColumn)
			: upsertValueFilterColumn(autoFilter.columns, op.column, op.values ?? [])
	const sortState = applyAdvancedSortUpdate(autoFilter.sortState, op, ref)
	return {
		ref,
		columns,
		...(sortState ? { sortState } : {}),
	}
}

function applyAdvancedAutoFilterUpdate(
	autoFilter: AutoFilter,
	op: SetAdvancedFilterOp,
): AutoFilter {
	return applyAutoFilterUpdate(autoFilter, op)
}

function upsertValueFilterColumn(
	columns: readonly FilterColumn[],
	column: number,
	values: readonly string[],
): FilterColumn[] {
	const updatedColumn: FilterColumn = { colId: column, kind: 'filters', values: [...values] }
	const next = columns.map((entry) =>
		entry.colId === column
			? {
					...updatedColumn,
					...(entry.hiddenButton !== undefined ? { hiddenButton: entry.hiddenButton } : {}),
					...(entry.showButton !== undefined ? { showButton: entry.showButton } : {}),
				}
			: cloneFilterColumn(entry),
	)
	if (!next.some((entry) => entry.colId === column)) next.push(updatedColumn)
	return next.sort((a, b) => a.colId - b.colId)
}

function applyAdvancedSortUpdate(
	sortState: SortState | undefined,
	op: FilterCriteriaUpdateOp,
	filterRef: string,
): SortState | undefined {
	if (op.sortRef === undefined && op.sortBy === undefined && op.descending === undefined) {
		return sortState ? cloneSortState(sortState) : undefined
	}
	const ref = op.sortRef ?? sortState?.ref ?? filterRef
	const existingConditions = sortState?.conditions ?? []
	const first = existingConditions[0]
	const firstRef = op.sortBy ?? first?.ref ?? ref
	const firstCondition = {
		...(first ?? { ref: firstRef }),
		ref: firstRef,
		...(op.descending !== undefined ? { descending: op.descending } : {}),
	}
	return {
		...(sortState ? cloneSortState(sortState) : {}),
		ref,
		conditions:
			existingConditions.length === 0
				? [firstCondition]
				: [firstCondition, ...existingConditions.slice(1).map((condition) => ({ ...condition }))],
	}
}

function cloneFilterColumn(column: FilterColumn): FilterColumn {
	return {
		...column,
		...(column.values ? { values: [...column.values] } : {}),
		...(column.dateGroupItems
			? { dateGroupItems: column.dateGroupItems.map((item) => ({ ...item })) }
			: {}),
		...(column.customFilters
			? { customFilters: column.customFilters.map((filter) => ({ ...filter })) }
			: {}),
	}
}

function cloneSortState(sortState: SortState): SortState {
	return {
		...sortState,
		conditions: sortState.conditions.map((condition) => ({ ...condition })),
	}
}

export function handleSetPrintArea(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPrintArea' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const rangeResult = safeParseRange(op.range)
	if (!rangeResult.ok) return rangeResult
	workbook.definedNames.set(
		'_xlnm.Print_Area',
		`${quoteSheetNameForFormula(op.sheet)}!${op.range}`,
		{
			kind: 'sheet',
			sheetId: sheetResult.value.id,
		},
	)
	return ok(patch([], [op.sheet]))
}

function quoteSheetNameForFormula(sheet: string): string {
	return `'${sheet.replace(/'/g, "''")}'`
}

export function handleSetComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setComment' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result

	const ref = op.ref.toUpperCase()
	const existing = findCommentEntry(result.value.comments, ref)
	const comment =
		op.author !== undefined
			? { ...existing?.[1], text: op.text, author: op.author }
			: { ...existing?.[1], text: op.text }
	if (existing && existing[0] === ref && commentsEqual(existing[1], comment)) {
		return ok(patch([], []))
	}
	if (existing && existing[0] !== ref) result.value.comments.delete(existing[0])
	result.value.comments.set(ref, comment)

	return ok(patch([ref], [op.sheet]))
}

export function handleSetThreadedComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setThreadedComment' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	const sheet = result.value
	sheet.ensureWritable()

	if (
		op.partPath === undefined &&
		op.threadedCommentId === undefined &&
		op.ref === undefined &&
		op.commentIndex === undefined
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'setThreadedComment requires partPath, threadedCommentId, ref, or commentIndex',
				{
					suggestedFix:
						'Inspect sheet threadedComments and provide a stable threadedCommentId or commentIndex.',
				},
			),
		)
	}
	if (
		op.commentIndex !== undefined &&
		(op.commentIndex < 0 || !Number.isInteger(op.commentIndex))
	) {
		return err(ascendError('VALIDATION_ERROR', 'commentIndex must be a non-negative integer'))
	}

	const matches = sheet.threadedComments
		.map((comment, index) => ({ comment, index }))
		.filter(({ comment, index }) => {
			if (op.commentIndex !== undefined && index !== op.commentIndex) return false
			if (op.partPath !== undefined && comment.partPath !== op.partPath) return false
			if (op.threadedCommentId !== undefined && comment.id !== op.threadedCommentId) return false
			if (op.ref !== undefined && comment.ref.toUpperCase() !== op.ref.toUpperCase()) return false
			return true
		})

	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching threaded comment found', {
				suggestedFix:
					'Inspect sheet threadedComments and provide a matching threadedCommentId, partPath, ref, or commentIndex.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setThreadedComment matched ${matches.length} comments`, {
				suggestedFix:
					'Provide a more specific selector, such as threadedCommentId or commentIndex.',
			}),
		)
	}

	const match = matches[0]
	if (!match) return err(ascendError('VALIDATION_ERROR', 'No matching threaded comment found'))
	sheet.threadedComments[match.index] = { ...match.comment, text: op.text }

	return ok(patch([`${sheet.name}!${match.comment.ref}`], [sheet.name], false))
}

export function handleDeleteComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteComment' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const ref = op.ref.toUpperCase()
	for (const [commentRef] of sheet.comments) {
		if (commentRef.toUpperCase() === ref) sheet.comments.delete(commentRef)
	}
	sheet.threadedComments = sheet.threadedComments.filter(
		(comment) => comment.ref.toUpperCase() !== ref,
	)
	return ok(patch([`${op.sheet}!${op.ref}`], [op.sheet]))
}

function findCommentEntry(
	comments: Workbook['sheets'][number]['comments'],
	ref: string,
):
	| [string, Workbook['sheets'][number]['comments'] extends Map<string, infer T> ? T : never]
	| null {
	for (const entry of comments) {
		if (entry[0].toUpperCase() === ref) return entry
	}
	return null
}

function commentsEqual(
	left: Workbook['sheets'][number]['comments'] extends Map<string, infer T> ? T : never,
	right: Workbook['sheets'][number]['comments'] extends Map<string, infer T> ? T : never,
): boolean {
	return JSON.stringify(left) === JSON.stringify(right)
}

export function handleSetHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setHyperlink' }>,
): Result<PatchResult> {
	const result = getSheet(workbook, op.sheet)
	if (!result.ok) return result
	if (!hasLinkDestination(op.url) && !hasLinkDestination(op.location)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setHyperlink requires url or location', {
				suggestedFix:
					'Provide url for an external hyperlink or location for an internal workbook reference such as Sheet2!A1.',
			}),
		)
	}
	const ref = op.ref.toUpperCase()
	const existing = findHyperlinkEntry(result.value.hyperlinks, ref)
	const next = {
		...(hasLinkDestination(op.url) ? { target: op.url } : {}),
		...(hasLinkDestination(op.location) ? { location: op.location } : {}),
		...(op.display !== undefined ? { display: op.display } : {}),
		...(op.tooltip !== undefined ? { tooltip: op.tooltip } : {}),
	}
	if (existing && existing[0] === ref && hyperlinksEqual(existing[1], next)) {
		return ok(patch([], []))
	}
	result.value.ensureWritable()
	if (existing && existing[0] !== ref) result.value.hyperlinks.delete(existing[0])
	result.value.hyperlinks.set(ref, next)
	return ok(patch([ref], [op.sheet]))
}

function findHyperlinkEntry(
	hyperlinks: Workbook['sheets'][number]['hyperlinks'],
	ref: string,
):
	| [string, Workbook['sheets'][number]['hyperlinks'] extends Map<string, infer T> ? T : never]
	| null {
	for (const entry of hyperlinks) {
		if (entry[0].toUpperCase() === ref) return entry
	}
	return null
}

function hyperlinksEqual(
	left: Workbook['sheets'][number]['hyperlinks'] extends Map<string, infer T> ? T : never,
	right: Workbook['sheets'][number]['hyperlinks'] extends Map<string, infer T> ? T : never,
): boolean {
	return (
		left.target === right.target &&
		left.location === right.location &&
		left.display === right.display &&
		left.tooltip === right.tooltip
	)
}

function hasLinkDestination(value: string | undefined): boolean {
	return typeof value === 'string' && value.trim().length > 0
}

export function handleDeleteHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteHyperlink' }>,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
	const ref = op.ref.toUpperCase()
	for (const [linkRef] of sheet.hyperlinks) {
		if (linkRef.toUpperCase() === ref) sheet.hyperlinks.delete(linkRef)
	}
	return ok(patch([`${op.sheet}!${op.ref}`], [op.sheet]))
}

export function handleSetDefinedName(
	_workbook: Workbook,
	op: Extract<Operation, { op: 'setDefinedName' }>,
): Result<PatchResult> {
	const nameError = definedNameError(op.name)
	if (nameError) return err(nameError)
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

function definedNameError(name: string): ReturnType<typeof ascendError> | null {
	const validation = validateExcelDefinedName(name)
	if (!validation) return null
	return ascendError('VALIDATION_ERROR', validation.message, {
		suggestedFix: validation.suggestedFix,
	})
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
	if (!Number.isInteger(op.from) || !Number.isInteger(op.to) || op.from > op.to || op.from < 0) {
		return err(ascendError('VALIDATION_ERROR', 'Invalid row group range'))
	}
	const booleanValidation = validateGroupBooleanOptions('groupRows', op)
	if (booleanValidation) return err(booleanValidation)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
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
	if (!Number.isInteger(op.from) || !Number.isInteger(op.to) || op.from > op.to || op.from < 0) {
		return err(ascendError('VALIDATION_ERROR', 'Invalid column group range'))
	}
	const booleanValidation = validateGroupBooleanOptions('groupCols', op)
	if (booleanValidation) return err(booleanValidation)
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	sheet.ensureWritable()
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

function validateGroupBooleanOptions(
	opName: 'groupRows' | 'groupCols',
	op: Extract<Operation, { op: 'groupRows' | 'groupCols' }>,
) {
	const values = op as Record<string, unknown>
	for (const field of ['collapsed', 'summaryBelow', 'summaryRight'] as const) {
		const value = values[field]
		if (value !== undefined && typeof value !== 'boolean') {
			return ascendError('VALIDATION_ERROR', `${opName} ${field} must be boolean`, {
				suggestedFix: `Set ${field}=true or ${field}=false.`,
			})
		}
	}
	return null
}
