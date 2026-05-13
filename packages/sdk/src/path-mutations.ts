import { parseA1Safe, parseRange, type Workbook } from '@ascend/core'
import type {
	ConditionalFormatRule,
	DataValidationRule,
	InputValue,
	Operation,
	StyleInput,
} from '@ascend/schema'
import { validateExcelTableName, validateExcelWorksheetName } from '@ascend/schema'
import type {
	PathMutation,
	PathMutationIssue,
	PathMutationPath,
	PathMutationResult,
} from './types.ts'

type PathSegments = readonly string[]
type ParsedPath =
	| { readonly ok: true; readonly segments: PathSegments }
	| { readonly ok: false; readonly message: string }
type SetAutoFilterPathValue = Extract<Operation, { op: 'setAutoFilter' }>
type SetTableColumnPathValue = Extract<Operation, { op: 'setTableColumn' }>
type SetTableStylePathValue = Extract<Operation, { op: 'setTableStyle' }>

interface CompiledPathMutation {
	readonly op: Operation
}

export const SUPPORTED_PATH_MUTATION_ROOTS = ['sheets', 'tables', 'names'] as const

export const SUPPORTED_PATH_MUTATION_SHAPES = [
	'/sheets/{sheet}/cells/{A1}/value',
	'/sheets/{sheet}/cells/{A1}/formula',
	'/sheets/{sheet}/cells/{A1}/comment',
	'/sheets/{sheet}/cells/{A1}/hyperlink',
	'/sheets/{sheet}/ranges/{A1:B2}/clear',
	'/sheets/{sheet}/ranges/{A1:B2}/style',
	'/sheets/{sheet}/ranges/{A1:B2}/numberFormat',
	'/sheets/{sheet}/ranges/{A1:B2}/validation',
	'/sheets/{sheet}/ranges/{A1:B2}/conditionalFormat',
	'/sheets/{sheet}/ranges/{A1:B2}/merge',
	'/sheets/{sheet}/autofilter',
	'/sheets/{sheet}/name',
	'/tables/{table}/name',
	'/tables/{table}/rows/append',
	'/tables/{table}/columns/{nameOrIndex}/name',
	'/tables/{table}/columns/{nameOrIndex}/formula',
	'/tables/{table}/columns/{nameOrIndex}/totalsRowFunction',
	'/tables/{table}/columns/{nameOrIndex}/totalsRowFormula',
	'/tables/{table}/columns/{nameOrIndex}/totalsRowLabel',
	'/tables/{table}/style',
	'/names/{name}/ref',
	'/sheets/{sheet}/names/{name}/ref',
] as const

export function compilePathMutations(
	workbook: Workbook,
	mutations: readonly PathMutation[],
): PathMutationResult {
	const ops: Operation[] = []
	const issues: PathMutationIssue[] = []

	for (const mutation of mutations) {
		const parsedPath = parseMutationPath(mutation.path)
		if (!parsedPath.ok) {
			issues.push(issue(mutation.path, 'invalid_path', parsedPath.message))
			continue
		}
		const compiled = compilePathMutation(workbook, mutation, parsedPath.segments)
		if ('op' in compiled) {
			pushOperation(ops, compiled.op)
		} else {
			issues.push(compiled)
		}
	}

	const replayable = issues.length === 0
	return {
		ops: replayable ? deferRenameOperations(ops) : [],
		mutationCount: mutations.length,
		issueCount: issues.length,
		issues,
		replayable,
	}
}

function compilePathMutation(
	workbook: Workbook,
	mutation: PathMutation,
	segments: PathSegments,
): CompiledPathMutation | PathMutationIssue {
	if (segments[0] === 'sheets') return compileSheetPath(workbook, mutation, segments)
	if (segments[0] === 'tables') return compileTablePath(workbook, mutation, segments)
	if (segments[0] === 'names') return compileWorkbookNamePath(mutation, segments)
	return issue(mutation.path, 'unsupported_path', `Unsupported path root "${segments[0]}".`, {
		supportedRoots: SUPPORTED_PATH_MUTATION_ROOTS,
		supportedShapes: SUPPORTED_PATH_MUTATION_SHAPES,
	})
}

function compileSheetPath(
	workbook: Workbook,
	mutation: PathMutation,
	segments: PathSegments,
): CompiledPathMutation | PathMutationIssue {
	const sheet = segments[1]
	if (!sheet) return issue(mutation.path, 'invalid_path', 'Sheet path must include a sheet name.')
	const sheetModel = workbook.getSheet(sheet)
	if (!sheetModel) {
		return issue(mutation.path, 'sheet_not_found', `Sheet "${sheet}" not found.`, {
			availableSheets: workbook.sheets.map((entry) => entry.name),
		})
	}
	if (segments.length === 3 && segments[2] === 'name') {
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(mutation.path, 'invalid_value', 'Sheet rename value must be a non-empty string.')
		}
		const validation = validateExcelWorksheetName(mutation.value)
		if (validation) {
			return issue(mutation.path, 'invalid_value', validation.message, {
				suggestedFix: validation.suggestedFix,
			})
		}
		return { op: { op: 'renameSheet', sheet, newName: mutation.value } }
	}
	if (segments[2] === 'cells' && segments.length === 5) {
		return compileCellPath(sheet, mutation, segments[3], segments[4])
	}
	if (segments[2] === 'ranges' && segments.length === 5) {
		const range = segments[3]
		if (!range) return issue(mutation.path, 'invalid_ref', 'Range path must include an A1 range.')
		if (!isValidRange(range)) {
			return issue(mutation.path, 'invalid_ref', `Invalid range reference "${range}".`)
		}
		return compileRangePath(sheet, range, mutation, segments[4])
	}
	if (segments[2] === 'autofilter' && segments.length === 3) {
		return compileAutoFilterPath(sheet, mutation)
	}
	if (segments[2] === 'names' && segments.length === 5 && segments[4] === 'ref') {
		const name = segments[3]
		if (!name)
			return issue(mutation.path, 'invalid_path', 'Sheet-scoped name path must include a name.')
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(
				mutation.path,
				'invalid_value',
				'Defined-name ref value must be a non-empty string.',
			)
		}
		return { op: { op: 'setDefinedName', name, scope: sheet, ref: mutation.value } }
	}
	return unsupportedShape(mutation.path, segments)
}

function compileCellPath(
	sheet: string,
	mutation: PathMutation,
	ref: string | undefined,
	field: string | undefined,
): CompiledPathMutation | PathMutationIssue {
	if (!ref || !parseA1Safe(ref)) {
		return issue(mutation.path, 'invalid_ref', `Invalid cell reference "${ref ?? ''}".`)
	}
	if (field === 'value') {
		const value = inputValue(mutation.value)
		if (!value.ok) return issue(mutation.path, 'invalid_value', value.message)
		return { op: { op: 'setCells', sheet, updates: [{ ref, value: value.value }] } }
	}
	if (field === 'formula') {
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(mutation.path, 'invalid_value', 'Formula value must be a non-empty string.')
		}
		return { op: { op: 'setFormula', sheet, ref, formula: mutation.value } }
	}
	if (field === 'comment') {
		return compileCommentPath(sheet, ref, mutation)
	}
	if (field === 'hyperlink') {
		return compileHyperlinkPath(sheet, ref, mutation)
	}
	return issue(mutation.path, 'unsupported_path', `Unsupported cell field "${field ?? ''}".`, {
		supportedFields: ['value', 'formula', 'comment', 'hyperlink'],
	})
}

function compileCommentPath(
	sheet: string,
	ref: string,
	mutation: PathMutation,
): CompiledPathMutation | PathMutationIssue {
	if (mutation.value === null) return { op: { op: 'deleteComment', sheet, ref } }
	if (typeof mutation.value === 'string') {
		return { op: { op: 'setComment', sheet, ref, text: mutation.value } }
	}
	if (!isRecord(mutation.value)) {
		return issue(
			mutation.path,
			'invalid_value',
			'Comment value must be a string, null, or { text, author? } object.',
		)
	}
	const text = mutation.value.text
	const author = mutation.value.author
	if (typeof text !== 'string') {
		return issue(mutation.path, 'invalid_value', 'Comment object must include string text.')
	}
	if (author !== undefined && typeof author !== 'string') {
		return issue(mutation.path, 'invalid_value', 'Comment author must be a string when provided.')
	}
	return { op: { op: 'setComment', sheet, ref, text, ...(author ? { author } : {}) } }
}

function compileHyperlinkPath(
	sheet: string,
	ref: string,
	mutation: PathMutation,
): CompiledPathMutation | PathMutationIssue {
	if (mutation.value === null) return { op: { op: 'deleteHyperlink', sheet, ref } }
	if (typeof mutation.value === 'string') {
		return { op: { op: 'setHyperlink', sheet, ref, url: mutation.value } }
	}
	if (!isRecord(mutation.value)) {
		return issue(
			mutation.path,
			'invalid_value',
			'Hyperlink value must be a string, null, or { url?, location?, display?, tooltip? } object.',
		)
	}
	const { url, location, display, tooltip } = mutation.value
	if (url !== undefined && typeof url !== 'string') {
		return issue(mutation.path, 'invalid_value', 'Hyperlink url must be a string when provided.')
	}
	if (location !== undefined && typeof location !== 'string') {
		return issue(
			mutation.path,
			'invalid_value',
			'Hyperlink location must be a string when provided.',
		)
	}
	if (!url && !location) {
		return issue(mutation.path, 'invalid_value', 'Hyperlink must include url or location.')
	}
	if (display !== undefined && typeof display !== 'string') {
		return issue(
			mutation.path,
			'invalid_value',
			'Hyperlink display must be a string when provided.',
		)
	}
	if (tooltip !== undefined && typeof tooltip !== 'string') {
		return issue(
			mutation.path,
			'invalid_value',
			'Hyperlink tooltip must be a string when provided.',
		)
	}
	return {
		op: {
			op: 'setHyperlink',
			sheet,
			ref,
			...(url ? { url } : {}),
			...(location ? { location } : {}),
			...(display ? { display } : {}),
			...(tooltip ? { tooltip } : {}),
		},
	}
}

function compileRangePath(
	sheet: string,
	range: string,
	mutation: PathMutation,
	field: string | undefined,
): CompiledPathMutation | PathMutationIssue {
	if (field === 'clear') {
		const what = mutation.value ?? 'all'
		if (what !== 'values' && what !== 'formulas' && what !== 'styles' && what !== 'all') {
			return issue(
				mutation.path,
				'invalid_value',
				'Range clear value must be values, formulas, styles, or all.',
			)
		}
		return { op: { op: 'clearRange', sheet, range, what } }
	}
	if (field === 'style') {
		if (!isRecord(mutation.value)) {
			return issue(mutation.path, 'invalid_value', 'Range style value must be a style object.')
		}
		return { op: { op: 'setStyle', sheet, range, style: mutation.value as StyleInput } }
	}
	if (field === 'numberFormat') {
		if (typeof mutation.value !== 'string') {
			return issue(mutation.path, 'invalid_value', 'Number format value must be a string.')
		}
		return { op: { op: 'setNumberFormat', sheet, range, format: mutation.value } }
	}
	if (field === 'validation') {
		if (mutation.value === null) return { op: { op: 'deleteDataValidation', sheet, range } }
		if (!isRecord(mutation.value)) {
			return issue(
				mutation.path,
				'invalid_value',
				'Validation value must be null or a data validation rule object.',
			)
		}
		return {
			op: {
				op: 'setDataValidation',
				sheet,
				range,
				rule: mutation.value as unknown as DataValidationRule,
			},
		}
	}
	if (field === 'conditionalFormat') {
		if (mutation.value === null) return { op: { op: 'deleteConditionalFormat', sheet, range } }
		if (!isRecord(mutation.value)) {
			return issue(
				mutation.path,
				'invalid_value',
				'Conditional format value must be null, a rule object, or { rule, mode?, reassignPriorities? }.',
			)
		}
		const maybeRule = mutation.value.rule
		const rule = (isRecord(maybeRule)
			? maybeRule
			: mutation.value) as unknown as ConditionalFormatRule
		const mode = mutation.value.mode
		const reassignPriorities = mutation.value.reassignPriorities
		if (mode !== undefined && mode !== 'replace' && mode !== 'append') {
			return issue(
				mutation.path,
				'invalid_value',
				'Conditional format mode must be replace or append.',
			)
		}
		if (reassignPriorities !== undefined && typeof reassignPriorities !== 'boolean') {
			return issue(
				mutation.path,
				'invalid_value',
				'Conditional format reassignPriorities must be boolean when provided.',
			)
		}
		return {
			op: {
				op: 'setConditionalFormat',
				sheet,
				range,
				rule,
				...(mode ? { mode } : {}),
				...(reassignPriorities !== undefined ? { reassignPriorities } : {}),
			},
		}
	}
	if (field === 'merge') {
		if (mutation.value === true) return { op: { op: 'mergeCells', sheet, range } }
		if (mutation.value === false || mutation.value === null) {
			return { op: { op: 'unmergeCells', sheet, range } }
		}
		return issue(mutation.path, 'invalid_value', 'Range merge value must be true, false, or null.')
	}
	return issue(mutation.path, 'unsupported_path', `Unsupported range field "${field ?? ''}".`, {
		supportedFields: ['clear', 'style', 'numberFormat', 'validation', 'conditionalFormat', 'merge'],
	})
}

function compileAutoFilterPath(
	sheet: string,
	mutation: PathMutation,
): CompiledPathMutation | PathMutationIssue {
	if (mutation.value === null) return { op: { op: 'clearAutoFilter', sheet } }
	if (typeof mutation.value === 'string') {
		if (!isValidRange(mutation.value)) {
			return issue(mutation.path, 'invalid_ref', `Invalid autofilter range "${mutation.value}".`)
		}
		return { op: { op: 'setAutoFilter', sheet, range: mutation.value } }
	}
	if (!isRecord(mutation.value)) {
		return issue(
			mutation.path,
			'invalid_value',
			'Autofilter value must be null, an A1 range string, or an object with range.',
		)
	}
	const range = mutation.value.range
	if (typeof range !== 'string' || !isValidRange(range)) {
		return issue(mutation.path, 'invalid_ref', 'Autofilter object must include a valid range.')
	}
	const op: SetAutoFilterPathValue = { op: 'setAutoFilter', sheet, range }
	return withOptionalAutoFilterFields(op, mutation)
}

function withOptionalAutoFilterFields(
	op: SetAutoFilterPathValue,
	mutation: PathMutation,
): CompiledPathMutation | PathMutationIssue {
	if (!isRecord(mutation.value)) return { op }
	const { column, values, sortRef, sortBy, descending } = mutation.value
	if (
		column !== undefined &&
		(typeof column !== 'number' || !Number.isInteger(column) || column < 0)
	) {
		return issue(
			mutation.path,
			'invalid_value',
			'Autofilter column must be a non-negative integer.',
		)
	}
	if (
		values !== undefined &&
		(!Array.isArray(values) || values.some((value) => typeof value !== 'string'))
	) {
		return issue(mutation.path, 'invalid_value', 'Autofilter values must be an array of strings.')
	}
	for (const [name, value] of [
		['sortRef', sortRef],
		['sortBy', sortBy],
	] as const) {
		if (value !== undefined && typeof value !== 'string') {
			return issue(mutation.path, 'invalid_value', `Autofilter ${name} must be a string.`)
		}
	}
	if (descending !== undefined && typeof descending !== 'boolean') {
		return issue(mutation.path, 'invalid_value', 'Autofilter descending must be boolean.')
	}
	const filterValues = values as readonly string[] | undefined
	const sortRefValue = sortRef as string | undefined
	const sortByValue = sortBy as string | undefined
	return {
		op: {
			...op,
			...(column !== undefined ? { column } : {}),
			...(filterValues !== undefined ? { values: filterValues } : {}),
			...(sortRefValue ? { sortRef: sortRefValue } : {}),
			...(sortByValue ? { sortBy: sortByValue } : {}),
			...(descending !== undefined ? { descending } : {}),
		},
	}
}

function compileTablePath(
	workbook: Workbook,
	mutation: PathMutation,
	segments: PathSegments,
): CompiledPathMutation | PathMutationIssue {
	const table = segments[1]
	if (!table) return issue(mutation.path, 'invalid_path', 'Table path must include a table name.')
	const tableModel = workbook.sheets
		.flatMap((sheet) => sheet.tables)
		.find((entry) => entry.name === table)
	if (!tableModel) {
		return issue(mutation.path, 'table_not_found', `Table "${table}" not found.`, {
			availableTables: workbook.sheets.flatMap((sheet) => sheet.tables.map((entry) => entry.name)),
		})
	}
	if (segments.length === 3 && segments[2] === 'name') {
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(mutation.path, 'invalid_value', 'Table rename value must be a non-empty string.')
		}
		const validation = validateExcelTableName(mutation.value)
		if (validation) {
			return issue(mutation.path, 'invalid_value', validation.message, {
				suggestedFix: validation.suggestedFix,
			})
		}
		return { op: { op: 'renameTable', table, newName: mutation.value } }
	}
	if (segments.length === 4 && segments[2] === 'rows' && segments[3] === 'append') {
		const rows = inputRows(mutation.value)
		if (!rows.ok) return issue(mutation.path, 'invalid_value', rows.message)
		return { op: { op: 'appendRows', table, rows: rows.rows } }
	}
	if (segments.length === 5 && segments[2] === 'columns') {
		return compileTableColumnPath(
			table,
			tableModel.columns.map((column) => column.name),
			mutation,
			segments[3],
			segments[4],
		)
	}
	if (segments.length === 3 && segments[2] === 'style') {
		return compileTableStylePath(table, mutation)
	}
	return unsupportedShape(mutation.path, segments)
}

function compileTableColumnPath(
	table: string,
	columns: readonly string[],
	mutation: PathMutation,
	columnSegment: string | undefined,
	field: string | undefined,
): CompiledPathMutation | PathMutationIssue {
	const column = tableColumnSelector(columnSegment)
	if (!column.ok) return issue(mutation.path, 'invalid_path', column.message)
	const exists =
		typeof column.selector === 'number'
			? column.selector >= 0 && column.selector < columns.length
			: columns.includes(column.selector)
	if (!exists) {
		return issue(
			mutation.path,
			'invalid_path',
			`Table column "${columnSegment ?? ''}" not found.`,
			{
				availableColumns: columns,
			},
		)
	}
	if (field === 'name') {
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(
				mutation.path,
				'invalid_value',
				'Table column name value must be a non-empty string.',
			)
		}
		return { op: { op: 'setTableColumn', table, column: column.selector, newName: mutation.value } }
	}
	if (field === 'formula')
		return compileTableColumnNullableString(table, column.selector, mutation, 'formula')
	if (field === 'totalsRowFunction') {
		return compileTableColumnNullableString(table, column.selector, mutation, 'totalsRowFunction')
	}
	if (field === 'totalsRowFormula') {
		return compileTableColumnNullableString(table, column.selector, mutation, 'totalsRowFormula')
	}
	if (field === 'totalsRowLabel') {
		return compileTableColumnNullableString(table, column.selector, mutation, 'totalsRowLabel')
	}
	return issue(
		mutation.path,
		'unsupported_path',
		`Unsupported table column field "${field ?? ''}".`,
		{
			supportedFields: [
				'name',
				'formula',
				'totalsRowFunction',
				'totalsRowFormula',
				'totalsRowLabel',
			],
		},
	)
}

function tableColumnSelector(
	segment: string | undefined,
): { ok: true; selector: string | number } | { ok: false; message: string } {
	if (!segment)
		return { ok: false, message: 'Table column path must include a column name or index.' }
	const asNumber = Number(segment)
	if (Number.isInteger(asNumber) && String(asNumber) === segment)
		return { ok: true, selector: asNumber }
	return { ok: true, selector: segment }
}

function compileTableColumnNullableString(
	table: string,
	column: string | number,
	mutation: PathMutation,
	field: keyof Pick<
		SetTableColumnPathValue,
		'formula' | 'totalsRowFunction' | 'totalsRowFormula' | 'totalsRowLabel'
	>,
): CompiledPathMutation | PathMutationIssue {
	if (mutation.value !== null && typeof mutation.value !== 'string') {
		return issue(
			mutation.path,
			'invalid_value',
			`Table column ${field} value must be a string or null.`,
		)
	}
	return { op: { op: 'setTableColumn', table, column, [field]: mutation.value } }
}

function compileTableStylePath(
	table: string,
	mutation: PathMutation,
): CompiledPathMutation | PathMutationIssue {
	if (mutation.value === null || typeof mutation.value === 'string') {
		return { op: { op: 'setTableStyle', table, styleName: mutation.value } }
	}
	if (!isRecord(mutation.value)) {
		return issue(
			mutation.path,
			'invalid_value',
			'Table style value must be a string, null, or object.',
		)
	}
	const { styleName, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes } =
		mutation.value
	if (styleName !== undefined && styleName !== null && typeof styleName !== 'string') {
		return issue(mutation.path, 'invalid_value', 'Table styleName must be a string or null.')
	}
	for (const [name, value] of [
		['showFirstColumn', showFirstColumn],
		['showLastColumn', showLastColumn],
		['showRowStripes', showRowStripes],
		['showColumnStripes', showColumnStripes],
	] as const) {
		if (value !== undefined && typeof value !== 'boolean') {
			return issue(mutation.path, 'invalid_value', `Table style ${name} must be boolean.`)
		}
	}
	const firstColumn = showFirstColumn as boolean | undefined
	const lastColumn = showLastColumn as boolean | undefined
	const rowStripes = showRowStripes as boolean | undefined
	const columnStripes = showColumnStripes as boolean | undefined
	const op: SetTableStylePathValue = {
		op: 'setTableStyle',
		table,
		...(styleName !== undefined ? { styleName } : {}),
		...(firstColumn !== undefined ? { showFirstColumn: firstColumn } : {}),
		...(lastColumn !== undefined ? { showLastColumn: lastColumn } : {}),
		...(rowStripes !== undefined ? { showRowStripes: rowStripes } : {}),
		...(columnStripes !== undefined ? { showColumnStripes: columnStripes } : {}),
	}
	return { op }
}

function compileWorkbookNamePath(
	mutation: PathMutation,
	segments: PathSegments,
): CompiledPathMutation | PathMutationIssue {
	const name = segments[1]
	if (!name) return issue(mutation.path, 'invalid_path', 'Name path must include a defined name.')
	if (segments.length === 3 && segments[2] === 'ref') {
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(
				mutation.path,
				'invalid_value',
				'Defined-name ref value must be a non-empty string.',
			)
		}
		return { op: { op: 'setDefinedName', name, ref: mutation.value } }
	}
	return unsupportedShape(mutation.path, segments)
}

function inputValue(
	value: unknown,
): { ok: true; value: InputValue } | { ok: false; message: string } {
	if (
		value === null ||
		value instanceof Date ||
		typeof value === 'string' ||
		typeof value === 'number' ||
		typeof value === 'boolean'
	) {
		return { ok: true, value }
	}
	return { ok: false, message: 'Cell value must be a scalar value, Date, or null.' }
}

function inputRows(
	value: unknown,
): { ok: true; rows: readonly (readonly InputValue[])[] } | { ok: false; message: string } {
	if (!Array.isArray(value))
		return { ok: false, message: 'Append rows value must be an array of rows.' }
	const rows: InputValue[][] = []
	for (const [rowIndex, row] of value.entries()) {
		if (!Array.isArray(row)) {
			return { ok: false, message: `Append row ${rowIndex} must be an array.` }
		}
		const values: InputValue[] = []
		for (const [colIndex, cell] of row.entries()) {
			const input = inputValue(cell)
			if (!input.ok) {
				return {
					ok: false,
					message: `Append row ${rowIndex}, column ${colIndex} ${input.message}`,
				}
			}
			values.push(input.value)
		}
		rows.push(values)
	}
	return { ok: true, rows }
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function pushOperation(ops: Operation[], op: Operation): void {
	const last = ops[ops.length - 1]
	if (last?.op === 'setCells' && op.op === 'setCells' && last.sheet === op.sheet) {
		ops[ops.length - 1] = {
			op: 'setCells',
			sheet: last.sheet,
			updates: [...last.updates, ...op.updates],
		}
		return
	}
	ops.push(op)
}

function deferRenameOperations(ops: readonly Operation[]): Operation[] {
	const immediate: Operation[] = []
	const deferred: Operation[] = []
	for (const op of ops) {
		if (op.op === 'renameSheet' || op.op === 'renameTable') {
			deferred.push(op)
			continue
		}
		pushOperation(immediate, op)
	}
	return [...immediate, ...deferred]
}

function parseMutationPath(path: PathMutationPath): ParsedPath {
	if (typeof path !== 'string') {
		if (path.length === 0) return { ok: false, message: 'Path must not be empty.' }
		const emptyIndex = path.findIndex((segment) => segment.length === 0)
		if (emptyIndex >= 0) {
			return { ok: false, message: `Path segment ${emptyIndex} must not be empty.` }
		}
		return { ok: true, segments: [...path] }
	}
	if (path.startsWith('/')) {
		return parsePointerPath(path)
	}
	return splitDotPath(path)
}

function splitDotPath(path: string): ParsedPath {
	if (path.length === 0) return { ok: false, message: 'Path must not be empty.' }
	const segments: string[] = []
	let segment = ''
	let escaped = false
	for (let i = 0; i < path.length; i++) {
		const char = path[i] ?? ''
		if (escaped) {
			if (char !== '.' && char !== '\\') {
				return { ok: false, message: `Invalid escaped character "\\${char}" in dot path.` }
			}
			segment += char
			escaped = false
			continue
		}
		if (char === '\\') {
			escaped = true
			continue
		}
		if (char === '.') {
			if (segment.length === 0) {
				return { ok: false, message: `Path segment ${segments.length} must not be empty.` }
			}
			segments.push(segment)
			segment = ''
			continue
		}
		segment += char
	}
	if (escaped) return { ok: false, message: 'Dot path must not end with an escape character.' }
	if (segment.length === 0) {
		return { ok: false, message: `Path segment ${segments.length} must not be empty.` }
	}
	segments.push(segment)
	return { ok: true, segments }
}

function parsePointerPath(path: string): ParsedPath {
	const encodedSegments = path.slice(1).split('/')
	if (encodedSegments.length === 0 || encodedSegments[0] === '') {
		return { ok: false, message: 'Path must not be empty.' }
	}
	const segments: string[] = []
	for (const [index, encodedSegment] of encodedSegments.entries()) {
		if (encodedSegment.length === 0) {
			return { ok: false, message: `Path segment ${index} must not be empty.` }
		}
		const decoded = decodePointerSegment(encodedSegment)
		if (!decoded.ok) return decoded
		segments.push(decoded.segment)
	}
	return { ok: true, segments }
}

function decodePointerSegment(
	segment: string,
):
	| { readonly ok: true; readonly segment: string }
	| { readonly ok: false; readonly message: string } {
	let decoded: string
	try {
		decoded = decodeURIComponent(segment)
	} catch {
		return { ok: false, message: `Invalid percent encoding in path segment "${segment}".` }
	}
	if (/(^|[^~])~(?:[^01]|$)/.test(decoded) || decoded === '~') {
		return { ok: false, message: `Invalid JSON Pointer escape in path segment "${segment}".` }
	}
	return { ok: true, segment: decoded.replace(/~1/g, '/').replace(/~0/g, '~') }
}

function isValidRange(range: string): boolean {
	try {
		parseRange(range)
		return true
	} catch {
		return false
	}
}

function unsupportedShape(path: PathMutationPath, segments: PathSegments): PathMutationIssue {
	return issue(path, 'unsupported_path', `Unsupported path shape "${segments.join('/')}".`, {
		supportedShapes: SUPPORTED_PATH_MUTATION_SHAPES,
	})
}

function issue(
	path: PathMutationPath,
	code: PathMutationIssue['code'],
	message: string,
	details?: Readonly<Record<string, unknown>>,
): PathMutationIssue {
	return {
		path,
		code,
		message,
		...(details ? { details } : {}),
	}
}
