import { parseA1Safe, parseRange, type Workbook } from '@ascend/core'
import type { InputValue, Operation } from '@ascend/schema'
import type {
	PathMutation,
	PathMutationIssue,
	PathMutationPath,
	PathMutationResult,
} from './types.ts'

type PathSegments = readonly string[]

interface CompiledPathMutation {
	readonly op: Operation
}

export function compilePathMutations(
	workbook: Workbook,
	mutations: readonly PathMutation[],
): PathMutationResult {
	const ops: Operation[] = []
	const issues: PathMutationIssue[] = []

	for (const mutation of mutations) {
		const segments = parseMutationPath(mutation.path)
		if (!segments) {
			issues.push(issue(mutation.path, 'invalid_path', 'Path must not be empty.'))
			continue
		}
		const compiled = compilePathMutation(workbook, mutation, segments)
		if ('op' in compiled) {
			pushOperation(ops, compiled.op)
		} else {
			issues.push(compiled)
		}
	}

	return {
		ops,
		mutationCount: mutations.length,
		issueCount: issues.length,
		issues,
		replayable: issues.length === 0,
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
		supportedRoots: ['sheets', 'tables', 'names'],
	})
}

function compileSheetPath(
	workbook: Workbook,
	mutation: PathMutation,
	segments: PathSegments,
): CompiledPathMutation | PathMutationIssue {
	const sheet = segments[1]
	if (!sheet) return issue(mutation.path, 'invalid_path', 'Sheet path must include a sheet name.')
	if (!workbook.getSheet(sheet)) {
		return issue(mutation.path, 'sheet_not_found', `Sheet "${sheet}" not found.`, {
			availableSheets: workbook.sheets.map((entry) => entry.name),
		})
	}
	if (segments.length === 3 && segments[2] === 'name') {
		if (typeof mutation.value !== 'string' || mutation.value.length === 0) {
			return issue(mutation.path, 'invalid_value', 'Sheet rename value must be a non-empty string.')
		}
		return { op: { op: 'renameSheet', sheet, newName: mutation.value } }
	}
	if (segments[2] === 'cells' && segments.length === 5) {
		return compileCellPath(sheet, mutation, segments[3], segments[4])
	}
	if (segments[2] === 'ranges' && segments.length === 5 && segments[4] === 'clear') {
		const range = segments[3]
		if (!range) return issue(mutation.path, 'invalid_ref', 'Range path must include an A1 range.')
		if (!isValidRange(range)) {
			return issue(mutation.path, 'invalid_ref', `Invalid range reference "${range}".`)
		}
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
	return issue(mutation.path, 'unsupported_path', `Unsupported cell field "${field ?? ''}".`, {
		supportedFields: ['value', 'formula'],
	})
}

function compileTablePath(
	workbook: Workbook,
	mutation: PathMutation,
	segments: PathSegments,
): CompiledPathMutation | PathMutationIssue {
	const table = segments[1]
	if (!table) return issue(mutation.path, 'invalid_path', 'Table path must include a table name.')
	if (!workbook.sheets.some((sheet) => sheet.tables.some((entry) => entry.name === table))) {
		return issue(mutation.path, 'table_not_found', `Table "${table}" not found.`, {
			availableTables: workbook.sheets.flatMap((sheet) => sheet.tables.map((entry) => entry.name)),
		})
	}
	if (segments.length === 4 && segments[2] === 'rows' && segments[3] === 'append') {
		const rows = inputRows(mutation.value)
		if (!rows.ok) return issue(mutation.path, 'invalid_value', rows.message)
		return { op: { op: 'appendRows', table, rows: rows.rows } }
	}
	return unsupportedShape(mutation.path, segments)
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

function parseMutationPath(path: PathMutationPath): PathSegments | null {
	if (typeof path !== 'string') {
		const segments = path.map((segment) => String(segment)).filter((segment) => segment.length > 0)
		return segments.length > 0 ? segments : null
	}
	if (path.startsWith('/')) {
		const segments = path
			.slice(1)
			.split('/')
			.map(decodePointerSegment)
			.filter((segment) => segment.length > 0)
		return segments.length > 0 ? segments : null
	}
	const segments = splitDotPath(path)
	return segments.length > 0 ? segments : null
}

function splitDotPath(path: string): string[] {
	const segments: string[] = []
	let segment = ''
	let escaped = false
	for (const char of path) {
		if (escaped) {
			segment += char
			escaped = false
			continue
		}
		if (char === '\\') {
			escaped = true
			continue
		}
		if (char === '.') {
			if (segment.length > 0) segments.push(segment)
			segment = ''
			continue
		}
		segment += char
	}
	if (escaped) segment += '\\'
	if (segment.length > 0) segments.push(segment)
	return segments
}

function decodePointerSegment(segment: string): string {
	return decodeURIComponent(segment).replace(/~1/g, '/').replace(/~0/g, '~')
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
		supportedShapes: [
			'/sheets/{sheet}/cells/{A1}/value',
			'/sheets/{sheet}/cells/{A1}/formula',
			'/sheets/{sheet}/ranges/{A1:B2}/clear',
			'/sheets/{sheet}/name',
			'/tables/{table}/rows/append',
			'/names/{name}/ref',
			'/sheets/{sheet}/names/{name}/ref',
		],
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
