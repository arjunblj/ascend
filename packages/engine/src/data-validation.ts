import type { Sheet, SheetDataValidation, Workbook } from '@ascend/core'
import { DEFAULT_STYLE_ID, parseRange } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import { isEmpty, topLeftScalar } from '@ascend/schema'
import { evaluateRelativeFormulaText } from './relative-formula.ts'

export interface ValidationResult {
	readonly valid: boolean
	readonly rule?: SheetDataValidation
	readonly message?: string
}

const INVALID_NUMERIC_FORMULA = Symbol('invalid numeric validation formula')

type NumericFormulaValue = number | null | typeof INVALID_NUMERIC_FORMULA

type ListCandidate =
	| { readonly kind: 'literal'; readonly value: string }
	| { readonly kind: 'cell'; readonly value: CellValue }

function cellInSqref(sqref: string, row: number, col: number): boolean {
	const parts = sqref.split(/\s+/)
	for (const part of parts) {
		try {
			const range = parseRange(part)
			if (
				row >= range.start.row &&
				row <= range.end.row &&
				col >= range.start.col &&
				col <= range.end.col
			) {
				return true
			}
		} catch {
			// invalid range part, skip
		}
	}
	return false
}

function findRuleForCell(sheet: Sheet, row: number, col: number): SheetDataValidation | undefined {
	for (const dv of sheet.dataValidations) {
		if (cellInSqref(dv.sqref, row, col)) return dv
	}
	return undefined
}

function findAnchorForCell(
	sqref: string,
	row: number,
	col: number,
): { row: number; col: number } | null {
	const parts = sqref.split(/\s+/)
	for (const part of parts) {
		try {
			const range = parseRange(part)
			if (
				row >= range.start.row &&
				row <= range.end.row &&
				col >= range.start.col &&
				col <= range.end.col
			) {
				return { row: range.start.row, col: range.start.col }
			}
		} catch {
			// invalid range part, skip
		}
	}
	return null
}

function parseListValues(formula1: string): string[] {
	const s = formula1.trim()
	if (s.startsWith('"') && s.endsWith('"')) {
		const inner = s.slice(1, -1)
		return inner.split(',').map((v) => v.trim())
	}
	return s.split(',').map((v) => v.trim())
}

function parseNumericFormula(formula: string): number | null {
	const s = formula.trim()
	const n = Number.parseFloat(s)
	return Number.isFinite(n) ? n : null
}

function evaluateFormulaScalar(
	formula: string | undefined,
	workbook: Workbook | undefined,
	sheet: Sheet,
	row: number,
	col: number,
): CellValue | null {
	if (!formula) return null
	if (!workbook) return null
	const sheetIndex = workbook.sheets.indexOf(sheet)
	if (sheetIndex < 0) return null
	const anchor = findAnchorForCell(
		findRuleForCell(sheet, row, col)?.sqref ?? `${row + 1}:${col + 1}`,
		row,
		col,
	) ?? { row, col }
	return evaluateRelativeFormulaText(
		formula,
		workbook,
		sheetIndex,
		anchor.row,
		anchor.col,
		row,
		col,
	)
}

function formulaToNumber(
	formula: string | undefined,
	workbook: Workbook | undefined,
	sheet: Sheet,
	row: number,
	col: number,
): NumericFormulaValue {
	if (!formula) return null
	const direct = parseNumericFormula(formula)
	if (direct !== null) return direct
	const value = evaluateFormulaScalar(formula, workbook, sheet, row, col)
	if (!value) return null
	return scalarToNumberOrInvalid(value)
}

function formulaToListCandidates(
	formula: string | undefined,
	workbook: Workbook | undefined,
	sheet: Sheet,
	row: number,
	col: number,
): ListCandidate[] {
	if (!formula) return []
	const trimmed = formula.trim()
	if (trimmed.startsWith('"') || !workbook || (trimmed.includes(',') && !/[=:!$]/.test(trimmed))) {
		return parseListValues(formula).map((value) => ({ kind: 'literal', value }))
	}
	const resolvedList = resolveListFormulaValues(trimmed, workbook, sheet)
	if (resolvedList.length > 0) {
		return resolvedList.map((value) => ({ kind: 'cell', value }))
	}
	const value = evaluateFormulaScalar(formula, workbook, sheet, row, col)
	if (!value) return parseListValues(formula).map((entry) => ({ kind: 'literal', value: entry }))
	if (value.kind === 'array') {
		const out: ListCandidate[] = []
		for (const line of value.rows) {
			for (const cell of line) out.push({ kind: 'cell', value: cell })
		}
		return out
	}
	return [{ kind: 'cell', value }]
}

function resolveListFormulaValues(
	formula: string,
	workbook: Workbook,
	sheet: Sheet,
	seenNames = new Set<string>(),
): CellValue[] {
	const normalized = formula.startsWith('=') ? formula.slice(1).trim() : formula
	const definedName = /^[A-Za-z_\\][\w.\\]*$/.test(normalized)
		? workbook.definedNames.resolve(normalized, sheet.id)
		: undefined
	if (definedName) {
		const key = definedName.name.toLowerCase()
		if (seenNames.has(key)) return []
		seenNames.add(key)
		return resolveListFormulaValues(definedName.formula, workbook, sheet, seenNames)
	}
	const structured = /^([^[]+)\[([^\]]+)\]$/.exec(normalized)
	if (!structured) return []
	const tableName = structured[1]?.trim()
	const columnName = structured[2]?.trim()
	if (!tableName || !columnName) return []
	for (const candidateSheet of workbook.sheets) {
		for (const table of candidateSheet.tables) {
			if (table.name.toLowerCase() !== tableName.toLowerCase()) continue
			const columnIndex = table.columns.findIndex(
				(column) => column.name.toLowerCase() === columnName.toLowerCase(),
			)
			if (columnIndex < 0) return []
			const col = table.ref.start.col + columnIndex
			const startRow = table.ref.start.row + (table.hasHeaders ? 1 : 0)
			const endRow = table.ref.end.row - (table.hasTotals ? 1 : 0)
			const values: CellValue[] = []
			for (let row = startRow; row <= endRow; row++) {
				const value = candidateSheet.cells.get(row, col)?.value
				if (value && !isEmpty(value)) values.push(value)
			}
			return values
		}
	}
	return []
}

function compareWithOperator(
	value: number,
	min: number,
	max: number,
	operator: string | undefined,
): boolean {
	switch (operator) {
		case 'between':
			return value >= min && value <= max
		case 'notBetween':
			return value < min || value > max
		case 'equal':
			return value === min
		case 'notEqual':
			return value !== min
		case 'greaterThan':
			return value > min
		case 'lessThan':
			return value < min
		case 'greaterThanOrEqual':
			return value >= min
		case 'lessThanOrEqual':
			return value <= min
		default:
			return value >= min && value <= max
	}
}

function scalarToNumber(v: CellValue): number | null {
	const s = topLeftScalar(v)
	switch (s.kind) {
		case 'number':
			return s.value
		case 'string': {
			const n = Number.parseFloat(s.value)
			return Number.isFinite(n) ? n : null
		}
		case 'boolean':
			return s.value ? 1 : 0
		case 'date':
			return s.serial
		case 'empty':
			return null
		default:
			return null
	}
}

function scalarToNumberOrInvalid(v: CellValue): NumericFormulaValue {
	const s = topLeftScalar(v)
	switch (s.kind) {
		case 'number':
			return s.value
		case 'date':
			return s.serial
		case 'boolean':
			return s.value ? 1 : 0
		case 'empty':
			return null
		case 'string': {
			const n = Number.parseFloat(s.value)
			return Number.isFinite(n) ? n : INVALID_NUMERIC_FORMULA
		}
		default:
			return INVALID_NUMERIC_FORMULA
	}
}

function scalarToString(v: CellValue): string {
	const s = topLeftScalar(v)
	switch (s.kind) {
		case 'string':
			return s.value
		case 'number':
			return String(s.value)
		case 'boolean':
			return s.value ? 'TRUE' : 'FALSE'
		case 'date':
			return String(s.serial)
		case 'richText':
			return s.runs.map((r) => r.text).join('')
		case 'empty':
			return ''
		default:
			return ''
	}
}

function numericValidationPasses(
	value: number,
	min: NumericFormulaValue,
	max: NumericFormulaValue,
	operator: string,
): boolean {
	if (min === INVALID_NUMERIC_FORMULA || max === INVALID_NUMERIC_FORMULA) return false
	if (min !== null && max !== null) return compareWithOperator(value, min, max, operator)
	if (operator === 'between' || operator === 'notBetween') return true
	if (min !== null) return compareWithOperator(value, min, min, operator)
	return true
}

function listCandidateMatches(candidate: ListCandidate, value: CellValue): boolean {
	const scalar = topLeftScalar(value)
	if (candidate.kind === 'literal') return scalarToString(value).trim() === candidate.value
	const allowed = topLeftScalar(candidate.value)
	if (allowed.kind === 'string' && scalar.kind === 'string') {
		return allowed.value.trim().toLowerCase() === scalar.value.trim().toLowerCase()
	}
	if (allowed.kind === 'number' && scalar.kind === 'number') return allowed.value === scalar.value
	if (allowed.kind === 'date' && scalar.kind === 'date') return allowed.serial === scalar.serial
	if (allowed.kind === 'boolean' && scalar.kind === 'boolean') return allowed.value === scalar.value
	return false
}

function formatListCandidates(candidates: readonly ListCandidate[]): string {
	return candidates
		.map((candidate) =>
			candidate.kind === 'literal' ? candidate.value : scalarToString(candidate.value).trim(),
		)
		.join(', ')
}

export function validateCellValue(
	sheet: Sheet,
	row: number,
	col: number,
	value: CellValue,
	workbook?: Workbook,
): ValidationResult {
	const rule = findRuleForCell(sheet, row, col)
	if (!rule) return { valid: true }

	const allowBlank = rule.allowBlank === true
	if (isEmpty(value) || (value.kind === 'string' && value.value === '')) {
		return allowBlank
			? { valid: true }
			: { valid: false, rule, message: rule.error ?? 'Value is required' }
	}

	const type = rule.type ?? 'none'
	const formula1 = rule.formula1
	const formula2 = rule.formula2
	const operator = rule.operator ?? 'between'

	switch (type) {
		case 'whole': {
			const num = scalarToNumber(value)
			if (num === null) {
				return { valid: false, rule, message: rule.error ?? 'Value must be a whole number' }
			}
			if (num !== Math.trunc(num)) {
				return { valid: false, rule, message: rule.error ?? 'Value must be a whole number' }
			}
			const min = formulaToNumber(formula1, workbook, sheet, row, col)
			const max = formulaToNumber(formula2, workbook, sheet, row, col)
			if (!numericValidationPasses(num, min, max, operator)) {
				return { valid: false, rule, message: rule.error ?? `Value must satisfy condition` }
			}
			return { valid: true }
		}
		case 'decimal': {
			const num = scalarToNumber(value)
			if (num === null) {
				return { valid: false, rule, message: rule.error ?? 'Value must be a number' }
			}
			const min = formulaToNumber(formula1, workbook, sheet, row, col)
			const max = formulaToNumber(formula2, workbook, sheet, row, col)
			if (!numericValidationPasses(num, min, max, operator)) {
				return { valid: false, rule, message: rule.error ?? `Value must satisfy condition` }
			}
			return { valid: true }
		}
		case 'list': {
			if (!formula1) return { valid: true }
			const allowed = formulaToListCandidates(formula1, workbook, sheet, row, col)
			const match = allowed.some((candidate) => listCandidateMatches(candidate, value))
			if (!match) {
				return {
					valid: false,
					rule,
					message: rule.error ?? `Value must be one of: ${formatListCandidates(allowed)}`,
				}
			}
			return { valid: true }
		}
		case 'textLength': {
			const str = scalarToString(value)
			const len = str.length
			const min = formulaToNumber(formula1, workbook, sheet, row, col)
			const max = formulaToNumber(formula2, workbook, sheet, row, col)
			if (min !== INVALID_NUMERIC_FORMULA && max !== INVALID_NUMERIC_FORMULA) {
				if (min !== null && max !== null) {
					if (!compareWithOperator(len, min, max, operator)) {
						return {
							valid: false,
							rule,
							message: rule.error ?? `Text length must be between ${min} and ${max}`,
						}
					}
				} else if (max !== null) {
					if (len > max) {
						return {
							valid: false,
							rule,
							message: rule.error ?? `Text length must be at most ${max}`,
						}
					}
				} else if (min !== null) {
					if (!compareWithOperator(len, min, min, operator)) {
						return {
							valid: false,
							rule,
							message: rule.error ?? `Text length must satisfy condition`,
						}
					}
				}
			} else {
				return { valid: false, rule, message: rule.error ?? `Text length must satisfy condition` }
			}
			return { valid: true }
		}
		case 'date': {
			const num = scalarToNumber(value)
			if (num === null) {
				return { valid: false, rule, message: rule.error ?? 'Value must be a date' }
			}
			const min = formulaToNumber(formula1, workbook, sheet, row, col)
			const max = formulaToNumber(formula2, workbook, sheet, row, col)
			if (!numericValidationPasses(num, min, max, operator)) {
				return { valid: false, rule, message: rule.error ?? `Date must satisfy condition` }
			}
			return { valid: true }
		}
		case 'time': {
			const num = scalarToNumber(value)
			if (num === null) {
				return { valid: false, rule, message: rule.error ?? 'Value must be a time' }
			}
			const min = formulaToNumber(formula1, workbook, sheet, row, col)
			const max = formulaToNumber(formula2, workbook, sheet, row, col)
			if (!numericValidationPasses(num, min, max, operator)) {
				return { valid: false, rule, message: rule.error ?? `Time must satisfy condition` }
			}
			return { valid: true }
		}
		case 'custom': {
			if (!formula1 || !workbook) return { valid: true }
			const existing = sheet.cells.get(row, col)
			let result: CellValue | null
			try {
				sheet.cells.set(row, col, {
					value,
					formula: existing?.formula ?? null,
					styleId: existing?.styleId ?? DEFAULT_STYLE_ID,
					formulaInfo: existing?.formulaInfo,
				})
				result = evaluateFormulaScalar(formula1, workbook, sheet, row, col)
			} finally {
				if (existing) {
					sheet.cells.set(row, col, existing)
				} else {
					sheet.cells.delete(row, col)
				}
			}
			if (!result) return { valid: true }
			const scalar = topLeftScalar(result)
			if (scalar.kind === 'error') {
				return { valid: false, rule, message: rule.error ?? scalar.value }
			}
			const pass =
				scalar.kind === 'boolean'
					? scalar.value
					: scalar.kind === 'number'
						? scalar.value !== 0
						: scalar.kind === 'string'
							? scalar.value.toUpperCase() === 'TRUE'
							: !isEmpty(scalar)
			return pass
				? { valid: true }
				: { valid: false, rule, message: rule.error ?? 'Custom validation failed' }
		}
		default:
			return { valid: true }
	}
}
