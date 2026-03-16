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
): number | null {
	if (!formula) return null
	const direct = parseNumericFormula(formula)
	if (direct !== null) return direct
	const value = evaluateFormulaScalar(formula, workbook, sheet, row, col)
	return value ? scalarToNumber(value) : null
}

function formulaToList(
	formula: string | undefined,
	workbook: Workbook | undefined,
	sheet: Sheet,
	row: number,
	col: number,
): string[] {
	if (!formula) return []
	const trimmed = formula.trim()
	if (trimmed.startsWith('"') || !workbook || (trimmed.includes(',') && !/[=:!$]/.test(trimmed))) {
		return parseListValues(formula)
	}
	const value = evaluateFormulaScalar(formula, workbook, sheet, row, col)
	if (!value) return parseListValues(formula)
	if (value.kind === 'array') {
		const out: string[] = []
		for (const line of value.rows) {
			for (const cell of line) out.push(scalarToString(cell).trim())
		}
		return out
	}
	return [scalarToString(value).trim()]
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

export function validateCellValue(
	sheet: Sheet,
	row: number,
	col: number,
	value: CellValue,
	workbook?: Workbook,
): ValidationResult {
	const rule = findRuleForCell(sheet, row, col)
	if (!rule) return { valid: true }

	const allowBlank = rule.allowBlank !== false
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
			if (min !== null && max !== null) {
				if (!compareWithOperator(num, min, max, operator)) {
					return {
						valid: false,
						rule,
						message: rule.error ?? `Value must be between ${min} and ${max}`,
					}
				}
			} else if (min !== null) {
				if (!compareWithOperator(num, min, min, operator)) {
					return { valid: false, rule, message: rule.error ?? `Value must satisfy condition` }
				}
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
			if (min !== null && max !== null) {
				if (!compareWithOperator(num, min, max, operator)) {
					return {
						valid: false,
						rule,
						message: rule.error ?? `Value must be between ${min} and ${max}`,
					}
				}
			} else if (min !== null) {
				if (!compareWithOperator(num, min, min, operator)) {
					return { valid: false, rule, message: rule.error ?? `Value must satisfy condition` }
				}
			}
			return { valid: true }
		}
		case 'list': {
			if (!formula1) return { valid: true }
			const allowed = formulaToList(formula1, workbook, sheet, row, col)
			const str = scalarToString(value)
			const normalized = str.trim()
			const match = allowed.some((a) => a === normalized || String(a) === normalized)
			if (!match) {
				return {
					valid: false,
					rule,
					message: rule.error ?? `Value must be one of: ${allowed.join(', ')}`,
				}
			}
			return { valid: true }
		}
		case 'textLength': {
			const str = scalarToString(value)
			const len = str.length
			const min = formulaToNumber(formula1, workbook, sheet, row, col)
			const max = formulaToNumber(formula2, workbook, sheet, row, col)
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
				if (len < min) {
					return {
						valid: false,
						rule,
						message: rule.error ?? `Text length must be at least ${min}`,
					}
				}
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
			if (min !== null && max !== null) {
				if (!compareWithOperator(num, min, max, operator)) {
					return {
						valid: false,
						rule,
						message: rule.error ?? `Date must be between ${min} and ${max}`,
					}
				}
			} else if (min !== null) {
				if (!compareWithOperator(num, min, min, operator)) {
					return { valid: false, rule, message: rule.error ?? `Date must satisfy condition` }
				}
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
