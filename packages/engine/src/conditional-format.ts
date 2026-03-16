import type { CellStyle, Sheet, SheetConditionalFormatRule, Workbook } from '@ascend/core'
import { expandRange, parseRange, toA1 } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import { isEmpty, topLeftScalar } from '@ascend/schema'
import { evaluateRelativeFormulaText } from './relative-formula.ts'

export interface ConditionalFormatResult {
	readonly ruleIndex: number
	readonly priority: number
	readonly type: string
	readonly format?: CellStyle
}

interface RuleWithContext {
	readonly cfIndex: number
	readonly ruleIndex: number
	readonly rule: SheetConditionalFormatRule
	readonly priority: number
	readonly sqref: string
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

function collectRulesInPriorityOrder(sheet: Sheet): RuleWithContext[] {
	const out: RuleWithContext[] = []
	let globalPriority = 1
	for (let cfIndex = 0; cfIndex < sheet.conditionalFormats.length; cfIndex++) {
		const cf = sheet.conditionalFormats[cfIndex]
		if (!cf) continue
		for (let ruleIndex = 0; ruleIndex < cf.rules.length; ruleIndex++) {
			const rule = cf.rules[ruleIndex]
			if (!rule) continue
			const priority = rule.priority ?? globalPriority++
			out.push({ cfIndex, ruleIndex, rule, priority, sqref: cf.sqref })
		}
	}
	out.sort((a, b) => a.priority - b.priority)
	return out
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

function parseNumericFormula(formula: string): number | null {
	const s = formula.trim()
	const n = Number.parseFloat(s)
	return Number.isFinite(n) ? n : null
}

function compareCellIs(
	cellNum: number | null,
	operator: string | undefined,
	v1: number | null,
	v2: number | null,
): boolean {
	const min = v1 ?? v2 ?? 0
	const max = v2 ?? v1 ?? min

	if (cellNum === null) return false

	switch (operator) {
		case 'between':
			return v1 !== null && v2 !== null && cellNum >= v1 && cellNum <= v2
		case 'notBetween':
			return v1 !== null && v2 !== null && (cellNum < v1 || cellNum > v2)
		case 'equal':
			return cellNum === min
		case 'notEqual':
			return cellNum !== min
		case 'greaterThan':
			return cellNum > min
		case 'lessThan':
			return cellNum < min
		case 'greaterThanOrEqual':
			return cellNum >= min
		case 'lessThanOrEqual':
			return cellNum <= min
		default:
			return cellNum >= min && cellNum <= max
	}
}

function unquoteFormula(formula: string): string {
	const s = formula.trim()
	if (s.startsWith('"') && s.endsWith('"')) return s.slice(1, -1).replace(/""/g, '"')
	return s
}

function evaluateFormulaNumber(
	workbook: Workbook,
	sheetIndex: number,
	anchorRow: number,
	anchorCol: number,
	row: number,
	col: number,
	formula: string | undefined,
): number | null {
	if (!formula) return null
	const direct = parseNumericFormula(formula)
	if (direct !== null) return direct
	const value = evaluateRelativeFormulaText(
		formula,
		workbook,
		sheetIndex,
		anchorRow,
		anchorCol,
		row,
		col,
	)
	return scalarToNumber(value)
}

function ruleMatches(
	rule: SheetConditionalFormatRule,
	workbook: Workbook,
	sheetIndex: number,
	anchorRow: number,
	anchorCol: number,
	cellValue: CellValue,
	row: number,
	col: number,
): boolean {
	const type = rule.type
	const formulas = rule.formulas ?? []
	const formula1 = formulas[0]
	const formula2 = formulas[1]
	const operator = rule.operator

	if (type === 'expression' || type === 'formula') {
		if (!formula1) return false
		const value = evaluateRelativeFormulaText(
			formula1,
			workbook,
			sheetIndex,
			anchorRow,
			anchorCol,
			row,
			col,
		)
		const scalar = topLeftScalar(value)
		if (scalar.kind === 'error') return false
		if (scalar.kind === 'boolean') return scalar.value
		if (scalar.kind === 'number') return scalar.value !== 0
		if (scalar.kind === 'string') return scalar.value.toUpperCase() === 'TRUE'
		return !isEmpty(scalar)
	}

	if (type === 'cellIs') {
		const cellNum = scalarToNumber(cellValue)
		return compareCellIs(
			cellNum,
			operator,
			evaluateFormulaNumber(workbook, sheetIndex, anchorRow, anchorCol, row, col, formula1),
			evaluateFormulaNumber(workbook, sheetIndex, anchorRow, anchorCol, row, col, formula2),
		)
	}

	if (type === 'containsText') {
		const text = scalarToString(cellValue)
		if (!formula1) return false
		const pattern = unquoteFormula(formula1)
		const idx = text.indexOf(pattern)
		return idx >= 0
	}

	if (type === 'notContainsText') {
		const text = scalarToString(cellValue)
		if (!formula1) return true
		const pattern = unquoteFormula(formula1)
		return text.indexOf(pattern) < 0
	}

	if (type === 'beginsWith') {
		const text = scalarToString(cellValue)
		if (!formula1) return false
		const prefix = unquoteFormula(formula1)
		return text.startsWith(prefix)
	}

	if (type === 'endsWith') {
		const text = scalarToString(cellValue)
		if (!formula1) return false
		const suffix = unquoteFormula(formula1)
		return text.endsWith(suffix)
	}

	if (type === 'containsBlanks') {
		return isEmpty(cellValue) || (cellValue.kind === 'string' && cellValue.value === '')
	}

	if (type === 'notContainsBlanks') {
		return !isEmpty(cellValue) && !(cellValue.kind === 'string' && cellValue.value === '')
	}

	return false
}

export function evaluateConditionalFormats(
	sheet: Sheet,
	workbook: Workbook,
	_calcContext?: unknown,
): Map<string, ConditionalFormatResult[]> {
	const result = new Map<string, ConditionalFormatResult[]>()
	const rules = collectRulesInPriorityOrder(sheet)
	const sheetIndex = workbook.sheets.indexOf(sheet)

	const cellsToCheck = new Set<string>()
	const anchors = new Map<string, { row: number; col: number }>()
	for (const cf of sheet.conditionalFormats) {
		const parts = cf.sqref.split(/\s+/)
		for (const part of parts) {
			try {
				const range = parseRange(part)
				anchors.set(part, { row: range.start.row, col: range.start.col })
				for (const ref of expandRange(range)) {
					cellsToCheck.add(toA1(ref))
				}
			} catch {
				// skip invalid range
			}
		}
	}

	for (const a1 of cellsToCheck) {
		const ref = parseRange(a1)
		const row = ref.start.row
		const col = ref.start.col
		const cellValue = sheet.cells.readValue(row, col)
		const matched: ConditionalFormatResult[] = []
		let stopIfTrue = false

		for (const { rule, sqref, ruleIndex, priority } of rules) {
			if (stopIfTrue) break
			if (!cellInSqref(sqref, row, col)) continue
			const part = sqref.split(/\s+/).find((candidate) => {
				try {
					const range = parseRange(candidate)
					return (
						row >= range.start.row &&
						row <= range.end.row &&
						col >= range.start.col &&
						col <= range.end.col
					)
				} catch {
					return false
				}
			})
			const anchor = part ? anchors.get(part) : undefined
			if (
				!ruleMatches(
					rule,
					workbook,
					sheetIndex,
					anchor?.row ?? row,
					anchor?.col ?? col,
					cellValue,
					row,
					col,
				)
			) {
				continue
			}

			const res: ConditionalFormatResult = {
				ruleIndex,
				priority,
				type: rule.type,
				...(rule.style !== undefined ? { format: rule.style } : {}),
			}
			matched.push(res)
			if (rule.stopIfTrue) stopIfTrue = true
		}

		if (matched.length > 0) {
			result.set(a1, matched)
		}
	}

	return result
}
