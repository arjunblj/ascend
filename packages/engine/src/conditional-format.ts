import type { CellStyle, Sheet, SheetConditionalFormatRule, Workbook } from '@ascend/core'
import { parseRange, type RangeRef, toA1 } from '@ascend/core'
import { hasWildcardPatternSyntax, wildcardMatch } from '@ascend/formulas'
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

interface ParsedSqrefPart {
	readonly ref: string
	readonly range: RangeRef
}

interface DuplicateUniqueValue {
	readonly key: string
	readonly textLike: boolean
}

function parseSqrefParts(sqref: string): ParsedSqrefPart[] {
	const parts = sqref.split(/\s+/)
	const parsed: ParsedSqrefPart[] = []
	for (const part of parts) {
		if (part.length === 0) continue
		try {
			parsed.push({ ref: part, range: parseRange(part) })
		} catch {
			// invalid range part, skip
		}
	}
	return parsed
}

function rangeContainsCell(range: RangeRef, row: number, col: number): boolean {
	return (
		row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col
	)
}

function findSqrefPart(
	parts: readonly ParsedSqrefPart[],
	row: number,
	col: number,
): ParsedSqrefPart | undefined {
	for (const part of parts) {
		if (rangeContainsCell(part.range, row, col)) return part
	}
	return undefined
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

function scalarToRankedNumber(v: CellValue): number | null {
	const s = topLeftScalar(v)
	switch (s.kind) {
		case 'number':
			return s.value
		case 'date':
			return s.serial
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

function duplicateUniqueValue(v: CellValue): DuplicateUniqueValue {
	const s = topLeftScalar(v)
	return {
		key: scalarToString(s).toLowerCase(),
		textLike: s.kind === 'string' || s.kind === 'richText',
	}
}

function compareText(a: string, b: string): number {
	return a.toLowerCase().localeCompare(b.toLowerCase())
}

function compareCellIsScalars(
	cellValue: CellValue,
	operator: string | undefined,
	v1: CellValue | null,
	v2: CellValue | null,
): boolean {
	if (isEmpty(topLeftScalar(cellValue))) return false

	const cellNum = scalarToNumber(cellValue)
	const v1Num = v1 ? scalarToNumber(v1) : null
	const v2Num = v2 ? scalarToNumber(v2) : null
	if (cellNum !== null && (v1Num !== null || v2Num !== null)) {
		return compareCellIs(cellNum, operator, v1Num, v2Num)
	}

	if (!v1) return false
	const cellText = scalarToString(cellValue)
	const v1Text = scalarToString(v1)
	const v2Text = v2 ? scalarToString(v2) : v1Text
	const cmp1 = compareText(cellText, v1Text)
	const cmp2 = compareText(cellText, v2Text)
	const minCmp = compareText(v1Text, v2Text) <= 0 ? cmp1 : cmp2
	const maxCmp = compareText(v1Text, v2Text) <= 0 ? cmp2 : cmp1

	switch (operator) {
		case 'equal':
			return cmp1 === 0
		case 'notEqual':
			return cmp1 !== 0
		case 'greaterThan':
			return cmp1 > 0
		case 'lessThan':
			return cmp1 < 0
		case 'greaterThanOrEqual':
			return cmp1 >= 0
		case 'lessThanOrEqual':
			return cmp1 <= 0
		case 'between':
			return minCmp >= 0 && maxCmp <= 0
		case 'notBetween':
			return minCmp < 0 || maxCmp > 0
		default:
			return minCmp >= 0 && maxCmp <= 0
	}
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

function textRulePattern(
	rule: SheetConditionalFormatRule,
	formula: string | undefined,
): string | null {
	if (rule.text !== undefined) return rule.text
	if (formula === undefined) return null
	return unquoteFormula(formula)
}

function evaluateFormulaValue(
	workbook: Workbook,
	sheetIndex: number,
	anchorRow: number,
	anchorCol: number,
	row: number,
	col: number,
	formula: string | undefined,
): CellValue | null {
	if (!formula) return null
	return evaluateRelativeFormulaText(formula, workbook, sheetIndex, anchorRow, anchorCol, row, col)
}

function evaluateFormulaBoolean(
	workbook: Workbook,
	sheetIndex: number,
	anchorRow: number,
	anchorCol: number,
	row: number,
	col: number,
	formula: string | undefined,
): boolean | null {
	if (!formula) return null
	const value = evaluateRelativeFormulaText(
		formula,
		workbook,
		sheetIndex,
		anchorRow,
		anchorCol,
		row,
		col,
	)
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'boolean':
			return scalar.value
		default:
			return null
	}
}

interface RangeContext {
	readonly allNumerics: number[]
	readonly valueCounts: ReadonlyMap<string, number>
	readonly values: readonly DuplicateUniqueValue[]
	readonly sortedNumerics: readonly number[]
	readonly numericMean: number | null
	readonly numericStdDev: number | null
}

const RANGE_CONTEXT_TYPES = new Set(['duplicateValues', 'uniqueValues', 'top10', 'aboveAverage'])
const OCCUPIED_ONLY_RULE_TYPES = new Set([
	'notContainsBlanks',
	'containsErrors',
	'duplicateValues',
	'uniqueValues',
	'top10',
	'aboveAverage',
	'timePeriod',
])
const EXCEL_MAX_COLS = 16_384

function needsRangeContext(type: string): boolean {
	return RANGE_CONTEXT_TYPES.has(type)
}

function canSkipBlankCellsForRule(rule: SheetConditionalFormatRule): boolean {
	return OCCUPIED_ONLY_RULE_TYPES.has(rule.type)
}

function buildRangeContext(sheet: Sheet, parts: readonly ParsedSqrefPart[]): RangeContext {
	const allNumerics: number[] = []
	const valueCounts = new Map<string, number>()
	const values: DuplicateUniqueValue[] = []
	for (const part of parts) {
		const range = part.range
		sheet.cells.forEachValueInRange(
			range.start.row,
			range.start.col,
			range.end.row,
			range.end.col,
			(value) => {
				if (!isEmpty(value)) {
					const comparable = duplicateUniqueValue(value)
					values.push(comparable)
					valueCounts.set(comparable.key, (valueCounts.get(comparable.key) ?? 0) + 1)
				}
				const n = scalarToRankedNumber(value)
				if (n !== null) allNumerics.push(n)
			},
		)
	}
	const sortedNumerics = [...allNumerics].sort((a, b) => a - b)
	const numericMean =
		allNumerics.length === 0
			? null
			: allNumerics.reduce((sum, value) => sum + value, 0) / allNumerics.length
	const numericStdDev =
		numericMean === null
			? null
			: Math.sqrt(
					allNumerics.reduce((sum, value) => sum + (value - numericMean) ** 2, 0) /
						allNumerics.length,
				)
	return { allNumerics, valueCounts, values, sortedNumerics, numericMean, numericStdDev }
}

function duplicateUniqueCount(cellValue: CellValue, rangeContext: RangeContext): number {
	const comparable = duplicateUniqueValue(cellValue)
	if (!comparable.textLike || !hasWildcardPatternSyntax(comparable.key)) {
		return rangeContext.valueCounts.get(comparable.key) ?? 0
	}
	let count = 0
	for (const candidate of rangeContext.values) {
		if (candidate.textLike && wildcardMatch(comparable.key, candidate.key)) count++
	}
	return count
}

const EXCEL_EPOCH_OFFSET = 25569
const MS_PER_DAY = 86400000

function todaySerial(): number {
	const now = new Date()
	const utcMidnight = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())
	return Math.floor(utcMidnight / MS_PER_DAY) + EXCEL_EPOCH_OFFSET
}

function serialToDateParts(serial: number): {
	year: number
	month: number
	day: number
} {
	const date = new Date((serial - EXCEL_EPOCH_OFFSET) * MS_PER_DAY)
	return {
		year: date.getUTCFullYear(),
		month: date.getUTCMonth(),
		day: date.getUTCDate(),
	}
}

function isoWeekStart(serial: number): number {
	const date = new Date((serial - EXCEL_EPOCH_OFFSET) * MS_PER_DAY)
	const dow = date.getUTCDay()
	const mondayOffset = dow === 0 ? -6 : 1 - dow
	return serial + mondayOffset
}

function matchTimePeriod(period: string, cellSerial: number, today: number): boolean {
	switch (period) {
		case 'today':
			return cellSerial === today
		case 'yesterday':
			return cellSerial === today - 1
		case 'tomorrow':
			return cellSerial === today + 1
		case 'last7Days':
			return cellSerial >= today - 6 && cellSerial <= today
		case 'thisWeek':
			return isoWeekStart(cellSerial) === isoWeekStart(today)
		case 'lastWeek':
			return isoWeekStart(cellSerial) === isoWeekStart(today) - 7
		case 'nextWeek':
			return isoWeekStart(cellSerial) === isoWeekStart(today) + 7
		case 'thisMonth': {
			const cell = serialToDateParts(cellSerial)
			const now = serialToDateParts(today)
			return cell.year === now.year && cell.month === now.month
		}
		case 'lastMonth': {
			const cell = serialToDateParts(cellSerial)
			const now = serialToDateParts(today)
			const prevMonth = now.month === 0 ? 11 : now.month - 1
			const prevYear = now.month === 0 ? now.year - 1 : now.year
			return cell.year === prevYear && cell.month === prevMonth
		}
		case 'nextMonth': {
			const cell = serialToDateParts(cellSerial)
			const now = serialToDateParts(today)
			const nextMonth = now.month === 11 ? 0 : now.month + 1
			const nextYear = now.month === 11 ? now.year + 1 : now.year
			return cell.year === nextYear && cell.month === nextMonth
		}
		default:
			return false
	}
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
	rangeContext?: RangeContext,
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
		return compareCellIsScalars(
			cellValue,
			operator,
			evaluateFormulaValue(workbook, sheetIndex, anchorRow, anchorCol, row, col, formula1),
			evaluateFormulaValue(workbook, sheetIndex, anchorRow, anchorCol, row, col, formula2),
		)
	}

	if (type === 'containsText') {
		const pattern = textRulePattern(rule, formula1)
		if (pattern === null) return false
		const formulaMatch = evaluateFormulaBoolean(
			workbook,
			sheetIndex,
			anchorRow,
			anchorCol,
			row,
			col,
			formula1,
		)
		if (formulaMatch !== null) return formulaMatch
		const text = scalarToString(cellValue)
		const idx = text.toLowerCase().indexOf(pattern.toLowerCase())
		return idx >= 0
	}

	if (type === 'notContainsText') {
		const pattern = textRulePattern(rule, formula1)
		if (pattern === null) return true
		const formulaMatch = evaluateFormulaBoolean(
			workbook,
			sheetIndex,
			anchorRow,
			anchorCol,
			row,
			col,
			formula1,
		)
		if (formulaMatch !== null) return formulaMatch
		const text = scalarToString(cellValue)
		return text.toLowerCase().indexOf(pattern.toLowerCase()) < 0
	}

	if (type === 'beginsWith') {
		const pattern = textRulePattern(rule, formula1)
		if (pattern === null) return false
		const formulaMatch = evaluateFormulaBoolean(
			workbook,
			sheetIndex,
			anchorRow,
			anchorCol,
			row,
			col,
			formula1,
		)
		if (formulaMatch !== null) return formulaMatch
		const text = scalarToString(cellValue)
		return text.toLowerCase().startsWith(pattern.toLowerCase())
	}

	if (type === 'endsWith') {
		const pattern = textRulePattern(rule, formula1)
		if (pattern === null) return false
		const formulaMatch = evaluateFormulaBoolean(
			workbook,
			sheetIndex,
			anchorRow,
			anchorCol,
			row,
			col,
			formula1,
		)
		if (formulaMatch !== null) return formulaMatch
		const text = scalarToString(cellValue)
		return text.toLowerCase().endsWith(pattern.toLowerCase())
	}

	if (type === 'containsBlanks') {
		return isEmpty(cellValue) || (cellValue.kind === 'string' && cellValue.value === '')
	}

	if (type === 'notContainsBlanks') {
		return !isEmpty(cellValue) && !(cellValue.kind === 'string' && cellValue.value === '')
	}

	if (type === 'containsErrors') {
		return topLeftScalar(cellValue).kind === 'error'
	}

	if (type === 'notContainsErrors') {
		return topLeftScalar(cellValue).kind !== 'error'
	}

	if (type === 'duplicateValues') {
		if (!rangeContext || isEmpty(cellValue)) return false
		return duplicateUniqueCount(cellValue, rangeContext) > 1
	}

	if (type === 'uniqueValues') {
		if (!rangeContext || isEmpty(cellValue)) return false
		return duplicateUniqueCount(cellValue, rangeContext) === 1
	}

	if (type === 'top10') {
		if (!rangeContext) return false
		const cellNum = scalarToRankedNumber(cellValue)
		if (cellNum === null) return false
		const sortedNumerics = rangeContext.sortedNumerics
		const total = sortedNumerics.length
		if (total === 0) return false

		const rank = rule.rank ?? 10
		const isPercent = rule.percent ?? false
		const isBottom = rule.bottom ?? false
		const n = isPercent ? Math.max(1, Math.ceil((total * rank) / 100)) : Math.min(rank, total)

		if (isBottom) {
			return cellNum <= (sortedNumerics[n - 1] ?? 0)
		}
		return cellNum >= (sortedNumerics[total - n] ?? 0)
	}

	if (type === 'aboveAverage') {
		if (!rangeContext || rangeContext.numericMean === null) return false
		const cellNum = scalarToRankedNumber(cellValue)
		if (cellNum === null) return false
		const above = rule.aboveAverage !== false
		const usesStdDev = rule.stdDev !== undefined && rangeContext.numericStdDev !== null
		const equal = !usesStdDev && rule.equalAverage === true
		const threshold = usesStdDev
			? rangeContext.numericMean + (above ? 1 : -1) * rule.stdDev * rangeContext.numericStdDev
			: rangeContext.numericMean
		if (above && equal) return cellNum >= rangeContext.numericMean
		if (above) return cellNum > threshold
		if (equal) return cellNum <= rangeContext.numericMean
		return cellNum < threshold
	}

	if (type === 'timePeriod') {
		const cellNum = scalarToNumber(cellValue)
		if (cellNum === null) return false
		const period = rule.timePeriod ?? rule.operator
		if (!period) return false
		return matchTimePeriod(period, Math.floor(cellNum), todaySerial())
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
	const parsedSqrefs = new Map<string, readonly ParsedSqrefPart[]>()
	const getParsedSqref = (sqref: string): readonly ParsedSqrefPart[] => {
		const cached = parsedSqrefs.get(sqref)
		if (cached) return cached
		const parsed = parseSqrefParts(sqref)
		parsedSqrefs.set(sqref, parsed)
		return parsed
	}

	const cellsToCheck = new Set<number>()
	for (const cf of sheet.conditionalFormats) {
		const parts = getParsedSqref(cf.sqref)
		const skipBlankCells = cf.rules.every(canSkipBlankCellsForRule)
		for (const part of parts) {
			const range = part.range
			if (skipBlankCells) {
				sheet.cells.forEachValueInRange(
					range.start.row,
					range.start.col,
					range.end.row,
					range.end.col,
					(_value, row, col) => {
						cellsToCheck.add(row * EXCEL_MAX_COLS + col)
					},
				)
				continue
			}
			for (let row = range.start.row; row <= range.end.row; row++) {
				for (let col = range.start.col; col <= range.end.col; col++) {
					cellsToCheck.add(row * EXCEL_MAX_COLS + col)
				}
			}
		}
	}

	const rangeContexts = new Map<string, RangeContext>()
	for (const { rule, sqref } of rules) {
		if (needsRangeContext(rule.type) && !rangeContexts.has(sqref)) {
			rangeContexts.set(sqref, buildRangeContext(sheet, getParsedSqref(sqref)))
		}
	}

	for (const cellKey of cellsToCheck) {
		const row = Math.floor(cellKey / EXCEL_MAX_COLS)
		const col = cellKey - row * EXCEL_MAX_COLS
		const cellValue = sheet.cells.readValue(row, col)
		const matched: ConditionalFormatResult[] = []
		let stopIfTrue = false

		for (const { rule, sqref, ruleIndex, priority } of rules) {
			if (stopIfTrue) break
			const part = findSqrefPart(getParsedSqref(sqref), row, col)
			if (!part) continue
			const anchor = part.range.start
			if (
				!ruleMatches(
					rule,
					workbook,
					sheetIndex,
					anchor.row,
					anchor.col,
					cellValue,
					row,
					col,
					rangeContexts.get(sqref),
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
			result.set(toA1({ row, col }), matched)
		}
	}

	return result
}
