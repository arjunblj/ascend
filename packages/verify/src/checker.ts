import type { Workbook } from '@ascend/core'
import { indexToColumn, toA1 } from '@ascend/core'
import { analyzeWorkbook, parseCellKey, type WorkbookAnalysis } from '@ascend/engine'
import { isError } from '@ascend/schema'

export interface CheckResult {
	readonly passed: boolean
	readonly issues: readonly CheckIssue[]
}

export interface CheckIssue {
	readonly rule: string
	readonly severity: 'error' | 'warning' | 'info'
	readonly message: string
	readonly refs?: readonly string[]
}

function checkBrokenRefs(_wb: Workbook, analysis: WorkbookAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		const cellAddr = `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`
		for (const ref of formula.refs) {
			if (ref.sheet && !analysis.sheetNameIndex.has(ref.sheet.toLowerCase())) {
				issues.push({
					rule: 'broken-refs',
					severity: 'error',
					message: `Reference to non-existent sheet "${ref.sheet}"`,
					refs: [cellAddr],
				})
			}
		}
	}

	return issues
}

function checkCircularRefs(wb: Workbook, analysis: WorkbookAnalysis): CheckIssue[] {
	const cycles = analysis.dependencyGraph.detectCycles()
	return cycles.map((cycle) => {
		const refs = cycle.map((key) => {
			const [si, row, col] = parseCellKey(key)
			const s = wb.sheets[si]
			const sheetName = s ? s.name : `Sheet${si}`
			return `${sheetName}!${toA1({ row, col })}`
		})
		return {
			rule: 'circular-refs',
			severity: 'error' as const,
			message: `Circular reference detected involving ${refs.length} cell(s)`,
			refs,
		}
	})
}

function checkFormulaErrors(wb: Workbook, analysis: WorkbookAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		const sheet = wb.sheets[formula.sheetIndex]
		const cell = sheet?.cells.get(formula.row, formula.col)
		if (!sheet || !cell || !cell.formula) continue
		if (isError(cell.value)) {
			issues.push({
				rule: 'formula-errors',
				severity: 'warning',
				message: `Formula evaluates to ${cell.value.value}`,
				refs: [`${sheet.name}!${toA1({ row: formula.row, col: formula.col })}`],
			})
		}
	}

	return issues
}

function checkOrphanedNames(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = new Set(wb.sheets.map((s) => s.name.toLowerCase()))

	for (const entry of wb.definedNames.list()) {
		const name = entry.name
		const ref = entry.formula
		const bang = ref.indexOf('!')
		if (bang !== -1) {
			const sheetPart = ref.substring(0, bang).replace(/^'|'$/g, '')
			if (!sheetNames.has(sheetPart.toLowerCase())) {
				issues.push({
					rule: 'orphaned-names',
					severity: 'warning',
					message: `Defined name "${name}" references non-existent sheet "${sheetPart}"`,
					refs: [ref],
				})
			}
		}
	}

	return issues
}

function checkTableIntegrity(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const sheet of wb.sheets) {
		for (const table of sheet.tables) {
			const rangeWidth = table.ref.end.col - table.ref.start.col + 1
			if (table.columns.length !== rangeWidth) {
				const rangeStr = `${indexToColumn(table.ref.start.col)}${table.ref.start.row + 1}:${indexToColumn(table.ref.end.col)}${table.ref.end.row + 1}`
				issues.push({
					rule: 'table-integrity',
					severity: 'error',
					message: `Table "${table.name}" has ${table.columns.length} columns but range spans ${rangeWidth}`,
					refs: [`${sheet.name}!${rangeStr}`],
				})
			}
		}
	}

	return issues
}

export function check(workbook: Workbook, analysis?: WorkbookAnalysis): CheckResult {
	const compiled = analysis ?? analyzeWorkbook(workbook)
	const issues = [
		...checkBrokenRefs(workbook, compiled),
		...checkCircularRefs(workbook, compiled),
		...checkFormulaErrors(workbook, compiled),
		...checkOrphanedNames(workbook),
		...checkTableIntegrity(workbook),
	]
	return { passed: issues.length === 0, issues }
}
