import type { Workbook } from '@ascend/core'
import { indexToColumn, toA1 } from '@ascend/core'
import {
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	cellHasFormula,
	parseCellKey,
	type WorkbookDependencyAnalysis,
	type WorkbookFormulaAnalysis,
} from '@ascend/engine'
import { isError, levenshtein } from '@ascend/schema'

export interface CheckResult {
	readonly passed: boolean
	readonly issues: readonly CheckIssue[]
}

export interface CheckIssue {
	readonly rule: string
	readonly severity: 'error' | 'warning' | 'info'
	readonly message: string
	readonly refs?: readonly string[]
	readonly suggestedFix?: string
}

function findClosestSheetName(target: string, sheetNames: readonly string[]): string | null {
	if (sheetNames.length === 0) return null
	let best: string | null = null
	let bestDist = Number.POSITIVE_INFINITY
	const targetLower = target.toLowerCase()
	for (const name of sheetNames) {
		const dist = levenshtein(targetLower, name.toLowerCase())
		if (dist < bestDist) {
			bestDist = dist
			best = name
		}
	}
	if (best !== null && bestDist <= Math.max(target.length, best.length) * 0.5) {
		return best
	}
	return null
}

function checkBrokenRefs(
	_wb: Workbook,
	analysis: WorkbookFormulaAnalysis,
	sheetNames: readonly string[],
): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		const cellAddr = `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`
		for (const ref of formula.refs) {
			if (ref.kind === 'sheetSpan') {
				const start = analysis.sheetNameIndex.get(ref.startSheet.toLowerCase())
				const end = analysis.sheetNameIndex.get(ref.endSheet.toLowerCase())
				if (start === undefined) {
					const closest = findClosestSheetName(ref.startSheet, sheetNames)
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Reference to non-existent sheet "${ref.startSheet}"`,
						refs: [cellAddr],
						...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
					})
				}
				if (end === undefined) {
					const closest = findClosestSheetName(ref.endSheet, sheetNames)
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Reference to non-existent sheet "${ref.endSheet}"`,
						refs: [cellAddr],
						...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
					})
				}
				if (start !== undefined && end !== undefined && start > end) {
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Invalid 3D sheet span "${ref.startSheet}:${ref.endSheet}"`,
						refs: [cellAddr],
					})
				}
				continue
			}
			if (ref.sheet?.startsWith('[')) continue
			if (ref.sheet && !analysis.sheetNameIndex.has(ref.sheet.toLowerCase())) {
				const closest = findClosestSheetName(ref.sheet, sheetNames)
				issues.push({
					rule: 'broken-refs',
					severity: 'error',
					message: `Reference to non-existent sheet "${ref.sheet}"`,
					refs: [cellAddr],
					...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
				})
			}
		}
	}

	return issues
}

function checkExternalRefs(analysis: WorkbookFormulaAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		const cellAddr = `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`
		for (const ref of formula.refs) {
			if (ref.kind === 'sheetSpan') continue
			if (ref.sheet?.startsWith('[')) {
				issues.push({
					rule: 'external-refs',
					severity: 'warning',
					message: `External workbook reference: ${ref.sheet}`,
					refs: [cellAddr],
					suggestedFix:
						'Replace external reference with a local copy of the data or a defined name',
				})
			}
		}
	}

	return issues
}

function checkCircularRefs(wb: Workbook, analysis: WorkbookDependencyAnalysis): CheckIssue[] {
	return analysis.cycles.map((cycle) => {
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
			suggestedFix: `Break the cycle by removing one of the references: ${refs.join(' → ')} → ${refs[0]}`,
		}
	})
}

function suggestedFixForError(errorType: string): string | null {
	switch (errorType) {
		case '#REF!':
			return 'Check that all referenced cells and ranges still exist; a row, column, or sheet may have been deleted'
		case '#NAME?':
			return 'Check for misspelled function names or undefined named ranges'
		case '#DIV/0!':
			return 'Add a check for zero before dividing (e.g. IF(B1=0, 0, A1/B1))'
		default:
			return null
	}
}

function checkFormulaErrors(wb: Workbook, analysis: WorkbookFormulaAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		const sheet = wb.sheets[formula.sheetIndex]
		const cell = sheet?.cells.get(formula.row, formula.col)
		if (!sheet || !cell || !cellHasFormula(cell)) continue
		if (isError(cell.value)) {
			const errorType = cell.value.value
			const fix = suggestedFixForError(errorType)
			issues.push({
				rule: 'formula-errors',
				severity: 'warning',
				message: `Formula evaluates to ${errorType}`,
				refs: [`${sheet.name}!${toA1({ row: formula.row, col: formula.col })}`],
				...(fix ? { suggestedFix: fix } : {}),
			})
		}
	}

	return issues
}

function checkOrphanedNames(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = new Set(wb.sheets.map((s) => s.name.toLowerCase()))
	const sheetNameList = wb.sheets.map((s) => s.name)

	for (const entry of wb.definedNames.list()) {
		const name = entry.name
		const ref = entry.formula
		const bang = ref.indexOf('!')
		if (bang !== -1) {
			const sheetPart = ref.substring(0, bang).replace(/^'|'$/g, '')
			if (!sheetNames.has(sheetPart.toLowerCase())) {
				const closest = findClosestSheetName(sheetPart, sheetNameList)
				issues.push({
					rule: 'orphaned-names',
					severity: 'warning',
					message: `Defined name "${name}" references non-existent sheet "${sheetPart}"`,
					refs: [ref],
					...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
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

export function check(
	workbook: Workbook,
	analysis?: {
		readonly formulas?: WorkbookFormulaAnalysis
		readonly dependencies?: WorkbookDependencyAnalysis
	},
): CheckResult {
	const formulas = analysis?.formulas ?? analyzeWorkbookFormulas(workbook)
	const dependencies = analysis?.dependencies ?? analyzeWorkbookDependencies(workbook)
	const sheetNames = workbook.sheets.map((s) => s.name)
	const issues = [
		...checkBrokenRefs(workbook, formulas, sheetNames),
		...checkExternalRefs(formulas),
		...checkCircularRefs(workbook, dependencies),
		...checkFormulaErrors(workbook, formulas),
		...checkOrphanedNames(workbook),
		...checkTableIntegrity(workbook),
	]
	return { passed: issues.length === 0, issues }
}
