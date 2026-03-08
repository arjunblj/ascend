import type { Workbook } from '@ascend/core'
import { indexToColumn, toA1 } from '@ascend/core'
import { cellKey, DependencyGraph, parseCellKey } from '@ascend/engine'
import { extractRefs, parseFormula } from '@ascend/formulas'
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

function checkBrokenRefs(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = new Set(wb.sheets.map((s) => s.name.toLowerCase()))

	for (const sheet of wb.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			const parsed = parseFormula(cell.formula)
			if (!parsed.ok) continue

			const refs = extractRefs(parsed.value)
			const cellAddr = `${sheet.name}!${toA1({ row, col })}`

			for (const ref of refs) {
				if (ref.sheet && !sheetNames.has(ref.sheet.toLowerCase())) {
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Reference to non-existent sheet "${ref.sheet}"`,
						refs: [cellAddr],
					})
				}
			}
		}
	}

	return issues
}

function checkCircularRefs(wb: Workbook): CheckIssue[] {
	const graph = new DependencyGraph()

	for (let si = 0; si < wb.sheets.length; si++) {
		const sheet = wb.sheets[si]
		if (!sheet) continue
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			const parsed = parseFormula(cell.formula)
			if (!parsed.ok) continue

			const refs = extractRefs(parsed.value)
			const deps: string[] = []

			for (const ref of refs) {
				const targetSheet = ref.sheet
					? wb.sheets.findIndex((s) => s.name.toLowerCase() === ref.sheet?.toLowerCase())
					: si
				if (targetSheet === -1) continue

				if (ref.kind === 'cell') {
					deps.push(cellKey(targetSheet, ref.ref.row, ref.ref.col))
				} else {
					for (let r = ref.start.row; r <= ref.end.row; r++) {
						for (let c = ref.start.col; c <= ref.end.col; c++) {
							deps.push(cellKey(targetSheet, r, c))
						}
					}
				}
			}

			graph.addFormula(cellKey(si, row, col), deps, false)
		}
	}

	const cycles = graph.detectCycles()
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

function checkFormulaErrors(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const sheet of wb.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			if (isError(cell.value)) {
				issues.push({
					rule: 'formula-errors',
					severity: 'warning',
					message: `Formula evaluates to ${cell.value.value}`,
					refs: [`${sheet.name}!${toA1({ row, col })}`],
				})
			}
		}
	}

	return issues
}

function checkOrphanedNames(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = new Set(wb.sheets.map((s) => s.name.toLowerCase()))

	for (const [name, ref] of wb.definedNames) {
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

export function check(workbook: Workbook): CheckResult {
	const issues = [
		...checkBrokenRefs(workbook),
		...checkCircularRefs(workbook),
		...checkFormulaErrors(workbook),
		...checkOrphanedNames(workbook),
		...checkTableIntegrity(workbook),
	]
	return { passed: issues.length === 0, issues }
}
