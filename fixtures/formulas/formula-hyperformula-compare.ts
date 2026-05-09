/**
 * Compare formula results: Ascend vs HyperFormula on a small shared scenario set.
 * Run: bun run fixtures/formulas/formula-hyperformula-compare.ts
 */
import { HyperFormula } from 'hyperformula'
import { createWorkbook, parseA1, type StyleId } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import type { CellValue } from '../../packages/schema/src/index.ts'
import { booleanValue, EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId

interface Scenario {
	name: string
	setup: Record<string, number | string | boolean>
	formula: string
}

export const hyperFormulaComparisonScenarios: readonly Scenario[] = [
	{ name: 'SUM', setup: { A1: 1, A2: 2, A3: 3 }, formula: '=SUM(A1:A3)' },
	{ name: 'AVERAGE', setup: { A1: 10, A2: 20 }, formula: '=AVERAGE(A1:A2)' },
	{ name: 'IF true', setup: { A1: 5 }, formula: '=IF(A1>3,"big","small")' },
	{ name: 'IF false', setup: { A1: 1 }, formula: '=IF(A1>3,"big","small")' },
	{
		name: 'VLOOKUP',
		setup: { A1: 'a', B1: 10, A2: 'b', B2: 20 },
		formula: '=VLOOKUP("b",A1:B2,2,FALSE)',
	},
	{ name: 'CONCATENATE', setup: { A1: 'Hello', A2: 'World' }, formula: '=CONCATENATE(A1," ",A2)' },
	{ name: 'LEN', setup: { A1: 'abcde' }, formula: '=LEN(A1)' },
	{ name: 'MAX', setup: { A1: 5, A2: 3, A3: 9 }, formula: '=MAX(A1:A3)' },
	{ name: 'MIN', setup: { A1: 5, A2: 3, A3: 9 }, formula: '=MIN(A1:A3)' },
	{ name: 'COUNT', setup: { A1: 1, A2: 2, A3: 'x' }, formula: '=COUNT(A1:A3)' },
	{ name: 'COUNTA', setup: { A1: 1, A2: 2, A3: 'x' }, formula: '=COUNTA(A1:A3)' },
	{ name: 'ABS', setup: { A1: -7 }, formula: '=ABS(A1)' },
	{ name: 'ROUND', setup: { A1: Math.PI }, formula: '=ROUND(A1,2)' },
	{ name: 'LEFT', setup: { A1: 'Hello' }, formula: '=LEFT(A1,3)' },
	{ name: 'RIGHT', setup: { A1: 'Hello' }, formula: '=RIGHT(A1,2)' },
]

function inputToCV(v: number | string | boolean): CellValue {
	if (typeof v === 'number') return numberValue(v)
	if (typeof v === 'boolean') return booleanValue(v)
	return stringValue(String(v))
}

function cvToString(cv: CellValue): string {
	switch (cv.kind) {
		case 'number':
			return String(cv.value)
		case 'string':
			return `"${cv.value}"`
		case 'boolean':
			return String(cv.value)
		case 'error':
			return cv.value
		case 'empty':
			return '(empty)'
		case 'date':
			return `date:${String(cv.serial)}`
		default:
			return JSON.stringify(cv)
	}
}

export interface HyperFormulaComparisonResult {
	scenarios: number
	mismatches: number
	skipped: number
	rows: Array<{
		name: string
		ascend: string
		hyperformula: string
		status: 'match' | 'skip' | 'diff'
	}>
}

export function runHyperFormulaComparison(
	options: { log?: boolean } = {},
): HyperFormulaComparisonResult {
	let mismatches = 0
	let skipped = 0
	const rows: HyperFormulaComparisonResult['rows'] = []
	for (const s of hyperFormulaComparisonScenarios) {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (const [a1, val] of Object.entries(s.setup)) {
			const { row, col } = parseA1(a1)
			sheet.cells.set(row, col, { value: inputToCV(val), formula: null, styleId: SID })
		}
		const f = s.formula.startsWith('=') ? s.formula.slice(1) : s.formula
		sheet.cells.set(10, 0, { value: EMPTY, formula: f, styleId: SID })
		recalculate(wb, defaultCalcContext())
		const ascendResult = sheet.cells.readValue(10, 0)

		const data: (number | string | boolean | null)[][] = Array.from({ length: 11 }, () => [null])
		for (const [a1, val] of Object.entries(s.setup)) {
			const { row, col } = parseA1(a1)
			while (data.length <= row) data.push([null])
			const r = data[row]
			if (r) {
				while (r.length <= col) r.push(null)
				r[col] = val
			}
		}
		const formulaRow = data[10]
		if (formulaRow) formulaRow[0] = s.formula
		const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
		hf.addSheet('Sheet1')
		hf.suspendEvaluation()
		hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
		hf.resumeEvaluation()
		const hfRaw = hf.getCellValue({ sheet: 0, row: 10, col: 0 })
		hf.destroy()

		let hfStr: string
		if (hfRaw === null || hfRaw === undefined) hfStr = '(empty)'
		else if (typeof hfRaw === 'number') hfStr = String(hfRaw)
		else if (typeof hfRaw === 'boolean') hfStr = String(hfRaw)
		else if (typeof hfRaw === 'string') hfStr = `"${hfRaw}"`
		else hfStr = String(hfRaw)

		const ascendStr = cvToString(ascendResult)
		const unsupportedByComparator = hfStr === '#NAME?' && ascendStr !== '#NAME?'
		const match = ascendStr === hfStr
		const status = match ? 'match' : unsupportedByComparator ? 'skip' : 'diff'
		if (unsupportedByComparator) skipped++
		else if (!match) mismatches++
		rows.push({ name: s.name, ascend: ascendStr, hyperformula: hfStr, status })
		if (options.log) {
			console.log(
				`${match ? 'MATCH' : unsupportedByComparator ? 'SKIP ' : 'DIFF '}  ${s.name.padEnd(16)} ascend=${ascendStr.padEnd(16)} hf=${hfStr}`,
			)
		}
	}
	if (options.log) {
		console.log(
			`\n${String(hyperFormulaComparisonScenarios.length)} scenarios, ${String(mismatches)} mismatches, ${String(skipped)} comparator skips`,
		)
	}
	return {
		scenarios: hyperFormulaComparisonScenarios.length,
		mismatches,
		skipped,
		rows,
	}
}

function main(): void {
	const { mismatches } = runHyperFormulaComparison({ log: true })
	if (mismatches > 0) process.exit(1)
}

if (import.meta.main) {
	main()
}
