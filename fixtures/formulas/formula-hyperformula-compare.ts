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
	knownHyperFormulaDivergence?: {
		readonly expectedAscend: string
		readonly note: string
	}
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
	{
		name: 'SUMPRODUCT',
		setup: { A1: 1, A2: 2, B1: 10, B2: 20 },
		formula: '=SUMPRODUCT(A1:A2,B1:B2)',
	},
	{
		name: 'COUNTIF',
		setup: { A1: 'red', A2: 'blue', A3: 'red' },
		formula: '=COUNTIF(A1:A3,"red")',
	},
	{
		name: 'SUMIF',
		setup: { A1: 'a', A2: 'b', A3: 'a', B1: 10, B2: 20, B3: 30 },
		formula: '=SUMIF(A1:A3,"a",B1:B3)',
	},
	{
		name: 'AVERAGEIF',
		setup: { A1: 'a', A2: 'b', A3: 'a', B1: 10, B2: 20, B3: 30 },
		formula: '=AVERAGEIF(A1:A3,"a",B1:B3)',
	},
	{ name: 'COUNTBLANK', setup: { A1: 1, A3: 3 }, formula: '=COUNTBLANK(A1:A3)' },
	{ name: 'MEDIAN', setup: { A1: 5, A2: 1, A3: 9 }, formula: '=MEDIAN(A1:A3)' },
	{
		name: 'STDEV.S',
		setup: { A1: 2, A2: 4, A3: 4, A4: 4, A5: 5, A6: 5, A7: 7, A8: 9 },
		formula: '=STDEV.S(A1:A8)',
	},
	{
		name: 'MATCH exact',
		setup: { A1: 'alpha', A2: 'beta', A3: 'gamma' },
		formula: '=MATCH("beta",A1:A3,0)',
	},
	{ name: 'INDEX 2D', setup: { A1: 10, B1: 20, A2: 30, B2: 40 }, formula: '=INDEX(A1:B2,2,2)' },
	{
		name: 'HLOOKUP',
		setup: { A1: 'q1', B1: 'q2', A2: 10, B2: 20 },
		formula: '=HLOOKUP("q2",A1:B2,2,FALSE)',
	},
	{
		name: 'XLOOKUP',
		setup: { A1: 'sku1', A2: 'sku2', B1: 7, B2: 11 },
		formula: '=XLOOKUP("sku2",A1:A2,B1:B2)',
	},
	{ name: 'CHOOSE', setup: {}, formula: '=CHOOSE(2,"first","second","third")' },
	{ name: 'CONCAT', setup: { A1: 'ab', A2: 'cd' }, formula: '=CONCAT(A1:A2)' },
	{ name: 'TEXTJOIN', setup: { A1: 'north', A2: 'south' }, formula: '=TEXTJOIN(",",TRUE,A1:A2)' },
	{ name: 'UPPER', setup: { A1: 'Alpha' }, formula: '=UPPER(A1)' },
	{ name: 'LOWER', setup: { A1: 'Alpha' }, formula: '=LOWER(A1)' },
	{ name: 'TRIM', setup: { A1: '  alpha   beta  ' }, formula: '=TRIM(A1)' },
	{ name: 'SUBSTITUTE', setup: { A1: 'a-b-a' }, formula: '=SUBSTITUTE(A1,"a","x")' },
	{ name: 'ROUNDUP', setup: { A1: 1.21 }, formula: '=ROUNDUP(A1,1)' },
	{ name: 'ROUNDDOWN', setup: { A1: 1.29 }, formula: '=ROUNDDOWN(A1,1)' },
	{ name: 'POWER', setup: { A1: 3, A2: 4 }, formula: '=POWER(A1,A2)' },
	{ name: 'SQRT', setup: { A1: 81 }, formula: '=SQRT(A1)' },
	{
		name: 'MOD',
		setup: { A1: -3, A2: 2 },
		formula: '=MOD(A1,A2)',
		knownHyperFormulaDivergence: {
			expectedAscend: '1',
			note: 'Excel MOD keeps the divisor sign for negative dividends.',
		},
	},
	{
		name: 'INT',
		setup: { A1: -1.2 },
		formula: '=INT(A1)',
		knownHyperFormulaDivergence: {
			expectedAscend: '-2',
			note: 'Excel INT floors negative numbers rather than truncating toward zero.',
		},
	},
	{ name: 'AND', setup: { A1: true, A2: 1 }, formula: '=AND(A1,A2)' },
	{ name: 'OR', setup: { A1: false, A2: 1 }, formula: '=OR(A1,A2)' },
	{ name: 'NOT', setup: { A1: false }, formula: '=NOT(A1)' },
	{ name: 'IFERROR', setup: { A1: 1 }, formula: '=IFERROR(A1,"fallback")' },
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

function numericallyClose(a: string, b: string): boolean {
	const an = Number(a)
	const bn = Number(b)
	if (!Number.isFinite(an) || !Number.isFinite(bn)) return false
	const scale = Math.max(1, Math.abs(an), Math.abs(bn))
	return Math.abs(an - bn) <= scale * 1e-10
}

export interface HyperFormulaComparisonResult {
	scenarios: number
	mismatches: number
	skipped: number
	rows: Array<{
		name: string
		ascend: string
		hyperformula: string
		status: 'match' | 'known-diff' | 'skip' | 'diff'
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
		const knownComparatorDivergence =
			s.knownHyperFormulaDivergence !== undefined &&
			ascendStr === s.knownHyperFormulaDivergence.expectedAscend
		const match = ascendStr === hfStr || numericallyClose(ascendStr, hfStr)
		const status = match
			? 'match'
			: knownComparatorDivergence
				? 'known-diff'
				: unsupportedByComparator
					? 'skip'
					: 'diff'
		if (unsupportedByComparator) skipped++
		else if (!match && !knownComparatorDivergence) mismatches++
		rows.push({ name: s.name, ascend: ascendStr, hyperformula: hfStr, status })
		if (options.log) {
			const note = knownComparatorDivergence ? ` (${s.knownHyperFormulaDivergence?.note})` : ''
			const label =
				status === 'known-diff'
					? 'KNOWN'
					: status === 'skip'
						? 'SKIP '
						: status === 'diff'
							? 'DIFF '
							: 'MATCH'
			console.log(
				`${label}  ${s.name.padEnd(16)} ascend=${ascendStr.padEnd(16)} hf=${hfStr}${note}`,
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
