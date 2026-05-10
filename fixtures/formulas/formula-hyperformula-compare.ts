/**
 * Compare formula results: Ascend vs HyperFormula on shared public and conformance cases.
 * Run: bun run fixtures/formulas/formula-hyperformula-compare.ts
 */

import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { HyperFormula } from 'hyperformula'
import { createWorkbook, parseA1, type StyleId } from '../../packages/core/src/index.ts'
import type { CalcContext } from '../../packages/engine/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import type { CellValue } from '../../packages/schema/src/index.ts'
import {
	booleanValue,
	EMPTY,
	numberValue,
	stringValue,
	topLeftScalar,
} from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId

interface Scenario {
	name: string
	setup: Record<string, number | string | boolean>
	setupFormulas?: Record<string, string>
	formula: string
	context?: {
		dateSystem?: '1900' | '1904'
		now?: string
		today?: string
		randomSeed?: number
		locale?: string
	}
	knownHyperFormulaDivergence?: {
		readonly expectedAscend: string
		readonly note: string
	}
}

interface ConformanceFixture {
	readonly function: string
	readonly cases: readonly ConformanceCase[]
}

interface ConformanceCase {
	readonly description: string
	readonly setup: Record<string, number | string | boolean>
	readonly setupFormulas?: Record<string, string>
	readonly formula: string
	readonly context?: Scenario['context']
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

const KNOWN_CONFORMANCE_DIVERGENCES = new Set([
	knownConformanceKey('date-functions.json', '1900 leap year bug: DATE(1900,2,29) works'),
	knownConformanceKey(
		'date-functions.json',
		'DATEVALUE parses ISO text for a fake settlement date',
	),
	knownConformanceKey('date-functions.json', 'TODAY returns the calculation context date'),
	knownConformanceKey('date-functions.json', 'NOW returns the calculation context timestamp'),
	knownConformanceKey('date-functions.json', 'DATE uses the 1904 system when requested'),
	knownConformanceKey('date-functions.json', 'YEAR interprets serial 0 in the 1904 system'),
	knownConformanceKey(
		'date-functions.json',
		'DATEVALUE respects the 1904 system for imported text dates',
	),
	knownConformanceKey('date-functions.json', 'DATE(1900,1,1) returns serial 1'),
	knownConformanceKey('stats-extended.json', 'SUBTOTAL code 8 computes sample VAR'),
	knownConformanceKey('stats-extended.json', 'CONFIDENCE.T returns t confidence interval width'),
	knownConformanceKey('stats-extended.json', 'F.DIST.RT returns right-tail F distribution'),
	knownConformanceKey('engineering.json', 'IMSQRT of negative one'),
	knownConformanceKey('engineering.json', 'IMPOWER squares a complex number'),
	knownConformanceKey('info-extended.json', 'ISEVEN truncates and returns TRUE for 2.9'),
	knownConformanceKey('text-functions.json', 'TEXT formats percentages for a dashboard cell'),
	knownConformanceKey('text-functions.json', 'VALUE parses a currency string with commas'),
	knownConformanceKey('info-functions.json', 'ISREF returns TRUE for a cell reference'),
	knownConformanceKey('info-functions.json', 'ISREF returns TRUE for a range reference'),
	knownConformanceKey('lookup-functions.json', 'MATCH supports wildcard matching in text lookups'),
	knownConformanceKey(
		'math-functions.json',
		'SUMIF treats a one-cell sum range as the top-left of a criteria-shaped range',
	),
	knownConformanceKey(
		'math-functions.json',
		'AVERAGEIF treats a one-cell average range as the top-left of a criteria-shaped range',
	),
	knownConformanceKey(
		'math-functions.json',
		'SUMIFS supports wildcard and nonblank criteria together',
	),
	knownConformanceKey('math-functions.json', 'COUNTIF treats empty criteria as blank matching'),
	knownConformanceKey(
		'math-functions.json',
		'COUNTIF treats <> with no operand as nonblank matching',
	),
	knownConformanceKey('math-functions.json', 'COUNTBLANK counts missing cells and empty strings'),
	knownConformanceKey('math-functions.json', 'INT rounds negative values down'),
	knownConformanceKey('spill-functions.json', 'FILTER returns fallback when no matches'),
	knownConformanceKey(
		'spill-functions.json',
		'FILTER with no matches and no fallback returns #CALC!',
	),
	knownConformanceKey('spill-functions.json', 'TRANSPOSE column to row, second element via INDEX'),
	knownConformanceKey('math-extended.json', 'RANDBETWEEN with seed'),
	knownConformanceKey('financial-functions.json', 'IPMT returns #NUM! when period is out of range'),
	knownConformanceKey('conformance-date-extended.json', 'DATE(1900,1,1) serial is 1'),
	knownConformanceKey('dynamic-functions.json', 'FILTER returns fallback when no matches'),
])

function knownConformanceKey(file: string, description: string): string {
	return `${file}\0${description}`
}

function inputToCV(v: number | string | boolean): CellValue {
	if (typeof v === 'number') return numberValue(v)
	if (typeof v === 'boolean') return booleanValue(v)
	return stringValue(String(v))
}

function cvToString(cv: CellValue): string {
	const scalar = topLeftScalar(cv)
	if (scalar !== cv) return cvToString(scalar)
	switch (scalar.kind) {
		case 'number':
			return String(scalar.value)
		case 'string':
			return `"${scalar.value}"`
		case 'boolean':
			return String(scalar.value)
		case 'error':
			return scalar.value
		case 'empty':
			return '(empty)'
		case 'date':
			return `date:${String(scalar.serial)}`
		default:
			return JSON.stringify(scalar)
	}
}

function numericallyClose(a: string, b: string): boolean {
	const an = Number(a)
	const bn = Number(b)
	if (!Number.isFinite(an) || !Number.isFinite(bn)) return false
	const scale = Math.max(1, Math.abs(an), Math.abs(bn))
	return Math.abs(an - bn) <= scale * 1e-10
}

function normalizeFormula(formula: string): string {
	return formula.startsWith('=') ? formula.slice(1) : formula
}

function runAscendScenario(s: Scenario): string {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	for (const [a1, val] of Object.entries(s.setup)) {
		const { row, col } = parseA1(a1)
		sheet.cells.set(row, col, { value: inputToCV(val), formula: null, styleId: SID })
	}
	if (s.setupFormulas) {
		for (const [a1, formula] of Object.entries(s.setupFormulas)) {
			const { row, col } = parseA1(a1)
			sheet.cells.set(row, col, {
				value: EMPTY,
				formula: normalizeFormula(formula),
				styleId: SID,
			})
		}
	}
	sheet.cells.set(10, 0, {
		value: EMPTY,
		formula: normalizeFormula(s.formula),
		styleId: SID,
	})
	const baseCtx = defaultCalcContext()
	const ctx: CalcContext = {
		...baseCtx,
		...(s.context?.dateSystem ? { dateSystem: s.context.dateSystem } : {}),
		...(s.context?.randomSeed !== undefined ? { randomSeed: s.context.randomSeed } : {}),
		...(s.context?.locale ? { locale: s.context.locale } : {}),
		...(s.context?.now ? { now: new Date(s.context.now) } : {}),
		...(s.context?.today ? { today: new Date(s.context.today) } : {}),
	}
	wb.calcSettings = {
		...wb.calcSettings,
		dateSystem: ctx.dateSystem,
		iterativeCalc: ctx.iterativeCalc,
	}
	recalculate(wb, ctx)
	return cvToString(sheet.cells.readValue(10, 0))
}

function runHyperFormulaScenario(s: Scenario): string {
	const data: (number | string | boolean | null)[][] = Array.from({ length: 11 }, () => [null])
	const setCell = (a1: string, value: number | string | boolean): void => {
		const { row, col } = parseA1(a1)
		while (data.length <= row) data.push([null])
		const targetRow = data[row]
		if (!targetRow) return
		while (targetRow.length <= col) targetRow.push(null)
		targetRow[col] = value
	}
	for (const [a1, val] of Object.entries(s.setup)) setCell(a1, val)
	if (s.setupFormulas) {
		for (const [a1, formula] of Object.entries(s.setupFormulas)) {
			setCell(a1, formula.startsWith('=') ? formula : `=${formula}`)
		}
	}
	const formulaRow = data[10]
	if (formulaRow) formulaRow[0] = s.formula.startsWith('=') ? s.formula : `=${s.formula}`

	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
	try {
		hf.addSheet('Sheet1')
		hf.suspendEvaluation()
		hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
		hf.resumeEvaluation()
		const hfRaw = hf.getCellValue({ sheet: 0, row: 10, col: 0 })
		if (hfRaw === null || hfRaw === undefined) return '(empty)'
		if (typeof hfRaw === 'number') return String(hfRaw)
		if (typeof hfRaw === 'boolean') return String(hfRaw)
		if (typeof hfRaw === 'string') return `"${hfRaw}"`
		return String(hfRaw)
	} catch (error) {
		return `HFERR:${error instanceof Error ? error.message : String(error)}`
	} finally {
		hf.destroy()
	}
}

function loadConformanceScenarios(): Array<Scenario & { readonly file: string }> {
	const fixturesDir = import.meta.dir
	const scenarios: Array<Scenario & { readonly file: string }> = []
	for (const file of readdirSync(fixturesDir).filter(
		(entry) => entry.endsWith('.json') && entry !== 'package.json',
	)) {
		const fixture = JSON.parse(readFileSync(join(fixturesDir, file), 'utf8')) as ConformanceFixture
		for (const testCase of fixture.cases ?? []) {
			scenarios.push({
				file,
				name: `${file}: ${testCase.description}`,
				setup: testCase.setup ?? {},
				...(testCase.setupFormulas ? { setupFormulas: testCase.setupFormulas } : {}),
				formula: testCase.formula,
				...(testCase.context ? { context: testCase.context } : {}),
				...(KNOWN_CONFORMANCE_DIVERGENCES.has(knownConformanceKey(file, testCase.description))
					? {
							knownHyperFormulaDivergence: {
								expectedAscend: '*',
								note: 'Documented comparator divergence against Excel/Ascend conformance semantics.',
							},
						}
					: {}),
			})
		}
	}
	return scenarios
}

export interface HyperFormulaComparisonResult {
	scenarios: number
	mismatches: number
	skipped: number
	knownDivergences: number
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
	let knownDivergences = 0
	const rows: HyperFormulaComparisonResult['rows'] = []
	const scenarios = [
		...hyperFormulaComparisonScenarios,
		...loadConformanceScenarios(),
	] satisfies readonly Scenario[]
	for (const s of scenarios) {
		const ascendStr = runAscendScenario(s)
		const hfStr = runHyperFormulaScenario(s)
		const unsupportedByComparator =
			(hfStr === '#NAME?' || hfStr === '#ERROR!' || hfStr.startsWith('HFERR:')) &&
			ascendStr !== hfStr
		const knownComparatorDivergence =
			s.knownHyperFormulaDivergence !== undefined &&
			(s.knownHyperFormulaDivergence.expectedAscend === '*' ||
				ascendStr === s.knownHyperFormulaDivergence.expectedAscend)
		const match = ascendStr === hfStr || numericallyClose(ascendStr, hfStr)
		const status = match
			? 'match'
			: knownComparatorDivergence
				? 'known-diff'
				: unsupportedByComparator
					? 'skip'
					: 'diff'
		if (status === 'skip') skipped++
		else if (status === 'known-diff') knownDivergences++
		else if (status === 'diff') mismatches++
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
			`\n${String(scenarios.length)} scenarios, ${String(mismatches)} mismatches, ${String(knownDivergences)} known comparator divergences, ${String(skipped)} comparator skips`,
		)
	}
	return {
		scenarios: scenarios.length,
		mismatches,
		skipped,
		knownDivergences,
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
