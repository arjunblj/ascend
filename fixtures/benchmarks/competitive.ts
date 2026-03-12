import { HyperFormula } from 'hyperformula'
import { createWorkbook, type StyleId } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId
const RUNS = 5

interface TimingResult {
	setupMs: number
	recalcMs: number
	totalMs: number
}

function medianOf(values: number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const mid = Math.floor(sorted.length / 2)
	return sorted.length % 2 === 0
		? ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2
		: (sorted[mid] ?? 0)
}

function colLetter(col: number): string {
	return String.fromCharCode(65 + col)
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

function formatThroughput(cellsPerSec: number): string {
	if (!Number.isFinite(cellsPerSec)) return 'n/a'
	if (cellsPerSec >= 1_000_000) return `${(cellsPerSec / 1_000_000).toFixed(2)}M/s`
	if (cellsPerSec >= 1_000) return `${(cellsPerSec / 1_000).toFixed(1)}K/s`
	return `${cellsPerSec.toFixed(1)}/s`
}

function ascendSumScenario(): TimingResult {
	const setupStart = performance.now()
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Missing sheet')
	for (let r = 0; r < 1000; r++) {
		for (let c = 0; c < 10; c++) {
			sheet.cells.set(r, c, { value: numberValue(r * 10 + c + 1), formula: null, styleId: SID })
		}
	}
	for (let i = 0; i < 100; i++) {
		const col = i % 10
		const letter = colLetter(col)
		sheet.cells.set(1000 + i, col, {
			value: EMPTY,
			formula: `SUM(${letter}1:${letter}1000)`,
			styleId: SID,
		})
	}
	const setupMs = performance.now() - setupStart

	const recalcStart = performance.now()
	recalculate(workbook, defaultCalcContext())
	const recalcMs = performance.now() - recalcStart

	return { setupMs, recalcMs, totalMs: setupMs + recalcMs }
}

function hfSumScenario(): TimingResult {
	const setupStart = performance.now()
	const data: (number | string | null)[][] = []
	for (let r = 0; r < 1000; r++) {
		const row: number[] = []
		for (let c = 0; c < 10; c++) {
			row.push(r * 10 + c + 1)
		}
		data.push(row)
	}
	for (let i = 0; i < 100; i++) {
		while (data.length <= 1000 + i) data.push(new Array<null>(10).fill(null))
		const col = i % 10
		const letter = colLetter(col)
		const formulaRow = data[1000 + i]
		if (formulaRow) formulaRow[col] = `=SUM(${letter}1:${letter}1000)`
	}
	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
	hf.addSheet('Sheet1')
	hf.suspendEvaluation()
	hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
	const setupMs = performance.now() - setupStart

	const recalcStart = performance.now()
	hf.resumeEvaluation()
	const recalcMs = performance.now() - recalcStart

	hf.destroy()
	return { setupMs, recalcMs, totalMs: setupMs + recalcMs }
}

function ascendVlookupScenario(): TimingResult {
	const setupStart = performance.now()
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Missing sheet')
	for (let r = 0; r < 5000; r++) {
		sheet.cells.set(r, 0, {
			value: stringValue(`key-${String(r + 1).padStart(5, '0')}`),
			formula: null,
			styleId: SID,
		})
		for (let c = 1; c < 5; c++) {
			sheet.cells.set(r, c, { value: numberValue((r + 1) * c), formula: null, styleId: SID })
		}
	}
	for (let i = 0; i < 500; i++) {
		const keyIndex = 5000 - ((i * 37) % 5000) - 1
		sheet.cells.set(i, 5, {
			value: stringValue(`key-${String(keyIndex + 1).padStart(5, '0')}`),
			formula: null,
			styleId: SID,
		})
		sheet.cells.set(i, 6, {
			value: EMPTY,
			formula: `VLOOKUP(F${i + 1},A$1:E$5000,3,FALSE)`,
			styleId: SID,
		})
	}
	const setupMs = performance.now() - setupStart

	const recalcStart = performance.now()
	recalculate(workbook, defaultCalcContext())
	const recalcMs = performance.now() - recalcStart

	return { setupMs, recalcMs, totalMs: setupMs + recalcMs }
}

function hfVlookupScenario(): TimingResult {
	const setupStart = performance.now()
	const data: (number | string | null)[][] = []
	for (let r = 0; r < 5000; r++) {
		const row: (number | string | null)[] = [
			`key-${String(r + 1).padStart(5, '0')}`,
			(r + 1) * 1,
			(r + 1) * 2,
			(r + 1) * 3,
			(r + 1) * 4,
		]
		if (r < 500) {
			const keyIndex = 5000 - ((r * 37) % 5000) - 1
			row.push(`key-${String(keyIndex + 1).padStart(5, '0')}`)
			row.push(`=VLOOKUP(F${r + 1},A$1:E$5000,3,FALSE)`)
		}
		data.push(row)
	}
	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
	hf.addSheet('Sheet1')
	hf.suspendEvaluation()
	hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
	const setupMs = performance.now() - setupStart

	const recalcStart = performance.now()
	hf.resumeEvaluation()
	const recalcMs = performance.now() - recalcStart

	hf.destroy()
	return { setupMs, recalcMs, totalMs: setupMs + recalcMs }
}

interface ScenarioConfig {
	readonly name: string
	readonly cells: number
	readonly ascend: () => TimingResult
	readonly hf: () => TimingResult
}

const scenarios: readonly ScenarioConfig[] = [
	{
		name: '1000×10 grid + 100 SUM formulas',
		cells: 1000 * 10 + 100,
		ascend: ascendSumScenario,
		hf: hfSumScenario,
	},
	{
		name: '5000×5 grid + 500 VLOOKUP formulas',
		cells: 5000 * 5 + 500 * 2,
		ascend: ascendVlookupScenario,
		hf: hfVlookupScenario,
	},
]

function runScenario(config: ScenarioConfig): void {
	const ascendResults: TimingResult[] = []
	const hfResults: TimingResult[] = []

	runGc()
	config.ascend()
	config.hf()
	runGc()

	for (let i = 0; i < RUNS; i++) {
		runGc()
		ascendResults.push(config.ascend())
		runGc()
		hfResults.push(config.hf())
	}

	const aSetup = medianOf(ascendResults.map((r) => r.setupMs))
	const aRecalc = medianOf(ascendResults.map((r) => r.recalcMs))
	const aTotal = medianOf(ascendResults.map((r) => r.totalMs))
	const aThroughput = aTotal > 0 ? (config.cells / aTotal) * 1000 : Number.POSITIVE_INFINITY

	const hSetup = medianOf(hfResults.map((r) => r.setupMs))
	const hRecalc = medianOf(hfResults.map((r) => r.recalcMs))
	const hTotal = medianOf(hfResults.map((r) => r.totalMs))
	const hThroughput = hTotal > 0 ? (config.cells / hTotal) * 1000 : Number.POSITIVE_INFINITY

	console.log(`\nCompetitive Benchmark: ${config.name}`)
	console.log('Engine          Setup(ms)  Recalc(ms)  Total(ms)  Throughput')
	console.log('─────────────────────────────────────────────────────────────')
	console.log(
		`${'Ascend'.padEnd(16)}${aSetup.toFixed(2).padStart(9)}  ${aRecalc.toFixed(2).padStart(10)}  ${aTotal.toFixed(2).padStart(9)}  ${formatThroughput(aThroughput)}`,
	)
	console.log(
		`${'HyperFormula'.padEnd(16)}${hSetup.toFixed(2).padStart(9)}  ${hRecalc.toFixed(2).padStart(10)}  ${hTotal.toFixed(2).padStart(9)}  ${formatThroughput(hThroughput)}`,
	)
}

for (const scenario of scenarios) {
	runScenario(scenario)
}
