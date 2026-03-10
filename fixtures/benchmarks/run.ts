import { createWorkbook, type StyleId, type Workbook } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { EMPTY, numberValue } from '../../packages/schema/src/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import {
	type BenchmarkCaseResult,
	createBenchmarkSuite,
	formatBytes,
	formatRate,
	summarizeSamples,
} from './results.ts'

const SID = 0 as StyleId

interface ScenarioInput {
	readonly workbook?: Workbook
	readonly bytes?: Uint8Array
	readonly rows: number
	readonly cols: number
	readonly cells: number
}

interface ScenarioRunResult {
	readonly assertions?: Record<string, string | number | boolean | null>
}

interface Scenario {
	readonly name: string
	readonly category: 'read' | 'write' | 'calc' | 'workflow'
	build(): ScenarioInput
	run(input: ScenarioInput): Promise<ScenarioRunResult | undefined> | ScenarioRunResult | undefined
}

function requireBytes(input: ScenarioInput): Uint8Array {
	if (!input.bytes) throw new Error('Scenario bytes were not built')
	return input.bytes
}

function requireWorkbook(input: ScenarioInput): Workbook {
	if (!input.workbook) throw new Error('Scenario workbook was not built')
	return input.workbook
}

function mustWrite(workbook: Workbook): Uint8Array {
	const result = writeXlsx(workbook)
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function setNumberCell(workbook: Workbook, row: number, col: number, value: number): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	sheet.cells.set(row, col, { value: numberValue(value), formula: null, styleId: SID })
}

function setFormulaCell(workbook: Workbook, row: number, col: number, formula: string): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	sheet.cells.set(row, col, { value: EMPTY, formula, styleId: SID })
}

function buildDenseWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			setNumberCell(workbook, r, c, r * cols + c + 1)
		}
	}
	return workbook
}

function buildStringDenseWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			const value = c % 2 === 0 ? `label-${r}-${c}` : `shared-${c % 5}`
			sheet.cells.set(r, c, {
				value: { kind: 'string', value },
				formula: null,
				styleId: SID,
			})
		}
	}
	return workbook
}

function buildSparseWorkbook(rows: number, cols: number, step: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rows; r += step) {
		for (let c = 0; c < cols; c++) {
			setNumberCell(workbook, r, c, r + c + 1)
		}
	}
	return workbook
}

function buildMultiSheetWorkbook(sheetCount: number, rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	for (let s = 0; s < sheetCount; s++) {
		const sheet = workbook.addSheet(`Sheet${s + 1}`)
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				sheet.cells.set(r, c, {
					value: numberValue((s + 1) * (r + 1) * (c + 1)),
					formula: null,
					styleId: SID,
				})
			}
		}
	}
	return workbook
}

function buildFormulaChainWorkbook(length: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	setNumberCell(workbook, 0, 0, 1)
	for (let r = 1; r < length; r++) {
		setFormulaCell(workbook, r, 0, `A${r}+1`)
	}
	return workbook
}

function buildRangeAggregationWorkbook(length: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < length; r++) {
		setNumberCell(workbook, r, 0, r + 1)
		setFormulaCell(workbook, r, 1, `SUM(A1:A${r + 1})`)
	}
	return workbook
}

const scenarios: readonly Scenario[] = [
	{
		name: 'write-dense-40k',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			return { workbook, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'write-large-100k',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			return { workbook, rows: 5000, cols: 20, cells: 100_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'write-multi-sheet',
		category: 'write',
		build() {
			const workbook = buildMultiSheetWorkbook(8, 1000, 10)
			return { workbook, rows: 1000, cols: 10, cells: 80_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'roundtrip-dense-40k',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
			mustWrite(result.value.workbook)
		},
	},
	{
		name: 'read-large-200k',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(10_000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 10_000, cols: 20, cells: 200_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-metadata-dense',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { mode: 'metadata-only' })
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-full-dense',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-full-string-dense',
		category: 'read',
		build() {
			const workbook = buildStringDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-full-sparse',
		category: 'read',
		build() {
			const workbook = buildSparseWorkbook(50_000, 8, 200)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 50_000, cols: 8, cells: 2_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-selective-sheet',
		category: 'read',
		build() {
			const workbook = buildMultiSheetWorkbook(4, 800, 10)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 800, cols: 10, cells: 32_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { sheets: ['Sheet3'] })
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-window-dense-values',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 20, cells: 100_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
			const window = wb.readWindow('Sheet1', 'A1:T5000', { rowLimit: 250 })
			if (!window) throw new Error('Dense window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
				},
			}
		},
	},
	{
		name: 'read-window-dense-values-compact',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 20, cells: 100_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
			const window = wb.readWindowCompact('Sheet1', 'A1:T5000', {
				rowLimit: 250,
				includeRefs: false,
			})
			if (!window) throw new Error('Compact dense window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
				},
			}
		},
	},
	{
		name: 'read-window-formula-chain-compact',
		category: 'read',
		build() {
			const workbook = buildFormulaChainWorkbook(6000)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 6000, cols: 1, cells: 6000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'formula' })
			const window = wb.readWindowCompact('Sheet1', 'A1:A6000', {
				rowLimit: 500,
				includeRefs: false,
			})
			if (!window) throw new Error('Formula chain window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
					formulaCount: window.cells.filter((cell) => cell.formula !== null).length,
				},
			}
		},
	},
	{
		name: 'read-window-sparse-wide',
		category: 'read',
		build() {
			const workbook = buildSparseWorkbook(100_000, 20, 500)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 100_000, cols: 20, cells: 4_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
			const window = wb.readWindow('Sheet1', 'A1:T100000', { rowLimit: 5000 })
			if (!window) throw new Error('Sparse window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
				},
			}
		},
	},
	{
		name: 'workflow-reopen-values-window',
		category: 'workflow',
		build() {
			const workbook = buildDenseWorkbook(4000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 4000, cols: 20, cells: 80_000 }
		},
		async run(input) {
			let totalCells = 0
			for (let i = 0; i < 3; i++) {
				const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
				const window = wb.readWindow('Sheet1', 'A1:T4000', { rowLimit: 200 })
				if (!window) throw new Error('Workflow benchmark failed to read Sheet1')
				totalCells += window.cells.length
			}
			return {
				assertions: {
					iterations: 3,
					totalCellsRead: totalCells,
				},
			}
		},
	},
	{
		name: 'formula-inspect-chain',
		category: 'workflow',
		build() {
			const workbook = buildFormulaChainWorkbook(6000)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 6000, cols: 1, cells: 6000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input))
			const info = wb.formula('Sheet1!A6000')
			if (!info) throw new Error('Formula inspect benchmark could not load formula target')
			return {
				assertions: {
					volatile: info.volatile,
					refCount: info.refs.length,
					functionCount: info.functions.length,
				},
			}
		},
	},
	{
		name: 'recalc-formula-chain',
		category: 'calc',
		build() {
			const workbook = buildFormulaChainWorkbook(6000)
			return { workbook, rows: 6000, cols: 1, cells: 6000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-range-aggregation',
		category: 'calc',
		build() {
			const workbook = buildRangeAggregationWorkbook(800)
			return { workbook, rows: 800, cols: 2, cells: 1600 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
]

function getRssBytes(): number | undefined {
	try {
		return process.memoryUsage.rss()
	} catch {
		return undefined
	}
}

function renderSummary(results: readonly BenchmarkCaseResult[]): string {
	const headers = [
		'scenario',
		'category',
		'median-ms',
		'p95-ms',
		'cells',
		'bytes',
		'throughput',
		'rss-delta',
	]
	const rows = results.map((result) => [
		result.name,
		result.category,
		result.metrics.medianMs.toFixed(2),
		result.metrics.p95Ms.toFixed(2),
		String(result.dimensions.cells ?? 'n/a'),
		typeof result.dimensions.bytes === 'number' ? formatBytes(result.dimensions.bytes) : 'n/a',
		result.metrics.throughputPerSec !== undefined
			? formatRate(result.metrics.throughputPerSec)
			: 'n/a',
		result.metrics.rssDeltaBytes !== undefined ? formatBytes(result.metrics.rssDeltaBytes) : 'n/a',
	])

	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')

	return [
		line(headers),
		widths.map((width) => '─'.repeat(width)).join('──'),
		...rows.map(line),
	].join('\n')
}

async function runScenario(scenario: Scenario, repeat: number): Promise<BenchmarkCaseResult> {
	const samples: Array<{
		readonly durationMs: number
		readonly throughputPerSec: number
		readonly rssDeltaBytes?: number
	}> = []
	let firstInput: ScenarioInput | undefined
	let assertions: Record<string, string | number | boolean | null> | undefined
	for (let i = 0; i < repeat; i++) {
		const input = scenario.build()
		firstInput ??= input
		const rssBefore = getRssBytes()
		const start = performance.now()
		const runResult = await scenario.run(input)
		const durationMs = performance.now() - start
		const rssAfter = getRssBytes()
		samples.push({
			durationMs,
			throughputPerSec:
				durationMs > 0 ? (input.cells / durationMs) * 1000 : Number.POSITIVE_INFINITY,
			rssDeltaBytes:
				rssBefore !== undefined && rssAfter !== undefined
					? Math.max(0, rssAfter - rssBefore)
					: undefined,
		})
		assertions ??= runResult?.assertions
	}
	const input = firstInput
	if (!input) throw new Error(`Scenario "${scenario.name}" did not produce input`)
	return {
		name: scenario.name,
		category: scenario.category,
		dimensions: {
			rows: input.rows,
			cols: input.cols,
			cells: input.cells,
			bytes: input.bytes?.byteLength ?? 0,
			repeat,
		},
		metrics: summarizeSamples(samples),
		...(repeat > 1 ? { samples } : {}),
		...(assertions ? { assertions } : {}),
	}
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

async function runScenarioIsolated(
	scenario: Scenario,
	repeat: number,
	json: boolean,
): Promise<BenchmarkCaseResult> {
	const proc = Bun.spawn(
		[
			'bun',
			'run',
			process.argv[1] ?? import.meta.path,
			'--scenario',
			scenario.name,
			'--repeat',
			String(repeat),
			'--json',
		],
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `Synthetic benchmark scenario "${scenario.name}" failed`)
	}
	const parsed = JSON.parse(stdout) as BenchmarkCaseResult
	if (!json) {
		console.log(`completed ${scenario.name}`)
	}
	return parsed
}

async function main(): Promise<void> {
	const json = process.argv.includes('--json')
	const scenarioName = readFlag('--scenario')
	const repeat = Math.max(1, Number.parseInt(readFlag('--repeat') ?? '1', 10) || 1)
	if (scenarioName) {
		const scenario = scenarios.find((entry) => entry.name === scenarioName)
		if (!scenario) throw new Error(`Unknown synthetic benchmark scenario "${scenarioName}"`)
		const result = await runScenario(scenario, repeat)
		if (json) {
			console.log(JSON.stringify(result, null, 2))
			return
		}
		console.log(renderSummary([result]))
		return
	}
	const results: BenchmarkCaseResult[] = []
	for (const scenario of scenarios) {
		results.push(await runScenarioIsolated(scenario, repeat, json))
	}
	const suite = createBenchmarkSuite({
		suite: 'ascend-synthetic-benchmarks',
		kind: 'synthetic',
		cases: results,
		metadata: {
			repeat,
		},
	})
	if (json) {
		console.log(JSON.stringify(suite, null, 2))
		return
	}

	console.log('Ascend benchmark summary')
	console.log(renderSummary(results))
}

await main()
