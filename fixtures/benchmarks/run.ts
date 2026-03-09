import { createWorkbook, type StyleId, type Workbook } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { EMPTY, numberValue } from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId

interface ScenarioInput {
	readonly workbook?: Workbook
	readonly bytes?: Uint8Array
	readonly rows: number
	readonly cols: number
	readonly cells: number
}

interface Scenario {
	readonly name: string
	readonly kind: 'read' | 'calc'
	build(): ScenarioInput
	run(input: ScenarioInput): void
}

interface ScenarioResult {
	readonly name: string
	readonly kind: 'read' | 'calc'
	readonly durationMs: number
	readonly rows: number
	readonly cols: number
	readonly cells: number
	readonly bytes: number
	readonly rssDeltaBytes?: number
	readonly throughputCellsPerSec: number
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
		kind: 'read',
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
		kind: 'read',
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
		kind: 'read',
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
		kind: 'read',
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
		kind: 'read',
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
		kind: 'read',
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
		kind: 'read',
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
		name: 'read-full-sparse',
		kind: 'read',
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
		kind: 'read',
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
		name: 'recalc-formula-chain',
		kind: 'calc',
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
		kind: 'calc',
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

function formatBytes(bytes: number): string {
	if (bytes < 1024) return `${bytes} B`
	if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
	return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

function formatRate(rate: number): string {
	if (!Number.isFinite(rate)) return 'n/a'
	if (rate >= 1_000_000) return `${(rate / 1_000_000).toFixed(2)}M cells/s`
	if (rate >= 1_000) return `${(rate / 1_000).toFixed(1)}K cells/s`
	return `${rate.toFixed(1)} cells/s`
}

function renderSummary(results: readonly ScenarioResult[]): string {
	const headers = ['scenario', 'ms', 'cells', 'bytes', 'throughput', 'rss-delta']
	const rows = results.map((result) => [
		result.name,
		result.durationMs.toFixed(2),
		String(result.cells),
		formatBytes(result.bytes),
		formatRate(result.throughputCellsPerSec),
		result.rssDeltaBytes !== undefined ? formatBytes(result.rssDeltaBytes) : 'n/a',
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

function runScenario(scenario: Scenario): ScenarioResult {
	const input = scenario.build()
	const bytes = input.bytes?.byteLength ?? 0
	const rssBefore = getRssBytes()
	const start = performance.now()
	scenario.run(input)
	const durationMs = performance.now() - start
	const rssAfter = getRssBytes()
	return {
		name: scenario.name,
		kind: scenario.kind,
		durationMs,
		rows: input.rows,
		cols: input.cols,
		cells: input.cells,
		bytes,
		rssDeltaBytes:
			rssBefore !== undefined && rssAfter !== undefined
				? Math.max(0, rssAfter - rssBefore)
				: undefined,
		throughputCellsPerSec:
			durationMs > 0 ? (input.cells / durationMs) * 1000 : Number.POSITIVE_INFINITY,
	}
}

function main(): void {
	const json = process.argv.includes('--json')
	const results = scenarios.map(runScenario)
	if (json) {
		console.log(JSON.stringify(results, null, 2))
		return
	}

	console.log('Ascend benchmark summary')
	console.log(renderSummary(results))
}

main()
