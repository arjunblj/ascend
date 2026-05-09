import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { TerminalInputParser } from '../../apps/tui/src/input/keymap.ts'
import { runRendererBakeoff } from '../../apps/tui/src/render/renderer-bakeoff.ts'
import { WorkbookTuiEngine } from '../../apps/tui/src/runtime/engine.ts'
import type { TerminalSize } from '../../apps/tui/src/runtime/types.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import { type BenchmarkCaseResult, createBenchmarkSuite, summarizeSamples } from './results.ts'
import { checkTuiTargets, formatTuiTargetResults } from './tui-targets.ts'

interface TuiScenario {
	readonly name: string
	readonly category: 'paint' | 'navigation' | 'edit' | 'command' | 'resize' | 'render'
	readonly dimensions: Record<string, string | number | boolean>
	setup?(size: TerminalSize): Promise<unknown>
	run(
		size: TerminalSize,
		context: unknown,
	): Promise<Record<string, string | number | boolean | null>>
	teardown?(context: unknown): Promise<void>
}

const DEFAULT_SIZE: TerminalSize = { rows: 32, cols: 120 }

const scenarios: readonly TuiScenario[] = [
	{
		name: 'file-hub-first-paint',
		category: 'paint',
		dimensions: { rows: DEFAULT_SIZE.rows, cols: DEFAULT_SIZE.cols },
		async run(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, lines: frame.lines.length }
		},
	},
	{
		name: 'warm-grid-navigation',
		category: 'navigation',
		dimensions: { rows: DEFAULT_SIZE.rows, cols: DEFAULT_SIZE.cols, event: 'ArrowRight' },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			engine.render(size)
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'key', key: 'ArrowRight' })
			const frame = engine.render(size)
			return { finalRef: engine.state().message, bytes: frame.stats.bytes }
		},
	},
	{
		name: 'formula-entry-commit',
		category: 'edit',
		dimensions: { rows: DEFAULT_SIZE.rows, cols: DEFAULT_SIZE.cols, formulas: 50 },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			for (let i = 0; i < 50; i++) {
				await dispatchText(engine, String(i + 1))
				await engine.dispatch({ kind: 'key', key: 'Enter' })
			}
			engine.render(size)
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'goto B1' })
			await dispatchText(engine, '=SUM(A1:A50)')
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'formula-edit-cursor-f4',
		category: 'edit',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			workflow: 'cursor-insert,f4-reference-cycle,commit',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await dispatchText(engine, '3')
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			engine.render(size)
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'goto B1' })
			await dispatchText(engine, '=A1*2')
			await engine.dispatch({ kind: 'key', key: 'Home' })
			await engine.dispatch({ kind: 'key', key: 'ArrowRight' })
			await engine.dispatch({ kind: 'key', key: 'F4' })
			await engine.dispatch({ kind: 'key', key: 'End' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'formula-trace-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			formulas: 50,
			command: 'trace-precedents',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await dispatchText(engine, '1')
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			for (let row = 2; row <= 50; row++) {
				await engine.dispatch({ kind: 'command', command: `goto A${row}` })
				await dispatchText(engine, `=A${row - 1}+1`)
				await engine.dispatch({ kind: 'key', key: 'Enter' })
			}
			await engine.dispatch({ kind: 'command', command: 'goto A50' })
			engine.render(size)
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({
				kind: 'command',
				command: 'trace precedents {"maxDepth":3}',
			})
			const frame = engine.render(size)
			await engine.dispatch({ kind: 'key', key: 'Escape' })
			return { bytes: frame.stats.bytes, focused: engine.state().workspace.focusedRegion }
		},
	},
	{
		name: 'formula-point-mode-workflow',
		category: 'edit',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			workflow: 'formula-entry,arrow-reference,commit',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await dispatchText(engine, '2')
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			engine.render(size)
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'goto B1' })
			await dispatchText(engine, '=')
			await engine.dispatch({ kind: 'key', key: 'ArrowLeft' })
			await dispatchText(engine, '*2')
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, mode: engine.state().mode }
		},
	},
	{
		name: 'command-palette-search',
		category: 'command',
		dimensions: { rows: DEFAULT_SIZE.rows, cols: DEFAULT_SIZE.cols, queryChar: 'f' },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: ':' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'key', key: 'text', text: 'f' })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, mode: engine.state().mode }
		},
	},
	{
		name: 'terminal-calibration-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			command: 'calibrate',
		},
		async setup(size) {
			return WorkbookTuiEngine.create({ size })
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'calibrate' })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, focused: engine.state().workspace.focusedRegion }
		},
	},
	{
		name: 'terminal-input-stream-parser',
		category: 'command',
		dimensions: {
			chunks: 5,
			iterations: 1000,
			features: 'split-csi,bracketed-paste,text,control',
		},
		async setup() {
			return [
				Buffer.from('\x1b['),
				Buffer.from('C\x1b[C'),
				Buffer.from('\rabc'),
				Buffer.from('\x1b[200~1\t2\n'),
				Buffer.from('3\t4\x1b[201~'),
			]
		},
		async run(_size, context) {
			if (!Array.isArray(context)) throw new Error('input parser benchmark missing chunks')
			let events = 0
			for (let i = 0; i < 1000; i++) {
				const parser = new TerminalInputParser()
				for (const chunk of context) {
					if (!(chunk instanceof Buffer)) throw new Error('input parser chunk is not a Buffer')
					events += parser.push(chunk).length
				}
				events += parser.flush().length
			}
			return { events }
		},
	},
	{
		name: 'dialog-command-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			dialogs: 'format,filter,validation',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'command', command: 'goto A1' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({
				kind: 'command',
				command: 'format {"bold":true,"numberFormat":"0.00"}',
			})
			await engine.dispatch({ kind: 'command', command: 'filter {"range":"A1:C100"}' })
			await engine.dispatch({
				kind: 'command',
				command:
					'validate {"range":"A1:A100","rule":{"type":"whole","operator":"between","formula1":"1","formula2":"100"}}',
			})
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'table-comment-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			commands: 'create-table,comment',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: 'Name\tValue\nA\t1\nB\t2' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({
				kind: 'command',
				command: 'table create {"ref":"A1:B3","name":"Revenue","hasHeaders":true}',
			})
			await engine.dispatch({ kind: 'command', command: 'goto B2' })
			await engine.dispatch({
				kind: 'command',
				command: 'comment {"text":"Review this","author":"Ascend"}',
			})
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'object-dialog-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			dialogs: 'chart-wizard,pivot-fields',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'chart' })
			const chartFrame = engine.render(size)
			await engine.dispatch({ kind: 'key', key: 'Escape' })
			await engine.dispatch({ kind: 'command', command: 'pivot' })
			const pivotFrame = engine.render(size)
			await engine.dispatch({ kind: 'key', key: 'Escape' })
			return { chartBytes: chartFrame.stats.bytes, pivotBytes: pivotFrame.stats.bytes }
		},
	},
	{
		name: 'object-inspector-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			commands: 'objects,escape',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'objects' })
			const inspectorFrame = engine.render(size)
			await engine.dispatch({ kind: 'key', key: 'Escape' })
			return { bytes: inspectorFrame.stats.bytes, focused: engine.state().workspace.focusedRegion }
		},
	},
	{
		name: 'print-preview-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			commands: 'set-print-area,set-page-setup',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({
				kind: 'command',
				command:
					'print {"range":"A1:D20","orientation":"landscape","fitToWidth":1,"fitToHeight":0}',
			})
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'file-save-export-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			commands: 'save-as,export-csv,export-json',
		},
		async setup(size) {
			const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-bench-'))
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: 'Name\tValue\nAlpha\t7\nBeta\t9' })
			return {
				engine,
				tempDir: dir,
				xlsxPath: join(dir, 'book.xlsx'),
				csvPath: join(dir, 'book.csv'),
				jsonPath: join(dir, 'book.json'),
			}
		},
		async teardown(context) {
			const { tempDir } = requireFileWorkflowContext(context)
			await rm(tempDir, { recursive: true, force: true })
		},
		async run(size, context) {
			const { engine, xlsxPath, csvPath, jsonPath } = requireFileWorkflowContext(context)
			await engine.dispatch({ kind: 'command', command: `save-as ${xlsxPath}` })
			await engine.dispatch({ kind: 'command', command: `export ${csvPath}` })
			await engine.dispatch({
				kind: 'command',
				command: `export ${JSON.stringify({ path: jsonPath, format: 'json' })}`,
			})
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'find-replace-workflow',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			cells: 1000,
			commands: 'find,replace-all',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: buildTsv(100, 10) })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({
				kind: 'command',
				command:
					'find {"range":"A1:J100","findText":"42","action":"find","lookIn":"values","matchEntireCell":true}',
			})
			await engine.dispatch({
				kind: 'command',
				command:
					'replace {"range":"A1:J100","findText":"99","replaceText":"100","action":"replaceAll","lookIn":"values","matchEntireCell":true}',
			})
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'dialog-form-apply',
		category: 'command',
		dimensions: {
			rows: DEFAULT_SIZE.rows,
			cols: DEFAULT_SIZE.cols,
			dialog: 'format-cells',
			fields: 'numberFormat,bold',
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'command', command: 'goto A1' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'command', command: 'format' })
			await dispatchText(engine, '0.00')
			await engine.dispatch({ kind: 'key', key: 'Tab' })
			await engine.dispatch({ kind: 'key', key: 'text', text: ' ' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'paste-10k-cells',
		category: 'edit',
		dimensions: { rows: 100, cols: 100, cells: 10_000 },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			return {
				engine,
				text: buildTsv(100, 100),
			}
		},
		async run(size, context) {
			const { engine, text } = requirePasteContext(context)
			await engine.dispatch({ kind: 'command', command: 'goto A1' })
			await engine.dispatch({ kind: 'key', key: 'text', text })
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, dirty: engine.state().dirty }
		},
	},
	{
		name: 'resize-frame',
		category: 'resize',
		dimensions: {
			fromRows: DEFAULT_SIZE.rows,
			fromCols: DEFAULT_SIZE.cols,
			toRows: 48,
			toCols: 160,
		},
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			return engine
		},
		async run(_size, context) {
			const engine = requireEngine(context)
			await engine.dispatch({ kind: 'resize', size: { rows: 48, cols: 160 } })
			const frame = engine.render({ rows: 48, cols: 160 })
			return { rows: frame.size.rows, cols: frame.size.cols, bytes: frame.stats.bytes }
		},
	},
	{
		name: 'renderer-bakeoff-ansi-baseline',
		category: 'render',
		dimensions: { rows: DEFAULT_SIZE.rows, cols: DEFAULT_SIZE.cols, frames: 40 },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			const frames = []
			for (let i = 0; i < 40; i++) {
				await engine.dispatch({ kind: 'key', key: i % 2 === 0 ? 'ArrowRight' : 'ArrowLeft' })
				frames.push(engine.render(size))
			}
			return frames
		},
		async run(_size, context) {
			if (!Array.isArray(context)) throw new Error('renderer bakeoff missing frames')
			const results = await runRendererBakeoff(context, ['ansi'])
			const ansi = results.find((result) => result.candidate === 'ansi')
			const opentui = results.find((result) => result.candidate === 'opentui')
			return {
				ansiStatus: ansi?.status ?? null,
				ansiP95Ms: ansi?.p95Ms ?? null,
				opentuiStatus: opentui?.status ?? null,
				bytesOut: ansi?.bytesOut ?? null,
			}
		},
	},
	{
		name: 'renderer-bakeoff-opentui-line-adapter',
		category: 'render',
		dimensions: { rows: DEFAULT_SIZE.rows, cols: DEFAULT_SIZE.cols, frames: 40 },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			const frames = []
			for (let i = 0; i < 40; i++) {
				await engine.dispatch({ kind: 'key', key: i % 2 === 0 ? 'ArrowRight' : 'ArrowLeft' })
				frames.push(engine.render(size))
			}
			return frames
		},
		async run(_size, context) {
			if (!Array.isArray(context)) throw new Error('renderer bakeoff missing frames')
			const results = await runRendererBakeoff(context, ['opentui'])
			const opentui = results.find((result) => result.candidate === 'opentui')
			return {
				opentuiStatus: opentui?.status ?? null,
				opentuiP95Ms: opentui?.p95Ms ?? null,
				bytesOut: opentui?.bytesOut ?? null,
			}
		},
	},
	{
		name: 'metadata-grid-paint-1m-x-20',
		category: 'paint',
		dimensions: { workbookRows: 1_000_000, workbookCols: 20, paintedRows: 32, paintedCols: 120 },
		async setup(size) {
			const engine = await WorkbookTuiEngine.create({ size })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'command', command: 'goto T1000000' })
			return engine
		},
		async run(size, context) {
			const engine = requireEngine(context)
			const frame = engine.render(size)
			return { bytes: frame.stats.bytes, sheet: engine.state().sheetName }
		},
	},
]

async function runScenario(
	scenario: TuiScenario,
	repeat: number,
	warmup: number,
	size: TerminalSize,
): Promise<BenchmarkCaseResult> {
	for (let i = 0; i < warmup; i++) await runScenarioSample(scenario, size)
	const samples: Array<{ durationMs: number; throughputPerSec?: number }> = []
	let assertions: Record<string, string | number | boolean | null> | undefined
	for (let i = 0; i < repeat; i++) {
		const context = scenario.setup ? await scenario.setup(size) : undefined
		const start = performance.now()
		try {
			assertions = await scenario.run(size, context)
			const durationMs = performance.now() - start
			samples.push({ durationMs })
		} finally {
			if (scenario.teardown) await scenario.teardown(context)
		}
	}
	return {
		name: scenario.name,
		category: scenario.category,
		dimensions: scenario.dimensions,
		metrics: summarizeSamples(samples),
		samples,
		...(assertions ? { assertions } : {}),
	}
}

async function runScenarioSample(
	scenario: TuiScenario,
	size: TerminalSize,
): Promise<Record<string, string | number | boolean | null>> {
	const context = scenario.setup ? await scenario.setup(size) : undefined
	try {
		return await scenario.run(size, context)
	} finally {
		if (scenario.teardown) await scenario.teardown(context)
	}
}

async function dispatchText(engine: WorkbookTuiEngine, text: string): Promise<void> {
	for (const char of text) {
		await engine.dispatch({ kind: 'key', key: 'text', text: char })
	}
}

function requireEngine(context: unknown): WorkbookTuiEngine {
	if (!(context instanceof WorkbookTuiEngine))
		throw new Error('TUI benchmark scenario missing engine')
	return context
}

function requirePasteContext(context: unknown): { engine: WorkbookTuiEngine; text: string } {
	if (
		typeof context !== 'object' ||
		context === null ||
		!('engine' in context) ||
		!('text' in context)
	) {
		throw new Error('TUI paste benchmark scenario missing context')
	}
	const candidate = context as { engine: unknown; text: unknown }
	if (!(candidate.engine instanceof WorkbookTuiEngine) || typeof candidate.text !== 'string') {
		throw new Error('TUI paste benchmark scenario has invalid context')
	}
	return candidate
}

function requireFileWorkflowContext(context: unknown): {
	engine: WorkbookTuiEngine
	tempDir: string
	xlsxPath: string
	csvPath: string
	jsonPath: string
} {
	if (
		typeof context !== 'object' ||
		context === null ||
		!('engine' in context) ||
		!('tempDir' in context) ||
		!('xlsxPath' in context) ||
		!('csvPath' in context) ||
		!('jsonPath' in context)
	) {
		throw new Error('TUI file workflow benchmark scenario missing context')
	}
	const candidate = context as {
		engine: unknown
		tempDir: unknown
		xlsxPath: unknown
		csvPath: unknown
		jsonPath: unknown
	}
	if (
		!(candidate.engine instanceof WorkbookTuiEngine) ||
		typeof candidate.tempDir !== 'string' ||
		typeof candidate.xlsxPath !== 'string' ||
		typeof candidate.csvPath !== 'string' ||
		typeof candidate.jsonPath !== 'string'
	) {
		throw new Error('TUI file workflow benchmark scenario has invalid context')
	}
	return candidate
}

function buildTsv(rows: number, cols: number): string {
	const lines: string[] = []
	for (let row = 0; row < rows; row++) {
		const cells: string[] = []
		for (let col = 0; col < cols; col++) cells.push(String(row * cols + col))
		lines.push(cells.join('\t'))
	}
	return `${lines.join('\n')}\n`
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index < 0) return undefined
	const value = process.argv[index + 1]
	return value && !value.startsWith('--') ? value : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function renderSummary(results: readonly BenchmarkCaseResult[]): string {
	return results
		.map((result) => {
			const metrics = result.metrics
			return `${result.name.padEnd(30)} median=${metrics.medianMs.toFixed(2)}ms p95=${metrics.p95Ms.toFixed(2)}ms`
		})
		.join('\n')
}

async function main(): Promise<void> {
	const json = hasFlag('--json')
	const checkTargets = hasFlag('--check-targets')
	const scenarioName = readFlag('--scenario')
	const repeat = Math.max(1, Number.parseInt(readFlag('--repeat') ?? '20', 10) || 20)
	const warmup = Math.max(0, Number.parseInt(readFlag('--warmup') ?? '5', 10) || 5)
	const rows = Math.max(8, Number.parseInt(readFlag('--rows') ?? String(DEFAULT_SIZE.rows), 10))
	const cols = Math.max(40, Number.parseInt(readFlag('--cols') ?? String(DEFAULT_SIZE.cols), 10))
	const size = { rows, cols }
	const selected = scenarioName
		? scenarios.filter((scenario) => scenario.name === scenarioName)
		: scenarios
	if (selected.length === 0) throw new Error(`Unknown TUI benchmark scenario "${scenarioName}"`)
	const tempDir = await mkdtemp(join(tmpdir(), 'ascend-tui-bench-'))
	try {
		await buildSmokeWorkbook(join(tempDir, 'smoke.xlsx'))
		const results: BenchmarkCaseResult[] = []
		for (const scenario of selected) results.push(await runScenario(scenario, repeat, warmup, size))
		const suite = createBenchmarkSuite({
			suite: 'ascend-tui-benchmarks',
			kind: 'synthetic',
			cases: results,
			metadata: { repeat, warmup, rows, cols },
		})
		const targetResults = checkTargets
			? checkTuiTargets(suite, { skipMissing: Boolean(scenarioName) })
			: []
		const targetFailures = targetResults.filter((result) => !result.passed)
		if (json) {
			console.log(
				JSON.stringify(
					checkTargets ? { ...suite, tuiTargetResults: targetResults } : suite,
					null,
					2,
				),
			)
			if (targetFailures.length > 0) process.exitCode = 1
			return
		}
		console.log('Ascend TUI benchmark summary')
		console.log(renderSummary(results))
		if (checkTargets) {
			console.log(formatTuiTargetResults(targetResults))
			if (targetFailures.length > 0) process.exitCode = 1
		}
	} finally {
		await rm(tempDir, { recursive: true, force: true })
	}
}

async function buildSmokeWorkbook(path: string): Promise<void> {
	const workbook = AscendWorkbook.create()
	workbook.applyAndRecalc([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
	await workbook.save(path)
}

await main()
