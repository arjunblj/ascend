#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { WorkbookTuiEngine } from '../../apps/tui/src/runtime/engine.ts'
import type { TelemetrySample, TerminalSize } from '../../apps/tui/src/runtime/types.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly previewRows: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface PaintSample {
	readonly fullOpenMs: number
	readonly fullRenderMs: number
	readonly fullTotalMs: number
	readonly fullHydrateMs: number | null
	readonly fullLayoutMs: number | null
	readonly fullFrameBytes: number
	readonly fullFrameCells: number
	readonly fullHydratedCells: number | null
	readonly fullPartial: boolean
	readonly previewOpenMs: number
	readonly previewRenderMs: number
	readonly previewTotalMs: number
	readonly previewHydrateMs: number | null
	readonly previewLayoutMs: number | null
	readonly previewFrameBytes: number
	readonly previewFrameCells: number
	readonly previewHydratedCells: number | null
	readonly previewPartial: boolean
	readonly previewReadOnly: boolean
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
	'sparse-wide',
])

const DEFAULT_SIZE: TerminalSize = { rows: 32, cols: 120 }

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function parseArgs(): Args {
	const workload = readOption(process.argv, '--workload') ?? 'mixed-10pct-text'
	if (!WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		previewRows: positiveInt(readOption(process.argv, '--preview-rows'), 500),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(process.argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(process.argv, '--warmup'), 1),
		json: hasFlag(process.argv, '--json'),
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function medianOptional(values: readonly (number | null | undefined)[]): number | undefined {
	const defined = values.filter((value): value is number => typeof value === 'number')
	return defined.length > 0 ? median(defined) : undefined
}

function latestLayoutTelemetry(samples: readonly TelemetrySample[]): TelemetrySample | undefined {
	for (let index = samples.length - 1; index >= 0; index--) {
		const sample = samples[index]
		if (sample?.layoutMs !== undefined || sample?.hydrateMs !== undefined) return sample
	}
	return undefined
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

async function measureFull(path: string, size: TerminalSize) {
	const openStart = performance.now()
	const engine = await WorkbookTuiEngine.create({ path, size })
	const openMs = performance.now() - openStart
	const renderStart = performance.now()
	const frame = engine.render(size)
	const renderMs = performance.now() - renderStart
	const state = engine.state()
	const telemetry = latestLayoutTelemetry(state.telemetry)
	const document = state.workspace.documents[0]
	return {
		openMs,
		renderMs,
		totalMs: openMs + renderMs,
		hydrateMs: telemetry?.hydrateMs ?? null,
		layoutMs: telemetry?.layoutMs ?? null,
		frameBytes: frame.stats.bytes,
		frameCells: frame.stats.fullFrameCells,
		hydratedCells: document?.info?.cellCount ?? null,
		partial: document?.info?.load.isPartial ?? false,
	}
}

async function measurePreview(path: string, size: TerminalSize, previewRows: number) {
	const openStart = performance.now()
	const engine = await WorkbookTuiEngine.create({
		path,
		loadOptions: { mode: 'values', maxRows: previewRows },
		size,
	})
	const openMs = performance.now() - openStart
	const renderStart = performance.now()
	const frame = engine.render(size)
	const renderMs = performance.now() - renderStart
	const state = engine.state()
	const telemetry = latestLayoutTelemetry(state.telemetry)
	const document = state.workspace.documents[0]
	return {
		openMs,
		renderMs,
		totalMs: openMs + renderMs,
		hydrateMs: telemetry?.hydrateMs ?? null,
		layoutMs: telemetry?.layoutMs ?? null,
		frameBytes: frame.stats.bytes,
		frameCells: frame.stats.fullFrameCells,
		hydratedCells: document?.info?.cellCount ?? null,
		partial: document?.info?.load.isPartial ?? false,
		readOnly: document?.readOnly ?? false,
	}
}

async function runSample(
	path: string,
	size: TerminalSize,
	previewRows: number,
): Promise<PaintSample> {
	const full = await measureFull(path, size)
	runGc()
	const preview = await measurePreview(path, size, previewRows)
	runGc()
	return {
		fullOpenMs: full.openMs,
		fullRenderMs: full.renderMs,
		fullTotalMs: full.totalMs,
		fullHydrateMs: full.hydrateMs,
		fullLayoutMs: full.layoutMs,
		fullFrameBytes: full.frameBytes,
		fullFrameCells: full.frameCells,
		fullHydratedCells: full.hydratedCells,
		fullPartial: full.partial,
		previewOpenMs: preview.openMs,
		previewRenderMs: preview.renderMs,
		previewTotalMs: preview.totalMs,
		previewHydrateMs: preview.hydrateMs,
		previewLayoutMs: preview.layoutMs,
		previewFrameBytes: preview.frameBytes,
		previewFrameCells: preview.frameCells,
		previewHydratedCells: preview.hydratedCells,
		previewPartial: preview.partial,
		previewReadOnly: preview.readOnly,
	}
}

function summarize(samples: readonly PaintSample[]) {
	const fullTotalMedianMs = median(samples.map((sample) => sample.fullTotalMs))
	const previewTotalMedianMs = median(samples.map((sample) => sample.previewTotalMs))
	const fullOpenMedianMs = median(samples.map((sample) => sample.fullOpenMs))
	const previewOpenMedianMs = median(samples.map((sample) => sample.previewOpenMs))
	return {
		fullTotalMedianMs,
		previewTotalMedianMs,
		speedupVsFull: fullTotalMedianMs / previewTotalMedianMs,
		fullOpenMedianMs,
		previewOpenMedianMs,
		openSpeedupVsFull: fullOpenMedianMs / previewOpenMedianMs,
		fullRenderMedianMs: median(samples.map((sample) => sample.fullRenderMs)),
		previewRenderMedianMs: median(samples.map((sample) => sample.previewRenderMs)),
		fullHydrateMedianMs: medianOptional(samples.map((sample) => sample.fullHydrateMs)),
		previewHydrateMedianMs: medianOptional(samples.map((sample) => sample.previewHydrateMs)),
		fullLayoutMedianMs: medianOptional(samples.map((sample) => sample.fullLayoutMs)),
		previewLayoutMedianMs: medianOptional(samples.map((sample) => sample.previewLayoutMs)),
		fullFrameBytesMedian: median(samples.map((sample) => sample.fullFrameBytes)),
		previewFrameBytesMedian: median(samples.map((sample) => sample.previewFrameBytes)),
		fullFrameCellsMedian: median(samples.map((sample) => sample.fullFrameCells)),
		previewFrameCellsMedian: median(samples.map((sample) => sample.previewFrameCells)),
		fullHydratedCellsMedian: medianOptional(samples.map((sample) => sample.fullHydratedCells)),
		previewHydratedCellsMedian: medianOptional(
			samples.map((sample) => sample.previewHydratedCells),
		),
		previewPartial: samples.every((sample) => sample.previewPartial),
		previewReadOnly: samples.every((sample) => sample.previewReadOnly),
		fullPartial: samples.some((sample) => sample.fullPartial),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const samples: PaintSample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath, DEFAULT_SIZE, args.previewRows)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath, DEFAULT_SIZE, args.previewRows))
			runGc()
		}
		const payload = {
			tool: 'tui-first-paint',
			args,
			size: DEFAULT_SIZE,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(data.xlsxPath, { force: true })
	}
}

await run()
