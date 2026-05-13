#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { indexToColumn, WorkbookDocument } from '../../packages/sdk/src/index.ts'
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

interface Sample {
	readonly directFullOpenMs: number
	readonly directFullWindowMs: number
	readonly directFullTotalMs: number
	readonly directFullHydratedCells: number
	readonly previewOpenMs: number
	readonly previewWindowMs: number
	readonly previewFirstWindowMs: number
	readonly previewHydratedCells: number
	readonly previewPartial: boolean
	readonly promoteMs: number
	readonly promotedHydratedCells: number
	readonly promotedPartial: boolean
	readonly previewThenPromoteTotalMs: number
	readonly firstWindowCells: number
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
	'sparse-wide',
])

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

async function timed<T>(
	fn: () => Promise<T> | T,
): Promise<{ readonly ms: number; readonly value: T }> {
	const start = performance.now()
	const value = await fn()
	return { ms: performance.now() - start, value }
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

async function runDirectFull(path: string, range: string, previewRows: number) {
	WorkbookDocument.clearCache()
	const opened = await timed(() => WorkbookDocument.open(path, { mode: 'full' }))
	const window = await timed(() =>
		opened.value.readWindowCompact(opened.value.sheets[0] ?? 'Data', range, {
			rowLimit: previewRows,
			includeRefs: false,
			omitEmpty: true,
			flatValues: true,
		}),
	)
	const info = opened.value.inspect()
	return {
		openMs: opened.ms,
		windowMs: window.ms,
		totalMs: opened.ms + window.ms,
		hydratedCells: info.cellCount,
		firstWindowCells: window.value?.cells.length ?? 0,
	}
}

async function runPreviewThenPromote(path: string, range: string, previewRows: number) {
	WorkbookDocument.clearCache()
	const preview = await timed(() =>
		WorkbookDocument.open(path, { mode: 'values', maxRows: previewRows }),
	)
	const previewWindow = await timed(() =>
		preview.value.readWindowCompact(preview.value.sheets[0] ?? 'Data', range, {
			rowLimit: previewRows,
			includeRefs: false,
			omitEmpty: true,
			flatValues: true,
		}),
	)
	const previewInfo = preview.value.inspect()
	const promoted = await timed(() => preview.value.withLoad({ mode: 'full' }))
	const promotedInfo = promoted.value.inspect()
	return {
		previewOpenMs: preview.ms,
		previewWindowMs: previewWindow.ms,
		previewFirstWindowMs: preview.ms + previewWindow.ms,
		previewHydratedCells: previewInfo.cellCount,
		previewPartial: previewInfo.load.isPartial,
		promoteMs: promoted.ms,
		promotedHydratedCells: promotedInfo.cellCount,
		promotedPartial: promotedInfo.load.isPartial,
		previewThenPromoteTotalMs: preview.ms + previewWindow.ms + promoted.ms,
		firstWindowCells: previewWindow.value?.cells.length ?? 0,
	}
}

async function runSample(path: string, range: string, previewRows: number): Promise<Sample> {
	const full = await runDirectFull(path, range, previewRows)
	runGc()
	const promoted = await runPreviewThenPromote(path, range, previewRows)
	runGc()
	return {
		directFullOpenMs: full.openMs,
		directFullWindowMs: full.windowMs,
		directFullTotalMs: full.totalMs,
		directFullHydratedCells: full.hydratedCells,
		...promoted,
		firstWindowCells: promoted.firstWindowCells,
	}
}

function summarize(samples: readonly Sample[]) {
	const directFullTotalMedianMs = median(samples.map((sample) => sample.directFullTotalMs))
	const previewFirstWindowMedianMs = median(samples.map((sample) => sample.previewFirstWindowMs))
	const promoteMedianMs = median(samples.map((sample) => sample.promoteMs))
	const previewThenPromoteTotalMedianMs = median(
		samples.map((sample) => sample.previewThenPromoteTotalMs),
	)
	return {
		directFullOpenMedianMs: median(samples.map((sample) => sample.directFullOpenMs)),
		directFullWindowMedianMs: median(samples.map((sample) => sample.directFullWindowMs)),
		directFullTotalMedianMs,
		previewOpenMedianMs: median(samples.map((sample) => sample.previewOpenMs)),
		previewWindowMedianMs: median(samples.map((sample) => sample.previewWindowMs)),
		previewFirstWindowMedianMs,
		promoteMedianMs,
		previewThenPromoteTotalMedianMs,
		firstWindowSpeedupVsFull: directFullTotalMedianMs / previewFirstWindowMedianMs,
		promotionOverheadVsDirectFull: previewThenPromoteTotalMedianMs / directFullTotalMedianMs,
		directFullHydratedCellsMedian: median(samples.map((sample) => sample.directFullHydratedCells)),
		previewHydratedCellsMedian: median(samples.map((sample) => sample.previewHydratedCells)),
		promotedHydratedCellsMedian: median(samples.map((sample) => sample.promotedHydratedCells)),
		firstWindowCellsMedian: median(samples.map((sample) => sample.firstWindowCells)),
		previewPartial: samples.every((sample) => sample.previewPartial),
		promotedPartial: samples.some((sample) => sample.promotedPartial),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const range = `A1:${indexToColumn(args.cols - 1)}${args.rows}`
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath, range, args.previewRows)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath, range, args.previewRows))
			runGc()
		}
		const payload = {
			tool: 'partial-promotion',
			args,
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
