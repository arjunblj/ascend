#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface OpenModeCase {
	readonly key: string
	readonly options: {
		readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
		readonly maxRows?: number
		readonly richMetadata?: boolean
	}
}

interface ModeSample {
	readonly openMs: number
	readonly inspectMs: number
	readonly totalMs: number
	readonly cellCount: number | null
	readonly sheetCount: number
	readonly loadedSheetCount: number
	readonly isPartial: boolean
	readonly cellsHydrated: boolean
	readonly richSheetMetadataHydrated: boolean
	readonly partialReasons: readonly string[]
}

interface Sample {
	readonly modes: Record<string, ModeSample>
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
])

const OPEN_MODES: readonly OpenModeCase[] = [
	{ key: 'metadataOnly', options: { mode: 'metadata-only' } },
	{ key: 'valuesCapped500', options: { mode: 'values', maxRows: 500 } },
	{ key: 'valuesFull', options: { mode: 'values' } },
	{ key: 'valuesFullRich', options: { mode: 'values', richMetadata: true } },
	{ key: 'formulaFull', options: { mode: 'formula' } },
	{ key: 'full', options: { mode: 'full' } },
	{ key: 'fullRich', options: { richMetadata: true } },
]

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

function ratio(numerator: number, denominator: number): number {
	return denominator === 0 ? 0 : numerator / denominator
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

async function measureOpenMode(xlsxPath: string, mode: OpenModeCase): Promise<ModeSample> {
	const openStart = performance.now()
	const workbook = await AscendWorkbook.open(xlsxPath, mode.options)
	const openMs = performance.now() - openStart
	const inspectStart = performance.now()
	const info = workbook.inspect()
	const inspectMs = performance.now() - inspectStart
	return {
		openMs,
		inspectMs,
		totalMs: openMs + inspectMs,
		cellCount: info.cellCount,
		sheetCount: info.sheetCount,
		loadedSheetCount: info.loadedSheetCount,
		isPartial: info.load.isPartial,
		cellsHydrated: info.load.cellsHydrated,
		richSheetMetadataHydrated: info.load.richSheetMetadataHydrated,
		partialReasons: info.load.partialReasons,
	}
}

async function runSample(xlsxPath: string): Promise<Sample> {
	const modes: Record<string, ModeSample> = {}
	for (const mode of OPEN_MODES) {
		runGc()
		modes[mode.key] = await measureOpenMode(xlsxPath, mode)
	}
	runGc()
	return { modes }
}

function summarize(samples: readonly Sample[]) {
	const modeSummaries = Object.fromEntries(
		OPEN_MODES.map((mode) => {
			const modeSamples = samples.map((sample) => sample.modes[mode.key])
			const openMedianMs = median(modeSamples.map((sample) => sample?.openMs ?? 0))
			const inspectMedianMs = median(modeSamples.map((sample) => sample?.inspectMs ?? 0))
			const totalMedianMs = median(modeSamples.map((sample) => sample?.totalMs ?? 0))
			return [
				mode.key,
				{
					openMedianMs,
					inspectMedianMs,
					totalMedianMs,
					cellCountMedian: median(modeSamples.map((sample) => sample?.cellCount ?? 0)),
					sheetCountMedian: median(modeSamples.map((sample) => sample?.sheetCount ?? 0)),
					loadedSheetCountMedian: median(
						modeSamples.map((sample) => sample?.loadedSheetCount ?? 0),
					),
					isPartial: modeSamples.every((sample) => sample?.isPartial === true),
					cellsHydrated: modeSamples.every((sample) => sample?.cellsHydrated === true),
					richSheetMetadataHydrated: modeSamples.every(
						(sample) => sample?.richSheetMetadataHydrated === true,
					),
					partialReasons: modeSamples.find((sample) => sample !== undefined)?.partialReasons ?? [],
				},
			]
		}),
	)
	const full = modeSummaries.full?.totalMedianMs ?? 0
	const fullRich = modeSummaries.fullRich?.totalMedianMs ?? 0
	const valuesFull = modeSummaries.valuesFull?.totalMedianMs ?? 0
	const valuesCapped = modeSummaries.valuesCapped500?.totalMedianMs ?? 0
	const valuesRich = modeSummaries.valuesFullRich?.totalMedianMs ?? 0
	return {
		modes: modeSummaries,
		ratios: {
			fullRichOverFull: ratio(fullRich, full),
			fullOverValuesFull: ratio(full, valuesFull),
			valuesFullRichOverValuesFull: ratio(valuesRich, valuesFull),
			fullOverValuesCapped500: ratio(full, valuesCapped),
			fullRichOverValuesCapped500: ratio(fullRich, valuesCapped),
		},
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath))
			runGc()
		}
		const payload = {
			tool: 'rich-reopen-breakdown',
			args,
			dimensions: {
				rows: data.rows,
				cols: data.cols,
				cells: data.cells,
				bytes: data.xlsxBytes.byteLength,
			},
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
