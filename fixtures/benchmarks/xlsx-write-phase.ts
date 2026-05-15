#!/usr/bin/env bun
import { createWorkbook } from '../../packages/core/src/index.ts'
import {
	type WriteXlsxOptions,
	writeDenseRowsXlsxStreaming,
	writeXlsx,
	writeXlsxStreaming,
} from '../../packages/io-xlsx/src/index.ts'
import {
	buildWorkloadValues,
	denseWriteAssertions,
	expectedOrderedWorkloadValuesHash,
	expectedWorkloadValuesHash,
	setCoreCellGenerated,
	type WorkloadName,
	workloadValue,
} from './competitive-io.ts'
import { writeHeapSnapshotFromEnv } from './heap-snapshot-on-exit.ts'
import { UPSTREAM_PROFILES } from './upstream-profiles.ts'

type StringMode = 'runner-default' | 'inline' | 'plain' | 'shared'
type ValueSource = 'generated' | 'materialized' | 'cached-materialized'
type WriterPath = 'dense-streaming' | 'workbook-streaming' | 'workbook-buffered'

interface Args {
	readonly profile?: string
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly stringMode: StringMode
	readonly valueSource: ValueSource
	readonly gcBetweenSamples: boolean
	readonly streaming: boolean
	readonly validate: boolean
	readonly json: boolean
}

interface PhaseSample {
	readonly writerPath: WriterPath
	readonly buildMs: number
	readonly writeMs: number
	readonly validateMs?: number
	readonly totalMs: number
	readonly cellsPerSecond: number
	readonly writeNsPerCell: number
	readonly bytes: number
	readonly rssAfterBytes: number
	readonly heapUsedBytes: number
}

interface PhaseRunContext {
	readonly cachedValues?: readonly (readonly ReturnType<typeof workloadValue>[])[] | undefined
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'mixed-closedxml-10text-5number',
	'plain-text',
	'string-heavy',
	'sparse-wide',
	'styles-heavy',
	'formula-heavy',
	'table-heavy',
	'feature-rich',
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
	const argv = process.argv.slice(2)
	const profileName = readOption(argv, '--profile')
	const profile = profileName
		? UPSTREAM_PROFILES.find((entry) => entry.name === profileName)
		: undefined
	if (profileName && !profile) {
		throw new Error(
			`Unsupported --profile "${profileName}". Expected one of: ${UPSTREAM_PROFILES.map((entry) => entry.name).join(', ')}`,
		)
	}
	if (profile && profile.category !== 'write') {
		throw new Error(
			`--profile "${profile.name}" is a read profile; xlsx-write-phase only supports write profiles`,
		)
	}
	const workload = readOption(argv, '--workload') ?? profile?.workload ?? 'dense-values'
	if (!WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	const stringMode = readOption(argv, '--string-mode') ?? 'runner-default'
	if (
		stringMode !== 'runner-default' &&
		stringMode !== 'inline' &&
		stringMode !== 'plain' &&
		stringMode !== 'shared'
	) {
		throw new Error('--string-mode must be runner-default, inline, plain, or shared')
	}
	const valueSource = readOption(argv, '--value-source') ?? 'generated'
	if (
		valueSource !== 'generated' &&
		valueSource !== 'materialized' &&
		valueSource !== 'cached-materialized'
	) {
		throw new Error('--value-source must be generated, materialized, or cached-materialized')
	}
	return {
		...(profile ? { profile: profile.name } : {}),
		rows: positiveInt(readOption(argv, '--rows'), profile?.rows ?? 2000),
		cols: positiveInt(readOption(argv, '--cols'), profile?.cols ?? 20),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 1),
		stringMode,
		valueSource,
		gcBetweenSamples: hasFlag(argv, '--gc-between-samples'),
		streaming: hasFlag(argv, '--streaming'),
		validate: hasFlag(argv, '--validate') || readOption(argv, '--validate') === 'true',
		json: hasFlag(argv, '--json'),
	}
}

function writeOptions(args: Args): WriteXlsxOptions {
	if (args.stringMode === 'shared') return { useSharedStrings: true }
	if (args.stringMode === 'inline') return { useSharedStrings: false }
	if (args.stringMode === 'plain') {
		return { useSharedStrings: false, usePlainStrings: true, omitDenseCellRefs: true }
	}
	return {
		useSharedStrings: args.workload === 'feature-rich' ? undefined : false,
		usePlainStrings: args.workload === 'string-heavy' || args.workload === 'plain-text',
		omitDenseCellRefs: args.workload === 'string-heavy' || args.workload === 'plain-text',
	}
}

function shouldUseDenseRowsWriter(workload: WorkloadName): boolean {
	return (
		workload === 'dense-values' ||
		workload === 'mixed-10pct-text' ||
		workload === 'mixed-50pct-text' ||
		workload === 'mixed-closedxml-10text-5number' ||
		workload === 'plain-text' ||
		workload === 'string-heavy'
	)
}

function shouldUseXmlSafeGeneratedStrings(workload: WorkloadName): boolean {
	return (
		workload === 'mixed-10pct-text' ||
		workload === 'mixed-50pct-text' ||
		workload === 'mixed-closedxml-10text-5number' ||
		workload === 'plain-text' ||
		workload === 'string-heavy'
	)
}

function denseValueType(workload: WorkloadName): 'number' | 'string' | undefined {
	if (workload === 'dense-values') return 'number'
	if (workload === 'plain-text' || workload === 'string-heavy') return 'string'
	return undefined
}

function denseValueTypes(
	workload: WorkloadName,
	cols: number,
): readonly ('number' | 'string' | undefined)[] | undefined {
	if (workload === 'mixed-50pct-text' && cols % 2 === 0) {
		return Array.from({ length: cols }, (_, col) => (col % 2 === 0 ? 'string' : 'number'))
	}
	if (workload === 'mixed-10pct-text' && cols % 10 === 0) {
		return Array.from({ length: cols }, (_, col) => (col % 10 === 0 ? 'string' : 'number'))
	}
	if (workload === 'mixed-closedxml-10text-5number') {
		return Array.from({ length: cols }, (_, col) => (col < 10 ? 'string' : 'number'))
	}
	return undefined
}

function shouldUseDirectDenseStreaming(args: Args): boolean {
	return args.streaming && shouldUseDenseRowsWriter(args.workload)
}

function writerPath(args: Args, useDirectDenseStreaming: boolean): WriterPath {
	if (useDirectDenseStreaming) return 'dense-streaming'
	return args.streaming ? 'workbook-streaming' : 'workbook-buffered'
}

function generatedCellCount(args: Args): number {
	if (args.workload !== 'sparse-wide') return args.rows * args.cols
	let count = args.rows
	if (args.cols > 1) count += args.rows
	for (let row = 0; row < args.rows; row++) {
		for (let col = 1; col < args.cols - 1; col++) {
			if ((row * 31 + col * 17) % 97 === 0) count++
		}
	}
	return count
}

function memory() {
	const current = process.memoryUsage()
	const rss = typeof current.rss === 'function' ? current.rss() : current.rss
	return {
		rssAfterBytes: rss,
		heapUsedBytes: current.heapUsed,
	}
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

async function runSample(args: Args, context: PhaseRunContext = {}): Promise<PhaseSample> {
	const totalStart = performance.now()
	const buildStart = performance.now()
	const useDirectDenseStreaming = shouldUseDirectDenseStreaming(args)
	const materializedValues =
		useDirectDenseStreaming && args.valueSource === 'materialized'
			? buildWorkloadValues(args.workload, args.rows, args.cols)
			: context.cachedValues
	const workbook = useDirectDenseStreaming ? undefined : createWorkbook()
	if (workbook) setCoreCellGenerated(workbook, args.rows, args.cols, args.workload)
	const buildMs = performance.now() - buildStart

	const writeStart = performance.now()
	const result = useDirectDenseStreaming
		? await writeDenseRowsXlsxStreaming({
				rows: args.rows,
				cols: args.cols,
				omitCellRefs: true,
				omitRowRefs: args.workload !== 'sparse-wide',
				cacheRepeatedRows: args.workload === 'mixed-closedxml-10text-5number',
				constantRows: args.workload === 'mixed-closedxml-10text-5number',
				stringsAreXmlSafe: shouldUseXmlSafeGeneratedStrings(args.workload),
				valueType: denseValueType(args.workload),
				valueTypes: denseValueTypes(args.workload, args.cols),
				allCellsPresent: true,
				valueAt: materializedValues
					? (row, col) => materializedValues[row]?.[col] ?? null
					: (row, col) => workloadValue(args.workload, row, col, args.cols),
			})
		: await writeWorkbookBytes(workbook, args)
	const writeMs = performance.now() - writeStart
	if (!result.ok) throw new Error(result.error.message)

	let validateMs: number | undefined
	if (args.validate) {
		const validateStart = performance.now()
		const materializeValues = args.rows * args.cols <= 500_000
		const values = materializeValues ? buildWorkloadValues(args.workload, args.rows, args.cols) : []
		denseWriteAssertions(result.value, {
			workloadName: args.workload,
			readSource: 'ascend-writer',
			sourceMode: useDirectDenseStreaming ? 'generated-write' : undefined,
			rows: args.rows,
			cols: args.cols,
			cells: materializeValues
				? values.reduce((count, row) => count + row.filter((value) => value !== null).length, 0)
				: generatedCellCount(args),
			values,
			semanticCellValuesHash: useDirectDenseStreaming
				? expectedOrderedWorkloadValuesHash(args.workload, args.rows, args.cols)
				: expectedWorkloadValuesHash(args.workload, args.rows, args.cols),
			orderedSemanticCellValuesHash: useDirectDenseStreaming
				? expectedOrderedWorkloadValuesHash(args.workload, args.rows, args.cols)
				: undefined,
			xlsxPath: '',
			xlsxBytes: result.value,
		})
		validateMs = performance.now() - validateStart
	}

	const totalMs = performance.now() - totalStart
	const sampleMemory = memory()
	writeHeapSnapshotFromEnv()
	return {
		writerPath: writerPath(args, useDirectDenseStreaming),
		buildMs,
		writeMs,
		...(validateMs === undefined ? {} : { validateMs }),
		totalMs,
		cellsPerSecond: generatedCellCount(args) / (writeMs / 1000),
		writeNsPerCell: (writeMs * 1_000_000) / generatedCellCount(args),
		bytes: result.value.byteLength,
		...sampleMemory,
	}
}

async function writeWorkbookBytes(
	workbook: ReturnType<typeof createWorkbook> | undefined,
	args: Args,
) {
	if (!workbook) throw new Error('Workbook build was skipped for a non-dense writer path')
	return args.streaming
		? await writeXlsxStreaming(workbook, undefined, writeOptions(args))
		: writeXlsx(workbook, undefined, writeOptions(args))
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function summarize(samples: readonly PhaseSample[], args: Args) {
	const buildMedianMs = median(samples.map((sample) => sample.buildMs))
	const writeMedianMs = median(samples.map((sample) => sample.writeMs))
	const validateMedianMs = samples.some((sample) => sample.validateMs !== undefined)
		? median(samples.map((sample) => sample.validateMs ?? 0))
		: undefined
	const totalMedianMs = median(samples.map((sample) => sample.totalMs))
	const phaseDurations = [
		['build', buildMedianMs],
		['write', writeMedianMs],
		...(validateMedianMs === undefined ? [] : ([['validate', validateMedianMs]] as const)),
	] as const
	const dominantPhase = phaseDurations.reduce(
		(best, current) => (current[1] > best[1] ? current : best),
		phaseDurations[0] ?? ['write', writeMedianMs],
	)[0]
	return {
		buildMedianMs,
		writeMedianMs,
		...(validateMedianMs === undefined ? {} : { validateMedianMs }),
		totalMedianMs,
		dominantPhase,
		writerPath: samples[0]?.writerPath ?? 'workbook-buffered',
		valueSource: args.valueSource,
		generatedValueCostIncludedInWrite:
			shouldUseDirectDenseStreaming(args) && args.valueSource === 'generated',
		writeMeasurementGuardrail:
			shouldUseDirectDenseStreaming(args) && args.valueSource === 'generated'
				? 'writeMs includes generated valueAt() cost; compare with --value-source cached-materialized to isolate writer/ZIP cost'
				: 'writeMs excludes per-cell workload generation cost',
		cellsPerSecondMedian: median(samples.map((sample) => sample.cellsPerSecond)),
		writeNsPerCellMedian: median(samples.map((sample) => sample.writeNsPerCell)),
		bytesMedian: median(samples.map((sample) => sample.bytes)),
		peakRssBytes: Math.max(...samples.map((sample) => sample.rssAfterBytes)),
	}
}

const args = parseArgs()
const cacheBuildStart = performance.now()
const cachedValues =
	shouldUseDirectDenseStreaming(args) && args.valueSource === 'cached-materialized'
		? buildWorkloadValues(args.workload, args.rows, args.cols)
		: undefined
const cacheBuildMs = performance.now() - cacheBuildStart
const context: PhaseRunContext = cachedValues ? { cachedValues } : {}
for (let i = 0; i < args.warmup; i++) {
	await runSample(args, context)
	if (args.gcBetweenSamples) runGc()
}
const samples: PhaseSample[] = []
for (let i = 0; i < args.repeat; i++) {
	samples.push(await runSample(args, context))
	if (args.gcBetweenSamples) runGc()
}
const payload = {
	tool: 'xlsx-write-phase',
	args,
	...(cachedValues ? { cacheBuildMs, cachedValueRows: cachedValues.length } : {}),
	summary: summarize(samples, args),
	samples,
}

if (args.json) {
	console.log(JSON.stringify(payload, null, 2))
} else {
	console.log(payload.summary)
}
