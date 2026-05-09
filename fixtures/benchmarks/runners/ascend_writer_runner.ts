#!/usr/bin/env bun
import { createWorkbook } from '../../../packages/core/src/index.ts'
import { writeDenseRowsXlsxStreaming, writeXlsx } from '../../../packages/io-xlsx/src/index.ts'
import {
	buildWorkloadValues,
	denseWriteAssertions,
	expectedOrderedWorkloadValuesHash,
	expectedWorkloadValuesHash,
	setCoreCellGenerated,
	type WorkloadName,
	workloadValue,
} from '../competitive-io.ts'

interface Args {
	readonly operation: 'write'
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly validationMode: 'final'
	readonly json: boolean
}

const SUPPORTED_WORKLOADS = new Set<string>([
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
	const args = process.argv.slice(2)
	const operation = readOption(args, '--operation')
	const workload = readOption(args, '--workload') ?? 'dense-values'
	if (operation !== 'write') throw new Error('--operation must be write')
	if (!SUPPORTED_WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	return {
		operation,
		rows: positiveInt(readOption(args, '--rows'), 2000),
		cols: positiveInt(readOption(args, '--cols'), 20),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(args, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(args, '--warmup'), 0),
		validationMode: 'final',
		json: hasFlag(args, '--json'),
	}
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

function memorySample(durationMs: number): {
	readonly durationMs: number
	readonly rssAfterBytes: number
	readonly rssAfterGcBytes: number
	readonly peakRssBytes: number
	readonly heapUsedBytes: number
	readonly heapTotalBytes: number
	readonly heapAfterGcBytes: number
} {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	runGc()
	const afterGc = process.memoryUsage()
	const rssAfterGc = typeof afterGc.rss === 'function' ? afterGc.rss() : afterGc.rss
	return {
		durationMs,
		rssAfterBytes: rss,
		rssAfterGcBytes: rssAfterGc,
		peakRssBytes: Math.max(rss, rssAfterGc),
		heapUsedBytes: memory.heapUsed,
		heapTotalBytes: memory.heapTotal,
		heapAfterGcBytes: afterGc.heapUsed,
	}
}

function generatedCellCount(args: Args): number {
	if (args.workload !== 'sparse-wide') return args.rows * args.cols
	if (args.cols <= 0) return 0
	let count = args.rows
	if (args.cols > 1) count += args.rows
	for (let row = 0; row < args.rows; row++) {
		for (let col = 1; col < args.cols - 1; col++) {
			if ((row * 31 + col * 17) % 97 === 0) count++
		}
	}
	return count
}

function shouldUsePlainStrings(workload: WorkloadName): boolean {
	return workload === 'string-heavy' || workload === 'plain-text'
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

function canUseGeneratedHashFromHarness(workload: WorkloadName): boolean {
	return shouldUseDenseRowsWriter(workload) || workload === 'sparse-wide'
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

async function writeWorkbook(args: Args): Promise<Uint8Array> {
	if (shouldUseDenseRowsWriter(args.workload)) {
		const options = {
			rows: args.rows,
			cols: args.cols,
			omitCellRefs: true,
			cacheRepeatedRows: args.workload === 'mixed-closedxml-10text-5number',
			constantRows: args.workload === 'mixed-closedxml-10text-5number',
			stringsAreXmlSafe: shouldUseXmlSafeGeneratedStrings(args.workload),
			valueType: denseValueType(args.workload),
			valueTypes: denseValueTypes(args.workload, args.cols),
			allCellsPresent: args.workload !== 'sparse-wide',
			valueAt: (row, col) => workloadValue(args.workload, row, col, args.cols),
		} as const
		const result = await writeDenseRowsXlsxStreaming(options)
		if (!result.ok) throw new Error(result.error.message)
		return result.value
	}
	const workbook = createWorkbook()
	setCoreCellGenerated(workbook, args.rows, args.cols, args.workload)
	const result = writeXlsx(workbook, undefined, {
		useSharedStrings: args.workload === 'feature-rich' ? undefined : false,
		usePlainStrings: shouldUsePlainStrings(args.workload),
		omitDenseCellRefs: shouldUsePlainStrings(args.workload),
	})
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

async function main(): Promise<void> {
	const args = parseArgs()
	for (let i = 0; i < args.warmup; i++) await writeWorkbook(args)
	const samples: ReturnType<typeof memorySample>[] = []
	let bytes: Uint8Array | undefined
	for (let i = 0; i < args.repeat; i++) {
		const start = performance.now()
		bytes = await writeWorkbook(args)
		samples.push(memorySample(performance.now() - start))
	}
	if (!bytes) throw new Error('No samples were produced')
	const shouldMaterializeExpectedValues = args.rows * args.cols <= 500_000
	const values = shouldMaterializeExpectedValues
		? buildWorkloadValues(args.workload, args.rows, args.cols)
		: []
	const shouldComputeExpectedHashes =
		shouldMaterializeExpectedValues || !canUseGeneratedHashFromHarness(args.workload)
	const input = {
		workloadName: args.workload,
		readSource: 'ascend-writer',
		sourceMode: 'generated-write',
		rows: args.rows,
		cols: args.cols,
		cells: shouldMaterializeExpectedValues
			? values.reduce((count, row) => count + row.filter((value) => value !== null).length, 0)
			: generatedCellCount(args),
		values,
		semanticCellValuesHash: shouldComputeExpectedHashes
			? expectedWorkloadValuesHash(args.workload, args.rows, args.cols)
			: '',
		orderedSemanticCellValuesHash: shouldComputeExpectedHashes
			? expectedOrderedWorkloadValuesHash(args.workload, args.rows, args.cols)
			: undefined,
		xlsxPath: '',
		xlsxBytes: bytes,
	} as const
	const payload = {
		assertions: {
			runnerVersion: 'workspace',
			...denseWriteAssertions(bytes, input),
			validationMode: args.validationMode,
			validationSamples: 1,
		},
		samples,
	}
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

await main()
