#!/usr/bin/env bun
import { readFile, rm } from 'node:fs/promises'
import type { Operation } from '@ascend/schema'
import {
	type AgentWorkflowProgressEvent,
	AscendWorkbook,
	auditPackageGraphRoundtrip,
	commitAgentPlan,
	indexToColumn,
} from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly updates: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface Sample {
	readonly commitMs: number
	readonly commitPostWriteMs: number
	readonly sourceOpenMs: number
	readonly sourceReadBytesMs: number
	readonly outputReadBytesMs: number
	readonly reopenOutputMs: number
	readonly checkMs: number
	readonly lintMs: number
	readonly preservationMs: number
	readonly outputPackageGraphMs: number
	readonly packageGraphAuditMs: number
	readonly breakdownTotalMs: number
	readonly outputBytes: number
	readonly checkIssues: number
	readonly lintWarnings: number
	readonly packageGraphIssues: number
	readonly valid: boolean
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
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
		updates: positiveInt(readOption(process.argv, '--updates'), 1_000),
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

function buildSetCellsOperation(count: number, rows: number, cols: number): Operation {
	const updates: Extract<Operation, { op: 'setCells' }>['updates'] = []
	for (let index = 0; index < count; index++) {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		updates.push({
			ref: `${indexToColumn(col)}${row + 1}`,
			value: 400_000 + index,
		})
	}
	return { op: 'setCells', sheet: 'Data', updates }
}

async function timedCommit(inputPath: string, outputPath: string, ops: readonly Operation[]) {
	let postWriteStarted: number | undefined
	let postWriteMs = 0
	const commit = await timed(() =>
		commitAgentPlan(inputPath, ops, {
			output: outputPath,
			approvals: [],
			onProgress: (event: AgentWorkflowProgressEvent) => {
				if (event.phase !== 'post-write') return
				if (event.status === 'started') {
					postWriteStarted = performance.now()
					return
				}
				if (postWriteStarted !== undefined) {
					postWriteMs = performance.now() - postWriteStarted
				}
			},
		}),
	)
	return { commitMs: commit.ms, postWriteMs, valid: commit.value.postWrite.valid }
}

async function verifyBreakdown(inputPath: string, outputPath: string) {
	const sourceOpen = await timed(() => AscendWorkbook.open(inputPath))
	const sourceGraph = sourceOpen.value.packageGraph()
	const sourceBytes = await timed(() => readFile(inputPath))
	const outputBytes = await timed(() => readFile(outputPath))
	const reopened = await timed(() => AscendWorkbook.open(outputPath, { richMetadata: true }))
	const check = await timed(() => reopened.value.check())
	const lint = await timed(() => reopened.value.lint())
	const preservation = await timed(() => reopened.value.writePlanSummary())
	const outputGraph = await timed(() => reopened.value.packageGraph())
	const audit = await timed(() =>
		auditPackageGraphRoundtrip(
			sourceGraph,
			new Uint8Array(
				sourceBytes.value.buffer,
				sourceBytes.value.byteOffset,
				sourceBytes.value.byteLength,
			),
			outputGraph.value,
			new Uint8Array(
				outputBytes.value.buffer,
				outputBytes.value.byteOffset,
				outputBytes.value.byteLength,
			),
		),
	)
	return {
		sourceOpenMs: sourceOpen.ms,
		sourceReadBytesMs: sourceBytes.ms,
		outputReadBytesMs: outputBytes.ms,
		reopenOutputMs: reopened.ms,
		checkMs: check.ms,
		lintMs: lint.ms,
		preservationMs: preservation.ms,
		outputPackageGraphMs: outputGraph.ms,
		packageGraphAuditMs: audit.ms,
		breakdownTotalMs:
			sourceOpen.ms +
			sourceBytes.ms +
			outputBytes.ms +
			reopened.ms +
			check.ms +
			lint.ms +
			preservation.ms +
			outputGraph.ms +
			audit.ms,
		outputBytes: outputBytes.value.byteLength,
		checkIssues: check.value.issues.length,
		lintWarnings: lint.value.warnings.length,
		packageGraphIssues: audit.value.issues.length,
	}
}

async function runSample(
	inputPath: string,
	outputPath: string,
	ops: readonly Operation[],
): Promise<Sample> {
	await rm(outputPath, { force: true })
	runGc()
	const commit = await timedCommit(inputPath, outputPath, ops)
	runGc()
	const breakdown = await verifyBreakdown(inputPath, outputPath)
	runGc()
	return {
		commitMs: commit.commitMs,
		commitPostWriteMs: commit.postWriteMs,
		...breakdown,
		valid: commit.valid && breakdown.checkIssues === 0 && breakdown.packageGraphIssues === 0,
	}
}

function summarize(samples: readonly Sample[]) {
	return {
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		commitPostWriteMedianMs: median(samples.map((sample) => sample.commitPostWriteMs)),
		sourceOpenMedianMs: median(samples.map((sample) => sample.sourceOpenMs)),
		sourceReadBytesMedianMs: median(samples.map((sample) => sample.sourceReadBytesMs)),
		outputReadBytesMedianMs: median(samples.map((sample) => sample.outputReadBytesMs)),
		reopenOutputMedianMs: median(samples.map((sample) => sample.reopenOutputMs)),
		checkMedianMs: median(samples.map((sample) => sample.checkMs)),
		lintMedianMs: median(samples.map((sample) => sample.lintMs)),
		preservationMedianMs: median(samples.map((sample) => sample.preservationMs)),
		outputPackageGraphMedianMs: median(samples.map((sample) => sample.outputPackageGraphMs)),
		packageGraphAuditMedianMs: median(samples.map((sample) => sample.packageGraphAuditMs)),
		breakdownTotalMedianMs: median(samples.map((sample) => sample.breakdownTotalMs)),
		outputBytesMedian: median(samples.map((sample) => sample.outputBytes)),
		checkIssuesMedian: median(samples.map((sample) => sample.checkIssues)),
		lintWarningsMedian: median(samples.map((sample) => sample.lintWarnings)),
		packageGraphIssuesMedian: median(samples.map((sample) => sample.packageGraphIssues)),
		valid: samples.every((sample) => sample.valid),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const outputPath = `${data.xlsxPath}.post-write-output.xlsx`
	const ops = [buildSetCellsOperation(args.updates, args.rows, args.cols)]
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath, outputPath, ops)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath, outputPath, ops))
			runGc()
		}
		const payload = {
			tool: 'post-write-breakdown',
			args,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(data.xlsxPath, { force: true })
		await rm(outputPath, { force: true })
	}
}

await run()
