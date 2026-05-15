#!/usr/bin/env bun
import { readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import type { Operation } from '@ascend/schema'
import {
	type AgentWorkflowProgressEvent,
	AscendWorkbook,
	auditPackageGraphRoundtrip,
	commitAgentPlan,
	indexToColumn,
	WorkbookDocument,
} from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly updates: number
	readonly inputFile?: string
	readonly sheet?: string
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface BenchmarkInput {
	readonly xlsxPath: string
	readonly sheet: string
	readonly rows: number
	readonly cols: number
	readonly cleanup: boolean
	readonly source: 'generated' | 'input-file'
}

interface Sample {
	readonly commitMs: number
	readonly commitWritePolicySnapshotMs: number
	readonly commitPackageGraphMs: number
	readonly commitApprovalAuditMs: number
	readonly commitLossAuditMs: number
	readonly commitPackageGraphAuditMs: number
	readonly commitApplyMs: number
	readonly commitRecalcMs: number
	readonly commitWritePlanSummaryMs: number
	readonly commitWritePolicyCheckMs: number
	readonly commitWritePolicyBuildMs: number
	readonly commitToBytesMs: number
	readonly commitWriteMs: number
	readonly commitWriteFileMs: number
	readonly commitRenameMs: number
	readonly commitOutputByteReadMs: number
	readonly commitOutputHashMs: number
	readonly commitPostWriteMs: number
	readonly commitPostWriteReopenMs: number
	readonly commitPostWriteCheckMs: number
	readonly commitPostWriteLintMs: number
	readonly commitPostWritePreservationMs: number
	readonly commitPostWritePackageGraphMs: number
	readonly commitPostWritePackageGraphAuditMs: number
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
	const inputFile = readOption(process.argv, '--input-file')
	const sheet = readOption(process.argv, '--sheet')
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		updates: positiveInt(readOption(process.argv, '--updates'), 1_000),
		...(inputFile !== undefined ? { inputFile } : {}),
		...(sheet !== undefined ? { sheet } : {}),
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

function buildSetCellsOperation(
	count: number,
	rows: number,
	cols: number,
	sheet: string,
): Operation {
	const updates: Extract<Operation, { op: 'setCells' }>['updates'] = []
	for (let index = 0; index < count; index++) {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		updates.push({
			ref: `${indexToColumn(col)}${row + 1}`,
			value: 400_000 + index,
		})
	}
	return { op: 'setCells', sheet, updates }
}

async function timedCommit(inputPath: string, outputPath: string, ops: readonly Operation[]) {
	let writeStarted: number | undefined
	let writeMs = 0
	let postWriteStarted: number | undefined
	let postWriteMs = 0
	const postWriteTimings: Record<string, number> = {}
	const commit = await timed(() =>
		commitAgentPlan(inputPath, ops, {
			output: outputPath,
			approvals: [],
			onProgress: (event: AgentWorkflowProgressEvent) => {
				if (event.phase === 'write') {
					if (event.status === 'started') {
						writeStarted = performance.now()
						return
					}
					if (writeStarted !== undefined) writeMs = performance.now() - writeStarted
					return
				}
				if (event.phase === 'post-write' && event.status === 'started') {
					postWriteStarted = performance.now()
					return
				}
				if (event.phase === 'post-write' && postWriteStarted !== undefined) {
					postWriteMs = performance.now() - postWriteStarted
					return
				}
				if (!event.phase.startsWith('post-write:') || event.status === 'started') return
				const duration = durationMsFromDetails(event.details)
				if (duration !== undefined) {
					postWriteTimings[event.phase.slice('post-write:'.length)] = duration
				}
			},
		}),
	)
	const timings = commit.value.postWrite.timings
	return {
		commitMs: commit.ms,
		commitWritePolicySnapshotMs: commit.value.timings.writePolicySnapshotMs,
		commitPackageGraphMs: commit.value.timings.packageGraphMs,
		commitApprovalAuditMs: commit.value.timings.approvalAuditMs,
		commitLossAuditMs: commit.value.timings.lossAuditMs,
		commitPackageGraphAuditMs: commit.value.timings.packageGraphAuditMs,
		commitApplyMs: commit.value.timings.applyMs,
		commitRecalcMs: commit.value.timings.recalcMs,
		commitWritePlanSummaryMs: commit.value.timings.writePlanSummaryMs,
		commitWritePolicyCheckMs: commit.value.timings.writePolicyCheckMs,
		commitWritePolicyBuildMs: commit.value.timings.writePolicyBuildMs,
		commitToBytesMs: commit.value.timings.toBytesMs,
		writeMs,
		commitWriteFileMs: commit.value.timings.writeFileMs,
		commitRenameMs: commit.value.timings.renameMs,
		commitOutputByteReadMs: commit.value.timings.outputByteReadMs,
		commitOutputHashMs: commit.value.timings.outputHashMs,
		postWriteMs,
		reopenMs: postWriteTimings.reopen ?? timings?.reopenMs ?? 0,
		checkMs: postWriteTimings.check ?? timings?.checkMs ?? 0,
		lintMs: postWriteTimings.lint ?? timings?.lintMs ?? 0,
		preservationMs: postWriteTimings.preservation ?? timings?.preservationMs ?? 0,
		packageGraphMs: postWriteTimings['package-graph'] ?? timings?.packageGraphMs ?? 0,
		packageGraphAuditMs:
			postWriteTimings['package-graph-audit'] ?? timings?.packageGraphAuditMs ?? 0,
		valid: commit.value.postWrite.valid,
	}
}

function durationMsFromDetails(details: unknown): number | undefined {
	if (details === null || typeof details !== 'object') return undefined
	const duration = (details as { durationMs?: unknown }).durationMs
	return typeof duration === 'number' && Number.isFinite(duration) ? duration : undefined
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
		commitWritePolicySnapshotMs: commit.commitWritePolicySnapshotMs,
		commitPackageGraphMs: commit.commitPackageGraphMs,
		commitApprovalAuditMs: commit.commitApprovalAuditMs,
		commitLossAuditMs: commit.commitLossAuditMs,
		commitPackageGraphAuditMs: commit.commitPackageGraphAuditMs,
		commitApplyMs: commit.commitApplyMs,
		commitRecalcMs: commit.commitRecalcMs,
		commitWritePlanSummaryMs: commit.commitWritePlanSummaryMs,
		commitWritePolicyCheckMs: commit.commitWritePolicyCheckMs,
		commitWritePolicyBuildMs: commit.commitWritePolicyBuildMs,
		commitToBytesMs: commit.commitToBytesMs,
		commitWriteMs: commit.writeMs,
		commitWriteFileMs: commit.commitWriteFileMs,
		commitRenameMs: commit.commitRenameMs,
		commitOutputByteReadMs: commit.commitOutputByteReadMs,
		commitOutputHashMs: commit.commitOutputHashMs,
		commitPostWriteMs: commit.postWriteMs,
		commitPostWriteReopenMs: commit.reopenMs,
		commitPostWriteCheckMs: commit.checkMs,
		commitPostWriteLintMs: commit.lintMs,
		commitPostWritePreservationMs: commit.preservationMs,
		commitPostWritePackageGraphMs: commit.packageGraphMs,
		commitPostWritePackageGraphAuditMs: commit.packageGraphAuditMs,
		...breakdown,
		valid: commit.valid && breakdown.checkIssues === 0 && breakdown.packageGraphIssues === 0,
	}
}

function summarize(samples: readonly Sample[]) {
	return {
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		commitWritePolicySnapshotMedianMs: median(
			samples.map((sample) => sample.commitWritePolicySnapshotMs),
		),
		commitPackageGraphMedianMs: median(samples.map((sample) => sample.commitPackageGraphMs)),
		commitApprovalAuditMedianMs: median(samples.map((sample) => sample.commitApprovalAuditMs)),
		commitLossAuditMedianMs: median(samples.map((sample) => sample.commitLossAuditMs)),
		commitPackageGraphAuditMedianMs: median(
			samples.map((sample) => sample.commitPackageGraphAuditMs),
		),
		commitApplyMedianMs: median(samples.map((sample) => sample.commitApplyMs)),
		commitRecalcMedianMs: median(samples.map((sample) => sample.commitRecalcMs)),
		commitWritePlanSummaryMedianMs: median(
			samples.map((sample) => sample.commitWritePlanSummaryMs),
		),
		commitWritePolicyCheckMedianMs: median(
			samples.map((sample) => sample.commitWritePolicyCheckMs),
		),
		commitWritePolicyBuildMedianMs: median(
			samples.map((sample) => sample.commitWritePolicyBuildMs),
		),
		commitToBytesMedianMs: median(samples.map((sample) => sample.commitToBytesMs)),
		commitWriteMedianMs: median(samples.map((sample) => sample.commitWriteMs)),
		commitWriteFileMedianMs: median(samples.map((sample) => sample.commitWriteFileMs)),
		commitRenameMedianMs: median(samples.map((sample) => sample.commitRenameMs)),
		commitOutputByteReadMedianMs: median(samples.map((sample) => sample.commitOutputByteReadMs)),
		commitOutputHashMedianMs: median(samples.map((sample) => sample.commitOutputHashMs)),
		commitPostWriteMedianMs: median(samples.map((sample) => sample.commitPostWriteMs)),
		commitPostWriteReopenMedianMs: median(samples.map((sample) => sample.commitPostWriteReopenMs)),
		commitPostWriteCheckMedianMs: median(samples.map((sample) => sample.commitPostWriteCheckMs)),
		commitPostWriteLintMedianMs: median(samples.map((sample) => sample.commitPostWriteLintMs)),
		commitPostWritePreservationMedianMs: median(
			samples.map((sample) => sample.commitPostWritePreservationMs),
		),
		commitPostWritePackageGraphMedianMs: median(
			samples.map((sample) => sample.commitPostWritePackageGraphMs),
		),
		commitPostWritePackageGraphAuditMedianMs: median(
			samples.map((sample) => sample.commitPostWritePackageGraphAuditMs),
		),
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

async function inferLoadedColumnCount(
	path: string,
	sheetName: string,
	fallbackCols: number,
): Promise<number> {
	const preview = await WorkbookDocument.open(path, {
		mode: 'values',
		sheets: [sheetName],
		maxRows: 100,
	})
	const sheetInfo = preview.inspect().sheets.find((sheet) => sheet.name === sheetName)
	WorkbookDocument.clearCache()
	return Math.max(1, sheetInfo?.colCount ?? fallbackCols)
}

async function resolveBenchmarkInput(args: Args): Promise<BenchmarkInput> {
	if (args.inputFile === undefined) {
		const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
		return {
			xlsxPath: data.xlsxPath,
			sheet: args.sheet ?? 'Data',
			rows: args.rows,
			cols: args.cols,
			cleanup: true,
			source: 'generated',
		}
	}
	const document = await WorkbookDocument.open(args.inputFile, {
		mode: 'metadata-only',
		...(args.sheet !== undefined ? { sheets: [args.sheet] } : {}),
	})
	const info = document.inspect()
	const sheetInfo =
		args.sheet !== undefined
			? info.sheets.find((sheet) => sheet.name === args.sheet)
			: info.sheets[0]
	if (!sheetInfo) {
		throw new Error(
			args.sheet !== undefined
				? `Sheet "${args.sheet}" not found in ${args.inputFile}`
				: `No sheets found in ${args.inputFile}`,
		)
	}
	const rows = Math.max(1, sheetInfo.rowCount ?? args.rows)
	WorkbookDocument.clearCache()
	const cols = Math.max(
		1,
		sheetInfo.colCount ?? (await inferLoadedColumnCount(args.inputFile, sheetInfo.name, args.cols)),
	)
	return {
		xlsxPath: args.inputFile,
		sheet: sheetInfo.name,
		rows,
		cols,
		cleanup: false,
		source: 'input-file',
	}
}

function outputPathFor(input: BenchmarkInput): string {
	return join(
		tmpdir(),
		`ascend-post-write-${process.pid}-${Date.now()}-${basename(input.xlsxPath)}.out.xlsx`,
	)
}

async function run() {
	const args = parseArgs()
	const data = await resolveBenchmarkInput(args)
	const outputPath = outputPathFor(data)
	const ops = [buildSetCellsOperation(args.updates, data.rows, data.cols, data.sheet)]
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
			input: data,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		if (data.cleanup) await rm(data.xlsxPath, { force: true })
		await rm(outputPath, { force: true })
	}
}

await run()
