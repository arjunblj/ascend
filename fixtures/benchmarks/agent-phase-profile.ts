#!/usr/bin/env bun
import { rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import type { Operation } from '@ascend/schema'
import {
	type AgentCommitTimings,
	type AgentWorkflowProgressEvent,
	commitAgentPlan,
	createAgentPlan,
	createPreparedAgentPlan,
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
	readonly timeoutMs?: number
	readonly progress: boolean
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

interface TimedRun<T> {
	readonly ms: number
	readonly value: T
	readonly phases: readonly PhaseTiming[]
}

interface PhaseTiming {
	readonly kind: 'plan' | 'commit' | 'repair-plan'
	readonly phase: string
	readonly ms: number
	readonly status: AgentWorkflowProgressEvent['status']
	readonly count?: number
}

interface SampleStats {
	readonly sampleCount: number
	readonly min: number
	readonly median: number
	readonly mean: number
	readonly p95: number
	readonly max: number
	readonly stddev: number
	readonly cv: number
}

interface Sample {
	readonly planMs: number
	readonly commitMs: number
	readonly totalMs: number
	readonly sharedPlanMs: number
	readonly sharedCommitMs: number
	readonly sharedTotalMs: number
	readonly planUnattributedMs: number
	readonly commitUnattributedMs: number
	readonly sharedPlanUnattributedMs: number
	readonly sharedCommitUnattributedMs: number
	readonly planPayloadBytes: number
	readonly commitPayloadBytes: number
	readonly sharedPlanPayloadBytes: number
	readonly sharedCommitPayloadBytes: number
	readonly commitOutputBytes: number
	readonly sharedCommitOutputBytes: number
	readonly commitTimingMs: AgentCommitTimings
	readonly sharedCommitTimingMs: AgentCommitTimings
	readonly planPhaseMs: Record<string, number>
	readonly commitPhaseMs: Record<string, number>
	readonly sharedPlanPhaseMs: Record<string, number>
	readonly sharedCommitPhaseMs: Record<string, number>
	readonly operationCount: number
	readonly updateCount: number
	readonly changedCells: number
	readonly commitChangedCells: number
	readonly sharedChangedCells: number
	readonly sharedCommitChangedCells: number
	readonly postWriteValid: boolean
	readonly sharedPostWriteValid: boolean
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
	const timeoutMs = readOption(process.argv, '--timeout-ms')
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		updates: positiveInt(readOption(process.argv, '--updates'), 1_000),
		...(inputFile !== undefined ? { inputFile } : {}),
		...(sheet !== undefined ? { sheet } : {}),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(process.argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(process.argv, '--warmup'), 1),
		...(timeoutMs !== undefined ? { timeoutMs: positiveInt(timeoutMs, 300_000) } : {}),
		progress: hasFlag(process.argv, '--progress'),
		json: hasFlag(process.argv, '--json'),
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function percentileSorted(sorted: readonly number[], percentile: number): number {
	if (sorted.length === 0) return 0
	const index = Math.min(sorted.length - 1, Math.ceil(sorted.length * percentile) - 1)
	return sorted[index] ?? 0
}

function sampleStats(values: readonly number[]): SampleStats {
	const sorted = [...values].filter(Number.isFinite).sort((a, b) => a - b)
	if (sorted.length === 0) {
		return { sampleCount: 0, min: 0, median: 0, mean: 0, p95: 0, max: 0, stddev: 0, cv: 0 }
	}
	const mean = sorted.reduce((sum, value) => sum + value, 0) / sorted.length
	const variance =
		sorted.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(sorted.length - 1, 1)
	const stddev = Math.sqrt(variance)
	return {
		sampleCount: sorted.length,
		min: sorted[0] ?? 0,
		median: median(sorted),
		mean,
		p95: percentileSorted(sorted, 0.95),
		max: sorted[sorted.length - 1] ?? 0,
		stddev,
		cv: mean === 0 ? 0 : stddev / mean,
	}
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
			value: 200_000 + index,
		})
	}
	return { op: 'setCells', sheet, updates }
}

async function timedWorkflow<T>(
	fn: (onProgress: (event: AgentWorkflowProgressEvent) => void) => Promise<T>,
	onPhase?: (phase: PhaseTiming) => void,
): Promise<TimedRun<T>> {
	const started = new Map<string, number>()
	const phases: PhaseTiming[] = []
	const start = performance.now()
	const value = await fn((event) => {
		const key = `${event.kind}:${event.phase}`
		const now = performance.now()
		if (event.status === 'started') {
			started.set(key, now)
			return
		}
		const phaseStart = started.get(key)
		const phaseTiming = {
			kind: event.kind,
			phase: event.phase,
			ms: phaseStart === undefined ? 0 : now - phaseStart,
			status: event.status,
			...(event.count !== undefined ? { count: event.count } : {}),
		} satisfies PhaseTiming
		phases.push(phaseTiming)
		onPhase?.(phaseTiming)
		started.delete(key)
	})
	return { ms: performance.now() - start, value, phases }
}

function phaseMap(phases: readonly PhaseTiming[]): Record<string, number> {
	const map: Record<string, number> = {}
	for (const phase of phases) map[phase.phase] = phase.ms
	return map
}

function topLevelPhaseMs(phases: readonly PhaseTiming[]): number {
	return phases.reduce((sum, phase) => (phase.phase.includes(':') ? sum : sum + phase.ms), 0)
}

function unattributedMs(totalMs: number, phases: readonly PhaseTiming[]): number {
	return Math.max(0, totalMs - topLevelPhaseMs(phases))
}

function payloadBytes(value: unknown): number {
	return JSON.stringify(value).length
}

async function runSample(
	inputPath: string,
	outputPath: string,
	ops: readonly Operation[],
	updateCount: number,
	onPhase?: (phase: PhaseTiming) => void,
): Promise<Sample> {
	await rm(outputPath, { force: true })
	const sharedOutputPath = `${outputPath}.shared.xlsx`
	await rm(sharedOutputPath, { force: true })
	runGc()
	const plan = await timedWorkflow(
		(onProgress) => createAgentPlan(inputPath, ops, { onProgress }),
		onPhase,
	)
	const commit = await timedWorkflow(
		(onProgress) =>
			commitAgentPlan(inputPath, ops, { output: outputPath, approvals: [], onProgress }),
		onPhase,
	)
	const sharedPlan = await timedWorkflow(
		(onProgress) => createPreparedAgentPlan(inputPath, ops, { onProgress }),
		onPhase,
	)
	const outputBytes = (await stat(outputPath)).size
	let sharedCommit: TimedRun<Awaited<ReturnType<typeof sharedPlan.value.commit>>>
	let sharedOutputBytes = 0
	try {
		sharedCommit = await timedWorkflow(
			(onProgress) =>
				sharedPlan.value.commit({ output: sharedOutputPath, approvals: [], onProgress }),
			onPhase,
		)
		sharedOutputBytes = (await stat(sharedOutputPath)).size
	} finally {
		await rm(sharedOutputPath, { force: true })
	}
	return {
		planMs: plan.ms,
		commitMs: commit.ms,
		totalMs: plan.ms + commit.ms,
		sharedPlanMs: sharedPlan.ms,
		sharedCommitMs: sharedCommit.ms,
		sharedTotalMs: sharedPlan.ms + sharedCommit.ms,
		planUnattributedMs: unattributedMs(plan.ms, plan.phases),
		commitUnattributedMs: unattributedMs(commit.ms, commit.phases),
		sharedPlanUnattributedMs: unattributedMs(sharedPlan.ms, sharedPlan.phases),
		sharedCommitUnattributedMs: unattributedMs(sharedCommit.ms, sharedCommit.phases),
		planPayloadBytes: payloadBytes(plan.value),
		commitPayloadBytes: payloadBytes(commit.value),
		sharedPlanPayloadBytes: payloadBytes(sharedPlan.value.plan),
		sharedCommitPayloadBytes: payloadBytes(sharedCommit.value),
		commitOutputBytes: outputBytes,
		sharedCommitOutputBytes: sharedOutputBytes,
		commitTimingMs: commit.value.timings,
		sharedCommitTimingMs: sharedCommit.value.timings,
		planPhaseMs: phaseMap(plan.phases),
		commitPhaseMs: phaseMap(commit.phases),
		sharedPlanPhaseMs: phaseMap(sharedPlan.phases),
		sharedCommitPhaseMs: phaseMap(sharedCommit.phases),
		operationCount: ops.length,
		updateCount,
		changedCells: plan.value.preview.changedCells.length,
		commitChangedCells: commit.value.apply.affectedCells.length,
		sharedChangedCells: sharedPlan.value.plan.preview.changedCells.length,
		sharedCommitChangedCells: sharedCommit.value.apply.affectedCells.length,
		postWriteValid: commit.value.postWrite.valid,
		sharedPostWriteValid: sharedCommit.value.postWrite.valid,
	}
}

function allPhaseNames(
	samples: readonly Sample[],
	key: 'planPhaseMs' | 'commitPhaseMs' | 'sharedPlanPhaseMs' | 'sharedCommitPhaseMs',
): string[] {
	return [...new Set(samples.flatMap((sample) => Object.keys(sample[key])))].sort()
}

function summarizePhases(
	samples: readonly Sample[],
	key: 'planPhaseMs' | 'commitPhaseMs' | 'sharedPlanPhaseMs' | 'sharedCommitPhaseMs',
): Record<string, number> {
	const summary: Record<string, number> = {}
	for (const phase of allPhaseNames(samples, key)) {
		summary[phase] = median(samples.map((sample) => sample[key][phase] ?? 0))
	}
	return summary
}

function summarizePhaseStats(
	samples: readonly Sample[],
	key: 'planPhaseMs' | 'commitPhaseMs' | 'sharedPlanPhaseMs' | 'sharedCommitPhaseMs',
): Record<string, SampleStats> {
	const summary: Record<string, SampleStats> = {}
	for (const phase of allPhaseNames(samples, key)) {
		summary[phase] = sampleStats(samples.map((sample) => sample[key][phase] ?? 0))
	}
	return summary
}

function summarizeCommitTimings(
	samples: readonly Sample[],
	key: 'commitTimingMs' | 'sharedCommitTimingMs',
): Record<string, number> {
	const summary: Record<string, number> = {}
	const timingKeys = Object.keys(samples[0]?.[key] ?? {}) as (keyof AgentCommitTimings)[]
	for (const timingKey of timingKeys) {
		summary[timingKey] = median(samples.map((sample) => sample[key][timingKey]))
	}
	return summary
}

function summarize(samples: readonly Sample[]) {
	return {
		planMedianMs: median(samples.map((sample) => sample.planMs)),
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		totalMedianMs: median(samples.map((sample) => sample.totalMs)),
		sharedPlanMedianMs: median(samples.map((sample) => sample.sharedPlanMs)),
		sharedCommitMedianMs: median(samples.map((sample) => sample.sharedCommitMs)),
		sharedTotalMedianMs: median(samples.map((sample) => sample.sharedTotalMs)),
		planUnattributedMedianMs: median(samples.map((sample) => sample.planUnattributedMs)),
		commitUnattributedMedianMs: median(samples.map((sample) => sample.commitUnattributedMs)),
		sharedPlanUnattributedMedianMs: median(
			samples.map((sample) => sample.sharedPlanUnattributedMs),
		),
		sharedCommitUnattributedMedianMs: median(
			samples.map((sample) => sample.sharedCommitUnattributedMs),
		),
		sharedWorkflowSpeedupVsCold:
			median(samples.map((sample) => sample.totalMs)) /
			median(samples.map((sample) => sample.sharedTotalMs)),
		planPayloadBytesMedian: median(samples.map((sample) => sample.planPayloadBytes)),
		commitPayloadBytesMedian: median(samples.map((sample) => sample.commitPayloadBytes)),
		sharedPlanPayloadBytesMedian: median(samples.map((sample) => sample.sharedPlanPayloadBytes)),
		sharedCommitPayloadBytesMedian: median(
			samples.map((sample) => sample.sharedCommitPayloadBytes),
		),
		commitOutputBytesMedian: median(samples.map((sample) => sample.commitOutputBytes)),
		sharedCommitOutputBytesMedian: median(samples.map((sample) => sample.sharedCommitOutputBytes)),
		operationCountMedian: median(samples.map((sample) => sample.operationCount)),
		updateCountMedian: median(samples.map((sample) => sample.updateCount)),
		changedCellsMedian: median(samples.map((sample) => sample.changedCells)),
		commitChangedCellsMedian: median(samples.map((sample) => sample.commitChangedCells)),
		sharedChangedCellsMedian: median(samples.map((sample) => sample.sharedChangedCells)),
		sharedCommitChangedCellsMedian: median(
			samples.map((sample) => sample.sharedCommitChangedCells),
		),
		postWriteValid: samples.every((sample) => sample.postWriteValid),
		sharedPostWriteValid: samples.every((sample) => sample.sharedPostWriteValid),
		planPhaseMedianMs: summarizePhases(samples, 'planPhaseMs'),
		commitPhaseMedianMs: summarizePhases(samples, 'commitPhaseMs'),
		sharedPlanPhaseMedianMs: summarizePhases(samples, 'sharedPlanPhaseMs'),
		sharedCommitPhaseMedianMs: summarizePhases(samples, 'sharedCommitPhaseMs'),
		planPhaseStats: summarizePhaseStats(samples, 'planPhaseMs'),
		commitPhaseStats: summarizePhaseStats(samples, 'commitPhaseMs'),
		sharedPlanPhaseStats: summarizePhaseStats(samples, 'sharedPlanPhaseMs'),
		sharedCommitPhaseStats: summarizePhaseStats(samples, 'sharedCommitPhaseMs'),
		commitTimingMedianMs: summarizeCommitTimings(samples, 'commitTimingMs'),
		sharedCommitTimingMedianMs: summarizeCommitTimings(samples, 'sharedCommitTimingMs'),
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

function phaseOutputPath(input: BenchmarkInput): string {
	return join(
		tmpdir(),
		`ascend-agent-phase-${process.pid}-${Date.now()}-${basename(input.xlsxPath)}.out.xlsx`,
	)
}

async function run() {
	const args = parseArgs()
	const data = await resolveBenchmarkInput(args)
	const outputPath = phaseOutputPath(data)
	const ops = [buildSetCellsOperation(args.updates, data.rows, data.cols, data.sheet)]
	const samples: Sample[] = []
	const timeout =
		args.timeoutMs !== undefined
			? setTimeout(() => {
					console.error(
						JSON.stringify({
							tool: 'agent-phase-profile',
							status: 'timeout',
							timeoutMs: args.timeoutMs,
							completedSamples: samples.length,
							input: data,
						}),
					)
					process.exit(124)
				}, args.timeoutMs)
			: undefined
	const reportPhase = args.progress
		? (sample: number, phase: PhaseTiming) => {
				console.error(JSON.stringify({ tool: 'agent-phase-profile', sample, ...phase }))
			}
		: undefined
	try {
		for (let i = 0; i < args.warmup; i++) {
			if (args.progress)
				console.error(JSON.stringify({ tool: 'agent-phase-profile', sample: -args.warmup + i }))
			await runSample(data.xlsxPath, outputPath, ops, args.updates, (phase) =>
				reportPhase?.(-args.warmup + i, phase),
			)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			if (args.progress)
				console.error(JSON.stringify({ tool: 'agent-phase-profile', sample: i + 1 }))
			samples.push(
				await runSample(data.xlsxPath, outputPath, ops, args.updates, (phase) =>
					reportPhase?.(i + 1, phase),
				),
			)
			runGc()
		}
		const payload = {
			tool: 'agent-phase-profile',
			args,
			input: data,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		if (timeout !== undefined) clearTimeout(timeout)
		if (data.cleanup) await rm(data.xlsxPath, { force: true })
		await rm(outputPath, { force: true })
	}
}

await run()
