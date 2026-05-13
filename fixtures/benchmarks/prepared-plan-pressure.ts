#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import type { PathMutation } from '../../packages/sdk/src/types.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly mutations: number
	readonly handles: number
	readonly maxHandles: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface Sample {
	readonly planTotalMs: number
	readonly planPerHandleMs: number
	readonly planPayloadBytesMedian: number
	readonly preparedHandlesCreated: number
	readonly maxHandles: number
	readonly estimatedEvictedHandles: number
	readonly rssRetainedAfterPlansMb: number
	readonly rssPerRetainedHandleMb: number
	readonly firstCommitStatus: number | null
	readonly firstHandleEvicted: boolean | null
	readonly latestCommitMs: number
	readonly latestCommitStatus: number
	readonly latestCommitPayloadBytes: number
	readonly latestCommitOk: boolean
}

interface ApiEnvelope {
	readonly ok?: boolean
	readonly data?: {
		readonly preparedPlan?: {
			readonly id?: string
			readonly ttlMs?: number
		}
	}
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
	const handles = positiveInt(readOption(process.argv, '--handles'), 96)
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		mutations: positiveInt(readOption(process.argv, '--mutations'), 25),
		handles,
		maxHandles: positiveInt(readOption(process.argv, '--max-handles'), Math.min(handles, 64)),
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

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function rssMb(): number {
	return process.memoryUsage().rss / (1024 * 1024)
}

function buildCellMutations(count: number, rows: number, cols: number): readonly PathMutation[] {
	const mutations: PathMutation[] = []
	for (let index = 0; index < count; index++) {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		mutations.push({
			path: `/sheets/Data/cells/${indexToColumn(col)}${row + 1}/value`,
			value: 700_000 + index,
		})
	}
	return mutations
}

async function post(
	apiFetch: typeof fetch,
	path: string,
	body: unknown,
): Promise<{
	readonly ms: number
	readonly status: number
	readonly text: string
	readonly payload: ApiEnvelope
}> {
	const start = performance.now()
	const response = await apiFetch(
		new Request(`http://ascend.local${path}`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(body),
		}),
	)
	const text = await response.text()
	return {
		ms: performance.now() - start,
		status: response.status,
		text,
		payload: JSON.parse(text) as ApiEnvelope,
	}
}

async function runSample(
	inputPath: string,
	outputPrefix: string,
	mutations: readonly PathMutation[],
	args: Args,
): Promise<Sample> {
	const apiFetch = createApiFetch({ preparedPlanMaxHandles: args.maxHandles })
	const handles: string[] = []
	const payloadBytes: number[] = []
	const planStart = performance.now()
	runGc()
	const rssBefore = rssMb()
	for (let index = 0; index < args.handles; index++) {
		const plan = await post(apiFetch, '/plan', {
			file: inputPath,
			mutations,
			compact: true,
			prepare: true,
			maxChangedCells: 25,
		})
		if (plan.status !== 200) throw new Error(`/plan failed: ${plan.text}`)
		const handle = plan.payload.data?.preparedPlan?.id
		if (!handle) throw new Error('Prepared plan did not return a handle')
		handles.push(handle)
		payloadBytes.push(plan.text.length)
	}
	const planTotalMs = performance.now() - planStart
	runGc()
	const rssAfterPlans = rssMb()
	const firstHandle = handles[0]
	const latestHandle = handles[handles.length - 1]
	let firstCommitStatus: number | null = null
	let firstHandleEvicted: boolean | null = null
	if (firstHandle && latestHandle && firstHandle !== latestHandle) {
		const firstCommit = await post(apiFetch, '/commit', {
			planHandle: firstHandle,
			output: `${outputPrefix}.first.xlsx`,
			approvals: [],
		})
		firstCommitStatus = firstCommit.status
		firstHandleEvicted = firstCommit.status !== 200
	}
	if (!latestHandle) throw new Error('No latest prepared handle captured')
	const latestCommit = await post(apiFetch, '/commit', {
		planHandle: latestHandle,
		output: `${outputPrefix}.latest.xlsx`,
		approvals: [],
	})
	if (latestCommit.status !== 200) throw new Error(`/commit latest failed: ${latestCommit.text}`)
	const retainedHandles = Math.min(args.handles, args.maxHandles)
	return {
		planTotalMs,
		planPerHandleMs: planTotalMs / args.handles,
		planPayloadBytesMedian: median(payloadBytes),
		preparedHandlesCreated: handles.length,
		maxHandles: args.maxHandles,
		estimatedEvictedHandles: Math.max(0, args.handles - args.maxHandles),
		rssRetainedAfterPlansMb: rssAfterPlans - rssBefore,
		rssPerRetainedHandleMb: (rssAfterPlans - rssBefore) / retainedHandles,
		firstCommitStatus,
		firstHandleEvicted,
		latestCommitMs: latestCommit.ms,
		latestCommitStatus: latestCommit.status,
		latestCommitPayloadBytes: latestCommit.text.length,
		latestCommitOk: latestCommit.payload.ok === true,
	}
}

function summarize(samples: readonly Sample[]) {
	return {
		planTotalMedianMs: median(samples.map((sample) => sample.planTotalMs)),
		planPerHandleMedianMs: median(samples.map((sample) => sample.planPerHandleMs)),
		planPayloadBytesMedian: median(samples.map((sample) => sample.planPayloadBytesMedian)),
		preparedHandlesCreatedMedian: median(samples.map((sample) => sample.preparedHandlesCreated)),
		maxHandlesMedian: median(samples.map((sample) => sample.maxHandles)),
		estimatedEvictedHandlesMedian: median(samples.map((sample) => sample.estimatedEvictedHandles)),
		rssRetainedAfterPlansMedianMb: median(samples.map((sample) => sample.rssRetainedAfterPlansMb)),
		rssPerRetainedHandleMedianMb: median(samples.map((sample) => sample.rssPerRetainedHandleMb)),
		firstHandleEvicted: samples.every((sample) => sample.firstHandleEvicted !== false),
		latestCommitMedianMs: median(samples.map((sample) => sample.latestCommitMs)),
		latestCommitPayloadBytesMedian: median(
			samples.map((sample) => sample.latestCommitPayloadBytes),
		),
		latestCommitOk: samples.every((sample) => sample.latestCommitOk),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const mutations = buildCellMutations(args.mutations, args.rows, args.cols)
	const outputPrefix = `${data.xlsxPath}.prepared-pressure`
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath, `${outputPrefix}.warmup-${i}`, mutations, args)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath, `${outputPrefix}.sample-${i}`, mutations, args))
			runGc()
		}
		const payload = {
			tool: 'prepared-plan-pressure',
			args,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(data.xlsxPath, { force: true })
		for (let i = 0; i < args.warmup; i++) {
			await rm(`${outputPrefix}.warmup-${i}.first.xlsx`, { force: true })
			await rm(`${outputPrefix}.warmup-${i}.latest.xlsx`, { force: true })
		}
		for (let i = 0; i < args.repeat; i++) {
			await rm(`${outputPrefix}.sample-${i}.first.xlsx`, { force: true })
			await rm(`${outputPrefix}.sample-${i}.latest.xlsx`, { force: true })
		}
	}
}

await run()
