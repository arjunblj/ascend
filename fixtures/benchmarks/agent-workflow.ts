#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly rowLimit: number
	readonly mutations: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface WorkflowSample {
	readonly totalMs: number
	readonly fullTotalMs: number
	readonly measuredSampleMs: number
	readonly inspectMs: number
	readonly readMs: number
	readonly planMs: number
	readonly fullPlanMs: number
	readonly commitMs: number
	readonly verifyMs: number
	readonly payloadBytes: number
	readonly fullPayloadBytes: number
	readonly inspectPayloadBytes: number
	readonly readPayloadBytes: number
	readonly planPayloadBytes: number
	readonly fullPlanPayloadBytes: number
	readonly commitPayloadBytes: number
	readonly verifyPayloadBytes: number
	readonly readCells: number
	readonly readPartial: boolean
	readonly readWindowRows: number | null
	readonly planChangedCellCount: number | null
	readonly planEmittedChangedCellCount: number | null
	readonly mutationCount: number
	readonly rssDeltaMb: number
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
		rowLimit: positiveInt(readOption(process.argv, '--row-limit'), 500),
		mutations: positiveInt(readOption(process.argv, '--mutations'), 1),
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

async function post(apiFetch: typeof fetch, path: string, body: unknown) {
	const start = performance.now()
	const response = await apiFetch(
		new Request(`http://ascend.local${path}`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(body),
		}),
	)
	const text = await response.text()
	const ms = performance.now() - start
	if (response.status !== 200) throw new Error(`${path} failed: ${text}`)
	return { ms, text, payload: JSON.parse(text) as ApiEnvelope }
}

interface ApiEnvelope {
	readonly data?: {
		readonly valid?: boolean
		readonly cells?: readonly unknown[]
		readonly load?: {
			readonly isPartial?: boolean
			readonly loadedSheets?: readonly string[]
		}
		readonly preview?: {
			readonly changedCellCount?: number
			readonly emittedChangedCellCount?: number
		}
	}
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function rssMb(): number {
	return process.memoryUsage().rss / (1024 * 1024)
}

function buildCellMutations(count: number, rows: number, cols: number) {
	return Array.from({ length: count }, (_, index) => {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		return {
			path: `/sheets/Data/cells/${indexToColumn(col)}${row + 1}/value`,
			value: 9001 + index,
		}
	})
}

async function runWorkflow(
	apiFetch: typeof fetch,
	inputPath: string,
	outputPath: string,
	range: string,
	rowLimit: number,
	mutationCount: number,
	rows: number,
	cols: number,
): Promise<WorkflowSample> {
	await rm(outputPath, { force: true })
	runGc()
	const rssBefore = rssMb()
	const totalStart = performance.now()
	const inspect = await post(apiFetch, '/inspect', { file: inputPath })
	const read = await post(apiFetch, '/read', {
		file: inputPath,
		range,
		format: 'compact',
		rowLimit,
	})
	const mutations = buildCellMutations(mutationCount, rows, cols)
	const fullPlan = await post(apiFetch, '/plan', { file: inputPath, mutations })
	const plan = await post(apiFetch, '/plan', {
		file: inputPath,
		mutations,
		compact: true,
		maxChangedCells: 25,
	})
	const commit = await post(apiFetch, '/commit', {
		file: inputPath,
		output: outputPath,
		mutations,
		approvals: [],
	})
	const verify = await post(apiFetch, '/check', { file: outputPath })
	const measuredSampleMs = performance.now() - totalStart
	const rssAfter = rssMb()
	const readLoad = read.payload.data?.load
	const compactWorkflowBytes =
		inspect.text.length +
		read.text.length +
		plan.text.length +
		commit.text.length +
		verify.text.length
	const fullWorkflowBytes =
		inspect.text.length +
		read.text.length +
		fullPlan.text.length +
		commit.text.length +
		verify.text.length
	return {
		totalMs: inspect.ms + read.ms + plan.ms + commit.ms + verify.ms,
		fullTotalMs: inspect.ms + read.ms + fullPlan.ms + commit.ms + verify.ms,
		measuredSampleMs,
		inspectMs: inspect.ms,
		readMs: read.ms,
		planMs: plan.ms,
		fullPlanMs: fullPlan.ms,
		commitMs: commit.ms,
		verifyMs: verify.ms,
		payloadBytes: compactWorkflowBytes,
		fullPayloadBytes: fullWorkflowBytes,
		inspectPayloadBytes: inspect.text.length,
		readPayloadBytes: read.text.length,
		planPayloadBytes: plan.text.length,
		fullPlanPayloadBytes: fullPlan.text.length,
		commitPayloadBytes: commit.text.length,
		verifyPayloadBytes: verify.text.length,
		readCells: read.payload.data?.cells?.length ?? 0,
		readPartial: readLoad?.isPartial === true,
		readWindowRows: readLoad?.loadedSheets ? rowLimit : null,
		planChangedCellCount: plan.payload.data?.preview?.changedCellCount ?? null,
		planEmittedChangedCellCount: plan.payload.data?.preview?.emittedChangedCellCount ?? null,
		mutationCount,
		rssDeltaMb: rssAfter - rssBefore,
		valid: verify.payload.data?.valid === true,
	}
}

function summarize(samples: readonly WorkflowSample[]) {
	return {
		totalMedianMs: median(samples.map((sample) => sample.totalMs)),
		fullTotalMedianMs: median(samples.map((sample) => sample.fullTotalMs)),
		measuredSampleMedianMs: median(samples.map((sample) => sample.measuredSampleMs)),
		inspectMedianMs: median(samples.map((sample) => sample.inspectMs)),
		readMedianMs: median(samples.map((sample) => sample.readMs)),
		planMedianMs: median(samples.map((sample) => sample.planMs)),
		fullPlanMedianMs: median(samples.map((sample) => sample.fullPlanMs)),
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		verifyMedianMs: median(samples.map((sample) => sample.verifyMs)),
		payloadBytesMedian: median(samples.map((sample) => sample.payloadBytes)),
		fullPayloadBytesMedian: median(samples.map((sample) => sample.fullPayloadBytes)),
		inspectPayloadBytesMedian: median(samples.map((sample) => sample.inspectPayloadBytes)),
		readPayloadBytesMedian: median(samples.map((sample) => sample.readPayloadBytes)),
		planPayloadBytesMedian: median(samples.map((sample) => sample.planPayloadBytes)),
		fullPlanPayloadBytesMedian: median(samples.map((sample) => sample.fullPlanPayloadBytes)),
		commitPayloadBytesMedian: median(samples.map((sample) => sample.commitPayloadBytes)),
		verifyPayloadBytesMedian: median(samples.map((sample) => sample.verifyPayloadBytes)),
		compactWorkflowSpeedupVsFull:
			median(samples.map((sample) => sample.fullTotalMs)) /
			median(samples.map((sample) => sample.totalMs)),
		planPayloadReduction:
			median(samples.map((sample) => sample.fullPlanPayloadBytes)) /
			median(samples.map((sample) => sample.planPayloadBytes)),
		readCellsMedian: median(samples.map((sample) => sample.readCells)),
		readWindowRowsMedian: medianOptional(samples.map((sample) => sample.readWindowRows)),
		planChangedCellCountMedian: medianOptional(
			samples.map((sample) => sample.planChangedCellCount),
		),
		planEmittedChangedCellCountMedian: medianOptional(
			samples.map((sample) => sample.planEmittedChangedCellCount),
		),
		mutationCountMedian: median(samples.map((sample) => sample.mutationCount)),
		rssDeltaMbMedian: median(samples.map((sample) => sample.rssDeltaMb)),
		readPartial: samples.every((sample) => sample.readPartial),
		valid: samples.every((sample) => sample.valid),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const range = `A1:${indexToColumn(args.cols - 1)}${args.rows}`
	const apiFetch = createApiFetch()
	const outputPath = `${data.xlsxPath}.workflow-output.xlsx`
	const samples: WorkflowSample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runWorkflow(
				apiFetch,
				data.xlsxPath,
				outputPath,
				range,
				args.rowLimit,
				args.mutations,
				args.rows,
				args.cols,
			)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(
				await runWorkflow(
					apiFetch,
					data.xlsxPath,
					outputPath,
					range,
					args.rowLimit,
					args.mutations,
					args.rows,
					args.cols,
				),
			)
			runGc()
		}
		const payload = {
			tool: 'agent-workflow',
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
