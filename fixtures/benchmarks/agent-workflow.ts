#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly rowLimit: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface WorkflowSample {
	readonly totalMs: number
	readonly inspectMs: number
	readonly readMs: number
	readonly planMs: number
	readonly commitMs: number
	readonly verifyMs: number
	readonly payloadBytes: number
	readonly inspectPayloadBytes: number
	readonly readPayloadBytes: number
	readonly planPayloadBytes: number
	readonly commitPayloadBytes: number
	readonly verifyPayloadBytes: number
	readonly readCells: number
	readonly readPartial: boolean
	readonly readWindowRows: number | null
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
	}
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function rssMb(): number {
	return process.memoryUsage().rss / (1024 * 1024)
}

async function runWorkflow(
	apiFetch: typeof fetch,
	inputPath: string,
	outputPath: string,
	range: string,
	rowLimit: number,
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
	const mutations = [{ path: '/sheets/Data/cells/A1/value', value: 9001 }]
	const plan = await post(apiFetch, '/plan', { file: inputPath, mutations })
	const commit = await post(apiFetch, '/commit', {
		file: inputPath,
		output: outputPath,
		mutations,
		approvals: [],
	})
	const verify = await post(apiFetch, '/check', { file: outputPath })
	const totalMs = performance.now() - totalStart
	const rssAfter = rssMb()
	const readLoad = read.payload.data?.load
	return {
		totalMs,
		inspectMs: inspect.ms,
		readMs: read.ms,
		planMs: plan.ms,
		commitMs: commit.ms,
		verifyMs: verify.ms,
		payloadBytes:
			inspect.text.length +
			read.text.length +
			plan.text.length +
			commit.text.length +
			verify.text.length,
		inspectPayloadBytes: inspect.text.length,
		readPayloadBytes: read.text.length,
		planPayloadBytes: plan.text.length,
		commitPayloadBytes: commit.text.length,
		verifyPayloadBytes: verify.text.length,
		readCells: read.payload.data?.cells?.length ?? 0,
		readPartial: readLoad?.isPartial === true,
		readWindowRows: readLoad?.loadedSheets ? rowLimit : null,
		mutationCount: mutations.length,
		rssDeltaMb: rssAfter - rssBefore,
		valid: verify.payload.data?.valid === true,
	}
}

function summarize(samples: readonly WorkflowSample[]) {
	return {
		totalMedianMs: median(samples.map((sample) => sample.totalMs)),
		inspectMedianMs: median(samples.map((sample) => sample.inspectMs)),
		readMedianMs: median(samples.map((sample) => sample.readMs)),
		planMedianMs: median(samples.map((sample) => sample.planMs)),
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		verifyMedianMs: median(samples.map((sample) => sample.verifyMs)),
		payloadBytesMedian: median(samples.map((sample) => sample.payloadBytes)),
		inspectPayloadBytesMedian: median(samples.map((sample) => sample.inspectPayloadBytes)),
		readPayloadBytesMedian: median(samples.map((sample) => sample.readPayloadBytes)),
		planPayloadBytesMedian: median(samples.map((sample) => sample.planPayloadBytes)),
		commitPayloadBytesMedian: median(samples.map((sample) => sample.commitPayloadBytes)),
		verifyPayloadBytesMedian: median(samples.map((sample) => sample.verifyPayloadBytes)),
		readCellsMedian: median(samples.map((sample) => sample.readCells)),
		readWindowRowsMedian: medianOptional(samples.map((sample) => sample.readWindowRows)),
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
			await runWorkflow(apiFetch, data.xlsxPath, outputPath, range, args.rowLimit)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runWorkflow(apiFetch, data.xlsxPath, outputPath, range, args.rowLimit))
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
