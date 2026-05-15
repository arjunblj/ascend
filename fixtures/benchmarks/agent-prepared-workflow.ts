#!/usr/bin/env bun
import { rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'

interface Args {
	readonly inputFile: string
	readonly sheet: string
	readonly range: string
	readonly rows: number
	readonly cols: number
	readonly rowLimit: number
	readonly mutations: number
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface ApiEnvelope {
	readonly data?: {
		readonly valid?: boolean
		readonly cells?: readonly unknown[]
		readonly load?: { readonly isPartial?: boolean }
		readonly preview?: {
			readonly changedCellCount?: number
			readonly emittedChangedCellCount?: number
		}
		readonly preparedPlan?: { readonly id?: string }
		readonly timings?: CommitTimings
		readonly postWrite?: { readonly timings?: PostWriteTimings; readonly valid?: boolean }
	}
}

interface CommitTimings {
	readonly writePolicySnapshotMs?: number
	readonly packageGraphMs?: number
	readonly packageGraphAuditMs?: number
	readonly applyMs?: number
	readonly writePlanSummaryMs?: number
	readonly writePolicyCheckMs?: number
	readonly writePolicyBuildMs?: number
	readonly toBytesMs?: number
	readonly writeFileMs?: number
	readonly outputByteReadMs?: number
	readonly outputHashMs?: number
}

interface PostWriteTimings {
	readonly reopenMs?: number
	readonly checkMs?: number
	readonly lintMs?: number
	readonly preservationMs?: number
	readonly packageGraphMs?: number
	readonly packageGraphAuditMs?: number
}

interface ApiResult {
	readonly ms: number
	readonly text: string
	readonly payload: ApiEnvelope
}

interface Sample {
	readonly totalMs: number
	readonly commitVerifiedTotalMs: number
	readonly inspectMs: number
	readonly readMs: number
	readonly planMs: number
	readonly commitMs: number
	readonly verifyMs: number
	readonly commitApplyMs: number | null
	readonly commitWritePlanSummaryMs: number | null
	readonly commitWritePolicyCheckMs: number | null
	readonly commitWritePolicyBuildMs: number | null
	readonly commitToBytesMs: number | null
	readonly commitWriteFileMs: number | null
	readonly commitOutputByteReadMs: number | null
	readonly commitOutputHashMs: number | null
	readonly postWriteReopenMs: number | null
	readonly postWriteCheckMs: number | null
	readonly postWriteLintMs: number | null
	readonly postWritePreservationMs: number | null
	readonly postWritePackageGraphMs: number | null
	readonly postWritePackageGraphAuditMs: number | null
	readonly payloadBytes: number
	readonly commitVerifiedPayloadBytes: number
	readonly outputBytes: number
	readonly readCells: number
	readonly readPartial: boolean
	readonly changedCellCount: number | null
	readonly emittedChangedCellCount: number | null
	readonly rssDeltaMb: number
	readonly valid: boolean
	readonly postWriteValid: boolean
}

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
	const inputFile = readOption(process.argv, '--input-file')
	if (!inputFile) throw new Error('Missing --input-file')
	return {
		inputFile,
		sheet: readOption(process.argv, '--sheet') ?? 'Data',
		range: readOption(process.argv, '--range') ?? 'A1:E65536',
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 5),
		rowLimit: positiveInt(readOption(process.argv, '--row-limit'), 500),
		mutations: positiveInt(readOption(process.argv, '--mutations'), 25),
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

function percentile(values: readonly number[], p: number): number {
	const sorted = [...values].sort((a, b) => a - b)
	if (sorted.length === 0) return 0
	const index = Math.ceil(sorted.length * p) - 1
	return sorted[Math.min(Math.max(index, 0), sorted.length - 1)] ?? 0
}

function mean(values: readonly number[]): number {
	return values.reduce((sum, value) => sum + value, 0) / values.length
}

function standardDeviation(values: readonly number[]): number {
	if (values.length < 2) return 0
	const avg = mean(values)
	const variance =
		values.reduce((sum, value) => {
			const delta = value - avg
			return sum + delta * delta
		}, 0) /
		(values.length - 1)
	return Math.sqrt(variance)
}

function seriesStats(values: readonly number[]) {
	const avg = mean(values)
	const stddev = standardDeviation(values)
	return {
		sampleCount: values.length,
		min: Math.min(...values),
		median: median(values),
		mean: avg,
		p95: percentile(values, 0.95),
		max: Math.max(...values),
		stddev,
		cv: avg === 0 ? 0 : stddev / avg,
	}
}

function optionalMedian(values: readonly (number | null)[]): number | undefined {
	const defined = values.filter((value): value is number => value !== null)
	return defined.length > 0 ? median(defined) : undefined
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function rssMb(): number {
	return process.memoryUsage().rss / (1024 * 1024)
}

async function post(apiFetch: typeof fetch, path: string, body: unknown): Promise<ApiResult> {
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

function pointerSegment(value: string): string {
	return encodeURIComponent(value.replace(/~/g, '~0').replace(/\//g, '~1'))
}

function buildCellMutations(count: number, rows: number, cols: number, sheetName: string) {
	const sheetSegment = pointerSegment(sheetName)
	return Array.from({ length: count }, (_, index) => {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		return {
			path: `/sheets/${sheetSegment}/cells/${indexToColumn(col)}${row + 1}/value`,
			value: 9001 + index,
		}
	})
}

function maybeNumber(value: number | undefined): number | null {
	return typeof value === 'number' ? value : null
}

async function runSample(apiFetch: typeof fetch, args: Args, outputPath: string): Promise<Sample> {
	await rm(outputPath, { force: true })
	runGc()
	const rssBefore = rssMb()
	const inspect = await post(apiFetch, '/inspect', { file: args.inputFile })
	const read = await post(apiFetch, '/read', {
		file: args.inputFile,
		sheet: args.sheet,
		range: args.range,
		format: 'compact',
		rowLimit: args.rowLimit,
	})
	const mutations = buildCellMutations(args.mutations, args.rows, args.cols, args.sheet)
	const plan = await post(apiFetch, '/plan', {
		file: args.inputFile,
		mutations,
		compact: true,
		maxChangedCells: args.mutations,
	})
	const planHandle = plan.payload.data?.preparedPlan?.id
	if (!planHandle) throw new Error('Prepared plan did not return a handle')
	const commit = await post(apiFetch, '/commit', {
		planHandle,
		output: outputPath,
		compact: true,
	})
	const verify = await post(apiFetch, '/check', { file: outputPath })
	const outputBytes = (await stat(outputPath)).size
	const timings = commit.payload.data?.timings
	const postWrite = commit.payload.data?.postWrite
	const postWriteTimings = postWrite?.timings
	const commitVerifiedTotalMs = inspect.ms + read.ms + plan.ms + commit.ms
	const totalMs = commitVerifiedTotalMs + verify.ms
	return {
		totalMs,
		commitVerifiedTotalMs,
		inspectMs: inspect.ms,
		readMs: read.ms,
		planMs: plan.ms,
		commitMs: commit.ms,
		verifyMs: verify.ms,
		commitApplyMs: maybeNumber(timings?.applyMs),
		commitWritePlanSummaryMs: maybeNumber(timings?.writePlanSummaryMs),
		commitWritePolicyCheckMs: maybeNumber(timings?.writePolicyCheckMs),
		commitWritePolicyBuildMs: maybeNumber(timings?.writePolicyBuildMs),
		commitToBytesMs: maybeNumber(timings?.toBytesMs),
		commitWriteFileMs: maybeNumber(timings?.writeFileMs),
		commitOutputByteReadMs: maybeNumber(timings?.outputByteReadMs),
		commitOutputHashMs: maybeNumber(timings?.outputHashMs),
		postWriteReopenMs: maybeNumber(postWriteTimings?.reopenMs),
		postWriteCheckMs: maybeNumber(postWriteTimings?.checkMs),
		postWriteLintMs: maybeNumber(postWriteTimings?.lintMs),
		postWritePreservationMs: maybeNumber(postWriteTimings?.preservationMs),
		postWritePackageGraphMs: maybeNumber(postWriteTimings?.packageGraphMs),
		postWritePackageGraphAuditMs: maybeNumber(postWriteTimings?.packageGraphAuditMs),
		payloadBytes:
			inspect.text.length +
			read.text.length +
			plan.text.length +
			commit.text.length +
			verify.text.length,
		commitVerifiedPayloadBytes:
			inspect.text.length + read.text.length + plan.text.length + commit.text.length,
		outputBytes,
		readCells: read.payload.data?.cells?.length ?? 0,
		readPartial: read.payload.data?.load?.isPartial === true,
		changedCellCount: plan.payload.data?.preview?.changedCellCount ?? null,
		emittedChangedCellCount: plan.payload.data?.preview?.emittedChangedCellCount ?? null,
		rssDeltaMb: rssMb() - rssBefore,
		valid: verify.payload.data?.valid === true,
		postWriteValid: postWrite?.valid === true,
	}
}

function summarize(samples: readonly Sample[]) {
	return {
		totalMedianMs: median(samples.map((sample) => sample.totalMs)),
		totalStats: seriesStats(samples.map((sample) => sample.totalMs)),
		commitVerifiedTotalMedianMs: median(samples.map((sample) => sample.commitVerifiedTotalMs)),
		commitVerifiedTotalStats: seriesStats(samples.map((sample) => sample.commitVerifiedTotalMs)),
		inspectMedianMs: median(samples.map((sample) => sample.inspectMs)),
		readMedianMs: median(samples.map((sample) => sample.readMs)),
		planMedianMs: median(samples.map((sample) => sample.planMs)),
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		verifyMedianMs: median(samples.map((sample) => sample.verifyMs)),
		commitApplyMedianMs: optionalMedian(samples.map((sample) => sample.commitApplyMs)),
		commitWritePlanSummaryMedianMs: optionalMedian(
			samples.map((sample) => sample.commitWritePlanSummaryMs),
		),
		commitWritePolicyCheckMedianMs: optionalMedian(
			samples.map((sample) => sample.commitWritePolicyCheckMs),
		),
		commitWritePolicyBuildMedianMs: optionalMedian(
			samples.map((sample) => sample.commitWritePolicyBuildMs),
		),
		commitToBytesMedianMs: optionalMedian(samples.map((sample) => sample.commitToBytesMs)),
		commitWriteFileMedianMs: optionalMedian(samples.map((sample) => sample.commitWriteFileMs)),
		commitOutputByteReadMedianMs: optionalMedian(
			samples.map((sample) => sample.commitOutputByteReadMs),
		),
		commitOutputHashMedianMs: optionalMedian(samples.map((sample) => sample.commitOutputHashMs)),
		postWriteReopenMedianMs: optionalMedian(samples.map((sample) => sample.postWriteReopenMs)),
		postWriteCheckMedianMs: optionalMedian(samples.map((sample) => sample.postWriteCheckMs)),
		postWriteLintMedianMs: optionalMedian(samples.map((sample) => sample.postWriteLintMs)),
		postWritePreservationMedianMs: optionalMedian(
			samples.map((sample) => sample.postWritePreservationMs),
		),
		postWritePackageGraphMedianMs: optionalMedian(
			samples.map((sample) => sample.postWritePackageGraphMs),
		),
		postWritePackageGraphAuditMedianMs: optionalMedian(
			samples.map((sample) => sample.postWritePackageGraphAuditMs),
		),
		payloadBytesMedian: median(samples.map((sample) => sample.payloadBytes)),
		commitVerifiedPayloadBytesMedian: median(
			samples.map((sample) => sample.commitVerifiedPayloadBytes),
		),
		outputBytesMedian: median(samples.map((sample) => sample.outputBytes)),
		readCellsMedian: median(samples.map((sample) => sample.readCells)),
		changedCellCountMedian: optionalMedian(samples.map((sample) => sample.changedCellCount)),
		emittedChangedCellCountMedian: optionalMedian(
			samples.map((sample) => sample.emittedChangedCellCount),
		),
		rssDeltaMbMedian: median(samples.map((sample) => sample.rssDeltaMb)),
		rssDeltaMbStats: seriesStats(samples.map((sample) => sample.rssDeltaMb)),
		readPartial: samples.every((sample) => sample.readPartial),
		valid: samples.every((sample) => sample.valid),
		postWriteValid: samples.every((sample) => sample.postWriteValid),
	}
}

async function run() {
	const args = parseArgs()
	const apiFetch = createApiFetch()
	const outputPath = join(
		tmpdir(),
		`ascend-agent-prepared-workflow-${process.pid}-${Date.now()}-${basename(args.inputFile)}.out.xlsx`,
	)
	const samples: Sample[] = []
	try {
		for (let index = 0; index < args.warmup; index++) {
			await runSample(apiFetch, args, outputPath)
			runGc()
		}
		for (let index = 0; index < args.repeat; index++) {
			samples.push(await runSample(apiFetch, args, outputPath))
			runGc()
		}
		const payload = {
			tool: 'agent-prepared-workflow',
			args,
			input: {
				xlsxPath: args.inputFile,
				sheet: args.sheet,
				range: args.range,
				rowLimit: args.rowLimit,
			},
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(outputPath, { force: true })
	}
}

await run()
