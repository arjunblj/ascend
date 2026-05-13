#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import type { ApplyResult, PathMutation, PreviewResult } from '../../packages/sdk/src/types.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly mutations: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface Sample {
	readonly directOpenMs: number
	readonly directCompileMs: number
	readonly directPreviewJournalMs: number
	readonly directApplyJournalMs: number
	readonly apiPlanMs: number
	readonly apiCompactPlanMs: number
	readonly apiCommitMs: number
	readonly apiPlanPayloadBytes: number
	readonly apiCompactPlanPayloadBytes: number
	readonly apiCommitPayloadBytes: number
	readonly mutationCount: number
	readonly compiledOps: number
	readonly compileIssues: number
	readonly previewJournalEntries: number
	readonly previewInverseOps: number
	readonly previewPreimages: number
	readonly applyJournalEntries: number
	readonly applyInverseOps: number
	readonly applyPreimages: number
	readonly rssDeltaMb: number
	readonly commitOk: boolean
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
		mutations: positiveInt(readOption(process.argv, '--mutations'), 1_000),
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

async function timed<T>(
	fn: () => Promise<T> | T,
): Promise<{ readonly ms: number; readonly value: T }> {
	const start = performance.now()
	const value = await fn()
	return { ms: performance.now() - start, value }
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
	readonly ok?: boolean
}

function buildCellMutations(count: number, rows: number, cols: number): readonly PathMutation[] {
	const mutations: PathMutation[] = []
	for (let index = 0; index < count; index++) {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		const ref = `${indexToColumn(col)}${row + 1}`
		mutations.push({
			path: `/sheets/Data/cells/${ref}/value`,
			value: 100_000 + index,
		})
	}
	return mutations
}

function journalEntries(result: PreviewResult | ApplyResult): number {
	return result.journal?.entries.length ?? 0
}

function journalInverseOps(result: PreviewResult | ApplyResult): number {
	return result.journal?.inverseOps.length ?? 0
}

function journalPreimages(result: PreviewResult | ApplyResult): number {
	return (
		result.journal?.entries.reduce(
			(total, entry) =>
				total +
				entry.preimages.reduce((entryTotal, preimage) => {
					if (preimage.kind === 'cells') return entryTotal + preimage.cells.length
					return entryTotal + 1
				}, 0),
			0,
		) ?? 0
	)
}

async function runSample(
	apiFetch: typeof fetch,
	inputPath: string,
	outputPath: string,
	mutations: readonly PathMutation[],
): Promise<Sample> {
	await rm(outputPath, { force: true })
	runGc()
	const rssBefore = rssMb()
	const opened = await timed(() => AscendWorkbook.open(inputPath))
	const compile = await timed(() => opened.value.compilePathMutations(mutations))
	if (compile.value.issueCount > 0) {
		throw new Error(`Path mutation compile failed with ${compile.value.issueCount} issue(s)`)
	}
	const preview = await timed(() => opened.value.preview(compile.value.ops, { journal: true }))
	const apply = await timed(() => opened.value.apply(compile.value.ops, { journal: true }))
	const apiPlan = await post(apiFetch, '/plan', { file: inputPath, mutations })
	const apiCompactPlan = await post(apiFetch, '/plan', {
		file: inputPath,
		mutations,
		compact: true,
		maxChangedCells: 25,
	})
	const apiCommit = await post(apiFetch, '/commit', {
		file: inputPath,
		output: outputPath,
		mutations,
		approvals: [],
	})
	const rssAfter = rssMb()
	return {
		directOpenMs: opened.ms,
		directCompileMs: compile.ms,
		directPreviewJournalMs: preview.ms,
		directApplyJournalMs: apply.ms,
		apiPlanMs: apiPlan.ms,
		apiCompactPlanMs: apiCompactPlan.ms,
		apiCommitMs: apiCommit.ms,
		apiPlanPayloadBytes: apiPlan.text.length,
		apiCompactPlanPayloadBytes: apiCompactPlan.text.length,
		apiCommitPayloadBytes: apiCommit.text.length,
		mutationCount: mutations.length,
		compiledOps: compile.value.ops.length,
		compileIssues: compile.value.issueCount,
		previewJournalEntries: journalEntries(preview.value),
		previewInverseOps: journalInverseOps(preview.value),
		previewPreimages: journalPreimages(preview.value),
		applyJournalEntries: journalEntries(apply.value),
		applyInverseOps: journalInverseOps(apply.value),
		applyPreimages: journalPreimages(apply.value),
		rssDeltaMb: rssAfter - rssBefore,
		commitOk: apiCommit.payload.ok === true,
	}
}

function summarize(samples: readonly Sample[]) {
	return {
		directOpenMedianMs: median(samples.map((sample) => sample.directOpenMs)),
		directCompileMedianMs: median(samples.map((sample) => sample.directCompileMs)),
		directPreviewJournalMedianMs: median(samples.map((sample) => sample.directPreviewJournalMs)),
		directApplyJournalMedianMs: median(samples.map((sample) => sample.directApplyJournalMs)),
		apiPlanMedianMs: median(samples.map((sample) => sample.apiPlanMs)),
		apiCompactPlanMedianMs: median(samples.map((sample) => sample.apiCompactPlanMs)),
		apiCommitMedianMs: median(samples.map((sample) => sample.apiCommitMs)),
		apiPlanPayloadBytesMedian: median(samples.map((sample) => sample.apiPlanPayloadBytes)),
		apiCompactPlanPayloadBytesMedian: median(
			samples.map((sample) => sample.apiCompactPlanPayloadBytes),
		),
		apiCommitPayloadBytesMedian: median(samples.map((sample) => sample.apiCommitPayloadBytes)),
		mutationCountMedian: median(samples.map((sample) => sample.mutationCount)),
		compiledOpsMedian: median(samples.map((sample) => sample.compiledOps)),
		compileIssuesMedian: median(samples.map((sample) => sample.compileIssues)),
		previewJournalEntriesMedian: median(samples.map((sample) => sample.previewJournalEntries)),
		previewInverseOpsMedian: median(samples.map((sample) => sample.previewInverseOps)),
		previewPreimagesMedian: median(samples.map((sample) => sample.previewPreimages)),
		applyJournalEntriesMedian: median(samples.map((sample) => sample.applyJournalEntries)),
		applyInverseOpsMedian: median(samples.map((sample) => sample.applyInverseOps)),
		applyPreimagesMedian: median(samples.map((sample) => sample.applyPreimages)),
		rssDeltaMbMedian: median(samples.map((sample) => sample.rssDeltaMb)),
		commitOk: samples.every((sample) => sample.commitOk),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const outputPath = `${data.xlsxPath}.path-mutation-output.xlsx`
	const mutations = buildCellMutations(args.mutations, args.rows, args.cols)
	const apiFetch = createApiFetch()
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(apiFetch, data.xlsxPath, outputPath, mutations)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(apiFetch, data.xlsxPath, outputPath, mutations))
			runGc()
		}
		const payload = {
			tool: 'path-mutation-overhead',
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
