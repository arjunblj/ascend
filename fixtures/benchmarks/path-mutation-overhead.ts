#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { createServer } from '../../apps/mcp/src/index.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import type { ApplyResult, PathMutation, PreviewResult } from '../../packages/sdk/src/types.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

type Surface = 'api' | 'both'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly mutations: number
	readonly surface: Surface
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
	readonly mcpPlanMs?: number
	readonly mcpCompactPlanMs?: number
	readonly mcpCommitMs?: number
	readonly apiPlanPayloadBytes: number
	readonly apiCompactPlanPayloadBytes: number
	readonly apiCommitPayloadBytes: number
	readonly mcpPlanPayloadBytes?: number
	readonly mcpCompactPlanPayloadBytes?: number
	readonly mcpCommitPayloadBytes?: number
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
	readonly mcpCommitOk?: boolean
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
	const surface = readOption(process.argv, '--surface') ?? 'api'
	if (surface !== 'api' && surface !== 'both') {
		throw new Error(`Unsupported --surface "${surface}". Use api or both.`)
	}
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		mutations: positiveInt(readOption(process.argv, '--mutations'), 1_000),
		surface,
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

function medianOptional(values: readonly (number | undefined)[]): number | undefined {
	const present = values.filter((value): value is number => value !== undefined)
	return present.length > 0 ? median(present) : undefined
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

interface SurfaceResponse {
	readonly ms: number
	readonly text: string
	readonly payload: ApiEnvelope
}

interface McpMutationClient {
	plan(body: unknown): Promise<SurfaceResponse>
	commit(body: unknown): Promise<SurfaceResponse>
}

function createMcpMutationClient(): McpMutationClient {
	const server = createServer()
	type McpTool = { handler: (args: unknown) => Promise<unknown> }
	const tools = (server as unknown as { _registeredTools: Record<string, McpTool> })
		._registeredTools
	const planTool = tools['ascend.plan']
	const commitTool = tools['ascend.commit']
	if (!planTool || !commitTool) throw new Error('MCP plan/commit tools are not registered')
	const call = async (handler: McpTool['handler'], body: unknown): Promise<SurfaceResponse> => {
		const start = performance.now()
		const result = await handler(body)
		const structured = (result as { readonly structuredContent?: ApiEnvelope }).structuredContent
		const payload = structured ?? { ok: false }
		const text = JSON.stringify(payload)
		const ms = performance.now() - start
		if (payload.ok !== true) throw new Error(`MCP tool failed: ${text}`)
		return { ms, text, payload }
	}
	return {
		plan: (body) => call(planTool.handler, body),
		commit: (body) => call(commitTool.handler, body),
	}
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
	mcpClient: McpMutationClient | undefined,
	inputPath: string,
	outputPath: string,
	mutations: readonly PathMutation[],
): Promise<Sample> {
	await rm(outputPath, { force: true })
	const mcpOutputPath = `${outputPath}.mcp.xlsx`
	await rm(mcpOutputPath, { force: true })
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
	const mcpPlan = mcpClient ? await mcpClient.plan({ file: inputPath, mutations }) : undefined
	const mcpCompactPlan = mcpClient
		? await mcpClient.plan({
				file: inputPath,
				mutations,
				compact: true,
				maxChangedCells: 25,
			})
		: undefined
	const mcpCommit = mcpClient
		? await mcpClient.commit({
				file: inputPath,
				output: mcpOutputPath,
				mutations,
				approvals: [],
			})
		: undefined
	const rssAfter = rssMb()
	await rm(mcpOutputPath, { force: true })
	return {
		directOpenMs: opened.ms,
		directCompileMs: compile.ms,
		directPreviewJournalMs: preview.ms,
		directApplyJournalMs: apply.ms,
		apiPlanMs: apiPlan.ms,
		apiCompactPlanMs: apiCompactPlan.ms,
		apiCommitMs: apiCommit.ms,
		...(mcpPlan ? { mcpPlanMs: mcpPlan.ms } : {}),
		...(mcpCompactPlan ? { mcpCompactPlanMs: mcpCompactPlan.ms } : {}),
		...(mcpCommit ? { mcpCommitMs: mcpCommit.ms } : {}),
		apiPlanPayloadBytes: apiPlan.text.length,
		apiCompactPlanPayloadBytes: apiCompactPlan.text.length,
		apiCommitPayloadBytes: apiCommit.text.length,
		...(mcpPlan ? { mcpPlanPayloadBytes: mcpPlan.text.length } : {}),
		...(mcpCompactPlan ? { mcpCompactPlanPayloadBytes: mcpCompactPlan.text.length } : {}),
		...(mcpCommit ? { mcpCommitPayloadBytes: mcpCommit.text.length } : {}),
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
		...(mcpCommit ? { mcpCommitOk: mcpCommit.payload.ok === true } : {}),
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
		mcpPlanMedianMs: medianOptional(samples.map((sample) => sample.mcpPlanMs)),
		mcpCompactPlanMedianMs: medianOptional(samples.map((sample) => sample.mcpCompactPlanMs)),
		mcpCommitMedianMs: medianOptional(samples.map((sample) => sample.mcpCommitMs)),
		apiPlanPayloadBytesMedian: median(samples.map((sample) => sample.apiPlanPayloadBytes)),
		apiCompactPlanPayloadBytesMedian: median(
			samples.map((sample) => sample.apiCompactPlanPayloadBytes),
		),
		apiCommitPayloadBytesMedian: median(samples.map((sample) => sample.apiCommitPayloadBytes)),
		mcpPlanPayloadBytesMedian: medianOptional(samples.map((sample) => sample.mcpPlanPayloadBytes)),
		mcpCompactPlanPayloadBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpCompactPlanPayloadBytes),
		),
		mcpCommitPayloadBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpCommitPayloadBytes),
		),
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
		mcpCommitOk: samples.some((sample) => sample.mcpCommitOk !== undefined)
			? samples.every((sample) => sample.mcpCommitOk !== false)
			: undefined,
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const outputPath = `${data.xlsxPath}.path-mutation-output.xlsx`
	const mutations = buildCellMutations(args.mutations, args.rows, args.cols)
	const apiFetch = createApiFetch()
	const mcpClient = args.surface === 'both' ? createMcpMutationClient() : undefined
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(apiFetch, mcpClient, data.xlsxPath, outputPath, mutations)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(apiFetch, mcpClient, data.xlsxPath, outputPath, mutations))
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
		await rm(`${outputPath}.mcp.xlsx`, { force: true })
	}
}

await run()
