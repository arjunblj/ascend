#!/usr/bin/env bun
import { rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { createServer } from '../../apps/mcp/src/index.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import {
	AscendWorkbook,
	configureSessionCache,
	WorkbookDocument,
} from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

type Surface = 'api' | 'both'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly rowLimit: number
	readonly mutations: number
	readonly surface: Surface
	readonly approvals?: readonly string[] | 'all'
	readonly inputFile?: string
	readonly sheet?: string
	readonly range?: string
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly sessionCacheMb?: number
	readonly json: boolean
}

interface BenchmarkInput {
	readonly xlsxPath: string
	readonly range: string
	readonly sheet: string
	readonly rows: number
	readonly cols: number
	readonly cleanup: boolean
	readonly source: 'generated' | 'input-file'
}

interface WorkflowSample {
	readonly totalMs: number
	readonly fullTotalMs: number
	readonly preparedTotalMs: number
	readonly commitVerifiedTotalMs: number
	readonly fullCommitVerifiedTotalMs: number
	readonly preparedCommitVerifiedTotalMs: number
	readonly measuredSampleMs: number
	readonly inspectMs: number
	readonly readMs: number
	readonly planMs: number
	readonly fullPlanMs: number
	readonly preparedPlanMs: number
	readonly commitMs: number
	readonly preparedCommitMs: number
	readonly commitWritePolicySnapshotMs: number | null
	readonly commitPackageGraphMs: number | null
	readonly commitPackageGraphAuditMs: number | null
	readonly commitApplyMs: number | null
	readonly commitWritePlanSummaryMs: number | null
	readonly commitWritePolicyCheckMs: number | null
	readonly commitWritePolicyBuildMs: number | null
	readonly commitToBytesMs: number | null
	readonly commitWriteFileMs: number | null
	readonly commitOutputByteReadMs: number | null
	readonly commitOutputHashMs: number | null
	readonly preparedCommitApplyMs: number | null
	readonly preparedCommitWritePlanSummaryMs: number | null
	readonly preparedCommitWritePolicyCheckMs: number | null
	readonly preparedCommitWritePolicyBuildMs: number | null
	readonly preparedCommitToBytesMs: number | null
	readonly preparedCommitWriteFileMs: number | null
	readonly preparedCommitOutputByteReadMs: number | null
	readonly preparedCommitOutputHashMs: number | null
	readonly preparedCommitWritePolicySnapshotMs: number | null
	readonly preparedCommitPackageGraphMs: number | null
	readonly preparedCommitPackageGraphAuditMs: number | null
	readonly commitPostWriteReopenMs: number | null
	readonly commitPostWriteCheckMs: number | null
	readonly commitPostWriteLintMs: number | null
	readonly commitPostWritePreservationMs: number | null
	readonly commitPostWritePackageGraphMs: number | null
	readonly commitPostWritePackageGraphAuditMs: number | null
	readonly preparedCommitPostWriteReopenMs: number | null
	readonly preparedCommitPostWriteCheckMs: number | null
	readonly preparedCommitPostWriteLintMs: number | null
	readonly preparedCommitPostWritePreservationMs: number | null
	readonly preparedCommitPostWritePackageGraphMs: number | null
	readonly preparedCommitPostWritePackageGraphAuditMs: number | null
	readonly verifyMs: number
	readonly preparedVerifyMs: number
	readonly payloadBytes: number
	readonly fullPayloadBytes: number
	readonly preparedPayloadBytes: number
	readonly commitVerifiedPayloadBytes: number
	readonly fullCommitVerifiedPayloadBytes: number
	readonly preparedCommitVerifiedPayloadBytes: number
	readonly inspectPayloadBytes: number
	readonly readPayloadBytes: number
	readonly planPayloadBytes: number
	readonly fullPlanPayloadBytes: number
	readonly preparedPlanPayloadBytes: number
	readonly commitPayloadBytes: number
	readonly preparedCommitPayloadBytes: number
	readonly commitOutputBytes: number
	readonly preparedCommitOutputBytes: number
	readonly verifyPayloadBytes: number
	readonly preparedVerifyPayloadBytes: number
	readonly readCells: number
	readonly readPartial: boolean
	readonly readWindowRows: number | null
	readonly planChangedCellCount: number | null
	readonly planEmittedChangedCellCount: number | null
	readonly preparedPlanChangedCellCount: number | null
	readonly preparedPlanEmittedChangedCellCount: number | null
	readonly compactHydratedOpenCount: number
	readonly commitVerifiedHydratedOpenCount: number
	readonly fullCommitVerifiedHydratedOpenCount: number
	readonly preparedCommitVerifiedHydratedOpenCount: number
	readonly fullHydratedOpenCount: number
	readonly preparedHydratedOpenCount: number
	readonly mcpTotalMs?: number
	readonly mcpPreparedTotalMs?: number
	readonly mcpCommitVerifiedTotalMs?: number
	readonly mcpPreparedCommitVerifiedTotalMs?: number
	readonly mcpInspectMs?: number
	readonly mcpReadMs?: number
	readonly mcpPlanMs?: number
	readonly mcpPreparedPlanMs?: number
	readonly mcpCommitMs?: number
	readonly mcpPreparedCommitMs?: number
	readonly mcpCommitPackageGraphMs?: number | null
	readonly mcpCommitPackageGraphAuditMs?: number | null
	readonly mcpCommitApplyMs?: number | null
	readonly mcpCommitWritePlanSummaryMs?: number | null
	readonly mcpCommitWritePolicyCheckMs?: number | null
	readonly mcpCommitWritePolicyBuildMs?: number | null
	readonly mcpCommitToBytesMs?: number | null
	readonly mcpCommitWriteFileMs?: number | null
	readonly mcpCommitOutputByteReadMs?: number | null
	readonly mcpCommitOutputHashMs?: number | null
	readonly mcpPreparedCommitApplyMs?: number | null
	readonly mcpPreparedCommitWritePlanSummaryMs?: number | null
	readonly mcpPreparedCommitWritePolicyCheckMs?: number | null
	readonly mcpPreparedCommitWritePolicyBuildMs?: number | null
	readonly mcpPreparedCommitToBytesMs?: number | null
	readonly mcpPreparedCommitWriteFileMs?: number | null
	readonly mcpPreparedCommitOutputByteReadMs?: number | null
	readonly mcpPreparedCommitOutputHashMs?: number | null
	readonly mcpPreparedCommitPackageGraphMs?: number | null
	readonly mcpPreparedCommitPackageGraphAuditMs?: number | null
	readonly mcpVerifyMs?: number
	readonly mcpPreparedVerifyMs?: number
	readonly mcpPayloadBytes?: number
	readonly mcpPreparedPayloadBytes?: number
	readonly mcpCommitVerifiedPayloadBytes?: number
	readonly mcpPreparedCommitVerifiedPayloadBytes?: number
	readonly mcpCommitOutputBytes?: number
	readonly mcpPreparedCommitOutputBytes?: number
	readonly mcpReadCells?: number
	readonly mcpReadPartial?: boolean
	readonly mcpCompactHydratedOpenCount?: number
	readonly mcpCommitVerifiedHydratedOpenCount?: number
	readonly mcpPreparedHydratedOpenCount?: number
	readonly mcpPreparedCommitVerifiedHydratedOpenCount?: number
	readonly mcpValid?: boolean
	readonly mcpPreparedValid?: boolean
	readonly planHydratedOpenCount: number
	readonly preparedPlanHydratedOpenCount: number
	readonly commitHydratedOpenCount: number
	readonly preparedCommitHydratedOpenCount: number
	readonly documentCacheHitCount: number
	readonly mutationCount: number
	readonly rssDeltaMb: number
	readonly valid: boolean
	readonly preparedValid: boolean
}

interface OpenStats {
	readonly workbookOpenCalls: number
	readonly workbookOpenSourceBytesCalls: number
	readonly workbookHydrations: number
	readonly documentOpenCalls: number
	readonly documentHydrations: number
	readonly documentCacheHits: number
}

type MutableOpenStats = {
	-readonly [K in keyof OpenStats]: OpenStats[K]
}

const openStats: MutableOpenStats = {
	workbookOpenCalls: 0,
	workbookOpenSourceBytesCalls: 0,
	workbookHydrations: 0,
	documentOpenCalls: 0,
	documentHydrations: 0,
	documentCacheHits: 0,
}

const seenDocuments = new WeakSet<WorkbookDocument>()

function installOpenStatsInstrumentation(): void {
	type WorkbookOpen = typeof AscendWorkbook.open
	type WorkbookOpenSourceBytes = typeof AscendWorkbook.openSourceBytes
	type DocumentOpen = typeof WorkbookDocument.open

	const originalWorkbookOpen = AscendWorkbook.open.bind(AscendWorkbook) as WorkbookOpen
	const originalWorkbookOpenSourceBytes = AscendWorkbook.openSourceBytes.bind(
		AscendWorkbook,
	) as WorkbookOpenSourceBytes
	const originalDocumentOpen = WorkbookDocument.open.bind(WorkbookDocument) as DocumentOpen

	Object.defineProperty(AscendWorkbook, 'open', {
		configurable: true,
		value: (async (...args: Parameters<WorkbookOpen>) => {
			openStats.workbookOpenCalls += 1
			openStats.workbookHydrations += 1
			return originalWorkbookOpen(...args)
		}) satisfies WorkbookOpen,
	})
	Object.defineProperty(AscendWorkbook, 'openSourceBytes', {
		configurable: true,
		value: (async (...args: Parameters<WorkbookOpenSourceBytes>) => {
			openStats.workbookOpenSourceBytesCalls += 1
			openStats.workbookHydrations += 1
			return originalWorkbookOpenSourceBytes(...args)
		}) satisfies WorkbookOpenSourceBytes,
	})
	Object.defineProperty(WorkbookDocument, 'open', {
		configurable: true,
		value: (async (...args: Parameters<DocumentOpen>) => {
			openStats.documentOpenCalls += 1
			const document = await originalDocumentOpen(...args)
			if (seenDocuments.has(document)) openStats.documentCacheHits += 1
			else {
				seenDocuments.add(document)
				openStats.documentHydrations += 1
			}
			return document
		}) satisfies DocumentOpen,
	})
}

function snapshotOpenStats(): OpenStats {
	return { ...openStats }
}

function diffOpenStats(after: OpenStats, before: OpenStats): OpenStats {
	return {
		workbookOpenCalls: after.workbookOpenCalls - before.workbookOpenCalls,
		workbookOpenSourceBytesCalls:
			after.workbookOpenSourceBytesCalls - before.workbookOpenSourceBytesCalls,
		workbookHydrations: after.workbookHydrations - before.workbookHydrations,
		documentOpenCalls: after.documentOpenCalls - before.documentOpenCalls,
		documentHydrations: after.documentHydrations - before.documentHydrations,
		documentCacheHits: after.documentCacheHits - before.documentCacheHits,
	}
}

function addOpenStats(...items: readonly OpenStats[]): OpenStats {
	return items.reduce<OpenStats>(
		(total, item) => ({
			workbookOpenCalls: total.workbookOpenCalls + item.workbookOpenCalls,
			workbookOpenSourceBytesCalls:
				total.workbookOpenSourceBytesCalls + item.workbookOpenSourceBytesCalls,
			workbookHydrations: total.workbookHydrations + item.workbookHydrations,
			documentOpenCalls: total.documentOpenCalls + item.documentOpenCalls,
			documentHydrations: total.documentHydrations + item.documentHydrations,
			documentCacheHits: total.documentCacheHits + item.documentCacheHits,
		}),
		{
			workbookOpenCalls: 0,
			workbookOpenSourceBytesCalls: 0,
			workbookHydrations: 0,
			documentOpenCalls: 0,
			documentHydrations: 0,
			documentCacheHits: 0,
		},
	)
}

function hydratedOpenCount(stats: OpenStats): number {
	return stats.workbookHydrations + stats.documentHydrations
}

installOpenStatsInstrumentation()

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
	const inputFile = readOption(process.argv, '--input-file')
	const sheet = readOption(process.argv, '--sheet')
	const range = readOption(process.argv, '--range')
	const approvals = parseApprovals(readOption(process.argv, '--approval'))
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		rowLimit: positiveInt(readOption(process.argv, '--row-limit'), 500),
		mutations: positiveInt(readOption(process.argv, '--mutations'), 1),
		surface,
		...(approvals !== undefined ? { approvals } : {}),
		...(inputFile !== undefined ? { inputFile } : {}),
		...(sheet !== undefined ? { sheet } : {}),
		...(range !== undefined ? { range } : {}),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(process.argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(process.argv, '--warmup'), 1),
		...(readOption(process.argv, '--session-cache-mb') !== undefined
			? {
					sessionCacheMb: positiveInt(readOption(process.argv, '--session-cache-mb'), 32),
				}
			: {}),
		json: hasFlag(process.argv, '--json'),
	}
}

function parseApprovals(raw: string | undefined): readonly string[] | 'all' | undefined {
	if (raw === undefined) return undefined
	const entries = raw
		.split(',')
		.map((entry) => entry.trim())
		.filter((entry) => entry.length > 0)
	if (entries.length === 0) return undefined
	return entries.some((entry) => entry.toLowerCase() === 'all') ? 'all' : entries
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

function optionalSeriesStats(values: readonly (number | null | undefined)[]) {
	const defined = values.filter((value): value is number => typeof value === 'number')
	return defined.length > 0 ? seriesStats(defined) : undefined
}

async function post(apiFetch: typeof fetch, path: string, body: unknown) {
	const beforeOpenStats = snapshotOpenStats()
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
	return {
		ms,
		text,
		payload: JSON.parse(text) as ApiEnvelope,
		openStats: diffOpenStats(snapshotOpenStats(), beforeOpenStats),
	}
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
		readonly preparedPlan?: { readonly id?: string }
		readonly timings?: {
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
		readonly postWrite?: {
			readonly timings?: {
				readonly reopenMs?: number
				readonly checkMs?: number
				readonly lintMs?: number
				readonly preservationMs?: number
				readonly packageGraphMs?: number
				readonly packageGraphAuditMs?: number
			}
		}
	}
}

interface SurfaceResponse {
	readonly ms: number
	readonly text: string
	readonly payload: ApiEnvelope
	readonly openStats: OpenStats
}

interface McpWorkflowClient {
	inspect(body: unknown): Promise<SurfaceResponse>
	read(body: unknown): Promise<SurfaceResponse>
	plan(body: unknown): Promise<SurfaceResponse>
	commit(body: unknown): Promise<SurfaceResponse>
	check(body: unknown): Promise<SurfaceResponse>
}

function createMcpWorkflowClient(): McpWorkflowClient {
	const server = createServer()
	type McpTool = { handler: (args: unknown) => Promise<unknown> }
	const tools = (server as unknown as { _registeredTools: Record<string, McpTool> })
		._registeredTools
	const required = {
		inspect: tools['ascend.inspect'],
		read: tools['ascend.read'],
		plan: tools['ascend.plan'],
		commit: tools['ascend.commit'],
		check: tools['ascend.check'],
	}
	for (const [name, tool] of Object.entries(required)) {
		if (!tool) throw new Error(`MCP tool ${name} is not registered`)
	}
	const call = async (handler: McpTool['handler'], body: unknown): Promise<SurfaceResponse> => {
		const beforeOpenStats = snapshotOpenStats()
		const start = performance.now()
		const result = await handler(body)
		const structured = (result as { readonly structuredContent?: ApiEnvelope }).structuredContent
		const payload = structured ?? {}
		const text = JSON.stringify(payload)
		const ms = performance.now() - start
		if ((payload as { readonly ok?: boolean }).ok !== true) {
			throw new Error(`MCP tool failed: ${text}`)
		}
		return { ms, text, payload, openStats: diffOpenStats(snapshotOpenStats(), beforeOpenStats) }
	}
	return {
		inspect: (body) => call(required.inspect.handler, body),
		read: (body) => call(required.read.handler, body),
		plan: (body) => call(required.plan.handler, body),
		commit: (body) => call(required.commit.handler, body),
		check: (body) => call(required.check.handler, body),
	}
}

interface PostWriteTimings {
	readonly reopenMs: number | null
	readonly checkMs: number | null
	readonly lintMs: number | null
	readonly preservationMs: number | null
	readonly packageGraphMs: number | null
	readonly packageGraphAuditMs: number | null
}

interface CommitTimings {
	readonly writePolicySnapshotMs: number | null
	readonly packageGraphMs: number | null
	readonly packageGraphAuditMs: number | null
	readonly applyMs: number | null
	readonly writePlanSummaryMs: number | null
	readonly writePolicyCheckMs: number | null
	readonly writePolicyBuildMs: number | null
	readonly toBytesMs: number | null
	readonly writeFileMs: number | null
	readonly outputByteReadMs: number | null
	readonly outputHashMs: number | null
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function rssMb(): number {
	return process.memoryUsage().rss / (1024 * 1024)
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

function pointerSegment(value: string): string {
	return encodeURIComponent(value.replace(/~/g, '~0').replace(/\//g, '~1'))
}

function postWriteTimings(payload: ApiEnvelope): PostWriteTimings {
	const timings = payload.data?.postWrite?.timings
	return {
		reopenMs: timings?.reopenMs ?? null,
		checkMs: timings?.checkMs ?? null,
		lintMs: timings?.lintMs ?? null,
		preservationMs: timings?.preservationMs ?? null,
		packageGraphMs: timings?.packageGraphMs ?? null,
		packageGraphAuditMs: timings?.packageGraphAuditMs ?? null,
	}
}

function commitTimings(payload: ApiEnvelope): CommitTimings {
	const timings = payload.data?.timings
	return {
		writePolicySnapshotMs: timings?.writePolicySnapshotMs ?? null,
		packageGraphMs: timings?.packageGraphMs ?? null,
		packageGraphAuditMs: timings?.packageGraphAuditMs ?? null,
		applyMs: timings?.applyMs ?? null,
		writePlanSummaryMs: timings?.writePlanSummaryMs ?? null,
		writePolicyCheckMs: timings?.writePolicyCheckMs ?? null,
		writePolicyBuildMs: timings?.writePolicyBuildMs ?? null,
		toBytesMs: timings?.toBytesMs ?? null,
		writeFileMs: timings?.writeFileMs ?? null,
		outputByteReadMs: timings?.outputByteReadMs ?? null,
		outputHashMs: timings?.outputHashMs ?? null,
	}
}

async function runMcpWorkflow(
	mcpClient: McpWorkflowClient,
	inputPath: string,
	outputPath: string,
	sheetName: string,
	range: string,
	rowLimit: number,
	mutations: readonly ReturnType<typeof buildCellMutations>[number][],
	approvals: readonly string[] | 'all',
): Promise<Partial<WorkflowSample>> {
	const preparedOutputPath = `${outputPath}.prepared.xlsx`
	await rm(outputPath, { force: true })
	await rm(preparedOutputPath, { force: true })
	const inspect = await mcpClient.inspect({ file: inputPath })
	const read = await mcpClient.read({
		file: inputPath,
		sheet: sheetName,
		range,
		format: 'compact',
		rowLimit,
	})
	const plan = await mcpClient.plan({
		file: inputPath,
		mutations,
		compact: true,
		prepare: false,
		maxChangedCells: 25,
	})
	const preparedPlan = await mcpClient.plan({
		file: inputPath,
		mutations,
		compact: true,
		maxChangedCells: 25,
	})
	const planHandle = preparedPlan.payload.data?.preparedPlan?.id
	if (!planHandle) throw new Error('MCP prepared plan did not return a handle')
	const commit = await mcpClient.commit({
		file: inputPath,
		output: outputPath,
		mutations,
		approvals,
		compact: true,
	})
	const preparedCommit = await mcpClient.commit({
		planHandle,
		output: preparedOutputPath,
		approvals,
		compact: true,
	})
	const verify = await mcpClient.check({ file: outputPath })
	const preparedVerify = await mcpClient.check({ file: preparedOutputPath })
	const commitOutputBytes = (await stat(outputPath)).size
	const preparedCommitOutputBytes = (await stat(preparedOutputPath)).size
	const sharedOpenStats = addOpenStats(inspect.openStats, read.openStats)
	const compactOpenStats = addOpenStats(
		sharedOpenStats,
		plan.openStats,
		commit.openStats,
		verify.openStats,
	)
	const compactCommitVerifiedOpenStats = addOpenStats(
		sharedOpenStats,
		plan.openStats,
		commit.openStats,
	)
	const preparedOpenStats = addOpenStats(
		sharedOpenStats,
		preparedPlan.openStats,
		preparedCommit.openStats,
		preparedVerify.openStats,
	)
	const preparedCommitVerifiedOpenStats = addOpenStats(
		sharedOpenStats,
		preparedPlan.openStats,
		preparedCommit.openStats,
	)
	const workflowBytes =
		inspect.text.length +
		read.text.length +
		plan.text.length +
		commit.text.length +
		verify.text.length
	const preparedWorkflowBytes =
		inspect.text.length +
		read.text.length +
		preparedPlan.text.length +
		preparedCommit.text.length +
		preparedVerify.text.length
	const mcpCommitTimings = commitTimings(commit.payload)
	const mcpPreparedCommitTimings = commitTimings(preparedCommit.payload)
	return {
		mcpTotalMs: inspect.ms + read.ms + plan.ms + commit.ms + verify.ms,
		mcpPreparedTotalMs:
			inspect.ms + read.ms + preparedPlan.ms + preparedCommit.ms + preparedVerify.ms,
		mcpCommitVerifiedTotalMs: inspect.ms + read.ms + plan.ms + commit.ms,
		mcpPreparedCommitVerifiedTotalMs: inspect.ms + read.ms + preparedPlan.ms + preparedCommit.ms,
		mcpInspectMs: inspect.ms,
		mcpReadMs: read.ms,
		mcpPlanMs: plan.ms,
		mcpPreparedPlanMs: preparedPlan.ms,
		mcpCommitMs: commit.ms,
		mcpPreparedCommitMs: preparedCommit.ms,
		mcpCommitPackageGraphMs: mcpCommitTimings.packageGraphMs,
		mcpCommitPackageGraphAuditMs: mcpCommitTimings.packageGraphAuditMs,
		mcpCommitApplyMs: mcpCommitTimings.applyMs,
		mcpCommitWritePlanSummaryMs: mcpCommitTimings.writePlanSummaryMs,
		mcpCommitWritePolicyCheckMs: mcpCommitTimings.writePolicyCheckMs,
		mcpCommitWritePolicyBuildMs: mcpCommitTimings.writePolicyBuildMs,
		mcpCommitToBytesMs: mcpCommitTimings.toBytesMs,
		mcpCommitWriteFileMs: mcpCommitTimings.writeFileMs,
		mcpCommitOutputByteReadMs: mcpCommitTimings.outputByteReadMs,
		mcpCommitOutputHashMs: mcpCommitTimings.outputHashMs,
		mcpPreparedCommitApplyMs: mcpPreparedCommitTimings.applyMs,
		mcpPreparedCommitWritePlanSummaryMs: mcpPreparedCommitTimings.writePlanSummaryMs,
		mcpPreparedCommitWritePolicyCheckMs: mcpPreparedCommitTimings.writePolicyCheckMs,
		mcpPreparedCommitWritePolicyBuildMs: mcpPreparedCommitTimings.writePolicyBuildMs,
		mcpPreparedCommitToBytesMs: mcpPreparedCommitTimings.toBytesMs,
		mcpPreparedCommitWriteFileMs: mcpPreparedCommitTimings.writeFileMs,
		mcpPreparedCommitOutputByteReadMs: mcpPreparedCommitTimings.outputByteReadMs,
		mcpPreparedCommitOutputHashMs: mcpPreparedCommitTimings.outputHashMs,
		mcpPreparedCommitPackageGraphMs: mcpPreparedCommitTimings.packageGraphMs,
		mcpPreparedCommitPackageGraphAuditMs: mcpPreparedCommitTimings.packageGraphAuditMs,
		mcpVerifyMs: verify.ms,
		mcpPreparedVerifyMs: preparedVerify.ms,
		mcpPayloadBytes: workflowBytes,
		mcpPreparedPayloadBytes: preparedWorkflowBytes,
		mcpCommitVerifiedPayloadBytes:
			inspect.text.length + read.text.length + plan.text.length + commit.text.length,
		mcpPreparedCommitVerifiedPayloadBytes:
			inspect.text.length +
			read.text.length +
			preparedPlan.text.length +
			preparedCommit.text.length,
		mcpCommitOutputBytes: commitOutputBytes,
		mcpPreparedCommitOutputBytes: preparedCommitOutputBytes,
		mcpReadCells: read.payload.data?.cells?.length ?? 0,
		mcpReadPartial: read.payload.data?.load?.isPartial === true,
		mcpCompactHydratedOpenCount: hydratedOpenCount(compactOpenStats),
		mcpCommitVerifiedHydratedOpenCount: hydratedOpenCount(compactCommitVerifiedOpenStats),
		mcpPreparedHydratedOpenCount: hydratedOpenCount(preparedOpenStats),
		mcpPreparedCommitVerifiedHydratedOpenCount: hydratedOpenCount(preparedCommitVerifiedOpenStats),
		mcpValid: verify.payload.data?.valid === true,
		mcpPreparedValid: preparedVerify.payload.data?.valid === true,
	}
}

async function runWorkflow(
	apiFetch: typeof fetch,
	mcpClient: McpWorkflowClient | undefined,
	inputPath: string,
	outputPath: string,
	sheetName: string,
	range: string,
	rowLimit: number,
	mutationCount: number,
	rows: number,
	cols: number,
	approvals: readonly string[] | 'all' = [],
): Promise<WorkflowSample> {
	await rm(outputPath, { force: true })
	const preparedOutputPath = `${outputPath}.prepared.xlsx`
	await rm(preparedOutputPath, { force: true })
	runGc()
	const rssBefore = rssMb()
	const totalStart = performance.now()
	const inspect = await post(apiFetch, '/inspect', { file: inputPath })
	const read = await post(apiFetch, '/read', {
		file: inputPath,
		sheet: sheetName,
		range,
		format: 'compact',
		rowLimit,
	})
	const mutations = buildCellMutations(mutationCount, rows, cols, sheetName)
	const fullPlan = await post(apiFetch, '/plan', { file: inputPath, mutations, prepare: false })
	const plan = await post(apiFetch, '/plan', {
		file: inputPath,
		mutations,
		compact: true,
		prepare: false,
		maxChangedCells: 25,
	})
	const preparedPlan = await post(apiFetch, '/plan', {
		file: inputPath,
		mutations,
		compact: true,
		maxChangedCells: 25,
	})
	const planHandle = preparedPlan.payload.data?.preparedPlan?.id
	if (!planHandle) throw new Error('Prepared plan did not return a handle')
	const commit = await post(apiFetch, '/commit', {
		file: inputPath,
		output: outputPath,
		mutations,
		approvals,
		compact: true,
	})
	const preparedCommit = await post(apiFetch, '/commit', {
		planHandle,
		output: preparedOutputPath,
		approvals,
		compact: true,
	})
	const verify = await post(apiFetch, '/check', { file: outputPath })
	const preparedVerify = await post(apiFetch, '/check', { file: preparedOutputPath })
	const commitOutputBytes = (await stat(outputPath)).size
	const preparedCommitOutputBytes = (await stat(preparedOutputPath)).size
	const measuredSampleMs = performance.now() - totalStart
	const mcpOutputPath = `${outputPath}.mcp.xlsx`
	const mcpWorkflow = mcpClient
		? await runMcpWorkflow(
				mcpClient,
				inputPath,
				mcpOutputPath,
				sheetName,
				range,
				rowLimit,
				mutations,
				approvals,
			)
		: {}
	const rssAfter = rssMb()
	const readLoad = read.payload.data?.load
	const commitTiming = commitTimings(commit.payload)
	const preparedCommitTiming = commitTimings(preparedCommit.payload)
	const commitPostWrite = postWriteTimings(commit.payload)
	const preparedCommitPostWrite = postWriteTimings(preparedCommit.payload)
	const sharedOpenStats = addOpenStats(inspect.openStats, read.openStats)
	const compactOpenStats = addOpenStats(
		sharedOpenStats,
		plan.openStats,
		commit.openStats,
		verify.openStats,
	)
	const compactCommitVerifiedOpenStats = addOpenStats(
		sharedOpenStats,
		plan.openStats,
		commit.openStats,
	)
	const fullOpenStats = addOpenStats(
		sharedOpenStats,
		fullPlan.openStats,
		commit.openStats,
		verify.openStats,
	)
	const fullCommitVerifiedOpenStats = addOpenStats(
		sharedOpenStats,
		fullPlan.openStats,
		commit.openStats,
	)
	const preparedOpenStats = addOpenStats(
		sharedOpenStats,
		preparedPlan.openStats,
		preparedCommit.openStats,
		preparedVerify.openStats,
	)
	const preparedCommitVerifiedOpenStats = addOpenStats(
		sharedOpenStats,
		preparedPlan.openStats,
		preparedCommit.openStats,
	)
	const compactWorkflowBytes =
		inspect.text.length +
		read.text.length +
		plan.text.length +
		commit.text.length +
		verify.text.length
	const compactCommitVerifiedWorkflowBytes =
		inspect.text.length + read.text.length + plan.text.length + commit.text.length
	const fullWorkflowBytes =
		inspect.text.length +
		read.text.length +
		fullPlan.text.length +
		commit.text.length +
		verify.text.length
	const fullCommitVerifiedWorkflowBytes =
		inspect.text.length + read.text.length + fullPlan.text.length + commit.text.length
	const preparedWorkflowBytes =
		inspect.text.length +
		read.text.length +
		preparedPlan.text.length +
		preparedCommit.text.length +
		preparedVerify.text.length
	const preparedCommitVerifiedWorkflowBytes =
		inspect.text.length + read.text.length + preparedPlan.text.length + preparedCommit.text.length
	return {
		totalMs: inspect.ms + read.ms + plan.ms + commit.ms + verify.ms,
		fullTotalMs: inspect.ms + read.ms + fullPlan.ms + commit.ms + verify.ms,
		preparedTotalMs: inspect.ms + read.ms + preparedPlan.ms + preparedCommit.ms + preparedVerify.ms,
		commitVerifiedTotalMs: inspect.ms + read.ms + plan.ms + commit.ms,
		fullCommitVerifiedTotalMs: inspect.ms + read.ms + fullPlan.ms + commit.ms,
		preparedCommitVerifiedTotalMs: inspect.ms + read.ms + preparedPlan.ms + preparedCommit.ms,
		measuredSampleMs,
		inspectMs: inspect.ms,
		readMs: read.ms,
		planMs: plan.ms,
		fullPlanMs: fullPlan.ms,
		preparedPlanMs: preparedPlan.ms,
		commitMs: commit.ms,
		preparedCommitMs: preparedCommit.ms,
		commitWritePolicySnapshotMs: commitTiming.writePolicySnapshotMs,
		commitPackageGraphMs: commitTiming.packageGraphMs,
		commitPackageGraphAuditMs: commitTiming.packageGraphAuditMs,
		commitApplyMs: commitTiming.applyMs,
		commitWritePlanSummaryMs: commitTiming.writePlanSummaryMs,
		commitWritePolicyCheckMs: commitTiming.writePolicyCheckMs,
		commitWritePolicyBuildMs: commitTiming.writePolicyBuildMs,
		commitToBytesMs: commitTiming.toBytesMs,
		commitWriteFileMs: commitTiming.writeFileMs,
		commitOutputByteReadMs: commitTiming.outputByteReadMs,
		commitOutputHashMs: commitTiming.outputHashMs,
		preparedCommitApplyMs: preparedCommitTiming.applyMs,
		preparedCommitWritePlanSummaryMs: preparedCommitTiming.writePlanSummaryMs,
		preparedCommitWritePolicyCheckMs: preparedCommitTiming.writePolicyCheckMs,
		preparedCommitWritePolicyBuildMs: preparedCommitTiming.writePolicyBuildMs,
		preparedCommitToBytesMs: preparedCommitTiming.toBytesMs,
		preparedCommitWriteFileMs: preparedCommitTiming.writeFileMs,
		preparedCommitOutputByteReadMs: preparedCommitTiming.outputByteReadMs,
		preparedCommitOutputHashMs: preparedCommitTiming.outputHashMs,
		preparedCommitWritePolicySnapshotMs: preparedCommitTiming.writePolicySnapshotMs,
		preparedCommitPackageGraphMs: preparedCommitTiming.packageGraphMs,
		preparedCommitPackageGraphAuditMs: preparedCommitTiming.packageGraphAuditMs,
		commitPostWriteReopenMs: commitPostWrite.reopenMs,
		commitPostWriteCheckMs: commitPostWrite.checkMs,
		commitPostWriteLintMs: commitPostWrite.lintMs,
		commitPostWritePreservationMs: commitPostWrite.preservationMs,
		commitPostWritePackageGraphMs: commitPostWrite.packageGraphMs,
		commitPostWritePackageGraphAuditMs: commitPostWrite.packageGraphAuditMs,
		preparedCommitPostWriteReopenMs: preparedCommitPostWrite.reopenMs,
		preparedCommitPostWriteCheckMs: preparedCommitPostWrite.checkMs,
		preparedCommitPostWriteLintMs: preparedCommitPostWrite.lintMs,
		preparedCommitPostWritePreservationMs: preparedCommitPostWrite.preservationMs,
		preparedCommitPostWritePackageGraphMs: preparedCommitPostWrite.packageGraphMs,
		preparedCommitPostWritePackageGraphAuditMs: preparedCommitPostWrite.packageGraphAuditMs,
		verifyMs: verify.ms,
		preparedVerifyMs: preparedVerify.ms,
		payloadBytes: compactWorkflowBytes,
		fullPayloadBytes: fullWorkflowBytes,
		preparedPayloadBytes: preparedWorkflowBytes,
		commitVerifiedPayloadBytes: compactCommitVerifiedWorkflowBytes,
		fullCommitVerifiedPayloadBytes: fullCommitVerifiedWorkflowBytes,
		preparedCommitVerifiedPayloadBytes: preparedCommitVerifiedWorkflowBytes,
		inspectPayloadBytes: inspect.text.length,
		readPayloadBytes: read.text.length,
		planPayloadBytes: plan.text.length,
		fullPlanPayloadBytes: fullPlan.text.length,
		preparedPlanPayloadBytes: preparedPlan.text.length,
		commitPayloadBytes: commit.text.length,
		preparedCommitPayloadBytes: preparedCommit.text.length,
		commitOutputBytes,
		preparedCommitOutputBytes,
		verifyPayloadBytes: verify.text.length,
		preparedVerifyPayloadBytes: preparedVerify.text.length,
		readCells: read.payload.data?.cells?.length ?? 0,
		readPartial: readLoad?.isPartial === true,
		readWindowRows: readLoad?.loadedSheets ? rowLimit : null,
		planChangedCellCount: plan.payload.data?.preview?.changedCellCount ?? null,
		planEmittedChangedCellCount: plan.payload.data?.preview?.emittedChangedCellCount ?? null,
		preparedPlanChangedCellCount: preparedPlan.payload.data?.preview?.changedCellCount ?? null,
		preparedPlanEmittedChangedCellCount:
			preparedPlan.payload.data?.preview?.emittedChangedCellCount ?? null,
		compactHydratedOpenCount: hydratedOpenCount(compactOpenStats),
		commitVerifiedHydratedOpenCount: hydratedOpenCount(compactCommitVerifiedOpenStats),
		fullCommitVerifiedHydratedOpenCount: hydratedOpenCount(fullCommitVerifiedOpenStats),
		preparedCommitVerifiedHydratedOpenCount: hydratedOpenCount(preparedCommitVerifiedOpenStats),
		fullHydratedOpenCount: hydratedOpenCount(fullOpenStats),
		preparedHydratedOpenCount: hydratedOpenCount(preparedOpenStats),
		planHydratedOpenCount: hydratedOpenCount(plan.openStats),
		preparedPlanHydratedOpenCount: hydratedOpenCount(preparedPlan.openStats),
		commitHydratedOpenCount: hydratedOpenCount(commit.openStats),
		preparedCommitHydratedOpenCount: hydratedOpenCount(preparedCommit.openStats),
		documentCacheHitCount: addOpenStats(
			inspect.openStats,
			read.openStats,
			fullPlan.openStats,
			plan.openStats,
			preparedPlan.openStats,
			commit.openStats,
			preparedCommit.openStats,
			verify.openStats,
			preparedVerify.openStats,
		).documentCacheHits,
		...mcpWorkflow,
		mutationCount,
		rssDeltaMb: rssAfter - rssBefore,
		valid: verify.payload.data?.valid === true,
		preparedValid: preparedVerify.payload.data?.valid === true,
	}
}

function summarize(samples: readonly WorkflowSample[]) {
	return {
		totalMedianMs: median(samples.map((sample) => sample.totalMs)),
		fullTotalMedianMs: median(samples.map((sample) => sample.fullTotalMs)),
		preparedTotalMedianMs: median(samples.map((sample) => sample.preparedTotalMs)),
		commitVerifiedTotalMedianMs: median(samples.map((sample) => sample.commitVerifiedTotalMs)),
		fullCommitVerifiedTotalMedianMs: median(
			samples.map((sample) => sample.fullCommitVerifiedTotalMs),
		),
		preparedCommitVerifiedTotalMedianMs: median(
			samples.map((sample) => sample.preparedCommitVerifiedTotalMs),
		),
		preparedCommitVerifiedTotalStats: seriesStats(
			samples.map((sample) => sample.preparedCommitVerifiedTotalMs),
		),
		measuredSampleMedianMs: median(samples.map((sample) => sample.measuredSampleMs)),
		inspectMedianMs: median(samples.map((sample) => sample.inspectMs)),
		readMedianMs: median(samples.map((sample) => sample.readMs)),
		planMedianMs: median(samples.map((sample) => sample.planMs)),
		fullPlanMedianMs: median(samples.map((sample) => sample.fullPlanMs)),
		preparedPlanMedianMs: median(samples.map((sample) => sample.preparedPlanMs)),
		preparedPlanStats: seriesStats(samples.map((sample) => sample.preparedPlanMs)),
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		preparedCommitMedianMs: median(samples.map((sample) => sample.preparedCommitMs)),
		commitWritePolicySnapshotMedianMs: medianOptional(
			samples.map((sample) => sample.commitWritePolicySnapshotMs),
		),
		commitPackageGraphMedianMs: medianOptional(
			samples.map((sample) => sample.commitPackageGraphMs),
		),
		commitPackageGraphAuditMedianMs: medianOptional(
			samples.map((sample) => sample.commitPackageGraphAuditMs),
		),
		commitApplyMedianMs: medianOptional(samples.map((sample) => sample.commitApplyMs)),
		commitWritePlanSummaryMedianMs: medianOptional(
			samples.map((sample) => sample.commitWritePlanSummaryMs),
		),
		commitWritePolicyCheckMedianMs: medianOptional(
			samples.map((sample) => sample.commitWritePolicyCheckMs),
		),
		commitWritePolicyBuildMedianMs: medianOptional(
			samples.map((sample) => sample.commitWritePolicyBuildMs),
		),
		commitToBytesMedianMs: medianOptional(samples.map((sample) => sample.commitToBytesMs)),
		commitWriteFileMedianMs: medianOptional(samples.map((sample) => sample.commitWriteFileMs)),
		commitOutputByteReadMedianMs: medianOptional(
			samples.map((sample) => sample.commitOutputByteReadMs),
		),
		commitOutputHashMedianMs: medianOptional(samples.map((sample) => sample.commitOutputHashMs)),
		preparedCommitApplyMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitApplyMs),
		),
		preparedCommitWritePolicySnapshotMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitWritePolicySnapshotMs),
		),
		preparedCommitPackageGraphMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPackageGraphMs),
		),
		preparedCommitPackageGraphAuditMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPackageGraphAuditMs),
		),
		preparedCommitWritePlanSummaryMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitWritePlanSummaryMs),
		),
		preparedCommitWritePolicyCheckMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitWritePolicyCheckMs),
		),
		preparedCommitWritePolicyBuildMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitWritePolicyBuildMs),
		),
		preparedCommitToBytesMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitToBytesMs),
		),
		preparedCommitWriteFileMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitWriteFileMs),
		),
		preparedCommitOutputByteReadMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitOutputByteReadMs),
		),
		preparedCommitOutputHashMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitOutputHashMs),
		),
		commitPostWriteReopenMedianMs: medianOptional(
			samples.map((sample) => sample.commitPostWriteReopenMs),
		),
		commitPostWriteCheckMedianMs: medianOptional(
			samples.map((sample) => sample.commitPostWriteCheckMs),
		),
		commitPostWriteLintMedianMs: medianOptional(
			samples.map((sample) => sample.commitPostWriteLintMs),
		),
		commitPostWritePreservationMedianMs: medianOptional(
			samples.map((sample) => sample.commitPostWritePreservationMs),
		),
		commitPostWritePackageGraphMedianMs: medianOptional(
			samples.map((sample) => sample.commitPostWritePackageGraphMs),
		),
		commitPostWritePackageGraphAuditMedianMs: medianOptional(
			samples.map((sample) => sample.commitPostWritePackageGraphAuditMs),
		),
		preparedCommitPostWriteReopenMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPostWriteReopenMs),
		),
		preparedCommitPostWriteReopenStats: optionalSeriesStats(
			samples.map((sample) => sample.preparedCommitPostWriteReopenMs),
		),
		preparedCommitPostWriteCheckMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPostWriteCheckMs),
		),
		preparedCommitPostWriteLintMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPostWriteLintMs),
		),
		preparedCommitPostWritePreservationMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPostWritePreservationMs),
		),
		preparedCommitPostWritePackageGraphMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPostWritePackageGraphMs),
		),
		preparedCommitPostWritePackageGraphAuditMedianMs: medianOptional(
			samples.map((sample) => sample.preparedCommitPostWritePackageGraphAuditMs),
		),
		verifyMedianMs: median(samples.map((sample) => sample.verifyMs)),
		preparedVerifyMedianMs: median(samples.map((sample) => sample.preparedVerifyMs)),
		preparedVerifyStats: seriesStats(samples.map((sample) => sample.preparedVerifyMs)),
		payloadBytesMedian: median(samples.map((sample) => sample.payloadBytes)),
		fullPayloadBytesMedian: median(samples.map((sample) => sample.fullPayloadBytes)),
		preparedPayloadBytesMedian: median(samples.map((sample) => sample.preparedPayloadBytes)),
		commitVerifiedPayloadBytesMedian: median(
			samples.map((sample) => sample.commitVerifiedPayloadBytes),
		),
		fullCommitVerifiedPayloadBytesMedian: median(
			samples.map((sample) => sample.fullCommitVerifiedPayloadBytes),
		),
		preparedCommitVerifiedPayloadBytesMedian: median(
			samples.map((sample) => sample.preparedCommitVerifiedPayloadBytes),
		),
		inspectPayloadBytesMedian: median(samples.map((sample) => sample.inspectPayloadBytes)),
		readPayloadBytesMedian: median(samples.map((sample) => sample.readPayloadBytes)),
		planPayloadBytesMedian: median(samples.map((sample) => sample.planPayloadBytes)),
		fullPlanPayloadBytesMedian: median(samples.map((sample) => sample.fullPlanPayloadBytes)),
		preparedPlanPayloadBytesMedian: median(
			samples.map((sample) => sample.preparedPlanPayloadBytes),
		),
		commitPayloadBytesMedian: median(samples.map((sample) => sample.commitPayloadBytes)),
		preparedCommitPayloadBytesMedian: median(
			samples.map((sample) => sample.preparedCommitPayloadBytes),
		),
		commitOutputBytesMedian: median(samples.map((sample) => sample.commitOutputBytes)),
		preparedCommitOutputBytesMedian: median(
			samples.map((sample) => sample.preparedCommitOutputBytes),
		),
		verifyPayloadBytesMedian: median(samples.map((sample) => sample.verifyPayloadBytes)),
		preparedVerifyPayloadBytesMedian: median(
			samples.map((sample) => sample.preparedVerifyPayloadBytes),
		),
		compactWorkflowSpeedupVsFull:
			median(samples.map((sample) => sample.fullTotalMs)) /
			median(samples.map((sample) => sample.totalMs)),
		commitVerifiedWorkflowSpeedupVsFull:
			median(samples.map((sample) => sample.fullCommitVerifiedTotalMs)) /
			median(samples.map((sample) => sample.commitVerifiedTotalMs)),
		preparedWorkflowSpeedupVsCompact:
			median(samples.map((sample) => sample.totalMs)) /
			median(samples.map((sample) => sample.preparedTotalMs)),
		preparedCommitVerifiedWorkflowSpeedupVsCompact:
			median(samples.map((sample) => sample.commitVerifiedTotalMs)) /
			median(samples.map((sample) => sample.preparedCommitVerifiedTotalMs)),
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
		preparedPlanChangedCellCountMedian: medianOptional(
			samples.map((sample) => sample.preparedPlanChangedCellCount),
		),
		preparedPlanEmittedChangedCellCountMedian: medianOptional(
			samples.map((sample) => sample.preparedPlanEmittedChangedCellCount),
		),
		compactHydratedOpenCountMedian: median(
			samples.map((sample) => sample.compactHydratedOpenCount),
		),
		commitVerifiedHydratedOpenCountMedian: median(
			samples.map((sample) => sample.commitVerifiedHydratedOpenCount),
		),
		fullCommitVerifiedHydratedOpenCountMedian: median(
			samples.map((sample) => sample.fullCommitVerifiedHydratedOpenCount),
		),
		preparedCommitVerifiedHydratedOpenCountMedian: median(
			samples.map((sample) => sample.preparedCommitVerifiedHydratedOpenCount),
		),
		fullHydratedOpenCountMedian: median(samples.map((sample) => sample.fullHydratedOpenCount)),
		preparedHydratedOpenCountMedian: median(
			samples.map((sample) => sample.preparedHydratedOpenCount),
		),
		mcpTotalMedianMs: medianOptional(samples.map((sample) => sample.mcpTotalMs)),
		mcpPreparedTotalMedianMs: medianOptional(samples.map((sample) => sample.mcpPreparedTotalMs)),
		mcpCommitVerifiedTotalMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitVerifiedTotalMs),
		),
		mcpPreparedCommitVerifiedTotalMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitVerifiedTotalMs),
		),
		mcpInspectMedianMs: medianOptional(samples.map((sample) => sample.mcpInspectMs)),
		mcpReadMedianMs: medianOptional(samples.map((sample) => sample.mcpReadMs)),
		mcpPlanMedianMs: medianOptional(samples.map((sample) => sample.mcpPlanMs)),
		mcpPreparedPlanMedianMs: medianOptional(samples.map((sample) => sample.mcpPreparedPlanMs)),
		mcpCommitMedianMs: medianOptional(samples.map((sample) => sample.mcpCommitMs)),
		mcpPreparedCommitMedianMs: medianOptional(samples.map((sample) => sample.mcpPreparedCommitMs)),
		mcpCommitPackageGraphMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitPackageGraphMs),
		),
		mcpCommitPackageGraphAuditMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitPackageGraphAuditMs),
		),
		mcpCommitApplyMedianMs: medianOptional(samples.map((sample) => sample.mcpCommitApplyMs)),
		mcpCommitWritePlanSummaryMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitWritePlanSummaryMs),
		),
		mcpCommitWritePolicyCheckMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitWritePolicyCheckMs),
		),
		mcpCommitWritePolicyBuildMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitWritePolicyBuildMs),
		),
		mcpCommitToBytesMedianMs: medianOptional(samples.map((sample) => sample.mcpCommitToBytesMs)),
		mcpCommitWriteFileMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitWriteFileMs),
		),
		mcpCommitOutputByteReadMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitOutputByteReadMs),
		),
		mcpCommitOutputHashMedianMs: medianOptional(
			samples.map((sample) => sample.mcpCommitOutputHashMs),
		),
		mcpPreparedCommitApplyMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitApplyMs),
		),
		mcpPreparedCommitPackageGraphMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitPackageGraphMs),
		),
		mcpPreparedCommitPackageGraphAuditMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitPackageGraphAuditMs),
		),
		mcpPreparedCommitWritePlanSummaryMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitWritePlanSummaryMs),
		),
		mcpPreparedCommitWritePolicyCheckMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitWritePolicyCheckMs),
		),
		mcpPreparedCommitWritePolicyBuildMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitWritePolicyBuildMs),
		),
		mcpPreparedCommitToBytesMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitToBytesMs),
		),
		mcpPreparedCommitWriteFileMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitWriteFileMs),
		),
		mcpPreparedCommitOutputByteReadMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitOutputByteReadMs),
		),
		mcpPreparedCommitOutputHashMedianMs: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitOutputHashMs),
		),
		mcpVerifyMedianMs: medianOptional(samples.map((sample) => sample.mcpVerifyMs)),
		mcpPreparedVerifyMedianMs: medianOptional(samples.map((sample) => sample.mcpPreparedVerifyMs)),
		mcpPayloadBytesMedian: medianOptional(samples.map((sample) => sample.mcpPayloadBytes)),
		mcpPreparedPayloadBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpPreparedPayloadBytes),
		),
		mcpCommitVerifiedPayloadBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpCommitVerifiedPayloadBytes),
		),
		mcpPreparedCommitVerifiedPayloadBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitVerifiedPayloadBytes),
		),
		mcpCommitOutputBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpCommitOutputBytes),
		),
		mcpPreparedCommitOutputBytesMedian: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitOutputBytes),
		),
		mcpReadCellsMedian: medianOptional(samples.map((sample) => sample.mcpReadCells)),
		mcpCompactHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.mcpCompactHydratedOpenCount),
		),
		mcpCommitVerifiedHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.mcpCommitVerifiedHydratedOpenCount),
		),
		mcpPreparedHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.mcpPreparedHydratedOpenCount),
		),
		mcpPreparedCommitVerifiedHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.mcpPreparedCommitVerifiedHydratedOpenCount),
		),
		mcpReadPartial: samples.some((sample) => sample.mcpReadPartial !== undefined)
			? samples.every((sample) => sample.mcpReadPartial === true)
			: undefined,
		mcpValid: samples.some((sample) => sample.mcpValid !== undefined)
			? samples.every((sample) => sample.mcpValid === true)
			: undefined,
		mcpPreparedValid: samples.some((sample) => sample.mcpPreparedValid !== undefined)
			? samples.every((sample) => sample.mcpPreparedValid === true)
			: undefined,
		planHydratedOpenCountMedian: median(samples.map((sample) => sample.planHydratedOpenCount)),
		preparedPlanHydratedOpenCountMedian: median(
			samples.map((sample) => sample.preparedPlanHydratedOpenCount),
		),
		commitHydratedOpenCountMedian: median(samples.map((sample) => sample.commitHydratedOpenCount)),
		preparedCommitHydratedOpenCountMedian: median(
			samples.map((sample) => sample.preparedCommitHydratedOpenCount),
		),
		documentCacheHitCountMedian: median(samples.map((sample) => sample.documentCacheHitCount)),
		mutationCountMedian: median(samples.map((sample) => sample.mutationCount)),
		rssDeltaMbMedian: median(samples.map((sample) => sample.rssDeltaMb)),
		rssDeltaMbStats: seriesStats(samples.map((sample) => sample.rssDeltaMb)),
		readPartial: samples.every((sample) => sample.readPartial),
		valid: samples.every((sample) => sample.valid),
		preparedValid: samples.every((sample) => sample.preparedValid),
	}
}

async function inferLoadedColumnCount(
	path: string,
	sheetName: string,
	fallbackCols: number,
	rowLimit: number,
): Promise<number> {
	const preview = await WorkbookDocument.open(path, {
		mode: 'values',
		sheets: [sheetName],
		maxRows: rowLimit,
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
			range: args.range ?? `A1:${indexToColumn(args.cols - 1)}${args.rows}`,
			sheet: args.sheet ?? 'Data',
			rows: args.rows,
			cols: args.cols,
			cleanup: true,
			source: 'generated',
		}
	}
	if (args.range !== undefined && args.sheet !== undefined) {
		return {
			xlsxPath: args.inputFile,
			range: args.range,
			sheet: args.sheet,
			rows: args.rows,
			cols: args.cols,
			cleanup: false,
			source: 'input-file',
		}
	}
	const document = await WorkbookDocument.open(args.inputFile, {
		mode: 'metadata-only',
		...(args.sheet !== undefined ? { sheets: [args.sheet] } : {}),
	})
	const info = document.inspect()
	const sheetInfo =
		(args.sheet !== undefined
			? info.sheets.find((sheet) => sheet.name === args.sheet)
			: info.sheets[0]) ?? info.sheets[0]
	if (!sheetInfo) throw new Error(`No sheets found in ${args.inputFile}`)
	const rows = Math.max(1, sheetInfo.rowCount ?? args.rows)
	WorkbookDocument.clearCache()
	const cols = Math.max(
		1,
		sheetInfo.colCount ??
			(await inferLoadedColumnCount(args.inputFile, sheetInfo.name, args.cols, args.rowLimit)),
	)
	return {
		xlsxPath: args.inputFile,
		range: args.range ?? `A1:${indexToColumn(cols - 1)}${rows}`,
		sheet: sheetInfo.name,
		rows,
		cols,
		cleanup: false,
		source: 'input-file',
	}
}

function workflowOutputPath(input: BenchmarkInput): string {
	return join(
		tmpdir(),
		`ascend-agent-workflow-${process.pid}-${Date.now()}-${basename(input.xlsxPath)}.out.xlsx`,
	)
}

async function run() {
	const args = parseArgs()
	if (args.sessionCacheMb !== undefined) {
		configureSessionCache({ maxCacheBytes: args.sessionCacheMb * 1024 * 1024 })
	}
	const data = await resolveBenchmarkInput(args)
	const apiFetch = createApiFetch()
	const mcpClient = args.surface === 'both' ? createMcpWorkflowClient() : undefined
	const outputPath = workflowOutputPath(data)
	const samples: WorkflowSample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runWorkflow(
				apiFetch,
				mcpClient,
				data.xlsxPath,
				outputPath,
				data.sheet,
				data.range,
				args.rowLimit,
				args.mutations,
				data.rows,
				data.cols,
				args.approvals ?? [],
			)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(
				await runWorkflow(
					apiFetch,
					mcpClient,
					data.xlsxPath,
					outputPath,
					data.sheet,
					data.range,
					args.rowLimit,
					args.mutations,
					data.rows,
					data.cols,
					args.approvals ?? [],
				),
			)
			runGc()
		}
		const payload = {
			tool: 'agent-workflow',
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
		await rm(`${outputPath}.prepared.xlsx`, { force: true })
		await rm(`${outputPath}.mcp.xlsx`, { force: true })
		await rm(`${outputPath}.mcp.xlsx.prepared.xlsx`, { force: true })
	}
}

await run()
