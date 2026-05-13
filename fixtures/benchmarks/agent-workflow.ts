#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { AscendWorkbook, WorkbookDocument } from '../../packages/sdk/src/index.ts'
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
	readonly preparedTotalMs: number
	readonly measuredSampleMs: number
	readonly inspectMs: number
	readonly readMs: number
	readonly planMs: number
	readonly fullPlanMs: number
	readonly preparedPlanMs: number
	readonly commitMs: number
	readonly preparedCommitMs: number
	readonly verifyMs: number
	readonly preparedVerifyMs: number
	readonly payloadBytes: number
	readonly fullPayloadBytes: number
	readonly preparedPayloadBytes: number
	readonly inspectPayloadBytes: number
	readonly readPayloadBytes: number
	readonly planPayloadBytes: number
	readonly fullPlanPayloadBytes: number
	readonly preparedPlanPayloadBytes: number
	readonly commitPayloadBytes: number
	readonly preparedCommitPayloadBytes: number
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
	readonly fullHydratedOpenCount: number
	readonly preparedHydratedOpenCount: number
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
	const preparedOutputPath = `${outputPath}.prepared.xlsx`
	await rm(preparedOutputPath, { force: true })
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
	const preparedPlan = await post(apiFetch, '/plan', {
		file: inputPath,
		mutations,
		compact: true,
		prepare: true,
		maxChangedCells: 25,
	})
	const planHandle = preparedPlan.payload.data?.preparedPlan?.id
	if (!planHandle) throw new Error('Prepared plan did not return a handle')
	const commit = await post(apiFetch, '/commit', {
		file: inputPath,
		output: outputPath,
		mutations,
		approvals: [],
	})
	const preparedCommit = await post(apiFetch, '/commit', {
		planHandle,
		output: preparedOutputPath,
		approvals: [],
	})
	const verify = await post(apiFetch, '/check', { file: outputPath })
	const preparedVerify = await post(apiFetch, '/check', { file: preparedOutputPath })
	const measuredSampleMs = performance.now() - totalStart
	const rssAfter = rssMb()
	const readLoad = read.payload.data?.load
	const sharedOpenStats = addOpenStats(inspect.openStats, read.openStats)
	const compactOpenStats = addOpenStats(
		sharedOpenStats,
		plan.openStats,
		commit.openStats,
		verify.openStats,
	)
	const fullOpenStats = addOpenStats(
		sharedOpenStats,
		fullPlan.openStats,
		commit.openStats,
		verify.openStats,
	)
	const preparedOpenStats = addOpenStats(
		sharedOpenStats,
		preparedPlan.openStats,
		preparedCommit.openStats,
		preparedVerify.openStats,
	)
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
	const preparedWorkflowBytes =
		inspect.text.length +
		read.text.length +
		preparedPlan.text.length +
		preparedCommit.text.length +
		preparedVerify.text.length
	return {
		totalMs: inspect.ms + read.ms + plan.ms + commit.ms + verify.ms,
		fullTotalMs: inspect.ms + read.ms + fullPlan.ms + commit.ms + verify.ms,
		preparedTotalMs: inspect.ms + read.ms + preparedPlan.ms + preparedCommit.ms + preparedVerify.ms,
		measuredSampleMs,
		inspectMs: inspect.ms,
		readMs: read.ms,
		planMs: plan.ms,
		fullPlanMs: fullPlan.ms,
		preparedPlanMs: preparedPlan.ms,
		commitMs: commit.ms,
		preparedCommitMs: preparedCommit.ms,
		verifyMs: verify.ms,
		preparedVerifyMs: preparedVerify.ms,
		payloadBytes: compactWorkflowBytes,
		fullPayloadBytes: fullWorkflowBytes,
		preparedPayloadBytes: preparedWorkflowBytes,
		inspectPayloadBytes: inspect.text.length,
		readPayloadBytes: read.text.length,
		planPayloadBytes: plan.text.length,
		fullPlanPayloadBytes: fullPlan.text.length,
		preparedPlanPayloadBytes: preparedPlan.text.length,
		commitPayloadBytes: commit.text.length,
		preparedCommitPayloadBytes: preparedCommit.text.length,
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
		measuredSampleMedianMs: median(samples.map((sample) => sample.measuredSampleMs)),
		inspectMedianMs: median(samples.map((sample) => sample.inspectMs)),
		readMedianMs: median(samples.map((sample) => sample.readMs)),
		planMedianMs: median(samples.map((sample) => sample.planMs)),
		fullPlanMedianMs: median(samples.map((sample) => sample.fullPlanMs)),
		preparedPlanMedianMs: median(samples.map((sample) => sample.preparedPlanMs)),
		commitMedianMs: median(samples.map((sample) => sample.commitMs)),
		preparedCommitMedianMs: median(samples.map((sample) => sample.preparedCommitMs)),
		verifyMedianMs: median(samples.map((sample) => sample.verifyMs)),
		preparedVerifyMedianMs: median(samples.map((sample) => sample.preparedVerifyMs)),
		payloadBytesMedian: median(samples.map((sample) => sample.payloadBytes)),
		fullPayloadBytesMedian: median(samples.map((sample) => sample.fullPayloadBytes)),
		preparedPayloadBytesMedian: median(samples.map((sample) => sample.preparedPayloadBytes)),
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
		verifyPayloadBytesMedian: median(samples.map((sample) => sample.verifyPayloadBytes)),
		preparedVerifyPayloadBytesMedian: median(
			samples.map((sample) => sample.preparedVerifyPayloadBytes),
		),
		compactWorkflowSpeedupVsFull:
			median(samples.map((sample) => sample.fullTotalMs)) /
			median(samples.map((sample) => sample.totalMs)),
		preparedWorkflowSpeedupVsCompact:
			median(samples.map((sample) => sample.totalMs)) /
			median(samples.map((sample) => sample.preparedTotalMs)),
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
		fullHydratedOpenCountMedian: median(samples.map((sample) => sample.fullHydratedOpenCount)),
		preparedHydratedOpenCountMedian: median(
			samples.map((sample) => sample.preparedHydratedOpenCount),
		),
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
		readPartial: samples.every((sample) => sample.readPartial),
		valid: samples.every((sample) => sample.valid),
		preparedValid: samples.every((sample) => sample.preparedValid),
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
		await rm(`${outputPath}.prepared.xlsx`, { force: true })
	}
}

await run()
