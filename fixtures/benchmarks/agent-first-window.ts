#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { createServer } from '../../apps/mcp/src/index.ts'
import { WorkbookTuiEngine } from '../../apps/tui/src/runtime/engine.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { AscendWorkbook, WorkbookDocument } from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly rowLimit: number
	readonly inputFile?: string
	readonly sheet?: string
	readonly range?: string
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface BenchmarkInput {
	readonly xlsxPath: string
	readonly range: string
	readonly sheet?: string
	readonly cleanup: boolean
	readonly source: 'generated' | 'input-file'
}

interface Sample {
	readonly fullOpenWindowMs?: number
	readonly cappedOpenWindowMs?: number
	readonly apiFirstWindowMs?: number
	readonly mcpFirstWindowMs?: number
	readonly tuiFirstPaintMs?: number
	readonly tuiOpenMs?: number
	readonly tuiRenderMs?: number
	readonly tuiHydrateMs?: number
	readonly cells: number
	readonly payloadBytes?: number
	readonly mcpPayloadBytes?: number
	readonly tuiFrameBytes?: number
	readonly fullRssDeltaMb?: number
	readonly fullRetainedRssDeltaMb?: number
	readonly cappedRssDeltaMb?: number
	readonly cappedRetainedRssDeltaMb?: number
	readonly apiRssDeltaMb?: number
	readonly apiRetainedRssDeltaMb?: number
	readonly mcpRssDeltaMb?: number
	readonly mcpRetainedRssDeltaMb?: number
	readonly tuiRssDeltaMb?: number
	readonly tuiRetainedRssDeltaMb?: number
	readonly fullHydratedCells?: number | null
	readonly cappedHydratedCells?: number | null
	readonly tuiHydratedCells?: number | null
	readonly apiPartial?: boolean
	readonly mcpPartial?: boolean
	readonly tuiPartial?: boolean
	readonly fullOpenCalls?: number
	readonly fullHydratedOpenCount?: number
	readonly fullDocumentCacheHits?: number
	readonly cappedOpenCalls?: number
	readonly cappedHydratedOpenCount?: number
	readonly cappedDocumentCacheHits?: number
	readonly apiOpenCalls?: number
	readonly apiHydratedOpenCount?: number
	readonly apiDocumentCacheHits?: number
	readonly mcpOpenCalls?: number
	readonly mcpHydratedOpenCount?: number
	readonly mcpDocumentCacheHits?: number
	readonly tuiOpenCalls?: number
	readonly tuiHydratedOpenCount?: number
	readonly tuiDocumentCacheHits?: number
}

interface OpenStats {
	readonly workbookOpenCalls: number
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
	workbookHydrations: 0,
	documentOpenCalls: 0,
	documentHydrations: 0,
	documentCacheHits: 0,
}

const seenDocuments = new WeakSet<WorkbookDocument>()

function installOpenStatsInstrumentation(): void {
	type WorkbookOpen = typeof AscendWorkbook.open
	type DocumentOpen = typeof WorkbookDocument.open
	const originalWorkbookOpen = AscendWorkbook.open.bind(AscendWorkbook) as WorkbookOpen
	const originalDocumentOpen = WorkbookDocument.open.bind(WorkbookDocument) as DocumentOpen
	Object.defineProperty(AscendWorkbook, 'open', {
		configurable: true,
		value: (async (...args: Parameters<WorkbookOpen>) => {
			openStats.workbookOpenCalls += 1
			openStats.workbookHydrations += 1
			return originalWorkbookOpen(...args)
		}) satisfies WorkbookOpen,
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
		workbookHydrations: after.workbookHydrations - before.workbookHydrations,
		documentOpenCalls: after.documentOpenCalls - before.documentOpenCalls,
		documentHydrations: after.documentHydrations - before.documentHydrations,
		documentCacheHits: after.documentCacheHits - before.documentCacheHits,
	}
}

installOpenStatsInstrumentation()

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
	'sparse-wide',
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
	const range = readOption(process.argv, '--range')
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		rowLimit: positiveInt(readOption(process.argv, '--row-limit'), 500),
		...(inputFile !== undefined ? { inputFile } : {}),
		...(sheet !== undefined ? { sheet } : {}),
		...(range !== undefined ? { range } : {}),
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

function medianOptional(values: readonly (number | undefined | null)[]): number | undefined {
	const defined = values.filter((value): value is number => typeof value === 'number')
	return defined.length > 0 ? median(defined) : undefined
}

async function time<T>(fn: () => Promise<T>): Promise<{ readonly ms: number; readonly result: T }> {
	const start = performance.now()
	const result = await fn()
	return { ms: performance.now() - start, result }
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function rssMb(): number {
	return process.memoryUsage().rss / (1024 * 1024)
}

async function runFullOpenWindow(
	path: string,
	sheetName: string | undefined,
	range: string,
	rowLimit: number,
): Promise<
	Pick<
		Sample,
		| 'fullOpenWindowMs'
		| 'cells'
		| 'fullHydratedCells'
		| 'fullRssDeltaMb'
		| 'fullRetainedRssDeltaMb'
		| 'fullOpenCalls'
		| 'fullHydratedOpenCount'
		| 'fullDocumentCacheHits'
	>
> {
	WorkbookDocument.clearCache()
	runGc()
	const rssBefore = rssMb()
	const beforeOpenStats = snapshotOpenStats()
	const measured = await time(async () => {
		const document = await WorkbookDocument.open(path, { mode: 'values' })
		const info = document.inspect()
		const targetSheet = sheetName ?? info.sheets[0]?.name ?? 'Sheet1'
		const window = document.readWindowCompact(targetSheet, range, {
			rowLimit,
			includeRefs: false,
			omitEmpty: true,
			flatValues: true,
		})
		return { info, window }
	})
	const rssAfter = rssMb()
	runGc()
	const rssAfterGc = rssMb()
	const openStats = diffOpenStats(snapshotOpenStats(), beforeOpenStats)
	return {
		fullOpenWindowMs: measured.ms,
		cells: measured.result.window?.cells.length ?? 0,
		fullHydratedCells: measured.result.info.cellCount,
		fullRssDeltaMb: rssAfter - rssBefore,
		fullRetainedRssDeltaMb: rssAfterGc - rssBefore,
		fullOpenCalls: openStats.documentOpenCalls,
		fullHydratedOpenCount: openStats.documentHydrations,
		fullDocumentCacheHits: openStats.documentCacheHits,
	}
}

async function runCappedOpenWindow(
	path: string,
	sheet: string | undefined,
	range: string,
	rowLimit: number,
): Promise<
	Pick<
		Sample,
		| 'cappedOpenWindowMs'
		| 'cells'
		| 'cappedHydratedCells'
		| 'cappedRssDeltaMb'
		| 'cappedRetainedRssDeltaMb'
		| 'cappedOpenCalls'
		| 'cappedHydratedOpenCount'
		| 'cappedDocumentCacheHits'
	>
> {
	WorkbookDocument.clearCache()
	runGc()
	const rssBefore = rssMb()
	const beforeOpenStats = snapshotOpenStats()
	const measured = await time(async () => {
		const preview = await WorkbookDocument.openFirstWindow(path, {
			range,
			...(sheet !== undefined ? { sheet } : {}),
			rowLimit,
		})
		return preview
	})
	const rssAfter = rssMb()
	runGc()
	const rssAfterGc = rssMb()
	const openStats = diffOpenStats(snapshotOpenStats(), beforeOpenStats)
	return {
		cappedOpenWindowMs: measured.ms,
		cells: measured.result.window.cells.length,
		cappedHydratedCells: measured.result.info.cellCount,
		cappedRssDeltaMb: rssAfter - rssBefore,
		cappedRetainedRssDeltaMb: rssAfterGc - rssBefore,
		cappedOpenCalls: openStats.documentOpenCalls,
		cappedHydratedOpenCount: openStats.documentHydrations,
		cappedDocumentCacheHits: openStats.documentCacheHits,
	}
}

async function runApiFirstWindow(
	path: string,
	sheet: string | undefined,
	range: string,
	rowLimit: number,
): Promise<
	Pick<
		Sample,
		| 'apiFirstWindowMs'
		| 'cells'
		| 'payloadBytes'
		| 'apiRssDeltaMb'
		| 'apiRetainedRssDeltaMb'
		| 'apiPartial'
		| 'apiOpenCalls'
		| 'apiHydratedOpenCount'
		| 'apiDocumentCacheHits'
	>
> {
	WorkbookDocument.clearCache()
	runGc()
	const rssBefore = rssMb()
	const beforeOpenStats = snapshotOpenStats()
	const apiFetch = createApiFetch()
	const body = JSON.stringify({
		file: path,
		range,
		...(sheet !== undefined ? { sheet } : {}),
		format: 'compact',
		preview: true,
		rowLimit,
	})
	const measured = await time(async () => {
		const response = await apiFetch(
			new Request('http://ascend.local/read', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json' },
				body,
			}),
		)
		const text = await response.text()
		if (response.status !== 200) throw new Error(text)
		return { text, payload: JSON.parse(text) as ApiEnvelope }
	})
	const rssAfter = rssMb()
	runGc()
	const rssAfterGc = rssMb()
	const openStats = diffOpenStats(snapshotOpenStats(), beforeOpenStats)
	const data = measured.result.payload.data
	return {
		apiFirstWindowMs: measured.ms,
		cells: data?.cells?.length ?? 0,
		payloadBytes: measured.result.text.length,
		apiRssDeltaMb: rssAfter - rssBefore,
		apiRetainedRssDeltaMb: rssAfterGc - rssBefore,
		apiPartial: data?.load?.isPartial ?? false,
		apiOpenCalls: openStats.documentOpenCalls,
		apiHydratedOpenCount: openStats.documentHydrations,
		apiDocumentCacheHits: openStats.documentCacheHits,
	}
}

interface ApiEnvelope {
	readonly data?: {
		readonly cells?: readonly unknown[]
		readonly load?: { readonly isPartial?: boolean }
	}
}

interface McpReadResult {
	readonly structuredContent?: {
		readonly ok?: boolean
		readonly data?: {
			readonly cells?: readonly unknown[]
			readonly load?: { readonly isPartial?: boolean }
		}
		readonly error?: unknown
	}
}

type McpReadHandler = (args: {
	readonly file: string
	readonly sheet?: string
	readonly range: string
	readonly format: 'compact'
	readonly preview: boolean
	readonly rowLimit: number
}) => Promise<McpReadResult>

async function runMcpFirstWindow(
	path: string,
	sheet: string | undefined,
	range: string,
	rowLimit: number,
): Promise<
	Pick<
		Sample,
		| 'mcpFirstWindowMs'
		| 'cells'
		| 'mcpPayloadBytes'
		| 'mcpRssDeltaMb'
		| 'mcpRetainedRssDeltaMb'
		| 'mcpPartial'
		| 'mcpOpenCalls'
		| 'mcpHydratedOpenCount'
		| 'mcpDocumentCacheHits'
	>
> {
	WorkbookDocument.clearCache()
	runGc()
	const rssBefore = rssMb()
	const beforeOpenStats = snapshotOpenStats()
	const server = createServer()
	const handler = (
		server as unknown as { _registeredTools: Record<string, { handler: McpReadHandler }> }
	)._registeredTools['ascend.read']?.handler
	if (!handler) throw new Error('MCP ascend.read handler not registered')
	const measured = await time(() =>
		handler({
			file: path,
			range,
			...(sheet !== undefined ? { sheet } : {}),
			format: 'compact',
			preview: true,
			rowLimit,
		}),
	)
	const content = measured.result.structuredContent
	if (content?.ok !== true) {
		throw new Error(`MCP ascend.read failed: ${JSON.stringify(content?.error ?? content)}`)
	}
	const rssAfter = rssMb()
	runGc()
	const rssAfterGc = rssMb()
	const openStats = diffOpenStats(snapshotOpenStats(), beforeOpenStats)
	const data = content.data
	return {
		mcpFirstWindowMs: measured.ms,
		cells: data?.cells?.length ?? 0,
		mcpPayloadBytes: JSON.stringify(content).length,
		mcpRssDeltaMb: rssAfter - rssBefore,
		mcpRetainedRssDeltaMb: rssAfterGc - rssBefore,
		mcpPartial: data?.load?.isPartial ?? false,
		mcpOpenCalls: openStats.documentOpenCalls,
		mcpHydratedOpenCount: openStats.documentHydrations,
		mcpDocumentCacheHits: openStats.documentCacheHits,
	}
}

async function runTuiFirstPaint(
	path: string,
	sheet: string | undefined,
	rowLimit: number,
): Promise<
	Pick<
		Sample,
		| 'tuiFirstPaintMs'
		| 'tuiOpenMs'
		| 'tuiRenderMs'
		| 'tuiHydrateMs'
		| 'tuiFrameBytes'
		| 'tuiRssDeltaMb'
		| 'tuiRetainedRssDeltaMb'
		| 'tuiHydratedCells'
		| 'tuiPartial'
		| 'tuiOpenCalls'
		| 'tuiHydratedOpenCount'
		| 'tuiDocumentCacheHits'
	>
> {
	WorkbookDocument.clearCache()
	runGc()
	const rssBefore = rssMb()
	const beforeOpenStats = snapshotOpenStats()
	const start = performance.now()
	const opened = await time(() =>
		WorkbookTuiEngine.create({
			path,
			...(sheet !== undefined ? { sheet } : {}),
			loadOptions: { mode: 'values', maxRows: rowLimit },
			size: { rows: 24, cols: 100 },
		}),
	)
	const rendered = await time(async () => opened.result.render({ rows: 24, cols: 100 }))
	const totalMs = performance.now() - start
	const rssAfter = rssMb()
	runGc()
	const rssAfterGc = rssMb()
	const openStats = diffOpenStats(snapshotOpenStats(), beforeOpenStats)
	const state = opened.result.state()
	const document = state.workspace.documents[0]
	const latestTelemetry = state.telemetry.at(-1)
	return {
		tuiFirstPaintMs: totalMs,
		tuiOpenMs: opened.ms,
		tuiRenderMs: rendered.ms,
		tuiHydrateMs: latestTelemetry?.hydrateMs,
		tuiFrameBytes: rendered.result.stats.bytes,
		tuiRssDeltaMb: rssAfter - rssBefore,
		tuiRetainedRssDeltaMb: rssAfterGc - rssBefore,
		tuiHydratedCells: document?.info?.cellCount ?? null,
		tuiPartial: document?.info?.load.isPartial ?? false,
		tuiOpenCalls: openStats.workbookOpenCalls + openStats.documentOpenCalls,
		tuiHydratedOpenCount: openStats.workbookHydrations + openStats.documentHydrations,
		tuiDocumentCacheHits: openStats.documentCacheHits,
	}
}

async function resolveBenchmarkInput(args: Args): Promise<BenchmarkInput> {
	if (args.inputFile === undefined) {
		const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
		return {
			xlsxPath: data.xlsxPath,
			range: args.range ?? `A1:${indexToColumn(args.cols - 1)}${args.rows}`,
			...(args.sheet !== undefined ? { sheet: args.sheet } : {}),
			cleanup: true,
			source: 'generated',
		}
	}
	if (args.range !== undefined) {
		return {
			xlsxPath: args.inputFile,
			range: args.range,
			...(args.sheet !== undefined ? { sheet: args.sheet } : {}),
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
	const cols = Math.max(1, sheetInfo.colCount ?? args.cols)
	WorkbookDocument.clearCache()
	return {
		xlsxPath: args.inputFile,
		range: `A1:${indexToColumn(cols - 1)}${rows}`,
		sheet: sheetInfo.name,
		cleanup: false,
		source: 'input-file',
	}
}

function summarize(samples: readonly Sample[]) {
	const fullOpenWindowMedianMs = medianOptional(samples.map((sample) => sample.fullOpenWindowMs))
	const cappedOpenWindowMedianMs = medianOptional(
		samples.map((sample) => sample.cappedOpenWindowMs),
	)
	const apiFirstWindowMedianMs = medianOptional(samples.map((sample) => sample.apiFirstWindowMs))
	const mcpFirstWindowMedianMs = medianOptional(samples.map((sample) => sample.mcpFirstWindowMs))
	const tuiFirstPaintMedianMs = medianOptional(samples.map((sample) => sample.tuiFirstPaintMs))
	return {
		fullOpenWindowMedianMs,
		cappedOpenWindowMedianMs,
		apiFirstWindowMedianMs,
		mcpFirstWindowMedianMs,
		tuiFirstPaintMedianMs,
		tuiOpenMedianMs: medianOptional(samples.map((sample) => sample.tuiOpenMs)),
		tuiRenderMedianMs: medianOptional(samples.map((sample) => sample.tuiRenderMs)),
		tuiHydrateMedianMs: medianOptional(samples.map((sample) => sample.tuiHydrateMs)),
		fullRssDeltaMbMedian: medianOptional(samples.map((sample) => sample.fullRssDeltaMb)),
		fullRetainedRssDeltaMbMedian: medianOptional(
			samples.map((sample) => sample.fullRetainedRssDeltaMb),
		),
		cappedRssDeltaMbMedian: medianOptional(samples.map((sample) => sample.cappedRssDeltaMb)),
		cappedRetainedRssDeltaMbMedian: medianOptional(
			samples.map((sample) => sample.cappedRetainedRssDeltaMb),
		),
		apiRssDeltaMbMedian: medianOptional(samples.map((sample) => sample.apiRssDeltaMb)),
		apiRetainedRssDeltaMbMedian: medianOptional(
			samples.map((sample) => sample.apiRetainedRssDeltaMb),
		),
		mcpRssDeltaMbMedian: medianOptional(samples.map((sample) => sample.mcpRssDeltaMb)),
		mcpRetainedRssDeltaMbMedian: medianOptional(
			samples.map((sample) => sample.mcpRetainedRssDeltaMb),
		),
		tuiRssDeltaMbMedian: medianOptional(samples.map((sample) => sample.tuiRssDeltaMb)),
		tuiRetainedRssDeltaMbMedian: medianOptional(
			samples.map((sample) => sample.tuiRetainedRssDeltaMb),
		),
		...(fullOpenWindowMedianMs !== undefined && cappedOpenWindowMedianMs !== undefined
			? { cappedSpeedupVsFull: fullOpenWindowMedianMs / cappedOpenWindowMedianMs }
			: {}),
		...(fullOpenWindowMedianMs !== undefined && apiFirstWindowMedianMs !== undefined
			? { apiSpeedupVsFull: fullOpenWindowMedianMs / apiFirstWindowMedianMs }
			: {}),
		...(fullOpenWindowMedianMs !== undefined && mcpFirstWindowMedianMs !== undefined
			? { mcpSpeedupVsFull: fullOpenWindowMedianMs / mcpFirstWindowMedianMs }
			: {}),
		...(fullOpenWindowMedianMs !== undefined && tuiFirstPaintMedianMs !== undefined
			? { tuiSpeedupVsFull: fullOpenWindowMedianMs / tuiFirstPaintMedianMs }
			: {}),
		cellsMedian: medianOptional(samples.map((sample) => sample.cells)),
		payloadBytesMedian: medianOptional(samples.map((sample) => sample.payloadBytes)),
		mcpPayloadBytesMedian: medianOptional(samples.map((sample) => sample.mcpPayloadBytes)),
		tuiFrameBytesMedian: medianOptional(samples.map((sample) => sample.tuiFrameBytes)),
		fullHydratedCellsMedian: medianOptional(samples.map((sample) => sample.fullHydratedCells)),
		cappedHydratedCellsMedian: medianOptional(samples.map((sample) => sample.cappedHydratedCells)),
		tuiHydratedCellsMedian: medianOptional(samples.map((sample) => sample.tuiHydratedCells)),
		fullOpenCallsMedian: medianOptional(samples.map((sample) => sample.fullOpenCalls)),
		fullHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.fullHydratedOpenCount),
		),
		fullDocumentCacheHitsMedian: medianOptional(
			samples.map((sample) => sample.fullDocumentCacheHits),
		),
		cappedOpenCallsMedian: medianOptional(samples.map((sample) => sample.cappedOpenCalls)),
		cappedHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.cappedHydratedOpenCount),
		),
		cappedDocumentCacheHitsMedian: medianOptional(
			samples.map((sample) => sample.cappedDocumentCacheHits),
		),
		apiOpenCallsMedian: medianOptional(samples.map((sample) => sample.apiOpenCalls)),
		apiHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.apiHydratedOpenCount),
		),
		apiDocumentCacheHitsMedian: medianOptional(
			samples.map((sample) => sample.apiDocumentCacheHits),
		),
		mcpOpenCallsMedian: medianOptional(samples.map((sample) => sample.mcpOpenCalls)),
		mcpHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.mcpHydratedOpenCount),
		),
		mcpDocumentCacheHitsMedian: medianOptional(
			samples.map((sample) => sample.mcpDocumentCacheHits),
		),
		tuiOpenCallsMedian: medianOptional(samples.map((sample) => sample.tuiOpenCalls)),
		tuiHydratedOpenCountMedian: medianOptional(
			samples.map((sample) => sample.tuiHydratedOpenCount),
		),
		tuiDocumentCacheHitsMedian: medianOptional(
			samples.map((sample) => sample.tuiDocumentCacheHits),
		),
		apiPartial: samples.some((sample) => sample.apiPartial === true),
		mcpPartial: samples.some((sample) => sample.mcpPartial === true),
		tuiPartial: samples.some((sample) => sample.tuiPartial === true),
	}
}

async function run() {
	const args = parseArgs()
	const data = await resolveBenchmarkInput(args)
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runFullOpenWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			await runCappedOpenWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			await runApiFirstWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			await runMcpFirstWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			await runTuiFirstPaint(data.xlsxPath, data.sheet, args.rowLimit)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			const full = await runFullOpenWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			runGc()
			const capped = await runCappedOpenWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			runGc()
			const api = await runApiFirstWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			runGc()
			const mcp = await runMcpFirstWindow(data.xlsxPath, data.sheet, data.range, args.rowLimit)
			runGc()
			const tui = await runTuiFirstPaint(data.xlsxPath, data.sheet, args.rowLimit)
			runGc()
			samples.push({ ...full, ...capped, ...api, ...mcp, ...tui, cells: mcp.cells })
		}
		const payload = {
			tool: 'agent-first-window',
			args,
			input: data,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		if (data.cleanup) await rm(data.xlsxPath, { force: true })
	}
}

await run()
