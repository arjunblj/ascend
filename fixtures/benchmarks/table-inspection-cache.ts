#!/usr/bin/env bun
import { statSync } from 'node:fs'
import { WorkbookDocument } from '../../packages/sdk/src/index.ts'

interface Args {
	readonly inputFile: string
	readonly sheet?: string
	readonly table?: string
	readonly rowLimit: number
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface Sample {
	readonly coldOpenMs: number
	readonly coldInspectMs: number
	readonly warmOpenMs: number
	readonly warmInspectMs: number
	readonly payloadBytes: number
	readonly tableCount: number
	readonly inspectedRows: number
	readonly totalRows: number
	readonly rssAfterBytes: number
	readonly heapUsedBytes: number
}

const DEFAULT_INPUT = 'fixtures/xlsx/calamine/table-multiple.xlsx'

function readOption(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
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
	const sheet = readOption('--sheet')
	const table = readOption('--table')
	return {
		inputFile: readOption('--input-file') ?? DEFAULT_INPUT,
		...(sheet !== undefined ? { sheet } : {}),
		...(table !== undefined ? { table } : {}),
		rowLimit: positiveInt(readOption('--row-limit'), 100),
		repeat: positiveInt(readOption('--repeat'), 5),
		warmup: nonNegativeInt(readOption('--warmup'), 1),
		json: hasFlag('--json'),
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function timed<T>(fn: () => T): { readonly ms: number; readonly value: T } {
	const start = performance.now()
	const value = fn()
	return { ms: performance.now() - start, value }
}

async function timedAsync<T>(
	fn: () => Promise<T> | T,
): Promise<{ readonly ms: number; readonly value: T }> {
	const start = performance.now()
	const value = await fn()
	return { ms: performance.now() - start, value }
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function tableInspectionPayload(document: WorkbookDocument, args: Args) {
	const workbook = document.inspect()
	const sheetName =
		args.sheet ??
		workbook.sheets.find((sheet) => (sheet.tableCount ?? 0) > 0)?.name ??
		workbook.sheets[0]?.name
	if (!sheetName) throw new Error(`No sheets found in ${args.inputFile}`)
	const sheet = document.inspectSheet(sheetName)
	if (!sheet) throw new Error(`Sheet "${sheetName}" not found in ${args.inputFile}`)
	const tableInfos = sheet.tables ?? []
	const selectedTables = args.table
		? tableInfos.filter((table) => table.name === args.table)
		: tableInfos
	if (args.table && selectedTables.length === 0) {
		throw new Error(`Table "${args.table}" not found on sheet "${sheetName}"`)
	}
	if (selectedTables.length === 0) {
		throw new Error(`No tables found on sheet "${sheetName}" in ${args.inputFile}`)
	}
	const tables = selectedTables.map((table) => {
		const handle = document.table(table.name)
		const rows = handle?.readRows({ limit: args.rowLimit })
		return {
			name: table.name,
			ref: table.ref,
			columnCount: table.columnDefs.length,
			rowCount: table.rowCount,
			hasHeaders: table.hasHeaders,
			hasTotals: table.hasTotals,
			hasFilter: table.autoFilter !== null,
			hasSort: table.sortState !== undefined,
			styleName: table.styleInfo?.name ?? null,
			window: rows
				? {
						returnedRows: rows.returnedRows,
						totalRows: rows.totalRows,
						hasMore: rows.hasMore,
						nextRowOffset: rows.nextRowOffset ?? null,
						rows: rows.rows,
					}
				: null,
		}
	})
	return {
		load: workbook.load,
		sheet: sheetName,
		tableCount: tableInfos.length,
		tables,
	}
}

async function runSample(args: Args): Promise<Sample> {
	WorkbookDocument.clearCache()
	runGc()
	const coldOpen = await timedAsync(() => WorkbookDocument.open(args.inputFile, { mode: 'full' }))
	const coldInspect = timed(() => tableInspectionPayload(coldOpen.value, args))
	const warmOpen = await timedAsync(() => WorkbookDocument.open(args.inputFile, { mode: 'full' }))
	const warmInspect = timed(() => tableInspectionPayload(warmOpen.value, args))
	const payload = JSON.stringify(warmInspect.value)
	const memory = process.memoryUsage()
	const tables = warmInspect.value.tables
	return {
		coldOpenMs: coldOpen.ms,
		coldInspectMs: coldInspect.ms,
		warmOpenMs: warmOpen.ms,
		warmInspectMs: warmInspect.ms,
		payloadBytes: Buffer.byteLength(payload),
		tableCount: warmInspect.value.tableCount,
		inspectedRows: tables.reduce((sum, table) => sum + (table.window?.returnedRows ?? 0), 0),
		totalRows: tables.reduce((sum, table) => sum + table.rowCount, 0),
		rssAfterBytes: memory.rss,
		heapUsedBytes: memory.heapUsed,
	}
}

function summarize(args: Args, samples: readonly Sample[]) {
	return {
		coldOpenMedianMs: median(samples.map((sample) => sample.coldOpenMs)),
		coldInspectMedianMs: median(samples.map((sample) => sample.coldInspectMs)),
		warmOpenMedianMs: median(samples.map((sample) => sample.warmOpenMs)),
		warmInspectMedianMs: median(samples.map((sample) => sample.warmInspectMs)),
		payloadBytesMedian: median(samples.map((sample) => sample.payloadBytes)),
		tableCountMedian: median(samples.map((sample) => sample.tableCount)),
		inspectedRowsMedian: median(samples.map((sample) => sample.inspectedRows)),
		totalRowsMedian: median(samples.map((sample) => sample.totalRows)),
		peakRssBytes: Math.max(...samples.map((sample) => sample.rssAfterBytes)),
		peakHeapUsedBytes: Math.max(...samples.map((sample) => sample.heapUsedBytes)),
		inputBytes: statSync(args.inputFile).size,
		cacheAssumption:
			'Warm table inspection reopens the same path and load options in one process after a cold WorkbookDocument load; this is a session-cache hit, not end-to-end open latency.',
		guardrail:
			'Do not compare warmOpenMedianMs or warmInspectMedianMs to unknown-workbook latency without adding coldOpenMedianMs.',
	}
}

async function run() {
	const args = parseArgs()
	const samples: Sample[] = []
	for (let i = 0; i < args.warmup; i++) await runSample(args)
	for (let i = 0; i < args.repeat; i++) samples.push(await runSample(args))
	const payload = {
		tool: 'table-inspection-cache',
		args,
		summary: summarize(args, samples),
		samples,
	}
	if (args.json) console.log(JSON.stringify(payload, null, 2))
	else console.log(payload.summary)
}

await run()
