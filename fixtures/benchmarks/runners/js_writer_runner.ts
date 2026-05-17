#!/usr/bin/env bun
import {
	buildWorkloadValues,
	type CompetitiveDataSet,
	denseWriteAssertions,
	expectedWorkloadValuesHash,
	type WorkloadName,
} from '../competitive-io.ts'

type Library = 'sheetjs' | 'exceljs'

interface Args {
	readonly operation: 'write'
	readonly library: Library
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly validationMode: 'final'
	readonly json: boolean
}

const SUPPORTED_WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'mixed-closedxml-10text-5number',
	'plain-text',
	'string-heavy',
	'sparse-wide',
	'styles-heavy',
	'formula-heavy',
	'table-heavy',
	'feature-rich',
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
	const args = process.argv.slice(2)
	const operation = readOption(args, '--operation')
	const library = readOption(args, '--library') ?? 'sheetjs'
	const workload = readOption(args, '--workload') ?? 'dense-values'
	if (operation !== 'write') throw new Error('--operation must be write')
	if (library !== 'sheetjs' && library !== 'exceljs') {
		throw new Error('--library must be sheetjs or exceljs')
	}
	if (!SUPPORTED_WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	return {
		operation,
		library,
		rows: positiveInt(readOption(args, '--rows'), 2000),
		cols: positiveInt(readOption(args, '--cols'), 20),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(args, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(args, '--warmup'), 0),
		validationMode: 'final',
		json: hasFlag(args, '--json'),
	}
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

function memorySnapshot(): {
	readonly rss: number
	readonly heapUsed: number
} {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	return { rss, heapUsed: memory.heapUsed }
}

function memorySample(
	durationMs: number,
	before: ReturnType<typeof memorySnapshot>,
): {
	readonly durationMs: number
	readonly rssDeltaBytes: number
	readonly retainedRssDeltaBytes: number
	readonly rssAfterBytes: number
	readonly rssAfterGcBytes: number
	readonly peakRssBytes: number
	readonly heapDeltaBytes: number
	readonly heapUsedBytes: number
	readonly heapTotalBytes: number
	readonly heapAfterGcBytes: number
} {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	runGc()
	const afterGc = process.memoryUsage()
	const rssAfterGc = typeof afterGc.rss === 'function' ? afterGc.rss() : afterGc.rss
	return {
		durationMs,
		rssDeltaBytes: Math.max(0, rss - before.rss),
		retainedRssDeltaBytes: Math.max(0, rssAfterGc - before.rss),
		rssAfterBytes: rss,
		rssAfterGcBytes: rssAfterGc,
		peakRssBytes: Math.max(rss, rssAfterGc),
		heapDeltaBytes: Math.max(0, memory.heapUsed - before.heapUsed),
		heapUsedBytes: memory.heapUsed,
		heapTotalBytes: memory.heapTotal,
		heapAfterGcBytes: afterGc.heapUsed,
	}
}

function buildInput(args: Args, bytes: Uint8Array): CompetitiveDataSet {
	const values = buildWorkloadValues(args.workload, args.rows, args.cols)
	return {
		workloadName: args.workload,
		readSource: 'ascend-writer',
		rows: args.rows,
		cols: args.cols,
		cells: values.reduce((count, row) => count + row.filter((value) => value !== null).length, 0),
		values,
		semanticCellValuesHash: expectedWorkloadValuesHash(args.workload, args.rows, args.cols),
		xlsxPath: '',
		xlsxBytes: bytes,
	}
}

function formulaForWorkload(workload: WorkloadName, row: number, col: number): string | undefined {
	if (workload !== 'formula-heavy' || col < 2) return undefined
	const currentRow = row + 1
	return `A${currentRow}+B${currentRow}+${col}`
}

async function writeWorkbook(args: Args): Promise<Uint8Array> {
	if (args.library === 'sheetjs') return writeSheetJs(args)
	return writeExcelJs(args)
}

async function writeSheetJs(args: Args): Promise<Uint8Array> {
	const sheetJs = await import('xlsx')
	const values = buildWorkloadValues(args.workload, args.rows, args.cols)
	const worksheet = sheetJs.utils.aoa_to_sheet(values)
	if (args.workload === 'formula-heavy') {
		for (let row = 0; row < args.rows; row++) {
			for (let col = 2; col < args.cols; col++) {
				const address = sheetJs.utils.encode_cell({ r: row, c: col })
				const cell = worksheet[address]
				if (!cell) continue
				cell.f = formulaForWorkload(args.workload, row, col)
			}
		}
	}
	const workbook = sheetJs.utils.book_new()
	sheetJs.utils.book_append_sheet(workbook, worksheet, 'Data')
	const bytes = sheetJs.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Uint8Array
	return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
}

async function writeExcelJs(args: Args): Promise<Uint8Array> {
	const ExcelJS = await import('exceljs')
	const workbook = new ExcelJS.Workbook()
	const sheet = workbook.addWorksheet('Data')
	const values = buildWorkloadValues(args.workload, args.rows, args.cols)
	for (const row of values) sheet.addRow(row)
	if (args.workload === 'formula-heavy') {
		for (let row = 0; row < args.rows; row++) {
			for (let col = 2; col < args.cols; col++) {
				const cell = sheet.getCell(row + 1, col + 1)
				cell.value = {
					formula: formulaForWorkload(args.workload, row, col),
					result: values[row]?.[col] ?? null,
				}
			}
		}
	}
	const bytes = (await workbook.xlsx.writeBuffer()) as ArrayBuffer
	return new Uint8Array(bytes)
}

async function main(): Promise<void> {
	const args = parseArgs()
	for (let i = 0; i < args.warmup; i++) await writeWorkbook(args)
	const samples: ReturnType<typeof memorySample>[] = []
	let bytes: Uint8Array | undefined
	for (let i = 0; i < args.repeat; i++) {
		runGc()
		const before = memorySnapshot()
		const start = performance.now()
		bytes = await writeWorkbook(args)
		samples.push(memorySample(performance.now() - start, before))
	}
	if (!bytes) throw new Error('No samples were produced')
	const input = buildInput(args, bytes)
	const payload = {
		assertions: {
			runnerVersion: args.library === 'sheetjs' ? (await import('xlsx')).version : '4.4.0',
			...denseWriteAssertions(bytes, input),
			validationMode: args.validationMode,
			validationSamples: 1,
		},
		samples,
	}
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

await main()
