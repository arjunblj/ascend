#!/usr/bin/env bun
import { indexToColumn } from '../../../packages/core/src/index.ts'
import { hashLines, hashOrderedLines } from '../competitive-io.ts'

type Library = 'sheetjs' | 'exceljs'

interface Args {
	readonly operation: 'read'
	readonly library: Library
	readonly file: string
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

type PrimitiveAssertion = string | number | boolean | null

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
	const argv = process.argv.slice(2)
	const operation = readOption(argv, '--operation')
	const library = readOption(argv, '--library') ?? 'sheetjs'
	const file = readOption(argv, '--file')
	if (operation !== 'read') throw new Error('--operation must be read')
	if (library !== 'sheetjs' && library !== 'exceljs') {
		throw new Error('--library must be sheetjs or exceljs')
	}
	if (!file) throw new Error('--file is required')
	return {
		operation,
		library,
		file,
		repeat: positiveInt(readOption(argv, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 0),
		json: hasFlag(argv, '--json'),
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

function scalarPayload(value: unknown): string | null {
	if (typeof value === 'number') return `n:${value}`
	if (typeof value === 'string') return `s:${value}`
	if (typeof value === 'boolean') return `b:${value}`
	if (typeof value === 'object' && value !== null && 'result' in value) {
		return scalarPayload((value as { readonly result?: unknown }).result)
	}
	if (typeof value === 'object' && value !== null && 'text' in value) {
		return scalarPayload((value as { readonly text?: unknown }).text)
	}
	return null
}

function valueAssertions(
	sheetCount: number,
	values: readonly string[],
	features: Record<string, PrimitiveAssertion> = {},
): Record<string, PrimitiveAssertion> {
	return {
		sheetCount,
		cellCount: values.length,
		semanticCellValuesHash: hashLines(values),
		orderedSemanticCellValuesHash: hashOrderedLines(values),
		...features,
	}
}

async function readWorkbook(args: Args): Promise<unknown> {
	if (args.library === 'sheetjs') return readSheetJs(args.file)
	return readExcelJs(args.file)
}

async function readSheetJs(file: string): Promise<import('xlsx').WorkBook> {
	const sheetJs = await import('xlsx')
	return sheetJs.readFile(file, { cellFormula: true, cellHTML: false, cellStyles: false })
}

async function readExcelJs(file: string): Promise<import('exceljs').Workbook> {
	const ExcelJS = await import('exceljs')
	const workbook = new ExcelJS.Workbook()
	await workbook.xlsx.readFile(file)
	return workbook
}

async function assertions(
	args: Args,
	workbook: unknown,
): Promise<Record<string, PrimitiveAssertion>> {
	if (args.library === 'sheetjs') {
		return sheetJsAssertions(workbook as import('xlsx').WorkBook)
	}
	return excelJsAssertions(workbook as import('exceljs').Workbook)
}

function sheetJsAssertions(workbook: import('xlsx').WorkBook): Record<string, PrimitiveAssertion> {
	const values: string[] = []
	const worksheet = workbook.Sheets.Data
	const sheetJs = require('xlsx') as typeof import('xlsx')
	if (worksheet) {
		for (const [ref, cell] of Object.entries(worksheet)) {
			if (ref.startsWith('!')) continue
			const decoded = sheetJs.utils.decode_cell(ref)
			const payload = scalarPayload((cell as { readonly v?: unknown }).v)
			if (payload === null) continue
			values.push(`Data!${indexToColumn(decoded.c)}${decoded.r + 1}\t${payload}`)
		}
	}
	return valueAssertions(workbook.SheetNames.length, values, sheetJsFeatureAssertions(workbook))
}

function sheetJsFeatureAssertions(
	workbook: import('xlsx').WorkBook,
): Record<string, PrimitiveAssertion> {
	const worksheet = workbook.Sheets.Data as Record<string, unknown> | undefined
	let readCommentCount = 0
	let readHyperlinkCount = 0
	if (worksheet) {
		for (const [ref, rawCell] of Object.entries(worksheet)) {
			if (ref.startsWith('!') || typeof rawCell !== 'object' || rawCell === null) continue
			const cell = rawCell as {
				readonly c?: readonly unknown[]
				readonly l?: unknown
			}
			if (Array.isArray(cell.c) && cell.c.length > 0) readCommentCount++
			if (cell.l !== undefined) readHyperlinkCount++
		}
	}
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount: 0,
		readConditionalFormatCount: 0,
		readDefinedNameCount: workbook.Workbook?.Names?.length ?? 0,
	}
}

function excelJsAssertions(
	workbook: import('exceljs').Workbook,
): Record<string, PrimitiveAssertion> {
	const values: string[] = []
	const sheet = workbook.getWorksheet('Data')
	if (sheet) {
		sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
			row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
				const payload = scalarPayload(cell.value)
				if (payload === null) return
				values.push(`Data!${indexToColumn(colNumber - 1)}${rowNumber}\t${payload}`)
			})
		})
	}
	return valueAssertions(workbook.worksheets.length, values, excelJsFeatureAssertions(workbook))
}

function excelJsFeatureAssertions(
	workbook: import('exceljs').Workbook,
): Record<string, PrimitiveAssertion> {
	const sheet = workbook.getWorksheet('Data')
	let readCommentCount = 0
	let readHyperlinkCount = 0
	if (sheet) {
		sheet.eachRow({ includeEmpty: false }, (row) => {
			row.eachCell({ includeEmpty: false }, (cell) => {
				if (cell.note) readCommentCount++
				const value = cell.value
				if (
					(typeof value === 'object' && value !== null && 'hyperlink' in value) ||
					(cell as unknown as { readonly hyperlink?: unknown }).hyperlink !== undefined
				) {
					readHyperlinkCount++
				}
			})
		})
	}
	const dataValidations = (
		sheet as unknown as { readonly dataValidations?: { readonly model?: Record<string, unknown> } }
	)?.dataValidations?.model
	const conditionalFormattings = (
		sheet as unknown as { readonly conditionalFormattings?: readonly unknown[] }
	)?.conditionalFormattings
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount: Object.keys(dataValidations ?? {}).length,
		readConditionalFormatCount: conditionalFormattings?.length ?? 0,
		readDefinedNameCount: workbook.definedNames.model.length,
	}
}

async function libraryVersion(args: Args): Promise<string> {
	if (args.library === 'sheetjs') return (await import('xlsx')).version
	try {
		const pkg = await Bun.file('node_modules/exceljs/package.json').json()
		return typeof pkg.version === 'string' ? pkg.version : 'unknown'
	} catch {
		return 'unknown'
	}
}

async function main(): Promise<void> {
	const args = parseArgs()
	for (let i = 0; i < args.warmup; i++) await readWorkbook(args)
	const samples: ReturnType<typeof memorySample>[] = []
	let workbook: unknown
	for (let i = 0; i < args.repeat; i++) {
		runGc()
		const before = memorySnapshot()
		const start = performance.now()
		workbook = await readWorkbook(args)
		samples.push(memorySample(performance.now() - start, before))
	}
	if (!workbook) throw new Error('No samples were produced')
	const payload = {
		assertions: {
			...(await assertions(args, workbook)),
			runnerVersion: await libraryVersion(args),
			runnerApi: args.library === 'sheetjs' ? 'xlsx.readFile' : 'exceljs.xlsx.readFile',
			runnerSource: 'path',
			runnerLoadMode: 'values',
		},
		samples,
	}
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

await main()
