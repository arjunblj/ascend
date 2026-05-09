#!/usr/bin/env bun
import { readXlsx } from '../../../packages/io-xlsx/src/index.ts'
import { summarizeAscendWorkbook, workbookShapeAssertions } from '../competitive-real-workbook.ts'

type Mode = 'formula' | 'values' | 'full' | 'metadata-only'
type Source = 'path' | 'bytes'

interface Args {
	readonly file: string
	readonly mode: Mode
	readonly source: Source
	readonly richMetadata: boolean
	readonly parseDates: boolean
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function parseArgs(): Args {
	const args = process.argv.slice(2)
	const file = readOption(args, '--file')
	const mode = readOption(args, '--mode') ?? 'values'
	const source = readOption(args, '--source') ?? 'bytes'
	if (!file) throw new Error('--file is required')
	if (mode !== 'formula' && mode !== 'values' && mode !== 'full' && mode !== 'metadata-only') {
		throw new Error('--mode must be formula, values, full, or metadata-only')
	}
	if (source !== 'path' && source !== 'bytes') {
		throw new Error('--source must be path or bytes')
	}
	return {
		file,
		mode,
		source,
		richMetadata: hasFlag(args, '--rich-metadata'),
		parseDates: !hasFlag(args, '--raw-values'),
		repeat: positiveInt(readOption(args, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(args, '--warmup'), 0),
		json: hasFlag(args, '--json'),
	}
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

async function readBytes(args: Args, preloaded: Uint8Array | undefined): Promise<Uint8Array> {
	return preloaded ?? Bun.file(args.file).bytes()
}

async function openWorkbook(args: Args, preloaded: Uint8Array | undefined) {
	const bytes = await readBytes(args, preloaded)
	const result = readXlsx(bytes, {
		mode: args.mode,
		...(args.richMetadata ? { richMetadata: true } : {}),
		...(args.parseDates ? {} : { parseDates: false }),
	})
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function readAssertions(
	result: Awaited<ReturnType<typeof openWorkbook>>,
	args: Args,
): Record<string, string | number | boolean | null> {
	if (args.mode === 'metadata-only') {
		return {
			metadataOnlyRead: true,
			sourceSheetCount: result.loadInfo.sourceSheetNames.length,
			loadedSheetCount: result.loadInfo.loadedSheetNames.length,
			loadedSheetNames: result.loadInfo.loadedSheetNames.join(','),
			hasAllSheets: result.loadInfo.hasAllSheets,
			cellsHydrated: result.loadInfo.cellsHydrated,
			cellCount: result.workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.cellCount(), 0),
			runnerVersion: 'workspace',
			runnerSource: args.source,
			runnerLoadMode: args.mode,
			runnerApi: 'readXlsx',
		}
	}
	return {
		...workbookShapeAssertions(
			summarizeAscendWorkbook({
				getWorkbookModel: () => result.workbook,
				report: result.report,
			}),
		),
		...readFeatureAssertions(result.workbook),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
		runnerParseDates: args.parseDates,
		runnerApi: 'readXlsx',
	}
}

function readFeatureAssertions(workbook: Awaited<ReturnType<typeof openWorkbook>>['workbook']): {
	readonly readCommentCount: number
	readonly readHyperlinkCount: number
	readonly readDataValidationCount: number
	readonly readConditionalFormatCount: number
	readonly readDefinedNameCount: number
} {
	let readCommentCount = 0
	let readHyperlinkCount = 0
	let readDataValidationCount = 0
	let readConditionalFormatCount = 0
	for (const sheet of workbook.sheets) {
		readCommentCount += sheet.comments.size
		readHyperlinkCount += sheet.hyperlinks.size
		readDataValidationCount += sheet.dataValidations.length
		readConditionalFormatCount += sheet.conditionalFormats.length
	}
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount,
		readConditionalFormatCount,
		readDefinedNameCount: workbook.definedNames.size,
	}
}

function memorySample(durationMs: number): {
	readonly durationMs: number
	readonly rssAfterBytes: number
	readonly peakRssBytes: number
	readonly heapUsedBytes: number
	readonly heapTotalBytes: number
} {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	return {
		durationMs,
		rssAfterBytes: rss,
		peakRssBytes: rss,
		heapUsedBytes: memory.heapUsed,
		heapTotalBytes: memory.heapTotal,
	}
}

async function main(): Promise<void> {
	const args = parseArgs()
	const preloaded = args.source === 'bytes' ? await Bun.file(args.file).bytes() : undefined
	for (let i = 0; i < args.warmup; i++) {
		const workbook = await openWorkbook(args, preloaded)
		readAssertions(workbook, args)
	}
	const samples: ReturnType<typeof memorySample>[] = []
	let assertions: Record<string, string | number | boolean | null> | undefined
	for (let i = 0; i < args.repeat; i++) {
		const start = performance.now()
		const workbook = await openWorkbook(args, preloaded)
		const durationMs = performance.now() - start
		assertions = readAssertions(workbook, args)
		samples.push(memorySample(durationMs))
	}
	const payload = { assertions: assertions ?? {}, samples }
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

await main()
