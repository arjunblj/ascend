#!/usr/bin/env bun
import { indexToColumn } from '../../../packages/core/src/index.ts'
import { readXlsx } from '../../../packages/io-xlsx/src/index.ts'
import type { CellValue } from '../../../packages/schema/src/index.ts'

type Mode = 'formula' | 'values' | 'full' | 'metadata-only'
type Source = 'path' | 'bytes'

interface Args {
	readonly file: string
	readonly mode: Mode
	readonly source: Source
	readonly richMetadata: boolean
	readonly parseDates: boolean
	readonly materializeCells: boolean
	readonly timeOpenOnly: boolean
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
		materializeCells: hasFlag(args, '--materialize-cells'),
		timeOpenOnly: hasFlag(args, '--time-open-only'),
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

function scalarPayload(value: unknown): string | null {
	if (typeof value !== 'object' || value === null || !('kind' in value)) return null
	const cellValue = value as { readonly kind?: unknown; readonly value?: unknown }
	switch (cellValue.kind) {
		case 'number':
			return typeof cellValue.value === 'number' ? `n:${cellValue.value}` : null
		case 'string':
			return typeof cellValue.value === 'string' ? `s:${cellValue.value}` : null
		case 'boolean':
			return typeof cellValue.value === 'boolean' ? `b:${cellValue.value}` : null
		default:
			return null
	}
}

function materializedReadAssertions(
	result: Awaited<ReturnType<typeof openWorkbook>>,
	args: Args,
): Record<string, string | number | boolean | null> {
	const sheet = result.workbook.sheets[0]
	const values: string[] = []
	if (sheet) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			const payload = scalarPayload(cell.value)
			if (payload !== null) values.push(`${sheet.name}!${indexToColumn(col)}${row + 1}\t${payload}`)
		}
	}
	return {
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
		runnerParseDates: args.parseDates,
		runnerApi: 'readXlsx',
		runnerMaterializedCells: true,
		runnerTimeOpenOnly: args.timeOpenOnly,
		sheetCount: sheet ? 1 : 0,
		sheetNamesHash: hashLines(sheet ? [`0:${sheet.name}`] : []),
		cellCount: values.length,
		semanticCellValuesHash: hashLines(values),
	}
}

function workbookShapeAssertions(
	workbook: Awaited<ReturnType<typeof openWorkbook>>['workbook'],
	reportStatus: string,
): Record<string, string | number | boolean | null> {
	let cellCount = 0
	let formulaCount = 0
	const sheetNames = workbook.sheets.map((sheet) => sheet.name)
	const usedRanges: string[] = []
	const semanticCellRefs: string[] = []
	const semanticCellValues: string[] = []
	const formulaTexts: string[] = []
	for (const sheet of workbook.sheets) {
		cellCount += sheet.cells.cellCount()
		formulaCount += sheet.cells.formulaCellCount()
		const usedRange = sheet.cells.usedRange()
		usedRanges.push(
			usedRange
				? `${sheet.name}!${indexToColumn(usedRange.start.col)}${usedRange.start.row + 1}:${indexToColumn(
						usedRange.end.col,
					)}${usedRange.end.row + 1}`
				: `${sheet.name}!empty`,
		)
		for (const [row, col, cell] of sheet.cells.iterate()) {
			const ref = `${sheet.name}!${indexToColumn(col)}${row + 1}`
			semanticCellRefs.push(ref)
			semanticCellValues.push(`${ref}\t${serializeCellValue(cell.value)}`)
			if (cell.formula) formulaTexts.push(`${ref}=${cell.formula}`)
		}
	}
	return {
		sheetCount: workbook.sheets.length,
		sheetNamesHash: hashLines(sheetNames.map((name, index) => `${index}:${name}`)),
		cellCount,
		physicalCellCount: null,
		formulaCount,
		usedRangeCount: usedRanges.length,
		firstUsedRange: usedRanges[0] ?? null,
		firstPhysicalUsedRange: null,
		usedRangesHash: hashLines(usedRanges),
		physicalUsedRangesHash: hashLines([]),
		semanticCellRefsHash: hashLines(semanticCellRefs),
		semanticCellValuesHash: hashLines(semanticCellValues),
		formulaTextHash: hashLines(formulaTexts),
		compatibility: reportStatus,
	}
}

function hashLines(lines: readonly string[]): string {
	const hash = new Bun.CryptoHasher('sha256')
	for (const line of [...lines].sort()) {
		hash.update(`${line.length}:`)
		hash.update(line)
		hash.update('\n')
	}
	return hash.digest('hex')
}

function canonicalNumber(value: number): string {
	return Object.is(value, -0) ? '0' : String(value)
}

function serializeCellValue(value: CellValue): string {
	switch (value.kind) {
		case 'empty':
			return 'empty'
		case 'number':
			return `n:${canonicalNumber(value.value)}`
		case 'date':
			return `n:${canonicalNumber(value.serial)}`
		case 'string':
			return `s:${value.value}`
		case 'richText':
			return `s:${value.runs.map((run) => run.text).join('')}`
		case 'boolean':
			return `b:${value.value ? 'true' : 'false'}`
		case 'error':
			return `e:${value.value}`
	}
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
		...workbookShapeAssertions(result.workbook, result.report.status),
		...readFeatureAssertions(result.workbook),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
		runnerParseDates: args.parseDates,
		runnerApi: 'readXlsx',
		runnerTimeOpenOnly: args.timeOpenOnly,
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

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

async function main(): Promise<void> {
	const args = parseArgs()
	const preloaded = args.source === 'bytes' ? await Bun.file(args.file).bytes() : undefined
	for (let i = 0; i < args.warmup; i++) {
		let warmupWorkbook: Awaited<ReturnType<typeof openWorkbook>> | undefined = await openWorkbook(
			args,
			preloaded,
		)
		if (!args.timeOpenOnly) {
			if (args.materializeCells) {
				materializedReadAssertions(warmupWorkbook, args)
			} else {
				readAssertions(warmupWorkbook, args)
			}
		}
		warmupWorkbook = undefined
		if (args.timeOpenOnly) runGc()
	}
	runGc()
	const samples: ReturnType<typeof memorySample>[] = []
	let assertions: Record<string, string | number | boolean | null> | undefined
	let finalWorkbook: Awaited<ReturnType<typeof openWorkbook>> | undefined
	for (let i = 0; i < args.repeat; i++) {
		if (args.timeOpenOnly) {
			finalWorkbook = undefined
			runGc()
			const start = performance.now()
			finalWorkbook = await openWorkbook(args, preloaded)
			samples.push(memorySample(performance.now() - start))
			continue
		}
		const start = performance.now()
		let workbook: Awaited<ReturnType<typeof openWorkbook>> | undefined = await openWorkbook(
			args,
			preloaded,
		)
		assertions = args.materializeCells
			? materializedReadAssertions(workbook, args)
			: readAssertions(workbook, args)
		samples.push(memorySample(performance.now() - start))
		workbook = undefined
	}
	if (args.timeOpenOnly && finalWorkbook) {
		assertions = args.materializeCells
			? materializedReadAssertions(finalWorkbook, args)
			: readAssertions(finalWorkbook, args)
	}
	const payload = { assertions: assertions ?? {}, samples }
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

await main()
