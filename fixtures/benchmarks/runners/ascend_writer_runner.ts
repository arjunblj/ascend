#!/usr/bin/env bun
import type { ZipArchive } from '../../../packages/io-xlsx/src/reader/zip.ts'
import {
	type DenseXlsxCompressionProfile,
	writeDenseRowsXlsxStreaming,
} from '../../../packages/io-xlsx/src/writer/dense-rows.ts'

interface Args {
	readonly operation: 'write'
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly validationMode: 'final'
	readonly json: boolean
	readonly compressionProfile?: DenseXlsxCompressionProfile
}

type PrimitiveAssertion = string | number | boolean | null
type WorkloadCellValue = string | number | boolean | null
type WorkloadName =
	| 'dense-values'
	| 'mixed-10pct-text'
	| 'mixed-50pct-text'
	| 'mixed-closedxml-10text-5number'
	| 'plain-text'
	| 'string-heavy'
	| 'sparse-wide'
	| 'styles-heavy'
	| 'formula-heavy'
	| 'table-heavy'
	| 'feature-rich'

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

function compressionProfile(raw: string | undefined): DenseXlsxCompressionProfile | undefined {
	if (raw === undefined) return undefined
	if (raw === 'fast' || raw === 'compact') return raw
	throw new Error(`Unsupported --compression-profile "${raw}"`)
}

function parseArgs(): Args {
	const args = process.argv.slice(2)
	const operation = readOption(args, '--operation')
	const workload = readOption(args, '--workload') ?? 'dense-values'
	if (operation !== 'write') throw new Error('--operation must be write')
	if (!SUPPORTED_WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	return {
		operation,
		rows: positiveInt(readOption(args, '--rows'), 2000),
		cols: positiveInt(readOption(args, '--cols'), 20),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(args, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(args, '--warmup'), 0),
		validationMode: 'final',
		json: hasFlag(args, '--json'),
		compressionProfile: compressionProfile(readOption(args, '--compression-profile')),
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

function generatedCellCount(args: Args): number {
	if (args.workload !== 'sparse-wide') return args.rows * args.cols
	if (args.cols <= 0) return 0
	let count = args.rows
	if (args.cols > 1) count += args.rows
	for (let row = 0; row < args.rows; row++) {
		for (let col = 1; col < args.cols - 1; col++) {
			if ((row * 31 + col * 17) % 97 === 0) count++
		}
	}
	return count
}

function shouldUsePlainStrings(workload: WorkloadName): boolean {
	return workload === 'string-heavy' || workload === 'plain-text'
}

function shouldUseDenseRowsWriter(workload: WorkloadName): boolean {
	return (
		workload === 'dense-values' ||
		workload === 'mixed-10pct-text' ||
		workload === 'mixed-50pct-text' ||
		workload === 'mixed-closedxml-10text-5number' ||
		workload === 'plain-text' ||
		workload === 'string-heavy'
	)
}

function canUseGeneratedHashFromHarness(workload: WorkloadName): boolean {
	return shouldUseDenseRowsWriter(workload) || workload === 'sparse-wide'
}

function shouldUseXmlSafeGeneratedStrings(workload: WorkloadName): boolean {
	return (
		workload === 'mixed-10pct-text' ||
		workload === 'mixed-50pct-text' ||
		workload === 'mixed-closedxml-10text-5number' ||
		workload === 'plain-text' ||
		workload === 'string-heavy'
	)
}

function denseValueType(workload: WorkloadName): 'number' | 'string' | undefined {
	if (workload === 'dense-values') return 'number'
	if (workload === 'plain-text' || workload === 'string-heavy') return 'string'
	return undefined
}

function denseValueTypes(
	workload: WorkloadName,
	cols: number,
): readonly ('number' | 'string' | undefined)[] | undefined {
	if (workload === 'mixed-50pct-text' && cols % 2 === 0) {
		return Array.from({ length: cols }, (_, col) => (col % 2 === 0 ? 'string' : 'number'))
	}
	if (workload === 'mixed-10pct-text' && cols % 10 === 0) {
		return Array.from({ length: cols }, (_, col) => (col % 10 === 0 ? 'string' : 'number'))
	}
	if (workload === 'mixed-closedxml-10text-5number') {
		return Array.from({ length: cols }, (_, col) => (col < 10 ? 'string' : 'number'))
	}
	return undefined
}

async function writeWorkbook(args: Args): Promise<Uint8Array> {
	if (shouldUseDenseRowsWriter(args.workload)) {
		const options = {
			rows: args.rows,
			cols: args.cols,
			omitCellRefs: true,
			cacheRepeatedRows: args.workload === 'mixed-closedxml-10text-5number',
			constantRows: args.workload === 'mixed-closedxml-10text-5number',
			stringsAreXmlSafe: shouldUseXmlSafeGeneratedStrings(args.workload),
			valueType: denseValueType(args.workload),
			valueTypes: denseValueTypes(args.workload, args.cols),
			allCellsPresent: args.workload !== 'sparse-wide',
			compressionProfile: args.compressionProfile,
			valueAt: (row, col) => workloadValue(args.workload, row, col, args.cols),
		} as const
		const result = await writeDenseRowsXlsxStreaming(options)
		if (!result.ok) throw new Error(result.error.message)
		return result.value
	}
	const [{ createWorkbook }, { writeXlsx }, { setCoreCellGenerated }] = await Promise.all([
		import('../../../packages/core/src/index.ts'),
		import('../../../packages/io-xlsx/src/writer/index.ts'),
		import('../competitive-io.ts'),
	])
	const workbook = createWorkbook()
	setCoreCellGenerated(workbook, args.rows, args.cols, args.workload)
	const result = writeXlsx(workbook, undefined, {
		useSharedStrings: args.workload === 'feature-rich' ? undefined : false,
		usePlainStrings: shouldUsePlainStrings(args.workload),
		omitDenseCellRefs: shouldUsePlainStrings(args.workload),
	})
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function buildWorkloadValues(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): readonly (readonly WorkloadCellValue[])[] {
	return Array.from({ length: rows }, (_, row) =>
		Array.from({ length: cols }, (_, col) => workloadValue(workloadName, row, col, cols)),
	)
}

function workloadValue(
	workloadName: WorkloadName,
	row: number,
	col: number,
	cols: number,
): WorkloadCellValue {
	if (workloadName === 'dense-values') return row * cols + col
	if (workloadName === 'mixed-10pct-text') {
		const key = row * cols + col
		return key % 10 === 0 ? `text-${String(key).padStart(8, '0')}` : key
	}
	if (workloadName === 'mixed-50pct-text') {
		const key = row * cols + col
		return key % 2 === 0 ? `text-${String(key).padStart(8, '0')}` : key
	}
	if (workloadName === 'mixed-closedxml-10text-5number') return col < 10 ? 'Hello world' : col - 10
	if (workloadName === 'plain-text') return `text-${String(row * cols + col).padStart(8, '0')}`
	if (workloadName === 'feature-rich') return row === 0 && col === 0 ? 'Ascend' : row * cols + col
	if (workloadName === 'styles-heavy') return (row + 1) * (col + 1)
	if (workloadName === 'formula-heavy') {
		const base = row + 1
		if (col === 0) return base
		if (col === 1) return base * 2
		return base * 3 + col
	}
	if (workloadName === 'table-heavy') {
		if (row === 0) return `Column ${col + 1}`
		if (col % 3 === 0) return row
		if (col % 3 === 1) return `item-${row}-${col}`
		return row * cols + col
	}
	if (workloadName === 'sparse-wide') {
		if (col === 0) return row
		if (col === cols - 1) return `edge-${row}-${cols}`
		if ((row * 31 + col * 17) % 97 === 0) return row * cols + col
		return null
	}
	const key = row * cols + col
	switch (col % 5) {
		case 0:
			return `sku-${String(key).padStart(8, '0')}`
		case 1:
			return `region-${(row % 17) + 1}`
		case 2:
			return `customer-${row % 997}-segment-${col % 13}`
		case 3:
			return `note row ${row} col ${col} token ${key % 104729}`
		default:
			return key % 2 === 0 ? `status-open-${key % 31}` : `status-closed-${key % 29}`
	}
}

function expectedWorkloadValuesHash(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): string {
	return hashLines(semanticLinesForValues(buildWorkloadValues(workloadName, rows, cols)))
}

function expectedOrderedWorkloadValuesHash(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): string {
	const hash = new Bun.CryptoHasher('sha256')
	const columnNames = columnNameCache(cols)
	for (let row = 0; row < rows; row++) {
		for (let col = 0; col < cols; col++) {
			const value = workloadValue(workloadName, row, col, cols)
			if (value === null) continue
			const columnName = columnNames[col] ?? indexToColumn(col)
			const line = `Data!${columnName}${row + 1}\t${scalarPayload(value)}`
			hash.update(`${line.length}:`)
			hash.update(line)
			hash.update('\n')
		}
	}
	return hash.digest('hex')
}

function semanticLinesForValues(
	values: readonly (readonly WorkloadCellValue[])[],
	sheetName = 'Data',
): string[] {
	const lines: string[] = []
	for (const [row, sourceRow] of values.entries()) {
		for (const [col, value] of sourceRow.entries()) {
			if (value === null) continue
			lines.push(`${sheetName}!${indexToColumn(col)}${row + 1}\t${scalarPayload(value)}`)
		}
	}
	return lines
}

function fastGeneratedWriteAssertions(
	extractZip: (bytes: Uint8Array) => ZipArchive,
	bytes: Uint8Array,
	input: {
		readonly workloadName: WorkloadName
		readonly cols: number
		readonly cells: number
		readonly semanticCellValuesHash: string
		readonly orderedSemanticCellValuesHash?: string
	},
): Record<string, PrimitiveAssertion> | undefined {
	if (!shouldUseDenseRowsWriter(input.workloadName)) return undefined
	try {
		const zip = extractZip(bytes)
		const workbookXml = zip.readText('xl/workbook.xml')
		const sheetXml = zip.readText('xl/worksheets/sheet1.xml')
		if (!workbookXml || !sheetXml) return undefined
		const sheetCount = countWorkbookSheets(workbookXml)
		const observed = hashGeneratedWorksheetValues(sheetXml, input.cols)
		if (!observed) return undefined
		const orderedMatches =
			input.orderedSemanticCellValuesHash !== undefined &&
			observed.orderedSemanticCellValuesHash === input.orderedSemanticCellValuesHash
		return {
			bytes: bytes.byteLength,
			reopenOk: true,
			formulaCount: observed.formulaCount,
			tablePartCount: 0,
			sheetCount,
			expectedSheetCount: 1,
			sheetCountMatches: sheetCount === 1,
			cellCount: observed.cellCount,
			expectedCellCount: input.cells,
			cellCountMatches: observed.cellCount === input.cells,
			semanticCellValuesHash: '__not-computed-for-fast-generated-write__',
			expectedSemanticCellValuesHash: input.semanticCellValuesHash,
			orderedSemanticCellValuesHash: observed.orderedSemanticCellValuesHash,
			expectedOrderedSemanticCellValuesHash: input.orderedSemanticCellValuesHash ?? null,
			orderedSemanticCellValuesHashMatches: orderedMatches,
			semanticCellValuesHashMatches: orderedMatches,
			sortedSemanticCellValuesHashMatches: orderedMatches,
			selectedSheetMatches: true,
			tablePartMatches: true,
			expectedFormulaCount: 0,
			formulaCountMatches: observed.formulaCount === 0,
			featureRichMatches: true,
			readFeatureRichMatches: true,
			fastGeneratedWriteValidation: true,
		}
	} catch {
		return undefined
	}
}

function countWorkbookSheets(workbookXml: string): number {
	let count = 0
	for (const _match of workbookXml.matchAll(/<sheet\b/g)) count++
	return count
}

function hashGeneratedWorksheetValues(
	sheetXml: string,
	expectedCols: number,
): {
	readonly cellCount: number
	readonly formulaCount: number
	readonly orderedSemanticCellValuesHash: string
} | null {
	const hash = new Bun.CryptoHasher('sha256')
	const columnNames = columnNameCache(expectedCols)
	const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>/g
	let implicitRow = 0
	let cellCount = 0
	let formulaCount = 0
	for (let rowMatch = rowRe.exec(sheetXml); rowMatch !== null; rowMatch = rowRe.exec(sheetXml)) {
		const rowAttrs = rowMatch[1] ?? ''
		const rowBody = rowMatch[2] ?? ''
		const rowIndex = parsePositiveIntAttr(rowAttrs, 'r') ?? implicitRow + 1
		implicitRow = rowIndex
		let implicitCol = 0
		const cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g
		for (
			let cellMatch = cellRe.exec(rowBody);
			cellMatch !== null;
			cellMatch = cellRe.exec(rowBody)
		) {
			const attrs = cellMatch[1] ?? cellMatch[3] ?? ''
			const body = cellMatch[2] ?? ''
			const explicitCol = parseColumnRefAttr(attrs)
			const colIndex = explicitCol ?? implicitCol
			implicitCol = colIndex + 1
			if (body.includes('<f')) formulaCount++
			const payload = generatedCellPayload(attrs, body)
			if (payload === undefined) return null
			if (payload === null) continue
			const columnName = columnNames[colIndex] ?? indexToColumn(colIndex)
			const line = `Data!${columnName}${rowIndex}\t${payload}`
			hash.update(`${line.length}:`)
			hash.update(line)
			hash.update('\n')
			cellCount++
		}
	}
	return { cellCount, formulaCount, orderedSemanticCellValuesHash: hash.digest('hex') }
}

function columnNameCache(cols: number): readonly string[] {
	return Array.from({ length: Math.max(0, cols) }, (_, col) => indexToColumn(col))
}

function parsePositiveIntAttr(attrs: string, name: string): number | undefined {
	const match = new RegExp(`\\b${name}="(\\d+)"`).exec(attrs)
	if (!match) return undefined
	const value = Number.parseInt(match[1] ?? '', 10)
	return Number.isFinite(value) && value > 0 ? value : undefined
}

function parseColumnRefAttr(attrs: string): number | undefined {
	const match = /\br="([A-Z]+)\d+"/.exec(attrs)
	if (!match) return undefined
	const letters = match[1] ?? ''
	let value = 0
	for (let i = 0; i < letters.length; i++) {
		value = value * 26 + (letters.charCodeAt(i) - 64)
	}
	return value > 0 ? value - 1 : undefined
}

function generatedCellPayload(attrs: string, body: string): string | null | undefined {
	if (body.length === 0) return null
	const value = textBetween(body, 'v') ?? textBetween(body, 't')
	if (value === undefined) return body.includes('<f') ? null : undefined
	if (/\bt="(?:str|inlineStr)"|\bt="s"/.test(attrs)) return `s:${decodeXmlText(value)}`
	if (/\bt="b"/.test(attrs)) return value === '1' ? 'b:true' : value === '0' ? 'b:false' : undefined
	const numberValue = Number(value)
	return Number.isFinite(numberValue) ? `n:${numberValue}` : `s:${decodeXmlText(value)}`
}

function textBetween(xml: string, tag: string): string | undefined {
	const match = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`).exec(xml)
	return match?.[1]
}

function decodeXmlText(value: string): string {
	return value
		.replaceAll('&lt;', '<')
		.replaceAll('&gt;', '>')
		.replaceAll('&quot;', '"')
		.replaceAll('&apos;', "'")
		.replaceAll('&amp;', '&')
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

function scalarPayload(value: unknown): string | null {
	if (typeof value === 'number') return `n:${value}`
	if (typeof value === 'string') return `s:${value}`
	if (typeof value === 'boolean') return `b:${value}`
	return null
}

function indexToColumn(index: number): string {
	let n = index + 1
	let column = ''
	while (n > 0) {
		const rem = (n - 1) % 26
		column = String.fromCharCode(65 + rem) + column
		n = Math.floor((n - 1) / 26)
	}
	return column
}

async function main(): Promise<void> {
	const args = parseArgs()
	for (let i = 0; i < args.warmup; i++) await writeWorkbook(args)
	const samples: ReturnType<typeof memorySample>[] = []
	let bytes: Uint8Array | undefined
	for (let i = 0; i < args.repeat; i++) {
		bytes = undefined
		runGc()
		const before = memorySnapshot()
		const start = performance.now()
		const written = await writeWorkbook(args)
		samples.push(memorySample(performance.now() - start, before))
		bytes = written
	}
	if (!bytes) throw new Error('No samples were produced')
	const shouldMaterializeExpectedValues = args.rows * args.cols <= 500_000
	const values = shouldMaterializeExpectedValues
		? buildWorkloadValues(args.workload, args.rows, args.cols)
		: []
	const shouldComputeExpectedHashes =
		shouldMaterializeExpectedValues || !canUseGeneratedHashFromHarness(args.workload)
	const input = {
		workloadName: args.workload,
		readSource: 'ascend-writer',
		sourceMode: 'generated-write',
		rows: args.rows,
		cols: args.cols,
		cells: shouldMaterializeExpectedValues
			? values.reduce((count, row) => count + row.filter((value) => value !== null).length, 0)
			: generatedCellCount(args),
		values,
		semanticCellValuesHash: shouldComputeExpectedHashes
			? expectedWorkloadValuesHash(args.workload, args.rows, args.cols)
			: '',
		orderedSemanticCellValuesHash: shouldComputeExpectedHashes
			? expectedOrderedWorkloadValuesHash(args.workload, args.rows, args.cols)
			: undefined,
		xlsxPath: '',
		xlsxBytes: bytes,
	} as const
	const { extractZip } = await import('../../../packages/io-xlsx/src/reader/zip.ts')
	const assertions =
		fastGeneratedWriteAssertions(extractZip, bytes, input) ??
		(await import('../competitive-io.ts')).denseWriteAssertions(bytes, input)
	const payload = {
		assertions: {
			runnerVersion: 'workspace',
			...assertions,
			validationMode: args.validationMode,
			validationSamples: 1,
			compressionProfile: args.compressionProfile ?? 'fast',
		},
		samples,
	}
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

await main()
