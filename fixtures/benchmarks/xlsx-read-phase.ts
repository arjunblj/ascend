#!/usr/bin/env bun
import {
	createWorkbook,
	DEFAULT_STYLE_ID,
	indexToColumn,
	type Workbook,
} from '../../packages/core/src/index.ts'
import { readXlsx, readXlsxRowsStream } from '../../packages/io-xlsx/src/index.ts'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import { booleanValue, type CellValue } from '../../packages/schema/src/index.ts'
import { SheetHandle } from '../../packages/sdk/src/index.ts'
import {
	buildGeneratedWriteDataSet,
	buildRawReadWorkloadDataSet,
	buildWorkloadDataSet,
	type CompetitiveDataSet,
	type ReadSource,
	type WorkloadName,
	workloadValue,
} from './competitive-io.ts'
import { writeHeapSnapshotFromEnv } from './heap-snapshot-on-exit.ts'
import {
	type Args as AscendRunnerArgs,
	directOrderedReadAssertions,
} from './runners/ascend_runner.ts'
import { UPSTREAM_PROFILES } from './upstream-profiles.ts'
import { scanWorksheetXmlBytes, scanWorksheetXmlStructure } from './xlsx-xml-tokenizer.ts'

interface Args {
	readonly profile?: string
	readonly inputFile?: string
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly readSource: ReadSource
	readonly repeat: number
	readonly warmup: number
	readonly phase: PhaseMode
	readonly gcBetweenSamples: boolean
	readonly json: boolean
}

interface PhaseSample {
	readonly zipOpenMs?: number
	readonly worksheetInflateMs?: number
	readonly worksheetDecodeMs?: number
	readonly directOrderedMs?: number
	readonly rowsStreamMs?: number
	readonly rowsStreamChunkedMs?: number
	readonly rowsWindowFirstRowMs?: number
	readonly rowsWindowMs?: number
	readonly worksheetXmlScanMs?: number
	readonly worksheetXmlByteScanMs?: number
	readonly readXlsxMs?: number
	readonly fullReadXlsxMs?: number
	readonly cappedReadWindowMs?: number
	readonly agentWindowMs?: number
	readonly cappedAgentWindowMs?: number
	readonly gridFillMs?: number
	readonly materializeHashMs?: number
	readonly totalHydratedMs?: number
	readonly totalAgentWindowMs?: number
	readonly totalCappedAgentWindowMs?: number
	readonly directOrderedCellsPerSecond?: number
	readonly rowsStreamCellsPerSecond?: number
	readonly rowsStreamChunkedCellsPerSecond?: number
	readonly rowsWindowCellsPerSecond?: number
	readonly worksheetXmlScanCellsPerSecond?: number
	readonly worksheetXmlByteScanCellsPerSecond?: number
	readonly readXlsxCellsPerSecond?: number
	readonly fullReadXlsxCellsPerSecond?: number
	readonly cappedReadWindowCellsPerSecond?: number
	readonly agentWindowCellsPerSecond?: number
	readonly cappedAgentWindowCellsPerSecond?: number
	readonly gridFillCellsPerSecond?: number
	readonly worksheetInflateBytesPerSecond?: number
	readonly worksheetDecodeBytesPerSecond?: number
	readonly bytes: number
	readonly rssAfterBytes: number
	readonly heapUsedBytes: number
}

type ReadXlsxOptions = NonNullable<Parameters<typeof readXlsx>[1]>

interface WorksheetXmlPart {
	readonly text: string
	readonly bytes: Uint8Array
}

type PhaseMode =
	| 'all'
	| 'zip'
	| 'direct'
	| 'rows'
	| 'rows-chunked'
	| 'rows-window'
	| 'xml'
	| 'read'
	| 'full-read'
	| 'hydrate'
	| 'agent-window'
	| 'capped-agent-window'
	| 'grid-fill'

const WORKLOADS = new Set<string>([
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
	'selected-sheet',
	'metadata-only',
	'warm-workflow',
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
	const argv = process.argv.slice(2)
	const profileName = readOption(argv, '--profile')
	const profile = profileName
		? UPSTREAM_PROFILES.find((entry) => entry.name === profileName)
		: undefined
	if (profileName && !profile) {
		throw new Error(
			`Unsupported --profile "${profileName}". Expected one of: ${UPSTREAM_PROFILES.map((entry) => entry.name).join(', ')}`,
		)
	}
	if (profile && profile.category !== 'read') {
		throw new Error(
			`--profile "${profile.name}" is a write profile; xlsx-read-phase only supports read profiles`,
		)
	}
	const workload = readOption(argv, '--workload') ?? profile?.workload ?? 'dense-values'
	if (!WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	const readSource = readOption(argv, '--read-source') ?? profile?.readSource ?? 'raw-ooxml'
	if (readSource !== 'ascend-writer' && readSource !== 'raw-ooxml') {
		throw new Error('--read-source must be ascend-writer or raw-ooxml')
	}
	return {
		...(profile ? { profile: profile.name } : {}),
		...(readOption(argv, '--input-file') ? { inputFile: readOption(argv, '--input-file') } : {}),
		rows: positiveInt(readOption(argv, '--rows'), profile?.rows ?? 2000),
		cols: positiveInt(readOption(argv, '--cols'), profile?.cols ?? 20),
		workload: workload as WorkloadName,
		readSource,
		repeat: positiveInt(readOption(argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 1),
		phase: parsePhase(readOption(argv, '--phase')),
		gcBetweenSamples: hasFlag(argv, '--gc-between-samples'),
		json: hasFlag(argv, '--json'),
	}
}

function parsePhase(raw: string | undefined): PhaseMode {
	if (!raw) return 'all'
	if (
		raw === 'all' ||
		raw === 'zip' ||
		raw === 'direct' ||
		raw === 'rows' ||
		raw === 'rows-chunked' ||
		raw === 'rows-window' ||
		raw === 'xml' ||
		raw === 'read' ||
		raw === 'full-read' ||
		raw === 'hydrate' ||
		raw === 'agent-window' ||
		raw === 'capped-agent-window' ||
		raw === 'grid-fill'
	) {
		return raw
	}
	throw new Error(
		'--phase must be all, zip, direct, rows, rows-chunked, rows-window, xml, read, full-read, hydrate, agent-window, capped-agent-window, or grid-fill',
	)
}

function memory() {
	const current = process.memoryUsage()
	const rss = typeof current.rss === 'function' ? current.rss() : current.rss
	return {
		rssAfterBytes: rss,
		heapUsedBytes: current.heapUsed,
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function medianOptional(values: readonly (number | undefined)[]): number | undefined {
	const defined = values.filter((value): value is number => value !== undefined)
	return defined.length > 0 ? median(defined) : undefined
}

function summarize(samples: readonly PhaseSample[]) {
	const zipOpenMedianMs = medianOptional(samples.map((sample) => sample.zipOpenMs))
	const worksheetInflateMedianMs = medianOptional(
		samples.map((sample) => sample.worksheetInflateMs),
	)
	const worksheetDecodeMedianMs = medianOptional(samples.map((sample) => sample.worksheetDecodeMs))
	const directOrderedMedianMs = medianOptional(samples.map((sample) => sample.directOrderedMs))
	const rowsStreamMedianMs = medianOptional(samples.map((sample) => sample.rowsStreamMs))
	const rowsStreamChunkedMedianMs = medianOptional(
		samples.map((sample) => sample.rowsStreamChunkedMs),
	)
	const rowsWindowFirstRowMedianMs = medianOptional(
		samples.map((sample) => sample.rowsWindowFirstRowMs),
	)
	const rowsWindowMedianMs = medianOptional(samples.map((sample) => sample.rowsWindowMs))
	const worksheetXmlScanMedianMs = medianOptional(
		samples.map((sample) => sample.worksheetXmlScanMs),
	)
	const worksheetXmlByteScanMedianMs = medianOptional(
		samples.map((sample) => sample.worksheetXmlByteScanMs),
	)
	const readXlsxMedianMs = medianOptional(samples.map((sample) => sample.readXlsxMs))
	const fullReadXlsxMedianMs = medianOptional(samples.map((sample) => sample.fullReadXlsxMs))
	const cappedReadWindowMedianMs = medianOptional(
		samples.map((sample) => sample.cappedReadWindowMs),
	)
	const agentWindowMedianMs = medianOptional(samples.map((sample) => sample.agentWindowMs))
	const cappedAgentWindowMedianMs = medianOptional(
		samples.map((sample) => sample.cappedAgentWindowMs),
	)
	const gridFillMedianMs = medianOptional(samples.map((sample) => sample.gridFillMs))
	const materializeHashMedianMs = medianOptional(samples.map((sample) => sample.materializeHashMs))
	const totalHydratedMedianMs = medianOptional(samples.map((sample) => sample.totalHydratedMs))
	const totalAgentWindowMedianMs = medianOptional(
		samples.map((sample) => sample.totalAgentWindowMs),
	)
	const totalCappedAgentWindowMedianMs = medianOptional(
		samples.map((sample) => sample.totalCappedAgentWindowMs),
	)
	const phaseDurations = [
		...(zipOpenMedianMs === undefined ? [] : ([['zipOpen', zipOpenMedianMs]] as const)),
		...(worksheetInflateMedianMs === undefined
			? []
			: ([['worksheetInflate', worksheetInflateMedianMs]] as const)),
		...(worksheetDecodeMedianMs === undefined
			? []
			: ([['worksheetDecode', worksheetDecodeMedianMs]] as const)),
		...(directOrderedMedianMs === undefined
			? []
			: ([['directOrdered', directOrderedMedianMs]] as const)),
		...(rowsStreamMedianMs === undefined ? [] : ([['rowsStream', rowsStreamMedianMs]] as const)),
		...(rowsStreamChunkedMedianMs === undefined
			? []
			: ([['rowsStreamChunked', rowsStreamChunkedMedianMs]] as const)),
		...(rowsWindowFirstRowMedianMs === undefined
			? []
			: ([['rowsWindowFirstRow', rowsWindowFirstRowMedianMs]] as const)),
		...(rowsWindowMedianMs === undefined ? [] : ([['rowsWindow', rowsWindowMedianMs]] as const)),
		...(worksheetXmlScanMedianMs === undefined
			? []
			: ([['worksheetXmlScan', worksheetXmlScanMedianMs]] as const)),
		...(worksheetXmlByteScanMedianMs === undefined
			? []
			: ([['worksheetXmlByteScan', worksheetXmlByteScanMedianMs]] as const)),
		...(readXlsxMedianMs === undefined ? [] : ([['readXlsx', readXlsxMedianMs]] as const)),
		...(fullReadXlsxMedianMs === undefined
			? []
			: ([['fullReadXlsx', fullReadXlsxMedianMs]] as const)),
		...(cappedReadWindowMedianMs === undefined
			? []
			: ([['cappedReadWindow', cappedReadWindowMedianMs]] as const)),
		...(agentWindowMedianMs === undefined ? [] : ([['agentWindow', agentWindowMedianMs]] as const)),
		...(cappedAgentWindowMedianMs === undefined
			? []
			: ([['cappedAgentWindow', cappedAgentWindowMedianMs]] as const)),
		...(gridFillMedianMs === undefined ? [] : ([['gridFill', gridFillMedianMs]] as const)),
		...(materializeHashMedianMs === undefined
			? []
			: ([['materializeHash', materializeHashMedianMs]] as const)),
	]
	const dominantPhase = phaseDurations.reduce(
		(best, current) => (current[1] > best[1] ? current : best),
		phaseDurations[0] ?? ['none', 0],
	)[0]
	return {
		...(zipOpenMedianMs === undefined ? {} : { zipOpenMedianMs }),
		...(worksheetInflateMedianMs === undefined ? {} : { worksheetInflateMedianMs }),
		...(worksheetDecodeMedianMs === undefined ? {} : { worksheetDecodeMedianMs }),
		...(directOrderedMedianMs === undefined ? {} : { directOrderedMedianMs }),
		...(rowsStreamMedianMs === undefined ? {} : { rowsStreamMedianMs }),
		...(rowsStreamChunkedMedianMs === undefined ? {} : { rowsStreamChunkedMedianMs }),
		...(rowsWindowFirstRowMedianMs === undefined ? {} : { rowsWindowFirstRowMedianMs }),
		...(rowsWindowMedianMs === undefined ? {} : { rowsWindowMedianMs }),
		...(readXlsxMedianMs === undefined || rowsStreamMedianMs === undefined
			? {}
			: { rowsStreamVsReadXlsxHeadroom: readXlsxMedianMs / rowsStreamMedianMs }),
		...(readXlsxMedianMs === undefined || rowsStreamChunkedMedianMs === undefined
			? {}
			: { rowsStreamChunkedVsReadXlsxHeadroom: readXlsxMedianMs / rowsStreamChunkedMedianMs }),
		...(readXlsxMedianMs === undefined || rowsWindowMedianMs === undefined
			? {}
			: { rowsWindowVsReadXlsxHeadroom: readXlsxMedianMs / rowsWindowMedianMs }),
		...(worksheetXmlScanMedianMs === undefined ? {} : { worksheetXmlScanMedianMs }),
		...(worksheetXmlByteScanMedianMs === undefined ? {} : { worksheetXmlByteScanMedianMs }),
		...(directOrderedMedianMs === undefined || worksheetXmlScanMedianMs === undefined
			? {}
			: { xmlScanVsDirectHeadroom: directOrderedMedianMs / worksheetXmlScanMedianMs }),
		...(readXlsxMedianMs === undefined || worksheetXmlScanMedianMs === undefined
			? {}
			: { xmlScanVsReadXlsxHeadroom: readXlsxMedianMs / worksheetXmlScanMedianMs }),
		...(worksheetXmlScanMedianMs === undefined || worksheetXmlByteScanMedianMs === undefined
			? {}
			: {
					xmlByteScanVsStringScanHeadroom: worksheetXmlScanMedianMs / worksheetXmlByteScanMedianMs,
				}),
		...(readXlsxMedianMs === undefined || worksheetXmlByteScanMedianMs === undefined
			? {}
			: { xmlByteScanVsReadXlsxHeadroom: readXlsxMedianMs / worksheetXmlByteScanMedianMs }),
		...(readXlsxMedianMs === undefined ? {} : { readXlsxMedianMs }),
		...(fullReadXlsxMedianMs === undefined ? {} : { fullReadXlsxMedianMs }),
		...(readXlsxMedianMs === undefined || fullReadXlsxMedianMs === undefined
			? {}
			: { fullReadVsValuesReadHeadroom: fullReadXlsxMedianMs / readXlsxMedianMs }),
		...(cappedReadWindowMedianMs === undefined ? {} : { cappedReadWindowMedianMs }),
		...(readXlsxMedianMs === undefined || cappedReadWindowMedianMs === undefined
			? {}
			: { cappedReadWindowVsReadXlsxHeadroom: readXlsxMedianMs / cappedReadWindowMedianMs }),
		...(agentWindowMedianMs === undefined ? {} : { agentWindowMedianMs }),
		...(cappedAgentWindowMedianMs === undefined ? {} : { cappedAgentWindowMedianMs }),
		...(gridFillMedianMs === undefined ? {} : { gridFillMedianMs }),
		...(materializeHashMedianMs === undefined ? {} : { materializeHashMedianMs }),
		...(totalHydratedMedianMs === undefined ? {} : { totalHydratedMedianMs }),
		...(totalAgentWindowMedianMs === undefined ? {} : { totalAgentWindowMedianMs }),
		...(totalCappedAgentWindowMedianMs === undefined ? {} : { totalCappedAgentWindowMedianMs }),
		...(totalAgentWindowMedianMs === undefined || totalCappedAgentWindowMedianMs === undefined
			? {}
			: {
					totalCappedAgentWindowVsFullAgentWindowHeadroom:
						totalAgentWindowMedianMs / totalCappedAgentWindowMedianMs,
				}),
		dominantPhase,
		...withOptionalMedian(
			'directOrderedCellsPerSecondMedian',
			samples.map((sample) => sample.directOrderedCellsPerSecond),
		),
		...withOptionalMedian(
			'rowsStreamCellsPerSecondMedian',
			samples.map((sample) => sample.rowsStreamCellsPerSecond),
		),
		...withOptionalMedian(
			'rowsStreamChunkedCellsPerSecondMedian',
			samples.map((sample) => sample.rowsStreamChunkedCellsPerSecond),
		),
		...withOptionalMedian(
			'rowsWindowCellsPerSecondMedian',
			samples.map((sample) => sample.rowsWindowCellsPerSecond),
		),
		...withOptionalMedian(
			'worksheetXmlScanCellsPerSecondMedian',
			samples.map((sample) => sample.worksheetXmlScanCellsPerSecond),
		),
		...withOptionalMedian(
			'worksheetXmlByteScanCellsPerSecondMedian',
			samples.map((sample) => sample.worksheetXmlByteScanCellsPerSecond),
		),
		...withOptionalMedian(
			'readXlsxCellsPerSecondMedian',
			samples.map((sample) => sample.readXlsxCellsPerSecond),
		),
		...withOptionalMedian(
			'fullReadXlsxCellsPerSecondMedian',
			samples.map((sample) => sample.fullReadXlsxCellsPerSecond),
		),
		...withOptionalMedian(
			'cappedReadWindowCellsPerSecondMedian',
			samples.map((sample) => sample.cappedReadWindowCellsPerSecond),
		),
		...withOptionalMedian(
			'agentWindowCellsPerSecondMedian',
			samples.map((sample) => sample.agentWindowCellsPerSecond),
		),
		...withOptionalMedian(
			'cappedAgentWindowCellsPerSecondMedian',
			samples.map((sample) => sample.cappedAgentWindowCellsPerSecond),
		),
		...withOptionalMedian(
			'gridFillCellsPerSecondMedian',
			samples.map((sample) => sample.gridFillCellsPerSecond),
		),
		...withOptionalMedian(
			'worksheetInflateBytesPerSecondMedian',
			samples.map((sample) => sample.worksheetInflateBytesPerSecond),
		),
		...withOptionalMedian(
			'worksheetDecodeBytesPerSecondMedian',
			samples.map((sample) => sample.worksheetDecodeBytesPerSecond),
		),
		bytesMedian: median(samples.map((sample) => sample.bytes)),
		peakRssBytes: Math.max(...samples.map((sample) => sample.rssAfterBytes)),
	}
}

function readOptionsForWorkload(workload: WorkloadName): ReadXlsxOptions {
	if (workload === 'selected-sheet') return { mode: 'values', sheets: ['Data'] }
	if (workload === 'metadata-only') return { mode: 'metadata-only' }
	return { mode: 'values' }
}

function assertReadPhaseWorkbook(workbook: Workbook, input: CompetitiveDataSet): void {
	if (input.workloadName !== 'metadata-only') {
		assertDenseWorkbookOrderedValues(workbook, input)
		return
	}
	const cellCount = workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.cellCount(), 0)
	if (cellCount !== 0) {
		throw new Error(`metadata-only read hydrated ${cellCount} cells`)
	}
}

function assertDenseWorkbookOrderedValues(workbook: Workbook, input: CompetitiveDataSet): void {
	const sheet = workbook.getSheet('Data')
	if (!sheet) throw new Error('Missing Data sheet')
	const hash = new Bun.CryptoHasher('sha256')
	const columnNames = Array.from({ length: input.cols }, (_, col) => indexToColumn(col))
	let cellCount = 0
	const usedRange = sheet.cells.usedRange()
	if (usedRange) {
		sheet.cells.forEachValueInRange(
			usedRange.start.row,
			usedRange.start.col,
			usedRange.end.row,
			usedRange.end.col,
			(value, row, col) => {
				const payload = cellValuePayload(value)
				if (payload === null) return
				const columnName = columnNames[col] ?? indexToColumn(col)
				const line = `Data!${columnName}${row + 1}\t${payload}`
				hash.update(`${line.length}:`)
				hash.update(line)
				hash.update('\n')
				cellCount++
			},
		)
	}
	if (cellCount !== input.cells) {
		throw new Error(`read hydrated ${cellCount} semantic cells; expected ${input.cells}`)
	}
	if (input.orderedSemanticCellValuesHash !== undefined) {
		const observedHash = hash.digest('hex')
		if (observedHash !== input.orderedSemanticCellValuesHash) {
			throw new Error(
				`ordered semantic hash mismatch: ${observedHash}; expected ${input.orderedSemanticCellValuesHash}`,
			)
		}
	}
}

function readAgentWindow(workbook: Workbook, input: CompetitiveDataSet): number {
	const sheet = workbook.getSheet('Data')
	if (!sheet) throw new Error('Missing Data sheet')
	const handle = new SheetHandle(
		'Data',
		() => sheet,
		(_row, _col, cell) => cell.formula,
	)
	const endCol = indexToColumn(Math.max(0, input.cols - 1))
	const rowLimit = Math.min(500, Math.max(1, input.rows))
	const window = handle.readWindowCompact(`A1:${endCol}${Math.max(1, input.rows)}`, {
		rowLimit,
		includeRefs: false,
		omitEmpty: true,
	})
	if (window.cells.length === 0 && input.cells > 0) {
		throw new Error('agent compact window returned no cells')
	}
	if (!window.hasMore && input.rows > rowLimit) {
		throw new Error('agent compact window did not report pagination')
	}
	return window.cells.length
}

function fillGeneratedCoreGrid(input: CompetitiveDataSet): number {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Data')
	if (input.workloadName !== 'sparse-wide') sheet.cells.setExpectedDensity('dense')
	for (let row = 0; row < input.rows; row++) {
		for (let col = 0; col < input.cols; col++) {
			const value = workloadValue(input.workloadName, row, col, input.cols)
			if (value === null) continue
			if (typeof value === 'number') {
				sheet.cells.setPlainNumber(row, col, value)
			} else if (typeof value === 'string') {
				sheet.cells.setStringResolved(row, col, value, null, DEFAULT_STYLE_ID)
			} else {
				sheet.cells.set(row, col, {
					value: booleanValue(value),
					formula: null,
					styleId: DEFAULT_STYLE_ID,
				})
			}
		}
	}
	const cellCount = sheet.cells.cellCount()
	if (cellCount !== input.cells) {
		throw new Error(`grid fill created ${cellCount} cells; expected ${input.cells}`)
	}
	return cellCount
}

function cellValuePayload(value: CellValue): string | null {
	switch (value.kind) {
		case 'number':
			return `n:${value.value}`
		case 'string':
			return `s:${value.value}`
		case 'boolean':
			return `b:${value.value}`
		default:
			return null
	}
}

function withOptionalMedian<T extends string>(
	name: T,
	values: readonly (number | undefined)[],
): Record<T, number> | Record<string, never> {
	const value = medianOptional(values)
	return value === undefined ? {} : ({ [name]: value } as Record<T, number>)
}

function runnerArgs(file: string): AscendRunnerArgs {
	return {
		operation: 'read',
		file,
		mode: 'values',
		source: 'path',
		richMetadata: false,
		orderedHashes: false,
		streamOrderedHashes: true,
		repeat: 1,
		warmup: 0,
		json: true,
	}
}

async function runSample(
	input: Awaited<ReturnType<typeof buildWorkloadDataSet>>,
	args: Args,
	worksheetXml?: WorksheetXmlPart,
): Promise<PhaseSample> {
	const sample: Partial<PhaseSample> = {
		bytes: input.xlsxBytes.byteLength,
	}
	if (args.phase === 'all' || args.phase === 'zip') {
		const zipStart = performance.now()
		const archive = extractZip(input.xlsxBytes)
		const zipOpenMs = performance.now() - zipStart
		const sheetPath = firstWorksheetPath(archive)
		const inflateStart = performance.now()
		const sheetBytes = archive.readBytes(sheetPath)
		const worksheetInflateMs = performance.now() - inflateStart
		if (!sheetBytes) throw new Error(`Missing worksheet XML part ${sheetPath}`)
		const decodeStart = performance.now()
		const sheetText = new TextDecoder('utf-8').decode(sheetBytes)
		const worksheetDecodeMs = performance.now() - decodeStart
		if (sheetText.length === 0) throw new Error(`Decoded empty worksheet XML part ${sheetPath}`)
		sample.zipOpenMs = zipOpenMs
		sample.worksheetInflateMs = worksheetInflateMs
		sample.worksheetDecodeMs = worksheetDecodeMs
		sample.worksheetInflateBytesPerSecond = sheetBytes.byteLength / (worksheetInflateMs / 1000)
		sample.worksheetDecodeBytesPerSecond = sheetBytes.byteLength / (worksheetDecodeMs / 1000)
	}

	if (args.phase === 'all' || args.phase === 'direct') {
		const directStart = performance.now()
		const directAssertions = await directOrderedReadAssertions(
			runnerArgs(input.xlsxPath),
			undefined,
		)
		const directOrderedMs = performance.now() - directStart
		if (!directAssertions) throw new Error('direct ordered scan did not produce assertions')
		sample.directOrderedMs = directOrderedMs
		sample.directOrderedCellsPerSecond = input.cells / (directOrderedMs / 1000)
	}

	if (args.phase === 'all' || args.phase === 'rows') {
		const rowsStart = performance.now()
		const rowsCells = await consumeRowsStream(input.xlsxBytes, false)
		const rowsStreamMs = performance.now() - rowsStart
		if (rowsCells !== input.cells) {
			throw new Error(`rows stream counted ${rowsCells} cells; expected ${input.cells}`)
		}
		sample.rowsStreamMs = rowsStreamMs
		sample.rowsStreamCellsPerSecond = input.cells / (rowsStreamMs / 1000)
	}

	if (args.phase === 'all' || args.phase === 'rows-chunked') {
		const chunkedStart = performance.now()
		const chunkedCells = await consumeRowsStream(input.xlsxBytes, true)
		const rowsStreamChunkedMs = performance.now() - chunkedStart
		if (chunkedCells !== input.cells) {
			throw new Error(`chunked rows stream counted ${chunkedCells} cells; expected ${input.cells}`)
		}
		sample.rowsStreamChunkedMs = rowsStreamChunkedMs
		sample.rowsStreamChunkedCellsPerSecond = input.cells / (rowsStreamChunkedMs / 1000)
	}

	if (args.phase === 'all' || args.phase === 'rows-window') {
		const rowLimit = Math.min(500, Math.max(1, input.rows))
		const windowStart = performance.now()
		const window = await consumeRowsWindow(input.xlsxBytes, rowLimit)
		const rowsWindowMs = performance.now() - windowStart
		if (window.cells <= 0 && input.cells > 0) {
			throw new Error('rows window stream returned no cells')
		}
		sample.rowsWindowFirstRowMs = window.firstRowMs
		sample.rowsWindowMs = rowsWindowMs
		sample.rowsWindowCellsPerSecond = window.cells / (rowsWindowMs / 1000)
	}

	if (args.phase === 'all' || args.phase === 'xml') {
		if (!worksheetXml) throw new Error('worksheet XML was not loaded for xml phase')
		const scanStart = performance.now()
		const scan = scanWorksheetXmlStructure(worksheetXml.text)
		const worksheetXmlScanMs = performance.now() - scanStart
		if (scan.cells !== input.cells) {
			throw new Error(`worksheet XML scan counted ${scan.cells} cells; expected ${input.cells}`)
		}
		const byteScanStart = performance.now()
		const byteScan = scanWorksheetXmlBytes(worksheetXml.bytes)
		const worksheetXmlByteScanMs = performance.now() - byteScanStart
		if (byteScan.cells !== input.cells) {
			throw new Error(
				`worksheet XML byte scan counted ${byteScan.cells} cells; expected ${input.cells}`,
			)
		}
		sample.worksheetXmlScanMs = worksheetXmlScanMs
		sample.worksheetXmlScanCellsPerSecond = input.cells / (worksheetXmlScanMs / 1000)
		sample.worksheetXmlByteScanMs = worksheetXmlByteScanMs
		sample.worksheetXmlByteScanCellsPerSecond = input.cells / (worksheetXmlByteScanMs / 1000)
	}

	if ((args.phase === 'all' && args.workload !== 'metadata-only') || args.phase === 'grid-fill') {
		const gridFillStart = performance.now()
		const gridFillCells = fillGeneratedCoreGrid(input)
		const gridFillMs = performance.now() - gridFillStart
		sample.gridFillMs = gridFillMs
		sample.gridFillCellsPerSecond = gridFillCells / (gridFillMs / 1000)
	}

	if (
		args.phase === 'all' ||
		args.phase === 'read' ||
		args.phase === 'hydrate' ||
		args.phase === 'agent-window'
	) {
		const readStart = performance.now()
		const read = readXlsx(input.xlsxBytes, readOptionsForWorkload(args.workload))
		const readXlsxMs = performance.now() - readStart
		if (!read.ok) throw new Error(read.error.message)
		sample.readXlsxMs = readXlsxMs
		sample.readXlsxCellsPerSecond = input.cells / (readXlsxMs / 1000)

		if (
			(args.phase === 'all' && args.workload !== 'metadata-only') ||
			args.phase === 'agent-window'
		) {
			const windowStart = performance.now()
			const windowCells = readAgentWindow(read.value.workbook, input)
			const agentWindowMs = performance.now() - windowStart
			sample.agentWindowMs = agentWindowMs
			sample.totalAgentWindowMs = readXlsxMs + agentWindowMs
			sample.agentWindowCellsPerSecond = windowCells / (agentWindowMs / 1000)
		}

		if (args.phase === 'all' || args.phase === 'hydrate') {
			const materializeStart = performance.now()
			assertReadPhaseWorkbook(read.value.workbook, input)
			const materializeHashMs = performance.now() - materializeStart
			sample.materializeHashMs = materializeHashMs
			sample.totalHydratedMs = readXlsxMs + materializeHashMs
		}
	}

	if (args.phase === 'full-read') {
		if (args.workload === 'metadata-only') {
			throw new Error('full-read is not supported for metadata-only workloads')
		}
		const fullReadStart = performance.now()
		const read = readXlsx(input.xlsxBytes)
		const fullReadXlsxMs = performance.now() - fullReadStart
		if (!read.ok) throw new Error(read.error.message)
		assertReadPhaseWorkbook(read.value.workbook, input)
		sample.fullReadXlsxMs = fullReadXlsxMs
		sample.fullReadXlsxCellsPerSecond = input.cells / (fullReadXlsxMs / 1000)
	}

	if (
		(args.phase === 'all' && args.workload !== 'metadata-only') ||
		args.phase === 'capped-agent-window'
	) {
		if (args.workload === 'metadata-only') {
			throw new Error('capped-agent-window is not supported for metadata-only workloads')
		}
		const rowLimit = Math.min(500, Math.max(1, input.rows))
		const readOptions = {
			...readOptionsForWorkload(args.workload),
			maxRows: rowLimit,
		} satisfies ReadXlsxOptions
		const cappedReadStart = performance.now()
		const read = readXlsx(input.xlsxBytes, readOptions)
		const cappedReadWindowMs = performance.now() - cappedReadStart
		if (!read.ok) throw new Error(read.error.message)
		const cappedWindowStart = performance.now()
		const windowCells = readAgentWindow(read.value.workbook, input)
		const cappedAgentWindowMs = performance.now() - cappedWindowStart
		sample.cappedReadWindowMs = cappedReadWindowMs
		sample.cappedAgentWindowMs = cappedAgentWindowMs
		sample.totalCappedAgentWindowMs = cappedReadWindowMs + cappedAgentWindowMs
		sample.cappedReadWindowCellsPerSecond = windowCells / (cappedReadWindowMs / 1000)
		sample.cappedAgentWindowCellsPerSecond = windowCells / (cappedAgentWindowMs / 1000)
	}

	const sampleMemory = memory()
	writeHeapSnapshotFromEnv()
	return {
		...sample,
		...sampleMemory,
	} as PhaseSample
}

const args = parseArgs()
let input = await loadInput(args)
const worksheetXml =
	args.phase === 'all' || args.phase === 'xml' ? loadFirstWorksheetXml(input.xlsxBytes) : undefined
if (
	args.phase === 'read' ||
	args.phase === 'direct' ||
	args.phase === 'rows' ||
	args.phase === 'rows-chunked' ||
	args.phase === 'rows-window' ||
	args.phase === 'xml' ||
	args.phase === 'zip' ||
	args.phase === 'full-read' ||
	args.phase === 'agent-window' ||
	args.phase === 'capped-agent-window' ||
	args.phase === 'grid-fill'
) {
	input = { ...input, values: [] }
	runGc()
}
for (let i = 0; i < args.warmup; i++) {
	if (args.gcBetweenSamples) runGc()
	await runSample(input, args, worksheetXml)
}
const samples: PhaseSample[] = []
for (let i = 0; i < args.repeat; i++) {
	if (args.gcBetweenSamples) runGc()
	samples.push(await runSample(input, args, worksheetXml))
}
const payload = {
	tool: 'xlsx-read-phase',
	args,
	summary: summarize(samples),
	samples,
}

if (args.json) {
	console.log(JSON.stringify(payload, null, 2))
} else {
	console.log(payload.summary)
}

function runGc(): void {
	const maybeBun = globalThis as typeof globalThis & { Bun?: { gc?: (force?: boolean) => void } }
	maybeBun.Bun?.gc?.(true)
}

function loadFirstWorksheetXml(bytes: Uint8Array): WorksheetXmlPart {
	const archive = extractZip(bytes)
	const sheetPath = firstWorksheetPath(archive)
	const sheetBytes = archive.readBytes(sheetPath)
	if (!sheetBytes) throw new Error(`Missing worksheet XML part ${sheetPath}`)
	return {
		text: new TextDecoder('utf-8').decode(sheetBytes),
		bytes: sheetBytes,
	}
}

function firstWorksheetPath(archive: ReturnType<typeof extractZip>): string {
	const sheet = [...archive.entries()].find((entry) =>
		/^xl\/worksheets\/sheet\d+\.xml$/i.test(entry.path),
	)
	if (!sheet) throw new Error('No worksheet XML part found')
	return sheet.path
}

async function consumeRowsStream(bytes: Uint8Array, chunkedSheetXml: boolean): Promise<number> {
	const result = await readXlsxRowsStream(bytes, { mode: 'values', chunkedSheetXml })
	if (!result.ok) throw new Error(result.error.message)
	let cells = 0
	for await (const row of result.value) cells += row.cells.length
	return cells
}

async function consumeRowsWindow(
	bytes: Uint8Array,
	rowLimit: number,
): Promise<{ readonly cells: number; readonly firstRowMs: number | undefined }> {
	const start = performance.now()
	const result = await readXlsxRowsStream(bytes, { mode: 'values', maxRows: rowLimit })
	if (!result.ok) throw new Error(result.error.message)
	let rows = 0
	let cells = 0
	let firstRowMs: number | undefined
	for await (const row of result.value) {
		firstRowMs ??= performance.now() - start
		rows++
		cells += row.cells.length
		if (rows >= rowLimit) break
	}
	return { cells, firstRowMs }
}

async function loadInput(args: Args): Promise<CompetitiveDataSet> {
	if (!args.inputFile) {
		if (args.readSource === 'raw-ooxml') {
			return buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
		}
		return buildWorkloadDataSet(args.workload, args.rows, args.cols, args.readSource)
	}
	const bytes = await Bun.file(args.inputFile).bytes()
	if (args.phase !== 'all' && args.phase !== 'hydrate') {
		return {
			workloadName: args.workload,
			readSource: args.readSource,
			sourceMode: 'full',
			rows: args.rows,
			cols: args.cols,
			cells: workloadCellCount(args.workload, args.rows, args.cols),
			values: [],
			semanticCellValuesHash: '',
			orderedSemanticCellValuesHash: '',
			xlsxPath: args.inputFile,
			xlsxBytes: bytes,
		}
	}
	return {
		...buildGeneratedWriteDataSet(args.workload, args.rows, args.cols),
		readSource: args.readSource,
		sourceMode: 'full',
		xlsxPath: args.inputFile,
		xlsxBytes: bytes,
	}
}

function workloadCellCount(workload: WorkloadName, rows: number, cols: number): number {
	if (workload !== 'sparse-wide') return rows * cols
	let count = 0
	for (let row = 0; row < rows; row++) {
		for (let col = 0; col < cols; col++) {
			if (workloadValue(workload, row, col, cols) !== null) count++
		}
	}
	return count
}
