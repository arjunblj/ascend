import { createWorkbook, indexToColumn, parseRange } from '../../packages/core/src/index.ts'
import type { Sheet } from '../../packages/core/src/workbook.ts'
import type { CellValue } from '../../packages/schema/src/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'

export interface ColumnarSidecarBenchmarkOptions {
	readonly rows?: number
	readonly cols?: number
	readonly repeats?: number
	readonly generation?: number
}

export interface NumericColumnSidecar {
	readonly rangeRef: string
	readonly rows: number
	readonly cols: number
	readonly generation: number
	readonly populatedCount: number
	readonly numericCount: number
	readonly checksum: number
	readonly values: readonly Float64Array[]
	readonly validity: readonly Uint8Array[]
}

export interface ColumnarSidecarBenchmarkResult {
	readonly source: 'synthetic' | 'fixture'
	readonly rows: number
	readonly cols: number
	readonly cells: number
	readonly repeats: number
	readonly range: string
	readonly generation: number
	readonly populatedCount: number
	readonly numericCount: number
	readonly checksum: number
	readonly estimatedSidecarPayloadBytes: number
	readonly workbookBuildMs: number
	readonly sidecarBuildMs: number
	readonly gridRepeatedScanMs: number
	readonly sidecarRepeatedScanMs: number
	readonly sidecarEndToEndMs: number
	readonly repeatedScanSpeedup: number
	readonly endToEndSpeedup: number
}

export interface ColumnarSidecarFixtureBenchmarkOptions {
	readonly fixture: string
	readonly sheet?: string
	readonly range?: string
	readonly repeats?: number
	readonly generation?: number
}

export interface ColumnarSidecarFixtureBenchmarkResult extends ColumnarSidecarBenchmarkResult {
	readonly source: 'fixture'
	readonly fixture: string
	readonly sheet: string
}

export interface ColumnarSidecarClaimReport {
	readonly generatedAt: string
	readonly allowedClaim: string
	readonly proofStatus: 'passed' | 'needs-more-evidence'
	readonly boundary: string
	readonly benchmark: ColumnarSidecarBenchmarkResult
	readonly generationInvalidation: {
		readonly sidecarGeneration: number
		readonly matchingGenerationValid: boolean
		readonly nextGenerationValid: boolean
	}
	readonly killCriterion: string
	readonly doNotPromoteYet: readonly string[]
	readonly nextProof: string
}

interface Timed<T> {
	readonly ms: number
	readonly value: T
}

const DEFAULT_ROWS = 50_000
const DEFAULT_COLS = 8
const DEFAULT_REPEATS = 40

export function runColumnarSidecarBenchmark(
	options: ColumnarSidecarBenchmarkOptions = {},
): ColumnarSidecarBenchmarkResult {
	const rows = positiveInteger(options.rows, DEFAULT_ROWS)
	const cols = positiveInteger(options.cols, DEFAULT_COLS)
	const repeats = positiveInteger(options.repeats, DEFAULT_REPEATS)
	const generation = positiveInteger(options.generation, 1)
	const range = `A1:${indexToColumn(cols - 1)}${rows}`

	const buildSheet = timeMs(() => buildNumericSheet(rows, cols))
	const sheet = buildSheet.value

	for (let col = 0; col < cols; col++) sumGridColumn(sheet, rows, col)

	const sidecarBuild = timeMs(() => buildNumericColumnSidecar(sheet, range, generation))
	const sidecar = sidecarBuild.value
	const gridScan = timeMs(() => repeatedGridRangeScan(sheet, range, repeats))
	const sidecarScan = timeMs(() => repeatedSidecarScan(sidecar, repeats))

	if (gridScan.value !== sidecarScan.value) {
		throw new Error(`Checksum mismatch: grid=${gridScan.value} sidecar=${sidecarScan.value}`)
	}

	const sidecarEndToEndMs = sidecarBuild.ms + sidecarScan.ms
	return {
		source: 'synthetic',
		rows,
		cols,
		cells: rows * cols,
		repeats,
		range,
		generation,
		populatedCount: sidecar.populatedCount,
		numericCount: sidecar.numericCount,
		checksum: gridScan.value,
		estimatedSidecarPayloadBytes: estimateSidecarPayloadBytes(sidecar),
		workbookBuildMs: roundMs(buildSheet.ms),
		sidecarBuildMs: roundMs(sidecarBuild.ms),
		gridRepeatedScanMs: roundMs(gridScan.ms),
		sidecarRepeatedScanMs: roundMs(sidecarScan.ms),
		sidecarEndToEndMs: roundMs(sidecarEndToEndMs),
		repeatedScanSpeedup: roundRatio(gridScan.ms / Math.max(sidecarScan.ms, Number.EPSILON)),
		endToEndSpeedup: roundRatio(gridScan.ms / Math.max(sidecarEndToEndMs, Number.EPSILON)),
	}
}

export async function runColumnarSidecarFixtureBenchmark(
	options: ColumnarSidecarFixtureBenchmarkOptions,
): Promise<ColumnarSidecarFixtureBenchmarkResult> {
	const repeats = positiveInteger(options.repeats, DEFAULT_REPEATS)
	const generation = positiveInteger(options.generation, 1)
	const open = await timeAsync(() => AscendWorkbook.open(options.fixture))
	const model = open.value.getWorkbookModel()
	const sheet = options.sheet ? model.getSheet(options.sheet) : model.sheets[0]
	if (!sheet) throw new Error(`Missing sheet ${options.sheet ?? '<first>'} in ${options.fixture}`)
	const used = sheet.cells.usedRange()
	const range =
		options.range ??
		(used
			? `${cellName(used.start.row, used.start.col)}:${cellName(used.end.row, used.end.col)}`
			: null)
	if (!range) throw new Error(`No used range in ${options.fixture}#${sheet.name}`)
	const parsed = parseRange(range)
	const rows = parsed.end.row - parsed.start.row + 1
	const cols = parsed.end.col - parsed.start.col + 1

	for (let col = 0; col < cols; col++) {
		sumGridRangeColumn(sheet, parsed.start.row, parsed.start.col + col, parsed.end.row)
	}

	const sidecarBuild = timeMs(() => buildNumericColumnSidecar(sheet, range, generation))
	const sidecar = sidecarBuild.value
	const gridScan = timeMs(() => repeatedGridRangeScan(sheet, range, repeats))
	const sidecarScan = timeMs(() => repeatedSidecarScan(sidecar, repeats))

	if (gridScan.value !== sidecarScan.value) {
		throw new Error(`Checksum mismatch: grid=${gridScan.value} sidecar=${sidecarScan.value}`)
	}

	const sidecarEndToEndMs = sidecarBuild.ms + sidecarScan.ms
	return {
		source: 'fixture',
		fixture: options.fixture,
		sheet: sheet.name,
		rows,
		cols,
		cells: rows * cols,
		repeats,
		range,
		generation,
		populatedCount: sidecar.populatedCount,
		numericCount: sidecar.numericCount,
		checksum: gridScan.value,
		estimatedSidecarPayloadBytes: estimateSidecarPayloadBytes(sidecar),
		workbookBuildMs: roundMs(open.ms),
		sidecarBuildMs: roundMs(sidecarBuild.ms),
		gridRepeatedScanMs: roundMs(gridScan.ms),
		sidecarRepeatedScanMs: roundMs(sidecarScan.ms),
		sidecarEndToEndMs: roundMs(sidecarEndToEndMs),
		repeatedScanSpeedup: roundRatio(gridScan.ms / Math.max(sidecarScan.ms, Number.EPSILON)),
		endToEndSpeedup: roundRatio(gridScan.ms / Math.max(sidecarEndToEndMs, Number.EPSILON)),
	}
}

export function columnarSidecarClaimReport(
	result: ColumnarSidecarBenchmarkResult,
): ColumnarSidecarClaimReport {
	const proofStatus = result.endToEndSpeedup > 1 ? 'passed' : 'needs-more-evidence'
	const fixtureResult = isFixtureBenchmarkResult(result) ? result : null
	return {
		generatedAt: new Date().toISOString(),
		allowedClaim:
			fixtureResult === null
				? 'Ascend has local evidence that a disposable numeric columnar sidecar can accelerate repeated dense range scans while preserving the workbook grid as source of truth.'
				: 'Ascend has public-fixture evidence that a disposable numeric columnar sidecar can preserve checksum parity for repeated range scans while keeping the workbook grid as source of truth.',
		proofStatus,
		boundary:
			fixtureResult === null
				? 'This is a synthetic benchmark over numeric/date-like values, not a production cache, Arrow ABI implementation, DuckDB integration, mixed-type table engine, or mutation invalidation system.'
				: 'This is a small public-fixture benchmark over numeric/date-like values in an imported workbook range, not a production cache, Arrow ABI implementation, DuckDB integration, mixed-type table engine, or mutation invalidation system.',
		benchmark: result,
		generationInvalidation: {
			sidecarGeneration: result.generation,
			matchingGenerationValid: isColumnarSidecarCurrent(result, result.generation),
			nextGenerationValid: isColumnarSidecarCurrent(result, result.generation + 1),
		},
		killCriterion:
			'Do not promote if real workbook tables fail checksum parity, sidecar build plus scan is not faster than grid scans for repeated workloads, or generation-aware invalidation cannot be made explicit.',
		doNotPromoteYet: [
			'Public SDK/API/MCP surface for sidecars.',
			'Claims about mixed values, formulas, filters, hidden rows, merged cells, table totals, or query-backed tables.',
			'Claims that sidecars replace workbook storage or preservation semantics.',
			...(fixtureResult === null
				? []
				: ['Broad performance claims until larger public real-workbook tables are benchmarked.']),
		],
		nextProof:
			'Run the benchmark against public real workbook tables and add generation invalidation probes before any performance-loop production work.',
	}
}

export function columnarSidecarClaimReportMarkdown(report: ColumnarSidecarClaimReport): string {
	const result = report.benchmark
	return [
		'# Columnar Sidecar Claim Report',
		'',
		`Generated: ${report.generatedAt}`,
		`Proof status: ${report.proofStatus}`,
		'',
		'## Claim wording allowed today',
		'',
		report.allowedClaim,
		'',
		'## Honest boundary',
		'',
		report.boundary,
		'',
		'## Proof summary',
		'',
		`Source: ${result.source}`,
		...(isFixtureBenchmarkResult(result)
			? [`Fixture: ${result.fixture}`, `Sheet: ${result.sheet}`]
			: []),
		`Range: ${result.range}`,
		`Cells: ${result.cells}`,
		`Repeats: ${result.repeats}`,
		`Generation: ${result.generation}`,
		`Numeric count: ${result.numericCount}`,
		`Estimated sidecar payload bytes: ${result.estimatedSidecarPayloadBytes}`,
		`Matching generation valid: ${report.generationInvalidation.matchingGenerationValid}`,
		`Next generation valid: ${report.generationInvalidation.nextGenerationValid}`,
		`Grid repeated scan ms: ${result.gridRepeatedScanMs}`,
		`Sidecar build ms: ${result.sidecarBuildMs}`,
		`Sidecar repeated scan ms: ${result.sidecarRepeatedScanMs}`,
		`End-to-end speedup: ${result.endToEndSpeedup}x`,
		'',
		'## Kill criterion',
		'',
		report.killCriterion,
		'',
		'## Do not promote yet',
		'',
		...report.doNotPromoteYet.map((entry) => `- ${entry}`),
		'',
		'## Next proof',
		'',
		report.nextProof,
	].join('\n')
}

function isFixtureBenchmarkResult(
	result: ColumnarSidecarBenchmarkResult,
): result is ColumnarSidecarFixtureBenchmarkResult {
	return result.source === 'fixture'
}

export function buildNumericColumnSidecar(
	sheet: Sheet,
	rangeRef: string,
	generation: number,
): NumericColumnSidecar {
	const range = parseRange(rangeRef)
	const rows = range.end.row - range.start.row + 1
	const cols = range.end.col - range.start.col + 1
	const values = Array.from({ length: cols }, () => new Float64Array(rows))
	const validity = Array.from({ length: cols }, () => new Uint8Array(rows))
	let populatedCount = 0
	let numericCount = 0
	let checksum = 0

	sheet.cells.forEachValueInRange(
		range.start.row,
		range.start.col,
		range.end.row,
		range.end.col,
		(value, row, col) => {
			populatedCount++
			const numeric = numericValue(value)
			if (numeric === null) return
			const rowOffset = row - range.start.row
			const colOffset = col - range.start.col
			const columnValues = values[colOffset]
			const columnValidity = validity[colOffset]
			if (!columnValues || !columnValidity) return
			columnValues[rowOffset] = numeric
			columnValidity[rowOffset] = 1
			numericCount++
			checksum += numeric
		},
	)

	return {
		rangeRef,
		rows,
		cols,
		generation,
		populatedCount,
		numericCount,
		checksum,
		values,
		validity,
	}
}

export function isColumnarSidecarCurrent(
	sidecar:
		| Pick<NumericColumnSidecar, 'generation'>
		| Pick<ColumnarSidecarBenchmarkResult, 'generation'>,
	currentGeneration: number,
): boolean {
	return Number.isInteger(currentGeneration) && currentGeneration === sidecar.generation
}

export function sumSidecarColumn(sidecar: NumericColumnSidecar, colOffset: number): number {
	const values = sidecar.values[colOffset]
	const validity = sidecar.validity[colOffset]
	if (!values || !validity) throw new Error(`Missing sidecar column ${colOffset}`)
	let sum = 0
	for (let row = 0; row < sidecar.rows; row++) {
		if (validity[row] === 1) sum += values[row] ?? 0
	}
	return sum
}

function estimateSidecarPayloadBytes(sidecar: NumericColumnSidecar): number {
	let bytes = 0
	for (const values of sidecar.values) bytes += values.byteLength
	for (const validity of sidecar.validity) bytes += validity.byteLength
	return bytes
}

function buildNumericSheet(rows: number, cols: number): Sheet {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Data')
	const rowValues = new Array<number>(cols)
	for (let row = 0; row < rows; row++) {
		for (let col = 0; col < cols; col++) {
			rowValues[col] = ((row + 1) * (col + 3)) % 10_007
		}
		sheet.cells.setPlainNumberSpan(row, 0, rowValues)
	}
	return sheet
}

function numericValue(value: CellValue): number | null {
	if (value.kind === 'number') return value.value
	if (value.kind === 'date') return value.serial
	return null
}

function repeatedGridRangeScan(sheet: Sheet, rangeRef: string, repeats: number): number {
	const range = parseRange(rangeRef)
	let checksum = 0
	for (let repeat = 0; repeat < repeats; repeat++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			checksum += sumGridRangeColumn(sheet, range.start.row, col, range.end.row)
		}
	}
	return checksum
}

function repeatedSidecarScan(sidecar: NumericColumnSidecar, repeats: number): number {
	let checksum = 0
	for (let repeat = 0; repeat < repeats; repeat++) {
		for (let col = 0; col < sidecar.cols; col++) checksum += sumSidecarColumn(sidecar, col)
	}
	return checksum
}

function sumGridColumn(sheet: Sheet, rows: number, col: number): number {
	return sheet.cells.aggregateNumericInRange(0, col, rows - 1, col).sum
}

function sumGridRangeColumn(sheet: Sheet, startRow: number, col: number, endRow: number): number {
	return sheet.cells.aggregateNumericInRange(startRow, col, endRow, col).sum
}

function cellName(row: number, col: number): string {
	return `${indexToColumn(col)}${row + 1}`
}

function positiveInteger(value: number | undefined, fallback: number): number {
	return Number.isInteger(value) && value !== undefined && value > 0 ? value : fallback
}

function timeMs<T>(fn: () => T): Timed<T> {
	const start = performance.now()
	const value = fn()
	return { ms: performance.now() - start, value }
}

async function timeAsync<T>(fn: () => Promise<T>): Promise<Timed<T>> {
	const start = performance.now()
	const value = await fn()
	return { ms: performance.now() - start, value }
}

function roundMs(value: number): number {
	return Number(value.toFixed(3))
}

function roundRatio(value: number): number {
	return Number(value.toFixed(2))
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	return process.argv[index + 1]
}

if (import.meta.main) {
	const fixture = readFlag('--fixture')
	const result = fixture
		? await runColumnarSidecarFixtureBenchmark({
				fixture,
				sheet: readFlag('--sheet'),
				range: readFlag('--range'),
				repeats: Number(readFlag('--repeats')) || undefined,
				generation: Number(readFlag('--generation')) || undefined,
			})
		: runColumnarSidecarBenchmark({
				rows: Number(readFlag('--rows')) || undefined,
				cols: Number(readFlag('--cols')) || undefined,
				repeats: Number(readFlag('--repeats')) || undefined,
				generation: Number(readFlag('--generation')) || undefined,
			})
	if (process.argv.includes('--claim-report')) {
		const report = columnarSidecarClaimReport(result)
		console.log(
			process.argv.includes('--json')
				? JSON.stringify(report, null, 2)
				: columnarSidecarClaimReportMarkdown(report),
		)
	} else {
		console.log(JSON.stringify(result, null, 2))
	}
}
