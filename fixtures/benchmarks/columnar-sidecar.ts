import { createWorkbook, indexToColumn, parseRange } from '../../packages/core/src/index.ts'
import type { Sheet } from '../../packages/core/src/workbook.ts'
import type { CellValue } from '../../packages/schema/src/index.ts'

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
	readonly rows: number
	readonly cols: number
	readonly cells: number
	readonly repeats: number
	readonly range: string
	readonly generation: number
	readonly populatedCount: number
	readonly numericCount: number
	readonly checksum: number
	readonly workbookBuildMs: number
	readonly sidecarBuildMs: number
	readonly gridRepeatedScanMs: number
	readonly sidecarRepeatedScanMs: number
	readonly sidecarEndToEndMs: number
	readonly repeatedScanSpeedup: number
	readonly endToEndSpeedup: number
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
	const gridScan = timeMs(() => repeatedGridScan(sheet, rows, cols, repeats))
	const sidecarScan = timeMs(() => repeatedSidecarScan(sidecar, repeats))

	if (gridScan.value !== sidecarScan.value) {
		throw new Error(`Checksum mismatch: grid=${gridScan.value} sidecar=${sidecarScan.value}`)
	}

	const sidecarEndToEndMs = sidecarBuild.ms + sidecarScan.ms
	return {
		rows,
		cols,
		cells: rows * cols,
		repeats,
		range,
		generation,
		populatedCount: sidecar.populatedCount,
		numericCount: sidecar.numericCount,
		checksum: gridScan.value,
		workbookBuildMs: roundMs(buildSheet.ms),
		sidecarBuildMs: roundMs(sidecarBuild.ms),
		gridRepeatedScanMs: roundMs(gridScan.ms),
		sidecarRepeatedScanMs: roundMs(sidecarScan.ms),
		sidecarEndToEndMs: roundMs(sidecarEndToEndMs),
		repeatedScanSpeedup: roundRatio(gridScan.ms / Math.max(sidecarScan.ms, Number.EPSILON)),
		endToEndSpeedup: roundRatio(gridScan.ms / Math.max(sidecarEndToEndMs, Number.EPSILON)),
	}
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

function repeatedGridScan(sheet: Sheet, rows: number, cols: number, repeats: number): number {
	let checksum = 0
	for (let repeat = 0; repeat < repeats; repeat++) {
		for (let col = 0; col < cols; col++) checksum += sumGridColumn(sheet, rows, col)
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

function positiveInteger(value: number | undefined, fallback: number): number {
	return Number.isInteger(value) && value !== undefined && value > 0 ? value : fallback
}

function timeMs<T>(fn: () => T): Timed<T> {
	const start = performance.now()
	const value = fn()
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
	const result = runColumnarSidecarBenchmark({
		rows: Number(readFlag('--rows')) || undefined,
		cols: Number(readFlag('--cols')) || undefined,
		repeats: Number(readFlag('--repeats')) || undefined,
		generation: Number(readFlag('--generation')) || undefined,
	})
	console.log(JSON.stringify(result, null, 2))
}
