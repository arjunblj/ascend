/**
 * Head-to-head XLSX read/write benchmarks in the standard benchmark envelope.
 *
 * Run:
 *   bun run fixtures/benchmarks/competitive-io.ts
 *   bun run fixtures/benchmarks/competitive-io.ts --json --repeat 5 > competitive-io.json
 *   bun run fixtures/benchmarks/competitive-io.ts --workload string-heavy --json
 *   bun run fixtures/benchmarks/competitive-io.ts --workload all --json
 *   bun run fixtures/benchmarks/competitive-io.ts --read-source raw-ooxml --json
 *   bun run fixtures/benchmarks/competitive-io.ts --workload all --read-source all --json
 *   bun run fixtures/benchmarks/competitive-io.ts --runner-manifest fixtures/benchmarks/runners/python-readers.manifest.json --workload sparse-wide --json
 *   bun run fixtures/benchmarks/competitive-io.ts --write-runner-manifest path/to/writers.json --workload table-heavy --json
 */
import { access, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strToU8, unzipSync, zipSync } from 'fflate'
import {
	createTableId,
	createWorkbook,
	indexToColumn,
	type Sheet,
	type StyleId,
	type Workbook,
} from '../../packages/core/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import { Ascend } from '../../packages/sdk/src/index.ts'
import {
	benchmarkProvenanceDimensions,
	type ExternalRunnerSpec,
	externalRunnerLicenseGateSatisfied,
	extractWorkbookFeatureSummary,
	normalizeExternalRunnerSpecs,
	normalizeExternalSampleAssertions,
	resolveExternalRunnerCommand,
} from './competitive-real-workbook.ts'
import {
	type BenchmarkCaseResult,
	createBenchmarkSuite,
	formatBytes,
	formatRate,
	summarizeSamples,
} from './results.ts'

const DEFAULT_ROWS = 2000
const DEFAULT_COLS = 20
const SID = 0 as StyleId

type PrimitiveAssertion = string | number | boolean | null
type ScalarCellValue = string | number | boolean
type WorkloadCellValue = ScalarCellValue | null
export type WorkloadName =
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
	| 'selected-sheet'
	| 'metadata-only'
	| 'warm-workflow'
type WorkloadSelection = WorkloadName | 'all'
export type ReadSource = 'ascend-writer' | 'raw-ooxml'
type ReadSourceSelection = ReadSource | 'all'
type CategorySelection = 'read' | 'write' | 'all'
type CompetitorSelection = 'js' | 'python' | 'external' | 'all'
type ValidationMode = 'each' | 'final'
type ExecutionScopeSelection = 'in-process' | 'external-process' | 'all'
type SourceMode = 'full' | 'generated-write'
type LibraryAllowlist = ReadonlySet<string> | undefined
const ASCEND_LIBRARY_ALIASES = [
	'ascend',
	'ascend-external',
	'ascend-external-values',
	'ascend-external-values-ordered',
	'ascend-external-bytes',
	'ascend-external-values-bytes',
	'ascend-readxlsx-raw-values-bytes',
	'ascend-readxlsx-raw-values-operation-bytes',
	'ascend-readxlsx-raw-values-operation-path',
	'ascend-readxlsx-row-stream-bytes',
	'ascend-readxlsx-cell-materialization-bytes',
	'ascend-readxlsx-selected-values',
	'ascend-readxlsx-values-bytes',
	'ascend-readxlsx-values-rich-metadata-path',
	'ascend-readxlsx-values-rich-metadata-bytes',
	'ascend-external-metadata-only-bytes',
	'ascend-external-writer',
	'ascend-external-writer-compact',
] as const
const LIBRARY_ALLOWLIST_ALIASES = new Map<string, readonly string[]>([
	['ascend', ASCEND_LIBRARY_ALIASES],
])
const FAST_GENERATED_WRITE_HASH_NOT_COMPUTED = '__not-computed-for-fast-generated-write__'
const ALL_WORKLOADS = [
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
] as const
const ALL_READ_SOURCES = ['ascend-writer', 'raw-ooxml'] as const

export interface CompetitiveDataSet {
	readonly workloadName: WorkloadName
	readonly readSource: ReadSource
	readonly sourceMode?: SourceMode
	readonly rows: number
	readonly cols: number
	readonly cells: number
	readonly values: readonly (readonly WorkloadCellValue[])[]
	readonly semanticCellValuesHash: string
	readonly orderedSemanticCellValuesHash?: string
	readonly xlsxPath: string
	readonly xlsxBytes: Uint8Array
}

export type DenseDataSet = CompetitiveDataSet

interface CompetitiveCase {
	readonly name: string
	readonly library: string
	readonly category: 'read' | 'write'
	readonly executionScope?: 'in-process' | 'external-process'
	readonly operationProfile?: string
	readonly runnerProvenance?: {
		readonly adapterVersion?: string
		readonly libraryVersion?: string
		readonly runtime?: string
	}
	readonly timingModel?: string
	readonly validationModel?: string
	readonly memoryModel?: string
	readonly capabilities?: {
		readonly valueOnlyRead?: boolean
		readonly metadataOnlyRead?: boolean
		readonly selectedSheetRead?: boolean
		readonly writeFormulas?: boolean
		readonly writeTables?: boolean
		readonly writeRichMetadata?: boolean
	}
	run(input: CompetitiveDataSet): Promise<{ assertions?: Record<string, PrimitiveAssertion> }>
	runBatched?(
		input: CompetitiveDataSet,
		repeat: number,
		warmup: number,
	): Promise<{
		assertions?: Record<string, PrimitiveAssertion>
		assertionsBySample?: readonly Record<string, PrimitiveAssertion>[]
		samples?: readonly MetricSample[]
	}>
}

export type ExternalReadRunnerSpec = ExternalRunnerSpec
export type ExternalWriteRunnerSpec = ExternalRunnerSpec

type WriteCapability = keyof Pick<
	NonNullable<CompetitiveCase['capabilities']>,
	'writeFormulas' | 'writeTables' | 'writeRichMetadata'
>

const REQUIRED_WRITE_CAPABILITY_BY_WORKLOAD: Partial<Record<WorkloadName, WriteCapability>> = {
	'formula-heavy': 'writeFormulas',
	'table-heavy': 'writeTables',
	'feature-rich': 'writeRichMetadata',
}

export function unsupportedExternalWriteWorkloadReason(
	spec: Pick<ExternalWriteRunnerSpec, 'name' | 'capabilities'>,
	workloadName: WorkloadName,
): string | undefined {
	const requiredCapability = REQUIRED_WRITE_CAPABILITY_BY_WORKLOAD[workloadName]
	if (requiredCapability && spec.capabilities?.[requiredCapability] !== true)
		return `unsupported write workload "${workloadName}": runner does not declare capabilities.${requiredCapability}=true`
	return undefined
}

function supportsGeneratedWriteWorkload(
	skipped: Array<{ library: string; reason: string }>,
	spec: Pick<ExternalWriteRunnerSpec, 'name' | 'capabilities'>,
	workloadName: WorkloadName,
): boolean {
	const reason = unsupportedExternalWriteWorkloadReason(spec, workloadName)
	if (reason) {
		skipped.push({ library: spec.name, reason })
		return false
	}
	return true
}

function supportsGeneratedRunnerWorkload(
	spec: Pick<ExternalRunnerSpec, 'workloads'>,
	workloadName: WorkloadName,
): boolean {
	return spec.workloads === undefined || spec.workloads.includes(workloadName)
}

export interface MetricSample {
	readonly durationMs: number
	readonly throughputPerSec?: number
	readonly rssDeltaBytes?: number
	readonly retainedRssDeltaBytes?: number
	readonly rssAfterBytes?: number
	readonly rssAfterGcBytes?: number
	readonly peakRssBytes?: number
	readonly heapDeltaBytes?: number
	readonly heapUsedBytes?: number
	readonly heapTotalBytes?: number
	readonly heapAfterGcBytes?: number
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function readPositiveIntFlag(name: string, fallback: number): number {
	const raw = readFlag(name)
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function readNonNegativeIntFlag(name: string, fallback: number): number {
	const raw = readFlag(name)
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

async function loadExternalReadRunnerSpecs(): Promise<ExternalReadRunnerSpec[]> {
	const manifestPath = readFlag('--runner-manifest')
	if (!manifestPath) return []
	const raw = await readFile(manifestPath, 'utf-8')
	return normalizeExternalRunnerSpecs(JSON.parse(raw) as unknown).filter(
		(spec) =>
			externalRunnerLicenseGateSatisfied(spec) &&
			(spec.categories === undefined || spec.categories.includes('read')),
	)
}

async function loadExternalWriteRunnerSpecs(): Promise<ExternalWriteRunnerSpec[]> {
	const manifestPath = readFlag('--write-runner-manifest')
	if (!manifestPath) return []
	const raw = await readFile(manifestPath, 'utf-8')
	return normalizeExternalRunnerSpecs(JSON.parse(raw) as unknown).filter(
		(spec) =>
			externalRunnerLicenseGateSatisfied(spec) &&
			(spec.categories === undefined || spec.categories.includes('write')),
	)
}

function readWorkloadFlag(): WorkloadSelection {
	const raw = readFlag('--workload')
	if (raw === undefined || raw === 'dense-values') return 'dense-values'
	if (raw === 'all') return 'all'
	if ((ALL_WORKLOADS as readonly string[]).includes(raw)) return raw as WorkloadName
	throw new Error(`Unsupported --workload "${raw}". Expected ${ALL_WORKLOADS.join(', ')}, or all.`)
}

function readSourceFlag(): ReadSourceSelection {
	const raw = readFlag('--read-source')
	if (raw === undefined || raw === 'ascend-writer') return 'ascend-writer'
	if (raw === 'raw-ooxml') return 'raw-ooxml'
	if (raw === 'all') return 'all'
	throw new Error(`Unsupported --read-source "${raw}". Expected ascend-writer, raw-ooxml, or all.`)
}

function readCategoryFlag(): CategorySelection {
	const raw = readFlag('--category')
	if (raw === undefined || raw === 'all') return 'all'
	if (raw === 'read' || raw === 'write') return raw
	throw new Error('Unsupported --category value. Expected read, write, or all.')
}

function readCompetitorFlag(): CompetitorSelection {
	const raw = readFlag('--competitor')
	if (raw === undefined || raw === 'all') return 'all'
	if (raw === 'js' || raw === 'python' || raw === 'external') return raw
	throw new Error('Unsupported --competitor value. Expected js, external, python, or all.')
}

function readValidationModeFlag(): ValidationMode {
	const raw = readFlag('--validation-mode')
	if (raw === undefined || raw === 'each') return 'each'
	if (raw === 'final') return raw
	throw new Error('Unsupported --validation-mode value. Expected each or final.')
}

function readExecutionScopeFlag(): ExecutionScopeSelection {
	const raw = readFlag('--execution-scope')
	if (raw === undefined || raw === 'all') return 'all'
	if (raw === 'in-process' || raw === 'external-process') return raw
	throw new Error(
		'Unsupported --execution-scope value. Expected in-process, external-process, or all.',
	)
}

function readSourceModeFlag(): SourceMode {
	const raw = readFlag('--source-mode')
	if (raw === undefined || raw === 'full') return 'full'
	if (raw === 'generated-write') return raw
	throw new Error('Unsupported --source-mode value. Expected full or generated-write.')
}

function assertSourceModeCompatible(input: {
	readonly sourceMode: SourceMode
	readonly categorySelection: CategorySelection
	readonly executionScopeSelection: ExecutionScopeSelection
}): void {
	if (input.sourceMode !== 'generated-write') return
	if (input.categorySelection !== 'write') {
		throw new Error('--source-mode generated-write requires --category write')
	}
	if (input.executionScopeSelection !== 'external-process') {
		throw new Error('--source-mode generated-write requires --execution-scope external-process')
	}
}

export function parseLibraryAllowlist(raw: string | undefined): LibraryAllowlist {
	if (raw === undefined || raw.trim() === '') return undefined
	const libraries = raw
		.split(',')
		.map((library) => library.trim())
		.filter(Boolean)
	if (libraries.length === 0) return undefined
	return new Set(
		libraries.flatMap((library) => LIBRARY_ALLOWLIST_ALIASES.get(library) ?? [library]),
	)
}

export function libraryAllowed(library: string, allowlist: LibraryAllowlist): boolean {
	return allowlist === undefined || allowlist.has(library)
}

export function competitorMatches(library: string, selection: CompetitorSelection): boolean {
	if (selection === 'all') return true
	if (selection === 'js')
		return library === 'ascend' || library === 'sheetjs' || library === 'exceljs'
	return (
		ASCEND_LIBRARY_ALIASES.includes(library as (typeof ASCEND_LIBRARY_ALIASES)[number]) ||
		library === 'sheetjs-metadata-only' ||
		library === 'openpyxl' ||
		library === 'openpyxl-write-only' ||
		library === 'openpyxl-metadata-only' ||
		library === 'python-calamine-metadata-only' ||
		library === 'xlsxwriter' ||
		library === 'xlsxwriter-constant-memory' ||
		library === 'pyexcelerate' ||
		library === 'pyexcelerate-range' ||
		library === 'pyexcelerate-cell' ||
		library === 'fastxlsx' ||
		library === 'pyopenxlsx' ||
		library === 'pyopenxlsx-bulk' ||
		library === 'pyfastexcel' ||
		library === 'fastexcel' ||
		library === 'fastexcel-java' ||
		library === 'python-calamine' ||
		library === 'rust-calamine' ||
		library === 'polars' ||
		library === 'polars-calamine' ||
		library === 'polars-xlsx2csv' ||
		library === 'polars-openpyxl' ||
		library === 'apache-poi' ||
		library === 'closedxml' ||
		library === 'npoi' ||
		library === 'excelize' ||
		library === 'duckdb-excel' ||
		library === 'rust-xlsxwriter'
	)
}

function workloadRows(workload: WorkloadName, fallbackRows: number | undefined): number {
	if (workload === 'metadata-only') return fallbackRows ?? 200
	return fallbackRows ?? (workload === 'sparse-wide' ? 5000 : DEFAULT_ROWS)
}

function workloadCols(workload: WorkloadName, fallbackCols: number | undefined): number {
	return fallbackCols ?? (workload === 'sparse-wide' ? 256 : DEFAULT_COLS)
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

function getRssBytes(): number | undefined {
	try {
		return process.memoryUsage.rss()
	} catch {
		return undefined
	}
}

function observedPeakRssBytes(values: readonly (number | undefined)[]): number | undefined {
	const defined = values.filter((value): value is number => value !== undefined)
	return defined.length > 0 ? Math.max(...defined) : undefined
}

interface CaseEvaluation {
	readonly status: string
	readonly assertions: Record<string, PrimitiveAssertion>
}

function isRankingEligible(status: string): boolean {
	return status === 'pass'
}

export function hashLines(lines: readonly string[]): string {
	const hash = new Bun.CryptoHasher('sha256')
	for (const line of [...lines].sort()) {
		hash.update(`${line.length}:`)
		hash.update(line)
		hash.update('\n')
	}
	return hash.digest('hex')
}

export function hashOrderedLines(lines: readonly string[]): string {
	const hash = new Bun.CryptoHasher('sha256')
	for (const line of lines) {
		hash.update(`${line.length}:`)
		hash.update(line)
		hash.update('\n')
	}
	return hash.digest('hex')
}

export function expectedDenseValuesHash(rows: number, cols: number): string {
	return expectedWorkloadValuesHash('dense-values', rows, cols)
}

export function expectedStringHeavyValuesHash(rows: number, cols: number): string {
	return expectedWorkloadValuesHash('string-heavy', rows, cols)
}

export function expectedSparseWideValuesHash(rows: number, cols: number): string {
	return expectedWorkloadValuesHash('sparse-wide', rows, cols)
}

export function expectedWorkloadValuesHash(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): string {
	const hash = new Bun.CryptoHasher('sha256')
	const columns = columnNameCache(cols)
		.map((name, col) => ({ name, col }))
		.sort((left, right) => (left.name < right.name ? -1 : left.name > right.name ? 1 : 0))
	const rowOrder = lexicographicRowOrder(rows)
	for (const { name: columnName, col } of columns) {
		for (const row of rowOrder) {
			const value = workloadValue(workloadName, row, col, cols)
			if (value === null) continue
			const line = `Data!${columnName}${row + 1}\t${scalarPayload(value)}`
			hash.update(`${line.length}:`)
			hash.update(line)
			hash.update('\n')
		}
	}
	return hash.digest('hex')
}

export function expectedOrderedWorkloadValuesHash(
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

export function buildWorkloadValues(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): readonly (readonly WorkloadCellValue[])[] {
	return Array.from({ length: rows }, (_, row) =>
		Array.from({ length: cols }, (_, col) => workloadValue(workloadName, row, col, cols)),
	)
}

export function workloadValue(
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
	if (workloadName === 'mixed-closedxml-10text-5number') {
		return col < 10 ? 'Hello world' : col - 10
	}
	if (workloadName === 'plain-text') {
		return `text-${String(row * cols + col).padStart(8, '0')}`
	}
	if (workloadName === 'selected-sheet') return row * cols + col
	if (workloadName === 'metadata-only') return row * cols + col
	if (workloadName === 'warm-workflow') return row * cols + col
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

function workloadValueAssertions(
	sheetCount: number,
	values: readonly string[],
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const semanticCellValuesHash = hashLines(values)
	const orderedSemanticCellValuesHash = hashOrderedLines(values)
	return {
		sheetCount,
		expectedSheetCount: 1,
		sheetCountMatches: sheetCount === 1,
		cellCount: values.length,
		expectedCellCount: input.cells,
		cellCountMatches: values.length === input.cells,
		semanticCellValuesHash,
		expectedSemanticCellValuesHash: input.semanticCellValuesHash,
		orderedSemanticCellValuesHash,
		expectedOrderedSemanticCellValuesHash: input.orderedSemanticCellValuesHash ?? null,
		orderedSemanticCellValuesHashMatches:
			input.orderedSemanticCellValuesHash !== undefined &&
			orderedSemanticCellValuesHash === input.orderedSemanticCellValuesHash,
		semanticCellValuesHashMatches: semanticCellValuesHash === input.semanticCellValuesHash,
	}
}

function coreCellPayload(value: unknown): string | null {
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

export function denseWorkbookAssertions(
	workbook: Workbook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const values: string[] = []
	const sheet = workbook.getSheet('Data')
	if (sheet) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			const payload = coreCellPayload(cell.value)
			if (payload === null) continue
			values.push(`Data!${indexToColumn(col)}${row + 1}\t${payload}`)
		}
	}
	return {
		...workloadValueAssertions(workbook.sheets.length, values, input),
		...workbookFeatureAssertions(workbook, input),
	}
}

function workbookFeatureAssertions(
	workbook: Workbook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const sheet = workbook.getSheet('Data')
	if (input.workloadName === 'styles-heavy') {
		return workbook.styles.size > 1 ? { readStyleCount: workbook.styles.size } : {}
	}
	if (input.workloadName === 'formula-heavy') {
		const formulaCellCount = sheet?.cells.formulaCellCount() ?? 0
		return formulaCellCount > 0 ? { readFormulaCellCount: formulaCellCount } : {}
	}
	if (input.workloadName === 'table-heavy') {
		const tableCount = sheet?.tables.length ?? 0
		return tableCount > 0 ? { readTableCount: tableCount } : {}
	}
	if (input.workloadName === 'feature-rich') {
		const featureCounts = {
			readCommentCount: sheet?.comments.size ?? 0,
			readHyperlinkCount: sheet?.hyperlinks.size ?? 0,
			readDataValidationCount: sheet?.dataValidations.length ?? 0,
			readConditionalFormatCount: sheet?.conditionalFormats.length ?? 0,
			readDefinedNameCount: workbook.definedNames.size,
		}
		return Object.values(featureCounts).some((count) => count > 0)
			? { ...featureCounts, ...coreFeatureRichSemanticAssertions(workbook, input) }
			: {}
	}
	return {}
}

function ascendSelectedSheetAssertions(
	workbook: Awaited<ReturnType<typeof Ascend.open>>,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const info = workbook.inspect()
	return {
		...denseWorkbookAssertions(workbook.getWorkbookModel(), input),
		selectedSheetRead: true,
		sourceSheetCount: info.load.sourceSheets.length,
		loadedSheetCount: info.load.loadedSheets.length,
		loadedSheetNames: info.load.loadedSheets.join(','),
		hasAllSheets: info.load.hasAllSheets,
		cellsHydrated: info.load.cellsHydrated,
	}
}

function ascendMetadataOnlyAssertions(
	workbook: Awaited<ReturnType<typeof Ascend.open>>,
): Record<string, PrimitiveAssertion> {
	const info = workbook.inspect()
	return {
		metadataOnlyRead: true,
		sourceSheetCount: info.load.sourceSheets.length,
		loadedSheetCount: info.load.loadedSheets.length,
		loadedSheetNames: info.load.loadedSheets.join(','),
		hasAllSheets: info.load.hasAllSheets,
		cellsHydrated: info.load.cellsHydrated,
		cellCount: info.cellCount,
		styleCellXfCount: info.styleSummary.cellXfCount,
		workbookViewCount: info.workbookViewCount,
	}
}

function workloadSheetJsAssertions(
	sheetJs: typeof import('xlsx'),
	workbook: import('xlsx').WorkBook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const values: string[] = []
	const worksheet = workbook.Sheets.Data
	if (worksheet) {
		for (const [ref, cell] of Object.entries(worksheet)) {
			if (ref.startsWith('!')) continue
			const decoded = sheetJs.utils.decode_cell(ref)
			const payload = scalarPayload((cell as { readonly v?: unknown }).v)
			if (payload === null) continue
			values.push(`Data!${indexToColumn(decoded.c)}${decoded.r + 1}\t${payload}`)
		}
	}
	return {
		...workloadValueAssertions(workbook.SheetNames.length, values, input),
		...sheetJsReadFeatureAssertions(workbook, input),
	}
}

function sheetJsSelectedSheetAssertions(
	sheetJs: typeof import('xlsx'),
	workbook: import('xlsx').WorkBook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const loadedSheetNames = Object.keys(workbook.Sheets)
	return {
		...workloadSheetJsAssertions(sheetJs, workbook, input),
		sheetCount: loadedSheetNames.length,
		selectedSheetRead: true,
		sourceSheetCount: workbook.SheetNames.length,
		loadedSheetCount: loadedSheetNames.length,
		loadedSheetNames: loadedSheetNames.join(','),
		hasAllSheets: loadedSheetNames.length === workbook.SheetNames.length,
		cellsHydrated: loadedSheetNames.length > 0,
	}
}

function workloadExcelJsAssertions(
	workbook: import('exceljs').Workbook,
	input: CompetitiveDataSet,
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
	return {
		...workloadValueAssertions(workbook.worksheets.length, values, input),
		...excelJsReadFeatureAssertions(workbook, input),
	}
}

function sheetJsReadFeatureAssertions(
	workbook: import('xlsx').WorkBook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	if (input.workloadName !== 'feature-rich') return {}
	const packageAssertions = sheetJsPackageFeatureAssertions(workbook, input)
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
		...packageAssertions,
		readCommentCount: Math.max(readCommentCount, Number(packageAssertions.readCommentCount ?? 0)),
		readHyperlinkCount: Math.max(
			readHyperlinkCount,
			Number(packageAssertions.readHyperlinkCount ?? 0),
		),
		readDataValidationCount: Number(packageAssertions.readDataValidationCount ?? 0),
		readConditionalFormatCount: Number(packageAssertions.readConditionalFormatCount ?? 0),
		readDefinedNameCount: Math.max(
			workbook.Workbook?.Names?.length ?? 0,
			Number(packageAssertions.readDefinedNameCount ?? 0),
		),
	}
}

function sheetJsPackageFeatureAssertions(
	workbook: import('xlsx').WorkBook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const files = (
		workbook as unknown as {
			readonly files?: Record<string, { readonly content?: Uint8Array | string }>
		}
	).files
	if (!files) return {}
	const workbookXml = sheetJsPackageFileText(files, 'xl/workbook.xml')
	const sheetXml = sheetJsPackageFileText(files, 'xl/worksheets/sheet1.xml')
	const sheetRelsXml = sheetJsPackageFileText(files, 'xl/worksheets/_rels/sheet1.xml.rels')
	const commentsXml = sheetJsPackageFileText(files, 'xl/comments1.xml')
	if (!workbookXml || !sheetXml) return {}
	const expected = expectedFeatureRichContract(input)
	const definedNameRef = definedNameValue(workbookXml, 'FeatureRange')
	const hyperlink = firstElementAttributes(sheetXml, 'hyperlink', 'ref', expected.hyperlinkRef)
	const hyperlinkRelId = hyperlink.get('r:id') ?? hyperlink.get('id') ?? ''
	const hyperlinkTarget = relationshipTarget(sheetRelsXml, hyperlinkRelId)
	const comment = commentEntry(commentsXml, expected.commentRef)
	const dataValidation = firstElementAttributes(sheetXml, 'dataValidation')
	const conditionalFormatting = elementWithAttributesAndBody(sheetXml, 'conditionalFormatting', {
		sqref: expected.conditionalFormatRef,
	})
	const cfRule = conditionalFormatting
		? firstElementAttributes(conditionalFormatting.body, 'cfRule')
		: new Map<string, string>()
	const definedNameMatches = normalizeFormulaRef(definedNameRef) === expected.featureRange
	const hyperlinkMatches =
		hyperlink.get('ref') === expected.hyperlinkRef &&
		(hyperlink.get('display') === undefined ||
			hyperlink.get('display') === expected.hyperlinkDisplay) &&
		hyperlink.get('tooltip') === expected.hyperlinkTooltip &&
		hyperlinkTarget === expected.hyperlinkTarget
	const commentMatches =
		comment.ref === expected.commentRef &&
		comment.author === expected.commentAuthor &&
		comment.text === expected.commentText
	const dataValidationMatches =
		dataValidation.get('type') === 'list' &&
		dataValidation.get('allowBlank') === '1' &&
		dataValidation.get('showInputMessage') === '1' &&
		dataValidation.get('sqref') === expected.validationRef &&
		tagText(sheetXml, 'formula1') === expected.validationFormula
	const conditionalFormattingMatches =
		conditionalFormatting !== null &&
		cfRule.get('type') === 'cellIs' &&
		cfRule.get('operator') === 'greaterThan' &&
		tagText(conditionalFormatting.body, 'formula') === expected.conditionalFormula
	return {
		readCommentCount: countXmlElements(commentsXml, 'comment'),
		readHyperlinkCount: countXmlElements(sheetXml, 'hyperlink'),
		readDataValidationCount: countXmlElements(sheetXml, 'dataValidation'),
		readConditionalFormatCount: countXmlElements(sheetXml, 'conditionalFormatting'),
		readDefinedNameCount: countXmlElements(workbookXml, 'definedName'),
		readFeatureRichSemanticMatches:
			definedNameMatches &&
			hyperlinkMatches &&
			commentMatches &&
			dataValidationMatches &&
			conditionalFormattingMatches,
		readFeatureRichDefinedNameMatches: definedNameMatches,
		readFeatureRichHyperlinkMatches: hyperlinkMatches,
		readFeatureRichCommentMatches: commentMatches,
		readFeatureRichDataValidationMatches: dataValidationMatches,
		readFeatureRichConditionalFormattingMatches: conditionalFormattingMatches,
	}
}

function sheetJsPackageFileText(
	files: Record<string, { readonly content?: Uint8Array | string }>,
	path: string,
): string {
	const content = files[path]?.content
	if (typeof content === 'string') return content
	if (content instanceof Uint8Array) return new TextDecoder().decode(content)
	return ''
}

function countXmlElements(xml: string, tagName: string): number {
	if (!xml) return 0
	return [...xml.matchAll(new RegExp(`<${tagName}\\b`, 'g'))].length
}

function excelJsReadFeatureAssertions(
	workbook: import('exceljs').Workbook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	if (input.workloadName !== 'feature-rich') return {}
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

function sheetJsMetadataOnlyAssertions(
	workbook: import('xlsx').WorkBook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	return {
		metadataOnlyRead: true,
		sourceSheetCount: workbook.SheetNames.length,
		loadedSheetCount: workbook.SheetNames.length,
		loadedSheetNames: workbook.SheetNames.join(','),
		hasAllSheets: true,
		cellsHydrated: false,
		expectedSourceSheetCount: expectedSourceSheetCount(input),
	}
}

export function denseWriteAssertions(
	bytes: Uint8Array,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const fastAssertions = generatedWriteAssertions(bytes, input)
	if (fastAssertions) return fastAssertions
	const read = readXlsx(bytes, { mode: 'values' })
	if (!read.ok) {
		return {
			bytes: bytes.byteLength,
			reopenOk: false,
			reopenError: read.error.message,
			sheetCountMatches: false,
			cellCountMatches: false,
			expectedSemanticCellValuesHash: input.semanticCellValuesHash,
			semanticCellValuesHashMatches: false,
		}
	}
	return {
		bytes: bytes.byteLength,
		reopenOk: true,
		formulaCount: xlsxWorksheetFormulaCount(bytes),
		tablePartCount: xlsxTablePartCount(bytes),
		...featureRichAssertions(bytes, input),
		...denseWorkbookAssertions(read.value.workbook, input),
	}
}

function generatedWriteAssertions(
	bytes: Uint8Array,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> | undefined {
	if (input.sourceMode !== 'generated-write') return undefined
	if (!canUseFastGeneratedWriteAssertions(input.workloadName)) return undefined
	try {
		const zip = extractZip(bytes)
		const workbookXml = zip.readText('xl/workbook.xml')
		const sheetXml = zip.readText('xl/worksheets/sheet1.xml')
		if (!workbookXml || !sheetXml) return undefined
		const sheetCount = countWorkbookSheets(workbookXml)
		const observed = hashGeneratedWorksheetValues(sheetXml, input.cols)
		if (!observed) return undefined
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
			semanticCellValuesHash: FAST_GENERATED_WRITE_HASH_NOT_COMPUTED,
			expectedSemanticCellValuesHash: input.semanticCellValuesHash,
			orderedSemanticCellValuesHash: observed.orderedSemanticCellValuesHash,
			expectedOrderedSemanticCellValuesHash: input.orderedSemanticCellValuesHash ?? null,
			orderedSemanticCellValuesHashMatches:
				input.orderedSemanticCellValuesHash !== undefined &&
				observed.orderedSemanticCellValuesHash === input.orderedSemanticCellValuesHash,
			semanticCellValuesHashMatches:
				input.orderedSemanticCellValuesHash !== undefined &&
				observed.orderedSemanticCellValuesHash === input.orderedSemanticCellValuesHash,
			fastGeneratedWriteValidation: true,
		}
	} catch {
		return undefined
	}
}

function canUseFastGeneratedWriteAssertions(workloadName: WorkloadName): boolean {
	return (
		workloadName === 'dense-values' ||
		workloadName === 'mixed-10pct-text' ||
		workloadName === 'mixed-50pct-text' ||
		workloadName === 'mixed-closedxml-10text-5number' ||
		workloadName === 'plain-text' ||
		workloadName === 'string-heavy' ||
		workloadName === 'sparse-wide'
	)
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

function lexicographicRowOrder(rows: number): readonly number[] {
	return Array.from({ length: Math.max(0, rows) }, (_, row) => row).sort((left, right) => {
		const leftText = `${left + 1}`
		const rightText = `${right + 1}`
		return leftText < rightText ? -1 : leftText > rightText ? 1 : 0
	})
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
	const type = /\bt="([^"]+)"/.exec(attrs)?.[1]
	if (type === 's') return undefined
	if (type === 'inlineStr') {
		const text = firstTagText(body, 't')
		return text === undefined ? null : `s:${decodeXmlText(text)}`
	}
	const value = firstTagText(body, 'v')
	if (value === undefined) return null
	if (type === 'b') return value === '1' || value === 'true' ? 'b:true' : 'b:false'
	if (type === 'str') return `s:${decodeXmlText(value)}`
	return `n:${value}`
}

function firstTagText(xml: string, tagName: string): string | undefined {
	const open = xml.indexOf(`<${tagName}>`)
	if (open < 0) return undefined
	const start = open + tagName.length + 2
	const end = xml.indexOf(`</${tagName}>`, start)
	return end < 0 ? undefined : xml.slice(start, end)
}

function decodeXmlText(text: string): string {
	return text
		.replaceAll('&lt;', '<')
		.replaceAll('&gt;', '>')
		.replaceAll('&quot;', '"')
		.replaceAll('&apos;', "'")
		.replaceAll('&amp;', '&')
}

function parseXmlAttributes(attrsText: string): Map<string, string> {
	const attrs = new Map<string, string>()
	for (const match of attrsText.matchAll(/([A-Za-z_][\w:.-]*)="([^"]*)"/g)) {
		const key = match[1]
		const value = match[2]
		if (key !== undefined && value !== undefined) attrs.set(key, decodeXmlText(value))
	}
	return attrs
}

function firstElementAttributes(
	xml: string,
	tagName: string,
	attrName?: string,
	attrValue?: string,
): Map<string, string> {
	for (const match of xml.matchAll(new RegExp(`<${tagName}\\b([^>]*)>`, 'g'))) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		if (attrName === undefined || attrs.get(attrName) === attrValue) return attrs
	}
	return new Map()
}

function elementWithAttributesAndBody(
	xml: string,
	tagName: string,
	expectedAttrs: Readonly<Record<string, string>>,
): { readonly attrs: ReadonlyMap<string, string>; readonly body: string } | null {
	const re = new RegExp(`<${tagName}\\b([^>]*)>([\\s\\S]*?)<\\/${tagName}>`, 'g')
	for (const match of xml.matchAll(re)) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		if (Object.entries(expectedAttrs).every(([key, value]) => attrs.get(key) === value)) {
			return { attrs, body: match[2] ?? '' }
		}
	}
	return null
}

function relationshipTarget(relsXml: string, relId: string): string {
	if (!relId) return ''
	for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		if (attrs.get('Id') === relId) return attrs.get('Target') ?? ''
	}
	return ''
}

function definedNameValue(workbookXml: string, name: string): string {
	const re = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g
	for (const match of workbookXml.matchAll(re)) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		if (attrs.get('name') === name) return decodeXmlText(match[2] ?? '')
	}
	return ''
}

function commentEntry(
	commentsXml: string,
	expectedRef: string,
): { readonly ref: string; readonly author: string; readonly text: string } {
	const authors = [...commentsXml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((match) =>
		decodeXmlText(match[1] ?? ''),
	)
	const commentRe = /<comment\b([^>]*)>([\s\S]*?)<\/comment>/g
	for (const match of commentsXml.matchAll(commentRe)) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		const ref = attrs.get('ref') ?? ''
		if (ref !== expectedRef) continue
		const authorId = Number.parseInt(attrs.get('authorId') ?? '', 10)
		const author = Number.isFinite(authorId) ? (authors[authorId] ?? '') : ''
		const text = [...(match[2] ?? '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
			.map((textMatch) => decodeXmlText(textMatch[1] ?? ''))
			.join('')
		return { ref, author, text }
	}
	return { ref: '', author: '', text: '' }
}

function normalizeFormulaRef(value: string): string {
	return value.startsWith('=') ? value.slice(1) : value
}

function tagText(xml: string, tagName: string): string {
	return decodeXmlText(firstTagText(xml, tagName) ?? '')
}

export function evaluateAssertions(
	category: 'read' | 'write',
	input: CompetitiveDataSet,
	assertions: Record<string, PrimitiveAssertion> | undefined,
): CaseEvaluation {
	const observed = assertions ?? {}
	if (observed.metadataOnlyRead === true) {
		const sourceSheetCountMatches = observed.sourceSheetCount === expectedSourceSheetCount(input)
		const loadedSheetCountMatches = observed.loadedSheetCount === expectedSourceSheetCount(input)
		const cellsNotHydrated = observed.cellsHydrated === false
		const matches = sourceSheetCountMatches && loadedSheetCountMatches && cellsNotHydrated
		return {
			status: matches ? 'pass' : 'semantic-mismatch',
			assertions: {
				...observed,
				expectedSourceSheetCount: expectedSourceSheetCount(input),
				sourceSheetCountMatches,
				loadedSheetCountMatches,
				cellsNotHydrated,
			},
		}
	}
	const sheetCountMatches = observed.sheetCountMatches === true || observed.sheetCount === 1
	const cellCountMatches = observed.cellCountMatches === true || observed.cellCount === input.cells
	const sortedSemanticCellValuesHashMatches =
		observed.semanticCellValuesHash === input.semanticCellValuesHash
	const orderedSemanticCellValuesHashMatches =
		input.orderedSemanticCellValuesHash !== undefined &&
		observed.orderedSemanticCellValuesHash === input.orderedSemanticCellValuesHash
	const semanticCellValuesHashMatches =
		sortedSemanticCellValuesHashMatches || orderedSemanticCellValuesHashMatches
	const selectedSheetMatches =
		observed.selectedSheetRead !== true ||
		(observed.sourceSheetCount === expectedSourceSheetCount(input) &&
			observed.loadedSheetCount === 1 &&
			observed.loadedSheetNames === 'Data' &&
			observed.hasAllSheets === false &&
			observed.cellsHydrated === true)
	const reopenOk = category === 'read' || observed.reopenOk !== false
	const tablePartMatches =
		category !== 'write' || input.workloadName !== 'table-heavy' || observed.tablePartCount > 0
	const expectedFormulaCount =
		input.workloadName === 'formula-heavy' ? input.rows * Math.max(0, input.cols - 2) : 0
	const formulaCountMatches =
		category !== 'write' ||
		input.workloadName !== 'formula-heavy' ||
		observed.formulaCount === expectedFormulaCount
	const featureRichMatches =
		category !== 'write' ||
		input.workloadName !== 'feature-rich' ||
		(observed.commentPartCount === 1 &&
			observed.vmlDrawingPartCount === 1 &&
			observed.worksheetHyperlinkCount === 1 &&
			observed.worksheetDataValidationCount === 1 &&
			observed.worksheetConditionalFormattingCount === 1 &&
			observed.definedNameCount === 1 &&
			observed.featureRichSemanticMatches === true)
	const readFeatureRichMatches =
		category !== 'read' ||
		input.workloadName !== 'feature-rich' ||
		(Number(observed.readCommentCount) > 0 &&
			Number(observed.readHyperlinkCount) > 0 &&
			Number(observed.readDataValidationCount) > 0 &&
			Number(observed.readConditionalFormatCount) > 0 &&
			Number(observed.readDefinedNameCount) > 0 &&
			observed.readFeatureRichSemanticMatches !== false)
	const matches =
		reopenOk &&
		sheetCountMatches &&
		cellCountMatches &&
		semanticCellValuesHashMatches &&
		selectedSheetMatches &&
		tablePartMatches &&
		formulaCountMatches &&
		featureRichMatches &&
		readFeatureRichMatches
	return {
		status: matches ? 'pass' : 'semantic-mismatch',
		assertions: {
			...observed,
			expectedSheetCount: 1,
			expectedCellCount: input.cells,
			expectedSemanticCellValuesHash: input.semanticCellValuesHash,
			expectedOrderedSemanticCellValuesHash: input.orderedSemanticCellValuesHash ?? null,
			reopenOk,
			sheetCountMatches,
			cellCountMatches,
			sortedSemanticCellValuesHashMatches,
			orderedSemanticCellValuesHashMatches,
			semanticCellValuesHashMatches,
			selectedSheetMatches,
			tablePartMatches,
			expectedFormulaCount,
			formulaCountMatches,
			featureRichMatches,
			readFeatureRichMatches,
		},
	}
}

function expectedSourceSheetCount(input: CompetitiveDataSet): number {
	return hasAuxiliarySheets(input.workloadName) ? 3 : 1
}

function xlsxTablePartCount(bytes: Uint8Array): number {
	try {
		const entries = unzipSync(bytes)
		return Object.keys(entries).filter((path) => /^xl\/tables\/.+\.xml$/.test(path)).length
	} catch {
		return 0
	}
}

function xlsxWorksheetFormulaCount(bytes: Uint8Array): number {
	try {
		const entries = unzipSync(bytes)
		let count = 0
		for (const [path, entry] of Object.entries(entries)) {
			if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) continue
			const xml = new TextDecoder().decode(entry)
			for (const _match of xml.matchAll(/<f\b/g)) count++
		}
		return count
	} catch {
		return 0
	}
}

function coreFeatureRichSemanticAssertions(
	workbook: Workbook,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	if (input.workloadName !== 'feature-rich') return {}
	const sheet = workbook.getSheet('Data')
	const expected = expectedFeatureRichContract(input)
	const definedNameRef = workbook.definedNames.get('FeatureRange') ?? ''
	const hyperlink = sheet?.hyperlinks.get('A1')
	const comment = sheet?.comments.get('B2')
	const validation = sheet?.dataValidations.find((entry) => entry.sqref === expected.validationRef)
	const conditionalFormat = sheet?.conditionalFormats.find(
		(entry) => entry.sqref === expected.conditionalFormatRef,
	)
	const conditionalRule = conditionalFormat?.rules[0]
	const definedNameMatches = normalizeFormulaRef(definedNameRef) === expected.featureRange
	const hyperlinkMatches =
		hyperlink?.target === expected.hyperlinkTarget &&
		hyperlink.display === expected.hyperlinkDisplay &&
		hyperlink.tooltip === expected.hyperlinkTooltip
	const commentMatches =
		comment?.text === expected.commentText && comment.author === expected.commentAuthor
	const dataValidationMatches =
		validation?.type === 'list' &&
		validation.allowBlank === true &&
		validation.showInputMessage === true &&
		validation.formula1 === expected.validationFormula
	const conditionalFormattingMatches =
		conditionalRule?.type === 'cellIs' &&
		conditionalRule.operator === 'greaterThan' &&
		conditionalRule.formulas[0] === expected.conditionalFormula
	const semanticMatches =
		definedNameMatches &&
		hyperlinkMatches &&
		commentMatches &&
		dataValidationMatches &&
		conditionalFormattingMatches
	return {
		readFeatureRichSemanticMatches: semanticMatches,
		readFeatureRichDefinedNameMatches: definedNameMatches,
		readFeatureRichHyperlinkMatches: hyperlinkMatches,
		readFeatureRichCommentMatches: commentMatches,
		readFeatureRichDataValidationMatches: dataValidationMatches,
		readFeatureRichConditionalFormattingMatches: conditionalFormattingMatches,
	}
}

function featureRichAssertions(
	bytes: Uint8Array,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	if (input.workloadName !== 'feature-rich') return {}
	const summary = extractWorkbookFeatureSummary(bytes)
	return {
		commentPartCount: summary.commentPartCount,
		vmlDrawingPartCount: summary.vmlDrawingPartCount,
		worksheetHyperlinkCount: summary.worksheetHyperlinkCount,
		worksheetDataValidationCount: summary.worksheetDataValidationCount,
		worksheetConditionalFormattingCount: summary.worksheetConditionalFormattingCount,
		definedNameCount: summary.definedNameCount,
		...xlsxFeatureRichSemanticAssertions(bytes, input),
	}
}

function xlsxFeatureRichSemanticAssertions(
	bytes: Uint8Array,
	input: CompetitiveDataSet,
): Record<string, PrimitiveAssertion> {
	const expected = expectedFeatureRichContract(input)
	try {
		const archive = extractZip(bytes)
		const workbookXml = archive.readText('xl/workbook.xml') ?? ''
		const sheetXml = archive.readText('xl/worksheets/sheet1.xml') ?? ''
		const sheetRelsXml = archive.readText('xl/worksheets/_rels/sheet1.xml.rels') ?? ''
		const commentsXml = archive.readText('xl/comments1.xml') ?? ''
		const definedNameRef = definedNameValue(workbookXml, 'FeatureRange')
		const hyperlink = firstElementAttributes(sheetXml, 'hyperlink', 'ref', 'A1')
		const hyperlinkRelId = hyperlink.get('r:id') ?? hyperlink.get('id') ?? ''
		const hyperlinkTarget = relationshipTarget(sheetRelsXml, hyperlinkRelId)
		const dataValidation = elementWithAttributesAndBody(sheetXml, 'dataValidation', {
			sqref: expected.validationRef,
		})
		const conditionalFormatting = elementWithAttributesAndBody(sheetXml, 'conditionalFormatting', {
			sqref: expected.conditionalFormatRef,
		})
		const cfRule = conditionalFormatting
			? firstElementAttributes(conditionalFormatting.body, 'cfRule')
			: new Map<string, string>()
		const comment = commentEntry(commentsXml, expected.commentRef)
		const definedNameMatches = normalizeFormulaRef(definedNameRef) === expected.featureRange
		const hyperlinkMatches =
			hyperlink.get('ref') === expected.hyperlinkRef &&
			(hyperlink.get('display') === undefined ||
				hyperlink.get('display') === expected.hyperlinkDisplay) &&
			hyperlink.get('tooltip') === expected.hyperlinkTooltip &&
			hyperlinkTarget === expected.hyperlinkTarget
		const commentMatches =
			comment.ref === expected.commentRef &&
			comment.author === expected.commentAuthor &&
			comment.text === expected.commentText
		const dataValidationMatches =
			dataValidation?.attrs.get('type') === 'list' &&
			dataValidation.attrs.get('allowBlank') === '1' &&
			dataValidation.attrs.get('showInputMessage') === '1' &&
			tagText(dataValidation.body, 'formula1') === expected.validationFormula
		const conditionalFormattingMatches =
			conditionalFormatting !== null &&
			cfRule.get('type') === 'cellIs' &&
			cfRule.get('operator') === 'greaterThan' &&
			tagText(conditionalFormatting.body, 'formula') === expected.conditionalFormula
		const semanticMatches =
			definedNameMatches &&
			hyperlinkMatches &&
			commentMatches &&
			dataValidationMatches &&
			conditionalFormattingMatches
		return {
			featureRichSemanticMatches: semanticMatches,
			featureRichDefinedNameMatches: definedNameMatches,
			featureRichHyperlinkMatches: hyperlinkMatches,
			featureRichCommentMatches: commentMatches,
			featureRichDataValidationMatches: dataValidationMatches,
			featureRichConditionalFormattingMatches: conditionalFormattingMatches,
			featureRichDefinedNameRef: normalizeFormulaRef(definedNameRef),
			featureRichHyperlinkTarget: hyperlinkTarget,
			featureRichCommentText: comment.text,
			featureRichDataValidationFormula: tagText(dataValidation?.body ?? '', 'formula1'),
			featureRichConditionalFormula: tagText(conditionalFormatting?.body ?? '', 'formula'),
		}
	} catch {
		return {
			featureRichSemanticMatches: false,
			featureRichDefinedNameMatches: false,
			featureRichHyperlinkMatches: false,
			featureRichCommentMatches: false,
			featureRichDataValidationMatches: false,
			featureRichConditionalFormattingMatches: false,
		}
	}
}

function expectedFeatureRichContract(input: CompetitiveDataSet): {
	readonly featureRange: string
	readonly hyperlinkRef: string
	readonly hyperlinkDisplay: string
	readonly hyperlinkTarget: string
	readonly hyperlinkTooltip: string
	readonly commentRef: string
	readonly commentAuthor: string
	readonly commentText: string
	readonly validationRef: string
	readonly validationFormula: string
	readonly conditionalFormatRef: string
	readonly conditionalFormula: string
} {
	return {
		featureRange: featureRange(input),
		hyperlinkRef: 'A1',
		hyperlinkDisplay: 'Ascend',
		hyperlinkTarget: 'https://example.com/ascend',
		hyperlinkTooltip: 'Open Ascend',
		commentRef: 'B2',
		commentAuthor: 'Ascend',
		commentText: 'Review',
		validationRef: `C2:C${Math.max(2, input.rows)}`,
		validationFormula: '"Q1,Q2,Q3"',
		conditionalFormatRef: `A1:A${Math.max(1, input.rows)}`,
		conditionalFormula: '0',
	}
}

export function coalesceRepeatCorrectnessStatus(statuses: readonly string[]): string {
	if (statuses.length === 0) return 'not-evaluated'
	const first = statuses[0] ?? 'not-evaluated'
	return statuses.every((status) => status === first) ? first : 'intermittent-mismatch'
}

function operationProfile(
	benchmarkCase: Pick<
		CompetitiveCase,
		'category' | 'executionScope' | 'capabilities' | 'operationProfile'
	>,
	input?: Pick<CompetitiveDataSet, 'workloadName'>,
): string {
	if (benchmarkCase.operationProfile) return benchmarkCase.operationProfile
	if (benchmarkCase.category === 'write') return writeOperationProfile(input?.workloadName)
	if (
		benchmarkCase.executionScope === 'external-process' &&
		benchmarkCase.capabilities?.valueOnlyRead !== true
	) {
		return 'read-formula-preserving'
	}
	return 'read-values'
}

export function writeOperationProfile(workloadName: WorkloadName | undefined): string {
	if (workloadName === 'styles-heavy') return 'write-styles'
	if (workloadName === 'formula-heavy') return 'write-formulas'
	if (workloadName === 'table-heavy') return 'write-tables'
	if (workloadName === 'feature-rich') return 'write-rich-metadata'
	return 'write-values'
}

function timingLane(
	benchmarkCase: Pick<CompetitiveCase, 'executionScope' | 'timingModel'>,
	input: CompetitiveDataSet,
): string {
	if (benchmarkCase.executionScope === 'external-process' && benchmarkCase.timingModel) {
		return `${benchmarkCase.timingModel}:${input.workloadName}`
	}
	return benchmarkCase.executionScope === 'external-process'
		? `external-internal-generated-${input.workloadName}`
		: `in-process-generated-${input.workloadName}`
}

function timingModel(benchmarkCase: {
	readonly executionScope?: CompetitiveCase['executionScope']
	readonly runBatched?: CompetitiveCase['runBatched']
	readonly timingModel?: string
}): string {
	if (benchmarkCase.timingModel) return benchmarkCase.timingModel
	if (benchmarkCase.executionScope === 'external-process') {
		return 'external-internal-operation-timing'
	}
	return benchmarkCase.runBatched ? 'operation-only' : 'operation-plus-validation'
}

function validationModel(
	benchmarkCase: Pick<CompetitiveCase, 'executionScope' | 'validationModel'>,
): string {
	if (benchmarkCase.validationModel) return benchmarkCase.validationModel
	return benchmarkCase.executionScope === 'external-process'
		? 'external-post-operation-assertions'
		: 'inline-assertions'
}

function coreCellValue(
	value: ScalarCellValue,
):
	| { kind: 'number'; value: number }
	| { kind: 'string'; value: string }
	| { kind: 'boolean'; value: boolean } {
	if (typeof value === 'number') return { kind: 'number', value }
	if (typeof value === 'boolean') return { kind: 'boolean', value }
	return { kind: 'string', value }
}

export function setCoreCell(
	workbook: Workbook,
	values: readonly (readonly WorkloadCellValue[])[],
	rows: number,
	cols: number,
	workloadName: WorkloadName = 'dense-values',
): void {
	setCoreCellFromSource(workbook, rows, cols, workloadName, (row, col) => values[row]?.[col])
}

export function setCoreCellGenerated(
	workbook: Workbook,
	rows: number,
	cols: number,
	workloadName: WorkloadName = 'dense-values',
): void {
	if (workloadName === 'sparse-wide') {
		setSparseWideCoreCells(workbook, rows, cols)
		return
	}
	setCoreCellFromSource(workbook, rows, cols, workloadName, (row, col) =>
		workloadValue(workloadName, row, col, cols),
	)
}

function setSparseWideCoreCells(workbook: Workbook, rows: number, cols: number): void {
	const sheet = workbook.addSheet('Data')
	if (cols <= 0) return
	const styleIds = workloadStyleIds(workbook, 'sparse-wide')
	for (let row = 0; row < rows; row++) {
		setCoreCellValue(sheet, styleIds, 'sparse-wide', row, 0, row)
		for (let col = 1; col < cols - 1; col++) {
			if ((row * 31 + col * 17) % 97 === 0) {
				setCoreCellValue(sheet, styleIds, 'sparse-wide', row, col, row * cols + col)
			}
		}
		if (cols > 1) {
			setCoreCellValue(sheet, styleIds, 'sparse-wide', row, cols - 1, `edge-${row}-${cols}`)
		}
	}
}

function setCoreCellFromSource(
	workbook: Workbook,
	rows: number,
	cols: number,
	workloadName: WorkloadName,
	valueAt: (row: number, col: number) => WorkloadCellValue | undefined,
): void {
	const sheet = workbook.addSheet('Data')
	if (workloadName !== 'sparse-wide') sheet.cells.setExpectedDensity('dense')
	const styleIds = workloadStyleIds(workbook, workloadName)
	for (let row = 0; row < rows; row++) {
		for (let col = 0; col < cols; col++) {
			const value = valueAt(row, col)
			if (value === undefined || value === null) continue
			setCoreCellValue(sheet, styleIds, workloadName, row, col, value)
		}
	}
	if (workloadName === 'table-heavy' && rows > 0 && cols > 0) {
		sheet.tables.push({
			id: createTableId(),
			name: 'DataTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: rows - 1, col: cols - 1 } },
			columns: Array.from({ length: cols }, (_, index) => ({ name: `Column ${index + 1}` })),
			hasHeaders: true,
			hasTotals: false,
			tableStyleInfo: { name: 'TableStyleMedium2', showRowStripes: true },
		})
	}
	if (workloadName === 'feature-rich') {
		workbook.definedNames.set(
			'FeatureRange',
			`Data!$A$1:$${indexToColumn(Math.max(0, cols - 1))}$${Math.max(1, rows)}`,
		)
		sheet.hyperlinks.set('A1', {
			target: 'https://example.com/ascend',
			display: 'Ascend',
			tooltip: 'Open Ascend',
		})
		if (rows > 1 && cols > 1) sheet.comments.set('B2', { text: 'Review', author: 'Ascend' })
		if (rows > 1 && cols > 2) {
			sheet.dataValidations.push({
				sqref: `C2:C${Math.max(2, rows)}`,
				type: 'list',
				allowBlank: true,
				showInputMessage: true,
				formula1: '"Q1,Q2,Q3"',
			})
		}
		if (rows > 0 && cols > 0) {
			workbook.differentialStyles.push({
				font: { bold: true },
				fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
			})
			sheet.conditionalFormats.push({
				sqref: `A1:A${Math.max(1, rows)}`,
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						dxfId: workbook.differentialStyles.length - 1,
						priority: 1,
						formulas: ['0'],
						style: workbook.differentialStyles[workbook.differentialStyles.length - 1],
					},
				],
			})
		}
	}
	if (hasAuxiliarySheets(workloadName)) {
		const summary = workbook.addSheet('Summary')
		const archive = workbook.addSheet('Archive')
		for (let row = 0; row < Math.min(rows, 100); row++) {
			summary.cells.set(row, 0, {
				value: coreCellValue(row),
				formula: null,
				styleId: SID,
			})
			archive.cells.set(row, 0, {
				value: coreCellValue(`archive-${row}`),
				formula: null,
				styleId: SID,
			})
		}
	}
}

function setCoreCellValue(
	sheet: Sheet,
	styleIds: readonly StyleId[],
	workloadName: WorkloadName,
	row: number,
	col: number,
	value: ScalarCellValue,
): void {
	const formula = formulaForWorkload(workloadName, row, col)
	const styleId = styleIds[(row + col) % styleIds.length] ?? SID
	if (typeof value === 'number') {
		sheet.cells.setNumberResolved(row, col, value, formula, styleId)
	} else if (typeof value === 'string') {
		sheet.cells.setStringResolved(row, col, value, formula, styleId)
	} else {
		sheet.cells.set(row, col, {
			value: coreCellValue(value),
			formula,
			styleId,
		})
	}
}

function hasAuxiliarySheets(workloadName: WorkloadName): boolean {
	return workloadName === 'selected-sheet' || workloadName === 'metadata-only'
}

function workloadStyleIds(workbook: Workbook, workloadName: WorkloadName): readonly StyleId[] {
	if (workloadName !== 'styles-heavy') return [SID]
	return [
		SID,
		workbook.styles.register({ font: { bold: true } }),
		workbook.styles.register({ numberFormat: '#,##0.00' }),
		workbook.styles.register({
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFE2F0D9' } },
		}),
		workbook.styles.register({
			font: { color: { kind: 'rgb', rgb: 'FF1F4E79' } },
			alignment: { horizontal: 'center' },
		}),
	]
}

function formulaForWorkload(workloadName: WorkloadName, row: number, col: number): string | null {
	if (workloadName !== 'formula-heavy' || col < 2) return null
	const currentRow = row + 1
	return `A${currentRow}+B${currentRow}+${col}`
}

interface RawOoxmlValueSource {
	readonly rows: number
	readonly cols: number
	valueAt(row: number, col: number): WorkloadCellValue | undefined
}

function buildRawOoxmlXlsx(workloadName: WorkloadName, source: RawOoxmlValueSource): Uint8Array {
	const { rows, cols } = source
	const columnNames = columnNameCache(cols)
	const sheetRows: string[] = []
	for (let row = 0; row < rows; row++) {
		let rowXml = `<row r="${row + 1}">`
		let hasCells = false
		for (let col = 0; col < cols; col++) {
			const value = source.valueAt(row, col)
			if (value === undefined || value === null) continue
			const ref = `${columnNames[col] ?? indexToColumn(col)}${row + 1}`
			const styleAttr = workloadName === 'styles-heavy' ? ` s="${((row + col) % 4) + 1}"` : ''
			const formula = formulaForWorkload(workloadName, row, col)
			const formulaXml = formula ? `<f>${escapeXml(formula)}</f>` : ''
			if (typeof value === 'number') {
				rowXml += `<c r="${ref}"${styleAttr}>${formulaXml}<v>${value}</v></c>`
			} else if (typeof value === 'boolean') {
				rowXml += `<c r="${ref}" t="b"${styleAttr}>${formulaXml}<v>${value ? 1 : 0}</v></c>`
			} else {
				rowXml += `<c r="${ref}" t="inlineStr"${styleAttr}>${formulaXml}<is><t>${escapeXml(value)}</t></is></c>`
			}
			hasCells = true
		}
		if (hasCells) sheetRows.push(`${rowXml}</row>`)
	}
	const hasStyles = workloadName === 'styles-heavy'
	const hasTable = workloadName === 'table-heavy' && rows > 0 && cols > 0
	const hasFeatureRich = workloadName === 'feature-rich'
	const hasAuxiliary = hasAuxiliarySheets(workloadName)
	const tableRef = `A1:${indexToColumn(Math.max(0, cols - 1))}${Math.max(1, rows)}`
	return zipSync({
		'[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasFeatureRich ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : ''}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  ${hasAuxiliary ? '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' : ''}
  ${hasAuxiliary ? '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' : ''}
  ${hasStyles ? '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' : ''}
  ${hasTable ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ''}
  ${hasFeatureRich ? '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' : ''}
</Types>`),
		'_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
		'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  ${hasAuxiliary ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' : ''}
  ${hasAuxiliary ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>' : ''}
  ${hasStyles ? `<Relationship Id="rId${hasAuxiliary ? 4 : 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` : ''}
</Relationships>`),
		...(hasTable
			? {
					'xl/worksheets/_rels/sheet1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`),
				}
			: {}),
		...(hasFeatureRich
			? {
					'xl/worksheets/_rels/sheet1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/ascend" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`),
				}
			: {}),
		'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    ${hasAuxiliary ? '<sheet name="Summary" sheetId="2" r:id="rId2"/>' : ''}
    ${hasAuxiliary ? '<sheet name="Archive" sheetId="3" r:id="rId3"/>' : ''}
  </sheets>
  ${hasFeatureRich ? `<definedNames><definedName name="FeatureRange">Data!$A$1:$${indexToColumn(Math.max(0, cols - 1))}$${Math.max(1, rows)}</definedName></definedNames>` : ''}
</workbook>`),
		'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${indexToColumn(Math.max(0, cols - 1))}${Math.max(1, rows)}"/>
  <sheetData>${sheetRows.join('')}</sheetData>
  ${hasFeatureRich ? `<conditionalFormatting sqref="A1:A${Math.max(1, rows)}"><cfRule type="cellIs" operator="greaterThan" priority="1"><formula>0</formula></cfRule></conditionalFormatting>` : ''}
  ${hasFeatureRich ? `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" sqref="C2:C${Math.max(2, rows)}"><formula1>"Q1,Q2,Q3"</formula1></dataValidation></dataValidations>` : ''}
  ${hasFeatureRich ? '<hyperlinks><hyperlink ref="A1" r:id="rId1" display="Ascend" tooltip="Open Ascend"/></hyperlinks>' : ''}
  ${hasFeatureRich ? '<legacyDrawing r:id="rId3"/>' : ''}
  ${hasTable ? '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' : ''}
</worksheet>`),
		...(hasAuxiliary
			? {
					'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A${Math.max(1, Math.min(rows, 100))}"/>
  <sheetData>${auxiliarySheetRows('summary', rows)}</sheetData>
</worksheet>`),
					'xl/worksheets/sheet3.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A${Math.max(1, Math.min(rows, 100))}"/>
  <sheetData>${auxiliarySheetRows('archive', rows)}</sheetData>
</worksheet>`),
				}
			: {}),
		...(hasStyles ? { 'xl/styles.xml': strToU8(rawStylesXml()) } : {}),
		...(hasTable
			? {
					'xl/tables/table1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="DataTable" displayName="DataTable" ref="${tableRef}" headerRowCount="1">
  <autoFilter ref="${tableRef}"/>
  <tableColumns count="${cols}">
    ${Array.from({ length: cols }, (_, index) => `<tableColumn id="${index + 1}" name="Column ${index + 1}"/>`).join('')}
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>
</table>`),
				}
			: {}),
		...(hasFeatureRich
			? {
					'xl/comments1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Ascend</author></authors>
  <commentList><comment ref="B2" authorId="0"><text><t>Review</t></text></comment></commentList>
</comments>`),
					'xl/drawings/vmlDrawing1.vml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  <v:shape id="_x0000_s1024" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/>
    <v:shadow on="t" color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Row>1</x:Row>
      <x:Column>1</x:Column>
    </x:ClientData>
  </v:shape>
</xml>`),
				}
			: {}),
	})
}

function auxiliarySheetRows(kind: 'summary' | 'archive', rows: number): string {
	return Array.from({ length: Math.min(rows, 100) }, (_, row) => {
		const ref = `A${row + 1}`
		const cell =
			kind === 'summary'
				? `<c r="${ref}"><v>${row}</v></c>`
				: `<c r="${ref}" t="inlineStr"><is><t>archive-${row}</t></is></c>`
		return `<row r="${row + 1}">${cell}</row>`
	}).join('')
}

function rawStylesXml(): string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
  <fonts count="3"><font/><font><b/></font><font><color rgb="FF1F4E79"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/></patternFill></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  </cellXfs>
</styleSheet>`
}

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
}

export async function buildWorkloadDataSet(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
	readSource: ReadSource = 'ascend-writer',
): Promise<CompetitiveDataSet> {
	const values = buildWorkloadValues(workloadName, rows, cols)
	const semanticLines = semanticLinesForValues(values)
	const bytes =
		readSource === 'raw-ooxml'
			? buildRawOoxmlXlsx(workloadName, {
					rows,
					cols,
					valueAt: (row, col) => values[row]?.[col],
				})
			: (() => {
					const workbook = createWorkbook()
					setCoreCell(workbook, values, rows, cols, workloadName)
					const result = writeXlsx(workbook, undefined, { omitDenseCellRefs: false })
					if (!result.ok) throw new Error(result.error.message)
					return result.value
				})()
	const xlsxPath = join(
		tmpdir(),
		`ascend-competitive-io-${readSource}-${workloadName}-${Date.now()}-${rows}x${cols}.xlsx`,
	)
	await writeFile(xlsxPath, bytes)
	return {
		workloadName,
		readSource,
		rows,
		cols,
		cells: semanticLines.length,
		values,
		semanticCellValuesHash: hashLines(semanticLines),
		orderedSemanticCellValuesHash: hashOrderedLines(semanticLines),
		xlsxPath,
		xlsxBytes: bytes,
	}
}

export async function buildRawReadWorkloadDataSet(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): Promise<CompetitiveDataSet> {
	const bytes = buildRawOoxmlXlsx(workloadName, {
		rows,
		cols,
		valueAt: (row, col) => workloadValue(workloadName, row, col, cols),
	})
	const xlsxPath = join(
		tmpdir(),
		`ascend-competitive-io-raw-ooxml-${workloadName}-${Date.now()}-${rows}x${cols}.xlsx`,
	)
	await writeFile(xlsxPath, bytes)
	return {
		workloadName,
		readSource: 'raw-ooxml',
		rows,
		cols,
		cells: generatedWorkloadCellCount(workloadName, rows, cols),
		values: [],
		semanticCellValuesHash: expectedWorkloadValuesHash(workloadName, rows, cols),
		orderedSemanticCellValuesHash: expectedOrderedWorkloadValuesHash(workloadName, rows, cols),
		xlsxPath,
		xlsxBytes: bytes,
	}
}

export function buildGeneratedWriteDataSet(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): CompetitiveDataSet {
	const semanticCellValuesHash = expectedWorkloadValuesHash(workloadName, rows, cols)
	const orderedSemanticCellValuesHash = expectedOrderedWorkloadValuesHash(workloadName, rows, cols)
	return {
		workloadName,
		readSource: 'ascend-writer',
		sourceMode: 'generated-write',
		rows,
		cols,
		cells: generatedWorkloadCellCount(workloadName, rows, cols),
		values: [],
		semanticCellValuesHash,
		orderedSemanticCellValuesHash,
		xlsxPath: '',
		xlsxBytes: new Uint8Array(),
	}
}

function generatedWorkloadCellCount(
	workloadName: WorkloadName,
	rows: number,
	cols: number,
): number {
	if (workloadName !== 'sparse-wide') return rows * cols
	let count = 0
	for (let row = 0; row < rows; row++) {
		for (let col = 0; col < cols; col++) {
			if (workloadValue(workloadName, row, col, cols) !== null) count++
		}
	}
	return count
}

export async function buildDenseDataSet(rows: number, cols: number): Promise<CompetitiveDataSet> {
	return buildWorkloadDataSet('dense-values', rows, cols)
}

function featureRange(input: CompetitiveDataSet): string {
	return `Data!$A$1:$${indexToColumn(Math.max(0, input.cols - 1))}$${Math.max(1, input.rows)}`
}

function applySheetJsFeatureRich(
	sheetJs: typeof import('xlsx'),
	worksheet: import('xlsx').WorkSheet,
	workbook: import('xlsx').WorkBook,
	input: CompetitiveDataSet,
): void {
	const cell = (worksheet.A1 ?? { t: 's', v: 'Ascend' }) as import('xlsx').CellObject
	cell.t = 's'
	cell.v = 'Ascend'
	sheetJs.utils.cell_set_hyperlink(cell, 'https://example.com/ascend', 'Open Ascend')
	sheetJs.utils.cell_add_comment(cell, 'Review', 'Ascend')
	worksheet.A1 = cell
	workbook.Workbook ??= {}
	workbook.Workbook.Names = [
		...(workbook.Workbook.Names ?? []),
		{ Name: 'FeatureRange', Ref: featureRange(input) },
	]
}

function applyExcelJsFeatureRich(
	workbook: import('exceljs').Workbook,
	sheet: import('exceljs').Worksheet,
	input: CompetitiveDataSet,
): void {
	sheet.getCell('A1').value = {
		text: 'Ascend',
		hyperlink: 'https://example.com/ascend',
		tooltip: 'Open Ascend',
	}
	if (input.rows > 1 && input.cols > 1) {
		sheet.getCell('B2').note = 'Review'
	}
	if (input.rows > 1 && input.cols > 2) {
		const dataValidations = (
			sheet as unknown as {
				readonly dataValidations?: { add(address: string, validation: unknown): unknown }
			}
		).dataValidations
		dataValidations?.add(`C2:C${input.rows}`, {
			type: 'list',
			allowBlank: true,
			showInputMessage: true,
			formulae: ['"Q1,Q2,Q3"'],
		})
	}
	sheet.addConditionalFormatting({
		ref: `A1:A${Math.max(1, input.rows)}`,
		rules: [
			{
				type: 'cellIs',
				operator: 'greaterThan',
				priority: 1,
				formulae: [0],
				style: { font: { bold: true } },
			},
		],
	})
	workbook.definedNames.add(featureRange(input), 'FeatureRange')
}

async function runTimedOperation<T>(
	input: CompetitiveDataSet,
	repeat: number,
	warmup: number,
	operation: () => T | Promise<T>,
): Promise<{ samples: readonly MetricSample[]; last: T }> {
	let last: T | undefined
	for (let i = 0; i < warmup; i++) {
		last = await operation()
	}
	const samples: MetricSample[] = []
	for (let i = 0; i < repeat; i++) {
		runGc()
		const rssBefore = getRssBytes()
		const heapBefore = process.memoryUsage().heapUsed
		const start = performance.now()
		last = await operation()
		const durationMs = performance.now() - start
		const memAfter = process.memoryUsage()
		const rssAfter = getRssBytes()
		runGc()
		const rssAfterGc = getRssBytes()
		const heapAfterGc = process.memoryUsage().heapUsed
		samples.push({
			durationMs,
			throughputPerSec:
				durationMs > 0 ? (input.cells / durationMs) * 1000 : Number.POSITIVE_INFINITY,
			rssDeltaBytes:
				rssBefore !== undefined && rssAfter !== undefined
					? Math.max(0, rssAfter - rssBefore)
					: undefined,
			retainedRssDeltaBytes:
				rssBefore !== undefined && rssAfterGc !== undefined
					? Math.max(0, rssAfterGc - rssBefore)
					: undefined,
			peakRssBytes: observedPeakRssBytes([rssBefore, rssAfter, rssAfterGc]),
			heapDeltaBytes: Math.max(0, memAfter.heapUsed - heapBefore),
			heapUsedBytes: memAfter.heapUsed,
			heapTotalBytes: memAfter.heapTotal,
			heapAfterGcBytes: heapAfterGc,
		})
	}
	if (last === undefined) {
		throw new Error('Timed operation did not produce a result')
	}
	return { samples, last }
}

function writeAscendGeneratedBytes(input: CompetitiveDataSet): Uint8Array {
	const workbook = createWorkbook()
	setCoreCellGenerated(workbook, input.rows, input.cols, input.workloadName)
	const result = writeXlsx(workbook, undefined, {
		useSharedStrings: input.workloadName === 'feature-rich' ? undefined : false,
		usePlainStrings: input.workloadName === 'string-heavy' || input.workloadName === 'plain-text',
		omitDenseCellRefs: input.workloadName === 'string-heavy' || input.workloadName === 'plain-text',
	})
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function writeSheetJsGeneratedBytes(
	sheetJs: typeof import('xlsx'),
	input: CompetitiveDataSet,
): Uint8Array {
	const worksheet = sheetJs.utils.aoa_to_sheet(input.values)
	const workbook = sheetJs.utils.book_new()
	if (input.workloadName === 'feature-rich') {
		applySheetJsFeatureRich(sheetJs, worksheet, workbook, input)
	}
	sheetJs.utils.book_append_sheet(workbook, worksheet, 'Data')
	return sheetJs.write(workbook, {
		type: 'buffer',
		bookType: 'xlsx',
	}) as Uint8Array
}

async function writeExcelJsGeneratedBytes(
	ExcelJS: typeof import('exceljs'),
	input: CompetitiveDataSet,
): Promise<Uint8Array> {
	const workbook = new ExcelJS.Workbook()
	const sheet = workbook.addWorksheet('Data')
	for (let row = 0; row < input.rows; row++) {
		sheet.addRow(input.values[row] ?? [])
	}
	if (input.workloadName === 'feature-rich') {
		applyExcelJsFeatureRich(workbook, sheet, input)
	}
	const bytes = (await workbook.xlsx.writeBuffer()) as ArrayBuffer
	return new Uint8Array(bytes)
}

async function loadCases(workloadName: WorkloadName): Promise<{
	cases: CompetitiveCase[]
	skipped: Array<{ library: string; reason: string }>
	externalReadRunnerSpecs: readonly ExternalReadRunnerSpec[]
	externalWriteRunnerSpecs: readonly ExternalWriteRunnerSpec[]
}> {
	const suffix = workloadName === 'dense-values' ? 'dense' : workloadName
	const readOperationProfile =
		workloadName === 'selected-sheet'
			? 'read-selected-values'
			: workloadName === 'metadata-only'
				? 'read-metadata-only'
				: workloadName === 'warm-workflow'
					? 'read-values-warm'
					: workloadName === 'feature-rich'
						? 'read-values-rich-metadata'
						: undefined
	const includeWriteCases =
		workloadName !== 'selected-sheet' &&
		workloadName !== 'metadata-only' &&
		workloadName !== 'warm-workflow'
	const cases: CompetitiveCase[] = [
		{
			name: `ascend:xlsx-read-${suffix}`,
			library: 'ascend',
			category: 'read',
			...(readOperationProfile ? { operationProfile: readOperationProfile } : {}),
			async run(input) {
				if (workloadName === 'selected-sheet') {
					const workbook = await Ascend.open(input.xlsxBytes, {
						mode: 'values',
						sheets: ['Data'],
					})
					return { assertions: ascendSelectedSheetAssertions(workbook, input) }
				}
				if (workloadName === 'metadata-only') {
					const workbook = await Ascend.open(input.xlsxBytes, { mode: 'metadata-only' })
					return { assertions: ascendMetadataOnlyAssertions(workbook) }
				}
				const workbook = await Ascend.open(input.xlsxBytes, {
					mode: 'values',
					...(workloadName === 'feature-rich' ? { richMetadata: true } : {}),
				})
				return { assertions: denseWorkbookAssertions(workbook.getWorkbookModel(), input) }
			},
			async runBatched(input, repeat, warmup) {
				const timed = await runTimedOperation(input, repeat, warmup, async () => {
					if (workloadName === 'selected-sheet') {
						return Ascend.open(input.xlsxBytes, {
							mode: 'values',
							sheets: ['Data'],
						})
					}
					if (workloadName === 'metadata-only') {
						return Ascend.open(input.xlsxBytes, { mode: 'metadata-only' })
					}
					return Ascend.open(input.xlsxBytes, {
						mode: 'values',
						...(workloadName === 'feature-rich' ? { richMetadata: true } : {}),
					})
				})
				return {
					samples: timed.samples,
					assertions:
						workloadName === 'selected-sheet'
							? ascendSelectedSheetAssertions(timed.last, input)
							: workloadName === 'metadata-only'
								? ascendMetadataOnlyAssertions(timed.last)
								: denseWorkbookAssertions(timed.last.getWorkbookModel(), input),
				}
			},
		},
		...(includeWriteCases
			? [
					{
						name: `ascend:xlsx-write-${suffix}`,
						library: 'ascend',
						category: 'write' as const,
						capabilities: {
							writeFormulas: true,
							writeTables: true,
							writeRichMetadata: true,
						},
						async run(input: CompetitiveDataSet) {
							const bytes = writeAscendGeneratedBytes(input)
							return { assertions: denseWriteAssertions(bytes, input) }
						},
						async runBatched(input: CompetitiveDataSet, repeat: number, warmup: number) {
							const timed = await runTimedOperation(input, repeat, warmup, () =>
								writeAscendGeneratedBytes(input),
							)
							return {
								samples: timed.samples,
								assertions: denseWriteAssertions(timed.last, input),
							}
						},
					},
				]
			: []),
	]
	const skipped: Array<{ library: string; reason: string }> = []
	const allExternalWriteRunnerSpecs = await loadExternalWriteRunnerSpecs()
	const workloadExternalWriteRunnerSpecs = allExternalWriteRunnerSpecs.filter((spec) =>
		supportsGeneratedRunnerWorkload(spec, workloadName),
	)
	const externalWriteRunnerSpecs = includeWriteCases
		? workloadExternalWriteRunnerSpecs.filter((spec) =>
				supportsGeneratedWriteWorkload(skipped, spec, workloadName),
			)
		: workloadExternalWriteRunnerSpecs
	const allExternalWriteRunnerNames = new Set(allExternalWriteRunnerSpecs.map((spec) => spec.name))

	let sheetJs: typeof import('xlsx') | undefined
	try {
		sheetJs = await import('xlsx')
	} catch (error) {
		skipped.push({
			library: 'sheetjs',
			reason: error instanceof Error ? error.message : 'module not available',
		})
	}
	if (sheetJs) {
		cases.push({
			name: `sheetjs:xlsx-read-${suffix}`,
			library: 'sheetjs',
			category: 'read',
			...(readOperationProfile ? { operationProfile: readOperationProfile } : {}),
			async run(input) {
				if (workloadName === 'selected-sheet') {
					const workbook = sheetJs.read(input.xlsxBytes, { type: 'buffer', sheets: 'Data' })
					return { assertions: sheetJsSelectedSheetAssertions(sheetJs, workbook, input) }
				}
				if (workloadName === 'metadata-only') {
					const workbook = sheetJs.read(input.xlsxBytes, { type: 'buffer', bookSheets: true })
					return { assertions: sheetJsMetadataOnlyAssertions(workbook, input) }
				}
				const workbook = sheetJs.read(input.xlsxBytes, {
					type: 'buffer',
					...(workloadName === 'feature-rich' ? { bookFiles: true } : {}),
				})
				return { assertions: workloadSheetJsAssertions(sheetJs, workbook, input) }
			},
			async runBatched(input, repeat, warmup) {
				const timed = await runTimedOperation(input, repeat, warmup, () => {
					if (workloadName === 'selected-sheet') {
						return sheetJs.read(input.xlsxBytes, { type: 'buffer', sheets: 'Data' })
					}
					return workloadName === 'metadata-only'
						? sheetJs.read(input.xlsxBytes, { type: 'buffer', bookSheets: true })
						: sheetJs.read(input.xlsxBytes, {
								type: 'buffer',
								...(workloadName === 'feature-rich' ? { bookFiles: true } : {}),
							})
				})
				return {
					samples: timed.samples,
					assertions:
						workloadName === 'selected-sheet'
							? sheetJsSelectedSheetAssertions(sheetJs, timed.last, input)
							: workloadName === 'metadata-only'
								? sheetJsMetadataOnlyAssertions(timed.last, input)
								: workloadSheetJsAssertions(sheetJs, timed.last, input),
				}
			},
		})
		if (
			includeWriteCases &&
			supportsGeneratedWriteWorkload(skipped, { name: 'sheetjs', capabilities: {} }, workloadName)
		) {
			cases.push({
				name: `sheetjs:xlsx-write-${suffix}`,
				library: 'sheetjs',
				category: 'write' as const,
				async run(input: CompetitiveDataSet) {
					const bytes = writeSheetJsGeneratedBytes(sheetJs, input)
					return { assertions: denseWriteAssertions(bytes, input) }
				},
				async runBatched(input: CompetitiveDataSet, repeat: number, warmup: number) {
					const timed = await runTimedOperation(input, repeat, warmup, () =>
						writeSheetJsGeneratedBytes(sheetJs, input),
					)
					return {
						samples: timed.samples,
						assertions: denseWriteAssertions(timed.last, input),
					}
				},
			})
		}
	}

	let ExcelJS: typeof import('exceljs') | undefined
	try {
		ExcelJS = await import('exceljs')
	} catch (error) {
		skipped.push({
			library: 'exceljs',
			reason: error instanceof Error ? error.message : 'module not available',
		})
	}
	if (ExcelJS) {
		if (workloadName === 'selected-sheet' || workloadName === 'metadata-only') {
			skipped.push({
				library: 'exceljs',
				reason:
					workloadName === 'selected-sheet'
						? 'selected-sheet read is unsupported by this harness without full workbook hydration'
						: 'metadata-only read is unsupported by this harness',
			})
		} else {
			cases.push({
				name: `exceljs:xlsx-read-${suffix}`,
				library: 'exceljs',
				category: 'read',
				...(readOperationProfile ? { operationProfile: readOperationProfile } : {}),
				async run(input) {
					const workbook = new ExcelJS.Workbook()
					await workbook.xlsx.load(Buffer.from(input.xlsxBytes))
					return { assertions: workloadExcelJsAssertions(workbook, input) }
				},
				async runBatched(input, repeat, warmup) {
					const timed = await runTimedOperation(input, repeat, warmup, async () => {
						const workbook = new ExcelJS.Workbook()
						await workbook.xlsx.load(Buffer.from(input.xlsxBytes))
						return workbook
					})
					return {
						samples: timed.samples,
						assertions: workloadExcelJsAssertions(timed.last, input),
					}
				},
			})
		}
		if (
			includeWriteCases &&
			supportsGeneratedWriteWorkload(
				skipped,
				{ name: 'exceljs', capabilities: { writeRichMetadata: true } },
				workloadName,
			)
		) {
			cases.push({
				name: `exceljs:xlsx-write-${suffix}`,
				library: 'exceljs',
				category: 'write',
				async run(input) {
					const bytes = await writeExcelJsGeneratedBytes(ExcelJS, input)
					return { assertions: denseWriteAssertions(bytes, input) }
				},
				async runBatched(input, repeat, warmup) {
					const timed = await runTimedOperation(input, repeat, warmup, () =>
						writeExcelJsGeneratedBytes(ExcelJS, input),
					)
					return {
						samples: timed.samples,
						assertions: denseWriteAssertions(timed.last, input),
					}
				},
			})
		}
	}

	if (includeWriteCases) {
		const manifestProvidesXlsxWriter =
			allExternalWriteRunnerNames.has('xlsxwriter') ||
			allExternalWriteRunnerNames.has('xlsxwriter-constant-memory')
		if (!manifestProvidesXlsxWriter) {
			if (await commandAvailable(['python3', '-c', 'import xlsxwriter, openpyxl'])) {
				if (
					supportsGeneratedWriteWorkload(
						skipped,
						{
							name: 'xlsxwriter',
							capabilities: {
								writeFormulas: true,
								writeTables: true,
								writeRichMetadata: true,
							},
						},
						workloadName,
					)
				) {
					cases.push({
						name: `xlsxwriter:xlsx-write-${suffix}`,
						library: 'xlsxwriter',
						category: 'write',
						executionScope: 'external-process',
						async run(input) {
							return {
								assertions: await runXlsxWriter(input, 1, 0).then((result) => result.assertions),
							}
						},
						async runBatched(input, repeat, warmup) {
							return runXlsxWriter(input, repeat, warmup)
						},
					})
				}
				if (
					supportsGeneratedWriteWorkload(
						skipped,
						{
							name: 'xlsxwriter-constant-memory',
							capabilities: { writeFormulas: true },
						},
						workloadName,
					)
				) {
					cases.push({
						name: `xlsxwriter-constant-memory:xlsx-write-${suffix}`,
						library: 'xlsxwriter-constant-memory',
						category: 'write',
						executionScope: 'external-process',
						async run(input) {
							return {
								assertions: await runXlsxWriter(input, 1, 0, true).then(
									(result) => result.assertions,
								),
							}
						},
						async runBatched(input, repeat, warmup) {
							return runXlsxWriter(input, repeat, warmup, true)
						},
					})
				}
			} else {
				skipped.push({
					library: 'xlsxwriter',
					reason: 'python3 modules xlsxwriter and openpyxl are not available',
				})
			}
		}
	}

	const openpyxlWriterRunner = 'fixtures/benchmarks/runners/openpyxl_writer_runner.py'
	const manifestProvidesOpenPyxl =
		allExternalWriteRunnerNames.has('openpyxl') ||
		allExternalWriteRunnerNames.has('openpyxl-write-only')
	if (!manifestProvidesOpenPyxl) {
		if (!(await fileExists(openpyxlWriterRunner))) {
			skipped.push({
				library: 'openpyxl',
				reason: `${openpyxlWriterRunner} is not available`,
			})
		} else if (
			includeWriteCases &&
			(await commandAvailable(['python3', '-c', 'import openpyxl']))
		) {
			if (
				supportsGeneratedWriteWorkload(
					skipped,
					{
						name: 'openpyxl',
						capabilities: { writeTables: true, writeRichMetadata: true },
					},
					workloadName,
				)
			) {
				cases.push({
					name: `openpyxl:xlsx-write-${suffix}`,
					library: 'openpyxl',
					category: 'write',
					executionScope: 'external-process',
					async run(input) {
						return {
							assertions: await runPythonWriteRunner(
								'OpenPyXL',
								openpyxlWriterRunner,
								input,
								1,
								0,
							).then((result) => result.assertions),
						}
					},
					async runBatched(input, repeat, warmup) {
						return runPythonWriteRunner('OpenPyXL', openpyxlWriterRunner, input, repeat, warmup)
					},
				})
			}
			if (
				supportsGeneratedWriteWorkload(
					skipped,
					{ name: 'openpyxl-write-only', capabilities: {} },
					workloadName,
				)
			) {
				cases.push({
					name: `openpyxl-write-only:xlsx-write-${suffix}`,
					library: 'openpyxl-write-only',
					category: 'write',
					executionScope: 'external-process',
					async run(input) {
						return {
							assertions: await runPythonWriteRunner(
								'OpenPyXL write-only',
								openpyxlWriterRunner,
								input,
								1,
								0,
								['--write-only'],
							).then((result) => result.assertions),
						}
					},
					async runBatched(input, repeat, warmup) {
						return runPythonWriteRunner(
							'OpenPyXL write-only',
							openpyxlWriterRunner,
							input,
							repeat,
							warmup,
							['--write-only'],
						)
					},
				})
			}
		} else if (includeWriteCases) {
			skipped.push({
				library: 'openpyxl',
				reason: 'python3 module openpyxl is not available',
			})
		}
	}

	const pyexcelerateRunner = 'fixtures/benchmarks/runners/pyexcelerate_runner.py'
	const manifestProvidesPyexcelerate =
		allExternalWriteRunnerNames.has('pyexcelerate') ||
		allExternalWriteRunnerNames.has('pyexcelerate-range') ||
		allExternalWriteRunnerNames.has('pyexcelerate-cell')
	if (!manifestProvidesPyexcelerate) {
		if (!(await fileExists(pyexcelerateRunner))) {
			skipped.push({
				library: 'pyexcelerate',
				reason: `${pyexcelerateRunner} is not available`,
			})
		} else if (
			includeWriteCases &&
			(await commandAvailable(['python3', '-c', 'import pyexcelerate, openpyxl']))
		) {
			const strategies = [
				{ library: 'pyexcelerate', label: 'PyExcelerate bulk-sheet', args: [] },
				{
					library: 'pyexcelerate-range',
					label: 'PyExcelerate range',
					args: ['--strategy', 'range'],
				},
				{
					library: 'pyexcelerate-cell',
					label: 'PyExcelerate cell',
					args: ['--strategy', 'cell'],
				},
			] as const
			for (const strategy of strategies) {
				if (
					!supportsGeneratedWriteWorkload(
						skipped,
						{ name: strategy.library, capabilities: {} },
						workloadName,
					)
				) {
					continue
				}
				cases.push({
					name: `${strategy.library}:xlsx-write-${suffix}`,
					library: strategy.library,
					category: 'write',
					executionScope: 'external-process',
					async run(input) {
						return {
							assertions: await runPythonWriteRunner(
								strategy.label,
								pyexcelerateRunner,
								input,
								1,
								0,
								strategy.args,
							).then((result) => result.assertions),
						}
					},
					async runBatched(input, repeat, warmup) {
						return runPythonWriteRunner(
							strategy.label,
							pyexcelerateRunner,
							input,
							repeat,
							warmup,
							strategy.args,
						)
					},
				})
			}
		} else if (includeWriteCases) {
			skipped.push({
				library: 'pyexcelerate',
				reason: 'python3 modules pyexcelerate and openpyxl are not available',
			})
		}
	}

	const externalReadRunnerSpecs = (await loadExternalReadRunnerSpecs()).filter((spec) =>
		supportsGeneratedRunnerWorkload(spec, workloadName),
	)
	for (const spec of externalReadRunnerSpecs) {
		if (spec.capabilities?.metadataOnlyRead === true && workloadName !== 'metadata-only') {
			continue
		}
		if (workloadName === 'metadata-only' && spec.capabilities?.metadataOnlyRead !== true) {
			continue
		}
		if (workloadName === 'selected-sheet' && spec.capabilities?.selectedSheetRead !== true) {
			continue
		}
		cases.push({
			name: `${spec.name}:xlsx-read-${suffix}`,
			library: spec.name,
			category: 'read',
			executionScope: 'external-process',
			runnerProvenance: {
				...(spec.adapterVersion ? { adapterVersion: spec.adapterVersion } : {}),
				...(spec.libraryVersion ? { libraryVersion: spec.libraryVersion } : {}),
				...(spec.runtime ? { runtime: spec.runtime } : {}),
			},
			...(spec.timingModel ? { timingModel: spec.timingModel } : {}),
			...(spec.validationModel ? { validationModel: spec.validationModel } : {}),
			...(spec.memoryModel ? { memoryModel: spec.memoryModel } : {}),
			...(readOperationProfile ? { operationProfile: readOperationProfile } : {}),
			...(spec.capabilities ? { capabilities: spec.capabilities } : {}),
			async run(input) {
				return {
					assertions: await runExternalReadRunner(spec, input, 1, 0).then(
						(result) => result.assertions,
					),
				}
			},
			async runBatched(input, repeat, warmup) {
				return runExternalReadRunner(spec, input, repeat, warmup)
			},
		})
	}

	if (includeWriteCases) {
		for (const spec of externalWriteRunnerSpecs) {
			cases.push({
				name: `${spec.name}:xlsx-write-${suffix}`,
				library: spec.name,
				category: 'write',
				executionScope: 'external-process',
				runnerProvenance: {
					...(spec.adapterVersion ? { adapterVersion: spec.adapterVersion } : {}),
					...(spec.libraryVersion ? { libraryVersion: spec.libraryVersion } : {}),
					...(spec.runtime ? { runtime: spec.runtime } : {}),
				},
				...(spec.timingModel ? { timingModel: spec.timingModel } : {}),
				...(spec.validationModel ? { validationModel: spec.validationModel } : {}),
				...(spec.memoryModel ? { memoryModel: spec.memoryModel } : {}),
				...(spec.capabilities ? { capabilities: spec.capabilities } : {}),
				async run(input) {
					return {
						assertions: await runExternalWriteRunner(spec, input, 1, 0).then(
							(result) => result.assertions,
						),
					}
				},
				async runBatched(input, repeat, warmup) {
					return runExternalWriteRunner(spec, input, repeat, warmup)
				},
			})
		}
	}

	return { cases, skipped, externalReadRunnerSpecs, externalWriteRunnerSpecs }
}

async function fileExists(path: string): Promise<boolean> {
	try {
		await access(path)
		return true
	} catch {
		return false
	}
}

async function commandAvailable(command: readonly string[]): Promise<boolean> {
	const proc = Bun.spawn(resolveExternalRunnerCommand(command), {
		stdout: 'pipe',
		stderr: 'pipe',
		cwd: process.cwd(),
	})
	const exitCode = await proc.exited
	return exitCode === 0
}

async function runXlsxWriter(
	input: CompetitiveDataSet,
	repeat: number,
	warmup: number,
	constantMemory = false,
): Promise<{
	assertions?: Record<string, PrimitiveAssertion>
	assertionsBySample?: readonly Record<string, PrimitiveAssertion>[]
	samples?: readonly MetricSample[]
}> {
	return runPythonWriteRunner(
		'XlsxWriter',
		'fixtures/benchmarks/runners/xlsxwriter_runner.py',
		input,
		repeat,
		warmup,
		constantMemory ? ['--constant-memory'] : [],
	)
}

async function runExternalReadRunner(
	spec: ExternalReadRunnerSpec,
	input: CompetitiveDataSet,
	repeat: number,
	warmup: number,
): Promise<{
	assertions?: Record<string, PrimitiveAssertion>
	samples?: readonly MetricSample[]
}> {
	const validationMode = readValidationModeFlag()
	const validationArgs =
		spec.capabilities?.finalValidation === true ? ['--validation-mode', validationMode] : []
	const selectedSheetArgs =
		input.workloadName === 'selected-sheet' && spec.capabilities?.selectedSheetRead === true
			? ['--selected-sheet', 'Data']
			: []
	const proc = Bun.spawn(
		[
			...resolveExternalRunnerCommand(spec.command),
			'--operation',
			'read',
			'--file',
			input.xlsxPath,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			...validationArgs,
			...selectedSheetArgs,
			'--json',
		],
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `${spec.name} runner exited with code ${exitCode}`)
	}
	const parsed = JSON.parse(stdout) as unknown
	const assertions = normalizeAssertions(parsed)
	return {
		assertions:
			spec.capabilities?.finalValidation === true
				? annotateFinalValidationAssertions(assertions, validationMode, repeat)
				: assertions,
		assertionsBySample: normalizeExternalSampleAssertions(parsed, repeat, spec.name),
		samples: normalizeExternalSamples(parsed, repeat, spec.name),
	}
}

async function runExternalWriteRunner(
	spec: ExternalWriteRunnerSpec,
	input: CompetitiveDataSet,
	repeat: number,
	warmup: number,
): Promise<{
	assertions?: Record<string, PrimitiveAssertion>
	assertionsBySample?: readonly Record<string, PrimitiveAssertion>[]
	samples?: readonly MetricSample[]
}> {
	const validationMode = readValidationModeFlag()
	const validationArgs =
		spec.capabilities?.finalValidation === true ? ['--validation-mode', validationMode] : []
	const proc = Bun.spawn(
		[
			...resolveExternalRunnerCommand(spec.command),
			'--operation',
			'write',
			'--rows',
			String(input.rows),
			'--cols',
			String(input.cols),
			'--workload',
			input.workloadName,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			...validationArgs,
			'--json',
		],
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `${spec.name} writer exited with code ${exitCode}`)
	}
	const parsed = JSON.parse(stdout) as unknown
	const assertions = normalizeAssertions(parsed)
	return {
		assertions:
			spec.capabilities?.finalValidation === true
				? annotateFinalValidationAssertions(assertions, validationMode, repeat)
				: assertions,
		assertionsBySample: normalizeExternalSampleAssertions(parsed, repeat, spec.name),
		samples: normalizeExternalSamples(parsed, repeat, spec.name),
	}
}

async function runPythonWriteRunner(
	runnerName: string,
	runnerPath: string,
	input: CompetitiveDataSet,
	repeat: number,
	warmup: number,
	extraArgs: readonly string[] = [],
): Promise<{
	assertions?: Record<string, PrimitiveAssertion>
	assertionsBySample?: readonly Record<string, PrimitiveAssertion>[]
	samples?: readonly MetricSample[]
}> {
	const validationMode = readValidationModeFlag()
	const proc = Bun.spawn(
		[
			...resolveExternalRunnerCommand(['python3']),
			runnerPath,
			'--operation',
			'write',
			'--rows',
			String(input.rows),
			'--cols',
			String(input.cols),
			'--workload',
			input.workloadName,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			'--validation-mode',
			validationMode,
			...extraArgs,
			'--json',
		],
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `${runnerName} runner exited with code ${exitCode}`)
	}
	const parsed = JSON.parse(stdout) as unknown
	return {
		assertions: normalizeAssertions(parsed),
		assertionsBySample: normalizeExternalSampleAssertions(parsed, repeat, runnerName),
		samples: normalizeExternalSamples(parsed, repeat, runnerName),
	}
}

export function normalizeAssertions(value: unknown): Record<string, PrimitiveAssertion> {
	const source =
		typeof value === 'object' &&
		value !== null &&
		'assertions' in value &&
		typeof value.assertions === 'object' &&
		value.assertions !== null
			? value.assertions
			: value
	if (typeof source !== 'object' || source === null || Array.isArray(source)) {
		throw new Error('External runner output must be a JSON object or { "assertions": object }')
	}
	const assertions: Record<string, PrimitiveAssertion> = {}
	for (const [key, entry] of Object.entries(source)) {
		if (
			typeof entry === 'string' ||
			typeof entry === 'number' ||
			typeof entry === 'boolean' ||
			entry === null
		) {
			assertions[key] = entry
		} else {
			throw new Error(`External runner assertion "${key}" must be a primitive value`)
		}
	}
	return assertions
}

export function normalizeExternalSamples(
	value: unknown,
	repeat: number,
	runnerName = 'External runner',
): readonly MetricSample[] {
	if (
		typeof value !== 'object' ||
		value === null ||
		!('samples' in value) ||
		!Array.isArray(value.samples)
	) {
		throw new Error(`${runnerName} did not report samples`)
	}
	const samples: MetricSample[] = []
	for (const [index, sample] of value.samples.entries()) {
		if (typeof sample !== 'object' || sample === null) {
			throw new Error(`${runnerName} sample ${index} must be an object`)
		}
		const durationMs = (sample as { durationMs?: unknown }).durationMs
		if (typeof durationMs !== 'number' || !Number.isFinite(durationMs) || durationMs <= 0) {
			throw new Error(`${runnerName} sample ${index} must provide a positive durationMs`)
		}
		samples.push({
			durationMs,
			...optionalSampleNumber(sample, 'throughputPerSec'),
			...optionalSampleNumber(sample, 'rssDeltaBytes'),
			...optionalSampleNumber(sample, 'retainedRssDeltaBytes'),
			...optionalSampleNumber(sample, 'rssAfterBytes'),
			...optionalSampleNumber(sample, 'rssAfterGcBytes'),
			...optionalSampleNumber(sample, 'peakRssBytes'),
			...optionalSampleNumber(sample, 'heapDeltaBytes'),
			...optionalSampleNumber(sample, 'heapUsedBytes'),
			...optionalSampleNumber(sample, 'heapTotalBytes'),
			...optionalSampleNumber(sample, 'heapAfterGcBytes'),
		})
	}
	if (samples.length !== repeat) {
		throw new Error(
			`${runnerName} reported ${samples.length} samples but repeat requested ${repeat}`,
		)
	}
	return samples
}

function optionalSampleNumber(
	sample: object,
	key: keyof Omit<MetricSample, 'durationMs'>,
): Partial<MetricSample> {
	const value = (sample as Record<string, unknown>)[key]
	if (value === undefined) return {}
	if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
		throw new Error(`External runner sample field "${key}" must be a non-negative finite number`)
	}
	return { [key]: value }
}

function annotateFinalValidationAssertions(
	assertions: Record<string, PrimitiveAssertion> | undefined,
	validationMode: string,
	repeat: number,
): Record<string, PrimitiveAssertion> | undefined {
	if (!assertions) return assertions
	return {
		...assertions,
		validationMode,
		validationSamples: validationMode === 'each' ? repeat : 1,
	}
}

async function runCompetitiveCase(
	benchmarkCase: CompetitiveCase,
	input: CompetitiveDataSet,
	repeat: number,
	warmup: number,
): Promise<BenchmarkCaseResult> {
	if (benchmarkCase.runBatched) {
		const result = await benchmarkCase.runBatched(input, repeat, warmup)
		const evaluatedBySample =
			result.assertionsBySample?.map((assertions) =>
				evaluateAssertions(benchmarkCase.category, input, assertions),
			) ?? []
		const evaluated =
			evaluatedBySample[0] ?? evaluateAssertions(benchmarkCase.category, input, result.assertions)
		const samples =
			result.samples?.map((sample) => ({
				...sample,
				throughputPerSec:
					sample.throughputPerSec ??
					(sample.durationMs > 0 ? (input.cells / sample.durationMs) * 1000 : undefined),
			})) ?? []
		if (samples.length === 0) throw new Error(`${benchmarkCase.name} did not report samples`)
		const correctnessStatuses =
			evaluatedBySample.length > 0 ? evaluatedBySample.map((entry) => entry.status) : []
		return buildResult(
			benchmarkCase,
			input,
			repeat,
			samples,
			evaluatedBySample.length > 0
				? coalesceRepeatCorrectnessStatus(correctnessStatuses)
				: evaluated.status,
			evaluated.assertions,
			correctnessStatuses,
		)
	}
	for (let i = 0; i < warmup; i++) {
		await benchmarkCase.run(input)
	}
	const samples: MetricSample[] = []
	let assertions: Record<string, PrimitiveAssertion> | undefined
	const correctnessStatuses: string[] = []
	for (let i = 0; i < repeat; i++) {
		runGc()
		const rssBefore = getRssBytes()
		const heapBefore = process.memoryUsage().heapUsed
		const start = performance.now()
		const result = await benchmarkCase.run(input)
		const durationMs = performance.now() - start
		const memAfter = process.memoryUsage()
		const rssAfter = getRssBytes()
		runGc()
		const rssAfterGc = getRssBytes()
		const heapAfterGc = process.memoryUsage().heapUsed
		samples.push({
			durationMs,
			throughputPerSec:
				durationMs > 0 ? (input.cells / durationMs) * 1000 : Number.POSITIVE_INFINITY,
			rssDeltaBytes:
				rssBefore !== undefined && rssAfter !== undefined
					? Math.max(0, rssAfter - rssBefore)
					: undefined,
			retainedRssDeltaBytes:
				rssBefore !== undefined && rssAfterGc !== undefined
					? Math.max(0, rssAfterGc - rssBefore)
					: undefined,
			peakRssBytes: observedPeakRssBytes([rssBefore, rssAfter, rssAfterGc]),
			heapDeltaBytes: Math.max(0, memAfter.heapUsed - heapBefore),
			heapUsedBytes: memAfter.heapUsed,
			heapTotalBytes: memAfter.heapTotal,
			heapAfterGcBytes: heapAfterGc,
		})
		const evaluated = evaluateAssertions(benchmarkCase.category, input, result.assertions)
		assertions ??= evaluated.assertions
		correctnessStatuses.push(evaluated.status)
	}
	return buildResult(
		benchmarkCase,
		input,
		repeat,
		samples,
		coalesceRepeatCorrectnessStatus(correctnessStatuses),
		assertions,
		correctnessStatuses,
	)
}

function buildResult(
	benchmarkCase: CompetitiveCase,
	input: CompetitiveDataSet,
	repeat: number,
	samples: readonly MetricSample[],
	correctnessStatus: string,
	assertions: Record<string, PrimitiveAssertion> | undefined,
	correctnessStatuses: readonly string[] = [],
): BenchmarkCaseResult {
	const assertedValidationSamples =
		typeof assertions?.validationSamples === 'number' ? assertions.validationSamples : undefined
	const validationSamples =
		assertedValidationSamples ??
		(correctnessStatuses.length > 0 ? correctnessStatuses.length : assertions ? 1 : 0)
	const validationMode =
		typeof assertions?.validationMode === 'string'
			? assertions.validationMode
			: validationSamples >= repeat
				? 'each'
				: 'final'
	const assertionBytes = typeof assertions?.bytes === 'number' ? assertions.bytes : undefined
	const measuredBytes =
		benchmarkCase.category === 'write' && assertionBytes !== undefined
			? assertionBytes
			: input.xlsxBytes.byteLength
	return {
		name: benchmarkCase.name,
		category: benchmarkCase.category,
		dimensions: {
			library: benchmarkCase.library,
			workload: input.workloadName,
			readSource: input.readSource,
			rows: input.rows,
			cols: input.cols,
			cells: input.cells,
			logicalCells: input.rows * input.cols,
			density: input.rows * input.cols > 0 ? input.cells / (input.rows * input.cols) : 0,
			bytes: measuredBytes,
			inputBytes: input.xlsxBytes.byteLength,
			...(benchmarkCase.category === 'write' && assertionBytes !== undefined
				? { outputBytes: assertionBytes }
				: {}),
			...(input.sourceMode ? { sourceMode: input.sourceMode } : {}),
			repeat,
			executionScope: benchmarkCase.executionScope ?? 'in-process',
			operationProfile: operationProfile(benchmarkCase, input),
			timingLane: timingLane(benchmarkCase, input),
			timingModel: timingModel(benchmarkCase),
			validationModel: validationModel(benchmarkCase),
			validationMode,
			validationSamples,
			...(benchmarkCase.memoryModel ? { memoryModel: benchmarkCase.memoryModel } : {}),
			...benchmarkProvenanceDimensions(assertions, benchmarkCase.runnerProvenance),
			correctnessStatus,
			rankingEligible: isRankingEligible(correctnessStatus),
			...(correctnessStatuses.length > 1
				? { repeatCorrectnessStatuses: correctnessStatuses.join(',') }
				: {}),
		},
		metrics: summarizeSamples(samples),
		...(repeat > 1 ? { samples } : {}),
		...(assertions ? { assertions } : {}),
	}
}

export function buildNonRankingResult(input: {
	readonly name: string
	readonly library: string
	readonly category: 'read' | 'write'
	readonly executionScope?: 'in-process' | 'external-process'
	readonly runnerProvenance?: {
		readonly adapterVersion?: string
		readonly libraryVersion?: string
		readonly runtime?: string
	}
	readonly timingModel?: string
	readonly validationModel?: string
	readonly memoryModel?: string
	readonly dataSet: CompetitiveDataSet
	readonly repeat: number
	readonly status: string
	readonly reason: string
}): BenchmarkCaseResult {
	return {
		name: input.name,
		category: input.category,
		dimensions: {
			library: input.library,
			workload: input.dataSet.workloadName,
			readSource: input.dataSet.readSource,
			rows: input.dataSet.rows,
			cols: input.dataSet.cols,
			cells: input.dataSet.cells,
			logicalCells: input.dataSet.rows * input.dataSet.cols,
			density:
				input.dataSet.rows * input.dataSet.cols > 0
					? input.dataSet.cells / (input.dataSet.rows * input.dataSet.cols)
					: 0,
			bytes: input.dataSet.xlsxBytes.byteLength,
			...(input.dataSet.sourceMode ? { sourceMode: input.dataSet.sourceMode } : {}),
			repeat: input.repeat,
			executionScope: input.executionScope ?? 'in-process',
			operationProfile: operationProfile(input, input.dataSet),
			timingLane: timingLane(input, input.dataSet),
			timingModel: timingModel(input),
			validationModel: validationModel(input),
			...(input.memoryModel ? { memoryModel: input.memoryModel } : {}),
			...benchmarkProvenanceDimensions(undefined, input.runnerProvenance),
			correctnessStatus: input.status,
			rankingEligible: false,
			errorReason: input.reason,
		},
		metrics: summarizeSamples([{ durationMs: 0 }]),
		assertions: { errorReason: input.reason },
	}
}

function renderSummary(
	results: readonly BenchmarkCaseResult[],
	skipped: readonly unknown[],
): string {
	const headers = [
		'case',
		'category',
		'median-ms',
		'p95-ms',
		'throughput',
		'rss-delta',
		'heap-delta',
	]
	const rows = results.map((result) => [
		result.name,
		result.category,
		result.metrics.medianMs.toFixed(2),
		result.metrics.p95Ms.toFixed(2),
		result.metrics.throughputPerSec !== undefined
			? formatRate(result.metrics.throughputPerSec)
			: 'n/a',
		result.metrics.rssDeltaBytes !== undefined ? formatBytes(result.metrics.rssDeltaBytes) : 'n/a',
		result.metrics.heapDeltaBytes !== undefined
			? formatBytes(result.metrics.heapDeltaBytes)
			: 'n/a',
	])
	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')
	const lines = [
		line(headers),
		widths.map((width) => '-'.repeat(width)).join('--'),
		...rows.map(line),
	]
	if (skipped.length > 0) {
		lines.push('')
		lines.push(`Skipped competitors: ${JSON.stringify(skipped)}`)
	}
	return lines.join('\n')
}

async function main(): Promise<void> {
	const workloadSelection = readWorkloadFlag()
	const readSource = readSourceFlag()
	const categorySelection = readCategoryFlag()
	const competitorSelection = readCompetitorFlag()
	const validationMode = readValidationModeFlag()
	const executionScopeSelection = readExecutionScopeFlag()
	const sourceMode = readSourceModeFlag()
	assertSourceModeCompatible({ sourceMode, categorySelection, executionScopeSelection })
	const libraryAllowlist = parseLibraryAllowlist(readFlag('--libraries'))
	const rowOverride =
		readFlag('--rows') === undefined ? undefined : readPositiveIntFlag('--rows', 1)
	const colOverride =
		readFlag('--cols') === undefined ? undefined : readPositiveIntFlag('--cols', 1)
	const repeat = readPositiveIntFlag('--repeat', 3)
	const warmup = readNonNegativeIntFlag('--warmup', 1)
	const json = hasFlag('--json')
	const workloads =
		workloadSelection === 'all' ? ALL_WORKLOADS : ([workloadSelection] as readonly WorkloadName[])
	const readSources =
		readSource === 'all' ? ALL_READ_SOURCES : ([readSource] as readonly ReadSource[])
	const results: BenchmarkCaseResult[] = []
	const failed: Array<{ case: string; library: string; category: string; reason: string }> = []
	const skipped: Array<{ library: string; reason: string }> = []
	const externalReadRunnerSpecsByWorkload: Record<string, readonly ExternalReadRunnerSpec[]> = {}
	const externalWriteRunnerSpecsByWorkload: Record<string, readonly ExternalWriteRunnerSpec[]> = {}
	const workloadDimensions: Array<{
		workload: WorkloadName
		readSource: ReadSource
		rows: number
		cols: number
	}> = []
	for (const workload of workloads) {
		const rows = workloadRows(workload, rowOverride)
		const cols = workloadCols(workload, colOverride)
		const loaded = await loadCases(workload)
		skipped.push(...loaded.skipped)
		externalReadRunnerSpecsByWorkload[workload] = loaded.externalReadRunnerSpecs
		externalWriteRunnerSpecsByWorkload[workload] = loaded.externalWriteRunnerSpecs
		for (const source of readSources) {
			workloadDimensions.push({ workload, readSource: source, rows, cols })
			const input =
				sourceMode === 'generated-write'
					? buildGeneratedWriteDataSet(workload, rows, cols)
					: await buildWorkloadDataSet(workload, rows, cols, source)
			for (const benchmarkCase of loaded.cases) {
				if (categorySelection !== 'all' && benchmarkCase.category !== categorySelection) {
					continue
				}
				if (!competitorMatches(benchmarkCase.library, competitorSelection)) {
					continue
				}
				if (!libraryAllowed(benchmarkCase.library, libraryAllowlist)) {
					continue
				}
				if (
					executionScopeSelection !== 'all' &&
					benchmarkCase.executionScope !== executionScopeSelection
				) {
					continue
				}
				if (
					readSource === 'all' &&
					source !== 'ascend-writer' &&
					benchmarkCase.category === 'write'
				) {
					continue
				}
				try {
					results.push(await runCompetitiveCase(benchmarkCase, input, repeat, warmup))
					if (!json) console.log(`completed ${benchmarkCase.name}`)
				} catch (error) {
					const reason = error instanceof Error ? error.message : String(error)
					failed.push({
						case: benchmarkCase.name,
						library: benchmarkCase.library,
						category: benchmarkCase.category,
						reason,
					})
					results.push(
						buildNonRankingResult({
							name: benchmarkCase.name,
							library: benchmarkCase.library,
							category: benchmarkCase.category,
							executionScope: benchmarkCase.executionScope,
							runnerProvenance: benchmarkCase.runnerProvenance,
							timingModel: benchmarkCase.timingModel,
							validationModel: benchmarkCase.validationModel,
							memoryModel: benchmarkCase.memoryModel,
							dataSet: input,
							repeat,
							status: 'error',
							reason,
						}),
					)
				}
			}
		}
	}
	const suite = createBenchmarkSuite({
		suite: 'ascend-competitive-io',
		kind: 'real-workbook',
		cases: results,
		metadata: {
			workload: workloadSelection,
			readSource,
			competitor: competitorSelection,
			validationMode,
			executionScope: executionScopeSelection,
			sourceMode,
			libraries: libraryAllowlist ? [...libraryAllowlist] : undefined,
			workloads: workloadDimensions,
			repeat,
			warmup,
			skipped,
			failed,
			externalReadRunnersByWorkload: externalReadRunnerSpecsByWorkload,
			externalWriteRunnersByWorkload: externalWriteRunnerSpecsByWorkload,
		},
	})
	if (json) {
		console.log(JSON.stringify(suite, null, 2))
		return
	}
	console.log('')
	console.log(renderSummary(results, [...skipped, ...failed]))
}

if (import.meta.main) {
	await main()
}
