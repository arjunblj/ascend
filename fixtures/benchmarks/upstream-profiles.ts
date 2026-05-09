/**
 * Upstream-shaped competitive profiles.
 *
 * These profiles encode dimensions published by other spreadsheet projects so
 * Ascend can be measured on their benchmark shapes, not only ours.
 *
 * Run:
 *   bun run fixtures/benchmarks/upstream-profiles.ts --profile all --json
 *   bun run fixtures/benchmarks/upstream-profiles.ts --profile fastexcel-reader-65536 --repeat 5 --json
 */
import { resolve } from 'node:path'
import type { ReadSource, WorkloadName } from './competitive-io.ts'
import type { BenchmarkCaseResult, BenchmarkSuiteResult } from './results.ts'
import { createBenchmarkSuite, formatBytes, formatRate } from './results.ts'

type UpstreamProfileName =
	| 'openpyxl-write-1000x50-10pct-text'
	| 'xlsxwriter-write-memory-200x50-50pct-text'
	| 'xlsxwriter-write-memory-400x50-50pct-text'
	| 'xlsxwriter-write-memory-800x50-50pct-text'
	| 'xlsxwriter-write-memory-1600x50-50pct-text'
	| 'xlsxwriter-write-memory-3200x50-50pct-text'
	| 'xlsxwriter-write-memory-6400x50-50pct-text'
	| 'xlsxwriter-write-memory-12800x50-50pct-text'
	| 'pyexcelerate-write-values-1000x100'
	| 'pyexcelerate-write-styles-1000x100'
	| 'apache-poi-ssperformance-xssf-50000x50'
	| 'excelize-generation-102400x50-plain-text'
	| 'fastexcel-writer-100000x4'
	| 'fastexcel-reader-65536'
	| 'rust-xlsxwriter-write-4000x50-50pct-text'
	| 'fastxlsx-read-5000x10-matrix'
	| 'fastxlsx-write-5000x10-matrix'
	| 'pyopenxlsx-read-1000x20'
	| 'pyopenxlsx-write-5000x10'
	| 'pyopenxlsx-bulk-write-50000x20'
	| 'pyfastexcel-write-50x30'
	| 'pyfastexcel-write-500x30'
	| 'pyfastexcel-write-5000x30'
	| 'pyfastexcel-write-50000x30'

type Category = 'read' | 'write'
type Competitor = 'js' | 'external' | 'all'

export interface UpstreamProfile {
	readonly name: UpstreamProfileName
	readonly sourceLibrary: string
	readonly sourceBenchmark: string
	readonly sourceUrl: string
	readonly category: Category
	readonly workload: WorkloadName
	readonly rows: number
	readonly cols: number
	readonly sheets: number
	readonly readSource?: ReadSource
	readonly competitor: Competitor
	readonly notes: string
}

export const UPSTREAM_PROFILES = [
	{
		name: 'openpyxl-write-1000x50-10pct-text',
		sourceLibrary: 'openpyxl',
		sourceBenchmark: 'openpyxl 3.0 write performance',
		sourceUrl: 'https://openpyxl.readthedocs.io/en/3.0/performance.html',
		category: 'write',
		workload: 'mixed-10pct-text',
		rows: 1000,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Rows=1000, cols=50, one sheet, 10% text; compares openpyxl and XlsxWriter.',
	},
	{
		name: 'xlsxwriter-write-memory-200x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 200,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Published XlsxWriter scaling table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'xlsxwriter-write-memory-400x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 400,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Published XlsxWriter scaling table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'xlsxwriter-write-memory-800x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 800,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Published XlsxWriter scaling table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'xlsxwriter-write-memory-1600x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 1600,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Published XlsxWriter scaling table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'xlsxwriter-write-memory-3200x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 3200,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Published XlsxWriter scaling table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'xlsxwriter-write-memory-6400x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 6400,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Published XlsxWriter scaling table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'xlsxwriter-write-memory-12800x50-50pct-text',
		sourceLibrary: 'XlsxWriter',
		sourceBenchmark: 'XlsxWriter memory/performance figures',
		sourceUrl: 'https://xlsxwriter.readthedocs.io/working_with_memory.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 12800,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'Largest published XlsxWriter table row count; 50 columns and 50/50 strings/numbers.',
	},
	{
		name: 'pyexcelerate-write-values-1000x100',
		sourceLibrary: 'PyExcelerate',
		sourceBenchmark: 'PyExcelerate value write benchmark',
		sourceUrl: 'https://github.com/kz26/PyExcelerate',
		category: 'write',
		workload: 'dense-values',
		rows: 1000,
		cols: 100,
		sheets: 1,
		competitor: 'all',
		notes: 'Published value benchmark shape for PyExcelerate, XlsxWriter, and openpyxl.',
	},
	{
		name: 'pyexcelerate-write-styles-1000x100',
		sourceLibrary: 'PyExcelerate',
		sourceBenchmark: 'PyExcelerate style write benchmark',
		sourceUrl: 'https://github.com/kz26/PyExcelerate',
		category: 'write',
		workload: 'styles-heavy',
		rows: 1000,
		cols: 100,
		sheets: 1,
		competitor: 'all',
		notes: 'Published style benchmark shape mapped onto Ascend style-heavy correctness checks.',
	},
	{
		name: 'apache-poi-ssperformance-xssf-50000x50',
		sourceLibrary: 'Apache POI',
		sourceBenchmark: 'SSPerformanceTest XSSF generation check',
		sourceUrl: 'https://poi.apache.org/help/faq',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 50000,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes: 'POI FAQ uses SSPerformanceTest with XSSF, 50,000 rows and 50 columns.',
	},
	{
		name: 'excelize-generation-102400x50-plain-text',
		sourceLibrary: 'Excelize',
		sourceBenchmark: 'Excelize performance comparison of similar libs',
		sourceUrl: 'https://xuri.me/excelize/en/performance.html',
		category: 'write',
		workload: 'plain-text',
		rows: 102400,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes:
			'Excelize publishes a generation benchmark for a 102400 x 50 plain text matrix across Go, Python, Java, PHP, and NodeJS Excel libraries.',
	},
	{
		name: 'fastexcel-writer-100000x4',
		sourceLibrary: 'fastexcel',
		sourceBenchmark: 'fastexcel writer benchmark',
		sourceUrl: 'https://github.com/dhatim/fastexcel',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 100000,
		cols: 4,
		sheets: 1,
		competitor: 'all',
		notes:
			'fastexcel publishes a single-worksheet writer benchmark with 100,000 rows and 4 columns.',
	},
	{
		name: 'fastexcel-reader-65536',
		sourceLibrary: 'fastexcel-reader',
		sourceBenchmark: 'fastexcel-reader 65,536-line read benchmark',
		sourceUrl: 'https://github.com/dhatim/fastexcel',
		category: 'read',
		workload: 'mixed-50pct-text',
		rows: 65536,
		cols: 10,
		sheets: 1,
		readSource: 'raw-ooxml',
		competitor: 'all',
		notes:
			'fastexcel publishes a 65,536-line read comparison; columns are fixed here to the value/label mixed reader lane.',
	},
	{
		name: 'rust-xlsxwriter-write-4000x50-50pct-text',
		sourceLibrary: 'rust_xlsxwriter',
		sourceBenchmark: 'rust_xlsxwriter hyperfine performance program',
		sourceUrl: 'https://rustxlsxwriter.github.io/performance.html',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 4000,
		cols: 50,
		sheets: 1,
		competitor: 'all',
		notes:
			'Official rust_xlsxwriter benchmark writes 4,000 rows x 50 columns with alternating strings and numbers.',
	},
	{
		name: 'fastxlsx-read-5000x10-matrix',
		sourceLibrary: 'fastxlsx',
		sourceBenchmark: 'FastXLSX reading performance matrix benchmark',
		sourceUrl: 'https://pypi.org/project/fastxlsx/',
		category: 'read',
		workload: 'mixed-50pct-text',
		rows: 5000,
		cols: 10,
		sheets: 1,
		readSource: 'raw-ooxml',
		competitor: 'all',
		notes: 'FastXLSX publishes a 5,000 x 10 matrix read benchmark against pycalamine and openpyxl.',
	},
	{
		name: 'fastxlsx-write-5000x10-matrix',
		sourceLibrary: 'fastxlsx',
		sourceBenchmark: 'FastXLSX writing performance matrix benchmark',
		sourceUrl: 'https://pypi.org/project/fastxlsx/',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 5000,
		cols: 10,
		sheets: 1,
		competitor: 'all',
		notes:
			'FastXLSX publishes a 5,000 x 10 matrix write benchmark against pyexcelerate, XlsxWriter, and openpyxl.',
	},
	{
		name: 'pyopenxlsx-read-1000x20',
		sourceLibrary: 'pyopenxlsx',
		sourceBenchmark: 'pyopenxlsx 20,000-cell read benchmark',
		sourceUrl: 'https://pypi.org/project/pyopenxlsx/0.3.2/',
		category: 'read',
		workload: 'dense-values',
		rows: 1000,
		cols: 20,
		sheets: 1,
		readSource: 'raw-ooxml',
		competitor: 'all',
		notes: 'pyopenxlsx publishes a 20,000-cell read comparison against openpyxl.',
	},
	{
		name: 'pyopenxlsx-write-5000x10',
		sourceLibrary: 'pyopenxlsx',
		sourceBenchmark: 'pyopenxlsx 50,000-cell write benchmark',
		sourceUrl: 'https://pypi.org/project/pyopenxlsx/0.3.2/',
		category: 'write',
		workload: 'dense-values',
		rows: 5000,
		cols: 10,
		sheets: 1,
		competitor: 'all',
		notes: 'pyopenxlsx publishes a 50,000-cell write comparison against openpyxl.',
	},
	{
		name: 'pyopenxlsx-bulk-write-50000x20',
		sourceLibrary: 'pyopenxlsx',
		sourceBenchmark: 'pyopenxlsx 1,000,000-cell bulk write benchmark',
		sourceUrl: 'https://pypi.org/project/pyopenxlsx/0.3.2/',
		category: 'write',
		workload: 'dense-values',
		rows: 50000,
		cols: 20,
		sheets: 1,
		competitor: 'all',
		notes:
			'pyopenxlsx publishes a 1,000,000-cell bulk write benchmark and resource-usage comparison against openpyxl.',
	},
	{
		name: 'pyfastexcel-write-50x30',
		sourceLibrary: 'pyfastexcel',
		sourceBenchmark: 'pyfastexcel 30-column write scaling benchmark',
		sourceUrl: 'https://pyfastexcel.readthedocs.io/en/latest/benchmark/',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 50,
		cols: 30,
		sheets: 1,
		competitor: 'all',
		notes: 'pyfastexcel publishes write scaling results for 50 rows x 30 columns.',
	},
	{
		name: 'pyfastexcel-write-500x30',
		sourceLibrary: 'pyfastexcel',
		sourceBenchmark: 'pyfastexcel 30-column write scaling benchmark',
		sourceUrl: 'https://pyfastexcel.readthedocs.io/en/latest/benchmark/',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 500,
		cols: 30,
		sheets: 1,
		competitor: 'all',
		notes: 'pyfastexcel publishes write scaling results for 500 rows x 30 columns.',
	},
	{
		name: 'pyfastexcel-write-5000x30',
		sourceLibrary: 'pyfastexcel',
		sourceBenchmark: 'pyfastexcel 30-column write scaling benchmark',
		sourceUrl: 'https://pyfastexcel.readthedocs.io/en/latest/benchmark/',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 5000,
		cols: 30,
		sheets: 1,
		competitor: 'all',
		notes: 'pyfastexcel publishes write scaling results for 5,000 rows x 30 columns.',
	},
	{
		name: 'pyfastexcel-write-50000x30',
		sourceLibrary: 'pyfastexcel',
		sourceBenchmark: 'pyfastexcel 30-column write scaling benchmark',
		sourceUrl: 'https://pyfastexcel.readthedocs.io/en/latest/benchmark/',
		category: 'write',
		workload: 'mixed-50pct-text',
		rows: 50000,
		cols: 30,
		sheets: 1,
		competitor: 'all',
		notes: 'pyfastexcel publishes write scaling results for 50,000 rows x 30 columns.',
	},
] as const satisfies readonly UpstreamProfile[]

const DEFAULT_READ_RUNNER_MANIFEST =
	'fixtures/benchmarks/runners/ascend-python-readers.manifest.json'
const DEFAULT_WRITE_RUNNER_MANIFEST = 'fixtures/benchmarks/runners/sota-writers.manifest.json'

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

function readCompetitorOverrideFlag(): Competitor | undefined {
	const raw = readFlag('--competitor')
	if (raw === undefined) return undefined
	if (raw === 'js' || raw === 'external' || raw === 'all') return raw
	throw new Error('Unsupported --competitor value. Expected js, external, or all.')
}

export function selectUpstreamProfiles(selection: string | undefined): readonly UpstreamProfile[] {
	if (selection === undefined || selection === 'all') return UPSTREAM_PROFILES
	const names = new Set(
		selection
			.split(',')
			.map((name) => name.trim())
			.filter(Boolean),
	)
	const selected = UPSTREAM_PROFILES.filter((profile) => names.has(profile.name))
	if (selected.length !== names.size) {
		const known = UPSTREAM_PROFILES.map((profile) => profile.name).join(', ')
		throw new Error(`Unsupported --profile value. Expected one or more of: ${known}`)
	}
	return selected
}

export function buildCompetitiveIoArgs(input: {
	readonly profile: UpstreamProfile
	readonly repeat: number
	readonly warmup: number
	readonly competitorOverride?: Competitor
	readonly libraries?: string
	readonly validationMode?: string
	readonly executionScope?: string
	readonly readRunnerManifest?: string
	readonly writeRunnerManifest?: string
}): string[] {
	const competitor = input.competitorOverride ?? input.profile.competitor
	const args = [
		'run',
		'fixtures/benchmarks/competitive-io.ts',
		'--json',
		'--category',
		input.profile.category,
		'--competitor',
		competitor,
		'--workload',
		input.profile.workload,
		'--rows',
		String(input.profile.rows),
		'--cols',
		String(input.profile.cols),
		'--repeat',
		String(input.repeat),
		'--warmup',
		String(input.warmup),
	]
	if (input.libraries) {
		args.push('--libraries', input.libraries)
	}
	if (input.validationMode) {
		args.push('--validation-mode', input.validationMode)
	}
	if (input.executionScope) {
		args.push('--execution-scope', input.executionScope)
	}
	if (input.profile.category === 'read') {
		args.push('--read-source', input.profile.readSource ?? 'raw-ooxml')
		args.push('--runner-manifest', input.readRunnerManifest ?? DEFAULT_READ_RUNNER_MANIFEST)
	}
	if (input.profile.category === 'write' && input.writeRunnerManifest) {
		args.push('--write-runner-manifest', input.writeRunnerManifest)
	}
	return args
}

async function runProfile(input: {
	readonly profile: UpstreamProfile
	readonly repeat: number
	readonly warmup: number
	readonly competitorOverride?: Competitor
	readonly libraries?: string
	readonly validationMode?: string
	readonly readRunnerManifest?: string
	readonly writeRunnerManifest?: string
}): Promise<BenchmarkSuiteResult> {
	const args = buildCompetitiveIoArgs(input)
	const proc = Bun.spawn(['bun', ...args], {
		cwd: resolve('.'),
		stdout: 'pipe',
		stderr: 'pipe',
	})
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(
			`Profile ${input.profile.name} failed with exit code ${exitCode}\n${stderr}\n${stdout}`,
		)
	}
	return JSON.parse(stdout) as BenchmarkSuiteResult
}

export function annotateUpstreamCases(
	suite: BenchmarkSuiteResult,
	profile: UpstreamProfile,
): readonly BenchmarkCaseResult[] {
	return suite.cases.map((entry) => ({
		...entry,
		name: `${profile.name}:${entry.name}`,
		dimensions: {
			...entry.dimensions,
			file: profile.name,
			upstreamProfile: profile.name,
			upstreamSourceLibrary: profile.sourceLibrary,
			upstreamSourceBenchmark: profile.sourceBenchmark,
			upstreamSourceUrl: profile.sourceUrl,
			upstreamSheets: profile.sheets,
		},
	}))
}

function renderSummary(suite: BenchmarkSuiteResult): string {
	const headers = ['profile', 'case', 'category', 'median-ms', 'throughput', 'status']
	const rows = suite.cases.map((entry) => [
		String(entry.dimensions.upstreamProfile ?? 'unknown'),
		entry.name.replace(`${String(entry.dimensions.upstreamProfile ?? '')}:`, ''),
		entry.category,
		entry.metrics.medianMs.toFixed(2),
		entry.metrics.throughputPerSec === undefined
			? 'n/a'
			: formatRate(entry.metrics.throughputPerSec),
		String(entry.dimensions.correctnessStatus ?? 'unknown'),
	])
	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')
	return [
		line(headers),
		widths.map((width) => '-'.repeat(width)).join('--'),
		...rows.map(line),
	].join('\n')
}

async function main(): Promise<void> {
	const profiles = selectUpstreamProfiles(readFlag('--profile'))
	const repeat = readPositiveIntFlag('--repeat', 3)
	const warmup = readNonNegativeIntFlag('--warmup', 1)
	const competitorOverride = readCompetitorOverrideFlag()
	const libraries = readFlag('--libraries')
	const validationMode = readFlag('--validation-mode')
	const executionScope = readFlag('--execution-scope')
	const readRunnerManifest = readFlag('--runner-manifest') ?? DEFAULT_READ_RUNNER_MANIFEST
	const writeRunnerManifest = readFlag('--write-runner-manifest') ?? DEFAULT_WRITE_RUNNER_MANIFEST
	const cases: BenchmarkCaseResult[] = []
	const childSuites: Array<Record<string, unknown>> = []
	for (const profile of profiles) {
		const suite = await runProfile({
			profile,
			repeat,
			warmup,
			competitorOverride,
			libraries,
			validationMode,
			executionScope,
			readRunnerManifest,
			writeRunnerManifest,
		})
		cases.push(...annotateUpstreamCases(suite, profile))
		childSuites.push({
			profile: profile.name,
			suite: suite.suite,
			generatedAt: suite.generatedAt,
			cases: suite.cases.length,
		})
		if (!hasFlag('--json')) {
			console.error(`completed ${profile.name}: ${formatBytes(sumBytes(suite.cases))} input data`)
		}
	}
	const suite = createBenchmarkSuite({
		suite: 'ascend-upstream-competitive-profiles',
		kind: 'real-workbook',
		cases,
		metadata: {
			profiles,
			repeat,
			warmup,
			competitorOverride,
			libraries,
			validationMode,
			childSuites,
			readRunnerManifest,
			writeRunnerManifest,
		},
	})
	if (hasFlag('--json')) {
		console.log(JSON.stringify(suite, null, 2))
		return
	}
	console.log(renderSummary(suite))
}

function sumBytes(cases: readonly BenchmarkCaseResult[]): number {
	return cases.reduce((total, entry) => {
		const bytes = entry.dimensions.bytes
		return total + (typeof bytes === 'number' ? bytes : 0)
	}, 0)
}

if (import.meta.main) {
	await main()
}
