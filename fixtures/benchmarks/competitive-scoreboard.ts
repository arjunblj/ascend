import { readFile } from 'node:fs/promises'
import type { BenchmarkCaseResult, BenchmarkSuiteResult } from './results.ts'
import { formatBytes, formatRate, welchsTTest } from './results.ts'

type RankingMetric =
	| 'medianMs'
	| 'p95Ms'
	| 'throughputPerSec'
	| 'rssDeltaBytes'
	| 'retainedRssDeltaBytes'
	| 'rssAfterBytes'
	| 'rssAfterGcBytes'
	| 'peakRssBytes'
	| 'heapUsedBytes'
	| 'heapTotalBytes'
	| 'heapAfterGcBytes'
	| 'heapDeltaBytes'

type ClaimProfileName =
	| 'best-js-generated-io'
	| 'xlsx-read-sota'
	| 'xlsx-write-sota'
	| 'upstream-xlsx-sota'

const ASCEND_READ_LIBRARIES = [
	'ascend',
	'ascend-external',
	'ascend-external-values',
	'ascend-external-values-ordered',
	'ascend-external-bytes',
	'ascend-external-values-bytes',
	'ascend-readxlsx-raw-values-bytes',
	'ascend-readxlsx-values-bytes',
	'ascend-readxlsx-values-rich-metadata-bytes',
	'ascend-external-metadata-only-bytes',
] as const
const ASCEND_WRITE_LIBRARIES = ['ascend', 'ascend-external-writer'] as const

const RANKING_METRICS = [
	'medianMs',
	'p95Ms',
	'throughputPerSec',
	'rssDeltaBytes',
	'retainedRssDeltaBytes',
	'rssAfterBytes',
	'rssAfterGcBytes',
	'peakRssBytes',
	'heapUsedBytes',
	'heapTotalBytes',
	'heapAfterGcBytes',
	'heapDeltaBytes',
] as const satisfies readonly RankingMetric[]

const EMPTY_HASH = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'

export interface CompetitiveScoreboardEntry {
	readonly caseName: string
	readonly library: string
	readonly category: string
	readonly workload: string
	readonly readSource: string
	readonly operationProfile: string
	readonly timingLane: string
	readonly file: string
	readonly correctnessStatus: string
	readonly rankingEligible: boolean
	readonly fidelityTier: number
	readonly fidelityLabel: string
	readonly featureObligationsMet: number
	readonly featureObligationsTotal: number
	readonly featureLabel: string
	readonly metric: RankingMetric
	readonly value: number | null
	readonly rank: number | null
	readonly speedRank: number | null
	readonly slowdownVsFastest: number | null
	readonly metricSamples?: readonly number[]
	readonly disqualifiedReason?: string
}

export interface CompetitiveScoreboardGroup {
	readonly category: string
	readonly workload: string
	readonly readSource: string
	readonly operationProfile: string
	readonly timingLane: string
	readonly file: string
	readonly metric: RankingMetric
	readonly winner: string | null
	readonly fastestEligible: string | null
	readonly entries: readonly CompetitiveScoreboardEntry[]
}

export interface CompetitiveScoreboardLibrarySummary {
	readonly library: string
	readonly cases: number
	readonly eligible: number
	readonly disqualified: number
	readonly wins: number
	readonly averageRank: number | null
	readonly geomeanSlowdownVsFastest: number | null
}

export interface CompetitiveScoreboard {
	readonly suite: string
	readonly metric: RankingMetric
	readonly groups: readonly CompetitiveScoreboardGroup[]
	readonly libraries: readonly CompetitiveScoreboardLibrarySummary[]
}

export interface CompetitiveScoreboardOptions {
	readonly metric?: RankingMetric
	readonly category?: string
}

export interface ScoreboardCoverageInspection {
	readonly failures: readonly string[]
	readonly gaps: readonly string[]
}

interface CompetitorRequirement {
	readonly label: string
	readonly libraries: readonly string[]
}

interface CaseRequirement {
	readonly category: string
	readonly operationProfile: string
	readonly workloads: readonly string[]
	readonly readSources?: readonly string[]
	readonly files?: readonly string[]
	readonly timingLanePrefix?: string
	readonly competitors?: readonly CompetitorRequirement[]
	readonly capabilityGaps?: readonly CompetitorRequirement[]
}

interface ClaimProfile {
	readonly name: ClaimProfileName
	readonly minRepeat: number
	readonly requireSamples: boolean
	readonly requiredMetrics: readonly RankingMetric[]
	readonly competitors: readonly CompetitorRequirement[]
	readonly cases: readonly CaseRequirement[]
}

const CLAIM_PROFILES: Record<ClaimProfileName, ClaimProfile> = {
	'best-js-generated-io': {
		name: 'best-js-generated-io',
		minRepeat: 3,
		requireSamples: true,
		requiredMetrics: ['medianMs', 'p95Ms', 'throughputPerSec'],
		competitors: [
			{ label: 'Ascend', libraries: ['ascend'] },
			{ label: 'SheetJS', libraries: ['sheetjs'] },
			{ label: 'ExcelJS', libraries: ['exceljs'] },
		],
		cases: [
			{
				category: 'read',
				operationProfile: 'read-values',
				workloads: ['dense-values', 'string-heavy', 'sparse-wide'],
				readSources: ['ascend-writer', 'raw-ooxml'],
				timingLanePrefix: 'in-process-generated-',
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['dense-values', 'string-heavy', 'sparse-wide'],
				readSources: ['ascend-writer'],
				timingLanePrefix: 'in-process-generated-',
			},
		],
	},
	'xlsx-read-sota': {
		name: 'xlsx-read-sota',
		minRepeat: 5,
		requireSamples: true,
		requiredMetrics: ['medianMs', 'p95Ms', 'throughputPerSec', 'peakRssBytes'],
		competitors: [
			{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES },
			{ label: 'SheetJS', libraries: ['sheetjs'] },
			{ label: 'ExcelJS', libraries: ['exceljs'] },
			{
				label: 'openpyxl',
				libraries: ['openpyxl', 'openpyxl-read-only', 'openpyxl-read-only-values'],
			},
			{
				label: 'Calamine',
				libraries: ['calamine', 'python-calamine', 'rust-calamine', 'fastexcel'],
			},
			{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
			{ label: 'ClosedXML', libraries: ['closedxml'] },
		],
		cases: [
			{
				category: 'read',
				operationProfile: 'read-values',
				workloads: [
					'dense-values',
					'sparse-wide',
					'string-heavy',
					'styles-heavy',
					'formula-heavy',
					'table-heavy',
				],
			},
			{
				category: 'read',
				operationProfile: 'read-values-rich-metadata',
				workloads: ['feature-rich'],
			},
			{
				category: 'read',
				operationProfile: 'read-selected-values',
				workloads: ['selected-sheet'],
				competitors: [{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES }],
				capabilityGaps: [
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{
						label: 'openpyxl',
						libraries: ['openpyxl', 'openpyxl-read-only', 'openpyxl-read-only-values'],
					},
					{
						label: 'Calamine',
						libraries: ['calamine', 'python-calamine', 'rust-calamine', 'fastexcel'],
					},
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
				],
			},
			{
				category: 'read',
				operationProfile: 'read-metadata-only',
				workloads: ['metadata-only'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'openpyxl', libraries: ['openpyxl-metadata-only'] },
				],
				capabilityGaps: [
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{
						label: 'Calamine',
						libraries: ['calamine', 'python-calamine', 'rust-calamine', 'fastexcel'],
					},
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
				],
			},
			{
				category: 'read',
				operationProfile: 'read-values-warm',
				workloads: ['warm-workflow'],
			},
		],
	},
	'xlsx-write-sota': {
		name: 'xlsx-write-sota',
		minRepeat: 5,
		requireSamples: true,
		requiredMetrics: ['medianMs', 'p95Ms', 'throughputPerSec', 'peakRssBytes'],
		competitors: [
			{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
			{ label: 'SheetJS', libraries: ['sheetjs'] },
			{ label: 'ExcelJS', libraries: ['exceljs'] },
			{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
			{
				label: 'pyexcelerate',
				libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
			},
			{ label: 'FastXLSX', libraries: ['fastxlsx'] },
			{ label: 'pyopenxlsx', libraries: ['pyopenxlsx'] },
			{ label: 'pyfastexcel', libraries: ['pyfastexcel'] },
			{ label: 'fastexcel Java', libraries: ['fastexcel-java'] },
			{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
			{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
			{ label: 'ClosedXML', libraries: ['closedxml'] },
			{ label: 'NPOI', libraries: ['npoi'] },
			{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
			{ label: 'Excelize', libraries: ['excelize'] },
		],
		cases: [
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['dense-values', 'plain-text', 'sparse-wide', 'string-heavy', 'styles-heavy'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{
						label: 'pyexcelerate',
						libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
					},
					{ label: 'fastexcel Java', libraries: ['fastexcel-java'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
					{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
					{ label: 'Excelize', libraries: ['excelize'] },
				],
				capabilityGaps: [
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{ label: 'NPOI', libraries: ['npoi'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['formula-heavy'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
				],
				capabilityGaps: [
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{
						label: 'pyexcelerate',
						libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
					},
					{ label: 'fastexcel Java', libraries: ['fastexcel-java'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
					{ label: 'NPOI', libraries: ['npoi'] },
					{ label: 'Excelize', libraries: ['excelize'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['table-heavy'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter'] },
					{ label: 'openpyxl', libraries: ['openpyxl'] },
					{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
				],
				capabilityGaps: [
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{
						label: 'pyexcelerate',
						libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
					},
					{ label: 'fastexcel Java', libraries: ['fastexcel-java'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
					{ label: 'NPOI', libraries: ['npoi'] },
					{ label: 'Excelize', libraries: ['excelize'] },
					{ label: 'XlsxWriter constant-memory', libraries: ['xlsxwriter-constant-memory'] },
					{ label: 'openpyxl write-only', libraries: ['openpyxl-write-only'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['feature-rich'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter'] },
					{ label: 'openpyxl', libraries: ['openpyxl'] },
				],
				capabilityGaps: [
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{
						label: 'pyexcelerate',
						libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
					},
					{ label: 'fastexcel Java', libraries: ['fastexcel-java'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
					{ label: 'NPOI', libraries: ['npoi'] },
					{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
					{ label: 'Excelize', libraries: ['excelize'] },
					{ label: 'XlsxWriter constant-memory', libraries: ['xlsxwriter-constant-memory'] },
					{ label: 'openpyxl write-only', libraries: ['openpyxl-write-only'] },
				],
			},
		],
	},
	'upstream-xlsx-sota': {
		name: 'upstream-xlsx-sota',
		minRepeat: 5,
		requireSamples: true,
		requiredMetrics: ['medianMs', 'p95Ms', 'throughputPerSec', 'peakRssBytes'],
		competitors: [
			{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES },
			{ label: 'SheetJS', libraries: ['sheetjs'] },
			{ label: 'ExcelJS', libraries: ['exceljs'] },
			{
				label: 'openpyxl',
				libraries: ['openpyxl', 'openpyxl-read-only', 'openpyxl-read-only-values'],
			},
			{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
			{
				label: 'pyexcelerate',
				libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
			},
			{ label: 'FastXLSX', libraries: ['fastxlsx'] },
			{ label: 'pyopenxlsx', libraries: ['pyopenxlsx'] },
			{
				label: 'Calamine',
				libraries: ['calamine', 'python-calamine', 'rust-calamine', 'fastexcel'],
			},
			{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
			{ label: 'ClosedXML', libraries: ['closedxml'] },
			{
				label: 'Polars',
				libraries: ['polars', 'polars-calamine', 'polars-xlsx2csv', 'polars-openpyxl'],
			},
			{ label: 'pyfastexcel', libraries: ['pyfastexcel'] },
			{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
			{ label: 'Excelize', libraries: ['excelize'] },
		],
		cases: [
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['mixed-10pct-text'],
				files: ['openpyxl-write-1000x50-10pct-text'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['plain-text'],
				files: ['excelize-generation-102400x50-plain-text'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{
						label: 'pyexcelerate',
						libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
					},
					{ label: 'FastXLSX', libraries: ['fastxlsx'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'Excelize', libraries: ['excelize'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['mixed-50pct-text'],
				files: [
					'xlsxwriter-write-memory-200x50-50pct-text',
					'xlsxwriter-write-memory-400x50-50pct-text',
					'xlsxwriter-write-memory-800x50-50pct-text',
					'xlsxwriter-write-memory-1600x50-50pct-text',
					'xlsxwriter-write-memory-3200x50-50pct-text',
					'xlsxwriter-write-memory-6400x50-50pct-text',
					'xlsxwriter-write-memory-12800x50-50pct-text',
					'apache-poi-ssperformance-xssf-50000x50',
					'fastexcel-writer-100000x4',
					'rust-xlsxwriter-write-4000x50-50pct-text',
					'fastxlsx-write-5000x10-matrix',
					'pyfastexcel-write-50x30',
					'pyfastexcel-write-500x30',
					'pyfastexcel-write-5000x30',
					'pyfastexcel-write-50000x30',
				],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{
						label: 'pyexcelerate',
						libraries: ['pyexcelerate', 'pyexcelerate-range', 'pyexcelerate-cell'],
					},
					{ label: 'FastXLSX', libraries: ['fastxlsx'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'fastexcel Java', libraries: ['fastexcel-java'] },
					{ label: 'NPOI', libraries: ['npoi'] },
					{ label: 'rust_xlsxwriter', libraries: ['rust-xlsxwriter', 'rust_xlsxwriter'] },
					{ label: 'pyfastexcel', libraries: ['pyfastexcel'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['dense-values'],
				files: [
					'pyexcelerate-write-values-1000x100',
					'pyopenxlsx-write-5000x10',
					'pyopenxlsx-bulk-write-50000x20',
				],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{ label: 'pyexcelerate bulk-sheet', libraries: ['pyexcelerate'] },
					{ label: 'pyexcelerate range', libraries: ['pyexcelerate-range'] },
					{ label: 'pyexcelerate cell', libraries: ['pyexcelerate-cell'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'pyopenxlsx', libraries: ['pyopenxlsx'] },
				],
			},
			{
				category: 'write',
				operationProfile: 'write-values',
				workloads: ['styles-heavy'],
				files: ['pyexcelerate-write-styles-1000x100'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_WRITE_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{ label: 'XlsxWriter', libraries: ['xlsxwriter', 'xlsxwriter-constant-memory'] },
					{ label: 'pyexcelerate bulk-sheet', libraries: ['pyexcelerate'] },
					{ label: 'pyexcelerate range', libraries: ['pyexcelerate-range'] },
					{ label: 'pyexcelerate cell', libraries: ['pyexcelerate-cell'] },
					{ label: 'openpyxl', libraries: ['openpyxl', 'openpyxl-write-only'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
				],
			},
			{
				category: 'read',
				operationProfile: 'read-values',
				workloads: ['mixed-50pct-text'],
				readSources: ['raw-ooxml'],
				files: ['fastexcel-reader-65536', 'fastxlsx-read-5000x10-matrix'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES },
					{ label: 'SheetJS', libraries: ['sheetjs'] },
					{ label: 'ExcelJS', libraries: ['exceljs'] },
					{
						label: 'openpyxl',
						libraries: ['openpyxl', 'openpyxl-read-only', 'openpyxl-read-only-values'],
					},
					{
						label: 'Calamine',
						libraries: ['calamine', 'python-calamine', 'rust-calamine', 'fastexcel'],
					},
					{ label: 'FastXLSX', libraries: ['fastxlsx'] },
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
					{
						label: 'Polars calamine',
						libraries: ['polars-calamine'],
					},
					{
						label: 'Polars xlsx2csv',
						libraries: ['polars-xlsx2csv'],
					},
					{
						label: 'Polars openpyxl',
						libraries: ['polars-openpyxl'],
					},
					{ label: 'Excelize', libraries: ['excelize'] },
				],
			},
			{
				category: 'read',
				operationProfile: 'read-values',
				workloads: ['dense-values'],
				readSources: ['raw-ooxml'],
				files: ['pyopenxlsx-read-1000x20'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES },
					{
						label: 'openpyxl',
						libraries: ['openpyxl', 'openpyxl-read-only', 'openpyxl-read-only-values'],
					},
					{ label: 'pyopenxlsx', libraries: ['pyopenxlsx'] },
				],
			},
			{
				category: 'read',
				operationProfile: 'read-values',
				workloads: ['calamine-nyc311-1m'],
				files: ['NYC_311_SR_2010-2020-sample-1M.xlsx'],
				competitors: [
					{ label: 'Ascend', libraries: ASCEND_READ_LIBRARIES },
					{
						label: 'Calamine',
						libraries: ['calamine', 'python-calamine', 'rust-calamine', 'fastexcel'],
					},
					{
						label: 'openpyxl',
						libraries: ['openpyxl', 'openpyxl-read-only', 'openpyxl-read-only-values'],
					},
					{ label: 'Apache POI', libraries: ['apache-poi', 'poi'] },
					{ label: 'ClosedXML', libraries: ['closedxml'] },
					{
						label: 'Polars calamine',
						libraries: ['polars-calamine'],
					},
					{
						label: 'Polars xlsx2csv',
						libraries: ['polars-xlsx2csv'],
					},
					{
						label: 'Polars openpyxl',
						libraries: ['polars-openpyxl'],
					},
					{ label: 'Excelize', libraries: ['excelize'] },
				],
			},
		],
	},
}

export function buildCompetitiveScoreboard(
	suite: BenchmarkSuiteResult,
	options: CompetitiveScoreboardOptions = {},
): CompetitiveScoreboard {
	const metric = options.metric ?? 'medianMs'
	const groups = groupCases(suite.cases, options.category).map((cases) => buildGroup(cases, metric))
	return {
		suite: suite.suite,
		metric,
		groups,
		libraries: summarizeLibraries(groups),
	}
}

export function assertScoreboardLeader(
	scoreboard: CompetitiveScoreboard,
	library: string,
): readonly string[] {
	const failures: string[] = []
	for (const group of scoreboard.groups) {
		if (group.winner === null || group.winner === library) continue
		const leader = group.entries.find((entry) => entry.library === library)
		const winner = group.entries.find((entry) => entry.library === group.winner)
		if (leader && winner && !isSignificantLeaderLoss(leader, winner)) continue
		failures.push(
			`${group.category}:${group.operationProfile}:${group.timingLane}:${group.file} winner=${group.winner} expected=${library}`,
		)
	}
	return failures
}

export function assertScoreboardProfileLeader(
	scoreboard: CompetitiveScoreboard,
	profileName: ClaimProfileName,
	library: string,
): readonly string[] {
	const profile = CLAIM_PROFILES[profileName]
	const failures: string[] = []
	for (const caseRequirement of profile.cases) {
		for (const workload of caseRequirement.workloads) {
			for (const readSource of caseRequirement.readSources ?? [undefined]) {
				for (const file of caseRequirement.files ?? [undefined]) {
					for (const group of scoreboard.groups) {
						if (!groupMatchesProfile(group, caseRequirement, workload, readSource, file)) {
							continue
						}
						const competitors = profileCompetitors(profile, caseRequirement)
						const winner = profileWinner(group, competitors)
						const leaderLibraries = profileLeaderLibraries(competitors, library)
						if (winner === null || leaderLibraries.has(winner.library)) continue
						const leader = bestProfileEntry(group, leaderLibraries)
						if (leader && !isSignificantLeaderLoss(leader, winner)) continue
						const sourceLabel = readSource ? `:${readSource}` : ''
						const fileLabel = file ? `:${file}` : ''
						failures.push(
							`${profile.name}:${group.category}:${group.operationProfile}:${group.timingLane}:${workload}${sourceLabel}${fileLabel} winner=${winner.library} expected=${library}`,
						)
					}
				}
			}
		}
	}
	return failures
}

function profileLeaderLibraries(
	competitors: readonly CompetitorRequirement[],
	library: string,
): ReadonlySet<string> {
	const competitor = competitors.find((entry) => entry.libraries.includes(library))
	return new Set(competitor?.libraries ?? [library])
}

function bestProfileEntry(
	group: CompetitiveScoreboardGroup,
	libraries: ReadonlySet<string>,
): CompetitiveScoreboardEntry | undefined {
	return group.entries
		.filter(
			(entry) => libraries.has(entry.library) && entry.rankingEligible && entry.value !== null,
		)
		.sort((a, b) => {
			if (a.fidelityTier !== b.fidelityTier) return b.fidelityTier - a.fidelityTier
			return compareMetricValues(a.value ?? 0, b.value ?? 0, group.metric)
		})[0]
}

export function assertScoreboardCoverage(
	suite: BenchmarkSuiteResult,
	profileName: ClaimProfileName,
): readonly string[] {
	return inspectScoreboardCoverage(suite, profileName).failures
}

export function inspectScoreboardCoverage(
	suite: BenchmarkSuiteResult,
	profileName: ClaimProfileName,
): ScoreboardCoverageInspection {
	const profile = CLAIM_PROFILES[profileName]
	const failures: string[] = []
	const gaps: string[] = []
	for (const caseRequirement of profile.cases) {
		for (const workload of caseRequirement.workloads) {
			for (const readSource of caseRequirement.readSources ?? [undefined]) {
				for (const file of caseRequirement.files ?? [undefined]) {
					const sourceLabel = readSource ? ` readSource=${readSource}` : ''
					const fileLabel = file ? ` file=${file}` : ''
					const competitors = profileCompetitors(profile, caseRequirement)
					const tupleFailures: string[] = []
					for (const competitor of competitors) {
						const failure = coverageFailureForCompetitor(
							suite,
							profile,
							caseRequirement,
							workload,
							readSource,
							file,
							sourceLabel,
							fileLabel,
							competitor,
							'missing',
						)
						if (failure) tupleFailures.push(failure)
					}
					failures.push(...tupleFailures)
					if (tupleFailures.length === 0) {
						const comparableFailure = comparableCoverageFailure(
							suite,
							profile,
							caseRequirement,
							workload,
							readSource,
							file,
							sourceLabel,
							fileLabel,
							competitors,
						)
						if (comparableFailure) failures.push(comparableFailure)
					}
					for (const competitor of caseRequirement.capabilityGaps ?? []) {
						gaps.push(
							`${profile.name} coverage-gap competitor=${competitor.label} category=${caseRequirement.category} operationProfile=${caseRequirement.operationProfile} workload=${workload}${sourceLabel}${fileLabel} reason=unsupported-operation`,
						)
					}
				}
			}
		}
	}
	return { failures, gaps }
}

function groupCases(
	cases: readonly BenchmarkCaseResult[],
	category: string | undefined,
): BenchmarkCaseResult[][] {
	const grouped = new Map<string, BenchmarkCaseResult[]>()
	for (const entry of cases) {
		if (category && entry.category !== category) continue
		const workload = dimensionString(entry, 'workload') ?? 'unknown-workload'
		const readSource = dimensionString(entry, 'readSource') ?? 'unknown-source'
		const file = dimensionString(entry, 'file') ?? 'unknown-file'
		const operationProfile = dimensionString(entry, 'operationProfile') ?? 'default'
		const timingLane = dimensionString(entry, 'timingLane') ?? 'default'
		const key = `${entry.category}\0${workload}\0${readSource}\0${operationProfile}\0${timingLane}\0${file}`
		const group = grouped.get(key)
		if (group) {
			group.push(entry)
		} else {
			grouped.set(key, [entry])
		}
	}
	return [...grouped.values()].sort((a, b) => {
		const left = groupSortKey(a)
		const right = groupSortKey(b)
		return left.localeCompare(right)
	})
}

function buildGroup(
	cases: readonly BenchmarkCaseResult[],
	metric: RankingMetric,
): CompetitiveScoreboardGroup {
	const category = cases[0]?.category ?? 'unknown'
	const workload = cases[0]
		? (dimensionString(cases[0], 'workload') ?? 'unknown-workload')
		: 'unknown-workload'
	const readSource = cases[0]
		? (dimensionString(cases[0], 'readSource') ?? 'unknown-source')
		: 'unknown-source'
	const operationProfile = cases[0]
		? (dimensionString(cases[0], 'operationProfile') ?? 'default')
		: 'default'
	const timingLane = cases[0] ? (dimensionString(cases[0], 'timingLane') ?? 'default') : 'default'
	const file = cases[0] ? (dimensionString(cases[0], 'file') ?? 'unknown-file') : 'unknown-file'
	const prelim = cases.map((entry) => buildEntry(entry, metric))
	const ranked = prelim
		.filter((entry) => entry.rankingEligible && entry.value !== null)
		.sort((a, b) => {
			if (a.fidelityTier !== b.fidelityTier) return b.fidelityTier - a.fidelityTier
			return compareMetricValues(a.value ?? 0, b.value ?? 0, metric)
		})
	const speedRanked = prelim
		.filter((entry) => entry.rankingEligible && entry.value !== null)
		.sort((a, b) => compareMetricValues(a.value ?? 0, b.value ?? 0, metric))
	const fastestValue = speedRanked[0]?.value ?? null
	const rankedByName = new Map(
		ranked.map((entry, index) => [
			entry.caseName,
			{
				rank: index + 1,
			},
		]),
	)
	const speedRankedByName = new Map(
		speedRanked.map((entry, index) => [
			entry.caseName,
			{
				speedRank: index + 1,
				slowdownVsFastest:
					fastestValue === null || entry.value === null
						? null
						: slowdownVsFastest(entry.value, fastestValue, metric),
			},
		]),
	)
	return {
		category,
		workload,
		readSource,
		operationProfile,
		timingLane,
		file,
		metric,
		winner: ranked[0]?.library ?? null,
		fastestEligible: speedRanked[0]?.library ?? null,
		entries: prelim
			.map((entry) => {
				const rank = rankedByName.get(entry.caseName)
				const speedRank = speedRankedByName.get(entry.caseName)
				return { ...entry, ...rank, ...speedRank }
			})
			.sort((a, b) => {
				if (a.rank !== null && b.rank !== null) return a.rank - b.rank
				if (a.rank !== null) return -1
				if (b.rank !== null) return 1
				return a.library.localeCompare(b.library)
			}),
	}
}

function buildEntry(entry: BenchmarkCaseResult, metric: RankingMetric): CompetitiveScoreboardEntry {
	const value = entry.metrics[metric] ?? null
	const correctnessStatus = dimensionString(entry, 'correctnessStatus') ?? 'not-evaluated'
	const fidelity = fidelityForStatus(correctnessStatus)
	const features = featureObligationsForEntry(entry)
	const rankingEligible = isCaseRankingEligible(entry)
	const samples = metricSamples(entry, metric)
	return {
		caseName: entry.name,
		library: dimensionString(entry, 'library') ?? inferLibrary(entry.name),
		category: entry.category,
		workload: dimensionString(entry, 'workload') ?? 'unknown-workload',
		readSource: dimensionString(entry, 'readSource') ?? 'unknown-source',
		operationProfile: dimensionString(entry, 'operationProfile') ?? 'default',
		timingLane: dimensionString(entry, 'timingLane') ?? 'default',
		file: dimensionString(entry, 'file') ?? 'unknown-file',
		correctnessStatus,
		rankingEligible,
		fidelityTier: fidelity.tier,
		fidelityLabel: fidelity.label,
		featureObligationsMet: features.met,
		featureObligationsTotal: features.total,
		featureLabel: featureLabel(features),
		metric,
		value,
		rank: null,
		speedRank: null,
		slowdownVsFastest: null,
		...(samples.length > 0 ? { metricSamples: samples } : {}),
		...(rankingEligible ? {} : { disqualifiedReason: `correctnessStatus=${correctnessStatus}` }),
	}
}

function featureObligationsForEntry(entry: BenchmarkCaseResult): { met: number; total: number } {
	const workload = dimensionString(entry, 'workload')
	if (entry.category === 'roundtrip') {
		const expectedPartHash = assertionString(entry, 'expectedFeaturePartNamesHash')
		const expectedInventoryHash = assertionString(entry, 'expectedFeatureInventoryHash')
		if (
			(!expectedPartHash || expectedPartHash === EMPTY_HASH) &&
			(!expectedInventoryHash || expectedInventoryHash === EMPTY_HASH)
		) {
			return { met: 0, total: 0 }
		}
		const checks = [
			assertionBoolean(entry, 'roundtripTablePartCountMatches'),
			assertionBoolean(entry, 'roundtripChartPartCountMatches'),
			assertionBoolean(entry, 'roundtripChartExPartCountMatches'),
			assertionBoolean(entry, 'roundtripDrawingPartCountMatches'),
			assertionBoolean(entry, 'roundtripVmlDrawingPartCountMatches'),
			assertionBoolean(entry, 'roundtripPivotTablePartCountMatches'),
			assertionBoolean(entry, 'roundtripPivotCachePartCountMatches'),
			assertionBoolean(entry, 'roundtripSlicerPartCountMatches'),
			assertionBoolean(entry, 'roundtripCommentPartCountMatches'),
			assertionBoolean(entry, 'roundtripThreadedCommentPartCountMatches'),
			assertionBoolean(entry, 'roundtripMediaPartCountMatches'),
			assertionBoolean(entry, 'roundtripExternalLinkPartCountMatches'),
			assertionBoolean(entry, 'roundtripConnectionPartCountMatches'),
			assertionBoolean(entry, 'roundtripCustomXmlPartCountMatches'),
			assertionBoolean(entry, 'roundtripWorksheetHyperlinkCountMatches'),
			assertionBoolean(entry, 'roundtripWorksheetDataValidationCountMatches'),
			assertionBoolean(entry, 'roundtripWorksheetConditionalFormattingCountMatches'),
			assertionBoolean(entry, 'roundtripDefinedNameCountMatches'),
			assertionBoolean(entry, 'roundtripFeaturePartNamesHashMatches'),
			assertionBoolean(entry, 'roundtripFeatureInventoryHashMatches'),
		]
		return {
			met: checks.filter((value) => value === true).length,
			total: checks.length,
		}
	}
	if (entry.category === 'read') {
		if (workload === 'styles-heavy' && hasAssertion(entry, 'readStyleCount')) {
			return {
				met: assertionNumber(entry, 'readStyleCount') > 1 ? 1 : 0,
				total: 1,
			}
		}
		if (workload === 'formula-heavy' && hasAssertion(entry, 'readFormulaCellCount')) {
			return {
				met: assertionNumber(entry, 'readFormulaCellCount') > 0 ? 1 : 0,
				total: 1,
			}
		}
		if (workload === 'table-heavy' && hasAssertion(entry, 'readTableCount')) {
			return {
				met: assertionNumber(entry, 'readTableCount') > 0 ? 1 : 0,
				total: 1,
			}
		}
		if (
			workload === 'feature-rich' &&
			[
				'readCommentCount',
				'readHyperlinkCount',
				'readDataValidationCount',
				'readConditionalFormatCount',
				'readDefinedNameCount',
			].some((key) => hasAssertion(entry, key))
		) {
			const checks = [
				assertionNumber(entry, 'readCommentCount') > 0,
				assertionNumber(entry, 'readHyperlinkCount') > 0,
				assertionNumber(entry, 'readDataValidationCount') > 0,
				assertionNumber(entry, 'readConditionalFormatCount') > 0,
				assertionNumber(entry, 'readDefinedNameCount') > 0,
			]
			return {
				met: checks.filter(Boolean).length,
				total: checks.length,
			}
		}
	}
	if (entry.category !== 'write') return { met: 0, total: 0 }
	if (workload === 'formula-heavy') {
		return {
			met: assertionBoolean(entry, 'formulaCountMatches') ? 1 : 0,
			total: 1,
		}
	}
	if (workload === 'table-heavy') {
		return {
			met: assertionNumber(entry, 'tablePartCount') > 0 ? 1 : 0,
			total: 1,
		}
	}
	if (workload === 'feature-rich') {
		const checks = [
			assertionNumber(entry, 'commentPartCount') === 1,
			assertionNumber(entry, 'vmlDrawingPartCount') === 1,
			assertionNumber(entry, 'worksheetHyperlinkCount') === 1,
			assertionNumber(entry, 'worksheetDataValidationCount') === 1,
			assertionNumber(entry, 'worksheetConditionalFormattingCount') === 1,
			assertionNumber(entry, 'definedNameCount') === 1,
		]
		return {
			met: checks.filter(Boolean).length,
			total: checks.length,
		}
	}
	return { met: 0, total: 0 }
}

function featureLabel(features: { met: number; total: number }): string {
	return features.total === 0 ? 'n/a' : `${features.met}/${features.total}`
}

function groupMatchesProfile(
	group: CompetitiveScoreboardGroup,
	requirement: CaseRequirement,
	workload: string,
	readSource: string | undefined,
	file: string | undefined,
): boolean {
	if (group.category !== requirement.category) return false
	if (group.operationProfile !== requirement.operationProfile) return false
	if (group.workload !== workload) return false
	if (readSource !== undefined && group.readSource !== readSource) return false
	if (file !== undefined && group.file !== file) return false
	return requirement.timingLanePrefix
		? group.timingLane.startsWith(requirement.timingLanePrefix)
		: true
}

function profileWinner(
	group: CompetitiveScoreboardGroup,
	competitors: readonly CompetitorRequirement[],
): CompetitiveScoreboardEntry | null {
	const libraries = new Set(competitors.flatMap((competitor) => competitor.libraries))
	return (
		group.entries
			.filter(
				(entry) => libraries.has(entry.library) && entry.rankingEligible && entry.value !== null,
			)
			.sort((a, b) => {
				if (a.fidelityTier !== b.fidelityTier) return b.fidelityTier - a.fidelityTier
				return compareMetricValues(a.value ?? 0, b.value ?? 0, group.metric)
			})[0] ?? null
	)
}

function profileCompetitors(
	profile: ClaimProfile,
	requirement: CaseRequirement,
): readonly CompetitorRequirement[] {
	return requirement.competitors ?? profile.competitors
}

function coverageFailureForCompetitor(
	suite: BenchmarkSuiteResult,
	profile: ClaimProfile,
	requirement: CaseRequirement,
	workload: string,
	readSource: string | undefined,
	file: string | undefined,
	sourceLabel: string,
	fileLabel: string,
	competitor: CompetitorRequirement,
	missingKind: 'missing',
): string | null {
	const matches = suite.cases.filter((entry) =>
		caseMatchesProfile(entry, requirement, workload, readSource, file, competitor),
	)
	if (matches.length === 0) {
		return `${profile.name} ${missingKind} competitor=${competitor.label} category=${requirement.category} operationProfile=${requirement.operationProfile} workload=${workload}${sourceLabel}${fileLabel}`
	}
	const eligible = matches.filter((entry) => isCaseRankingEligible(entry))
	if (eligible.length === 0) {
		const statuses = [
			...new Set(
				matches.map((entry) => dimensionString(entry, 'correctnessStatus') ?? 'not-evaluated'),
			),
		].join(',')
		return `${profile.name} ineligible competitor=${competitor.label} category=${requirement.category} operationProfile=${requirement.operationProfile} workload=${workload}${sourceLabel}${fileLabel} correctnessStatus=${statuses}`
	}
	const valid = eligible.find((entry) => caseSatisfiesEvidence(entry, profile))
	if (!valid) {
		return `${profile.name} weak-evidence competitor=${competitor.label} category=${requirement.category} operationProfile=${requirement.operationProfile} workload=${workload}${sourceLabel}${fileLabel} requiredRepeat=${profile.minRepeat} requiredMetrics=${profile.requiredMetrics.join(',')}`
	}
	return null
}

function comparableCoverageFailure(
	suite: BenchmarkSuiteResult,
	profile: ClaimProfile,
	requirement: CaseRequirement,
	workload: string,
	readSource: string | undefined,
	file: string | undefined,
	sourceLabel: string,
	fileLabel: string,
	competitors: readonly CompetitorRequirement[],
): string | null {
	const lanes = [
		...new Set(
			suite.cases
				.filter((entry) => caseMatchesProfileTuple(entry, requirement, workload, readSource, file))
				.map((entry) => dimensionString(entry, 'timingLane') ?? 'default'),
		),
	]
	const hasSharedLane = lanes.some((lane) =>
		competitors.every((competitor) =>
			suite.cases.some(
				(entry) =>
					(dimensionString(entry, 'timingLane') ?? 'default') === lane &&
					caseMatchesProfile(entry, requirement, workload, readSource, file, competitor) &&
					isCaseRankingEligible(entry) &&
					caseSatisfiesEvidence(entry, profile),
			),
		),
	)
	return hasSharedLane
		? null
		: `${profile.name} missing-comparable category=${requirement.category} operationProfile=${requirement.operationProfile} workload=${workload}${sourceLabel}${fileLabel} requiredCompetitors=${competitors.map((competitor) => competitor.label).join(',')}`
}

function caseMatchesProfileTuple(
	entry: BenchmarkCaseResult,
	requirement: CaseRequirement,
	workload: string,
	readSource: string | undefined,
	file: string | undefined,
): boolean {
	if (entry.category !== requirement.category) return false
	if ((dimensionString(entry, 'operationProfile') ?? 'default') !== requirement.operationProfile) {
		return false
	}
	if (dimensionString(entry, 'workload') !== workload) return false
	if (readSource !== undefined && dimensionString(entry, 'readSource') !== readSource) return false
	if (file !== undefined && dimensionString(entry, 'file') !== file) return false
	const timingLane = dimensionString(entry, 'timingLane') ?? 'default'
	return requirement.timingLanePrefix ? timingLane.startsWith(requirement.timingLanePrefix) : true
}

function caseMatchesProfile(
	entry: BenchmarkCaseResult,
	requirement: CaseRequirement,
	workload: string,
	readSource: string | undefined,
	file: string | undefined,
	competitor: CompetitorRequirement,
): boolean {
	if (
		!competitor.libraries.includes(dimensionString(entry, 'library') ?? inferLibrary(entry.name))
	) {
		return false
	}
	return caseMatchesProfileTuple(entry, requirement, workload, readSource, file)
}

function caseSatisfiesEvidence(entry: BenchmarkCaseResult, profile: ClaimProfile): boolean {
	const repeat = dimensionNumber(entry, 'repeat')
	if (repeat === undefined || repeat < profile.minRepeat) return false
	if (entry.metrics.sampleCount < profile.minRepeat) return false
	if (profile.requireSamples && (!entry.samples || entry.samples.length < profile.minRepeat)) {
		return false
	}
	return profile.requiredMetrics.every((metric) => typeof entry.metrics[metric] === 'number')
}

function isCaseRankingEligible(entry: BenchmarkCaseResult): boolean {
	const correctnessStatus = dimensionString(entry, 'correctnessStatus') ?? 'not-evaluated'
	return fidelityForStatus(correctnessStatus).tier > 0
}

function isSignificantLeaderLoss(
	leader: CompetitiveScoreboardEntry,
	winner: CompetitiveScoreboardEntry,
): boolean {
	const leaderSamples = leader.metricSamples
	const winnerSamples = winner.metricSamples
	if (!leaderSamples || !winnerSamples || leaderSamples.length < 3 || winnerSamples.length < 3) {
		return true
	}
	return welchsTTest(winnerSamples, leaderSamples).pValue < 0.05
}

function summarizeLibraries(
	groups: readonly CompetitiveScoreboardGroup[],
): CompetitiveScoreboardLibrarySummary[] {
	const summaries = new Map<
		string,
		{
			cases: number
			eligible: number
			disqualified: number
			wins: number
			ranks: number[]
			slowdowns: number[]
		}
	>()
	for (const group of groups) {
		for (const entry of group.entries) {
			const summary = summaries.get(entry.library) ?? {
				cases: 0,
				eligible: 0,
				disqualified: 0,
				wins: 0,
				ranks: [],
				slowdowns: [],
			}
			summary.cases++
			if (entry.rankingEligible && entry.rank !== null) {
				summary.eligible++
				summary.ranks.push(entry.rank)
				if (entry.slowdownVsFastest !== null) summary.slowdowns.push(entry.slowdownVsFastest)
				if (entry.rank === 1) summary.wins++
			} else {
				summary.disqualified++
			}
			summaries.set(entry.library, summary)
		}
	}
	return [...summaries.entries()]
		.map(([library, summary]) => ({
			library,
			cases: summary.cases,
			eligible: summary.eligible,
			disqualified: summary.disqualified,
			wins: summary.wins,
			averageRank: summary.ranks.length > 0 ? mean(summary.ranks) : null,
			geomeanSlowdownVsFastest:
				summary.slowdowns.length > 0 ? geometricMean(summary.slowdowns) : null,
		}))
		.sort((a, b) => {
			if (b.wins !== a.wins) return b.wins - a.wins
			if (b.eligible !== a.eligible) return b.eligible - a.eligible
			if (a.geomeanSlowdownVsFastest !== null && b.geomeanSlowdownVsFastest !== null) {
				return a.geomeanSlowdownVsFastest - b.geomeanSlowdownVsFastest
			}
			return a.library.localeCompare(b.library)
		})
}

function compareMetricValues(left: number, right: number, metric: RankingMetric): number {
	return metric === 'throughputPerSec' ? right - left : left - right
}

function slowdownVsFastest(value: number, winner: number, metric: RankingMetric): number | null {
	if (value <= 0 || winner <= 0) return null
	return metric === 'throughputPerSec' ? winner / value : value / winner
}

function metricSamples(entry: BenchmarkCaseResult, metric: RankingMetric): readonly number[] {
	if (!entry.samples) return []
	return entry.samples
		.map((sample) => {
			switch (metric) {
				case 'medianMs':
				case 'p95Ms':
					return sample.durationMs
				case 'throughputPerSec':
					return sample.throughputPerSec
				case 'rssDeltaBytes':
					return sample.rssDeltaBytes
				case 'retainedRssDeltaBytes':
					return sample.retainedRssDeltaBytes
				case 'rssAfterBytes':
					return sample.rssAfterBytes
				case 'rssAfterGcBytes':
					return sample.rssAfterGcBytes
				case 'peakRssBytes':
					return sample.peakRssBytes
				case 'heapUsedBytes':
					return sample.heapUsedBytes
				case 'heapTotalBytes':
					return sample.heapTotalBytes
				case 'heapAfterGcBytes':
					return sample.heapAfterGcBytes
				case 'heapDeltaBytes':
					return sample.heapDeltaBytes
				default:
					return undefined
			}
		})
		.filter((value): value is number => value !== undefined)
}

function fidelityForStatus(status: string): { tier: number; label: string } {
	switch (status) {
		case 'exact-package-match':
			return { tier: 3, label: 'exact-package' }
		case 'pass':
			return { tier: 2, label: 'semantic-read' }
		case 'semantic-roundtrip-pass':
			return { tier: 2, label: 'semantic-roundtrip' }
		default:
			return { tier: 0, label: 'not-ranked' }
	}
}

function groupSortKey(cases: readonly BenchmarkCaseResult[]): string {
	const first = cases[0]
	if (!first) return ''
	return `${first.category}:${dimensionString(first, 'workload') ?? 'unknown-workload'}:${dimensionString(first, 'readSource') ?? 'unknown-source'}:${dimensionString(first, 'operationProfile') ?? 'default'}:${dimensionString(first, 'timingLane') ?? 'default'}:${dimensionString(first, 'file') ?? first.name}`
}

function dimensionString(entry: BenchmarkCaseResult, key: string): string | undefined {
	const value = entry.dimensions[key]
	return typeof value === 'string' ? value : undefined
}

function dimensionNumber(entry: BenchmarkCaseResult, key: string): number | undefined {
	const value = entry.dimensions[key]
	return typeof value === 'number' ? value : undefined
}

function assertionNumber(entry: BenchmarkCaseResult, key: string): number {
	const value = entry.assertions?.[key]
	return typeof value === 'number' ? value : 0
}

function hasAssertion(entry: BenchmarkCaseResult, key: string): boolean {
	return entry.assertions?.[key] !== undefined
}

function assertionString(entry: BenchmarkCaseResult, key: string): string | undefined {
	const value = entry.assertions?.[key]
	return typeof value === 'string' ? value : undefined
}

function assertionBoolean(entry: BenchmarkCaseResult, key: string): boolean | undefined {
	const value = entry.assertions?.[key]
	return typeof value === 'boolean' ? value : undefined
}

function inferLibrary(name: string): string {
	return name.split(':')[0] ?? name
}

function mean(values: readonly number[]): number {
	return values.reduce((sum, value) => sum + value, 0) / values.length
}

function geometricMean(values: readonly number[]): number {
	return Math.exp(values.reduce((sum, value) => sum + Math.log(value), 0) / values.length)
}

async function main(): Promise<void> {
	const args = process.argv.slice(2)
	const json = args.includes('--json')
	const metric = readRankingMetric(args)
	const category = readOption(args, '--category')
	const leader = readOption(args, '--assert-leader')
	const profile = readClaimProfile(args)
	const profileLeader = readOption(args, '--assert-profile-leader')
	if (profileLeader && !profile) {
		throw new Error('--assert-profile-leader requires --require-profile')
	}
	const inputPath = args.find((arg) => !arg.startsWith('--') && !isOptionValue(args, arg))
	if (!inputPath) {
		throw new Error(
			'Usage: bun run fixtures/benchmarks/competitive-scoreboard.ts <suite.json> [--json] [--metric medianMs|p95Ms|throughputPerSec|peakRssBytes|rssAfterBytes|rssDeltaBytes] [--category read|roundtrip] [--assert-leader ascend] [--require-profile best-js-generated-io|xlsx-read-sota|xlsx-write-sota|upstream-xlsx-sota] [--assert-profile-leader ascend]',
		)
	}
	const suite = JSON.parse(await readFile(inputPath, 'utf-8')) as BenchmarkSuiteResult
	const scoreboard = buildCompetitiveScoreboard(suite, { metric, category })
	const leaderFailures = leader ? assertScoreboardLeader(scoreboard, leader) : []
	const coverageInspection = profile
		? inspectScoreboardCoverage(suite, profile)
		: { failures: [], gaps: [] }
	const profileLeaderFailures =
		profile && profileLeader
			? assertScoreboardProfileLeader(scoreboard, profile, profileLeader)
			: []
	if (json) {
		console.log(
			JSON.stringify(
				{
					...scoreboard,
					leaderFailures,
					coverageFailures: coverageInspection.failures,
					coverageGaps: coverageInspection.gaps,
					profileLeaderFailures,
				},
				null,
				2,
			),
		)
	} else {
		console.log(
			renderScoreboard(
				scoreboard,
				leaderFailures,
				coverageInspection.failures,
				coverageInspection.gaps,
				profileLeaderFailures,
			),
		)
	}
	process.exit(
		leaderFailures.length > 0 ||
			coverageInspection.failures.length > 0 ||
			profileLeaderFailures.length > 0
			? 1
			: 0,
	)
}

export function renderScoreboard(
	scoreboard: CompetitiveScoreboard,
	failures: readonly string[],
	coverageFailures: readonly string[] = [],
	coverageGaps: readonly string[] = [],
	profileLeaderFailures: readonly string[] = [],
): string {
	const lines = [`Competitive scoreboard: ${scoreboard.suite} (${scoreboard.metric})`, '']
	lines.push('Library summary')
	for (const library of scoreboard.libraries) {
		const slowdown =
			library.geomeanSlowdownVsFastest === null
				? 'n/a'
				: `${library.geomeanSlowdownVsFastest.toFixed(2)}x`
		const averageRank = library.averageRank === null ? 'n/a' : library.averageRank.toFixed(2)
		lines.push(
			`  ${library.library}: wins=${library.wins} eligible=${library.eligible}/${library.cases} disqualified=${library.disqualified} avg-rank=${averageRank} geomean-slowdown=${slowdown}`,
		)
	}
	for (const group of scoreboard.groups) {
		lines.push('')
		const fastest =
			group.fastestEligible && group.fastestEligible !== group.winner
				? ` fastest=${group.fastestEligible}`
				: ''
		lines.push(`${renderGroupLabel(group)} winner=${group.winner ?? 'none'}${fastest}`)
		for (const entry of group.entries) {
			const rank = entry.rank === null ? 'DQ' : `#${entry.rank}`
			const speedRank = entry.speedRank === null ? 'speed=DQ' : `speed=#${entry.speedRank}`
			const value = entry.value === null ? 'n/a' : formatMetric(entry.metric, entry.value)
			const slowdown =
				entry.slowdownVsFastest === null ? '' : ` ${entry.slowdownVsFastest.toFixed(2)}x`
			const features = entry.featureObligationsTotal === 0 ? '' : ` features=${entry.featureLabel}`
			const reason = entry.disqualifiedReason ? ` ${entry.disqualifiedReason}` : ''
			lines.push(
				`  ${rank} ${entry.library} ${value}${slowdown} ${speedRank} ${entry.fidelityLabel} ${entry.correctnessStatus}${features}${reason}`,
			)
		}
	}
	if (failures.length > 0) {
		lines.push('')
		lines.push('Leader assertion failures')
		for (const failure of failures) lines.push(`  ${failure}`)
	}
	if (coverageFailures.length > 0) {
		lines.push('')
		lines.push('Coverage assertion failures')
		for (const failure of coverageFailures) lines.push(`  ${failure}`)
	}
	if (coverageGaps.length > 0) {
		lines.push('')
		lines.push('Coverage gaps')
		for (const gap of coverageGaps) lines.push(`  ${gap}`)
	}
	if (profileLeaderFailures.length > 0) {
		lines.push('')
		lines.push('Profile leader assertion failures')
		for (const failure of profileLeaderFailures) lines.push(`  ${failure}`)
	}
	return lines.join('\n')
}

function renderGroupLabel(group: CompetitiveScoreboardGroup): string {
	const parts = [group.category, group.operationProfile]
	if (group.timingLane !== 'default') parts.push(group.timingLane)
	if (group.workload !== 'unknown-workload') parts.push(group.workload)
	if (group.readSource !== 'unknown-source') parts.push(group.readSource)
	if (group.file !== 'unknown-file') parts.push(group.file)
	return parts.join(':')
}

function formatMetric(metric: RankingMetric, value: number): string {
	if (metric === 'throughputPerSec') return formatRate(value)
	if (isByteMetric(metric)) return formatBytes(value)
	return `${value.toFixed(2)} ms`
}

function isByteMetric(metric: RankingMetric): boolean {
	return metric.endsWith('Bytes')
}

function readOption(args: readonly string[], flag: string): string | undefined {
	const index = args.indexOf(flag)
	return index >= 0 ? args[index + 1] : undefined
}

function readRankingMetric(args: readonly string[]): RankingMetric | undefined {
	const value = readOption(args, '--metric')
	if (value === undefined) return undefined
	if (isRankingMetric(value)) return value
	throw new Error(
		`Unknown ranking metric "${value}". Expected one of: ${RANKING_METRICS.join(', ')}`,
	)
}

function readClaimProfile(args: readonly string[]): ClaimProfileName | undefined {
	const value = readOption(args, '--require-profile')
	if (value === undefined) return undefined
	if (isClaimProfile(value)) return value
	throw new Error(
		`Unknown claim profile "${value}". Expected one of: ${Object.keys(CLAIM_PROFILES).join(', ')}`,
	)
}

function isClaimProfile(value: string): value is ClaimProfileName {
	return Object.hasOwn(CLAIM_PROFILES, value)
}

function isRankingMetric(value: string): value is RankingMetric {
	return RANKING_METRICS.includes(value as RankingMetric)
}

function isOptionValue(args: readonly string[], arg: string): boolean {
	return args.some((entry, index) => entry.startsWith('--') && args[index + 1] === arg)
}

if (import.meta.main) {
	await main()
}
