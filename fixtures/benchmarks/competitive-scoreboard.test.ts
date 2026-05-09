import { describe, expect, test } from 'bun:test'
import {
	assertScoreboardCoverage,
	assertScoreboardLeader,
	assertScoreboardProfileLeader,
	buildCompetitiveScoreboard,
	inspectScoreboardCoverage,
	renderScoreboard,
} from './competitive-scoreboard.ts'
import type { BenchmarkSuiteResult } from './results.ts'

describe('buildCompetitiveScoreboard', () => {
	test('ranks only correctness-eligible competitor cases', () => {
		const suite = suiteWithCases([
			caseResult('ascend:read:book.xlsx', 'read', 'ascend', 'book.xlsx', true, 'pass', 10),
			caseResult('sheetjs:read:book.xlsx', 'read', 'sheetjs', 'book.xlsx', true, 'pass', 15),
			caseResult(
				'exceljs:read:book.xlsx',
				'read',
				'exceljs',
				'book.xlsx',
				false,
				'semantic-mismatch',
				5,
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(scoreboard.groups).toHaveLength(1)
		expect(scoreboard.groups[0]?.winner).toBe('ascend')
		expect(scoreboard.groups[0]?.entries.map((entry) => [entry.library, entry.rank])).toEqual([
			['ascend', 1],
			['sheetjs', 2],
			['exceljs', null],
		])
		expect(scoreboard.libraries[0]?.library).toBe('ascend')
		expect(scoreboard.libraries.find((entry) => entry.library === 'exceljs')?.disqualified).toBe(1)
	})

	test('derives ranking eligibility from correctness status', () => {
		const suite = suiteWithCases([
			caseResult('ascend:read:book.xlsx', 'read', 'ascend', 'book.xlsx', true, 'pass', 10),
			caseResult(
				'sheetjs:read:book.xlsx',
				'read',
				'sheetjs',
				'book.xlsx',
				true,
				'semantic-mismatch',
				5,
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(scoreboard.groups[0]?.winner).toBe('ascend')
		expect(scoreboard.groups[0]?.entries.map((entry) => [entry.library, entry.rank])).toEqual([
			['ascend', 1],
			['sheetjs', null],
		])
		expect(scoreboard.libraries.find((entry) => entry.library === 'sheetjs')?.disqualified).toBe(1)
	})

	test('supports leader assertions across grouped workbook workloads', () => {
		const suite = suiteWithCases([
			caseResult('ascend:read:a.xlsx', 'read', 'ascend', 'a.xlsx', true, 'pass', 10),
			caseResult('sheetjs:read:a.xlsx', 'read', 'sheetjs', 'a.xlsx', true, 'pass', 9),
			caseResult('ascend:read:b.xlsx', 'read', 'ascend', 'b.xlsx', true, 'pass', 5),
			caseResult('sheetjs:read:b.xlsx', 'read', 'sheetjs', 'b.xlsx', true, 'pass', 7),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(assertScoreboardLeader(scoreboard, 'ascend')).toEqual([
			'read:default:default:a.xlsx winner=sheetjs expected=ascend',
		])
	})

	test('treats external Ascend adapters as Ascend for leader assertions', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend-external-writer',
				category: 'write',
				workload: 'plain-text',
				repeat: 5,
				medianMs: 10,
				file: 'excelize-generation-102400x50-plain-text',
				timingLane: 'external-internal-generated-plain-text',
			}),
			matrixCase({
				library: 'excelize',
				category: 'write',
				workload: 'plain-text',
				repeat: 5,
				medianMs: 12,
				file: 'excelize-generation-102400x50-plain-text',
				timingLane: 'external-internal-generated-plain-text',
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(scoreboard.groups[0]?.winner).toBe('ascend-external-writer')
		expect(assertScoreboardLeader(scoreboard, 'ascend')).toEqual([])
	})

	test('groups distinct operation profiles separately', () => {
		const suite = suiteWithCases([
			caseResult(
				'ascend:read-formula:book.xlsx',
				'read',
				'ascend',
				'book.xlsx',
				true,
				'pass',
				15,
				undefined,
				'read-formula-preserving',
			),
			caseResult(
				'ascend:read-values:book.xlsx',
				'read',
				'ascend',
				'book.xlsx',
				true,
				'pass',
				10,
				undefined,
				'read-values',
			),
			caseResult(
				'fastexcel:read:book.xlsx',
				'read',
				'fastexcel',
				'book.xlsx',
				true,
				'pass',
				8,
				undefined,
				'read-values',
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(scoreboard.groups.map((group) => [group.operationProfile, group.winner])).toEqual([
			['read-formula-preserving', 'ascend'],
			['read-values', 'fastexcel'],
		])
	})

	test('groups distinct timing lanes separately', () => {
		const suite = suiteWithCases([
			caseResult(
				'ascend:read-values:book.xlsx',
				'read',
				'ascend',
				'book.xlsx',
				true,
				'pass',
				10,
				undefined,
				'read-values',
				'in-process-preloaded-bytes',
			),
			caseResult(
				'fastexcel:read:book.xlsx',
				'read',
				'fastexcel',
				'book.xlsx',
				true,
				'pass',
				8,
				undefined,
				'read-values',
				'external-internal-file-path',
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(
			scoreboard.groups.map((group) => [group.operationProfile, group.timingLane, group.winner]),
		).toEqual([
			['read-values', 'external-internal-file-path', 'fastexcel'],
			['read-values', 'in-process-preloaded-bytes', 'ascend'],
		])
	})

	test('renders generated IO group labels with workload and read source', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'dense-values',
				readSource: 'raw-ooxml',
				repeat: 3,
				timingLane: 'in-process-generated-dense-values',
			}),
			matrixCase({
				library: 'sheetjs',
				category: 'read',
				workload: 'dense-values',
				readSource: 'raw-ooxml',
				repeat: 3,
				timingLane: 'in-process-generated-dense-values',
				medianMs: 3,
			}),
		])
		const output = renderScoreboard(buildCompetitiveScoreboard(suite), [])

		expect(output).toContain(
			'read:read-values:in-process-generated-dense-values:dense-values:raw-ooxml winner=',
		)
		expect(output).not.toContain('read:read-values:unknown-file winner=')
	})

	test('does not fail leader assertions for statistically insignificant sample noise', () => {
		const suite = suiteWithCases([
			caseResult(
				'ascend:read:noisy.xlsx',
				'read',
				'ascend',
				'noisy.xlsx',
				true,
				'pass',
				11,
				[10, 11, 12, 9, 13],
			),
			caseResult(
				'sheetjs:read:noisy.xlsx',
				'read',
				'sheetjs',
				'noisy.xlsx',
				true,
				'pass',
				10,
				[9, 10, 11, 8, 12],
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(scoreboard.groups[0]?.winner).toBe('sheetjs')
		expect(assertScoreboardLeader(scoreboard, 'ascend')).toEqual([])
	})

	test('ranks peak RSS as a lower-better metric', () => {
		const suite = suiteWithCases([
			memoryCase(
				'ascend:read:book.xlsx',
				'ascend',
				24_000_000,
				[24_000_000, 25_000_000, 23_000_000],
			),
			memoryCase(
				'sheetjs:read:book.xlsx',
				'sheetjs',
				48_000_000,
				[48_000_000, 49_000_000, 47_000_000],
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite, { metric: 'peakRssBytes' })
		expect(scoreboard.groups[0]?.winner).toBe('ascend')
		expect(scoreboard.groups[0]?.entries.map((entry) => [entry.library, entry.rank])).toEqual([
			['ascend', 1],
			['sheetjs', 2],
		])
	})

	test('leader assertions compare samples for the selected metric', () => {
		const suite = suiteWithCases([
			memoryCase(
				'ascend:read:noisy-memory.xlsx',
				'ascend',
				105,
				[100, 110, 120, 130, 140],
				[1000, 1000, 1000, 1000, 1000],
			),
			memoryCase(
				'sheetjs:read:noisy-memory.xlsx',
				'sheetjs',
				100,
				[99, 109, 119, 129, 139],
				[1, 1, 1, 1, 1],
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite, { metric: 'peakRssBytes' })
		expect(scoreboard.groups[0]?.winner).toBe('sheetjs')
		expect(assertScoreboardLeader(scoreboard, 'ascend')).toEqual([])
	})

	test('prefers exact package fidelity over faster semantic rewrites', () => {
		const suite = suiteWithCases([
			caseResult(
				'ascend:no-op-roundtrip:book.xlsm',
				'roundtrip',
				'ascend',
				'book.xlsm',
				true,
				'exact-package-match',
				20,
			),
			caseResult(
				'sheetjs:no-op-roundtrip:book.xlsm',
				'roundtrip',
				'sheetjs',
				'book.xlsm',
				true,
				'semantic-roundtrip-pass',
				10,
			),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		const group = scoreboard.groups[0]
		expect(group?.winner).toBe('ascend')
		expect(group?.fastestEligible).toBe('sheetjs')
		expect(group?.entries.map((entry) => [entry.library, entry.rank, entry.speedRank])).toEqual([
			['ascend', 1, 2],
			['sheetjs', 2, 1],
		])
		expect(assertScoreboardLeader(scoreboard, 'ascend')).toEqual([])
	})

	test('surfaces generated write feature obligation coverage in entries', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'write',
				workload: 'feature-rich',
				repeat: 5,
				featureAssertions: {
					commentPartCount: 1,
					vmlDrawingPartCount: 1,
					worksheetHyperlinkCount: 1,
					worksheetDataValidationCount: 1,
					worksheetConditionalFormattingCount: 1,
					definedNameCount: 1,
				},
			}),
			matrixCase({
				library: 'sheetjs',
				category: 'write',
				workload: 'feature-rich',
				repeat: 5,
				correctnessStatus: 'semantic-mismatch',
				featureAssertions: {
					commentPartCount: 1,
					vmlDrawingPartCount: 1,
					worksheetHyperlinkCount: 1,
					definedNameCount: 1,
				},
			}),
			matrixCase({
				library: 'xlsxwriter',
				category: 'write',
				workload: 'table-heavy',
				repeat: 5,
				featureAssertions: { tablePartCount: 1 },
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		const featureGroup = scoreboard.groups.find((group) => group.workload === 'feature-rich')
		const tableGroup = scoreboard.groups.find((group) => group.workload === 'table-heavy')

		expect(
			featureGroup?.entries.map((entry) => [
				entry.library,
				entry.featureObligationsMet,
				entry.featureObligationsTotal,
				entry.featureLabel,
			]),
		).toEqual([
			['ascend', 6, 6, '6/6'],
			['sheetjs', 4, 6, '4/6'],
		])
		expect(tableGroup?.entries[0]?.featureLabel).toBe('1/1')
	})

	test('surfaces generated read feature obligation coverage when adapters report it', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'feature-rich',
				repeat: 5,
				featureAssertions: {
					readCommentCount: 1,
					readHyperlinkCount: 1,
					readDataValidationCount: 1,
					readConditionalFormatCount: 1,
					readDefinedNameCount: 1,
				},
			}),
			matrixCase({
				library: 'sheetjs',
				category: 'read',
				workload: 'feature-rich',
				repeat: 5,
			}),
			matrixCase({
				library: 'exceljs',
				category: 'read',
				workload: 'feature-rich',
				repeat: 5,
				featureAssertions: {
					readCommentCount: 1,
					readHyperlinkCount: 1,
					readDataValidationCount: 4,
					readConditionalFormatCount: 1,
					readDefinedNameCount: 1,
				},
			}),
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'table-heavy',
				repeat: 5,
				featureAssertions: { readTableCount: 1 },
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		const featureGroup = scoreboard.groups.find((group) => group.workload === 'feature-rich')
		const tableGroup = scoreboard.groups.find((group) => group.workload === 'table-heavy')

		expect(
			featureGroup?.entries.map((entry) => [
				entry.library,
				entry.featureObligationsMet,
				entry.featureObligationsTotal,
				entry.featureLabel,
			]),
		).toEqual([
			['ascend', 5, 5, '5/5'],
			['sheetjs', 0, 0, 'n/a'],
			['exceljs', 5, 5, '5/5'],
		])
		expect(tableGroup?.entries[0]?.featureLabel).toBe('1/1')
	})

	test('surfaces real roundtrip feature fingerprint coverage in entries', () => {
		const suite = suiteWithCases([
			roundtripFeatureCase('ascend', {
				roundtripFeatureInventoryHashMatches: true,
			}),
			roundtripFeatureCase('sheetjs', {
				correctnessStatus: 'package-roundtrip-mismatch',
				roundtripFeatureInventoryHashMatches: false,
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)

		expect(
			scoreboard.groups[0]?.entries.map((entry) => [
				entry.library,
				entry.featureObligationsMet,
				entry.featureObligationsTotal,
				entry.featureLabel,
			]),
		).toEqual([
			['ascend', 20, 20, '20/20'],
			['sheetjs', 19, 20, '19/20'],
		])
	})

	test('omits real roundtrip feature coverage for feature-empty workbooks', () => {
		const suite = suiteWithCases([
			roundtripFeatureCase('ascend', {
				expectedFeatureInventoryHash:
					'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
				expectedFeaturePartNamesHash:
					'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)

		expect(scoreboard.groups[0]?.entries[0]?.featureLabel).toBe('n/a')
	})

	test('passes best-js generated IO coverage only when the full matrix has enough evidence', () => {
		const suite = suiteWithCases(
			['dense-values', 'string-heavy', 'sparse-wide'].flatMap((workload) =>
				['read', 'write'].flatMap((category) =>
					(category === 'read' ? ['ascend-writer', 'raw-ooxml'] : ['ascend-writer']).flatMap(
						(readSource) =>
							['ascend', 'sheetjs', 'exceljs'].map((library) =>
								matrixCase({ library, category, workload, repeat: 3, readSource }),
							),
					),
				),
			),
		)
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toEqual([])
	})

	test('coverage profiles require competitors in a shared exact timing lane', () => {
		const suite = suiteWithCases(
			['dense-values', 'string-heavy', 'sparse-wide'].flatMap((workload) =>
				['read', 'write'].flatMap((category) =>
					(category === 'read' ? ['ascend-writer', 'raw-ooxml'] : ['ascend-writer']).flatMap(
						(readSource) =>
							['ascend', 'sheetjs', 'exceljs'].map((library) =>
								matrixCase({
									library,
									category,
									workload,
									repeat: 3,
									readSource,
									timingLane:
										workload === 'dense-values' &&
										category === 'read' &&
										readSource === 'ascend-writer' &&
										library === 'exceljs'
											? 'in-process-generated-dense-values-alt'
											: `in-process-generated-${workload}`,
								}),
							),
					),
				),
			),
		)
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toContain(
			'best-js-generated-io missing-comparable category=read operationProfile=read-values workload=dense-values readSource=ascend-writer requiredCompetitors=Ascend,SheetJS,ExcelJS',
		)
	})

	test('coverage profiles accept a shared lane when extra partial lanes exist', () => {
		const fullMatrix = ['dense-values', 'string-heavy', 'sparse-wide'].flatMap((workload) =>
			['read', 'write'].flatMap((category) =>
				(category === 'read' ? ['ascend-writer', 'raw-ooxml'] : ['ascend-writer']).flatMap(
					(readSource) =>
						['ascend', 'sheetjs', 'exceljs'].map((library) =>
							matrixCase({ library, category, workload, repeat: 3, readSource }),
						),
				),
			),
		)
		const suite = suiteWithCases([
			...fullMatrix,
			matrixCase({
				library: 'exceljs',
				category: 'read',
				workload: 'dense-values',
				repeat: 3,
				readSource: 'ascend-writer',
				timingLane: 'in-process-generated-dense-values-alt',
			}),
		])
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toEqual([])
	})

	test('upstream coverage requires real-workbook competitors in a shared timing lane', () => {
		const suite = suiteWithCases(
			[
				['ascend-external-values', 'external-internal-file-path-alt'],
				['rust-calamine', 'external-internal-file-path'],
				['openpyxl-read-only-values', 'external-internal-file-path'],
				['apache-poi', 'external-internal-file-path'],
				['closedxml', 'external-internal-file-path'],
				['polars-calamine', 'external-internal-file-path'],
				['polars-xlsx2csv', 'external-internal-file-path'],
				['polars-openpyxl', 'external-internal-file-path'],
				['excelize', 'external-internal-file-path'],
			].map(([library, timingLane]) =>
				matrixCase({
					library: library ?? '',
					category: 'read',
					workload: 'calamine-nyc311-1m',
					repeat: 5,
					file: 'NYC_311_SR_2010-2020-sample-1M.xlsx',
					timingLane,
					peakRssBytes: 1024,
				}),
			),
		)
		expect(assertScoreboardCoverage(suite, 'upstream-xlsx-sota')).toContain(
			'upstream-xlsx-sota missing-comparable category=read operationProfile=read-values workload=calamine-nyc311-1m file=NYC_311_SR_2010-2020-sample-1M.xlsx requiredCompetitors=Ascend,Calamine,openpyxl,Apache POI,ClosedXML,Polars calamine,Polars xlsx2csv,Polars openpyxl,Excelize',
		)
	})

	test('coverage profiles report missing competitors and weak evidence', () => {
		const suite = suiteWithCases([
			matrixCase({ library: 'ascend', category: 'read', workload: 'dense-values', repeat: 1 }),
			matrixCase({
				library: 'sheetjs',
				category: 'read',
				workload: 'dense-values',
				repeat: 3,
				correctnessStatus: 'semantic-mismatch',
			}),
		])
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toContain(
			'best-js-generated-io weak-evidence competitor=Ascend category=read operationProfile=read-values workload=dense-values readSource=ascend-writer requiredRepeat=3 requiredMetrics=medianMs,p95Ms,throughputPerSec',
		)
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toContain(
			'best-js-generated-io ineligible competitor=SheetJS category=read operationProfile=read-values workload=dense-values readSource=ascend-writer correctnessStatus=semantic-mismatch',
		)
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toContain(
			'best-js-generated-io missing competitor=ExcelJS category=read operationProfile=read-values workload=dense-values readSource=ascend-writer',
		)
		expect(assertScoreboardCoverage(suite, 'best-js-generated-io')).toContain(
			'best-js-generated-io missing competitor=Ascend category=read operationProfile=read-values workload=dense-values readSource=raw-ooxml',
		)
	})

	test('profile leader assertions ignore non-profile external lanes', () => {
		const suite = suiteWithCases([
			matrixCase({ library: 'ascend', category: 'write', workload: 'dense-values', repeat: 3 }),
			matrixCase({
				library: 'sheetjs',
				category: 'write',
				workload: 'dense-values',
				repeat: 3,
				medianMs: 5,
			}),
			matrixCase({
				library: 'exceljs',
				category: 'write',
				workload: 'dense-values',
				repeat: 3,
				medianMs: 8,
			}),
			matrixCase({
				library: 'xlsxwriter',
				category: 'write',
				workload: 'dense-values',
				repeat: 3,
				medianMs: 0.5,
				timingLane: 'external-internal-generated-dense-values',
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(assertScoreboardLeader(scoreboard, 'ascend')).toEqual([
			'write:write-values:external-internal-generated-dense-values:unknown-file winner=xlsxwriter expected=ascend',
		])
		expect(assertScoreboardProfileLeader(scoreboard, 'best-js-generated-io', 'ascend')).toEqual([])
	})

	test('profile leader assertions catch profile competitor losses', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'dense-values',
				repeat: 3,
				medianMs: 10,
			}),
			matrixCase({
				library: 'sheetjs',
				category: 'read',
				workload: 'dense-values',
				repeat: 3,
				medianMs: 5,
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(assertScoreboardProfileLeader(scoreboard, 'best-js-generated-io', 'ascend')).toEqual([
			'best-js-generated-io:read:read-values:in-process-generated-dense-values:dense-values:ascend-writer winner=sheetjs expected=ascend',
		])
	})

	test('profile leader assertions accept profile aliases for expected library', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend-external-values',
				category: 'read',
				workload: 'dense-values',
				repeat: 5,
				operationProfile: 'read-values',
				timingLane: 'external-internal-generated-dense-values',
				peakRssBytes: 1_000,
				medianMs: 5,
			}),
			matrixCase({
				library: 'fastexcel',
				category: 'read',
				workload: 'dense-values',
				repeat: 5,
				operationProfile: 'read-values',
				timingLane: 'external-internal-generated-dense-values',
				peakRssBytes: 1_000,
				medianMs: 8,
			}),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)
		expect(assertScoreboardProfileLeader(scoreboard, 'xlsx-read-sota', 'ascend')).toEqual([])
	})

	test('xlsx read SOTA profile distinguishes required competitors from capability gaps', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'selected-sheet',
				repeat: 5,
				operationProfile: 'read-selected-values',
				peakRssBytes: 1_000,
			}),
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'metadata-only',
				repeat: 5,
				operationProfile: 'read-metadata-only',
				peakRssBytes: 1_000,
			}),
			matrixCase({
				library: 'sheetjs',
				category: 'read',
				workload: 'metadata-only',
				repeat: 5,
				operationProfile: 'read-metadata-only',
				peakRssBytes: 2_000,
			}),
			matrixCase({
				library: 'openpyxl-metadata-only',
				category: 'read',
				workload: 'metadata-only',
				repeat: 5,
				operationProfile: 'read-metadata-only',
				peakRssBytes: 3_000,
			}),
		])
		const inspection = inspectScoreboardCoverage(suite, 'xlsx-read-sota')

		expect(inspection.failures).not.toContain(
			'xlsx-read-sota missing competitor=SheetJS category=read operationProfile=read-selected-values workload=selected-sheet',
		)
		expect(inspection.gaps).toContain(
			'xlsx-read-sota coverage-gap competitor=SheetJS category=read operationProfile=read-selected-values workload=selected-sheet reason=unsupported-operation',
		)
		expect(inspection.failures).not.toContain(
			'xlsx-read-sota missing competitor=ExcelJS category=read operationProfile=read-metadata-only workload=metadata-only',
		)
		expect(inspection.gaps).toContain(
			'xlsx-read-sota coverage-gap competitor=ExcelJS category=read operationProfile=read-metadata-only workload=metadata-only reason=unsupported-operation',
		)
		expect(inspection.failures).not.toContain(
			'xlsx-read-sota missing competitor=openpyxl category=read operationProfile=read-metadata-only workload=metadata-only',
		)
		expect(inspection.gaps).not.toContain(
			'xlsx-read-sota coverage-gap competitor=openpyxl category=read operationProfile=read-metadata-only workload=metadata-only reason=unsupported-operation',
		)
	})

	test('xlsx read SOTA profile accepts Ascend external runner evidence', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend-external-values',
				category: 'read',
				workload: 'dense-values',
				repeat: 5,
				operationProfile: 'read-values',
				timingLane: 'external-internal-generated-dense-values',
				peakRssBytes: 1_000,
			}),
		])

		expect(assertScoreboardCoverage(suite, 'xlsx-read-sota')).not.toContain(
			'xlsx-read-sota missing competitor=Ascend category=read operationProfile=read-values workload=dense-values',
		)
	})

	test('xlsx read SOTA profile still requires warm workflow competitors', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'read',
				workload: 'warm-workflow',
				repeat: 5,
				operationProfile: 'read-values-warm',
				peakRssBytes: 1_000,
			}),
		])

		expect(assertScoreboardCoverage(suite, 'xlsx-read-sota')).toContain(
			'xlsx-read-sota missing competitor=openpyxl category=read operationProfile=read-values-warm workload=warm-workflow',
		)
	})

	test('xlsx write SOTA profile requires generated write evidence across language leaders', () => {
		const workloads = [
			'dense-values',
			'plain-text',
			'sparse-wide',
			'string-heavy',
			'styles-heavy',
			'formula-heavy',
			'table-heavy',
			'feature-rich',
		]
		const libraries = [
			'ascend',
			'sheetjs',
			'exceljs',
			'xlsxwriter',
			'pyexcelerate',
			'openpyxl',
			'apache-poi',
			'fastexcel-java',
			'closedxml',
			'npoi',
			'rust-xlsxwriter',
			'excelize',
		]
		const suite = suiteWithCases(
			workloads.flatMap((workload) =>
				libraries.map((library) =>
					matrixCase({
						library,
						category: 'write',
						workload,
						repeat: 5,
						peakRssBytes: 1_000,
					}),
				),
			),
		)

		expect(assertScoreboardCoverage(suite, 'xlsx-write-sota')).toEqual([])
	})

	test('xlsx write SOTA profile reports missing non-JS writers', () => {
		const suite = suiteWithCases([
			matrixCase({
				library: 'ascend',
				category: 'write',
				workload: 'dense-values',
				repeat: 5,
				peakRssBytes: 1_000,
			}),
		])

		expect(assertScoreboardCoverage(suite, 'xlsx-write-sota')).toContain(
			'xlsx-write-sota missing competitor=XlsxWriter category=write operationProfile=write-values workload=dense-values',
		)
	})

	test('xlsx roundtrip SOTA profile accepts evaluated correctness losers and requires Ascend leader', () => {
		const suite = suiteWithCases([
			editRoundtripCase('ascend', 'semantic-roundtrip-pass', 3, 3),
			editRoundtripCase('sheetjs', 'package-roundtrip-mismatch', 1, 1),
			editRoundtripCase('exceljs', 'package-roundtrip-mismatch', 5, 5),
			editRoundtripCase('openpyxl', 'package-roundtrip-mismatch', 7, 7),
			editRoundtripCase('excelize', 'feature-roundtrip-mismatch', 2, 2),
		])
		const scoreboard = buildCompetitiveScoreboard(suite)

		expect(assertScoreboardCoverage(suite, 'xlsx-roundtrip-sota')).toEqual([])
		expect(assertScoreboardProfileLeader(scoreboard, 'xlsx-roundtrip-sota', 'ascend')).toEqual([])
		expect(scoreboard.groups[0]?.winner).toBe('ascend')
	})

	test('xlsx roundtrip SOTA profile requires external edited-roundtrip competitors', () => {
		const suite = suiteWithCases([editRoundtripCase('ascend', 'semantic-roundtrip-pass', 3, 3)])

		expect(assertScoreboardCoverage(suite, 'xlsx-roundtrip-sota')).toContain(
			'xlsx-roundtrip-sota missing competitor=openpyxl category=edit-roundtrip operationProfile=edit-roundtrip workload=real-workbook file=styles_formulas.xlsx',
		)
		expect(assertScoreboardCoverage(suite, 'xlsx-roundtrip-sota')).toContain(
			'xlsx-roundtrip-sota missing competitor=Excelize category=edit-roundtrip operationProfile=edit-roundtrip workload=real-workbook file=styles_formulas.xlsx',
		)
	})

	test('upstream SOTA profile requires cross-language and method-level runners on published shapes', () => {
		const suite = suiteWithCases([])
		const failures = assertScoreboardCoverage(suite, 'upstream-xlsx-sota')

		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=pyexcelerate range category=write operationProfile=write-values workload=dense-values file=pyexcelerate-write-values-1000x100',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=pyexcelerate cell category=write operationProfile=write-values workload=styles-heavy file=pyexcelerate-write-styles-1000x100',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=fastexcel Java category=write operationProfile=write-values workload=mixed-50pct-text file=fastexcel-writer-100000x4',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=Excelize category=read operationProfile=read-values workload=mixed-50pct-text readSource=raw-ooxml file=fastexcel-reader-65536',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=Apache POI category=read operationProfile=read-values workload=calamine-nyc311-1m file=NYC_311_SR_2010-2020-sample-1M.xlsx',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=Polars calamine category=read operationProfile=read-values workload=calamine-nyc311-1m file=NYC_311_SR_2010-2020-sample-1M.xlsx',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=Polars xlsx2csv category=read operationProfile=read-values workload=calamine-nyc311-1m file=NYC_311_SR_2010-2020-sample-1M.xlsx',
		)
		expect(failures).toContain(
			'upstream-xlsx-sota missing competitor=Polars openpyxl category=read operationProfile=read-values workload=calamine-nyc311-1m file=NYC_311_SR_2010-2020-sample-1M.xlsx',
		)
	})
})

function suiteWithCases(cases: BenchmarkSuiteResult['cases']): BenchmarkSuiteResult {
	return {
		formatVersion: 1,
		suite: 'test-suite',
		kind: 'real-workbook',
		generatedAt: '2026-01-01T00:00:00.000Z',
		runtime: { platform: 'test', arch: 'test' },
		git: {},
		cases,
	}
}

function caseResult(
	name: string,
	category: string,
	library: string,
	file: string,
	rankingEligible: boolean,
	correctnessStatus: string,
	medianMs: number,
	samples?: readonly number[],
	operationProfile?: string,
	timingLane?: string,
): BenchmarkSuiteResult['cases'][number] {
	return {
		name,
		category,
		dimensions: {
			library,
			file,
			correctnessStatus,
			rankingEligible,
			...(operationProfile ? { operationProfile } : {}),
			...(timingLane ? { timingLane } : {}),
		},
		metrics: {
			sampleCount: 1,
			minMs: medianMs,
			medianMs,
			meanMs: medianMs,
			p95Ms: medianMs,
			maxMs: medianMs,
		},
		...(samples ? { samples: samples.map((durationMs) => ({ durationMs })) } : {}),
	}
}

function memoryCase(
	name: string,
	library: string,
	peakRssBytes: number,
	peakSamples: readonly number[],
	durationSamples: readonly number[] = peakSamples.map((_, index) => index + 1),
): BenchmarkSuiteResult['cases'][number] {
	return {
		name,
		category: 'read',
		dimensions: {
			library,
			file: 'book.xlsx',
			correctnessStatus: 'pass',
			rankingEligible: true,
		},
		metrics: {
			sampleCount: peakSamples.length,
			minMs: Math.min(...durationSamples),
			medianMs: durationSamples[Math.floor(durationSamples.length / 2)] ?? durationSamples[0] ?? 0,
			meanMs:
				durationSamples.reduce((sum, value) => sum + value, 0) /
				Math.max(1, durationSamples.length),
			p95Ms: Math.max(...durationSamples),
			maxMs: Math.max(...durationSamples),
			peakRssBytes,
		},
		samples: peakSamples.map((samplePeakRssBytes, index) => ({
			durationMs: durationSamples[index] ?? 0,
			peakRssBytes: samplePeakRssBytes,
		})),
	}
}

function roundtripFeatureCase(
	library: string,
	overrides: Record<string, string | boolean> = {},
): BenchmarkSuiteResult['cases'][number] {
	const matches = {
		roundtripTablePartCountMatches: true,
		roundtripChartPartCountMatches: true,
		roundtripChartExPartCountMatches: true,
		roundtripDrawingPartCountMatches: true,
		roundtripVmlDrawingPartCountMatches: true,
		roundtripPivotTablePartCountMatches: true,
		roundtripPivotCachePartCountMatches: true,
		roundtripSlicerPartCountMatches: true,
		roundtripCommentPartCountMatches: true,
		roundtripThreadedCommentPartCountMatches: true,
		roundtripMediaPartCountMatches: true,
		roundtripExternalLinkPartCountMatches: true,
		roundtripConnectionPartCountMatches: true,
		roundtripCustomXmlPartCountMatches: true,
		roundtripWorksheetHyperlinkCountMatches: true,
		roundtripWorksheetDataValidationCountMatches: true,
		roundtripWorksheetConditionalFormattingCountMatches: true,
		roundtripDefinedNameCountMatches: true,
		roundtripFeaturePartNamesHashMatches: true,
		roundtripFeatureInventoryHashMatches: true,
	}
	const correctnessStatus =
		typeof overrides.correctnessStatus === 'string'
			? overrides.correctnessStatus
			: 'semantic-roundtrip-pass'
	return {
		name: `${library}:no-op-roundtrip:book.xlsx`,
		category: 'roundtrip',
		dimensions: {
			library,
			file: 'book.xlsx',
			correctnessStatus,
			rankingEligible: correctnessStatus === 'semantic-roundtrip-pass',
		},
		metrics: {
			sampleCount: 1,
			minMs: 1,
			medianMs: 1,
			meanMs: 1,
			p95Ms: 1,
			maxMs: 1,
		},
		assertions: {
			expectedFeaturePartNamesHash: 'feature-parts',
			expectedFeatureInventoryHash: 'feature-inventory',
			...matches,
			...overrides,
		},
	}
}

function editRoundtripCase(
	library: string,
	correctnessStatus: string,
	medianMs: number,
	peakRssBytes: number,
): BenchmarkSuiteResult['cases'][number] {
	const repeat = 3
	const durationSamples = [medianMs - 0.1, medianMs, medianMs + 0.1]
	return {
		name: `${library}:edit-roundtrip:styles_formulas.xlsx`,
		category: 'edit-roundtrip',
		dimensions: {
			library,
			workload: 'real-workbook',
			file: 'styles_formulas.xlsx',
			repeat,
			operationProfile: 'edit-roundtrip',
			timingLane: 'external-internal-file-path-materialization-timing',
			correctnessStatus,
			rankingEligible: correctnessStatus === 'semantic-roundtrip-pass',
		},
		metrics: {
			sampleCount: repeat,
			minMs: Math.min(...durationSamples),
			medianMs,
			meanMs: medianMs,
			p95Ms: Math.max(...durationSamples),
			maxMs: Math.max(...durationSamples),
			throughputPerSec: 1000 / medianMs,
			peakRssBytes,
		},
		samples: durationSamples.map((durationMs) => ({
			durationMs,
			throughputPerSec: 1000 / durationMs,
			peakRssBytes,
		})),
	}
}

function matrixCase(input: {
	readonly library: string
	readonly category: string
	readonly workload: string
	readonly repeat: number
	readonly correctnessStatus?: string
	readonly medianMs?: number
	readonly timingLane?: string
	readonly readSource?: string
	readonly operationProfile?: string
	readonly file?: string
	readonly peakRssBytes?: number
	readonly featureAssertions?: Record<string, number>
}): BenchmarkSuiteResult['cases'][number] {
	const operationProfile =
		input.operationProfile ?? (input.category === 'read' ? 'read-values' : 'write-values')
	const timingLane = input.timingLane ?? `in-process-generated-${input.workload}`
	const medianMs = input.medianMs ?? 2
	const peakRssBytes = input.peakRssBytes
	const midpoint = Math.floor(input.repeat / 2)
	const durationSamples = Array.from(
		{ length: input.repeat },
		(_, index) => medianMs + (index - midpoint) * 0.1,
	)
	return {
		name: `${input.library}:${input.category}:${input.workload}`,
		category: input.category,
		dimensions: {
			library: input.library,
			workload: input.workload,
			readSource: input.readSource ?? 'ascend-writer',
			repeat: input.repeat,
			operationProfile,
			timingLane,
			...(input.file ? { file: input.file } : {}),
			correctnessStatus: input.correctnessStatus ?? 'pass',
			rankingEligible: (input.correctnessStatus ?? 'pass') === 'pass',
		},
		metrics: {
			sampleCount: input.repeat,
			minMs: Math.min(...durationSamples),
			medianMs,
			meanMs: medianMs,
			p95Ms: Math.max(...durationSamples),
			maxMs: Math.max(...durationSamples),
			throughputPerSec: 1000,
			...(peakRssBytes === undefined ? {} : { peakRssBytes }),
		},
		samples: durationSamples.map((durationMs) => ({
			durationMs,
			throughputPerSec: 1000,
			...(peakRssBytes === undefined ? {} : { peakRssBytes }),
		})),
		...(input.featureAssertions ? { assertions: input.featureAssertions } : {}),
	}
}
