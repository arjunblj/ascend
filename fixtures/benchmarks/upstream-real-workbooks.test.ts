import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { normalizeExternalRunnerSpecs } from './competitive-real-workbook.ts'
import {
	annotateUpstreamRealCases,
	assertProfileWorkbookReady,
	buildCompetitiveRealWorkbookArgs,
	missingWorkbookMessage,
	selectUpstreamRealProfiles,
	UPSTREAM_REAL_WORKBOOK_PROFILES,
	validateUpstreamProfileSuite,
} from './upstream-real-workbooks.ts'

describe('upstream real workbook profiles', () => {
	test('exports calamine NYC 311 benchmark metadata', () => {
		expect(UPSTREAM_REAL_WORKBOOK_PROFILES).toHaveLength(1)
		const profile = UPSTREAM_REAL_WORKBOOK_PROFILES[0]
		expect(profile.name).toBe('calamine-nyc311-1m')
		expect(profile.sourceKind).toBe('pinned-artifact')
		expect(profile.replayStatus).toBe('exact-artifact')
		expect(profile.timingBoundary).toContain('Competitive real-workbook read lane')
		expect(profile.timingBoundary).toContain('pinned locally materialized XLSX artifact')
		expect(profile.timingBoundary).toContain(
			'not dataset acquisition or CSV-to-XLSX materialization',
		)
		expect(profile.expectedRows).toBe(1_000_001)
		expect(profile.expectedCols).toBe(41)
		expect(profile.expectedRangeCells).toBe(41_000_041)
		expect(profile.expectedNonEmptyCells).toBe(28_056_975)
		expect(profile.expectedFirstUsedRange).toBe('NYC_311_SR_2010-2020-sample-1M!A1:AO1000001')
		expect(profile.expectedXlsxBytes).toBe(249_316_631)
		expect(profile.expectedXlsxSha256).toBe(
			'74a9b50621cf9b0fe8cdb2d4072b5535a2c0e2d83247bb38a37a3b3d809202ea',
		)
		expect(profile.upstreamPublishedXlsxBytes).toBe(186_000_000)
		expect(profile.datasetUrl).toContain('NYC_311_SR_2010-2020-sample-1M.7z')
	})

	test('selects named profiles', () => {
		expect(selectUpstreamRealProfiles(undefined)).toEqual(UPSTREAM_REAL_WORKBOOK_PROFILES)
		expect(selectUpstreamRealProfiles('all')).toEqual(UPSTREAM_REAL_WORKBOOK_PROFILES)
		expect(selectUpstreamRealProfiles('calamine-nyc311-1m')).toEqual(
			UPSTREAM_REAL_WORKBOOK_PROFILES,
		)
		expect(() => selectUpstreamRealProfiles('missing-profile')).toThrow('Unsupported --profile')
	})

	test('builds competitive real workbook args', () => {
		const profile = UPSTREAM_REAL_WORKBOOK_PROFILES[0]
		expect(
			buildCompetitiveRealWorkbookArgs({
				profile,
				repeat: 5,
				warmup: 3,
				runnerManifest: 'readers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-real-workbook.ts',
			'--json',
			'--category',
			'read',
			'--runner-manifest',
			'readers.json',
			'--expected-shape-sidecar',
			'fixtures/benchmarks/upstream-real-workbooks/nyc311-shape.json',
			'--repeat',
			'5',
			'--warmup',
			'3',
			'research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx',
		])
	})

	test('NYC 311 SOTA manifest spans every required reader competitor on one timing lane', () => {
		const manifest = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/nyc311-sota-readers.manifest.json', 'utf-8'),
		) as unknown
		const specs = normalizeExternalRunnerSpecs(manifest)
		const names = new Set(specs.map((spec) => spec.name))
		expect(names).toEqual(
			new Set([
				'ascend-external-values-ordered',
				'rust-calamine',
				'openpyxl-read-only-values',
				'apache-poi',
				'closedxml',
				'polars-calamine',
				'polars-xlsx2csv',
				'polars-openpyxl',
				'excelize',
			]),
		)
		expect(new Set(specs.map((spec) => spec.timingModel))).toEqual(
			new Set(['external-internal-file-path-values-timing']),
		)
		expect(specs.every((spec) => spec.capabilities?.valueOnlyRead === true)).toBe(true)
	})

	test('missing workbook message includes acquisition details', () => {
		const message = missingWorkbookMessage(UPSTREAM_REAL_WORKBOOK_PROFILES[0])
		expect(message).toContain('Missing upstream workbook')
		expect(message).toContain('raw.githubusercontent.com')
		expect(message).toContain(
			'sha256=5c5f876b097ed6b51d52a5309c029ac605e959204cfb64a41f847bdc3ef3165b',
		)
		expect(message).toContain('materialize_nyc311_xlsx.py')
		expect(message).toContain(
			'sha256=74a9b50621cf9b0fe8cdb2d4072b5535a2c0e2d83247bb38a37a3b3d809202ea',
		)
		expect(message).toContain('nonEmptyCells=28056975')
	})

	test('rejects local upstream workbooks with the wrong pinned size or hash', async () => {
		const tempDir = await mkdtemp(join(tmpdir(), 'ascend-upstream-profile-'))
		try {
			const path = join(tempDir, 'fixture.xlsx')
			await writeFile(path, 'wrong workbook')
			await expect(
				assertProfileWorkbookReady({
					...UPSTREAM_REAL_WORKBOOK_PROFILES[0],
					localPath: path,
					expectedXlsxBytes: 999,
				}),
			).rejects.toThrow('Unexpected upstream workbook size')
			await expect(
				assertProfileWorkbookReady({
					...UPSTREAM_REAL_WORKBOOK_PROFILES[0],
					localPath: path,
					expectedXlsxBytes: 14,
					expectedXlsxSha256: '0'.repeat(64),
				}),
			).rejects.toThrow('Unexpected upstream workbook sha256')
		} finally {
			await rm(tempDir, { recursive: true, force: true })
		}
	})

	test('annotates cases with upstream profile dimensions for scoreboard coverage', () => {
		const profile = UPSTREAM_REAL_WORKBOOK_PROFILES[0]
		const cases = annotateUpstreamRealCases(
			{
				formatVersion: 1,
				suite: 'fixture',
				kind: 'real-workbook',
				generatedAt: '2026-05-08T00:00:00.000Z',
				runtime: { platform: 'darwin', arch: 'arm64' },
				git: {},
				cases: [
					{
						name: 'ascend:read-values:NYC_311_SR_2010-2020-sample-1M.xlsx',
						category: 'read',
						dimensions: {
							library: 'ascend',
							file: 'NYC_311_SR_2010-2020-sample-1M.xlsx',
							operationProfile: 'read-values',
							correctnessStatus: 'pass',
							rankingEligible: true,
						},
						metrics: {
							sampleCount: 1,
							minMs: 1,
							medianMs: 1,
							meanMs: 1,
							p95Ms: 1,
							maxMs: 1,
						},
					},
				],
			},
			profile,
		)
		expect(cases[0]?.dimensions.workload).toBe('calamine-nyc311-1m')
		expect(cases[0]?.dimensions.upstreamSourceKind).toBe('pinned-artifact')
		expect(cases[0]?.dimensions.upstreamReplayStatus).toBe('exact-artifact')
		expect(cases[0]?.dimensions.upstreamTimingBoundary).toBe(profile.timingBoundary)
		expect(cases[0]?.dimensions.upstreamExpectedNonEmptyCells).toBe(28_056_975)
		expect(cases[0]?.dimensions.upstreamExpectedXlsxSha256).toBe(
			'74a9b50621cf9b0fe8cdb2d4072b5535a2c0e2d83247bb38a37a3b3d809202ea',
		)
	})

	test('validates benchmark output against the upstream shape contract', () => {
		const profile = UPSTREAM_REAL_WORKBOOK_PROFILES[0]
		const suite = {
			formatVersion: 1,
			suite: 'fixture',
			kind: 'real-workbook',
			generatedAt: '2026-05-08T00:00:00.000Z',
			runtime: { platform: 'darwin', arch: 'arm64' },
			git: {},
			cases: [
				{
					name: 'ascend:read-values:NYC_311_SR_2010-2020-sample-1M.xlsx',
					category: 'read',
					dimensions: {
						library: 'ascend',
						file: 'NYC_311_SR_2010-2020-sample-1M.xlsx',
						operationProfile: 'read-values',
						correctnessStatus: 'pass',
						rankingEligible: true,
						sheets: 1,
						cells: 28_056_975,
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
						expectedFirstUsedRange: 'NYC_311_SR_2010-2020-sample-1M!A1:AO1000001',
					},
				},
			],
		}
		expect(() => validateUpstreamProfileSuite(profile, suite)).not.toThrow()
		expect(() =>
			validateUpstreamProfileSuite(profile, {
				...suite,
				cases: [
					{
						...suite.cases[0],
						dimensions: {
							...suite.cases[0].dimensions,
							correctnessStatus: 'semantic-mismatch',
							rankingEligible: false,
						},
					},
				],
			}),
		).toThrow('upstream shape contract')
		expect(() =>
			validateUpstreamProfileSuite(profile, {
				...suite,
				cases: [
					{
						...suite.cases[0],
						dimensions: { ...suite.cases[0].dimensions, cells: 1 },
					},
				],
			}),
		).toThrow('upstream shape contract')
	})
})
