import { describe, expect, test } from 'bun:test'
import {
	annotateUpstreamCases,
	assertExactUpstreamReplayProfiles,
	buildCompetitiveIoArgs,
	buildIsolatedLibraryFailureSuite,
	isExactUpstreamReplayStatus,
	isKilledRunnerReason,
	selectUpstreamProfiles,
	shouldIsolateLibrariesForProfile,
	splitLibraryList,
	summarizeUpstreamReplayCoverage,
	UPSTREAM_PROFILE_SETS,
	UPSTREAM_PROFILES,
	validateUpstreamProfileSuite,
} from './upstream-profiles.ts'

describe('upstream competitive profiles', () => {
	test('exports third-party benchmark shapes from primary sources', () => {
		expect(UPSTREAM_PROFILES.map((profile) => profile.name)).toEqual([
			'openpyxl-write-1000x50-10pct-text',
			'xlsxwriter-write-memory-200x50-50pct-text',
			'xlsxwriter-write-memory-400x50-50pct-text',
			'xlsxwriter-write-memory-800x50-50pct-text',
			'xlsxwriter-write-memory-1600x50-50pct-text',
			'xlsxwriter-write-memory-3200x50-50pct-text',
			'xlsxwriter-write-memory-6400x50-50pct-text',
			'xlsxwriter-write-memory-12800x50-50pct-text',
			'pyexcelerate-write-values-1000x100',
			'pyexcelerate-write-styles-1000x100',
			'apache-poi-ssperformance-xssf-50000x50',
			'excelize-generation-102400x50-plain-text',
			'closedxml-save-text-1000000x10',
			'closedxml-load-text-1000000x10',
			'closedxml-save-mixed-250000x15',
			'closedxml-load-mixed-250000x15',
			'fastexcel-writer-100000x4',
			'fastexcel-reader-65536',
			'rust-xlsxwriter-write-4000x50-50pct-text',
			'fastxlsx-read-5000x10-matrix',
			'fastxlsx-write-5000x10-matrix',
			'pyopenxlsx-read-1000x20',
			'pyopenxlsx-write-5000x10',
			'pyopenxlsx-bulk-write-50000x20',
			'pyfastexcel-write-50x30',
			'pyfastexcel-write-500x30',
			'pyfastexcel-write-5000x30',
			'pyfastexcel-write-50000x30',
		])
		expect(UPSTREAM_PROFILES.every((profile) => profile.sourceUrl.startsWith('https://'))).toBe(
			true,
		)
		expect(
			UPSTREAM_PROFILES.filter(
				(profile) =>
					profile.name !== 'excelize-generation-102400x50-plain-text' &&
					profile.name !== 'fastxlsx-read-5000x10-matrix',
			).every(
				(profile) =>
					profile.sourceKind === 'published-shape' &&
					profile.replayStatus === 'shape-clone' &&
					profile.timingBoundary.includes('competitive-io') &&
					profile.timingBoundary.includes('not the upstream project native timing harness'),
			),
		).toBe(true)
	})

	test('marks top SOTA target lanes with exact upstream replay provenance', () => {
		const excelize = selectUpstreamProfiles('excelize-generation-102400x50-plain-text')[0]
		const fastxlsx = selectUpstreamProfiles('fastxlsx-read-5000x10-matrix')[0]
		expect(excelize).toBeTruthy()
		expect(fastxlsx).toBeTruthy()
		if (!excelize || !fastxlsx) return

		for (const profile of [excelize, fastxlsx]) {
			expect(profile.sourceKind).toBe('upstream-script')
			expect(profile.replayStatus).toBe('exact-script')
			expect(isExactUpstreamReplayStatus(profile.replayStatus)).toBe(true)
			expect(profile.upstreamRepo).toMatch(/^https:\/\/github\.com\//)
			expect(profile.upstreamCommand).toBeTruthy()
		}
		expect(excelize.timingBoundary).toContain('102400 x 50')
		expect(fastxlsx.timingBoundary).toContain('5000 x 10')
	})

	test('selects one or more named profiles', () => {
		expect(selectUpstreamProfiles(undefined)).toHaveLength(UPSTREAM_PROFILES.length)
		expect(selectUpstreamProfiles('all')).toHaveLength(UPSTREAM_PROFILES.length)
		expect(
			selectUpstreamProfiles(
				'openpyxl-write-1000x50-10pct-text,pyexcelerate-write-values-1000x100',
			).map((profile) => profile.name),
		).toEqual(['openpyxl-write-1000x50-10pct-text', 'pyexcelerate-write-values-1000x100'])
		expect(() => selectUpstreamProfiles('missing-profile')).toThrow('Unsupported --profile')
	})

	test('selects curated profile sets for tight feedback loops', () => {
		expect(UPSTREAM_PROFILE_SETS['write-smoke']).toContain('pyopenxlsx-write-5000x10')
		expect(
			selectUpstreamProfiles(undefined, 'write-memory').map((profile) => profile.name),
		).toEqual([
			'xlsxwriter-write-memory-200x50-50pct-text',
			'xlsxwriter-write-memory-400x50-50pct-text',
			'xlsxwriter-write-memory-800x50-50pct-text',
			'xlsxwriter-write-memory-1600x50-50pct-text',
			'xlsxwriter-write-memory-3200x50-50pct-text',
			'xlsxwriter-write-memory-6400x50-50pct-text',
			'xlsxwriter-write-memory-12800x50-50pct-text',
		])
		expect(selectUpstreamProfiles(undefined, 'read-smoke').map((profile) => profile.name)).toEqual([
			'fastexcel-reader-65536',
			'fastxlsx-read-5000x10-matrix',
			'pyopenxlsx-read-1000x20',
		])
		expect(
			selectUpstreamProfiles('excelize-generation-102400x50-plain-text', 'read-smoke').map(
				(profile) => profile.name,
			),
		).toEqual([
			'excelize-generation-102400x50-plain-text',
			'fastexcel-reader-65536',
			'fastxlsx-read-5000x10-matrix',
			'pyopenxlsx-read-1000x20',
		])
	})

	test('exact replay gate rejects published-shape clones for public benchmark claims', () => {
		const profile = selectUpstreamProfiles('pyexcelerate-write-values-1000x100')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		expect(isExactUpstreamReplayStatus(profile.replayStatus)).toBe(false)
		expect(() => assertExactUpstreamReplayProfiles([profile], 'claim gate')).toThrow(
			'claim gate contains 1 profile(s) that are not exact upstream replays',
		)
		expect(() => assertExactUpstreamReplayProfiles([profile], 'claim gate')).toThrow(
			'pyexcelerate-write-values-1000x100: sourceKind=published-shape replayStatus=shape-clone',
		)
	})

	test('exact replay coverage summary exposes the public-proof denominator', () => {
		const coverage = summarizeUpstreamReplayCoverage(UPSTREAM_PROFILES)

		expect(coverage.total).toBe(UPSTREAM_PROFILES.length)
		expect(coverage.exact).toBe(2)
		expect(coverage.nonExact).toBe(UPSTREAM_PROFILES.length - 2)
		expect(coverage.byStatus['shape-clone']).toBe(UPSTREAM_PROFILES.length - 2)
		expect(coverage.byStatus['exact-script']).toBe(2)
		expect(coverage.nonExactProfiles).toContain('pyexcelerate-write-values-1000x100')
		expect(coverage.nonExactProfiles).not.toContain('excelize-generation-102400x50-plain-text')
		expect(coverage.nonExactProfiles).not.toContain('fastxlsx-read-5000x10-matrix')
	})

	test('splits library lists for isolated heavy profile runs', () => {
		expect(splitLibraryList(undefined)).toEqual([])
		expect(splitLibraryList(' ascend-external-writer, excelize ,, xlsxwriter ')).toEqual([
			'ascend-external-writer',
			'excelize',
			'xlsxwriter',
		])
	})

	test('auto-isolates large profiles and external read profiles', () => {
		const largeProfile = selectUpstreamProfiles('excelize-generation-102400x50-plain-text')[0]
		const smallProfile = selectUpstreamProfiles('pyexcelerate-write-values-1000x100')[0]
		const readProfile = selectUpstreamProfiles('fastexcel-reader-65536')[0]
		expect(largeProfile).toBeTruthy()
		expect(smallProfile).toBeTruthy()
		expect(readProfile).toBeTruthy()
		if (!largeProfile || !smallProfile || !readProfile) return

		expect(
			shouldIsolateLibrariesForProfile({
				profile: largeProfile,
				libraries: 'ascend-external-writer,excelize',
			}),
		).toBe(true)
		expect(
			shouldIsolateLibrariesForProfile({
				profile: smallProfile,
				libraries: 'ascend-external-writer,pyexcelerate',
			}),
		).toBe(false)
		expect(
			shouldIsolateLibrariesForProfile({
				profile: smallProfile,
				libraries: 'ascend-external-writer,pyexcelerate',
				isolationMode: 'always',
			}),
		).toBe(true)
		expect(
			shouldIsolateLibrariesForProfile({
				profile: largeProfile,
				libraries: 'ascend-external-writer',
			}),
		).toBe(false)
		expect(
			shouldIsolateLibrariesForProfile({
				profile: largeProfile,
				libraries: 'ascend-external-writer,excelize',
				isolationMode: 'never',
			}),
		).toBe(false)
		expect(
			shouldIsolateLibrariesForProfile({
				profile: largeProfile,
				libraries: 'ascend-external-writer,excelize',
				competitorOverride: 'external',
				executionScope: 'external-process',
			}),
		).toBe(false)
		expect(
			shouldIsolateLibrariesForProfile({
				profile: readProfile,
				libraries: 'ascend-readxlsx-raw-values-operation-bytes,fastexcel,pyopenxlsx',
				competitorOverride: 'external',
				executionScope: 'external-process',
			}),
		).toBe(true)
	})

	test('records isolated library failures as non-ranking cases', () => {
		const profile = selectUpstreamProfiles('fastexcel-reader-65536')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		const suite = buildIsolatedLibraryFailureSuite({
			profile,
			library: 'pyopenxlsx',
			repeat: 3,
			warmup: 1,
			executionScope: 'external-process',
			errorReason: 'Profile fastexcel-reader-65536 timed out after 120000ms',
		})
		const result = suite.cases[0]
		expect(result?.name).toBe('pyopenxlsx:isolated-runner-error')
		expect(result?.dimensions.library).toBe('pyopenxlsx')
		expect(result?.dimensions.correctnessStatus).toBe('error')
		expect(result?.dimensions.rankingEligible).toBe(false)
		expect(result?.dimensions.errorReason).toContain('timed out')
		expect(result?.metrics.medianMs).toBe(0)
	})

	test('validates child benchmark suites against published profile shape', () => {
		const profile = selectUpstreamProfiles('pyexcelerate-write-values-1000x100')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		const suite = {
			formatVersion: 1,
			suite: 'fixture',
			kind: 'real-workbook',
			generatedAt: '2026-05-09T00:00:00.000Z',
			runtime: { platform: 'test', arch: 'test' },
			git: {},
			cases: [
				{
					name: 'ascend:write-values',
					category: 'write',
					dimensions: {
						library: 'ascend',
						workload: 'dense-values',
						readSource: 'ascend-writer',
						rows: 1000,
						cols: 100,
						cells: 100_000,
						logicalCells: 100_000,
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
		}
		expect(() => validateUpstreamProfileSuite(profile, suite)).not.toThrow()
		expect(() =>
			validateUpstreamProfileSuite(profile, {
				...suite,
				cases: [
					{
						...suite.cases[0],
						dimensions: { ...suite.cases[0].dimensions, rows: 999, cells: 99_900 },
					},
				],
			}),
		).toThrow('upstream shape contract')
	})

	test('recognizes killed external runners for isolated retry decisions', () => {
		expect(isKilledRunnerReason('ascend-external-writer writer exited with code 137')).toBe(true)
		expect(isKilledRunnerReason(new Error('signal: killed while running excelize'))).toBe(true)
		expect(isKilledRunnerReason('external runner failed with SIGKILL')).toBe(true)
		expect(isKilledRunnerReason('external runner failed with exit code 1')).toBe(false)
		expect(isKilledRunnerReason(undefined)).toBe(false)
	})

	test('annotates upstream cases with repro and profile commands', () => {
		const profile = selectUpstreamProfiles('fastexcel-reader-65536')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		const cases = annotateUpstreamCases(
			{
				formatVersion: 1,
				suite: 'child',
				kind: 'real-workbook',
				generatedAt: new Date(0).toISOString(),
				runtime: { platform: 'test', arch: 'arm64' },
				git: {},
				cases: [
					{
						name: 'ascend',
						category: 'read',
						dimensions: {},
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
			{
				repeat: 5,
				warmup: 2,
				library: 'ascend-readxlsx-raw-values-operation-bytes',
				libraryIsolationMode: 'always',
				timeoutMs: 120_000,
			},
		)
		expect(cases[0]?.name).toBe('fastexcel-reader-65536:ascend')
		expect(cases[0]?.reproCommand).toContain('--profile fastexcel-reader-65536')
		expect(cases[0]?.reproCommand).toContain(
			'--libraries ascend-readxlsx-raw-values-operation-bytes',
		)
		expect(cases[0]?.reproCommand).toContain('--isolate-libraries always')
		expect(cases[0]?.reproCommand).toContain('--timeout-ms 120000')
		expect(cases[0]?.profileCommand).toContain('fixtures/benchmarks/profile-bun.ts')
		expect(cases[0]?.profileCommand).toContain(cases[0]?.reproCommand ?? '')
		expect(cases[0]?.dimensions.upstreamSourceKind).toBe('published-shape')
		expect(cases[0]?.dimensions.upstreamReplayStatus).toBe('shape-clone')
		expect(cases[0]?.dimensions.upstreamTimingBoundary).toContain(
			'not the upstream project native timing harness',
		)
	})

	test('builds competitive-io args for read profiles with external runner manifest', () => {
		const profile = selectUpstreamProfiles('fastexcel-reader-65536')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		expect(
			buildCompetitiveIoArgs({
				profile,
				repeat: 5,
				warmup: 2,
				readRunnerManifest: 'readers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'read',
			'--competitor',
			'all',
			'--workload',
			'mixed-50pct-text',
			'--rows',
			'65536',
			'--cols',
			'10',
			'--repeat',
			'5',
			'--warmup',
			'2',
			'--read-source',
			'raw-ooxml',
			'--runner-manifest',
			'readers.json',
		])
	})

	test('builds competitive-io args for write profiles with external writer manifest', () => {
		const profile = selectUpstreamProfiles('fastexcel-writer-100000x4')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		expect(
			buildCompetitiveIoArgs({
				profile,
				repeat: 5,
				warmup: 2,
				writeRunnerManifest: 'writers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'write',
			'--competitor',
			'all',
			'--workload',
			'mixed-50pct-text',
			'--rows',
			'100000',
			'--cols',
			'4',
			'--repeat',
			'5',
			'--warmup',
			'2',
			'--write-runner-manifest',
			'writers.json',
		])
	})

	test('builds competitive-io args for ClosedXML text-only read and write profiles', () => {
		const writeProfile = selectUpstreamProfiles('closedxml-save-text-1000000x10')[0]
		const readProfile = selectUpstreamProfiles('closedxml-load-text-1000000x10')[0]
		expect(writeProfile).toBeTruthy()
		expect(readProfile).toBeTruthy()
		if (!writeProfile || !readProfile) return

		expect(
			buildCompetitiveIoArgs({
				profile: writeProfile,
				repeat: 3,
				warmup: 1,
				competitorOverride: 'external',
				libraries: 'ascend-external-writer,closedxml',
				validationMode: 'final',
				writeRunnerManifest: 'writers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'write',
			'--competitor',
			'external',
			'--workload',
			'plain-text',
			'--rows',
			'1000000',
			'--cols',
			'10',
			'--repeat',
			'3',
			'--warmup',
			'1',
			'--libraries',
			'ascend-external-writer,closedxml',
			'--validation-mode',
			'final',
			'--write-runner-manifest',
			'writers.json',
		])
		expect(
			buildCompetitiveIoArgs({
				profile: readProfile,
				repeat: 3,
				warmup: 1,
				competitorOverride: 'external',
				libraries: 'ascend-readxlsx-raw-values-operation-bytes,closedxml',
				readRunnerManifest: 'readers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'read',
			'--competitor',
			'external',
			'--workload',
			'plain-text',
			'--rows',
			'1000000',
			'--cols',
			'10',
			'--repeat',
			'3',
			'--warmup',
			'1',
			'--libraries',
			'ascend-readxlsx-raw-values-operation-bytes,closedxml',
			'--read-source',
			'raw-ooxml',
			'--runner-manifest',
			'readers.json',
		])
	})

	test('builds competitive-io args for ClosedXML mixed read and write profiles', () => {
		const writeProfile = selectUpstreamProfiles('closedxml-save-mixed-250000x15')[0]
		const readProfile = selectUpstreamProfiles('closedxml-load-mixed-250000x15')[0]
		expect(writeProfile).toBeTruthy()
		expect(readProfile).toBeTruthy()
		if (!writeProfile || !readProfile) return

		expect(
			buildCompetitiveIoArgs({
				profile: writeProfile,
				repeat: 3,
				warmup: 1,
				competitorOverride: 'external',
				libraries: 'ascend-external-writer,closedxml',
				validationMode: 'final',
				writeRunnerManifest: 'writers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'write',
			'--competitor',
			'external',
			'--workload',
			'mixed-closedxml-10text-5number',
			'--rows',
			'250000',
			'--cols',
			'15',
			'--repeat',
			'3',
			'--warmup',
			'1',
			'--libraries',
			'ascend-external-writer,closedxml',
			'--validation-mode',
			'final',
			'--write-runner-manifest',
			'writers.json',
		])
		expect(
			buildCompetitiveIoArgs({
				profile: readProfile,
				repeat: 3,
				warmup: 1,
				competitorOverride: 'external',
				libraries: 'ascend-readxlsx-raw-values-operation-bytes,closedxml',
				readRunnerManifest: 'readers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'read',
			'--competitor',
			'external',
			'--workload',
			'mixed-closedxml-10text-5number',
			'--rows',
			'250000',
			'--cols',
			'15',
			'--repeat',
			'3',
			'--warmup',
			'1',
			'--libraries',
			'ascend-readxlsx-raw-values-operation-bytes,closedxml',
			'--read-source',
			'raw-ooxml',
			'--runner-manifest',
			'readers.json',
		])
	})

	test('can force a profile to external runners while preserving upstream dimensions', () => {
		const profile = selectUpstreamProfiles('excelize-generation-102400x50-plain-text')[0]
		expect(profile).toBeTruthy()
		if (!profile) return

		expect(
			buildCompetitiveIoArgs({
				profile,
				repeat: 5,
				warmup: 1,
				competitorOverride: 'external',
				libraries: 'ascend-external-writer,excelize',
				validationMode: 'final',
				executionScope: 'external-process',
				writeRunnerManifest: 'writers.json',
			}),
		).toEqual([
			'run',
			'fixtures/benchmarks/competitive-io.ts',
			'--json',
			'--category',
			'write',
			'--competitor',
			'external',
			'--workload',
			'plain-text',
			'--rows',
			'102400',
			'--cols',
			'50',
			'--repeat',
			'5',
			'--warmup',
			'1',
			'--libraries',
			'ascend-external-writer,excelize',
			'--validation-mode',
			'final',
			'--execution-scope',
			'external-process',
			'--source-mode',
			'generated-write',
			'--write-runner-manifest',
			'writers.json',
		])
	})
})
