import { describe, expect, test } from 'bun:test'
import {
	buildCompetitiveIoArgs,
	selectUpstreamProfiles,
	UPSTREAM_PROFILES,
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
			'--write-runner-manifest',
			'writers.json',
		])
	})
})
