import { describe, expect, test } from 'bun:test'
import type { BenchmarkCaseResult, BenchmarkSuiteResult } from './results.ts'
import { checkThroughputTargets } from './targets.ts'

const runtime = { platform: 'test', arch: 'arm64' } as const
const git = {} as const

function benchmarkCase(
	name: string,
	category: string,
	throughputPerSec: number,
): BenchmarkCaseResult {
	return {
		name,
		category,
		dimensions: {},
		metrics: {
			sampleCount: 1,
			minMs: 1,
			medianMs: 1,
			meanMs: 1,
			p95Ms: 1,
			maxMs: 1,
			throughputPerSec,
		},
	}
}

function suite(cases: readonly BenchmarkCaseResult[], set?: string): BenchmarkSuiteResult {
	return {
		formatVersion: 1,
		suite: 'test',
		kind: 'synthetic',
		generatedAt: '2026-05-09T00:00:00.000Z',
		runtime,
		git,
		cases,
		...(set ? { metadata: { set } } : {}),
	}
}

describe('benchmark throughput targets', () => {
	test('category checks use the best scenario in each category for non-smoke suites', () => {
		const results = checkThroughputTargets(
			suite([
				benchmarkCase('slow-read', 'read', 1_000_000),
				benchmarkCase('fast-read', 'read', 3_500_000),
				benchmarkCase('write', 'write', 1_600_000),
				benchmarkCase('calc', 'calc', 600_000),
			]),
		)

		expect(results).toHaveLength(3)
		expect(results.every((result) => result.scope === 'category')).toBe(true)
		expect(results.every((result) => result.passed)).toBe(true)
		expect(results.find((result) => result.target.metric === 'read throughput')?.scenarioName).toBe(
			'fast-read',
		)
	})

	test('smoke suites check every named scenario floor, not just best category throughput', () => {
		const results = checkThroughputTargets(
			suite(
				[
					benchmarkCase('read-values-dense', 'read', 3_500_000),
					benchmarkCase('write-csv-large', 'write', 4_000_000),
					benchmarkCase('recalc-incremental', 'calc', 2_000_000),
					benchmarkCase('recalc-quickselect', 'calc', 100_000),
				],
				'smoke',
			),
		)

		const quickselect = results.find(
			(result) => result.target.metric === 'recalc-quickselect throughput',
		)
		const missing = results.find(
			(result) => result.target.metric === 'recalc-criteria-caching throughput',
		)

		expect(quickselect?.scope).toBe('scenario')
		expect(quickselect?.passed).toBe(false)
		expect(missing?.passed).toBe(false)
		expect(missing?.actualCellsPerSec).toBeNull()
	})

	test('target ratio can relax CI smoke floors without changing strict defaults', () => {
		const strict = checkThroughputTargets(
			suite(
				[
					benchmarkCase('read-values-dense', 'read', 2_900_000),
					benchmarkCase('write-csv-large', 'write', 4_000_000),
					benchmarkCase('recalc-incremental', 'calc', 2_000_000),
					benchmarkCase('recalc-if-short-circuit', 'calc', 120_000),
				],
				'smoke',
			),
		)
		const relaxed = checkThroughputTargets(
			suite(
				[
					benchmarkCase('read-values-dense', 'read', 2_900_000),
					benchmarkCase('write-csv-large', 'write', 4_000_000),
					benchmarkCase('recalc-incremental', 'calc', 2_000_000),
					benchmarkCase('recalc-if-short-circuit', 'calc', 120_000),
				],
				'smoke',
			),
			{ minRatio: 0.75 },
		)

		expect(strict.find((result) => result.target.metric === 'read throughput')?.passed).toBe(false)
		expect(relaxed.find((result) => result.target.metric === 'read throughput')?.passed).toBe(true)
		expect(
			strict.find((result) => result.target.metric === 'recalc-if-short-circuit throughput')
				?.passed,
		).toBe(false)
		expect(
			relaxed.find((result) => result.target.metric === 'recalc-if-short-circuit throughput')
				?.requiredCellsPerSec,
		).toBe(112_500)
	})
})
