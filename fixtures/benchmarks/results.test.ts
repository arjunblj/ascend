import { describe, expect, test } from 'bun:test'
import { compareSuites, createBenchmarkSuite, summarizeSamples, welchsTTest } from './results.ts'

describe('summarizeSamples', () => {
	test('aggregates duration and throughput metrics', () => {
		const summary = summarizeSamples([
			{ durationMs: 30, throughputPerSec: 1000, rssDeltaBytes: 10, peakRssBytes: 300 },
			{ durationMs: 10, throughputPerSec: 1200, rssDeltaBytes: 20, peakRssBytes: 100 },
			{ durationMs: 20, throughputPerSec: 1100, rssDeltaBytes: 30, peakRssBytes: 200 },
		])
		expect(summary.sampleCount).toBe(3)
		expect(summary.minMs).toBe(10)
		expect(summary.medianMs).toBe(20)
		expect(summary.stddevMs).toBe(10)
		expect(summary.cvMs).toBe(0.5)
		expect(summary.p95Ms).toBe(30)
		expect(summary.maxMs).toBe(30)
		expect(summary.throughputPerSec).toBe(1100)
		expect(summary.rssDeltaBytes).toBe(20)
		expect(summary.peakRssBytes).toBe(200)
	})
})

describe('welchsTTest', () => {
	test('identical samples yield t=0 and p≈1', () => {
		const result = welchsTTest([1, 2, 3], [1, 2, 3])
		expect(result.t).toBe(0)
		expect(result.pValue).toBe(1)
		expect(result.df).toBeGreaterThan(0)
	})

	test('clearly different samples yield low p-value', () => {
		const result = welchsTTest([1, 2, 3], [4, 5, 6])
		expect(Math.abs(result.t)).toBeGreaterThan(3)
		expect(result.pValue).toBeLessThan(0.01)
	})

	test('requires at least 2 samples per group', () => {
		expect(() => welchsTTest([1], [1, 2, 3])).toThrow()
		expect(() => welchsTTest([1, 2, 3], [1])).toThrow()
	})
})

describe('compareSuites', () => {
	test('flags regressions and improvements with thresholds', () => {
		const baseline = createBenchmarkSuite({
			suite: 'baseline',
			kind: 'synthetic',
			cases: [
				{
					name: 'read-case',
					category: 'read',
					dimensions: { cells: 1000 },
					metrics: {
						sampleCount: 1,
						minMs: 100,
						medianMs: 100,
						meanMs: 100,
						p95Ms: 100,
						maxMs: 100,
						throughputPerSec: 10_000,
						rssDeltaBytes: 1000,
					},
				},
			],
		})
		const candidate = createBenchmarkSuite({
			suite: 'candidate',
			kind: 'synthetic',
			cases: [
				{
					name: 'read-case',
					category: 'read',
					dimensions: { cells: 1000 },
					metrics: {
						sampleCount: 1,
						minMs: 125,
						medianMs: 125,
						meanMs: 125,
						p95Ms: 125,
						maxMs: 125,
						throughputPerSec: 13_000,
						rssDeltaBytes: 800,
					},
					reproCommand: 'bun run bench --scenario read-case --json',
					profileCommand:
						'bun run fixtures/benchmarks/profile-bun.ts -- bun run bench --scenario read-case --json',
				},
			],
		})
		const comparison = compareSuites(baseline, candidate)
		expect(comparison.summary.regressed).toBe(1)
		expect(comparison.summary.improved).toBe(0)
		expect(comparison.cases[0]?.status).toBe('regressed')
		expect(
			comparison.cases[0]?.comparisons.some(
				(entry) => entry.name === 'medianMs' && entry.status === 'regressed',
			),
		).toBe(true)
		expect(
			comparison.cases[0]?.comparisons.some(
				(entry) => entry.name === 'throughputPerSec' && entry.status === 'improved',
			),
		).toBe(true)
		expect(
			comparison.cases[0]?.comparisons.some(
				(entry) => entry.name === 'rssDeltaBytes' && entry.status === 'improved',
			),
		).toBe(true)
		expect(comparison.cases[0]?.reproCommand).toBe('bun run bench --scenario read-case --json')
		expect(comparison.cases[0]?.profileCommand).toContain('profile-bun.ts')
	})

	test('statistically insignificant differences do not flag as regression', () => {
		// High variance: baseline [100,50,150] vs candidate [115,65,165]
		// Median increases 15% (exceeds 10% threshold) but t-test p-value is high
		const baseline = createBenchmarkSuite({
			suite: 'baseline',
			kind: 'synthetic',
			cases: [
				{
					name: 'noisy-case',
					category: 'read',
					dimensions: { cells: 1000 },
					metrics: summarizeSamples([{ durationMs: 100 }, { durationMs: 50 }, { durationMs: 150 }]),
					samples: [{ durationMs: 100 }, { durationMs: 50 }, { durationMs: 150 }],
				},
			],
		})
		const candidate = createBenchmarkSuite({
			suite: 'candidate',
			kind: 'synthetic',
			cases: [
				{
					name: 'noisy-case',
					category: 'read',
					dimensions: { cells: 1000 },
					metrics: summarizeSamples([{ durationMs: 115 }, { durationMs: 65 }, { durationMs: 165 }]),
					samples: [{ durationMs: 115 }, { durationMs: 65 }, { durationMs: 165 }],
				},
			],
		})
		const comparison = compareSuites(baseline, candidate)
		expect(comparison.summary.regressed).toBe(0)
		expect(comparison.cases[0]?.status).not.toBe('regressed')
	})

	test('configurable thresholds override defaults', () => {
		const baseline = createBenchmarkSuite({
			suite: 'baseline',
			kind: 'synthetic',
			cases: [
				{
					name: 'case',
					category: 'read',
					dimensions: {},
					metrics: {
						sampleCount: 1,
						minMs: 100,
						medianMs: 100,
						meanMs: 100,
						p95Ms: 100,
						maxMs: 100,
					},
				},
			],
		})
		const candidate = createBenchmarkSuite({
			suite: 'candidate',
			kind: 'synthetic',
			cases: [
				{
					name: 'case',
					category: 'read',
					dimensions: {},
					metrics: {
						sampleCount: 1,
						minMs: 115,
						medianMs: 115,
						meanMs: 115,
						p95Ms: 115,
						maxMs: 115,
					},
				},
			],
		})
		const strict = compareSuites(baseline, candidate, {
			thresholds: { medianMs: 0.05 },
		})
		const lenient = compareSuites(baseline, candidate, {
			thresholds: { medianMs: 0.25 },
		})
		expect(strict.summary.regressed).toBe(1)
		expect(lenient.summary.regressed).toBe(0)
	})
})
