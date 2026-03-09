import { describe, expect, test } from 'bun:test'
import { compareSuites, createBenchmarkSuite, summarizeSamples } from './results.ts'

describe('summarizeSamples', () => {
	test('aggregates duration and throughput metrics', () => {
		const summary = summarizeSamples([
			{ durationMs: 30, throughputPerSec: 1000, rssDeltaBytes: 10 },
			{ durationMs: 10, throughputPerSec: 1200, rssDeltaBytes: 20 },
			{ durationMs: 20, throughputPerSec: 1100, rssDeltaBytes: 30 },
		])
		expect(summary.sampleCount).toBe(3)
		expect(summary.minMs).toBe(10)
		expect(summary.medianMs).toBe(20)
		expect(summary.p95Ms).toBe(30)
		expect(summary.maxMs).toBe(30)
		expect(summary.throughputPerSec).toBe(1100)
		expect(summary.rssDeltaBytes).toBe(20)
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
	})
})
