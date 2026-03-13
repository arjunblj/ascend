import type { BenchmarkSuiteResult } from './results.ts'

export interface ThroughputTarget {
	readonly category: string
	readonly metric: string
	readonly minCellsPerSec: number
}

export const throughputTargets: readonly ThroughputTarget[] = [
	{ category: 'read', metric: 'read throughput', minCellsPerSec: 1_000_000 },
	{ category: 'write', metric: 'write throughput', minCellsPerSec: 500_000 },
	{ category: 'calc', metric: 'recalc throughput', minCellsPerSec: 100_000 },
]

export interface TargetCheckResult {
	readonly target: ThroughputTarget
	readonly actualCellsPerSec: number | null
	readonly scenarioName: string | null
	readonly passed: boolean
}

export function checkThroughputTargets(suite: BenchmarkSuiteResult): readonly TargetCheckResult[] {
	return throughputTargets.map((target) => {
		const matching = suite.cases.filter(
			(c) => c.category === target.category && typeof c.metrics.throughputPerSec === 'number',
		)
		if (matching.length === 0) {
			return {
				target,
				actualCellsPerSec: null,
				scenarioName: null,
				passed: false,
			}
		}
		const throughputs = matching
			.map((entry) => entry.metrics.throughputPerSec as number)
			.sort((a, b) => a - b)
		const median = throughputs[Math.floor(throughputs.length / 2)] ?? 0
		const slowest =
			matching.reduce((a, b) =>
				(a.metrics.throughputPerSec ?? Number.POSITIVE_INFINITY) <
				(b.metrics.throughputPerSec ?? Number.POSITIVE_INFINITY)
					? a
					: b,
			) ?? matching[0]
		return {
			target,
			actualCellsPerSec: median,
			scenarioName: `median(${matching.length}) / slowest=${slowest?.name ?? 'n/a'}`,
			passed: median >= target.minCellsPerSec,
		}
	})
}

export function formatTargetResults(results: readonly TargetCheckResult[]): string {
	const lines: string[] = ['Throughput Target Check', '═'.repeat(72)]
	for (const r of results) {
		const status = r.passed ? 'PASS' : 'FAIL'
		const actual = r.actualCellsPerSec !== null ? formatThroughput(r.actualCellsPerSec) : 'no data'
		const required = formatThroughput(r.target.minCellsPerSec)
		lines.push(
			`  [${status}] ${r.target.metric.padEnd(20)} ${actual.padStart(10)} (target: ${required})${r.scenarioName ? `  [${r.scenarioName}]` : ''}`,
		)
	}
	const passed = results.filter((r) => r.passed).length
	lines.push('─'.repeat(72))
	lines.push(`  ${passed}/${results.length} targets met`)
	return lines.join('\n')
}

function formatThroughput(cellsPerSec: number): string {
	if (cellsPerSec >= 1_000_000) return `${(cellsPerSec / 1_000_000).toFixed(2)}M/s`
	if (cellsPerSec >= 1_000) return `${(cellsPerSec / 1_000).toFixed(1)}K/s`
	return `${cellsPerSec.toFixed(1)}/s`
}
