import type { BenchmarkSuiteResult } from './results.ts'

export interface ThroughputTarget {
	readonly category: string
	readonly metric: string
	readonly minCellsPerSec: number
}

export const throughputTargets: readonly ThroughputTarget[] = [
	{ category: 'read', metric: 'read throughput', minCellsPerSec: 3_000_000 },
	{ category: 'write', metric: 'write throughput', minCellsPerSec: 1_500_000 },
	{ category: 'calc', metric: 'recalc throughput', minCellsPerSec: 500_000 },
]

export interface ScenarioThroughputTarget {
	readonly name: string
	readonly metric: string
	readonly minCellsPerSec: number
}

export const smokeScenarioThroughputTargets: readonly ScenarioThroughputTarget[] = [
	{ name: 'read-full-dense', metric: 'read-full-dense throughput', minCellsPerSec: 800_000 },
	{ name: 'read-values-dense', metric: 'read-values-dense throughput', minCellsPerSec: 2_500_000 },
	{ name: 'write-csv-large', metric: 'write-csv-large throughput', minCellsPerSec: 3_000_000 },
	{
		name: 'workflow-sdk-edit-cycle',
		metric: 'workflow-sdk-edit-cycle throughput',
		minCellsPerSec: 80_000,
	},
	{
		name: 'workflow-sdk-defined-names-edit-cycle',
		metric: 'workflow-sdk-defined-names-edit-cycle throughput',
		minCellsPerSec: 200_000,
	},
	{
		name: 'recalc-incremental',
		metric: 'recalc-incremental throughput',
		minCellsPerSec: 1_000_000,
	},
	{
		name: 'recalc-if-short-circuit',
		metric: 'recalc-if-short-circuit throughput',
		minCellsPerSec: 150_000,
	},
	{
		name: 'recalc-lookup-exact-incremental',
		metric: 'recalc-lookup-exact-incremental throughput',
		minCellsPerSec: 1_000_000,
	},
	{
		name: 'recalc-dynamic-spill-churn',
		metric: 'recalc-dynamic-spill-churn throughput',
		minCellsPerSec: 250_000,
	},
	{
		name: 'recalc-criteria-caching',
		metric: 'recalc-criteria-caching throughput',
		minCellsPerSec: 500_000,
	},
	{ name: 'recalc-quickselect', metric: 'recalc-quickselect throughput', minCellsPerSec: 500_000 },
	{
		name: 'structural-insert-rows-recalc',
		metric: 'structural-insert-rows-recalc throughput',
		minCellsPerSec: 100_000,
	},
	{ name: 'read-csv-large', metric: 'read-csv-large throughput', minCellsPerSec: 2_000_000 },
]

export interface TargetCheckResult {
	readonly target: ThroughputTarget | ScenarioThroughputTarget
	readonly actualCellsPerSec: number | null
	readonly requiredCellsPerSec: number
	readonly scenarioName: string | null
	readonly passed: boolean
	readonly scope: 'category' | 'scenario'
}

export interface ThroughputTargetOptions {
	readonly minRatio?: number
}

export function checkThroughputTargets(
	suite: BenchmarkSuiteResult,
	options: ThroughputTargetOptions = {},
): readonly TargetCheckResult[] {
	const minRatio = options.minRatio ?? 1
	const results = throughputTargets.map((target) => checkCategoryTarget(suite, target, minRatio))
	if (suite.metadata?.set === 'smoke') {
		for (const target of smokeScenarioThroughputTargets) {
			results.push(checkScenarioTarget(suite, target, minRatio))
		}
	}
	return results
}

export function formatTargetResults(results: readonly TargetCheckResult[]): string {
	const lines: string[] = ['Throughput Target Check', '═'.repeat(72)]
	for (const r of results) {
		const status = r.passed ? 'PASS' : 'FAIL'
		const actual = r.actualCellsPerSec !== null ? formatThroughput(r.actualCellsPerSec) : 'no data'
		const required = formatThroughput(r.requiredCellsPerSec)
		lines.push(
			`  [${status}] ${r.target.metric.padEnd(44)} ${actual.padStart(10)} (target: ${required})${r.scenarioName ? `  [${r.scenarioName}]` : ''}`,
		)
	}
	const passed = results.filter((r) => r.passed).length
	lines.push('─'.repeat(72))
	lines.push(`  ${passed}/${results.length} targets met`)
	return lines.join('\n')
}

function checkCategoryTarget(
	suite: BenchmarkSuiteResult,
	target: ThroughputTarget,
	minRatio: number,
): TargetCheckResult {
	const requiredCellsPerSec = target.minCellsPerSec * minRatio
	const matching = suite.cases.filter((c) => c.category === target.category)
	if (matching.length === 0) {
		return {
			target,
			actualCellsPerSec: null,
			requiredCellsPerSec,
			scenarioName: null,
			passed: false,
			scope: 'category',
		}
	}
	const best = matching.reduce((a, b) =>
		(a.metrics.throughputPerSec ?? 0) > (b.metrics.throughputPerSec ?? 0) ? a : b,
	)
	const actual = best.metrics.throughputPerSec ?? 0
	return {
		target,
		actualCellsPerSec: actual,
		requiredCellsPerSec,
		scenarioName: best.name,
		passed: actual >= requiredCellsPerSec,
		scope: 'category',
	}
}

function checkScenarioTarget(
	suite: BenchmarkSuiteResult,
	target: ScenarioThroughputTarget,
	minRatio: number,
): TargetCheckResult {
	const requiredCellsPerSec = target.minCellsPerSec * minRatio
	const result = suite.cases.find((c) => c.name === target.name)
	if (!result) {
		return {
			target,
			actualCellsPerSec: null,
			requiredCellsPerSec,
			scenarioName: target.name,
			passed: false,
			scope: 'scenario',
		}
	}
	const actual = result.metrics.throughputPerSec ?? 0
	return {
		target,
		actualCellsPerSec: actual,
		requiredCellsPerSec,
		scenarioName: result.name,
		passed: actual >= requiredCellsPerSec,
		scope: 'scenario',
	}
}

function formatThroughput(cellsPerSec: number): string {
	if (cellsPerSec >= 1_000_000) return `${(cellsPerSec / 1_000_000).toFixed(2)}M/s`
	if (cellsPerSec >= 1_000) return `${(cellsPerSec / 1_000).toFixed(1)}K/s`
	return `${cellsPerSec.toFixed(1)}/s`
}
