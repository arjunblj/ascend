import { readFile } from 'node:fs/promises'
import {
	type BenchmarkMetricComparison,
	type BenchmarkSuiteResult,
	compareSuites,
	formatBytes,
	formatRate,
} from './results.ts'

async function main(): Promise<void> {
	const [baselinePath, candidatePath] = process.argv.slice(2)
	if (!baselinePath || !candidatePath) {
		throw new Error(
			'Usage: bun run fixtures/benchmarks/compare.ts <baseline.json> <candidate.json> [--json]',
		)
	}
	const [baseline, candidate] = await Promise.all([
		loadSuite(baselinePath),
		loadSuite(candidatePath),
	])
	const comparison = compareSuites(baseline, candidate)
	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(comparison, null, 2))
		process.exit(comparison.summary.regressed > 0 ? 1 : 0)
	}
	console.log(renderComparison(comparison))
	process.exit(comparison.summary.regressed > 0 ? 1 : 0)
}

async function loadSuite(path: string): Promise<BenchmarkSuiteResult> {
	const raw = await readFile(path, 'utf-8')
	return JSON.parse(raw) as BenchmarkSuiteResult
}

function renderComparison(comparison: ReturnType<typeof compareSuites>): string {
	const lines = [
		`Benchmark compare: ${comparison.baselineSuite} -> ${comparison.candidateSuite}`,
		`regressed=${comparison.summary.regressed} improved=${comparison.summary.improved} unchanged=${comparison.summary.unchanged} missing=${comparison.summary.missing} added=${comparison.summary.added}`,
	]
	for (const entry of comparison.cases) {
		lines.push(``)
		lines.push(`${statusLabel(entry.status)} ${entry.name}`)
		if (entry.comparisons.length === 0) continue
		for (const metric of entry.comparisons) {
			lines.push(`  ${renderMetric(metric)}`)
		}
	}
	return lines.join('\n')
}

function renderMetric(metric: BenchmarkMetricComparison): string {
	const baseline = formatMetricValue(metric.name, metric.baseline)
	const candidate = formatMetricValue(metric.name, metric.candidate)
	const deltaPct =
		metric.deltaPct === null
			? 'n/a'
			: `${metric.deltaPct >= 0 ? '+' : ''}${(metric.deltaPct * 100).toFixed(1)}%`
	return `${statusLabel(metric.status)} ${metric.name}: ${baseline} -> ${candidate} (${deltaPct})`
}

function formatMetricValue(name: BenchmarkMetricComparison['name'], value: number): string {
	switch (name) {
		case 'medianMs':
		case 'p95Ms':
			return `${value.toFixed(2)} ms`
		case 'throughputPerSec':
			return formatRate(value)
		case 'rssDeltaBytes':
		case 'retainedRssDeltaBytes':
			return formatBytes(value)
	}
}

function statusLabel(status: 'new' | 'missing' | 'improved' | 'regressed' | 'unchanged'): string {
	switch (status) {
		case 'improved':
			return 'IMPROVED'
		case 'regressed':
			return 'REGRESSED'
		case 'new':
			return 'NEW'
		case 'missing':
			return 'MISSING'
		case 'unchanged':
			return 'UNCHANGED'
	}
}

await main()
