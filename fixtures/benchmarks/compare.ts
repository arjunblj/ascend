import { readFile } from 'node:fs/promises'
import {
	type BenchmarkMetricComparison,
	type BenchmarkSuiteResult,
	type CompareSuitesConfig,
	compareSuites,
	formatBytes,
	formatRate,
} from './results.ts'

function parseThresholds(argv: string[]): {
	thresholds?: CompareSuitesConfig['thresholds']
	explain: boolean
	rest: string[]
} {
	const rest: string[] = []
	const thresholds: NonNullable<CompareSuitesConfig['thresholds']> = {}
	let explain = false
	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i]
		if (arg === '--threshold-median' && argv[i + 1]) {
			thresholds.medianMs = Number.parseFloat(argv[++i] ?? '')
		} else if (arg === '--threshold-p95' && argv[i + 1]) {
			thresholds.p95Ms = Number.parseFloat(argv[++i] ?? '')
		} else if (arg === '--threshold-throughput' && argv[i + 1]) {
			thresholds.throughputPerSec = Number.parseFloat(argv[++i] ?? '')
		} else if (arg === '--threshold-rss' && argv[i + 1]) {
			const v = Number.parseFloat(argv[++i] ?? '')
			thresholds.rssDeltaBytes = v
			thresholds.retainedRssDeltaBytes = v
		} else if (arg === '--explain') {
			explain = true
		} else {
			rest.push(arg)
		}
	}
	return {
		thresholds: Object.keys(thresholds).length > 0 ? thresholds : undefined,
		explain,
		rest,
	}
}

async function main(): Promise<void> {
	const { thresholds, explain, rest } = parseThresholds(process.argv.slice(2))
	const [baselinePath, candidatePath] = rest
	if (!baselinePath || !candidatePath) {
		throw new Error(
			'Usage: bun run fixtures/benchmarks/compare.ts <baseline.json> <candidate.json> [--json] [--explain] [--threshold-median N] [--threshold-p95 N] [--threshold-throughput N] [--threshold-rss N]',
		)
	}
	const [baseline, candidate] = await Promise.all([
		loadSuite(baselinePath),
		loadSuite(candidatePath),
	])
	const config: CompareSuitesConfig | undefined = thresholds ? { thresholds } : undefined
	const comparison = compareSuites(baseline, candidate, config)
	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(comparison, null, 2))
		process.exit(comparison.summary.regressed > 0 ? 1 : 0)
	}
	console.log(renderComparison(comparison, { explain }))
	process.exit(comparison.summary.regressed > 0 ? 1 : 0)
}

async function loadSuite(path: string): Promise<BenchmarkSuiteResult> {
	const raw = await readFile(path, 'utf-8')
	return JSON.parse(raw) as BenchmarkSuiteResult
}

function renderComparison(
	comparison: ReturnType<typeof compareSuites>,
	options: { readonly explain: boolean },
): string {
	const lines = [
		`Benchmark compare: ${comparison.baselineSuite} -> ${comparison.candidateSuite}`,
		`regressed=${comparison.summary.regressed} improved=${comparison.summary.improved} unchanged=${comparison.summary.unchanged} missing=${comparison.summary.missing} added=${comparison.summary.added}`,
	]
	for (const entry of comparison.cases) {
		lines.push(``)
		lines.push(`${statusLabel(entry.status)} ${entry.name}`)
		for (const metric of entry.comparisons) {
			lines.push(`  ${renderMetric(metric)}`)
		}
		if (options.explain && (entry.status === 'regressed' || entry.status === 'new')) {
			if (entry.reproCommand) lines.push(`  repro: ${entry.reproCommand}`)
			if (entry.profileCommand) lines.push(`  profile: ${entry.profileCommand}`)
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
		case 'heapDeltaBytes':
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
