/**
 * Benchmark regression CI check.
 * Runs the smoke benchmark set and fails if any scenario is more than 15% slower than baseline.
 * Baseline thresholds are medianMs * 1.15 (15% regression tolerance).
 */
import { join } from 'node:path'

const REGRESSION_THRESHOLD_PCT = 0.15

const BASELINE_MEDIAN_MS: Record<string, number> = {
	'read-full-dense': 99.79,
	'workflow-sdk-edit-cycle': 66.51,
	'workflow-sdk-defined-names-edit-cycle': 39.9,
	'recalc-incremental': 0.34,
	'recalc-if-short-circuit': 133.18,
	'recalc-lookup-exact-incremental': 229.0,
	'recalc-dynamic-spill-churn': 4.98,
	'recalc-criteria-caching': 55.24,
	'recalc-quickselect': 25.65,
	'structural-insert-rows-recalc': 58.6,
	'read-csv-large': 59.32,
}

interface BenchmarkCaseResult {
	readonly name: string
	readonly metrics: { readonly medianMs: number }
}

async function runSmokeBenchmarks(): Promise<BenchmarkCaseResult[]> {
	const runPath = join(import.meta.dir, 'run.ts')
	const proc = Bun.spawn(
		['bun', 'run', runPath, '--set', 'smoke', '--repeat', '2', '--warmup', '1', '--json'],
		{ stdout: 'pipe', stderr: 'pipe', cwd: join(import.meta.dir, '../..') },
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(`Benchmark run failed: ${stderr || 'unknown error'}`)
	}
	const suite = JSON.parse(stdout)
	return suite.cases ?? []
}

async function main(): Promise<void> {
	const results = await runSmokeBenchmarks()
	const maxThresholdMs = (name: string) => {
		const base = BASELINE_MEDIAN_MS[name]
		return base !== undefined ? base * (1 + REGRESSION_THRESHOLD_PCT) : Number.POSITIVE_INFINITY
	}
	const summary: Array<{ name: string; status: string; medianMs: number; thresholdMs: number }> = []
	let hasFailure = false
	for (const r of results) {
		const threshold = maxThresholdMs(r.name)
		const pass = r.metrics.medianMs <= threshold
		if (!pass) hasFailure = true
		summary.push({
			name: r.name,
			status: pass ? 'PASS' : 'FAIL',
			medianMs: r.metrics.medianMs,
			thresholdMs: threshold,
		})
	}
	console.log('Benchmark CI check (smoke set, 15% regression threshold)\n')
	console.log(
		`${'Scenario'.padEnd(40)}${'Status'.padEnd(8)}${'Median(ms)'.padEnd(14)}Threshold(ms)`,
	)
	console.log('-'.repeat(75))
	for (const s of summary) {
		const base = BASELINE_MEDIAN_MS[s.name]
		const pct = base !== undefined ? ` (+${((s.medianMs / base - 1) * 100).toFixed(1)}%)` : ''
		console.log(
			s.name.padEnd(40) +
				s.status.padEnd(8) +
				s.medianMs.toFixed(2).padEnd(14) +
				s.thresholdMs.toFixed(2) +
				pct,
		)
	}
	console.log('')
	if (hasFailure) {
		console.error('One or more scenarios exceeded the 15% regression threshold.')
		process.exit(1)
	}
	console.log('All scenarios passed.')
}

main().catch((err) => {
	console.error('CI check failed:', err)
	process.exit(1)
})
