/**
 * Benchmark regression CI check.
 * Runs the smoke benchmark set and fails if any scenario is more than 15% slower than baseline.
 * Baseline thresholds are medianMs * 1.15 (15% regression tolerance).
 * Memory validation is enabled by default; use --no-check-memory to disable.
 */
import { join } from 'node:path'

const REGRESSION_THRESHOLD_PCT = 0.15

const BASELINE_MEDIAN_MS: Record<string, number> = {
	'read-full-dense': 11.5,
	'read-values-dense': 10.49,
	'workflow-sdk-edit-cycle': 52.58,
	'workflow-sdk-defined-names-edit-cycle': 24.28,
	'recalc-incremental': 0.53,
	'recalc-if-short-circuit': 133.09,
	'recalc-lookup-exact-incremental': 115.45,
	'recalc-dynamic-spill-churn': 3.86,
	'recalc-criteria-caching': 52.7,
	'recalc-quickselect': 21.16,
	'structural-insert-rows-recalc': 57.19,
	'read-csv-large': 52.61,
}

/** Memory ceilings (bytes) for key scenarios. Set at 1.5x observed values. */
const MEMORY_CEILING_BYTES: Record<string, number> = {
	'read-full-dense': 15 * 1024 * 1024, // 15 MB (observed median 0)
	'write-dense-40k': 1_286_000, // ~1.5x of median 857314
	'recalc-formula-chain': 44_000_000, // ~1.5x of observed max ~29MB
}

const MEMORY_SCENARIOS = Object.keys(MEMORY_CEILING_BYTES) as readonly string[]

interface BenchmarkCaseResult {
	readonly name: string
	readonly metrics: {
		readonly medianMs: number
		readonly heapDeltaBytes?: number
	}
}

async function runScenario(name: string): Promise<BenchmarkCaseResult> {
	const runPath = join(import.meta.dir, 'run.ts')
	const proc = Bun.spawn(
		['bun', 'run', runPath, '--scenario', name, '--repeat', '5', '--warmup', '1', '--json'],
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
	return JSON.parse(stdout) as BenchmarkCaseResult
}

async function runSmokeBenchmarks(): Promise<BenchmarkCaseResult[]> {
	const runPath = join(import.meta.dir, 'run.ts')
	const proc = Bun.spawn(
		['bun', 'run', runPath, '--set', 'smoke', '--repeat', '5', '--warmup', '1', '--json'],
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

async function runMemoryScenarios(): Promise<Map<string, BenchmarkCaseResult>> {
	const results = new Map<string, BenchmarkCaseResult>()
	const smoke = await runSmokeBenchmarks()
	for (const r of smoke) {
		results.set(r.name, r)
	}
	for (const name of MEMORY_SCENARIOS) {
		if (results.has(name)) continue
		const r = await runScenario(name)
		results.set(r.name, r)
	}
	return results
}

function formatBytes(bytes: number): string {
	if (bytes < 1024) return `${bytes} B`
	if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
	return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

async function main(): Promise<void> {
	const checkMemory = !process.argv.includes('--no-check-memory')
	const results = checkMemory
		? [...(await runMemoryScenarios()).values()]
		: await runSmokeBenchmarks()

	const maxThresholdMs = (name: string) => {
		const base = BASELINE_MEDIAN_MS[name]
		return base !== undefined ? base * (1 + REGRESSION_THRESHOLD_PCT) : Number.POSITIVE_INFINITY
	}
	const summary: Array<{
		name: string
		status: string
		medianMs: number
		thresholdMs: number
		memoryStatus?: string
		heapDeltaBytes?: number
		memoryCeilingBytes?: number
	}> = []
	let hasTimingFailure = false
	let hasMemoryFailure = false

	for (const r of results) {
		const threshold = maxThresholdMs(r.name)
		const timingPass = r.metrics.medianMs <= threshold
		if (!timingPass) hasTimingFailure = true

		let memoryStatus: string | undefined
		let heapDeltaBytes: number | undefined
		let memoryCeilingBytes: number | undefined
		if (checkMemory) {
			const ceiling = MEMORY_CEILING_BYTES[r.name]
			if (ceiling !== undefined) {
				heapDeltaBytes = r.metrics.heapDeltaBytes ?? 0
				memoryCeilingBytes = ceiling
				const memoryPass = heapDeltaBytes <= ceiling
				if (!memoryPass) hasMemoryFailure = true
				memoryStatus = memoryPass ? 'PASS' : 'FAIL'
			}
		}

		summary.push({
			name: r.name,
			status: timingPass ? 'PASS' : 'FAIL',
			medianMs: r.metrics.medianMs,
			thresholdMs: threshold,
			memoryStatus,
			heapDeltaBytes,
			memoryCeilingBytes,
		})
	}

	console.log('Benchmark CI check (smoke set, 15% regression threshold)')
	if (checkMemory) {
		console.log('Memory validation: enabled (use --no-check-memory to disable)\n')
	} else {
		console.log('Memory validation: disabled\n')
	}
	const hasMemoryCol = checkMemory && summary.some((s) => s.memoryStatus !== undefined)
	const header = hasMemoryCol
		? `${'Scenario'.padEnd(40)}${'Status'.padEnd(8)}${'Median(ms)'.padEnd(14)}Threshold(ms)  Memory`
		: `${'Scenario'.padEnd(40)}${'Status'.padEnd(8)}${'Median(ms)'.padEnd(14)}Threshold(ms)`
	console.log(header)
	console.log('-'.repeat(hasMemoryCol ? 90 : 75))
	for (const s of summary) {
		const base = BASELINE_MEDIAN_MS[s.name]
		const pct = base !== undefined ? ` (+${((s.medianMs / base - 1) * 100).toFixed(1)}%)` : ''
		const memCol =
			s.memoryStatus !== undefined
				? `  ${s.memoryStatus} ${s.heapDeltaBytes !== undefined ? formatBytes(s.heapDeltaBytes) : 'n/a'}/${s.memoryCeilingBytes !== undefined ? formatBytes(s.memoryCeilingBytes) : 'n/a'}`
				: ''
		console.log(
			s.name.padEnd(40) +
				s.status.padEnd(8) +
				s.medianMs.toFixed(2).padEnd(14) +
				s.thresholdMs.toFixed(2) +
				pct +
				memCol,
		)
	}
	console.log('')
	if (hasTimingFailure) {
		console.error('One or more scenarios exceeded the 15% regression threshold.')
		process.exit(1)
	}
	if (hasMemoryFailure) {
		console.error('One or more scenarios exceeded the memory ceiling (heapDeltaBytes).')
		process.exit(1)
	}
	console.log('All scenarios passed.')
}

main().catch((err) => {
	console.error('CI check failed:', err)
	process.exit(1)
})
