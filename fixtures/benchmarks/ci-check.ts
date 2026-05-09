/**
 * Benchmark regression CI check.
 * Runs the smoke benchmark set and fails if any category or smoke scenario misses
 * its throughput floor.
 * Memory validation is enabled by default; use --no-check-memory to disable.
 */
import { join } from 'node:path'
import type { BenchmarkCaseResult, BenchmarkSuiteResult } from './results.ts'
import { checkThroughputTargets, formatTargetResults } from './targets.ts'

/** Memory ceilings (bytes) for key scenarios. Set at 1.5x observed values. */
const MEMORY_CEILING_BYTES: Record<string, number> = {
	'read-full-dense': 15 * 1024 * 1024, // 15 MB (observed median 0)
	'write-dense-40k': 1_286_000, // ~1.5x of median 857314
	'recalc-formula-chain': 44_000_000, // ~1.5x of observed max ~29MB
}

const MEMORY_SCENARIOS = Object.keys(MEMORY_CEILING_BYTES) as readonly string[]

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

async function runSmokeBenchmarks(): Promise<BenchmarkSuiteResult> {
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
	return JSON.parse(stdout) as BenchmarkSuiteResult
}

async function runMemoryScenarios(
	smokeSuite: BenchmarkSuiteResult,
): Promise<Map<string, BenchmarkCaseResult>> {
	const results = new Map<string, BenchmarkCaseResult>()
	for (const r of smokeSuite.cases) {
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
	const smokeSuite = await runSmokeBenchmarks()
	const targetResults = checkThroughputTargets(smokeSuite)
	const memoryResults = checkMemory ? await runMemoryScenarios(smokeSuite) : new Map()
	const memorySummary: Array<{
		name: string
		status: string
		heapDeltaBytes?: number
		memoryCeilingBytes?: number
	}> = []
	let hasMemoryFailure = false

	if (checkMemory) {
		for (const [name, ceiling] of Object.entries(MEMORY_CEILING_BYTES)) {
			const result = memoryResults.get(name)
			const heapDeltaBytes = result?.metrics.heapDeltaBytes ?? 0
			const memoryPass = heapDeltaBytes <= ceiling
			if (!memoryPass) hasMemoryFailure = true

			memorySummary.push({
				name,
				status: memoryPass ? 'PASS' : 'FAIL',
				heapDeltaBytes,
				memoryCeilingBytes: ceiling,
			})
		}
	}

	console.log('Benchmark CI check (smoke throughput targets)')
	console.log(formatTargetResults(targetResults))
	console.log('')
	if (checkMemory) {
		console.log('Memory validation: enabled (use --no-check-memory to disable)\n')
		console.log(`${'Scenario'.padEnd(40)}${'Status'.padEnd(8)}Heap delta / ceiling`)
		console.log('-'.repeat(75))
		for (const s of memorySummary) {
			console.log(
				s.name.padEnd(40) +
					s.status.padEnd(8) +
					`${s.heapDeltaBytes !== undefined ? formatBytes(s.heapDeltaBytes) : 'n/a'} / ${s.memoryCeilingBytes !== undefined ? formatBytes(s.memoryCeilingBytes) : 'n/a'}`,
			)
		}
		console.log('')
	} else {
		console.log('Memory validation: disabled\n')
	}

	const failedTargets = targetResults.filter((entry) => !entry.passed)
	if (failedTargets.length > 0) {
		console.error('One or more throughput targets failed.')
		for (const item of failedTargets) {
			const actual =
				item.actualCellsPerSec === null ? 'no data' : `${item.actualCellsPerSec.toFixed(1)}/s`
			console.error(
				`- ${item.target.metric}: actual ${actual}, required ${item.target.minCellsPerSec.toFixed(1)}/s`,
			)
		}
		process.exit(1)
	}
	if (hasMemoryFailure) {
		console.error('One or more scenarios exceeded the memory ceiling (heapDeltaBytes).')
		for (const item of memorySummary.filter((entry) => entry.status === 'FAIL')) {
			console.error(
				`- ${item.name}: actual ${item.heapDeltaBytes !== undefined ? formatBytes(item.heapDeltaBytes) : 'n/a'}, ceiling ${item.memoryCeilingBytes !== undefined ? formatBytes(item.memoryCeilingBytes) : 'n/a'}`,
			)
		}
		process.exit(1)
	}
	console.log('All scenarios passed.')
}

main().catch((err) => {
	console.error('CI check failed:', err)
	process.exit(1)
})
