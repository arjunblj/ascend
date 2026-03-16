import { spawnSync } from 'node:child_process'

interface ScenarioResult {
	readonly scenario: string
	readonly chunkBits: number
	readonly throughputPerSec: number
	readonly medianMs: number
}

const scenarios = [
	'read-large-200k',
	'write-dense-40k',
	'recalc-formula-chain',
	'recalc-incremental',
]
const variants = [6, 5] as const

function runScenarioWithChunkBits(scenario: string, chunkBits: number): ScenarioResult {
	const proc = spawnSync(
		'bun',
		[
			'run',
			'fixtures/benchmarks/run.ts',
			'--scenario',
			scenario,
			'--repeat',
			'3',
			'--warmup',
			'1',
			'--json',
		],
		{
			cwd: process.cwd(),
			encoding: 'utf8',
			env: { ...process.env, ASCEND_CHUNK_BITS: String(chunkBits) },
		},
	)
	if (proc.status !== 0) {
		throw new Error(proc.stderr || `Benchmark failed for ${scenario} (CHUNK_BITS=${chunkBits})`)
	}
	const output = proc.stdout.trim()
	if (output.length === 0) throw new Error(`No output for ${scenario} (CHUNK_BITS=${chunkBits})`)
	const parsed = JSON.parse(output) as {
		metrics: { throughputPerSec?: number; medianMs: number }
	}
	return {
		scenario,
		chunkBits,
		throughputPerSec: parsed.metrics.throughputPerSec ?? 0,
		medianMs: parsed.metrics.medianMs,
	}
}

function formatRate(value: number): string {
	if (value >= 1_000_000) return `${(value / 1_000_000).toFixed(2)}M/s`
	if (value >= 1_000) return `${(value / 1_000).toFixed(1)}K/s`
	return `${value.toFixed(1)}/s`
}

function formatDelta(candidate: number, baseline: number): string {
	if (baseline === 0) return 'n/a'
	const pct = ((candidate - baseline) / baseline) * 100
	const sign = pct >= 0 ? '+' : ''
	return `${sign}${pct.toFixed(1)}%`
}

function main(): void {
	const byScenario = new Map<string, ScenarioResult[]>()
	for (const scenario of scenarios) {
		for (const bits of variants) {
			const result = runScenarioWithChunkBits(scenario, bits)
			const existing = byScenario.get(scenario)
			if (existing) existing.push(result)
			else byScenario.set(scenario, [result])
		}
	}

	console.log('Chunk Size A/B Benchmark')
	console.log('='.repeat(88))
	console.log(
		[
			'Scenario'.padEnd(28),
			'64x64 Throughput'.padStart(16),
			'32x32 Throughput'.padStart(16),
			'Throughput Δ'.padStart(14),
			'Median Δ'.padStart(10),
		].join(' '),
	)
	console.log('-'.repeat(88))

	for (const scenario of scenarios) {
		const results = byScenario.get(scenario) ?? []
		const baseline = results.find((r) => r.chunkBits === 6)
		const candidate = results.find((r) => r.chunkBits === 5)
		if (!baseline || !candidate) continue
		const throughputDelta = formatDelta(candidate.throughputPerSec, baseline.throughputPerSec)
		const medianDelta = formatDelta(baseline.medianMs, candidate.medianMs)
		console.log(
			[
				scenario.padEnd(28),
				formatRate(baseline.throughputPerSec).padStart(16),
				formatRate(candidate.throughputPerSec).padStart(16),
				throughputDelta.padStart(14),
				medianDelta.padStart(10),
			].join(' '),
		)
	}
}

main()
