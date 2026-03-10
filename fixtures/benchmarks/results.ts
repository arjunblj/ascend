interface MetricSample {
	readonly durationMs: number
	readonly throughputPerSec?: number
	readonly rssDeltaBytes?: number
	readonly retainedRssDeltaBytes?: number
	readonly rssAfterBytes?: number
	readonly rssAfterGcBytes?: number
	readonly heapUsedBytes?: number
	readonly heapTotalBytes?: number
	readonly heapAfterGcBytes?: number
}

export interface BenchmarkMetricSummary {
	readonly sampleCount: number
	readonly minMs: number
	readonly medianMs: number
	readonly meanMs: number
	readonly p95Ms: number
	readonly maxMs: number
	readonly throughputPerSec?: number
	readonly rssDeltaBytes?: number
	readonly retainedRssDeltaBytes?: number
	readonly rssAfterBytes?: number
	readonly rssAfterGcBytes?: number
	readonly heapUsedBytes?: number
	readonly heapTotalBytes?: number
	readonly heapAfterGcBytes?: number
}

export interface BenchmarkCaseResult {
	readonly name: string
	readonly category: string
	readonly dimensions: Record<string, string | number | boolean>
	readonly metrics: BenchmarkMetricSummary
	readonly samples?: readonly MetricSample[]
	readonly assertions?: Record<string, string | number | boolean | null>
}

export interface BenchmarkRuntimeInfo {
	readonly bunVersion?: string
	readonly nodeVersion?: string
	readonly platform: string
	readonly arch: string
}

export interface BenchmarkGitInfo {
	readonly sha?: string
	readonly branch?: string
}

export interface BenchmarkSuiteResult {
	readonly formatVersion: 1
	readonly suite: string
	readonly kind: 'synthetic' | 'real-workbook'
	readonly generatedAt: string
	readonly runtime: BenchmarkRuntimeInfo
	readonly git: BenchmarkGitInfo
	readonly cases: readonly BenchmarkCaseResult[]
	readonly metadata?: Record<string, unknown>
}

export interface BenchmarkMetricComparison {
	readonly name:
		| 'medianMs'
		| 'p95Ms'
		| 'throughputPerSec'
		| 'rssDeltaBytes'
		| 'retainedRssDeltaBytes'
	readonly direction: 'lower-better' | 'higher-better'
	readonly baseline: number
	readonly candidate: number
	readonly delta: number
	readonly deltaPct: number | null
	readonly status: 'improved' | 'regressed' | 'unchanged'
}

export interface BenchmarkCaseComparison {
	readonly name: string
	readonly status: 'new' | 'missing' | 'improved' | 'regressed' | 'unchanged'
	readonly comparisons: readonly BenchmarkMetricComparison[]
}

export interface BenchmarkSuiteComparison {
	readonly baselineSuite: string
	readonly candidateSuite: string
	readonly cases: readonly BenchmarkCaseComparison[]
	readonly summary: {
		readonly improved: number
		readonly regressed: number
		readonly unchanged: number
		readonly missing: number
		readonly added: number
	}
}

const DURATION_REGRESSION_THRESHOLD = 0.1
const P95_REGRESSION_THRESHOLD = 0.15
const RSS_REGRESSION_THRESHOLD = 0.15
const THROUGHPUT_REGRESSION_THRESHOLD = 0.1

export function createBenchmarkSuite(input: {
	readonly suite: string
	readonly kind: 'synthetic' | 'real-workbook'
	readonly cases: readonly BenchmarkCaseResult[]
	readonly metadata?: Record<string, unknown>
}): BenchmarkSuiteResult {
	return {
		formatVersion: 1,
		suite: input.suite,
		kind: input.kind,
		generatedAt: new Date().toISOString(),
		runtime: getRuntimeInfo(),
		git: getGitInfo(),
		cases: input.cases,
		...(input.metadata ? { metadata: input.metadata } : {}),
	}
}

export function summarizeSamples(samples: readonly MetricSample[]): BenchmarkMetricSummary {
	if (samples.length === 0) {
		throw new Error('Cannot summarize empty benchmark samples')
	}
	const durations = samples.map((sample) => sample.durationMs)
	return {
		sampleCount: samples.length,
		minMs: Math.min(...durations),
		medianMs: median(durations),
		meanMs: mean(durations),
		p95Ms: percentile(durations, 0.95),
		maxMs: Math.max(...durations),
		...withDefined(
			'throughputPerSec',
			medianDefined(samples.map((sample) => sample.throughputPerSec)),
		),
		...withDefined('rssDeltaBytes', medianDefined(samples.map((sample) => sample.rssDeltaBytes))),
		...withDefined(
			'retainedRssDeltaBytes',
			medianDefined(samples.map((sample) => sample.retainedRssDeltaBytes)),
		),
		...withDefined('rssAfterBytes', medianDefined(samples.map((sample) => sample.rssAfterBytes))),
		...withDefined(
			'rssAfterGcBytes',
			medianDefined(samples.map((sample) => sample.rssAfterGcBytes)),
		),
		...withDefined('heapUsedBytes', medianDefined(samples.map((sample) => sample.heapUsedBytes))),
		...withDefined('heapTotalBytes', medianDefined(samples.map((sample) => sample.heapTotalBytes))),
		...withDefined(
			'heapAfterGcBytes',
			medianDefined(samples.map((sample) => sample.heapAfterGcBytes)),
		),
	}
}

export function compareSuites(
	baseline: BenchmarkSuiteResult,
	candidate: BenchmarkSuiteResult,
): BenchmarkSuiteComparison {
	const baselineCases = new Map(baseline.cases.map((entry) => [entry.name, entry]))
	const candidateCases = new Map(candidate.cases.map((entry) => [entry.name, entry]))
	const names = [...new Set([...baselineCases.keys(), ...candidateCases.keys()])].sort((a, b) =>
		a.localeCompare(b),
	)
	const cases: BenchmarkCaseComparison[] = []
	for (const name of names) {
		const base = baselineCases.get(name)
		const next = candidateCases.get(name)
		if (!base && next) {
			cases.push({ name, status: 'new', comparisons: [] })
			continue
		}
		if (base && !next) {
			cases.push({ name, status: 'missing', comparisons: [] })
			continue
		}
		if (!base || !next) continue
		const comparisons = [
			compareMetric('medianMs', 'lower-better', base.metrics.medianMs, next.metrics.medianMs),
			compareMetric('p95Ms', 'lower-better', base.metrics.p95Ms, next.metrics.p95Ms),
			compareOptionalMetric(
				'throughputPerSec',
				'higher-better',
				base.metrics.throughputPerSec,
				next.metrics.throughputPerSec,
			),
			compareOptionalMetric(
				'rssDeltaBytes',
				'lower-better',
				base.metrics.rssDeltaBytes,
				next.metrics.rssDeltaBytes,
			),
			compareOptionalMetric(
				'retainedRssDeltaBytes',
				'lower-better',
				base.metrics.retainedRssDeltaBytes,
				next.metrics.retainedRssDeltaBytes,
			),
		].filter((entry): entry is BenchmarkMetricComparison => entry !== undefined)
		const status = comparisons.some((entry) => entry.status === 'regressed')
			? 'regressed'
			: comparisons.some((entry) => entry.status === 'improved')
				? 'improved'
				: 'unchanged'
		cases.push({ name, status, comparisons })
	}
	return {
		baselineSuite: baseline.suite,
		candidateSuite: candidate.suite,
		cases,
		summary: {
			improved: cases.filter((entry) => entry.status === 'improved').length,
			regressed: cases.filter((entry) => entry.status === 'regressed').length,
			unchanged: cases.filter((entry) => entry.status === 'unchanged').length,
			missing: cases.filter((entry) => entry.status === 'missing').length,
			added: cases.filter((entry) => entry.status === 'new').length,
		},
	}
}

export function formatBytes(bytes: number): string {
	if (bytes < 1024) return `${bytes} B`
	if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
	return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

export function formatRate(rate: number): string {
	if (!Number.isFinite(rate)) return 'n/a'
	if (rate >= 1_000_000) return `${(rate / 1_000_000).toFixed(2)}M/s`
	if (rate >= 1_000) return `${(rate / 1_000).toFixed(1)}K/s`
	return `${rate.toFixed(1)}/s`
}

export function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const mid = Math.floor(sorted.length / 2)
	return sorted.length % 2 === 0
		? ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2
		: (sorted[mid] ?? 0)
}

export function percentile(values: readonly number[], ratio: number): number {
	const sorted = [...values].sort((a, b) => a - b)
	const index = Math.max(0, Math.min(sorted.length - 1, Math.ceil(sorted.length * ratio) - 1))
	return sorted[index] ?? 0
}

export function mean(values: readonly number[]): number {
	return values.reduce((sum, value) => sum + value, 0) / values.length
}

export function medianDefined(values: ReadonlyArray<number | undefined>): number | undefined {
	const defined = values.filter((value): value is number => value !== undefined)
	return defined.length > 0 ? median(defined) : undefined
}

function compareMetric(
	name: BenchmarkMetricComparison['name'],
	direction: BenchmarkMetricComparison['direction'],
	baseline: number,
	candidate: number,
): BenchmarkMetricComparison {
	const delta = candidate - baseline
	const deltaPct = baseline === 0 ? null : delta / baseline
	const threshold =
		name === 'medianMs'
			? DURATION_REGRESSION_THRESHOLD
			: name === 'p95Ms'
				? P95_REGRESSION_THRESHOLD
				: name === 'throughputPerSec'
					? THROUGHPUT_REGRESSION_THRESHOLD
					: RSS_REGRESSION_THRESHOLD
	const isRegression =
		direction === 'lower-better'
			? deltaPct !== null && deltaPct > threshold
			: deltaPct !== null && deltaPct < -threshold
	const isImprovement =
		direction === 'lower-better'
			? deltaPct !== null && deltaPct < -threshold
			: deltaPct !== null && deltaPct > threshold
	return {
		name,
		direction,
		baseline,
		candidate,
		delta,
		deltaPct,
		status: isRegression ? 'regressed' : isImprovement ? 'improved' : 'unchanged',
	}
}

function compareOptionalMetric(
	name: BenchmarkMetricComparison['name'],
	direction: BenchmarkMetricComparison['direction'],
	baseline: number | undefined,
	candidate: number | undefined,
): BenchmarkMetricComparison | undefined {
	if (baseline === undefined || candidate === undefined) return undefined
	return compareMetric(name, direction, baseline, candidate)
}

function getRuntimeInfo(): BenchmarkRuntimeInfo {
	return {
		bunVersion: process.versions.bun,
		nodeVersion: process.versions.node,
		platform: process.platform,
		arch: process.arch,
	}
}

function getGitInfo(): BenchmarkGitInfo {
	return {
		...withDefined('sha', runGit(['rev-parse', 'HEAD'])),
		...withDefined('branch', runGit(['rev-parse', '--abbrev-ref', 'HEAD'])),
	}
}

function runGit(args: string[]): string | undefined {
	try {
		const proc = Bun.spawnSync({
			cmd: ['git', ...args],
			stdout: 'pipe',
			stderr: 'pipe',
		})
		if (proc.exitCode !== 0) return undefined
		const text = new TextDecoder().decode(proc.stdout).trim()
		return text.length > 0 ? text : undefined
	} catch {
		return undefined
	}
}

function withDefined<Key extends string, Value>(
	key: Key,
	value: Value | undefined,
): Partial<Record<Key, Value>> {
	return value === undefined ? {} : ({ [key]: value } as Partial<Record<Key, Value>>)
}
