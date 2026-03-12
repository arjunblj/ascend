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
	readonly heapDeltaBytes?: number
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
	readonly heapDeltaBytes?: number
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
		| 'heapDeltaBytes'
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

const DEFAULT_MEDIAN_THRESHOLD = 0.1
const DEFAULT_P95_THRESHOLD = 0.15
const DEFAULT_RSS_THRESHOLD = 0.15
const DEFAULT_THROUGHPUT_THRESHOLD = 0.1
const T_TEST_MIN_SAMPLES = 3
const T_TEST_SIGNIFICANCE_LEVEL = 0.05

export interface CompareSuitesThresholds {
	readonly medianMs?: number
	readonly p95Ms?: number
	readonly throughputPerSec?: number
	readonly rssDeltaBytes?: number
	readonly retainedRssDeltaBytes?: number
	readonly heapDeltaBytes?: number
}

export interface CompareSuitesConfig {
	readonly thresholds?: CompareSuitesThresholds
}

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
		...withDefined('heapDeltaBytes', medianDefined(samples.map((sample) => sample.heapDeltaBytes))),
	}
}

export function compareSuites(
	baseline: BenchmarkSuiteResult,
	candidate: BenchmarkSuiteResult,
	config?: CompareSuitesConfig,
): BenchmarkSuiteComparison {
	const thresholds = config?.thresholds ?? {}
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
		const comparisons = buildComparisons(base, next, thresholds)
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

function buildComparisons(
	base: BenchmarkCaseResult,
	next: BenchmarkCaseResult,
	thresholds: CompareSuitesThresholds,
): BenchmarkMetricComparison[] {
	const baseDurations = base.samples?.map((s) => s.durationMs)
	const nextDurations = next.samples?.map((s) => s.durationMs)
	const durationTTest =
		baseDurations &&
		nextDurations &&
		baseDurations.length >= T_TEST_MIN_SAMPLES &&
		nextDurations.length >= T_TEST_MIN_SAMPLES
			? welchsTTest(baseDurations, nextDurations)
			: null

	const baseThroughput = base.samples
		?.map((s) => s.throughputPerSec)
		.filter((v): v is number => v !== undefined)
	const nextThroughput = next.samples
		?.map((s) => s.throughputPerSec)
		.filter((v): v is number => v !== undefined)
	const throughputTTest =
		baseThroughput &&
		nextThroughput &&
		baseThroughput.length >= T_TEST_MIN_SAMPLES &&
		nextThroughput.length >= T_TEST_MIN_SAMPLES
			? welchsTTest(baseThroughput, nextThroughput)
			: null

	const baseRss = base.samples
		?.map((s) => s.rssDeltaBytes)
		.filter((v): v is number => v !== undefined)
	const nextRss = next.samples
		?.map((s) => s.rssDeltaBytes)
		.filter((v): v is number => v !== undefined)
	const rssTTest =
		baseRss &&
		nextRss &&
		baseRss.length >= T_TEST_MIN_SAMPLES &&
		nextRss.length >= T_TEST_MIN_SAMPLES
			? welchsTTest(baseRss, nextRss)
			: null

	const baseRetained = base.samples
		?.map((s) => s.retainedRssDeltaBytes)
		.filter((v): v is number => v !== undefined)
	const nextRetained = next.samples
		?.map((s) => s.retainedRssDeltaBytes)
		.filter((v): v is number => v !== undefined)
	const retainedTTest =
		baseRetained &&
		nextRetained &&
		baseRetained.length >= T_TEST_MIN_SAMPLES &&
		nextRetained.length >= T_TEST_MIN_SAMPLES
			? welchsTTest(baseRetained, nextRetained)
			: null

	const baseHeapDelta = base.samples
		?.map((s) => s.heapDeltaBytes)
		.filter((v): v is number => v !== undefined)
	const nextHeapDelta = next.samples
		?.map((s) => s.heapDeltaBytes)
		.filter((v): v is number => v !== undefined)
	const heapDeltaTTest =
		baseHeapDelta &&
		nextHeapDelta &&
		baseHeapDelta.length >= T_TEST_MIN_SAMPLES &&
		nextHeapDelta.length >= T_TEST_MIN_SAMPLES
			? welchsTTest(baseHeapDelta, nextHeapDelta)
			: null

	return [
		compareMetric(
			'medianMs',
			'lower-better',
			base.metrics.medianMs,
			next.metrics.medianMs,
			thresholds.medianMs ?? DEFAULT_MEDIAN_THRESHOLD,
			durationTTest?.pValue,
		),
		compareMetric(
			'p95Ms',
			'lower-better',
			base.metrics.p95Ms,
			next.metrics.p95Ms,
			thresholds.p95Ms ?? DEFAULT_P95_THRESHOLD,
			durationTTest?.pValue,
		),
		compareOptionalMetric(
			'throughputPerSec',
			'higher-better',
			base.metrics.throughputPerSec,
			next.metrics.throughputPerSec,
			thresholds.throughputPerSec ?? DEFAULT_THROUGHPUT_THRESHOLD,
			throughputTTest?.pValue,
		),
		compareOptionalMetric(
			'rssDeltaBytes',
			'lower-better',
			base.metrics.rssDeltaBytes,
			next.metrics.rssDeltaBytes,
			thresholds.rssDeltaBytes ?? DEFAULT_RSS_THRESHOLD,
			rssTTest?.pValue,
		),
		compareOptionalMetric(
			'retainedRssDeltaBytes',
			'lower-better',
			base.metrics.retainedRssDeltaBytes,
			next.metrics.retainedRssDeltaBytes,
			thresholds.retainedRssDeltaBytes ?? DEFAULT_RSS_THRESHOLD,
			retainedTTest?.pValue,
		),
		compareOptionalMetric(
			'heapDeltaBytes',
			'lower-better',
			base.metrics.heapDeltaBytes,
			next.metrics.heapDeltaBytes,
			thresholds.heapDeltaBytes ?? DEFAULT_RSS_THRESHOLD,
			heapDeltaTTest?.pValue,
		),
	].filter((entry): entry is BenchmarkMetricComparison => entry !== undefined)
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

export function variance(values: readonly number[]): number {
	if (values.length < 2) return 0
	const m = mean(values)
	return values.reduce((sum, x) => sum + (x - m) ** 2, 0) / (values.length - 1)
}

export interface WelchsTTestResult {
	readonly t: number
	readonly df: number
	readonly pValue: number
}

/**
 * Welch's t-test for comparing two independent samples.
 * Returns t-statistic, degrees of freedom (Welch-Satterthwaite), and p-value.
 * P-value uses approximation: p = exp(-0.717*|t| - 0.416*t²) for two-tailed test.
 */
export function welchsTTest(
	baseline: readonly number[],
	candidate: readonly number[],
): WelchsTTestResult {
	if (baseline.length < 2 || candidate.length < 2) {
		throw new Error('Welch t-test requires at least 2 samples in each group')
	}
	const n1 = baseline.length
	const n2 = candidate.length
	const m1 = mean(baseline)
	const m2 = mean(candidate)
	const v1 = variance(baseline)
	const v2 = variance(candidate)
	const se1Sq = v1 / n1
	const se2Sq = v2 / n2
	const sePooled = Math.sqrt(se1Sq + se2Sq)
	if (sePooled === 0) {
		return { t: 0, df: n1 + n2 - 2, pValue: 1 }
	}
	const t = (m1 - m2) / sePooled
	const df = (se1Sq + se2Sq) ** 2 / (se1Sq ** 2 / (n1 - 1) + se2Sq ** 2 / (n2 - 1))
	const pValue = Math.exp(-0.717 * Math.abs(t) - 0.416 * t * t)
	return { t, df, pValue }
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
	threshold: number,
	pValue: number | undefined,
): BenchmarkMetricComparison {
	const delta = candidate - baseline
	const deltaPct = baseline === 0 ? null : delta / baseline
	const thresholdExceeded =
		direction === 'lower-better'
			? deltaPct !== null && deltaPct > threshold
			: deltaPct !== null && deltaPct < -threshold
	const improvementExceeded =
		direction === 'lower-better'
			? deltaPct !== null && deltaPct < -threshold
			: deltaPct !== null && deltaPct > threshold
	const statisticallySignificant = pValue === undefined || pValue < T_TEST_SIGNIFICANCE_LEVEL
	const isRegression = thresholdExceeded && statisticallySignificant
	const isImprovement = improvementExceeded && statisticallySignificant
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
	threshold: number,
	pValue: number | undefined,
): BenchmarkMetricComparison | undefined {
	if (baseline === undefined || candidate === undefined) return undefined
	return compareMetric(name, direction, baseline, candidate, threshold, pValue)
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
