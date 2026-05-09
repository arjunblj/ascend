/**
 * Upstream real-workbook competitive profiles.
 *
 * These profiles represent published real workbook benchmarks. They do not
 * synthesize data; the source workbook must be acquired into the documented
 * local path before running.
 *
 * Run:
 *   bun run fixtures/benchmarks/upstream-real-workbooks.ts --profile calamine-nyc311-1m --json
 */
import { createHash } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import type { BenchmarkCaseResult, BenchmarkSuiteResult } from './results.ts'
import { createBenchmarkSuite, formatBytes, formatRate } from './results.ts'

type UpstreamRealProfileName = 'calamine-nyc311-1m'

export interface UpstreamRealWorkbookProfile {
	readonly name: UpstreamRealProfileName
	readonly sourceLibrary: string
	readonly sourceBenchmark: string
	readonly sourceUrl: string
	readonly datasetUrl: string
	readonly archiveBytes: number
	readonly archiveSha256: string
	readonly sourceCsvBytes: number
	readonly sourceCsvSha256: string
	readonly localPath: string
	readonly category: 'read'
	readonly worksheet: string
	readonly expectedRows: number
	readonly expectedCols: number
	readonly expectedRangeCells: number
	readonly expectedNonEmptyCells: number
	readonly expectedFirstUsedRange: string
	readonly expectedXlsxBytes: number
	readonly expectedXlsxSha256: string
	readonly expectedShapeSidecar: string
	readonly upstreamPublishedXlsxBytes: number
	readonly materializer: string
	readonly materializerLibrary: string
	readonly materializerLibraryVersion: string
	readonly notes: string
}

export const UPSTREAM_REAL_WORKBOOK_PROFILES = [
	{
		name: 'calamine-nyc311-1m',
		sourceLibrary: 'calamine',
		sourceBenchmark: 'calamine README performance benchmark',
		sourceUrl: 'https://docs.rs/crate/calamine/latest/source/README.md',
		datasetUrl:
			'https://raw.githubusercontent.com/wiki/jqnatividad/qsv/files/NYC_311_SR_2010-2020-sample-1M.7z',
		archiveBytes: 48_111_517,
		archiveSha256: '5c5f876b097ed6b51d52a5309c029ac605e959204cfb64a41f847bdc3ef3165b',
		sourceCsvBytes: 538_951_068,
		sourceCsvSha256: '18f0dd774a6c4b79da3dbf3aa0cd878d374dab132226af2c629d9eef9595061b',
		localPath: 'research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx',
		category: 'read',
		worksheet: 'NYC_311_SR_2010-2020-sample-1M',
		expectedRows: 1_000_001,
		expectedCols: 41,
		expectedRangeCells: 41_000_041,
		expectedNonEmptyCells: 28_056_975,
		expectedFirstUsedRange: 'NYC_311_SR_2010-2020-sample-1M!A1:AO1000001',
		expectedXlsxBytes: 249_316_631,
		expectedXlsxSha256: '74a9b50621cf9b0fe8cdb2d4072b5535a2c0e2d83247bb38a37a3b3d809202ea',
		expectedShapeSidecar: 'fixtures/benchmarks/upstream-real-workbooks/nyc311-shape.json',
		upstreamPublishedXlsxBytes: 186_000_000,
		materializer: 'fixtures/benchmarks/materialize_nyc311_xlsx.py',
		materializerLibrary: 'XlsxWriter',
		materializerLibraryVersion: '3.2.9',
		notes:
			'Published calamine read benchmark shape, materialized from the pinned public CSV archive. The calamine README reports a 186 MB XLSX; the public artifact is a CSV archive, so this profile pins the locally materialized XLSX bytes separately.',
	},
] as const satisfies readonly UpstreamRealWorkbookProfile[]

const DEFAULT_RUNNER_MANIFEST = 'fixtures/benchmarks/runners/ascend-python-readers.manifest.json'

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function readPositiveIntFlag(name: string, fallback: number): number {
	const raw = readFlag(name)
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function readNonNegativeIntFlag(name: string, fallback: number): number {
	const raw = readFlag(name)
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

export function selectUpstreamRealProfiles(
	selection: string | undefined,
): readonly UpstreamRealWorkbookProfile[] {
	if (selection === undefined || selection === 'all') return UPSTREAM_REAL_WORKBOOK_PROFILES
	const names = new Set(
		selection
			.split(',')
			.map((name) => name.trim())
			.filter(Boolean),
	)
	const selected = UPSTREAM_REAL_WORKBOOK_PROFILES.filter((profile) => names.has(profile.name))
	if (selected.length !== names.size) {
		const known = UPSTREAM_REAL_WORKBOOK_PROFILES.map((profile) => profile.name).join(', ')
		throw new Error(`Unsupported --profile value. Expected one or more of: ${known}`)
	}
	return selected
}

export function buildCompetitiveRealWorkbookArgs(input: {
	readonly profile: UpstreamRealWorkbookProfile
	readonly repeat: number
	readonly warmup: number
	readonly runnerManifest?: string
	readonly competitor?: 'all' | 'in-process' | 'external'
}): string[] {
	return [
		'run',
		'fixtures/benchmarks/competitive-real-workbook.ts',
		'--json',
		'--category',
		input.profile.category,
		'--runner-manifest',
		input.runnerManifest ?? DEFAULT_RUNNER_MANIFEST,
		'--expected-shape-sidecar',
		input.profile.expectedShapeSidecar,
		...(input.competitor && input.competitor !== 'all' ? ['--competitor', input.competitor] : []),
		'--repeat',
		String(input.repeat),
		'--warmup',
		String(input.warmup),
		input.profile.localPath,
	]
}

export function missingWorkbookMessage(profile: UpstreamRealWorkbookProfile): string {
	return [
		`Missing upstream workbook for ${profile.name}: ${profile.localPath}`,
		`Acquire ${profile.datasetUrl}`,
		`Expected archive bytes=${profile.archiveBytes} sha256=${profile.archiveSha256}`,
		`Extract NYC_311_SR_2010-2020-sample-1M.csv with bytes=${profile.sourceCsvBytes} sha256=${profile.sourceCsvSha256}`,
		`Materialize with: python3 ${profile.materializer} <csv> ${profile.localPath} --json`,
		`Expected XLSX bytes=${profile.expectedXlsxBytes} sha256=${profile.expectedXlsxSha256}`,
		`Expected worksheet=${profile.worksheet} rows=${profile.expectedRows} cols=${profile.expectedCols} nonEmptyCells=${profile.expectedNonEmptyCells}`,
	].join('\n')
}

async function sha256File(path: string): Promise<string> {
	const digest = createHash('sha256')
	for await (const chunk of createReadStream(path)) {
		digest.update(chunk)
	}
	return digest.digest('hex')
}

export async function assertProfileWorkbookReady(
	profile: UpstreamRealWorkbookProfile,
): Promise<void> {
	const path = resolve(profile.localPath)
	let info: Awaited<ReturnType<typeof stat>>
	try {
		info = await stat(path)
	} catch {
		throw new Error(missingWorkbookMessage(profile))
	}
	if (info.size !== profile.expectedXlsxBytes) {
		throw new Error(
			[
				`Unexpected upstream workbook size for ${profile.name}: ${profile.localPath}`,
				`Actual bytes=${info.size}`,
				`Expected bytes=${profile.expectedXlsxBytes}`,
				missingWorkbookMessage(profile),
			].join('\n'),
		)
	}
	const actualSha256 = await sha256File(path)
	if (actualSha256 !== profile.expectedXlsxSha256) {
		throw new Error(
			[
				`Unexpected upstream workbook sha256 for ${profile.name}: ${profile.localPath}`,
				`Actual sha256=${actualSha256}`,
				`Expected sha256=${profile.expectedXlsxSha256}`,
				missingWorkbookMessage(profile),
			].join('\n'),
		)
	}
}

export function validateUpstreamProfileSuite(
	profile: UpstreamRealWorkbookProfile,
	suite: BenchmarkSuiteResult,
): void {
	if (suite.cases.length === 0) {
		throw new Error(`Profile ${profile.name} produced no benchmark cases`)
	}
	const mismatches: string[] = []
	for (const entry of suite.cases) {
		if (entry.dimensions.file !== profile.localPath.split('/').at(-1)) {
			mismatches.push(`${entry.name}: file=${entry.dimensions.file}`)
		}
		if (entry.dimensions.correctnessStatus !== 'pass') {
			mismatches.push(
				`${entry.name}: correctnessStatus=${String(entry.dimensions.correctnessStatus)}`,
			)
		}
		if (entry.dimensions.rankingEligible !== true) {
			mismatches.push(`${entry.name}: rankingEligible=${String(entry.dimensions.rankingEligible)}`)
		}
		if (!Number.isFinite(entry.metrics.medianMs) || entry.metrics.medianMs <= 0) {
			mismatches.push(`${entry.name}: medianMs=${String(entry.metrics.medianMs)}`)
		}
		if (!Number.isFinite(entry.metrics.sampleCount) || entry.metrics.sampleCount <= 0) {
			mismatches.push(`${entry.name}: sampleCount=${String(entry.metrics.sampleCount)}`)
		}
		if (entry.dimensions.sheets !== 1) {
			mismatches.push(`${entry.name}: sheets=${entry.dimensions.sheets}`)
		}
		if (entry.dimensions.cells !== profile.expectedNonEmptyCells) {
			mismatches.push(`${entry.name}: cells=${entry.dimensions.cells}`)
		}
		const firstUsedRange =
			entry.assertions?.expectedFirstUsedRange ?? entry.assertions?.firstUsedRange
		if (firstUsedRange !== profile.expectedFirstUsedRange) {
			mismatches.push(`${entry.name}: firstUsedRange=${String(firstUsedRange)}`)
		}
	}
	if (mismatches.length > 0) {
		throw new Error(
			[
				`Profile ${profile.name} benchmark output does not match the upstream shape contract`,
				...mismatches,
			].join('\n'),
		)
	}
}

async function runProfile(input: {
	readonly profile: UpstreamRealWorkbookProfile
	readonly repeat: number
	readonly warmup: number
	readonly runnerManifest?: string
	readonly competitor?: 'all' | 'in-process' | 'external'
}): Promise<BenchmarkSuiteResult> {
	await assertProfileWorkbookReady(input.profile)
	const args = buildCompetitiveRealWorkbookArgs(input)
	const proc = Bun.spawn(['bun', ...args], {
		cwd: resolve('.'),
		stdout: 'pipe',
		stderr: 'pipe',
	})
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(
			`Profile ${input.profile.name} failed with exit code ${exitCode}\n${stderr}\n${stdout}`,
		)
	}
	const suite = JSON.parse(stdout) as BenchmarkSuiteResult
	validateUpstreamProfileSuite(input.profile, suite)
	return suite
}

export function annotateUpstreamRealCases(
	suite: BenchmarkSuiteResult,
	profile: UpstreamRealWorkbookProfile,
): readonly BenchmarkCaseResult[] {
	return suite.cases.map((entry) => ({
		...entry,
		name: `${profile.name}:${entry.name}`,
		dimensions: {
			...entry.dimensions,
			workload: profile.name,
			upstreamProfile: profile.name,
			upstreamSourceLibrary: profile.sourceLibrary,
			upstreamSourceBenchmark: profile.sourceBenchmark,
			upstreamSourceUrl: profile.sourceUrl,
			upstreamDatasetUrl: profile.datasetUrl,
			upstreamWorksheet: profile.worksheet,
			upstreamExpectedRows: profile.expectedRows,
			upstreamExpectedCols: profile.expectedCols,
			upstreamExpectedRangeCells: profile.expectedRangeCells,
			upstreamExpectedNonEmptyCells: profile.expectedNonEmptyCells,
			upstreamExpectedFirstUsedRange: profile.expectedFirstUsedRange,
			upstreamExpectedXlsxBytes: profile.expectedXlsxBytes,
			upstreamExpectedXlsxSha256: profile.expectedXlsxSha256,
			upstreamExpectedShapeSidecar: profile.expectedShapeSidecar,
			upstreamPublishedXlsxBytes: profile.upstreamPublishedXlsxBytes,
			upstreamArchiveSha256: profile.archiveSha256,
			upstreamSourceCsvSha256: profile.sourceCsvSha256,
			upstreamMaterializer: profile.materializer,
			upstreamMaterializerLibrary: profile.materializerLibrary,
			upstreamMaterializerLibraryVersion: profile.materializerLibraryVersion,
		},
	}))
}

function renderSummary(suite: BenchmarkSuiteResult): string {
	const headers = ['profile', 'case', 'category', 'median-ms', 'throughput', 'status']
	const rows = suite.cases.map((entry) => [
		String(entry.dimensions.upstreamProfile ?? 'unknown'),
		entry.name.replace(`${String(entry.dimensions.upstreamProfile ?? '')}:`, ''),
		entry.category,
		entry.metrics.medianMs.toFixed(2),
		entry.metrics.throughputPerSec === undefined
			? 'n/a'
			: formatRate(entry.metrics.throughputPerSec),
		String(entry.dimensions.correctnessStatus ?? 'unknown'),
	])
	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')
	return [
		line(headers),
		widths.map((width) => '-'.repeat(width)).join('--'),
		...rows.map(line),
	].join('\n')
}

async function main(): Promise<void> {
	const profiles = selectUpstreamRealProfiles(readFlag('--profile'))
	const repeat = readPositiveIntFlag('--repeat', 3)
	const warmup = readNonNegativeIntFlag('--warmup', 1)
	const runnerManifest = readFlag('--runner-manifest') ?? DEFAULT_RUNNER_MANIFEST
	const competitor = readFlag('--competitor') as 'all' | 'in-process' | 'external' | undefined
	const cases: BenchmarkCaseResult[] = []
	const childSuites: Array<Record<string, unknown>> = []
	for (const profile of profiles) {
		const suite = await runProfile({ profile, repeat, warmup, runnerManifest, competitor })
		cases.push(...annotateUpstreamRealCases(suite, profile))
		childSuites.push({
			profile: profile.name,
			suite: suite.suite,
			generatedAt: suite.generatedAt,
			cases: suite.cases.length,
		})
		if (!hasFlag('--json')) {
			console.error(`completed ${profile.name}: ${formatBytes(profile.expectedXlsxBytes)} source`)
		}
	}
	const suite = createBenchmarkSuite({
		suite: 'ascend-upstream-real-workbook-profiles',
		kind: 'real-workbook',
		cases,
		metadata: {
			profiles,
			repeat,
			warmup,
			childSuites,
			runnerManifest,
			competitor: competitor ?? 'all',
		},
	})
	if (hasFlag('--json')) {
		console.log(JSON.stringify(suite, null, 2))
		return
	}
	console.log(renderSummary(suite))
}

if (import.meta.main) {
	await main()
}
