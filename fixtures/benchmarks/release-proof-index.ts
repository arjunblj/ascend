import { createHash } from 'node:crypto'
import { basename } from 'node:path'
import {
	type PackageActionProofCaseResult,
	type PackageActionProofResult,
	packageActionProofMarkdown,
	runPackageActionProof,
} from './package-action-proof.ts'
import {
	runSafeOpenProof,
	type SafeOpenProofCaseResult,
	type SafeOpenProofResult,
	safeOpenProofMarkdown,
} from './safe-open-proof.ts'

export type ReleaseProofIndexArtifactName = 'safe-open-proof' | 'package-action-proof'
export type ReleaseProofIndexExcludedEvidenceName = 'practical-latency-contracts'
export type ReleaseProofReadinessOwner = 'correctness' | 'performance' | 'product' | 'release'
export type ReleaseProofReadinessStatus = 'missing' | 'satisfied'

export interface ReleaseProofIndexOptions {
	readonly includeTimings?: boolean
}

export interface ReleaseProofIndexArtifact {
	readonly name: ReleaseProofIndexArtifactName
	readonly command: string
	readonly claim: string
	readonly publicationStatus: 'local-proof-ready' | 'needs-release-packaging'
	readonly publicationBlockers: readonly string[]
	readonly readyWhen: readonly ReleaseProofReadinessRequirement[]
	readonly headlineClaimAllowed: boolean
	readonly releaseGate: 'ready' | 'blocked-by-publication-policy'
	readonly sha256: string
	readonly stableShapeSha256: string
	readonly jsonBytes: number
	readonly markdownBytes: number
	readonly fixtureProvenance: ReleaseProofFixtureProvenance
	readonly summary: Readonly<Record<string, string | number | boolean>>
	readonly boundary: string
}

export interface ReleaseProofReadinessRequirement {
	readonly id: string
	readonly status: ReleaseProofReadinessStatus
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly requirement: string
	readonly evidence?: string
}

export interface ReleaseProofFixtureProvenance {
	readonly publicFixtureCases: number
	readonly generatedWorkbookCases: number
	readonly generatedEdgePackageCases: number
	readonly malformedCases: number
	readonly publicFixtureNames: readonly string[]
	readonly generatedCaseNames: readonly string[]
	readonly deterministicGeneratedCaseNames: readonly string[]
	readonly generatedCaseSha256: Readonly<Record<string, string>>
	readonly boundary: string
}

export interface ReleaseProofIndexExcludedEvidence {
	readonly name: ReleaseProofIndexExcludedEvidenceName
	readonly command: string
	readonly reason: string
	readonly eligibilityRule: string
	readonly ownerLoop: 'performance' | 'product' | 'release'
	readonly boundary: string
}

export interface ReleaseProofIndexResult {
	readonly generatedAt: string
	readonly artifactCount: number
	readonly excludedEvidenceCount: number
	readonly signed: false
	readonly attestation: false
	readonly readiness: ReleaseProofReadinessSummary
	readonly boundary: string
	readonly artifacts: readonly ReleaseProofIndexArtifact[]
	readonly excludedEvidence: readonly ReleaseProofIndexExcludedEvidence[]
}

export interface ReleaseProofReadinessSummary {
	readonly releaseGate: 'ready' | 'blocked-by-publication-policy'
	readonly headlineClaimsAllowed: boolean
	readonly totalRequirementCount: number
	readonly missingRequirementCount: number
	readonly satisfiedRequirementCount: number
	readonly missingByOwnerLoop: Readonly<Record<ReleaseProofReadinessOwner, number>>
	readonly missingByArtifact: Readonly<Record<ReleaseProofIndexArtifactName, readonly string[]>>
	readonly boundary: string
}

export async function runReleaseProofIndex(
	options: ReleaseProofIndexOptions = {},
): Promise<ReleaseProofIndexResult> {
	const includeTimings = options.includeTimings ?? false
	const safeOpen = await runSafeOpenProof({
		repeat: includeTimings ? 3 : 1,
		warmup: includeTimings ? 1 : 0,
		includeTimings,
	})
	const packageAction = await runPackageActionProof({ includeTimings })
	const artifacts = [
		safeOpenArtifact(safeOpen, includeTimings),
		packageActionArtifact(packageAction, includeTimings),
	]
	return {
		generatedAt: new Date().toISOString(),
		artifactCount: artifacts.length,
		excludedEvidenceCount: EXCLUDED_EVIDENCE.length,
		signed: false,
		attestation: false,
		readiness: releaseReadinessSummary(artifacts),
		boundary:
			'Digest index for local release evidence artifacts. This is not signed provenance, SLSA, in-toto attestation, or tamper-evident storage.',
		artifacts,
		excludedEvidence: EXCLUDED_EVIDENCE,
	}
}

export function releaseProofIndexMarkdown(result: ReleaseProofIndexResult): string {
	return [
		'# Release Proof Evidence Index',
		'',
		`Generated: ${result.generatedAt}`,
		result.boundary,
		'',
		'| Artifact | Claim | Command | Publication status | Release gate | Headline claim allowed | Publication blockers | Ready when | JSON bytes | Markdown bytes | Fixture provenance | SHA-256 | Stable shape SHA-256 | Summary | Boundary |',
		'| --- | --- | --- | --- | --- | --- | --- | --- | ---: | ---: | --- | --- | --- | --- | --- |',
		...result.artifacts.map(markdownRow),
		'',
		'## Release Readiness Gate',
		'',
		`Release gate: ${result.readiness.releaseGate}`,
		`Headline claims allowed: ${result.readiness.headlineClaimsAllowed}`,
		`ReadyWhen requirements: total=${result.readiness.totalRequirementCount}, missing=${result.readiness.missingRequirementCount}, satisfied=${result.readiness.satisfiedRequirementCount}`,
		`Missing by owner loop: ${formatOwnerCounts(result.readiness.missingByOwnerLoop)}`,
		`Missing by artifact: ${formatMissingByArtifact(result.readiness.missingByArtifact)}`,
		result.readiness.boundary,
		'',
		'## Excluded Evidence',
		'',
		'| Evidence | Command | Reason | Eligibility rule | Owner loop | Boundary |',
		'| --- | --- | --- | --- | --- | --- |',
		...result.excludedEvidence.map(excludedEvidenceMarkdownRow),
		'',
		`Signed: ${result.signed}`,
		`Attestation: ${result.attestation}`,
	].join('\n')
}

const EXCLUDED_EVIDENCE: readonly ReleaseProofIndexExcludedEvidence[] = [
	{
		name: 'practical-latency-contracts',
		command: 'bun run fixtures/benchmarks/practical-latency-contracts.ts --json',
		reason:
			'Latency contract reports are diagnostic benchmark evidence, not release proof artifacts.',
		eligibilityRule:
			'Eligible for release proof only after a tracked-clean run over standardized public inputs with product-approved threshold wording.',
		ownerLoop: 'performance',
		boundary:
			'No local timing report in this index is a release performance threshold, signed provenance, or headline product claim.',
	},
]

function safeOpenArtifact(
	result: SafeOpenProofResult,
	includeTimings: boolean,
): ReleaseProofIndexArtifact {
	const json = JSON.stringify(result, null, 2)
	const markdown = safeOpenProofMarkdown(result)
	const statusCounts = countBy(result.cases, (entry) => entry.status)
	const reviewBeforeHydration = result.cases.filter(
		(entry) => entry.reviewBeforeHydration === true,
	).length
	return {
		name: 'safe-open-proof',
		command: includeTimings
			? 'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json'
			: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
		claim: 'safe unknown workbook opening',
		publicationStatus: 'needs-release-packaging',
		publicationBlockers: [
			'signed and unknown-part cases are durable code-generated packages, not public binary fixtures',
			'local timing evidence is proof-run data, not a release performance threshold',
		],
		readyWhen: [
			{
				id: 'public-edge-fixtures',
				status: 'missing',
				ownerLoop: 'product',
				requirement:
					'replace generated signed/unknown-part packages with public binary fixtures or explicitly approve disclosed generated edge packages',
				evidence: 'safe-open fixture scan currently finds no checked-in public binary replacements',
			},
			{
				id: 'release-latency-run',
				status: 'missing',
				ownerLoop: 'performance',
				requirement:
					'run tracked-clean release-environment open-plan latency evidence over standardized public inputs with approved threshold wording',
			},
			{
				id: 'publication-boundary',
				status: 'missing',
				ownerLoop: 'release',
				requirement:
					'approve boundary language that excludes malware scanning, sandboxing, file trust, active-content safety, and signed provenance',
			},
		],
		headlineClaimAllowed: false,
		releaseGate: 'blocked-by-publication-policy',
		sha256: sha256(json),
		stableShapeSha256: sha256(stableJson(stripRunNoise(result))),
		jsonBytes: utf8Bytes(json),
		markdownBytes: utf8Bytes(markdown),
		fixtureProvenance: safeOpenFixtureProvenance(result.cases),
		summary: {
			cases: result.cases.length,
			ok: statusCounts.ok ?? 0,
			rejected: statusCounts.rejected ?? 0,
			reviewBeforeHydration,
			malformedRejected: result.cases.some(
				(entry) => entry.name === 'malformed' && entry.status === 'rejected',
			),
		},
		boundary:
			'Pre-hydration package-feature routing only; not malware scanning, sandboxing, active-content safety, or malformed-package recovery.',
	}
}

function packageActionArtifact(
	result: PackageActionProofResult,
	includeTimings: boolean,
): ReleaseProofIndexArtifact {
	const json = JSON.stringify(result, null, 2)
	const markdown = packageActionProofMarkdown(result)
	const counts = result.combinedCommitActionCounts
	return {
		name: 'package-action-proof',
		command: includeTimings
			? 'bun run fixtures/benchmarks/package-action-proof.ts --json'
			: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
		claim: 'auditable package-part mutation',
		publicationStatus: 'needs-release-packaging',
		publicationBlockers: [
			'synthetic edge packages must stay disclosed unless replaced by public binary fixtures',
			'proof is local evidence, not signed provenance or third-party attestation',
		],
		readyWhen: [
			{
				id: 'edge-fixture-policy',
				status: 'missing',
				ownerLoop: 'product',
				requirement:
					'accept disclosed generated edge packages as release proof or replace them with public binary fixtures',
				evidence:
					'current proof uses generated calc-chain, signature-invalidation, unknown-part, and docProps edge packages',
			},
			{
				id: 'provenance-boundary',
				status: 'missing',
				ownerLoop: 'release',
				requirement:
					'approve local-proof wording that does not imply SLSA, in-toto, signed provenance, or third-party attestation',
			},
			{
				id: 'unsupported-feature-boundary',
				status: 'missing',
				ownerLoop: 'correctness',
				requirement:
					'approve boundaries for signatures, chart byte passthrough, Excel recalculation equivalence, and unsupported feature semantics',
			},
		],
		headlineClaimAllowed: false,
		releaseGate: 'blocked-by-publication-policy',
		sha256: sha256(json),
		stableShapeSha256: sha256(stableJson(stripRunNoise(result))),
		jsonBytes: utf8Bytes(json),
		markdownBytes: utf8Bytes(markdown),
		fixtureProvenance: packageActionFixtureProvenance(result.cases),
		summary: {
			cases: result.cases.length,
			passthrough: counts.passthrough,
			regenerate: counts.regenerate,
			add: counts.add,
			drop: counts.drop,
			error: counts.error,
			allActionsCovered: Object.values(counts).every((count) => count > 0),
			sourceGraphEverywhere: result.cases.every(
				(entry) => entry.commitCoverage.sourceGraphIncluded,
			),
		},
		boundary:
			'Local package-part evidence only; not signed provenance, Excel recalculation equivalence, or semantic understanding of every unsupported feature.',
	}
}

function markdownRow(row: ReleaseProofIndexArtifact): string {
	return [
		row.name,
		row.claim,
		`\`${row.command}\``,
		row.publicationStatus,
		row.releaseGate,
		String(row.headlineClaimAllowed),
		row.publicationBlockers.join('; '),
		formatReadyWhen(row.readyWhen),
		String(row.jsonBytes),
		String(row.markdownBytes),
		formatFixtureProvenance(row.fixtureProvenance),
		`\`${row.sha256}\``,
		`\`${row.stableShapeSha256}\``,
		formatSummary(row.summary),
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function formatReadyWhen(requirements: readonly ReleaseProofReadinessRequirement[]): string {
	return requirements
		.map(
			(requirement) =>
				`${requirement.id}(${requirement.status},${requirement.ownerLoop})=${requirement.requirement}`,
		)
		.join('; ')
}

function releaseReadinessSummary(
	artifacts: readonly ReleaseProofIndexArtifact[],
): ReleaseProofReadinessSummary {
	const missingByOwnerLoop: Record<ReleaseProofReadinessOwner, number> = {
		correctness: 0,
		performance: 0,
		product: 0,
		release: 0,
	}
	const missingByArtifact: Record<ReleaseProofIndexArtifactName, string[]> = {
		'safe-open-proof': [],
		'package-action-proof': [],
	}
	let totalRequirementCount = 0
	let missingRequirementCount = 0
	let satisfiedRequirementCount = 0
	for (const artifact of artifacts) {
		for (const requirement of artifact.readyWhen) {
			totalRequirementCount += 1
			if (requirement.status === 'satisfied') {
				satisfiedRequirementCount += 1
			} else {
				missingRequirementCount += 1
				missingByOwnerLoop[requirement.ownerLoop] += 1
				missingByArtifact[artifact.name]?.push(requirement.id)
			}
		}
	}
	const headlineClaimsAllowed =
		missingRequirementCount === 0 &&
		artifacts.every((artifact) => artifact.headlineClaimAllowed && artifact.releaseGate === 'ready')
	return {
		releaseGate: headlineClaimsAllowed ? 'ready' : 'blocked-by-publication-policy',
		headlineClaimsAllowed,
		totalRequirementCount,
		missingRequirementCount,
		satisfiedRequirementCount,
		missingByOwnerLoop,
		missingByArtifact,
		boundary:
			'Aggregate release readiness is a publication gate over local proof artifacts. It is not signed provenance, attestation verification, or a substitute for owner approval of each missing requirement.',
	}
}

function formatOwnerCounts(counts: Readonly<Record<ReleaseProofReadinessOwner, number>>): string {
	return (['correctness', 'performance', 'product', 'release'] as const)
		.map((owner) => `${owner}=${counts[owner]}`)
		.join(', ')
}

function formatMissingByArtifact(
	missingByArtifact: Readonly<Record<ReleaseProofIndexArtifactName, readonly string[]>>,
): string {
	return (['safe-open-proof', 'package-action-proof'] as const)
		.map((artifact) => `${artifact}=${missingByArtifact[artifact].join(',') || 'none'}`)
		.join('; ')
}

function safeOpenFixtureProvenance(
	cases: readonly SafeOpenProofCaseResult[],
): ReleaseProofFixtureProvenance {
	const byKind = countBy(cases, (entry) => entry.kind)
	return {
		publicFixtureCases: byKind.file ?? 0,
		generatedWorkbookCases: 0,
		generatedEdgePackageCases: byKind.synthetic ?? 0,
		malformedCases: byKind.malformed ?? 0,
		publicFixtureNames: cases.filter((entry) => entry.kind === 'file').map((entry) => entry.name),
		generatedCaseNames: cases.filter((entry) => entry.kind !== 'file').map((entry) => entry.name),
		deterministicGeneratedCaseNames: cases
			.filter((entry) => entry.kind !== 'file')
			.map((entry) => entry.name),
		generatedCaseSha256: Object.fromEntries(
			cases
				.filter((entry) => entry.kind !== 'file')
				.map((entry) => [entry.name, entry.inputSha256] as const),
		),
		boundary:
			'Public fixture cases are checked-in workbook files; generated edge and malformed case digests identify harness-generated bytes, not signed attestations or public binary fixtures.',
	}
}

function packageActionFixtureProvenance(
	cases: readonly PackageActionProofCaseResult[],
): ReleaseProofFixtureProvenance {
	const byKind = countBy(cases, (entry) => entry.sourceKind)
	return {
		publicFixtureCases: byKind['public-fixture'] ?? 0,
		generatedWorkbookCases: byKind['generated-workbook'] ?? 0,
		generatedEdgePackageCases: byKind['generated-edge-package'] ?? 0,
		malformedCases: 0,
		publicFixtureNames: cases
			.filter((entry) => entry.sourceKind === 'public-fixture')
			.map((entry) => entry.name),
		generatedCaseNames: cases
			.filter((entry) => entry.sourceKind !== 'public-fixture')
			.map((entry) => entry.name),
		deterministicGeneratedCaseNames: cases
			.filter((entry) => entry.sourceKind === 'generated-edge-package')
			.map((entry) => entry.name),
		generatedCaseSha256: Object.fromEntries(
			cases
				.filter((entry) => entry.sourceKind === 'generated-edge-package')
				.map((entry) => [entry.name, entry.inputSha256] as const),
		),
		boundary:
			'Public fixture cases are checked-in workbook files; generated edge package digests identify deterministic local harness inputs that must stay disclosed. Generated workbook cases are named but not digested because workbook package timestamps can vary.',
	}
}

function formatFixtureProvenance(provenance: ReleaseProofFixtureProvenance): string {
	return [
		`public=${provenance.publicFixtureCases}`,
		`generatedWorkbook=${provenance.generatedWorkbookCases}`,
		`generatedEdge=${provenance.generatedEdgePackageCases}`,
		`malformed=${provenance.malformedCases}`,
		`generatedCases=${provenance.generatedCaseNames.join(',') || 'none'}`,
		`deterministicGenerated=${provenance.deterministicGeneratedCaseNames.join(',') || 'none'}`,
		`generatedDigests=${formatGeneratedCaseDigests(provenance.generatedCaseSha256)}`,
	].join('; ')
}

function formatGeneratedCaseDigests(digests: Readonly<Record<string, string>>): string {
	const entries = Object.entries(digests)
	return entries.length > 0
		? entries.map(([name, digest]) => `${name}:${digest.slice(0, 12)}`).join(',')
		: 'none'
}

function excludedEvidenceMarkdownRow(row: ReleaseProofIndexExcludedEvidence): string {
	return [
		row.name,
		`\`${row.command}\``,
		row.reason,
		row.eligibilityRule,
		row.ownerLoop,
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function countBy<T>(
	values: readonly T[],
	keyFor: (value: T) => string,
): Readonly<Record<string, number>> {
	const counts: Record<string, number> = {}
	for (const value of values) {
		const key = keyFor(value)
		counts[key] = (counts[key] ?? 0) + 1
	}
	return counts
}

function formatSummary(summary: Readonly<Record<string, string | number | boolean>>): string {
	return Object.entries(summary)
		.map(([key, value]) => `${key}=${value}`)
		.join(', ')
}

function stripRunNoise(value: unknown): unknown {
	if (Array.isArray(value)) return value.map(stripRunNoise)
	if (!value || typeof value !== 'object') return value
	const entries = Object.entries(value as Record<string, unknown>)
		.filter(([key]) => key !== 'generatedAt' && key !== 'inputSha256' && !key.endsWith('MedianMs'))
		.map(([key, entry]) => [key, stripRunNoise(entry)] as const)
	return Object.fromEntries(entries)
}

function stableJson(value: unknown): string {
	if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`
	if (value && typeof value === 'object') {
		return `{${Object.entries(value as Record<string, unknown>)
			.sort(([left], [right]) => left.localeCompare(right))
			.map(([key, entry]) => `${JSON.stringify(key)}:${stableJson(entry)}`)
			.join(',')}}`
	}
	return JSON.stringify(value)
}

function sha256(value: string): string {
	return createHash('sha256').update(value).digest('hex')
}

function utf8Bytes(value: string): number {
	return new TextEncoder().encode(value).byteLength
}

if (import.meta.main) {
	const json = process.argv.includes('--json')
	const result = await runReleaseProofIndex({
		includeTimings: !process.argv.includes('--no-timings'),
	})
	console.log(json ? JSON.stringify(result, null, 2) : releaseProofIndexMarkdown(result))
	if (!json) {
		console.error(`Indexed ${result.artifactCount} release proof evidence artifacts.`)
		console.error(`Run with --json for machine-readable output from ${basename(import.meta.path)}.`)
	}
}
