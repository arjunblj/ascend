import { createHash } from 'node:crypto'
import { basename } from 'node:path'
import {
	type PackageActionProofResult,
	packageActionProofMarkdown,
	runPackageActionProof,
} from './package-action-proof.ts'
import {
	runSafeOpenProof,
	type SafeOpenProofResult,
	safeOpenProofMarkdown,
} from './safe-open-proof.ts'

export type ReleaseProofIndexArtifactName = 'safe-open-proof' | 'package-action-proof'
export type ReleaseProofIndexExcludedEvidenceName = 'practical-latency-contracts'

export interface ReleaseProofIndexOptions {
	readonly includeTimings?: boolean
}

export interface ReleaseProofIndexArtifact {
	readonly name: ReleaseProofIndexArtifactName
	readonly command: string
	readonly claim: string
	readonly publicationStatus: 'local-proof-ready' | 'needs-release-packaging'
	readonly publicationBlockers: readonly string[]
	readonly headlineClaimAllowed: boolean
	readonly releaseGate: 'ready' | 'blocked-by-publication-policy'
	readonly sha256: string
	readonly stableShapeSha256: string
	readonly jsonBytes: number
	readonly markdownBytes: number
	readonly summary: Readonly<Record<string, string | number | boolean>>
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
	readonly boundary: string
	readonly artifacts: readonly ReleaseProofIndexArtifact[]
	readonly excludedEvidence: readonly ReleaseProofIndexExcludedEvidence[]
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
		'| Artifact | Claim | Command | Publication status | Release gate | Headline claim allowed | Publication blockers | JSON bytes | Markdown bytes | SHA-256 | Stable shape SHA-256 | Summary | Boundary |',
		'| --- | --- | --- | --- | --- | --- | --- | ---: | ---: | --- | --- | --- | --- |',
		...result.artifacts.map(markdownRow),
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
		headlineClaimAllowed: false,
		releaseGate: 'blocked-by-publication-policy',
		sha256: sha256(json),
		stableShapeSha256: sha256(stableJson(stripRunNoise(result))),
		jsonBytes: utf8Bytes(json),
		markdownBytes: utf8Bytes(markdown),
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
		headlineClaimAllowed: false,
		releaseGate: 'blocked-by-publication-policy',
		sha256: sha256(json),
		stableShapeSha256: sha256(stableJson(stripRunNoise(result))),
		jsonBytes: utf8Bytes(json),
		markdownBytes: utf8Bytes(markdown),
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
		String(row.jsonBytes),
		String(row.markdownBytes),
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
		.filter(([key]) => key !== 'generatedAt' && !key.endsWith('MedianMs'))
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
