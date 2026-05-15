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
export type ReleaseProofDeferredClaimName =
	| 'formula-language-service-primitives'
	| 'token-bounded-agent-view'
	| 'retained-viewport-patch-history'
	| 'columnar-scan-sidecars'
	| 'formula-oracle-routing'
	| 'agent-workflow-observability'

export interface ReleaseProofIndexOptions {
	readonly includeTimings?: boolean
}

export interface ReleaseProofIndexArtifact {
	readonly name: ReleaseProofIndexArtifactName
	readonly command: string
	readonly compactReportCommand?: string
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
	readonly deferredClaimCount: number
	readonly signed: false
	readonly attestation: false
	readonly readiness: ReleaseProofReadinessSummary
	readonly boundary: string
	readonly artifacts: readonly ReleaseProofIndexArtifact[]
	readonly excludedEvidence: readonly ReleaseProofIndexExcludedEvidence[]
	readonly deferredClaims: readonly ReleaseProofDeferredClaim[]
}

export interface ReleaseProofOwnerHandoffIndex {
	readonly generatedAt: string
	readonly releaseGate: ReleaseProofReadinessSummary['releaseGate']
	readonly headlineClaimsAllowed: boolean
	readonly implementationSurfacePromotionAllowed: boolean
	readonly missingRequirementCount: number
	readonly implementationHandoffs: readonly ReleaseProofImplementationHandoff[]
	readonly deferredClaims: readonly ReleaseProofDeferredClaim[]
	readonly excludedEvidence: readonly ReleaseProofIndexExcludedEvidence[]
	readonly boundary: string
}

export interface ReleaseProofReadinessSummary {
	readonly releaseGate: 'ready' | 'blocked-by-publication-policy'
	readonly headlineClaimsAllowed: boolean
	readonly implementationSurfacePromotionAllowed: boolean
	readonly implementationSurfacePromotionBoundary: string
	readonly totalRequirementCount: number
	readonly missingRequirementCount: number
	readonly satisfiedRequirementCount: number
	readonly missingByOwnerLoop: Readonly<Record<ReleaseProofReadinessOwner, number>>
	readonly missingByArtifact: Readonly<Record<ReleaseProofIndexArtifactName, readonly string[]>>
	readonly nextOwnerActions: readonly ReleaseProofNextOwnerAction[]
	readonly implementationHandoffs: readonly ReleaseProofImplementationHandoff[]
	readonly boundary: string
}

export interface ReleaseProofNextOwnerAction {
	readonly rank: number
	readonly artifact: ReleaseProofIndexArtifactName
	readonly requirementId: string
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly priority: 'claim-evidence' | 'claim-boundary' | 'publication-policy'
	readonly nextStepKind:
		| 'owner-decision-or-fixture-replacement'
		| 'owner-boundary-approval'
		| 'validation-run'
		| 'owner-decision-or-harness-expansion'
		| 'publication-policy'
	readonly rationale: string
}

export interface ReleaseProofImplementationHandoff {
	readonly rank: number
	readonly artifact: ReleaseProofIndexArtifactName
	readonly claim: string
	readonly proofRequired: ReleaseProofClaimProofRequired
	readonly ownerLoops: readonly ReleaseProofReadinessOwner[]
	readonly proofCommand: string
	readonly compactReportCommand?: string
	readonly implementationSurfacePromotionAllowed: boolean
	readonly blockingRequirementIds: readonly string[]
	readonly nextStepKinds: readonly ReleaseProofNextOwnerAction['nextStepKind'][]
	readonly boundary: string
}

export interface ReleaseProofClaimProofRequired {
	readonly fixture: string
	readonly benchmark: string
	readonly surface: string
	readonly validationGate: string
	readonly competitorContrast: string
	readonly honestBoundary: string
	readonly killCriterion: string
}

export interface ReleaseProofDeferredClaim {
	readonly name: ReleaseProofDeferredClaimName
	readonly claim: string
	readonly status: 'do-not-promote-yet' | 'proof-backed-hold'
	readonly ownerLoops: readonly ReleaseProofReadinessOwner[]
	readonly reason: string
	readonly proofNeeded: string
	readonly killCriterion: string
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
		deferredClaimCount: DEFERRED_CLAIMS.length,
		signed: false,
		attestation: false,
		readiness: releaseReadinessSummary(artifacts),
		boundary:
			'Digest index for local release evidence artifacts. This is not signed provenance, SLSA, in-toto attestation, or tamper-evident storage.',
		artifacts,
		excludedEvidence: EXCLUDED_EVIDENCE,
		deferredClaims: DEFERRED_CLAIMS.map(cloneDeferredClaim),
	}
}

export function releaseProofIndexMarkdown(result: ReleaseProofIndexResult): string {
	return [
		'# Release Proof Evidence Index',
		'',
		`Generated: ${result.generatedAt}`,
		result.boundary,
		'',
		'| Artifact | Claim | Command | Compact report command | Publication status | Release gate | Headline claim allowed | Publication blockers | Ready when | JSON bytes | Markdown bytes | Fixture provenance | SHA-256 | Stable shape SHA-256 | Summary | Boundary |',
		'| --- | --- | --- | --- | --- | --- | --- | --- | --- | ---: | ---: | --- | --- | --- | --- | --- |',
		...result.artifacts.map(markdownRow),
		'',
		'## Release Readiness Gate',
		'',
		`Release gate: ${result.readiness.releaseGate}`,
		`Headline claims allowed: ${result.readiness.headlineClaimsAllowed}`,
		`Implementation surface promotion allowed: ${result.readiness.implementationSurfacePromotionAllowed}`,
		result.readiness.implementationSurfacePromotionBoundary,
		`ReadyWhen requirements: total=${result.readiness.totalRequirementCount}, missing=${result.readiness.missingRequirementCount}, satisfied=${result.readiness.satisfiedRequirementCount}`,
		`Missing by owner loop: ${formatOwnerCounts(result.readiness.missingByOwnerLoop)}`,
		`Missing by artifact: ${formatMissingByArtifact(result.readiness.missingByArtifact)}`,
		`Next owner actions: ${formatNextOwnerActions(result.readiness.nextOwnerActions)}`,
		`Implementation handoffs: ${formatImplementationHandoffs(result.readiness.implementationHandoffs)}`,
		result.readiness.boundary,
		'',
		'## Excluded Evidence',
		'',
		'| Evidence | Command | Reason | Eligibility rule | Owner loop | Boundary |',
		'| --- | --- | --- | --- | --- | --- |',
		...result.excludedEvidence.map(excludedEvidenceMarkdownRow),
		'',
		'## Deferred Claims',
		'',
		'| Claim | Status | Owner loops | Reason | Proof needed | Kill criterion | Boundary |',
		'| --- | --- | --- | --- | --- | --- | --- |',
		...result.deferredClaims.map(deferredClaimMarkdownRow),
		'',
		`Signed: ${result.signed}`,
		`Attestation: ${result.attestation}`,
	].join('\n')
}

export function releaseProofOwnerHandoffIndex(
	result: ReleaseProofIndexResult,
): ReleaseProofOwnerHandoffIndex {
	return {
		generatedAt: result.generatedAt,
		releaseGate: result.readiness.releaseGate,
		headlineClaimsAllowed: result.readiness.headlineClaimsAllowed,
		implementationSurfacePromotionAllowed: result.readiness.implementationSurfacePromotionAllowed,
		missingRequirementCount: result.readiness.missingRequirementCount,
		implementationHandoffs: result.readiness.implementationHandoffs,
		deferredClaims: result.deferredClaims,
		excludedEvidence: result.excludedEvidence,
		boundary:
			'Compact owner handoff index for release proof routing. It is not a release artifact bundle, signed attestation, or product surface.',
	}
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

const DEFERRED_CLAIMS: readonly ReleaseProofDeferredClaim[] = [
	{
		name: 'formula-language-service-primitives',
		claim: 'formula language-service primitives',
		status: 'proof-backed-hold',
		ownerLoops: ['product', 'correctness'],
		reason:
			'Current proof supports spans, diagnostics, hover, completions, binding roles, and rejection-first prepareRename only; edit-producing rename is frozen.',
		proofNeeded:
			'Workbook-context ownership for defined names, table columns, sheet refs, external refs, and operation-owned edit plans before any rename promotion.',
		killCriterion:
			'Do not promote rename while prepareRename must refuse workbook-context targets or while no operation owns cross-workbook edits.',
		boundary:
			'No edit-producing rename, no table/defined-name rename, and no claim that all formula references can be rewritten safely.',
	},
	{
		name: 'token-bounded-agent-view',
		claim: 'token-bounded agent view',
		status: 'proof-backed-hold',
		ownerLoops: ['product'],
		reason:
			'Current proof is useful for deterministic compact views, but release copy still needs a concrete product example and budget-boundary wording.',
		proofNeeded:
			'One public workbook example showing requested budget, estimated tokens, omissions, locators, and recovery path across existing surfaces.',
		killCriterion:
			'Do not publish exact-token wording when structural floors exceed tiny budgets or omitted evidence is not recoverable by locator.',
		boundary:
			'Approximate token counts only; compact views do not replace workbook inspection, proof artifacts, or omitted workbook evidence.',
	},
	{
		name: 'retained-viewport-patch-history',
		claim: 'retained viewport patch history',
		status: 'proof-backed-hold',
		ownerLoops: ['product', 'performance'],
		reason:
			'SDK/API/MCP retained patch proof exists, but CLI is excluded and the claim must not drift into collaboration or transaction-isolation language.',
		proofNeeded:
			'Owner-approved product wording for bounded history, invalidation reasons, retention caps, and cross-surface recovery without CLI claims.',
		killCriterion:
			'Do not promote collaboration, sync, CRDT, or unlimited-history wording without multi-writer convergence and storage retention proof.',
		boundary:
			'Bounded per-window patch history only; not MVCC transaction isolation, collaborative editing, or unlimited audit history.',
	},
	{
		name: 'columnar-scan-sidecars',
		claim: 'columnar scan sidecars',
		status: 'do-not-promote-yet',
		ownerLoops: ['performance'],
		reason:
			'Sidecar probes show promise, but the evidence is still benchmark/research shaped rather than product-surface shaped.',
		proofNeeded:
			'Repeated-scan wins, build cost, invalidation cost, memory caps, and checksum parity over multiple public real workbook table/range shapes.',
		killCriterion:
			'Do not promote if build plus invalidation erases repeated-scan gains, parity fails, or memory overhead is not bounded.',
		boundary:
			'Disposable sidecar only; not a storage engine, workbook rewrite, or guaranteed acceleration for sparse or single-pass reads.',
	},
	{
		name: 'formula-oracle-routing',
		claim: 'formula oracle routing',
		status: 'do-not-promote-yet',
		ownerLoops: ['correctness'],
		reason:
			'Mismatch routing is a correctness research program until oracle classes are complete and reproducible without private corpora.',
		proofNeeded:
			'Runnable public corpus artifacts with mismatch classes, HyperFormula/LibreOffice/Excel/static-golden routing, skip counters, and divergence counters.',
		killCriterion:
			'Do not publish Excel-compatible formula wording while cached values, private corpora, or unsupported oracle classes are required.',
		boundary:
			'No blanket Excel-compatibility claim and no claim that every mismatch class has an automated oracle.',
	},
	{
		name: 'agent-workflow-observability',
		claim: 'agent workflow observability',
		status: 'do-not-promote-yet',
		ownerLoops: ['product'],
		reason:
			'Trace ideas are promising, but proof must show they improve repair, audit, or recovery decisions rather than duplicating logs.',
		proofNeeded:
			'Public inspect/plan/commit/reopen/diff/audit workflow traces with failure taxonomy and recovery prompts over existing agent surfaces.',
		killCriterion:
			'Do not promote if traces are only verbose logs or if they do not improve a concrete repair, audit, or recovery workflow.',
		boundary:
			'No claim of autonomous correctness, signed audit trail, or complete observability across every workbook feature.',
	},
]

function cloneDeferredClaim(claim: ReleaseProofDeferredClaim): ReleaseProofDeferredClaim {
	return {
		...claim,
		ownerLoops: [...claim.ownerLoops],
	}
}

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
		compactReportCommand:
			'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json',
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
			{
				id: 'compact-report-publication-policy',
				status: 'missing',
				ownerLoop: 'release',
				requirement:
					'define artifact storage, privacy filtering, and canonicalization policy before compact report digests are indexed or published',
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
		compactReportCommand:
			'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json',
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
					'current proof uses public calc-chain, macro, and chart fixtures, plus generated docProps, signature-invalidation, and unknown-part edge packages',
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
			{
				id: 'streaming-matrix-boundary',
				status: 'missing',
				ownerLoop: 'performance',
				requirement:
					'approve that one representative streaming writer proof is sufficient for release wording, or expand package-action proof to streaming variants for every package-action scenario before claiming streaming parity',
				evidence: 'current proof reports one streamingProofCase and one streamingRegeneratePart',
			},
			{
				id: 'compact-report-publication-policy',
				status: 'missing',
				ownerLoop: 'release',
				requirement:
					'define artifact storage, privacy filtering, and canonicalization policy before compact report digests are indexed or published',
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
			streamingProofCases: result.cases.filter((entry) => entry.streamingProof !== undefined)
				.length,
			streamingRegenerateParts: result.cases.reduce(
				(count, entry) => count + (entry.streamingProof?.streamingRegeneratePartPaths.length ?? 0),
				0,
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
		row.compactReportCommand ? `\`${row.compactReportCommand}\`` : 'n/a',
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
	const missingRequirements: Array<{
		readonly artifact: ReleaseProofIndexArtifactName
		readonly requirement: ReleaseProofReadinessRequirement
	}> = []
	for (const artifact of artifacts) {
		for (const requirement of artifact.readyWhen) {
			totalRequirementCount += 1
			if (requirement.status === 'satisfied') {
				satisfiedRequirementCount += 1
			} else {
				missingRequirementCount += 1
				missingByOwnerLoop[requirement.ownerLoop] += 1
				missingByArtifact[artifact.name]?.push(requirement.id)
				missingRequirements.push({ artifact: artifact.name, requirement })
			}
		}
	}
	const headlineClaimsAllowed =
		missingRequirementCount === 0 &&
		artifacts.every((artifact) => artifact.headlineClaimAllowed && artifact.releaseGate === 'ready')
	const implementationSurfacePromotionAllowed = headlineClaimsAllowed
	const nextOwnerActions = missingRequirements.map(rankMissingRequirement).sort((left, right) => {
		const byRank = left.rank - right.rank
		if (byRank !== 0) return byRank
		return `${left.artifact}:${left.requirementId}`.localeCompare(
			`${right.artifact}:${right.requirementId}`,
		)
	})
	return {
		releaseGate: headlineClaimsAllowed ? 'ready' : 'blocked-by-publication-policy',
		headlineClaimsAllowed,
		implementationSurfacePromotionAllowed,
		implementationSurfacePromotionBoundary: implementationSurfacePromotionAllowed
			? 'Implementation surfaces may be considered only after all release proof gates are satisfied and owner loops approve the product shape.'
			: 'Current release proof blockers are owner decisions, validation runs, optional harness expansion, or publication policy. They do not authorize new SDK, CLI, API, or MCP surfaces.',
		totalRequirementCount,
		missingRequirementCount,
		satisfiedRequirementCount,
		missingByOwnerLoop,
		missingByArtifact,
		nextOwnerActions,
		implementationHandoffs: buildImplementationHandoffs(
			artifacts,
			nextOwnerActions,
			implementationSurfacePromotionAllowed,
		),
		boundary:
			'Aggregate release readiness is a publication gate over local proof artifacts. It is not signed provenance, attestation verification, or a substitute for owner approval of each missing requirement.',
	}
}

function buildImplementationHandoffs(
	artifacts: readonly ReleaseProofIndexArtifact[],
	nextOwnerActions: readonly ReleaseProofNextOwnerAction[],
	implementationSurfacePromotionAllowed: boolean,
): readonly ReleaseProofImplementationHandoff[] {
	return artifacts.map((artifact, index) => {
		const artifactActions = nextOwnerActions.filter((action) => action.artifact === artifact.name)
		return {
			rank: index + 1,
			artifact: artifact.name,
			claim: artifact.claim,
			proofRequired: claimProofRequired(artifact.name),
			ownerLoops: uniqueOwnerLoops(artifact.readyWhen.map((requirement) => requirement.ownerLoop)),
			proofCommand: artifact.command,
			compactReportCommand: artifact.compactReportCommand,
			implementationSurfacePromotionAllowed,
			blockingRequirementIds: artifact.readyWhen
				.filter((requirement) => requirement.status === 'missing')
				.map((requirement) => requirement.id),
			nextStepKinds: uniqueNextStepKinds(artifactActions.map((action) => action.nextStepKind)),
			boundary:
				'Owner handoff for proof, validation, boundary approval, and publication policy only; it is not permission to add new SDK, CLI, API, or MCP surfaces.',
		}
	})
}

function claimProofRequired(
	artifact: ReleaseProofIndexArtifactName,
): ReleaseProofClaimProofRequired {
	switch (artifact) {
		case 'safe-open-proof':
			return {
				fixture:
					'Public clean, formula-heavy, macro, pivot, ActiveX, chart, signed, unknown-part, and malformed workbook/package cases; generated signed and unknown-part cases must stay disclosed unless replaced by public binary fixtures.',
				benchmark:
					'Release-environment open-plan latency over standardized public inputs, compared with full hydration and approved threshold wording.',
				surface:
					'Existing SDK open planner, CLI open-plan, API open-plan endpoint, and MCP open-plan tool only; no new opener surface.',
				validationGate:
					'Run the safe-open proof harness, focused open-plan tests, malformed package checks, release-proof-index, typecheck, Biome, and changed tests when code changes.',
				competitorContrast:
					'Microsoft Protected View is trust/read-only UX for potentially unsafe files; Ascend claim is OSS pre-hydration package-feature routing.',
				honestBoundary:
					'Not malware scanning, sandboxing, file trust, active-content safety, signed provenance, or malformed-package recovery.',
				killCriterion:
					'Do not publish headline wording if generated signed/unknown packages are hidden, product rejects disclosed generated topology fixtures, or latency wording lacks an approved release-environment run.',
			}
		case 'package-action-proof':
			return {
				fixture:
					'Public and generated package cases covering docProps passthrough, worksheet regeneration, sheet add, public calc-chain drop, signature invalidation, macro/sidecar accounting, chart sidecar accounting, and unknown-part error.',
				benchmark:
					'Package-proof overhead in bytes and milliseconds for plan/commit evidence, including compact versus expanded report paths and the representative streaming writer proof.',
				surface:
					'Existing SDK evidence and compact CLI/API/MCP proof summaries only; no new mutation surface.',
				validationGate:
					'Run the package-action proof harness, plan/commit/reopen/diff/audit checks, journal/package compatibility tests, release-proof-index, schema/typecheck/Biome, and changed tests when code changes.',
				competitorContrast:
					'openpyxl and SheetJS document preservation/write boundaries; Ascend claim is explicit per-part action accounting.',
				honestBoundary:
					'Not signed provenance, SLSA, in-toto attestation, Excel recalculation equivalence, chart byte passthrough, or semantic understanding of every unsupported feature.',
				killCriterion:
					'Do not publish stronger wording if synthetic edge packages are hidden, chart XML is described as byte-passthrough, one streaming proof is described as full matrix parity, or local digests imply attestation.',
			}
	}
}

function rankMissingRequirement(input: {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly requirement: ReleaseProofReadinessRequirement
}): ReleaseProofNextOwnerAction {
	const { artifact, requirement } = input
	switch (requirement.id) {
		case 'public-edge-fixtures':
		case 'edge-fixture-policy':
			return {
				rank: 10,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'claim-evidence',
				nextStepKind: 'owner-decision-or-fixture-replacement',
				rationale:
					'Fixture disclosure or replacement decides whether the proof evidence can support release wording without private or hidden generated inputs.',
			}
		case 'unsupported-feature-boundary':
			return {
				rank: 20,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'claim-boundary',
				nextStepKind: 'owner-boundary-approval',
				rationale:
					'Correctness must approve unsupported-feature boundaries before package-action wording can stay honest.',
			}
		case 'release-latency-run':
			return {
				rank: 30,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'claim-evidence',
				nextStepKind: 'validation-run',
				rationale:
					'Performance evidence is needed before any release wording implies a latency threshold or speed claim.',
			}
		case 'streaming-matrix-boundary':
			return {
				rank: 40,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'claim-boundary',
				nextStepKind: 'owner-decision-or-harness-expansion',
				rationale:
					'Streaming wording must stay limited to one representative proof unless a broader matrix is approved.',
			}
		case 'publication-boundary':
		case 'provenance-boundary':
			return {
				rank: 50,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'publication-policy',
				nextStepKind: 'publication-policy',
				rationale:
					'Release wording must avoid trust, active-content safety, signed provenance, and attestation implications.',
			}
		case 'compact-report-publication-policy':
			return {
				rank: 60,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'publication-policy',
				nextStepKind: 'publication-policy',
				rationale:
					'Compact report storage and canonicalization are needed before digest publication, but not before using generated local proof reports.',
			}
		default:
			return {
				rank: 100,
				artifact,
				requirementId: requirement.id,
				ownerLoop: requirement.ownerLoop,
				priority: 'publication-policy',
				nextStepKind: 'publication-policy',
				rationale: 'Unclassified missing release-readiness requirement.',
			}
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

function formatNextOwnerActions(actions: readonly ReleaseProofNextOwnerAction[]): string {
	return actions
		.map(
			(action) =>
				`${action.rank}:${action.artifact}/${action.requirementId}(${action.ownerLoop},${action.priority},${action.nextStepKind})`,
		)
		.join('; ')
}

function formatImplementationHandoffs(
	handoffs: readonly ReleaseProofImplementationHandoff[],
): string {
	return handoffs
		.map(
			(handoff) =>
				`${handoff.rank}:${handoff.artifact}(${handoff.ownerLoops.join('+')};promotion=${handoff.implementationSurfacePromotionAllowed};blockers=${handoff.blockingRequirementIds.join(',') || 'none'};kill=${handoff.proofRequired.killCriterion})`,
		)
		.join('; ')
}

function uniqueOwnerLoops(
	ownerLoops: readonly ReleaseProofReadinessOwner[],
): readonly ReleaseProofReadinessOwner[] {
	const order: readonly ReleaseProofReadinessOwner[] = [
		'correctness',
		'performance',
		'product',
		'release',
	]
	const seen = new Set(ownerLoops)
	return order.filter((owner) => seen.has(owner))
}

function uniqueNextStepKinds(
	nextStepKinds: readonly ReleaseProofNextOwnerAction['nextStepKind'][],
): readonly ReleaseProofNextOwnerAction['nextStepKind'][] {
	const seen = new Set(nextStepKinds)
	return [
		'owner-decision-or-fixture-replacement',
		'owner-boundary-approval',
		'validation-run',
		'owner-decision-or-harness-expansion',
		'publication-policy',
	].filter((kind) => seen.has(kind))
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

function deferredClaimMarkdownRow(row: ReleaseProofDeferredClaim): string {
	return [
		row.claim,
		row.status,
		row.ownerLoops.join('+'),
		row.reason,
		row.proofNeeded,
		row.killCriterion,
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
	const ownerHandoffsJson = process.argv.includes('--owner-handoffs-json')
	const result = await runReleaseProofIndex({
		includeTimings: !process.argv.includes('--no-timings'),
	})
	console.log(
		ownerHandoffsJson
			? JSON.stringify(releaseProofOwnerHandoffIndex(result), null, 2)
			: json
				? JSON.stringify(result, null, 2)
				: releaseProofIndexMarkdown(result),
	)
	if (!json && !ownerHandoffsJson) {
		console.error(`Indexed ${result.artifactCount} release proof evidence artifacts.`)
		console.error(
			`Run with --json or --owner-handoffs-json for machine-readable output from ${basename(import.meta.path)}.`,
		)
	}
}
