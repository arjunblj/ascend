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
	readonly fixturePolicy: ReleaseProofFixturePolicy
	readonly performancePolicy: ReleaseProofPerformancePolicy
	readonly correctnessPolicy: ReleaseProofCorrectnessPolicy
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
	readonly fixturePolicy: ReleaseProofFixturePolicy
	readonly performancePolicy: ReleaseProofPerformancePolicy
	readonly correctnessPolicy: ReleaseProofCorrectnessPolicy
	readonly nextOwnerActions: readonly ReleaseProofNextOwnerAction[]
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
	readonly acceptanceEvidence: string
	readonly forbiddenShortcut: string
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
	readonly blockingActions: readonly ReleaseProofNextOwnerAction[]
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

export interface ReleaseProofFixturePolicy {
	readonly currentDecision: 'owner-approval-required'
	readonly generatedStructuralFixturesAllowedWhen: readonly string[]
	readonly publicBinaryFixturesRequiredWhen: readonly string[]
	readonly approvalChecklist: readonly ReleaseProofFixturePolicyApprovalItem[]
	readonly trackedFixtureScanCommands: Readonly<Record<ReleaseProofIndexArtifactName, string>>
	readonly currentGeneratedStructuralCases: Readonly<
		Record<ReleaseProofIndexArtifactName, readonly string[]>
	>
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly boundary: string
}

export interface ReleaseProofFixturePolicyApprovalItem {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly gateId: string
	readonly ownerLoop: 'product' | 'release'
	readonly status: 'pending-owner-decision'
	readonly decisionNeeded: string
	readonly acceptanceEvidence: string
	readonly rejectIf: string
	readonly validationCommand: string
}

export interface ReleaseProofSourceReference {
	readonly label: string
	readonly url: string
}

export interface ReleaseProofPerformancePolicy {
	readonly currentDecision: 'owner-approval-required'
	readonly approvalChecklist: readonly ReleaseProofPerformancePolicyApprovalItem[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly boundary: string
}

export interface ReleaseProofPerformancePolicyApprovalItem {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly gateId: string
	readonly ownerLoop: 'performance'
	readonly status: 'pending-owner-decision'
	readonly decisionNeeded: string
	readonly acceptanceEvidence: string
	readonly rejectIf: string
	readonly validationCommand: string
}

export interface ReleaseProofCorrectnessPolicy {
	readonly currentDecision: 'owner-approval-required'
	readonly approvalChecklist: readonly ReleaseProofCorrectnessPolicyApprovalItem[]
	readonly unsupportedFeatureBoundaries: readonly ReleaseProofUnsupportedFeatureBoundary[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly boundary: string
}

export interface ReleaseProofCorrectnessPolicyApprovalItem {
	readonly artifact: 'package-action-proof'
	readonly gateId: 'unsupported-feature-boundary'
	readonly ownerLoop: 'correctness'
	readonly status: 'pending-owner-decision'
	readonly decisionNeeded: string
	readonly acceptanceEvidence: string
	readonly rejectIf: string
	readonly validationCommand: string
}

export interface ReleaseProofUnsupportedFeatureBoundary {
	readonly feature: string
	readonly currentEvidence: string
	readonly allowedWording: string
	readonly forbiddenWording: string
}

const FIXTURE_POLICY: ReleaseProofFixturePolicy = {
	currentDecision: 'owner-approval-required',
	generatedStructuralFixturesAllowedWhen: [
		'edge case is package-topology-only and does not depend on private workbook contents',
		'fixture bytes are generated by tracked harness code and labeled as generated or synthetic',
		'tracked public fixture scan found no replacement candidate for the same package shape',
		'release proof index keeps the owner gate missing until product explicitly accepts generated evidence',
		'claim wording stays below trust, malware scanning, file safety, signed provenance, or vendor-behavior claims',
	],
	publicBinaryFixturesRequiredWhen: [
		'claim depends on real-world authoring behavior, vendor repair behavior, UI behavior, or workbook semantics',
		'claim would imply user trust, malware scanning, active-content safety, signed provenance, or attestation',
		'generated fixture would hide licensing, privacy, provenance, or large/private workbook uncertainty',
	],
	approvalChecklist: [
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			ownerLoop: 'product',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Accept disclosed generated signed/unknown structural packages for guarded safe-open topology proof, or require public binary replacements.',
			acceptanceEvidence:
				'Safe-open proof labels generated signed and unknown-part cases, tracked scan finds no replacement, and claim wording excludes trust, malware scanning, active-content safety, and signed provenance.',
			rejectIf:
				'Generated fixtures are hidden, treated as public binaries, used for trust/safety wording, or replaceable by approved tracked public binaries.',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			ownerLoop: 'product',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Accept disclosed generated signature/unknown structural packages for guarded package-action proof, or require public binary replacements.',
			acceptanceEvidence:
				'Package-action proof labels generated signature-invalidation and unknown-part error cases, tracked scan finds no replacement, and claim wording stays limited to package action accounting.',
			rejectIf:
				'Generated fixtures are hidden, treated as public binaries, used for provenance/trust wording, or replaceable by approved tracked public binaries.',
			validationCommand: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
		},
		{
			artifact: 'safe-open-proof',
			gateId: 'publication-boundary',
			ownerLoop: 'release',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Approve safe-open publication wording below malware scanning, sandboxing, file trust, active-content safety, signed provenance, and malformed-package recovery.',
			acceptanceEvidence:
				'Release wording states pre-hydration package-feature routing only and keeps generated structural fixtures disclosed.',
			rejectIf:
				'Copy implies Protected View equivalence, malware safety, trust certification, signed provenance, or recovery of arbitrary malformed packages.',
			validationCommand:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'provenance-boundary',
			ownerLoop: 'release',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Approve package-action publication wording below SLSA, in-toto, Sigstore, GitHub artifact attestation, and signed-provenance thresholds.',
			acceptanceEvidence:
				'Release wording says local package-part evidence and does not claim signer identity, transparency-log inclusion, tamper-evident storage, or build provenance.',
			rejectIf:
				'Copy calls local digests signed provenance, attestation, tamper-evident storage, malware safety, or active-content trust.',
			validationCommand:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
		},
	],
	trackedFixtureScanCommands: {
		'safe-open-proof': 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
		'package-action-proof': 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
	},
	currentGeneratedStructuralCases: {
		'safe-open-proof': ['signed', 'unknown-part', 'malformed'],
		'package-action-proof': ['signature-invalidation-drop', 'unknown-part-error'],
	},
	sourceReferences: [
		{
			label: 'GitHub repository limits',
			url: 'https://docs.github.com/en/repositories/creating-and-managing-repositories/repository-limits',
		},
		{
			label: 'GitHub large files',
			url: 'https://docs.github.com/github/managing-large-files/working-with-large-files',
		},
		{
			label: 'OpenSSF Scorecard binary artifacts',
			url: 'https://github.com/ossf/scorecard/blob/main/docs/checks.md',
		},
		{
			label: 'SLSA provenance',
			url: 'https://slsa.dev/spec/v1.0-rc1/provenance',
		},
		{
			label: 'GitHub artifact attestations',
			url: 'https://docs.github.com/actions/concepts/security/artifact-attestations',
		},
	],
	boundary:
		'Fixture policy is an owner-decision aid for local proof artifacts. It is not approval to publish generated fixtures as public binaries, not signed provenance, and not a license or privacy review.',
}

const PERFORMANCE_POLICY: ReleaseProofPerformancePolicy = {
	currentDecision: 'owner-approval-required',
	approvalChecklist: [
		{
			artifact: 'safe-open-proof',
			gateId: 'release-latency-run',
			ownerLoop: 'performance',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Approve a tracked-clean release-environment safe-open latency run over standardized public inputs and non-threshold wording.',
			acceptanceEvidence:
				'Timed safe-open proof uses approved public inputs, repeat/warmup policy, environment notes, and wording that reports observations without a release threshold.',
			rejectIf:
				'Run uses private corpora, dirty worktree state, one-off local timing, machine-specific ratios, or copy that implies a latency SLA.',
			validationCommand:
				'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'streaming-matrix-boundary',
			ownerLoop: 'performance',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Accept one representative dirty-sheet streaming proof for narrow wording, or require a broader streaming matrix before any parity claim.',
			acceptanceEvidence:
				'Package-action proof reports one streaming proof case with regenerated dirty worksheet and passthrough-byte equality, and release wording says representative proof only.',
			rejectIf:
				'Copy says full streaming parity, covers add/drop/error streaming behavior, or implies macro/chart streaming preservation without a broader matrix.',
			validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
		},
	],
	sourceReferences: [
		{
			label: 'Bun benchmarking',
			url: 'https://bun.sh/docs/project/benchmarking',
		},
		{
			label: 'hyperfine benchmarking',
			url: 'https://github.com/sharkdp/hyperfine',
		},
		{
			label: 'hyperfine manual',
			url: 'https://man.archlinux.org/man/hyperfine.1.en',
		},
	],
	boundary:
		'Performance policy is an owner-decision aid for local benchmark evidence. It is not a release performance threshold, SLA, streaming parity claim, or production optimization mandate.',
}

const CORRECTNESS_POLICY: ReleaseProofCorrectnessPolicy = {
	currentDecision: 'owner-approval-required',
	approvalChecklist: [
		{
			artifact: 'package-action-proof',
			gateId: 'unsupported-feature-boundary',
			ownerLoop: 'correctness',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Approve package-action unsupported-feature wording as per-part accounting and fail-closed evidence, not semantic support for unsupported workbook features.',
			acceptanceEvidence:
				'Package-action proof covers signatures, calc chain, chart/drawing sidecars, macros/ActiveX, unknown parts, and representative streaming scope with allowed/forbidden wording.',
			rejectIf:
				'Copy claims chart XML byte passthrough, signature preservation or verification, Excel-fresh cached formulas, macro/ActiveX safety, unknown-part understanding, or full streaming parity.',
			validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
		},
	],
	unsupportedFeatureBoundaries: [
		{
			feature: 'digital-signatures',
			currentEvidence:
				'Generated signature-invalidation case reports signature origin/signature XML as dropped after package mutation.',
			allowedWording:
				'Ascend detects signature package parts and reports invalidation/drop evidence when a write changes the package.',
			forbiddenWording: 'Ascend preserves, verifies, re-signs, or attests signatures.',
		},
		{
			feature: 'calc-chain',
			currentEvidence:
				'Public calc-chain fixture reports xl/calcChain.xml drop for formula topology edits.',
			allowedWording:
				'Ascend reports calc-chain drop/regeneration decisions when edits make cached calculation order unsafe.',
			forbiddenWording:
				'Ascend proves Excel recalculation equivalence or cached formula freshness.',
		},
		{
			feature: 'chart-drawing-sidecars',
			currentEvidence:
				'Public chart fixture accounts for drawing sidecars separately while chart XML can regenerate.',
			allowedWording:
				'Ascend accounts for chart/drawing sidecars separately from regenerated workbook parts.',
			forbiddenWording: 'Chart XML is byte-passthrough or every chart feature is understood.',
		},
		{
			feature: 'macros-activex',
			currentEvidence:
				'Public macro fixture reports macro-bearing parts as package evidence; safe-open routes macro/ActiveX to review.',
			allowedWording: 'Ascend records macro/ActiveX package preservation and review routing.',
			forbiddenWording:
				'Macros or ActiveX are safe, sandboxed, scanned, or executable through Ascend.',
		},
		{
			feature: 'unknown-parts',
			currentEvidence:
				'Generated unknown-part case reports one error action and failed post-write audit.',
			allowedWording: 'Ascend can fail closed with an explicit unknown-part error.',
			forbiddenWording: 'Ascend preserves or understands arbitrary unknown parts.',
		},
		{
			feature: 'streaming-scope',
			currentEvidence:
				'One representative streaming dirty-sheet proof reports a regenerated worksheet and passthrough-byte equality.',
			allowedWording: 'One representative streaming writer package-action proof exists.',
			forbiddenWording: 'Full streaming parity across all package-action scenarios.',
		},
	],
	sourceReferences: [
		{
			label: 'OOXML calculation chain',
			url: 'https://ooxml.info/docs/12/12.3/12.3.1/',
		},
		{
			label: 'Microsoft macro security',
			url: 'https://support.microsoft.com/en-gb/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373',
		},
		{
			label: 'Microsoft ActiveX settings',
			url: 'https://support.microsoft.com/en-us/office/enable-or-disable-activex-settings-in-office-files-f1303e08-a3f8-41c5-a17e-b0b8898743ed',
		},
		{
			label: 'SheetJS VBA blobs',
			url: 'https://docs.sheetjs.com/docs/csf/features/vba/',
		},
		{
			label: 'OOXML digital signatures',
			url: 'https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EHROM.html',
		},
	],
	boundary:
		'Correctness policy is an owner-decision aid for claim wording. It does not prove semantic support for unsupported workbook features, Excel recalculation equivalence, macro/ActiveX safety, signature verification, or streaming parity.',
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
		fixturePolicy: cloneFixturePolicy(),
		performancePolicy: clonePerformancePolicy(),
		correctnessPolicy: cloneCorrectnessPolicy(),
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
		'## Next Owner Actions',
		'',
		'| Rank | Artifact | Gate | Owner loop | Priority | Next step | Acceptance evidence | Forbidden shortcut |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- |',
		...result.readiness.nextOwnerActions.map(nextOwnerActionMarkdownRow),
		'',
		'## Fixture Policy',
		'',
		`Current decision: ${result.fixturePolicy.currentDecision}`,
		`Tracked fixture scans: ${formatFixturePolicyCommands(result.fixturePolicy.trackedFixtureScanCommands)}`,
		`Generated structural cases: ${formatGeneratedStructuralCases(result.fixturePolicy.currentGeneratedStructuralCases)}`,
		result.fixturePolicy.boundary,
		'',
		'Generated structural fixtures are allowed only when:',
		...result.fixturePolicy.generatedStructuralFixturesAllowedWhen.map((entry) => `- ${entry}`),
		'',
		'Public binary fixtures are required when:',
		...result.fixturePolicy.publicBinaryFixturesRequiredWhen.map((entry) => `- ${entry}`),
		'',
		'Approval checklist:',
		'',
		'| Artifact | Gate | Owner | Status | Decision needed | Acceptance evidence | Reject if | Validation command |',
		'| --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.fixturePolicy.approvalChecklist.map(fixturePolicyApprovalMarkdownRow),
		'',
		'Source references:',
		...result.fixturePolicy.sourceReferences.map(
			(reference) => `- ${reference.label}: ${reference.url}`,
		),
		'',
		'## Performance Policy',
		'',
		`Current decision: ${result.performancePolicy.currentDecision}`,
		result.performancePolicy.boundary,
		'',
		'Approval checklist:',
		'',
		'| Artifact | Gate | Owner | Status | Decision needed | Acceptance evidence | Reject if | Validation command |',
		'| --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.performancePolicy.approvalChecklist.map(performancePolicyApprovalMarkdownRow),
		'',
		'Source references:',
		...result.performancePolicy.sourceReferences.map(
			(reference) => `- ${reference.label}: ${reference.url}`,
		),
		'',
		'## Correctness Policy',
		'',
		`Current decision: ${result.correctnessPolicy.currentDecision}`,
		result.correctnessPolicy.boundary,
		'',
		'Unsupported feature boundaries:',
		'',
		'| Feature | Current evidence | Allowed wording | Forbidden wording |',
		'| --- | --- | --- | --- |',
		...result.correctnessPolicy.unsupportedFeatureBoundaries.map(correctnessBoundaryMarkdownRow),
		'',
		'Approval checklist:',
		'',
		'| Artifact | Gate | Owner | Status | Decision needed | Acceptance evidence | Reject if | Validation command |',
		'| --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.correctnessPolicy.approvalChecklist.map(correctnessPolicyApprovalMarkdownRow),
		'',
		'Source references:',
		...result.correctnessPolicy.sourceReferences.map(
			(reference) => `- ${reference.label}: ${reference.url}`,
		),
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
		fixturePolicy: cloneFixturePolicy(),
		performancePolicy: clonePerformancePolicy(),
		correctnessPolicy: cloneCorrectnessPolicy(),
		nextOwnerActions: result.readiness.nextOwnerActions,
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
				evidence:
					'safe-open fixture scan over tracked fixtures currently finds no public signed/unknown binary replacements',
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
					'package-action fixture scan over tracked fixtures finds public docProps/calc-chain/customXml/macro/chart candidates but 0 signature/unknown-path replacements; current proof still uses generated signature-invalidation and unknown-part edge packages',
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
			blockingActions: artifactActions.map(cloneNextOwnerAction),
			nextStepKinds: uniqueNextStepKinds(artifactActions.map((action) => action.nextStepKind)),
			boundary:
				'Owner handoff for proof, validation, boundary approval, and publication policy only; it is not permission to add new SDK, CLI, API, or MCP surfaces.',
		}
	})
}

function cloneNextOwnerAction(action: ReleaseProofNextOwnerAction): ReleaseProofNextOwnerAction {
	return { ...action }
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
					'Public and generated package cases covering public docProps passthrough, worksheet regeneration, sheet add, public calc-chain drop, signature invalidation, macro/sidecar accounting, chart sidecar accounting, and unknown-part error.',
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
				acceptanceEvidence:
					artifact === 'safe-open-proof'
						? 'Product accepts disclosed generated signed/unknown structural packages for guarded topology proof, or replaces them with approved public binary fixtures.'
						: 'Product accepts disclosed generated signature/unknown structural packages for guarded package-action proof, or replaces them with approved public binary fixtures.',
				forbiddenShortcut:
					'Do not hide generated fixture provenance, vendor license-unclear binaries, or imply real-world trust behavior from structural package topology alone.',
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
				acceptanceEvidence:
					'Correctness approves allowed/forbidden wording for signatures, calc chains, chart/drawing sidecars, macros/ActiveX, unknown parts, and streaming scope.',
				forbiddenShortcut:
					'Do not describe chart XML as byte-passthrough, signatures as preserved or verified, cached formulas as Excel-fresh, or unknown parts as understood.',
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
				acceptanceEvidence:
					'Performance reruns tracked-clean release-environment open-plan evidence over standardized public inputs and approves non-threshold wording.',
				forbiddenShortcut:
					'Do not use dirty-worktree, private-corpus, one-off local timing, or machine-specific ratios as release performance claims.',
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
				acceptanceEvidence:
					'Performance accepts one representative dirty-sheet streaming proof for narrow wording, or expands the matrix to add/drop/error and public macro/chart cases.',
				forbiddenShortcut:
					'Do not call one representative streaming case full streaming parity or imply streaming coverage for every package-action scenario.',
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
				acceptanceEvidence:
					artifact === 'safe-open-proof'
						? 'Release approves safe-open boundary language that excludes malware scanning, sandboxing, file trust, active-content safety, signed provenance, and malformed-package recovery.'
						: 'Release approves local package-action proof wording below SLSA, in-toto, Sigstore, GitHub artifact attestation, and signed-provenance thresholds.',
				forbiddenShortcut:
					'Do not call local digests signed provenance, tamper-evident storage, attestation, malware safety, or active-content trust.',
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
				acceptanceEvidence:
					'Release defines artifact storage path, retention/privacy filtering, canonicalization subject, and verification expectations for compact reports.',
				forbiddenShortcut:
					'Do not index or publish compact report digests before storage, privacy, canonicalization, and verification policy exists.',
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
				acceptanceEvidence:
					'Owner supplies explicit acceptance evidence or replaces this requirement with a classified release-readiness gate.',
				forbiddenShortcut: 'Do not treat an unclassified missing requirement as satisfied.',
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
				`${action.rank}:${action.artifact}/${action.requirementId}(${action.ownerLoop},${action.priority},${action.nextStepKind};accept=${action.acceptanceEvidence};forbid=${action.forbiddenShortcut})`,
		)
		.join('; ')
}

function nextOwnerActionMarkdownRow(action: ReleaseProofNextOwnerAction): string {
	return [
		String(action.rank),
		action.artifact,
		action.requirementId,
		action.ownerLoop,
		action.priority,
		action.nextStepKind,
		action.acceptanceEvidence,
		action.forbiddenShortcut,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
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

function cloneFixturePolicy(): ReleaseProofFixturePolicy {
	return {
		...FIXTURE_POLICY,
		generatedStructuralFixturesAllowedWhen: [
			...FIXTURE_POLICY.generatedStructuralFixturesAllowedWhen,
		],
		publicBinaryFixturesRequiredWhen: [...FIXTURE_POLICY.publicBinaryFixturesRequiredWhen],
		approvalChecklist: FIXTURE_POLICY.approvalChecklist.map((item) => ({ ...item })),
		trackedFixtureScanCommands: { ...FIXTURE_POLICY.trackedFixtureScanCommands },
		currentGeneratedStructuralCases: Object.fromEntries(
			Object.entries(FIXTURE_POLICY.currentGeneratedStructuralCases).map(([artifact, cases]) => [
				artifact,
				[...cases],
			]),
		) as ReleaseProofFixturePolicy['currentGeneratedStructuralCases'],
		sourceReferences: FIXTURE_POLICY.sourceReferences.map((reference) => ({ ...reference })),
	}
}

function fixturePolicyApprovalMarkdownRow(row: ReleaseProofFixturePolicyApprovalItem): string {
	return [
		row.artifact,
		row.gateId,
		row.ownerLoop,
		row.status,
		row.decisionNeeded,
		row.acceptanceEvidence,
		row.rejectIf,
		`\`${row.validationCommand}\``,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function clonePerformancePolicy(): ReleaseProofPerformancePolicy {
	return {
		...PERFORMANCE_POLICY,
		approvalChecklist: PERFORMANCE_POLICY.approvalChecklist.map((item) => ({ ...item })),
		sourceReferences: PERFORMANCE_POLICY.sourceReferences.map((reference) => ({ ...reference })),
	}
}

function performancePolicyApprovalMarkdownRow(
	row: ReleaseProofPerformancePolicyApprovalItem,
): string {
	return [
		row.artifact,
		row.gateId,
		row.ownerLoop,
		row.status,
		row.decisionNeeded,
		row.acceptanceEvidence,
		row.rejectIf,
		`\`${row.validationCommand}\``,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneCorrectnessPolicy(): ReleaseProofCorrectnessPolicy {
	return {
		...CORRECTNESS_POLICY,
		approvalChecklist: CORRECTNESS_POLICY.approvalChecklist.map((item) => ({ ...item })),
		unsupportedFeatureBoundaries: CORRECTNESS_POLICY.unsupportedFeatureBoundaries.map(
			(boundary) => ({ ...boundary }),
		),
		sourceReferences: CORRECTNESS_POLICY.sourceReferences.map((reference) => ({
			...reference,
		})),
	}
}

function correctnessBoundaryMarkdownRow(row: ReleaseProofUnsupportedFeatureBoundary): string {
	return [row.feature, row.currentEvidence, row.allowedWording, row.forbiddenWording]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function correctnessPolicyApprovalMarkdownRow(
	row: ReleaseProofCorrectnessPolicyApprovalItem,
): string {
	return [
		row.artifact,
		row.gateId,
		row.ownerLoop,
		row.status,
		row.decisionNeeded,
		row.acceptanceEvidence,
		row.rejectIf,
		`\`${row.validationCommand}\``,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function formatFixturePolicyCommands(
	commands: Readonly<Record<ReleaseProofIndexArtifactName, string>>,
): string {
	return (['safe-open-proof', 'package-action-proof'] as const)
		.map((artifact) => `${artifact}=\`${commands[artifact]}\``)
		.join('; ')
}

function formatGeneratedStructuralCases(
	casesByArtifact: Readonly<Record<ReleaseProofIndexArtifactName, readonly string[]>>,
): string {
	return (['safe-open-proof', 'package-action-proof'] as const)
		.map((artifact) => `${artifact}=${casesByArtifact[artifact].join(',') || 'none'}`)
		.join('; ')
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
