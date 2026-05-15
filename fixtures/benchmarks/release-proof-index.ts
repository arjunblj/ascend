import { createHash } from 'node:crypto'
import { basename } from 'node:path'
import {
	type PackageActionFixtureScanResult,
	runPackageActionFixtureScan,
} from './package-action-fixture-scan.ts'
import {
	type PackageActionCompactReleaseReport,
	type PackageActionProofCaseResult,
	type PackageActionProofResult,
	packageActionCompactReleaseReport,
	packageActionProofMarkdown,
	runPackageActionProof,
} from './package-action-proof.ts'
import { runSafeOpenFixtureScan, type SafeOpenFixtureScanResult } from './safe-open-fixture-scan.ts'
import {
	runSafeOpenProof,
	type SafeOpenCompactReleaseReport,
	type SafeOpenProofCaseResult,
	type SafeOpenProofResult,
	safeOpenCompactReleaseReport,
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
export type ReleaseProofPortfolioClaimName =
	| ReleaseProofIndexArtifactName
	| ReleaseProofDeferredClaimName
	| 'release-proof-bundle'
	| 'property-journal-laws'
export type ReleaseProofPortfolioClaimStatus =
	| 'claim-wording-allowed-today'
	| 'needs-one-more-fold-in'
	| 'speculative-do-not-promote'
export type ReleaseProofPortfolioHandoffDecision =
	| 'top-implementation-handoff'
	| 'proof-packaging-only'
	| 'do-not-promote-yet'

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

export interface ReleaseProofCompactReportPublicationEvidence {
	readonly ownerLoop: 'release'
	readonly status: 'local-summary-present-publication-policy-required'
	readonly ownerApprovalRequired: true
	readonly compactReportDigestsIndexed: false
	readonly allCompactCommandsPresent: boolean
	readonly compactReportsEmbedForbiddenPayloadFields: boolean
	readonly generatedAtIncluded: boolean
	readonly missingPolicyRequirements: readonly string[]
	readonly policyDecisions: readonly ReleaseProofCompactReportPublicationPolicyDecision[]
	readonly reports: readonly ReleaseProofCompactReportPublicationEvidenceItem[]
	readonly boundary: string
}

export interface ReleaseProofCompactReportPublicationPolicyDecision {
	readonly requirement: string
	readonly ownerLoop: 'release'
	readonly status: 'pending-owner-decision'
	readonly decisionNeeded: string
	readonly acceptanceEvidence: string
	readonly rejectIf: string
}

export type ReleaseProofPackageActionKind = 'passthrough' | 'regenerate' | 'add' | 'drop' | 'error'

export interface ReleaseProofStreamingMatrixEvidence {
	readonly artifact: 'package-action-proof'
	readonly gateId: 'streaming-matrix-boundary'
	readonly ownerLoop: 'performance'
	readonly status: 'representative-proof-present-owner-approval-required'
	readonly ownerApprovalRequired: true
	readonly validationCommand: string
	readonly representativeProofCases: number
	readonly streamingRegenerateParts: number
	readonly coveredActionKinds: readonly ReleaseProofPackageActionKind[]
	readonly missingActionKinds: readonly ReleaseProofPackageActionKind[]
	readonly coveredCaseNames: readonly string[]
	readonly nonStreamingCaseNames: readonly string[]
	readonly publicNonStreamingCaseNames: readonly string[]
	readonly generatedNonStreamingCaseNames: readonly string[]
	readonly streamingIssueCaseNames: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofCompactReportPublicationEvidenceItem {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly gateId: 'compact-report-publication-policy'
	readonly command: string
	readonly jsonBytes: number
	readonly topLevelFields: readonly string[]
	readonly forbiddenPayloadFieldsPresent: readonly string[]
	readonly readyWhenGatePresent: boolean
	readonly generatedAtIncluded: boolean
	readonly headlineClaimAllowed: false
	readonly releaseGate: 'blocked-by-publication-policy'
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
	readonly fixturePolicyEvidence: ReleaseProofFixturePolicyEvidence
	readonly generatedFixtureDecisionEvidence: ReleaseProofGeneratedFixtureDecisionEvidence
	readonly performancePolicy: ReleaseProofPerformancePolicy
	readonly safeOpenLatencyValidationEvidence: ReleaseProofSafeOpenLatencyValidationEvidence
	readonly correctnessPolicy: ReleaseProofCorrectnessPolicy
	readonly correctnessBoundaryEvidence: ReleaseProofCorrectnessBoundaryEvidence
	readonly releasePackageabilityEvidence: ReleaseProofPackageabilityEvidence
	readonly streamingMatrixEvidence: ReleaseProofStreamingMatrixEvidence
	readonly compactReportPublicationEvidence: ReleaseProofCompactReportPublicationEvidence
	readonly readiness: ReleaseProofReadinessSummary
	readonly boundary: string
	readonly claimPortfolio: readonly ReleaseProofPortfolioClaim[]
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
	readonly fixturePolicyEvidence: ReleaseProofFixturePolicyEvidence
	readonly generatedFixtureDecisionEvidence: ReleaseProofGeneratedFixtureDecisionEvidence
	readonly performancePolicy: ReleaseProofPerformancePolicy
	readonly safeOpenLatencyValidationEvidence: ReleaseProofSafeOpenLatencyValidationEvidence
	readonly correctnessPolicy: ReleaseProofCorrectnessPolicy
	readonly correctnessBoundaryEvidence: ReleaseProofCorrectnessBoundaryEvidence
	readonly releasePackageabilityEvidence: ReleaseProofPackageabilityEvidence
	readonly streamingMatrixEvidence: ReleaseProofStreamingMatrixEvidence
	readonly compactReportPublicationEvidence: ReleaseProofCompactReportPublicationEvidence
	readonly nextOwnerActions: readonly ReleaseProofNextOwnerAction[]
	readonly claimBlockerBoard: readonly ReleaseProofClaimBlockerBoardRow[]
	readonly implementationHandoffs: readonly ReleaseProofImplementationHandoff[]
	readonly claimPortfolio: readonly ReleaseProofPortfolioClaim[]
	readonly deferredClaims: readonly ReleaseProofDeferredClaim[]
	readonly excludedEvidence: readonly ReleaseProofIndexExcludedEvidence[]
	readonly boundary: string
}

export interface ReleaseProofPackageabilityEvidence {
	readonly ownerLoop: 'release'
	readonly status: 'local-tarball-smokes-present-publication-policy-required'
	readonly ownerApprovalRequired: true
	readonly sdkSmokeCommand: string
	readonly appSmokeCommand: string
	readonly coveredEvidence: readonly string[]
	readonly missingPolicyRequirements: readonly string[]
	readonly forbiddenClaims: readonly string[]
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
	readonly claimBlockerBoard: readonly ReleaseProofClaimBlockerBoardRow[]
	readonly implementationHandoffs: readonly ReleaseProofImplementationHandoff[]
	readonly boundary: string
}

export interface ReleaseProofClaimBlockerBoardRow {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly claim: string
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly blockerCount: number
	readonly requirementIds: readonly string[]
	readonly actionRanks: readonly number[]
	readonly nextStepKinds: readonly ReleaseProofNextOwnerAction['nextStepKind'][]
	readonly acceptanceEvidence: readonly string[]
	readonly forbiddenShortcuts: readonly string[]
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

export interface ReleaseProofPortfolioClaim {
	readonly rank: number
	readonly name: ReleaseProofPortfolioClaimName
	readonly claim: string
	readonly northStarLink: string
	readonly status: ReleaseProofPortfolioClaimStatus
	readonly evidenceNeeded: ReleaseProofClaimProofRequired
	readonly killCriterion: string
	readonly likelyHandoffOwner: readonly ReleaseProofReadinessOwner[]
	readonly handoffDecision: ReleaseProofPortfolioHandoffDecision
	readonly proofCommand?: string
	readonly boundary: string
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
	readonly safeOpenFixtureAcceptanceChecklist: readonly ReleaseProofSafeOpenFixtureAcceptanceItem[]
	readonly packageActionFixtureAcceptanceChecklist: readonly ReleaseProofPackageActionFixtureAcceptanceItem[]
	readonly approvalChecklist: readonly ReleaseProofFixturePolicyApprovalItem[]
	readonly trackedFixtureScanCommands: Readonly<Record<ReleaseProofIndexArtifactName, string>>
	readonly currentGeneratedStructuralCases: Readonly<
		Record<ReleaseProofIndexArtifactName, readonly string[]>
	>
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly boundary: string
}

export interface ReleaseProofSafeOpenFixtureAcceptanceItem {
	readonly caseName: 'signed' | 'unknown-part' | 'malformed'
	readonly generatedCaseKind: 'generated-edge-package' | 'generated-malformed-package'
	readonly acceptableAsTopologyProofWhen: string
	readonly requiresPublicBinaryWhen: string
	readonly validationCommand: string
	readonly gateEffect: 'keeps-public-edge-fixtures-missing-until-owner-approval'
}

export interface ReleaseProofPackageActionFixtureAcceptanceItem {
	readonly caseName: 'signature-invalidation-drop' | 'unknown-part-error'
	readonly generatedCaseKind: 'generated-edge-package'
	readonly acceptableAsPackageActionProofWhen: string
	readonly requiresPublicBinaryWhen: string
	readonly forbiddenClaim: string
	readonly validationCommand: string
	readonly gateEffect: 'keeps-edge-fixture-policy-missing-until-owner-approval'
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

export interface ReleaseProofFixturePolicyEvidence {
	readonly ownerLoop: 'product'
	readonly status: 'tracked-scan-complete-owner-approval-required'
	readonly ownerApprovalRequired: true
	readonly allScansUseTrackedCorpus: boolean
	readonly publicReplacementGapsRemain: boolean
	readonly safeOpen: ReleaseProofSafeOpenFixturePolicyEvidence
	readonly packageAction: ReleaseProofPackageActionFixturePolicyEvidence
	readonly boundary: string
}

export interface ReleaseProofGeneratedFixtureDecisionEvidence {
	readonly ownerLoop: 'product'
	readonly status: 'generated-structural-cases-disclosed-owner-approval-required'
	readonly ownerApprovalRequired: true
	readonly allGeneratedStructuralCasesDisclosed: boolean
	readonly publicReplacementGapsRemain: boolean
	readonly validationCommand: string
	readonly cases: readonly ReleaseProofGeneratedFixtureDecisionCase[]
	readonly boundary: string
}

export interface ReleaseProofGeneratedFixtureDecisionCase {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly gateId: 'public-edge-fixtures' | 'edge-fixture-policy'
	readonly caseName: string
	readonly generatedKind: 'generated-edge-package' | 'generated-malformed-package'
	readonly replacementEvidence: string
	readonly ownerDecisionNeeded: string
	readonly recommendedOwnerAction: string
	readonly allowedUse: string
	readonly forbiddenUse: string
}

export interface ReleaseProofSafeOpenFixturePolicyEvidence {
	readonly artifact: 'safe-open-proof'
	readonly gateId: 'public-edge-fixtures'
	readonly validationCommand: string
	readonly corpus: SafeOpenFixtureScanResult['corpus']
	readonly scanned: number
	readonly rejected: number
	readonly replacementStatus: SafeOpenFixtureScanResult['replacementStatus']
	readonly riskFamilyCounts: Readonly<Record<string, number>>
	readonly signatureOrUnknownMatches: number
	readonly currentGeneratedStructuralCases: readonly string[]
	readonly externalCandidateEvidence: readonly ReleaseProofExternalFixtureCandidateEvidence[]
	readonly boundary: string
}

export interface ReleaseProofExternalFixtureCandidateEvidence {
	readonly artifact: 'safe-open-proof'
	readonly gateId: 'public-edge-fixtures'
	readonly caseName: 'unknown-part'
	readonly status: 'external-candidate-owner-review-required'
	readonly candidateId: string
	readonly repositoryUrl: string
	readonly sourceUrl: string
	readonly licenseEvidenceUrl: string
	readonly license: string
	readonly sha256: string
	readonly packageManifestSha256: string
	readonly recommendedMode: string
	readonly reviewBeforeHydration: boolean
	readonly riskFamily: string
	readonly partCount: number
	readonly relationshipCount: number
	readonly sampleUnknownPart: string
	readonly ownerDecisionNeeded: string
	readonly gateEffect: 'does-not-satisfy-public-edge-fixtures'
	readonly boundary: string
}

export interface ReleaseProofPackageActionFixturePolicyEvidence {
	readonly artifact: 'package-action-proof'
	readonly gateId: 'edge-fixture-policy'
	readonly validationCommand: string
	readonly corpus: PackageActionFixtureScanResult['corpus']
	readonly scanned: number
	readonly rejected: number
	readonly replacementStatus: PackageActionFixtureScanResult['replacementStatus']
	readonly featureCounts: PackageActionFixtureScanResult['featureCounts']
	readonly currentGeneratedStructuralCases: readonly string[]
	readonly missingReplacementFeatures: readonly string[]
	readonly boundary: string
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

export interface ReleaseProofSafeOpenLatencyValidationEvidence {
	readonly artifact: 'safe-open-proof'
	readonly gateId: 'release-latency-run'
	readonly ownerLoop: 'performance'
	readonly status:
		| 'timed-evidence-absent-owner-run-required'
		| 'local-timed-diagnostic-owner-approval-required'
	readonly ownerApprovalRequired: true
	readonly releaseClaimAllowed: false
	readonly thresholdClaimAllowed: false
	readonly validationCommand: string
	readonly repeat: number
	readonly warmup: number
	readonly timedCaseCount: number
	readonly publicTimedCaseNames: readonly string[]
	readonly generatedTimedCaseNames: readonly string[]
	readonly malformedRejected: boolean
	readonly publicOpenPlanMedianMs: Readonly<Record<string, number>>
	readonly publicFullOpenRatio: Readonly<Record<string, number>>
	readonly missingPolicyRequirements: readonly string[]
	readonly boundary: string
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

export interface ReleaseProofCorrectnessBoundaryEvidence {
	readonly artifact: 'package-action-proof'
	readonly gateId: 'unsupported-feature-boundary'
	readonly ownerLoop: 'correctness'
	readonly status: 'evidence-present-owner-approval-required'
	readonly allCurrentEvidencePresent: boolean
	readonly missingFeatureNames: readonly string[]
	readonly ownerEscalationRequired: boolean
	readonly ownerApprovalRequired: true
	readonly validationCommand: string
	readonly featureChecks: readonly ReleaseProofCorrectnessBoundaryFeatureCheck[]
	readonly boundary: string
}

export interface ReleaseProofCorrectnessBoundaryFeatureCheck {
	readonly feature: string
	readonly evidencePresent: boolean
	readonly evidenceSources: readonly string[]
	readonly proofChecks: readonly string[]
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
	safeOpenFixtureAcceptanceChecklist: [
		{
			caseName: 'signed',
			generatedCaseKind: 'generated-edge-package',
			acceptableAsTopologyProofWhen:
				'Owner accepts package-topology evidence for signature-related parts only, with generated provenance disclosed and no signature verification or trust wording.',
			requiresPublicBinaryWhen:
				'Claim wording depends on real signed workbook behavior, signature validity, vendor repair UX, or user trust.',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			gateEffect: 'keeps-public-edge-fixtures-missing-until-owner-approval',
		},
		{
			caseName: 'unknown-part',
			generatedCaseKind: 'generated-edge-package',
			acceptableAsTopologyProofWhen:
				'Owner accepts package-topology evidence that unknown package features route to review before hydration, with no preservation or understanding claim.',
			requiresPublicBinaryWhen:
				'Claim wording depends on arbitrary third-party unknown-part preservation, vendor behavior, or real-world workbook semantics.',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			gateEffect: 'keeps-public-edge-fixtures-missing-until-owner-approval',
		},
		{
			caseName: 'malformed',
			generatedCaseKind: 'generated-malformed-package',
			acceptableAsTopologyProofWhen:
				'Owner accepts generated bad bytes as fail-closed rejection-path evidence only.',
			requiresPublicBinaryWhen:
				'Claim wording depends on vendor repair equivalence, recovery of arbitrary malformed files, or real-world corrupt workbook provenance.',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
			gateEffect: 'keeps-public-edge-fixtures-missing-until-owner-approval',
		},
	],
	packageActionFixtureAcceptanceChecklist: [
		{
			caseName: 'signature-invalidation-drop',
			generatedCaseKind: 'generated-edge-package',
			acceptableAsPackageActionProofWhen:
				'Owner accepts generated signature package topology as local evidence that generated edits drop or invalidate signature-related parts.',
			requiresPublicBinaryWhen:
				'Claim wording depends on a real signed workbook, signature validity, re-signing behavior, vendor trust UX, or real-world authoring provenance.',
			forbiddenClaim:
				'Do not claim signature preservation, verification, re-signing, attestation, SLSA, in-toto, or signed provenance.',
			validationCommand: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			gateEffect: 'keeps-edge-fixture-policy-missing-until-owner-approval',
		},
		{
			caseName: 'unknown-part-error',
			generatedCaseKind: 'generated-edge-package',
			acceptableAsPackageActionProofWhen:
				'Owner accepts generated unknown package topology as local fail-closed package-action evidence with an explicit error action.',
			requiresPublicBinaryWhen:
				'Claim wording depends on arbitrary third-party unknown-part preservation, understanding, recovery, or real workbook semantics.',
			forbiddenClaim:
				'Do not claim arbitrary unknown-part preservation, understanding, safe recovery, trust, or signed provenance.',
			validationCommand: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			gateEffect: 'keeps-edge-fixture-policy-missing-until-owner-approval',
		},
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

const SAFE_OPEN_EXTERNAL_FIXTURE_CANDIDATES: readonly ReleaseProofExternalFixtureCandidateEvidence[] =
	[
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			caseName: 'unknown-part',
			status: 'external-candidate-owner-review-required',
			candidateId: 'excelforge-book1-unknown-part',
			repositoryUrl: 'https://github.com/node-projects/excelForge',
			sourceUrl:
				'https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx',
			licenseEvidenceUrl:
				'https://raw.githubusercontent.com/node-projects/excelForge/master/package.json',
			license: 'MIT',
			sha256: '9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122',
			packageManifestSha256: 'cae1feec581eed864255cff45fa23a7e2c085cb0f2c2628d1a0187fc39de3ef7',
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			riskFamily: 'preservedOther',
			partCount: 50,
			relationshipCount: 37,
			sampleUnknownPart: 'docMetadata/LabelInfo.xml',
			ownerDecisionNeeded:
				'Product/release must decide whether to vendor this MIT-package-manifest-backed workbook as a public unknown-part fixture with attribution policy.',
			gateEffect: 'does-not-satisfy-public-edge-fixtures',
			boundary:
				'External candidate evidence is a pointer for owner review only. The workbook is not vendored, attribution policy is not approved, and this does not address the signed-workbook fixture gap.',
		},
	]

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
				'Accept representative streaming proofs covering passthrough/regenerate/add/drop for narrow wording, or require a broader streaming matrix before any parity claim.',
			acceptanceEvidence:
				'Package-action proof reports three streaming proof cases covering passthrough/regenerate/add/drop with passthrough-byte equality, and release wording says representative proofs only.',
			rejectIf:
				'Copy says full streaming parity, covers error streaming behavior, or implies macro/chart streaming preservation without a broader matrix.',
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
	const safeOpenFixtureScan = runSafeOpenFixtureScan()
	const packageActionFixtureScan = runPackageActionFixtureScan()
	const safeOpenCompact = safeOpenCompactReleaseReport(safeOpen)
	const packageActionCompact = packageActionCompactReleaseReport(packageAction)
	const artifacts = [
		safeOpenArtifact(safeOpen, includeTimings),
		packageActionArtifact(packageAction, includeTimings),
	]
	const fixtureEvidence = fixturePolicyEvidence(safeOpenFixtureScan, packageActionFixtureScan)
	return {
		generatedAt: new Date().toISOString(),
		artifactCount: artifacts.length,
		excludedEvidenceCount: EXCLUDED_EVIDENCE.length,
		deferredClaimCount: DEFERRED_CLAIMS.length,
		signed: false,
		attestation: false,
		fixturePolicy: cloneFixturePolicy(),
		fixturePolicyEvidence: fixtureEvidence,
		generatedFixtureDecisionEvidence: generatedFixtureDecisionEvidence(fixtureEvidence),
		performancePolicy: clonePerformancePolicy(),
		safeOpenLatencyValidationEvidence: safeOpenLatencyValidationEvidence(safeOpen),
		correctnessPolicy: cloneCorrectnessPolicy(),
		correctnessBoundaryEvidence: correctnessBoundaryEvidence(safeOpen, packageAction),
		releasePackageabilityEvidence: releasePackageabilityEvidence(),
		streamingMatrixEvidence: streamingMatrixEvidence(packageAction),
		compactReportPublicationEvidence: compactReportPublicationEvidence(
			safeOpenCompact,
			packageActionCompact,
			artifacts,
		),
		readiness: releaseReadinessSummary(artifacts),
		boundary:
			'Digest index for local release evidence artifacts. This is not signed provenance, SLSA, in-toto attestation, or tamper-evident storage.',
		claimPortfolio: CLAIM_PORTFOLIO.map(clonePortfolioClaim),
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
		`Claim blocker board: ${formatClaimBlockerBoard(result.readiness.claimBlockerBoard)}`,
		`Implementation handoffs: ${formatImplementationHandoffs(result.readiness.implementationHandoffs)}`,
		result.readiness.boundary,
		'',
		'## Release Packageability Evidence',
		'',
		`Status: ${result.releasePackageabilityEvidence.status}`,
		`SDK smoke command: \`${result.releasePackageabilityEvidence.sdkSmokeCommand}\``,
		`App smoke command: \`${result.releasePackageabilityEvidence.appSmokeCommand}\``,
		`Owner approval required: ${result.releasePackageabilityEvidence.ownerApprovalRequired}`,
		result.releasePackageabilityEvidence.boundary,
		'',
		'Covered packageability evidence:',
		...result.releasePackageabilityEvidence.coveredEvidence.map((entry) => `- ${entry}`),
		'',
		'Missing packageability policy requirements:',
		...result.releasePackageabilityEvidence.missingPolicyRequirements.map((entry) => `- ${entry}`),
		'',
		'Forbidden packageability claims:',
		...result.releasePackageabilityEvidence.forbiddenClaims.map((entry) => `- ${entry}`),
		'',
		'## Claim Blocker Board',
		'',
		'| Artifact | Claim | Owner loop | Blockers | Next steps | Acceptance evidence | Forbidden shortcuts | Boundary |',
		'| --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.readiness.claimBlockerBoard.map(claimBlockerBoardMarkdownRow),
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
		'Safe-open generated case acceptance checklist:',
		'',
		'| Case | Generated kind | Acceptable as topology proof when | Requires public binary when | Validation command | Gate effect |',
		'| --- | --- | --- | --- | --- | --- |',
		...result.fixturePolicy.safeOpenFixtureAcceptanceChecklist.map(
			safeOpenFixtureAcceptanceMarkdownRow,
		),
		'',
		'Package-action generated case acceptance checklist:',
		'',
		'| Case | Generated kind | Acceptable as package-action proof when | Requires public binary when | Forbidden claim | Validation command | Gate effect |',
		'| --- | --- | --- | --- | --- | --- | --- |',
		...result.fixturePolicy.packageActionFixtureAcceptanceChecklist.map(
			packageActionFixtureAcceptanceMarkdownRow,
		),
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
		'Fixture policy evidence:',
		'',
		`Status: ${result.fixturePolicyEvidence.status}`,
		`All scans use tracked corpus: ${result.fixturePolicyEvidence.allScansUseTrackedCorpus}`,
		`Public replacement gaps remain: ${result.fixturePolicyEvidence.publicReplacementGapsRemain}`,
		`Owner approval required: ${result.fixturePolicyEvidence.ownerApprovalRequired}`,
		result.fixturePolicyEvidence.boundary,
		'',
		'| Artifact | Gate | Command | Corpus | Scanned | Rejected | Replacement status | Generated structural cases | Gap evidence | Boundary |',
		'| --- | --- | --- | --- | ---: | ---: | --- | --- | --- | --- |',
		fixturePolicyEvidenceMarkdownRow(result.fixturePolicyEvidence.safeOpen),
		fixturePolicyEvidenceMarkdownRow(result.fixturePolicyEvidence.packageAction),
		'',
		'External fixture candidates:',
		'',
		'| Artifact | Gate | Case | Status | Candidate | License | SHA-256 | Routing | Gate effect | Boundary |',
		'| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.fixturePolicyEvidence.safeOpen.externalCandidateEvidence.map(
			externalFixtureCandidateEvidenceMarkdownRow,
		),
		'',
		'Generated fixture decision evidence:',
		'',
		`Status: ${result.generatedFixtureDecisionEvidence.status}`,
		`All generated structural cases disclosed: ${result.generatedFixtureDecisionEvidence.allGeneratedStructuralCasesDisclosed}`,
		`Public replacement gaps remain: ${result.generatedFixtureDecisionEvidence.publicReplacementGapsRemain}`,
		`Owner approval required: ${result.generatedFixtureDecisionEvidence.ownerApprovalRequired}`,
		result.generatedFixtureDecisionEvidence.boundary,
		'',
		'| Artifact | Gate | Case | Kind | Replacement evidence | Owner decision needed | Recommended owner action | Allowed use | Forbidden use |',
		'| --- | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.generatedFixtureDecisionEvidence.cases.map(
			generatedFixtureDecisionEvidenceMarkdownRow,
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
		'Safe-open latency validation evidence:',
		'',
		`Status: ${result.safeOpenLatencyValidationEvidence.status}`,
		`Timed case count: ${result.safeOpenLatencyValidationEvidence.timedCaseCount}`,
		`Public timed cases: ${result.safeOpenLatencyValidationEvidence.publicTimedCaseNames.join(',') || 'none'}`,
		`Generated timed cases: ${result.safeOpenLatencyValidationEvidence.generatedTimedCaseNames.join(',') || 'none'}`,
		`Release claim allowed: ${result.safeOpenLatencyValidationEvidence.releaseClaimAllowed}`,
		`Threshold claim allowed: ${result.safeOpenLatencyValidationEvidence.thresholdClaimAllowed}`,
		`Owner approval required: ${result.safeOpenLatencyValidationEvidence.ownerApprovalRequired}`,
		result.safeOpenLatencyValidationEvidence.boundary,
		'',
		'Missing latency policy requirements:',
		...result.safeOpenLatencyValidationEvidence.missingPolicyRequirements.map(
			(entry) => `- ${entry}`,
		),
		'',
		'## Streaming Matrix Evidence',
		'',
		`Status: ${result.streamingMatrixEvidence.status}`,
		`Covered action kinds: ${result.streamingMatrixEvidence.coveredActionKinds.join(',') || 'none'}`,
		`Missing action kinds: ${result.streamingMatrixEvidence.missingActionKinds.join(',') || 'none'}`,
		`Covered cases: ${result.streamingMatrixEvidence.coveredCaseNames.join(',') || 'none'}`,
		`Non-streaming cases: ${result.streamingMatrixEvidence.nonStreamingCaseNames.join(',') || 'none'}`,
		`Public non-streaming cases: ${result.streamingMatrixEvidence.publicNonStreamingCaseNames.join(',') || 'none'}`,
		`Owner approval required: ${result.streamingMatrixEvidence.ownerApprovalRequired}`,
		result.streamingMatrixEvidence.boundary,
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
		'Correctness boundary evidence:',
		'',
		`Status: ${result.correctnessBoundaryEvidence.status}`,
		`All current evidence present: ${result.correctnessBoundaryEvidence.allCurrentEvidencePresent}`,
		`Missing feature names: ${result.correctnessBoundaryEvidence.missingFeatureNames.join(',') || 'none'}`,
		`Owner escalation required: ${result.correctnessBoundaryEvidence.ownerEscalationRequired}`,
		`Owner approval required: ${result.correctnessBoundaryEvidence.ownerApprovalRequired}`,
		result.correctnessBoundaryEvidence.boundary,
		'',
		'| Feature | Evidence present | Sources | Checks | Allowed wording | Forbidden wording |',
		'| --- | --- | --- | --- | --- | --- |',
		...result.correctnessBoundaryEvidence.featureChecks.map(correctnessBoundaryEvidenceMarkdownRow),
		'',
		'## Compact Report Publication Evidence',
		'',
		`Status: ${result.compactReportPublicationEvidence.status}`,
		`Compact report digests indexed: ${result.compactReportPublicationEvidence.compactReportDigestsIndexed}`,
		`All compact commands present: ${result.compactReportPublicationEvidence.allCompactCommandsPresent}`,
		`Forbidden payload fields embedded: ${result.compactReportPublicationEvidence.compactReportsEmbedForbiddenPayloadFields}`,
		`GeneratedAt included: ${result.compactReportPublicationEvidence.generatedAtIncluded}`,
		`Owner approval required: ${result.compactReportPublicationEvidence.ownerApprovalRequired}`,
		result.compactReportPublicationEvidence.boundary,
		'',
		'Missing publication policy requirements:',
		...result.compactReportPublicationEvidence.missingPolicyRequirements.map(
			(entry) => `- ${entry}`,
		),
		'',
		'Publication policy decisions:',
		'',
		'| Requirement | Owner | Status | Decision needed | Acceptance evidence | Reject if |',
		'| --- | --- | --- | --- | --- | --- |',
		...result.compactReportPublicationEvidence.policyDecisions.map(
			compactReportPublicationPolicyDecisionMarkdownRow,
		),
		'',
		'| Artifact | Gate | Command | JSON bytes | Top-level fields | Forbidden payload fields | ReadyWhen gate present | GeneratedAt included | Headline claim allowed | Release gate | Boundary |',
		'| --- | --- | --- | ---: | --- | --- | --- | --- | --- | --- | --- |',
		...result.compactReportPublicationEvidence.reports.map(
			compactReportPublicationEvidenceMarkdownRow,
		),
		'',
		'## Excluded Evidence',
		'',
		'| Evidence | Command | Reason | Eligibility rule | Owner loop | Boundary |',
		'| --- | --- | --- | --- | --- | --- |',
		...result.excludedEvidence.map(excludedEvidenceMarkdownRow),
		'',
		'## Ranked Claim Portfolio',
		'',
		'| Rank | Claim | Status | North Star link | Owner loops | Handoff decision | Proof command | Kill criterion | Boundary |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.claimPortfolio.map(portfolioClaimMarkdownRow),
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
		fixturePolicyEvidence: cloneFixturePolicyEvidence(result.fixturePolicyEvidence),
		generatedFixtureDecisionEvidence: cloneGeneratedFixtureDecisionEvidence(
			result.generatedFixtureDecisionEvidence,
		),
		performancePolicy: clonePerformancePolicy(),
		safeOpenLatencyValidationEvidence: cloneSafeOpenLatencyValidationEvidence(
			result.safeOpenLatencyValidationEvidence,
		),
		correctnessPolicy: cloneCorrectnessPolicy(),
		correctnessBoundaryEvidence: cloneCorrectnessBoundaryEvidence(
			result.correctnessBoundaryEvidence,
		),
		releasePackageabilityEvidence: cloneReleasePackageabilityEvidence(
			result.releasePackageabilityEvidence,
		),
		compactReportPublicationEvidence: cloneCompactReportPublicationEvidence(
			result.compactReportPublicationEvidence,
		),
		streamingMatrixEvidence: cloneStreamingMatrixEvidence(result.streamingMatrixEvidence),
		nextOwnerActions: result.readiness.nextOwnerActions,
		claimBlockerBoard: result.readiness.claimBlockerBoard,
		implementationHandoffs: result.readiness.implementationHandoffs,
		claimPortfolio: result.claimPortfolio.map(clonePortfolioClaim),
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

function releasePackageabilityEvidence(): ReleaseProofPackageabilityEvidence {
	return {
		ownerLoop: 'release',
		status: 'local-tarball-smokes-present-publication-policy-required',
		ownerApprovalRequired: true,
		sdkSmokeCommand: 'bun run release:sdk:smoke',
		appSmokeCommand: 'bun run release:apps:smoke',
		coveredEvidence: [
			'SDK tarball installs into a temp consumer and verifies create/open/plan/commit/reopen/check/recalc plus bundled agent docs.',
			'CLI/API/MCP app tarballs install into a temp consumer without workspace dependencies.',
			'Installed CLI bin reports version, completes create/write/inspect/plan/commit/check/read over a temp workbook, and searches bundled docs.',
			'Installed API createApiFetch handles write/inspect/plan/commit/check/read plus /capabilities over a temp workbook.',
			'Installed MCP package registers tool/resource callbacks and completes write/inspect/plan/commit/check/read plus docs search and capabilities over a temp workbook.',
		],
		missingPolicyRequirements: [
			'artifact storage path',
			'registry publication workflow',
			'signed provenance or explicit non-provenance wording',
			'API listener lifecycle smoke',
			'stdio MCP protocol-session smoke',
			'retention and privacy filtering for release smoke output',
		],
		forbiddenClaims: [
			'published registry artifacts',
			'signed provenance',
			'SLSA, in-toto, Sigstore, or GitHub artifact attestation',
			'production API lifecycle readiness',
			'full MCP protocol compatibility',
		],
		boundary:
			'Packageability evidence proves local tarball install and basic installed workflow smoke only. It is not registry publication, signed provenance, artifact retention policy, production server lifecycle proof, or a real MCP protocol session.',
	}
}

function cloneReleasePackageabilityEvidence(
	evidence: ReleaseProofPackageabilityEvidence,
): ReleaseProofPackageabilityEvidence {
	return {
		...evidence,
		coveredEvidence: [...evidence.coveredEvidence],
		missingPolicyRequirements: [...evidence.missingPolicyRequirements],
		forbiddenClaims: [...evidence.forbiddenClaims],
	}
}

const CLAIM_PORTFOLIO: readonly ReleaseProofPortfolioClaim[] = [
	portfolioClaim({
		rank: 1,
		name: 'safe-open-proof',
		claim: 'safe unknown workbook opening',
		northStarLink: 'Preservation-first XLSX and trustworthy agent workflows.',
		status: 'claim-wording-allowed-today',
		evidenceNeeded: claimProofRequired('safe-open-proof'),
		likelyHandoffOwner: ['product', 'performance', 'release'],
		handoffDecision: 'top-implementation-handoff',
		proofCommand: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
		boundary:
			'Hand off proof packaging and owner decisions only; do not add opener surfaces while release blockers remain.',
	}),
	portfolioClaim({
		rank: 2,
		name: 'package-action-proof',
		claim: 'auditable package-part mutation',
		northStarLink: 'Trustworthy mutation planning and preservation-first writes.',
		status: 'claim-wording-allowed-today',
		evidenceNeeded: claimProofRequired('package-action-proof'),
		likelyHandoffOwner: ['correctness', 'product', 'performance', 'release'],
		handoffDecision: 'top-implementation-handoff',
		proofCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
		boundary:
			'Hand off package-accounting proof and owner boundary decisions only; do not add mutation surfaces while release blockers remain.',
	}),
	portfolioClaim({
		rank: 3,
		name: 'formula-language-service-primitives',
		claim: 'formula language-service primitives',
		northStarLink: 'Formula intelligence without unsafe workbook mutation.',
		status: 'needs-one-more-fold-in',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Public formula corpus plus explicit LET, defined-name, table, external-ref, sheet/range, and parse-failure refusal snapshots.',
			benchmark:
				'Formula-assist proof over public formulas with sampled reference spans, binding roles, refusal counts, and owner-approved latency wording only after a timed run.',
			surface:
				'Existing SDK/CLI/API/MCP formula-assist surfaces; no edit-producing rename surface.',
			validationGate:
				'Run formula-assist proof, parser/span/binding tests, cross-surface formula-assist tests, and typecheck/Biome when code changes.',
			competitorContrast:
				'LSP prepareRename allows refusal before rename; HyperFormula is a formula engine baseline, while Ascend should claim workbook-preserving formula assistance.',
			honestBoundary:
				'No edit-producing rename, defined-name rename, table-column rename, sheet/range rename, or external-ref rename.',
			killCriterion:
				'Do not promote rename until workbook-context symbol ownership and operation-owned edit plans exist.',
		}),
		likelyHandoffOwner: ['product', 'correctness'],
		handoffDecision: 'proof-packaging-only',
		proofCommand:
			'bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json',
		boundary: 'Rejection-first proof only; rename remains frozen.',
	}),
	portfolioClaim({
		rank: 4,
		name: 'token-bounded-agent-view',
		claim: 'token-bounded agent view',
		northStarLink: 'World-class agent DX under strict context budgets.',
		status: 'needs-one-more-fold-in',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Dense table, wide sparse, formula-heavy, metadata-heavy, and public workbook examples with omission locators.',
			benchmark:
				'Full versus budgeted estimate, compression ratio, omitted counters, recovery-locator coverage, and structural floor reporting.',
			surface: 'Existing SDK, CLI, API, and MCP agent-view/read surfaces only.',
			validationGate:
				'Deterministic truncation, cross-surface JSON shape, omitted-evidence recovery, and no hidden summarization without counters.',
			competitorContrast:
				'Univer exposes agent spreadsheet operations; Ascend claim is deterministic local evidence under token budgets.',
			honestBoundary: 'Token counts are approximate and omitted evidence is absent by design.',
			killCriterion:
				'Do not publish exact-token wording when structural floors exceed tiny budgets or omitted evidence is not recoverable by locator.',
		}),
		likelyHandoffOwner: ['product'],
		handoffDecision: 'proof-packaging-only',
		boundary: 'Needs a product-shaped example, not another surface.',
	}),
	portfolioClaim({
		rank: 5,
		name: 'retained-viewport-patch-history',
		claim: 'retained viewport patch history',
		northStarLink: 'Real-world performance and UI/agent efficiency.',
		status: 'needs-one-more-fold-in',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Retained patch, skipped token, invalid token, expired history, projection change, metadata invalidation, and changed-source cases.',
			benchmark:
				'Patch bytes, retained history size, invalidation rates, and generation-token retention caps.',
			surface: 'SDK interactive patch stream plus API/MCP compact recovery; CLI excluded.',
			validationGate:
				'Viewport proof harness, SDK interactive contract tests, API/MCP compact changedSince tests, and retention cap assertions.',
			competitorContrast:
				'Database MVCC retains readable versions, but this claim is bounded patch history, not transaction isolation or CRDT collaboration.',
			honestBoundary:
				'Bounded per-window history only, not unlimited history or multi-writer sync.',
			killCriterion:
				'Do not promote collaboration, sync, CRDT, or unlimited-history wording without multi-writer convergence and storage retention proof.',
		}),
		likelyHandoffOwner: ['product', 'performance'],
		handoffDecision: 'proof-packaging-only',
		boundary: 'Product proof is bounded and CLI remains excluded.',
	}),
	portfolioClaim({
		rank: 6,
		name: 'release-proof-bundle',
		claim: 'release proof bundle',
		northStarLink: 'Trustworthy releases and auditability without fake attestations.',
		status: 'needs-one-more-fold-in',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'One real public workbook workflow per top claim with inspect, plan, commit, reopen, diff, audit, and digest evidence.',
			benchmark:
				'Bundle generation overhead, output size, canonical report bytes, and compact versus expanded report cost.',
			surface:
				'Stable SDK report schema first; CLI/API/MCP references only after artifact storage and privacy semantics stabilize.',
			validationGate:
				'Golden proof fixtures, digest checks, reopen/diff/audit checks, package graph audit checks, and failure cases.',
			competitorContrast:
				'Generic spreadsheet libraries read and write files; Ascend should explain the decision trail.',
			honestBoundary:
				'Not signed, tamper-evident, SLSA, in-toto, certified provenance, or third-party attestation.',
			killCriterion:
				'Do not publish bundle wording until storage, retention/privacy filtering, canonicalization subject, and offline verification policy are approved.',
		}),
		likelyHandoffOwner: ['product', 'release'],
		handoffDecision: 'proof-packaging-only',
		boundary:
			'Proof bundle work follows the top two claims and must stay below attestation language.',
	}),
	portfolioClaim({
		rank: 7,
		name: 'formula-oracle-routing',
		claim: 'formula oracle routing',
		northStarLink: 'Correctness credibility for real formula behavior.',
		status: 'speculative-do-not-promote',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Runnable public corpus artifacts by mismatch class: cached-only, volatile, numeric drift, unsupported function, external refs, dynamic arrays, structured refs, and date systems.',
			benchmark:
				'Per-oracle route overhead, corpus completion time, skip counters, and divergence counters.',
			surface: 'Completed JSON artifacts and CLI report only; no MCP/API promotion.',
			validationGate:
				'Converter tests, artifact verifier, skipped/divergence counters, and no threshold changes without evidence.',
			competitorContrast:
				'HyperFormula is the strongest OSS formula baseline; Excel and LibreOffice are behavior oracles with automation limits.',
			honestBoundary: 'No blanket Excel-compatible formula claim.',
			killCriterion:
				'Do not publish compatibility claims while private corpora, cached values, or unsupported oracle classes are required.',
		}),
		likelyHandoffOwner: ['correctness'],
		handoffDecision: 'do-not-promote-yet',
		boundary:
			'Correctness research only until mismatch classes are reproducible from public artifacts.',
	}),
	portfolioClaim({
		rank: 8,
		name: 'property-journal-laws',
		claim: 'property-style journal laws',
		northStarLink: 'Trustworthy mutation planning with generated inverse-law evidence.',
		status: 'speculative-do-not-promote',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Generated operation sequences covering cells, rows, columns, sheets, formulas, styles, tables, and explicit lossy boundaries.',
			benchmark:
				'Shrink time, failing-case minimization size, seed stability, and changed-test integration cost.',
			surface: 'Test harness only until laws and exclusions stabilize.',
			validationGate:
				'fast-check shrinking, deterministic seeds, explicit exclusions, and journal compatibility assertions.',
			competitorContrast:
				'Property-based tests prove invariants over operation spaces instead of adding hand-written fixture examples.',
			honestBoundary:
				'Generated laws are scoped to covered operations and excluded lossy metadata/style boundaries.',
			killCriterion:
				'Do not promote broad inverse-law claims until generated coverage is shrinkable and lossy boundaries are explicit.',
		}),
		likelyHandoffOwner: ['correctness'],
		handoffDecision: 'do-not-promote-yet',
		boundary: 'Testing strategy, not a release product claim.',
	}),
	portfolioClaim({
		rank: 9,
		name: 'columnar-scan-sidecars',
		claim: 'columnar scan sidecars',
		northStarLink: 'Real-world performance without replacing workbook truth.',
		status: 'speculative-do-not-promote',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Real workbook tables and ranges with numbers, dates, blanks, strings, formulas, filters, hidden rows, and styles.',
			benchmark:
				'Repeated scans, sidecar build cost, invalidation cost, memory overhead, and checksum parity against canonical workbook reads.',
			surface: 'Benchmark harness only; no SDK, CLI, API, or MCP product surface.',
			validationGate:
				'Generation-key invalidation, checksum parity, memory caps, and benchmark guard before production.',
			competitorContrast:
				'DuckDB reads XLSX ranges into typed SQL tables; Arrow supplies the columnar scan substrate.',
			honestBoundary:
				'Not a storage engine, workbook rewrite, or guaranteed faster path for sparse or single-pass reads.',
			killCriterion:
				'Do not promote if build plus invalidation erases repeated-scan gains, parity fails, or memory overhead is not bounded.',
		}),
		likelyHandoffOwner: ['performance'],
		handoffDecision: 'do-not-promote-yet',
		boundary: 'Performance research only.',
	}),
	portfolioClaim({
		rank: 10,
		name: 'agent-workflow-observability',
		claim: 'agent workflow observability',
		northStarLink: 'World-class agent DX and recoverable workflow audits.',
		status: 'speculative-do-not-promote',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Public inspect, plan, commit, reopen, diff, and audit workflow traces with failure taxonomy and recovery prompts.',
			benchmark:
				'Trace payload size, redaction overhead, failure-class coverage, and repair/recovery decision improvement.',
			surface: 'Existing trace artifacts only until proof shows concrete repair or audit value.',
			validationGate:
				'Trace golden tests, redaction checks, failure taxonomy snapshots, and recovery prompt validation.',
			competitorContrast:
				'Agent logs explain what happened; Ascend should prove traces improve workbook repair and audit decisions.',
			honestBoundary:
				'No autonomous correctness, signed audit trail, or complete observability claim.',
			killCriterion:
				'Do not promote if traces are only verbose logs or do not improve a concrete repair, audit, or recovery workflow.',
		}),
		likelyHandoffOwner: ['product'],
		handoffDecision: 'do-not-promote-yet',
		boundary: 'Observability must prove recovery value before promotion.',
	}),
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

const RELEASE_PACKAGE_ACTION_KINDS: readonly ReleaseProofPackageActionKind[] = [
	'passthrough',
	'regenerate',
	'add',
	'drop',
	'error',
]

function cloneDeferredClaim(claim: ReleaseProofDeferredClaim): ReleaseProofDeferredClaim {
	return {
		...claim,
		ownerLoops: [...claim.ownerLoops],
	}
}

function portfolioClaim(
	claim: Omit<ReleaseProofPortfolioClaim, 'killCriterion'>,
): ReleaseProofPortfolioClaim {
	return {
		...claim,
		killCriterion: claim.evidenceNeeded.killCriterion,
		likelyHandoffOwner: [...claim.likelyHandoffOwner],
	}
}

function claimEvidenceNeeded(
	evidenceNeeded: ReleaseProofClaimProofRequired,
): ReleaseProofClaimProofRequired {
	return evidenceNeeded
}

function clonePortfolioClaim(claim: ReleaseProofPortfolioClaim): ReleaseProofPortfolioClaim {
	return {
		...claim,
		evidenceNeeded: { ...claim.evidenceNeeded },
		likelyHandoffOwner: [...claim.likelyHandoffOwner],
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
					'approve that representative streaming writer proofs are sufficient for release wording, or expand package-action proof to streaming variants for every package-action scenario before claiming streaming parity',
				evidence:
					'current proof reports three streamingProofCases covering passthrough/regenerate/add/drop and two streamingRegenerateParts; error remains non-streaming',
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

function correctnessBoundaryEvidence(
	safeOpen: SafeOpenProofResult,
	packageAction: PackageActionProofResult,
): ReleaseProofCorrectnessBoundaryEvidence {
	const featureChecks: ReleaseProofCorrectnessBoundaryFeatureCheck[] = [
		correctnessBoundaryFeatureCheck({
			feature: 'digital-signatures',
			evidencePresent:
				packageCaseHasAction(packageAction, 'signature-invalidation-drop', 'drop') &&
				packageCasePostWritePassed(packageAction, 'signature-invalidation-drop'),
			evidenceSources: [
				'package-action-proof/signature-invalidation-drop',
				'safe-open-proof/signed',
			],
			proofChecks: [
				'generated edge package is disclosed',
				'commit proof records a drop action for signature package parts',
				'safe-open proof routes signature package features to metadata-only review',
			],
		}),
		correctnessBoundaryFeatureCheck({
			feature: 'calc-chain',
			evidencePresent:
				packageCaseIsPublicFixture(packageAction, 'calc-chain-drop') &&
				packageCaseHasAction(packageAction, 'calc-chain-drop', 'drop') &&
				packageCasePostWritePassed(packageAction, 'calc-chain-drop'),
			evidenceSources: ['package-action-proof/calc-chain-drop'],
			proofChecks: [
				'public fixture is used',
				'commit proof records a drop action for xl/calcChain.xml',
				'post-write audit passes after dropping unsafe calculation-order metadata',
			],
		}),
		correctnessBoundaryFeatureCheck({
			feature: 'chart-drawing-sidecars',
			evidencePresent:
				packageCaseIsPublicFixture(packageAction, 'chart-sidecar-accounting') &&
				packageCaseHasAction(packageAction, 'chart-sidecar-accounting', 'passthrough') &&
				packageCaseHasAction(packageAction, 'chart-sidecar-accounting', 'regenerate') &&
				packageCasePostWritePassed(packageAction, 'chart-sidecar-accounting'),
			evidenceSources: ['package-action-proof/chart-sidecar-accounting'],
			proofChecks: [
				'public chart fixture is used',
				'commit proof records passthrough sidecars and regenerated workbook-owned parts',
				'post-write audit passes without claiming semantic chart support',
			],
		}),
		correctnessBoundaryFeatureCheck({
			feature: 'macros-activex',
			evidencePresent:
				packageCaseIsPublicFixture(packageAction, 'macro-passthrough') &&
				packageCaseHasAction(packageAction, 'macro-passthrough', 'passthrough') &&
				safeOpenRiskRoutedToReview(safeOpen, 'macro', 'preservedMacro') &&
				safeOpenRiskRoutedToReview(safeOpen, 'activex', 'preservedActiveX'),
			evidenceSources: [
				'package-action-proof/macro-passthrough',
				'safe-open-proof/macro',
				'safe-open-proof/activex',
			],
			proofChecks: [
				'public macro fixture is used for package-action evidence',
				'macro-bearing parts are recorded as package evidence',
				'safe-open proof routes macro and ActiveX risk families to review before hydration',
			],
		}),
		correctnessBoundaryFeatureCheck({
			feature: 'unknown-parts',
			evidencePresent:
				packageCaseHasAction(packageAction, 'unknown-part-error', 'error') &&
				!packageCasePostWritePassed(packageAction, 'unknown-part-error') &&
				safeOpenRiskRoutedToReview(safeOpen, 'unknown-part', 'preservedOther'),
			evidenceSources: ['package-action-proof/unknown-part-error', 'safe-open-proof/unknown-part'],
			proofChecks: [
				'generated edge package is disclosed',
				'commit proof records an error action for the unknown package part',
				'post-write audit fails closed and safe-open routes unknown package features to review',
			],
		}),
		correctnessBoundaryFeatureCheck({
			feature: 'streaming-scope',
			evidencePresent:
				packageCaseHasRepresentativeStreamingProof(packageAction, 'docprops-passthrough') &&
				packageCaseHasRepresentativeStreamingProof(packageAction, 'calc-chain-drop'),
			evidenceSources: [
				'package-action-proof/docprops-passthrough',
				'package-action-proof/add-sheet-part',
				'package-action-proof/calc-chain-drop',
			],
			proofChecks: [
				'representative streaming proofs cover passthrough/regenerate/add/drop',
				'streaming proof records regenerated worksheet parts and passthrough byte equality',
				'error and macro/chart streaming remain outside the proof boundary',
			],
		}),
	]
	const missingFeatureNames = featureChecks
		.filter((entry) => !entry.evidencePresent)
		.map((entry) => entry.feature)
	return {
		artifact: 'package-action-proof',
		gateId: 'unsupported-feature-boundary',
		ownerLoop: 'correctness',
		status: 'evidence-present-owner-approval-required',
		allCurrentEvidencePresent: featureChecks.every((entry) => entry.evidencePresent),
		missingFeatureNames,
		ownerEscalationRequired: missingFeatureNames.length > 0,
		ownerApprovalRequired: true,
		validationCommand:
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
		featureChecks,
		boundary:
			'Evidence check for unsupported-feature wording only. It does not satisfy the owner gate, approve release copy, or prove semantic support for signatures, calc chains, charts, macros, ActiveX, unknown parts, or full streaming parity.',
	}
}

function correctnessBoundaryFeatureCheck(input: {
	readonly feature: string
	readonly evidencePresent: boolean
	readonly evidenceSources: readonly string[]
	readonly proofChecks: readonly string[]
}): ReleaseProofCorrectnessBoundaryFeatureCheck {
	const wording = correctnessBoundaryWording(input.feature)
	return {
		feature: input.feature,
		evidencePresent: input.evidencePresent,
		evidenceSources: [...input.evidenceSources],
		proofChecks: [...input.proofChecks],
		allowedWording: wording.allowedWording,
		forbiddenWording: wording.forbiddenWording,
	}
}

function correctnessBoundaryWording(feature: string): ReleaseProofUnsupportedFeatureBoundary {
	const boundary = CORRECTNESS_POLICY.unsupportedFeatureBoundaries.find(
		(entry) => entry.feature === feature,
	)
	if (!boundary) throw new Error(`Missing correctness boundary wording for ${feature}`)
	return boundary
}

function packageCase(
	result: PackageActionProofResult,
	name: string,
): PackageActionProofCaseResult | undefined {
	return result.cases.find((entry) => entry.name === name)
}

function packageCaseHasAction(
	result: PackageActionProofResult,
	name: string,
	action: keyof PackageActionProofCaseResult['commitActionCounts'],
): boolean {
	return (packageCase(result, name)?.commitActionCounts[action] ?? 0) > 0
}

function packageCasePostWritePassed(result: PackageActionProofResult, name: string): boolean {
	return packageCase(result, name)?.postWriteAuditsPassed === true
}

function packageCaseIsPublicFixture(result: PackageActionProofResult, name: string): boolean {
	return packageCase(result, name)?.sourceKind === 'public-fixture'
}

function packageCaseHasRepresentativeStreamingProof(
	result: PackageActionProofResult,
	name: string,
): boolean {
	const streamingProof = packageCase(result, name)?.streamingProof
	return (
		streamingProof !== undefined &&
		streamingProof.streamingRegeneratePartPaths.length > 0 &&
		streamingProof.passthroughBytesEqualCount > 0 &&
		streamingProof.issueCount === 0
	)
}

function streamingMatrixEvidence(
	result: PackageActionProofResult,
): ReleaseProofStreamingMatrixEvidence {
	const streamingCases = result.cases.filter((entry) => entry.streamingProof !== undefined)
	const nonStreamingCases = result.cases.filter((entry) => entry.streamingProof === undefined)
	const coveredActionKinds = RELEASE_PACKAGE_ACTION_KINDS.filter((action) =>
		streamingCases.some((entry) => (entry.streamingProof?.actionCounts[action] ?? 0) > 0),
	)
	const covered = new Set(coveredActionKinds)
	return {
		artifact: 'package-action-proof',
		gateId: 'streaming-matrix-boundary',
		ownerLoop: 'performance',
		status: 'representative-proof-present-owner-approval-required',
		ownerApprovalRequired: true,
		validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
		representativeProofCases: streamingCases.length,
		streamingRegenerateParts: streamingCases.reduce(
			(count, entry) => count + (entry.streamingProof?.streamingRegeneratePartPaths.length ?? 0),
			0,
		),
		coveredActionKinds,
		missingActionKinds: RELEASE_PACKAGE_ACTION_KINDS.filter((action) => !covered.has(action)),
		coveredCaseNames: streamingCases.map((entry) => entry.name),
		nonStreamingCaseNames: nonStreamingCases.map((entry) => entry.name),
		publicNonStreamingCaseNames: nonStreamingCases
			.filter((entry) => entry.sourceKind === 'public-fixture')
			.map((entry) => entry.name),
		generatedNonStreamingCaseNames: nonStreamingCases
			.filter((entry) => entry.sourceKind !== 'public-fixture')
			.map((entry) => entry.name),
		streamingIssueCaseNames: streamingCases
			.filter((entry) => (entry.streamingProof?.issueCount ?? 0) > 0)
			.map((entry) => entry.name),
		boundary:
			'Streaming matrix evidence proves representative streaming package-action cases covering passthrough/regenerate/add/drop. It does not prove full streaming parity, error streaming behavior, or streaming coverage for public macro/chart fixtures.',
	}
}

function cloneStreamingMatrixEvidence(
	evidence: ReleaseProofStreamingMatrixEvidence,
): ReleaseProofStreamingMatrixEvidence {
	return {
		...evidence,
		coveredActionKinds: [...evidence.coveredActionKinds],
		missingActionKinds: [...evidence.missingActionKinds],
		coveredCaseNames: [...evidence.coveredCaseNames],
		nonStreamingCaseNames: [...evidence.nonStreamingCaseNames],
		publicNonStreamingCaseNames: [...evidence.publicNonStreamingCaseNames],
		generatedNonStreamingCaseNames: [...evidence.generatedNonStreamingCaseNames],
		streamingIssueCaseNames: [...evidence.streamingIssueCaseNames],
	}
}

function safeOpenRiskRoutedToReview(
	result: SafeOpenProofResult,
	name: string,
	riskFamily: string,
): boolean {
	const proofCase = result.cases.find((entry) => entry.name === name)
	return (
		proofCase?.status === 'ok' &&
		proofCase.reviewBeforeHydration === true &&
		proofCase.riskFamilies.includes(riskFamily)
	)
}

const COMPACT_REPORT_FORBIDDEN_PAYLOAD_FIELDS = new Set([
	'inputBytes',
	'outputBytes',
	'inputSha256',
	'sourceSha256',
	'outputSha256',
	'sha256',
	'stableShapeSha256',
	'proofJsonBytes',
	'streamingRegeneratePartPaths',
])

function compactReportPublicationEvidence(
	safeOpen: SafeOpenCompactReleaseReport,
	packageAction: PackageActionCompactReleaseReport,
	artifacts: readonly ReleaseProofIndexArtifact[],
): ReleaseProofCompactReportPublicationEvidence {
	const reports = [
		compactReportPublicationEvidenceItem('safe-open-proof', safeOpen, artifacts),
		compactReportPublicationEvidenceItem('package-action-proof', packageAction, artifacts),
	]
	const policyDecisions = compactReportPublicationPolicyDecisions()
	return {
		ownerLoop: 'release',
		status: 'local-summary-present-publication-policy-required',
		ownerApprovalRequired: true,
		compactReportDigestsIndexed: false,
		allCompactCommandsPresent: reports.every((report) => report.command.length > 0),
		compactReportsEmbedForbiddenPayloadFields: reports.some(
			(report) => report.forbiddenPayloadFieldsPresent.length > 0,
		),
		generatedAtIncluded: reports.every((report) => report.generatedAtIncluded),
		missingPolicyRequirements: policyDecisions.map((entry) => entry.requirement),
		policyDecisions,
		reports,
		boundary:
			'Compact report publication evidence proves only that local claim-safe summaries exist. It does not publish compact report digests, define artifact storage, canonicalize bytes, or create signed provenance.',
	}
}

function compactReportPublicationPolicyDecisions(): ReleaseProofCompactReportPublicationPolicyDecision[] {
	return [
		{
			requirement: 'artifact storage path',
			ownerLoop: 'release',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Choose the durable location and naming policy for compact release reports before any digest or URL is published.',
			acceptanceEvidence:
				'Release owns a path convention that separates local proof summaries from full artifacts and generated/private workbook bytes.',
			rejectIf:
				'Reports are published from temporary paths, local worktrees, private corpora, or locations that imply tamper-evident storage without an attestation system.',
		},
		{
			requirement: 'retention and privacy filtering',
			ownerLoop: 'release',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Define retention duration and privacy filtering for compact reports, fixture names, command strings, and any future expanded evidence.',
			acceptanceEvidence:
				'Release policy states what is kept, what is redacted, and how private workbook paths, generated bytes, and per-part digests stay out of public summaries.',
			rejectIf:
				'Compact reports include workbook bytes, private workbook paths, source/output digests, generated package bytes, or unpublished corpus identifiers.',
		},
		{
			requirement: 'canonicalization subject',
			ownerLoop: 'release',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Define exactly which JSON shape is canonical for compact report verification and which timestamp or environment fields are excluded.',
			acceptanceEvidence:
				'Release policy names the canonical JSON subject, noise-stripping rules, stable-shape expectations, and compatibility expectations for future fields.',
			rejectIf:
				'Digest wording depends on noncanonical pretty JSON, wall-clock timestamps, local paths, or environment-specific field ordering.',
		},
		{
			requirement: 'offline verification expectations',
			ownerLoop: 'release',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Define whether an offline verifier should reproduce compact reports from source, compare digests, or only inspect owner-approved summaries.',
			acceptanceEvidence:
				'Release policy states the verifier command, expected inputs, failure behavior, and boundary below SLSA, in-toto, Sigstore, and GitHub artifact attestation claims.',
			rejectIf:
				'Copy implies signed provenance, third-party attestation, transparency-log inclusion, tamper-evident storage, or identity-bound build provenance without implementing it.',
		},
	]
}

function compactReportPublicationEvidenceItem(
	artifact: ReleaseProofIndexArtifactName,
	report: SafeOpenCompactReleaseReport | PackageActionCompactReleaseReport,
	artifacts: readonly ReleaseProofIndexArtifact[],
): ReleaseProofCompactReportPublicationEvidenceItem {
	const readyWhenGatePresent =
		artifacts
			.find((entry) => entry.name === artifact)
			?.readyWhen.some((entry) => entry.id === 'compact-report-publication-policy') ?? false
	return {
		artifact,
		gateId: 'compact-report-publication-policy',
		command: report.command,
		jsonBytes: utf8Bytes(JSON.stringify(report)),
		topLevelFields: Object.keys(report).sort(),
		forbiddenPayloadFieldsPresent: compactReportForbiddenPayloadFields(report),
		readyWhenGatePresent,
		generatedAtIncluded: typeof report.generatedAt === 'string' && report.generatedAt.length > 0,
		headlineClaimAllowed: report.headlineClaimAllowed,
		releaseGate: report.releaseGate,
		boundary: report.boundary,
	}
}

function compactReportForbiddenPayloadFields(value: unknown): string[] {
	const found = new Set<string>()
	collectForbiddenCompactReportFields(value, found)
	return Array.from(found).sort()
}

function collectForbiddenCompactReportFields(value: unknown, found: Set<string>): void {
	if (Array.isArray(value)) {
		for (const item of value) collectForbiddenCompactReportFields(item, found)
		return
	}
	if (value === null || typeof value !== 'object') return
	for (const [key, nested] of Object.entries(value)) {
		if (COMPACT_REPORT_FORBIDDEN_PAYLOAD_FIELDS.has(key)) found.add(key)
		collectForbiddenCompactReportFields(nested, found)
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
	const claimBlockerBoard = buildClaimBlockerBoard(artifacts, nextOwnerActions)
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
		claimBlockerBoard,
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

function buildClaimBlockerBoard(
	artifacts: readonly ReleaseProofIndexArtifact[],
	nextOwnerActions: readonly ReleaseProofNextOwnerAction[],
): readonly ReleaseProofClaimBlockerBoardRow[] {
	const rows: ReleaseProofClaimBlockerBoardRow[] = []
	for (const artifact of artifacts) {
		for (const ownerLoop of uniqueOwnerLoops(
			artifact.readyWhen.map((requirement) => requirement.ownerLoop),
		)) {
			const actions = nextOwnerActions.filter(
				(action) => action.artifact === artifact.name && action.ownerLoop === ownerLoop,
			)
			if (actions.length === 0) continue
			rows.push({
				artifact: artifact.name,
				claim: artifact.claim,
				ownerLoop,
				blockerCount: actions.length,
				requirementIds: actions.map((action) => action.requirementId),
				actionRanks: actions.map((action) => action.rank),
				nextStepKinds: uniqueNextStepKinds(actions.map((action) => action.nextStepKind)),
				acceptanceEvidence: actions.map((action) => action.acceptanceEvidence),
				forbiddenShortcuts: actions.map((action) => action.forbiddenShortcut),
				boundary:
					'Claim blocker board row is derived from missing readyWhen gates and owner actions. It is routing evidence only, not gate satisfaction or permission to add product surfaces.',
			})
		}
	}
	return rows.sort((left, right) => {
		const artifactOrder =
			artifacts.findIndex((artifact) => artifact.name === left.artifact) -
			artifacts.findIndex((artifact) => artifact.name === right.artifact)
		if (artifactOrder !== 0) return artifactOrder
		return ownerLoopRank(left.ownerLoop) - ownerLoopRank(right.ownerLoop)
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
					'Do not publish stronger wording if synthetic edge packages are hidden, chart XML is described as byte-passthrough, representative streaming proofs are described as full matrix parity, or local digests imply attestation.',
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
					'Streaming wording must stay limited to representative proof cases unless a broader matrix is approved.',
				acceptanceEvidence:
					'Performance accepts representative streaming proofs covering passthrough/regenerate/add/drop for narrow wording, or expands the matrix to error and public macro/chart cases.',
				forbiddenShortcut:
					'Do not call representative streaming cases full streaming parity or imply streaming coverage for every package-action scenario.',
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

function formatClaimBlockerBoard(rows: readonly ReleaseProofClaimBlockerBoardRow[]): string {
	return rows
		.map(
			(row) =>
				`${row.artifact}/${row.ownerLoop}=${row.requirementIds.join(',')}:${row.nextStepKinds.join('+')}`,
		)
		.join('; ')
}

function claimBlockerBoardMarkdownRow(row: ReleaseProofClaimBlockerBoardRow): string {
	return [
		row.artifact,
		row.claim,
		row.ownerLoop,
		row.requirementIds.join(','),
		row.nextStepKinds.join(','),
		row.acceptanceEvidence.join('; '),
		row.forbiddenShortcuts.join('; '),
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
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

function ownerLoopRank(ownerLoop: ReleaseProofReadinessOwner): number {
	return (['correctness', 'performance', 'product', 'release'] as const).indexOf(ownerLoop)
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

function safeOpenFixtureAcceptanceMarkdownRow(
	row: ReleaseProofSafeOpenFixtureAcceptanceItem,
): string {
	return [
		row.caseName,
		row.generatedCaseKind,
		row.acceptableAsTopologyProofWhen,
		row.requiresPublicBinaryWhen,
		`\`${row.validationCommand}\``,
		row.gateEffect,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function packageActionFixtureAcceptanceMarkdownRow(
	row: ReleaseProofPackageActionFixtureAcceptanceItem,
): string {
	return [
		row.caseName,
		row.generatedCaseKind,
		row.acceptableAsPackageActionProofWhen,
		row.requiresPublicBinaryWhen,
		row.forbiddenClaim,
		`\`${row.validationCommand}\``,
		row.gateEffect,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
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

function fixturePolicyEvidence(
	safeOpen: SafeOpenFixtureScanResult,
	packageAction: PackageActionFixtureScanResult,
): ReleaseProofFixturePolicyEvidence {
	const packageActionMissingReplacementFeatures = [
		packageAction.featureCounts.signaturePackage === 0 ? 'signaturePackage' : undefined,
		packageAction.featureCounts.syntheticUnknownPathFamily === 0
			? 'syntheticUnknownPathFamily'
			: undefined,
	].filter((entry): entry is string => entry !== undefined)
	const safeOpenEvidence: ReleaseProofSafeOpenFixturePolicyEvidence = {
		artifact: 'safe-open-proof',
		gateId: 'public-edge-fixtures',
		validationCommand: FIXTURE_POLICY.trackedFixtureScanCommands['safe-open-proof'],
		corpus: safeOpen.corpus,
		scanned: safeOpen.scanned,
		rejected: safeOpen.rejected,
		replacementStatus: safeOpen.replacementStatus,
		riskFamilyCounts: { ...safeOpen.riskFamilyCounts },
		signatureOrUnknownMatches: safeOpen.signatureOrUnknownMatches.length,
		currentGeneratedStructuralCases: [
			...FIXTURE_POLICY.currentGeneratedStructuralCases['safe-open-proof'],
		],
		externalCandidateEvidence: SAFE_OPEN_EXTERNAL_FIXTURE_CANDIDATES.map((entry) => ({
			...entry,
		})),
		boundary: safeOpen.boundary,
	}
	const packageActionEvidence: ReleaseProofPackageActionFixturePolicyEvidence = {
		artifact: 'package-action-proof',
		gateId: 'edge-fixture-policy',
		validationCommand: FIXTURE_POLICY.trackedFixtureScanCommands['package-action-proof'],
		corpus: packageAction.corpus,
		scanned: packageAction.scanned,
		rejected: packageAction.rejected,
		replacementStatus: packageAction.replacementStatus,
		featureCounts: { ...packageAction.featureCounts },
		currentGeneratedStructuralCases: [
			...FIXTURE_POLICY.currentGeneratedStructuralCases['package-action-proof'],
		],
		missingReplacementFeatures: packageActionMissingReplacementFeatures,
		boundary: packageAction.boundary,
	}
	return {
		ownerLoop: 'product',
		status: 'tracked-scan-complete-owner-approval-required',
		ownerApprovalRequired: true,
		allScansUseTrackedCorpus:
			safeOpen.corpus === 'tracked-git-fixtures' && packageAction.corpus === 'tracked-git-fixtures',
		publicReplacementGapsRemain:
			safeOpen.replacementStatus === 'no-public-binary-replacement-found' ||
			packageAction.replacementStatus === 'remaining-generated-edge-cases',
		safeOpen: safeOpenEvidence,
		packageAction: packageActionEvidence,
		boundary:
			'Fixture scan evidence is local tracked-corpus evidence for owner decisions. It does not prove that no suitable public fixtures exist elsewhere, approve generated fixtures as public binaries, or satisfy product release gates.',
	}
}

function fixturePolicyEvidenceMarkdownRow(
	row: ReleaseProofSafeOpenFixturePolicyEvidence | ReleaseProofPackageActionFixturePolicyEvidence,
): string {
	const gapEvidence =
		row.artifact === 'safe-open-proof'
			? `signatureOrUnknownMatches=${row.signatureOrUnknownMatches}`
			: `missingReplacementFeatures=${row.missingReplacementFeatures.join(',') || 'none'}; featureCounts=${formatFixtureFeatureCounts(row.featureCounts)}`
	return [
		row.artifact,
		row.gateId,
		`\`${row.validationCommand}\``,
		row.corpus,
		String(row.scanned),
		String(row.rejected),
		row.replacementStatus,
		row.currentGeneratedStructuralCases.join(',') || 'none',
		gapEvidence,
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function externalFixtureCandidateEvidenceMarkdownRow(
	row: ReleaseProofExternalFixtureCandidateEvidence,
): string {
	return [
		row.artifact,
		row.gateId,
		row.caseName,
		row.status,
		`${row.candidateId} (${row.sourceUrl})`,
		`${row.license} (${row.licenseEvidenceUrl})`,
		row.sha256,
		`${row.riskFamily}; ${row.recommendedMode}; reviewBeforeHydration=${row.reviewBeforeHydration}; parts=${row.partCount}; relationships=${row.relationshipCount}; sample=${row.sampleUnknownPart}`,
		row.gateEffect,
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneFixturePolicyEvidence(
	evidence: ReleaseProofFixturePolicyEvidence,
): ReleaseProofFixturePolicyEvidence {
	return {
		...evidence,
		safeOpen: {
			...evidence.safeOpen,
			riskFamilyCounts: { ...evidence.safeOpen.riskFamilyCounts },
			currentGeneratedStructuralCases: [...evidence.safeOpen.currentGeneratedStructuralCases],
			externalCandidateEvidence: evidence.safeOpen.externalCandidateEvidence.map((entry) => ({
				...entry,
			})),
		},
		packageAction: {
			...evidence.packageAction,
			featureCounts: { ...evidence.packageAction.featureCounts },
			currentGeneratedStructuralCases: [...evidence.packageAction.currentGeneratedStructuralCases],
			missingReplacementFeatures: [...evidence.packageAction.missingReplacementFeatures],
		},
	}
}

function generatedFixtureDecisionEvidence(
	fixtureEvidence: ReleaseProofFixturePolicyEvidence,
): ReleaseProofGeneratedFixtureDecisionEvidence {
	const cases: ReleaseProofGeneratedFixtureDecisionCase[] = [
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			caseName: 'signed',
			generatedKind: 'generated-edge-package',
			replacementEvidence: `tracked safe-open scan found signatureOrUnknownMatches=${fixtureEvidence.safeOpen.signatureOrUnknownMatches} across ${fixtureEvidence.safeOpen.scanned} fixtures`,
			ownerDecisionNeeded:
				'Accept disclosed generated signature package topology as safe-open routing proof, or provide an approved public signed workbook fixture.',
			recommendedOwnerAction:
				'Accept disclosed generated topology for local safe-open package-feature routing only, while keeping the release gate missing until product and release approve the wording.',
			allowedUse:
				'Pre-hydration package-feature routing evidence that detects signature-related package parts.',
			forbiddenUse:
				'Do not claim signature verification, signature preservation, trust, malware scanning, signed provenance, or real-world vendor behavior.',
		},
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			caseName: 'unknown-part',
			generatedKind: 'generated-edge-package',
			replacementEvidence: `tracked safe-open scan found signatureOrUnknownMatches=${fixtureEvidence.safeOpen.signatureOrUnknownMatches} across ${fixtureEvidence.safeOpen.scanned} fixtures; external candidate excelforge-book1-unknown-part awaits owner review and is not vendored`,
			ownerDecisionNeeded:
				'Accept disclosed generated unknown-part topology as safe-open routing proof, or provide an approved public unknown-part workbook fixture.',
			recommendedOwnerAction:
				'Accept disclosed generated topology for local unknown-part routing proof only, while keeping the release gate missing until product and release approve the wording.',
			allowedUse:
				'Pre-hydration package-feature routing evidence that unknown package features require review before hydration.',
			forbiddenUse:
				'Do not claim arbitrary unknown-part preservation, understanding, recovery, trust, malware scanning, or signed provenance.',
		},
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			caseName: 'malformed',
			generatedKind: 'generated-malformed-package',
			replacementEvidence: `tracked safe-open scan rejected ${fixtureEvidence.safeOpen.rejected} fixture(s); malformed proof input remains generated structural bytes`,
			ownerDecisionNeeded:
				'Accept disclosed generated malformed-package bytes as rejection-path proof, or provide an approved public malformed workbook fixture.',
			recommendedOwnerAction:
				'Accept disclosed generated malformed bytes for fail-closed rejection-path proof only, while rejecting any recovery or vendor-repair wording.',
			allowedUse: 'Fail-closed malformed-package rejection evidence for the safe-open harness.',
			forbiddenUse:
				'Do not claim recovery of arbitrary malformed files, vendor repair equivalence, malware safety, or file trust.',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			caseName: 'signature-invalidation-drop',
			generatedKind: 'generated-edge-package',
			replacementEvidence: `tracked package-action scan found signaturePackage=${fixtureEvidence.packageAction.featureCounts.signaturePackage} across ${fixtureEvidence.packageAction.scanned} fixtures`,
			ownerDecisionNeeded:
				'Accept disclosed generated signature package topology as package-action invalidation proof, or provide an approved public signed workbook fixture.',
			recommendedOwnerAction:
				'Accept disclosed generated topology for local signature-part drop/invalidation accounting only, while keeping provenance and trust wording forbidden.',
			allowedUse:
				'Local package-action accounting evidence that signature package parts are dropped or invalidated after mutation.',
			forbiddenUse:
				'Do not claim signature preservation, re-signing, verification, attestation, SLSA, in-toto, or signed provenance.',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			caseName: 'unknown-part-error',
			generatedKind: 'generated-edge-package',
			replacementEvidence: `tracked package-action scan found syntheticUnknownPathFamily=${fixtureEvidence.packageAction.featureCounts.syntheticUnknownPathFamily} across ${fixtureEvidence.packageAction.scanned} fixtures`,
			ownerDecisionNeeded:
				'Accept disclosed generated unknown-part topology as fail-closed package-action proof, or provide an approved public unknown-part workbook fixture.',
			recommendedOwnerAction:
				'Accept disclosed generated topology for local fail-closed unknown-part package-action proof only, while keeping arbitrary preservation and trust wording forbidden.',
			allowedUse:
				'Local package-action accounting evidence that an unsupported unknown package part can fail closed with an explicit error action.',
			forbiddenUse:
				'Do not claim arbitrary unknown-part preservation, understanding, safe recovery, trust, or signed provenance.',
		},
	]
	return {
		ownerLoop: 'product',
		status: 'generated-structural-cases-disclosed-owner-approval-required',
		ownerApprovalRequired: true,
		allGeneratedStructuralCasesDisclosed: cases.every((entry) =>
			FIXTURE_POLICY.currentGeneratedStructuralCases[entry.artifact].includes(entry.caseName),
		),
		publicReplacementGapsRemain: fixtureEvidence.publicReplacementGapsRemain,
		validationCommand:
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
		cases,
		boundary:
			'Generated fixture decision evidence is a product-owner input. It does not approve generated fixtures as public binaries, satisfy fixture gates, prove license clearance, or support trust, safety, or signed-provenance claims.',
	}
}

function generatedFixtureDecisionEvidenceMarkdownRow(
	row: ReleaseProofGeneratedFixtureDecisionCase,
): string {
	return [
		row.artifact,
		row.gateId,
		row.caseName,
		row.generatedKind,
		row.replacementEvidence,
		row.ownerDecisionNeeded,
		row.recommendedOwnerAction,
		row.allowedUse,
		row.forbiddenUse,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneGeneratedFixtureDecisionEvidence(
	evidence: ReleaseProofGeneratedFixtureDecisionEvidence,
): ReleaseProofGeneratedFixtureDecisionEvidence {
	return {
		...evidence,
		cases: evidence.cases.map((entry) => ({ ...entry })),
	}
}

function formatFixtureFeatureCounts(
	counts: PackageActionFixtureScanResult['featureCounts'],
): string {
	return [
		`docPropsCore=${counts.docPropsCore}`,
		`docPropsCustom=${counts.docPropsCustom}`,
		`calcChain=${counts.calcChain}`,
		`customXml=${counts.customXml}`,
		`macro=${counts.macro}`,
		`chartOrDrawing=${counts.chartOrDrawing}`,
		`signaturePackage=${counts.signaturePackage}`,
		`syntheticUnknownPathFamily=${counts.syntheticUnknownPathFamily}`,
	].join(',')
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

function safeOpenLatencyValidationEvidence(
	result: SafeOpenProofResult,
): ReleaseProofSafeOpenLatencyValidationEvidence {
	const timedCases = result.cases.filter((entry) => entry.openPlanMedianMs !== undefined)
	const publicTimedCases = timedCases.filter((entry) => entry.kind === 'file')
	const generatedTimedCases = timedCases.filter((entry) => entry.kind !== 'file')
	return {
		artifact: 'safe-open-proof',
		gateId: 'release-latency-run',
		ownerLoop: 'performance',
		status:
			timedCases.length > 0
				? 'local-timed-diagnostic-owner-approval-required'
				: 'timed-evidence-absent-owner-run-required',
		ownerApprovalRequired: true,
		releaseClaimAllowed: false,
		thresholdClaimAllowed: false,
		validationCommand:
			'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json',
		repeat: result.repeat,
		warmup: result.warmup,
		timedCaseCount: timedCases.length,
		publicTimedCaseNames: publicTimedCases.map((entry) => entry.name),
		generatedTimedCaseNames: generatedTimedCases.map((entry) => entry.name),
		malformedRejected: result.cases.some(
			(entry) => entry.name === 'malformed' && entry.status === 'rejected',
		),
		publicOpenPlanMedianMs: numericCaseMetricByName(publicTimedCases, 'openPlanMedianMs'),
		publicFullOpenRatio: numericCaseMetricByName(publicTimedCases, 'fullOpenRatio'),
		missingPolicyRequirements: [
			'tracked-clean release environment',
			'standardized public input set',
			'approved repeat and warmup policy',
			'non-threshold release wording',
		],
		boundary:
			'Safe-open latency validation evidence is performance-owner routing data. It is not a release threshold, SLA, speed claim, or approval to use machine-specific local timing in public wording.',
	}
}

function numericCaseMetricByName(
	cases: readonly SafeOpenProofCaseResult[],
	metric: 'openPlanMedianMs' | 'fullOpenRatio',
): Readonly<Record<string, number>> {
	return Object.fromEntries(
		cases
			.map((entry) => [entry.name, entry[metric]] as const)
			.filter((entry): entry is readonly [string, number] => entry[1] !== undefined),
	)
}

function cloneSafeOpenLatencyValidationEvidence(
	evidence: ReleaseProofSafeOpenLatencyValidationEvidence,
): ReleaseProofSafeOpenLatencyValidationEvidence {
	return {
		...evidence,
		publicTimedCaseNames: [...evidence.publicTimedCaseNames],
		generatedTimedCaseNames: [...evidence.generatedTimedCaseNames],
		publicOpenPlanMedianMs: { ...evidence.publicOpenPlanMedianMs },
		publicFullOpenRatio: { ...evidence.publicFullOpenRatio },
		missingPolicyRequirements: [...evidence.missingPolicyRequirements],
	}
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

function correctnessBoundaryEvidenceMarkdownRow(
	row: ReleaseProofCorrectnessBoundaryFeatureCheck,
): string {
	return [
		row.feature,
		String(row.evidencePresent),
		row.evidenceSources.join('; '),
		row.proofChecks.join('; '),
		row.allowedWording,
		row.forbiddenWording,
	]
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

function cloneCorrectnessBoundaryEvidence(
	evidence: ReleaseProofCorrectnessBoundaryEvidence,
): ReleaseProofCorrectnessBoundaryEvidence {
	return {
		...evidence,
		featureChecks: evidence.featureChecks.map((entry) => ({
			...entry,
			evidenceSources: [...entry.evidenceSources],
			proofChecks: [...entry.proofChecks],
		})),
	}
}

function compactReportPublicationEvidenceMarkdownRow(
	row: ReleaseProofCompactReportPublicationEvidenceItem,
): string {
	return [
		row.artifact,
		row.gateId,
		`\`${row.command}\``,
		String(row.jsonBytes),
		row.topLevelFields.join(','),
		row.forbiddenPayloadFieldsPresent.join(',') || 'none',
		String(row.readyWhenGatePresent),
		String(row.generatedAtIncluded),
		String(row.headlineClaimAllowed),
		row.releaseGate,
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function compactReportPublicationPolicyDecisionMarkdownRow(
	row: ReleaseProofCompactReportPublicationPolicyDecision,
): string {
	return [
		row.requirement,
		row.ownerLoop,
		row.status,
		row.decisionNeeded,
		row.acceptanceEvidence,
		row.rejectIf,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneCompactReportPublicationEvidence(
	evidence: ReleaseProofCompactReportPublicationEvidence,
): ReleaseProofCompactReportPublicationEvidence {
	return {
		...evidence,
		missingPolicyRequirements: [...evidence.missingPolicyRequirements],
		policyDecisions: evidence.policyDecisions.map((decision) => ({ ...decision })),
		reports: evidence.reports.map((report) => ({
			...report,
			topLevelFields: [...report.topLevelFields],
			forbiddenPayloadFieldsPresent: [...report.forbiddenPayloadFieldsPresent],
		})),
	}
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

function portfolioClaimMarkdownRow(row: ReleaseProofPortfolioClaim): string {
	return [
		String(row.rank),
		row.claim,
		row.status,
		row.northStarLink,
		row.likelyHandoffOwner.join('+'),
		row.handoffDecision,
		row.proofCommand ?? 'none',
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
