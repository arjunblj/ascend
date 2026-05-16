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
	type SafeOpenProofTimingEnvironment,
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
	readonly fixtureAcquisitionPlan: ReleaseProofFixtureAcquisitionPlan
	readonly generatedFixtureDecisionEvidence: ReleaseProofGeneratedFixtureDecisionEvidence
	readonly performancePolicy: ReleaseProofPerformancePolicy
	readonly safeOpenLatencyValidationEvidence: ReleaseProofSafeOpenLatencyValidationEvidence
	readonly correctnessPolicy: ReleaseProofCorrectnessPolicy
	readonly correctnessBoundaryEvidence: ReleaseProofCorrectnessBoundaryEvidence
	readonly trustCompletenessBoundaryEvidence: ReleaseProofTrustCompletenessBoundaryEvidence
	readonly releasePackageabilityEvidence: ReleaseProofPackageabilityEvidence
	readonly streamingMatrixEvidence: ReleaseProofStreamingMatrixEvidence
	readonly compactReportPublicationEvidence: ReleaseProofCompactReportPublicationEvidence
	readonly readiness: ReleaseProofReadinessSummary
	readonly qssLeapfrogReleaseMatrix: ReleaseProofQssLeapfrogReleaseMatrix
	readonly releaseDecisionBoard: ReleaseProofReleaseDecisionBoard
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
	readonly fixtureAcquisitionPlan: ReleaseProofFixtureAcquisitionPlan
	readonly generatedFixtureDecisionEvidence: ReleaseProofGeneratedFixtureDecisionEvidence
	readonly performancePolicy: ReleaseProofPerformancePolicy
	readonly safeOpenLatencyValidationEvidence: ReleaseProofSafeOpenLatencyValidationEvidence
	readonly correctnessPolicy: ReleaseProofCorrectnessPolicy
	readonly correctnessBoundaryEvidence: ReleaseProofCorrectnessBoundaryEvidence
	readonly trustCompletenessBoundaryEvidence: ReleaseProofTrustCompletenessBoundaryEvidence
	readonly releasePackageabilityEvidence: ReleaseProofPackageabilityEvidence
	readonly streamingMatrixEvidence: ReleaseProofStreamingMatrixEvidence
	readonly compactReportPublicationEvidence: ReleaseProofCompactReportPublicationEvidence
	readonly nextOwnerActions: readonly ReleaseProofNextOwnerAction[]
	readonly claimBlockerBoard: readonly ReleaseProofClaimBlockerBoardRow[]
	readonly implementationHandoffs: readonly ReleaseProofImplementationHandoff[]
	readonly qssLeapfrogReleaseMatrix: ReleaseProofQssLeapfrogReleaseMatrix
	readonly releaseDecisionBoard: ReleaseProofReleaseDecisionBoard
	readonly claimPortfolio: readonly ReleaseProofPortfolioClaim[]
	readonly deferredClaims: readonly ReleaseProofDeferredClaim[]
	readonly excludedEvidence: readonly ReleaseProofIndexExcludedEvidence[]
	readonly boundary: string
}

export interface ReleaseProofFixtureDecisionPacket {
	readonly ownerLoop: 'product'
	readonly status: 'owner-decision-required'
	readonly releaseGate: ReleaseProofReadinessSummary['releaseGate']
	readonly headlineClaimsAllowed: boolean
	readonly ownerApprovalRequired: true
	readonly publicReplacementGapsRemain: boolean
	readonly trackedScans: readonly ReleaseProofFixtureDecisionTrackedScan[]
	readonly approvalChecklist: readonly ReleaseProofFixturePolicyApprovalItem[]
	readonly generatedCases: readonly ReleaseProofGeneratedFixtureDecisionCase[]
	readonly validationCommands: readonly string[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly forbiddenShortcuts: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofFixtureDecisionTrackedScan {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly gateId: 'public-edge-fixtures' | 'edge-fixture-policy'
	readonly command: string
	readonly corpus: 'tracked-git-fixtures'
	readonly scanned: number
	readonly replacementStatus: string
	readonly generatedStructuralCases: readonly string[]
	readonly publicReplacementGap: true
	readonly boundary: string
}

export interface ReleaseProofQssLeapfrogReleaseMatrix {
	readonly status: 'top-two-only'
	readonly northStar: string
	readonly competitor: 'Quadratic/QSS'
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly rows: readonly ReleaseProofQssLeapfrogReleaseMatrixRow[]
	readonly activeReleaseBlockers: readonly ReleaseProofClaimBlockerBoardRow[]
	readonly archivedResearchNotes: readonly ReleaseProofQssArchivedResearchNote[]
	readonly boundary: string
}

export interface ReleaseProofQssLeapfrogReleaseMatrixRow {
	readonly rank: number
	readonly artifact: ReleaseProofIndexArtifactName
	readonly claim: string
	readonly qssLikelyDoesWell: readonly string[]
	readonly ascendBetterWhereProven: readonly string[]
	readonly acceptedEvidence: readonly ReleaseProofQssAcceptedEvidenceItem[]
	readonly missingEvidence: readonly string[]
	readonly ownerActions: readonly ReleaseProofNextOwnerAction[]
	readonly claimsWeMustNotMake: readonly string[]
	readonly weakClaimDisposition: readonly ReleaseProofQssWeakClaimDisposition[]
	readonly boundary: string
}

export interface ReleaseProofQssAcceptedEvidenceItem {
	readonly evidenceId: string
	readonly kind: 'test' | 'benchmark' | 'proof-artifact' | 'rc-gate'
	readonly command: string
	readonly path: string
	readonly acceptedScope: string
	readonly boundary: string
}

export interface ReleaseProofQssWeakClaimDisposition {
	readonly weakClaim: string
	readonly disposition: 'downgrade' | 'blocker' | 'kill'
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly action: string
	readonly stopCondition: string
}

export interface ReleaseProofQssArchivedResearchNote {
	readonly name: ReleaseProofDeferredClaimName | ReleaseProofIndexExcludedEvidenceName
	readonly status: 'archived-research-note'
	readonly ownerLoops: readonly ReleaseProofReadinessOwner[]
	readonly reason: string
	readonly killCriterion: string
}

export interface ReleaseProofReleaseDecisionBoard {
	readonly status: 'top-two-only'
	readonly releaseGate: ReleaseProofReadinessSummary['releaseGate']
	readonly headlineClaimsAllowed: boolean
	readonly implementationSurfacePromotionAllowed: boolean
	readonly missingRequirementCount: number
	readonly rows: readonly ReleaseProofReleaseDecisionBoardRow[]
	readonly doNotPromoteYet: readonly ReleaseProofReleaseDecisionDoNotPromoteItem[]
	readonly boundary: string
}

export interface ReleaseProofReleaseDecisionBoardRow {
	readonly rank: number
	readonly artifact: ReleaseProofIndexArtifactName
	readonly claimWordingAllowedToday: string
	readonly evidenceWeHave: readonly ReleaseProofQssAcceptedEvidenceItem[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerActions: readonly ReleaseProofNextOwnerAction[]
	readonly ownerDecisionArtifacts: readonly ReleaseProofOwnerDecisionArtifact[]
	readonly headlineClaimAllowed: boolean
	readonly implementationSurfacePromotionAllowed: boolean
	readonly proofRequired: ReleaseProofClaimProofRequired
	readonly acceptedEvidence: readonly ReleaseProofQssAcceptedEvidenceItem[]
	readonly claimsWeMustNotMake: readonly string[]
	readonly aPlusBlockingOwnerActions: readonly ReleaseProofNextOwnerAction[]
	readonly boundary: string
}

export interface ReleaseProofOwnerDecisionArtifact {
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly artifactId: string
	readonly path: string
	readonly validationCommand: string
	readonly decision: string
	readonly nextAction: string
	readonly forbiddenShortcut: string
	readonly boundary: string
}

export interface ReleaseProofReleaseDecisionDoNotPromoteItem {
	readonly name: ReleaseProofQssArchivedResearchNote['name']
	readonly status: 'do-not-promote-yet'
	readonly ownerLoops: readonly ReleaseProofReadinessOwner[]
	readonly reason: string
	readonly killCriterion: string
	readonly boundary: string
}

export interface ReleaseProofPackageabilityEvidence {
	readonly ownerLoop: 'release'
	readonly status: 'local-tarball-smokes-present-publication-policy-required'
	readonly ownerApprovalRequired: true
	readonly sdkSmokeCommand: string
	readonly appSmokeCommand: string
	readonly rcGateCommand: string
	readonly coveredEvidence: readonly string[]
	readonly missingPolicyRequirements: readonly string[]
	readonly forbiddenClaims: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofFixtureAcquisitionPlan {
	readonly ownerLoop: 'product'
	readonly status: 'ranked-owner-review-required'
	readonly validationCommand: string
	readonly taskCount: number
	readonly tasks: readonly ReleaseProofFixtureAcquisitionTask[]
	readonly boundary: string
}

export interface ReleaseProofFixtureAcquisitionTask {
	readonly rank: number
	readonly caseName: 'signed-package' | 'malformed-package'
	readonly relatedArtifacts: readonly ReleaseProofIndexArtifactName[]
	readonly relatedGates: readonly string[]
	readonly task: string
	readonly evidenceAlreadyPresent: string
	readonly proofStillMissing: string
	readonly validationCommand: string
	readonly competitorOrSpecReference: string
	readonly killCriterion: string
	readonly ownerDecision: string
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
	readonly validationCommand: string
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
	readonly caseName: 'signed' | 'malformed'
	readonly generatedCaseKind: 'generated-edge-package' | 'generated-malformed-package'
	readonly acceptableAsTopologyProofWhen: string
	readonly requiresPublicBinaryWhen: string
	readonly validationCommand: string
	readonly gateEffect: 'keeps-public-edge-fixtures-missing-until-owner-approval'
}

export interface ReleaseProofPackageActionFixtureAcceptanceItem {
	readonly caseName: 'signature-invalidation-drop'
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
	readonly status: 'external-candidate-owner-review-required' | 'vendored-public-fixture'
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
	readonly gateEffect: 'does-not-satisfy-public-edge-fixtures' | 'satisfies-unknown-part-only'
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
	readonly externalCandidateEvidence: readonly ReleaseProofPackageActionExternalFixtureCandidateEvidence[]
	readonly boundary: string
}

export interface ReleaseProofPackageActionExternalFixtureCandidateEvidence {
	readonly artifact: 'package-action-proof'
	readonly gateId: 'edge-fixture-policy'
	readonly caseName: 'unknown-part-error'
	readonly status: 'external-candidate-owner-review-required' | 'vendored-public-fixture'
	readonly candidateId: string
	readonly sourceUrl: string
	readonly licenseEvidenceUrl: string
	readonly license: string
	readonly inputSha256: string
	readonly operationSummary: string
	readonly planWritePolicyOk: boolean
	readonly commitWritePolicyOk: boolean
	readonly postWriteAuditsPassed: boolean
	readonly actionCounts: Readonly<Record<ReleaseProofPackageActionKind, number>>
	readonly unknownPartPath: string
	readonly unknownPartContentType: string
	readonly unknownPartErrorAction: true
	readonly passthroughBytesEqual: true
	readonly issueCount: number
	readonly packageIssueRefs: readonly string[]
	readonly gateEffect: 'does-not-satisfy-edge-fixture-policy' | 'satisfies-unknown-part-only'
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
	readonly runProfile: ReleaseProofSafeOpenLatencyRunProfile
	readonly runProfileSatisfied: boolean
	readonly runProfileFailures: readonly string[]
	readonly repeat: number
	readonly warmup: number
	readonly timingEnvironmentCaptured: boolean
	readonly timingEnvironment?: SafeOpenProofTimingEnvironment
	readonly timedCaseCount: number
	readonly publicTimedCaseNames: readonly string[]
	readonly generatedTimedCaseNames: readonly string[]
	readonly malformedRejected: boolean
	readonly publicOpenPlanMedianMs: Readonly<Record<string, number>>
	readonly publicOpenPlanP95Ms: Readonly<Record<string, number>>
	readonly publicOpenPlanCv: Readonly<Record<string, number>>
	readonly publicFullOpenMedianMs: Readonly<Record<string, number>>
	readonly publicFullOpenP95Ms: Readonly<Record<string, number>>
	readonly publicFullOpenCv: Readonly<Record<string, number>>
	readonly publicFullOpenRatio: Readonly<Record<string, number>>
	readonly missingPolicyRequirements: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofSafeOpenLatencyRunProfile {
	readonly profileId: 'safe-open-release-latency-owner-review'
	readonly artifact: 'safe-open-proof'
	readonly gateId: 'release-latency-run'
	readonly ownerLoop: 'performance'
	readonly command: string
	readonly minimumRepeat: number
	readonly minimumWarmup: number
	readonly requiredCaseKind: 'file'
	readonly requiredPublicCaseNames: readonly string[]
	readonly requireTimingEnvironment: true
	readonly requiredMetrics: readonly string[]
	readonly cvGuard: ReleaseProofSafeOpenLatencyCvGuard
	readonly forbiddenUses: readonly string[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly boundary: string
}

export interface ReleaseProofSafeOpenLatencyCvGuard {
	readonly metric: 'publicOpenPlanCv'
	readonly maxRecommendedCv: number
	readonly onExceed: 'diagnostic-only-rerun-or-owner-review'
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

export interface ReleaseProofTrustCompletenessBoundaryEvidence {
	readonly ownerLoop: 'correctness'
	readonly status: 'boundary-pinned-owner-scope'
	readonly validationCommand: string
	readonly releaseTrustMatrixPath: string
	readonly outOfScopeClasses: readonly ReleaseProofTrustCompletenessBoundaryClass[]
	readonly requiredPromotionEvidence: string
	readonly doesNotCloseGates: readonly ReleaseProofReadinessOwner[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly boundary: string
}

export interface ReleaseProofTrustCompletenessBoundaryClass {
	readonly name: string
	readonly promoteOnlyWhen: string
	readonly ownerAction: string
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
	],
	approvalChecklist: [
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			ownerLoop: 'product',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Accept disclosed generated signed and malformed structural packages for guarded safe-open topology proof, or require public binary replacements.',
			acceptanceEvidence:
				'Safe-open proof labels generated signed and malformed cases, uses a vendored public unknown-part fixture, and claim wording excludes trust, malware scanning, active-content safety, and signed provenance.',
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
				'Accept disclosed generated signature topology for guarded package-action proof, or require a public signed workbook replacement.',
			acceptanceEvidence:
				'Package-action proof labels generated signature-invalidation, uses a vendored public unknown-part fixture for explicit error accounting, and claim wording stays limited to package action accounting.',
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
		'safe-open-proof': ['signed', 'malformed'],
		'package-action-proof': ['signature-invalidation-drop'],
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
			label: 'SLSA 1.2 build provenance distribution',
			url: 'https://slsa.dev/spec/v1.2/distributing-provenance',
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
			status: 'vendored-public-fixture',
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
				'Product/release must keep this MIT-package-manifest-backed workbook approved as a public unknown-part fixture with attribution policy.',
			gateEffect: 'satisfies-unknown-part-only',
			boundary:
				'Vendored public fixture evidence covers unknown-part package routing only. It does not address the signed-workbook fixture gap or prove arbitrary unknown-part preservation.',
		},
	]

const PACKAGE_ACTION_EXTERNAL_FIXTURE_CANDIDATES: readonly ReleaseProofPackageActionExternalFixtureCandidateEvidence[] =
	[
		{
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			caseName: 'unknown-part-error',
			status: 'vendored-public-fixture',
			candidateId: 'excelforge-book1-unknown-part-mutation',
			sourceUrl:
				'https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx',
			licenseEvidenceUrl:
				'https://raw.githubusercontent.com/node-projects/excelForge/master/package.json',
			license: 'MIT',
			inputSha256: '9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122',
			operationSummary: 'setCells Projekt 1!A1 = "probe"',
			planWritePolicyOk: false,
			commitWritePolicyOk: false,
			postWriteAuditsPassed: false,
			actionCounts: { passthrough: 42, regenerate: 6, add: 0, drop: 0, error: 1 },
			unknownPartPath: 'docMetadata/LabelInfo.xml',
			unknownPartContentType: 'application/vnd.ms-office.classificationlabels+xml',
			unknownPartErrorAction: true,
			passthroughBytesEqual: true,
			issueCount: 1,
			packageIssueRefs: ['Projekt 1!A1'],
			gateEffect: 'satisfies-unknown-part-only',
			boundary:
				'Vendored public mutation evidence covers explicit unknown-part error accounting only. It does not address signed-workbook evidence or prove arbitrary unknown-part preservation.',
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
				'Timed safe-open proof satisfies the machine-readable owner-review profile: approved public inputs, repeat/warmup policy, timing environment, required metrics, CV guard, and wording that reports observations without a release threshold.',
			rejectIf:
				'Run uses private corpora, dirty worktree state, one-off local timing, machine-specific ratios, or copy that implies a latency SLA.',
			validationCommand:
				'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'streaming-matrix-boundary',
			ownerLoop: 'performance',
			status: 'pending-owner-decision',
			decisionNeeded:
				'Accept representative streaming proofs covering passthrough/regenerate/add/drop for narrow wording, or require a broader streaming matrix before any parity claim.',
			acceptanceEvidence:
				'Package-action proof reports five streaming proof cases covering passthrough/regenerate/add/drop plus public macro/chart package accounting with passthrough-byte equality, and release wording says representative proofs only.',
			rejectIf:
				'Copy says full streaming parity, covers generated edge/error streaming behavior, or implies semantic preservation for unsupported workbook features.',
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
		{
			label: 'Google Benchmark repeated statistics',
			url: 'https://github.com/google/benchmark/blob/main/docs/user_guide.md',
		},
	],
	boundary:
		'Performance policy is an owner-decision aid for local benchmark evidence. It is not a release performance threshold, SLA, streaming parity claim, or production optimization mandate.',
}

const SAFE_OPEN_LATENCY_RUN_PROFILE: ReleaseProofSafeOpenLatencyRunProfile = {
	profileId: 'safe-open-release-latency-owner-review',
	artifact: 'safe-open-proof',
	gateId: 'release-latency-run',
	ownerLoop: 'performance',
	command: 'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
	minimumRepeat: 10,
	minimumWarmup: 3,
	requiredCaseKind: 'file',
	requiredPublicCaseNames: ['clean', 'formula-heavy', 'macro', 'pivot', 'activex', 'chart'],
	requireTimingEnvironment: true,
	requiredMetrics: [
		'openPlanSampleCount',
		'openPlanMedianMs',
		'openPlanP95Ms',
		'openPlanCv',
		'fullOpenSampleCount',
		'fullOpenMedianMs',
		'fullOpenP95Ms',
		'fullOpenCv',
		'fullOpenRatio',
	],
	cvGuard: {
		metric: 'publicOpenPlanCv',
		maxRecommendedCv: 0.25,
		onExceed: 'diagnostic-only-rerun-or-owner-review',
		boundary:
			'CV is an owner-review guardrail for noisy local timing. Exceeding it keeps the result diagnostic unless performance explicitly accepts the variance or reruns under cleaner conditions.',
	},
	forbiddenUses: [
		'release threshold',
		'SLA',
		'QSS performance comparison',
		'hardware-normalized benchmark',
		'private-corpus evidence',
		'generated-only input evidence',
	],
	sourceReferences: [
		{
			label: 'hyperfine warmup and run counts',
			url: 'https://man.archlinux.org/man/hyperfine.1.en',
		},
		{
			label: 'Google Benchmark repeated statistics and CV',
			url: 'https://github.com/google/benchmark/blob/main/docs/user_guide.md',
		},
	],
	boundary:
		'Owner-review profile only. Satisfying it can make local timing evidence reviewable, but it still does not authorize a release speed claim without performance and release approval.',
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
				'Public unknown-part evidence routes unknown package features to review or fails closed with explicit package-action errors.',
			allowedWording: 'Ascend can fail closed with an explicit unknown-part error.',
			forbiddenWording: 'Ascend preserves or understands arbitrary unknown parts.',
		},
		{
			feature: 'encrypted-files',
			currentEvidence:
				'Public Calamine password-protected fixture opens with the fixture password and rejects missing or wrong passwords before workbook hydration.',
			allowedWording:
				'Ascend can decrypt supported OOXML password-protected workbooks when the caller supplies the password, and fail closed on missing or wrong passwords.',
			forbiddenWording:
				'Ascend recovers passwords, removes protection, verifies file trust, scans encrypted files for malware, or supports every Excel encryption variant.',
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
	const readiness = releaseReadinessSummary(artifacts)
	const fixtureEvidence = fixturePolicyEvidence(safeOpenFixtureScan, packageActionFixtureScan)
	const qssMatrix = qssLeapfrogReleaseMatrix(artifacts, readiness)
	return {
		generatedAt: new Date().toISOString(),
		artifactCount: artifacts.length,
		excludedEvidenceCount: EXCLUDED_EVIDENCE.length,
		deferredClaimCount: DEFERRED_CLAIMS.length,
		signed: false,
		attestation: false,
		fixturePolicy: cloneFixturePolicy(),
		fixturePolicyEvidence: fixtureEvidence,
		fixtureAcquisitionPlan: fixtureAcquisitionPlan(),
		generatedFixtureDecisionEvidence: generatedFixtureDecisionEvidence(fixtureEvidence),
		performancePolicy: clonePerformancePolicy(),
		safeOpenLatencyValidationEvidence: safeOpenLatencyValidationEvidence(safeOpen),
		correctnessPolicy: cloneCorrectnessPolicy(),
		correctnessBoundaryEvidence: correctnessBoundaryEvidence(safeOpen, packageAction),
		trustCompletenessBoundaryEvidence: trustCompletenessBoundaryEvidence(),
		releasePackageabilityEvidence: releasePackageabilityEvidence(),
		streamingMatrixEvidence: streamingMatrixEvidence(packageAction),
		compactReportPublicationEvidence: compactReportPublicationEvidence(
			safeOpenCompact,
			packageActionCompact,
			artifacts,
		),
		readiness,
		qssLeapfrogReleaseMatrix: qssMatrix,
		releaseDecisionBoard: releaseDecisionBoard(artifacts, readiness, qssMatrix),
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
		'## QSS Leapfrog Release Matrix',
		'',
		`Status: ${result.qssLeapfrogReleaseMatrix.status}`,
		`Competitor: ${result.qssLeapfrogReleaseMatrix.competitor}`,
		result.qssLeapfrogReleaseMatrix.northStar,
		result.qssLeapfrogReleaseMatrix.boundary,
		'',
		'QSS source references:',
		...result.qssLeapfrogReleaseMatrix.sourceReferences.map(
			(reference) => `- ${reference.label}: ${reference.url}`,
		),
		'',
		'| Rank | Claim | QSS likely does well | Ascend better where proven | Accepted evidence | Missing evidence | Owner/action | Must not claim | Weak claim disposition | Boundary |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.qssLeapfrogReleaseMatrix.rows.map(qssLeapfrogReleaseMatrixMarkdownRow),
		'',
		'Active release blockers:',
		...result.qssLeapfrogReleaseMatrix.activeReleaseBlockers.map(
			(row) =>
				`- ${row.artifact}/${row.ownerLoop}: ${row.requirementIds.join(',')} (${row.nextStepKinds.join(',')})`,
		),
		'',
		'Archived research notes:',
		...result.qssLeapfrogReleaseMatrix.archivedResearchNotes.map(
			(note) => `- ${note.name}: ${note.reason}`,
		),
		'',
		'## Release Decision Board',
		'',
		result.releaseDecisionBoard.boundary,
		'',
		'| Rank | Claim | Evidence we have | Evidence missing | QSS contrast | Allowed wording | Forbidden wording | Next owner action | Owner decision artifact | Headline claim allowed | Implementation promotion allowed | Exact proof | Must not claim | A+ blocking owner action | Boundary |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.releaseDecisionBoard.rows.map(releaseDecisionBoardMarkdownRow),
		'',
		'Do not promote yet:',
		...result.releaseDecisionBoard.doNotPromoteYet.map((item) => `- ${item.name}: ${item.reason}`),
		'',
		'## Release Packageability Evidence',
		'',
		`Status: ${result.releasePackageabilityEvidence.status}`,
		`SDK smoke command: \`${result.releasePackageabilityEvidence.sdkSmokeCommand}\``,
		`App smoke command: \`${result.releasePackageabilityEvidence.appSmokeCommand}\``,
		`RC gate command: \`${result.releasePackageabilityEvidence.rcGateCommand}\``,
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
		'| Rank | Artifact | Gate | Owner loop | Priority | Next step | Validation command | Acceptance evidence | Forbidden shortcut |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- | --- |',
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
		...result.fixturePolicyEvidence.packageAction.externalCandidateEvidence.map(
			packageActionExternalCandidateEvidenceMarkdownRow,
		),
		'',
		'Fixture acquisition plan:',
		'',
		`Status: ${result.fixtureAcquisitionPlan.status}`,
		`Task count: ${result.fixtureAcquisitionPlan.taskCount}`,
		result.fixtureAcquisitionPlan.boundary,
		'',
		'| Rank | Case | Artifacts | Gates | Task | Evidence already present | Proof still missing | Validation | Reference | Kill criterion | Owner decision | Boundary |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.fixtureAcquisitionPlan.tasks.map(fixtureAcquisitionTaskMarkdownRow),
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
		`Run profile: ${result.safeOpenLatencyValidationEvidence.runProfile.profileId}`,
		`Run profile command: \`${result.safeOpenLatencyValidationEvidence.runProfile.command}\``,
		`Run profile satisfied: ${result.safeOpenLatencyValidationEvidence.runProfileSatisfied}`,
		`Run profile failures: ${result.safeOpenLatencyValidationEvidence.runProfileFailures.join('; ') || 'none'}`,
		`Run profile minimums: repeat ${result.safeOpenLatencyValidationEvidence.runProfile.minimumRepeat}, warmup ${result.safeOpenLatencyValidationEvidence.runProfile.minimumWarmup}`,
		`Run profile public cases: ${result.safeOpenLatencyValidationEvidence.runProfile.requiredPublicCaseNames.join(',')}`,
		`Run profile CV guard: ${result.safeOpenLatencyValidationEvidence.runProfile.cvGuard.metric} <= ${result.safeOpenLatencyValidationEvidence.runProfile.cvGuard.maxRecommendedCv}`,
		`Timed case count: ${result.safeOpenLatencyValidationEvidence.timedCaseCount}`,
		`Timing environment captured: ${result.safeOpenLatencyValidationEvidence.timingEnvironmentCaptured}`,
		`Public timed cases: ${result.safeOpenLatencyValidationEvidence.publicTimedCaseNames.join(',') || 'none'}`,
		`Generated timed cases: ${result.safeOpenLatencyValidationEvidence.generatedTimedCaseNames.join(',') || 'none'}`,
		`Public open-plan median ms: ${JSON.stringify(result.safeOpenLatencyValidationEvidence.publicOpenPlanMedianMs)}`,
		`Public open-plan p95 ms: ${JSON.stringify(result.safeOpenLatencyValidationEvidence.publicOpenPlanP95Ms)}`,
		`Public open-plan CV: ${JSON.stringify(result.safeOpenLatencyValidationEvidence.publicOpenPlanCv)}`,
		`Public full-open median ms: ${JSON.stringify(result.safeOpenLatencyValidationEvidence.publicFullOpenMedianMs)}`,
		`Public full-open p95 ms: ${JSON.stringify(result.safeOpenLatencyValidationEvidence.publicFullOpenP95Ms)}`,
		`Public full-open CV: ${JSON.stringify(result.safeOpenLatencyValidationEvidence.publicFullOpenCv)}`,
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
		'Correctness/trust completeness boundary:',
		'',
		`Status: ${result.trustCompletenessBoundaryEvidence.status}`,
		`Validation command: \`${result.trustCompletenessBoundaryEvidence.validationCommand}\``,
		`Matrix path: ${result.trustCompletenessBoundaryEvidence.releaseTrustMatrixPath}`,
		`Does not close gates: ${result.trustCompletenessBoundaryEvidence.doesNotCloseGates.join(',')}`,
		result.trustCompletenessBoundaryEvidence.requiredPromotionEvidence,
		result.trustCompletenessBoundaryEvidence.boundary,
		'',
		'| Out-of-scope class | Promote only when | Owner action |',
		'| --- | --- | --- |',
		...result.trustCompletenessBoundaryEvidence.outOfScopeClasses.map(
			trustCompletenessBoundaryMarkdownRow,
		),
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
		fixtureAcquisitionPlan: cloneFixtureAcquisitionPlan(result.fixtureAcquisitionPlan),
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
		trustCompletenessBoundaryEvidence: cloneTrustCompletenessBoundaryEvidence(
			result.trustCompletenessBoundaryEvidence,
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
		qssLeapfrogReleaseMatrix: cloneQssLeapfrogReleaseMatrix(result.qssLeapfrogReleaseMatrix),
		releaseDecisionBoard: cloneReleaseDecisionBoard(result.releaseDecisionBoard),
		claimPortfolio: result.claimPortfolio.map(clonePortfolioClaim),
		deferredClaims: result.deferredClaims,
		excludedEvidence: result.excludedEvidence,
		boundary:
			'Compact owner handoff index for release proof routing. It is not a release artifact bundle, signed attestation, or product surface.',
	}
}

export function releaseProofFixtureDecisionPacket(
	result: ReleaseProofIndexResult,
): ReleaseProofFixtureDecisionPacket {
	const productApprovalChecklist = FIXTURE_POLICY.approvalChecklist.filter(
		(item) => item.ownerLoop === 'product',
	)
	const trackedScans: ReleaseProofFixtureDecisionTrackedScan[] = [
		{
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			command: FIXTURE_POLICY.trackedFixtureScanCommands['safe-open-proof'],
			corpus: result.fixturePolicyEvidence.safeOpen.corpus,
			scanned: result.fixturePolicyEvidence.safeOpen.scanned,
			replacementStatus: result.fixturePolicyEvidence.safeOpen.replacementStatus,
			generatedStructuralCases: [
				...FIXTURE_POLICY.currentGeneratedStructuralCases['safe-open-proof'],
			],
			publicReplacementGap: true,
			boundary:
				'Tracked safe-open scan found a public unknown-part replacement, but no public signed-package replacement in the checked-in fixture corpus.',
		},
		{
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			command: FIXTURE_POLICY.trackedFixtureScanCommands['package-action-proof'],
			corpus: result.fixturePolicyEvidence.packageAction.corpus,
			scanned: result.fixturePolicyEvidence.packageAction.scanned,
			replacementStatus: result.fixturePolicyEvidence.packageAction.replacementStatus,
			generatedStructuralCases: [
				...FIXTURE_POLICY.currentGeneratedStructuralCases['package-action-proof'],
			],
			publicReplacementGap: true,
			boundary:
				'Tracked package-action scan found a public unknown-path replacement, but no public signature-package replacement in the checked-in fixture corpus.',
		},
	]
	return {
		ownerLoop: 'product',
		status: 'owner-decision-required',
		releaseGate: result.readiness.releaseGate,
		headlineClaimsAllowed: result.readiness.headlineClaimsAllowed,
		ownerApprovalRequired: true,
		publicReplacementGapsRemain: result.fixturePolicyEvidence.publicReplacementGapsRemain,
		trackedScans,
		approvalChecklist: productApprovalChecklist.map((item) => ({ ...item })),
		generatedCases: result.generatedFixtureDecisionEvidence.cases.map((entry) => ({ ...entry })),
		validationCommands: [
			...new Set([
				...trackedScans.map((scan) => scan.command),
				result.generatedFixtureDecisionEvidence.validationCommand,
			]),
		],
		sourceReferences: FIXTURE_POLICY.sourceReferences.map((entry) => ({ ...entry })),
		forbiddenShortcuts: [
			'Do not hide generated fixture provenance.',
			'Do not treat generated structural packages as public binary replacements.',
			'Do not use generated topology evidence for trust, malware scanning, active-content safety, signed provenance, attestation, or real-world vendor behavior claims.',
			'Do not vendor external candidate workbooks until license, attribution, privacy, and provenance are owner-approved.',
		],
		boundary:
			'Compact product fixture decision packet. It is not fixture approval, public binary provenance, release wording approval, or permission to add private or large workbook data.',
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
		rcGateCommand: 'bun run release:rc:gate',
		coveredEvidence: [
			'SDK tarball installs into a temp consumer and verifies create/open/plan/commit/reopen/check/recalc plus bundled agent docs.',
			'CLI/API/MCP app tarballs install into a temp consumer without workspace dependencies.',
			'Unified RC gate builds JS artifacts, packs SDK/CLI/API/MCP tarballs, installs only those tarballs into an isolated consumer app, rejects workspace/file dependency leakage, and runs SDK/CLI/API/MCP workbook proof.',
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

const QSS_SOURCE_REFERENCES: readonly ReleaseProofSourceReference[] = [
	{
		label: 'Quadratic docs: AI, Python, SQL, JavaScript spreadsheet',
		url: 'https://docs.quadratichq.com/',
	},
	{
		label: 'Quadratic docs: navigation, 60 FPS WASM/WebGL interaction',
		url: 'https://docs.quadratichq.com/spreadsheet/navigating',
	},
	{
		label: 'Quadratic docs: formulas',
		url: 'https://docs.quadratichq.com/formulas/getting-started',
	},
]

function qssLeapfrogReleaseMatrix(
	artifacts: readonly ReleaseProofIndexArtifact[],
	readiness: ReleaseProofReadinessSummary,
): ReleaseProofQssLeapfrogReleaseMatrix {
	return {
		status: 'top-two-only',
		northStar:
			'Position Ascend as the trust/proof/runtime layer for agentic spreadsheet work where local release evidence proves an advantage over QSS-style AI/code spreadsheet UX.',
		competitor: 'Quadratic/QSS',
		sourceReferences: QSS_SOURCE_REFERENCES.map((reference) => ({ ...reference })),
		rows: artifacts.map((artifact, index) =>
			qssLeapfrogReleaseMatrixRow(artifact, index + 1, readiness.nextOwnerActions),
		),
		activeReleaseBlockers: readiness.claimBlockerBoard.map(cloneClaimBlockerBoardRow),
		archivedResearchNotes: [
			...DEFERRED_CLAIMS.map(
				(claim): ReleaseProofQssArchivedResearchNote => ({
					name: claim.name,
					status: 'archived-research-note',
					ownerLoops: [...claim.ownerLoops],
					reason: claim.reason,
					killCriterion: claim.killCriterion,
				}),
			),
			...EXCLUDED_EVIDENCE.map(
				(evidence): ReleaseProofQssArchivedResearchNote => ({
					name: evidence.name,
					status: 'archived-research-note',
					ownerLoops: [evidence.ownerLoop],
					reason: evidence.reason,
					killCriterion: evidence.eligibilityRule,
				}),
			),
		],
		boundary:
			'This matrix is a release-priority gate for the top two claims only. Formula rename, token-bounded agent view, viewport history, columnar sidecars, oracle routing, and agent observability remain archived until they change top-claim implementation priority.',
	}
}

function qssLeapfrogReleaseMatrixRow(
	artifact: ReleaseProofIndexArtifact,
	rank: number,
	nextOwnerActions: readonly ReleaseProofNextOwnerAction[],
): ReleaseProofQssLeapfrogReleaseMatrixRow {
	const ownerActions = nextOwnerActions.filter((action) => action.artifact === artifact.name)
	const missingEvidence = artifact.readyWhen
		.filter((requirement) => requirement.status === 'missing')
		.map((requirement) => `${requirement.id}: ${requirement.requirement}`)
	if (artifact.name === 'safe-open-proof') {
		return {
			rank,
			artifact: artifact.name,
			claim: artifact.claim,
			qssLikelyDoesWell: [
				'AI/code-first spreadsheet UX with Python, SQL, JavaScript, formulas, and database connections.',
				'Browser-native interactive spreadsheet experience with WASM/WebGL navigation and collaboration-oriented product shape.',
			],
			ascendBetterWhereProven: [
				'Pre-hydration package-feature routing for unknown, macro, ActiveX, signature, and malformed workbook risk families.',
				'Local proof artifacts expose fixture provenance, generated-case disclosure, readyWhen gates, and explicit forbidden wording before release claims are allowed.',
				'Packageability evidence shows the existing SDK/CLI/API/MCP artifacts can be installed and exercised from local tarballs without workspace dependency leakage.',
			],
			acceptedEvidence: safeOpenQssEvidence(),
			missingEvidence,
			ownerActions: ownerActions.map(cloneNextOwnerAction),
			claimsWeMustNotMake: [
				'Malware scanning, active-content safety, sandboxing, trust certification, or Microsoft Protected View equivalence.',
				'Signed workbook verification, signer identity, SLSA, in-toto, Sigstore, or GitHub artifact attestation.',
				'Release latency, threshold, or QSS performance win before the performance owner runs the approved public-input validation.',
				'Complete public edge fixture coverage while signed or malformed cases remain generated or owner-unapproved.',
			],
			weakClaimDisposition: [
				{
					weakClaim: 'safe unknown workbook opening',
					disposition: 'downgrade',
					ownerLoop: 'product',
					action:
						'Use cautious pre-hydration package-feature routing wording until public-edge fixture policy is approved.',
					stopCondition:
						'Stop if generated signed or malformed cases are hidden or treated as public binaries.',
				},
				{
					weakClaim: 'QSS-beating open latency',
					disposition: 'blocker',
					ownerLoop: 'performance',
					action:
						'Run the tracked-clean release-latency validation over approved public inputs before any timing comparison.',
					stopCondition:
						'Stop if evidence uses private corpus inputs, generated-only inputs, or one-shot local timings.',
				},
				{
					weakClaim: 'safe file/trust/security claim',
					disposition: 'kill',
					ownerLoop: 'release',
					action:
						'Forbid security/trust wording; publish only package-feature routing boundaries after release approval.',
					stopCondition:
						'Stop if wording implies malware safety, sandboxing, signature trust, or attestation.',
				},
			],
			boundary:
				'Ascend can claim inspectable routing evidence for unknown workbooks, not user safety or full-fidelity opening of arbitrary Excel files.',
		}
	}
	return {
		rank,
		artifact: artifact.name,
		claim: artifact.claim,
		qssLikelyDoesWell: [
			'AI-assisted analysis, code cells, database-connected data workflows, and polished spreadsheet interaction.',
			'Modern browser spreadsheet UX that may be stronger for exploratory analysis than a proof-focused runtime layer.',
		],
		ascendBetterWhereProven: [
			'Per-part write accounting classifies workbook package changes as passthrough, regenerate, add, drop, or error.',
			'Proof artifacts expose source graph evidence, package journal issues, compact/full reports, and fail-closed boundaries for unsupported package features.',
			'Post-write verification separates expected generated table-part changes from unresolved package graph drift.',
		],
		acceptedEvidence: packageActionQssEvidence(),
		missingEvidence,
		ownerActions: ownerActions.map(cloneNextOwnerAction),
		claimsWeMustNotMake: [
			'Signed provenance, SLSA, in-toto, Sigstore, GitHub artifact attestation, tamper-evident storage, or registry publication.',
			'Semantic understanding or preservation of every unsupported workbook feature.',
			'Full streaming parity while the matrix is representative and owner-gated.',
			'Arbitrary unknown-part preservation or recovery when generated and external candidates remain policy-gated.',
		],
		weakClaimDisposition: [
			{
				weakClaim: 'auditable package-part mutation',
				disposition: 'downgrade',
				ownerLoop: 'correctness',
				action:
					'Use local per-part package action accounting wording until unsupported-feature boundaries are owner-approved.',
				stopCondition: 'Stop if wording claims semantic support for unsupported features.',
			},
			{
				weakClaim: 'complete streaming package-action parity',
				disposition: 'blocker',
				ownerLoop: 'performance',
				action:
					'Keep streaming wording representative unless performance expands and approves the streaming action matrix.',
				stopCondition: 'Stop if representative proofs are described as full streaming parity.',
			},
			{
				weakClaim: 'proof bundle as provenance or attestation',
				disposition: 'kill',
				ownerLoop: 'release',
				action:
					'Forbid provenance/attestation wording unless a real signed publication and verification workflow exists.',
				stopCondition:
					'Stop if local digests, compact reports, or tarball smoke output are described as signed provenance.',
			},
		],
		boundary:
			'Ascend can claim local package action accounting and fail-closed audit evidence, not universal Excel compatibility or signed provenance.',
	}
}

function safeOpenQssEvidence(): readonly ReleaseProofQssAcceptedEvidenceItem[] {
	return [
		{
			evidenceId: 'safe-open-proof-harness',
			kind: 'proof-artifact',
			command: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
			path: 'fixtures/benchmarks/safe-open-proof.ts',
			acceptedScope:
				'Local safe-open proof: case counts, review-before-hydration routes, fixture provenance, malformed rejection, and honest boundary fields.',
			boundary: 'No timing, malware, trust, Protected View, or signed-provenance claim.',
		},
		{
			evidenceId: 'safe-open-proof-tests',
			kind: 'test',
			command: 'bun test fixtures/benchmarks/safe-open-proof.test.ts',
			path: 'fixtures/benchmarks/safe-open-proof.test.ts',
			acceptedScope:
				'Committed regression coverage for safe-open proof shape and claim-safe output.',
			boundary: 'Tests validate proof shape, not external vendor behavior.',
		},
		{
			evidenceId: 'safe-open-fixture-scan',
			kind: 'benchmark',
			command: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			path: 'fixtures/benchmarks/safe-open-fixture-scan.ts',
			acceptedScope:
				'Tracked corpus scan showing public unknown-part fixture coverage and remaining signed fixture gap.',
			boundary: 'Tracked corpus evidence only; not proof that no public fixture exists elsewhere.',
		},
		{
			evidenceId: 'release-proof-index-owner-handoff',
			kind: 'proof-artifact',
			command:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			path: 'fixtures/benchmarks/release-proof-index.ts',
			acceptedScope:
				'Machine-readable readyWhen gates, active blockers, deferred claims, fixture policy, and release boundary.',
			boundary: 'Owner routing only; headlineClaimsAllowed remains false.',
		},
		{
			evidenceId: 'release-rc-gate',
			kind: 'rc-gate',
			command: 'bun run release:rc:gate',
			path: 'scripts/release-rc-gate.ts',
			acceptedScope:
				'Local RC packageability gate for SDK/CLI/API/MCP tarballs and installed workbook proof.',
			boundary: 'Local tarball proof only; not registry publication or attestation.',
		},
	]
}

function packageActionQssEvidence(): readonly ReleaseProofQssAcceptedEvidenceItem[] {
	return [
		{
			evidenceId: 'package-action-proof-harness',
			kind: 'proof-artifact',
			command: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
			path: 'fixtures/benchmarks/package-action-proof.ts',
			acceptedScope:
				'Local package-action proof for passthrough/regenerate/add/drop/error accounting, source graph evidence, and post-write audit status.',
			boundary:
				'No signed provenance, semantic support for every feature, or full streaming parity claim.',
		},
		{
			evidenceId: 'package-action-proof-tests',
			kind: 'test',
			command: 'bun test fixtures/benchmarks/package-action-proof.test.ts',
			path: 'fixtures/benchmarks/package-action-proof.test.ts',
			acceptedScope:
				'Committed regression coverage for package-action proof shape and compact report boundaries.',
			boundary: 'Tests validate proof shape and fixture cases, not public release policy.',
		},
		{
			evidenceId: 'package-action-fixture-scan',
			kind: 'benchmark',
			command: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			path: 'fixtures/benchmarks/package-action-fixture-scan.ts',
			acceptedScope:
				'Tracked corpus scan for package features and remaining generated signature/unknown edge cases.',
			boundary: 'Tracked corpus evidence only; generated edge cases remain owner-gated.',
		},
		{
			evidenceId: 'release-proof-index-owner-handoff',
			kind: 'proof-artifact',
			command:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			path: 'fixtures/benchmarks/release-proof-index.ts',
			acceptedScope:
				'Machine-readable readyWhen gates, unsupported-feature boundary, streaming matrix, provenance boundary, and owner actions.',
			boundary: 'Owner routing only; headlineClaimsAllowed remains false.',
		},
		{
			evidenceId: 'copy-sheet-table-package-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "prepared copySheet commits reopen workbook-unique table identities"',
			path: 'packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Committed post-write proof that generated copied table package parts are expected and reopened-valid.',
			boundary:
				'Specific table-copy package topology proof, not universal package mutation support.',
		},
		{
			evidenceId: 'release-rc-gate',
			kind: 'rc-gate',
			command: 'bun run release:rc:gate',
			path: 'scripts/release-rc-gate.ts',
			acceptedScope:
				'Local RC packageability gate for SDK/CLI/API/MCP tarballs and installed workbook proof.',
			boundary: 'Local tarball proof only; not registry publication or attestation.',
		},
	]
}

function releaseDecisionBoard(
	artifacts: readonly ReleaseProofIndexArtifact[],
	readiness: ReleaseProofReadinessSummary,
	qssMatrix: ReleaseProofQssLeapfrogReleaseMatrix,
): ReleaseProofReleaseDecisionBoard {
	return {
		status: 'top-two-only',
		releaseGate: readiness.releaseGate,
		headlineClaimsAllowed: readiness.headlineClaimsAllowed,
		implementationSurfacePromotionAllowed: readiness.implementationSurfacePromotionAllowed,
		missingRequirementCount: readiness.missingRequirementCount,
		rows: qssMatrix.rows.map((row) => {
			const artifact = artifacts.find((candidate) => candidate.name === row.artifact)
			const handoff = readiness.implementationHandoffs.find(
				(candidate) => candidate.artifact === row.artifact,
			)
			return {
				rank: row.rank,
				artifact: row.artifact,
				claimWordingAllowedToday: row.claim,
				evidenceWeHave: row.acceptedEvidence.map((item) => ({ ...item })),
				evidenceMissing: [...row.missingEvidence],
				qssContrast: releaseDecisionQssContrast(row),
				allowedWording: releaseDecisionAllowedWording(row.artifact),
				forbiddenWording: [...row.claimsWeMustNotMake],
				nextOwnerActions: row.ownerActions.map(cloneNextOwnerAction),
				ownerDecisionArtifacts: ownerDecisionArtifactsFor(row.artifact),
				headlineClaimAllowed: artifact?.headlineClaimAllowed ?? false,
				implementationSurfacePromotionAllowed:
					handoff?.implementationSurfacePromotionAllowed ??
					readiness.implementationSurfacePromotionAllowed,
				proofRequired: claimProofRequired(row.artifact),
				acceptedEvidence: row.acceptedEvidence.map((item) => ({ ...item })),
				claimsWeMustNotMake: [...row.claimsWeMustNotMake],
				aPlusBlockingOwnerActions: row.ownerActions.map(cloneNextOwnerAction),
				boundary:
					'Release decision row only. It names allowed local claim wording, exact proof pointers, forbidden shortcuts, and owner blockers without satisfying any gate.',
			}
		}),
		doNotPromoteYet: qssMatrix.archivedResearchNotes.map((note) => ({
			name: note.name,
			status: 'do-not-promote-yet',
			ownerLoops: [...note.ownerLoops],
			reason: note.reason,
			killCriterion: note.killCriterion,
			boundary:
				'Archived research note for release stewardship. Do not turn this into release wording or a new implementation surface until it changes the top-two claim gate.',
		})),
		boundary:
			'Top-two release-decision artifact for claim stewardship. It is derived from committed release proof gates and must not be treated as a product surface, benchmark threshold, signed provenance, or owner approval.',
	}
}

function releaseDecisionBoardMarkdownRow(row: ReleaseProofReleaseDecisionBoardRow): string {
	return [
		String(row.rank),
		row.claimWordingAllowedToday,
		row.evidenceWeHave
			.map((item) => `${item.evidenceId}=\`${item.command}\` (${item.path})`)
			.join('; '),
		row.evidenceMissing.join('; '),
		row.qssContrast.join('; '),
		row.allowedWording,
		row.forbiddenWording.join('; '),
		row.nextOwnerActions
			.map(
				(action) =>
					`${action.ownerLoop}/${action.requirementId}: ${action.nextStepKind}=\`${action.validationCommand}\``,
			)
			.join('; '),
		row.ownerDecisionArtifacts
			.map(
				(artifact) =>
					`${artifact.ownerLoop}/${artifact.artifactId}: \`${artifact.validationCommand}\` (${artifact.path})`,
			)
			.join('; '),
		String(row.headlineClaimAllowed),
		String(row.implementationSurfacePromotionAllowed),
		row.acceptedEvidence
			.map((item) => `${item.evidenceId}=\`${item.command}\` (${item.path})`)
			.join('; '),
		row.claimsWeMustNotMake.join('; '),
		row.aPlusBlockingOwnerActions
			.map((action) => `${action.ownerLoop}/${action.requirementId}: ${action.nextStepKind}`)
			.join('; '),
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneReleaseDecisionBoard(
	board: ReleaseProofReleaseDecisionBoard,
): ReleaseProofReleaseDecisionBoard {
	return {
		...board,
		doNotPromoteYet: board.doNotPromoteYet.map((item) => ({
			...item,
			ownerLoops: [...item.ownerLoops],
		})),
		rows: board.rows.map((row) => ({
			...row,
			proofRequired: { ...row.proofRequired },
			evidenceWeHave: row.evidenceWeHave.map((item) => ({ ...item })),
			evidenceMissing: [...row.evidenceMissing],
			qssContrast: [...row.qssContrast],
			forbiddenWording: [...row.forbiddenWording],
			nextOwnerActions: row.nextOwnerActions.map(cloneNextOwnerAction),
			ownerDecisionArtifacts: row.ownerDecisionArtifacts.map(cloneOwnerDecisionArtifact),
			acceptedEvidence: row.acceptedEvidence.map((item) => ({ ...item })),
			claimsWeMustNotMake: [...row.claimsWeMustNotMake],
			aPlusBlockingOwnerActions: row.aPlusBlockingOwnerActions.map(cloneNextOwnerAction),
		})),
	}
}

function cloneOwnerDecisionArtifact(
	artifact: ReleaseProofOwnerDecisionArtifact,
): ReleaseProofOwnerDecisionArtifact {
	return { ...artifact }
}

function ownerDecisionArtifactsFor(
	artifact: ReleaseProofIndexArtifactName,
): readonly ReleaseProofOwnerDecisionArtifact[] {
	switch (artifact) {
		case 'safe-open-proof':
			return [
				{
					ownerLoop: 'correctness',
					artifactId: 'excel-behavior-compatibility-matrix',
					path: 'docs/EXCEL_BEHAVIOR_COMPATIBILITY_MATRIX.md',
					validationCommand:
						'bun test packages/sdk/src/excel-behavior-compatibility-matrix.test.ts',
					decision:
						'Use the ranked compatibility matrix for safe wording on common workbook support, active content, unknown parts, encryption, malformed packages, and signature blockers.',
					nextAction:
						'Compatibility owner adds the next real/public workbook gap from the matrix; product/release owners keep signed and generated malformed/signature policy blockers explicit.',
					forbiddenShortcut:
						'Do not turn the matrix into full Excel compatibility, file trust, malware scanning, password recovery, or signature preservation wording.',
					boundary:
						'Owner decision artifact only. It routes compatibility work and claim wording without satisfying release proof gates by itself.',
				},
				{
					ownerLoop: 'performance',
					artifactId: 'performance-claim-baseline-matrix',
					path: 'docs/PERFORMANCE_CLAIM_BASELINE_MATRIX.md',
					validationCommand:
						'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts',
					decision:
						'Use the performance matrix as a defer decision: no broad XLSX read, SOTA, or QSS-leapfrog speed claim is promotable from the current partial baseline.',
					nextAction:
						'Benchmarking owner runs or expands the full clean `xlsx-read-sota` profile before optimizing or promoting speed wording.',
					forbiddenShortcut:
						'Do not count unavailable runners, blocked runners, dirty-worktree timings, or one-workload medians as speed wins.',
					boundary:
						'Owner decision artifact only. It blocks weak performance wording and names the next benchmark loop; it is not a release speed claim.',
				},
			]
		case 'package-action-proof':
			return [
				{
					ownerLoop: 'correctness',
					artifactId: 'excel-behavior-compatibility-matrix',
					path: 'docs/EXCEL_BEHAVIOR_COMPATIBILITY_MATRIX.md',
					validationCommand:
						'bun test packages/sdk/src/excel-behavior-compatibility-matrix.test.ts',
					decision:
						'Use the compatibility matrix to keep package-action wording preservation-first for visual, pivot, active-content, external-link, unknown-part, encrypted, malformed, and signature surfaces.',
					nextAction:
						'Correctness owner promotes an individual package surface only after real/public open -> inspect -> edit -> save/reopen -> verify evidence exists.',
					forbiddenShortcut:
						'Do not claim semantic support for every unsupported package feature, byte passthrough for chart XML, signature verification, or arbitrary unknown-part preservation.',
					boundary:
						'Owner decision artifact only. It constrains package-action wording without replacing package-action proof gates.',
				},
			]
	}
}

function releaseDecisionQssContrast(
	row: ReleaseProofQssLeapfrogReleaseMatrixRow,
): readonly string[] {
	return [
		...row.qssLikelyDoesWell.map((entry) => `QSS likely does well: ${entry}`),
		...row.ascendBetterWhereProven.map((entry) => `Ascend proven today: ${entry}`),
	]
}

function releaseDecisionAllowedWording(artifact: ReleaseProofIndexArtifactName): string {
	switch (artifact) {
		case 'safe-open-proof':
			return 'Ascend provides local pre-hydration package-feature routing evidence for unknown workbook risk families, with generated fixture provenance disclosed and headline safety wording blocked.'
		case 'package-action-proof':
			return 'Ascend provides local per-part package action accounting for workbook edits, with unsupported feature boundaries and representative streaming scope disclosed.'
	}
}

function qssLeapfrogReleaseMatrixMarkdownRow(row: ReleaseProofQssLeapfrogReleaseMatrixRow): string {
	return [
		String(row.rank),
		row.claim,
		row.qssLikelyDoesWell.join('; '),
		row.ascendBetterWhereProven.join('; '),
		row.acceptedEvidence
			.map((item) => `${item.evidenceId}=\`${item.command}\` (${item.path})`)
			.join('; '),
		row.missingEvidence.join('; '),
		row.ownerActions
			.map((action) => `${action.ownerLoop}/${action.requirementId}: ${action.nextStepKind}`)
			.join('; '),
		row.claimsWeMustNotMake.join('; '),
		row.weakClaimDisposition
			.map((item) => `${item.disposition}:${item.weakClaim}->${item.ownerLoop}`)
			.join('; '),
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneQssLeapfrogReleaseMatrix(
	matrix: ReleaseProofQssLeapfrogReleaseMatrix,
): ReleaseProofQssLeapfrogReleaseMatrix {
	return {
		...matrix,
		sourceReferences: matrix.sourceReferences.map((reference) => ({ ...reference })),
		rows: matrix.rows.map((row) => ({
			...row,
			qssLikelyDoesWell: [...row.qssLikelyDoesWell],
			ascendBetterWhereProven: [...row.ascendBetterWhereProven],
			acceptedEvidence: row.acceptedEvidence.map((item) => ({ ...item })),
			missingEvidence: [...row.missingEvidence],
			ownerActions: row.ownerActions.map(cloneNextOwnerAction),
			claimsWeMustNotMake: [...row.claimsWeMustNotMake],
			weakClaimDisposition: row.weakClaimDisposition.map((item) => ({ ...item })),
		})),
		activeReleaseBlockers: matrix.activeReleaseBlockers.map(cloneClaimBlockerBoardRow),
		archivedResearchNotes: matrix.archivedResearchNotes.map((note) => ({
			...note,
			ownerLoops: [...note.ownerLoops],
		})),
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
			'signed and malformed cases remain disclosed generated topology/rejection packages, not public binary fixtures',
			'local timing evidence is proof-run data, not a release performance threshold',
		],
		readyWhen: [
			{
				id: 'public-edge-fixtures',
				status: 'missing',
				ownerLoop: 'product',
				requirement:
					'replace generated signed/malformed packages with public binary fixtures or explicitly approve disclosed generated edge packages',
				evidence:
					'safe-open proof uses public unknown-part and encrypted password workbooks, but still lacks a public signed workbook fixture and keeps malformed rejection generated',
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
					'accept disclosed generated signature topology as release proof or replace it with a public binary fixture',
				evidence:
					'package-action fixture scan over tracked fixtures finds public docProps/calc-chain/customXml/macro/chart/unknown-path candidates but 0 signature-package replacements; current proof still uses a generated signature-invalidation edge package',
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
					'current proof reports five streamingProofCases covering passthrough/regenerate/add/drop plus public macro/chart package accounting, with generated edge/error cases still non-streaming',
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
				'generated signature topology is disclosed',
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
				'public unknown-part fixture is used',
				'commit proof records an error action for the unknown package part',
				'post-write audit fails closed and safe-open routes unknown package features to review',
			],
		}),
		correctnessBoundaryFeatureCheck({
			feature: 'encrypted-files',
			evidencePresent:
				safeOpenCaseOk(safeOpen, 'encrypted-password') &&
				safeOpenCaseRejected(safeOpen, 'encrypted-missing-password') &&
				safeOpenCaseRejected(safeOpen, 'encrypted-wrong-password'),
			evidenceSources: [
				'safe-open-proof/encrypted-password',
				'safe-open-proof/encrypted-missing-password',
				'safe-open-proof/encrypted-wrong-password',
			],
			proofChecks: [
				'public encrypted Calamine fixture is used',
				'valid password opens with full-mode planning',
				'missing or wrong password rejects before hydration',
				'boundary does not claim password recovery, protection removal, malware scanning, or file trust',
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
				'package-action-proof/macro-passthrough',
				'package-action-proof/chart-sidecar-accounting',
			],
			proofChecks: [
				'representative streaming proofs cover passthrough/regenerate/add/drop',
				'public macro and chart cases have streaming package-action proof coverage',
				'streaming proof records regenerated worksheet parts and passthrough byte equality',
				'generated edge error and signature streaming remain outside the proof boundary',
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

function trustCompletenessBoundaryEvidence(): ReleaseProofTrustCompletenessBoundaryEvidence {
	return {
		ownerLoop: 'correctness',
		status: 'boundary-pinned-owner-scope',
		validationCommand: 'bun test packages/sdk/src/release-trust-matrix.test.ts',
		releaseTrustMatrixPath: 'docs/RELEASE_TRUST_MATRIX.md',
		outOfScopeClasses: [
			{
				name: 'Broad formula function coverage',
				promoteOnlyWhen:
					'a supported edit can write stale formula bindings, stale analysis, or a wrong cached result',
				ownerAction:
					'Keep generic function breadth in formula correctness work, not release trust scope.',
			},
			{
				name: 'Product/DX orchestration such as progressive open or viewport merge helpers',
				promoteOnlyWhen: 'the orchestration hides journal, verifier, or package drift from agents',
				ownerAction:
					'Route ergonomics work to product/DX unless it changes fail-closed trust evidence.',
			},
			{
				name: 'Reader/writer performance and benchmark tuning',
				promoteOnlyWhen:
					'the benchmark work preserves the same commit/reopen/verify contract or exposes correctness drift',
				ownerAction:
					'Keep throughput claims in benchmark-owned proof and out of correctness release scope.',
			},
			{
				name: 'More malformed-field enumeration',
				promoteOnlyWhen:
					'the malformed field can mutate the wrong surface, corrupt an exact inverse, or be silently blessed after write',
				ownerAction: 'Reject enumeration work that does not change a supported edit trust path.',
			},
			{
				name: 'New unknown Excel feature implementation',
				promoteOnlyWhen:
					'unknown package state is not preserved, reported honestly, or kept blocking on drift',
				ownerAction:
					'Require machine-readable preservation or drift evidence before expanding feature scope.',
			},
		],
		requiredPromotionEvidence:
			'Promote a new correctness item only when it names a top release claim and proves a silent corruption, exact-journal, or post-write drift path not already covered.',
		doesNotCloseGates: ['product', 'performance', 'release'],
		sourceReferences: [
			{
				label: 'SLSA 1.2 build provenance distribution',
				url: 'https://slsa.dev/spec/v1.2/distributing-provenance',
			},
			{
				label: 'GitHub artifact attestations',
				url: 'https://docs.github.com/en/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds',
			},
			{
				label: 'Microsoft Protected View',
				url: 'https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653',
			},
		],
		boundary:
			'Correctness/trust completeness boundary only. It does not approve headline release claims, product surfaces, performance wording, fixture policy, provenance, or publication gates.',
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
			'Streaming matrix evidence proves representative streaming package-action cases covering passthrough/regenerate/add/drop plus public macro/chart package accounting. It does not prove full streaming parity, generated edge/error streaming behavior, or semantic preservation for unsupported workbook features.',
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

function safeOpenCaseOk(result: SafeOpenProofResult, name: string): boolean {
	return result.cases.some((entry) => entry.name === name && entry.status === 'ok')
}

function safeOpenCaseRejected(result: SafeOpenProofResult, name: string): boolean {
	return result.cases.some((entry) => entry.name === name && entry.status === 'rejected')
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

function cloneClaimBlockerBoardRow(
	row: ReleaseProofClaimBlockerBoardRow,
): ReleaseProofClaimBlockerBoardRow {
	return {
		...row,
		requirementIds: [...row.requirementIds],
		actionRanks: [...row.actionRanks],
		nextStepKinds: [...row.nextStepKinds],
		acceptanceEvidence: [...row.acceptanceEvidence],
		forbiddenShortcuts: [...row.forbiddenShortcuts],
	}
}

function claimProofRequired(
	artifact: ReleaseProofIndexArtifactName,
): ReleaseProofClaimProofRequired {
	switch (artifact) {
		case 'safe-open-proof':
			return {
				fixture:
					'Public clean, formula-heavy, macro, pivot, ActiveX, chart, unknown-part, encrypted valid-password and fail-closed password-rejection, plus signed and malformed workbook/package cases; generated signed and malformed cases must stay disclosed unless replaced by public binary fixtures.',
				benchmark:
					'Release-environment open-plan latency over standardized public inputs, compared with full hydration and approved threshold wording.',
				surface:
					'Existing SDK open planner, CLI open-plan, API open-plan endpoint, and MCP open-plan tool only; no new opener surface.',
				validationGate:
					'Run the safe-open proof harness, focused SDK/CLI/API/MCP open-plan tests, malformed package checks, release-proof-index, typecheck, Biome, and changed tests when code changes.',
				competitorContrast:
					'Microsoft Protected View is trust/read-only UX for potentially unsafe files; Ascend claim is OSS pre-hydration package-feature routing.',
				honestBoundary:
					'Not malware scanning, sandboxing, file trust, active-content safety, signed provenance, or malformed-package recovery.',
				killCriterion:
					'Do not publish headline wording if generated signed or malformed packages are hidden, product rejects disclosed generated topology fixtures, or latency wording lacks an approved release-environment run.',
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
				validationCommand:
					artifact === 'safe-open-proof'
						? 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json'
						: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
				acceptanceEvidence:
					artifact === 'safe-open-proof'
						? 'Product accepts disclosed generated signed/malformed structural packages for guarded topology proof, while the vendored public unknown-part fixture remains approved.'
						: 'Product accepts disclosed generated signature topology for guarded package-action proof, while the vendored public unknown-part fixture remains approved.',
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
				validationCommand:
					'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
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
				validationCommand:
					'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
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
				validationCommand:
					'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
				acceptanceEvidence:
					'Performance accepts representative streaming proofs covering passthrough/regenerate/add/drop plus public macro/chart package accounting for narrow wording, or expands the matrix to generated edge/error cases.',
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
				validationCommand:
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
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
				validationCommand:
					artifact === 'safe-open-proof'
						? 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json'
						: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json',
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
				validationCommand:
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
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
				`${action.rank}:${action.artifact}/${action.requirementId}(${action.ownerLoop},${action.priority},${action.nextStepKind};cmd=${action.validationCommand};accept=${action.acceptanceEvidence};forbid=${action.forbiddenShortcut})`,
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
		`\`${action.validationCommand}\``,
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
		packageAction.featureCounts.unknownPathFamily === 0 ? 'unknownPathFamily' : undefined,
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
		externalCandidateEvidence: PACKAGE_ACTION_EXTERNAL_FIXTURE_CANDIDATES.map((entry) => ({
			...entry,
			actionCounts: { ...entry.actionCounts },
			packageIssueRefs: [...entry.packageIssueRefs],
		})),
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

function packageActionExternalCandidateEvidenceMarkdownRow(
	row: ReleaseProofPackageActionExternalFixtureCandidateEvidence,
): string {
	return [
		row.artifact,
		row.gateId,
		row.caseName,
		row.status,
		`${row.candidateId} (${row.sourceUrl})`,
		`${row.license} (${row.licenseEvidenceUrl})`,
		row.inputSha256,
		`error=${row.actionCounts.error}; auditsPassed=${row.postWriteAuditsPassed}; part=${row.unknownPartPath}; contentType=${row.unknownPartContentType}; refs=${row.packageIssueRefs.join(',')}`,
		row.gateEffect,
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function fixtureAcquisitionPlan(): ReleaseProofFixtureAcquisitionPlan {
	return {
		ownerLoop: 'product',
		status: 'ranked-owner-review-required',
		validationCommand:
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
		taskCount: 2,
		tasks: [
			{
				rank: 1,
				caseName: 'signed-package',
				relatedArtifacts: ['safe-open-proof', 'package-action-proof'],
				relatedGates: ['public-edge-fixtures', 'edge-fixture-policy'],
				task: 'Acquire or generate under explicit policy a license-clear public signed XLSX package fixture.',
				evidenceAlreadyPresent:
					'Tracked fixture scans now include a public unknown-part match for safe-open, but still find signaturePackage=0 for package-action and no public signed workbook fixture.',
				proofStillMissing:
					'Approved public signed workbook bytes or owner acceptance of generated signature topology, plus wording that excludes signature verification, trust, and attestation.',
				validationCommand:
					'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json && bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
				competitorOrSpecReference:
					'OPC digital signatures use signature origin and signature XML parts, but signature trust must be handled by the package consumer.',
				killCriterion:
					'Do not use a signed workbook if provenance, redistribution rights, private contents, certificate meaning, or trust wording cannot be reviewed.',
				ownerDecision:
					'Product/release either accepts disclosed generated signature topology for narrow package-routing proof or sources an approved public signed workbook.',
				boundary:
					'This is topology evidence only. It must not become a signature validation, identity, malware safety, or signed-provenance claim.',
			},
			{
				rank: 2,
				caseName: 'malformed-package',
				relatedArtifacts: ['safe-open-proof'],
				relatedGates: ['public-edge-fixtures'],
				task: 'Decide whether generated malformed bytes are sufficient for fail-closed rejection proof or whether a public malformed workbook fixture is required.',
				evidenceAlreadyPresent:
					'Safe-open proof already includes a generated malformed package rejection path and the tracked scan reports one rejected fixture.',
				proofStillMissing:
					'Owner approval for generated bad bytes as rejection-path proof, or an approved public malformed fixture with clear redistribution rights.',
				validationCommand: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
				competitorOrSpecReference:
					'OPC defines package structure; malformed package behavior is an application rejection boundary, not a repair guarantee.',
				killCriterion:
					'Do not publish recovery, repair, vendor-equivalence, malware safety, or file-trust wording from malformed rejection evidence.',
				ownerDecision:
					'Product/release accepts generated malformed rejection proof for narrow fail-closed wording or requires a public malformed workbook.',
				boundary:
					'Malformed proof supports fail-closed rejection only and is lower leverage than unknown-part or signed fixture acquisition.',
			},
		],
		boundary:
			'Fixture acquisition planning ranks owner work for public proof gaps. It is not fixture approval, not license review, and not gate satisfaction.',
	}
}

function fixtureAcquisitionTaskMarkdownRow(row: ReleaseProofFixtureAcquisitionTask): string {
	return [
		String(row.rank),
		row.caseName,
		row.relatedArtifacts.join(','),
		row.relatedGates.join(','),
		row.task,
		row.evidenceAlreadyPresent,
		row.proofStillMissing,
		row.validationCommand,
		row.competitorOrSpecReference,
		row.killCriterion,
		row.ownerDecision,
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function cloneFixtureAcquisitionPlan(
	plan: ReleaseProofFixtureAcquisitionPlan,
): ReleaseProofFixtureAcquisitionPlan {
	return {
		...plan,
		tasks: plan.tasks.map((task) => ({
			...task,
			relatedArtifacts: [...task.relatedArtifacts],
			relatedGates: [...task.relatedGates],
		})),
	}
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
			externalCandidateEvidence: evidence.packageAction.externalCandidateEvidence.map((entry) => ({
				...entry,
				actionCounts: { ...entry.actionCounts },
				packageIssueRefs: [...entry.packageIssueRefs],
			})),
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
		`unknownPathFamily=${counts.unknownPathFamily}`,
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
	const publicOpenPlanMedianMs = numericCaseMetricByName(publicTimedCases, 'openPlanMedianMs')
	const publicOpenPlanP95Ms = numericCaseMetricByName(publicTimedCases, 'openPlanP95Ms')
	const publicOpenPlanCv = numericCaseMetricByName(publicTimedCases, 'openPlanCv')
	const publicFullOpenMedianMs = numericCaseMetricByName(publicTimedCases, 'fullOpenMedianMs')
	const publicFullOpenP95Ms = numericCaseMetricByName(publicTimedCases, 'fullOpenP95Ms')
	const publicFullOpenCv = numericCaseMetricByName(publicTimedCases, 'fullOpenCv')
	const publicFullOpenRatio = numericCaseMetricByName(publicTimedCases, 'fullOpenRatio')
	const runProfileFailures = safeOpenLatencyRunProfileFailures(result, publicTimedCases, {
		publicOpenPlanMedianMs,
		publicOpenPlanP95Ms,
		publicOpenPlanCv,
		publicFullOpenMedianMs,
		publicFullOpenP95Ms,
		publicFullOpenCv,
		publicFullOpenRatio,
	})
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
			'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
		runProfile: cloneSafeOpenLatencyRunProfile(SAFE_OPEN_LATENCY_RUN_PROFILE),
		runProfileSatisfied: runProfileFailures.length === 0,
		runProfileFailures,
		repeat: result.repeat,
		warmup: result.warmup,
		timingEnvironmentCaptured: result.timingEnvironment !== undefined,
		...(result.timingEnvironment ? { timingEnvironment: result.timingEnvironment } : {}),
		timedCaseCount: timedCases.length,
		publicTimedCaseNames: publicTimedCases.map((entry) => entry.name),
		generatedTimedCaseNames: generatedTimedCases.map((entry) => entry.name),
		malformedRejected: result.cases.some(
			(entry) => entry.name === 'malformed' && entry.status === 'rejected',
		),
		publicOpenPlanMedianMs,
		publicOpenPlanP95Ms,
		publicOpenPlanCv,
		publicFullOpenMedianMs,
		publicFullOpenP95Ms,
		publicFullOpenCv,
		publicFullOpenRatio,
		missingPolicyRequirements: [
			'tracked-clean release environment',
			'standardized public input set',
			'performance-owner approval of the safe-open release-latency owner-review profile',
			'non-threshold release wording',
		],
		boundary:
			'Safe-open latency validation evidence is performance-owner routing data. It is not a release threshold, SLA, speed claim, or approval to use machine-specific local timing in public wording.',
	}
}

function safeOpenLatencyRunProfileFailures(
	result: SafeOpenProofResult,
	publicTimedCases: readonly SafeOpenProofCaseResult[],
	metrics: {
		readonly publicOpenPlanMedianMs: Readonly<Record<string, number>>
		readonly publicOpenPlanP95Ms: Readonly<Record<string, number>>
		readonly publicOpenPlanCv: Readonly<Record<string, number>>
		readonly publicFullOpenMedianMs: Readonly<Record<string, number>>
		readonly publicFullOpenP95Ms: Readonly<Record<string, number>>
		readonly publicFullOpenCv: Readonly<Record<string, number>>
		readonly publicFullOpenRatio: Readonly<Record<string, number>>
	},
): readonly string[] {
	const failures: string[] = []
	const profile = SAFE_OPEN_LATENCY_RUN_PROFILE
	if (result.repeat < profile.minimumRepeat) {
		failures.push(`repeat ${result.repeat} below profile minimum ${profile.minimumRepeat}`)
	}
	if (result.warmup < profile.minimumWarmup) {
		failures.push(`warmup ${result.warmup} below profile minimum ${profile.minimumWarmup}`)
	}
	if (!result.timingEnvironment) {
		failures.push('timing environment metadata missing')
	}
	const publicTimedNames = new Set(publicTimedCases.map((entry) => entry.name))
	for (const name of profile.requiredPublicCaseNames) {
		if (!publicTimedNames.has(name)) {
			failures.push(`required public case ${name} missing timed evidence`)
			continue
		}
		if (metrics.publicOpenPlanMedianMs[name] === undefined) {
			failures.push(`required metric openPlanMedianMs missing for ${name}`)
		}
		if (metrics.publicOpenPlanP95Ms[name] === undefined) {
			failures.push(`required metric openPlanP95Ms missing for ${name}`)
		}
		if (metrics.publicFullOpenRatio[name] === undefined) {
			failures.push(`required metric fullOpenRatio missing for ${name}`)
		}
		if (metrics.publicFullOpenMedianMs[name] === undefined) {
			failures.push(`required metric fullOpenMedianMs missing for ${name}`)
		}
		if (metrics.publicFullOpenP95Ms[name] === undefined) {
			failures.push(`required metric fullOpenP95Ms missing for ${name}`)
		}
		if (metrics.publicFullOpenCv[name] === undefined) {
			failures.push(`required metric fullOpenCv missing for ${name}`)
		}
		const cv = metrics.publicOpenPlanCv[name]
		if (cv === undefined) {
			failures.push(`required metric openPlanCv missing for ${name}`)
		} else if (cv > profile.cvGuard.maxRecommendedCv) {
			failures.push(
				`public open-plan CV ${cv} for ${name} exceeds owner-review guard ${profile.cvGuard.maxRecommendedCv}`,
			)
		}
	}
	return failures
}

function numericCaseMetricByName(
	cases: readonly SafeOpenProofCaseResult[],
	metric:
		| 'openPlanMedianMs'
		| 'openPlanP95Ms'
		| 'openPlanCv'
		| 'fullOpenMedianMs'
		| 'fullOpenP95Ms'
		| 'fullOpenCv'
		| 'fullOpenRatio',
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
		runProfile: cloneSafeOpenLatencyRunProfile(evidence.runProfile),
		runProfileFailures: [...evidence.runProfileFailures],
		...(evidence.timingEnvironment ? { timingEnvironment: { ...evidence.timingEnvironment } } : {}),
		publicTimedCaseNames: [...evidence.publicTimedCaseNames],
		generatedTimedCaseNames: [...evidence.generatedTimedCaseNames],
		publicOpenPlanMedianMs: { ...evidence.publicOpenPlanMedianMs },
		publicOpenPlanP95Ms: { ...evidence.publicOpenPlanP95Ms },
		publicOpenPlanCv: { ...evidence.publicOpenPlanCv },
		publicFullOpenMedianMs: { ...evidence.publicFullOpenMedianMs },
		publicFullOpenP95Ms: { ...evidence.publicFullOpenP95Ms },
		publicFullOpenCv: { ...evidence.publicFullOpenCv },
		publicFullOpenRatio: { ...evidence.publicFullOpenRatio },
		missingPolicyRequirements: [...evidence.missingPolicyRequirements],
	}
}

function cloneSafeOpenLatencyRunProfile(
	profile: ReleaseProofSafeOpenLatencyRunProfile,
): ReleaseProofSafeOpenLatencyRunProfile {
	return {
		...profile,
		requiredPublicCaseNames: [...profile.requiredPublicCaseNames],
		requiredMetrics: [...profile.requiredMetrics],
		cvGuard: { ...profile.cvGuard },
		forbiddenUses: [...profile.forbiddenUses],
		sourceReferences: profile.sourceReferences.map((reference) => ({ ...reference })),
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

function trustCompletenessBoundaryMarkdownRow(
	row: ReleaseProofTrustCompletenessBoundaryClass,
): string {
	return [row.name, row.promoteOnlyWhen, row.ownerAction]
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

function cloneTrustCompletenessBoundaryEvidence(
	evidence: ReleaseProofTrustCompletenessBoundaryEvidence,
): ReleaseProofTrustCompletenessBoundaryEvidence {
	return {
		...evidence,
		outOfScopeClasses: evidence.outOfScopeClasses.map((entry) => ({ ...entry })),
		doesNotCloseGates: [...evidence.doesNotCloseGates],
		sourceReferences: evidence.sourceReferences.map((entry) => ({ ...entry })),
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
	const releaseDecisionJson = process.argv.includes('--release-decision-json')
	const fixtureDecisionJson = process.argv.includes('--fixture-decision-json')
	const result = await runReleaseProofIndex({
		includeTimings: !process.argv.includes('--no-timings'),
	})
	console.log(
		releaseDecisionJson
			? JSON.stringify(result.releaseDecisionBoard, null, 2)
			: fixtureDecisionJson
				? JSON.stringify(releaseProofFixtureDecisionPacket(result), null, 2)
				: ownerHandoffsJson
					? JSON.stringify(releaseProofOwnerHandoffIndex(result), null, 2)
					: json
						? JSON.stringify(result, null, 2)
						: releaseProofIndexMarkdown(result),
	)
	if (!json && !ownerHandoffsJson && !releaseDecisionJson && !fixtureDecisionJson) {
		console.error(`Indexed ${result.artifactCount} release proof evidence artifacts.`)
		console.error(
			`Run with --json, --owner-handoffs-json, --release-decision-json, or --fixture-decision-json for machine-readable output from ${basename(import.meta.path)}.`,
		)
	}
}
