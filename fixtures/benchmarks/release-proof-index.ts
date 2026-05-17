import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, statSync } from 'node:fs'
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
const HIGH_RISK_PACKAGE_CONTRACT_COMMAND =
	'bun test fixtures/corpus/high-risk-package-contract.test.ts --timeout 30000'
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
	| 'local-evidence-wording-owner-gated'
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
	readonly researchHygieneDecisionPacket: ReleaseProofResearchHygieneDecisionPacket
	readonly nextOwnerActions: readonly ReleaseProofNextOwnerAction[]
	readonly claimBlockerBoard: readonly ReleaseProofClaimBlockerBoardRow[]
	readonly implementationHandoffs: readonly ReleaseProofImplementationHandoff[]
	readonly qssLeapfrogReleaseMatrix: ReleaseProofQssLeapfrogReleaseMatrix
	readonly releaseDecisionBoard: ReleaseProofReleaseDecisionBoard
	readonly claimDecisionCoverage: ReleaseProofClaimDecisionCoverage
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

export interface ReleaseProofCorrectnessBoundaryDecisionPacket {
	readonly ownerLoop: 'correctness'
	readonly status: 'owner-decision-required'
	readonly releaseGate: ReleaseProofReadinessSummary['releaseGate']
	readonly headlineClaimsAllowed: boolean
	readonly ownerApprovalRequired: true
	readonly artifact: 'package-action-proof'
	readonly gateId: 'unsupported-feature-boundary'
	readonly allCurrentEvidencePresent: boolean
	readonly missingFeatureNames: readonly string[]
	readonly approvalChecklist: readonly ReleaseProofCorrectnessPolicyApprovalItem[]
	readonly featureChecks: readonly ReleaseProofCorrectnessBoundaryFeatureCheck[]
	readonly validationCommands: readonly string[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly forbiddenShortcuts: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofPerformanceBoundaryDecisionPacket {
	readonly ownerLoop: 'performance'
	readonly status: 'owner-decision-required'
	readonly releaseGate: ReleaseProofReadinessSummary['releaseGate']
	readonly headlineClaimsAllowed: boolean
	readonly ownerApprovalRequired: true
	readonly broadSpeedClaimAllowed: false
	readonly benchmarkBlocker: {
		readonly artifactId: 'performance-claim-baseline-matrix'
		readonly path: 'docs/PERFORMANCE_CLAIM_BASELINE_MATRIX.md'
		readonly validationCommand: 'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts'
		readonly claim: 'broad XLSX read speed leadership'
		readonly releaseDecision: 'claim-downgrade-do-not-promote'
		readonly evidenceWeHave: readonly string[]
		readonly evidenceMissing: readonly string[]
		readonly qssContrast: readonly string[]
		readonly allowedWording: string
		readonly forbiddenWording: readonly string[]
		readonly nextAction: string
		readonly nextOwnerAction: string
		readonly benchmarkCommands: readonly string[]
		readonly acceptanceEvidence: readonly string[]
		readonly stopCondition: string
	}
	readonly approvalChecklist: readonly ReleaseProofPerformancePolicyApprovalItem[]
	readonly validationCommands: readonly string[]
	readonly sourceReferences: readonly ReleaseProofSourceReference[]
	readonly forbiddenShortcuts: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofResearchHygieneDecisionPacket {
	readonly ownerLoops: readonly ['product', 'release']
	readonly status: 'claim-downgrade-do-not-promote'
	readonly releaseGate: ReleaseProofReadinessSummary['releaseGate']
	readonly headlineClaimsAllowed: boolean
	readonly ownerApprovalRequired: true
	readonly claim: 'research-surface-hygiene'
	readonly workBlockDisposition: 'claim-downgrade-do-not-promote'
	readonly dirtyInventoryCommand: string
	readonly inventorySnapshot: ReleaseProofResearchHygieneInventorySnapshot
	readonly localExcelCorpus: ReleaseProofResearchHygieneLocalExcelCorpus
	readonly loopManagerState: ReleaseProofResearchHygieneLoopManagerState
	readonly releaseRoutingSummary: ReleaseProofResearchHygieneReleaseRoutingSummary
	readonly validationCommands: readonly string[]
	readonly ownerFiles: readonly string[]
	readonly classificationBuckets: readonly ReleaseProofResearchHygieneClassificationBucket[]
	readonly failureEvidence: readonly string[]
	readonly acceptanceCriteria: string
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly qssContrast: readonly string[]
	readonly nextOwnerAction: string
	readonly stopCondition: string
	readonly boundary: string
}

export interface ReleaseProofResearchHygieneReleaseRoutingSummary {
	readonly acceptedReleaseEvidence: readonly string[]
	readonly blockedClaims: readonly string[]
	readonly ownerReadyImplementationTasks: readonly string[]
	readonly archiveDeferMaterial: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofResearchHygieneClassificationBucket {
	readonly bucket: 'accepted-evidence' | 'active-owner-blocker' | 'archive-only'
	readonly requirement: string
	readonly forbiddenShortcut: string
}

export interface ReleaseProofResearchHygieneInventorySnapshot {
	readonly command: string
	readonly status: 'inventory-collected' | 'inventory-command-failed'
	readonly decision:
		| 'owner-classification-required'
		| 'current-inventory-classified-release-routing-required'
		| 'no-unclassified-paths-currently-visible'
		| 'inventory-rerun-required'
	readonly dirtyPathCount: number
	readonly classifiedPathCount: number
	readonly unclassifiedPathCount: number
	readonly statusCodeCounts: Readonly<Record<string, number>>
	readonly classificationCounts: Readonly<Record<string, number>>
	readonly rootCounts: Readonly<Record<'research' | 'scripts' | 'tmp' | 'other', number>>
	readonly untrackedDirectoryCount: number
	readonly modifiedFileCount: number
	readonly classifiedEntries: readonly ReleaseProofResearchHygieneInventoryEntry[]
	readonly unclassifiedEntries: readonly ReleaseProofResearchHygieneInventoryEntry[]
	readonly failure?: string
	readonly boundary: string
}

type ReleaseProofResearchHygieneClassification =
	| 'accepted-evidence'
	| 'active-owner-blocker'
	| 'archive-only'
	| 'unclassified-owner-decision-required'

export interface ReleaseProofResearchHygieneInventoryEntry {
	readonly statusCode: string
	readonly path: string
	readonly classification: ReleaseProofResearchHygieneClassification
	readonly reason: string
	readonly nextOwnerAction: string
}

export interface ReleaseProofResearchHygieneLocalExcelCorpus {
	readonly path: 'research/excel-corpus'
	readonly status:
		| 'local-corpus-inventory-present'
		| 'local-corpus-missing'
		| 'local-corpus-unreadable'
	readonly releaseDecision: 'active-owner-blocker-not-release-evidence'
	readonly fileCount: number
	readonly workbookCount: number
	readonly totalBytes: number
	readonly largestFile?: string
	readonly manifestPath: 'research/excel-corpus/manifest.json'
	readonly manifestEntryCount: number
	readonly manifestMissingFiles: readonly string[]
	readonly files: readonly ReleaseProofResearchHygieneLocalExcelCorpusFile[]
	readonly ownerAction: string
	readonly forbiddenWording: readonly string[]
	readonly failure?: string
	readonly boundary: string
}

export interface ReleaseProofResearchHygieneLocalExcelCorpusFile {
	readonly path: string
	readonly sizeBytes: number
	readonly manifestListed: boolean
	readonly workbookKind: 'xlsx' | 'xlsm' | 'metadata'
	readonly featureSignals: readonly string[]
	readonly releaseStatus: 'local-only-owner-review-required'
	readonly nextOwnerAction: string
}

export interface ReleaseProofResearchHygieneLoopManagerState {
	readonly paths: readonly ['scripts/ascend-loop-manager.ts', 'tmp/ascend-loop-manager']
	readonly status:
		| 'loop-manager-state-present'
		| 'loop-manager-state-missing'
		| 'loop-manager-state-unreadable'
	readonly releaseDecision: 'active-owner-blocker-not-release-evidence'
	readonly fileCount: number
	readonly totalBytes: number
	readonly files: readonly ReleaseProofResearchHygieneLoopManagerStateFile[]
	readonly ownerAction: string
	readonly forbiddenWording: readonly string[]
	readonly failure?: string
	readonly boundary: string
}

export interface ReleaseProofResearchHygieneLoopManagerStateFile {
	readonly path: string
	readonly sizeBytes: number
	readonly stateKind:
		| 'manager-script'
		| 'current-board'
		| 'manual-steers'
		| 'north-star-reflection'
		| 'manager-state-file'
	readonly releaseStatus: 'operational-state-owner-review-required'
	readonly nextOwnerAction: string
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
	readonly name: ReleaseProofPortfolioClaimName | ReleaseProofIndexExcludedEvidenceName
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
	readonly topClaimOwnerActionQueue: readonly ReleaseProofTopClaimOwnerAction[]
	readonly doNotPromoteYet: readonly ReleaseProofReleaseDecisionDoNotPromoteItem[]
	readonly doNotPromoteDispositionSummary: ReleaseProofReleaseDecisionDispositionSummary
	readonly releaseWordingDecisionSummary: ReleaseProofReleaseWordingDecisionSummary
	readonly todayCommitClaimMatrix: readonly ReleaseProofTodayCommitClaimMatrixRow[]
	readonly claimDecisionContractCoverage: ReleaseProofClaimDecisionContractCoverage
	readonly blockedOwnerActionQueue: readonly ReleaseProofBlockedOwnerAction[]
	readonly benchmarkCorpusOwnerActionQueue: readonly ReleaseProofBenchmarkCorpusOwnerAction[]
	readonly implementationReadyOwnerActionQueue: readonly ReleaseProofImplementationReadyOwnerAction[]
	readonly claimDowngradeOwnerActionQueue: readonly ReleaseProofClaimDowngradeOwnerAction[]
	readonly benchmarkCorpusRunContractCoverage: ReleaseProofBenchmarkCorpusRunContractCoverage
	readonly ownerActionQueueCoverage: ReleaseProofOwnerActionQueueCoverage
	readonly ownerActionExecutionContractCoverage: ReleaseProofOwnerActionExecutionContractCoverage
	readonly boundary: string
}

export interface ReleaseProofTodayCommitClaimMatrixRow {
	readonly claimArea:
		| 'safe-agent-workflows'
		| 'formula-calc-behavior'
		| 'signed-encrypted-macro-handling'
		| 'write-performance'
		| 'external-baselines'
		| 'research-proof-surface'
	readonly commits: readonly string[]
	readonly releaseOrSotaClaimBecameMoreTrue: string
	readonly evidenceProvesIt: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly nextOwnerAction: string
	readonly boundary: string
}

export interface ReleaseProofReleaseDecisionDispositionSummary {
	readonly implementationReadyBlockerNames: readonly ReleaseProofReleaseDecisionDoNotPromoteItem['name'][]
	readonly benchmarkCorpusBlockerNames: readonly ReleaseProofReleaseDecisionDoNotPromoteItem['name'][]
	readonly claimDowngradeDoNotPromoteNames: readonly ReleaseProofReleaseDecisionDoNotPromoteItem['name'][]
	readonly boundary: string
}

export interface ReleaseProofReleaseWordingDecisionSummary {
	readonly status: 'headline-claims-blocked-local-wording-only'
	readonly headlineClaimsAllowed: false
	readonly localAllowedClaimNames: readonly ReleaseProofIndexArtifactName[]
	readonly doNotPromoteClaimNames: readonly ReleaseProofReleaseDecisionDoNotPromoteItem['name'][]
	readonly localAllowedWordingByClaim: Readonly<Record<ReleaseProofIndexArtifactName, string>>
	readonly doNotPromoteAllowedWordingByClaim: Readonly<
		Record<ReleaseProofReleaseDecisionDoNotPromoteItem['name'], string>
	>
	readonly forbiddenWordingByClaim: Readonly<
		Record<
			ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name'],
			readonly string[]
		>
	>
	readonly boundary: string
}

export interface ReleaseProofClaimDecisionContractCoverage {
	readonly status: 'all-release-claim-decisions-self-contained' | 'claim-decision-contract-gap'
	readonly decisionCount: number
	readonly topClaimDecisionCount: number
	readonly doNotPromoteDecisionCount: number
	readonly missingEvidenceWeHaveKeys: readonly string[]
	readonly missingEvidenceMissingKeys: readonly string[]
	readonly missingQssContrastKeys: readonly string[]
	readonly missingAllowedWordingKeys: readonly string[]
	readonly missingForbiddenWordingKeys: readonly string[]
	readonly missingNextOwnerActionKeys: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofClaimDecisionCoverage {
	readonly status: 'all-handoff-claims-covered-by-release-decision-board' | 'coverage-gap'
	readonly topClaimNames: readonly ReleaseProofIndexArtifactName[]
	readonly doNotPromoteNames: readonly ReleaseProofReleaseDecisionDoNotPromoteItem['name'][]
	readonly portfolioClaimCount: number
	readonly deferredClaimCount: number
	readonly excludedEvidenceCount: number
	readonly uncoveredPortfolioClaimNames: readonly ReleaseProofPortfolioClaimName[]
	readonly uncoveredDeferredClaimNames: readonly ReleaseProofDeferredClaimName[]
	readonly uncoveredExcludedEvidenceNames: readonly ReleaseProofIndexExcludedEvidenceName[]
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
	readonly workBlockDisposition:
		| 'implementation-ready-blocker'
		| 'benchmark-corpus-blocker'
		| 'claim-downgrade-do-not-promote'
	readonly ownerLoops: readonly ReleaseProofReadinessOwner[]
	readonly reason: string
	readonly evidenceWeHave: readonly string[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerAction: string
	readonly validationCommands: readonly string[]
	readonly killCriterion: string
	readonly boundary: string
}

export interface ReleaseProofTopClaimOwnerAction {
	readonly artifact: ReleaseProofIndexArtifactName
	readonly claim: string
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly requirementId: string
	readonly workBlockDisposition: ReleaseProofReleaseDecisionDoNotPromoteItem['workBlockDisposition']
	readonly rank: number
	readonly priority: ReleaseProofNextOwnerAction['priority']
	readonly nextStepKind: ReleaseProofNextOwnerAction['nextStepKind']
	readonly evidenceWeHave: readonly ReleaseProofQssAcceptedEvidenceItem[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly validationCommand: string
	readonly acceptanceEvidence: string
	readonly forbiddenShortcut: string
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerAction: string
	readonly boundary: string
}

export interface ReleaseProofBlockedOwnerAction {
	readonly name: ReleaseProofReleaseDecisionDoNotPromoteItem['name']
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly workBlockDisposition: ReleaseProofReleaseDecisionDoNotPromoteItem['workBlockDisposition']
	readonly evidenceWeHave: readonly string[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerAction: string
	readonly validationCommands: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofBenchmarkCorpusOwnerAction {
	readonly sourceQueue: 'top-claim-owner-action' | 'blocked-owner-action'
	readonly name: ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name']
	readonly claim: string
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly requirementId?: string
	readonly workBlockDisposition: 'benchmark-corpus-blocker'
	readonly ownerFiles: readonly string[]
	readonly validationCommands: readonly string[]
	readonly commandsToRun: readonly string[]
	readonly failureEvidence: readonly string[]
	readonly acceptanceCriteria: string
	readonly runInputScope: string
	readonly runEnvironment: string
	readonly requiredOutputEvidence: readonly string[]
	readonly promotionCondition: string
	readonly stopCondition: string
	readonly evidenceWeHave: readonly string[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerAction: string
	readonly boundary: string
}

export interface ReleaseProofImplementationReadyOwnerAction {
	readonly sourceQueue: 'top-claim-owner-action' | 'blocked-owner-action'
	readonly name: ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name']
	readonly claim: string
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly requirementId?: string
	readonly workBlockDisposition: 'implementation-ready-blocker'
	readonly ownerFiles: readonly string[]
	readonly validationCommands: readonly string[]
	readonly commandsToRun: readonly string[]
	readonly failureEvidence: readonly string[]
	readonly acceptanceCriteria: string
	readonly evidenceWeHave: readonly string[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerAction: string
	readonly boundary: string
}

export interface ReleaseProofClaimDowngradeOwnerAction {
	readonly sourceQueue: 'blocked-owner-action'
	readonly name: ReleaseProofReleaseDecisionDoNotPromoteItem['name']
	readonly claim: string
	readonly ownerLoop: ReleaseProofReadinessOwner
	readonly workBlockDisposition: 'claim-downgrade-do-not-promote'
	readonly ownerFiles: readonly string[]
	readonly validationCommands: readonly string[]
	readonly commandsToRun: readonly string[]
	readonly failureEvidence: readonly string[]
	readonly acceptanceCriteria: string
	readonly evidenceWeHave: readonly string[]
	readonly evidenceMissing: readonly string[]
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWording: readonly string[]
	readonly nextOwnerAction: string
	readonly boundary: string
}

export interface ReleaseProofBenchmarkCorpusRunContractCoverage {
	readonly status:
		| 'all-benchmark-corpus-actions-have-run-contract'
		| 'benchmark-corpus-run-contract-gap'
	readonly actionCount: number
	readonly missingInputScopeActionKeys: readonly string[]
	readonly missingRunEnvironmentActionKeys: readonly string[]
	readonly missingRequiredOutputEvidenceActionKeys: readonly string[]
	readonly missingPromotionConditionActionKeys: readonly string[]
	readonly missingStopConditionActionKeys: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofOwnerActionQueueCoverage {
	readonly status: 'all-owner-actions-covered-by-disposition-queues' | 'coverage-gap'
	readonly sourceTopClaimActionCount: number
	readonly sourceBlockedActionCount: number
	readonly benchmarkCorpusActionCount: number
	readonly implementationReadyActionCount: number
	readonly claimDowngradeActionCount: number
	readonly coveredActionCount: number
	readonly uncoveredTopClaimActionKeys: readonly string[]
	readonly uncoveredBlockedActionKeys: readonly string[]
	readonly boundary: string
}

export interface ReleaseProofOwnerActionExecutionContractCoverage {
	readonly status:
		| 'all-disposition-owner-actions-have-execution-contract'
		| 'execution-contract-gap'
	readonly actionCount: number
	readonly benchmarkCorpusActionCount: number
	readonly implementationReadyActionCount: number
	readonly claimDowngradeActionCount: number
	readonly missingOwnerFileActionKeys: readonly string[]
	readonly missingCommandActionKeys: readonly string[]
	readonly missingFailureEvidenceActionKeys: readonly string[]
	readonly missingAcceptanceCriteriaActionKeys: readonly string[]
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
	readonly rejectedFixtures: readonly string[]
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
		`Release wording summary: status=${result.releaseDecisionBoard.releaseWordingDecisionSummary.status}; localAllowed=${result.releaseDecisionBoard.releaseWordingDecisionSummary.localAllowedClaimNames.join(',')}; doNotPromote=${result.releaseDecisionBoard.releaseWordingDecisionSummary.doNotPromoteClaimNames.join(',')}`,
		`Claim decision contract coverage: status=${result.releaseDecisionBoard.claimDecisionContractCoverage.status}; decisions=${result.releaseDecisionBoard.claimDecisionContractCoverage.decisionCount}; missingEvidenceHave=${result.releaseDecisionBoard.claimDecisionContractCoverage.missingEvidenceWeHaveKeys.join(',')}; missingEvidenceMissing=${result.releaseDecisionBoard.claimDecisionContractCoverage.missingEvidenceMissingKeys.join(',')}; missingQss=${result.releaseDecisionBoard.claimDecisionContractCoverage.missingQssContrastKeys.join(',')}; missingAllowed=${result.releaseDecisionBoard.claimDecisionContractCoverage.missingAllowedWordingKeys.join(',')}; missingForbidden=${result.releaseDecisionBoard.claimDecisionContractCoverage.missingForbiddenWordingKeys.join(',')}; missingNext=${result.releaseDecisionBoard.claimDecisionContractCoverage.missingNextOwnerActionKeys.join(',')}`,
		'',
		'| Rank | Claim | Evidence we have | Evidence missing | QSS contrast | Allowed wording | Forbidden wording | Next owner action | Owner decision artifact | Headline claim allowed | Implementation promotion allowed | Exact proof | Must not claim | A+ blocking owner action | Boundary |',
		'| ---: | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |',
		...result.releaseDecisionBoard.rows.map(releaseDecisionBoardMarkdownRow),
		'Top claim owner action queue:',
		...result.releaseDecisionBoard.topClaimOwnerActionQueue.map(
			(row) =>
				`- ${row.ownerLoop}/${row.artifact}/${row.requirementId}: ${row.nextStepKind}. Command: \`${row.validationCommand}\`. Accept: ${row.acceptanceEvidence} Forbidden: ${row.forbiddenShortcut}`,
		),
		'Benchmark/corpus owner action queue:',
		...result.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue.map(
			(row) =>
				`- ${row.ownerLoop}/${row.name}${row.requirementId ? `/${row.requirementId}` : ''}: ${row.sourceQueue}. Files: ${row.ownerFiles.map((file) => `\`${file}\``).join('; ')} Commands: ${row.commandsToRun.map((command) => `\`${command}\``).join('; ')} Input: ${row.runInputScope} Environment: ${row.runEnvironment} Required output: ${row.requiredOutputEvidence.join('; ')} Promote only if: ${row.promotionCondition} Stop: ${row.stopCondition} Failure evidence: ${row.failureEvidence.join('; ')} Accept: ${row.acceptanceCriteria} Next: ${row.nextOwnerAction}`,
		),
		'Implementation-ready owner action queue:',
		...result.releaseDecisionBoard.implementationReadyOwnerActionQueue.map(
			(row) =>
				`- ${row.ownerLoop}/${row.name}${row.requirementId ? `/${row.requirementId}` : ''}: ${row.sourceQueue}. Files: ${row.ownerFiles.map((file) => `\`${file}\``).join('; ')} Commands: ${row.commandsToRun.map((command) => `\`${command}\``).join('; ')} Failure evidence: ${row.failureEvidence.join('; ')} Accept: ${row.acceptanceCriteria} Next: ${row.nextOwnerAction}`,
		),
		'Claim downgrade owner action queue:',
		...result.releaseDecisionBoard.claimDowngradeOwnerActionQueue.map(
			(row) =>
				`- ${row.ownerLoop}/${row.name}: ${row.sourceQueue}. Files: ${row.ownerFiles.map((file) => `\`${file}\``).join('; ')} Commands: ${row.commandsToRun.map((command) => `\`${command}\``).join('; ')} Failure evidence: ${row.failureEvidence.join('; ')} Accept: ${row.acceptanceCriteria} Next: ${row.nextOwnerAction}`,
		),
		`Benchmark/corpus run contract coverage: status=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.status}; actions=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.actionCount}; missingInputScope=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.missingInputScopeActionKeys.join(',')}; missingRunEnvironment=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.missingRunEnvironmentActionKeys.join(',')}; missingRequiredOutput=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.missingRequiredOutputEvidenceActionKeys.join(',')}; missingPromotion=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.missingPromotionConditionActionKeys.join(',')}; missingStop=${result.releaseDecisionBoard.benchmarkCorpusRunContractCoverage.missingStopConditionActionKeys.join(',')}`,
		`Owner action queue coverage: status=${result.releaseDecisionBoard.ownerActionQueueCoverage.status}; top=${result.releaseDecisionBoard.ownerActionQueueCoverage.sourceTopClaimActionCount}; blocked=${result.releaseDecisionBoard.ownerActionQueueCoverage.sourceBlockedActionCount}; covered=${result.releaseDecisionBoard.ownerActionQueueCoverage.coveredActionCount}; uncoveredTop=${result.releaseDecisionBoard.ownerActionQueueCoverage.uncoveredTopClaimActionKeys.join(',')}; uncoveredBlocked=${result.releaseDecisionBoard.ownerActionQueueCoverage.uncoveredBlockedActionKeys.join(',')}`,
		`Owner action execution contract coverage: status=${result.releaseDecisionBoard.ownerActionExecutionContractCoverage.status}; actions=${result.releaseDecisionBoard.ownerActionExecutionContractCoverage.actionCount}; missingFiles=${result.releaseDecisionBoard.ownerActionExecutionContractCoverage.missingOwnerFileActionKeys.join(',')}; missingCommands=${result.releaseDecisionBoard.ownerActionExecutionContractCoverage.missingCommandActionKeys.join(',')}; missingFailureEvidence=${result.releaseDecisionBoard.ownerActionExecutionContractCoverage.missingFailureEvidenceActionKeys.join(',')}; missingAcceptanceCriteria=${result.releaseDecisionBoard.ownerActionExecutionContractCoverage.missingAcceptanceCriteriaActionKeys.join(',')}`,
		'',
		'Do not promote yet:',
		`Disposition summary: implementation-ready-blocker=${result.releaseDecisionBoard.doNotPromoteDispositionSummary.implementationReadyBlockerNames.join(',')}; benchmark-corpus-blocker=${result.releaseDecisionBoard.doNotPromoteDispositionSummary.benchmarkCorpusBlockerNames.join(',')}; claim-downgrade-do-not-promote=${result.releaseDecisionBoard.doNotPromoteDispositionSummary.claimDowngradeDoNotPromoteNames.join(',')}`,
		...result.releaseDecisionBoard.doNotPromoteYet.map(
			(item) =>
				`- ${item.name}: Disposition: ${item.workBlockDisposition}. ${item.reason} Missing: ${item.evidenceMissing.join('; ')} Allowed: ${item.allowedWording} Forbidden: ${item.forbiddenWording.join('; ')} Commands: ${item.validationCommands.map((command) => `\`${command}\``).join('; ')} Next: ${item.nextOwnerAction}`,
		),
		'Blocked owner action queue:',
		...result.releaseDecisionBoard.blockedOwnerActionQueue.map(
			(row) =>
				`- ${row.ownerLoop}/${row.name}: ${row.workBlockDisposition}. Commands: ${row.validationCommands.map((command) => `\`${command}\``).join('; ')} Next: ${row.nextOwnerAction}`,
		),
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
		researchHygieneDecisionPacket: releaseProofResearchHygieneDecisionPacket(result),
		streamingMatrixEvidence: cloneStreamingMatrixEvidence(result.streamingMatrixEvidence),
		nextOwnerActions: result.readiness.nextOwnerActions,
		claimBlockerBoard: result.readiness.claimBlockerBoard,
		implementationHandoffs: result.readiness.implementationHandoffs,
		qssLeapfrogReleaseMatrix: cloneQssLeapfrogReleaseMatrix(result.qssLeapfrogReleaseMatrix),
		releaseDecisionBoard: cloneReleaseDecisionBoard(result.releaseDecisionBoard),
		claimDecisionCoverage: claimDecisionCoverage(result),
		claimPortfolio: result.claimPortfolio.map(clonePortfolioClaim),
		deferredClaims: result.deferredClaims,
		excludedEvidence: result.excludedEvidence,
		boundary:
			'Compact owner handoff index for release proof routing. It is not a release artifact bundle, signed attestation, or product surface.',
	}
}

function claimDecisionCoverage(result: ReleaseProofIndexResult): ReleaseProofClaimDecisionCoverage {
	const coveredNames = new Set<string>([
		...result.releaseDecisionBoard.rows.map((row) => row.artifact),
		...result.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name),
	])
	const uncoveredPortfolioClaimNames = result.claimPortfolio
		.map((claim) => claim.name)
		.filter((name) => !coveredNames.has(name))
	const uncoveredDeferredClaimNames = result.deferredClaims
		.map((claim) => claim.name)
		.filter((name) => !coveredNames.has(name))
	const uncoveredExcludedEvidenceNames = result.excludedEvidence
		.map((evidence) => evidence.name)
		.filter((name) => !coveredNames.has(name))
	const uncoveredCount =
		uncoveredPortfolioClaimNames.length +
		uncoveredDeferredClaimNames.length +
		uncoveredExcludedEvidenceNames.length
	return {
		status:
			uncoveredCount === 0
				? 'all-handoff-claims-covered-by-release-decision-board'
				: 'coverage-gap',
		topClaimNames: result.releaseDecisionBoard.rows.map((row) => row.artifact),
		doNotPromoteNames: result.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name),
		portfolioClaimCount: result.claimPortfolio.length,
		deferredClaimCount: result.deferredClaims.length,
		excludedEvidenceCount: result.excludedEvidence.length,
		uncoveredPortfolioClaimNames,
		uncoveredDeferredClaimNames,
		uncoveredExcludedEvidenceNames,
		boundary:
			'Coverage summary only. Every owner-handoff claim, deferred claim, and excluded diagnostic must resolve to a top release-decision row or a do-not-promote decision before it can influence release wording.',
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

export function releaseProofCorrectnessBoundaryDecisionPacket(
	result: ReleaseProofIndexResult,
): ReleaseProofCorrectnessBoundaryDecisionPacket {
	const approvalChecklist = result.correctnessPolicy.approvalChecklist.map((item) => ({ ...item }))
	const featureChecks = result.correctnessBoundaryEvidence.featureChecks.map((entry) => ({
		...entry,
		evidenceSources: [...entry.evidenceSources],
		proofChecks: [...entry.proofChecks],
	}))
	return {
		ownerLoop: 'correctness',
		status: 'owner-decision-required',
		releaseGate: result.readiness.releaseGate,
		headlineClaimsAllowed: result.readiness.headlineClaimsAllowed,
		ownerApprovalRequired: true,
		artifact: result.correctnessBoundaryEvidence.artifact,
		gateId: result.correctnessBoundaryEvidence.gateId,
		allCurrentEvidencePresent: result.correctnessBoundaryEvidence.allCurrentEvidencePresent,
		missingFeatureNames: [...result.correctnessBoundaryEvidence.missingFeatureNames],
		approvalChecklist,
		featureChecks,
		validationCommands: [
			...new Set([
				result.correctnessBoundaryEvidence.validationCommand,
				HIGH_RISK_PACKAGE_CONTRACT_COMMAND,
				...approvalChecklist.map((item) => item.validationCommand),
			]),
		],
		sourceReferences: result.correctnessPolicy.sourceReferences.map((entry) => ({ ...entry })),
		forbiddenShortcuts: [
			'Do not claim semantic support for every unsupported workbook feature.',
			'Do not claim chart XML byte passthrough, signature preservation or verification, Excel-fresh cached formulas, macro/ActiveX safety, unknown-part understanding, or full streaming parity.',
			'Do not treat evidence-present status as owner approval or release wording approval.',
		],
		boundary:
			'Compact correctness unsupported-feature boundary packet. It is not owner approval, semantic support for unsupported features, release wording approval, or permission to promote package-action parity claims.',
	}
}

export function releaseProofPerformanceBoundaryDecisionPacket(
	result: ReleaseProofIndexResult,
): ReleaseProofPerformanceBoundaryDecisionPacket {
	const approvalChecklist = result.performancePolicy.approvalChecklist.map((item) => ({ ...item }))
	const benchmarkValidationCommand =
		'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts'
	const benchmarkCommands = [
		'env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all.json',
		'env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all-scoreboard.json',
		'env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata-scoreboard.json',
	]
	return {
		ownerLoop: 'performance',
		status: 'owner-decision-required',
		releaseGate: result.readiness.releaseGate,
		headlineClaimsAllowed: result.readiness.headlineClaimsAllowed,
		ownerApprovalRequired: true,
		broadSpeedClaimAllowed: false,
		benchmarkBlocker: {
			artifactId: 'performance-claim-baseline-matrix',
			path: 'docs/PERFORMANCE_CLAIM_BASELINE_MATRIX.md',
			validationCommand: benchmarkValidationCommand,
			claim: 'broad XLSX read speed leadership',
			releaseDecision: 'claim-downgrade-do-not-promote',
			evidenceWeHave: [
				'Focused ClosedXML value-read head-to-head evidence is accepted for comparable value-read rows only.',
				'Same-lane selected-sheet external-process evidence at commit 39163862 is accepted for Ascend, SheetJS, and OpenPyXL Data-sheet-only rows.',
				'Current-worktree python-calamine selected-sheet runner proof at commit 79d6cefd is accepted as evidence that Calamine can project the Data sheet on the same selected-sheet lane, but not as clean repeat-5 speed evidence.',
				'Commit bbf875b4 adds an Ascend readXlsx selected-sheet external-process row to the same selected-sheet timing lane, making it benchmarkable without using a separate in-process timing model.',
				'Commit df349629 records a current same-lane selected-sheet repeat-15 baseline: readXlsx selected-sheet is faster than SheetJS, openpyxl, and python-calamine on the completed comparable rows, while SDK Ascend.open remains a timing-boundary row and unsupported competitors remain non-wins.',
				'Same-lane metadata-only external-process evidence at commit fa3a13dc is accepted for Ascend, SheetJS, and OpenPyXL metadata-only rows.',
				'Commit b6925afe adds a python-calamine metadata-only baseline to the same metadata-only timing lane and scoreboard coverage; the synthetic scoreboard fixture shows Calamine can be the metadata-only row winner, so this is downgrade evidence rather than an Ascend speed win.',
				'Commit bb31bebe pins current FastXLSX carry-forward evidence: FastXLSX runs in an isolated Python 3.12 environment on the same cell-materialization lane, Ascend wins comparable value/warm rows by median, but FastXLSX uses less memory and one noisy table-heavy repeat-5 group briefly favored FastXLSX.',
				'Commit f8846cf8 records a current clean metadata-only Calamine recheck: Calamine remains the plain sheet-list metadata-only median winner, a capsule-skip optimization target failed to improve Ascend and was reverted, and no production optimization is carried forward from that target.',
				'Commit 1908f3f5 records the current metadata-only Calamine recheck at cc689bcc: python-calamine-metadata-only wins the comparable plain sheet-list/no-cell-hydration row, a file-backed metadata-only path-open candidate was measured and killed, and metadata-only speed wording remains downgraded.',
				'Commit 187548bf pins dense-values write evidence: a noisy full external repeat-5 row lost to native writers, a focused repeat-15 fastest-writer rerun had Ascend as median winner, ClosedXML was unavailable, and no write optimization is justified from this row.',
				'Commits 67b900ed, e22eb86a, 0d0c9632, 905ecb5e, c297ba4c, 27af69d4, 7d61a2ef, 183e7ebf, and eca32509 pin plain-text, string-heavy, dense-values, styles-heavy, formula-heavy, table-heavy, and feature-rich generated-write rows as scoped comparable evidence, not broad XLSX write/SOTA proof.',
				'Current full-profile and merged selected-sheet/metadata-only scoreboards from commit 9ddfff91 report no leader failures or profile leader failures.',
			],
			evidenceMissing: [
				'ClosedXML coverage policy for missing/error rows outside comparable value-read.',
				'Feature-rich SheetJS and Calamine semantic-support evidence or an explicit not-comparable policy.',
				'Unsupported selected-sheet policy for ExcelJS, Apache POI, and ClosedXML rows, plus unsupported metadata-only competitor policy for ExcelJS, Apache POI, and ClosedXML rows.',
				'Clean repeat-5 selected-sheet rerun that includes the new ascend-readXlsx selected-sheet row before any readXlsx-specific speed wording.',
				'Same-timing SDK Ascend.open selected-sheet open-only row or benchmark bug classification before any SDK selected-sheet speed wording.',
				'Current full-profile gate that carries the isolated FastXLSX setup forward while keeping feature-rich and memory-footprint boundaries explicit.',
				'Named metadata-only production cost center from profiling before any further metadata-only optimization work.',
				'Clean multi-workload XLSX write SOTA gate before any broad write-speed or QSS/SOTA write wording.',
				'Tracked-clean release-environment approval before any public speed wording.',
			],
			qssContrast: [
				'QSS-leapfrog speed wording is blocked because current evidence is scoped to completed comparable rows, not all broad XLSX read workloads.',
				'Competitor rows with unsupported operations, semantic mismatches, or unavailable runners are non-wins and cannot be counted against QSS or SOTA claims.',
			],
			allowedWording:
				'Allowed wording: Ascend has bounded local evidence of comparable selected-sheet and metadata-only read rows, while broad XLSX read-speed leadership remains blocked and Calamine metadata-only coverage must not be counted as an Ascend win.',
			forbiddenWording: [
				'Do not claim Ascend is the fastest XLSX reader.',
				'Do not claim Ascend leads metadata-only reads while the Calamine metadata-only baseline is comparable and may win the row.',
				'Do not claim Ascend beats Calamine on metadata-only open or that a capsule-skip optimization improved the current measured workflow.',
				'Do not claim Ascend beats FastXLSX on memory, feature-rich rich-metadata reads, or every XLSX workflow.',
				'Do not claim Ascend is SOTA for XLSX write, beats every generated XLSX writer, beats omitted/unsupported/blocked writers, produces the smallest XLSX, or proves byte/order-equivalent output against every writer.',
				'Do not claim SOTA, QSS-leapfrog read speed, or broad speed leadership from the current partial baseline.',
				'Do not count unavailable runners, unsupported operations, feature-rich semantic mismatches, or dirty-worktree timings as wins.',
			],
			nextAction:
				'Downgrade broad read/write speed wording and stop production optimization from this evidence: the current full-profile and merged selected-sheet/metadata-only scoreboards have no leader failures, the Calamine metadata-only baseline is not an Ascend win, commit df349629 keeps selected-sheet wording scoped to readXlsx same-lane rows and blocks SDK Ascend.open speed wording on a timing boundary, commits f8846cf8 and 1908f3f5 kill metadata-only optimization targets after measured Calamine non-wins, generated-write wins remain scoped to plain-text/string-heavy/dense-values/styles-heavy/formula-heavy/table-heavy/feature-rich comparable rows with OpenPyXL feature-rich kept as not comparable, current FastXLSX value/warm rows are scoped wins with memory and feature-rich boundaries, and ClosedXML coverage, feature-rich SheetJS/Calamine semantic mismatches, remaining unsupported selected-sheet/metadata-only competitor rows, and clean multi-workload xlsx-write-sota coverage remain non-wins.',
			nextOwnerAction:
				'Benchmarking owner either resolves one explicit blocker row (ClosedXML coverage policy, feature-rich SheetJS/Calamine semantic policy, remaining unsupported selected-sheet/metadata-only competitor policy, current full-profile FastXLSX carry-forward policy, same-timing SDK selected-sheet open-only row, clean multi-workload xlsx-write-sota coverage, or a profiling-named public workflow cost center) with the commands below, or keeps the broad speed claim downgraded and stops.',
			benchmarkCommands,
			acceptanceEvidence: [
				'Clean detached worktree or clean release benchmark environment.',
				'ClosedXML is measured as ran/won for comparable value-read rows in the focused head-to-head run and remains not comparable for selected-sheet and metadata-only.',
				'The selected-sheet same-lane external-process run at commit 39163862 is accepted as scoped evidence: Ascend, SheetJS, and OpenPyXL all loaded only the Data sheet; Ascend had the fastest median among those completed comparable rows.',
				'Commit df349629 records current selected-sheet same-lane evidence: readXlsx selected-sheet wins completed comparable rows against SheetJS, openpyxl, and python-calamine, but SDK Ascend.open selected-sheet remains a timing-boundary row and unsupported selected-sheet competitors stay non-wins.',
				'Selected-sheet wording remains scoped because ExcelJS, Apache POI, and ClosedXML are unsupported-operation gaps, while python-calamine selected-sheet has only current-worktree projection proof and must not be counted as a clean speed win.',
				'The metadata-only same-lane external-process run at commit fa3a13dc is accepted as scoped evidence: Ascend, SheetJS, and OpenPyXL all loaded workbook metadata without hydrating cells; Ascend had the fastest median among those completed comparable rows.',
				'Commit b6925afe adds python-calamine metadata-only runner and scoreboard coverage on the same metadata-only lane; treat the row as comparable coverage and downgrade broad speed wording because Calamine may win that row.',
				'Commit bb31bebe pins current FastXLSX evidence: same-lane value/warm rows are comparable and mostly Ascend median wins, but FastXLSX lower RSS and the feature-rich boundary keep the evidence scoped.',
				'Commit f8846cf8 records the current metadata-only Calamine recheck and rejected capsule-skip optimization: Calamine wins the plain sheet-list metadata-only median, patched Ascend did not improve median or memory, and the patch was reverted.',
				'Commit 1908f3f5 records a newer metadata-only Calamine loss: Calamine wins the current comparable plain sheet-list row at cc689bcc, a file-backed metadata-only path-open candidate was measured and rejected, and production work should stop until profiling identifies a narrower safe-open cost center.',
				'Commit 187548bf records scoped dense-values write evidence and defer wording: the repeat-15 fastest-writer rerun is an Ascend median win, but the noisy full row, unavailable ClosedXML runner, file-size and memory boundaries block broad write-speed wording.',
				'Commits 67b900ed, e22eb86a, 0d0c9632, 905ecb5e, c297ba4c, 27af69d4, 7d61a2ef, 183e7ebf, and eca32509 record scoped generated-write evidence: plain-text, string-heavy, dense-values, styles-heavy, formula-heavy, table-heavy, and feature-rich rows can be cited only as bounded comparable rows with memory/file-size/order-equivalence and unsupported-runner/not-comparable boundaries.',
				'The current full-profile run at commit 9ddfff91 and merged selected-sheet/metadata-only scoreboard report no leader failures or profile leader failures, but coverage still fails for ClosedXML missing/error rows and feature-rich semantic mismatches.',
				'Median, p95, CV/noise, memory, environment, runner/library versions, command, input shape, and semantic comparability are recorded for each comparable row.',
				'Failed, missing, or semantically mismatched runners are not counted as wins.',
			],
			stopCondition:
				'Stop production optimization from this evidence. Continue only as blocker work for ClosedXML coverage policy, feature-rich SheetJS/Calamine semantic support or not-comparable policy, remaining unsupported selected-sheet/metadata-only competitor policy, current full-profile FastXLSX carry-forward policy, same-timing SDK selected-sheet open-only row, clean multi-workload xlsx-write-sota coverage, or a profiling-named public workflow cost center.',
		},
		approvalChecklist,
		validationCommands: [
			...new Set([
				benchmarkValidationCommand,
				...benchmarkCommands,
				...approvalChecklist.map((item) => item.validationCommand),
			]),
		],
		sourceReferences: result.performancePolicy.sourceReferences.map((entry) => ({ ...entry })),
		forbiddenShortcuts: [
			'Do not claim Ascend is the fastest XLSX reader.',
			'Do not claim SOTA or QSS-leapfrog read speed from partial profile rows.',
			'Do not count unavailable runners, blocked runners, unsupported operations, semantic mismatches, dirty-worktree timings, or one-workload medians as wins.',
			'Do not turn local safe-open timing into a release threshold, SLA, or QSS performance comparison.',
		],
		boundary:
			'Compact performance boundary packet. It is not a release speed claim, benchmark promotion, production optimization mandate, or owner approval.',
	}
}

export function releaseProofResearchHygieneDecisionPacket(
	result: ReleaseProofIndexResult,
): ReleaseProofResearchHygieneDecisionPacket {
	const decision = result.releaseDecisionBoard.doNotPromoteYet.find(
		(item) => item.name === 'research-surface-hygiene',
	)
	const downgradeRows = result.releaseDecisionBoard.claimDowngradeOwnerActionQueue.filter(
		(row) => row.name === 'research-surface-hygiene',
	)
	const ownerFiles = [...new Set(downgradeRows.flatMap((row) => row.ownerFiles))]
	const validationCommands = [...new Set(downgradeRows.flatMap((row) => row.commandsToRun))]
	const dirtyInventoryCommand =
		'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager'
	const inventorySnapshot = researchHygieneInventorySnapshot(dirtyInventoryCommand)
	return {
		ownerLoops: ['product', 'release'],
		status: 'claim-downgrade-do-not-promote',
		releaseGate: result.readiness.releaseGate,
		headlineClaimsAllowed: result.readiness.headlineClaimsAllowed,
		ownerApprovalRequired: true,
		claim: 'research-surface-hygiene',
		workBlockDisposition: 'claim-downgrade-do-not-promote',
		dirtyInventoryCommand,
		inventorySnapshot,
		localExcelCorpus: researchHygieneLocalExcelCorpus(),
		loopManagerState: researchHygieneLoopManagerState(),
		releaseRoutingSummary: researchHygieneReleaseRoutingSummary(result, inventorySnapshot),
		validationCommands,
		ownerFiles,
		classificationBuckets: [
			{
				bucket: 'accepted-evidence',
				requirement:
					'Only cite a research path after it is represented in the release-proof index with evidence we have, evidence missing, QSS contrast, allowed wording, forbidden wording, and next owner action.',
				forbiddenShortcut:
					'Do not cite raw research, tmp, or experiment files directly in release wording.',
			},
			{
				bucket: 'active-owner-blocker',
				requirement:
					'Convert unfinished research into an owner action with files, commands, failure evidence, acceptance criteria, and a stop condition.',
				forbiddenShortcut:
					'Do not let broad research notes create new product surfaces or benchmark work without a named owner loop.',
			},
			{
				bucket: 'archive-only',
				requirement:
					'Keep stale or non-release research out of release wording and owner queues unless new evidence changes a top claim.',
				forbiddenShortcut:
					'Do not promote archive-only material as product, correctness, performance, QSS, or release evidence.',
			},
		],
		failureEvidence: decision?.evidenceMissing ?? [],
		acceptanceCriteria:
			decision?.nextOwnerAction ??
			'Product/release owner classifies each untriaged path before citing any research-derived claim.',
		allowedWording: decision?.allowedWording ?? RESEARCH_SURFACE_HYGIENE_BLOCKER.allowedWording,
		forbiddenWording: decision?.forbiddenWording ?? [
			...RESEARCH_SURFACE_HYGIENE_BLOCKER.forbiddenWording,
		],
		qssContrast: decision?.qssContrast ?? [
			'QSS comparison is blocked until research evidence is classified into release decisions instead of broad notes.',
		],
		nextOwnerAction: decision?.nextOwnerAction ?? RESEARCH_SURFACE_HYGIENE_BLOCKER.ownerAction,
		stopCondition:
			'Stop only when every path reported by the dirty inventory command is classified as accepted evidence, active owner blocker, or archive-only, and release-proof-index tests pass.',
		boundary:
			'Compact research hygiene decision packet. It is not research approval, not archive deletion, not owner classification by itself, and not permission to cite untriaged research in release wording.',
	}
}

function researchHygieneReleaseRoutingSummary(
	result: ReleaseProofIndexResult,
	inventorySnapshot: ReleaseProofResearchHygieneInventorySnapshot,
): ReleaseProofResearchHygieneReleaseRoutingSummary {
	const formatEntry = (entry: ReleaseProofResearchHygieneInventoryEntry) =>
		`${entry.path}: ${entry.reason} Next: ${entry.nextOwnerAction}`
	return {
		acceptedReleaseEvidence: inventorySnapshot.classifiedEntries
			.filter((entry) => entry.classification === 'accepted-evidence')
			.map(formatEntry),
		blockedClaims: result.releaseDecisionBoard.doNotPromoteYet.map(
			(item) =>
				`${item.name}: ${item.allowedWording} Forbidden: ${item.forbiddenWording.join('; ')}`,
		),
		ownerReadyImplementationTasks:
			result.releaseDecisionBoard.implementationReadyOwnerActionQueue.map(
				(row) =>
					`${row.ownerLoop}/${row.name}${row.requirementId ? `/${row.requirementId}` : ''}: ${row.nextOwnerAction}`,
			),
		archiveDeferMaterial: inventorySnapshot.classifiedEntries
			.filter((entry) => entry.classification === 'archive-only')
			.map(formatEntry),
		boundary:
			'Compact routing summary only. It separates accepted evidence, blocked claims, owner-ready implementation work, and archive/defer material without promoting raw research or creating new work surfaces.',
	}
}

function researchHygieneInventorySnapshot(
	command: string,
): ReleaseProofResearchHygieneInventorySnapshot {
	try {
		const stdout = execFileSync(
			'git',
			[
				'status',
				'--short',
				'research',
				'scripts/ascend-loop-manager.ts',
				'tmp/ascend-loop-manager',
			],
			{
				cwd: process.cwd(),
				encoding: 'utf8',
			},
		)
		const entries = stdout
			.split('\n')
			.map((line) => line.trimEnd())
			.filter((line) => line.length > 0)
			.map((line) => {
				const path = line.slice(3)
				return {
					statusCode: line.slice(0, 2).trim(),
					path,
					...classifyResearchHygienePath(path),
				}
			})
		const classifiedEntries = entries.filter(
			(entry) => entry.classification !== 'unclassified-owner-decision-required',
		)
		const unclassifiedEntries = entries.filter(
			(entry) => entry.classification === 'unclassified-owner-decision-required',
		)
		const statusCodeCounts = countBy(entries, (entry) => entry.statusCode)
		const classificationCounts = countBy(entries, (entry) => entry.classification)
		const rootCounts = researchHygieneInventoryRootCounts(entries)
		return {
			command,
			status: 'inventory-collected',
			decision:
				unclassifiedEntries.length > 0
					? 'owner-classification-required'
					: entries.length > 0
						? 'current-inventory-classified-release-routing-required'
						: 'no-unclassified-paths-currently-visible',
			dirtyPathCount: entries.length,
			classifiedPathCount: classifiedEntries.length,
			unclassifiedPathCount: unclassifiedEntries.length,
			statusCodeCounts,
			classificationCounts,
			rootCounts,
			untrackedDirectoryCount: entries.filter(
				(entry) => entry.statusCode === '??' && entry.path.endsWith('/'),
			).length,
			modifiedFileCount: entries.filter((entry) => entry.statusCode === 'M').length,
			classifiedEntries,
			unclassifiedEntries,
			boundary:
				'Inventory snapshot is current git-status routing evidence only. Classification routes paths to release-proof citation, active owner blocker, or archive-only handling; it is not permission to cite raw research paths in release wording.',
		}
	} catch (error) {
		return {
			command,
			status: 'inventory-command-failed',
			decision: 'inventory-rerun-required',
			dirtyPathCount: 0,
			classifiedPathCount: 0,
			unclassifiedPathCount: 0,
			statusCodeCounts: {},
			classificationCounts: {},
			rootCounts: { research: 0, scripts: 0, tmp: 0, other: 0 },
			untrackedDirectoryCount: 0,
			modifiedFileCount: 0,
			classifiedEntries: [],
			unclassifiedEntries: [],
			failure: error instanceof Error ? error.message : String(error),
			boundary:
				'Inventory snapshot failed to collect. Owners must rerun the dirty inventory command before citing any research-derived claim.',
		}
	}
}

function classifyResearchHygienePath(
	path: string,
): Omit<ReleaseProofResearchHygieneInventoryEntry, 'statusCode' | 'path'> {
	if (
		path === 'research/experiments/index.md' ||
		path === 'research/experiments/runs/2026/2026-05-15-fixture-decision-packet/'
	) {
		return {
			classification: 'accepted-evidence',
			reason:
				'Current fixture-decision research is already folded into release-proof-index compact fixture-decision output; cite the proof packet, not the raw research log.',
			nextOwnerAction:
				'Keep the release-proof-index fixture-decision packet as the canonical citation and do not promote raw experiment notes.',
		}
	}

	if (path === 'research/excel-corpus/') {
		return {
			classification: 'active-owner-blocker',
			reason:
				'Local workbook corpus material may contain useful public-fixture candidates, but it is not release evidence until license, provenance, and tracked-fixture policy are resolved.',
			nextOwnerAction:
				'Product/release owner classifies each workbook as vendored public fixture, private/local-only diagnostic, or deletion/archive candidate before any claim uses it.',
		}
	}

	if (path === 'scripts/ascend-loop-manager.ts' || path === 'tmp/ascend-loop-manager/') {
		return {
			classification: 'active-owner-blocker',
			reason:
				'Loop-manager script and board state are operational steering surfaces, not product proof or release evidence.',
			nextOwnerAction:
				'Release owner decides whether to track, ignore, or archive the manager tool/state; do not cite it for Ascend runtime claims.',
		}
	}

	if (path.startsWith('research/')) {
		return {
			classification: 'archive-only',
			reason:
				'Broad research note is not represented as a current release-proof-index claim decision and must stay out of release wording.',
			nextOwnerAction:
				'Leave as archive-only unless a future owner converts one specific finding into release-proof evidence with commands, forbidden wording, and a next action.',
		}
	}

	return {
		classification: 'unclassified-owner-decision-required',
		reason: 'Path is outside the current research hygiene routing rules.',
		nextOwnerAction:
			'Product/release owner must classify this path before citing it or expanding the release-proof index.',
	}
}

type ResearchHygieneManifestEntry = {
	readonly file?: unknown
	readonly size_bytes?: unknown
	readonly features?: Record<string, unknown>
}

function researchHygieneLocalExcelCorpus(): ReleaseProofResearchHygieneLocalExcelCorpus {
	const corpusPath = 'research/excel-corpus'
	const manifestPath = 'research/excel-corpus/manifest.json' as const
	const ownerAction =
		'Product/release owner reviews each local corpus workbook for source URL, redistribution license, privacy, size, and release relevance; classify each as approved public fixture, private/local-only diagnostic, or archive/deletion candidate before any release claim cites it.'
	const forbiddenWording = [
		'Do not cite research/excel-corpus workbooks as public release fixtures until provenance, license, privacy, size, and vendoring policy are approved.',
		'Do not use the local NYC 311 materialized workbook as a release benchmark input unless acquisition, materialization, and size policy are explicitly approved.',
	]

	try {
		if (!existsSync(corpusPath)) {
			return {
				path: corpusPath,
				status: 'local-corpus-missing',
				releaseDecision: 'active-owner-blocker-not-release-evidence',
				fileCount: 0,
				workbookCount: 0,
				totalBytes: 0,
				manifestPath,
				manifestEntryCount: 0,
				manifestMissingFiles: [],
				files: [],
				ownerAction,
				forbiddenWording,
				boundary:
					'Local Excel corpus directory is absent. This satisfies no fixture evidence, creates no release claim, and is not permission to use these files in release wording.',
			}
		}

		const files = execFileSync('find', [corpusPath, '-maxdepth', '1', '-type', 'f', '-print'], {
			cwd: process.cwd(),
			encoding: 'utf8',
		})
			.split('\n')
			.map((line) => line.trim())
			.filter((line) => line.length > 0)
			.sort()

		const manifest = readResearchHygieneManifest(manifestPath)
		const manifestByFile = new Map(manifest.map((entry) => [String(entry.file), entry]))
		const corpusFiles = files.map((path) => {
			const fileName = basename(path)
			const manifestEntry = manifestByFile.get(fileName)
			const sizeBytes = statSync(path).size
			return {
				path,
				sizeBytes,
				manifestListed: Boolean(manifestEntry),
				workbookKind: researchHygieneWorkbookKind(path),
				featureSignals: researchHygieneFeatureSignals(manifestEntry),
				releaseStatus: 'local-only-owner-review-required' as const,
				nextOwnerAction:
					'Review source, license, privacy, size, and fixture relevance before promoting this local corpus file into tracked release evidence.',
			}
		})
		const workbookFiles = corpusFiles.filter((file) => file.workbookKind !== 'metadata')
		const largestFile = corpusFiles.reduce<ReleaseProofResearchHygieneLocalExcelCorpusFile | null>(
			(current, file) => (!current || file.sizeBytes > current.sizeBytes ? file : current),
			null,
		)
		const corpusFileNames = new Set(corpusFiles.map((file) => basename(file.path)))
		const manifestMissingFiles = manifest
			.map((entry) => String(entry.file))
			.filter((file) => !corpusFileNames.has(file))
			.sort()

		return {
			path: corpusPath,
			status: 'local-corpus-inventory-present',
			releaseDecision: 'active-owner-blocker-not-release-evidence',
			fileCount: corpusFiles.length,
			workbookCount: workbookFiles.length,
			totalBytes: corpusFiles.reduce((sum, file) => sum + file.sizeBytes, 0),
			largestFile: largestFile?.path,
			manifestPath,
			manifestEntryCount: manifest.length,
			manifestMissingFiles,
			files: corpusFiles,
			ownerAction,
			forbiddenWording,
			boundary:
				'Local Excel corpus inventory is owner-routing evidence only. It is not vendored public fixture approval, not license review, not privacy review, and not permission to use these files in release wording or benchmark claims.',
		}
	} catch (error) {
		return {
			path: corpusPath,
			status: 'local-corpus-unreadable',
			releaseDecision: 'active-owner-blocker-not-release-evidence',
			fileCount: 0,
			workbookCount: 0,
			totalBytes: 0,
			manifestPath,
			manifestEntryCount: 0,
			manifestMissingFiles: [],
			files: [],
			ownerAction,
			forbiddenWording,
			failure: error instanceof Error ? error.message : String(error),
			boundary:
				'Local Excel corpus inventory could not be read. Product/release owners must rerun the research hygiene packet before citing corpus-derived evidence; this is not permission to use these files in release wording.',
		}
	}
}

function readResearchHygieneManifest(path: string): readonly ResearchHygieneManifestEntry[] {
	if (!existsSync(path)) return []
	const parsed = JSON.parse(readFileSync(path, 'utf8')) as unknown
	if (!Array.isArray(parsed)) return []
	return parsed.filter((entry): entry is ResearchHygieneManifestEntry => {
		return typeof entry === 'object' && entry !== null && 'file' in entry
	})
}

function researchHygieneWorkbookKind(
	path: string,
): ReleaseProofResearchHygieneLocalExcelCorpusFile['workbookKind'] {
	if (path.endsWith('.xlsx')) return 'xlsx'
	if (path.endsWith('.xlsm')) return 'xlsm'
	return 'metadata'
}

function researchHygieneFeatureSignals(
	entry: ResearchHygieneManifestEntry | undefined,
): readonly string[] {
	if (!entry?.features) return []
	return Object.entries(entry.features)
		.filter(([, value]) => value === true)
		.map(([feature]) => feature)
		.sort()
}

function researchHygieneLoopManagerState(): ReleaseProofResearchHygieneLoopManagerState {
	const scriptPath = 'scripts/ascend-loop-manager.ts'
	const statePath = 'tmp/ascend-loop-manager'
	const ownerAction =
		'Release owner decides whether the loop-manager script and tmp steering state should be tracked operational tooling, ignored local state, or archived; keep it out of product, correctness, performance, and QSS claims until that decision is made.'
	const forbiddenWording = [
		'Do not cite scripts/ascend-loop-manager.ts or tmp/ascend-loop-manager as Ascend runtime proof, product capability, benchmark evidence, or release artifact.',
		'Do not treat manager board text, manual steers, or session-derived state as owner approval, public documentation, or claim evidence.',
	]

	try {
		if (!existsSync(scriptPath) && !existsSync(statePath)) {
			return {
				paths: [scriptPath, statePath],
				status: 'loop-manager-state-missing',
				releaseDecision: 'active-owner-blocker-not-release-evidence',
				fileCount: 0,
				totalBytes: 0,
				files: [],
				ownerAction,
				forbiddenWording,
				boundary:
					'Loop-manager operational state is absent. This creates no release claim and is not permission to cite manager state in release wording.',
			}
		}

		const filePaths = [
			...(existsSync(scriptPath) ? [scriptPath] : []),
			...(existsSync(statePath)
				? execFileSync('find', [statePath, '-maxdepth', '1', '-type', 'f', '-print'], {
						cwd: process.cwd(),
						encoding: 'utf8',
					})
						.split('\n')
						.map((line) => line.trim())
						.filter((line) => line.length > 0)
						.sort()
				: []),
		]
		const files = filePaths.map((path) => ({
			path,
			sizeBytes: statSync(path).size,
			stateKind: researchHygieneLoopManagerStateKind(path),
			releaseStatus: 'operational-state-owner-review-required' as const,
			nextOwnerAction:
				'Release owner classifies this manager state as tracked operational tooling, ignored local state, or archive-only material before any release workflow cites it.',
		}))

		return {
			paths: [scriptPath, statePath],
			status: 'loop-manager-state-present',
			releaseDecision: 'active-owner-blocker-not-release-evidence',
			fileCount: files.length,
			totalBytes: files.reduce((sum, file) => sum + file.sizeBytes, 0),
			files,
			ownerAction,
			forbiddenWording,
			boundary:
				'Loop-manager state inventory is owner-routing evidence only. It is operational steering state, not Ascend runtime proof, not product documentation, not benchmark evidence, and not permission to use these files in release wording.',
		}
	} catch (error) {
		return {
			paths: [scriptPath, statePath],
			status: 'loop-manager-state-unreadable',
			releaseDecision: 'active-owner-blocker-not-release-evidence',
			fileCount: 0,
			totalBytes: 0,
			files: [],
			ownerAction,
			forbiddenWording,
			failure: error instanceof Error ? error.message : String(error),
			boundary:
				'Loop-manager operational state inventory could not be read. Release owners must rerun the research hygiene packet before citing manager-derived evidence; this is not permission to use these files in release wording.',
		}
	}
}

function researchHygieneLoopManagerStateKind(
	path: string,
): ReleaseProofResearchHygieneLoopManagerStateFile['stateKind'] {
	if (path === 'scripts/ascend-loop-manager.ts') return 'manager-script'
	if (path.endsWith('/board.md')) return 'current-board'
	if (path.endsWith('/manual-steers.md')) return 'manual-steers'
	if (path.endsWith('/north-star-reflection.md')) return 'north-star-reflection'
	return 'manager-state-file'
}

function researchHygieneInventoryRootCounts(
	entries: readonly ReleaseProofResearchHygieneInventoryEntry[],
): Readonly<Record<'research' | 'scripts' | 'tmp' | 'other', number>> {
	const counts = { research: 0, scripts: 0, tmp: 0, other: 0 }
	for (const entry of entries) {
		if (entry.path.startsWith('research/')) {
			counts.research += 1
		} else if (entry.path.startsWith('scripts/')) {
			counts.scripts += 1
		} else if (entry.path.startsWith('tmp/')) {
			counts.tmp += 1
		} else {
			counts.other += 1
		}
	}
	return counts
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
			...CLAIM_PORTFOLIO.slice(2).map(portfolioClaimArchivedResearchNote),
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
			'This matrix is a release-priority gate for the top two claims only. Formula rename, token-bounded agent view, viewport history, release proof bundle, oracle routing, property journal laws, columnar sidecars, agent observability, and untriaged research-surface work remain archived until they change top-claim implementation priority.',
	}
}

function portfolioClaimArchivedResearchNote(
	claim: ReleaseProofPortfolioClaim,
): ReleaseProofQssArchivedResearchNote {
	const deferredClaim = DEFERRED_CLAIMS.find((candidate) => candidate.name === claim.name)
	return {
		name: claim.name,
		status: 'archived-research-note',
		ownerLoops: deferredClaim ? [...deferredClaim.ownerLoops] : [...claim.likelyHandoffOwner],
		reason: deferredClaim?.reason ?? claim.boundary,
		killCriterion: deferredClaim?.killCriterion ?? claim.killCriterion,
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
			'Execution or semantic support for active content, Power Query, Data Model, chart style/color, analytical sidecar, or calc metadata parts.',
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
			evidenceId: 'safe-open-encrypted-password-no-echo-tests',
			kind: 'test',
			command:
				'bun test packages/sdk/src/open-plan.test.ts apps/cli/src/cli.test.ts apps/api/src/server.test.ts apps/mcp/src/index.test.ts -t "encrypted workbook passwords|open-plan --password|plan --password" --timeout 30000',
			path: 'packages/sdk/src/open-plan.test.ts; apps/cli/src/cli.test.ts; apps/api/src/server.test.ts; apps/mcp/src/index.test.ts',
			acceptedScope:
				'Commit 92acf61e tightens SDK/CLI/API/MCP encrypted-workbook password no-echo assertions while preserving fail-closed encrypted edit/export behavior.',
			boundary:
				'No password recovery, encryption removal, re-encrypted export support, or guarantee that arbitrary logs outside these surfaces are scrubbed.',
		},
		{
			evidenceId: 'sdk-encrypted-output-fail-closed-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/sdk.test.ts -t "opens encrypted XLSX fixtures|encrypted workbook saves fail closed unless decrypted export is explicit" --timeout 30000',
			path: 'packages/sdk/src/workbook.ts; packages/sdk/src/sdk.test.ts',
			acceptedScope:
				'Commit bd6e4358 proves encrypted workbook file saves fail closed unless decrypted export is explicit, leave blocked outputs absent, and approved decrypted outputs reopen with the intended edit.',
			boundary:
				'Encrypted output fail-closed evidence only; it does not provide re-encryption support, password recovery, encrypted output writing, or arbitrary log scrubbing.',
		},
		{
			evidenceId: 'sdk-encrypted-session-package-inspection-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/sdk.test.ts -t "rawPackagePart inspects decrypted package parts after password open" --timeout 30000',
			path: 'packages/sdk/src/session.ts; packages/sdk/src/sdk.test.ts',
			acceptedScope:
				'Commit 7ff308c9 makes WorkbookDocument packageGraph/rawPackagePart inspect decrypted package bytes after password open, including metadata-only sessions, while preserving fail-closed behavior for dirty encrypted exports.',
			boundary:
				'Encrypted package-inspection evidence only; it does not provide re-encryption support, password recovery, encrypted output writing, malware scanning, or file trust.',
		},
		{
			evidenceId: 'sdk-encrypted-agent-commit-policy-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "encrypted workbook commits" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/workbook.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commits 5ab9365d, a9ae9276, and d8d9ae1a make SDK agent commits fail closed before apply/write for encrypted sources unless allowDecryptedExport is explicit, and prove explicit decrypted XLSX/text outputs can reopen while the encrypted source remains unchanged.',
			boundary:
				'SDK agent commit policy evidence only; it does not provide re-encryption support, password recovery, encrypted output writing, or arbitrary log scrubbing.',
		},
		{
			evidenceId: 'sdk-signed-output-fail-closed-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/sdk.test.ts -t "dirty signed workbooks require explicit signature invalidation before export|signed workbook saves fail closed unless signature invalidation is explicit" --timeout 30000',
			path: 'packages/sdk/src/workbook.ts; packages/sdk/src/sdk.test.ts',
			acceptedScope:
				'Commits 4502acbf and 17fa6741 make dirty signed workbook byte export and file save fail closed unless signature invalidation is explicit, while keeping package-graph audit possible and proving approved unsigned outputs reopen without signature findings.',
			boundary:
				'Synthetic signed-package topology fail-closed evidence only; it does not prove real signer identity, signature verification, re-signing, certificate trust, attestation, or public signed-workbook fixture coverage.',
		},
		{
			evidenceId: 'sdk-vba-signature-invalidation-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/sdk.test.ts -t "VBA project signatures fail closed while approved exports preserve macros" --timeout 30000',
			path: 'packages/io-xlsx/src/writer/index.ts; packages/sdk/src/sdk.test.ts',
			acceptedScope:
				'Commit e5d0ee17 treats VBA project signature parts as signature capsules: edited macro-enabled workbooks fail closed unless signature invalidation is explicit, approved exports preserve the VBA project, drop vbaProjectSignature, reopen with macro inventory intact, and remove signature findings.',
			boundary:
				'Synthetic VBA-signature package topology evidence only; it does not prove macro safety, macro execution, signature verification, re-signing, signer identity, malware scanning, or public signed-VBA fixture coverage.',
		},
		{
			evidenceId: 'sdk-high-risk-stream-export-approval-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/sdk.test.ts -t "high-risk workbook streams require the same explicit export approvals" --timeout 30000',
			path: 'packages/sdk/src/workbook.ts; packages/sdk/src/sdk.test.ts',
			acceptedScope:
				'Commit ca3c3296 makes toStream use the same explicit approvals as toBytes for edited encrypted and signed workbooks, proving blocked streams reject, approved decrypted/signed-invalidating streams reopen, and original high-risk workbooks remain guarded.',
			boundary:
				'High-risk stream export approval evidence only; it does not provide re-encryption support, signature preservation, re-signing, signature validation, malware scanning, or file trust.',
		},
		{
			evidenceId: 'sdk-high-risk-text-export-approval-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/sdk.test.ts -t "high-risk workbook text saves require the same explicit export approvals" --timeout 30000',
			path: 'packages/sdk/src/workbook.ts; packages/sdk/src/sdk.test.ts',
			acceptedScope:
				'Commits b7e8eccc and a004fb4a make values-only text exports from encrypted, signed, macro, and signed-macro workbooks fail closed unless the exact decrypted-export, signature-loss, and active-content-loss approvals are present; the combined signed-macro case requires both signature and active-content approval and still keeps XLSX/XLSM byte export fail-closed without signature invalidation.',
			boundary:
				'High-risk values-only text export evidence only; it does not provide re-encryption, signature preservation, re-signing, macro safety, active-content execution, malware scanning, package-preserving text output, or file trust.',
		},
		{
			evidenceId: 'sdk-signed-agent-text-commit-policy-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "signed workbook commits can explicitly write unsigned text output and verify it" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commits d8d9ae1a and 90f4c248 prove approved SDK agent commits from signed sources can write unsigned TSV output, reopen/verify the text output, record not-applicable-non-xlsx package audit policy, and leave the signed source bytes unchanged.',
			boundary:
				'Synthetic signed-package approved text-output evidence only; it does not prove signature preservation, re-signing, signature validation, signer identity, or public signed-workbook fixture coverage.',
		},
		{
			evidenceId: 'mcp-agent-workflow-open-plan-first',
			kind: 'test',
			command: 'bun test apps/mcp/src/index.test.ts -t "agent workflow" --timeout 30000',
			path: 'apps/mcp/src/index.ts; apps/mcp/src/index.test.ts',
			acceptedScope:
				'Commit d4ee22e1 makes MCP agent workflow guidance start with ascend.open_plan before hydrating unknown XLSX/XLSM files and keeps encrypted-password handling explicit.',
			boundary:
				'MCP workflow guidance evidence only; it does not prove every outside-user workflow, latency, or safe-open publication policy.',
		},
		{
			evidenceId: 'agent-safe-edit-example-proof',
			kind: 'proof-artifact',
			command:
				'bun run examples/agent-safe-edit.ts /private/tmp/ascend-agent-safe-edit-input.xlsx /private/tmp/ascend-agent-safe-edit-output.xlsx',
			path: 'examples/agent-safe-edit.ts',
			acceptedScope:
				'Commit cab4bff1 makes the agent-safe-edit example reopen the committed workbook, verify the edited formula cell, run check/lint, and emit concrete follow-up verification commands after commit. Commit adc9c1e1 makes the root SDK example emit the shared proofBundle.safeToUse, whatChanged, and whySafe gate evidence using createAgentWorkflowProofSummary. Commit b93f3493 makes the SDK example print reopened-output post-write proof fields for data connections, formula state, security, and visuals.',
			boundary:
				'Local example workflow proof only; it does not prove package publication, arbitrary workbook safety, performance, external trust, or every SDK/CLI/API/MCP workflow.',
		},
		{
			evidenceId: 'api-agent-safe-edit-example-proof',
			kind: 'proof-artifact',
			command:
				'bun run examples/agent-safe-edit-http.ts /private/tmp/ascend-agent-safe-edit-http-input.xlsx /private/tmp/ascend-agent-safe-edit-http-output.xlsx',
			path: 'examples/agent-safe-edit-http.ts; examples/agent-safe-edit-http.test.ts',
			acceptedScope:
				'Commit 01d08512 adds a runnable HTTP API safe-edit example that opens/plans, inspects, reads, prepares a plan, commits, reopens, checks, lints, verifies the edited formula cell, and has a focused example test for the workflow contract. Commit 2569626c indexes the runnable HTTP workflow in SDK agent-doc search so outside users can discover it. Commit 37794635 switches it to the packaged @ascend/api import for installed-consumer proof. Commit 7afcd630 makes the HTTP example emit proofBundle.safeToUse, whatChanged, and whySafe evidence. Commit a8e15d9b adds an explicit trust-preflight step with trust posture, finding count, and next actions before inspect/read. Commit b93f3493 makes the HTTP example print and test reopened-output post-write proof fields for data connections, formula state, security, and visuals.',
			boundary:
				'Local generated-workbook API workflow proof only; it does not prove arbitrary workbook safety, public workbook behavior, package publication, performance, external trust, or every SDK/CLI/API/MCP workflow.',
		},
		{
			evidenceId: 'mcp-agent-safe-edit-example-proof',
			kind: 'proof-artifact',
			command: 'bun test examples/agent-safe-edit-mcp.test.ts --timeout 30000',
			path: 'examples/agent-safe-edit-mcp.ts; examples/agent-safe-edit-mcp.test.ts; packages/sdk/src/agent-docs.ts; packages/sdk/src/agent-docs.test.ts',
			acceptedScope:
				'Commit de45eb83 adds a runnable MCP safe-edit example and focused test covering agent_workflow discovery, open-plan, inspect, read, prepared plan, commit, reopen, check, lint, and edited formula verification, and indexes the runnable MCP workflow in SDK agent-doc search. Commit 37794635 switches it to the packaged @ascend/mcp import for installed-consumer proof. Commit 7afcd630 makes the MCP example emit proofBundle.safeToUse, whatChanged, and whySafe evidence. Commit a8e15d9b adds an explicit trust-preflight step with trust posture, finding count, and next actions before inspect/read. Commit b93f3493 makes the MCP example print and test reopened-output post-write proof fields for data connections, formula state, security, and visuals.',
			boundary:
				'Local generated-workbook MCP workflow proof only; it does not prove arbitrary workbook safety, public workbook behavior, package publication, performance, external trust, or every SDK/CLI/API/MCP workflow.',
		},
		{
			evidenceId: 'examples-package-safe-edit-scripts-proof',
			kind: 'test',
			command: 'bun test examples/package-scripts.test.ts --timeout 30000',
			path: 'examples/package.json; examples/package-scripts.test.ts; examples/README.md',
			acceptedScope:
				'Commit bbea493c adds examples package scripts for SDK, HTTP, and MCP safe-edit workflows and proves each script runs inspect/open-plan, prepared plan, commit, reopen/check/lint, and edited formula verification from the examples package. Commit 37794635 adds package dependencies so the runnable HTTP/MCP examples use packaged app imports. Commit 7afcd630 verifies HTTP and MCP scripts surface proofBundle.safeToUse and changed-cell evidence. Commit adc9c1e1 verifies the SDK script surfaces proofBundle.safeToUse and changed-cell evidence too. Commit a8e15d9b proves the HTTP and MCP scripts include trust-preflight in the workflow. Commit b93f3493 proves package scripts surface reopened-output post-write proof fields for data connections, formula state, security, and visuals.',
			boundary:
				'Local examples-package workflow proof only; it does not prove package publication, registry install behavior, arbitrary workbook safety, performance, external trust, or every SDK/CLI/API/MCP workflow.',
		},
		{
			evidenceId: 'root-package-safe-edit-scripts-proof',
			kind: 'test',
			command: 'bun test examples/root-scripts.test.ts --timeout 30000',
			path: 'package.json; examples/root-scripts.test.ts; examples/README.md; apps/cli/src/commands/agent-init.ts; apps/api/src/server.ts; apps/mcp/src/index.ts; scripts/release-apps-smoke.ts',
			acceptedScope:
				'Commit a09660be adds root package scripts for SDK, HTTP, and MCP safe-edit workflows, proves each root script runs the generated-workbook workflow, and updates CLI/API/MCP workflow discovery plus installed app smoke checks to point at the root commands. Commit 7afcd630 verifies HTTP and MCP root scripts surface proofBundle.safeToUse and changed-cell evidence. Commit adc9c1e1 verifies the SDK root script surfaces proofBundle.safeToUse and changed-cell evidence too. Commit a8e15d9b proves the HTTP and MCP root scripts include trust-preflight in the workflow. Commit b93f3493 proves root scripts surface reopened-output post-write proof fields for data connections, formula state, security, and visuals.',
			boundary:
				'Local root-package script evidence only; it does not prove package publication, registry install behavior, arbitrary workbook safety, performance, external trust, or every SDK/CLI/API/MCP workflow.',
		},
		{
			evidenceId: 'workflow-example-proof-context',
			kind: 'test',
			command:
				'bun test apps/cli/src/cli.test.ts apps/api/src/server.test.ts apps/mcp/src/index.test.ts -t "agent-init prints the canonical agent workflow contract|/agent-workflow exposes the API safe edit contract|ascend.agent_workflow exposes machine-readable safe edit guidance" --timeout 30000',
			path: 'apps/cli/src/commands/agent-init.ts; apps/cli/src/cli.test.ts; apps/api/src/server.ts; apps/api/src/server.test.ts; apps/mcp/src/index.ts; apps/mcp/src/index.test.ts; scripts/release-apps-smoke.ts',
			acceptedScope:
				'Commit f3347e17 exposes workflow example proof context on CLI agent-init and API/MCP agent-workflow surfaces: repository-root workdir, source-checkout/Bun prerequisites, and the root-scripts proof command are visible to outside users and guarded by installed app smoke checks.',
			boundary:
				'Workflow-discovery proof-context evidence only; it does not prove package publication, registry install behavior, arbitrary workbook safety, example runtime success beyond the cited root-script proof, performance, or external trust.',
		},
		{
			evidenceId: 'installed-sdk-safe-edit-example-proof',
			kind: 'test',
			command: 'bun test examples/package-install-safe-edit.test.ts --timeout 30000',
			path: 'examples/package-install-safe-edit.ts; examples/package-install-safe-edit.test.ts; scripts/release-sdk-smoke.ts',
			acceptedScope:
				'Commit 5915794f makes the installed SDK safe-edit example emit a machine-readable proof bundle with safeToUse, changed-cell before/after evidence, open-plan/trust/plan/commit/reopen gates, hashes, check/lint validity, post-write validity, and release-sdk-smoke assertions. Commit 56cd4aa0 moves the outside-user command to the packaged bin `node_modules/.bin/ascend-sdk-safe-edit` and keeps the focused installed example proof passing. Commit a5fa3006 promotes the proof bundle into an SDK `createAgentWorkflowProofSummary` helper with plan-linked, write-policy, reopen-verify, and package-graph gates reused by the installed example. This closes the prior blank-vs-empty proof contract mismatch for the generated example.',
			boundary:
				'Installed SDK generated-workbook proof only; it does not prove registry publication, arbitrary workbook safety, hosted service readiness, external trust, performance, or public workbook behavior.',
		},
		{
			evidenceId: 'installed-cli-safe-edit-example-proof',
			kind: 'test',
			command:
				'bun test apps/cli/src/cli.test.ts -t "example-safe-edit runs the packaged inspect plan commit reopen verify workflow" --timeout 30000',
			path: 'apps/cli/src/commands/example-safe-edit.ts; apps/cli/src/cli.test.ts; scripts/release-apps-smoke.ts',
			acceptedScope:
				'Commit 3d630232 adds the installed CLI `ascend example-safe-edit <file.xlsx> <out.xlsx>` workflow with proofBundle.safeToUse, changed-cell evidence, safety gates, and reopen/check/lint verification. Commit 7e7183df makes the CLI example use the shared SDK createAgentWorkflowProofSummary helper and proves plan-linked, write-policy, reopen-verify, and package-graph gates through the CLI test and installed app smoke.',
			boundary:
				'Installed CLI generated-workbook proof only; it does not prove registry publication, arbitrary workbook safety, hosted service readiness, external trust, performance, or public workbook behavior.',
		},
		{
			evidenceId: 'installed-sdk-workflow-discovery-proof',
			kind: 'test',
			command:
				'bun test apps/cli/src/cli.test.ts apps/api/src/server.test.ts apps/mcp/src/index.test.ts -t "agent-init prints the canonical agent workflow contract|/agent-workflow exposes the API safe edit contract|ascend.agent_workflow exposes machine-readable safe edit guidance" --timeout 30000',
			path: 'apps/cli/src/commands/agent-init.ts; apps/cli/src/cli.test.ts; apps/api/src/server.ts; apps/api/src/server.test.ts; apps/mcp/src/index.ts; apps/mcp/src/index.test.ts; scripts/release-apps-smoke.ts',
			acceptedScope:
				'Commits f8d63593, cc689bcc, and 56cd4aa0 expose the installed SDK safe-edit package-bin command and proof-bundle context through CLI agent-init, API /agent-workflow, MCP ascend.agent_workflow, the MCP workflow resource, API package tests, and installed app smoke checks.',
			boundary:
				'Installed workflow-discovery evidence only; it does not prove registry download success, arbitrary workbook safety, hosted service readiness, public workbook behavior, package publication, or full API package test health.',
		},
		{
			evidenceId: 'api-custom-ui-active-content-proof',
			kind: 'test',
			command:
				'bun test apps/api/src/server.test.ts -t "/active-content reports custom UI callbacks" --timeout 30000',
			path: 'apps/api/src/server.ts; apps/api/src/server.test.ts; docs/EXCEL_BEHAVIOR_COMPATIBILITY_MATRIX.md; packages/sdk/src/excel-behavior-compatibility-matrix.test.ts',
			acceptedScope:
				'Commit 3653fd6f reports generated RibbonX custom UI callbacks through API /active-content with relationship provenance, preservedCustomUi compatibility metadata, blocked execution policy, and active-content warnings.',
			boundary:
				'Generated custom UI agent-context proof only; it does not replace broader public custom UI fixture coverage or authorize Custom UI safety, macro safety, active-content safety, or trust wording.',
		},
		{
			evidenceId: 'public-shape-macro-active-content-proof',
			kind: 'test',
			command:
				'bun test fixtures/corpus/active-content-contract.test.ts -t "drawing shape macro bindings" --timeout 60000',
			path: 'packages/core/src/active-content.ts; packages/io-xlsx/src/reader/active-content.ts; packages/io-xlsx/src/reader/index.ts; packages/sdk/src/read-view.ts; packages/sdk/src/workbook-trust.ts; fixtures/corpus/active-content-contract.test.ts',
			acceptedScope:
				'Commit c44c5480 reports public LibreOffice drawing shape macro bindings as blocked active content with shape identity, trust findings, read-view warnings, and safe-edit reopen/package-preservation proof.',
			boundary:
				'Public shape-macro reporting and preservation evidence only; it does not authorize macro execution, malware scanning, active-content safety, trusted-source behavior, or broad custom UI safety wording.',
		},
		{
			evidenceId: 'public-shape-macro-commit-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public shape macro drawings as blocked active content after safe edit" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts; packages/sdk/src/index.ts; packages/sdk/src/index-exports.test.ts',
			acceptedScope:
				'Commit 9df35fd6 proves safe edits against a public LibreOffice shape-macro workbook preserve source bytes, reopen output, keep package graph audits passing, and report post-write shape macros as blocked active content with shape identity. Commit fef9294f expands post-write visual proof to report drawing-object details including drawing part path, source, kind, id, name, text, relationship ids, and chart relationship targets, with root SDK proof types exported.',
			boundary:
				'Public shape-macro drawing-object commit audit evidence only; it does not authorize macro execution, malware scanning, active-content safety, trusted-source behavior, broad custom UI safety wording, or full drawing/layout parity.',
		},
		{
			evidenceId: 'public-activex-blocked-execution-policy-proof',
			kind: 'test',
			command:
				'bun test fixtures/corpus/active-content-contract.test.ts packages/io-xlsx/src/reader/active-content.test.ts -t "ActiveX control|ActiveX controls" --timeout 60000',
			path: 'packages/io-xlsx/src/reader/index.ts; packages/io-xlsx/src/reader/active-content.test.ts; fixtures/corpus/active-content-contract.test.ts',
			acceptedScope:
				'Commit 2459f79a marks public ActiveX and macro-bearing form-control inventory with blocked executionPolicy metadata while preserving ActiveX XML, binary, worksheet, and VML identity through safe edit/reopen evidence.',
			boundary:
				'ActiveX/form-control execution metadata only; it does not execute, sandbox, scan, trust, or semantically edit ActiveX controls.',
		},
		{
			evidenceId: 'chart-series-source-fail-closed-proof',
			kind: 'test',
			command:
				'bun test packages/engine/src/operations.test.ts -t "setChartSeriesSource rejects source fields absent from parsed chart series" --timeout 30000',
			path: 'packages/engine/src/operations/visual-ops.ts; packages/engine/src/operations.test.ts',
			acceptedScope:
				'Commit 44f43d12 makes setChartSeriesSource fail closed when an edit would insert name/category/value source fields absent from the parsed chart series, preserving the workbook model unchanged.',
			boundary:
				'Chart series-source insertion guard only; it does not prove full chart authoring, layout editing, chart XML semantic preservation, or chart XML byte passthrough.',
		},
		{
			evidenceId: 'public-chart-source-commit-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public chart source edits through save and reopen audits" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commit 1eaf28ff proves a public ClosedXML chart source edit is an expected package graph rewrite, commits through save/reopen, preserves source bytes, and reopens with the edited chart series source.',
			boundary:
				'Public chart-source commit evidence only; it does not prove full chart authoring, chart layout editing, arbitrary chart XML preservation, or byte-equivalent chart output.',
		},
		{
			evidenceId: 'chart-source-structural-edit-proof',
			kind: 'test',
			command:
				'bun test packages/engine/src/operations.test.ts packages/sdk/src/agent-workflow.test.ts -t "row and column shifts update chart source refs|moveRange rewrites chart source refs that reference the moved range|moveRange rejects partial chart source refs before mutation|commits structural row edits with shifted chart source refs after save and reopen|commits moved chart source refs after save and reopen" --timeout 30000',
			path: 'packages/engine/src/operations.test.ts; packages/engine/src/structural/formula-rewrite.ts; packages/engine/src/operations/structural-ops.ts; packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commit 0fc4370a proves row and column structural edits update local chart name/category/value source references, leave external 3D chart references untouched, mark rewritten chart parts as expected post-write package graph changes, and commit a public ClosedXML row insert through save and reopen with shifted chart source refs. Commit c06bba18 extends the same chart-source routing to moveRange by rebasing fully moved chart refs, failing closed on partial moved chart refs before mutation, and committing a public ClosedXML chart-source move through save and reopen.',
			boundary:
				'Chart-source structural edit evidence only; it does not prove full chart authoring, chart layout editing, structured-reference chart rewriting, external workbook reference rewriting, arbitrary chart XML preservation, or byte-equivalent chart output.',
		},
		{
			evidenceId: 'sparkline-structural-edit-proof',
			kind: 'test',
			command:
				'bun test packages/engine/src/operations.test.ts packages/sdk/src/advanced-filter-sparkline.test.ts -t "row and column shifts update sparkline source and location ranges|structural row edits shift sparkline ranges through SDK save and reopen" --timeout 30000',
			path: 'packages/engine/src/operations.test.ts; packages/engine/src/structural/sheet-topology.ts; packages/sdk/src/advanced-filter-sparkline.test.ts',
			acceptedScope:
				'Commit 8357e5de proves structural row/column edits shift sparkline group source ranges, location ranges, date-axis ranges, and individual sparkline refs, with SDK save/reopen coverage for a generated workbook containing advanced filters and sparklines.',
			boundary:
				'Sparkline structural edit evidence only; it does not prove Excel rendering parity, arbitrary sparkline OOXML preservation, unsupported sparkline settings, or public workbook generality.',
		},
		{
			evidenceId: 'public-calc-chain-formula-commit-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public calc-chain formula edits through save and reopen audits" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commit 31141dd4 proves public calc-chain formula edits discard stale calcChain package parts as expected, reopen cleanly, surface post-write formulaState with calcChainState absent, and preserve the source bytes.',
			boundary:
				'Public calc-chain formula commit evidence only; it does not prove Excel-fresh formula values, full calculation parity, formula engine completeness, or arbitrary calc-chain preservation.',
		},
		{
			evidenceId: 'public-hidden-sheet-topology-commit-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public hidden-sheet workbook views through save and reopen audits" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts; packages/sdk/src/index.ts; packages/sdk/src/index-exports.test.ts',
			acceptedScope:
				'Commit 74f63b73 proves public hidden-sheet workbook topology is included in post-write proof: safe edits reopen cleanly, compact commit results report hidden sheet names and sheet states, root SDK exports the topology types, and source bytes remain unchanged. Commit 033be30c adds reopened workbook-view details for activeTab and firstSheet and proves the public hidden-sheet workbook view survives save/reopen.',
			boundary:
				'Public hidden-sheet topology commit evidence only; it does not prove workbook protection security, visibility authorization, full workbook-view parity, or arbitrary topology editing support.',
		},
		{
			evidenceId: 'public-workbook-view-protection-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public workbook structure protection and view metadata through reopen audits|commits public workbook strong-hash protection with honest metadata reporting" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts; packages/sdk/src/index.ts; packages/sdk/src/index-exports.test.ts',
			acceptedScope:
				'Commit 033be30c proves a public Apache POI workbook with structure protection can be safely edited while post-write proof reports reopened workbook protection plus workbookViewDetails such as tabRatio, and root SDK exports the workbook-view proof type. Commit 1a620712 proves public strong-hash workbook protection is reported after save/reopen with algorithm name, spin count, hash/salt presence, and reported-not-validated verification while raw hash and salt values stay out of agent-facing proof JSON.',
			boundary:
				'Public workbook-view and workbook-protection reporting evidence only; it does not validate password correctness, disclose raw hash or salt values, enforce protection security, authorize protected structure edits, prove full workbook-view parity, or guarantee arbitrary topology preservation.',
		},
		{
			evidenceId: 'public-protected-range-hash-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public protected ranges with honest hash reporting after save and reopen" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts; packages/sdk/src/index.ts; packages/sdk/src/index-exports.test.ts',
			acceptedScope:
				'Commit d519c5a7 proves public protected-range metadata is included in post-write proof with password/hash/security-descriptor counts and per-range details after safe edit, save, and reopen; root SDK exports the protected-range security type.',
			boundary:
				'Public protected-range reporting evidence only; it does not validate password hashes, enforce protection security, authorize protected-range edits, or prove Excel protection equivalence.',
		},
		{
			evidenceId: 'query-table-topology-fail-closed-proof',
			kind: 'test',
			command:
				'bun test fixtures/corpus/external-refresh-contract.test.ts -t "fails closed on public query-table column topology edits without dirtying output" --timeout 60000',
			path: 'fixtures/corpus/external-refresh-contract.test.ts; packages/sdk/src/workbook.ts',
			acceptedScope:
				'Commit a477e2ed proves public query-table column-topology edits fail closed with unsupported journal evidence while preserving connection metadata, table metadata, package graph integrity, and byte preservation on save/reopen.',
			boundary:
				'Public query-table topology fail-closed evidence only; it does not prove query refresh execution, external data trust, arbitrary query-table editing, or broader cross-workbook refresh support.',
		},
		{
			evidenceId: 'cli-agent-facing-open-diagnostics-proof',
			kind: 'test',
			command:
				'bun test apps/cli/src/cli.test.ts apps/cli/src/file-errors.test.ts -t "custom UI callbacks|open-plan --json reports missing files|raw ENOENT noise|missing ops sidecar|structured missing workflow input guidance|structured missing workbook guidance|verification commands --json return structured missing input guidance|read surfaces --json return structured missing input guidance|proof and recovery commands --json return structured missing input guidance|calc and export --json return structured missing input guidance|direct mutation commands --json return structured missing input guidance before opening files|utility commands --json return structured missing input guidance|CLI JSON validation errors include command-specific guidance and fallback codes|formula edit commands --json return structured missing input guidance|unknown command --json returns a failure envelope|unknown inspect flag --json returns actionable retry details|doctor unsupported flag --json returns command flag details|read table --json exposes table metadata and paginated rows|trace --json reports missing cells with structured retry guidance|read name --json exposes parsed name metadata|read --json reports missing sheets with structured retry guidance|inspect --json reports missing sheets with structured retry guidance|agent-view --json reports missing sheets with structured retry guidance|find searches for values|ops --json exposes operation schemas with examples|inspect --json reports unknown details with structured retry guidance|formula show returns parsed formula info" --timeout 30000',
			path: 'apps/cli/src/index.ts; apps/cli/src/output/json.ts; apps/cli/src/commands/open-plan.ts; apps/cli/src/commands/inspect.ts; apps/cli/src/commands/read.ts; apps/cli/src/commands/agent-view.ts; apps/cli/src/commands/plan.ts; apps/cli/src/commands/commit.ts; apps/cli/src/commands/check.ts; apps/cli/src/commands/diff.ts; apps/cli/src/commands/lint.ts; apps/cli/src/commands/trace.ts; apps/cli/src/commands/dump.ts; apps/cli/src/commands/template-merge.ts; apps/cli/src/commands/repair-plan.ts; apps/cli/src/commands/calc.ts; apps/cli/src/commands/export.ts; apps/cli/src/commands/preview.ts; apps/cli/src/commands/write.ts; apps/cli/src/commands/create.ts; apps/cli/src/commands/find.ts; apps/cli/src/commands/formula.ts; apps/cli/src/commands/list.ts; apps/cli/src/cli.test.ts; apps/cli/src/file-errors.test.ts',
			acceptedScope:
				'Commit d837689e makes CLI inspect --detail active-content report generated RibbonX custom UI callbacks and makes CLI open-plan --json return retryable FILE_NOT_FOUND guidance for missing workbook paths; commit 9b155c7f extends missing-file guidance to non-JSON CLI output without raw ENOENT noise; commit 3d4d5374 reports missing plan/commit ops sidecar paths directly and leaves commit output absent; commit ff38e8f4 makes CLI plan/commit --json return structured retryable missing workflow input guidance for file/ops gaps; commit b81bd22e makes CLI open-plan/inspect --json return structured retryable missing workbook guidance; commit 7e31d148 makes CLI check/lint/trace/diff --json return structured retryable missing verification input guidance; commit 21c5a987 makes CLI read/agent-view --json return structured retryable missing read input guidance; commit 5ab3470e makes CLI dump/template-merge/repair-plan --json return structured retryable missing proof and recovery input guidance; commit 6e190776 makes CLI calc/export --json return structured retryable missing calc/export input guidance; commit 4a2abfdb makes CLI preview/write --json return structured retryable missing direct-edit input guidance before opening files; commit a736a539 makes CLI create/list/find/formula --json return structured retryable missing utility input guidance; commit c7455eeb makes legacy CLI string errors emitted with --json return coded retryable machine failures; commit 4b867889 makes CLI formula assist/set/fill --json return structured retryable missing formula-edit input guidance; commit 2a996840 makes CLI inspect/read --json invalid-argument failures include command-specific retry details and suggested fixes; commit 8826b61b extends command-specific JSON guidance to find --match, dump conflicting flags, template-merge data shape, and export format failures; commit 78e06818 makes CLI unknown command and unknown flag --json dispatch errors include retryable details, allowed commands/flags, and agent workflow guidance; commit f20db00b makes CLI read table misses and trace missing-cell failures return structured retryable guidance; commit 9cd51118 makes CLI read missing defined names and missing sheets return structured retryable guidance with available names/sheets; commit baf140bd makes CLI inspect missing-sheet failures return structured retryable guidance with available sheets; commit 82b2bed5 makes CLI agent-view missing-sheet failures return structured retryable guidance with available sheets; commit 83246775 makes CLI find missing-sheet failures return structured retryable guidance with available sheets; commit 42f881b5 makes CLI ops missing-operation lookup failures return structured retryable guidance with available operations; commit 315cd030 makes CLI inspect unknown-detail failures return structured retryable guidance with allowed detail values; commit 4b8b82b6 makes CLI formula unknown-subcommand failures return structured retryable guidance with allowed subcommands and the closest suggestion; commit 59d0d65c makes CLI formula-show missing-formula failures return structured retryable guidance with load context.',
			boundary:
				'CLI agent-facing diagnostics only; it does not prove public custom UI fixture coverage, file recovery, path discovery, Custom UI safety, active-content safety, or trust wording.',
		},
		{
			evidenceId: 'cli-agent-init-workflow-examples-proof',
			kind: 'test',
			command:
				'bun test apps/cli/src/cli.test.ts -t "agent-init prints the canonical agent workflow contract" --timeout 30000',
			path: 'apps/cli/src/commands/agent-init.ts; apps/cli/src/cli.test.ts',
			acceptedScope:
				'Commit 0d1b33f7 makes CLI agent-init surface runnable SDK, HTTP, and MCP safe-edit example commands alongside the canonical workflow contract, API endpoints, MCP tools, and safety defaults.',
			boundary:
				'CLI workflow-discovery evidence only; it does not prove package publication, arbitrary workbook safety, performance, external trust, or every SDK/API/MCP workflow.',
		},
		{
			evidenceId: 'api-open-workflow-reference-proof',
			kind: 'test',
			command:
				'bun test apps/api/src/server.test.ts -t "jsonFailureError wraps string failures|string API failures return coded JSON envelopes|missing workbook references|missing inputs|missing template data|before opening workbooks|trace reports missing target cells|/agent-workflow exposes the API safe edit contract" --timeout 30000',
			path: 'apps/api/src/server.ts; apps/api/src/response.ts; apps/api/src/server.test.ts',
			acceptedScope:
				'Commits 215d6e57, 7c1a9708, 8ce0fbe2, 43781bef, ea67f3b3, 6490a0e6, e090fe13, 8d12c141, 2cb02045, 7303b787, ac7d8006, and 091a4318 return structured retryable missing-workbook-reference errors for API plan/commit/open-plan/inspect/active-content/trust-report/package-graph/raw-part/visuals/pivots/dump/template-merge/read/agent-view/repair-plan/check/lint/trace/write/preview/calc/diff/export workflow requests instead of generic missing-file responses. Commit 346410a9 returns structured retryable missing-range errors for API read/agent-view ranges, commit 7303b787 returns structured retryable missing trace-cell errors for trace cells, commit 091a4318 returns structured retryable missing export-format errors, commit 2b884ed6 returns structured retryable missing template data errors, commit 542523a4 rejects invalid replay filters, Pivot materialization mode, read format, and unsupported export format before opening workbooks, commit 0a7d9d32 makes legacy string API failures return coded machine envelopes, commit ceb94425 makes API /trace missing target cells return structured retryable guidance, commit 25ca9b21 makes direct jsonFailureError string helper failures return coded machine envelopes, commit ec0b98e9 makes unsupported API routes return structured retryable guidance with supported route inventory, commit 98752c84 exposes a machine-readable API safe-edit workflow contract for open-plan, inspect, read, plan, commit, reopen-verify, and repair, commit 407499cd exposes runnable SDK/API/MCP workflow example commands in that API contract, and commit a8e15d9b adds trust-preflight with POST /trust-report to the API workflow contract.',
			boundary:
				'API request-shape diagnostics only; it does not prove file recovery, path discovery, source workbook existence, edit correctness, latency, or trust wording.',
		},
		{
			evidenceId: 'mcp-open-workflow-reference-proof',
			kind: 'test',
			command:
				'bun test apps/mcp/src/index.test.ts -t "missing workbook references|string MCP tool errors return coded JSON failures|ascend.read_table reports missing tables|ascend.agent_workflow exposes machine-readable safe edit guidance" --timeout 30000',
			path: 'apps/mcp/src/index.ts; apps/mcp/src/response.ts; apps/mcp/src/index.test.ts',
			acceptedScope:
				'Commit da273900 returns structured retryable missing-workbook-reference errors for MCP commit requests without file or planHandle instead of generic missing-file responses. Commit daa3ecb5 makes legacy MCP string tool errors return coded retryable machine failures. Commit f16085b3 makes MCP ascend.read_table missing-table failures return structured retryable guidance with available table names. Commit 77695508 exposes a machine-readable MCP safe-edit workflow contract for open-plan, inspect/read, plan, commit, reopen-verify, and repair. Commit 407499cd exposes runnable SDK/API/MCP workflow example commands in that MCP contract. Commit a8e15d9b adds trust-preflight with ascend.trust_report to the MCP workflow contract.',
			boundary:
				'MCP request-shape diagnostics only; it does not prove file recovery, path discovery, source workbook existence, edit correctness, latency, or trust wording.',
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
				'Local RC packageability gate for SDK/CLI/API/MCP tarballs and installed workbook proof. Commit 0931685e fixes bundled SDK agent docs resolution from file URLs, preserving installed docs-search smoke coverage; commit f1f76e36 adds installed CLI agent-init plus API/MCP agent-workflow discovery smoke checks; commit 37794635 packages runnable HTTP/MCP workflow examples and extends installed SDK smoke coverage for bundled example discovery; commit 0d1b33f7 extends installed app smoke coverage for CLI agent-init runnable example commands; commit 407499cd extends installed app smoke coverage for API/MCP workflow example commands; commit a09660be updates installed app smoke coverage to assert root-level safe-edit example commands; commit f3347e17 extends installed app smoke coverage for example proof context; commit 58a7f579 exposes workflow examples and proof context in packaged MCP resource/app smoke output.',
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
				'Tracked corpus scan for package features, public unknown-path coverage, and the remaining generated signature edge case.',
			boundary: 'Tracked corpus evidence only; generated signature topology remains owner-gated.',
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
			evidenceId: 'query-table-refresh-agent-commit-proof',
			kind: 'test',
			command:
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public query-table refresh metadata edits through save and reopen" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commit 868add46 proves an agent commit can edit public query-table refresh metadata with required preservedQueryTable approval, save, reopen, retain inspectable refresh metadata, keep source bytes unchanged, and treat the intentionally rewritten query-table part as expected post-write package graph evidence.',
			boundary:
				'Public query-table refresh metadata edit evidence only; it does not execute queries, refresh external data, prove arbitrary query-table editing, or authorize external data trust wording.',
		},
		{
			evidenceId: 'query-table-post-write-data-connections-proof',
			kind: 'test',
			command:
				'bun test fixtures/corpus/external-refresh-contract.test.ts -t "commit proof reports reopened public query-table connection metadata" --timeout 30000',
			path: 'packages/core/src/connection.ts; packages/io-xlsx/src/reader/connections.ts; packages/sdk/src/agent-workflow.ts; packages/sdk/src/read-view.ts; packages/sdk/src/types.ts; fixtures/corpus/external-refresh-contract.test.ts',
			acceptedScope:
				'Commit caa08959 makes SDK post-write verification report reopened workbook/query-table connection metadata after an approved public query-table refresh metadata edit, with safe-open blocker coverage for the dangling-thumbnail query-table fixture. Commit 62566e09 expands reader, SDK inspect/refresh metadata, and post-write proof to report public connection type, description, deleted/background/keepAlive/interval flags, source file, command text, and a hasConnectionString boolean without exposing connection-string contents.',
			boundary:
				'Public post-write data-connection reporting evidence only; it does not execute connections, validate external data freshness, expose connection-string secrets, prove credential safety, or prove arbitrary query-table editing.',
		},
		{
			evidenceId: 'workbook-connection-scheduling-commit-proof',
			kind: 'test',
			command:
				'bun test packages/engine/src/operations.test.ts packages/io-xlsx/src/writer/writer.test.ts packages/sdk/src/ops-schema.test.ts packages/sdk/src/agent-workflow.test.ts -t "setConnectionRefresh|updates XML-legal single-quoted connection attributes without duplicating them|setConnectionRefresh is exposed with connection refresh guidance|commits public workbook connection scheduling edits through save and reopen" --timeout 30000',
			path: 'packages/schema/src/operations.ts; packages/engine/src/operations/connection-ops.ts; packages/io-xlsx/src/writer/connection.ts; packages/io-xlsx/src/writer/index.ts; packages/sdk/src/ops.ts; packages/sdk/src/agent-workflow.test.ts',
			acceptedScope:
				'Commit 1cb093fe proves public workbook connection scheduling metadata edits for backgroundRefresh, keepAlive, refreshInterval, refreshOnLoad, saveData, and refreshedVersion validate in operation schemas, persist into OOXML, save, reopen, and report in postWrite.dataConnections without executing the connection.',
			boundary:
				'Public workbook-connection scheduling edit evidence only; it does not execute connections, prove external-data freshness, validate credentials, support arbitrary connection authoring, or allow workbook-only scheduling fields on query-table parts.',
		},
		{
			evidenceId: 'external-link-source-binding-proof',
			kind: 'test',
			command:
				'bun test packages/io-xlsx/src/reader/external-links.test.ts packages/io-xlsx/src/package-graph.test.ts packages/io-xlsx/src/package-graph-fidelity.test.ts -t "parses DDE and OLE external link source metadata|inventories OLE external link relationship binding across save and reopen|normalizes content types, relationship identity, owners, and feature families|uses OPC source, id, target, and TargetMode fields for adversarial relationship drift|classifies external workbook path relationships as external-link metadata" --timeout 60000',
			path: 'packages/core/src/workbook.ts; packages/io-xlsx/src/reader/external-links.ts; packages/io-xlsx/src/reader/external-links.test.ts; packages/io-xlsx/src/reader/relationships.ts; packages/io-xlsx/src/package-graph.ts; packages/io-xlsx/src/package-graph.test.ts; packages/io-xlsx/src/package-graph-fidelity.test.ts; packages/sdk/src/agent-workflow.ts; packages/sdk/src/types.ts',
			acceptedScope:
				'Commits 8576a860, ab5d635a, 4a61e9af, and 66111c9e prove external-link source bindings are inventoried for DDE/OLE/externalBook-style relationships, externalLinkPath and nonstandard external-link parts are classified as preservedExternalLink, hyperlink relationships are classified as preservedHyperlink, and relationship source/id/target/TargetMode drift remains audited.',
			boundary:
				'External-link and hyperlink package relationship evidence only; it does not execute links, fetch external workbooks, prove linked-data freshness, validate linked workbook contents, repair arbitrary broken links, or make external references safe to trust.',
		},
		{
			evidenceId: 'opaque-relationship-classification-proof',
			kind: 'test',
			command:
				'bun test packages/io-xlsx/src/package-graph.test.ts -t "normalizes content types, relationship identity, owners, and feature families" --timeout 30000',
			path: 'packages/io-xlsx/src/package-graph.ts; packages/io-xlsx/src/package-graph.test.ts; packages/io-xlsx/src/reader/relationships.ts',
			acceptedScope:
				'Commits 1410d809, 8e351092, 75698a98, 346da8b1, and 844336cb classify package relationships and parts for active content/control/embedding, Power Query/Data Model, chart style/color, custom XML/VBA/signature, calc-chain, and sheet metadata surfaces with owner scopes and preservation policies.',
			boundary:
				'Package relationship classification evidence only; it does not execute active content, refresh Power Query, understand Data Model semantics, edit chart styles/colors, prove calc metadata freshness, or guarantee arbitrary unknown-part recovery.',
		},
		{
			evidenceId: 'public-formula-cache-post-write-proof',
			kind: 'test',
			command:
				'bun test fixtures/corpus/formula-binding-contract.test.ts -t "commit proof reports missing public formula caches after save and reopen" --timeout 30000',
			path: 'packages/sdk/src/agent-workflow.ts; packages/sdk/src/index.ts; packages/sdk/src/index-exports.test.ts; fixtures/corpus/formula-binding-contract.test.ts',
			acceptedScope:
				'Commit f7338c91 proves a public ClosedXML formula workbook can be committed safely while post-write formulaState reports reopened formula cells, missing cached formula values, missing-cache locations, calcChain presence, and warnings instead of claiming fresh recalculation. Commit 019f457e adds cached formula value locations and cached formula value samples to post-write formula proof and exports the SDK proof type.',
			boundary:
				'Public formula-cache reporting evidence only; it does not prove Excel recalculation equivalence, fresh cached values, full formula parity, or formula oracle coverage.',
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
	const doNotPromoteYet = qssMatrix.archivedResearchNotes.map(releaseDecisionDoNotPromoteItem)
	const rows = qssMatrix.rows.map((row) => {
		const artifact = artifacts.find((candidate) => candidate.name === row.artifact)
		const handoff = readiness.implementationHandoffs.find(
			(candidate) => candidate.artifact === row.artifact,
		)
		const allowedWording = releaseDecisionAllowedWording(row.artifact)
		return {
			rank: row.rank,
			artifact: row.artifact,
			claimWordingAllowedToday: allowedWording,
			evidenceWeHave: row.acceptedEvidence.map((item) => ({ ...item })),
			evidenceMissing: [...row.missingEvidence],
			qssContrast: releaseDecisionQssContrast(row),
			allowedWording,
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
	})
	const topClaimOwnerActionQueue = releaseDecisionTopClaimOwnerActionQueue(rows)
	const blockedOwnerActionQueue = releaseDecisionBlockedOwnerActionQueue(doNotPromoteYet)
	const benchmarkCorpusOwnerActionQueue = releaseDecisionBenchmarkCorpusOwnerActionQueue(
		topClaimOwnerActionQueue,
		blockedOwnerActionQueue,
	)
	const implementationReadyOwnerActionQueue = releaseDecisionImplementationReadyOwnerActionQueue(
		topClaimOwnerActionQueue,
		blockedOwnerActionQueue,
	)
	const claimDowngradeOwnerActionQueue =
		releaseDecisionClaimDowngradeOwnerActionQueue(blockedOwnerActionQueue)
	const benchmarkCorpusRunContractCoverage = releaseDecisionBenchmarkCorpusRunContractCoverage(
		benchmarkCorpusOwnerActionQueue,
	)
	return {
		status: 'top-two-only',
		releaseGate: readiness.releaseGate,
		headlineClaimsAllowed: readiness.headlineClaimsAllowed,
		implementationSurfacePromotionAllowed: readiness.implementationSurfacePromotionAllowed,
		missingRequirementCount: readiness.missingRequirementCount,
		rows,
		topClaimOwnerActionQueue,
		doNotPromoteYet,
		doNotPromoteDispositionSummary: releaseDecisionDispositionSummary(doNotPromoteYet),
		releaseWordingDecisionSummary: releaseWordingDecisionSummary(rows, doNotPromoteYet),
		todayCommitClaimMatrix: todayCommitClaimMatrix(),
		claimDecisionContractCoverage: releaseDecisionClaimDecisionContractCoverage(
			rows,
			doNotPromoteYet,
		),
		blockedOwnerActionQueue,
		benchmarkCorpusOwnerActionQueue,
		implementationReadyOwnerActionQueue,
		claimDowngradeOwnerActionQueue,
		benchmarkCorpusRunContractCoverage,
		ownerActionQueueCoverage: releaseDecisionOwnerActionQueueCoverage(
			topClaimOwnerActionQueue,
			blockedOwnerActionQueue,
			benchmarkCorpusOwnerActionQueue,
			implementationReadyOwnerActionQueue,
			claimDowngradeOwnerActionQueue,
		),
		ownerActionExecutionContractCoverage: releaseDecisionOwnerActionExecutionContractCoverage(
			benchmarkCorpusOwnerActionQueue,
			implementationReadyOwnerActionQueue,
			claimDowngradeOwnerActionQueue,
		),
		boundary:
			'Top-two release-decision artifact for claim stewardship. It is derived from committed release proof gates and must not be treated as a product surface, benchmark threshold, signed provenance, or owner approval.',
	}
}

function todayCommitClaimMatrix(): readonly ReleaseProofTodayCommitClaimMatrixRow[] {
	return [
		{
			claimArea: 'safe-agent-workflows',
			commits: [
				'de45eb83',
				'37794635',
				'a09660be',
				'407499cd',
				'f3347e17',
				'58a7f579',
				'5981764c',
				'5915794f',
				'f8d63593',
				'cc689bcc',
				'56cd4aa0',
				'3d630232',
				'5028438e',
				'1ed2be29',
				'a5fa3006',
				'7afcd630',
				'adc9c1e1',
				'7e7183df',
				'a8e15d9b',
				'223a1ec7',
				'868add46',
				'caa08959',
				'62566e09',
				'1cb093fe',
				'8576a860',
				'ab5d635a',
				'4a61e9af',
				'66111c9e',
				'91dabea8',
				'4d272f77',
				'1eaf28ff',
				'0fc4370a',
				'c06bba18',
				'8357e5de',
				'31141dd4',
				'f7338c91',
				'019f457e',
				'a2960803',
				'74f63b73',
				'033be30c',
				'1a620712',
				'd519c5a7',
				'62f45cb5',
				'b93f3493',
			],
			releaseOrSotaClaimBecameMoreTrue:
				'Ascend is more credible as an agent-native spreadsheet runtime because SDK, installed-SDK, CLI, HTTP API, MCP, root-package, public query-table/data-connection reporting, public workbook-connection scheduling edits, external-link source-binding/package classification, public chart/visual, chart-source structural row/column/range-move edit proof, sparkline structural edit proof, public calc-chain/formula-cache, public hidden-sheet/workbook-view, public workbook-protection, protected-range examples, and generated safe-edit post-write proof fields expose runnable open-plan/trust/inspect/plan/commit/reopen/verify workflows plus shared proof-bundle and post-write proof context.',
			evidenceProvesIt: [
				'bun test examples/agent-safe-edit.test.ts --timeout 30000',
				'bun test examples/agent-safe-edit-mcp.test.ts --timeout 30000',
				'bun test examples/root-scripts.test.ts --timeout 30000',
				'bun test apps/cli/src/cli.test.ts apps/api/src/server.test.ts apps/mcp/src/index.test.ts -t "agent-init prints the canonical agent workflow contract|/agent-workflow exposes the API safe edit contract|ascend.agent_workflow exposes machine-readable safe edit guidance" --timeout 30000',
				'bun test apps/api/api.test.ts -t "operations and capabilities endpoints expose agent schemas" --timeout 30000',
				'bun test examples/package-install-safe-edit.test.ts --timeout 30000',
				'bun test apps/cli/src/cli.test.ts -t "example-safe-edit runs the packaged inspect plan commit reopen verify workflow" --timeout 30000',
				'bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow" --timeout 30000',
				'bun test apps/api/src/server.test.ts -t "dump emits replayable operation batches" --timeout 30000',
				'bun test apps/mcp/src/index.test.ts -t "ascend.commit accepts prepared plan handles" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "workflow proof summary explains changed cells and safety gates" --timeout 30000',
				'bun test examples/package-scripts.test.ts --timeout 30000',
				'bun test examples/root-scripts.test.ts --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public query-table refresh metadata edits through save and reopen" --timeout 30000',
				'bun test fixtures/corpus/external-refresh-contract.test.ts -t "commit proof reports reopened public query-table connection metadata" --timeout 30000',
				'bun test packages/io-xlsx/src/reader/connections.test.ts packages/sdk/src/connection-inventory.test.ts fixtures/corpus/external-refresh-contract.test.ts -t "connection part inventory|connection SDK inventory|reports and preserves public query-table refresh surfaces without executing them|commit proof reports reopened public query-table connection metadata" --timeout 60000',
				'bun test packages/engine/src/operations.test.ts packages/io-xlsx/src/writer/writer.test.ts packages/sdk/src/ops-schema.test.ts packages/sdk/src/agent-workflow.test.ts -t "setConnectionRefresh|updates XML-legal single-quoted connection attributes without duplicating them|setConnectionRefresh is exposed with connection refresh guidance|commits public workbook connection scheduling edits through save and reopen" --timeout 30000',
				'bun test packages/io-xlsx/src/reader/external-links.test.ts packages/io-xlsx/src/package-graph.test.ts packages/io-xlsx/src/package-graph-fidelity.test.ts -t "parses DDE and OLE external link source metadata|inventories OLE external link relationship binding across save and reopen|normalizes content types, relationship identity, owners, and feature families|uses OPC source, id, target, and TargetMode fields for adversarial relationship drift|classifies external workbook path relationships as external-link metadata" --timeout 60000',
				'bun test apps/cli/src/cli.test.ts -t "agent-init prints the canonical agent workflow contract" --timeout 30000',
				'bun test apps/api/src/server.test.ts -t "/agent-workflow exposes the API safe edit contract" --timeout 30000',
				'bun test apps/mcp/src/index.test.ts -t "ascend.agent_workflow exposes machine-readable safe edit guidance|agent resources return canonical workflow context" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public chart source edits through save and reopen audits" --timeout 30000',
				'bun test packages/engine/src/operations.test.ts packages/sdk/src/agent-workflow.test.ts -t "row and column shifts update chart source refs|moveRange rewrites chart source refs that reference the moved range|moveRange rejects partial chart source refs before mutation|commits structural row edits with shifted chart source refs after save and reopen|commits moved chart source refs after save and reopen" --timeout 30000',
				'bun test packages/engine/src/operations.test.ts packages/sdk/src/advanced-filter-sparkline.test.ts -t "row and column shifts update sparkline source and location ranges|structural row edits shift sparkline ranges through SDK save and reopen" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public calc-chain formula edits through save and reopen audits" --timeout 30000',
				'bun test fixtures/corpus/formula-binding-contract.test.ts -t "commit proof reports missing public formula caches after save and reopen" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public hidden-sheet workbook views through save and reopen audits" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public workbook structure protection and view metadata through reopen audits" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public workbook strong-hash protection with honest metadata reporting" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public protected ranges with honest hash reporting after save and reopen" --timeout 30000',
				'bun test examples/agent-safe-edit.test.ts examples/agent-safe-edit-http.test.ts examples/agent-safe-edit-mcp.test.ts examples/package-scripts.test.ts examples/root-scripts.test.ts --timeout 30000',
				'fixtures/benchmarks/release-proof-index.test.ts pins workflow example evidence, proof context, and RC-gate scope.',
				'Blocked adjacent API package sweep: bun test apps/api/api.test.ts --timeout 30000 currently fails outside workflow discovery because the export-format assertion expects "Unsupported format" while the current response is "Unsupported export format: weird".',
			],
			allowedWording:
				'Ascend provides local runnable SDK, installed-SDK package-bin, CLI, API, and MCP safe-edit workflow examples with root commands, explicit trust-preflight steps, shared SDK proof-summary gates, workflow discovery that points agents to postWrite.dataConnections, postWrite.formulaState, and postWrite.visuals, generated safe-edit example output that surfaces reopened-output post-write proof fields for data connections, formula state, security, and visuals, public query-table refresh/data-connection metadata including source file, command text, refresh flags, and connection-string presence without secret disclosure, workbook connection scheduling metadata edits for backgroundRefresh, keepAlive, refreshInterval, refreshOnLoad, saveData, and refreshedVersion without connection execution, external-link source binding and package relationship classification for inspection/preservation, chart-source/visual object reporting, local chart source refs shifted or rebased by structural row/column/range-move edits with reopened public-workbook proof, local sparkline source/location refs shifted by structural row/column edits with SDK save/reopen proof, calc-chain/formula-cache reporting with cached and missing value locations, hidden-sheet/workbook-view topology, workbook structure-protection and strong-hash metadata reporting without raw hash/salt disclosure, protected-range metadata commit proof, and machine-readable proof context for generated/public workbooks.',
			forbiddenWording: [
				'Do not claim arbitrary workbook safety, package publication, registry download proof, hosted service readiness, public-workbook generality, connection execution, arbitrary connection authoring, credential safety, external-data freshness, external-link execution, linked-workbook fetch/validation, arbitrary broken-link repair, fresh recalculation, full drawing/layout parity, full chart authoring, chart layout editing, full sparkline rendering or arbitrary sparkline OOXML preservation, structured-reference chart rewriting, external workbook chart-reference rewriting, complete workbook-view parity, password correctness validation, raw protection hash/salt disclosure, protection enforcement/security validation, complete API package health, or complete workflow observability from the local generated/public-workbook examples.',
			],
			ownerLoop: 'release',
			nextOwnerAction:
				'Release-code owner stops adding discovery-only workflow affordances; next accepted work must either fix the unrelated API export-format test expectation, run package publication/install smoke in a clean release environment, or add one public inspect/plan/commit/reopen/diff/audit/trace workflow that changes an agent repair decision.',
			boundary:
				'Compact current-commit claim row only; evidence is local workflow proof, not a broad product or publication claim.',
		},
		{
			claimArea: 'formula-calc-behavior',
			commits: ['f6a71088', '104c38c0', '64d82251', 'ba8ea5fb', '0604c9f3'],
			releaseOrSotaClaimBecameMoreTrue:
				'Ascend formula/calc behavior is more credible for dynamic-array-style workflows because common IS predicates, ERROR.TYPE, SWITCH, and IFS now map over range or array operands and spill through focused conditional contexts instead of collapsing to a top-left scalar.',
			evidenceProvesIt: [
				'bun test packages/engine/src/calc.test.ts packages/formulas/src/functions/functions.test.ts -t "error predicates map over arrays inside IF conditions|common IS predicates spill boolean masks for range operands|ERROR.TYPE spills error codes for range operands|IS predicates do not coerce text as numbers|SWITCH spills results for array expressions|SWITCH array expressions preserve lazy branch evaluation|IFS spills results for array conditions|IFS array conditions preserve selected array results and lazy branch evaluation" --timeout 30000',
			],
			allowedWording:
				'Ascend has local regression proof for array/range mapping of ISERROR, ISERR, ISNA, ISBLANK, ISTEXT, ISLOGICAL, ISNONTEXT, ISEVEN, ISODD, ERROR.TYPE, SWITCH, and IFS in focused formula-engine cases, including spilled IF/IFERROR/SWITCH/IFS masks and lazy branch handling.',
			forbiddenWording: [
				'Do not claim Excel-compatible formulas, dynamic-array completeness, full INFO-function parity, external-oracle parity, fresh cached values, or QSS/SOTA formula superiority from these local unit regressions.',
			],
			ownerLoop: 'correctness',
			nextOwnerAction:
				'Formula/Calc owner should promote these cases into the oracle corpus by adding LibreOffice/Excel-ground-truth workbook rows or accepted-mismatch retirements, then rerun formula-corpus-correctness with --max-mismatches 0 before any parity wording changes.',
			boundary:
				'Compact current-commit claim row only; evidence is local formula-engine regression coverage, not an external Excel oracle or broad formula-compatibility claim.',
		},
		{
			claimArea: 'signed-encrypted-macro-handling',
			commits: [
				'7ff308c9',
				'e5d0ee17',
				'ca3c3296',
				'2401fa0a',
				'b7e8eccc',
				'a004fb4a',
				'9df35fd6',
				'fef9294f',
			],
			releaseOrSotaClaimBecameMoreTrue:
				'Ascend is stricter and more explainable on high-risk workbook handling: encrypted package inspection uses decrypted session bytes, signed/VBA signature edits fail closed, streams and values-only text exports require explicit high-risk approvals, and public shape macros remain blocked active content with drawing-object proof after safe edit.',
			evidenceProvesIt: [
				'bun test packages/sdk/src/sdk.test.ts -t "rawPackagePart inspects decrypted package parts after password open" --timeout 30000',
				'bun test packages/sdk/src/sdk.test.ts -t "VBA project signatures fail closed while approved exports preserve macros" --timeout 30000',
				'bun test packages/sdk/src/sdk.test.ts -t "high-risk workbook streams require the same explicit export approvals" --timeout 30000',
				'bun test packages/sdk/src/sdk.test.ts -t "high-risk workbook text saves require the same explicit export approvals" --timeout 30000',
				'bun test packages/sdk/src/agent-workflow.test.ts -t "commits public shape macro drawings as blocked active content after safe edit" --timeout 30000',
			],
			allowedWording:
				'Ascend has local proof that encrypted, signed, macro, signed-macro, and public shape-macro workflows fail closed or report blocked active content with drawing-object identity unless the exact loss/decryption/signature approvals are explicit.',
			forbiddenWording: [
				'Do not claim macro safety, malware scanning, re-encryption, signature preservation, re-signing, signer trust, safe execution of active content, or full drawing/layout parity.',
			],
			ownerLoop: 'correctness',
			nextOwnerAction:
				'Excel compatibility owner should stop adding synthetic high-risk policy cases unless they close a public signed/macro fixture policy gap; next work is public fixture/provenance coverage or an explicit do-not-promote decision for signed fixture wording.',
			boundary:
				'Compact current-commit claim row only; synthetic high-risk topology proof remains below broad trust or safety wording.',
		},
		{
			claimArea: 'write-performance',
			commits: [
				'67b900ed',
				'e22eb86a',
				'bd162937',
				'0d0c9632',
				'905ecb5e',
				'c297ba4c',
				'27af69d4',
				'2029dfab',
				'9cab723f',
				'7d61a2ef',
				'183e7ebf',
				'eca32509',
			],
			releaseOrSotaClaimBecameMoreTrue:
				'Dense-values, string-heavy, styles-heavy, formula-heavy, table-heavy, and feature-rich generated XLSX write rows now have scoped current comparable evidence, but broad XLSX write/SOTA/QSS speed wording remains downgraded.',
			evidenceProvesIt: [
				'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts --timeout 30000',
				'performance owner artifact in release-proof-index names the plain-text baseline, string-heavy baseline, 0d0c9632 string-heavy optimization proof, 905ecb5e styles-heavy write baseline win, c297ba4c dense-values current repeat-15 comparable win, 27af69d4 and 9cab723f string-heavy current repeat-15 comparable wins, 2029dfab current styles-heavy repeat-15 comparable win, 7d61a2ef formula-heavy current comparable win, 183e7ebf table-heavy current comparable win, and eca32509 feature-rich XlsxWriter win plus OpenPyXL not-comparable boundary as bounded evidence.',
			],
			allowedWording:
				'Ascend has bounded local evidence of dense-values, string-heavy, styles-heavy, formula-heavy, and table-heavy generated write rows where focused comparable external reruns favor Ascend, plus feature-rich evidence against semantically comparable XlsxWriter with OpenPyXL explicitly not comparable; treat them as scoped generated-write evidence, not broad speed leadership.',
			forbiddenWording: [
				'Do not claim fastest XLSX writer, broad write SOTA, QSS-leapfrog speed, smallest XLSX output, byte/order equivalence against every writer, or external workflow dominance from scoped generated rows.',
			],
			ownerLoop: 'performance',
			nextOwnerAction:
				'Benchmark owner stops production optimization from these rows unless the next block is a clean multi-workload xlsx-write-sota gate, an explicit external-baseline blocker, or a named public workflow loss that identifies a production cost center.',
			boundary:
				'Compact current-commit claim row only; scoped generated write rows cannot promote broad performance claims.',
		},
		{
			claimArea: 'external-baselines',
			commits: ['fa3a13dc', 'b6925afe', 'bb31bebe', 'f8846cf8', '9ddfff91', 'df349629', '1908f3f5'],
			releaseOrSotaClaimBecameMoreTrue:
				'External read baselines are more comparable across current selected-sheet, metadata-only, FastXLSX, SheetJS, ExcelJS, Calamine, and ClosedXML boundaries, but they mostly sharpen downgrade decisions rather than promote speed leadership.',
			evidenceProvesIt: [
				'fixtures/benchmarks/release-proof-index.ts performance boundary packet records same-lane selected-sheet/metadata-only evidence, df349629 current selected-sheet same-lane read proof, 1908f3f5 current metadata-only Calamine loss/kill decision, FastXLSX scoped rows, and current full-profile/merged scoreboard status.',
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --performance-boundary-json',
			],
			allowedWording:
				'Ascend has bounded local comparable external-baseline evidence for selected-sheet, metadata-only, and FastXLSX rows with unsupported/mismatch/timing-boundary disclosures.',
			forbiddenWording: [
				'Do not claim fastest XLSX reader, broad SOTA read speed, QSS speed leadership, or a Calamine metadata-only win.',
			],
			ownerLoop: 'performance',
			nextOwnerAction:
				'Benchmark owner resolves one explicit blocker row: ClosedXML coverage, feature-rich SheetJS/Calamine semantic policy, remaining unsupported selected-sheet/metadata-only competitors, FastXLSX environment coverage, same-timing SDK selected-sheet open-only row, or a profiling-named public workflow cost center.',
			boundary:
				'Compact current-commit claim row only; external baseline evidence is scoped and mostly blocks broad speed claims.',
		},
		{
			claimArea: 'research-proof-surface',
			commits: ['62f45cb5'],
			releaseOrSotaClaimBecameMoreTrue:
				'Research material is more release-safe because the dirty research/tmp surface is routed into accepted evidence, blocked claims, owner-ready implementation tasks, or archive/defer material instead of being cited directly.',
			evidenceProvesIt: [
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
				'bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 120000',
			],
			allowedWording:
				'Raw research and tmp material is classified as release routing evidence only; cite the proof-index packets, not the raw notes.',
			forbiddenWording: [
				'Do not cite research/ or tmp/ paths as product, correctness, performance, QSS, or release evidence unless a future commit folds one finding into the proof index with commands and owner action.',
			],
			ownerLoop: 'product',
			nextOwnerAction:
				'Product/release owner reviews classifiedEntries, resolves local Excel corpus and loop-manager state as active owner blockers, and archives broad research notes unless they change a claim or implementation target.',
			boundary:
				'Compact current-commit claim row only; research classification is not approval to publish or cite raw research.',
		},
	]
}

function releaseDecisionDispositionSummary(
	items: readonly ReleaseProofReleaseDecisionDoNotPromoteItem[],
): ReleaseProofReleaseDecisionDispositionSummary {
	return {
		implementationReadyBlockerNames: items
			.filter((item) => item.workBlockDisposition === 'implementation-ready-blocker')
			.map((item) => item.name),
		benchmarkCorpusBlockerNames: items
			.filter((item) => item.workBlockDisposition === 'benchmark-corpus-blocker')
			.map((item) => item.name),
		claimDowngradeDoNotPromoteNames: items
			.filter((item) => item.workBlockDisposition === 'claim-downgrade-do-not-promote')
			.map((item) => item.name),
		boundary:
			'Routing summary for blocked claims only. It groups do-not-promote decisions by the next work block type without changing claim wording, owner approval, or release gates.',
	}
}

function releaseWordingDecisionSummary(
	rows: readonly ReleaseProofReleaseDecisionBoardRow[],
	doNotPromoteYet: readonly ReleaseProofReleaseDecisionDoNotPromoteItem[],
): ReleaseProofReleaseWordingDecisionSummary {
	return {
		status: 'headline-claims-blocked-local-wording-only',
		headlineClaimsAllowed: false,
		localAllowedClaimNames: rows.map((row) => row.artifact),
		doNotPromoteClaimNames: doNotPromoteYet.map((item) => item.name),
		localAllowedWordingByClaim: Object.fromEntries(
			rows.map((row) => [row.artifact, row.allowedWording]),
		) as Readonly<Record<ReleaseProofIndexArtifactName, string>>,
		doNotPromoteAllowedWordingByClaim: Object.fromEntries(
			doNotPromoteYet.map((item) => [item.name, item.allowedWording]),
		) as Readonly<Record<ReleaseProofReleaseDecisionDoNotPromoteItem['name'], string>>,
		forbiddenWordingByClaim: Object.fromEntries([
			...rows.map((row) => [row.artifact, [...row.forbiddenWording]] as const),
			...doNotPromoteYet.map((item) => [item.name, [...item.forbiddenWording]] as const),
		]) as Readonly<
			Record<
				ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name'],
				readonly string[]
			>
		>,
		boundary:
			'Release wording decision summary only. It is for copy review and owner routing; headline claims remain blocked until release gates and owner approvals pass.',
	}
}

type ReleaseProofClaimDecisionContractSubject = {
	readonly key: string
	readonly evidenceWeHaveCount: number
	readonly evidenceMissingCount: number
	readonly qssContrast: readonly string[]
	readonly allowedWording: string
	readonly forbiddenWordingCount: number
	readonly nextOwnerActionCount: number
}

function releaseDecisionClaimDecisionContractCoverage(
	rows: readonly ReleaseProofReleaseDecisionBoardRow[],
	doNotPromoteYet: readonly ReleaseProofReleaseDecisionDoNotPromoteItem[],
): ReleaseProofClaimDecisionContractCoverage {
	const decisions: readonly ReleaseProofClaimDecisionContractSubject[] = [
		...rows.map((row) => ({
			key: `top-claim:${row.artifact}`,
			evidenceWeHaveCount: row.evidenceWeHave.length,
			evidenceMissingCount: row.evidenceMissing.length,
			qssContrast: row.qssContrast,
			allowedWording: row.allowedWording,
			forbiddenWordingCount: row.forbiddenWording.length,
			nextOwnerActionCount: row.nextOwnerActions.length,
		})),
		...doNotPromoteYet.map((item) => ({
			key: `do-not-promote:${item.name}`,
			evidenceWeHaveCount: item.evidenceWeHave.length,
			evidenceMissingCount: item.evidenceMissing.length,
			qssContrast: item.qssContrast,
			allowedWording: item.allowedWording,
			forbiddenWordingCount: item.forbiddenWording.length,
			nextOwnerActionCount: item.nextOwnerAction.trim().length > 0 ? 1 : 0,
		})),
	]
	const missingEvidenceWeHaveKeys = decisions
		.filter((decision) => decision.evidenceWeHaveCount === 0)
		.map((decision) => decision.key)
	const missingEvidenceMissingKeys = decisions
		.filter((decision) => decision.evidenceMissingCount === 0)
		.map((decision) => decision.key)
	const missingQssContrastKeys = decisions
		.filter(
			(decision) =>
				decision.qssContrast.length === 0 ||
				decision.qssContrast.some((item) =>
					item.includes(
						'QSS contrast is blocked until this diagnostic evidence changes a top-two release claim.',
					),
				),
		)
		.map((decision) => decision.key)
	const missingAllowedWordingKeys = decisions
		.filter((decision) => decision.allowedWording.trim().length === 0)
		.map((decision) => decision.key)
	const missingForbiddenWordingKeys = decisions
		.filter((decision) => decision.forbiddenWordingCount === 0)
		.map((decision) => decision.key)
	const missingNextOwnerActionKeys = decisions
		.filter((decision) => decision.nextOwnerActionCount === 0)
		.map((decision) => decision.key)
	const status =
		missingEvidenceWeHaveKeys.length === 0 &&
		missingEvidenceMissingKeys.length === 0 &&
		missingQssContrastKeys.length === 0 &&
		missingAllowedWordingKeys.length === 0 &&
		missingForbiddenWordingKeys.length === 0 &&
		missingNextOwnerActionKeys.length === 0
			? 'all-release-claim-decisions-self-contained'
			: 'claim-decision-contract-gap'
	return {
		status,
		decisionCount: decisions.length,
		topClaimDecisionCount: rows.length,
		doNotPromoteDecisionCount: doNotPromoteYet.length,
		missingEvidenceWeHaveKeys,
		missingEvidenceMissingKeys,
		missingQssContrastKeys,
		missingAllowedWordingKeys,
		missingForbiddenWordingKeys,
		missingNextOwnerActionKeys,
		boundary:
			'Claim decision contract coverage only. It proves every release decision or do-not-promote claim names evidence we have, evidence missing, QSS contrast, allowed wording, forbidden wording, and next owner action without approving release claims.',
	}
}

function releaseDecisionTopClaimOwnerActionQueue(
	rows: readonly ReleaseProofReleaseDecisionBoardRow[],
): readonly ReleaseProofTopClaimOwnerAction[] {
	return rows.flatMap((row) =>
		row.nextOwnerActions.map((action) => ({
			artifact: row.artifact,
			claim: row.claimWordingAllowedToday,
			ownerLoop: action.ownerLoop,
			requirementId: action.requirementId,
			workBlockDisposition: releaseDecisionTopClaimWorkBlockDisposition(action),
			rank: action.rank,
			priority: action.priority,
			nextStepKind: action.nextStepKind,
			evidenceWeHave: row.evidenceWeHave.map((item) => ({ ...item })),
			evidenceMissing: [...row.evidenceMissing],
			qssContrast: [...row.qssContrast],
			validationCommand: action.validationCommand,
			acceptanceEvidence: action.acceptanceEvidence,
			forbiddenShortcut: action.forbiddenShortcut,
			allowedWording: row.allowedWording,
			forbiddenWording: [...row.forbiddenWording],
			nextOwnerAction: `${action.rationale} Validate with \`${action.validationCommand}\`; acceptance evidence: ${action.acceptanceEvidence}`,
			boundary:
				'Owner-action queue row for top claims only. It exposes missing readyWhen gates and exact validation commands without satisfying gates or promoting stronger wording.',
		})),
	)
}

function releaseDecisionTopClaimWorkBlockDisposition(
	action: ReleaseProofNextOwnerAction,
): ReleaseProofReleaseDecisionDoNotPromoteItem['workBlockDisposition'] {
	switch (action.nextStepKind) {
		case 'owner-decision-or-fixture-replacement':
		case 'validation-run':
		case 'owner-decision-or-harness-expansion':
			return 'benchmark-corpus-blocker'
		case 'owner-boundary-approval':
		case 'publication-policy':
			return 'implementation-ready-blocker'
	}
}

function releaseDecisionBlockedOwnerActionQueue(
	items: readonly ReleaseProofReleaseDecisionDoNotPromoteItem[],
): readonly ReleaseProofBlockedOwnerAction[] {
	return items.flatMap((item) =>
		item.ownerLoops.map((ownerLoop) => ({
			name: item.name,
			ownerLoop,
			workBlockDisposition: item.workBlockDisposition,
			evidenceWeHave: [...item.evidenceWeHave],
			evidenceMissing: [...item.evidenceMissing],
			qssContrast: [...item.qssContrast],
			allowedWording: item.allowedWording,
			forbiddenWording: [...item.forbiddenWording],
			nextOwnerAction: item.nextOwnerAction,
			validationCommands: [...item.validationCommands],
			boundary:
				'Owner-action queue row for blocked claims only. It lets owner loops filter exact next work without promoting the claim, satisfying evidence, or changing release gates.',
		})),
	)
}

function releaseDecisionBenchmarkCorpusOwnerActionQueue(
	topClaimActions: readonly ReleaseProofTopClaimOwnerAction[],
	blockedActions: readonly ReleaseProofBlockedOwnerAction[],
): readonly ReleaseProofBenchmarkCorpusOwnerAction[] {
	return [
		...topClaimActions
			.filter((action) => action.workBlockDisposition === 'benchmark-corpus-blocker')
			.map((action) => ({
				...releaseDecisionBenchmarkCorpusRunContract(action.artifact, action.requirementId),
				sourceQueue: 'top-claim-owner-action' as const,
				name: action.artifact,
				claim: action.claim,
				ownerLoop: action.ownerLoop,
				requirementId: action.requirementId,
				workBlockDisposition: action.workBlockDisposition,
				ownerFiles: releaseDecisionBenchmarkCorpusOwnerFiles(action.artifact, action.requirementId),
				validationCommands: [action.validationCommand],
				commandsToRun: [action.validationCommand],
				failureEvidence: [...action.evidenceMissing],
				acceptanceCriteria: action.acceptanceEvidence,
				evidenceWeHave: action.evidenceWeHave.map(
					(item) => `${item.evidenceId}: \`${item.command}\` (${item.path})`,
				),
				evidenceMissing: [...action.evidenceMissing],
				qssContrast: [...action.qssContrast],
				allowedWording: action.allowedWording,
				forbiddenWording: [...action.forbiddenWording],
				nextOwnerAction: action.nextOwnerAction,
				boundary:
					'Benchmark/corpus owner-action queue row derived from top release claims. It tells the benchmark loop exactly what to run without satisfying the gate or authorizing stronger wording.',
			})),
		...blockedActions
			.filter((action) => action.workBlockDisposition === 'benchmark-corpus-blocker')
			.map((action) => ({
				...releaseDecisionBenchmarkCorpusRunContract(action.name),
				sourceQueue: 'blocked-owner-action' as const,
				name: action.name,
				claim: action.name,
				ownerLoop: action.ownerLoop,
				workBlockDisposition: action.workBlockDisposition,
				ownerFiles: releaseDecisionBenchmarkCorpusOwnerFiles(action.name),
				validationCommands: [...action.validationCommands],
				commandsToRun: [...action.validationCommands],
				failureEvidence: [...action.evidenceMissing],
				acceptanceCriteria: action.nextOwnerAction,
				evidenceWeHave: [...action.evidenceWeHave],
				evidenceMissing: [...action.evidenceMissing],
				qssContrast: [...action.qssContrast],
				allowedWording: action.allowedWording,
				forbiddenWording: [...action.forbiddenWording],
				nextOwnerAction: action.nextOwnerAction,
				boundary:
					'Benchmark/corpus owner-action queue row derived from do-not-promote decisions. It tells the benchmark loop exactly what to run without promoting the claim or satisfying the blocker.',
			})),
	]
}

function releaseDecisionBenchmarkCorpusOwnerFiles(
	name: ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name'],
	requirementId?: string,
): readonly string[] {
	if (name === 'safe-open-proof' && requirementId === 'public-edge-fixtures') {
		return ['fixtures/benchmarks/safe-open-fixture-scan.ts']
	}
	if (name === 'safe-open-proof' && requirementId === 'release-latency-run') {
		return ['fixtures/benchmarks/safe-open-proof.ts']
	}
	if (name === 'package-action-proof' && requirementId === 'edge-fixture-policy') {
		return ['fixtures/benchmarks/package-action-fixture-scan.ts']
	}
	if (name === 'package-action-proof' && requirementId === 'streaming-matrix-boundary') {
		return ['fixtures/benchmarks/package-action-proof.ts']
	}
	switch (name) {
		case 'formula-oracle-routing':
			return [
				'fixtures/benchmarks/formula-corpus-correctness.ts',
				'fixtures/benchmarks/formula-corpus-correctness.test.ts',
				'fixtures/xlsx/libreoffice/manifest.ts',
			]
		case 'columnar-scan-sidecars':
			return [
				'fixtures/benchmarks/columnar-sidecar.ts',
				'fixtures/benchmarks/columnar-sidecar.test.ts',
				'fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx',
			]
		case 'practical-latency-contracts':
			return [
				'fixtures/benchmarks/practical-latency-contracts.ts',
				'fixtures/benchmarks/practical-latency-contracts.test.ts',
			]
		default:
			return []
	}
}

type ReleaseProofBenchmarkCorpusRunContract = Pick<
	ReleaseProofBenchmarkCorpusOwnerAction,
	| 'runInputScope'
	| 'runEnvironment'
	| 'requiredOutputEvidence'
	| 'promotionCondition'
	| 'stopCondition'
>

function releaseDecisionBenchmarkCorpusRunContract(
	name: ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name'],
	requirementId?: string,
): ReleaseProofBenchmarkCorpusRunContract {
	if (name === 'safe-open-proof' && requirementId === 'public-edge-fixtures') {
		return {
			runInputScope:
				'Tracked safe-open fixture corpus: public unknown-part fixture plus disclosed generated signed and malformed structural packages.',
			runEnvironment:
				'Current repo fixture scan only; do not count private or license-unclear binary replacements unless product approves them.',
			requiredOutputEvidence: [
				'JSON scan output names generated structural cases, public replacement gaps, and tracked-corpus status.',
				'Owner decision either accepts disclosed generated topology fixtures or names public binary replacements.',
			],
			promotionCondition:
				'Promote only if product accepts disclosed generated signed/malformed topology proof or replaces those cases with approved public fixtures.',
			stopCondition:
				'Stop after the fixture scan and owner decision; do not broaden fixture search unless product rejects the disclosed generated cases.',
		}
	}
	if (name === 'safe-open-proof' && requirementId === 'release-latency-run') {
		return {
			runInputScope:
				'Standardized public safe-open cases from the release-latency owner-review profile.',
			runEnvironment:
				'Tracked-clean release environment with repeat 10, warmup 3, timing metadata, and no dirty-worktree or private-corpus timings counted.',
			requiredOutputEvidence: [
				'Open-plan and full-open median, p95, and CV for each required public case.',
				'Timing environment metadata, worktree `releaseClaimable` status, tracked-dirty file list, and product-approved non-threshold wording.',
			],
			promotionCondition:
				'Promote only bounded latency wording after the public profile passes CV/noise guardrails and product approves non-threshold copy.',
			stopCondition:
				'Stop at a bounded performance decision: accepted public profile, explicit no-target decision, or rerun blocker with failure output.',
		}
	}
	if (name === 'package-action-proof' && requirementId === 'edge-fixture-policy') {
		return {
			runInputScope:
				'Tracked package-action fixture corpus with disclosed generated signature-invalidation topology and public unknown-part candidates.',
			runEnvironment:
				'Current repo fixture scan only; generated topology remains owner-gated and cannot be hidden as public binary evidence.',
			requiredOutputEvidence: [
				'JSON scan output names generated structural cases, signature replacement gaps, and tracked-corpus status.',
				'Owner decision accepts disclosed generated signature topology or names public binary replacements.',
			],
			promotionCondition:
				'Promote only if product accepts disclosed generated signature topology for guarded proof or replaces it with approved public fixtures.',
			stopCondition:
				'Stop after fixture scan and owner decision; do not imply real signature verification or signed provenance.',
		}
	}
	if (name === 'package-action-proof' && requirementId === 'streaming-matrix-boundary') {
		return {
			runInputScope:
				'Package-action proof cases covering passthrough, regenerate, add, drop, macro/chart accounting, and known non-streaming cases.',
			runEnvironment:
				'No-timings proof run; representative streaming scope only, not a speed benchmark or full streaming parity run.',
			requiredOutputEvidence: [
				'Covered action kinds, missing action kinds, non-streaming cases, and public non-streaming cases.',
				'Performance owner decision accepting representative scope or naming the matrix expansion required.',
			],
			promotionCondition:
				'Promote only narrow representative streaming wording after owner approval; otherwise keep full streaming parity forbidden.',
			stopCondition:
				'Stop at representative-scope approval or an explicit matrix-expansion blocker.',
		}
	}
	if (name === 'formula-oracle-routing') {
		return {
			runInputScope:
				'Full public LibreOffice formula corpus manifest with formula-fidelity tag, 34 workbooks, 418 formulas, and 402 cached formulas compared.',
			runEnvironment:
				'Local public-corpus run only; cached-value routing evidence cannot count private corpora or unavailable Excel oracle execution as parity.',
			requiredOutputEvidence: [
				'Thresholded JSON proves 23 accepted mismatches, 0 unaccepted mismatches, 0 semantic mismatches, 22 numeric-drift mismatches, 1 stale-oracle mismatch, and 34 semantic-perfect workbooks.',
				'Strict zero-mismatch command remains captured as blocker evidence until accepted mismatches have owner-approved wording and non-accepted oracle routes have real artifacts.',
				'Named HyperFormula, LibreOffice, Excel, or static-golden adapter gaps before any Excel-compatible formula wording.',
			],
			promotionCondition:
				'Promote only after public corpus artifacts and real oracle adapters emit reproducible non-accepted route, skip, divergence, and verifier evidence with owner-approved thresholds.',
			stopCondition:
				'Stop at the public-corpus threshold gate plus a named corpus-class or adapter blocker; do not add formula compatibility wording from cached-value routing alone.',
		}
	}
	if (name === 'columnar-scan-sidecars') {
		return {
			runInputScope:
				'Public SEC MMF workbook Table 9 claim report plus follow-on structurally diverse public table/range fixtures.',
			runEnvironment:
				'Benchmark-only sidecar run with repeats, checksum parity, build cost, invalidation cost, and memory or payload guardrails.',
			requiredOutputEvidence: [
				'Repeated-scan median, p95 or noise/CV, sidecar build cost, invalidation cost, memory overhead, and checksum parity.',
				'Decision that the result remains benchmark-only or names a real SDK/API/MCP product surface.',
			],
			promotionCondition:
				'Promote only if diverse public fixtures show end-to-end wins including build/invalidation cost and bounded memory with checksum parity.',
			stopCondition:
				'Stop at benchmark-only evidence, a production-surface owner blocker, or a do-not-promote speed/product decision.',
		}
	}
	if (name === 'practical-latency-contracts') {
		return {
			runInputScope:
				'Public-tracked practical workflow contract preset covering first-view, edit-verify, and repeated-inspection envelopes.',
			runEnvironment:
				'Tracked-clean worktree; run the dry-run first, then repeat 3 warmup 1 with no private inputs or dirty-worktree timings counted.',
			requiredOutputEvidence: [
				'Summary/profile JSON with median, p95, CV/noise, input provenance, and memory or payload guardrails.',
				'One profile-backed production target or explicit no-target decision plus product-approved non-threshold wording.',
			],
			promotionCondition:
				'Promote only after tracked-clean public profile artifacts support non-threshold workflow wording approved by product and performance.',
			stopCondition:
				'Stop at accepted public profile evidence, explicit no-target decision, or rerun blocker with command failure evidence.',
		}
	}
	return {
		runInputScope: '',
		runEnvironment: '',
		requiredOutputEvidence: [],
		promotionCondition: '',
		stopCondition: '',
	}
}

function releaseDecisionImplementationReadyOwnerActionQueue(
	topClaimActions: readonly ReleaseProofTopClaimOwnerAction[],
	blockedActions: readonly ReleaseProofBlockedOwnerAction[],
): readonly ReleaseProofImplementationReadyOwnerAction[] {
	return [
		...topClaimActions
			.filter((action) => action.workBlockDisposition === 'implementation-ready-blocker')
			.map((action) => ({
				sourceQueue: 'top-claim-owner-action' as const,
				name: action.artifact,
				claim: action.claim,
				ownerLoop: action.ownerLoop,
				requirementId: action.requirementId,
				workBlockDisposition: action.workBlockDisposition,
				ownerFiles: releaseDecisionImplementationOwnerFiles(action.artifact, action.requirementId),
				validationCommands: releaseDecisionImplementationValidationCommands(action),
				commandsToRun: releaseDecisionImplementationValidationCommands(action),
				failureEvidence: [...action.evidenceMissing],
				acceptanceCriteria: action.acceptanceEvidence,
				evidenceWeHave: action.evidenceWeHave.map(
					(item) => `${item.evidenceId}: \`${item.command}\` (${item.path})`,
				),
				evidenceMissing: [...action.evidenceMissing],
				qssContrast: [...action.qssContrast],
				allowedWording: action.allowedWording,
				forbiddenWording: [...action.forbiddenWording],
				nextOwnerAction: action.nextOwnerAction,
				boundary:
					'Implementation-ready owner-action queue row derived from top release claims. It tells correctness, product, or release owners exactly what to validate without satisfying the gate or authorizing stronger wording.',
			})),
		...blockedActions
			.filter((action) => action.workBlockDisposition === 'implementation-ready-blocker')
			.map((action) => ({
				sourceQueue: 'blocked-owner-action' as const,
				name: action.name,
				claim: action.name,
				ownerLoop: action.ownerLoop,
				workBlockDisposition: action.workBlockDisposition,
				ownerFiles: releaseDecisionImplementationOwnerFiles(action.name),
				validationCommands: [...action.validationCommands],
				commandsToRun: [...action.validationCommands],
				failureEvidence: [...action.evidenceMissing],
				acceptanceCriteria: action.nextOwnerAction,
				evidenceWeHave: [...action.evidenceWeHave],
				evidenceMissing: [...action.evidenceMissing],
				qssContrast: [...action.qssContrast],
				allowedWording: action.allowedWording,
				forbiddenWording: [...action.forbiddenWording],
				nextOwnerAction: action.nextOwnerAction,
				boundary:
					'Implementation-ready owner-action queue row derived from do-not-promote decisions. It tells correctness, product, or release owners exactly what to validate without promoting the claim or satisfying the blocker.',
			})),
	]
}

function releaseDecisionImplementationOwnerFiles(
	name: ReleaseProofIndexArtifactName | ReleaseProofReleaseDecisionDoNotPromoteItem['name'],
	requirementId?: string,
): readonly string[] {
	if (
		(name === 'safe-open-proof' || name === 'package-action-proof') &&
		(requirementId === 'publication-boundary' || requirementId === 'provenance-boundary')
	) {
		return ['fixtures/benchmarks/release-proof-index.ts']
	}
	if (name === 'safe-open-proof' && requirementId === 'compact-report-publication-policy') {
		return ['fixtures/benchmarks/safe-open-proof.ts']
	}
	if (name === 'package-action-proof' && requirementId === 'compact-report-publication-policy') {
		return ['fixtures/benchmarks/package-action-proof.ts']
	}
	if (name === 'package-action-proof' && requirementId === 'unsupported-feature-boundary') {
		return [
			'fixtures/benchmarks/package-action-proof.ts',
			'fixtures/corpus/high-risk-package-contract.test.ts',
		]
	}
	switch (name) {
		case 'formula-language-service-primitives':
			return [
				'fixtures/benchmarks/formula-assist-proof.ts',
				'packages/sdk/src/formula-edit.test.ts',
				'apps/cli/src/cli.test.ts',
				'apps/api/src/server.test.ts',
				'apps/mcp/src/index.test.ts',
			]
		case 'token-bounded-agent-view':
			return [
				'fixtures/benchmarks/agent-view-budget-proof.test.ts',
				'fixtures/benchmarks/agent-view-recovery-proof.test.ts',
				'packages/sdk/src/sdk.test.ts',
				'apps/cli/src/cli.test.ts',
				'apps/api/src/server.test.ts',
				'apps/mcp/src/index.test.ts',
			]
		case 'retained-viewport-patch-history':
			return [
				'fixtures/benchmarks/viewport-patch-proof.test.ts',
				'packages/sdk/src/interactive-contract.test.ts',
				'apps/api/src/server.test.ts',
				'apps/mcp/src/index.test.ts',
			]
		case 'release-proof-bundle':
			return [
				'scripts/release-rc-gate.ts',
				'fixtures/benchmarks/safe-open-proof.ts',
				'fixtures/benchmarks/package-action-proof.ts',
				'fixtures/benchmarks/release-proof-index.ts',
			]
		case 'property-journal-laws':
			return [
				'fixtures/benchmarks/journal-law-proof.test.ts',
				'packages/sdk/src/journal-exactness.test.ts',
			]
		case 'agent-workflow-observability':
			return [
				'packages/sdk/src/agent-workflow.test.ts',
				'apps/cli/src/cli.test.ts',
				'apps/api/src/server.test.ts',
				'apps/mcp/src/index.test.ts',
			]
		default:
			return []
	}
}

function releaseDecisionImplementationValidationCommands(
	action: ReleaseProofTopClaimOwnerAction,
): readonly string[] {
	if (
		action.artifact === 'package-action-proof' &&
		action.requirementId === 'unsupported-feature-boundary'
	) {
		return [HIGH_RISK_PACKAGE_CONTRACT_COMMAND, action.validationCommand]
	}
	return [action.validationCommand]
}

function releaseDecisionClaimDowngradeOwnerActionQueue(
	blockedActions: readonly ReleaseProofBlockedOwnerAction[],
): readonly ReleaseProofClaimDowngradeOwnerAction[] {
	return blockedActions
		.filter((action) => action.workBlockDisposition === 'claim-downgrade-do-not-promote')
		.map((action) => ({
			sourceQueue: 'blocked-owner-action' as const,
			name: action.name,
			claim: action.name,
			ownerLoop: action.ownerLoop,
			workBlockDisposition: action.workBlockDisposition,
			ownerFiles: releaseDecisionClaimDowngradeOwnerFiles(action.name),
			validationCommands: [...action.validationCommands],
			commandsToRun: [...action.validationCommands],
			failureEvidence: [...action.evidenceMissing],
			acceptanceCriteria: action.nextOwnerAction,
			evidenceWeHave: [...action.evidenceWeHave],
			evidenceMissing: [...action.evidenceMissing],
			qssContrast: [...action.qssContrast],
			allowedWording: action.allowedWording,
			forbiddenWording: [...action.forbiddenWording],
			nextOwnerAction: action.nextOwnerAction,
			boundary:
				'Claim downgrade owner-action queue row derived from do-not-promote decisions. It tells product or release owners exactly what to classify, archive, or keep forbidden without promoting the claim.',
		}))
}

function releaseDecisionClaimDowngradeOwnerFiles(
	name: ReleaseProofReleaseDecisionDoNotPromoteItem['name'],
): readonly string[] {
	if (name === 'columnar-scan-sidecars') {
		return [
			'fixtures/benchmarks/release-proof-index.ts',
			'fixtures/benchmarks/release-proof-index.test.ts',
			'fixtures/benchmarks/columnar-sidecar.ts',
			'fixtures/benchmarks/columnar-sidecar.test.ts',
		]
	}
	if (name === 'property-journal-laws') {
		return [
			'fixtures/benchmarks/release-proof-index.ts',
			'fixtures/benchmarks/release-proof-index.test.ts',
			'fixtures/benchmarks/journal-law-proof.ts',
			'fixtures/benchmarks/journal-law-proof.test.ts',
			'packages/sdk/src/journal-exactness.test.ts',
		]
	}
	if (name !== 'research-surface-hygiene') return []
	return [
		'research/',
		'research/experiments/',
		'scripts/ascend-loop-manager.ts',
		'tmp/ascend-loop-manager/',
		'fixtures/benchmarks/release-proof-index.test.ts',
	]
}

function releaseDecisionOwnerActionQueueCoverage(
	topClaimActions: readonly ReleaseProofTopClaimOwnerAction[],
	blockedActions: readonly ReleaseProofBlockedOwnerAction[],
	benchmarkCorpusActions: readonly ReleaseProofBenchmarkCorpusOwnerAction[],
	implementationReadyActions: readonly ReleaseProofImplementationReadyOwnerAction[],
	claimDowngradeActions: readonly ReleaseProofClaimDowngradeOwnerAction[],
): ReleaseProofOwnerActionQueueCoverage {
	const coveredKeys = new Set([
		...benchmarkCorpusActions.map(releaseDecisionDispositionActionKey),
		...implementationReadyActions.map(releaseDecisionDispositionActionKey),
		...claimDowngradeActions.map(releaseDecisionDispositionActionKey),
	])
	const uncoveredTopClaimActionKeys = topClaimActions
		.map(releaseDecisionTopClaimActionKey)
		.filter((key) => !coveredKeys.has(key))
	const uncoveredBlockedActionKeys = blockedActions
		.map(releaseDecisionBlockedActionKey)
		.filter((key) => !coveredKeys.has(key))
	const coveredActionCount =
		benchmarkCorpusActions.length + implementationReadyActions.length + claimDowngradeActions.length
	const status =
		uncoveredTopClaimActionKeys.length === 0 &&
		uncoveredBlockedActionKeys.length === 0 &&
		coveredActionCount === topClaimActions.length + blockedActions.length
			? 'all-owner-actions-covered-by-disposition-queues'
			: 'coverage-gap'
	return {
		status,
		sourceTopClaimActionCount: topClaimActions.length,
		sourceBlockedActionCount: blockedActions.length,
		benchmarkCorpusActionCount: benchmarkCorpusActions.length,
		implementationReadyActionCount: implementationReadyActions.length,
		claimDowngradeActionCount: claimDowngradeActions.length,
		coveredActionCount,
		uncoveredTopClaimActionKeys,
		uncoveredBlockedActionKeys,
		boundary:
			'Owner action queue coverage only. It proves routing rows are present in the disposition queues without satisfying any claim gate, approving wording, or executing validation commands.',
	}
}

function releaseDecisionBenchmarkCorpusRunContractCoverage(
	benchmarkCorpusActions: readonly ReleaseProofBenchmarkCorpusOwnerAction[],
): ReleaseProofBenchmarkCorpusRunContractCoverage {
	const missingInputScopeActionKeys = benchmarkCorpusActions
		.filter((action) => action.runInputScope.trim().length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingRunEnvironmentActionKeys = benchmarkCorpusActions
		.filter((action) => action.runEnvironment.trim().length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingRequiredOutputEvidenceActionKeys = benchmarkCorpusActions
		.filter((action) => action.requiredOutputEvidence.length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingPromotionConditionActionKeys = benchmarkCorpusActions
		.filter((action) => action.promotionCondition.trim().length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingStopConditionActionKeys = benchmarkCorpusActions
		.filter((action) => action.stopCondition.trim().length === 0)
		.map(releaseDecisionDispositionActionKey)
	const status =
		missingInputScopeActionKeys.length === 0 &&
		missingRunEnvironmentActionKeys.length === 0 &&
		missingRequiredOutputEvidenceActionKeys.length === 0 &&
		missingPromotionConditionActionKeys.length === 0 &&
		missingStopConditionActionKeys.length === 0
			? 'all-benchmark-corpus-actions-have-run-contract'
			: 'benchmark-corpus-run-contract-gap'
	return {
		status,
		actionCount: benchmarkCorpusActions.length,
		missingInputScopeActionKeys,
		missingRunEnvironmentActionKeys,
		missingRequiredOutputEvidenceActionKeys,
		missingPromotionConditionActionKeys,
		missingStopConditionActionKeys,
		boundary:
			'Benchmark/corpus run contract coverage only. It proves each benchmark or corpus blocker names input scope, run environment, required output evidence, promotion condition, and stop condition without executing commands or promoting claims.',
	}
}

type ReleaseProofDispositionOwnerAction =
	| ReleaseProofBenchmarkCorpusOwnerAction
	| ReleaseProofImplementationReadyOwnerAction
	| ReleaseProofClaimDowngradeOwnerAction

function releaseDecisionOwnerActionExecutionContractCoverage(
	benchmarkCorpusActions: readonly ReleaseProofBenchmarkCorpusOwnerAction[],
	implementationReadyActions: readonly ReleaseProofImplementationReadyOwnerAction[],
	claimDowngradeActions: readonly ReleaseProofClaimDowngradeOwnerAction[],
): ReleaseProofOwnerActionExecutionContractCoverage {
	const actions: readonly ReleaseProofDispositionOwnerAction[] = [
		...benchmarkCorpusActions,
		...implementationReadyActions,
		...claimDowngradeActions,
	]
	const missingOwnerFileActionKeys = actions
		.filter((action) => action.ownerFiles.length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingCommandActionKeys = actions
		.filter((action) => action.commandsToRun.length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingFailureEvidenceActionKeys = actions
		.filter((action) => action.failureEvidence.length === 0)
		.map(releaseDecisionDispositionActionKey)
	const missingAcceptanceCriteriaActionKeys = actions
		.filter((action) => action.acceptanceCriteria.trim().length === 0)
		.map(releaseDecisionDispositionActionKey)
	const status =
		missingOwnerFileActionKeys.length === 0 &&
		missingCommandActionKeys.length === 0 &&
		missingFailureEvidenceActionKeys.length === 0 &&
		missingAcceptanceCriteriaActionKeys.length === 0
			? 'all-disposition-owner-actions-have-execution-contract'
			: 'execution-contract-gap'
	return {
		status,
		actionCount: actions.length,
		benchmarkCorpusActionCount: benchmarkCorpusActions.length,
		implementationReadyActionCount: implementationReadyActions.length,
		claimDowngradeActionCount: claimDowngradeActions.length,
		missingOwnerFileActionKeys,
		missingCommandActionKeys,
		missingFailureEvidenceActionKeys,
		missingAcceptanceCriteriaActionKeys,
		boundary:
			'Execution contract coverage only. It proves disposition queue rows name owner files, commands to run, failure evidence, and acceptance criteria without executing those commands or satisfying release gates.',
	}
}

function releaseDecisionTopClaimActionKey(action: ReleaseProofTopClaimOwnerAction): string {
	return `top-claim-owner-action:${action.ownerLoop}:${action.artifact}:${action.requirementId}:${action.workBlockDisposition}`
}

function releaseDecisionBlockedActionKey(action: ReleaseProofBlockedOwnerAction): string {
	return `blocked-owner-action:${action.ownerLoop}:${action.name}:none:${action.workBlockDisposition}`
}

function releaseDecisionDispositionActionKey(
	action:
		| ReleaseProofBenchmarkCorpusOwnerAction
		| ReleaseProofImplementationReadyOwnerAction
		| ReleaseProofClaimDowngradeOwnerAction,
): string {
	const requirementId = 'requirementId' in action ? action.requirementId : undefined
	return `${action.sourceQueue}:${action.ownerLoop}:${action.name}:${requirementId ?? 'none'}:${action.workBlockDisposition}`
}

function releaseDecisionDoNotPromoteItem(
	note: ReleaseProofQssArchivedResearchNote,
): ReleaseProofReleaseDecisionDoNotPromoteItem {
	const portfolioClaim = CLAIM_PORTFOLIO.find((claim) => claim.name === note.name)
	const deferredClaim = DEFERRED_CLAIMS.find((claim) => claim.name === note.name)
	const excludedEvidence = EXCLUDED_EVIDENCE.find((evidence) => evidence.name === note.name)
	const proof = portfolioClaim?.evidenceNeeded
	const formulaLanguageServiceBlocker =
		note.name === 'formula-language-service-primitives'
			? FORMULA_LANGUAGE_SERVICE_BLOCKER
			: undefined
	const releaseProofBundleBlocker =
		note.name === 'release-proof-bundle' ? RELEASE_PROOF_BUNDLE_BLOCKER : undefined
	const propertyJournalLawBlocker =
		note.name === 'property-journal-laws' ? PROPERTY_JOURNAL_LAW_BLOCKER : undefined
	const researchHygieneBlocker =
		note.name === 'research-surface-hygiene' ? RESEARCH_SURFACE_HYGIENE_BLOCKER : undefined
	const columnarSidecarBlocker =
		note.name === 'columnar-scan-sidecars' ? COLUMNAR_SIDECAR_BLOCKER : undefined
	const formulaOracleRoutingBlocker =
		note.name === 'formula-oracle-routing' ? FORMULA_ORACLE_ROUTING_BLOCKER : undefined
	const tokenBoundedAgentViewBlocker =
		note.name === 'token-bounded-agent-view' ? TOKEN_BOUNDED_AGENT_VIEW_BLOCKER : undefined
	const retainedViewportPatchBlocker =
		note.name === 'retained-viewport-patch-history' ? RETAINED_VIEWPORT_PATCH_BLOCKER : undefined
	const agentWorkflowObservabilityBlocker =
		note.name === 'agent-workflow-observability' ? AGENT_WORKFLOW_OBSERVABILITY_BLOCKER : undefined
	const practicalLatencyContractsBlocker =
		note.name === 'practical-latency-contracts' ? PRACTICAL_LATENCY_CONTRACTS_BLOCKER : undefined
	const allowedWording =
		formulaLanguageServiceBlocker?.allowedWording ??
		releaseProofBundleBlocker?.allowedWording ??
		propertyJournalLawBlocker?.allowedWording ??
		researchHygieneBlocker?.allowedWording ??
		columnarSidecarBlocker?.allowedWording ??
		formulaOracleRoutingBlocker?.allowedWording ??
		tokenBoundedAgentViewBlocker?.allowedWording ??
		retainedViewportPatchBlocker?.allowedWording ??
		agentWorkflowObservabilityBlocker?.allowedWording ??
		practicalLatencyContractsBlocker?.allowedWording ??
		`Do not promote ${note.name} as release wording today; use it only as owner planning or research evidence.`
	return {
		name: note.name,
		status: 'do-not-promote-yet',
		workBlockDisposition: releaseDecisionWorkBlockDisposition(note.name),
		ownerLoops: [...note.ownerLoops],
		reason: note.reason,
		evidenceWeHave: [
			note.reason,
			...(formulaLanguageServiceBlocker?.evidenceWeHave ?? []),
			...(releaseProofBundleBlocker?.evidenceWeHave ?? []),
			...(propertyJournalLawBlocker?.evidenceWeHave ?? []),
			...(researchHygieneBlocker?.evidenceWeHave ?? []),
			...(columnarSidecarBlocker?.evidenceWeHave ?? []),
			...(formulaOracleRoutingBlocker?.evidenceWeHave ?? []),
			...(tokenBoundedAgentViewBlocker?.evidenceWeHave ?? []),
			...(retainedViewportPatchBlocker?.evidenceWeHave ?? []),
			...(agentWorkflowObservabilityBlocker?.evidenceWeHave ?? []),
			...(practicalLatencyContractsBlocker?.evidenceWeHave ?? []),
			...(portfolioClaim?.proofCommand
				? [`Existing proof command: \`${portfolioClaim.proofCommand}\`.`]
				: []),
			...(excludedEvidence
				? [`Excluded diagnostic command: \`${excludedEvidence.command}\`.`]
				: []),
		],
		evidenceMissing: [
			...(deferredClaim ? [deferredClaim.proofNeeded] : []),
			...(excludedEvidence ? [excludedEvidence.eligibilityRule] : []),
			...(formulaLanguageServiceBlocker?.evidenceMissing ?? []),
			...(releaseProofBundleBlocker?.evidenceMissing ?? []),
			...(propertyJournalLawBlocker?.evidenceMissing ?? []),
			...(researchHygieneBlocker?.evidenceMissing ?? []),
			...(columnarSidecarBlocker?.evidenceMissing ?? []),
			...(formulaOracleRoutingBlocker?.evidenceMissing ?? []),
			...(tokenBoundedAgentViewBlocker?.evidenceMissing ?? []),
			...(retainedViewportPatchBlocker?.evidenceMissing ?? []),
			...(agentWorkflowObservabilityBlocker?.evidenceMissing ?? []),
			...(practicalLatencyContractsBlocker?.evidenceMissing ?? []),
			...(proof ? [proof.fixture, proof.benchmark, proof.surface, proof.validationGate] : []),
		],
		qssContrast: practicalLatencyContractsBlocker?.qssContrast ?? [
			proof?.competitorContrast ??
				'QSS contrast is blocked until this diagnostic evidence changes a top-two release claim.',
		],
		allowedWording,
		forbiddenWording: [
			note.killCriterion,
			...(formulaLanguageServiceBlocker?.forbiddenWording ?? []),
			...(releaseProofBundleBlocker?.forbiddenWording ?? []),
			...(propertyJournalLawBlocker?.forbiddenWording ?? []),
			...(researchHygieneBlocker?.forbiddenWording ?? []),
			...(columnarSidecarBlocker?.forbiddenWording ?? []),
			...(formulaOracleRoutingBlocker?.forbiddenWording ?? []),
			...(tokenBoundedAgentViewBlocker?.forbiddenWording ?? []),
			...(retainedViewportPatchBlocker?.forbiddenWording ?? []),
			...(agentWorkflowObservabilityBlocker?.forbiddenWording ?? []),
			...(practicalLatencyContractsBlocker?.forbiddenWording ?? []),
			...(proof ? [proof.honestBoundary] : []),
			...(excludedEvidence ? [excludedEvidence.boundary] : []),
		],
		nextOwnerAction:
			formulaLanguageServiceBlocker?.ownerAction ??
			releaseProofBundleBlocker?.ownerAction ??
			propertyJournalLawBlocker?.ownerAction ??
			researchHygieneBlocker?.ownerAction ??
			columnarSidecarBlocker?.ownerAction ??
			formulaOracleRoutingBlocker?.ownerAction ??
			tokenBoundedAgentViewBlocker?.ownerAction ??
			retainedViewportPatchBlocker?.ownerAction ??
			agentWorkflowObservabilityBlocker?.ownerAction ??
			practicalLatencyContractsBlocker?.ownerAction ??
			deferredClaim?.proofNeeded ??
			excludedEvidence?.eligibilityRule ??
			proof?.validationGate ??
			'No owner action is release-blocking until this claim changes the top-two release gate.',
		validationCommands: releaseDecisionValidationCommands(note.name),
		killCriterion: note.killCriterion,
		boundary:
			'Archived research note for release stewardship. Do not turn this into release wording or a new implementation surface until it changes the top-two claim gate.',
	}
}

function releaseDecisionValidationCommands(
	name: ReleaseProofQssArchivedResearchNote['name'],
): readonly string[] {
	switch (name) {
		case 'safe-open-proof':
		case 'package-action-proof':
			return [
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json',
			]
		case 'formula-language-service-primitives':
			return [
				'bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json',
				'bun test packages/sdk/src/formula-edit.test.ts --timeout 30000',
				'bun test apps/cli/src/cli.test.ts -t "formula assist returns formula IDE help without opening a workbook" --timeout 30000',
				'bun test apps/api/src/server.test.ts -t "formula-assist exposes diagnostics, completions, signature help, and reference edits" --timeout 30000',
				'bun test apps/mcp/src/index.test.ts -t "ascend.formula_assist exposes formula IDE helpers for agents" --timeout 30000',
			]
		case 'token-bounded-agent-view':
			return [
				'bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts',
				'bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets without losing shape facts" --timeout 30000',
				'bun test apps/cli/src/cli.test.ts -t "agent-view --tokens returns budget metadata" --timeout 30000',
				'bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata" --timeout 30000',
				'bun test apps/mcp/src/index.test.ts -t "ascend.agent_view exposes token budget metadata" --timeout 30000',
			]
		case 'retained-viewport-patch-history':
			return [
				'bun test fixtures/benchmarks/viewport-patch-proof.test.ts',
				'bun test packages/sdk/src/interactive-contract.test.ts -t "interactive viewport tokens from other sessions refreshes and retained-log gaps force fresh reads|interactive viewport patch results expose invalidation reasons without silent nulls|interactive pull patches refuse metadata and layout edits that need fresh viewport state|interactive pull patches reject tokens older than the retained change log" --timeout 30000',
				'bun test apps/api/src/server.test.ts -t "compact changedSince reads invalidate when the requested window changes|compact changedSince reads return a fresh window after source changes" --timeout 30000',
				'bun test apps/mcp/src/index.test.ts -t "ascend.read compact changedSince invalidates when the requested window changes|ascend.read compact changedSince invalidates when selected columns change|ascend.read compact changedSince returns a fresh window after source changes" --timeout 30000',
			]
		case 'release-proof-bundle':
			return [
				'bun run release:rc:gate',
				'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json',
				'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json',
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			]
		case 'formula-oracle-routing':
			return [
				'bun test fixtures/benchmarks/formula-corpus-correctness.test.ts fixtures/xlsx/libreoffice-fixtures.test.ts --timeout 60000',
				'bun run fixtures/benchmarks/formula-corpus-correctness.ts --corpus-root fixtures/xlsx/libreoffice --manifest fixtures/xlsx/libreoffice/manifest.ts --tag formula-fidelity --max-mismatches 23 --max-unaccepted-mismatches 0 --max-semantic-mismatches 0 --max-errors 0 --min-workbooks 34 --min-formulas 418 --min-compared-formulas 402 --min-semantic-perfect-workbooks 34 --json',
				'bun run fixtures/benchmarks/formula-corpus-correctness.ts --corpus-root fixtures/xlsx/libreoffice --manifest fixtures/xlsx/libreoffice/manifest.ts --tag formula-fidelity --max-mismatches 0 --max-accepted-mismatches 0 --max-unaccepted-mismatches 0 --max-semantic-mismatches 0 --max-volatile-oracle-skips 0 --max-errors 0 --json',
			]
		case 'property-journal-laws':
			return [
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json',
				'bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000',
			]
		case 'columnar-scan-sidecars':
			return [
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json',
				'bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000',
			]
		case 'agent-workflow-observability':
			return [
				'bun test packages/sdk/src/agent-workflow.test.ts -t "keeps unresolved external-link package graph failures visible to agents|release proof bundle links plan, commit, reopen, diff, and audit evidence|prepared agent commits surface post-write audit failures as blocking model output" --timeout 30000',
				'bun test apps/cli/src/cli.test.ts -t "plan invalid ops return structured batch repair details|check surfaces structured issue metadata for agent repair|trace shows values and respects max depth|trace shows precedents for formula cell" --timeout 30000',
				'bun test apps/api/src/server.test.ts -t "trace returns structured partial-load diagnostics for capped formula views|plan invalid ops return structured batch repair details|prepared path mutation handles surface post-write audit failures as blocked output|direct path mutation commits surface post-write audit failures as blocked output" --timeout 30000',
				'bun test apps/mcp/src/index.test.ts -t "ascend.trace returns structured partial-load diagnostics for capped formula views|ascend.plan invalid ops return structured batch repair details|prepared MCP path mutation handles surface post-write audit failures as blocked output|direct MCP path mutation commits surface post-write audit failures as blocked output" --timeout 30000',
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			]
		case 'research-surface-hygiene':
			return [
				'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
				'bun test fixtures/benchmarks/release-proof-index.test.ts',
			]
		case 'practical-latency-contracts':
			return [
				'bun test fixtures/benchmarks/practical-latency-contracts.test.ts --timeout 30000',
				'bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 1 --warmup 0 --dry-run --json',
				'bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
			]
	}
}

function releaseDecisionWorkBlockDisposition(
	name: ReleaseProofQssArchivedResearchNote['name'],
): ReleaseProofReleaseDecisionDoNotPromoteItem['workBlockDisposition'] {
	switch (name) {
		case 'safe-open-proof':
		case 'package-action-proof':
			return 'claim-downgrade-do-not-promote'
		case 'formula-oracle-routing':
		case 'practical-latency-contracts':
			return 'benchmark-corpus-blocker'
		case 'columnar-scan-sidecars':
		case 'property-journal-laws':
		case 'research-surface-hygiene':
			return 'claim-downgrade-do-not-promote'
		case 'formula-language-service-primitives':
		case 'token-bounded-agent-view':
		case 'retained-viewport-patch-history':
		case 'release-proof-bundle':
		case 'agent-workflow-observability':
			return 'implementation-ready-blocker'
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
		doNotPromoteDispositionSummary: {
			...board.doNotPromoteDispositionSummary,
			implementationReadyBlockerNames: [
				...board.doNotPromoteDispositionSummary.implementationReadyBlockerNames,
			],
			benchmarkCorpusBlockerNames: [
				...board.doNotPromoteDispositionSummary.benchmarkCorpusBlockerNames,
			],
			claimDowngradeDoNotPromoteNames: [
				...board.doNotPromoteDispositionSummary.claimDowngradeDoNotPromoteNames,
			],
		},
		releaseWordingDecisionSummary: {
			...board.releaseWordingDecisionSummary,
			localAllowedClaimNames: [...board.releaseWordingDecisionSummary.localAllowedClaimNames],
			doNotPromoteClaimNames: [...board.releaseWordingDecisionSummary.doNotPromoteClaimNames],
			localAllowedWordingByClaim: {
				...board.releaseWordingDecisionSummary.localAllowedWordingByClaim,
			},
			doNotPromoteAllowedWordingByClaim: {
				...board.releaseWordingDecisionSummary.doNotPromoteAllowedWordingByClaim,
			},
			forbiddenWordingByClaim: Object.fromEntries(
				Object.entries(board.releaseWordingDecisionSummary.forbiddenWordingByClaim).map(
					([name, forbiddenWording]) => [name, [...forbiddenWording]],
				),
			) as ReleaseProofReleaseWordingDecisionSummary['forbiddenWordingByClaim'],
		},
		todayCommitClaimMatrix: board.todayCommitClaimMatrix.map((row) => ({
			...row,
			commits: [...row.commits],
			evidenceProvesIt: [...row.evidenceProvesIt],
			forbiddenWording: [...row.forbiddenWording],
		})),
		claimDecisionContractCoverage: {
			...board.claimDecisionContractCoverage,
			missingEvidenceWeHaveKeys: [...board.claimDecisionContractCoverage.missingEvidenceWeHaveKeys],
			missingEvidenceMissingKeys: [
				...board.claimDecisionContractCoverage.missingEvidenceMissingKeys,
			],
			missingQssContrastKeys: [...board.claimDecisionContractCoverage.missingQssContrastKeys],
			missingAllowedWordingKeys: [...board.claimDecisionContractCoverage.missingAllowedWordingKeys],
			missingForbiddenWordingKeys: [
				...board.claimDecisionContractCoverage.missingForbiddenWordingKeys,
			],
			missingNextOwnerActionKeys: [
				...board.claimDecisionContractCoverage.missingNextOwnerActionKeys,
			],
		},
		doNotPromoteYet: board.doNotPromoteYet.map((item) => ({
			...item,
			ownerLoops: [...item.ownerLoops],
			evidenceWeHave: [...item.evidenceWeHave],
			evidenceMissing: [...item.evidenceMissing],
			qssContrast: [...item.qssContrast],
			forbiddenWording: [...item.forbiddenWording],
			validationCommands: [...item.validationCommands],
		})),
		blockedOwnerActionQueue: board.blockedOwnerActionQueue.map((row) => ({
			...row,
			evidenceWeHave: [...row.evidenceWeHave],
			evidenceMissing: [...row.evidenceMissing],
			qssContrast: [...row.qssContrast],
			forbiddenWording: [...row.forbiddenWording],
			validationCommands: [...row.validationCommands],
		})),
		benchmarkCorpusOwnerActionQueue: board.benchmarkCorpusOwnerActionQueue.map((row) => ({
			...row,
			ownerFiles: [...row.ownerFiles],
			validationCommands: [...row.validationCommands],
			commandsToRun: [...row.commandsToRun],
			failureEvidence: [...row.failureEvidence],
			requiredOutputEvidence: [...row.requiredOutputEvidence],
			evidenceWeHave: [...row.evidenceWeHave],
			evidenceMissing: [...row.evidenceMissing],
			qssContrast: [...row.qssContrast],
			forbiddenWording: [...row.forbiddenWording],
		})),
		implementationReadyOwnerActionQueue: board.implementationReadyOwnerActionQueue.map((row) => ({
			...row,
			ownerFiles: [...row.ownerFiles],
			validationCommands: [...row.validationCommands],
			commandsToRun: [...row.commandsToRun],
			failureEvidence: [...row.failureEvidence],
			evidenceWeHave: [...row.evidenceWeHave],
			evidenceMissing: [...row.evidenceMissing],
			qssContrast: [...row.qssContrast],
			forbiddenWording: [...row.forbiddenWording],
		})),
		claimDowngradeOwnerActionQueue: board.claimDowngradeOwnerActionQueue.map((row) => ({
			...row,
			ownerFiles: [...row.ownerFiles],
			validationCommands: [...row.validationCommands],
			commandsToRun: [...row.commandsToRun],
			failureEvidence: [...row.failureEvidence],
			evidenceWeHave: [...row.evidenceWeHave],
			evidenceMissing: [...row.evidenceMissing],
			qssContrast: [...row.qssContrast],
			forbiddenWording: [...row.forbiddenWording],
		})),
		ownerActionQueueCoverage: {
			...board.ownerActionQueueCoverage,
			uncoveredTopClaimActionKeys: [...board.ownerActionQueueCoverage.uncoveredTopClaimActionKeys],
			uncoveredBlockedActionKeys: [...board.ownerActionQueueCoverage.uncoveredBlockedActionKeys],
		},
		benchmarkCorpusRunContractCoverage: {
			...board.benchmarkCorpusRunContractCoverage,
			missingInputScopeActionKeys: [
				...board.benchmarkCorpusRunContractCoverage.missingInputScopeActionKeys,
			],
			missingRunEnvironmentActionKeys: [
				...board.benchmarkCorpusRunContractCoverage.missingRunEnvironmentActionKeys,
			],
			missingRequiredOutputEvidenceActionKeys: [
				...board.benchmarkCorpusRunContractCoverage.missingRequiredOutputEvidenceActionKeys,
			],
			missingPromotionConditionActionKeys: [
				...board.benchmarkCorpusRunContractCoverage.missingPromotionConditionActionKeys,
			],
			missingStopConditionActionKeys: [
				...board.benchmarkCorpusRunContractCoverage.missingStopConditionActionKeys,
			],
		},
		ownerActionExecutionContractCoverage: {
			...board.ownerActionExecutionContractCoverage,
			missingOwnerFileActionKeys: [
				...board.ownerActionExecutionContractCoverage.missingOwnerFileActionKeys,
			],
			missingCommandActionKeys: [
				...board.ownerActionExecutionContractCoverage.missingCommandActionKeys,
			],
			missingFailureEvidenceActionKeys: [
				...board.ownerActionExecutionContractCoverage.missingFailureEvidenceActionKeys,
			],
			missingAcceptanceCriteriaActionKeys: [
				...board.ownerActionExecutionContractCoverage.missingAcceptanceCriteriaActionKeys,
			],
		},
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
		topClaimOwnerActionQueue: board.topClaimOwnerActionQueue.map((row) => ({
			...row,
			evidenceWeHave: row.evidenceWeHave.map((item) => ({ ...item })),
			evidenceMissing: [...row.evidenceMissing],
			forbiddenWording: [...row.forbiddenWording],
			qssContrast: [...row.qssContrast],
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
						'Use the ranked compatibility matrix for safe wording on common workbook support, protection metadata, active content, unknown parts, encryption, malformed packages, and signature blockers.',
					nextAction:
						'Compatibility owner treats formula-binding fixtures, public LibreOffice cached-value parity, public table header/totals/current-row structured-reference fixtures, classic/x14 conditional-format fixtures, bounded chart series-source fixtures, public VBA/ActiveX/form-control preservation fixtures, and public external-link/query-table refresh metadata fixtures as accepted evidence, then adds Excel-ground-truth formula/cached-result fixtures outside the LibreOffice set, remaining lower-frequency table/conditional-format edge cases, broader chart authoring fixtures, broader public active-content/custom-UI fixtures, and broader public cross-workbook refresh fixtures from the matrix.',
					forbiddenShortcut:
						'Do not turn the matrix into full Excel compatibility, full chart editing support, worksheet/workbook/protected-range protection security, file trust, malware scanning, password recovery, or signature preservation wording.',
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
						'Benchmarking owner treats the focused ClosedXML, same-lane selected-sheet, same-lane metadata-only, current-worktree python-calamine selected-sheet runner proof, current full-profile/merged scoreboard runs, commit df349629 current selected-sheet same-lane read proof, commit 1908f3f5 current metadata-only Calamine loss, commit 67b900ed plain-text write baseline, commit e22eb86a string-heavy write baseline, commit 0d0c9632 string-heavy write optimization proof, commit 905ecb5e styles-heavy write baseline win, commit c297ba4c dense-values current repeat-15 comparable win, commits 27af69d4 and 9cab723f string-heavy current repeat-15 comparable wins, commit 2029dfab current styles-heavy repeat-15 comparable win, commit 7d61a2ef current formula-heavy comparable win, commit 183e7ebf current table-heavy comparable win, and commit eca32509 feature-rich XlsxWriter win plus OpenPyXL not-comparable boundary as accepted bounded evidence, downgrades broad speed wording, and stops production optimization unless the next work is explicit blocker resolution for ClosedXML coverage, feature-rich semantic mismatches, remaining unsupported selected-sheet/metadata-only competitors, FastXLSX environment coverage, a same-timing SDK selected-sheet open-only row, a clean multi-workload xlsx-write-sota gate, or another named public workflow loss.',
					forbiddenShortcut:
						'Do not count unavailable runners, blocked runners, dirty-worktree timings, one-workload medians, the 2000x20 plain-text/dense/string/styles rows, or focused fastest-comparable reruns as broad XLSX write/SOTA/QSS speed wins.',
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
						'Use the compatibility matrix to keep package-action wording preservation-first for protection metadata, visual, pivot, active-content, external-link, unknown-part, encrypted, malformed, and signature surfaces.',
					nextAction:
						'Correctness owner promotes an individual package surface only after real/public open -> inspect -> edit -> save/reopen -> verify evidence exists.',
					forbiddenShortcut:
						'Do not claim semantic support for every unsupported package feature, worksheet/workbook/protected-range protection as file security, byte passthrough for chart XML, signature verification, or arbitrary unknown-part preservation.',
					boundary:
						'Owner decision artifact only. It constrains package-action wording without replacing package-action proof gates.',
				},
				{
					ownerLoop: 'performance',
					artifactId: 'package-action-streaming-matrix-evidence',
					path: 'fixtures/benchmarks/package-action-proof.ts',
					validationCommand:
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					decision:
						'Use the streaming matrix as an owner decision point: current evidence is representative proof, not full streaming parity.',
					nextAction:
						'Performance owner either accepts representative passthrough/regenerate/add/drop streaming proof for narrow package-action wording, or expands generated edge/error and signature cases before any parity wording.',
					forbiddenShortcut:
						'Do not describe representative streaming cases, public macro/chart accounting, or non-streaming generated edge cases as full streaming writer parity.',
					boundary:
						'Owner decision artifact only. It routes the streaming blocker without authorizing a speed claim or package-action parity claim.',
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
		status: 'local-evidence-wording-owner-gated',
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
		status: 'local-evidence-wording-owner-gated',
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
				'No new property-generation harness until a release owner ties it to a specific public workflow claim; existing deterministic cases remain local correctness guardrails.',
			benchmark:
				'No benchmark or changed-test integration claim; shrink time and failing-case minimization stay out of scope for release wording.',
			surface:
				'Test-strategy-only evidence; no product undo, audit, rollback, SDK, CLI, API, or MCP surface.',
			validationGate:
				'Release-proof-index downgrade first; deterministic journal-law and journal-exactness tests remain regression guards only.',
			competitorContrast:
				'Property-based tests prove invariants over operation spaces instead of adding hand-written fixture examples.',
			honestBoundary:
				'Generated laws are scoped to covered operations and excluded lossy metadata/style boundaries.',
			killCriterion:
				'Do not start fast-check/shrinking work or promote broad inverse-law claims without a product-approved release workflow claim.',
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
				'No additional sidecar fixtures until product defines a repeated-scan user workflow and a release claim that sidecar evidence can answer.',
			benchmark:
				'No benchmark-loop expansion from current evidence; only after product approval would repeated scans need build cost, invalidation cost, memory overhead, and checksum parity against canonical workbook reads.',
			surface:
				'No SDK, CLI, API, or MCP product surface; existing sidecar code remains downgraded research evidence.',
			validationGate:
				'Release-proof-index stop decision first; generation-key invalidation, checksum parity, and memory caps are prerequisites only if product revives the surface.',
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
	portfolioClaim({
		rank: 11,
		name: 'research-surface-hygiene',
		claim: 'research surface as release evidence',
		northStarLink: 'Honest machine-readable proof instead of expanding notes.',
		status: 'speculative-do-not-promote',
		evidenceNeeded: claimEvidenceNeeded({
			fixture:
				'Inventory of current research files split into accepted evidence, active owner blockers, and archive-only material.',
			benchmark:
				'No benchmark claim; measure only count of unclassified research artifacts and whether each has an owner decision.',
			surface:
				'Release-proof index and owner handoff only; no SDK, CLI, API, MCP, or documentation surface.',
			validationGate:
				'Release-proof-index tests prove each promoted research item names evidence we have, evidence missing, QSS contrast, allowed wording, forbidden wording, and next owner action.',
			competitorContrast:
				'QSS comparison is blocked until research evidence is classified into release decisions instead of broad notes.',
			honestBoundary:
				'Untriaged research files are not release evidence and must not be cited for product, correctness, or performance claims.',
			killCriterion:
				'Do not promote any new research-derived claim while its source material is unclassified or lacks an owner-ready next action.',
		}),
		likelyHandoffOwner: ['product', 'release'],
		handoffDecision: 'do-not-promote-yet',
		boundary:
			'Research hygiene blocker only; classify or archive existing material before starting new research surfaces.',
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
			'Sidecar probes are downgraded research evidence; without a product-defined repeated-scan workflow they should not keep benchmark or production optimization work alive.',
		proofNeeded:
			'Product-approved repeated-scan workflow, bounded memory semantics, generation-key invalidation contract, and a release claim that sidecar evidence could answer before any renewed benchmark work.',
		killCriterion:
			'Do not promote or benchmark-expand while the surface is undefined, build plus invalidation costs are unknown, parity can fail, or memory overhead is not bounded.',
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
	{
		name: 'research-surface-hygiene',
		claim: 'research surface as release evidence',
		status: 'do-not-promote-yet',
		ownerLoops: ['product', 'release'],
		reason:
			'The research/product lane must not promote broad or dirty research material until it is classified as accepted evidence, active blocker, or archive-only.',
		proofNeeded:
			'Classify current research files into accepted evidence, active owner blockers, and archive-only material, then expose only owner-ready decisions through the release-proof index.',
		killCriterion:
			'Do not promote any research-derived claim while its source material is unclassified or lacks an owner-ready next action.',
		boundary:
			'Untriaged research files are not release evidence and must not be cited for product, correctness, or performance claims.',
	},
]

const FORMULA_LANGUAGE_SERVICE_BLOCKER = {
	ownerAction:
		'Product/correctness owner records one public SDK/CLI/API/MCP formula-assist workflow with diagnostics, reference spans, binding roles, completions, signature help, LET prepareRename success, workbook-context prepareRename refusal, and no edit-producing rename; then defines Workbook-context ownership and operation-owned edit plans before any rename wording. Validate with `bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json`, `bun test packages/sdk/src/formula-edit.test.ts --timeout 30000`, and focused CLI/API/MCP formula-assist tests.',
	allowedWording:
		'Do not promote formula-language-service-primitives as release wording today. Allowed wording: local formula-assist evidence covers diagnostics, spans, hovers, completions, signature help, binding roles, and rejection-first prepareRename only; rename remains blocked.',
	evidenceWeHave: [
		'Formula-assist proof command exists: `bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json` records sampled formula diagnostics, reference spans, binding roles, refusal counts, and claim boundaries.',
		'SDK formula-edit tests cover bundled formulaAssist, hover/code actions, reference kinds, binding roles, LET prepareRename success, and workbook-context/reference refusal in `packages/sdk/src/formula-edit.test.ts`.',
		'CLI/API/MCP formula-assist tests expose diagnostics, completions, signature help, reference edits, and workbook-context rename refusal in `apps/cli/src/cli.test.ts`, `apps/api/src/server.test.ts`, and `apps/mcp/src/index.test.ts`.',
		'API formula-assist input diagnostics are accepted as workflow polish: `0e253fe4 fix(api): structure formula assist inputs` returns a structured retryable missing-formula error instead of a generic missing-formula string in `apps/api/src/server.ts` and `apps/api/src/server.test.ts`.',
	],
	evidenceMissing: [
		'One public formula corpus workflow showing sampled formulas, spans, binding roles, refusal counts, and recovery wording across SDK/CLI/API/MCP.',
		'Workbook-context ownership and operation-owned edit plans for defined names, table names, table columns, sheet/range references, and external workbook references before any edit-producing rename.',
		'Owner-approved latency evidence for formula-assist only; no timing or practical-efficiency wording from proof shape alone.',
	],
	forbiddenWording: [
		'Do not claim edit-producing rename, safe rewrite of workbook-context references, defined-name rename, table/table-column rename, sheet/range rename, external-ref rename, or Excel-compatible formula language from formula-assist evidence.',
	],
} as const

const RELEASE_PROOF_BUNDLE_BLOCKER = {
	ownerAction:
		'Product/release owner records one public workbook workflow per top claim with inspect, plan, commit, reopen, diff, audit, compact report digest, `bun run release:rc:gate`, both compact proof commands, and `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`; keep bundle wording blocked until artifact storage, privacy filtering, canonicalization subject, and offline verification policy are approved.',
	allowedWording:
		'Do not promote release-proof-bundle as release wording today. Allowed wording: local proof plumbing can link plan, commit, reopen, diff, audit, and compact report evidence for owner review, below artifact storage, privacy, canonicalization, and offline verification policy.',
	evidenceWeHave: [
		'Local RC gate exists: `bun run release:rc:gate` proves SDK/CLI/API/MCP tarballs can be installed and exercised from a fresh temp app.',
		'Compact proof commands exist for top claims: `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json` and `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json`.',
		'Release-proof index already carries stable digests, compact report publication blockers, and top-claim owner handoffs.',
	],
	evidenceMissing: [
		'One public workbook workflow per top claim that includes inspect, plan, commit, reopen, diff, audit, and compact report digest evidence.',
		'Approved artifact storage path, retention/privacy filtering, canonicalization subject, and offline verification expectations for compact reports.',
		'Passing validation sequence: `bun run release:rc:gate`, both compact proof commands, and `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`.',
	],
	forbiddenWording: [
		'Do not call a release proof bundle signed provenance, tamper-evident storage, SLSA, in-toto, certified provenance, third-party attestation, or registry publication evidence.',
	],
} as const

const PROPERTY_JOURNAL_LAW_BLOCKER = {
	ownerAction:
		'Correctness owner downgrades property-journal-laws to permanent test-strategy-only evidence and stops fast-check/shrinking harness work from the current release evidence. Keep deterministic journal-law and journal-exactness tests as local regression guards only; restart property generation only if product first defines a public workflow claim that generated inverse laws would directly prove. Validate the downgrade with `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json` and `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`.',
	allowedWording:
		'Do not promote property-journal-laws as release wording today. Allowed wording: deterministic local journal-law tests are regression guardrails for covered exact operations and explicit lossy boundaries, not product undo, audit, rollback, or property-based testing claims.',
	evidenceWeHave: [
		'Journal-law proof command exists: `bun test fixtures/benchmarks/journal-law-proof.test.ts --timeout 30000` covers generated exact inverse-law sequences, lossy metadata boundaries, operation-family counts, issue reasons, and claim-safe markdown.',
		'Journal-law claim report exists: `bun run fixtures/benchmarks/journal-law-proof.ts --exact-cases 48 --sequence-length 5 --claim-report --json` emits exact case count, lossy boundary count, exact operation families, lossy issue reasons, do-not-promote wording, and next proof.',
		'SDK journal exactness tests cover allowed lossy issue reasons, representative exact inverse restoration, saved package-state lossiness, formula-binding lossiness, dynamic spill lossiness, table-style metadata boundaries, and journal surface classifications in `packages/sdk/src/journal-exactness.test.ts`.',
		'Claim report wording already limits the proof to deterministic local journal evidence and reports style, table-style, package-part, data-validation, and conditional-format lossy boundaries.',
		'No product owner has tied shrinkable property generation to a release workflow claim; current evidence is therefore a test-strategy downgrade, not an implementation blocker.',
	],
	evidenceMissing: [
		'Product-approved public workflow claim explaining why generated inverse laws would change release wording rather than add harness surface area.',
		'Only after product approval: shrinkable and replayable property generation with stable seeds, replay paths, minimized failing cases, operation-family coverage thresholds, and failure triage workflow.',
		'Public inverse operations or explicit release wording exclusions for style exactness, table-style exactness, package-part preservation, data-validation ordering/duplicate metadata, and conditional-format ordering/duplicate metadata.',
	],
	forbiddenWording: [
		'Do not claim property-based testing, shrinkable generated coverage, full undo coverage, exact rollback for every operation, style/table-style exactness, package-byte restoration, product audit, or signed audit/release attestation from deterministic journal-law proof.',
	],
} as const

const COLUMNAR_SIDECAR_BLOCKER = {
	ownerAction:
		'Performance owner downgrades columnar-scan-sidecars to do-not-promote and stops benchmark expansion from the current evidence. Do not add more sidecar fixtures, SDK/API/MCP surfaces, or production optimization work unless product first defines a real repeated-scan user workflow, bounded memory semantics, generation-key invalidation contract, and a release claim that sidecar evidence could answer; validate the stop decision with `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json`.',
	allowedWording:
		'Do not promote columnar-scan-sidecars as release wording today. Allowed wording: existing disposable sidecar experiments are downgraded research evidence showing checksum-parity probes on limited public ranges while the workbook grid remains source of truth.',
	evidenceWeHave: [
		'Columnar sidecar proof command exists: `bun test fixtures/benchmarks/columnar-sidecar.test.ts --timeout 30000` covers synthetic checksum parity, generation invalidation, claim-safe markdown, a tracked public fixture range, and an externally sourced public workbook range.',
		'External claim-report command exists: `bun run fixtures/benchmarks/columnar-sidecar.ts --fixture fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx --sheet "Table 9" --repeats 8 --claim-report --json` reports checksum parity and boundaries for one public numeric/date-like imported range.',
		'Claim report already labels the sidecar disposable and keeps the workbook grid as source of truth.',
		'No release owner has defined a production sidecar surface or repeated-scan user workflow; current evidence is therefore a stop decision, not a benchmark-loop expansion.',
	],
	evidenceMissing: [
		'Product-approved repeated-scan workflow that explains why a sidecar belongs in an agent-native spreadsheet runtime instead of remaining an experiment.',
		'Bounded memory semantics, generation-key invalidation contract, and source-of-truth guarantees for any real SDK/API/MCP product surface.',
		'Only after product approval: multiple larger and structurally diverse external public workbook tables/ranges plus end-to-end wins including build cost, invalidation cost, memory caps, checksum parity, and noise/CV.',
	],
	forbiddenWording: [
		'Do not claim a production cache, Arrow ABI, DuckDB integration, storage engine, workbook rewrite path, mixed-type table engine, guaranteed acceleration, QSS/SOTA speed win, SDK/API/MCP sidecar product surface, or benchmark-backed production optimization target from current columnar sidecar evidence.',
	],
} as const

const FORMULA_ORACLE_ROUTING_BLOCKER = {
	ownerAction:
		'Formula/Calc correctness owner keeps formula-oracle-routing out of release wording, reruns `bun test fixtures/benchmarks/formula-corpus-correctness.test.ts fixtures/xlsx/libreoffice-fixtures.test.ts --timeout 60000` and the full public LibreOffice command `bun run fixtures/benchmarks/formula-corpus-correctness.ts --corpus-root fixtures/xlsx/libreoffice --manifest fixtures/xlsx/libreoffice/manifest.ts --tag formula-fidelity --max-mismatches 23 --max-unaccepted-mismatches 0 --max-semantic-mismatches 0 --max-errors 0 --min-workbooks 34 --min-formulas 418 --min-compared-formulas 402 --min-semantic-perfect-workbooks 34 --json`, then adds public oracle-class fixtures or real HyperFormula/LibreOffice/Excel/static-golden adapters that make non-accepted route counts, skip counters, divergence counters, and verifier artifacts reproducible before any Excel-compatible formula wording.',
	allowedWording:
		'Do not promote formula-oracle-routing as release wording today. Allowed wording: public cached-value corpus routing can report sampled formula counts, skips, mismatch classes, and oracle gaps; it does not prove Excel-compatible formulas.',
	evidenceWeHave: [
		'Formula corpus correctness tests cover TypeScript corpus manifests, cached-value comparisons, date-system-sensitive formulas, control-character strings, mismatch references, no-cached-value skips, volatile oracle skips, accepted numeric drift, stale oracle routing, assertion gates, and CLI threshold gates in `fixtures/benchmarks/formula-corpus-correctness.test.ts`.',
		'Public LibreOffice fixture gate in `fixtures/xlsx/libreoffice-fixtures.test.ts` over `fixtures/xlsx/libreoffice/manifest.ts` pins 34 workbooks, 418 formulas, 402 compared formulas, 23 accepted mismatches, 0 unaccepted mismatches, 0 semantic mismatches, 22 numeric-drift mismatches, 1 stale-oracle mismatch, 32 perfect workbooks, and 34 semantic-perfect workbooks.',
		'Strict zero-mismatch release wording is currently blocked by real public output: `formula corpus correctness failed: mismatches 23 exceeded 0; accepted mismatches 23 exceeded 0` from the full LibreOffice formula-fidelity corpus with zero-mismatch thresholds.',
		'HyperFormula comparator smoke tests exist in `fixtures/benchmarks/formula-sota.test.ts`, but they are performance/correctness comparator evidence, not a formula-corpus oracle replacement.',
	],
	evidenceMissing: [
		'Runnable public corpus artifacts for cached-only, volatile, numeric drift, unsupported function, external refs, dynamic arrays, structured refs, and date-system mismatch classes with stable expected route counts.',
		'Actual HyperFormula, LibreOffice, Excel, and static-golden oracle adapters that emit skip counters, divergence counters, oracle artifacts, and reproducible failure output without private corpora; the current public LibreOffice route counts exercise accepted-mismatch only, not real adapter execution.',
		'Owner-approved thresholds and artifact verifier before changing formula compatibility, cached-value freshness, or QSS/SOTA wording.',
	],
	forbiddenWording: [
		'Do not claim Excel-compatible formulas, full formula parity, complete oracle automation, fresh cached values, zero mismatches, QSS/SOTA formula superiority, or HyperFormula/LibreOffice/Excel oracle execution from cached-value routing evidence alone.',
	],
} as const

const TOKEN_BOUNDED_AGENT_VIEW_BLOCKER = {
	ownerAction:
		'Product owner records one public end-to-end agent-view example that starts from a strict `maxApproxTokens` request, shows omitted sample-row, column-sample, and formula-pattern locators, then recovers omitted evidence through narrower reads or an unbudgeted same-range view; validate with `bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts` plus the focused SDK/CLI/API/MCP agent-view token-budget tests before any release wording.',
	allowedWording:
		'Do not promote token-bounded-agent-view as release wording today. Allowed wording: deterministic compact views can expose approximate token budgets, structural floors, omitted-evidence counters, and recovery locators for owner review.',
	evidenceWeHave: [
		'Budget proof command exists: `bun test fixtures/benchmarks/agent-view-budget-proof.test.ts` covers dense-table, wide-sparse, formula-heavy, metadata-heavy, and public-formula-stress cases for deterministic budget metadata, shape preservation, and counted omissions.',
		'Recovery proof command exists: `bun test fixtures/benchmarks/agent-view-recovery-proof.test.ts` proves same-range unbudgeted recovery, compact omitted-evidence locators, narrow sample-row recovery, and formula-pattern example recovery.',
		'Committed SDK/CLI/API/MCP tests expose budget metadata through `packages/sdk/src/sdk.test.ts`, `apps/cli/src/cli.test.ts`, `apps/api/src/server.test.ts`, and `apps/mcp/src/index.test.ts`.',
	],
	evidenceMissing: [
		'One public product example that demonstrates an agent using omitted-evidence locators to recover missing rows, column samples, and formula-pattern examples instead of trusting the compact view as complete.',
		'Owner-approved wording for approximate token estimates, structural floor behavior, omitted-evidence recovery, and when an unbudgeted read is required.',
		'Focused cross-surface validation commands covering `fixtures/benchmarks/agent-view-budget-proof.test.ts`, `fixtures/benchmarks/agent-view-recovery-proof.test.ts`, and the SDK/CLI/API/MCP agent-view token-budget tests without unrelated API server coverage.',
	],
	forbiddenWording: [
		'Do not claim exact model-token counts, complete workbook context under every budget, hidden summarization, or automatic recovery of omitted evidence.',
	],
} as const

const RETAINED_VIEWPORT_PATCH_BLOCKER = {
	ownerAction:
		'Product/performance owner records one public SDK/API/MCP compact-read workflow with retained patch, skipped token, invalid token, expired history, projection-change invalidation, metadata invalidation, patch bytes, retention cap, and recovery action; validate with `bun test fixtures/benchmarks/viewport-patch-proof.test.ts`, the focused SDK interactive viewport patch tests, and the focused API/MCP compact changedSince tests before any retained-history wording.',
	allowedWording:
		'Do not promote retained-viewport-patch-history as release wording today. Allowed wording: SDK/API/MCP proof can describe bounded single-session retained patches, invalidation reasons, and fresh-window recovery; CLI and collaboration claims stay excluded.',
	evidenceWeHave: [
		'Viewport proof command exists: `bun test fixtures/benchmarks/viewport-patch-proof.test.ts` covers retained patches, skipped retained tokens, invalid tokens, cross-session tokens, expired history, projection changes, and metadata invalidation.',
		'SDK interactive contract tests cover viewport patch helpers and invalidation reasons in `packages/sdk/src/interactive-contract.test.ts`.',
		'API and MCP compact `changedSince` tests cover empty patches, projection changes, invalid tokens, selected-column changes, and changed-source invalidations in `apps/api/src/server.test.ts` and `apps/mcp/src/index.test.ts`.',
	],
	evidenceMissing: [
		'One public workflow showing patch bytes, retained history size or cap, invalidation reasons, and recovery action across SDK/API/MCP without CLI claims.',
		'Owner-approved product wording for bounded single-session history, invalidation reasons, retention caps, and fresh-window recovery.',
		'Performance-approved patch-size and invalidation-rate evidence before any practical latency or efficiency wording.',
	],
	forbiddenWording: [
		'Do not claim collaboration, sync, CRDT, multi-writer convergence, transaction isolation, unlimited history, or a signed audit trail from viewport patch evidence.',
	],
} as const

const AGENT_WORKFLOW_OBSERVABILITY_BLOCKER = {
	ownerAction:
		'Product owner keeps agent-workflow-observability out of release wording and stops adding more prepared-plan smoke/failure-recovery plumbing from the current evidence. Next work is one public inspect/plan/commit/reopen/diff/audit/trace/repair workflow with failure taxonomy, trace payload size, compact/redacted artifact behavior, recovery prompts, and proof that trace output changes the next repair or audit decision. Own the workflow proof in `packages/sdk/src/agent-workflow.test.ts`, `apps/cli/src/cli.test.ts`, `apps/api/src/server.test.ts`, and `apps/mcp/src/index.test.ts`; validate with `bun test packages/sdk/src/agent-workflow.test.ts -t "keeps unresolved external-link package graph failures visible to agents|release proof bundle links plan, commit, reopen, diff, and audit evidence|prepared agent commits surface post-write audit failures as blocking model output" --timeout 30000`, `bun test apps/cli/src/cli.test.ts -t "plan invalid ops return structured batch repair details|check surfaces structured issue metadata for agent repair|trace shows values and respects max depth|trace shows precedents for formula cell" --timeout 30000`, `bun test apps/api/src/server.test.ts -t "trace returns structured partial-load diagnostics for capped formula views|plan invalid ops return structured batch repair details|prepared path mutation handles surface post-write audit failures as blocked output|direct path mutation commits surface post-write audit failures as blocked output" --timeout 30000`, `bun test apps/mcp/src/index.test.ts -t "ascend.trace returns structured partial-load diagnostics for capped formula views|ascend.plan invalid ops return structured batch repair details|prepared MCP path mutation handles surface post-write audit failures as blocked output|direct MCP path mutation commits surface post-write audit failures as blocked output" --timeout 30000`, and `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` before any observability wording.',
	allowedWording:
		'Do not promote agent-workflow-observability as release wording today. Allowed wording: existing trace, repair, prepared-plan failure-recovery, and installed SDK prepared-plan smoke surfaces are internal workflow evidence until a public workflow proves decision improvement.',
	evidenceWeHave: [
		'SDK agent workflow tests cover package graph visibility, post-write audit blocking, compact trace artifact counts, and release proof bundle links in `packages/sdk/src/agent-workflow.test.ts`.',
		'Prepared-plan failure recovery is accepted as internal workflow proof: `dfa0202b test(sdk): prove prepared write rollback`, `b35a8157 test(api): prove prepared write failure recovery`, and `786e886c fix(api): classify export failures as conflicts` cover rollback/retry semantics and structured export failure classification across SDK/API/MCP paths.',
		'Installed SDK prepared-plan smoke evidence is accepted as internal workflow proof: `568e32cd test(release): smoke sdk prepared plans` exercises `createPreparedAgentPlan` from packaged `@ascend/sdk`, plan digest/input hash guards, commit, reopen/check, post-write audits, and formula recalculation in `scripts/release-sdk-smoke.ts`.',
		'CLI trace and repair tests expose trace depth, formula precedents, structured batch repair details, and check metadata for agent repair in `apps/cli/src/cli.test.ts`.',
		'API and MCP tests expose capped formula-view trace diagnostics, structured repair details, compact trace artifact counts, and blocked post-write audit output in `apps/api/src/server.test.ts` and `apps/mcp/src/index.test.ts`.',
		'API repair-plan input diagnostics are accepted as internal workflow proof: `2cb02045 fix(api): structure repair plan inputs` returns a structured retryable missing-workbook-reference error for `/repair-plan` instead of a generic missing-file response.',
		'Runnable safe-edit examples are accepted as local workflow proof: `cab4bff1 test(examples): prove agent safe edit workflow` covers SDK inspect/plan/commit/reopen/check/lint verification, `01d08512 test(examples): add runnable api safe edit workflow` covers the same generated-workbook path through HTTP API calls, `de45eb83 test(examples): add runnable mcp safe edit workflow` covers it through MCP tools, `bbea493c test(examples): add package safe edit scripts` proves all three workflows run from examples package scripts, and `a09660be feat(cli): add root safe edit workflow scripts` proves the same workflows through root package scripts.',
		'Agent-doc search indexes runnable HTTP and MCP safe-edit workflows through `2569626c fix(sdk): index runnable http workflow example` and `de45eb83 test(examples): add runnable mcp safe edit workflow`, making outside-user workflow examples discoverable from the SDK docs surface.',
		'CLI agent-init surfaces runnable SDK, HTTP, and MCP safe-edit example commands through `0d1b33f7 feat(cli): surface runnable agent workflow examples`, so a CLI user can find the same outside-user workflows without knowing repo paths.',
		'API and MCP workflow discovery surfaces expose runnable SDK, HTTP, and MCP safe-edit example commands through `407499cd feat(api): expose runnable workflow examples`, and installed app smoke checks guard those example fields.',
		'Workflow discovery surfaces expose runnable-example proof context through `f3347e17 feat(api): expose workflow example proof context`, including repository-root workdir, setup prerequisites, and `bun test examples/root-scripts.test.ts` as the proof command.',
		'Packaged proof surfaces expose workflow examples through `58a7f579 test(release): expose workflow examples in packaged proof`, including MCP agent-workflow resource text and installed smoke output.',
		'Installed SDK safe-edit workflow proof is accepted after `5915794f feat(sdk): summarize installed safe edit proof` and `56cd4aa0 feat(sdk): add safe edit package bin`: `bun test examples/package-install-safe-edit.test.ts --timeout 30000` now proves safeToUse, changed-cell before/after evidence, gate outcomes, hashes, check/lint validity, post-write audit validity, and the outside-user `node_modules/.bin/ascend-sdk-safe-edit` command from the installed-SDK example. This closes the earlier `5981764c` blank/empty proof contract blocker for generated-workbook wording.',
		'Reusable SDK workflow proof summary is accepted after `a5fa3006 feat(sdk): expose workflow proof summary`: `bun test packages/sdk/src/agent-workflow.test.ts -t "workflow proof summary explains changed cells and safety gates" --timeout 30000` proves changed-cell before/after evidence plus plan-linked, plan, write-policy, commit, reopen-verify, and package-graph safety gates, and the installed SDK example now uses that helper.',
		'Root SDK safe-edit proof-bundle output is accepted after `adc9c1e1 feat(examples): show sdk workflow proof bundle`: `bun test examples/agent-safe-edit.test.ts --timeout 30000`, `bun test examples/package-scripts.test.ts --timeout 30000`, and `bun test examples/root-scripts.test.ts --timeout 30000` prove the runnable SDK example emits proofBundle.safeToUse, changed-cell evidence, and the shared whySafe gates from package and root commands.',
		'Installed safe-edit workflow discovery is accepted after `f8d63593 feat(apps): expose installed safe edit workflow`, `cc689bcc test(api): prove installed workflow discovery`, `56cd4aa0 feat(sdk): add safe edit package bin`, `91dabea8 feat(apps): expose connection proof workflow field`, `4d272f77 feat(apps): expose visual proof workflow field`, and `a2960803 feat(apps): expose formula proof workflow field`: CLI, API, MCP, MCP resources, API package tests, and installed app smoke output expose the installed SDK package-bin command plus proofBundle, postWrite.dataConnections, postWrite.visuals, and postWrite.formulaState output contracts. The full `bun test apps/api/api.test.ts --timeout 30000` sweep remains blocked by an unrelated export-format message expectation, so package health wording is not accepted.',
		'Installed CLI safe-edit workflow proof is accepted after `3d630232 feat(cli): add packaged safe edit workflow` and `7e7183df fix(cli): share safe edit proof summary`: `bun test apps/cli/src/cli.test.ts -t "example-safe-edit runs the packaged inspect plan commit reopen verify workflow" --timeout 30000` proves `ascend example-safe-edit <file.xlsx> <out.xlsx>` runs inspect, plan, commit, reopen, and verify with the shared SDK proof summary, `proofBundle.safeToUse`, changed-cell evidence, plan-linked/write-policy/reopen-verify/package-graph gates, and reopened check/lint output.',
		'CLI commit proof-bundle output is accepted after `5028438e feat(cli): summarize commit proof bundle` and `223a1ec7 fix(cli): align proof gate naming`: `bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow" --timeout 30000` proves `ascend commit --proof --json` and compact `--proof` output include safeToUse, whatChanged, whySafe, input-guard/write-policy/commit/reopen/package-graph gates, and output hashes.',
		'API and MCP commit proof-bundle output is accepted after `1ed2be29 feat(apps): expose commit proof bundles`: `bun test apps/api/src/server.test.ts -t "dump emits replayable operation batches" --timeout 30000` and `bun test apps/mcp/src/index.test.ts -t "ascend.commit accepts prepared plan handles" --timeout 30000` prove prepared commit responses can include proofBundle.safeToUse, whatChanged, whySafe, and write-policy/commit/reopen/package-graph gates.',
		'API and MCP runnable example proof bundles are accepted after `7afcd630 feat(examples): show api mcp proof bundles`: `bun test examples/package-scripts.test.ts --timeout 30000` and `bun test examples/root-scripts.test.ts --timeout 30000` prove HTTP and MCP examples expose proofBundle.safeToUse and changed-cell evidence from package and root commands.',
		'API and MCP trust-preflight workflow proof is accepted after `a8e15d9b feat(apps): add trust preflight workflows`: `bun test apps/api/src/server.test.ts -t "/agent-workflow exposes the API safe edit contract" --timeout 30000`, `bun test apps/mcp/src/index.test.ts -t "ascend.agent_workflow exposes machine-readable safe edit guidance" --timeout 30000`, `bun test examples/agent-safe-edit-http.test.ts examples/agent-safe-edit-mcp.test.ts --timeout 30000`, `bun test examples/package-scripts.test.ts --timeout 30000`, and `bun test examples/root-scripts.test.ts --timeout 30000` prove API/MCP workflow discovery and runnable examples include a trust-preflight step before inspect/read.',
		'Safe-edit post-write proof output is accepted after `b93f3493 test(release): surface safe-edit post-write proof`: `bun test examples/agent-safe-edit.test.ts examples/agent-safe-edit-http.test.ts examples/agent-safe-edit-mcp.test.ts examples/package-scripts.test.ts examples/root-scripts.test.ts --timeout 30000` proves SDK, HTTP, MCP, package, and root generated-workbook safe-edit outputs include reopened-output dataConnections, formulaState, security, and visuals proof fields instead of hiding post-write proof behind the commit summary.',
		'Post-write data-connection proof is accepted after `caa08959 fix(sdk): report post-write data connections`, `62566e09 fix(sdk): report connection proof metadata`, and `1cb093fe feat(sdk): edit workbook connection scheduling metadata`: `bun test fixtures/corpus/external-refresh-contract.test.ts -t "commit proof reports reopened public query-table connection metadata" --timeout 30000`, `bun test packages/io-xlsx/src/reader/connections.test.ts packages/sdk/src/connection-inventory.test.ts fixtures/corpus/external-refresh-contract.test.ts -t "connection part inventory|connection SDK inventory|reports and preserves public query-table refresh surfaces without executing them|commit proof reports reopened public query-table connection metadata" --timeout 60000`, and `bun test packages/sdk/src/agent-workflow.test.ts -t "commits public workbook connection scheduling edits through save and reopen" --timeout 30000` prove approved public query-table/data-connection commits return reopened connection counts, states, names, ids, part paths, type/description/source/command/refresh flags, `hasConnectionString`, and workbook connection scheduling metadata edits without exposing connection-string contents or executing connections in `postWrite.dataConnections`.',
		'Post-write formula proof is accepted after `f7338c91 test(corpus): prove post-write formula cache reporting` and `019f457e fix(sdk): report cached formula proof locations`: `bun test fixtures/corpus/formula-binding-contract.test.ts -t "commit proof reports missing public formula caches after save and reopen" --timeout 30000` proves public formula commits report calcChain state, formula cache state, cached/missing counts, cached value locations, missing cached locations, and warnings in `postWrite.formulaState` without claiming fresh recalculation.',
		'Workflow docs list inspect, plan, commit, verify, trace, and repair-plan recovery paths in `docs/AGENT_WORKFLOW.md`, but documentation is guidance rather than release proof.',
	],
	evidenceMissing: [
		'One public workbook workflow showing inspect, plan, commit, reopen, diff, audit, trace, and repair-plan output with a failure taxonomy and recovery prompts.',
		'Trace payload size, compact/redacted artifact behavior, failure-class coverage, and evidence that trace output changes a concrete repair or audit decision.',
		'Owner-approved decision that prepared-plan rollback/retry plus installed-consumer smoke evidence is enough for prepared-plan plumbing and should not continue as low-value API/SDK/MCP work without a new public workflow failure class.',
		'Owner-approved decision whether CLI/API/MCP proofBundle gates are enough for safe workflow wording or whether the next release-code block must add a single public workbook workflow tying proofBundle output to diff, trace, and repair-plan decisions.',
		'Golden trace fixtures, redaction/privacy checks, and owner-approved recovery wording before publishing observability language.',
	],
	forbiddenWording: [
		'Do not claim autonomous correctness, complete observability, signed audit trail, repair automation, root-cause diagnosis, privacy-safe redaction, registry-published CLI safe-edit availability, or that traces/proofBundle gates alone prove workbook safety from current workflow evidence.',
	],
} as const

const PRACTICAL_LATENCY_CONTRACTS_BLOCKER = {
	ownerAction:
		'Performance owner keeps practical-latency-contracts out of release wording, reruns `bun test fixtures/benchmarks/practical-latency-contracts.test.ts --timeout 30000`, dry-runs `bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 1 --warmup 0 --dry-run --json`, then runs the real public-tracked contract `bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 3 --warmup 1 --json` from a tracked-clean worktree; promote only if summary/profile artifacts name first-view, edit-verify, and repeated-inspection envelopes, public/generated input provenance, median/p95/CV, memory or payload guardrails, one profile-backed production target, and product-approved non-threshold wording.',
	allowedWording:
		'Do not promote practical-latency-contracts as release wording today. Allowed wording: public-tracked latency contracts are benchmark-owner diagnostics for first-view, edit-verify, and repeated-inspection envelopes until a tracked-clean run and non-threshold wording are approved.',
	evidenceWeHave: [
		'Practical latency contract tests cover the public-tracked preset, generated edit-input labeling, envelope target selection, noisy-above-floor target handling, and hot-cache guardrails in `fixtures/benchmarks/practical-latency-contracts.test.ts`.',
		'The benchmark emits summary JSON, markdown, profile-summary markdown, input provenance, worktree guardrails, user-visible envelope decisions, diagnostic ceilings, memory/payload guardrails, and profile commands from `fixtures/benchmarks/practical-latency-contracts.ts`.',
		'Release-proof index already treats the existing practical-latency report as excluded diagnostic evidence rather than a release proof artifact.',
	],
	evidenceMissing: [
		'Tracked-clean public-tracked run of `bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 3 --warmup 1 --json` with no tracked dirty files.',
		'Published summary/profile artifacts showing first-view, edit-verify, and repeated-inspection envelopes with median, p95, CV/noise, memory or payload guardrails, and one profile-backed production target or explicit no-target decision.',
		'Product/performance approval that any wording is non-threshold, public-input scoped, and not a QSS/SOTA speed claim.',
	],
	qssContrast: [
		'QSS can sell visible spreadsheet responsiveness from an integrated product; Ascend must prove public-input first-view, edit-verify, and repeated-inspection envelopes instead of internal phase timings.',
	],
	forbiddenWording: [
		'Do not claim release latency, SLA, fastest XLSX reader, QSS/SOTA speed win, production optimization target, or user-visible workflow improvement from dirty-worktree, dry-run, private-input, local one-off, diagnostic-ceiling, or hot-cache-only timing evidence.',
	],
} as const

const RESEARCH_SURFACE_HYGIENE_BLOCKER = {
	ownerAction:
		'Product/release owner runs `git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager` and `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json`, reviews `classifiedEntries`, resolves active-owner-blocker rows for the local Excel corpus and loop-manager state, keeps archive-only rows out of release wording, and reruns `bun test fixtures/benchmarks/release-proof-index.test.ts` before citing any research-derived claim.',
	allowedWording:
		'Do not promote research-surface-hygiene as release wording today. Allowed wording: current research/tmp material is routed as accepted-evidence, active-owner-blocker, or archive-only, and raw research paths remain non-citeable for release claims.',
	evidenceWeHave: [
		'`bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json` classifies the current visible research inventory into accepted-evidence, active-owner-blocker, and archive-only rows.',
		'The current fixture-decision research entry is routed to accepted evidence only through the compact fixture-decision proof packet; raw experiment notes remain non-citeable.',
		'The local Excel corpus (`research/excel-corpus/`) plus loop-manager script/state (`scripts/ascend-loop-manager.ts`, `tmp/ascend-loop-manager/`) are active owner blockers; broad research notes are archive-only unless a future owner converts one finding into release-proof evidence.',
	],
	evidenceMissing: [
		'Product/release owner review of `classifiedEntries` from `git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager` and `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json`, including whether the local Excel corpus becomes approved public fixtures, private/local-only diagnostics, or archive/deletion candidates.',
		'Release owner decision for whether `scripts/ascend-loop-manager.ts` and `tmp/ascend-loop-manager/` should be tracked, ignored, or archived as operational steering state.',
		'Passing `bun test fixtures/benchmarks/release-proof-index.test.ts` proving any promoted research item has evidence we have, evidence missing, QSS contrast, allowed wording, forbidden wording, and next owner action.',
	],
	forbiddenWording: [
		'Do not cite `research/` or `tmp/` files as product, correctness, performance, QSS, or release evidence until they are classified and routed through the release-proof index.',
	],
} as const

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
				'VBA project signature parts are dropped on approved edited exports while the VBA project remains inventoried',
				'values-only text export from signed macro workbooks requires both signature-loss and active-content-loss approvals',
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
				'high-risk package corpus contract proves unknown package parts route to review and mutation stays fail-closed',
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
				'high-risk-package-contract/encrypted-password',
			],
			proofChecks: [
				'public encrypted Calamine fixture is used',
				'valid password opens with full-mode planning',
				'missing or wrong password rejects before hydration',
				'high-risk package corpus contract proves no-op encrypted bytes stay encrypted and edited encrypted export fails closed unless decrypted plain XLSX output is explicitly requested',
				'SDK session packageGraph/rawPackagePart inspect decrypted package bytes after password open, including metadata-only sessions',
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
		rejectedFixtures: [...safeOpen.rejectedFixtures],
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
					'Safe-open proof already includes a generated malformed package rejection path. The tracked scan reports one rejected fixture, `fixtures/xlsx/calamine/pass_protected.xlsx`, but that is an encrypted/password fixture and not a public malformed-package replacement.',
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
			rejectedFixtures: [...evidence.safeOpen.rejectedFixtures],
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
			replacementEvidence: `tracked safe-open scan rejected ${fixtureEvidence.safeOpen.rejected} fixture(s): ${fixtureEvidence.safeOpen.rejectedFixtures.join(', ') || 'none'}; none are accepted public malformed-package replacements, so malformed proof input remains generated structural bytes`,
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
	const correctnessBoundaryJson = process.argv.includes('--correctness-boundary-json')
	const performanceBoundaryJson = process.argv.includes('--performance-boundary-json')
	const researchHygieneJson = process.argv.includes('--research-hygiene-json')
	const result = await runReleaseProofIndex({
		includeTimings: !process.argv.includes('--no-timings'),
	})
	console.log(
		releaseDecisionJson
			? JSON.stringify(result.releaseDecisionBoard, null, 2)
			: performanceBoundaryJson
				? JSON.stringify(releaseProofPerformanceBoundaryDecisionPacket(result), null, 2)
				: correctnessBoundaryJson
					? JSON.stringify(releaseProofCorrectnessBoundaryDecisionPacket(result), null, 2)
					: researchHygieneJson
						? JSON.stringify(releaseProofResearchHygieneDecisionPacket(result), null, 2)
						: fixtureDecisionJson
							? JSON.stringify(releaseProofFixtureDecisionPacket(result), null, 2)
							: ownerHandoffsJson
								? JSON.stringify(releaseProofOwnerHandoffIndex(result), null, 2)
								: json
									? JSON.stringify(result, null, 2)
									: releaseProofIndexMarkdown(result),
	)
	if (
		!json &&
		!ownerHandoffsJson &&
		!releaseDecisionJson &&
		!fixtureDecisionJson &&
		!correctnessBoundaryJson &&
		!performanceBoundaryJson &&
		!researchHygieneJson
	) {
		console.error(`Indexed ${result.artifactCount} release proof evidence artifacts.`)
		console.error(
			`Run with --json, --owner-handoffs-json, --release-decision-json, --fixture-decision-json, --correctness-boundary-json, --performance-boundary-json, or --research-hygiene-json for machine-readable output from ${basename(import.meta.path)}.`,
		)
	}
}
