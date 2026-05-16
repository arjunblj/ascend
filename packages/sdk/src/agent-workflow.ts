import { createHash } from 'node:crypto'
import { copyFile, readFile, rename, stat, unlink, writeFile } from 'node:fs/promises'
import { dirname, extname, join, resolve } from 'node:path'
import {
	type ActiveContentInfo,
	type ChartPartInfo,
	type ExternalReferenceInfo,
	indexToColumn,
	parseA1Safe,
	parseRange,
	type RangeRef,
	type SheetDrawingObjectRef,
	type SheetImageRef,
	type SheetState,
	type Workbook,
} from '@ascend/core'
import { normalizeFormulaInput, parseFormula } from '@ascend/formulas'
import type {
	XlsxPackageGraph,
	XlsxPackageGraphFidelityIssue,
	XlsxPackageLossPolicy,
	XlsxPackageOwnerScope,
} from '@ascend/io-xlsx'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
	extractZip,
	inspectXlsxPackageGraph,
} from '@ascend/io-xlsx'
import {
	AscendException,
	ascendError,
	type CalcSettings,
	type FeatureReport,
	type Operation,
} from '@ascend/schema'
import { listCapabilities, summarizeCapabilities } from './capabilities.ts'
import { collectFormulaReferences } from './formula-info.ts'
import {
	MUTATION_JOURNAL_ISSUE_SCHEMA,
	MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
	type MutationJournal,
	type MutationJournalIssue,
} from './journal.ts'
import { WorkbookDocument, type WorkbookLoadOptions } from './session.ts'
import type { CheckIssue, FormulaReferenceInfo } from './types.ts'
import { AscendWorkbook } from './workbook.ts'

type WorkbookWritePlanInfo = ReturnType<AscendWorkbook['writePlanSummary']>
type WorkbookWritePlanPartInfo = WorkbookWritePlanInfo['parts'][number]
type PackageActionProofArchive = ReturnType<typeof extractZip>

export interface AgentPlanResult {
	readonly file: string
	readonly inputSha256: string
	readonly operationCount: number
	readonly planDigest: string
	readonly needsApproval: boolean
	readonly approvals: readonly ApprovalRequirement[]
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly preview: ReturnType<AscendWorkbook['preview']>
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly writePolicy: WritePolicyReport
	readonly unsupportedFeatures: readonly unknown[]
	readonly lossAudit: LossAudit
	readonly packageGraphAudit: PackageGraphAudit
	readonly capabilities: ReturnType<typeof summarizeCapabilities>
}

export interface CompactAgentPlanOptions {
	readonly maxChangedCells?: number
}

export interface CompactAgentCommitOptions {
	readonly maxAffectedCells?: number
}

export interface CompactAgentPreview {
	readonly wouldSucceed: boolean
	readonly changedCellCount: number
	readonly emittedChangedCellCount: number
	readonly changedCells: ReturnType<AscendWorkbook['preview']>['changedCells']
	readonly changedRanges: readonly CompactAffectedRange[]
	readonly recalcScope: number
	readonly warningCount: number
	readonly errorCount: number
	readonly warnings: ReturnType<AscendWorkbook['preview']>['warnings']
	readonly errors: ReturnType<AscendWorkbook['preview']>['errors']
	readonly journalSummary?: CompactJournalSummary
}

export interface CompactAgentPlanResult extends Omit<AgentPlanResult, 'preview'> {
	readonly preview: CompactAgentPreview
}

export interface PreparedAgentPlan {
	readonly file: string
	readonly inputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly plan: AgentPlanResult
	commit(options?: AgentCommitOptions): Promise<AgentCommitResult>
}

export interface AgentCommitOptions {
	readonly output?: string
	readonly inPlace?: boolean
	readonly backup?: string
	readonly expectSha256?: string
	readonly password?: string
	readonly allowDecryptedExport?: boolean
	readonly allowLoss?: readonly string[] | 'all'
	readonly approvals?: readonly string[] | 'all'
	readonly onProgress?: AgentWorkflowProgressHandler
}

export interface AgentPlanOptions {
	readonly password?: string
	readonly onProgress?: AgentWorkflowProgressHandler
}

export interface AgentCommitResult {
	readonly file: string
	readonly output: string
	readonly backup?: string
	readonly inputSha256: string
	readonly outputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly approvals: readonly ApprovalRequirement[]
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly apply: ReturnType<AscendWorkbook['apply']>
	readonly recalc: ReturnType<AscendWorkbook['recalc']> | null
	readonly timings: AgentCommitTimings
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly writePolicy: WritePolicyReport
	readonly postWrite: AgentPostWriteVerification
	readonly lossAudit: LossAudit
	readonly packageGraphAudit: PackageGraphAudit
}

export interface AgentWorkflowProofGate {
	readonly gate: string
	readonly ok: boolean
	readonly evidence: Record<string, unknown>
}

export interface AgentWorkflowProofChangedCell {
	readonly ref: string
	readonly before: AgentPlanResult['preview']['cellChanges'][number]['before']
	readonly after: AgentPlanResult['preview']['cellChanges'][number]['after']
	readonly formulaBefore: AgentPlanResult['preview']['cellChanges'][number]['formulaBefore']
	readonly formulaAfter: AgentPlanResult['preview']['cellChanges'][number]['formulaAfter']
}

export interface AgentWorkflowProofSummaryOptions {
	readonly preflightGates?: readonly AgentWorkflowProofGate[]
	readonly defaultSheetName?: string
}

export interface AgentWorkflowProofSummary {
	readonly safeToUse: boolean
	readonly whatChanged: readonly AgentWorkflowProofChangedCell[]
	readonly whySafe: readonly AgentWorkflowProofGate[]
	readonly evidence: {
		readonly inputSha256: string
		readonly planDigest: string
		readonly outputSha256: string
		readonly operationCount: number
		readonly changedCellCount: number
		readonly postWriteValid: boolean
		readonly auditsPassed: boolean
		readonly reopened: boolean
		readonly checkValid: boolean
		readonly lintClean: boolean
		readonly writePolicyOk: boolean
		readonly packageGraphAuditOk: boolean
	}
}

export interface ReleaseProofDiffEvidence {
	readonly sheetDiffCount: number
	readonly changedSheets?: readonly string[]
}

export interface ReleaseProofBundleOptions {
	readonly diff?: ReleaseProofDiffEvidence
	readonly sourceBytes?: Uint8Array
	readonly outputBytes?: Uint8Array
	readonly claimBoundaries?: readonly string[]
}

export interface ReleaseProofArtifactDigest {
	readonly name: string
	readonly digest: string
	readonly summary: string
}

export interface ReleaseProofConsistencyCheck {
	readonly check: string
	readonly ok: boolean
	readonly details?: unknown
}

export type PackageActionKind = 'passthrough' | 'regenerate' | 'add' | 'drop' | 'error'

export interface PackageActionProofOptions {
	readonly sourcePackageGraph?: XlsxPackageGraph
	readonly sourceBytes?: Uint8Array
	readonly outputBytes?: Uint8Array
	readonly writePolicy?: WritePolicyReport
	readonly packageGraphAudit?: PackageGraphAudit
	readonly claimBoundaries?: readonly string[]
}

export type AgentCommitPackageActionProofOptions = Omit<
	PackageActionProofOptions,
	'sourceBytes' | 'outputBytes' | 'writePolicy' | 'packageGraphAudit'
>

export interface PackageActionProof {
	readonly formatVersion: 1
	readonly kind: 'ascend-package-action-proof'
	readonly totalActions: number
	readonly byAction: Readonly<Record<PackageActionKind, number>>
	readonly coverage: PackageActionProofCoverage
	readonly actions: readonly PackageActionProofEntry[]
	readonly issues: readonly string[]
	readonly claimBoundaries: readonly string[]
}

export interface PackageActionProofCoverage {
	readonly proofScope: 'package-part-actions-with-audit-summaries'
	readonly sourceGraphIncluded: boolean
	readonly sourcePartCount?: number
	readonly sourceRelationshipCount?: number
	readonly writePolicyIncluded: boolean
	readonly packageGraphAuditIncluded: boolean
	readonly relationshipAuditIssueCount: number
	readonly bytePreservationAuditIssueCount: number
	readonly sourceByteDigestCount: number
	readonly outputByteDigestCount: number
	readonly matchingByteDigestCount: number
	readonly mismatchedByteDigestCount: number
}

export interface PackageActionProofEntry {
	readonly action: PackageActionKind
	readonly reason: string
	readonly partPath?: string
	readonly sourcePresent?: boolean
	readonly origin?: WorkbookWritePlanPartInfo['origin']
	readonly owner?: WorkbookWritePlanPartInfo['owner']
	readonly contentType?: string
	readonly streaming?: boolean
	readonly sourceSha256?: string
	readonly outputSha256?: string
	readonly bytesEqual?: boolean
	readonly diagnosticCodes?: readonly WritePolicyDiagnostic['code'][]
	readonly auditIssueCodes?: readonly XlsxPackageGraphFidelityIssue['code'][]
}

export interface ReleaseProofBundle {
	readonly formatVersion: 1
	readonly kind: 'ascend-release-proof-bundle'
	readonly proofKind: 'local-evidence'
	readonly subject: {
		readonly file: string
		readonly output: string
		readonly inputSha256: string
		readonly outputSha256: string
		readonly planDigest: string
	}
	readonly operations: {
		readonly count: number
		readonly planArtifactDigest?: string
		readonly commitArtifactDigest?: string
		readonly digestsMatch: boolean
	}
	readonly plan: {
		readonly traceDigest: string
		readonly needsApproval: boolean
		readonly checkValid: boolean
		readonly lintClean: boolean
		readonly lossAuditOk: boolean
		readonly packageGraphAuditOk: boolean
		readonly writePolicyOk: boolean
		readonly phases: readonly AgentTracePhase[]
		readonly artifacts: readonly ReleaseProofArtifactDigest[]
	}
	readonly commit: {
		readonly traceDigest: string
		readonly outputSha256: string
		readonly checkValid: boolean
		readonly lintClean: boolean
		readonly lossAuditOk: boolean
		readonly packageGraphAuditOk: boolean
		readonly writePolicyOk: boolean
		readonly phases: readonly AgentTracePhase[]
		readonly artifacts: readonly ReleaseProofArtifactDigest[]
	}
	readonly reopen: {
		readonly valid: boolean
		readonly reopened: true
		readonly auditsPassed: boolean
		readonly outputSha256: string
		readonly checkValid: boolean
		readonly lintClean: boolean
		readonly packageGraphAuditOk: boolean
		readonly unresolvedPackageGraphIssueCount: number
	}
	readonly diff: {
		readonly included: boolean
		readonly sheetDiffCount?: number
		readonly changedSheets?: readonly string[]
	}
	readonly packageActions: {
		readonly plan: PackageActionProof
		readonly commit: PackageActionProof
		readonly countsMatch: boolean
		readonly issueCount: number
	}
	readonly consistency: {
		readonly valid: boolean
		readonly checks: readonly ReleaseProofConsistencyCheck[]
		readonly issues: readonly string[]
	}
	readonly claimBoundaries: readonly string[]
}

export interface CompactAgentCommitResult
	extends Pick<
		AgentCommitResult,
		| 'file'
		| 'output'
		| 'backup'
		| 'inputSha256'
		| 'outputSha256'
		| 'planDigest'
		| 'operationCount'
		| 'approvals'
		| 'modelOutput'
		| 'timings'
	> {
	readonly trace: CompactAgentTraceSummary
	readonly apply: CompactApplySummary
	readonly recalc: CompactRecalcSummary
	readonly check: CompactCheckSummary
	readonly lint: CompactLintSummary
	readonly preservation: CompactWritePlanSummary
	readonly writePolicy: CompactWritePolicySummary
	readonly postWrite: CompactAgentPostWriteVerification
	readonly lossAudit: CompactLossAuditSummary
	readonly packageGraphAudit: CompactPackageGraphAuditSummary
}

export interface CompactAgentTraceSummary {
	readonly kind: AgentWorkflowTrace['kind']
	readonly file: string
	readonly inputSha256: string
	readonly outputSha256?: string
	readonly planDigest?: string
	readonly traceDigest: string
	readonly phaseCount: number
	readonly warningCount: number
	readonly blockedCount: number
	readonly failedCount: number
	readonly artifactCount: number
	readonly phases: readonly AgentTracePhase[]
}

export interface CompactApplySummary {
	readonly applied: boolean
	readonly affectedCellCount: number
	readonly emittedAffectedCellCount: number
	readonly affectedCellRefs: readonly string[]
	readonly affectedRanges: readonly CompactAffectedRange[]
	readonly recalcRequired: boolean
	readonly warningCount: number
	readonly errorCount: number
	readonly journalSummary?: CompactJournalSummary
}

export interface CompactAffectedRange {
	readonly sheet: string
	readonly range: string
}

export interface CompactJournalSummary {
	readonly schemaVersion: number
	readonly schemaId: string
	readonly supported: boolean
	readonly exact: boolean
	readonly inverseOpCount: number
	readonly issueCount: number
	readonly issues: readonly MutationJournalIssue[]
}

export interface CompactRecalcSummary {
	readonly required: boolean
	readonly changedCellCount: number
	readonly errorCount: number
	readonly durationMs: number | null
}

export interface CompactCheckSummary {
	readonly valid: boolean
	readonly issueCount: number
	readonly errorCount: number
	readonly warningCount: number
	readonly infoCount: number
}

export interface CompactLintSummary {
	readonly clean: boolean
	readonly warningCount: number
	readonly errorCount: number
	readonly parseErrorCount: number
}

export interface CompactWritePlanSummary {
	readonly totalParts: number
	readonly byOrigin: ReturnType<AscendWorkbook['writePlanSummary']>['byOrigin']
	readonly skippedCapsuleCount: number
}

export interface CompactWritePolicySummary {
	readonly ok: boolean
	readonly diagnosticCount: number
	readonly blockerCount: number
	readonly warningCount: number
	readonly summary: WritePolicyReport['summary']
}

export interface CompactAgentPostWriteVerification {
	readonly valid: boolean
	readonly auditsPassed: boolean
	readonly output: string
	readonly outputSha256: string
	readonly reopened: true
	readonly timings?: AgentPostWriteVerificationTimings
	readonly check: CompactCheckSummary
	readonly lint: CompactLintSummary
	readonly preservation: CompactWritePlanSummary
	readonly opaquePayloads: PostWriteOpaquePayloadSummary
	readonly comments: PostWriteCommentSummary
	readonly hyperlinks: PostWriteHyperlinkSummary
	readonly dataValidations: PostWriteDataValidationSummary
	readonly dataConnections: PostWriteDataConnectionSummary
	readonly tables: PostWriteTableSummary
	readonly definedNames: PostWriteDefinedNameSummary
	readonly externalReferences: PostWriteExternalReferenceSummary
	readonly formulaState: PostWriteFormulaSummary
	readonly workbookTopology: PostWriteWorkbookTopologySummary
	readonly analytics: PostWriteAnalyticsSummary
	readonly activeContent: PostWriteActiveContentSummary
	readonly visuals: PostWriteVisualSummary
	readonly security: PostWriteSecuritySummary
	readonly packageGraphAudit: CompactPackageGraphAuditSummary
	readonly expectedPackageGraphIssueCount: number
	readonly unresolvedPackageGraphIssueCount: number
}

export interface CompactLossAuditSummary {
	readonly ok: boolean
	readonly blockedFeatureCount: number
	readonly blockedPackagePartCount: number
	readonly allowedLoss: readonly string[] | 'all'
	readonly policy: LossAudit['policy']
}

export interface CompactPackageGraphAuditSummary {
	readonly ok: boolean
	readonly issueCount: number
	readonly emittedIssueCount: number
	readonly issues: readonly XlsxPackageGraphFidelityIssue[]
	readonly policy: PackageGraphAudit['policy']
}

export interface AgentPostWriteVerification {
	readonly valid: boolean
	readonly auditsPassed: boolean
	readonly output: string
	readonly outputSha256: string
	readonly reopened: true
	readonly timings?: AgentPostWriteVerificationTimings
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly opaquePayloads: PostWriteOpaquePayloadSummary
	readonly comments: PostWriteCommentSummary
	readonly hyperlinks: PostWriteHyperlinkSummary
	readonly dataValidations: PostWriteDataValidationSummary
	readonly dataConnections: PostWriteDataConnectionSummary
	readonly tables: PostWriteTableSummary
	readonly definedNames: PostWriteDefinedNameSummary
	readonly externalReferences: PostWriteExternalReferenceSummary
	readonly formulaState: PostWriteFormulaSummary
	readonly workbookTopology: PostWriteWorkbookTopologySummary
	readonly analytics: PostWriteAnalyticsSummary
	readonly activeContent: PostWriteActiveContentSummary
	readonly visuals: PostWriteVisualSummary
	readonly security: PostWriteSecuritySummary
	readonly packageGraphAudit: PackageGraphAudit
	readonly expectedPackageGraphIssueCount: number
	readonly unresolvedPackageGraphIssueCount: number
}

export interface PostWriteOpaquePayloadSummary {
	readonly generatedWithOpaquePayloads: number
	readonly x14ConditionalFormatExtensionPayloads: number
	readonly x14DataValidationExtensionPayloads: number
	readonly worksheetParts: readonly string[]
	readonly preservationMode: 'generated-with-opaque-payload' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteCommentSummary {
	readonly legacyCommentLocations: number
	readonly threadedCommentLocations: number
	readonly legacyDrawingLocations: number
	readonly locations: readonly string[]
	readonly threadedCommentPartPaths: readonly string[]
	readonly verification: 'reopened-output'
}

export interface PostWriteHyperlinkSummary {
	readonly total: number
	readonly externalTargets: number
	readonly internalLocations: number
	readonly displayed: number
	readonly withTooltips: number
	readonly locations: readonly string[]
	readonly targets: readonly string[]
	readonly internalLocationTargets: readonly string[]
	readonly links: readonly PostWriteHyperlinkEntry[]
	readonly preservationMode: 'generated' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteHyperlinkEntry {
	readonly sheetName: string
	readonly ref: string
	readonly location: string
	readonly target?: string
	readonly internalLocation?: string
	readonly display?: string
	readonly tooltip?: string
}

export interface PostWriteDataValidationSummary {
	readonly total: number
	readonly formulaBacked: number
	readonly listValidations: number
	readonly x14Validations: number
	readonly ranges: readonly string[]
	readonly types: readonly string[]
	readonly validations: readonly PostWriteDataValidationEntry[]
	readonly preservationMode: 'generated' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteDataValidationEntry {
	readonly sheetName: string
	readonly sqref: string
	readonly location: string
	readonly type?: string
	readonly operator?: string
	readonly source?: string
	readonly formula1?: string
	readonly formula2?: string
	readonly allowBlank?: boolean
	readonly showInputMessage?: boolean
	readonly showErrorMessage?: boolean
}

export interface PostWriteDataConnectionSummary {
	readonly total: number
	readonly workbookConnections: number
	readonly queryTables: number
	readonly refreshOnOpen: number
	readonly notSaved: number
	readonly unknown: number
	readonly cached: number
	readonly partPaths: readonly string[]
	readonly names: readonly string[]
	readonly connectionIds: readonly number[]
	readonly connections: readonly PostWriteDataConnectionEntry[]
	readonly preservationMode: 'preserve-exact' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteDataConnectionEntry {
	readonly kind: 'connection' | 'queryTable' | 'powerQueryMashup'
	readonly partPath: string
	readonly state: 'cached' | 'refresh-on-open' | 'not-saved' | 'unknown'
	readonly sheetName?: string
	readonly name?: string
	readonly connectionId?: number
	readonly refreshOnLoad?: boolean
	readonly saveData?: boolean
	readonly refreshedVersion?: number
}

export interface PostWriteTableSummary {
	readonly tableLocations: number
	readonly queryTableLocations: number
	readonly tableAutoFilterLocations: number
	readonly tableNames: readonly string[]
	readonly locations: readonly string[]
	readonly tablePartPaths: readonly string[]
	readonly queryTablePartPaths: readonly string[]
	readonly preservationMode: 'preserve-exact' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteDefinedNameSummary {
	readonly total: number
	readonly workbookScoped: number
	readonly sheetScoped: number
	readonly hidden: number
	readonly names: readonly PostWriteDefinedNameEntry[]
	readonly verification: 'reopened-output'
}

export interface PostWriteDefinedNameEntry {
	readonly name: string
	readonly formula: string
	readonly scope: 'workbook' | 'sheet'
	readonly sheet?: string
	readonly hidden?: boolean
}

export interface PostWriteExternalReferenceSummary {
	readonly total: number
	readonly boundByExternalBookRelId: number
	readonly fallbackPathRelationships: number
	readonly missingPathRelationships: number
	readonly partPaths: readonly string[]
	readonly targets: readonly string[]
	readonly parts: readonly PostWriteExternalReferenceEntry[]
	readonly preservationMode: 'preserve-exact' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteExternalReferenceEntry {
	readonly partPath: string
	readonly relId?: string
	readonly externalBookRelId?: string
	readonly linkRelId?: string
	readonly linkBindingStatus?: ExternalReferenceInfo['linkBindingStatus']
	readonly linkRelationshipKind?: ExternalReferenceInfo['linkRelationshipKind']
	readonly target?: string
	readonly targetMode?: string
}

export interface PostWriteFormulaSummary {
	readonly calcChainState: 'present' | 'absent' | 'not-applicable'
	readonly calcChainParts: readonly string[]
	readonly recalculationRequested: boolean
	readonly calcSettings: CalcSettings
	readonly formulaCells: number
	readonly cachedFormulaValues: number
	readonly missingCachedFormulaValues: number
	readonly formulaCacheState: 'no-formulas' | 'all-cached' | 'partially-cached' | 'all-missing'
	readonly cachedValueKinds: readonly PostWriteFormulaCacheValueKindCount[]
	readonly missingCachedFormulaLocationSample: readonly string[]
	readonly warnings: readonly string[]
	readonly verification: 'reopened-output'
}

export interface PostWriteFormulaCacheValueKindCount {
	readonly kind: string
	readonly count: number
}

export interface PostWriteWorkbookTopologySummary {
	readonly sheets: number
	readonly visibleSheets: number
	readonly hiddenSheets: number
	readonly veryHiddenSheets: number
	readonly hiddenSheetNames: readonly string[]
	readonly veryHiddenSheetNames: readonly string[]
	readonly workbookViews: number
	readonly activeTabs: readonly number[]
	readonly firstSheets: readonly number[]
	readonly workbookViewDetails: readonly PostWriteWorkbookViewEntry[]
	readonly sheetStates: readonly PostWriteSheetTopologyEntry[]
	readonly verification: 'reopened-output'
}

export interface PostWriteWorkbookViewEntry {
	readonly index: number
	readonly activeTab?: number
	readonly firstSheet?: number
	readonly visibility?: string
	readonly tabRatio?: number
}

export interface PostWriteSheetTopologyEntry {
	readonly sheetName: string
	readonly state: SheetState
}

export interface PostWriteAnalyticsSummary {
	readonly pivotCaches: number
	readonly pivotTables: number
	readonly slicerCaches: number
	readonly slicers: number
	readonly timelineCaches: number
	readonly timelines: number
	readonly partPaths: readonly string[]
	readonly pivotCacheDetails: readonly PostWritePivotCacheEntry[]
	readonly pivotTableDetails: readonly PostWritePivotTableEntry[]
	readonly slicerCacheDetails: readonly PostWriteSlicerCacheEntry[]
	readonly timelineCacheDetails: readonly PostWriteTimelineCacheEntry[]
	readonly requiresExternalRefresh: boolean
	readonly preservationMode: 'preserve-exact' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWritePivotCacheEntry {
	readonly partPath: string
	readonly cacheId?: number
	readonly sourceType?: string
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly sourceName?: string
	readonly refreshOnLoad?: boolean
	readonly invalid?: boolean
	readonly saveData?: boolean
	readonly recordCount?: number
	readonly recordsPartPath?: string
	readonly outputState: 'cached' | 'stale' | 'refresh-on-open' | 'not-saved' | 'unknown'
	readonly requiresExternalRefresh: boolean
	readonly linkedPivotTableNames: readonly string[]
}

export interface PostWritePivotTableEntry {
	readonly partPath: string
	readonly name?: string
	readonly sheetName: string
	readonly cacheId?: number
	readonly locationRef?: string
}

export interface PostWriteSlicerCacheEntry {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly slicerPartPaths: readonly string[]
}

export interface PostWriteTimelineCacheEntry {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly timelinePartPaths: readonly string[]
	readonly selection?: { readonly startDate: string; readonly endDate: string }
}

export interface PostWriteActiveContentSummary {
	readonly total: number
	readonly vbaProjects: number
	readonly activeXControls: number
	readonly formControls: number
	readonly macroSheets: number
	readonly vbaSignatures: number
	readonly digitalSignatures: number
	readonly customUi: number
	readonly shapeMacros: number
	readonly unknownActiveContent: number
	readonly partPaths: readonly string[]
	readonly entries: readonly PostWriteActiveContentEntry[]
	readonly executionPolicy: 'blocked' | 'none'
	readonly preservationMode: 'preserve-exact' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteActiveContentEntry {
	readonly kind: ActiveContentInfo['kind']
	readonly partPath: string
	readonly contentType: string
	readonly anchor: ActiveContentInfo['anchor']
	readonly sheetName?: string
	readonly sourcePartPath?: string
	readonly relType?: string
	readonly sourceRelationshipId?: string
	readonly relationshipCount: number
	readonly byteSize?: number
	readonly opaque?: boolean
	readonly executionPolicy?: ActiveContentInfo['executionPolicy']
	readonly invalidationPolicy?: ActiveContentInfo['invalidationPolicy']
	readonly resigningPolicy?: ActiveContentInfo['resigningPolicy']
	readonly shapeMacro?: ActiveContentInfo['shapeMacro']
}

export interface PostWriteSecuritySummary {
	readonly workbookProtected: boolean
	readonly workbookLocks: readonly string[]
	readonly workbookPasswordProtected: boolean
	readonly workbookRevisionPasswordProtected: boolean
	readonly protectedSheets: number
	readonly protectedSheetNames: readonly string[]
	readonly sheetPasswordProtected: number
	readonly sheetStrongHashProtected: number
	readonly protectedRanges: number
	readonly protectedRangeLocations: readonly string[]
	readonly protectedRangePasswordProtected: number
	readonly protectedRangeStrongHashProtected: number
	readonly protectedRangeSecurityDescriptors: number
	readonly protectedRangeDetails: readonly PostWriteProtectedRangeSecurityEntry[]
	readonly sheets: readonly PostWriteSheetSecurityEntry[]
	readonly passwordHashVerification: 'reported-not-validated' | 'none'
	readonly preservationMode: 'generated' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteSheetSecurityEntry {
	readonly sheetName: string
	readonly protected: boolean
	readonly passwordProtected: boolean
	readonly strongHashProtected: boolean
	readonly allowedActions: readonly string[]
	readonly protectedRanges: number
	readonly protectedRangeLocations: readonly string[]
}

export interface PostWriteProtectedRangeSecurityEntry {
	readonly sheetName: string
	readonly name?: string
	readonly sqref: string
	readonly location: string
	readonly passwordProtected: boolean
	readonly strongHashProtected: boolean
	readonly hasSecurityDescriptor: boolean
}

export interface PostWriteVisualSummary {
	readonly sheetsWithVisuals: number
	readonly images: number
	readonly drawingObjects: number
	readonly drawingMlObjects: number
	readonly vmlObjects: number
	readonly chartParts: number
	readonly chartSheets: number
	readonly drawingPartPaths: readonly string[]
	readonly mediaPartPaths: readonly string[]
	readonly chartPartPaths: readonly string[]
	readonly vmlPartPaths: readonly string[]
	readonly sheets: readonly PostWriteSheetVisualEntry[]
	readonly charts: readonly PostWriteChartEntry[]
	readonly preservationMode: 'preserve-exact' | 'none'
	readonly verification: 'reopened-output'
}

export interface PostWriteSheetVisualEntry {
	readonly sheetName: string
	readonly hasDrawingMl: boolean
	readonly hasVml: boolean
	readonly imageCount: number
	readonly drawingObjectCount: number
	readonly drawingMlObjectCount: number
	readonly vmlObjectCount: number
	readonly drawingPartPaths: readonly string[]
	readonly mediaPartPaths: readonly string[]
	readonly vmlPartPaths: readonly string[]
}

export interface PostWriteChartEntry {
	readonly partPath: string
	readonly sheetName?: string
	readonly chartType?: string
	readonly title?: string
	readonly seriesCount: number
	readonly series: readonly PostWriteChartSeriesEntry[]
}

export interface PostWriteChartSeriesEntry {
	readonly index: number
	readonly nameRef?: string
	readonly nameText?: string
	readonly categoryRef?: string
	readonly valueRef?: string
}

export interface AgentPostWriteVerificationTimings {
	readonly outputSnapshotMs: number
	readonly reopenMs: number
	readonly checkMs: number
	readonly lintMs: number
	readonly preservationMs: number
	readonly packageGraphMs: number
	readonly packageGraphAuditMs: number
}

export interface AgentCommitTimings {
	readonly writePolicySnapshotMs: number
	readonly packageGraphMs: number
	readonly approvalAuditMs: number
	readonly lossAuditMs: number
	readonly packageGraphAuditMs: number
	readonly applyMs: number
	readonly recalcMs: number
	readonly writePlanSummaryMs: number
	readonly writePolicyCheckMs: number
	readonly writePolicyBuildMs: number
	readonly toBytesMs: number
	readonly writeFileMs: number
	readonly renameMs: number
	readonly outputByteReadMs: number
	readonly outputHashMs: number
}

export interface LossAudit {
	readonly ok: boolean
	readonly blockedFeatures: readonly FeatureReport[]
	readonly blockedPackageParts: readonly LossAuditPackagePart[]
	readonly allowedLoss: readonly string[] | 'all'
	readonly policy: 'block-preserved-and-unsupported'
}

export interface LossAuditPackagePart {
	readonly partPath: string
	readonly featureFamily: string
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly preservationMode: WritePolicyPreservationMode
	readonly ownerScope: XlsxPackageOwnerScope
	readonly bytePreservationExpected: boolean
	readonly contentType: string
	readonly contentTypeSource: string
	readonly sourceRelationshipPart?: string
	readonly sourceRelationshipId?: string
	readonly sourceRelationshipType?: string
	readonly sourceRelationshipRawTarget?: string
	readonly sourceRelationshipResolvedTarget?: string
	readonly sourceRelationshipTargetMode?: string
	readonly reason: string
}

export interface PackageGraphAudit {
	readonly ok: boolean
	readonly issues: readonly XlsxPackageGraphFidelityIssue[]
	readonly policy: 'read-integrity' | 'safe-edit-roundtrip' | 'not-applicable-non-xlsx'
}

export interface WritePolicyReport {
	readonly ok: boolean
	readonly diagnostics: readonly WritePolicyDiagnostic[]
	readonly summary: {
		readonly generatedParts: number
		readonly copiedThroughParts: number
		readonly skippedCapsules: number
		readonly invalidatedSignatures: number
		readonly approvalRequiredFeatures: number
		readonly packageGraphIssues: number
		readonly externalReferences: number
		readonly externalReferenceBindingIssues: number
		readonly chartSourceIntegrityIssues: number
		readonly legacyCommentLocations: number
		readonly threadedCommentLocations: number
		readonly commentIntegrityIssues: number
		readonly x14ConditionalFormatExtensionPayloads: number
		readonly x14DataValidationExtensionPayloads: number
		readonly calcChainPolicy: 'not-present' | 'preserved' | 'discarded-for-formula-topology'
		readonly preservationModes: WritePolicyPreservationModeSummary
	}
}

export interface WritePolicyPreservationModeSummary {
	readonly preserveExactParts: number
	readonly generatedParts: number
	readonly generatedWithOpaquePayloads: number
	readonly invalidatedOnEditParts: number
	readonly discardedForRecalcParts: number
	readonly inspectOnlyParts: number
	readonly reviewRequiredParts: number
	readonly unsupportedFeatures: number
	readonly lossyApprovalRequiredFeatures: number
}

export type WritePolicyPreservationMode =
	| 'preserve-exact'
	| 'generated'
	| 'generated-with-opaque-payload'
	| 'invalidated-on-edit'
	| 'discarded-for-recalc'
	| 'inspect-only'
	| 'review-required'
	| 'unsupported'
	| 'lossy-approval-required'

export interface WritePolicyDiagnostic {
	readonly code:
		| 'generated-replacement-parts'
		| 'copied-through-package-parts'
		| 'skipped-preservation-capsules'
		| 'signature-invalidation'
		| 'calc-chain-preserved'
		| 'calc-chain-discarded'
		| 'active-content-preserved'
		| 'visual-sidecar-preservation-risk'
		| 'visual-edit-preservation-risk'
		| 'drawingml-vml-drift-risk'
		| 'chart-source-ref-drift-risk'
		| 'analytics-preservation-risk'
		| 'analytics-pivot-refresh-risk'
		| 'table-preservation-risk'
		| 'legacy-comment-preservation-risk'
		| 'threaded-comment-preservation-risk'
		| 'conditional-format-extension-preservation'
		| 'data-validation-extension-preservation'
		| 'external-link-dependency'
		| 'external-link-binding-risk'
		| 'external-link-package-risk'
		| 'package-graph-audit-issue'
		| 'pre-write-check-error'
		| 'approval-required-feature'
	readonly severity: 'info' | 'warning' | 'blocker'
	readonly message: string
	readonly suggestedAction: string
	readonly locations?: readonly string[]
	readonly partPaths?: readonly string[]
	readonly packageParts?: readonly WritePolicyPackagePart[]
	readonly featureFamily?: string
	readonly ownerScope?: XlsxPackageOwnerScope
	readonly preservationPolicy?: XlsxPackageLossPolicy
	readonly preservationMode?: WritePolicyPreservationMode
	readonly details?: unknown
}

export interface WritePolicyPackagePart {
	readonly partPath: string
	readonly featureFamily: string
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly preservationMode: WritePolicyPreservationMode
	readonly ownerScope: XlsxPackageOwnerScope
	readonly bytePreservationExpected: boolean
	readonly contentType: string
	readonly contentTypeSource: string
	readonly sourceRelationshipPart?: string
	readonly sourceRelationshipId?: string
	readonly sourceRelationshipType?: string
	readonly sourceRelationshipRawType?: string
	readonly sourceRelationshipRawTarget?: string
	readonly sourceRelationshipResolvedTarget?: string
	readonly sourceRelationshipTargetMode?: string
}

export interface AgentWorkflowTrace {
	readonly formatVersion: 1
	readonly kind: 'plan' | 'commit' | 'repair-plan'
	readonly file: string
	readonly inputSha256: string
	readonly outputSha256?: string
	readonly planDigest?: string
	readonly operationCount?: number
	readonly traceDigest: string
	readonly phases: readonly AgentTracePhase[]
	readonly artifacts: readonly AgentTraceArtifact[]
}

export interface AgentTracePhase {
	readonly phase: string
	readonly status: 'ok' | 'warning' | 'blocked' | 'failed'
	readonly summary: string
	readonly count?: number
	readonly refs?: readonly string[]
	readonly details?: unknown
}

export interface AgentTraceArtifact {
	readonly name: string
	readonly digest: string
	readonly summary: string
}

export type AgentWorkflowProgressHandler = (
	event: AgentWorkflowProgressEvent,
) => void | Promise<void>

export interface AgentWorkflowProgressEvent {
	readonly formatVersion: 1
	readonly sequence: number
	readonly kind: AgentWorkflowTrace['kind']
	readonly phase: string
	readonly status: AgentTracePhase['status'] | 'started' | 'skipped'
	readonly summary: string
	readonly count?: number
	readonly refs?: readonly string[]
	readonly details?: unknown
}

export interface AgentModelOutput {
	readonly summary: string
	readonly blocked: boolean
	readonly warnings: readonly string[]
	readonly nextActions: readonly string[]
	readonly digests: {
		readonly inputSha256: string
		readonly outputSha256?: string
		readonly planDigest?: string
		readonly traceDigest: string
	}
	readonly counts: {
		readonly operations?: number
		readonly changedCells?: number
		readonly recalcErrors?: number
		readonly checkIssues?: number
		readonly lintWarnings?: number
		readonly blockedFeatures?: number
		readonly packageGraphIssues?: number
		readonly postWritePackageGraphIssues?: number
		readonly postWriteLintFailures?: number
		readonly preservationParts?: number
		readonly writePolicyDiagnostics?: number
	}
}

interface ExpectedPostWritePackageGraphChanges {
	readonly removedPartPaths: ReadonlySet<string>
	readonly addedPartPaths: ReadonlySet<string>
	readonly rewrittenPartPaths: ReadonlySet<string>
}

export interface ApprovalRequirement {
	readonly id: string
	readonly kind: 'lossy-write' | 'destructive-operation'
	readonly severity: 'medium' | 'high' | 'critical'
	readonly title: string
	readonly reason: string
	readonly operationIndex?: number
	readonly operation?: string
	readonly feature?: string
	readonly tier?: FeatureReport['tier']
	readonly satisfies: readonly string[]
}

export interface RepairPlanResult {
	readonly file: string
	readonly inputSha256: string
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly unsupportedFeatures: readonly unknown[]
	readonly actions: readonly RepairAction[]
}

export interface RepairAction {
	readonly code: string
	readonly title: string
	readonly command?: string
	readonly details?: unknown
}

interface AgentSourceIdentity {
	readonly size: number
	readonly mtimeMs: number
}

export async function createAgentPlan(
	file: string,
	ops: readonly Operation[],
	options: AgentPlanOptions = {},
): Promise<AgentPlanResult> {
	const progress = createProgressEmitter('plan', options.onProgress)
	await progress('hash-input', 'started', 'Hashing input workbook.')
	const source = await readStableAgentSource(file)
	const inputSha256 = sha256Bytes(source.sourceBytes)
	await progress('hash-input', 'ok', 'Input workbook hash captured.')
	await progress('load-workbook', 'started', 'Opening workbook.')
	const wb = await openWorkbookFromBytes(file, source.sourceBytes, passwordOpenOptions(options))
	await progress('load-workbook', 'ok', 'Workbook opened.')
	return createAgentPlanFromWorkbook(file, inputSha256, wb, ops, { progress })
}

export async function createPreparedAgentPlan(
	file: string,
	ops: readonly Operation[],
	options: AgentPlanOptions = {},
): Promise<PreparedAgentPlan> {
	const progress = createProgressEmitter('plan', options.onProgress)
	await progress('hash-input', 'started', 'Hashing input workbook.')
	const source = await readStableAgentSource(file)
	const inputSha256 = sha256Bytes(source.sourceBytes)
	await progress('hash-input', 'ok', 'Input workbook hash captured.')
	await progress('load-workbook', 'started', 'Opening workbook.')
	const wb = await openWorkbookFromBytes(file, source.sourceBytes, passwordOpenOptions(options))
	await progress('load-workbook', 'ok', 'Workbook opened.')
	const plan = await createAgentPlanFromWorkbook(file, inputSha256, wb, ops, { progress })
	const planDigest = plan.planDigest
	let committed = false
	return {
		file,
		inputSha256,
		planDigest,
		operationCount: ops.length,
		plan,
		async commit(commitOptions: AgentCommitOptions = {}): Promise<AgentCommitResult> {
			if (committed) {
				throw new AscendException(
					ascendError('VALIDATION_ERROR', 'Prepared agent plan has already been committed', {
						suggestedFix: 'Create a fresh prepared plan before committing another output.',
					}),
				)
			}
			const currentSource = await readStableAgentSource(file)
			const actualSha256 = sha256Bytes(currentSource.sourceBytes)
			if (actualSha256 !== inputSha256) {
				throw new AscendException(
					ascendError('VALIDATION_ERROR', 'Input workbook changed after agent plan was prepared', {
						details: {
							expected: source.identity,
							actual: currentSource.identity,
							expectedSha256: inputSha256,
							actualSha256,
							planDigest,
						},
						suggestedFix: 'Re-run ascend plan and commit with the new input workbook.',
					}),
				)
			}
			const committedResult = await commitAgentPlanFromWorkbook(
				file,
				inputSha256,
				wb,
				ops,
				{ ...commitOptions, expectSha256: commitOptions.expectSha256 ?? inputSha256 },
				{ sourceBytes: source.sourceBytes, preparedCheck: plan.check },
			)
			committed = true
			return committedResult
		},
	}
}

export async function createAgentPlanFromWorkbook(
	file: string,
	inputSha256: string,
	wb: AscendWorkbook,
	ops: readonly Operation[],
	options: {
		readonly progress?: ReturnType<typeof createProgressEmitter>
		readonly onProgress?: AgentWorkflowProgressHandler
	} = {},
): Promise<AgentPlanResult> {
	const progress = options.progress ?? createProgressEmitter('plan', options.onProgress)
	const load = wb.inspect().load
	if (load.isPartial) {
		throw new AscendException(
			ascendError(
				'VALIDATION_ERROR',
				'Cannot create an agent write plan from a partial workbook view',
				{
					details: { load },
					suggestedFix: 'Reopen the workbook with a full load before planning or committing edits.',
				},
			),
		)
	}
	const workbookModel = wb.getWorkbookModel()
	const writePolicyWorkbook = canUseLiveWritePolicyWorkbook(workbookModel, ops)
		? workbookModel
		: snapshotWritePolicyWorkbook(workbookModel)
	await progress('preview', 'started', 'Previewing operations.', { count: ops.length })
	const preview = wb.preview(ops, { journal: true })
	await progressFromPhase(previewPhase(preview), progress)
	const packageGraph = wb.packageGraph()
	await progress('loss-audit', 'started', 'Auditing preserved and unsupported features.')
	const lossAudit = auditLossPolicy(wb.report.features, [], packageGraph)
	await progressFromPhase(lossAuditPhase(lossAudit), progress)
	await progress('package-graph-audit', 'started', 'Auditing package graph fidelity.')
	const packageGraphAudit = auditPackageGraphIntegrity(packageGraph)
	await progressFromPhase(packageGraphAuditPhase(packageGraphAudit), progress)
	await progress('approval-audit', 'started', 'Auditing explicit approval requirements.')
	const approvals = buildApprovalRequirements(wb.report.features, ops)
	await progressFromPhase(approvalPhase(approvals), progress)
	await progress('check', 'started', 'Running structural checks.')
	const check = wb.check()
	await progressFromPhase(checkPhase(check), progress)
	await progress('lint', 'started', 'Running formula lint.')
	const lint = wb.lint()
	await progressFromPhase(lintPhase(lint), progress)
	await progress('preservation-audit', 'started', 'Summarizing package preservation.')
	const preservation = preview.writePlan ?? wb.writePlanSummary()
	await progressFromPhase(preservationPhase(preservation), progress)
	await progress('write-policy', 'started', 'Explaining write preservation and loss policy.')
	const writePolicy = buildWritePolicyReport(
		wb.report.features,
		packageGraph,
		preservation,
		packageGraphAudit,
		writePolicyWorkbook,
		ops,
		wb.inspect().externalReferenceDetails,
		check.issues,
	)
	await progressFromPhase(writePolicyPhase(writePolicy), progress)
	const planDigest = digestPlan(inputSha256, ops)
	await progress('finalize', 'ok', 'Plan digest and trace finalized.')
	const trace = finalizeTrace({
		kind: 'plan',
		file,
		inputSha256,
		planDigest,
		operationCount: ops.length,
		phases: [
			okPhase('hash-input', 'Input workbook hash captured.'),
			previewPhase(preview),
			checkPhase(check),
			lintPhase(lint),
			lossAuditPhase(lossAudit),
			packageGraphAuditPhase(packageGraphAudit),
			approvalPhase(approvals),
			preservationPhase(preservation),
			writePolicyPhase(writePolicy),
		],
		artifacts: [
			artifact('ops', ops, `${ops.length} operation(s)`),
			artifact('preview', preview, `${preview.changedCells.length} changed cell(s)`),
			artifact('lossAudit', lossAudit, `${lossAudit.blockedFeatures.length} blocked feature(s)`),
			artifact(
				'packageGraphAudit',
				packageGraphAudit,
				`${packageGraphAudit.issues.length} package graph issue(s)`,
			),
			artifact('approvals', approvals, `${approvals.length} approval requirement(s)`),
			artifact('preservation', preservation, `${preservation.totalParts} package part(s)`),
			artifact(
				'writePolicy',
				writePolicy,
				`${writePolicy.diagnostics.length} write policy diagnostic(s)`,
			),
		],
	})
	return {
		file,
		inputSha256,
		operationCount: ops.length,
		planDigest,
		needsApproval: approvals.length > 0,
		approvals,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		preview,
		check,
		lint,
		preservation,
		writePolicy,
		unsupportedFeatures: wb.report.features,
		lossAudit,
		packageGraphAudit,
		capabilities: summarizeCapabilities(listCapabilities({ gapsOnly: true })),
	}
}

export function compactAgentPlanResult(
	result: AgentPlanResult,
	options: CompactAgentPlanOptions = {},
): CompactAgentPlanResult {
	const maxChangedCells = Math.max(0, Math.floor(options.maxChangedCells ?? 50))
	const changedCells = result.preview.changedCells.slice(0, maxChangedCells)
	return {
		...result,
		preview: {
			wouldSucceed: result.preview.wouldSucceed,
			changedCellCount: result.preview.changedCells.length,
			emittedChangedCellCount: changedCells.length,
			changedCells,
			changedRanges: compactRangesFromRefs(result.preview.changedCells.map((cell) => cell.ref)),
			recalcScope: result.preview.recalcScope,
			warningCount: result.preview.warnings.length,
			errorCount: result.preview.errors.length,
			warnings: result.preview.warnings,
			errors: result.preview.errors,
			...(result.preview.journal
				? { journalSummary: compactJournalSummary(result.preview.journal) }
				: {}),
		},
	}
}

export function compactAgentCommitResult(
	result: AgentCommitResult,
	options: CompactAgentCommitOptions = {},
): CompactAgentCommitResult {
	const maxAffectedCells = Math.max(0, Math.floor(options.maxAffectedCells ?? 50))
	const affectedCellRefs = result.apply.affectedCells.slice(0, maxAffectedCells)
	return {
		file: result.file,
		output: result.output,
		...(result.backup ? { backup: result.backup } : {}),
		inputSha256: result.inputSha256,
		outputSha256: result.outputSha256,
		planDigest: result.planDigest,
		operationCount: result.operationCount,
		approvals: result.approvals,
		timings: result.timings,
		trace: compactTraceSummary(result.trace),
		modelOutput: result.modelOutput,
		apply: {
			applied: result.apply.errors.length === 0,
			affectedCellCount: result.apply.affectedCells.length,
			emittedAffectedCellCount: affectedCellRefs.length,
			affectedCellRefs,
			affectedRanges: compactRangesFromDirtyRegions(result.apply.dirtyRegions),
			recalcRequired: result.apply.recalcRequired,
			warningCount: result.apply.warnings?.length ?? 0,
			errorCount: result.apply.errors.length,
			...(result.apply.journal
				? { journalSummary: compactJournalSummary(result.apply.journal) }
				: {}),
		},
		recalc: {
			required: result.recalc !== null,
			changedCellCount: result.recalc?.changed.length ?? 0,
			errorCount: result.recalc?.errors.length ?? 0,
			durationMs: result.recalc?.duration ?? null,
		},
		check: compactCheckSummary(result.check),
		lint: compactLintSummary(result.lint),
		preservation: compactWritePlanSummary(result.preservation),
		writePolicy: compactWritePolicySummary(result.writePolicy),
		postWrite: compactPostWriteVerification(result.postWrite),
		lossAudit: compactLossAuditSummary(result.lossAudit),
		packageGraphAudit: compactPackageGraphAuditSummary(result.packageGraphAudit),
	}
}

export function createAgentWorkflowProofSummary(
	plan: AgentPlanResult,
	commit: AgentCommitResult,
	options: AgentWorkflowProofSummaryOptions = {},
): AgentWorkflowProofSummary {
	const whatChanged = plan.preview.cellChanges.map((cell) => ({
		ref: qualifyProofCellRef(cell.ref, options.defaultSheetName),
		before: cell.before,
		after: cell.after,
		formulaBefore: cell.formulaBefore,
		formulaAfter: cell.formulaAfter,
	}))
	const whySafe = [
		...(options.preflightGates ?? []),
		{
			gate: 'plan-linked',
			ok: commit.inputSha256 === plan.inputSha256 && commit.planDigest === plan.planDigest,
			evidence: {
				planInputSha256: plan.inputSha256,
				commitInputSha256: commit.inputSha256,
				planDigest: plan.planDigest,
				commitPlanDigest: commit.planDigest,
			},
		},
		{
			gate: 'plan',
			ok: plan.preview.wouldSucceed && plan.approvals.length === 0,
			evidence: {
				planDigest: plan.planDigest,
				changedCellCount: whatChanged.length,
				approvalCount: plan.approvals.length,
				errorCount: plan.preview.errors.length,
			},
		},
		{
			gate: 'write-policy',
			ok: commit.writePolicy.ok,
			evidence: {
				diagnosticCount: commit.writePolicy.diagnostics.length,
				blockerCount: commit.writePolicy.diagnostics.filter(
					(diagnostic) => diagnostic.severity === 'blocker',
				).length,
			},
		},
		{
			gate: 'commit',
			ok: commit.postWrite.valid && commit.postWrite.auditsPassed,
			evidence: {
				outputSha256: commit.outputSha256,
				postWriteValid: commit.postWrite.valid,
				auditsPassed: commit.postWrite.auditsPassed,
			},
		},
		{
			gate: 'reopen-verify',
			ok: commit.postWrite.reopened && commit.postWrite.check.valid && commit.postWrite.lint.clean,
			evidence: {
				reopened: commit.postWrite.reopened,
				checkValid: commit.postWrite.check.valid,
				checkIssueCount: commit.postWrite.check.issues.length,
				lintClean: commit.postWrite.lint.clean,
				lintWarningCount: commit.postWrite.lint.warnings.length,
			},
		},
		{
			gate: 'package-graph',
			ok:
				commit.postWrite.packageGraphAudit.ok &&
				commit.postWrite.unresolvedPackageGraphIssueCount === 0,
			evidence: {
				packageGraphAuditOk: commit.postWrite.packageGraphAudit.ok,
				expectedPackageGraphIssueCount: commit.postWrite.expectedPackageGraphIssueCount,
				unresolvedPackageGraphIssueCount: commit.postWrite.unresolvedPackageGraphIssueCount,
			},
		},
	] as const satisfies readonly AgentWorkflowProofGate[]
	return {
		safeToUse: whySafe.every((gate) => gate.ok),
		whatChanged,
		whySafe,
		evidence: {
			inputSha256: plan.inputSha256,
			planDigest: plan.planDigest,
			outputSha256: commit.outputSha256,
			operationCount: commit.operationCount,
			changedCellCount: whatChanged.length,
			postWriteValid: commit.postWrite.valid,
			auditsPassed: commit.postWrite.auditsPassed,
			reopened: commit.postWrite.reopened,
			checkValid: commit.postWrite.check.valid,
			lintClean: commit.postWrite.lint.clean,
			writePolicyOk: commit.writePolicy.ok,
			packageGraphAuditOk: commit.postWrite.packageGraphAudit.ok,
		},
	}
}

function qualifyProofCellRef(ref: string, defaultSheetName: string | undefined): string {
	return defaultSheetName && !ref.includes('!') ? `${defaultSheetName}!${ref}` : ref
}

function compactRangesFromDirtyRegions(
	dirtyRegions: ReturnType<AscendWorkbook['apply']>['dirtyRegions'],
): readonly CompactAffectedRange[] {
	return dirtyRegions.map(({ sheet, range }) => ({ sheet, range }))
}

function compactRangesFromRefs(refs: readonly string[]): readonly CompactAffectedRange[] {
	const bySheet = new Map<
		string,
		{ minRow: number; minCol: number; maxRow: number; maxCol: number }
	>()
	for (const fullRef of refs) {
		const bang = fullRef.lastIndexOf('!')
		const sheet = bang === -1 ? '' : fullRef.slice(0, bang).replace(/^'|'$/g, '')
		const ref = bang === -1 ? fullRef : fullRef.slice(bang + 1)
		if (!sheet || !ref) continue
		try {
			const range = parseRange(ref)
			const current = bySheet.get(sheet)
			if (current) {
				current.minRow = Math.min(current.minRow, range.start.row)
				current.minCol = Math.min(current.minCol, range.start.col)
				current.maxRow = Math.max(current.maxRow, range.end.row)
				current.maxCol = Math.max(current.maxCol, range.end.col)
			} else {
				bySheet.set(sheet, {
					minRow: range.start.row,
					minCol: range.start.col,
					maxRow: range.end.row,
					maxCol: range.end.col,
				})
			}
		} catch {}
	}
	return [...bySheet.entries()].map(([sheet, range]) => ({
		sheet,
		range: `${indexToColumn(range.minCol)}${range.minRow + 1}:${indexToColumn(range.maxCol)}${range.maxRow + 1}`,
	}))
}

function compactJournalSummary(journal: MutationJournal): CompactJournalSummary {
	return {
		schemaVersion: MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
		schemaId: MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
		supported: journal.supported,
		exact: journal.exact,
		inverseOpCount: journal.inverseOps.length,
		issueCount: journal.issues.length,
		issues: journal.issues,
	}
}

export function createReleaseProofBundle(
	plan: AgentPlanResult,
	commit: AgentCommitResult,
	options: ReleaseProofBundleOptions = {},
): ReleaseProofBundle {
	const planOpsArtifact = traceArtifactByName(plan.trace, 'ops')
	const commitOpsArtifact = traceArtifactByName(commit.trace, 'ops')
	const planPackageActions = createPackageActionProof(plan.preservation, {
		...(options.sourceBytes ? { sourceBytes: options.sourceBytes } : {}),
		writePolicy: plan.writePolicy,
		packageGraphAudit: plan.packageGraphAudit,
	})
	const commitPackageActions = createPackageActionProof(commit.preservation, {
		...(options.sourceBytes ? { sourceBytes: options.sourceBytes } : {}),
		...(options.outputBytes ? { outputBytes: options.outputBytes } : {}),
		writePolicy: commit.writePolicy,
		packageGraphAudit: commit.packageGraphAudit,
	})
	const checks: ReleaseProofConsistencyCheck[] = [
		releaseProofCheck('input-hash-linked', plan.inputSha256 === commit.inputSha256, {
			planInputSha256: plan.inputSha256,
			commitInputSha256: commit.inputSha256,
		}),
		releaseProofCheck('plan-digest-linked', plan.planDigest === commit.planDigest, {
			planDigest: plan.planDigest,
			commitPlanDigest: commit.planDigest,
		}),
		releaseProofCheck('operation-count-linked', plan.operationCount === commit.operationCount, {
			planOperationCount: plan.operationCount,
			commitOperationCount: commit.operationCount,
		}),
		releaseProofCheck(
			'operation-artifact-linked',
			planOpsArtifact?.digest === commitOpsArtifact?.digest,
			{
				planOpsDigest: planOpsArtifact?.digest ?? null,
				commitOpsDigest: commitOpsArtifact?.digest ?? null,
			},
		),
		releaseProofCheck(
			'post-write-output-hash-linked',
			commit.outputSha256 === commit.postWrite.outputSha256,
			{
				outputSha256: commit.outputSha256,
				postWriteOutputSha256: commit.postWrite.outputSha256,
			},
		),
		releaseProofCheck('post-write-reopened', commit.postWrite.valid && commit.postWrite.reopened),
		releaseProofCheck('post-write-audits-passed', commit.postWrite.auditsPassed, {
			unresolvedPackageGraphIssueCount: commit.postWrite.unresolvedPackageGraphIssueCount,
		}),
		releaseProofCheck('plan-package-graph-audit-ok', plan.packageGraphAudit.ok, {
			issueCount: plan.packageGraphAudit.issues.length,
		}),
		releaseProofCheck('commit-package-graph-audit-ok', commit.postWrite.packageGraphAudit.ok, {
			issueCount: commit.postWrite.packageGraphAudit.issues.length,
		}),
		releaseProofCheck('plan-package-action-proof-clean', planPackageActions.issues.length === 0, {
			issueCount: planPackageActions.issues.length,
		}),
		releaseProofCheck(
			'commit-package-action-proof-clean',
			commitPackageActions.issues.length === 0,
			{ issueCount: commitPackageActions.issues.length },
		),
	]
	const issues = checks
		.filter((check) => !check.ok)
		.map((check) => `release proof check failed: ${check.check}`)
	return {
		formatVersion: 1,
		kind: 'ascend-release-proof-bundle',
		proofKind: 'local-evidence',
		subject: {
			file: plan.file,
			output: commit.output,
			inputSha256: plan.inputSha256,
			outputSha256: commit.outputSha256,
			planDigest: plan.planDigest,
		},
		operations: {
			count: plan.operationCount,
			...(planOpsArtifact ? { planArtifactDigest: planOpsArtifact.digest } : {}),
			...(commitOpsArtifact ? { commitArtifactDigest: commitOpsArtifact.digest } : {}),
			digestsMatch: planOpsArtifact?.digest === commitOpsArtifact?.digest,
		},
		plan: {
			traceDigest: plan.trace.traceDigest,
			needsApproval: plan.needsApproval,
			checkValid: plan.check.valid,
			lintClean: plan.lint.clean,
			lossAuditOk: plan.lossAudit.ok,
			packageGraphAuditOk: plan.packageGraphAudit.ok,
			writePolicyOk: plan.writePolicy.ok,
			phases: plan.trace.phases,
			artifacts: plan.trace.artifacts.map(releaseProofArtifactDigest),
		},
		commit: {
			traceDigest: commit.trace.traceDigest,
			outputSha256: commit.outputSha256,
			checkValid: commit.check.valid,
			lintClean: commit.lint.clean,
			lossAuditOk: commit.lossAudit.ok,
			packageGraphAuditOk: commit.packageGraphAudit.ok,
			writePolicyOk: commit.writePolicy.ok,
			phases: commit.trace.phases,
			artifacts: commit.trace.artifacts.map(releaseProofArtifactDigest),
		},
		reopen: {
			valid: commit.postWrite.valid,
			reopened: commit.postWrite.reopened,
			auditsPassed: commit.postWrite.auditsPassed,
			outputSha256: commit.postWrite.outputSha256,
			checkValid: commit.postWrite.check.valid,
			lintClean: commit.postWrite.lint.clean,
			packageGraphAuditOk: commit.postWrite.packageGraphAudit.ok,
			unresolvedPackageGraphIssueCount: commit.postWrite.unresolvedPackageGraphIssueCount,
		},
		diff: options.diff
			? {
					included: true,
					sheetDiffCount: options.diff.sheetDiffCount,
					...(options.diff.changedSheets ? { changedSheets: options.diff.changedSheets } : {}),
				}
			: { included: false },
		packageActions: {
			plan: planPackageActions,
			commit: commitPackageActions,
			countsMatch: packageActionCountsEqual(
				planPackageActions.byAction,
				commitPackageActions.byAction,
			),
			issueCount: planPackageActions.issues.length + commitPackageActions.issues.length,
		},
		consistency: {
			valid: issues.length === 0,
			checks,
			issues,
		},
		claimBoundaries: options.claimBoundaries ?? DEFAULT_RELEASE_PROOF_CLAIM_BOUNDARIES,
	}
}

export function createPackageActionProof(
	preservation: WorkbookWritePlanInfo,
	options: PackageActionProofOptions = {},
): PackageActionProof {
	const sourceArchive = packageActionProofArchive(options.sourceBytes)
	const outputArchive = packageActionProofArchive(options.outputBytes)
	const sourcePackageGraph =
		options.sourcePackageGraph ??
		(options.sourceBytes ? inspectXlsxPackageGraph(options.sourceBytes) : undefined)
	const sourcePartByPath = sourcePackageGraph
		? new Map(sourcePackageGraph.parts.map((part) => [part.path, part]))
		: undefined
	const plannedPaths = new Set(preservation.parts.map((part) => part.path))
	const diagnosticsByPath = diagnosticsByPartPath(options.writePolicy?.diagnostics ?? [])
	const auditIssuesByPath = auditIssuesByPartPath(options.packageGraphAudit?.issues ?? [])
	const actions: PackageActionProofEntry[] = []

	for (const part of preservation.parts) {
		const sourcePresent = sourcePartByPath?.has(part.path)
		const action = packageActionForPlannedPart(part, sourcePresent)
		const diagnosticCodes = diagnosticsByPath.get(part.path)?.map((diagnostic) => diagnostic.code)
		const auditIssueCodes = auditIssuesByPath.get(part.path)?.map((issue) => issue.code)
		actions.push({
			action,
			reason: packageActionReason(action, part, sourcePresent),
			partPath: part.path,
			...(sourcePresent !== undefined ? { sourcePresent } : {}),
			origin: part.origin,
			owner: part.owner,
			...(part.contentType ? { contentType: part.contentType } : {}),
			streaming: part.streaming,
			...packageActionDigestEvidence(part.path, sourceArchive, outputArchive),
			...(diagnosticCodes?.length ? { diagnosticCodes: Array.from(new Set(diagnosticCodes)) } : {}),
			...(auditIssueCodes?.length ? { auditIssueCodes: Array.from(new Set(auditIssueCodes)) } : {}),
		})
	}

	for (const partPath of preservation.skippedCapsules) {
		if (plannedPaths.has(partPath)) continue
		const sourcePresent = sourcePartByPath?.has(partPath)
		const diagnosticCodes = diagnosticsByPath.get(partPath)?.map((diagnostic) => diagnostic.code)
		const auditIssueCodes = auditIssuesByPath.get(partPath)?.map((issue) => issue.code)
		actions.push({
			action: 'drop',
			reason:
				'Preservation capsule is listed as skipped and will not be emitted by the planned write.',
			partPath,
			...(sourcePresent !== undefined ? { sourcePresent } : {}),
			...packageActionDigestEvidence(partPath, sourceArchive, outputArchive),
			...(diagnosticCodes?.length ? { diagnosticCodes: Array.from(new Set(diagnosticCodes)) } : {}),
			...(auditIssueCodes?.length ? { auditIssueCodes: Array.from(new Set(auditIssueCodes)) } : {}),
		})
	}

	for (const diagnostic of options.writePolicy?.diagnostics ?? []) {
		if (diagnostic.severity !== 'blocker') continue
		const partPaths = diagnostic.partPaths?.length ? diagnostic.partPaths : [undefined]
		for (const partPath of partPaths) {
			actions.push({
				action: 'error',
				reason: diagnostic.message,
				...(partPath ? { partPath } : {}),
				...(partPath && sourcePartByPath ? { sourcePresent: sourcePartByPath.has(partPath) } : {}),
				...packageActionDigestEvidence(partPath, sourceArchive, outputArchive),
				diagnosticCodes: [diagnostic.code],
			})
		}
	}

	for (const issue of options.packageGraphAudit?.issues ?? []) {
		const partPaths = packageGraphIssuePackagePaths(issue)
		for (const partPath of partPaths.length > 0 ? partPaths : [undefined]) {
			actions.push({
				action: 'error',
				reason: issue.message,
				...(partPath ? { partPath } : {}),
				...(partPath && sourcePartByPath ? { sourcePresent: sourcePartByPath.has(partPath) } : {}),
				...packageActionDigestEvidence(partPath, sourceArchive, outputArchive),
				auditIssueCodes: [issue.code],
			})
		}
	}

	const byAction = emptyPackageActionCounts()
	for (const action of actions) byAction[action.action] += 1
	const auditIssues = options.packageGraphAudit?.issues ?? []
	const sourceByteDigestCount = actions.filter((action) => action.sourceSha256 !== undefined).length
	const outputByteDigestCount = actions.filter((action) => action.outputSha256 !== undefined).length
	const issues = actions
		.filter((action) => action.action === 'error')
		.map((action) => (action.partPath ? `${action.partPath}: ${action.reason}` : action.reason))
	return {
		formatVersion: 1,
		kind: 'ascend-package-action-proof',
		totalActions: actions.length,
		byAction,
		coverage: {
			proofScope: 'package-part-actions-with-audit-summaries',
			sourceGraphIncluded: sourcePackageGraph !== undefined,
			...(sourcePackageGraph
				? {
						sourcePartCount: sourcePackageGraph.parts.length,
						sourceRelationshipCount: sourcePackageGraph.relationships?.length ?? 0,
					}
				: {}),
			writePolicyIncluded: options.writePolicy !== undefined,
			packageGraphAuditIncluded: options.packageGraphAudit !== undefined,
			relationshipAuditIssueCount: auditIssues.filter(isPackageRelationshipAuditIssue).length,
			bytePreservationAuditIssueCount: auditIssues.filter(isPackageBytePreservationAuditIssue)
				.length,
			sourceByteDigestCount,
			outputByteDigestCount,
			matchingByteDigestCount: actions.filter((action) => action.bytesEqual === true).length,
			mismatchedByteDigestCount: actions.filter((action) => action.bytesEqual === false).length,
		},
		actions,
		issues,
		claimBoundaries: options.claimBoundaries ?? DEFAULT_PACKAGE_ACTION_PROOF_CLAIM_BOUNDARIES,
	}
}

export function createAgentCommitPackageActionProof(
	result: AgentCommitResult,
	options: AgentCommitPackageActionProofOptions = {},
): PackageActionProof {
	const bytes = COMMIT_PACKAGE_ACTION_PROOF_BYTES.get(result)
	return createPackageActionProof(result.preservation, {
		...options,
		...(bytes ? { sourceBytes: bytes.sourceBytes, outputBytes: bytes.outputBytes } : {}),
		writePolicy: result.writePolicy,
		packageGraphAudit: result.packageGraphAudit,
	})
}

function compactTraceSummary(trace: AgentWorkflowTrace): CompactAgentTraceSummary {
	return {
		kind: trace.kind,
		file: trace.file,
		inputSha256: trace.inputSha256,
		...(trace.outputSha256 ? { outputSha256: trace.outputSha256 } : {}),
		...(trace.planDigest ? { planDigest: trace.planDigest } : {}),
		traceDigest: trace.traceDigest,
		phaseCount: trace.phases.length,
		warningCount: trace.phases.filter((phase) => phase.status === 'warning').length,
		blockedCount: trace.phases.filter((phase) => phase.status === 'blocked').length,
		failedCount: trace.phases.filter((phase) => phase.status === 'failed').length,
		artifactCount: trace.artifacts.length,
		phases: trace.phases.map(({ phase, status, summary, count, refs }) => ({
			phase,
			status,
			summary,
			...(count !== undefined ? { count } : {}),
			...(refs ? { refs } : {}),
		})),
	}
}

function compactCheckSummary(check: ReturnType<AscendWorkbook['check']>): CompactCheckSummary {
	return {
		valid: check.valid,
		issueCount: check.issues.length,
		errorCount: check.issues.filter((issue) => issue.severity === 'error').length,
		warningCount: check.issues.filter((issue) => issue.severity === 'warning').length,
		infoCount: check.issues.filter((issue) => issue.severity === 'info').length,
	}
}

function compactLintSummary(lint: ReturnType<AscendWorkbook['lint']>): CompactLintSummary {
	return {
		clean: lint.clean,
		warningCount: lint.warnings.length,
		errorCount: lint.warnings.filter((warning) => warning.severity === 'error').length,
		parseErrorCount: lint.warnings.filter((warning) => warning.rule === 'parse-error').length,
	}
}

function compactWritePlanSummary(
	preservation: ReturnType<AscendWorkbook['writePlanSummary']>,
): CompactWritePlanSummary {
	return {
		totalParts: preservation.totalParts,
		byOrigin: preservation.byOrigin,
		skippedCapsuleCount: preservation.skippedCapsules.length,
	}
}

function compactWritePolicySummary(writePolicy: WritePolicyReport): CompactWritePolicySummary {
	return {
		ok: writePolicy.ok,
		diagnosticCount: writePolicy.diagnostics.length,
		blockerCount: writePolicy.diagnostics.filter((diagnostic) => diagnostic.severity === 'blocker')
			.length,
		warningCount: writePolicy.diagnostics.filter((diagnostic) => diagnostic.severity === 'warning')
			.length,
		summary: writePolicy.summary,
	}
}

const DEFAULT_RELEASE_PROOF_CLAIM_BOUNDARIES = [
	'This bundle is local workbook evidence, not a signed SLSA or GitHub artifact attestation.',
	'The bundle reports Ascend plan, commit, reopen, diff, and audit results; it does not claim Excel recalculation equivalence.',
	'Private workbook data should not be embedded unless the caller explicitly chooses to persist full artifacts.',
	'Signed provenance claims require an external attestation envelope and verifier roots.',
] as const

const DEFAULT_PACKAGE_ACTION_PROOF_CLAIM_BOUNDARIES = [
	'This proof is local package-part evidence, not signed provenance or third-party attestation.',
	'Package actions describe Ascend write-plan and audit evidence; they do not prove Excel semantic recalculation equivalence.',
	'Passthrough byte equality is proven only when source and output byte digests are present and bytesEqual is true.',
	'Drop and error actions require caller review before claiming workbook feature preservation.',
] as const

const COMMIT_PACKAGE_ACTION_PROOF_BYTES = new WeakMap<
	AgentCommitResult,
	{ readonly sourceBytes: Uint8Array; readonly outputBytes: Uint8Array }
>()

function traceArtifactByName(
	trace: AgentWorkflowTrace,
	name: string,
): AgentTraceArtifact | undefined {
	return trace.artifacts.find((artifact) => artifact.name === name)
}

function releaseProofArtifactDigest(artifact: AgentTraceArtifact): ReleaseProofArtifactDigest {
	return {
		name: artifact.name,
		digest: artifact.digest,
		summary: artifact.summary,
	}
}

function releaseProofCheck(
	check: string,
	ok: boolean,
	details?: unknown,
): ReleaseProofConsistencyCheck {
	return {
		check,
		ok,
		...(details !== undefined ? { details } : {}),
	}
}

function packageActionForPlannedPart(
	part: WorkbookWritePlanPartInfo,
	sourcePresent: boolean | undefined,
): PackageActionKind {
	if (part.origin !== 'generated') return 'passthrough'
	return sourcePresent === false ? 'add' : 'regenerate'
}

function packageActionReason(
	action: PackageActionKind,
	part: WorkbookWritePlanPartInfo,
	sourcePresent: boolean | undefined,
): string {
	switch (action) {
		case 'passthrough':
			return `${part.origin} package part will be copied through from preserved source or capsule bytes.`
		case 'add':
			return 'Generated package part is absent from the source package graph and will be added.'
		case 'regenerate':
			return sourcePresent === undefined
				? 'Generated package part will be emitted; source graph was not provided, so add vs regenerate cannot be proven.'
				: 'Generated package part replaces an existing source package part.'
		case 'drop':
			return 'Package part will not be emitted by the planned write.'
		case 'error':
			return 'Package action is blocked by diagnostic or package graph audit evidence.'
	}
}

function packageActionProofArchive(
	bytes: Uint8Array | undefined,
): PackageActionProofArchive | undefined {
	if (!bytes) return undefined
	return extractZip(bytes)
}

function packageActionDigestEvidence(
	partPath: string | undefined,
	sourceArchive: PackageActionProofArchive | undefined,
	outputArchive: PackageActionProofArchive | undefined,
): Pick<PackageActionProofEntry, 'sourceSha256' | 'outputSha256' | 'bytesEqual'> {
	if (!partPath) return {}
	const sourceBytes = sourceArchive?.readBytes(partPath)
	const outputBytes = outputArchive?.readBytes(partPath)
	const sourceSha256 = sourceBytes ? sha256Bytes(sourceBytes) : undefined
	const outputSha256 = outputBytes ? sha256Bytes(outputBytes) : undefined
	return {
		...(sourceSha256 ? { sourceSha256 } : {}),
		...(outputSha256 ? { outputSha256 } : {}),
		...(sourceSha256 && outputSha256 ? { bytesEqual: sourceSha256 === outputSha256 } : {}),
	}
}

function emptyPackageActionCounts(): Record<PackageActionKind, number> {
	return {
		passthrough: 0,
		regenerate: 0,
		add: 0,
		drop: 0,
		error: 0,
	}
}

function isPackageRelationshipAuditIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	return (
		issue.code === 'package_relationship_duplicate_id' ||
		issue.code === 'package_relationship_source' ||
		issue.code === 'package_relationship_target' ||
		issue.code === 'package_preserved_relationship' ||
		issue.code === 'package_preserved_relationship_identity'
	)
}

function isPackageBytePreservationAuditIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	return issue.code === 'package_preserved_part_bytes'
}

function packageActionCountsEqual(
	left: Readonly<Record<PackageActionKind, number>>,
	right: Readonly<Record<PackageActionKind, number>>,
): boolean {
	for (const action of Object.keys(left) as PackageActionKind[]) {
		if (left[action] !== right[action]) return false
	}
	return true
}

function diagnosticsByPartPath(
	diagnostics: readonly WritePolicyDiagnostic[],
): Map<string, WritePolicyDiagnostic[]> {
	const byPath = new Map<string, WritePolicyDiagnostic[]>()
	for (const diagnostic of diagnostics) {
		for (const partPath of diagnostic.partPaths ?? []) {
			const existing = byPath.get(partPath)
			if (existing) existing.push(diagnostic)
			else byPath.set(partPath, [diagnostic])
		}
	}
	return byPath
}

function auditIssuesByPartPath(
	issues: readonly XlsxPackageGraphFidelityIssue[],
): Map<string, XlsxPackageGraphFidelityIssue[]> {
	const byPath = new Map<string, XlsxPackageGraphFidelityIssue[]>()
	for (const issue of issues) {
		for (const partPath of packageGraphIssuePackagePaths(issue)) {
			const existing = byPath.get(partPath)
			if (existing) existing.push(issue)
			else byPath.set(partPath, [issue])
		}
	}
	return byPath
}

function compactPostWriteVerification(
	postWrite: AgentPostWriteVerification,
): CompactAgentPostWriteVerification {
	return {
		valid: postWrite.valid,
		output: postWrite.output,
		outputSha256: postWrite.outputSha256,
		reopened: postWrite.reopened,
		auditsPassed: postWrite.auditsPassed,
		...(postWrite.timings ? { timings: postWrite.timings } : {}),
		check: compactCheckSummary(postWrite.check),
		lint: compactLintSummary(postWrite.lint),
		preservation: compactWritePlanSummary(postWrite.preservation),
		opaquePayloads: postWrite.opaquePayloads,
		comments: postWrite.comments,
		hyperlinks: postWrite.hyperlinks,
		dataValidations: postWrite.dataValidations,
		dataConnections: postWrite.dataConnections,
		tables: postWrite.tables,
		definedNames: postWrite.definedNames,
		externalReferences: postWrite.externalReferences,
		formulaState: postWrite.formulaState,
		workbookTopology: postWrite.workbookTopology,
		analytics: postWrite.analytics,
		activeContent: postWrite.activeContent,
		visuals: postWrite.visuals,
		security: postWrite.security,
		packageGraphAudit: compactPackageGraphAuditSummary(postWrite.packageGraphAudit),
		expectedPackageGraphIssueCount: postWrite.expectedPackageGraphIssueCount,
		unresolvedPackageGraphIssueCount: postWrite.unresolvedPackageGraphIssueCount,
	}
}

function compactLossAuditSummary(lossAudit: LossAudit): CompactLossAuditSummary {
	return {
		ok: lossAudit.ok,
		blockedFeatureCount: lossAudit.blockedFeatures.length,
		blockedPackagePartCount: lossAudit.blockedPackageParts.length,
		allowedLoss: lossAudit.allowedLoss,
		policy: lossAudit.policy,
	}
}

function compactPackageGraphAuditSummary(
	audit: PackageGraphAudit,
): CompactPackageGraphAuditSummary {
	const issues = audit.issues.slice(0, 20)
	return {
		ok: audit.ok,
		issueCount: audit.issues.length,
		emittedIssueCount: issues.length,
		issues,
		policy: audit.policy,
	}
}

async function resolveCommitOutputTarget(
	file: string,
	inputSha256: string,
	options: AgentCommitOptions,
	progress: ReturnType<typeof createProgressEmitter>,
): Promise<string> {
	if (options.expectSha256 && options.expectSha256 !== inputSha256) {
		await progress('hash-guard', 'failed', 'Input hash did not match expected SHA-256.', {
			details: { expected: options.expectSha256, actual: inputSha256 },
		})
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Input workbook hash does not match --expect-sha256', {
				details: { expected: options.expectSha256, actual: inputSha256 },
				suggestedFix: 'Re-run ascend plan and commit with the new inputSha256.',
			}),
		)
	}
	await progress(
		'hash-guard',
		'ok',
		options.expectSha256
			? 'Input hash matched expected SHA-256.'
			: 'No input hash guard requested.',
	)

	const output = options.inPlace ? file : options.output
	if (!output) {
		await progress('output-policy', 'failed', 'Commit output target is missing.')
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', 'Commit requires --output unless --in-place is set', {
				suggestedFix: 'Pass --output out.xlsx or --in-place with an optional --backup path.',
			}),
		)
	}
	if (!options.inPlace && options.backup) {
		await progress('output-policy', 'failed', '--backup was provided without --in-place.')
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', '--backup is only valid with --in-place commits', {
				suggestedFix:
					'Use --output for non-destructive writes, or add --in-place when creating a backup.',
			}),
		)
	}
	await progress('output-policy', 'ok', `Commit output target resolved to ${output}.`)
	return output
}

export async function commitAgentPlan(
	file: string,
	ops: readonly Operation[],
	options: AgentCommitOptions = {},
): Promise<AgentCommitResult> {
	const progress = createProgressEmitter('commit', options.onProgress)
	await progress('hash-input', 'started', 'Hashing input workbook.')
	const source = await readStableAgentSource(file)
	const inputSha256 = sha256Bytes(source.sourceBytes)
	await progress('hash-input', 'ok', 'Input workbook hash captured.')
	const output = await resolveCommitOutputTarget(file, inputSha256, options, progress)
	await progress('load-workbook', 'started', 'Opening workbook.')
	const wb = await openWorkbookFromBytes(file, source.sourceBytes, passwordOpenOptions(options))
	await progress('load-workbook', 'ok', 'Workbook opened.')
	return commitAgentPlanFromWorkbook(file, inputSha256, wb, ops, options, {
		progress,
		output,
		sourceBytes: source.sourceBytes,
	})
}

export async function commitAgentPlanFromWorkbook(
	file: string,
	inputSha256: string,
	wb: AscendWorkbook,
	ops: readonly Operation[],
	options: AgentCommitOptions = {},
	internal: {
		readonly progress?: ReturnType<typeof createProgressEmitter>
		readonly output?: string
		readonly sourceBytes?: Uint8Array
		readonly preparedCheck?: ReturnType<AscendWorkbook['check']>
	} = {},
): Promise<AgentCommitResult> {
	const progress = internal.progress ?? createProgressEmitter('commit', options.onProgress)
	const load = wb.inspect().load
	if (load.isPartial) {
		throw new AscendException(
			ascendError(
				'VALIDATION_ERROR',
				'Cannot commit an agent write plan from a partial workbook view',
				{
					details: { load },
					suggestedFix: 'Reopen the workbook with a full load before planning or committing edits.',
				},
			),
		)
	}
	const output =
		internal.output ?? (await resolveCommitOutputTarget(file, inputSha256, options, progress))
	if (wb.sourceInfo().sourceWasEncrypted && !options.allowDecryptedExport) {
		const error = encryptedWorkbookCommitExportError(output)
		await progress('write-policy', 'failed', error.ascendError.message, error.ascendError.details)
		throw error
	}
	const workbookModel = wb.getWorkbookModel()
	const packageGraphResult = await timedCommitStep(() => wb.packageGraph())
	const packageGraph = packageGraphResult.value
	const sourceBytesForPackageAudit =
		wb.sourceInfo().sourceWasEncrypted && options.allowDecryptedExport
			? wb.toBytes({ allowDecryptedExport: true })
			: undefined
	const useLiveWritePolicyWorkbook = canUseLiveWritePolicyWorkbook(workbookModel, ops)
	const writePolicySnapshotResult = useLiveWritePolicyWorkbook
		? { value: workbookModel, ms: 0 }
		: await timedCommitStep(() => snapshotWritePolicyWorkbook(workbookModel))
	const writePolicyWorkbook = writePolicySnapshotResult.value
	const expectedPostWritePackageGraphChanges = expectedPackageGraphChangesForOperations(
		writePolicyWorkbook,
		ops,
		packageGraph,
	)
	await progress('approval-audit', 'started', 'Auditing explicit approval requirements.')
	const approvalsResult = await timedCommitStep(() =>
		buildApprovalRequirements(wb.report.features, ops),
	)
	const approvals = approvalsResult.value
	const effectiveAllowLoss = mergeAllowLoss(
		options.allowLoss,
		approvalSatisfiedLossFeatures(approvals, options.approvals),
	)
	await progress('loss-audit', 'started', 'Auditing preserved and unsupported features.')
	const lossAuditResult = await timedCommitStep(() =>
		auditLossPolicy(wb.report.features, effectiveAllowLoss, packageGraph),
	)
	const lossAudit = lossAuditResult.value
	await progressFromPhase(lossAuditPhase(lossAudit), progress)
	await progress('package-graph-audit', 'started', 'Auditing package graph fidelity.')
	const packageGraphAuditResult = await timedCommitStep(() =>
		auditPackageGraphIntegrity(packageGraph),
	)
	const packageGraphAudit = packageGraphAuditResult.value
	const blockedApprovals = unsatisfiedApprovalRequirements(approvals, lossAudit, options.approvals)
	await progressFromPhase(packageGraphAuditPhase(packageGraphAudit), progress)
	await progressFromPhase(approvalPhase(approvals, options.approvals, blockedApprovals), progress)
	if (blockedApprovals.length > 0) {
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Commit requires explicit approval', {
				details: { approvals: blockedApprovals, lossAudit },
				suggestedFix:
					'Inspect plan.approvals. If the action is intentional, pass --approval <id> for each requirement, --approval all, or use --allow-loss for lossy workbook features.',
			}),
		)
	}
	const rollbackSnapshot = wb.createMutationRollbackSnapshot()
	let apply: ReturnType<AscendWorkbook['apply']>
	let recalc: ReturnType<AscendWorkbook['recalc']> | null = null
	let recalcMs = 0
	let preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	let writePolicy: WritePolicyReport
	let writePolicyCheck: ReturnType<AscendWorkbook['check']>
	let applyMs = 0
	let preservationMs = 0
	let writePolicyCheckMs = 0
	let writePolicyBuildMs = 0
	try {
		await progress('apply', 'started', 'Applying operations.', { count: ops.length })
		const applyResult = await timedCommitStep(() =>
			wb.apply(ops, { transaction: true, journal: true }),
		)
		apply = applyResult.value
		applyMs = applyResult.ms
		await progressFromPhase(applyPhase(apply), progress)
		if (apply.errors.length > 0)
			throw new AscendException(apply.errors[0] ?? ascendError('VALIDATION_ERROR', 'Apply failed'))

		if (apply.recalcRequired) {
			await progress('recalc', 'started', 'Recalculating formulas.')
			const recalcResult = await timedCommitStep(() => wb.recalc())
			recalc = recalcResult.value
			recalcMs = recalcResult.ms
			await progressFromPhase(recalcPhase(recalc), progress)
			if (recalc.errors.length > 0) {
				const first = recalc.errors[0]
				throw new AscendException(
					first
						? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
								refs: [first.ref],
								details: { evalError: first.error },
								suggestedFix:
									'Run ascend plan to inspect the failing formula state before committing.',
							})
						: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
				)
			}
		} else {
			await progressFromPhase(recalcPhase(recalc), progress)
		}

		await progress('preservation-audit', 'started', 'Summarizing package preservation.')
		const preservationResult = await timedCommitStep(() => wb.writePlanSummary())
		preservation = preservationResult.value
		preservationMs = preservationResult.ms
		await progressFromPhase(preservationPhase(preservation), progress)
		await progress('write-policy', 'started', 'Explaining write preservation and loss policy.')
		const writePolicyCheckResult =
			internal.preparedCheck && canReusePreparedCommitCheck(writePolicyWorkbook, ops)
				? { value: internal.preparedCheck, ms: 0 }
				: await timedCommitStep(() => wb.check())
		writePolicyCheck = writePolicyCheckResult.value
		writePolicyCheckMs = writePolicyCheckResult.ms
		const writePolicyResult = await timedCommitStep(() =>
			buildWritePolicyReport(
				wb.report.features,
				packageGraph,
				preservation,
				packageGraphAudit,
				writePolicyWorkbook,
				ops,
				wb.inspect().externalReferenceDetails,
				writePolicyCheck.issues,
			),
		)
		writePolicy = writePolicyResult.value
		writePolicyBuildMs = writePolicyResult.ms
		await progressFromPhase(writePolicyPhase(writePolicy), progress)
		if (writePolicy.diagnostics.some((diagnostic) => diagnostic.severity === 'blocker')) {
			throw new AscendException(
				ascendError('VALIDATION_ERROR', 'Commit blocked by write policy', {
					details: { writePolicy },
					suggestedFix:
						'Inspect writePolicy.diagnostics and repair blocker diagnostics before committing.',
				}),
			)
		}
	} catch (error) {
		wb.restoreMutationRollbackSnapshot(rollbackSnapshot)
		throw error
	}
	if (options.inPlace && options.backup) {
		await progress('backup', 'started', `Creating in-place backup at ${options.backup}.`)
		try {
			await copyFile(file, options.backup)
		} catch (error) {
			wb.restoreMutationRollbackSnapshot(rollbackSnapshot)
			const backupError = workbookBackupError(file, options.backup, error)
			await progress(
				'backup',
				'failed',
				backupError.ascendError.message,
				backupError.ascendError.details,
			)
			throw backupError
		}
		await progress('backup', 'ok', `In-place backup created at ${options.backup}.`)
	}
	const sourceBytes = internal.sourceBytes ?? (await fileBytes(file))
	const sourcePackageBytes = sourceBytesForPackageAudit ?? sourceBytes
	await progress('write', 'started', `Writing workbook to ${output}.`)
	let writeTimings: AtomicWorkbookWriteTimings
	try {
		writeTimings = await writeWorkbookAtomically(wb, output, {
			...(options.allowDecryptedExport ? { allowDecryptedExport: true } : {}),
			allowSignatureInvalidation: true,
			...(shouldAllowActiveContentLossForTextExport(wb, output, effectiveAllowLoss)
				? { allowActiveContentLoss: true }
				: {}),
		})
	} catch (error) {
		wb.restoreMutationRollbackSnapshot(rollbackSnapshot)
		const writeError = atomicWorkbookWriteError(output, error)
		await progress(
			'write',
			'failed',
			writeError.ascendError.message,
			writeError.ascendError.details,
		)
		throw writeError
	}
	const outputBytesResult = await timedCommitStep(() => fileBytes(output))
	const outputBytes = outputBytesResult.value
	const outputSha256Result = await timedCommitStep(() => sha256Bytes(outputBytes))
	const outputSha256 = outputSha256Result.value
	await progress('write', 'ok', `Workbook written to ${output}.`)
	const timings: AgentCommitTimings = {
		writePolicySnapshotMs: writePolicySnapshotResult.ms,
		packageGraphMs: packageGraphResult.ms,
		approvalAuditMs: approvalsResult.ms,
		lossAuditMs: lossAuditResult.ms,
		packageGraphAuditMs: packageGraphAuditResult.ms,
		applyMs,
		recalcMs,
		writePlanSummaryMs: preservationMs,
		writePolicyCheckMs,
		writePolicyBuildMs,
		toBytesMs: writeTimings.toBytesMs,
		writeFileMs: writeTimings.writeFileMs,
		renameMs: writeTimings.renameMs,
		outputByteReadMs: outputBytesResult.ms,
		outputHashMs: outputSha256Result.ms,
	}
	await progress('post-write', 'started', 'Reopening written workbook for verification.')
	let postWrite: AgentPostWriteVerification
	try {
		postWrite = await verifyWrittenWorkbook(
			output,
			outputSha256,
			packageGraph,
			sourcePackageBytes,
			expectedPostWritePackageGraphChanges,
			progress,
			ops,
			wb.getWorkbookModel(),
		)
	} catch (error) {
		wb.restoreMutationRollbackSnapshot(rollbackSnapshot)
		throw error
	}
	await progressFromPhase(postWritePhase(postWrite, expectedPostWritePackageGraphChanges), progress)
	await progress('check', 'started', 'Reusing pre-write structural checks.')
	const check = writePolicyCheck
	await progressFromPhase(checkPhase(check), progress)
	await progress('lint', 'started', 'Running formula lint.')
	const lint = wb.lint()
	await progressFromPhase(lintPhase(lint), progress)
	const planDigest = digestPlan(inputSha256, ops)
	await progress('finalize', 'ok', 'Commit digest and trace finalized.')
	const trace = finalizeTrace({
		kind: 'commit',
		file,
		inputSha256,
		outputSha256,
		planDigest,
		operationCount: ops.length,
		phases: [
			okPhase('hash-input', 'Input workbook hash captured.'),
			options.expectSha256
				? okPhase('hash-guard', 'Input hash matched expected SHA-256.')
				: okPhase('hash-guard', 'No input hash guard requested.'),
			lossAuditPhase(lossAudit),
			packageGraphAuditPhase(packageGraphAudit),
			approvalPhase(approvals, options.approvals, blockedApprovals),
			applyPhase(apply),
			recalcPhase(recalc),
			preservationPhase(preservation),
			writePolicyPhase(writePolicy),
			okPhase('write', `Workbook written to ${output}.`),
			postWritePhase(postWrite, expectedPostWritePackageGraphChanges),
			checkPhase(check),
			lintPhase(lint),
		],
		artifacts: [
			artifact('ops', ops, `${ops.length} operation(s)`),
			artifact('apply', apply, `${apply.affectedCells.length} affected cell(s)`),
			artifact(
				'recalc',
				recalc,
				recalc ? `${recalc.changed.length} recalculated cell(s)` : 'not required',
			),
			artifact('lossAudit', lossAudit, `${lossAudit.blockedFeatures.length} blocked feature(s)`),
			artifact(
				'packageGraphAudit',
				packageGraphAudit,
				`${packageGraphAudit.issues.length} package graph issue(s)`,
			),
			artifact('approvals', approvals, `${approvals.length} approval requirement(s)`),
			artifact('preservation', preservation, `${preservation.totalParts} package part(s)`),
			artifact(
				'writePolicy',
				writePolicy,
				`${writePolicy.diagnostics.length} write policy diagnostic(s)`,
			),
			artifact(
				'postWrite',
				postWrite,
				postWrite.valid
					? 'written workbook reopened and verified'
					: 'written workbook failed verification',
			),
		],
	})

	const result: AgentCommitResult = {
		file,
		output,
		...(options.inPlace && options.backup ? { backup: options.backup } : {}),
		inputSha256,
		outputSha256,
		planDigest,
		operationCount: ops.length,
		approvals,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		apply,
		recalc,
		timings,
		check,
		lint,
		preservation,
		writePolicy,
		postWrite,
		lossAudit,
		packageGraphAudit,
	}
	COMMIT_PACKAGE_ACTION_PROOF_BYTES.set(result, { sourceBytes: sourcePackageBytes, outputBytes })
	return result
}

async function verifyWrittenWorkbook(
	output: string,
	outputSha256: string,
	sourceGraph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
	expectedPackageGraphChanges: ExpectedPostWritePackageGraphChanges,
	progress: ReturnType<typeof createProgressEmitter>,
	ops: readonly Operation[],
	sourceWorkbook: Workbook,
): Promise<AgentPostWriteVerification> {
	const outputSnapshot = await timedPostWriteStep(
		progress,
		'output-snapshot',
		'Reading post-write output snapshot.',
		'Post-write output snapshot matched written bytes.',
		async () => {
			const snapshot = await readStablePostWriteOutputSnapshot(output)
			if (snapshot.identity.sha256 !== outputSha256) {
				throw postWriteOutputChangedError(output, outputSha256, snapshot.identity.sha256)
			}
			return snapshot
		},
	)
	const reopened = await timedPostWriteStep(
		progress,
		'reopen',
		'Reopening written workbook.',
		'Written workbook reopened.',
		() =>
			openPostWriteDocument(
				outputSnapshot.value.identity,
				outputSnapshot.value.bytes,
				postWriteOpenOptions(sourceGraph, ops, sourceWorkbook),
			),
	)
	const check = await timedPostWriteStep(
		progress,
		'check',
		'Running post-write structural checks.',
		'Post-write structural checks completed.',
		() => reopened.value.check(),
	)
	const lint = await timedPostWriteStep(
		progress,
		'lint',
		'Running post-write formula lint.',
		'Post-write formula lint completed.',
		() => reopened.value.lint(),
	)
	const preservation = await timedPostWriteStep(
		progress,
		'preservation',
		'Summarizing post-write preservation.',
		'Post-write preservation summary completed.',
		() => reopened.value.writePlanSummary(),
	)
	const workbook = reopened.value.getWorkbookModel()
	const opaquePayloads = postWriteOpaquePayloadSummary(workbook)
	const comments = postWriteCommentSummary(workbook)
	const hyperlinks = postWriteHyperlinkSummary(workbook)
	const dataValidations = postWriteDataValidationSummary(workbook)
	const dataConnections = postWriteDataConnectionSummary(workbook)
	const tables = postWriteTableSummary(workbook)
	const definedNames = postWriteDefinedNameSummary(workbook)
	const externalReferences = postWriteExternalReferenceSummary(workbook)
	const analytics = postWriteAnalyticsSummary(workbook)
	const activeContent = postWriteActiveContentSummary(workbook)
	const visuals = postWriteVisualSummary(workbook)
	const security = postWriteSecuritySummary(workbook)
	const workbookTopology = postWriteWorkbookTopologySummary(workbook)
	const outputIsXlsx = reopened.value.inspect().sourceFormat === 'xlsx'
	const outputGraph = outputIsXlsx
		? await timedPostWriteStep(
				progress,
				'package-graph',
				'Reading post-write package graph.',
				'Post-write package graph read completed.',
				() => reopened.value.packageGraph(),
			)
		: { value: null, ms: 0 }
	const formulaState = postWriteFormulaSummary(workbook, outputGraph.value, outputIsXlsx)
	const packageGraphAudit = outputGraph.value
		? await timedPostWriteStep(
				progress,
				'package-graph-audit',
				'Auditing post-write package graph roundtrip.',
				'Post-write package graph roundtrip audit completed.',
				() =>
					auditPackageGraphRoundtrip(
						sourceGraph,
						sourceBytes,
						outputGraph.value,
						outputSnapshot.value.bytes,
					),
			)
		: { value: nonXlsxPackageGraphAudit(), ms: 0 }
	const unresolvedPackageGraphIssues = packageGraphAudit.value.issues.filter(
		(issue) => !isExpectedPostWritePackageGraphIssue(issue, expectedPackageGraphChanges),
	)
	const lintFailures = postWriteLintFailures(lint.value)
	const expectedPackageGraphIssueCount =
		packageGraphAudit.value.issues.length - unresolvedPackageGraphIssues.length
	return {
		valid: check.value.valid,
		auditsPassed:
			check.value.valid && lintFailures.length === 0 && unresolvedPackageGraphIssues.length === 0,
		output,
		outputSha256,
		reopened: true,
		timings: {
			outputSnapshotMs: outputSnapshot.ms,
			reopenMs: reopened.ms,
			checkMs: check.ms,
			lintMs: lint.ms,
			preservationMs: preservation.ms,
			packageGraphMs: outputGraph.ms,
			packageGraphAuditMs: packageGraphAudit.ms,
		},
		check: check.value,
		lint: lint.value,
		preservation: preservation.value,
		opaquePayloads,
		comments,
		hyperlinks,
		dataValidations,
		dataConnections,
		tables,
		definedNames,
		externalReferences,
		formulaState,
		workbookTopology,
		analytics,
		activeContent,
		visuals,
		security,
		packageGraphAudit: packageGraphAudit.value,
		expectedPackageGraphIssueCount,
		unresolvedPackageGraphIssueCount: unresolvedPackageGraphIssues.length,
	}
}

function postWriteCommentSummary(workbook: Workbook): PostWriteCommentSummary {
	const legacyComments = collectLegacyCommentLocations(workbook)
	const threadedComments = collectThreadedCommentLocations(workbook)
	return {
		legacyCommentLocations: legacyComments.length,
		threadedCommentLocations: threadedComments.length,
		legacyDrawingLocations: legacyComments.filter((comment) => comment.hasLegacyDrawing).length,
		locations: uniqueStrings([
			...legacyComments.map((comment) => comment.location),
			...threadedComments.map((comment) => comment.location),
		]),
		threadedCommentPartPaths: uniqueStrings(
			threadedComments.flatMap((comment) => (comment.partPath ? [comment.partPath] : [])),
		),
		verification: 'reopened-output',
	}
}

function postWriteHyperlinkSummary(workbook: Workbook): PostWriteHyperlinkSummary {
	const links = workbook.sheets.flatMap((sheet) =>
		[...sheet.hyperlinks.entries()].map(([ref, hyperlink]) => ({
			sheetName: sheet.name,
			ref,
			location: `${sheet.name}!${ref}`,
			...(hyperlink.target ? { target: hyperlink.target } : {}),
			...(hyperlink.location ? { internalLocation: hyperlink.location } : {}),
			...(hyperlink.display ? { display: hyperlink.display } : {}),
			...(hyperlink.tooltip ? { tooltip: hyperlink.tooltip } : {}),
		})),
	)
	return {
		total: links.length,
		externalTargets: links.filter((link) => link.target !== undefined).length,
		internalLocations: links.filter((link) => link.internalLocation !== undefined).length,
		displayed: links.filter((link) => link.display !== undefined).length,
		withTooltips: links.filter((link) => link.tooltip !== undefined).length,
		locations: links.map((link) => link.location),
		targets: uniqueStrings(links.flatMap((link) => (link.target ? [link.target] : []))),
		internalLocationTargets: uniqueStrings(
			links.flatMap((link) => (link.internalLocation ? [link.internalLocation] : [])),
		),
		links,
		preservationMode: links.length > 0 ? 'generated' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteDataValidationSummary(workbook: Workbook): PostWriteDataValidationSummary {
	const validations = workbook.sheets.flatMap((sheet) =>
		sheet.dataValidations.map((validation) => ({
			sheetName: sheet.name,
			sqref: validation.sqref,
			location: `${sheet.name}!${validation.sqref}`,
			...(validation.type !== undefined ? { type: validation.type } : {}),
			...(validation.operator !== undefined ? { operator: validation.operator } : {}),
			...(validation.source !== undefined ? { source: validation.source } : {}),
			...(validation.formula1 !== undefined ? { formula1: validation.formula1 } : {}),
			...(validation.formula2 !== undefined ? { formula2: validation.formula2 } : {}),
			...(validation.allowBlank !== undefined ? { allowBlank: validation.allowBlank } : {}),
			...(validation.showInputMessage !== undefined
				? { showInputMessage: validation.showInputMessage }
				: {}),
			...(validation.showErrorMessage !== undefined
				? { showErrorMessage: validation.showErrorMessage }
				: {}),
		})),
	)
	return {
		total: validations.length,
		formulaBacked: validations.filter(
			(validation) => validation.formula1 !== undefined || validation.formula2 !== undefined,
		).length,
		listValidations: validations.filter((validation) => validation.type === 'list').length,
		x14Validations: validations.filter((validation) => validation.source === 'x14').length,
		ranges: validations.map((validation) => validation.location),
		types: uniqueStrings(
			validations.flatMap((validation) => (validation.type ? [validation.type] : [])),
		),
		validations,
		preservationMode: validations.length > 0 ? 'generated' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteDataConnectionSummary(workbook: Workbook): PostWriteDataConnectionSummary {
	const connections = workbook.connectionParts.map((part) => ({
		kind: part.kind,
		partPath: part.partPath,
		state: postWriteDataConnectionState(part),
		...(part.sheetName !== undefined ? { sheetName: part.sheetName } : {}),
		...(part.name !== undefined ? { name: part.name } : {}),
		...(part.connectionId !== undefined ? { connectionId: part.connectionId } : {}),
		...(part.refreshOnLoad !== undefined ? { refreshOnLoad: part.refreshOnLoad } : {}),
		...(part.saveData !== undefined ? { saveData: part.saveData } : {}),
		...(part.refreshedVersion !== undefined ? { refreshedVersion: part.refreshedVersion } : {}),
	}))
	return {
		total: connections.length,
		workbookConnections: connections.filter((connection) => connection.kind === 'connection')
			.length,
		queryTables: connections.filter((connection) => connection.kind === 'queryTable').length,
		refreshOnOpen: connections.filter((connection) => connection.state === 'refresh-on-open')
			.length,
		notSaved: connections.filter(
			(connection) => connection.state === 'not-saved' || connection.saveData === false,
		).length,
		unknown: connections.filter((connection) => connection.state === 'unknown').length,
		cached: connections.filter((connection) => connection.state === 'cached').length,
		partPaths: uniqueStrings(connections.map((connection) => connection.partPath)),
		names: uniqueStrings(
			connections.flatMap((connection) => (connection.name ? [connection.name] : [])),
		),
		connectionIds: uniqueNumbers(
			connections.flatMap((connection) =>
				connection.connectionId !== undefined ? [connection.connectionId] : [],
			),
		),
		connections,
		preservationMode: connections.length > 0 ? 'preserve-exact' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteDataConnectionState(
	part: Workbook['connectionParts'][number],
): PostWriteDataConnectionEntry['state'] {
	if (part.refreshOnLoad) return 'refresh-on-open'
	if (part.saveData === false) return 'not-saved'
	if (part.refreshedVersion === undefined) return 'unknown'
	return 'cached'
}

function postWriteTableSummary(workbook: Workbook): PostWriteTableSummary {
	const tables = workbook.sheets.flatMap((sheet) =>
		sheet.tables.map((table) => ({
			sheetName: sheet.name,
			tableName: table.name,
			location: `${sheet.name}!${rangeRefToA1(table.ref)}`,
			...(table.partPath ? { partPath: table.partPath } : {}),
			hasAutoFilter: table.autoFilter != null,
			...(table.queryTable?.partPath ? { queryTablePartPath: table.queryTable.partPath } : {}),
		})),
	)
	return {
		tableLocations: tables.length,
		queryTableLocations: tables.filter((table) => table.queryTablePartPath !== undefined).length,
		tableAutoFilterLocations: tables.filter((table) => table.hasAutoFilter).length,
		tableNames: tables.map((table) => table.tableName),
		locations: tables.map((table) => table.location),
		tablePartPaths: uniqueStrings(
			tables.flatMap((table) => (table.partPath ? [table.partPath] : [])),
		),
		queryTablePartPaths: uniqueStrings(
			tables.flatMap((table) => (table.queryTablePartPath ? [table.queryTablePartPath] : [])),
		),
		preservationMode: tables.length > 0 ? 'preserve-exact' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteDefinedNameSummary(workbook: Workbook): PostWriteDefinedNameSummary {
	const names = workbook.definedNames.list().map((entry) => {
		let sheet: string | undefined
		if (entry.scope.kind === 'sheet') {
			const scope = entry.scope as Extract<typeof entry.scope, { readonly kind: 'sheet' }>
			sheet = workbook.sheets.find((candidate) => candidate.id === scope.sheetId)?.name
		}
		return {
			name: entry.name,
			formula: entry.formula,
			scope: entry.scope.kind,
			...(sheet ? { sheet } : {}),
			...(entry.hidden ? { hidden: true } : {}),
		}
	})
	return {
		total: names.length,
		workbookScoped: names.filter((entry) => entry.scope === 'workbook').length,
		sheetScoped: names.filter((entry) => entry.scope === 'sheet').length,
		hidden: names.filter((entry) => entry.hidden === true).length,
		names,
		verification: 'reopened-output',
	}
}

function postWriteExternalReferenceSummary(workbook: Workbook): PostWriteExternalReferenceSummary {
	const parts = workbook.externalReferenceDetails.map((entry) => ({
		partPath: entry.partPath,
		...(entry.relId ? { relId: entry.relId } : {}),
		...(entry.externalBookRelId ? { externalBookRelId: entry.externalBookRelId } : {}),
		...(entry.linkRelId ? { linkRelId: entry.linkRelId } : {}),
		...(entry.linkBindingStatus ? { linkBindingStatus: entry.linkBindingStatus } : {}),
		...(entry.linkRelationshipKind ? { linkRelationshipKind: entry.linkRelationshipKind } : {}),
		...(entry.target ? { target: entry.target } : {}),
		...(entry.targetMode ? { targetMode: entry.targetMode } : {}),
	}))
	return {
		total: parts.length,
		boundByExternalBookRelId: parts.filter(
			(entry) => entry.linkBindingStatus === 'externalBookRelId',
		).length,
		fallbackPathRelationships: parts.filter(
			(entry) => entry.linkBindingStatus === 'fallbackPathRelationship',
		).length,
		missingPathRelationships: parts.filter(
			(entry) => entry.linkBindingStatus === 'missingPathRelationship',
		).length,
		partPaths: uniqueStrings(parts.map((entry) => entry.partPath)),
		targets: uniqueStrings(parts.flatMap((entry) => (entry.target ? [entry.target] : []))),
		parts,
		preservationMode: parts.length > 0 ? 'preserve-exact' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteFormulaSummary(
	workbook: Workbook,
	packageGraph: XlsxPackageGraph | null,
	outputIsXlsx: boolean,
): PostWriteFormulaSummary {
	const calcChainParts = outputIsXlsx
		? (packageGraph?.parts ?? [])
				.filter((part) => part.featureFamily === 'preservedCalcChain')
				.map((part) => part.path)
		: []
	const calcSettings = workbook.calcSettings
	const recalculationRequested =
		calcSettings.fullCalcOnLoad ||
		calcSettings.forceFullCalc === true ||
		calcSettings.calcCompleted === false ||
		calcSettings.calcOnSave === true
	const formulaCache = postWriteFormulaCacheSummary(workbook)
	const warnings: string[] = []
	if (calcSettings.calcMode === 'manual') {
		warnings.push('Workbook is in manual calculation mode in the reopened output.')
	}
	if (recalculationRequested) {
		warnings.push('Reopened output calculation settings request recalculation on open or save.')
	}
	if (formulaCache.missingCachedFormulaValues > 0) {
		warnings.push(
			`${formulaCache.missingCachedFormulaValues} reopened formula cell(s) do not carry cached formula values; Ascend is reporting the missing cache rather than claiming recalculation.`,
		)
	}
	if (formulaCache.cachedFormulaValues > 0) {
		warnings.push(
			'Reopened formula cache values are stored workbook values, not proof that Ascend recalculated formulas equivalently to Excel.',
		)
	}
	if (calcChainParts.length > 0) {
		warnings.push(
			'CalcChain is present in the reopened output and is dependency-order metadata, not proof of fresh formula values.',
		)
	}
	return {
		calcChainState: outputIsXlsx
			? calcChainParts.length > 0
				? 'present'
				: 'absent'
			: 'not-applicable',
		calcChainParts,
		recalculationRequested,
		calcSettings,
		...formulaCache,
		warnings,
		verification: 'reopened-output',
	}
}

function postWriteFormulaCacheSummary(
	workbook: Workbook,
): Pick<
	PostWriteFormulaSummary,
	| 'formulaCells'
	| 'cachedFormulaValues'
	| 'missingCachedFormulaValues'
	| 'formulaCacheState'
	| 'cachedValueKinds'
	| 'missingCachedFormulaLocationSample'
> {
	let formulaCells = 0
	let cachedFormulaValues = 0
	let missingCachedFormulaValues = 0
	const cachedValueKindCounts = new Map<string, number>()
	const missingCachedFormulaLocationSample: string[] = []
	for (const sheet of workbook.sheets) {
		const range = sheet.cells.usedRange()
		if (!range) continue
		sheet.cells.forEachCellInRange(range, (row, col, cell) => {
			if (cell.formula === null) return
			formulaCells++
			if (cell.value.kind === 'empty') {
				missingCachedFormulaValues++
				if (missingCachedFormulaLocationSample.length < 25) {
					missingCachedFormulaLocationSample.push(`${sheet.name}!${indexToColumn(col)}${row + 1}`)
				}
				return
			}
			cachedFormulaValues++
			cachedValueKindCounts.set(
				cell.value.kind,
				(cachedValueKindCounts.get(cell.value.kind) ?? 0) + 1,
			)
		})
	}
	return {
		formulaCells,
		cachedFormulaValues,
		missingCachedFormulaValues,
		formulaCacheState: formulaCacheState(
			formulaCells,
			cachedFormulaValues,
			missingCachedFormulaValues,
		),
		cachedValueKinds: [...cachedValueKindCounts.entries()]
			.map(([kind, count]) => ({ kind, count }))
			.sort((left, right) => left.kind.localeCompare(right.kind)),
		missingCachedFormulaLocationSample,
	}
}

function formulaCacheState(
	formulaCells: number,
	cachedFormulaValues: number,
	missingCachedFormulaValues: number,
): PostWriteFormulaSummary['formulaCacheState'] {
	if (formulaCells === 0) return 'no-formulas'
	if (cachedFormulaValues === formulaCells) return 'all-cached'
	if (missingCachedFormulaValues === formulaCells) return 'all-missing'
	return 'partially-cached'
}

function postWriteWorkbookTopologySummary(workbook: Workbook): PostWriteWorkbookTopologySummary {
	const sheetStates = workbook.sheets.map((sheet) => ({
		sheetName: sheet.name,
		state: sheet.state,
	}))
	const workbookViewDetails = workbook.workbookViews.map((view, index) => ({
		index,
		...(view.activeTab !== undefined ? { activeTab: view.activeTab } : {}),
		...(view.firstSheet !== undefined ? { firstSheet: view.firstSheet } : {}),
		...(view.visibility !== undefined ? { visibility: view.visibility } : {}),
		...(view.tabRatio !== undefined ? { tabRatio: view.tabRatio } : {}),
	}))
	const hiddenSheetNames = sheetStates
		.filter((sheet) => sheet.state === 'hidden')
		.map((sheet) => sheet.sheetName)
	const veryHiddenSheetNames = sheetStates
		.filter((sheet) => sheet.state === 'veryHidden')
		.map((sheet) => sheet.sheetName)
	return {
		sheets: sheetStates.length,
		visibleSheets: sheetStates.filter((sheet) => sheet.state === 'visible').length,
		hiddenSheets: hiddenSheetNames.length,
		veryHiddenSheets: veryHiddenSheetNames.length,
		hiddenSheetNames,
		veryHiddenSheetNames,
		workbookViews: workbook.workbookViews.length,
		activeTabs: workbook.workbookViews.flatMap((view) =>
			view.activeTab === undefined ? [] : [view.activeTab],
		),
		firstSheets: workbook.workbookViews.flatMap((view) =>
			view.firstSheet === undefined ? [] : [view.firstSheet],
		),
		workbookViewDetails,
		sheetStates,
		verification: 'reopened-output',
	}
}

function postWriteAnalyticsSummary(workbook: Workbook): PostWriteAnalyticsSummary {
	const pivotTableDetails = workbook.pivotTables.map((pivot) => ({
		partPath: pivot.partPath,
		...(pivot.name ? { name: pivot.name } : {}),
		sheetName: pivot.sheetName,
		...(pivot.cacheId !== undefined ? { cacheId: pivot.cacheId } : {}),
		...(pivot.locationRef ? { locationRef: pivot.locationRef } : {}),
	}))
	const pivotCacheDetails = workbook.pivotCaches.map((cache) => {
		const linkedPivots = workbook.pivotTables.filter(
			(pivot) => cache.cacheId !== undefined && pivot.cacheId === cache.cacheId,
		)
		const outputState = pivotCacheOutputState(workbook, cache, linkedPivots)
		return {
			partPath: cache.partPath,
			...(cache.cacheId !== undefined ? { cacheId: cache.cacheId } : {}),
			...(cache.sourceType ? { sourceType: cache.sourceType } : {}),
			...(cache.sourceSheet ? { sourceSheet: cache.sourceSheet } : {}),
			...(cache.sourceRef ? { sourceRef: cache.sourceRef } : {}),
			...(cache.sourceName ? { sourceName: cache.sourceName } : {}),
			...(cache.refreshOnLoad !== undefined ? { refreshOnLoad: cache.refreshOnLoad } : {}),
			...(cache.invalid !== undefined ? { invalid: cache.invalid } : {}),
			...(cache.saveData !== undefined ? { saveData: cache.saveData } : {}),
			...(cache.recordCount !== undefined ? { recordCount: cache.recordCount } : {}),
			...(cache.recordsPartPath ? { recordsPartPath: cache.recordsPartPath } : {}),
			outputState,
			requiresExternalRefresh: outputState !== 'cached',
			linkedPivotTableNames: linkedPivotNames(linkedPivots, []),
		}
	})
	const slicerCacheDetails = workbook.slicerCaches.map((cache) => ({
		partPath: cache.partPath,
		...(cache.name ? { name: cache.name } : {}),
		...(cache.sourceName ? { sourceName: cache.sourceName } : {}),
		...(cache.pivotCacheId !== undefined ? { pivotCacheId: cache.pivotCacheId } : {}),
		pivotTableNames: cache.pivotTableNames,
		slicerPartPaths: workbook.slicers
			.filter((slicer) => slicer.cacheName === cache.name)
			.map((slicer) => slicer.partPath),
	}))
	const timelineCacheDetails = workbook.timelineCaches.map((cache) => ({
		partPath: cache.partPath,
		...(cache.name ? { name: cache.name } : {}),
		...(cache.sourceName ? { sourceName: cache.sourceName } : {}),
		...(cache.pivotCacheId !== undefined ? { pivotCacheId: cache.pivotCacheId } : {}),
		pivotTableNames: cache.pivotTableNames,
		timelinePartPaths: workbook.timelines
			.filter((timeline) => timeline.cacheName === cache.name)
			.map((timeline) => timeline.partPath),
		...(cache.state?.selection ? { selection: cache.state.selection } : {}),
	}))
	const partPaths = uniqueStrings([
		...pivotCacheDetails.flatMap((entry) => [
			entry.partPath,
			...(entry.recordsPartPath ? [entry.recordsPartPath] : []),
		]),
		...pivotTableDetails.map((entry) => entry.partPath),
		...slicerCacheDetails.flatMap((entry) => [entry.partPath, ...entry.slicerPartPaths]),
		...timelineCacheDetails.flatMap((entry) => [entry.partPath, ...entry.timelinePartPaths]),
	])
	const analyticsPartCount =
		workbook.pivotCaches.length +
		workbook.pivotTables.length +
		workbook.slicerCaches.length +
		workbook.slicers.length +
		workbook.timelineCaches.length +
		workbook.timelines.length
	return {
		pivotCaches: workbook.pivotCaches.length,
		pivotTables: workbook.pivotTables.length,
		slicerCaches: workbook.slicerCaches.length,
		slicers: workbook.slicers.length,
		timelineCaches: workbook.timelineCaches.length,
		timelines: workbook.timelines.length,
		partPaths,
		pivotCacheDetails,
		pivotTableDetails,
		slicerCacheDetails,
		timelineCacheDetails,
		requiresExternalRefresh: pivotCacheDetails.some((entry) => entry.requiresExternalRefresh),
		preservationMode: analyticsPartCount > 0 ? 'preserve-exact' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteActiveContentSummary(workbook: Workbook): PostWriteActiveContentSummary {
	const entries = workbook.activeContent.map((entry) => ({
		kind: entry.kind,
		partPath: entry.partPath,
		contentType: entry.contentType,
		anchor: entry.anchor,
		...(entry.sheetName ? { sheetName: entry.sheetName } : {}),
		...(entry.sourcePartPath ? { sourcePartPath: entry.sourcePartPath } : {}),
		...(entry.relType ? { relType: entry.relType } : {}),
		...(entry.sourceRelationshipId ? { sourceRelationshipId: entry.sourceRelationshipId } : {}),
		relationshipCount: entry.relationshipCount,
		...(entry.byteSize !== undefined ? { byteSize: entry.byteSize } : {}),
		...(entry.opaque !== undefined ? { opaque: entry.opaque } : {}),
		...(entry.executionPolicy ? { executionPolicy: entry.executionPolicy } : {}),
		...(entry.invalidationPolicy ? { invalidationPolicy: entry.invalidationPolicy } : {}),
		...(entry.resigningPolicy ? { resigningPolicy: entry.resigningPolicy } : {}),
		...(entry.shapeMacro ? { shapeMacro: entry.shapeMacro } : {}),
	}))
	return {
		total: entries.length,
		vbaProjects: entries.filter((entry) => entry.kind === 'vbaProject').length,
		activeXControls: entries.filter(
			(entry) =>
				entry.kind === 'activeX' &&
				entry.relType !==
					'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary',
		).length,
		formControls: entries.filter((entry) => entry.kind === 'formControl').length,
		macroSheets: entries.filter((entry) => entry.kind === 'macroSheet').length,
		vbaSignatures: entries.filter((entry) => entry.kind === 'vbaSignature').length,
		digitalSignatures: entries.filter((entry) => entry.kind === 'digitalSignature').length,
		customUi: entries.filter((entry) => entry.kind === 'customUi').length,
		shapeMacros: entries.filter((entry) => entry.kind === 'shapeMacro').length,
		unknownActiveContent: entries.filter((entry) => entry.kind === 'unknownActiveContent').length,
		partPaths: uniqueStrings(entries.map((entry) => entry.partPath)),
		entries,
		executionPolicy: entries.length > 0 ? 'blocked' : 'none',
		preservationMode: entries.length > 0 ? 'preserve-exact' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteSecuritySummary(workbook: Workbook): PostWriteSecuritySummary {
	const workbookProtection = workbook.workbookProtection
	const workbookLocks = workbookProtection
		? (['lockStructure', 'lockWindows', 'lockRevision'] as const).filter(
				(key) => workbookProtection[key] === true,
			)
		: []
	const workbookPasswordProtected = Boolean(
		workbookProtection?.workbookPassword ||
			workbookProtection?.workbookHashValue ||
			workbookProtection?.workbookAlgorithmName ||
			workbookProtection?.workbookSaltValue ||
			workbookProtection?.workbookSpinCount,
	)
	const workbookRevisionPasswordProtected = Boolean(
		workbookProtection?.revisionsPassword ||
			workbookProtection?.revisionsHashValue ||
			workbookProtection?.revisionsAlgorithmName ||
			workbookProtection?.revisionsSaltValue ||
			workbookProtection?.revisionsSpinCount,
	)
	const sheets = workbook.sheets
		.map((sheet) => {
			const protection = sheet.protection
			const protectedRangeLocations = sheet.protectedRanges.map(
				(range) => `${sheet.name}!${range.sqref}`,
			)
			const passwordProtected = Boolean(protection?.password)
			const strongHashProtected = Boolean(
				protection?.algorithmName ||
					protection?.hashValue ||
					protection?.saltValue ||
					protection?.spinCount,
			)
			return {
				sheetName: sheet.name,
				protected: protection !== null,
				passwordProtected,
				strongHashProtected,
				allowedActions: protection ? sheetProtectionAllowedActions(protection) : [],
				protectedRanges: sheet.protectedRanges.length,
				protectedRangeLocations,
			}
		})
		.filter((sheet) => sheet.protected || sheet.protectedRanges > 0)
	const protectedRangeDetails = workbook.sheets.flatMap((sheet) =>
		sheet.protectedRanges.map((range) => {
			const strongHashProtected = Boolean(
				range.algorithmName || range.hashValue || range.saltValue || range.spinCount,
			)
			return {
				sheetName: sheet.name,
				...(range.name ? { name: range.name } : {}),
				sqref: range.sqref,
				location: `${sheet.name}!${range.sqref}`,
				passwordProtected: Boolean(range.password),
				strongHashProtected,
				hasSecurityDescriptor: Boolean(range.securityDescriptor),
			}
		}),
	)
	const protectedRangeLocations = uniqueStrings(
		protectedRangeDetails.map((range) => range.location),
	)
	const hasPasswordOrHash =
		workbookPasswordProtected ||
		workbookRevisionPasswordProtected ||
		sheets.some((sheet) => sheet.passwordProtected || sheet.strongHashProtected) ||
		protectedRangeDetails.some((range) => range.passwordProtected || range.strongHashProtected)
	const hasSecurity =
		workbookProtection !== null || sheets.length > 0 || protectedRangeLocations.length > 0
	return {
		workbookProtected: workbookProtection !== null,
		workbookLocks,
		workbookPasswordProtected,
		workbookRevisionPasswordProtected,
		protectedSheets: sheets.filter((sheet) => sheet.protected).length,
		protectedSheetNames: sheets.filter((sheet) => sheet.protected).map((sheet) => sheet.sheetName),
		sheetPasswordProtected: sheets.filter((sheet) => sheet.passwordProtected).length,
		sheetStrongHashProtected: sheets.filter((sheet) => sheet.strongHashProtected).length,
		protectedRanges: protectedRangeLocations.length,
		protectedRangeLocations,
		protectedRangePasswordProtected: protectedRangeDetails.filter(
			(range) => range.passwordProtected,
		).length,
		protectedRangeStrongHashProtected: protectedRangeDetails.filter(
			(range) => range.strongHashProtected,
		).length,
		protectedRangeSecurityDescriptors: protectedRangeDetails.filter(
			(range) => range.hasSecurityDescriptor,
		).length,
		protectedRangeDetails,
		sheets,
		passwordHashVerification: hasPasswordOrHash ? 'reported-not-validated' : 'none',
		preservationMode: hasSecurity ? 'generated' : 'none',
		verification: 'reopened-output',
	}
}

function sheetProtectionAllowedActions(
	protection: NonNullable<Workbook['sheets'][number]['protection']>,
): string[] {
	return [
		'formatCells',
		'formatColumns',
		'formatRows',
		'insertColumns',
		'insertRows',
		'insertHyperlinks',
		'deleteColumns',
		'deleteRows',
		'selectLockedCells',
		'sort',
		'autoFilter',
		'pivotTables',
		'selectUnlockedCells',
	].filter((key) => protection[key as keyof typeof protection] === true)
}

function postWriteVisualSummary(workbook: Workbook): PostWriteVisualSummary {
	const sheets = workbook.sheets
		.map((sheet) => {
			const drawingMlObjectRefs = sheet.drawingObjectRefs.filter(
				(object) => object.source !== 'vml',
			)
			const vmlObjectRefs = sheet.drawingObjectRefs.filter((object) => object.source === 'vml')
			return {
				sheetName: sheet.name,
				hasDrawingMl: sheet.drawingRefs.hasDrawing,
				hasVml: sheet.drawingRefs.hasLegacyDrawing,
				imageCount: sheet.imageRefs.length,
				drawingObjectCount: sheet.drawingObjectRefs.length,
				drawingMlObjectCount: drawingMlObjectRefs.length,
				vmlObjectCount: vmlObjectRefs.length,
				drawingPartPaths: uniqueStrings([
					...sheet.imageRefs.map((image) => image.drawingPartPath),
					...drawingMlObjectRefs.map((object) => object.drawingPartPath),
				]),
				mediaPartPaths: uniqueStrings(sheet.imageRefs.map((image) => image.targetPath)),
				vmlPartPaths: uniqueStrings(vmlObjectRefs.map((object) => object.drawingPartPath)),
			}
		})
		.filter(
			(sheet) =>
				sheet.hasDrawingMl || sheet.hasVml || sheet.imageCount > 0 || sheet.drawingObjectCount > 0,
		)
	const charts = workbook.chartParts.map((chart) => ({
		partPath: chart.partPath,
		...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
		...(chart.chartType ? { chartType: chart.chartType } : {}),
		...(chart.title ? { title: chart.title } : {}),
		seriesCount: chart.series.length,
		series: chart.series.map((series, index) => ({
			index,
			...(series.nameRef ? { nameRef: series.nameRef } : {}),
			...(series.nameText ? { nameText: series.nameText } : {}),
			...(series.categoryRef ? { categoryRef: series.categoryRef } : {}),
			...(series.valueRef ? { valueRef: series.valueRef } : {}),
		})),
	}))
	const chartSheetPartPaths = workbook.chartSheets.flatMap(
		(chartSheet) => chartSheet.chartPartPaths,
	)
	const visualCount =
		sheets.length +
		workbook.chartParts.length +
		workbook.chartSheets.length +
		chartSheetPartPaths.length
	return {
		sheetsWithVisuals: sheets.length,
		images: sheets.reduce((total, sheet) => total + sheet.imageCount, 0),
		drawingObjects: sheets.reduce((total, sheet) => total + sheet.drawingObjectCount, 0),
		drawingMlObjects: sheets.reduce((total, sheet) => total + sheet.drawingMlObjectCount, 0),
		vmlObjects: sheets.reduce((total, sheet) => total + sheet.vmlObjectCount, 0),
		chartParts: workbook.chartParts.length,
		chartSheets: workbook.chartSheets.length,
		drawingPartPaths: uniqueStrings(sheets.flatMap((sheet) => sheet.drawingPartPaths)),
		mediaPartPaths: uniqueStrings(sheets.flatMap((sheet) => sheet.mediaPartPaths)),
		chartPartPaths: uniqueStrings([
			...charts.map((chart) => chart.partPath),
			...chartSheetPartPaths,
		]),
		vmlPartPaths: uniqueStrings(sheets.flatMap((sheet) => sheet.vmlPartPaths)),
		sheets,
		charts,
		preservationMode: visualCount > 0 ? 'preserve-exact' : 'none',
		verification: 'reopened-output',
	}
}

function postWriteOpaquePayloadSummary(workbook: Workbook): PostWriteOpaquePayloadSummary {
	const conditionalFormats = collectX14ConditionalFormatExtensionPayloads(workbook)
	const dataValidations = collectX14DataValidationExtensionPayloads(workbook)
	const generatedWithOpaquePayloads = conditionalFormats.length + dataValidations.length
	const worksheetParts = uniqueStrings([
		...conditionalFormats.flatMap((entry) => (entry.sheetPartPath ? [entry.sheetPartPath] : [])),
		...dataValidations.flatMap((entry) => (entry.sheetPartPath ? [entry.sheetPartPath] : [])),
	])
	return {
		generatedWithOpaquePayloads,
		x14ConditionalFormatExtensionPayloads: conditionalFormats.length,
		x14DataValidationExtensionPayloads: dataValidations.length,
		worksheetParts,
		preservationMode: generatedWithOpaquePayloads > 0 ? 'generated-with-opaque-payload' : 'none',
		verification: 'reopened-output',
	}
}

const SIMPLE_POST_WRITE_FEATURE_FAMILIES = new Set<string>([
	'packageContentTypes',
	'packageRelationships',
	'preservedDocumentProperties',
	'preservedStyles',
	'sharedStrings',
	'workbook',
	'worksheet',
])

function postWriteOpenOptions(
	sourceGraph: XlsxPackageGraph,
	ops: readonly Operation[],
	sourceWorkbook: Workbook,
): Pick<WorkbookLoadOptions, 'mode' | 'richMetadata' | 'formulaModeHydrateValues'> {
	if (
		ops.every((op) => op.op === 'setCells') &&
		sourceGraph.parts.every((part) => SIMPLE_POST_WRITE_FEATURE_FAMILIES.has(part.featureFamily)) &&
		!needsRichPostWriteReopen(sourceWorkbook)
	) {
		return {
			mode: 'formula',
			formulaModeHydrateValues: workbookHasFormulaCells(sourceWorkbook),
		}
	}
	return { mode: 'full', richMetadata: true }
}

function workbookHasFormulaCells(workbook: Workbook): boolean {
	return workbook.sheets.some(
		(sheet) => sheet.cells.formulaCellCount() > 0 || sheet.cells.formulaInfoCellCount() > 0,
	)
}

function needsRichPostWriteReopen(workbook: Workbook): boolean {
	if (
		workbook.workbookProtection ||
		workbook.chartParts.length > 0 ||
		workbook.chartSheets.length > 0 ||
		workbook.macroSheets.length > 0 ||
		workbook.pivotCaches.length > 0 ||
		workbook.pivotTables.length > 0 ||
		workbook.slicerCaches.length > 0 ||
		workbook.slicers.length > 0 ||
		workbook.timelineCaches.length > 0 ||
		workbook.timelines.length > 0 ||
		workbook.connectionParts.length > 0 ||
		workbook.dataModelParts.length > 0 ||
		workbook.activeContent.length > 0 ||
		workbook.externalReferenceDetails.length > 0
	) {
		return true
	}
	return workbook.sheets.some(
		(sheet) =>
			sheet.tables.length > 0 ||
			sheet.comments.size > 0 ||
			sheet.threadedComments.length > 0 ||
			sheet.hyperlinks.size > 0 ||
			sheet.ignoredErrors.length > 0 ||
			sheet.dataValidations.length > 0 ||
			sheet.conditionalFormats.length > 0 ||
			sheet.imageRefs.length > 0 ||
			sheet.drawingObjectRefs.length > 0 ||
			sheet.sparklineGroups.length > 0 ||
			sheet.x14ConditionalFormats.length > 0 ||
			sheet.x14DataValidations.length > 0 ||
			sheet.advancedFilters.length > 0 ||
			sheet.autoFilter !== null ||
			sheet.sortState !== null ||
			sheet.protection !== null ||
			sheet.protectedRanges.length > 0 ||
			sheet.preservedExtLst !== null ||
			sheet.preservedCustomSheetViews !== null ||
			sheet.preservedControlsXml !== null,
	)
}

export function auditLossPolicy(
	features: readonly FeatureReport[],
	allowLoss: readonly string[] | 'all' = [],
	packageGraph?: XlsxPackageGraph,
): LossAudit {
	const allowed = allowLoss === 'all' ? 'all' : allowLoss.map((entry) => entry.toLowerCase())
	const blockedFeatures = features.filter((feature) => {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') return false
		if (isSafePackagePreservationFeature(feature)) return false
		if (allowed === 'all') return false
		return (
			!allowed.includes(feature.feature.toLowerCase()) &&
			!allowed.includes(lossFeatureTierKey(feature)) &&
			!allowed.includes(lossFeatureApprovalId(feature))
		)
	})
	const blockedPackageParts = packageGraph
		? buildBlockedPackageParts(packageGraph, blockedFeatures)
		: []
	return {
		ok: blockedFeatures.length === 0,
		blockedFeatures,
		blockedPackageParts,
		allowedLoss: allowLoss,
		policy: 'block-preserved-and-unsupported',
	}
}

export function auditPackageGraphIntegrity(packageGraph: XlsxPackageGraph): PackageGraphAudit {
	const issues = auditXlsxPackageGraphReadIntegrity(packageGraph)
	return {
		ok: issues.length === 0,
		issues,
		policy: 'read-integrity',
	}
}

export function auditPackageGraphRoundtrip(
	sourceGraph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
	outputGraph: XlsxPackageGraph,
	outputBytes: Uint8Array,
): PackageGraphAudit {
	const issues = [
		...auditXlsxPackageGraphReadIntegrity(outputGraph),
		...auditXlsxPackageGraphSafeEditIntegrity(sourceGraph, outputGraph),
		...auditXlsxPackageGraphBytePreservation(sourceGraph, sourceBytes, outputBytes),
	]
	return {
		ok: issues.length === 0,
		issues,
		policy: 'safe-edit-roundtrip',
	}
}

function nonXlsxPackageGraphAudit(): PackageGraphAudit {
	return {
		ok: true,
		issues: [],
		policy: 'not-applicable-non-xlsx',
	}
}

export async function createRepairPlan(file: string): Promise<RepairPlanResult> {
	const inputSha256 = await fileSha256(file)
	const wb = await AscendWorkbook.open(file)
	const check = wb.check()
	const lint = wb.lint()
	const unsupportedFeatures = wb.report.features
	const actions: RepairAction[] = []
	const checkErrors = check.issues.filter((issue) => issue.severity === 'error')
	if (checkErrors.length > 0) {
		actions.push({
			code: 'run-check',
			title: 'Resolve structural check failures before writing.',
			command: `ascend check ${file} --json`,
			details: { errors: checkErrors },
		})
	}
	if (lint.warnings.length > 0) {
		actions.push({
			code: 'run-lint',
			title: 'Resolve formula lint failures before committing formula edits.',
			command: `ascend lint ${file} --json`,
			details: { warnings: lint.warnings },
		})
	}
	if (unsupportedFeatures.length > 0) {
		actions.push({
			code: 'inspect-unsupported',
			title: 'Review preserved or unsupported workbook features before editing.',
			command: `ascend inspect ${file} --json --verbose`,
			details: { features: unsupportedFeatures },
		})
	}
	if (actions.length === 0) {
		actions.push({
			code: 'plan-before-commit',
			title: 'Workbook is ready for the agent edit workflow.',
			command: `ascend plan ${file} --ops ops.json --json`,
		})
	}
	const trace = finalizeTrace({
		kind: 'repair-plan',
		file,
		inputSha256,
		phases: [
			okPhase('hash-input', 'Input workbook hash captured.'),
			checkPhase(check),
			lintPhase(lint),
			featureAuditPhase(unsupportedFeatures),
			actions.length === 0
				? okPhase('repair-actions', 'No repair actions were needed.')
				: warningPhase(
						'repair-actions',
						`${actions.length} repair action(s) suggested.`,
						actions.length,
					),
		],
		artifacts: [
			artifact('check', check, `${check.issues.length} check issue(s)`),
			artifact('lint', lint, `${lint.warnings.length} lint warning(s)`),
			artifact('features', unsupportedFeatures, `${unsupportedFeatures.length} feature(s)`),
			artifact('actions', actions, `${actions.length} action(s)`),
		],
	})
	return {
		file,
		inputSha256,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		check,
		lint,
		unsupportedFeatures,
		actions,
	}
}

export function digestPlan(inputSha256: string, ops: readonly Operation[]): string {
	return sha256Text(stableStringify({ inputSha256, ops }))
}

export function sha256Bytes(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

async function fileSha256(file: string): Promise<string> {
	return sha256Bytes(await fileBytes(file))
}

async function fileBytes(file: string): Promise<Uint8Array> {
	if (typeof Bun !== 'undefined') return Bun.file(file).bytes()
	const bytes = await readFile(file)
	return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

async function readStableAgentSource(file: string): Promise<{
	readonly sourceBytes: Uint8Array
	readonly identity: AgentSourceIdentity
}> {
	let lastSource: {
		readonly sourceBytes: Uint8Array
		readonly identity: AgentSourceIdentity
	} | null = null
	for (let attempt = 0; attempt < 2; attempt++) {
		const before = await readSourceIdentity(file)
		const sourceBytes = await fileBytes(file)
		const after = await readSourceIdentity(file)
		lastSource = { sourceBytes, identity: after }
		if (isAgentSourceIdentityEqual(before, after)) return lastSource
	}
	if (!lastSource) {
		const sourceBytes = await fileBytes(file)
		return { sourceBytes, identity: await readSourceIdentity(file) }
	}
	return lastSource
}

async function readSourceIdentity(file: string): Promise<AgentSourceIdentity> {
	const info = await stat(file)
	return { size: info.size, mtimeMs: info.mtimeMs }
}

function isAgentSourceIdentityEqual(
	left: AgentSourceIdentity,
	right: AgentSourceIdentity,
): boolean {
	return left.size === right.size && left.mtimeMs === right.mtimeMs
}

function passwordOpenOptions(options: Pick<AgentCommitOptions, 'password'>): {
	readonly password?: string
} {
	return options.password === undefined ? {} : { password: options.password }
}

async function openWorkbookFromBytes(
	file: string,
	bytes: Uint8Array,
	options: Omit<NonNullable<Parameters<typeof AscendWorkbook.open>[1]>, 'sourceExtension'> = {},
): Promise<AscendWorkbook> {
	const sourceExtension = extname(file).replace(/^\./, '').toLowerCase()
	return AscendWorkbook.open(bytes, {
		...options,
		...(sourceExtension ? { sourceExtension } : {}),
	})
}

async function openPostWriteDocument(
	identity: PostWriteOutputIdentity,
	bytes: Uint8Array,
	options: NonNullable<Parameters<typeof WorkbookDocument.open>[1]> = {},
): Promise<WorkbookDocument> {
	return WorkbookDocument.openPathSnapshot(identity.path, bytes, identity, options)
}

interface PostWriteOutputIdentity {
	readonly path: string
	readonly size: number
	readonly mtimeMs: number
	readonly ctimeMs: number
	readonly sha256: string
}

interface PostWriteOutputSnapshot {
	readonly bytes: Uint8Array
	readonly identity: PostWriteOutputIdentity
}

async function readStablePostWriteOutputSnapshot(file: string): Promise<PostWriteOutputSnapshot> {
	const path = resolve(file)
	let lastSnapshot: PostWriteOutputSnapshot | null = null
	for (let attempt = 0; attempt < 2; attempt++) {
		const before = await stat(path)
		const bytes = await fileBytes(path)
		const sha256 = sha256Bytes(bytes)
		const after = await stat(path)
		lastSnapshot = {
			bytes,
			identity: {
				path,
				size: after.size,
				mtimeMs: after.mtimeMs,
				ctimeMs: after.ctimeMs,
				sha256,
			},
		}
		if (
			before.size === after.size &&
			before.mtimeMs === after.mtimeMs &&
			before.ctimeMs === after.ctimeMs
		) {
			return lastSnapshot
		}
	}
	const observedSha256 = lastSnapshot?.identity.sha256
	throw new AscendException(
		ascendError('EXPORT_ERROR', `Written workbook was unstable before verification: ${file}`, {
			retryable: true,
			retryStrategy: 'modified',
			details: {
				output: file,
				...(observedSha256 ? { observedSha256 } : {}),
			},
			suggestedFix:
				'Retry the commit after ensuring no other process is modifying the output workbook.',
		}),
	)
}

function postWriteOutputChangedError(
	output: string,
	expectedSha256: string,
	actualSha256: string,
): AscendException {
	return new AscendException(
		ascendError(
			'EXPORT_ERROR',
			`Written workbook changed before post-write verification: ${output}`,
			{
				retryable: true,
				retryStrategy: 'modified',
				details: {
					output,
					expectedSha256,
					actualSha256,
				},
				suggestedFix:
					'Retry the commit after ensuring no other process is modifying the output workbook.',
			},
		),
	)
}

function sha256Text(text: string): string {
	return createHash('sha256').update(text).digest('hex')
}

function finalizeTrace(
	trace: Omit<AgentWorkflowTrace, 'formatVersion' | 'traceDigest'>,
): AgentWorkflowTrace {
	const base = { formatVersion: 1 as const, ...trace }
	return { ...base, traceDigest: sha256Text(stableStringify(base)) }
}

function createProgressEmitter(
	kind: AgentWorkflowTrace['kind'],
	handler: AgentWorkflowProgressHandler | undefined,
): (
	phase: string,
	status: AgentWorkflowProgressEvent['status'],
	summary: string,
	extras?: Pick<AgentWorkflowProgressEvent, 'count' | 'refs' | 'details'>,
) => Promise<void> {
	let sequence = 0
	return async (phase, status, summary, extras = {}) => {
		if (!handler) return
		sequence += 1
		try {
			await handler({
				formatVersion: 1,
				sequence,
				kind,
				phase,
				status,
				summary,
				...definedProgressExtras(extras),
			})
		} catch {
			// Progress handlers are observers; workbook writes and proof generation must stay authoritative.
		}
	}
}

async function progressFromPhase(
	phase: AgentTracePhase,
	emit: ReturnType<typeof createProgressEmitter>,
): Promise<void> {
	await emit(phase.phase, phase.status, phase.summary, {
		...(phase.count !== undefined ? { count: phase.count } : {}),
		...(phase.refs ? { refs: phase.refs } : {}),
		...(phase.details !== undefined ? { details: phase.details } : {}),
	})
}

async function timedPostWriteStep<T>(
	emit: ReturnType<typeof createProgressEmitter>,
	name: string,
	startedSummary: string,
	okSummary: string,
	fn: () => T | Promise<T>,
): Promise<{ readonly value: T; readonly ms: number }> {
	const phase = `post-write:${name}`
	await emit(phase, 'started', startedSummary)
	const start = performance.now()
	try {
		const value = await fn()
		const ms = performance.now() - start
		await emit(phase, 'ok', okSummary, { details: { durationMs: ms } })
		return { value, ms }
	} catch (error) {
		const ms = performance.now() - start
		await emit(phase, 'failed', `${startedSummary} failed.`, { details: { durationMs: ms } })
		throw error
	}
}

async function timedCommitStep<T>(
	fn: () => T | Promise<T>,
): Promise<{ readonly value: T; readonly ms: number }> {
	const start = performance.now()
	const value = await fn()
	return { value, ms: performance.now() - start }
}

function definedProgressExtras(
	extras: Pick<AgentWorkflowProgressEvent, 'count' | 'refs' | 'details'>,
): Pick<AgentWorkflowProgressEvent, 'count' | 'refs' | 'details'> {
	return {
		...(extras.count !== undefined ? { count: extras.count } : {}),
		...(extras.refs ? { refs: extras.refs } : {}),
		...(extras.details !== undefined ? { details: extras.details } : {}),
	}
}

function artifact(name: string, value: unknown, summary: string): AgentTraceArtifact {
	return {
		name,
		digest: sha256Text(stableStringify(value)),
		summary,
	}
}

function okPhase(phase: string, summary: string, count?: number): AgentTracePhase {
	return phaseResult(phase, 'ok', summary, count)
}

function warningPhase(phase: string, summary: string, count?: number): AgentTracePhase {
	return phaseResult(phase, 'warning', summary, count)
}

function phaseResult(
	phase: string,
	status: AgentTracePhase['status'],
	summary: string,
	count?: number,
	details?: unknown,
): AgentTracePhase {
	return {
		phase,
		status,
		summary,
		...(count !== undefined ? { count } : {}),
		...(details !== undefined ? { details } : {}),
	}
}

function previewPhase(preview: ReturnType<AscendWorkbook['preview']>): AgentTracePhase {
	if (preview.errors.length > 0)
		return phaseResult(
			'preview',
			'failed',
			`Preview failed with ${preview.errors.length} error(s).`,
			preview.errors.length,
			preview.errors,
		)
	if (preview.warnings.length > 0)
		return phaseResult(
			'preview',
			'warning',
			`Preview succeeded with ${preview.warnings.length} warning(s).`,
			preview.warnings.length,
			preview.warnings,
		)
	return okPhase(
		'preview',
		`${preview.changedCells.length} changed cell(s) previewed.`,
		preview.changedCells.length,
	)
}

function applyPhase(apply: ReturnType<AscendWorkbook['apply']>): AgentTracePhase {
	if (apply.errors.length > 0)
		return phaseResult(
			'apply',
			'failed',
			`Apply failed with ${apply.errors.length} error(s).`,
			apply.errors.length,
			apply.errors,
		)
	if ((apply.warnings?.length ?? 0) > 0)
		return phaseResult(
			'apply',
			'warning',
			`Apply completed with ${apply.warnings?.length ?? 0} warning(s).`,
			apply.warnings?.length ?? 0,
			apply.warnings,
		)
	return okPhase(
		'apply',
		`${apply.affectedCells.length} affected cell(s).`,
		apply.affectedCells.length,
	)
}

function recalcPhase(recalc: ReturnType<AscendWorkbook['recalc']> | null): AgentTracePhase {
	if (!recalc) return okPhase('recalc', 'Recalculation not required.')
	if (recalc.errors.length > 0)
		return phaseResult(
			'recalc',
			'failed',
			`Recalculation failed with ${recalc.errors.length} error(s).`,
			recalc.errors.length,
			recalc.errors,
		)
	return okPhase('recalc', `${recalc.changed.length} recalculated cell(s).`, recalc.changed.length)
}

function checkPhase(check: ReturnType<AscendWorkbook['check']>): AgentTracePhase {
	const errors = check.issues.filter((issue) => issue.severity === 'error')
	if (errors.length > 0)
		return phaseResult(
			'check',
			'failed',
			`Structural check found ${errors.length} error(s).`,
			errors.length,
			errors,
		)
	if (check.issues.length > 0)
		return phaseResult(
			'check',
			'warning',
			`Structural check found ${check.issues.length} issue(s).`,
			check.issues.length,
			check.issues,
		)
	return okPhase('check', 'Structural check passed.', 0)
}

function lintPhase(lint: ReturnType<AscendWorkbook['lint']>): AgentTracePhase {
	if (lint.warnings.length > 0)
		return phaseResult(
			'lint',
			'warning',
			`Formula lint found ${lint.warnings.length} warning(s).`,
			lint.warnings.length,
			lint.warnings,
		)
	return okPhase('lint', 'Formula lint passed.', 0)
}

function lossAuditPhase(lossAudit: LossAudit): AgentTracePhase {
	if (!lossAudit.ok)
		return phaseResult(
			'loss-audit',
			'blocked',
			`${lossAudit.blockedFeatures.length} preserved or unsupported feature(s) require approval.`,
			lossAudit.blockedFeatures.length,
			{
				blockedFeatures: lossAudit.blockedFeatures,
				blockedPackageParts: lossAudit.blockedPackageParts,
			},
		)
	return okPhase('loss-audit', 'No unapproved preserved or unsupported feature loss detected.', 0)
}

function packageGraphAuditPhase(audit: PackageGraphAudit): AgentTracePhase {
	if (!audit.ok) {
		return phaseResult(
			'package-graph-audit',
			'warning',
			`${audit.issues.length} package graph fidelity issue(s) require inspection.`,
			audit.issues.length,
			{ issues: audit.issues },
		)
	}
	if (audit.policy === 'not-applicable-non-xlsx') {
		return okPhase('package-graph-audit', 'Package graph audit skipped for non-XLSX output.', 0)
	}
	return okPhase('package-graph-audit', 'Package graph fidelity audit passed.', 0)
}

function approvalPhase(
	approvals: readonly ApprovalRequirement[],
	granted: readonly string[] | 'all' = [],
	blockedApprovals?: readonly ApprovalRequirement[],
): AgentTracePhase {
	if (approvals.length === 0)
		return okPhase('approval-audit', 'No approval-gated action detected.', 0)
	const unsatisfied =
		blockedApprovals ?? approvals.filter((approval) => !isApprovalSatisfied(approval, granted))
	if (unsatisfied.length > 0) {
		return phaseResult(
			'approval-audit',
			'blocked',
			`${unsatisfied.length} action(s) require explicit approval.`,
			unsatisfied.length,
			unsatisfied,
		)
	}
	return okPhase(
		'approval-audit',
		`${approvals.length} approval requirement(s) satisfied.`,
		approvals.length,
	)
}

function featureAuditPhase(features: readonly unknown[]): AgentTracePhase {
	if (features.length > 0)
		return warningPhase(
			'feature-audit',
			`${features.length} preserved or unsupported feature(s) require review.`,
			features.length,
		)
	return okPhase('feature-audit', 'No preserved or unsupported features detected.', 0)
}

function preservationPhase(
	preservation: ReturnType<AscendWorkbook['writePlanSummary']>,
): AgentTracePhase {
	return okPhase(
		'preservation-audit',
		`${preservation.totalParts} workbook package part(s) planned for write/preservation.`,
		preservation.totalParts,
	)
}

function writePolicyPhase(writePolicy: WritePolicyReport): AgentTracePhase {
	const blockers = writePolicy.diagnostics.filter((diagnostic) => diagnostic.severity === 'blocker')
	if (blockers.length > 0) {
		return phaseResult(
			'write-policy',
			'blocked',
			`${blockers.length} write policy blocker(s) require action before commit.`,
			blockers.length,
			writePolicy,
		)
	}
	const warnings = writePolicy.diagnostics.filter((diagnostic) => diagnostic.severity === 'warning')
	if (warnings.length > 0) {
		return phaseResult(
			'write-policy',
			'warning',
			`${warnings.length} write policy warning(s) require inspection.`,
			warnings.length,
			writePolicy,
		)
	}
	return phaseResult(
		'write-policy',
		'ok',
		'Write policy has no warning or blocker diagnostics.',
		writePolicy.diagnostics.length,
		writePolicy,
	)
}

function postWritePhase(
	postWrite: AgentPostWriteVerification,
	expectedPackageGraphChanges: ExpectedPostWritePackageGraphChanges = {
		removedPartPaths: new Set(),
		addedPartPaths: new Set(),
		rewrittenPartPaths: new Set(),
	},
): AgentTracePhase {
	const checkErrors = postWrite.check.issues.filter((issue) => issue.severity === 'error')
	const lintFailures = postWriteLintFailures(postWrite.lint)
	if (!postWrite.valid || checkErrors.length > 0) {
		return phaseResult(
			'post-write',
			'failed',
			`Written workbook reopened but structural verification found ${checkErrors.length} error(s).`,
			checkErrors.length,
			{ output: postWrite.output, outputSha256: postWrite.outputSha256, errors: checkErrors },
		)
	}
	if (lintFailures.length > 0) {
		return phaseResult(
			'post-write',
			'blocked',
			`Written workbook reopened but formula lint found ${lintFailures.length} blocking issue(s).`,
			lintFailures.length,
			{
				output: postWrite.output,
				outputSha256: postWrite.outputSha256,
				check: postWrite.check,
				lint: postWrite.lint,
				lintFailures,
				packageGraphAudit: postWrite.packageGraphAudit,
			},
		)
	}
	if (!postWrite.packageGraphAudit.ok) {
		const unresolvedPackageGraphIssues = postWrite.packageGraphAudit.issues.filter(
			(issue) => !isExpectedPostWritePackageGraphIssue(issue, expectedPackageGraphChanges),
		)
		if (unresolvedPackageGraphIssues.length === 0) {
			return phaseResult(
				'post-write',
				'warning',
				`Written workbook reopened with ${postWrite.packageGraphAudit.issues.length} expected package graph change(s) from approved operations.`,
				postWrite.packageGraphAudit.issues.length,
				{
					output: postWrite.output,
					outputSha256: postWrite.outputSha256,
					check: postWrite.check,
					lint: postWrite.lint,
					packageGraphAudit: postWrite.packageGraphAudit,
					expectedPackageGraphChanges: {
						removedPartPaths: [...expectedPackageGraphChanges.removedPartPaths],
						addedPartPaths: [...expectedPackageGraphChanges.addedPartPaths],
					},
				},
			)
		}
		return phaseResult(
			'post-write',
			'blocked',
			`Written workbook reopened but package graph roundtrip audit found ${unresolvedPackageGraphIssues.length} unresolved issue(s).`,
			postWrite.check.issues.length +
				postWrite.lint.warnings.length +
				unresolvedPackageGraphIssues.length,
			{
				output: postWrite.output,
				outputSha256: postWrite.outputSha256,
				check: postWrite.check,
				lint: postWrite.lint,
				packageGraphAudit: postWrite.packageGraphAudit,
				unresolvedPackageGraphIssues,
			},
		)
	}
	if (postWrite.check.issues.length > 0 || postWrite.lint.warnings.length > 0) {
		return phaseResult(
			'post-write',
			'warning',
			`Written workbook reopened with ${postWrite.check.issues.length} check issue(s), ${postWrite.lint.warnings.length} lint warning(s), and ${postWrite.packageGraphAudit.issues.length} package graph issue(s).`,
			postWrite.check.issues.length +
				postWrite.lint.warnings.length +
				postWrite.packageGraphAudit.issues.length,
			{
				output: postWrite.output,
				outputSha256: postWrite.outputSha256,
				check: postWrite.check,
				lint: postWrite.lint,
				packageGraphAudit: postWrite.packageGraphAudit,
			},
		)
	}
	return phaseResult(
		'post-write',
		'ok',
		'Written workbook reopened and passed post-write package graph verification.',
		0,
		{
			output: postWrite.output,
			outputSha256: postWrite.outputSha256,
			packageGraphAudit: postWrite.packageGraphAudit,
		},
	)
}

function modelOutputFromTrace(trace: AgentWorkflowTrace): AgentModelOutput {
	const blocked = trace.phases.some(
		(phase) => phase.status === 'blocked' || phase.status === 'failed',
	)
	const warnings = trace.phases
		.filter((phase) => phase.status === 'warning')
		.map((phase) => `${phase.phase}: ${phase.summary}`)
	const nextActions = buildNextActions(trace)
	const counts: Partial<Record<keyof AgentModelOutput['counts'], number>> = {}
	setCount(counts, 'operations', trace.operationCount)
	setCount(counts, 'changedCells', traceArtifactCount(trace, 'preview'))
	setCount(counts, 'recalcErrors', phaseCount(trace, 'recalc', 'failed'))
	setCount(counts, 'checkIssues', phaseCount(trace, 'check'))
	setCount(counts, 'lintWarnings', phaseCount(trace, 'lint', 'warning'))
	setCount(counts, 'blockedFeatures', phaseCount(trace, 'loss-audit', 'blocked'))
	setCount(counts, 'packageGraphIssues', phaseCount(trace, 'package-graph-audit', 'warning'))
	setCount(counts, 'postWritePackageGraphIssues', postWritePackageGraphIssueCount(trace))
	setCount(counts, 'postWriteLintFailures', postWriteLintFailureCount(trace))
	setCount(counts, 'preservationParts', phaseCount(trace, 'preservation-audit'))
	setCount(counts, 'writePolicyDiagnostics', phaseCount(trace, 'write-policy'))
	return {
		summary: blocked
			? `${trace.kind} needs attention before it can be treated as safe.`
			: `${trace.kind} completed its safety workflow.`,
		blocked,
		warnings,
		nextActions,
		digests: {
			inputSha256: trace.inputSha256,
			...(trace.outputSha256 ? { outputSha256: trace.outputSha256 } : {}),
			...(trace.planDigest ? { planDigest: trace.planDigest } : {}),
			traceDigest: trace.traceDigest,
		},
		counts,
	}
}

function setCount<K extends keyof AgentModelOutput['counts']>(
	counts: Partial<Record<keyof AgentModelOutput['counts'], number>>,
	key: K,
	value: number | undefined,
): void {
	if (value !== undefined) counts[key] = value
}

function buildNextActions(trace: AgentWorkflowTrace): string[] {
	if (
		trace.phases.some((phase) => phase.phase === 'approval-audit' && phase.status === 'blocked')
	) {
		return [
			'Inspect approvals and rerun commit with --approval <id> only if the action is intentional.',
		]
	}
	if (trace.phases.some((phase) => phase.phase === 'loss-audit' && phase.status === 'blocked')) {
		return [
			'Inspect lossAudit.blockedFeatures and lossAudit.blockedPackageParts.',
			'Commit only with explicit allowLoss entries if the write is intentionally lossy.',
		]
	}
	if (trace.phases.some((phase) => phase.status === 'failed')) {
		return ['Inspect failed trace phases and run repair-plan before committing.']
	}
	if ((postWritePackageGraphIssueCount(trace) ?? 0) > 0) {
		return [
			'Inspect postWrite.packageGraphAudit.issues before treating the output workbook as package-fidelity safe.',
			'Compare the output workbook hash and retain the traceDigest with the edit record.',
		]
	}
	if ((postWriteLintFailureCount(trace) ?? 0) > 0) {
		return [
			'Inspect postWrite.lint.warnings before treating the output workbook formulas as safe.',
			'Repair formula parse/error lint failures and rerun the saved-output workflow.',
		]
	}
	if (trace.phases.some((phase) => phase.phase === 'write-policy' && phase.status === 'warning')) {
		return [
			'Inspect writePolicy.diagnostics before committing around invalidated or copied-through package features.',
			...(trace.kind === 'plan' ? ['Commit with this planDigest and inputSha256 when ready.'] : []),
		]
	}
	if (
		trace.phases.some(
			(phase) => phase.phase === 'package-graph-audit' && phase.status === 'warning',
		)
	) {
		return [
			'Inspect packageGraphAudit.issues before editing around malformed or unclassified package structure.',
			...(trace.kind === 'plan' ? ['Commit with this planDigest and inputSha256 when ready.'] : []),
		]
	}
	if (trace.kind === 'plan') return ['Commit with this planDigest and inputSha256 when ready.']
	if (trace.kind === 'commit')
		return ['Verify the output workbook hash and retain the traceDigest.']
	return ['Follow the suggested repair actions in order.']
}

function postWritePackageGraphIssueCount(trace: AgentWorkflowTrace): number | undefined {
	const postWrite = trace.phases.find((phase) => phase.phase === 'post-write')
	const audit = packageGraphAuditFromDetails(postWrite?.details)
	return audit ? audit.issues.length : undefined
}

function postWriteLintFailureCount(trace: AgentWorkflowTrace): number | undefined {
	const postWrite = trace.phases.find((phase) => phase.phase === 'post-write')
	const lint = lintFromDetails(postWrite?.details)
	return lint ? postWriteLintFailures(lint).length : undefined
}

function postWriteLintFailures(
	lint: ReturnType<AscendWorkbook['lint']>,
): ReturnType<AscendWorkbook['lint']>['warnings'] {
	return lint.warnings.filter(
		(warning) => warning.severity === 'error' || warning.rule === 'parse-error',
	)
}

function expectedPackageGraphChangesForOperations(
	workbook: Workbook,
	ops: readonly Operation[],
	sourceGraph: XlsxPackageGraph,
): ExpectedPostWritePackageGraphChanges {
	const removedPartPaths = new Set<string>()
	const addedPartPaths = new Set<string>()
	const rewrittenPartPaths = new Set<string>()
	let nextTableNumber = nextGeneratedTablePartNumber(workbook, sourceGraph)
	for (const op of ops) {
		if (op.op === 'deleteSheet') {
			const sheet = workbook.getSheet(op.sheet)
			const partPath = sheet?.preservedXml?.partPath
			if (partPath) removedPartPaths.add(partPath)
		}
		if (op.op === 'copySheet') {
			const sheet = workbook.getSheet(op.sheet)
			if (sheet) {
				for (const table of sheet.tables) {
					if (table.tableType === 'queryTable' || table.queryTable) continue
					addedPartPaths.add(`xl/tables/table${nextTableNumber}.xml`)
					nextTableNumber++
				}
			}
		}
		if (op.op === 'setConnectionRefresh' && op.partPath) {
			rewrittenPartPaths.add(op.partPath)
		}
		if (op.op === 'setChartSeriesSource' && op.partPath) {
			rewrittenPartPaths.add(op.partPath)
		}
	}
	if (ops.some((op) => operationInvalidatesCalcChain(workbook, op))) {
		for (const part of sourceGraph.parts) {
			if (part.featureFamily === 'preservedCalcChain') removedPartPaths.add(part.path)
		}
	}
	return { removedPartPaths, addedPartPaths, rewrittenPartPaths }
}

function isExpectedPostWritePackageGraphIssue(
	issue: XlsxPackageGraphFidelityIssue,
	expected: ExpectedPostWritePackageGraphChanges,
): boolean {
	const targetPath = packageGraphIssueTargetPath(issue)
	if (!targetPath) return false
	if (
		expected.addedPartPaths.has(targetPath) &&
		issue.code === 'package_content_type_override' &&
		issue.featureFamily === 'preservedTable'
	) {
		return true
	}
	if (
		expected.rewrittenPartPaths.has(targetPath) &&
		issue.code === 'package_preserved_part_bytes'
	) {
		return true
	}
	if (!expected.removedPartPaths.has(targetPath)) return false
	return (
		issue.code === 'package_content_type_override' ||
		(issue.featureFamily === 'preservedCalcChain' &&
			issue.code === 'package_preserved_relationship') ||
		(issue.code === 'package_preserved_relationship' && issue.featureFamily === 'worksheet')
	)
}

function nextGeneratedTablePartNumber(workbook: Workbook, sourceGraph: XlsxPackageGraph): number {
	let next = 1
	const visit = (partPath: string | undefined) => {
		const match = /^xl\/tables\/table(\d+)\.xml$/i.exec(partPath ?? '')
		if (!match) return
		next = Math.max(next, Number(match[1]) + 1)
	}
	for (const part of sourceGraph.parts) visit(part.path)
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) visit(table.partPath)
	}
	return next
}

function operationInvalidatesCalcChain(workbook: Workbook, op: Operation): boolean {
	switch (op.op) {
		case 'addSheet':
		case 'deleteSheet':
		case 'renameSheet':
		case 'moveSheet':
		case 'setDefinedName':
		case 'deleteDefinedName':
		case 'setPivotCache':
		case 'setPivotFieldItem':
		case 'setConnectionRefresh':
		case 'setTimelineRange':
		case 'rewriteExternalLink':
		case 'setFormula':
		case 'fillFormula':
		case 'insertRows':
		case 'deleteRows':
		case 'insertCols':
		case 'deleteCols':
		case 'createTable':
		case 'appendRows':
		case 'sortRange':
		case 'copySheet':
		case 'copyRange':
		case 'moveRange':
		case 'deleteTable':
		case 'renameTable':
		case 'resizeTable':
		case 'setTableColumn':
		case 'setSlicerCacheItem':
			return true
		case 'clearRange':
			return op.what === 'formulas' || op.what === 'all'
		case 'setCells': {
			const sheet = workbook.getSheet(op.sheet)
			if (!sheet) return false
			return op.updates.some((update) => {
				const ref = parseA1Safe(update.ref)
				if (!ref) return false
				return (
					sheet.cells.readFormula(ref.row, ref.col) !== undefined ||
					sheet.cells.readFormulaInfo(ref.row, ref.col) !== undefined
				)
			})
		}
		case 'setRichText':
		case 'setNumberFormat':
		case 'setStyle':
		case 'setDocumentProperties':
		case 'setWorkbookProperties':
		case 'setWorkbookView':
		case 'setCalcSettings':
		case 'setTheme':
		case 'setWorkbookProtection':
		case 'setRowHeight':
		case 'setColWidth':
		case 'freezePane':
		case 'mergeCells':
		case 'unmergeCells':
		case 'deleteComment':
		case 'deleteHyperlink':
		case 'deleteDataValidation':
		case 'setAutoFilter':
		case 'clearAutoFilter':
		case 'setSheetProtection':
		case 'setTabColor':
		case 'hideSheet':
		case 'hideRows':
		case 'hideCols':
		case 'deleteConditionalFormat':
		case 'setPageSetup':
		case 'setPrintArea':
		case 'setTableStyle':
		case 'setConditionalFormat':
		case 'setDataValidation':
		case 'setHyperlink':
		case 'setComment':
		case 'setThreadedComment':
		case 'replaceImage':
		case 'insertImage':
		case 'deleteImage':
		case 'setDrawingText':
		case 'setChartSeriesSource':
		case 'setSparklineGroup':
		case 'setAdvancedFilter':
			return false
	}
	return false
}

function packageGraphIssueTargetPath(issue: XlsxPackageGraphFidelityIssue): string | undefined {
	if (issue.partPath) return issue.partPath
	if (!issue.expected || typeof issue.expected !== 'object') return undefined
	const expected = issue.expected as {
		readonly partPath?: unknown
		readonly resolvedTarget?: unknown
	}
	if (typeof expected.partPath === 'string') return expected.partPath
	if (typeof expected.resolvedTarget === 'string') return expected.resolvedTarget
	return undefined
}

function packageGraphAuditFromDetails(details: unknown): PackageGraphAudit | undefined {
	if (!details || typeof details !== 'object') return undefined
	const audit = (details as { packageGraphAudit?: unknown }).packageGraphAudit
	if (!audit || typeof audit !== 'object') return undefined
	const issues = (audit as { issues?: unknown }).issues
	const policy = (audit as { policy?: unknown }).policy
	const ok = (audit as { ok?: unknown }).ok
	if (!Array.isArray(issues) || typeof policy !== 'string' || typeof ok !== 'boolean') {
		return undefined
	}
	return audit as PackageGraphAudit
}

function lintFromDetails(details: unknown): ReturnType<AscendWorkbook['lint']> | undefined {
	if (!details || typeof details !== 'object') return undefined
	const lint = (details as { lint?: unknown }).lint
	if (!lint || typeof lint !== 'object') return undefined
	const clean = (lint as { clean?: unknown }).clean
	const warnings = (lint as { warnings?: unknown }).warnings
	if (typeof clean !== 'boolean' || !Array.isArray(warnings)) return undefined
	return lint as ReturnType<AscendWorkbook['lint']>
}

function snapshotWritePolicyWorkbook(workbook: Workbook): Workbook {
	const snapshot = workbook.clone()
	for (const sheet of snapshot.sheets) sheet.ensureWritable()
	return snapshot
}

function canUseLiveWritePolicyWorkbook(
	workbook: Workbook,
	operations: readonly Operation[],
): boolean {
	return canReusePreparedCommitCheck(workbook, operations)
}

function canReusePreparedCommitCheck(
	workbook: Workbook,
	operations: readonly Operation[],
): boolean {
	return (
		operations.every((operation) => operation.op === 'setCells') &&
		workbook.sheets.every(
			(sheet) => sheet.cells.formulaCellCount() === 0 && sheet.cells.formulaInfoCellCount() === 0,
		)
	)
}

function buildWritePolicyReport(
	features: readonly FeatureReport[],
	packageGraph: XlsxPackageGraph,
	preservation: ReturnType<AscendWorkbook['writePlanSummary']>,
	packageGraphAudit: PackageGraphAudit,
	workbook: Workbook,
	operations: readonly Operation[] = [],
	externalReferences: readonly ExternalReferenceInfo[] = [],
	checkIssues: readonly CheckIssue[] = [],
): WritePolicyReport {
	const diagnostics: WritePolicyDiagnostic[] = []
	const copiedThroughParts = preservation.parts.filter((part) => part.origin !== 'generated')
	const generatedParts = preservation.parts.filter((part) => part.origin === 'generated')
	const skipped = preservation.skippedCapsules
	const packagePartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const packageGraphPartByPath = packagePartByPath
	const signatureParts = packageGraph.parts.filter(
		(part) => part.featureFamily === 'preservedSignature',
	)
	const skippedSignatureParts = signatureParts.filter((part) => skipped.includes(part.path))
	if (skippedSignatureParts.length > 0) {
		diagnostics.push({
			code: 'signature-invalidation',
			severity: 'warning',
			message: `${skippedSignatureParts.length} digital signature package part(s) will be omitted because generated edits invalidate signed package content.`,
			suggestedAction:
				'Commit only with explicit approval and re-sign the workbook outside Ascend if a trusted signature is required.',
			partPaths: skippedSignatureParts.map((part) => part.path),
			packageParts: packagePartDetailsFromParts(skippedSignatureParts),
			featureFamily: 'preservedSignature',
			preservationPolicy: 'invalidate-on-edit',
			preservationMode: 'invalidated-on-edit',
		})
	}
	const calcChainParts = packageGraph.parts.filter(
		(part) => part.featureFamily === 'preservedCalcChain',
	)
	const skippedCalcChainParts = calcChainParts.filter((part) => skipped.includes(part.path))
	const calcChainPolicy =
		calcChainParts.length === 0
			? 'not-present'
			: skippedCalcChainParts.length > 0
				? 'discarded-for-formula-topology'
				: 'preserved'
	const checkErrors = checkIssues.filter(
		(issue) =>
			issue.severity === 'error' &&
			!isExpectedDiscardedCalcChainCheckIssue(issue, skippedCalcChainParts),
	)
	if (checkErrors.length > 0) {
		diagnostics.push({
			code: 'pre-write-check-error',
			severity: 'blocker',
			message: `${checkErrors.length} structural check error(s) would be written to the output workbook.`,
			suggestedAction:
				'Run ascend check or repair-plan, fix the reported workbook integrity errors, and re-run plan before committing.',
			details: { checkIssues: checkErrors },
		})
	}
	if (generatedParts.length > 0) {
		diagnostics.push({
			code: 'generated-replacement-parts',
			severity: 'info',
			message: `${generatedParts.length} package part(s) will be generated or regenerated by Ascend.`,
			suggestedAction:
				'Inspect preservation.parts where origin is generated when auditing replacement package XML.',
			partPaths: generatedParts.map((part) => part.path),
			preservationMode: 'generated',
		})
	}
	if (copiedThroughParts.length > 0) {
		const packageParts = packagePartDetails(
			packagePartByPath,
			copiedThroughParts.map((part) => part.path),
		)
		diagnostics.push({
			code: 'copied-through-package-parts',
			severity: 'info',
			message: `${copiedThroughParts.length} package part(s) will be copied through from preserved source or capsule bytes.`,
			suggestedAction:
				'Expect byte-preservation audit coverage for copied-through package parts after write.',
			partPaths: copiedThroughParts.map((part) => part.path),
			...(packageParts.length > 0 ? { packageParts } : {}),
			...(packageParts[0]?.preservationMode
				? { preservationMode: packageParts[0].preservationMode }
				: {}),
		})
	}
	if (skippedCalcChainParts.length > 0) {
		diagnostics.push({
			code: 'calc-chain-discarded',
			severity: 'warning',
			message:
				'The imported calculation chain will be discarded because the planned edit changes formula topology or recalculation freshness.',
			suggestedAction:
				'Allow Excel or another spreadsheet application to rebuild calcChain on open; do not treat the old order as authoritative.',
			partPaths: skippedCalcChainParts.map((part) => part.path),
			packageParts: packagePartDetailsFromParts(skippedCalcChainParts),
			featureFamily: 'preservedCalcChain',
			preservationPolicy: 'discard-on-recalc',
			preservationMode: 'discarded-for-recalc',
		})
	} else if (calcChainParts.length > 0) {
		diagnostics.push({
			code: 'calc-chain-preserved',
			severity: 'info',
			message:
				'The imported calculation chain is planned for preservation because this edit does not alter formula topology.',
			suggestedAction:
				'Post-write package graph audit should confirm the calcChain part and workbook relationship survive.',
			partPaths: calcChainParts.map((part) => part.path),
			packageParts: packagePartDetailsFromParts(calcChainParts),
			featureFamily: 'preservedCalcChain',
			preservationPolicy: 'discard-on-recalc',
			preservationMode: 'preserve-exact',
		})
	}
	const skippedNonSignatureCalc = skipped.filter(
		(path) =>
			!signatureParts.some((part) => part.path === path) &&
			!calcChainParts.some((part) => part.path === path) &&
			!preservation.parts.some((part) => part.path === path),
	)
	if (skippedNonSignatureCalc.length > 0) {
		const packageParts = packagePartDetails(packagePartByPath, skippedNonSignatureCalc)
		diagnostics.push({
			code: 'skipped-preservation-capsules',
			severity: 'warning',
			message: `${skippedNonSignatureCalc.length} preservation capsule(s) will not be copied into the written package.`,
			suggestedAction:
				'Inspect skippedCapsules and confirm each skipped part is intentionally regenerated or no longer package-reachable.',
			partPaths: skippedNonSignatureCalc,
			...(packageParts.length > 0 ? { packageParts } : {}),
			...(packageParts[0]?.preservationMode
				? { preservationMode: packageParts[0].preservationMode }
				: {}),
		})
	}
	for (const part of packageGraph.parts.filter((part) =>
		isActiveContentFeature(part.featureFamily),
	)) {
		diagnostics.push({
			code: 'active-content-preserved',
			severity: 'warning',
			message: `${part.featureFamily} part ${part.path} is active or security-sensitive content planned for preservation, not execution.`,
			suggestedAction:
				'Require explicit approval before writing and never imply macro, ActiveX, callback, or protected-payload execution support.',
			partPaths: [part.path],
			packageParts: packagePartDetailsFromParts([part]),
			featureFamily: part.featureFamily,
			ownerScope: part.ownerScope,
			preservationPolicy: part.preservationPolicy,
			preservationMode: preservationModeForPackagePolicy(part.preservationPolicy),
		})
	}
	const visualFamilies = new Set<string>()
	const visualPartPaths: string[] = []
	const analyticsPartPaths: string[] = []
	for (const part of packageGraph.parts) {
		if (isVisualSidecar(part.featureFamily)) {
			visualFamilies.add(part.featureFamily)
			visualPartPaths.push(part.path)
		}
		if (isAnalyticalSidecar(part.featureFamily)) analyticsPartPaths.push(part.path)
	}
	const visualWriteRisk = buildVisualWriteRisk(workbook, operations, packageGraph.parts)
	const chartIntegrityIssues = collectChartIntegrityIssues(checkIssues)
	const visualHasOperationRisk =
		visualWriteRisk.relatedOperations.length > 0 ||
		visualWriteRisk.chartSourceRefDrift.length > 0 ||
		visualWriteRisk.drawingmlVmlDrift.length > 0
	const visualPackageGraphIssues = packageGraphAudit.issues.filter((issue) =>
		isVisualPackageGraphIssue(issue),
	)
	const visualIssuePartPaths = uniqueStrings([
		...visualPackageGraphIssues.flatMap((issue) =>
			packageGraphIssuePackagePaths(issue).filter(isVisualPackagePartPath),
		),
		...chartIntegrityIssues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isVisualPackagePartPath),
		),
	])
	const visualDiagnosticPartPaths = uniqueStrings([...visualPartPaths, ...visualIssuePartPaths])
	for (const issue of visualPackageGraphIssues) {
		if (issue.featureFamily && isVisualSidecar(issue.featureFamily)) {
			visualFamilies.add(issue.featureFamily)
		}
	}
	if (
		visualDiagnosticPartPaths.length > 0 ||
		visualPackageGraphIssues.length > 0 ||
		chartIntegrityIssues.length > 0
	) {
		const copiedThroughVisualPartPaths = copiedThroughParts
			.map((part) => part.path)
			.filter((path) => visualPartPaths.includes(path))
		const generatedOrReplacementVisualParts = generatedParts
			.filter(
				(part) => visualDiagnosticPartPaths.includes(part.path) || isDrawingMlVisualPath(part.path),
			)
			.map((part) => ({ partPath: part.path, origin: part.origin }))
		const visualWarningIssueCount = visualPackageGraphIssues.filter(
			(issue) => issue.severity === 'error' || issue.severity === 'warning',
		).length
		const visualHasPackageRisk = visualWarningIssueCount > 0
		const visualHasIntegrityRisk = chartIntegrityIssues.length > 0
		const visualPartCount = uniqueStrings(
			visualDiagnosticPartPaths.length > 0
				? visualDiagnosticPartPaths
				: visualPackageGraphIssues.flatMap((issue) => packageGraphIssuePackagePaths(issue)),
		)
		diagnostics.push({
			code: 'visual-sidecar-preservation-risk',
			severity:
				visualHasOperationRisk || visualHasPackageRisk || visualHasIntegrityRisk
					? 'warning'
					: 'info',
			message:
				visualHasOperationRisk || visualHasPackageRisk || visualHasIntegrityRisk
					? `${visualPartCount.length} chart, drawing, image, or VML sidecar package path(s) have operation-scoped write risk, package graph issues, or verify issues.`
					: `${visualPartCount.length} chart, drawing, image, or VML sidecar part(s) will be copied through; no visual or chart-source edit is planned.`,
			suggestedAction:
				visualHasOperationRisk || visualHasPackageRisk || visualHasIntegrityRisk
					? 'Inspect visualInventory, verify chart source refs, and check postWrite.packageGraphAudit before treating visuals as fidelity-safe.'
					: 'No action is required for unrelated cell edits; retain visualInventory/packageGraphAudit with the edit record when visual fidelity is audited.',
			...(visualDiagnosticPartPaths.length > 0 ? { partPaths: visualDiagnosticPartPaths } : {}),
			...(visualDiagnosticPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, visualDiagnosticPartPaths) }
				: {}),
			featureFamily: [...visualFamilies].sort().join(','),
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				operationScoped: visualHasOperationRisk,
				packageGraphAudit: {
					policy: packageGraphAudit.policy,
					ok: packageGraphAudit.ok,
					issueCount: packageGraphAudit.issues.length,
					visualIssueCount: visualPackageGraphIssues.length,
					issues: visualPackageGraphIssues,
				},
				verifyIssues: chartIntegrityIssues,
				copiedThroughVisualParts: packagePartDetails(
					packagePartByPath,
					copiedThroughVisualPartPaths,
				),
				generatedOrReplacementVisualParts,
				chartSourceRefs: collectChartSourceRefSummary(workbook),
				drawingModel: collectDrawingModelSummary(workbook, packageGraph.parts),
				recommendedInspection:
					'Inspect visualInventory before visual edits; after write, inspect postWrite.check, postWrite.packageGraphAudit, chart source refs, DrawingML drawing relationships/media bytes, and any separate VML legacy drawing/comment layout.',
				relatedOperations: visualWriteRisk.relatedOperations,
				chartSourceRefDrift: visualWriteRisk.chartSourceRefDrift,
				drawingmlVmlDrift: visualWriteRisk.drawingmlVmlDrift,
			},
		})
	}
	const analyticsPackageGraphIssues = packageGraphAudit.issues.filter((issue) =>
		isAnalyticalPackageGraphIssue(issue),
	)
	if (
		analyticsPartPaths.length > 0 ||
		analyticsPackageGraphIssues.length > 0 ||
		visualWriteRisk.analyticsPivotRefreshRisk.length > 0
	) {
		const copiedThroughAnalyticsPartPaths = copiedThroughParts
			.map((part) => part.path)
			.filter((path) => analyticsPartPaths.includes(path))
		const generatedOrReplacementAnalyticsParts = generatedParts
			.filter(
				(part) => analyticsPartPaths.includes(part.path) || isAnalyticalPackagePartPath(part.path),
			)
			.map((part) => ({ partPath: part.path, origin: part.origin }))
		const operationScoped = visualWriteRisk.analyticsPivotRefreshRisk.length > 0
		const warningIssueCount = analyticsPackageGraphIssues.filter(
			(issue) => issue.severity === 'error' || issue.severity === 'warning',
		).length
		const analyticsPartCount = uniqueStrings([
			...analyticsPartPaths,
			...analyticsPackageGraphIssues.flatMap((issue) => packageGraphIssuePackagePaths(issue)),
		]).length
		diagnostics.push({
			code: 'analytics-preservation-risk',
			severity: operationScoped || warningIssueCount > 0 ? 'warning' : 'info',
			message:
				operationScoped || warningIssueCount > 0
					? `${analyticsPartCount} pivot, slicer, or timeline package part(s) have operation-scoped write risk or package graph issues.`
					: `${analyticsPartCount} pivot, slicer, or timeline package part(s) will be preserved; no analytical cache/filter edit is planned.`,
			suggestedAction:
				'Inspect pivotRefreshPlans, pivotCaches, pivotTables, slicerCaches, timelineCaches, packageGraphAudit.issues, and postWrite.packageGraphAudit before trusting analytical output after write.',
			...(analyticsPartPaths.length > 0 ? { partPaths: analyticsPartPaths } : {}),
			...(analyticsPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, analyticsPartPaths) }
				: {}),
			featureFamily: 'preservedPivot,preservedSlicer,preservedTimeline',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				operationScoped,
				copiedThroughAnalyticsParts: packagePartDetails(
					packagePartByPath,
					copiedThroughAnalyticsPartPaths,
				),
				generatedOrReplacementAnalyticsParts,
				packageGraphAudit: {
					policy: packageGraphAudit.policy,
					ok: packageGraphAudit.ok,
					issueCount: packageGraphAudit.issues.length,
					analyticsIssueCount: analyticsPackageGraphIssues.length,
					issues: analyticsPackageGraphIssues,
				},
				pivotCacheRisks: collectPivotCacheRisks(workbook),
				slicerTimelineCacheDependencies: collectSlicerTimelineCacheDependencies(workbook),
				unsupportedHeadlessRefresh:
					'Ascend can preserve and edit pivot, slicer, and timeline metadata, but cannot refresh pivot caches or recalculate PivotTable output cells headlessly.',
				recommendedInspection: [
					'inspect --detail pivots',
					'inspect --detail slicers',
					'inspect --detail timelines',
					'pivotRefreshPlans',
					'packageGraphAudit.issues',
					'postWrite.packageGraphAudit',
				],
				analyticsPivotRefreshRisk: visualWriteRisk.analyticsPivotRefreshRisk,
			},
		})
	}
	if (visualWriteRisk.relatedOperations.length > 0) {
		const relatedPartPaths = uniqueStrings(
			visualWriteRisk.relatedOperations.flatMap((operation) => operation.partPaths),
		)
		diagnostics.push({
			code: 'visual-edit-preservation-risk',
			severity: visualWriteRisk.relatedOperations.some((operation) => operation.matchCount !== 1)
				? 'warning'
				: 'info',
			message: `${visualWriteRisk.relatedOperations.length} planned drawing, image, or chart operation(s) should be checked against visualInventory selectors before commit.`,
			suggestedAction:
				'Use targetPath/relId/imageIndex for image replacement, drawingPartPath/id/drawingObjectIndex for drawing text, and partPath/chartIndex/seriesIndex for chart source edits; verify postWrite.packageGraphAudit after write.',
			...(relatedPartPaths.length > 0 ? { partPaths: relatedPartPaths } : {}),
			...(relatedPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, relatedPartPaths) }
				: {}),
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				relatedOperations: visualWriteRisk.relatedOperations,
			},
		})
	}
	if (visualWriteRisk.drawingmlVmlDrift.length > 0) {
		const driftPartPaths = uniqueStrings(
			visualWriteRisk.drawingmlVmlDrift.flatMap((entry) => entry.partPaths),
		)
		diagnostics.push({
			code: 'drawingml-vml-drift-risk',
			severity: 'warning',
			message: `${visualWriteRisk.drawingmlVmlDrift.length} visual operation(s) target sheet(s) that contain both DrawingML and VML drawing state.`,
			suggestedAction:
				'Select the exact drawing object or image from visualInventory and verify DrawingML plus VML/comment layout metadata after write; image and DrawingML text operations do not imply legacy VML shape updates.',
			...(driftPartPaths.length > 0 ? { partPaths: driftPartPaths } : {}),
			...(driftPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, driftPartPaths) }
				: {}),
			featureFamily: 'preservedDrawing,preservedVml',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				drift: visualWriteRisk.drawingmlVmlDrift,
			},
		})
	}
	if (visualWriteRisk.chartSourceRefDrift.length > 0 || chartIntegrityIssues.length > 0) {
		const driftPartPaths = uniqueStrings([
			...visualWriteRisk.chartSourceRefDrift.flatMap((entry) => entry.partPaths),
			...chartIntegrityIssues.flatMap((issue) =>
				checkIssuePackagePartPaths(issue).filter(isVisualPackagePartPath),
			),
		])
		diagnostics.push({
			code: 'chart-source-ref-drift-risk',
			severity: 'warning',
			message: chartSourceRefRiskMessage(
				visualWriteRisk.chartSourceRefDrift.length,
				chartIntegrityIssues.length,
			),
			suggestedAction:
				'Use setChartSeriesSource with explicit partPath, chartIndex, and seriesIndex when a structural edit should move chart name/category/value refs; repair verify chart source issues before write and verify chart series refs after write.',
			...(driftPartPaths.length > 0 ? { partPaths: driftPartPaths } : {}),
			...(driftPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, driftPartPaths) }
				: {}),
			featureFamily: 'preservedChart',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				chartSourceRefDrift: visualWriteRisk.chartSourceRefDrift,
				verifyIssues: chartIntegrityIssues,
				chartSourceRefs: collectChartSourceRefSummary(workbook),
			},
		})
	}
	if (visualWriteRisk.analyticsPivotRefreshRisk.length > 0) {
		const refreshPartPaths = uniqueStrings(
			visualWriteRisk.analyticsPivotRefreshRisk.flatMap((entry) => entry.partPaths),
		)
		diagnostics.push({
			code: 'analytics-pivot-refresh-risk',
			severity: 'warning',
			message: `${visualWriteRisk.analyticsPivotRefreshRisk.length} planned pivot, slicer, or timeline operation(s) update analytical filter/cache state; linked pivot output can remain stale until refreshed.`,
			suggestedAction:
				'Ascend cannot refresh pivots headlessly; after write, open the workbook in Excel or another pivot-aware engine, refresh affected pivots/caches, and recalculate formulas before trusting analytical output cells.',
			...(refreshPartPaths.length > 0 ? { partPaths: refreshPartPaths } : {}),
			...(refreshPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, refreshPartPaths) }
				: {}),
			featureFamily: 'preservedPivot,preservedSlicer,preservedTimeline',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				operationScoped: true,
				analyticsPivotRefreshRisk: visualWriteRisk.analyticsPivotRefreshRisk,
				unsupportedHeadlessRefresh:
					'Pivot cache refresh and PivotTable output recalculation require Excel or another pivot-aware engine.',
				recommendedInspection: [
					'inspect --detail pivots',
					'inspect --detail slicers',
					'inspect --detail timelines',
					'pivotRefreshPlans',
					'postWrite.packageGraphAudit',
				],
			},
		})
	}
	const tableIntegrityIssues = collectTableIntegrityIssues(checkIssues)
	const tableLocations = collectTableLocations(workbook)
	const tablePackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isTablePackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const tablePartPaths = collectTablePartPaths(
		packageGraph.parts,
		tableLocations,
		tableIntegrityIssues,
		tablePackageIssuePartPaths,
	)
	const tablePackageGraphIssues = packageGraphIssuesForParts(packageGraphAudit, tablePartPaths, [
		'preservedTable',
		'preservedQueryTable',
	])
	if (
		tableLocations.length > 0 ||
		tableIntegrityIssues.length > 0 ||
		tablePackageGraphIssues.length > 0
	) {
		const relatedOperations = collectTableRelatedOperations(workbook, operations)
		diagnostics.push({
			code: 'table-preservation-risk',
			severity:
				tableIntegrityIssues.length > 0 ||
				tablePackageGraphIssues.some((issue) => issue.severity === 'error')
					? 'warning'
					: 'info',
			message:
				tablePackageGraphIssues.length > 0 || tableIntegrityIssues.length > 0
					? `${tablePackageGraphIssues.length + tableIntegrityIssues.length} table or queryTable integrity issue(s) require inspection before write.`
					: `${tableLocations.length} table location(s) depend on worksheet table relationships and table/queryTable sidecar preservation.`,
			suggestedAction:
				'Use inspectSheet(sheet).tables for table part paths, queryTable sidecars, filters, sort state, style metadata, and topology preconditions; verify postWrite.check plus postWrite.packageGraphAudit after table or structural edits.',
			locations: tableLocations.map((entry) => entry.location),
			...(tablePartPaths.length > 0 ? { partPaths: tablePartPaths } : {}),
			...(tablePartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, tablePartPaths) }
				: {}),
			featureFamily: 'preservedTable,preservedQueryTable',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				tables: tableLocations,
				relatedOperations,
				packageGraphIssues: tablePackageGraphIssues,
				verifyIssues: tableIntegrityIssues,
				preconditions: [
					'workbook-unique table names',
					'workbook-unique table ids',
					'non-overlapping same-sheet table ranges',
					'table/queryTable relationship binding preserved',
				],
			},
		})
	}
	const commentIntegrityIssues = collectCommentIntegrityIssues(checkIssues)
	const legacyCommentLocations = collectLegacyCommentLocations(workbook)
	const legacyCommentPackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isLegacyCommentPackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const legacyCommentPartPaths = collectLegacyCommentPartPaths(
		packageGraph.parts,
		legacyCommentLocations,
		commentIntegrityIssues.legacy,
		legacyCommentPackageIssuePartPaths,
	)
	const legacyCommentPackageGraphIssues = packageGraphIssuesForParts(
		packageGraphAudit,
		legacyCommentPartPaths,
		['preservedComments', 'preservedVml'],
	)
	if (
		legacyCommentLocations.length > 0 ||
		commentIntegrityIssues.legacy.length > 0 ||
		legacyCommentPackageGraphIssues.length > 0
	) {
		const relatedOperations = collectLegacyCommentRelatedOperations(
			legacyCommentLocations,
			operations,
		)
		diagnostics.push({
			code: 'legacy-comment-preservation-risk',
			severity: 'warning',
			message:
				legacyCommentPackageGraphIssues.length > 0 || commentIntegrityIssues.legacy.length > 0
					? `${legacyCommentPackageGraphIssues.length + commentIntegrityIssues.legacy.length} legacy comment package or verify issue(s) require inspection before write.`
					: legacyCommentLocations.length > 0
						? `${legacyCommentLocations.length} legacy comment location(s) depend on comments XML${legacyCommentPartPaths.some((path) => path.includes('/drawings/')) ? ' and VML drawing' : ''} package preservation.`
						: `${commentIntegrityIssues.legacy.length} legacy comment verify issue(s) require inspection before write.`,
			suggestedAction:
				'For text-only edits, prefer setComment on the existing sheet/ref. Inspect inspectSheet(sheet).comments legacyDrawing/VML metadata before layout changes, then verify postWrite.check and postWrite.packageGraphAudit.',
			locations: legacyCommentLocations.map((entry) => entry.location),
			...(legacyCommentPartPaths.length > 0 ? { partPaths: legacyCommentPartPaths } : {}),
			...(legacyCommentPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, legacyCommentPartPaths) }
				: {}),
			featureFamily: 'preservedComments,preservedVml',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				comments: legacyCommentLocations,
				relatedOperations,
				packageGraphIssues: legacyCommentPackageGraphIssues,
				verifyIssues: commentIntegrityIssues.legacy,
				safeTextEdit: 'setComment',
				structuralEditVerification:
					'For row or column inserts/deletes near comments, inspectSheet(sheet).comments legacyDrawing/VML metadata before and after the write.',
			},
		})
	}
	const threadedCommentLocations = collectThreadedCommentLocations(workbook)
	const threadedCommentPackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isThreadedCommentPackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const threadedCommentPartPaths = collectThreadedCommentPartPaths(
		packageGraph.parts,
		threadedCommentLocations,
		commentIntegrityIssues.threaded,
		threadedCommentPackageIssuePartPaths,
	)
	const threadedCommentPackageGraphIssues = packageGraphIssuesForParts(
		packageGraphAudit,
		threadedCommentPartPaths,
		['preservedThreadedComments'],
	)
	if (
		threadedCommentLocations.length > 0 ||
		commentIntegrityIssues.threaded.length > 0 ||
		threadedCommentPackageGraphIssues.length > 0
	) {
		const relatedOperations = collectThreadedCommentRelatedOperations(
			threadedCommentLocations,
			operations,
		)
		diagnostics.push({
			code: 'threaded-comment-preservation-risk',
			severity: 'warning',
			message:
				threadedCommentPackageGraphIssues.length > 0 || commentIntegrityIssues.threaded.length > 0
					? `${threadedCommentPackageGraphIssues.length + commentIntegrityIssues.threaded.length} threaded comment package or verify issue(s) require inspection before write.`
					: threadedCommentLocations.length > 0
						? `${threadedCommentLocations.length} threaded comment location(s) depend on threaded comment ids, parent ids, person ids, and persons sidecar preservation.`
						: `${commentIntegrityIssues.threaded.length} threaded comment verify issue(s) require inspection before write.`,
			suggestedAction:
				'For text-only edits, use setThreadedComment with partPath and threadedCommentId from inspectSheet(sheet).threadedComments. Preserve parentId/personId/persons metadata and verify postWrite.check plus postWrite.packageGraphAudit.',
			locations: threadedCommentLocations.map((entry) => entry.location),
			...(threadedCommentPartPaths.length > 0 ? { partPaths: threadedCommentPartPaths } : {}),
			...(threadedCommentPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, threadedCommentPartPaths) }
				: {}),
			featureFamily: 'preservedThreadedComments',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				threadedComments: threadedCommentLocations,
				relatedOperations,
				packageGraphIssues: threadedCommentPackageGraphIssues,
				verifyIssues: commentIntegrityIssues.threaded,
				safeTextEdit: 'setThreadedComment',
				structuralEditVerification:
					'For row or column inserts/deletes near threaded comments, inspectSheet(sheet).threadedComments ids/person metadata before and after the write.',
			},
		})
	}
	const x14ConditionalFormatExtensionPayloads =
		collectX14ConditionalFormatExtensionPayloads(workbook)
	if (x14ConditionalFormatExtensionPayloads.length > 0) {
		const relatedOperations = collectX14ConditionalFormatSemanticEdits(
			workbook,
			operations,
			x14ConditionalFormatExtensionPayloads,
		)
		const partPaths = uniqueStrings(
			x14ConditionalFormatExtensionPayloads.flatMap((entry) =>
				entry.sheetPartPath ? [entry.sheetPartPath] : [],
			),
		)
		diagnostics.push({
			code: 'conditional-format-extension-preservation',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message:
				relatedOperations.length > 0
					? x14ConditionalFormatRiskMessage(
							x14ConditionalFormatExtensionPayloads.length,
							relatedOperations,
						)
					: `${x14ConditionalFormatExtensionPayloads.length} x14 conditional-format extension payload(s) will be preserved as opaque worksheet XML; no conditional-format semantic edit is planned.`,
			suggestedAction: x14ConditionalFormatSuggestedAction(relatedOperations.length > 0),
			...(partPaths.length > 0 ? { partPaths } : {}),
			featureFamily: 'x14ConditionalFormatting',
			preservationPolicy: 'generated',
			preservationMode: 'generated-with-opaque-payload',
			details: {
				provenance: 'worksheet-extLst',
				preservationMode: 'opaque-payload-preserved-in-generated-worksheet-xml',
				semanticEditRisk: relatedOperations.length > 0,
				relatedOperations,
				x14ConditionalFormats: x14ConditionalFormatExtensionPayloads,
			},
		})
	}
	const x14DataValidationExtensionPayloads = collectX14DataValidationExtensionPayloads(workbook)
	if (x14DataValidationExtensionPayloads.length > 0) {
		const relatedOperations = collectX14DataValidationSemanticEdits(
			workbook,
			operations,
			x14DataValidationExtensionPayloads,
		)
		const partPaths = uniqueStrings(
			x14DataValidationExtensionPayloads.flatMap((entry) =>
				entry.sheetPartPath ? [entry.sheetPartPath] : [],
			),
		)
		diagnostics.push({
			code: 'data-validation-extension-preservation',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message:
				relatedOperations.length > 0
					? x14DataValidationRiskMessage(
							x14DataValidationExtensionPayloads.length,
							relatedOperations,
						)
					: `${x14DataValidationExtensionPayloads.length} x14 data-validation extension payload(s) will be preserved as opaque worksheet XML; no data-validation semantic edit is planned.`,
			suggestedAction: x14DataValidationSuggestedAction(relatedOperations.length > 0),
			...(partPaths.length > 0 ? { partPaths } : {}),
			featureFamily: 'x14DataValidation',
			preservationPolicy: 'generated',
			preservationMode: 'generated-with-opaque-payload',
			details: {
				provenance: 'worksheet-extLst',
				preservationMode: 'opaque-payload-preserved-in-generated-worksheet-xml',
				semanticEditRisk: relatedOperations.length > 0,
				relatedOperations,
				x14DataValidations: x14DataValidationExtensionPayloads,
			},
		})
	}
	const externalLinkRisk = buildExternalLinkRisk(workbook, operations, externalReferences)
	const externalLinkIntegrityIssues = collectExternalLinkIntegrityIssues(checkIssues)
	const externalLinkPackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isExternalLinkPackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const externalLinkPartPaths = collectExternalLinkPartPaths(
		packageGraph.parts,
		externalReferences,
		externalLinkIntegrityIssues,
		externalLinkPackageIssuePartPaths,
	)
	const externalLinkPackageGraphIssues = packageGraphIssuesForParts(
		packageGraphAudit,
		externalLinkPartPaths,
		['preservedExternalLink'],
	)
	if (externalReferences.length > 0) {
		const externalReferencePartPaths = uniqueStrings(
			externalReferences.map((entry) => entry.partPath),
		)
		const relatedOperations = externalLinkRisk.relatedOperations.filter(
			(operation) => operation.externalReference?.partPath !== undefined,
		)
		diagnostics.push({
			code: 'external-link-dependency',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message:
				relatedOperations.length > 0
					? `${externalReferences.length} external workbook dependency link(s) are preserved by relationship-id metadata and ${relatedOperations.length} planned operation(s) reference them.`
					: `${externalReferences.length} external workbook dependency link(s) are preserved by relationship-id metadata.`,
			suggestedAction:
				relatedOperations.length > 0
					? 'Inspect details.externalLinks and details.relatedOperations before committing formula or defined-name edits that depend on external workbooks.'
					: 'Inspect externalReferenceDetails and externalReferenceUsages before rewriting workbook paths or trusting imported formula dependencies.',
			partPaths: externalReferencePartPaths,
			packageParts: packagePartDetails(packagePartByPath, externalReferencePartPaths),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				externalLinks: externalLinkRisk.linkGroups.filter(
					(group) => group.externalReference !== undefined,
				),
				relatedOperations,
				operationScoped: relatedOperations.length > 0,
			},
		})
	}
	if (externalReferences.length === 0 && externalLinkRisk.relatedOperations.length > 0) {
		diagnostics.push({
			code: 'external-link-dependency',
			severity: 'warning',
			message: `${externalLinkRisk.relatedOperations.length} planned operation(s) contain external workbook formula or defined-name references without imported external-link package metadata.`,
			suggestedAction:
				'Keep the formula or defined name text symbolic, inspect externalReferenceUsages after apply, and add external-link package metadata before relying on workbook path rebinding.',
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				externalLinks: externalLinkRisk.linkGroups,
				relatedOperations: externalLinkRisk.relatedOperations,
				operationScoped: true,
			},
		})
	}
	const externalReferenceBindingIssues = externalReferences.filter(
		(entry) =>
			entry.linkBindingStatus !== undefined && entry.linkBindingStatus !== 'externalBookRelId',
	)
	if (externalReferenceBindingIssues.length > 0) {
		const issuePartPaths = uniqueStrings(
			externalReferenceBindingIssues.map((entry) => entry.partPath),
		)
		const fallbackCount = externalReferenceBindingIssues.filter(
			(entry) => entry.linkBindingStatus === 'fallbackPathRelationship',
		).length
		const missingCount = externalReferenceBindingIssues.filter(
			(entry) => entry.linkBindingStatus === 'missingPathRelationship',
		).length
		const bindingIssueGroups = externalLinkRisk.linkGroups.filter((group) =>
			group.externalReference
				? externalReferenceBindingIssues.some(
						(entry) => entry.partPath === group.externalReference?.partPath,
					)
				: false,
		)
		const relatedOperations = bindingIssueGroups.flatMap((group) => group.relatedOperations)
		diagnostics.push({
			code: 'external-link-binding-risk',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message: `${externalReferenceBindingIssues.length} external workbook dependency link(s) have ambiguous package binding (${fallbackCount} fallback, ${missingCount} missing).`,
			suggestedAction:
				relatedOperations.length > 0
					? 'Use rewriteExternalLink with a stable partPath/linkRelId selector before or with these operation-scoped formula and defined-name edits.'
					: 'Use rewriteExternalLink with a stable partPath/linkRelId selector to repair or intentionally preserve the externalBook relationship binding before editing formulas or names that depend on it.',
			partPaths: issuePartPaths,
			packageParts: packagePartDetails(packagePartByPath, issuePartPaths),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				bindingIssueCounts: { fallback: fallbackCount, missing: missingCount },
				externalLinks: bindingIssueGroups,
				relatedOperations,
				operationScoped: relatedOperations.length > 0,
				rewriteExternalLinkRecommendations: bindingIssueGroups.map(
					externalLinkRewriteRecommendation,
				),
			},
		})
	}
	if (externalLinkIntegrityIssues.length > 0 || externalLinkPackageGraphIssues.length > 0) {
		diagnostics.push({
			code: 'external-link-package-risk',
			severity: 'warning',
			message: `${externalLinkIntegrityIssues.length + externalLinkPackageGraphIssues.length} external-link package issue(s) require inspection before write.`,
			suggestedAction:
				'Inspect externalReferenceDetails, external-link package relationships, externalBook r:id bindings, and packageGraphAudit before committing; use rewriteExternalLink only with a stable partPath/linkRelId selector.',
			...(externalLinkPartPaths.length > 0 ? { partPaths: externalLinkPartPaths } : {}),
			...(externalLinkPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, externalLinkPartPaths) }
				: {}),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			preservationMode: 'preserve-exact',
			details: {
				externalLinks: externalLinkRisk.linkGroups,
				packageGraphIssues: externalLinkPackageGraphIssues,
				verifyIssues: externalLinkIntegrityIssues,
				rewriteExternalLinkRecommendations: externalLinkRisk.linkGroups
					.filter((group) => group.externalReference !== undefined)
					.map(externalLinkRewriteRecommendation),
			},
		})
	}
	const genericPackageGraphIssues = packageGraphAudit.issues.filter(
		(issue) => !isRoutedPackageGraphIssue(issue),
	)
	if (genericPackageGraphIssues.length > 0) {
		const issuePartPaths = uniqueStrings(
			genericPackageGraphIssues.flatMap((issue) => packageGraphIssuePackagePaths(issue)),
		)
		const packageParts = packagePartDetails(packagePartByPath, issuePartPaths)
		diagnostics.push({
			code: 'package-graph-audit-issue',
			severity: 'warning',
			message: `${genericPackageGraphIssues.length} package graph issue(s) were found before write planning.`,
			suggestedAction:
				'Inspect packageGraphAudit.issues and resolve or explicitly accept package graph risk before committing.',
			partPaths: issuePartPaths,
			...(packageParts.length > 0 ? { packageParts } : {}),
			details: { packageGraphIssues: genericPackageGraphIssues },
		})
	}
	for (const feature of features) {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') continue
		if (isSafePackagePreservationFeature(feature)) continue
		const packageParts = packagePartDetails(packagePartByPath, feature.locations)
		const preservationMode =
			feature.tier === 'unsupported'
				? 'unsupported'
				: (packageParts[0]?.preservationMode ?? 'lossy-approval-required')
		diagnostics.push({
			code: 'approval-required-feature',
			severity: 'warning',
			message: `${feature.feature} (${feature.tier}) requires explicit approval before write.`,
			suggestedAction:
				'Inspect plan.approvals and pass only the corresponding approval id or allow-loss entry when intentional.',
			partPaths: feature.locations,
			...(packageParts.length > 0 ? { packageParts } : {}),
			featureFamily: feature.feature,
			preservationMode,
		})
	}
	const warningsOrBlockers = diagnostics.some((diagnostic) => diagnostic.severity !== 'info')
	const lossyApprovalRequiredFeatures = diagnostics.filter(
		(diagnostic) => diagnostic.code === 'approval-required-feature',
	).length
	return {
		ok: !warningsOrBlockers,
		diagnostics,
		summary: {
			generatedParts: generatedParts.length,
			copiedThroughParts: copiedThroughParts.length,
			skippedCapsules: skipped.length,
			invalidatedSignatures: skippedSignatureParts.length,
			approvalRequiredFeatures: diagnostics.filter(
				(diagnostic) => diagnostic.code === 'approval-required-feature',
			).length,
			packageGraphIssues: packageGraphAudit.issues.length,
			externalReferences: externalReferences.length,
			externalReferenceBindingIssues: externalReferenceBindingIssues.length,
			chartSourceIntegrityIssues: chartIntegrityIssues.length,
			legacyCommentLocations: legacyCommentLocations.length,
			threadedCommentLocations: threadedCommentLocations.length,
			commentIntegrityIssues:
				commentIntegrityIssues.legacy.length + commentIntegrityIssues.threaded.length,
			x14ConditionalFormatExtensionPayloads: x14ConditionalFormatExtensionPayloads.length,
			x14DataValidationExtensionPayloads: x14DataValidationExtensionPayloads.length,
			calcChainPolicy,
			preservationModes: {
				preserveExactParts: copiedThroughParts.filter(
					(part) => packageGraphPartByPath.get(part.path)?.preservationPolicy === 'preserve-exact',
				).length,
				generatedParts: generatedParts.length,
				generatedWithOpaquePayloads:
					x14ConditionalFormatExtensionPayloads.length + x14DataValidationExtensionPayloads.length,
				invalidatedOnEditParts: skippedSignatureParts.length,
				discardedForRecalcParts: skippedCalcChainParts.length,
				inspectOnlyParts: packageGraph.parts.filter(
					(part) => part.preservationPolicy === 'inspect-only',
				).length,
				reviewRequiredParts: packageGraph.parts.filter(
					(part) => part.preservationPolicy === 'unknown-review-required',
				).length,
				unsupportedFeatures: features.filter((feature) => feature.tier === 'unsupported').length,
				lossyApprovalRequiredFeatures,
			},
		},
	}
}

interface VisualWriteRisk {
	readonly relatedOperations: readonly VisualRelatedOperation[]
	readonly drawingmlVmlDrift: readonly VisualDrawingDriftRisk[]
	readonly chartSourceRefDrift: readonly ChartSourceRefDriftRisk[]
	readonly analyticsPivotRefreshRisk: readonly AnalyticsPivotRefreshRisk[]
}

interface VisualRelatedOperation {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheetName?: string | undefined
	readonly targetKind: 'image' | 'drawingText' | 'chartSource'
	readonly selector: Readonly<Record<string, unknown>>
	readonly matchCount: number
	readonly matches: readonly Record<string, unknown>[]
	readonly partPaths: readonly string[]
	readonly recommendation: string
}

interface VisualDrawingDriftRisk {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheetName: string
	readonly partPaths: readonly string[]
	readonly drawingMlObjectCount: number
	readonly vmlObjectCount: number
	readonly imageCount: number
	readonly recommendation: string
}

interface ChartSourceRefDriftRisk {
	readonly chartPartPath: string
	readonly sheetName?: string | undefined
	readonly chartIndex: number
	readonly seriesIndex: number
	readonly sourceRefs: readonly ChartSourceRefAudit[]
	readonly relatedOperations: readonly ChartSourceRelatedOperation[]
	readonly partPaths: readonly string[]
	readonly recommendation: Readonly<Record<string, unknown>>
}

interface ChartSourceRefAudit {
	readonly sourceKind: 'nameRef' | 'categoryRef' | 'valueRef'
	readonly ref: string
	readonly referencedSheets: readonly string[]
}

interface ChartSourceRelatedOperation {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheetName: string
	readonly rangeImpact: 'rows' | 'columns' | 'sheet'
	readonly at?: number | undefined
	readonly count?: number | undefined
	readonly newName?: string | undefined
}

interface AnalyticsPivotRefreshRisk {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly targetKind: 'pivotCache' | 'pivotFieldItem' | 'slicerCacheItem' | 'timelineRange'
	readonly selector: Readonly<Record<string, unknown>>
	readonly matchCount: number
	readonly matches: readonly Record<string, unknown>[]
	readonly linkedPivotTableNames: readonly string[]
	readonly partPaths: readonly string[]
	readonly recommendation: string
}

function buildVisualWriteRisk(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): VisualWriteRisk {
	return {
		relatedOperations: collectVisualRelatedOperations(workbook, operations),
		drawingmlVmlDrift: collectDrawingMlVmlDriftRisks(workbook, operations, parts),
		chartSourceRefDrift: collectChartSourceRefDriftRisks(workbook, operations, parts),
		analyticsPivotRefreshRisk: collectAnalyticsPivotRefreshRisks(workbook, operations, parts),
	}
}

function collectVisualRelatedOperations(
	workbook: Workbook,
	operations: readonly Operation[],
): VisualRelatedOperation[] {
	return operations.flatMap((operation, operationIndex) => {
		switch (operation.op) {
			case 'replaceImage':
			case 'deleteImage':
				return [imageReplacementRecommendation(workbook, operation, operationIndex)]
			case 'insertImage':
				return [imageInsertRecommendation(workbook, operation, operationIndex)]
			case 'setDrawingText':
				return [drawingTextRecommendation(workbook, operation, operationIndex)]
			case 'setChartSeriesSource':
				return [chartSourceReplacementRecommendation(workbook, operation, operationIndex)]
			default:
				return []
		}
	})
}

function imageReplacementRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'replaceImage' | 'deleteImage' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const sheet = workbook.getSheet(operation.sheet)
	const matches = sheet
		? sheet.imageRefs
				.map((image, index) => ({ image, index }))
				.filter(({ image, index }) => imageSelectorMatches(image, index, operation))
				.map(({ image, index }) => ({
					index,
					drawingPartPath: image.drawingPartPath,
					targetPath: image.targetPath,
					relId: image.relId,
					...(image.name ? { name: image.name } : {}),
				}))
		: []
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		targetKind: 'image',
		selector: compactRecord({
			targetPath: operation.targetPath,
			relId: operation.relId,
			name: operation.name,
			imageIndex: operation.imageIndex,
		}),
		matchCount: matches.length,
		matches,
		partPaths: uniqueStrings(matches.flatMap((match) => [match.drawingPartPath, match.targetPath])),
		recommendation:
			matches.length === 1
				? `${operation.op} has a unique image selector; keep targetPath or imageIndex in the edit record and verify the drawing relationship plus media bytes after write.`
				: `Use inspect --detail images or visualInventory to choose one image by targetPath, relId, name, or imageIndex before ${operation.op}.`,
	}
}

function imageInsertRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'insertImage' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const sheet = workbook.getSheet(operation.sheet)
	const collisions = sheet
		? sheet.imageRefs
				.map((image, index) => ({ image, index }))
				.filter(
					({ image }) =>
						(operation.targetPath !== undefined && image.targetPath === operation.targetPath) ||
						(operation.relId !== undefined && image.relId === operation.relId),
				)
				.map(({ image, index }) => ({
					index,
					drawingPartPath: image.drawingPartPath,
					targetPath: image.targetPath,
					relId: image.relId,
				}))
		: []
	const drawingPartPath =
		operation.drawingPartPath ?? sheet?.imageRefs[0]?.drawingPartPath ?? 'xl/drawings/drawing1.xml'
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		targetKind: 'image',
		selector: compactRecord({
			drawingPartPath: operation.drawingPartPath,
			targetPath: operation.targetPath,
			relId: operation.relId,
			name: operation.name,
		}),
		matchCount: collisions.length === 0 ? 1 : collisions.length,
		matches: collisions,
		partPaths: uniqueStrings([
			drawingPartPath,
			...(operation.targetPath ? [operation.targetPath] : []),
		]),
		recommendation:
			collisions.length === 0
				? 'insertImage can allocate missing media and relationship ids; provide an anchor when placement matters and verify the generated drawing relationship after write.'
				: 'insertImage selector collides with existing image identity; omit targetPath/relId or provide unique values before commit.',
	}
}

function drawingTextRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setDrawingText' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const sheet = workbook.getSheet(operation.sheet)
	const matches = sheet
		? sheet.drawingObjectRefs
				.map((object, index) => ({ object, index }))
				.filter(({ object, index }) => drawingObjectSelectorMatches(object, index, operation))
				.map(({ object, index }) => ({
					index,
					drawingPartPath: object.drawingPartPath,
					kind: object.kind,
					...(object.source ? { source: object.source } : {}),
					...(object.id !== undefined ? { id: object.id } : {}),
					...(object.name ? { name: object.name } : {}),
					editableText: object.text !== undefined,
				}))
		: []
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		targetKind: 'drawingText',
		selector: compactRecord({
			drawingPartPath: operation.drawingPartPath,
			id: operation.id,
			name: operation.name,
			drawingObjectIndex: operation.drawingObjectIndex,
		}),
		matchCount: matches.length,
		matches,
		partPaths: uniqueStrings(matches.map((match) => match.drawingPartPath)),
		recommendation:
			matches.length === 1 && matches[0]?.editableText === true
				? 'setDrawingText has a unique text-bearing object; preserve drawingPartPath plus id/name or drawingObjectIndex and verify the drawing part after write.'
				: 'Use inspect --detail visuals or visualInventory to select exactly one text-bearing DrawingML/VML object before setDrawingText.',
	}
}

function chartSourceReplacementRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setChartSeriesSource' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const matches = matchingCharts(workbook, operation).map((chart) => ({
		partPath: chart.partPath,
		...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
		seriesCount: chart.series.length,
		seriesExists: chart.series[operation.seriesIndex] !== undefined,
		currentSeries: chart.series[operation.seriesIndex],
		replacement: compactRecord({
			nameRef: operation.nameRef,
			categoryRef: operation.categoryRef,
			valueRef: operation.valueRef,
		}),
	}))
	return {
		operationIndex,
		op: operation.op,
		...(operation.sheet ? { sheetName: operation.sheet } : {}),
		targetKind: 'chartSource',
		selector: compactRecord({
			partPath: operation.partPath,
			sheet: operation.sheet,
			chartIndex: operation.chartIndex,
			seriesIndex: operation.seriesIndex,
		}),
		matchCount: matches.length,
		matches,
		partPaths: uniqueStrings(matches.map((match) => match.partPath)),
		recommendation:
			matches.length === 1 && matches[0]?.seriesExists === true
				? 'setChartSeriesSource has a unique chart and existing series; keep partPath/chartIndex/seriesIndex and verify name/category/value refs after write.'
				: 'Use inspect --detail visuals or visualInventory to select one chart by partPath or chartIndex and an existing seriesIndex before editing source refs.',
	}
}

function collectDrawingMlVmlDriftRisks(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): VisualDrawingDriftRisk[] {
	return operations.flatMap((operation, operationIndex) => {
		if (!isSheetVisualOperation(operation)) return []
		const sheet = workbook.getSheet(operation.sheet)
		if (!sheet?.drawingRefs.hasDrawing || !sheet.drawingRefs.hasLegacyDrawing) return []
		const drawingMlObjectCount = sheet.drawingObjectRefs.filter(
			(object) => object.source !== 'vml',
		).length
		const vmlObjectCount = sheet.drawingObjectRefs.filter(
			(object) => object.source === 'vml',
		).length
		return [
			{
				operationIndex,
				op: operation.op,
				sheetName: operation.sheet,
				drawingMlObjectCount,
				vmlObjectCount,
				imageCount: sheet.imageRefs.length,
				partPaths: sheetVisualPartPaths(sheet.imageRefs, sheet.drawingObjectRefs, parts),
				recommendation:
					'DrawingML and VML are separate package graphs; verify the selected object source and legacy VML/comment layout after write.',
			},
		]
	})
}

function collectChartSourceRefDriftRisks(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): ChartSourceRefDriftRisk[] {
	const risks: ChartSourceRefDriftRisk[] = []
	for (const [chartIndex, chart] of workbook.chartParts.entries()) {
		for (const [seriesIndex, series] of chart.series.entries()) {
			const sourceRefs = chartSourceRefs(workbook, series, chart.sheetName)
			const referencedSheets = new Set(sourceRefs.flatMap((source) => source.referencedSheets))
			const relatedOperations = chartSourceRelatedOperations(operations, referencedSheets)
			if (sourceRefs.length === 0 || relatedOperations.length === 0) continue
			risks.push({
				chartPartPath: chart.partPath,
				...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
				chartIndex,
				seriesIndex,
				sourceRefs,
				relatedOperations,
				partPaths: chartRelatedPartPaths(chart.partPath, parts),
				recommendation: {
					op: 'setChartSeriesSource',
					partPath: chart.partPath,
					chartIndex,
					seriesIndex,
					current: compactRecord({
						nameRef: series.nameRef,
						categoryRef: series.categoryRef,
						valueRef: series.valueRef,
					}),
				},
			})
		}
	}
	return risks
}

function collectAnalyticsPivotRefreshRisks(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): AnalyticsPivotRefreshRisk[] {
	return operations.flatMap((operation, operationIndex) => {
		switch (operation.op) {
			case 'setPivotCache':
				return [pivotCacheRefreshRisk(workbook, operation, operationIndex)]
			case 'setPivotFieldItem':
				return [pivotFieldItemRefreshRisk(workbook, operation, operationIndex)]
			case 'setSlicerCacheItem':
				return [slicerCacheItemRefreshRisk(workbook, operation, operationIndex, parts)]
			case 'setTimelineRange':
				return [timelineRangeRefreshRisk(workbook, operation, operationIndex, parts)]
			default:
				return []
		}
	})
}

function pivotCacheRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setPivotCache' }>,
	operationIndex: number,
): AnalyticsPivotRefreshRisk {
	const matches = workbook.pivotCaches
		.filter((cache) => pivotCacheSelectorMatches(workbook, cache, operation))
		.map((cache) => {
			const linkedPivots = workbook.pivotTables.filter(
				(pivot) => cache.cacheId !== undefined && pivot.cacheId === cache.cacheId,
			)
			return compactRecord({
				pivotCache: cache.partPath,
				cacheId: cache.cacheId,
				sourceSheet: cache.sourceSheet,
				sourceRef: cache.sourceRef,
				nextSourceSheet: operation.sourceSheet,
				nextSourceRef: operation.sourceRef,
				refreshOnLoad: operation.refreshOnLoad ?? cache.refreshOnLoad,
				invalid: operation.invalid ?? cache.invalid,
				saveData: operation.saveData ?? cache.saveData,
				outputState: pivotCacheOutputState(workbook, cache, linkedPivots),
				linkedPivotTableNames: linkedPivotNames(linkedPivots, []),
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCacheRecordsPartPath: cache.recordsPartPath,
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			Array.isArray(match.linkedPivotTableNames)
				? match.linkedPivotTableNames.filter((name): name is string => typeof name === 'string')
				: [],
		),
	)
	const partPaths = uniqueStrings(
		matches.flatMap((match) =>
			[
				match.pivotCache,
				match.pivotCacheRecordsPartPath,
				...(Array.isArray(match.pivotTablePartPaths) ? match.pivotTablePartPaths : []),
			].filter((path): path is string => typeof path === 'string' && path.length > 0),
		),
	)
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'pivotCache',
		selector: compactRecord({
			cacheId: operation.cacheId,
			partPath: operation.partPath,
			pivotTable: operation.pivotTable,
			sourceSheet: operation.sourceSheet,
			sourceRef: operation.sourceRef,
			refreshOnLoad: operation.refreshOnLoad,
			invalid: operation.invalid,
			saveData: operation.saveData,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1
				? 'setPivotCache matched one cache; mark invalid/refreshOnLoad when changing source metadata, then refresh linked pivots in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail pivots to select one pivot cache by cacheId, partPath, or pivotTable before changing source or freshness metadata.',
	}
}

function pivotFieldItemRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setPivotFieldItem' }>,
	operationIndex: number,
): AnalyticsPivotRefreshRisk {
	const matches = workbook.pivotTables
		.filter((pivot) => pivotTableSelectorMatches(pivot, operation))
		.map((pivot) => {
			const field = pivot.fields[operation.fieldIndex]
			const item = field?.items?.[operation.itemIndex]
			const cache = pivotCacheForTable(workbook, pivot)
			return compactRecord({
				pivotTable: pivot.name,
				sheetName: pivot.sheetName,
				partPath: pivot.partPath,
				cacheId: pivot.cacheId,
				cachePartPath: cache?.partPath,
				cacheRecordsPartPath: cache?.recordsPartPath,
				fieldIndex: operation.fieldIndex,
				fieldName: field?.name,
				fieldExists: field !== undefined,
				itemIndex: operation.itemIndex,
				itemCaption: item?.caption,
				itemExists: item !== undefined,
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			typeof match.pivotTable === 'string' && match.pivotTable.length > 0 ? [match.pivotTable] : [],
		),
	)
	const partPaths = uniqueStrings(
		matches.flatMap((match) =>
			[match.partPath, match.cachePartPath, match.cacheRecordsPartPath].filter(
				(path): path is string => typeof path === 'string' && path.length > 0,
			),
		),
	)
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'pivotFieldItem',
		selector: compactRecord({
			pivotTable: operation.pivotTable,
			partPath: operation.partPath,
			sheet: operation.sheet,
			fieldIndex: operation.fieldIndex,
			itemIndex: operation.itemIndex,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1 && matches[0]?.fieldExists === true && matches[0]?.itemExists === true
				? 'setPivotFieldItem matched one pivot field item; refresh the linked pivot cache/table in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail pivots to select one pivot table by name or partPath and an existing fieldIndex/itemIndex before editing filter state.',
	}
}

function slicerCacheItemRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setSlicerCacheItem' }>,
	operationIndex: number,
	parts: readonly XlsxPackageGraph['parts'][number][],
): AnalyticsPivotRefreshRisk {
	const matches = workbook.slicerCaches
		.filter((cache) => slicerCacheSelectorMatches(cache, operation))
		.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			const slicerPartPaths = workbook.slicers
				.filter((slicer) => slicer.cacheName === cache.name)
				.map((slicer) => slicer.partPath)
			return compactRecord({
				slicerCache: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				item: operation.item,
				itemExists: cache.items?.some((item) => item.index === operation.item) === true,
				linkedPivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				slicerPartPaths,
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				pivotCacheRecordsPartPaths: linkedCaches.flatMap((pivotCache) =>
					pivotCache.recordsPartPath ? [pivotCache.recordsPartPath] : [],
				),
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			Array.isArray(match.linkedPivotTableNames)
				? match.linkedPivotTableNames.filter((name): name is string => typeof name === 'string')
				: [],
		),
	)
	const partPaths = uniqueStrings([
		...matches.flatMap((match) =>
			[
				match.partPath,
				...(Array.isArray(match.slicerPartPaths) ? match.slicerPartPaths : []),
				...(Array.isArray(match.pivotTablePartPaths) ? match.pivotTablePartPaths : []),
				...(Array.isArray(match.pivotCachePartPaths) ? match.pivotCachePartPaths : []),
				...(Array.isArray(match.pivotCacheRecordsPartPaths)
					? match.pivotCacheRecordsPartPaths
					: []),
			].filter((path): path is string => typeof path === 'string' && path.length > 0),
		),
		...operationScopedAnalyticalPartPaths(operation.partPath, parts),
	])
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'slicerCacheItem',
		selector: compactRecord({
			slicerCache: operation.slicerCache,
			partPath: operation.partPath,
			item: operation.item,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1
				? 'setSlicerCacheItem matched one slicer cache; refresh linked pivot tables/caches in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail slicers to select one slicer cache by name or partPath before editing item state.',
	}
}

function timelineRangeRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setTimelineRange' }>,
	operationIndex: number,
	parts: readonly XlsxPackageGraph['parts'][number][],
): AnalyticsPivotRefreshRisk {
	const matches = workbook.timelineCaches
		.filter((cache) => timelineCacheSelectorMatches(cache, operation))
		.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			const timelinePartPaths = workbook.timelines
				.filter((timeline) => timeline.cacheName === cache.name)
				.map((timeline) => timeline.partPath)
			return compactRecord({
				timelineCache: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				startDate: operation.startDate,
				endDate: operation.endDate,
				linkedPivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				timelinePartPaths,
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				pivotCacheRecordsPartPaths: linkedCaches.flatMap((pivotCache) =>
					pivotCache.recordsPartPath ? [pivotCache.recordsPartPath] : [],
				),
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			Array.isArray(match.linkedPivotTableNames)
				? match.linkedPivotTableNames.filter((name): name is string => typeof name === 'string')
				: [],
		),
	)
	const partPaths = uniqueStrings([
		...matches.flatMap((match) =>
			[
				match.partPath,
				...(Array.isArray(match.timelinePartPaths) ? match.timelinePartPaths : []),
				...(Array.isArray(match.pivotTablePartPaths) ? match.pivotTablePartPaths : []),
				...(Array.isArray(match.pivotCachePartPaths) ? match.pivotCachePartPaths : []),
				...(Array.isArray(match.pivotCacheRecordsPartPaths)
					? match.pivotCacheRecordsPartPaths
					: []),
			].filter((path): path is string => typeof path === 'string' && path.length > 0),
		),
		...operationScopedAnalyticalPartPaths(operation.partPath, parts),
	])
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'timelineRange',
		selector: compactRecord({
			timelineCache: operation.timelineCache,
			partPath: operation.partPath,
			startDate: operation.startDate,
			endDate: operation.endDate,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1
				? 'setTimelineRange matched one timeline cache; refresh linked pivot tables/caches in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail timelines to select one timeline cache by name or partPath before editing date range state.',
	}
}

function imageSelectorMatches(
	image: SheetImageRef,
	index: number,
	operation: Extract<Operation, { op: 'replaceImage' | 'deleteImage' }>,
): boolean {
	if (
		operation.imageIndex === undefined &&
		operation.targetPath === undefined &&
		operation.relId === undefined &&
		operation.name === undefined
	) {
		return false
	}
	if (operation.imageIndex !== undefined && index !== operation.imageIndex) return false
	if (operation.targetPath !== undefined && image.targetPath !== operation.targetPath) return false
	if (operation.relId !== undefined && image.relId !== operation.relId) return false
	if (operation.name !== undefined && image.name !== operation.name) return false
	return true
}

function drawingObjectSelectorMatches(
	object: SheetDrawingObjectRef,
	index: number,
	operation: Extract<Operation, { op: 'setDrawingText' }>,
): boolean {
	if (
		operation.drawingObjectIndex === undefined &&
		operation.drawingPartPath === undefined &&
		operation.id === undefined &&
		operation.name === undefined
	) {
		return false
	}
	if (operation.drawingObjectIndex !== undefined && index !== operation.drawingObjectIndex) {
		return false
	}
	if (
		operation.drawingPartPath !== undefined &&
		object.drawingPartPath !== operation.drawingPartPath
	) {
		return false
	}
	if (operation.id !== undefined && object.id !== operation.id) return false
	if (operation.name !== undefined && object.name !== operation.name) return false
	return true
}

function matchingCharts(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setChartSeriesSource' }>,
): ChartPartInfo[] {
	let candidates = workbook.chartParts
	if (operation.partPath !== undefined) {
		candidates = candidates.filter((chart) => chart.partPath === operation.partPath)
	}
	if (operation.sheet !== undefined) {
		candidates = candidates.filter((chart) => chart.sheetName === operation.sheet)
	}
	if (operation.chartIndex !== undefined) {
		const chart = candidates[operation.chartIndex]
		return chart ? [chart] : []
	}
	return candidates
}

function pivotTableSelectorMatches(
	pivot: Workbook['pivotTables'][number],
	operation: Extract<Operation, { op: 'setPivotFieldItem' }>,
): boolean {
	if (operation.partPath !== undefined && pivot.partPath !== operation.partPath) return false
	if (operation.pivotTable !== undefined && pivot.name !== operation.pivotTable) return false
	if (operation.sheet !== undefined && pivot.sheetName !== operation.sheet) return false
	return true
}

function pivotCacheSelectorMatches(
	workbook: Workbook,
	cache: Workbook['pivotCaches'][number],
	operation: Extract<Operation, { op: 'setPivotCache' }>,
): boolean {
	if (operation.cacheId !== undefined && cache.cacheId !== operation.cacheId) return false
	if (operation.partPath !== undefined && cache.partPath !== operation.partPath) return false
	if (operation.pivotTable !== undefined) {
		return workbook.pivotTables.some(
			(pivot) =>
				pivot.name === operation.pivotTable &&
				pivot.cacheId !== undefined &&
				pivot.cacheId === cache.cacheId,
		)
	}
	return operation.cacheId !== undefined || operation.partPath !== undefined
}

function pivotCacheForTable(
	workbook: Workbook,
	pivot: Workbook['pivotTables'][number],
): Workbook['pivotCaches'][number] | undefined {
	if (pivot.cacheId === undefined) return undefined
	return workbook.pivotCaches.find((cache) => cache.cacheId === pivot.cacheId)
}

function slicerCacheSelectorMatches(
	cache: Workbook['slicerCaches'][number],
	operation: Extract<Operation, { op: 'setSlicerCacheItem' }>,
): boolean {
	if (operation.partPath !== undefined && cache.partPath !== operation.partPath) return false
	if (
		operation.slicerCache !== undefined &&
		cache.name !== operation.slicerCache &&
		cache.partPath !== operation.slicerCache
	) {
		return false
	}
	return true
}

function timelineCacheSelectorMatches(
	cache: Workbook['timelineCaches'][number],
	operation: Extract<Operation, { op: 'setTimelineRange' }>,
): boolean {
	if (operation.partPath !== undefined && cache.partPath !== operation.partPath) return false
	if (
		operation.timelineCache !== undefined &&
		cache.name !== operation.timelineCache &&
		cache.partPath !== operation.timelineCache
	) {
		return false
	}
	return true
}

function linkedPivotTables(
	workbook: Workbook,
	pivotTableNames: readonly string[],
	pivotCacheId: number | undefined,
): Workbook['pivotTables'] {
	const names = new Set(pivotTableNames)
	return workbook.pivotTables.filter((pivot) => {
		if (pivot.name !== undefined && names.has(pivot.name)) return true
		if (pivotCacheId !== undefined && pivot.cacheId === pivotCacheId) return true
		return false
	})
}

function linkedPivotCaches(
	workbook: Workbook,
	pivots: readonly Workbook['pivotTables'][number][],
	pivotCacheId: number | undefined,
): Workbook['pivotCaches'] {
	const cacheIds = new Set(
		[
			pivotCacheId,
			...pivots.flatMap((pivot) => (pivot.cacheId === undefined ? [] : [pivot.cacheId])),
		].filter((cacheId): cacheId is number => cacheId !== undefined),
	)
	return workbook.pivotCaches.filter((cache) => {
		if (cache.cacheId === undefined) return false
		return cacheIds.has(cache.cacheId)
	})
}

function linkedPivotNames(
	pivots: readonly Workbook['pivotTables'][number][],
	pivotTableNames: readonly string[],
): string[] {
	return uniqueStrings([
		...pivotTableNames,
		...pivots.flatMap((pivot) =>
			pivot.name !== undefined && pivot.name.length > 0 ? [pivot.name] : [],
		),
	])
}

function collectPivotCacheRisks(
	workbook: Workbook,
): ReadonlyArray<Readonly<Record<string, unknown>>> {
	return workbook.pivotCaches.map((cache) => {
		const linkedPivots = workbook.pivotTables.filter(
			(pivot) => cache.cacheId !== undefined && pivot.cacheId === cache.cacheId,
		)
		const outputState = pivotCacheOutputState(workbook, cache, linkedPivots)
		return compactRecord({
			partPath: cache.partPath,
			cacheId: cache.cacheId,
			sourceType: cache.sourceType,
			sourceSheet: cache.sourceSheet,
			sourceRef: cache.sourceRef,
			sourceName: cache.sourceName,
			refreshOnLoad: cache.refreshOnLoad,
			enableRefresh: cache.enableRefresh,
			invalid: cache.invalid,
			saveData: cache.saveData,
			recordCount: cache.recordCount,
			recordsPartPath: cache.recordsPartPath,
			outputState,
			requiresExternalRefresh: outputState !== 'cached',
			linkedPivotTables: linkedPivots.map((pivot) =>
				compactRecord({
					name: pivot.name,
					partPath: pivot.partPath,
					sheetName: pivot.sheetName,
					locationRef: pivot.locationRef,
				}),
			),
			warnings: pivotCacheRiskWarnings(cache, outputState),
			recommendedOps:
				outputState === 'cached'
					? []
					: [
							compactRecord({
								op: 'setPivotCache',
								partPath: cache.partPath,
								cacheId: cache.cacheId,
								refreshOnLoad: true,
								invalid: true,
								saveData: false,
							}),
						],
		})
	})
}

function collectSlicerTimelineCacheDependencies(
	workbook: Workbook,
): ReadonlyArray<Readonly<Record<string, unknown>>> {
	return [
		...workbook.slicerCaches.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			return compactRecord({
				kind: 'slicerCache',
				name: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				pivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				slicerPartPaths: workbook.slicers
					.filter((slicer) => slicer.cacheName === cache.name)
					.map((slicer) => slicer.partPath),
			})
		}),
		...workbook.timelineCaches.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			return compactRecord({
				kind: 'timelineCache',
				name: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				pivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				timelinePartPaths: workbook.timelines
					.filter((timeline) => timeline.cacheName === cache.name)
					.map((timeline) => timeline.partPath),
				state: cache.state,
			})
		}),
	]
}

function pivotCacheOutputState(
	workbook: Workbook,
	cache: Workbook['pivotCaches'][number],
	linkedPivots: readonly Workbook['pivotTables'][number][],
): 'cached' | 'stale' | 'refresh-on-open' | 'not-saved' | 'unknown' {
	if (cache.invalid) return 'stale'
	if (cache.refreshOnLoad) return 'refresh-on-open'
	if (cache.saveData === false) return 'not-saved'
	if (linkedPivots.some((pivot) => pivotSavedOutputState(workbook, pivot) === 'not-saved')) {
		return 'not-saved'
	}
	if (cache.recordsPartPath === undefined && cache.recordCount === undefined) return 'unknown'
	if (linkedPivots.some((pivot) => pivotSavedOutputState(workbook, pivot) === 'unknown')) {
		return 'unknown'
	}
	return 'cached'
}

function pivotSavedOutputState(
	workbook: Workbook,
	pivot: Workbook['pivotTables'][number],
): 'cached' | 'not-saved' | 'unknown' {
	if (!pivot.locationRef) return 'unknown'
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return 'unknown'
	}
	const sheet = workbook.sheets.find((entry) => entry.name === pivot.sheetName)
	if (!sheet) return 'unknown'
	for (let row = bounds.start.row; row <= bounds.end.row; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			const kind = (sheet.cells.get(row, col)?.value as { kind?: string } | undefined)?.kind
			if (kind !== undefined && kind !== 'empty') return 'cached'
		}
	}
	return 'not-saved'
}

function pivotCacheRiskWarnings(
	cache: Workbook['pivotCaches'][number],
	outputState: 'cached' | 'stale' | 'refresh-on-open' | 'not-saved' | 'unknown',
): string[] {
	const warnings: string[] = []
	if (outputState === 'stale') warnings.push('Pivot cache is marked invalid; output is stale.')
	if (outputState === 'refresh-on-open') {
		warnings.push('Pivot cache requests refresh on open; saved output may change.')
	}
	if (outputState === 'not-saved') {
		warnings.push('Pivot cache records or PivotTable output cells are not fully saved.')
	}
	if (outputState === 'unknown') warnings.push('Pivot cache/output freshness is unknown.')
	if (cache.sourceSheet !== undefined || cache.sourceRef !== undefined) {
		warnings.push(
			'Cache source metadata is inspectable, but output is not recalculated headlessly.',
		)
	}
	return warnings
}

function operationScopedAnalyticalPartPaths(
	partPath: string | undefined,
	parts: readonly XlsxPackageGraph['parts'][number][],
): string[] {
	if (partPath === undefined) return []
	return parts.some((part) => part.path === partPath && isAnalyticalSidecar(part.featureFamily))
		? [partPath]
		: []
}

function isSheetVisualOperation(
	operation: Operation,
): operation is Extract<
	Operation,
	{ op: 'replaceImage' | 'insertImage' | 'deleteImage' | 'setDrawingText' }
> {
	return (
		operation.op === 'replaceImage' ||
		operation.op === 'insertImage' ||
		operation.op === 'deleteImage' ||
		operation.op === 'setDrawingText'
	)
}

function sheetVisualPartPaths(
	imageRefs: readonly SheetImageRef[],
	drawingObjectRefs: readonly SheetDrawingObjectRef[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): string[] {
	const drawingPartPaths = uniqueStrings([
		...imageRefs.map((image) => image.drawingPartPath),
		...drawingObjectRefs.map((object) => object.drawingPartPath),
	])
	const mediaPartPaths = imageRefs.map((image) => image.targetPath)
	const vmlPartPaths = parts
		.filter((part) => part.featureFamily === 'preservedVml')
		.map((part) => part.path)
	return uniqueStrings([...drawingPartPaths, ...mediaPartPaths, ...vmlPartPaths])
}

function collectChartSourceRefSummary(workbook: Workbook): readonly Record<string, unknown>[] {
	return workbook.chartParts.map((chart, chartIndex) => ({
		partPath: chart.partPath,
		...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
		...(chart.chartType ? { chartType: chart.chartType } : {}),
		...(chart.title ? { title: chart.title } : {}),
		chartIndex,
		series: chart.series.map((series, seriesIndex) => ({
			seriesIndex,
			sourceRefs: chartSourceRefs(workbook, series, chart.sheetName),
		})),
	}))
}

function collectDrawingModelSummary(
	workbook: Workbook,
	parts: readonly XlsxPackageGraph['parts'][number][],
): Readonly<Record<string, unknown>> {
	return {
		sheets: workbook.sheets
			.filter(
				(sheet) =>
					sheet.drawingRefs.hasDrawing ||
					sheet.drawingRefs.hasLegacyDrawing ||
					sheet.imageRefs.length > 0 ||
					sheet.drawingObjectRefs.length > 0,
			)
			.map((sheet) => {
				const drawingMlObjectRefs = sheet.drawingObjectRefs.filter(
					(object) => object.source !== 'vml',
				)
				const vmlObjectRefs = sheet.drawingObjectRefs.filter((object) => object.source === 'vml')
				return {
					sheetName: sheet.name,
					hasDrawingMl: sheet.drawingRefs.hasDrawing,
					hasVml: sheet.drawingRefs.hasLegacyDrawing,
					imageCount: sheet.imageRefs.length,
					drawingMlObjectCount: drawingMlObjectRefs.length,
					vmlObjectCount: vmlObjectRefs.length,
					drawingMlPartPaths: uniqueStrings([
						...sheet.imageRefs.map((image) => image.drawingPartPath),
						...drawingMlObjectRefs.map((object) => object.drawingPartPath),
					]),
					vmlPartPaths: uniqueStrings(vmlObjectRefs.map((object) => object.drawingPartPath)),
				}
			}),
		packagePartCounts: {
			drawingMl: parts.filter((part) => part.featureFamily === 'preservedDrawing').length,
			media: parts.filter((part) => part.featureFamily === 'preservedMedia').length,
			chart: parts.filter((part) => part.featureFamily === 'preservedChart').length,
			chartSidecar: parts.filter(
				(part) =>
					part.featureFamily === 'preservedChartStyle' ||
					part.featureFamily === 'preservedChartColor',
			).length,
			vml: parts.filter((part) => part.featureFamily === 'preservedVml').length,
		},
		distinction:
			'DrawingML drawings, images, and charts use drawing/chart/media relationships; VML drawings are separate legacy shape/comment layout parts.',
	}
}

function chartSourceRefs(
	workbook: Workbook,
	series: ChartPartInfo['series'][number],
	defaultSheetName?: string,
): ChartSourceRefAudit[] {
	return [
		chartSourceRef(workbook, 'nameRef', series.nameRef, defaultSheetName),
		chartSourceRef(workbook, 'categoryRef', series.categoryRef, defaultSheetName),
		chartSourceRef(workbook, 'valueRef', series.valueRef, defaultSheetName),
	].filter((entry): entry is ChartSourceRefAudit => entry !== undefined)
}

function chartSourceRef(
	workbook: Workbook,
	sourceKind: ChartSourceRefAudit['sourceKind'],
	ref: string | undefined,
	defaultSheetName?: string,
): ChartSourceRefAudit | undefined {
	if (!ref) return undefined
	return {
		sourceKind,
		ref,
		referencedSheets: chartSourceReferencedSheets(workbook, ref, defaultSheetName),
	}
}

function chartSourceReferencedSheets(
	workbook: Workbook,
	ref: string,
	defaultSheetName?: string,
): string[] {
	const parsed = parseFormula(normalizeFormulaInput(ref))
	if (parsed.ok) {
		const references = flattenFormulaReferences(collectFormulaReferences(parsed.value))
		const hasExternalReference = references.some(
			(reference) =>
				reference.scope?.kind === 'external' || reference.scope?.kind === 'externalSheetSpan',
		)
		const sheets = references.flatMap((reference) => {
			if (reference.kind === 'structured') {
				return chartStructuredReferenceSheetNames(workbook, reference, defaultSheetName)
			}
			if (reference.scope?.kind === 'sheet') return [reference.scope.sheet]
			if (reference.scope?.kind === 'sheetSpan') {
				return [reference.scope.startSheet, reference.scope.endSheet]
			}
			if (reference.scope?.kind === 'local' && defaultSheetName) return [defaultSheetName]
			return []
		})
		if (sheets.length > 0) return uniqueStrings(sheets)
		if (hasExternalReference) return []
	}
	const regexMatch = /^'([^']|'')+'!|^([^'!:]+)!/.exec(ref)
	if (regexMatch?.[0]) {
		return [regexMatch[0].slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'")]
	}
	return defaultSheetName ? [defaultSheetName] : []
}

function chartStructuredReferenceSheetNames(
	workbook: Workbook,
	reference: Extract<FormulaReferenceInfo, { kind: 'structured' }>,
	defaultSheetName?: string,
): string[] {
	const tableName = reference.table.trim()
	if (tableName.length === 0) return defaultSheetName ? [defaultSheetName] : []
	const normalizedTableName = tableName.toLowerCase()
	return uniqueStrings(
		workbook.sheets.flatMap((sheet) =>
			sheet.tables.some((table) => table.name.toLowerCase() === normalizedTableName)
				? [sheet.name]
				: [],
		),
	)
}

function chartSourceRelatedOperations(
	operations: readonly Operation[],
	referencedSheets: ReadonlySet<string>,
): ChartSourceRelatedOperation[] {
	const related: ChartSourceRelatedOperation[] = []
	operations.forEach((operation, operationIndex) => {
		switch (operation.op) {
			case 'insertRows':
			case 'deleteRows':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'rows',
					at: operation.at,
					count: operation.count,
				})
				break
			case 'insertCols':
			case 'deleteCols':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'columns',
					at: operation.at,
					count: operation.count,
				})
				break
			case 'renameSheet':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'sheet',
					newName: operation.newName,
				})
				break
			case 'deleteSheet':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'sheet',
				})
				break
		}
	})
	return related
}

function chartRelatedPartPaths(
	chartPartPath: string,
	parts: readonly XlsxPackageGraph['parts'][number][],
): string[] {
	return uniqueStrings(
		parts
			.filter(
				(part) =>
					part.path === chartPartPath ||
					part.featureFamily === 'preservedChartStyle' ||
					part.featureFamily === 'preservedChartColor',
			)
			.map((part) => part.path),
	)
}

function compactRecord(
	value: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
	return Object.fromEntries(
		Object.entries(value).filter(([, entryValue]) => entryValue !== undefined),
	)
}

interface ExternalLinkRisk {
	readonly linkGroups: readonly ExternalLinkRiskGroup[]
	readonly relatedOperations: readonly ExternalLinkRelatedOperation[]
}

interface ExternalLinkRiskGroup {
	readonly key: string
	readonly workbook?: string | undefined
	readonly sheet?: string | undefined
	readonly sheetSpans?: readonly { readonly startSheet: string; readonly endSheet: string }[]
	readonly references: readonly string[]
	readonly externalReference?: ExternalReferenceInfo | undefined
	readonly bindingRisk?: ExternalLinkBindingRisk | undefined
	readonly relatedOperations: readonly ExternalLinkRelatedOperation[]
}

interface ExternalLinkBindingRisk {
	readonly status: NonNullable<ExternalReferenceInfo['linkBindingStatus']>
	readonly externalBookRelId?: string | undefined
	readonly linkRelId?: string | undefined
	readonly linkRelationshipPart?: string | undefined
	readonly linkRelationshipKind?: ExternalReferenceInfo['linkRelationshipKind'] | undefined
	readonly linkRelationshipRawTarget?: string | undefined
	readonly target?: string | undefined
	readonly targetMode?: string | undefined
}

interface ExternalLinkRelatedOperation {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sourceKind:
		| 'cellFormula'
		| 'fillFormula'
		| 'definedName'
		| 'deletedDefinedName'
		| 'externalLinkRewrite'
		| 'dataValidation'
		| 'conditionalFormat'
		| 'tableColumnFormula'
		| 'tableTotalsRowFormula'
		| 'chartSeriesName'
		| 'chartSeriesCategory'
		| 'chartSeriesValue'
		| 'sparklineGroupRange'
	readonly sheetName?: string | undefined
	readonly sourceRef?: string | undefined
	readonly range?: string | undefined
	readonly name?: string | undefined
	readonly scope?: string | undefined
	readonly formula?: string | undefined
	readonly workbook?: string | undefined
	readonly sheet?: string | undefined
	readonly sheetSpan?: { readonly startSheet: string; readonly endSheet: string } | undefined
	readonly references?: readonly string[] | undefined
	readonly externalReference?: ExternalReferenceInfo | undefined
	readonly selector?: ExternalLinkRewriteSelector | undefined
	readonly newTarget?: string | undefined
	readonly targetMode?: string | undefined
}

interface ExternalLinkRewriteSelector {
	readonly partPath?: string | undefined
	readonly relId?: string | undefined
	readonly linkRelId?: string | undefined
	readonly target?: string | undefined
}

interface ExternalFormulaReferenceGroup {
	readonly workbook: string
	readonly sheet?: string | undefined
	readonly sheetSpan?: { readonly startSheet: string; readonly endSheet: string } | undefined
	readonly references: readonly string[]
	readonly externalReference?: ExternalReferenceInfo | undefined
}

function buildExternalLinkRisk(
	workbook: Workbook,
	operations: readonly Operation[],
	externalReferences: readonly ExternalReferenceInfo[],
): ExternalLinkRisk {
	const relatedOperations = collectExternalLinkRelatedOperations(
		workbook,
		operations,
		externalReferences,
	)
	const groupByKey = new Map<string, ExternalLinkRiskGroup>()
	for (const reference of externalReferences) {
		groupByKey.set(externalLinkGroupKey(reference), {
			key: externalLinkGroupKey(reference),
			externalReference: copyExternalReferenceInfo(reference),
			...(externalLinkBindingRisk(reference)
				? { bindingRisk: externalLinkBindingRisk(reference) }
				: {}),
			references: [],
			relatedOperations: [],
		})
	}
	for (const operation of relatedOperations) {
		const key = operation.externalReference
			? externalLinkGroupKey(operation.externalReference)
			: `token:${operation.workbook ?? operation.selector?.target ?? operation.operationIndex}`
		const existing = groupByKey.get(key)
		const references = uniqueStrings([
			...(existing?.references ?? []),
			...(operation.references ?? []),
		])
		const sheetSpans = uniqueSheetSpans([
			...(existing?.sheetSpans ?? []),
			...(operation.sheetSpan ? [operation.sheetSpan] : []),
		])
		groupByKey.set(key, {
			key,
			...((operation.workbook ?? existing?.workbook)
				? { workbook: operation.workbook ?? existing?.workbook }
				: {}),
			...((operation.sheet ?? existing?.sheet)
				? { sheet: operation.sheet ?? existing?.sheet }
				: {}),
			...(sheetSpans.length > 0 ? { sheetSpans } : {}),
			references,
			...((operation.externalReference ?? existing?.externalReference)
				? {
						externalReference: copyExternalReferenceInfo(
							operation.externalReference ?? existing?.externalReference,
						),
					}
				: {}),
			...(externalLinkBindingRisk(operation.externalReference ?? existing?.externalReference)
				? {
						bindingRisk: externalLinkBindingRisk(
							operation.externalReference ?? existing?.externalReference,
						),
					}
				: {}),
			relatedOperations: [...(existing?.relatedOperations ?? []), operation],
		})
	}
	return {
		linkGroups: [...groupByKey.values()],
		relatedOperations,
	}
}

function collectExternalLinkRelatedOperations(
	workbook: Workbook,
	operations: readonly Operation[],
	externalReferences: readonly ExternalReferenceInfo[],
): ExternalLinkRelatedOperation[] {
	const related: ExternalLinkRelatedOperation[] = []
	operations.forEach((operation, operationIndex) => {
		switch (operation.op) {
			case 'setFormula':
				pushExternalLinkFormulaOperations(related, externalReferences, operationIndex, operation, [
					{
						sourceKind: 'cellFormula',
						sheetName: operation.sheet,
						sourceRef: `${operation.sheet}!${operation.ref}`,
						formula: operation.formula,
					},
				])
				break
			case 'fillFormula':
				pushExternalLinkFormulaOperations(related, externalReferences, operationIndex, operation, [
					{
						sourceKind: 'fillFormula',
						sheetName: operation.sheet,
						range: operation.range,
						formula: operation.formula,
					},
				])
				break
			case 'setDefinedName':
				pushExternalLinkFormulaOperations(related, externalReferences, operationIndex, operation, [
					{
						sourceKind: 'definedName',
						name: operation.name,
						...(operation.scope ? { scope: operation.scope } : {}),
						formula: operation.ref,
					},
				])
				break
			case 'deleteDefinedName':
				for (const name of matchingDefinedNames(workbook, operation)) {
					pushExternalLinkFormulaOperations(
						related,
						externalReferences,
						operationIndex,
						operation,
						[
							{
								sourceKind: 'deletedDefinedName',
								name: name.name,
								...(definedNameScopeName(workbook, name.scope)
									? { scope: definedNameScopeName(workbook, name.scope) }
									: {}),
								formula: name.formula,
							},
						],
					)
				}
				break
			case 'setDataValidation':
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					[operation.rule.formula1, operation.rule.formula2]
						.filter((formula): formula is string => formula !== undefined)
						.map((formula) => ({
							sourceKind: 'dataValidation',
							sheetName: operation.sheet,
							sourceRef: `${operation.sheet}!${operation.range}`,
							range: operation.range,
							formula,
						})),
				)
				break
			case 'setConditionalFormat':
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					conditionalFormatOperationFormulaTexts(operation.rule).map((formula) => ({
						sourceKind: 'conditionalFormat',
						sheetName: operation.sheet,
						sourceRef: `${operation.sheet}!${operation.range}`,
						range: operation.range,
						formula,
					})),
				)
				break
			case 'setTableColumn': {
				const sourceRef = `${operation.table}[${String(operation.column)}]`
				const sources: ExternalLinkFormulaOperationSource[] = []
				if (operation.formula) {
					sources.push({
						sourceKind: 'tableColumnFormula',
						sourceRef,
						name: operation.table,
						formula: operation.formula,
					})
				}
				if (operation.totalsRowFormula) {
					sources.push({
						sourceKind: 'tableTotalsRowFormula',
						sourceRef,
						name: operation.table,
						formula: operation.totalsRowFormula,
					})
				}
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					sources,
				)
				break
			}
			case 'setChartSeriesSource': {
				const sourceRef = chartOperationSourceRef(operation)
				const sources: ExternalLinkFormulaOperationSource[] = []
				if (operation.nameRef) {
					sources.push({
						sourceKind: 'chartSeriesName',
						sourceRef,
						formula: operation.nameRef,
					})
				}
				if (operation.categoryRef) {
					sources.push({
						sourceKind: 'chartSeriesCategory',
						sourceRef,
						formula: operation.categoryRef,
					})
				}
				if (operation.valueRef) {
					sources.push({
						sourceKind: 'chartSeriesValue',
						sourceRef,
						formula: operation.valueRef,
					})
				}
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					sources,
				)
				break
			}
			case 'setSparklineGroup':
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					operation.range
						? [
								{
									sourceKind: 'sparklineGroupRange',
									sheetName: operation.sheet,
									sourceRef: `${operation.sheet}!sparklineGroup${operation.groupIndex}`,
									range: operation.range,
									formula: operation.range,
								},
							]
						: [],
				)
				break
			case 'rewriteExternalLink':
				for (const reference of externalReferences.filter((entry) =>
					rewriteExternalLinkMatches(operation, entry),
				)) {
					related.push({
						operationIndex,
						op: operation.op,
						sourceKind: 'externalLinkRewrite',
						externalReference: copyExternalReferenceInfo(reference),
						selector: externalLinkRewriteSelector(operation),
						newTarget: operation.newTarget,
						...(operation.targetMode ? { targetMode: operation.targetMode } : {}),
					})
				}
				break
		}
	})
	return related
}

type ExternalLinkFormulaOperationSource = Pick<
	ExternalLinkRelatedOperation,
	'sourceKind' | 'sheetName' | 'sourceRef' | 'range' | 'name' | 'scope' | 'formula'
> & {
	readonly formula: string
}

function pushExternalLinkFormulaOperations(
	related: ExternalLinkRelatedOperation[],
	externalReferences: readonly ExternalReferenceInfo[],
	operationIndex: number,
	operation: Operation,
	sources: readonly ExternalLinkFormulaOperationSource[],
): void {
	for (const source of sources) {
		for (const group of externalFormulaReferenceGroups(source.formula, externalReferences)) {
			related.push({
				operationIndex,
				op: operation.op,
				...source,
				workbook: group.workbook,
				...(group.sheet ? { sheet: group.sheet } : {}),
				...(group.sheetSpan ? { sheetSpan: group.sheetSpan } : {}),
				references: group.references,
				...(group.externalReference
					? { externalReference: copyExternalReferenceInfo(group.externalReference) }
					: {}),
			})
		}
	}
}

function conditionalFormatOperationFormulaTexts(
	rule: Extract<Operation, { op: 'setConditionalFormat' }>['rule'],
): string[] {
	return [
		...[rule.formula, rule.formula2].filter((formula): formula is string => formula !== undefined),
		...conditionalFormatValueObjectFormulas(rule.colorScale?.cfvo),
		...conditionalFormatValueObjectFormulas(rule.dataBar?.cfvo),
		...conditionalFormatValueObjectFormulas(rule.iconSet?.cfvo),
	]
}

function conditionalFormatValueObjectFormulas(
	values:
		| readonly {
				readonly type?: string
				readonly value?: string
		  }[]
		| undefined,
): string[] {
	return (
		values?.flatMap((value) => (value.type === 'formula' && value.value ? [value.value] : [])) ?? []
	)
}

function chartOperationSourceRef(
	operation: Extract<Operation, { op: 'setChartSeriesSource' }>,
): string {
	if (operation.partPath) return `${operation.partPath}#series${operation.seriesIndex}`
	if (operation.sheet)
		return `${operation.sheet}!chart${operation.chartIndex ?? 0}#series${operation.seriesIndex}`
	return `chart${operation.chartIndex ?? 0}#series${operation.seriesIndex}`
}

function externalFormulaReferenceGroups(
	formula: string,
	externalReferences: readonly ExternalReferenceInfo[],
): ExternalFormulaReferenceGroup[] {
	const parsed = parseFormula(normalizeFormulaInput(formula))
	if (!parsed.ok) return []
	const groups = new Map<string, ExternalFormulaReferenceGroup>()
	for (const reference of flattenFormulaReferences(collectFormulaReferences(parsed.value))) {
		if (reference.scope?.kind !== 'external' && reference.scope?.kind !== 'externalSheetSpan') {
			continue
		}
		const key =
			reference.scope.kind === 'external'
				? `${reference.scope.workbook}\u0000${reference.scope.sheet}`
				: `${reference.scope.workbook}\u0000${reference.scope.startSheet}\u0000${reference.scope.endSheet}`
		const existing = groups.get(key)
		const references = uniqueStrings([...(existing?.references ?? []), reference.text])
		groups.set(key, {
			workbook: reference.scope.workbook,
			...(reference.scope.kind === 'external'
				? { sheet: reference.scope.sheet }
				: {
						sheetSpan: {
							startSheet: reference.scope.startSheet,
							endSheet: reference.scope.endSheet,
						},
					}),
			references,
			...(existing?.externalReference
				? { externalReference: existing.externalReference }
				: {
						...(resolveExternalReference(externalReferences, reference.scope.workbook)
							? {
									externalReference: resolveExternalReference(
										externalReferences,
										reference.scope.workbook,
									),
								}
							: {}),
					}),
		})
	}
	return [...groups.values()]
}

function flattenFormulaReferences(
	references: readonly FormulaReferenceInfo[],
): FormulaReferenceInfo[] {
	const flattened: FormulaReferenceInfo[] = []
	for (const reference of references) {
		if (reference.kind === 'union' || reference.kind === 'intersection') {
			flattened.push(...flattenFormulaReferences(reference.members))
			continue
		}
		flattened.push(reference)
	}
	return flattened
}

function resolveExternalReference(
	details: readonly ExternalReferenceInfo[],
	workbookToken: string,
): ExternalReferenceInfo | undefined {
	const numericIndex = parseExternalReferenceIndex(workbookToken)
	if (numericIndex !== undefined) return copyExternalReferenceInfo(details[numericIndex])
	const tokenName = externalReferenceBasename(workbookToken)
	const matches = details.filter((entry) => {
		if (entry.target === workbookToken || entry.partPath === workbookToken) return true
		return entry.target !== undefined && externalReferenceBasename(entry.target) === tokenName
	})
	return matches.length === 1 ? copyExternalReferenceInfo(matches[0]) : undefined
}

function uniqueSheetSpans(
	sheetSpans: readonly { readonly startSheet: string; readonly endSheet: string }[],
): readonly { readonly startSheet: string; readonly endSheet: string }[] {
	const seen = new Set<string>()
	const result: { readonly startSheet: string; readonly endSheet: string }[] = []
	for (const span of sheetSpans) {
		const key = `${span.startSheet}\u0000${span.endSheet}`
		if (seen.has(key)) continue
		seen.add(key)
		result.push(span)
	}
	return result
}

function parseExternalReferenceIndex(workbookToken: string): number | undefined {
	if (!/^\d+$/.test(workbookToken)) return undefined
	const index = Number(workbookToken)
	return Number.isSafeInteger(index) && index > 0 ? index - 1 : undefined
}

function externalReferenceBasename(path: string): string {
	const normalized = path.replace(/\\/g, '/')
	return normalized.slice(normalized.lastIndexOf('/') + 1)
}

function matchingDefinedNames(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'deleteDefinedName' }>,
): ReturnType<Workbook['definedNames']['list']> {
	return workbook.definedNames.list().filter((name) => {
		if (name.name !== operation.name) return false
		if (!operation.scope) return name.scope.kind === 'workbook'
		return definedNameScopeName(workbook, name.scope) === operation.scope
	})
}

function definedNameScopeName(
	workbook: Workbook,
	scope: ReturnType<Workbook['definedNames']['list']>[number]['scope'],
): string | undefined {
	if (scope.kind === 'workbook') return undefined
	return workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name
}

function rewriteExternalLinkMatches(
	operation: Extract<Operation, { op: 'rewriteExternalLink' }>,
	reference: ExternalReferenceInfo,
): boolean {
	if (operation.partPath !== undefined && reference.partPath !== operation.partPath) return false
	if (operation.relId !== undefined && reference.relId !== operation.relId) return false
	if (operation.linkRelId !== undefined && reference.linkRelId !== operation.linkRelId) return false
	if (operation.target !== undefined && reference.target !== operation.target) return false
	return (
		operation.partPath !== undefined ||
		operation.relId !== undefined ||
		operation.linkRelId !== undefined ||
		operation.target !== undefined
	)
}

function externalLinkRewriteSelector(
	operation: Extract<Operation, { op: 'rewriteExternalLink' }>,
): ExternalLinkRewriteSelector {
	return {
		...(operation.partPath ? { partPath: operation.partPath } : {}),
		...(operation.relId ? { relId: operation.relId } : {}),
		...(operation.linkRelId ? { linkRelId: operation.linkRelId } : {}),
		...(operation.target ? { target: operation.target } : {}),
	}
}

function externalLinkGroupKey(reference: ExternalReferenceInfo): string {
	return `part:${reference.partPath}`
}

function externalLinkBindingRisk(
	reference: ExternalReferenceInfo | undefined,
): ExternalLinkBindingRisk | undefined {
	if (!reference?.linkBindingStatus || reference.linkBindingStatus === 'externalBookRelId') {
		return undefined
	}
	return {
		status: reference.linkBindingStatus,
		...(reference.externalBookRelId ? { externalBookRelId: reference.externalBookRelId } : {}),
		...(reference.linkRelId ? { linkRelId: reference.linkRelId } : {}),
		...(reference.linkRelationshipPart
			? { linkRelationshipPart: reference.linkRelationshipPart }
			: {}),
		...(reference.linkRelationshipKind
			? { linkRelationshipKind: reference.linkRelationshipKind }
			: {}),
		...(reference.linkRelationshipRawTarget
			? { linkRelationshipRawTarget: reference.linkRelationshipRawTarget }
			: {}),
		...(reference.target ? { target: reference.target } : {}),
		...(reference.targetMode ? { targetMode: reference.targetMode } : {}),
	}
}

function externalLinkRewriteRecommendation(group: ExternalLinkRiskGroup): {
	readonly op: 'rewriteExternalLink'
	readonly partPath?: string | undefined
	readonly linkRelId?: string | undefined
	readonly target?: string | undefined
	readonly newTarget: '<new-target>'
	readonly targetMode?: string | undefined
} {
	const reference = group.externalReference
	return {
		op: 'rewriteExternalLink',
		...(reference?.partPath ? { partPath: reference.partPath } : {}),
		...(reference?.linkRelId ? { linkRelId: reference.linkRelId } : {}),
		...(reference?.target ? { target: reference.target } : {}),
		newTarget: '<new-target>',
		...(reference?.targetMode ? { targetMode: reference.targetMode } : {}),
	}
}

function copyExternalReferenceInfo(
	reference: ExternalReferenceInfo | undefined,
): ExternalReferenceInfo | undefined {
	return reference ? { ...reference } : undefined
}

interface LegacyCommentLocation {
	readonly sheetName: string
	readonly ref: string
	readonly location: string
	readonly author?: string
	readonly hasLegacyDrawing: boolean
	readonly shapeId?: string
	readonly vmlTarget?: { readonly row?: number; readonly column?: number }
}

interface ThreadedCommentLocation {
	readonly sheetName: string
	readonly ref: string
	readonly location: string
	readonly commentIndex: number
	readonly partPath?: string
	readonly id?: string
	readonly parentId?: string
	readonly personId?: string
	readonly author?: string
}

interface TableLocation {
	readonly sheetName: string
	readonly tableName: string
	readonly location: string
	readonly partPath?: string
	readonly tableId?: string
	readonly queryTablePartPath?: string
	readonly queryTableRelationshipId?: string
}

interface TableRelatedOperation {
	readonly operationIndex: number
	readonly op: string
	readonly sheetName?: string
	readonly tableName?: string
	readonly ref?: string
	readonly at?: number
	readonly count?: number
	readonly axis?: 'row' | 'column'
}

interface CommentRelatedOperation {
	readonly operationIndex: number
	readonly op: string
	readonly sheetName?: string
	readonly ref?: string
	readonly source?: string
	readonly target?: string
	readonly targetSheet?: string
	readonly mode?: string
	readonly at?: number
	readonly count?: number
	readonly axis?: 'row' | 'column'
	readonly range?: string
	readonly affectedLocations?: readonly string[]
	readonly partPath?: string
	readonly threadedCommentId?: string
	readonly commentIndex?: number
}

function collectCommentIntegrityIssues(checkIssues: readonly CheckIssue[]): {
	readonly legacy: readonly CheckIssue[]
	readonly threaded: readonly CheckIssue[]
} {
	return {
		legacy: checkIssues.filter(isLegacyCommentIntegrityIssue),
		threaded: checkIssues.filter(isThreadedCommentIntegrityIssue),
	}
}

function collectTableIntegrityIssues(checkIssues: readonly CheckIssue[]): readonly CheckIssue[] {
	return checkIssues.filter(
		(issue) =>
			issue.rule === 'table-integrity' ||
			issue.rule === 'table-package-integrity' ||
			issue.rule === 'table-query-integrity' ||
			(issue.rule === 'package-graph-integrity' &&
				(checkIssueFeatureFamilies(issue).some(isTableFeatureFamily) ||
					checkIssuePackagePartPaths(issue).some(isTablePackagePartPath))),
	)
}

function collectChartIntegrityIssues(checkIssues: readonly CheckIssue[]): readonly CheckIssue[] {
	return checkIssues.filter((issue) => issue.rule === 'chart-series-integrity')
}

function chartSourceRefRiskMessage(driftCount: number, integrityIssueCount: number): string {
	if (driftCount > 0 && integrityIssueCount > 0) {
		return `${driftCount} chart source reference group(s) may drift under planned topology edits and ${integrityIssueCount} chart source verify issue(s) require repair.`
	}
	if (integrityIssueCount > 0) {
		return `${integrityIssueCount} chart source verify issue(s) require repair before write.`
	}
	return `${driftCount} chart source reference group(s) may drift under planned sheet, row, or column topology edits.`
}

function collectExternalLinkIntegrityIssues(
	checkIssues: readonly CheckIssue[],
): readonly CheckIssue[] {
	return checkIssues.filter(isExternalLinkIntegrityIssue)
}

function collectExternalLinkPartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	externalReferences: readonly ExternalReferenceInfo[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	return uniqueStrings([
		...parts
			.filter((part) => isExternalLinkFeatureFamily(part.featureFamily))
			.map((part) => part.path),
		...externalReferences.flatMap((reference) => [
			reference.partPath,
			...(reference.linkRelationshipPart ? [reference.linkRelationshipPart] : []),
		]),
		...issues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isExternalLinkPackagePartPath),
		),
		...packageIssuePartPaths.filter(isExternalLinkPackagePartPath),
	])
}

function collectTableLocations(workbook: Workbook): TableLocation[] {
	return workbook.sheets.flatMap((sheet) =>
		sheet.tables.map((table) => {
			const location = `${sheet.name}!${rangeRefToA1(table.ref)}`
			return {
				sheetName: sheet.name,
				tableName: table.name,
				location,
				...(table.partPath ? { partPath: table.partPath } : {}),
				tableId: String(table.id),
				...(table.queryTable?.partPath ? { queryTablePartPath: table.queryTable.partPath } : {}),
				...(table.queryTable?.relationshipId
					? { queryTableRelationshipId: table.queryTable.relationshipId }
					: {}),
			}
		}),
	)
}

function collectTablePartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	locations: readonly TableLocation[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	return uniqueStrings([
		...parts.filter((part) => isTableFeatureFamily(part.featureFamily)).map((part) => part.path),
		...locations.flatMap((location) => [
			...(location.partPath ? [location.partPath] : []),
			...(location.queryTablePartPath ? [location.queryTablePartPath] : []),
		]),
		...issues.flatMap((issue) => checkIssuePackagePartPaths(issue).filter(isTablePackagePartPath)),
		...packageIssuePartPaths.filter(isTablePackagePartPath),
	])
}

function collectTableRelatedOperations(
	workbook: Workbook,
	operations: readonly Operation[],
): TableRelatedOperation[] {
	const tableSheets = new Set(
		workbook.sheets.filter((sheet) => sheet.tables.length > 0).map((sheet) => sheet.name),
	)
	const related: TableRelatedOperation[] = []
	for (const [operationIndex, op] of operations.entries()) {
		switch (op.op) {
			case 'createTable':
				related.push({
					operationIndex,
					op: op.op,
					sheetName: op.sheet,
					tableName: op.name,
					ref: op.ref,
				})
				break
			case 'appendRows':
			case 'deleteTable':
			case 'renameTable':
			case 'resizeTable':
			case 'setTableColumn':
			case 'setTableStyle':
				related.push({
					operationIndex,
					op: op.op,
					tableName: op.table,
					...('ref' in op ? { ref: op.ref } : {}),
				})
				break
			case 'insertRows':
			case 'deleteRows':
				if (!tableSheets.has(op.sheet)) break
				related.push({
					operationIndex,
					op: op.op,
					sheetName: op.sheet,
					at: op.at,
					count: op.count,
					axis: 'row',
				})
				break
			case 'insertCols':
			case 'deleteCols':
				if (!tableSheets.has(op.sheet)) break
				related.push({
					operationIndex,
					op: op.op,
					sheetName: op.sheet,
					at: op.at,
					count: op.count,
					axis: 'column',
				})
				break
		}
	}
	return related
}

function collectLegacyCommentLocations(workbook: Workbook): LegacyCommentLocation[] {
	return workbook.sheets.flatMap((sheet) =>
		[...sheet.comments.entries()].map(([ref, comment]) => ({
			sheetName: sheet.name,
			ref,
			location: `${sheet.name}!${ref}`,
			...(comment.author ? { author: comment.author } : {}),
			hasLegacyDrawing: comment.legacyDrawing !== undefined,
			...(comment.legacyDrawing?.shapeId ? { shapeId: comment.legacyDrawing.shapeId } : {}),
			...(comment.legacyDrawing
				? {
						vmlTarget: {
							...(comment.legacyDrawing.row !== undefined
								? { row: comment.legacyDrawing.row }
								: {}),
							...(comment.legacyDrawing.column !== undefined
								? { column: comment.legacyDrawing.column }
								: {}),
						},
					}
				: {}),
		})),
	)
}

function collectThreadedCommentLocations(workbook: Workbook): ThreadedCommentLocation[] {
	return workbook.sheets.flatMap((sheet) =>
		sheet.threadedComments.map((comment, commentIndex) => ({
			sheetName: sheet.name,
			ref: comment.ref,
			location: `${sheet.name}!${comment.ref}`,
			commentIndex,
			...(comment.partPath ? { partPath: comment.partPath } : {}),
			...(comment.id ? { id: comment.id } : {}),
			...(comment.parentId ? { parentId: comment.parentId } : {}),
			...(comment.personId ? { personId: comment.personId } : {}),
			...(comment.author ? { author: comment.author } : {}),
		})),
	)
}

function collectLegacyCommentPartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	locations: readonly LegacyCommentLocation[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	const hasVmlLayout = locations.some((location) => location.hasLegacyDrawing)
	return uniqueStrings([
		...parts
			.filter(
				(part) =>
					part.featureFamily === 'preservedComments' ||
					(hasVmlLayout && part.featureFamily === 'preservedVml'),
			)
			.map((part) => part.path),
		...issues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isLegacyCommentPackagePartPath),
		),
		...packageIssuePartPaths.filter(isLegacyCommentPackagePartPath),
	])
}

function collectThreadedCommentPartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	locations: readonly ThreadedCommentLocation[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	return uniqueStrings([
		...parts
			.filter((part) => part.featureFamily === 'preservedThreadedComments')
			.map((part) => part.path),
		...locations.flatMap((location) => (location.partPath ? [location.partPath] : [])),
		...issues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isThreadedCommentPackagePartPath),
		),
		...packageIssuePartPaths.filter(isThreadedCommentPackagePartPath),
	])
}

function isLegacyCommentIntegrityIssue(issue: CheckIssue): boolean {
	if (issue.rule === 'legacy-comment-drawing-integrity') return true
	if (issue.rule === 'drawing-integrity' && issue.details?.kind === 'orphan-vml-drawing-part') {
		return checkIssuePackagePartPaths(issue).some(isLegacyCommentPackagePartPath)
	}
	if (issue.rule !== 'package-graph-integrity') return false
	if (checkIssueFeatureFamilies(issue).some(isThreadedCommentFeatureFamily)) return false
	if (checkIssuePackagePartPaths(issue).some(isThreadedCommentPackagePartPath)) return false
	if (checkIssueFeatureFamilies(issue).some(isLegacyCommentFeatureFamily)) return true
	return (
		checkIssuePackagePartPaths(issue).some(isLegacyCommentPackagePartPath) ||
		checkIssueContentTypes(issue).some(isLegacyCommentContentType)
	)
}

function isThreadedCommentIntegrityIssue(issue: CheckIssue): boolean {
	if (issue.rule === 'threaded-comment-integrity') return true
	if (issue.rule !== 'package-graph-integrity') return false
	if (checkIssueFeatureFamilies(issue).some(isThreadedCommentFeatureFamily)) return true
	return (
		checkIssuePackagePartPaths(issue).some(isThreadedCommentPackagePartPath) ||
		checkIssueContentTypes(issue).some(isThreadedCommentContentType)
	)
}

function isExternalLinkIntegrityIssue(issue: CheckIssue): boolean {
	if (issue.rule === 'external-link-integrity') return true
	if (issue.rule !== 'package-graph-integrity') return false
	if (checkIssueFeatureFamilies(issue).some(isExternalLinkFeatureFamily)) return true
	return checkIssuePackagePartPaths(issue).some(isExternalLinkPackagePartPath)
}

function checkIssuePackagePartPaths(issue: CheckIssue): string[] {
	const details = issue.details ?? {}
	const values = [
		...(issue.refs ?? []).filter(isPackagePartPath),
		stringDetail(details, 'partPath'),
		stringDetail(details, 'drawingPartPath'),
		stringDetail(details, 'firstPartPath'),
		stringDetail(details, 'duplicatePartPath'),
		stringDetail(details, 'relationshipPartPath'),
	].filter((value): value is string => value !== undefined && isPackagePartPath(value))
	return uniqueStrings(values)
}

function checkIssueContentTypes(issue: CheckIssue): string[] {
	const details = issue.details ?? {}
	return [
		stringDetail(details, 'contentType'),
		stringDetail(details, 'expected'),
		stringDetail(details, 'actual'),
	].filter((value): value is string => typeof value === 'string' && value.includes('/'))
}

function checkIssueFeatureFamilies(issue: CheckIssue): string[] {
	const featureFamily = stringDetail(issue.details ?? {}, 'featureFamily')
	return featureFamily ? [featureFamily] : []
}

function stringDetail(details: Readonly<Record<string, unknown>>, key: string): string | undefined {
	const value = details[key]
	return typeof value === 'string' ? value : undefined
}

function isPackagePartPath(value: string): boolean {
	return value === '[Content_Types].xml' || /^[^!#]+\/[^!#]+\.[A-Za-z0-9]+$/i.test(value)
}

function isAnalyticalPackagePartPath(path: string): boolean {
	return (
		/^xl\/pivotTables\/[^/]+\.xml$/i.test(path) ||
		/^xl\/pivotCache\/[^/]+\.xml$/i.test(path) ||
		/^xl\/slicerCaches\/[^/]+\.xml$/i.test(path) ||
		/^xl\/slicers\/[^/]+\.xml$/i.test(path) ||
		/^xl\/timelineCaches\/[^/]+\.xml$/i.test(path) ||
		/^xl\/timelines\/[^/]+\.xml$/i.test(path)
	)
}

function isLegacyCommentPackagePartPath(path: string): boolean {
	return /^xl\/comments\d+\.xml$/i.test(path) || /^xl\/drawings\/[^/]+\.vml$/i.test(path)
}

function isThreadedCommentPackagePartPath(path: string): boolean {
	return /^xl\/threadedComments\/[^/]+\.xml$/i.test(path) || /^xl\/persons\/[^/]+\.xml$/i.test(path)
}

function isTablePackagePartPath(path: string): boolean {
	return /^xl\/tables\/[^/]+\.xml$/i.test(path) || /^xl\/queryTables\/[^/]+\.xml$/i.test(path)
}

function isExternalLinkPackagePartPath(path: string): boolean {
	return /^xl\/externalLinks\/[^/]+\.xml$/i.test(path)
}

function isLegacyCommentFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedComments' || featureFamily === 'preservedVml'
}

function isThreadedCommentFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedThreadedComments'
}

function isTableFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedTable' || featureFamily === 'preservedQueryTable'
}

function isExternalLinkFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedExternalLink'
}

function isAnalyticalPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isAnalyticalSidecar(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isAnalyticalPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isAnalyticalContentType)
	)
}

function isTablePackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isTableFeatureFamily(issue.featureFamily)) return true
	return packageGraphIssuePackagePaths(issue).some(isTablePackagePartPath)
}

function isExternalLinkPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isExternalLinkFeatureFamily(issue.featureFamily)) return true
	return packageGraphIssuePackagePaths(issue).some(isExternalLinkPackagePartPath)
}

function isLegacyCommentPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isLegacyCommentFeatureFamily(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isLegacyCommentPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isLegacyCommentContentType)
	)
}

function isThreadedCommentPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isThreadedCommentFeatureFamily(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isThreadedCommentPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isThreadedCommentContentType)
	)
}

function isVisualPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isVisualSidecar(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isVisualPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isVisualContentType)
	)
}

function isRoutedPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (isLegacyCommentPackageGraphIssue(issue) || isThreadedCommentPackageGraphIssue(issue)) {
		return true
	}
	if (isVisualPackageGraphIssue(issue)) return true
	if (isAnalyticalPackageGraphIssue(issue)) return true
	if (!issue.featureFamily) return false
	return (
		isTableFeatureFamily(issue.featureFamily) || isExternalLinkFeatureFamily(issue.featureFamily)
	)
}

function isLegacyCommentContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml' ||
		contentType === 'application/vnd.openxmlformats-officedocument.vmlDrawing'
	)
}

function isThreadedCommentContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.ms-excel.threadedcomments+xml' ||
		contentType === 'application/vnd.ms-excel.person+xml'
	)
}

function isAnalyticalContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml' ||
		contentType ===
			'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml' ||
		contentType ===
			'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml' ||
		contentType === 'application/vnd.ms-excel.slicerCache+xml' ||
		contentType === 'application/vnd.ms-excel.slicer+xml' ||
		contentType === 'application/vnd.ms-excel.timelineCache+xml' ||
		contentType === 'application/vnd.ms-excel.timeline+xml'
	)
}

function isVisualContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.openxmlformats-officedocument.drawing+xml' ||
		contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' ||
		contentType === 'application/vnd.ms-office.chartstyle+xml' ||
		contentType === 'application/vnd.ms-office.chartcolorstyle+xml' ||
		contentType === 'application/vnd.openxmlformats-officedocument.vmlDrawing' ||
		contentType.startsWith('image/')
	)
}

function packageGraphIssuePackagePaths(issue: XlsxPackageGraphFidelityIssue): string[] {
	const values = [
		issue.partPath,
		issue.sourcePartPath,
		issue.relationshipPartPath,
		typeof issue.actual === 'string' ? issue.actual : undefined,
		typeof issue.expected === 'string' ? issue.expected : undefined,
	].filter((value): value is string => value !== undefined && isPackagePartPath(value))
	return uniqueStrings(values)
}

function packageGraphIssueContentTypes(issue: XlsxPackageGraphFidelityIssue): string[] {
	return [
		issue.contentType,
		typeof issue.expected === 'string' ? issue.expected : undefined,
	].filter((value): value is string => typeof value === 'string' && value.includes('/'))
}

function packageGraphIssuesForParts(
	audit: PackageGraphAudit,
	partPaths: readonly string[],
	featureFamilies: readonly string[] = [],
): readonly XlsxPackageGraphFidelityIssue[] {
	if (partPaths.length === 0 && featureFamilies.length === 0) return []
	const pathSet = new Set(partPaths)
	const featureFamilySet = new Set(featureFamilies)
	return audit.issues.filter((issue) => {
		if (issue.partPath !== undefined && pathSet.has(issue.partPath)) return true
		if (issue.sourcePartPath !== undefined && pathSet.has(issue.sourcePartPath)) return true
		if (issue.relationshipPartPath !== undefined && pathSet.has(issue.relationshipPartPath)) {
			return true
		}
		if (typeof issue.actual === 'string' && pathSet.has(issue.actual)) return true
		return issue.featureFamily !== undefined && featureFamilySet.has(issue.featureFamily)
	})
}

function collectLegacyCommentRelatedOperations(
	locations: readonly LegacyCommentLocation[],
	operations: readonly Operation[],
): CommentRelatedOperation[] {
	const related: CommentRelatedOperation[] = []
	for (const [operationIndex, op] of operations.entries()) {
		if (op.op === 'setComment' || op.op === 'deleteComment') {
			related.push({
				operationIndex,
				op: op.op,
				sheetName: op.sheet,
				ref: op.ref,
			})
			continue
		}
		if ((op.op === 'copyRange' || op.op === 'moveRange') && copiesComments(op.mode)) {
			related.push({
				operationIndex,
				op: op.op,
				sheetName: op.sheet,
				source: op.source,
				target: op.target,
				...(op.targetSheet ? { targetSheet: op.targetSheet } : {}),
				...(op.mode ? { mode: op.mode } : {}),
			})
			continue
		}
		const topology = collectCommentTopologyRelatedOperation(locations, op, operationIndex)
		if (topology) related.push(topology)
	}
	return related
}

function collectThreadedCommentRelatedOperations(
	locations: readonly ThreadedCommentLocation[],
	operations: readonly Operation[],
): CommentRelatedOperation[] {
	const related: CommentRelatedOperation[] = []
	for (const [operationIndex, op] of operations.entries()) {
		if (op.op === 'setThreadedComment') {
			related.push({
				operationIndex,
				op: op.op,
				sheetName: op.sheet,
				...(op.ref ? { ref: op.ref } : {}),
				...(op.partPath ? { partPath: op.partPath } : {}),
				...(op.threadedCommentId ? { threadedCommentId: op.threadedCommentId } : {}),
				...(op.commentIndex !== undefined ? { commentIndex: op.commentIndex } : {}),
			})
			continue
		}
		const topology = collectCommentTopologyRelatedOperation(locations, op, operationIndex)
		if (topology) related.push(topology)
	}
	return related
}

function copiesComments(mode: string | undefined): boolean {
	return mode === undefined || mode === 'all' || mode === 'comments'
}

function collectCommentTopologyRelatedOperation(
	locations: readonly (LegacyCommentLocation | ThreadedCommentLocation)[],
	operation: Operation,
	operationIndex: number,
): CommentRelatedOperation | undefined {
	if (!isRowColumnTopologyOperation(operation)) return undefined
	const affectedLocations = locations
		.filter(
			(location) =>
				location.sheetName === operation.sheet &&
				commentLocationAffectedByTopologyOperation(location, operation),
		)
		.map((location) => location.location)
	if (affectedLocations.length === 0) return undefined
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		at: operation.at,
		count: operation.count,
		axis: operation.op === 'insertRows' || operation.op === 'deleteRows' ? 'row' : 'column',
		range: topologyOperationRange(operation),
		affectedLocations: uniqueStrings(affectedLocations),
	}
}

function commentLocationAffectedByTopologyOperation(
	location: LegacyCommentLocation | ThreadedCommentLocation,
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
): boolean {
	let range: RangeRef
	try {
		range = parseRange(location.ref)
	} catch {
		return true
	}
	const axis = operation.op === 'insertRows' || operation.op === 'deleteRows' ? 'row' : 'col'
	return topologyOperationAffectsRange(operation, axis, range)
}

interface X14ConditionalFormatExtensionPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly source: 'x14:conditionalFormatting'
	readonly sqref: string
	readonly index: number
	readonly priority?: number
	readonly type?: string
	readonly preservedAttributeNames: readonly string[]
	readonly preservedChildElements: readonly string[]
}

function collectX14ConditionalFormatExtensionPayloads(
	workbook: Workbook,
): X14ConditionalFormatExtensionPayload[] {
	const payloads: X14ConditionalFormatExtensionPayload[] = []
	for (const sheet of workbook.sheets) {
		for (const format of sheet.x14ConditionalFormats) {
			if (format.deleted) continue
			const preservedAttributeNames = Object.keys(format.preservedRuleAttributes ?? {})
				.filter((name) => !X14_CONDITIONAL_FORMAT_MODELED_ATTRIBUTES.has(name))
				.sort()
			const preservedChildElements = uniqueStrings(
				(format.preservedRuleChildXml ?? []).map(preservedChildElementName),
			).sort()
			if (preservedAttributeNames.length === 0 && preservedChildElements.length === 0) {
				continue
			}
			payloads.push({
				sheetName: sheet.name,
				...(sheet.preservedXml?.partPath ? { sheetPartPath: sheet.preservedXml.partPath } : {}),
				source: 'x14:conditionalFormatting',
				sqref: format.sqref,
				index: format.index,
				...(format.priority !== undefined ? { priority: format.priority } : {}),
				...(format.type ? { type: format.type } : {}),
				preservedAttributeNames,
				preservedChildElements,
			})
		}
	}
	return payloads
}

interface X14DataValidationExtensionPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly source: 'x14:dataValidations'
	readonly sqref: string
	readonly index: number
	readonly type?: string
	readonly operator?: string
	readonly hasFormula1: boolean
	readonly hasFormula2: boolean
	readonly preservedAttributeNames: readonly string[]
	readonly preservedChildElements: readonly string[]
}

function collectX14DataValidationExtensionPayloads(
	workbook: Workbook,
): X14DataValidationExtensionPayload[] {
	const payloads: X14DataValidationExtensionPayload[] = []
	for (const sheet of workbook.sheets) {
		for (const validation of sheet.x14DataValidations) {
			if (validation.deleted) continue
			const preservedAttributeNames = Object.keys(validation.preservedAttributes ?? {})
				.filter((name) => !X14_DATA_VALIDATION_MODELED_ATTRIBUTES.has(name))
				.sort()
			const preservedChildElements = uniqueStrings(
				(validation.preservedChildXml ?? []).map(preservedChildElementName),
			).sort()
			if (preservedAttributeNames.length === 0 && preservedChildElements.length === 0) {
				continue
			}
			payloads.push({
				sheetName: sheet.name,
				...(sheet.preservedXml?.partPath ? { sheetPartPath: sheet.preservedXml.partPath } : {}),
				source: 'x14:dataValidations',
				sqref: validation.sqref,
				index: validation.index,
				...(validation.type ? { type: validation.type } : {}),
				...(validation.operator ? { operator: validation.operator } : {}),
				hasFormula1: validation.formula1 !== undefined,
				hasFormula2: validation.formula2 !== undefined,
				preservedAttributeNames,
				preservedChildElements,
			})
		}
	}
	return payloads
}

const X14_CONDITIONAL_FORMAT_MODELED_ATTRIBUTES = new Set(['type', 'priority', 'id'])

const X14_DATA_VALIDATION_MODELED_ATTRIBUTES = new Set([
	'sqref',
	'type',
	'operator',
	'errorStyle',
	'imeMode',
	'allowBlank',
	'showInputMessage',
	'showErrorMessage',
	'showDropDown',
	'promptTitle',
	'prompt',
	'errorTitle',
	'error',
])

interface X14SemanticEdit {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheet?: string
	readonly sheetName: string
	readonly range?: string
	readonly table?: string
	readonly column?: string | number
	readonly newName?: string
	readonly at?: number
	readonly count?: number
	readonly priority?: number
	readonly ruleIndex?: number
	readonly affectedPayloads: readonly X14AffectedPayload[]
	readonly reason: string
}

interface X14AffectedPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly source:
		| X14ConditionalFormatExtensionPayload['source']
		| X14DataValidationExtensionPayload['source']
	readonly sqref: string
	readonly index: number
	readonly priority?: number
}

function collectX14ConditionalFormatSemanticEdits(
	workbook: Workbook,
	operations: readonly Operation[],
	payloads: readonly X14ConditionalFormatExtensionPayload[],
): X14SemanticEdit[] {
	const edits: X14SemanticEdit[] = []
	operations.forEach((operation, operationIndex) => {
		if (operation.op === 'setConditionalFormat' || operation.op === 'deleteConditionalFormat') {
			const affectedPayloads = payloads.filter((payload) =>
				conditionalFormatOperationAffectsPayload(operation, payload),
			)
			if (affectedPayloads.length === 0) return
			edits.push({
				operationIndex,
				op: operation.op,
				sheet: operation.sheet,
				sheetName: operation.sheet,
				...(operation.range ? { range: operation.range } : {}),
				...(operation.op === 'deleteConditionalFormat' && operation.priority !== undefined
					? { priority: operation.priority }
					: {}),
				...(operation.op === 'deleteConditionalFormat' && operation.ruleIndex !== undefined
					? { ruleIndex: operation.ruleIndex }
					: {}),
				affectedPayloads: affectedPayloads.map(x14AffectedPayload),
				reason:
					'conditional-format operation may change rule order, priority, or range semantics around preserved x14 extension XML',
			})
			return
		}
		const topologyEdit = x14TopologyEdit(operation, operationIndex, payloads)
		if (topologyEdit) edits.push(topologyEdit)
		const tableEdit = x14ConditionalFormatTableReferenceEdit(
			workbook,
			operation,
			operationIndex,
			payloads,
		)
		if (tableEdit) edits.push(tableEdit)
	})
	return edits
}

function collectX14DataValidationSemanticEdits(
	workbook: Workbook,
	operations: readonly Operation[],
	payloads: readonly X14DataValidationExtensionPayload[],
): X14SemanticEdit[] {
	const edits: X14SemanticEdit[] = []
	operations.forEach((operation, operationIndex) => {
		if (operation.op === 'setDataValidation' || operation.op === 'deleteDataValidation') {
			const affectedPayloads = payloads.filter((payload) =>
				dataValidationOperationAffectsPayload(operation, payload),
			)
			if (affectedPayloads.length === 0) return
			edits.push({
				operationIndex,
				op: operation.op,
				sheet: operation.sheet,
				sheetName: operation.sheet,
				range: operation.range,
				affectedPayloads: affectedPayloads.map(x14AffectedPayload),
				reason:
					'data-validation operation overlaps preserved x14 extension XML for the same worksheet range',
			})
			return
		}
		const topologyEdit = x14TopologyEdit(operation, operationIndex, payloads)
		if (topologyEdit) edits.push(topologyEdit)
		const tableEdit = x14DataValidationTableReferenceEdit(
			workbook,
			operation,
			operationIndex,
			payloads,
		)
		if (tableEdit) edits.push(tableEdit)
	})
	return edits
}

function x14TopologyEdit(
	operation: Operation,
	operationIndex: number,
	payloads: readonly (X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload)[],
): X14SemanticEdit | undefined {
	if (!isRowColumnTopologyOperation(operation)) return undefined
	const affectedPayloads = payloads.filter((payload) =>
		topologyOperationAffectsPayloadSqref(operation, payload),
	)
	if (affectedPayloads.length === 0) return undefined
	return {
		operationIndex,
		op: operation.op,
		sheet: operation.sheet,
		sheetName: operation.sheet,
		range: topologyOperationRange(operation),
		at: operation.at,
		count: operation.count,
		affectedPayloads: affectedPayloads.map(x14AffectedPayload),
		reason:
			'row or column topology operation can rewrite shifted x14 sqref metadata or tombstone fully deleted x14 entries',
	}
}

function x14ConditionalFormatTableReferenceEdit(
	workbook: Workbook,
	operation: Operation,
	operationIndex: number,
	payloads: readonly X14ConditionalFormatExtensionPayload[],
): X14SemanticEdit | undefined {
	if (!isTableReferenceRewriteOperation(operation)) return undefined
	const affectedPayloads = payloads.filter((payload) =>
		x14ConditionalFormatPayloadReferencesTable(workbook, payload, operation),
	)
	if (affectedPayloads.length === 0) return undefined
	return x14TableReferenceEdit(operation, operationIndex, affectedPayloads)
}

function x14DataValidationTableReferenceEdit(
	workbook: Workbook,
	operation: Operation,
	operationIndex: number,
	payloads: readonly X14DataValidationExtensionPayload[],
): X14SemanticEdit | undefined {
	if (!isTableReferenceRewriteOperation(operation)) return undefined
	const affectedPayloads = payloads.filter((payload) =>
		x14DataValidationPayloadReferencesTable(workbook, payload, operation),
	)
	if (affectedPayloads.length === 0) return undefined
	return x14TableReferenceEdit(operation, operationIndex, affectedPayloads)
}

function x14TableReferenceEdit(
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
	operationIndex: number,
	affectedPayloads: readonly (
		| X14ConditionalFormatExtensionPayload
		| X14DataValidationExtensionPayload
	)[],
): X14SemanticEdit {
	const sheetName = affectedPayloads[0]?.sheetName ?? ''
	return {
		operationIndex,
		op: operation.op,
		...(sheetName ? { sheet: sheetName, sheetName } : { sheetName }),
		table: operation.table,
		...(operation.op === 'setTableColumn' ? { column: operation.column } : {}),
		...(operation.newName ? { newName: operation.newName } : {}),
		affectedPayloads: affectedPayloads.map(x14AffectedPayload),
		reason:
			'table rename or column rename rewrites structured references in x14 formulas that reference the table',
	}
}

function isRowColumnTopologyOperation(
	operation: Operation,
): operation is Extract<
	Operation,
	{ op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }
> {
	return (
		operation.op === 'insertRows' ||
		operation.op === 'deleteRows' ||
		operation.op === 'insertCols' ||
		operation.op === 'deleteCols'
	)
}

function topologyOperationAffectsPayloadSqref(
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
): boolean {
	if (operation.sheet !== payload.sheetName) return false
	const ranges = parseSqrefRanges(payload.sqref)
	if (ranges.length === 0) return true
	const axis = operation.op === 'insertRows' || operation.op === 'deleteRows' ? 'row' : 'col'
	return ranges.some((range) => topologyOperationAffectsRange(operation, axis, range))
}

function topologyOperationAffectsRange(
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
	axis: 'row' | 'col',
	range: RangeRef,
): boolean {
	const end =
		axis === 'row'
			? Math.max(range.start.row, range.end.row)
			: Math.max(range.start.col, range.end.col)
	const opStart = operation.at
	const opEnd = operation.at + operation.count - 1
	if (operation.op === 'insertRows' || operation.op === 'insertCols') {
		return opStart <= end
	}
	return opStart <= end && opEnd >= 0 && operation.count > 0
}

function topologyOperationRange(
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
): string {
	const start = operation.at
	const end = operation.at + operation.count - 1
	if (operation.op === 'insertRows' || operation.op === 'deleteRows') {
		return `${start + 1}:${end + 1}`
	}
	return `${indexToColumn(start)}:${indexToColumn(end)}`
}

function isTableReferenceRewriteOperation(
	operation: Operation,
): operation is Extract<Operation, { op: 'renameTable' | 'setTableColumn' }> {
	return (
		operation.op === 'renameTable' ||
		(operation.op === 'setTableColumn' && operation.newName !== undefined)
	)
}

function x14ConditionalFormatPayloadReferencesTable(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	return formulaTextsReferenceTable(
		x14ConditionalFormatPayloadFormulaTexts(workbook, payload),
		workbook,
		payload,
		operation,
	)
}

function x14DataValidationPayloadReferencesTable(
	workbook: Workbook,
	payload: X14DataValidationExtensionPayload,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	return formulaTextsReferenceTable(
		x14DataValidationPayloadFormulaTexts(workbook, payload),
		workbook,
		payload,
		operation,
	)
}

function x14ConditionalFormatPayloadFormulaTexts(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload,
): string[] {
	const format = workbook.sheets
		.find((sheet) => sheet.name === payload.sheetName)
		?.x14ConditionalFormats.find((entry) => entry.index === payload.index)
	return uniqueStrings([
		...(format?.formulas ?? []),
		...conditionalFormatValueObjectFormulas(format?.dataBar?.cfvo),
		...conditionalFormatValueObjectFormulas(format?.iconSet?.cfvo),
	])
}

function x14DataValidationPayloadFormulaTexts(
	workbook: Workbook,
	payload: X14DataValidationExtensionPayload,
): string[] {
	const validation = workbook.sheets
		.find((sheet) => sheet.name === payload.sheetName)
		?.x14DataValidations.find((entry) => entry.index === payload.index)
	return [validation?.formula1, validation?.formula2].filter(
		(formula): formula is string => formula !== undefined,
	)
}

function formulaTextsReferenceTable(
	formulas: readonly string[],
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	if (formulas.length === 0) return false
	return formulas.some((formula) =>
		formulaStructuredReferences(formula).some((reference) =>
			structuredReferenceMatchesTableOperation(workbook, payload, reference, operation),
		),
	)
}

function formulaStructuredReferences(
	formula: string,
): Extract<FormulaReferenceInfo, { kind: 'structured' }>[] {
	const parsed = parseFormula(normalizeFormulaInput(formula))
	if (!parsed.ok) return []
	return flattenFormulaReferences(collectFormulaReferences(parsed.value)).filter(
		(reference): reference is Extract<FormulaReferenceInfo, { kind: 'structured' }> =>
			reference.kind === 'structured',
	)
}

function structuredReferenceMatchesTableOperation(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
	reference: Extract<FormulaReferenceInfo, { kind: 'structured' }>,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	const tableName = reference.table || tableNameForLocalStructuredReference(workbook, payload)
	if (!tableName || !tableNameMatchesTableOperation(tableName, operation)) return false
	if (operation.op === 'renameTable') return true
	const columns = [String(operation.column), operation.newName].flatMap((column) =>
		column === undefined ? [] : [column.toLowerCase()],
	)
	if (!reference.column && !reference.endColumn) return true
	return (
		(reference.column !== undefined && columns.includes(reference.column.toLowerCase())) ||
		(reference.endColumn !== undefined && columns.includes(reference.endColumn.toLowerCase()))
	)
}

function tableNameMatchesTableOperation(
	tableName: string,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	const normalized = tableName.toLowerCase()
	if (normalized === operation.table.toLowerCase()) return true
	return operation.op === 'renameTable' && normalized === operation.newName.toLowerCase()
}

function tableNameForLocalStructuredReference(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
): string | undefined {
	const sheet = workbook.sheets.find((entry) => entry.name === payload.sheetName)
	if (!sheet) return undefined
	return sheet.tables.find((table) => rangesMayOverlap(toRangeStringSafe(table.ref), payload.sqref))
		?.name
}

function toRangeStringSafe(range: RangeRef): string {
	return `${indexToColumn(Math.min(range.start.col, range.end.col))}${Math.min(range.start.row, range.end.row) + 1}:${indexToColumn(Math.max(range.start.col, range.end.col))}${Math.max(range.start.row, range.end.row) + 1}`
}

function conditionalFormatOperationAffectsPayload(
	operation: Extract<Operation, { op: 'setConditionalFormat' | 'deleteConditionalFormat' }>,
	payload: X14ConditionalFormatExtensionPayload,
): boolean {
	if (operation.sheet !== payload.sheetName) return false
	if (operation.op === 'deleteConditionalFormat') {
		if (operation.ruleIndex !== undefined) return operation.ruleIndex === payload.index
		if (operation.priority !== undefined) return operation.priority === payload.priority
	}
	if (!operation.range) return true
	return rangesMayOverlap(operation.range, payload.sqref)
}

function dataValidationOperationAffectsPayload(
	operation: Extract<Operation, { op: 'setDataValidation' | 'deleteDataValidation' }>,
	payload: X14DataValidationExtensionPayload,
): boolean {
	if (operation.sheet !== payload.sheetName) return false
	return rangesMayOverlap(operation.range, payload.sqref)
}

function rangesMayOverlap(left: string, right: string): boolean {
	const leftRanges = parseSqrefRanges(left)
	const rightRanges = parseSqrefRanges(right)
	if (leftRanges.length === 0 || rightRanges.length === 0) return true
	return leftRanges.some((leftRange) =>
		rightRanges.some((rightRange) => rangesOverlap(leftRange, rightRange)),
	)
}

function parseSqrefRanges(sqref: string): RangeRef[] {
	const ranges: RangeRef[] = []
	for (const part of sqref.trim().split(/\s+/)) {
		if (!part) continue
		try {
			ranges.push(parseRange(part))
		} catch {
			return []
		}
	}
	return ranges
}

function rangesOverlap(left: RangeRef, right: RangeRef): boolean {
	const leftStartRow = Math.min(left.start.row, left.end.row)
	const leftEndRow = Math.max(left.start.row, left.end.row)
	const leftStartCol = Math.min(left.start.col, left.end.col)
	const leftEndCol = Math.max(left.start.col, left.end.col)
	const rightStartRow = Math.min(right.start.row, right.end.row)
	const rightEndRow = Math.max(right.start.row, right.end.row)
	const rightStartCol = Math.min(right.start.col, right.end.col)
	const rightEndCol = Math.max(right.start.col, right.end.col)
	return (
		leftStartRow <= rightEndRow &&
		leftEndRow >= rightStartRow &&
		leftStartCol <= rightEndCol &&
		leftEndCol >= rightStartCol
	)
}

function x14AffectedPayload(
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
): X14AffectedPayload {
	return {
		sheetName: payload.sheetName,
		...(payload.sheetPartPath ? { sheetPartPath: payload.sheetPartPath } : {}),
		source: payload.source,
		sqref: payload.sqref,
		index: payload.index,
		...('priority' in payload && payload.priority !== undefined
			? { priority: payload.priority }
			: {}),
	}
}

function x14ConditionalFormatSuggestedAction(hasSemanticEditRisk: boolean): string {
	if (hasSemanticEditRisk) {
		return 'Inspect inspectSheet(sheet).x14ConditionalFormats entries by sheetName, sheetPartPath, index, priority, and sqref before committing conditional-format edits; verify postWrite.packageGraphAudit after write.'
	}
	return 'No action is required for unrelated cell edits; keep inspectSheet(sheet).x14ConditionalFormats provenance with the edit record if auditing opaque x14 payload preservation.'
}

function x14DataValidationSuggestedAction(hasSemanticEditRisk: boolean): string {
	if (hasSemanticEditRisk) {
		return 'Inspect inspectSheet(sheet).x14DataValidations entries by sheetName, sheetPartPath, index, and sqref before committing data-validation edits; verify postWrite.packageGraphAudit after write.'
	}
	return 'No action is required for unrelated cell edits; keep inspectSheet(sheet).x14DataValidations provenance with the edit record if auditing opaque x14 payload preservation.'
}

function x14ConditionalFormatRiskMessage(
	payloadCount: number,
	relatedOperations: readonly X14SemanticEdit[],
): string {
	if (
		relatedOperations.every(
			(operation) =>
				operation.op === 'setConditionalFormat' || operation.op === 'deleteConditionalFormat',
		)
	) {
		return `${payloadCount} preserved x14 conditional-format extension payload(s) share a sheet with planned conditional-format semantic edits.`
	}
	return `${payloadCount} preserved x14 conditional-format extension payload(s) intersect planned conditional-format, topology, or table-reference edits.`
}

function x14DataValidationRiskMessage(
	payloadCount: number,
	relatedOperations: readonly X14SemanticEdit[],
): string {
	if (
		relatedOperations.every(
			(operation) =>
				operation.op === 'setDataValidation' || operation.op === 'deleteDataValidation',
		)
	) {
		return `${payloadCount} preserved x14 data-validation extension payload(s) overlap planned data-validation semantic edits.`
	}
	return `${payloadCount} preserved x14 data-validation extension payload(s) intersect planned data-validation, topology, or table-reference edits.`
}

function preservedChildElementName(xml: string): string {
	return /^<\s*([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)/.exec(xml)?.[1] ?? 'unknown'
}

function packagePartDetails(
	packagePartByPath: ReadonlyMap<string, XlsxPackageGraph['parts'][number]>,
	paths: readonly string[],
): WritePolicyPackagePart[] {
	return packagePartDetailsFromParts(
		uniqueStrings(paths)
			.map((path) => packagePartByPath.get(path))
			.filter((part): part is XlsxPackageGraph['parts'][number] => part !== undefined),
	)
}

function packagePartDetailsFromParts(
	parts: readonly XlsxPackageGraph['parts'][number][],
): WritePolicyPackagePart[] {
	return parts.map((part) => ({
		partPath: part.path,
		featureFamily: part.featureFamily,
		preservationPolicy: part.preservationPolicy,
		preservationMode: preservationModeForPackagePolicy(part.preservationPolicy),
		ownerScope: part.ownerScope,
		bytePreservationExpected: part.bytePreservationExpected,
		contentType: part.contentType,
		contentTypeSource: part.contentTypeSource,
		...(part.sourceRelationshipPart ? { sourceRelationshipPart: part.sourceRelationshipPart } : {}),
		...(part.sourceRelationshipId ? { sourceRelationshipId: part.sourceRelationshipId } : {}),
		...(part.sourceRelationshipType ? { sourceRelationshipType: part.sourceRelationshipType } : {}),
		...(part.sourceRelationshipRawType
			? { sourceRelationshipRawType: part.sourceRelationshipRawType }
			: {}),
		...(part.sourceRelationshipRawTarget
			? { sourceRelationshipRawTarget: part.sourceRelationshipRawTarget }
			: {}),
		...(part.sourceRelationshipResolvedTarget
			? { sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget }
			: {}),
		...(part.sourceRelationshipTargetMode
			? { sourceRelationshipTargetMode: part.sourceRelationshipTargetMode }
			: {}),
	}))
}

function preservationModeForPackagePolicy(
	policy: XlsxPackageLossPolicy,
): WritePolicyPreservationMode {
	switch (policy) {
		case 'generated':
			return 'generated'
		case 'preserve-exact':
			return 'preserve-exact'
		case 'discard-on-recalc':
			return 'discarded-for-recalc'
		case 'invalidate-on-edit':
			return 'invalidated-on-edit'
		case 'inspect-only':
			return 'inspect-only'
		case 'unknown-review-required':
			return 'review-required'
	}
}

function isExpectedDiscardedCalcChainCheckIssue(
	issue: CheckIssue,
	skippedCalcChainParts: readonly XlsxPackageGraph['parts'][number][],
): boolean {
	if (issue.rule !== 'package-graph-integrity') return false
	if (issue.details?.code !== 'package_relationship_target') return false
	if (issue.details.featureFamily !== 'preservedCalcChain') return false
	const actual = typeof issue.details.actual === 'string' ? issue.details.actual : undefined
	if (!actual) return false
	return skippedCalcChainParts.some((part) => part.path === actual)
}

function uniqueStrings(values: readonly string[]): string[] {
	return [...new Set(values)]
}

function uniqueNumbers(values: readonly number[]): number[] {
	return [...new Set(values)]
}

function rangeRefToA1(range: RangeRef): string {
	return `${indexToColumn(range.start.col)}${range.start.row + 1}:${indexToColumn(range.end.col)}${range.end.row + 1}`
}

function isActiveContentFeature(featureFamily: string): boolean {
	return (
		featureFamily === 'preservedMacro' ||
		featureFamily === 'preservedActiveX' ||
		featureFamily === 'preservedControl' ||
		featureFamily === 'preservedCustomUi' ||
		featureFamily === 'preservedEmbedding' ||
		featureFamily === 'preservedVendorSecurity'
	)
}

function isVisualSidecar(featureFamily: string): boolean {
	return (
		featureFamily === 'preservedDrawing' ||
		featureFamily === 'preservedChart' ||
		featureFamily === 'preservedChartSheet' ||
		featureFamily === 'preservedChartStyle' ||
		featureFamily === 'preservedChartColor' ||
		featureFamily === 'preservedMedia' ||
		featureFamily === 'preservedVml'
	)
}

function isDrawingMlVisualPath(path: string): boolean {
	return (
		/^xl\/drawings\/[^/]+\.xml$/i.test(path) ||
		/^xl\/drawings\/_rels\/[^/]+\.xml\.rels$/i.test(path) ||
		/^xl\/media\/[^/]+$/i.test(path) ||
		/^xl\/charts\/[^/]+\.xml$/i.test(path) ||
		/^xl\/charts\/_rels\/[^/]+\.xml\.rels$/i.test(path)
	)
}

function isVisualPackagePartPath(path: string): boolean {
	return isDrawingMlVisualPath(path) || /^xl\/drawings\/[^/]+\.vml$/i.test(path)
}

function isAnalyticalSidecar(featureFamily: string): boolean {
	return (
		featureFamily === 'preservedPivot' ||
		featureFamily === 'preservedSlicer' ||
		featureFamily === 'preservedTimeline'
	)
}

function buildBlockedPackageParts(
	packageGraph: XlsxPackageGraph,
	blockedFeatures: readonly FeatureReport[],
): LossAuditPackagePart[] {
	const blockedFamilies = new Set(blockedFeatures.map((feature) => feature.feature.toLowerCase()))
	if (blockedFamilies.size === 0) return []
	return packageGraph.parts
		.filter((part) => blockedFamilies.has(part.featureFamily.toLowerCase()))
		.map((part) => ({
			partPath: part.path,
			featureFamily: part.featureFamily,
			preservationPolicy: part.preservationPolicy,
			preservationMode: preservationModeForPackagePolicy(part.preservationPolicy),
			ownerScope: part.ownerScope,
			bytePreservationExpected: part.bytePreservationExpected,
			contentType: part.contentType,
			contentTypeSource: part.contentTypeSource,
			...(part.sourceRelationshipPart
				? { sourceRelationshipPart: part.sourceRelationshipPart }
				: {}),
			...(part.sourceRelationshipId ? { sourceRelationshipId: part.sourceRelationshipId } : {}),
			...(part.sourceRelationshipType
				? { sourceRelationshipType: part.sourceRelationshipType }
				: {}),
			...(part.sourceRelationshipRawTarget
				? { sourceRelationshipRawTarget: part.sourceRelationshipRawTarget }
				: {}),
			...(part.sourceRelationshipResolvedTarget
				? { sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget }
				: {}),
			...(part.sourceRelationshipTargetMode
				? { sourceRelationshipTargetMode: part.sourceRelationshipTargetMode }
				: {}),
			reason: packagePartLossReason(part.preservationPolicy),
		}))
}

function packagePartLossReason(policy: XlsxPackageLossPolicy): string {
	switch (policy) {
		case 'invalidate-on-edit':
			return 'Generated workbook edits invalidate this package feature unless the workbook is re-signed outside Ascend.'
		case 'unknown-review-required':
			return 'Ascend cannot classify this preserved package part, so an agent must explicitly approve any write around it.'
		case 'preserve-exact':
			return 'This package part is expected to be byte-preserved, but writing around preserved features still requires explicit agent approval.'
		case 'discard-on-recalc':
			return 'This package part may be discarded when formula recalculation invalidates the cached dependency state.'
		case 'inspect-only':
			return 'This package part is inspect-only and cannot be safely regenerated by Ascend.'
		case 'generated':
			return 'This package part is regenerated by Ascend during writes.'
	}
}

function buildApprovalRequirements(
	features: readonly FeatureReport[],
	ops: readonly Operation[],
): ApprovalRequirement[] {
	const approvals: ApprovalRequirement[] = []
	for (const feature of features) {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') continue
		if (isSafePackagePreservationFeature(feature)) continue
		const id = lossFeatureApprovalId(feature)
		approvals.push({
			id,
			kind: 'lossy-write',
			severity: feature.tier === 'unsupported' ? 'critical' : 'high',
			title: `Approve write with ${feature.feature}`,
			reason:
				'Workbook contains preserved or unsupported package features that require explicit approval before writing.',
			feature: feature.feature,
			tier: feature.tier,
			satisfies: [id],
		})
	}
	for (const [operationIndex, op] of ops.entries()) {
		const destructive = destructiveOperationApproval(op, operationIndex)
		if (destructive) approvals.push(destructive)
	}
	return approvals
}

function isSafePackagePreservationFeature(feature: FeatureReport): boolean {
	if (feature.feature === 'preservedDocumentProperties' && feature.tier === 'preserved') return true
	if (feature.feature === 'preservedCalcChain' && feature.tier === 'preserved') return true
	if (feature.feature !== 'preservedOther' || feature.tier !== 'preserved') return false
	if (feature.locations.length === 0) return false
	return feature.locations.every(isSafePackagePreservationLocation)
}

function isSafePackagePreservationLocation(location: string): boolean {
	return (
		location === 'docProps/core.xml' ||
		location === 'docProps/app.xml' ||
		location === 'docProps/custom.xml'
	)
}

function destructiveOperationApproval(
	op: Operation,
	operationIndex: number,
): ApprovalRequirement | null {
	switch (op.op) {
		case 'deleteSheet':
			return destructiveApproval(op, operationIndex, 'critical', `Delete sheet "${op.sheet}"`)
		case 'deleteRows':
			return destructiveApproval(
				op,
				operationIndex,
				'high',
				`Delete ${op.count} row(s) from "${op.sheet}"`,
			)
		case 'deleteCols':
			return destructiveApproval(
				op,
				operationIndex,
				'high',
				`Delete ${op.count} column(s) from "${op.sheet}"`,
			)
		case 'clearRange':
			return op.what === 'all'
				? destructiveApproval(
						op,
						operationIndex,
						'high',
						`Clear all content in ${op.sheet}!${op.range}`,
					)
				: null
		case 'deleteTable':
			return destructiveApproval(op, operationIndex, 'medium', `Delete table "${op.table}"`)
		case 'deleteDefinedName':
			return destructiveApproval(op, operationIndex, 'medium', `Delete defined name "${op.name}"`)
		default:
			return null
	}
}

function destructiveApproval(
	op: Operation,
	operationIndex: number,
	severity: ApprovalRequirement['severity'],
	title: string,
): ApprovalRequirement {
	const opKey = op.op.toLowerCase()
	return {
		id: `op:${operationIndex}:${opKey}`,
		kind: 'destructive-operation',
		severity,
		title,
		reason: 'Operation can remove workbook content or metadata and requires explicit approval.',
		operationIndex,
		operation: op.op,
		satisfies: [`op:${operationIndex}:${opKey}`],
	}
}

function approvalSatisfiedLossFeatures(
	approvals: readonly ApprovalRequirement[],
	granted: readonly string[] | 'all' | undefined,
): readonly string[] {
	if (!granted) return []
	return approvals
		.filter((approval) => approval.kind === 'lossy-write' && isApprovalSatisfied(approval, granted))
		.map((approval) => approval.id)
}

function unsatisfiedApprovalRequirements(
	approvals: readonly ApprovalRequirement[],
	lossAudit: LossAudit,
	granted: readonly string[] | 'all' | undefined,
): ApprovalRequirement[] {
	const blockedLossKeys = new Set(
		lossAudit.blockedFeatures.map((feature) => lossFeatureApprovalId(feature)),
	)
	return approvals.filter((approval) => {
		if (approval.kind === 'lossy-write') {
			return blockedLossKeys.has(approval.id.toLowerCase())
		}
		return !isApprovalSatisfied(approval, granted)
	})
}

function isApprovalSatisfied(
	approval: ApprovalRequirement,
	granted: readonly string[] | 'all' | undefined,
): boolean {
	if (granted === 'all') return true
	const normalized = new Set((granted ?? []).map((entry) => entry.toLowerCase()))
	return normalized.has(approval.id.toLowerCase())
}

function lossFeatureTierKey(feature: FeatureReport): string {
	return `${feature.feature}:${feature.tier}`.toLowerCase()
}

function lossFeatureApprovalId(feature: FeatureReport): string {
	return `loss:${lossFeatureTierKey(feature)}:${shortDigest(feature.locations)}`
}

function mergeAllowLoss(
	base: readonly string[] | 'all' | undefined,
	extra: readonly string[],
): readonly string[] | 'all' {
	if (base === 'all') return 'all'
	return [...new Set([...(base ?? []), ...extra])]
}

function shortDigest(value: unknown): string {
	return sha256Text(stableStringify(value)).slice(0, 10)
}

function phaseCount(
	trace: AgentWorkflowTrace,
	phaseName: string,
	status?: AgentTracePhase['status'],
): number | undefined {
	const phase = trace.phases.find(
		(candidate) => candidate.phase === phaseName && (status ? candidate.status === status : true),
	)
	return phase?.count
}

function traceArtifactCount(trace: AgentWorkflowTrace, name: string): number | undefined {
	const artifact = trace.artifacts.find((entry) => entry.name === name)
	if (!artifact) return undefined
	const match = /^(\d+)/.exec(artifact.summary)
	return match ? Number(match[1]) : undefined
}

interface AtomicWorkbookWriteTimings {
	readonly toBytesMs: number
	readonly writeFileMs: number
	readonly renameMs: number
}

const ACTIVE_CONTENT_TEXT_LOSS_FEATURES = new Set([
	'preservedMacro',
	'preservedActiveX',
	'preservedControl',
	'preservedCustomUi',
	'preservedDrawing',
])

function shouldAllowActiveContentLossForTextExport(
	wb: AscendWorkbook,
	output: string,
	effectiveAllowLoss: readonly string[] | 'all',
): boolean {
	const ext = extname(output).toLowerCase()
	if (ext !== '.csv' && ext !== '.tsv') return false
	if (
		!wb
			.inspect()
			.activeContent.some(
				(entry) => entry.kind !== 'digitalSignature' && entry.kind !== 'vbaSignature',
			)
	) {
		return false
	}
	if (effectiveAllowLoss === 'all') return true
	const allowed = new Set(effectiveAllowLoss.map((entry) => entry.toLowerCase()))
	const activeContentFeatures = wb.report.features.filter(
		(feature) =>
			ACTIVE_CONTENT_TEXT_LOSS_FEATURES.has(feature.feature) &&
			(feature.tier === 'preserved' || feature.tier === 'unsupported'),
	)
	if (activeContentFeatures.length === 0) return false
	return activeContentFeatures.every(
		(feature) =>
			allowed.has(feature.feature.toLowerCase()) ||
			allowed.has(lossFeatureTierKey(feature)) ||
			allowed.has(lossFeatureApprovalId(feature)),
	)
}

async function writeWorkbookAtomically(
	wb: AscendWorkbook,
	output: string,
	options: Pick<AgentCommitOptions, 'allowDecryptedExport'> & {
		readonly allowSignatureInvalidation?: boolean
		readonly allowActiveContentLoss?: boolean
	} = {},
): Promise<AtomicWorkbookWriteTimings> {
	const ext = extname(output)
	const temp = join(dirname(output), `.${Date.now()}.${process.pid}.ascend-tmp${ext}`)
	try {
		if (ext === '.csv' || ext === '.tsv') {
			const save = await timedCommitStep(() =>
				wb.save(temp, {
					...(options.allowDecryptedExport ? { allowDecryptedExport: true } : {}),
					...(options.allowSignatureInvalidation ? { allowSignatureInvalidation: true } : {}),
					...(options.allowActiveContentLoss ? { allowActiveContentLoss: true } : {}),
				}),
			)
			const renameResult = await timedCommitStep(() => rename(temp, output))
			return {
				toBytesMs: 0,
				writeFileMs: save.ms,
				renameMs: renameResult.ms,
			}
		}
		const bytes = await timedCommitStep(() =>
			wb.toBytes({
				...(options.allowDecryptedExport ? { allowDecryptedExport: true } : {}),
				...(options.allowSignatureInvalidation ? { allowSignatureInvalidation: true } : {}),
			}),
		)
		const write = await timedCommitStep(() => writeFile(temp, bytes.value))
		const renameResult = await timedCommitStep(() => rename(temp, output))
		return {
			toBytesMs: bytes.ms,
			writeFileMs: write.ms,
			renameMs: renameResult.ms,
		}
	} catch (error) {
		await unlink(temp).catch(() => undefined)
		throw error
	}
}

function encryptedWorkbookCommitExportError(output: string): AscendException {
	return new AscendException(
		ascendError(
			'EXPORT_ERROR',
			'Cannot export an edited encrypted workbook without re-encryption support.',
			{
				details: {
					output,
					sourceWasEncrypted: true,
					reEncryptionSupported: false,
					requestedExport: 'xlsx',
				},
				suggestedFix:
					'Pass allowDecryptedExport: true only when the caller explicitly accepts a decrypted plain XLSX output, or reopen the original encrypted workbook without editing.',
			},
		),
	)
}

function atomicWorkbookWriteError(output: string, error: unknown): AscendException {
	if (error instanceof AscendException) return error
	const cause = error instanceof Error ? error.message : String(error)
	const causeCode =
		typeof error === 'object' && error !== null && 'code' in error
			? String((error as { readonly code?: unknown }).code)
			: undefined
	return new AscendException(
		ascendError('EXPORT_ERROR', `Failed to write workbook atomically to ${output}`, {
			retryable: true,
			retryStrategy: 'modified',
			details: {
				output,
				operation: 'atomic-workbook-write',
				cause,
				...(causeCode ? { causeCode } : {}),
			},
			suggestedFix:
				'Choose a writable output file path and retry the same plan with the same input hash guard.',
		}),
	)
}

function workbookBackupError(file: string, backup: string, error: unknown): AscendException {
	const cause = error instanceof Error ? error.message : String(error)
	const causeCode =
		typeof error === 'object' && error !== null && 'code' in error
			? String((error as { readonly code?: unknown }).code)
			: undefined
	return new AscendException(
		ascendError('EXPORT_ERROR', `Failed to create in-place backup at ${backup}`, {
			retryable: true,
			retryStrategy: 'modified',
			details: {
				file,
				backup,
				cause,
				...(causeCode ? { causeCode } : {}),
			},
			suggestedFix:
				'Choose a writable backup file path or remove the backup option before retrying the same input hash guard.',
		}),
	)
}

function stableStringify(value: unknown): string {
	const parts: string[] = []
	appendStableStringify(value, parts, false)
	return parts.join('')
}

function appendStableStringify(value: unknown, parts: string[], undefinedAsLiteral: boolean): void {
	if (Array.isArray(value)) {
		parts.push('[')
		for (let index = 0; index < value.length; index++) {
			if (index > 0) parts.push(',')
			appendStableStringify(value[index], parts, false)
		}
		parts.push(']')
		return
	}
	if (value && typeof value === 'object') {
		const entries = Object.entries(value as Record<string, unknown>).sort(([a], [b]) =>
			a.localeCompare(b),
		)
		parts.push('{')
		for (let index = 0; index < entries.length; index++) {
			if (index > 0) parts.push(',')
			const [key, val] = entries[index] as [string, unknown]
			parts.push(JSON.stringify(key), ':')
			appendStableStringify(val, parts, true)
		}
		parts.push('}')
		return
	}
	const encoded = JSON.stringify(value)
	if (encoded !== undefined) parts.push(encoded)
	else if (undefinedAsLiteral) parts.push('undefined')
}
