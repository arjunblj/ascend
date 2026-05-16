import { expect, test } from 'bun:test'
import type {
	AgentCommitTimings,
	AgentPostWriteVerification,
	AgentPostWriteVerificationTimings,
	AgentWorkflowProofChangedCell,
	AgentWorkflowProofGate,
	AgentWorkflowProofSummary,
	PostWriteActiveContentEntry,
	PostWriteActiveContentSummary,
	PostWriteAnalyticsSummary,
	PostWriteChartEntry,
	PostWriteChartSeriesEntry,
	PostWriteCommentSummary,
	PostWriteDataConnectionEntry,
	PostWriteDataConnectionSummary,
	PostWriteDataValidationEntry,
	PostWriteDataValidationSummary,
	PostWriteDefinedNameEntry,
	PostWriteDefinedNameSummary,
	PostWriteExternalReferenceEntry,
	PostWriteExternalReferenceSummary,
	PostWriteFormulaCacheValueKindCount,
	PostWriteFormulaSummary,
	PostWriteOpaquePayloadSummary,
	PostWritePivotCacheEntry,
	PostWritePivotTableEntry,
	PostWriteProtectedRangeSecurityEntry,
	PostWriteSecuritySummary,
	PostWriteSheetSecurityEntry,
	PostWriteSheetTopologyEntry,
	PostWriteSheetVisualEntry,
	PostWriteSlicerCacheEntry,
	PostWriteTableSummary,
	PostWriteTimelineCacheEntry,
	PostWriteVisualSummary,
	PostWriteWorkbookTopologySummary,
	WritePolicyPreservationMode,
	WritePolicyPreservationModeSummary,
} from './index.ts'
import { createAgentWorkflowProofSummary } from './index.ts'

type CommitProofExportSurface = {
	readonly activeContentEntry?: PostWriteActiveContentEntry
	readonly activeContentSummary?: PostWriteActiveContentSummary
	readonly analytics?: PostWriteAnalyticsSummary
	readonly chart?: PostWriteChartEntry
	readonly chartSeries?: PostWriteChartSeriesEntry
	readonly comment?: PostWriteCommentSummary
	readonly proofChangedCell?: AgentWorkflowProofChangedCell
	readonly proofGate?: AgentWorkflowProofGate
	readonly proofSummary?: AgentWorkflowProofSummary
	readonly dataConnection?: PostWriteDataConnectionEntry
	readonly dataConnections?: PostWriteDataConnectionSummary
	readonly dataValidation?: PostWriteDataValidationEntry
	readonly dataValidations?: PostWriteDataValidationSummary
	readonly commitTimings?: AgentCommitTimings
	readonly definedName?: PostWriteDefinedNameEntry
	readonly definedNames?: PostWriteDefinedNameSummary
	readonly externalReference?: PostWriteExternalReferenceEntry
	readonly externalReferences?: PostWriteExternalReferenceSummary
	readonly formulaCacheValueKind?: PostWriteFormulaCacheValueKindCount
	readonly formula?: PostWriteFormulaSummary
	readonly opaquePayloads?: PostWriteOpaquePayloadSummary
	readonly pivotCache?: PostWritePivotCacheEntry
	readonly pivotTable?: PostWritePivotTableEntry
	readonly protectedRangeSecurity?: PostWriteProtectedRangeSecurityEntry
	readonly security?: PostWriteSecuritySummary
	readonly sheetSecurity?: PostWriteSheetSecurityEntry
	readonly sheetTopology?: PostWriteSheetTopologyEntry
	readonly sheetVisual?: PostWriteSheetVisualEntry
	readonly slicerCache?: PostWriteSlicerCacheEntry
	readonly table?: PostWriteTableSummary
	readonly timelineCache?: PostWriteTimelineCacheEntry
	readonly verification?: AgentPostWriteVerification
	readonly verificationTimings?: AgentPostWriteVerificationTimings
	readonly visual?: PostWriteVisualSummary
	readonly workbookTopology?: PostWriteWorkbookTopologySummary
	readonly writePolicyMode?: WritePolicyPreservationMode
	readonly writePolicyModes?: WritePolicyPreservationModeSummary
}

test('root SDK exports commit proof audit types for release consumers', () => {
	const exportNames: ReadonlyArray<keyof CommitProofExportSurface> = [
		'activeContentEntry',
		'activeContentSummary',
		'analytics',
		'chart',
		'chartSeries',
		'comment',
		'proofChangedCell',
		'proofGate',
		'proofSummary',
		'dataConnection',
		'dataConnections',
		'dataValidation',
		'dataValidations',
		'commitTimings',
		'definedName',
		'definedNames',
		'externalReference',
		'externalReferences',
		'formulaCacheValueKind',
		'formula',
		'opaquePayloads',
		'pivotCache',
		'pivotTable',
		'protectedRangeSecurity',
		'security',
		'sheetSecurity',
		'sheetTopology',
		'sheetVisual',
		'slicerCache',
		'table',
		'timelineCache',
		'verification',
		'verificationTimings',
		'visual',
		'workbookTopology',
		'writePolicyMode',
		'writePolicyModes',
	]

	expect(exportNames).toContain('verification')
	expect(exportNames).toContain('writePolicyModes')
	expect(typeof createAgentWorkflowProofSummary).toBe('function')
})
