import { expect, test } from 'bun:test'
import type {
	AgentCommitTimings,
	AgentPostWriteVerification,
	AgentPostWriteVerificationTimings,
	PostWriteActiveContentEntry,
	PostWriteActiveContentSummary,
	PostWriteAnalyticsSummary,
	PostWriteChartEntry,
	PostWriteCommentSummary,
	PostWriteDefinedNameEntry,
	PostWriteDefinedNameSummary,
	PostWriteExternalReferenceEntry,
	PostWriteExternalReferenceSummary,
	PostWriteFormulaSummary,
	PostWriteOpaquePayloadSummary,
	PostWritePivotCacheEntry,
	PostWritePivotTableEntry,
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

type CommitProofExportSurface = {
	readonly activeContentEntry?: PostWriteActiveContentEntry
	readonly activeContentSummary?: PostWriteActiveContentSummary
	readonly analytics?: PostWriteAnalyticsSummary
	readonly chart?: PostWriteChartEntry
	readonly comment?: PostWriteCommentSummary
	readonly commitTimings?: AgentCommitTimings
	readonly definedName?: PostWriteDefinedNameEntry
	readonly definedNames?: PostWriteDefinedNameSummary
	readonly externalReference?: PostWriteExternalReferenceEntry
	readonly externalReferences?: PostWriteExternalReferenceSummary
	readonly formula?: PostWriteFormulaSummary
	readonly opaquePayloads?: PostWriteOpaquePayloadSummary
	readonly pivotCache?: PostWritePivotCacheEntry
	readonly pivotTable?: PostWritePivotTableEntry
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
		'comment',
		'commitTimings',
		'definedName',
		'definedNames',
		'externalReference',
		'externalReferences',
		'formula',
		'opaquePayloads',
		'pivotCache',
		'pivotTable',
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
})
