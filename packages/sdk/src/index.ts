/** Convert a zero-based column index to an Excel column letter (`0 → "A"`, `25 → "Z"`, `26 → "AA"`). */
export { indexToColumn, parseA1 } from '@ascend/core'
export type { AgentDocEntry, AgentDocKind, AgentDocSearchResult } from './agent-docs.ts'
export { loadAgentDocs, readAgentDoc, searchAgentDocs } from './agent-docs.ts'
export type {
	AgentCommitOptions,
	AgentCommitResult,
	AgentModelOutput,
	AgentPlanResult,
	AgentPostWriteVerification,
	AgentTraceArtifact,
	AgentTracePhase,
	AgentWorkflowProgressEvent,
	AgentWorkflowProgressHandler,
	AgentWorkflowTrace,
	ApprovalRequirement,
	LossAudit,
	RepairAction,
	RepairPlanResult,
} from './agent-workflow.ts'
export {
	auditLossPolicy,
	commitAgentPlan,
	createAgentPlan,
	createRepairPlan,
	digestPlan,
	sha256Bytes,
} from './agent-workflow.ts'
export type {
	CapabilityFilters,
	CapabilityPriority,
	CapabilityStatus,
	CapabilitySummary,
	CapabilitySurface,
	ExcelCapability,
} from './capabilities.ts'
export {
	EXCEL_CAPABILITIES,
	getCapability,
	isCapabilityGap,
	listCapabilities,
	summarizeCapabilities,
} from './capabilities.ts'
export type { ExportFormat, FormatDisplayOptions } from './format-helpers.ts'
/** Helpers for formatting cell values and inferring export formats. */
export {
	ensureOutputExtension,
	escapeDelimitedCell,
	formatDisplayCellValue,
	inferExportFormat,
	normalizeExportFormat,
	toA1Ref,
} from './format-helpers.ts'
export type {
	OperationApprovalMetadata,
	OperationInvalidExample,
	OperationJsonSchema,
	OperationSchema,
	ParseOperationsResult,
} from './ops.ts'
/** Typed operation builders and introspection. */
export * as ops from './ops.ts'
/** List all supported patch operations or generate JSON-schema-style descriptors for agents. */
export { getOperationsSchema, listOperations, parseOperations } from './ops.ts'
/** Base class for read-only workbook views (inspect, read, formula, check, lint, trace, diff). */
export { WorkbookReadView } from './read-view.ts'
export type {
	/** A cell selector: `"Sheet1!A1"`, `{ sheet: "Sheet1", ref: "A1" }`, or `{ sheet: "Sheet1", row: 0, col: 0 }`. */
	CellSelector,
	CellSelectorObject,
	/** A range selector: `"Sheet1!A1:B10"` or `{ sheet: "Sheet1", range: "A1:B10" }`. */
	RangeSelector,
	RangeSelectorObject,
} from './ref-selectors.ts'
export type {
	SessionCacheOptions,
	WorkbookLoadOptions,
	WorkbookSessionOpenOptions,
} from './session.ts'
/** Cache and session management for multi-request read-only workbook access. */
export { configureSessionCache, WorkbookDocument, WorkbookSession } from './session.ts'
/** Fluent handle for reading/writing a single sheet inside an `AscendWorkbook`. */
export { SheetHandle } from './sheet-handle.ts'
export type { WorkbookRowStreamSource } from './stream.ts'
/** Stream workbook rows one chunk at a time for large-file processing. */
export { streamWorkbookRows } from './stream.ts'
/** Handle for structured table operations (read rows/columns, append). */
export { TableHandle } from './table-handle.ts'
export type {
	ActiveContentInfo,
	ActiveContentKind,
	AgentColumnSummary,
	AgentFormulaPatternInfo,
	AgentReadOptions,
	AgentSampleRow,
	AgentViewOptions,
	AgentViewResult,
	ApplyAndRecalcResult,
	ApplyResult,
	BatchResult,
	CellFormulaReferenceInfo,
	CellInfo,
	ChangedCell,
	CheckIssue,
	CheckResult,
	CommentSummary,
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	CompoundFormulaReferenceInfo,
	ConditionalFormatRuleSummary,
	DataValidationSummary,
	DefinedNameInfo,
	FlatCellValue,
	FormulaBindingSummary,
	FormulaCellEntry,
	FormulaInfo,
	FormulaReferenceInfo,
	FormulaReferenceScope,
	GetPivotDataFilter,
	GetPivotDataMatchInfo,
	GetPivotDataQuery,
	GetPivotDataResult,
	HyperlinkSummary,
	ImplicitIntersectionFormulaReferenceInfo,
	LintResult,
	LintWarning,
	MergeRangeSummary,
	NameFormulaReferenceInfo,
	PageMetadataSummary,
	PivotAreaInfo,
	PivotAreaReferenceInfo,
	PivotAxisFieldItemInfo,
	PivotAxisItemInfo,
	PivotCacheFieldGroupDiscreteItemInfo,
	PivotCacheFieldGroupInfo,
	PivotCacheFieldGroupRangeInfo,
	PivotCacheFieldInfo,
	PivotCacheInfo,
	PivotCacheRecordInfo,
	PivotCacheRecordsInfo,
	PivotCacheRecordValueInfo,
	PivotCacheRecordValueKind,
	PivotCacheRecordValueKindCount,
	PivotCacheSharedItemsInfo,
	PivotChartFormatInfo,
	PivotDataFieldInfo,
	PivotFieldInfo,
	PivotFieldReference,
	PivotFormatInfo,
	PivotRefreshCacheRecordsInfo,
	PivotRefreshOutputState,
	PivotRefreshPlanInfo,
	PivotRefreshRecommendedOp,
	PivotRefreshTableInfo,
	PivotTableInfo,
	PivotTableLocationInfo,
	PivotTableOptionsInfo,
	PivotTableStyleInfo,
	PreviewResult,
	RangeFormulaReferenceInfo,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	RecalcResult,
	SheetCommentInfo,
	SheetHyperlinkInfo,
	SheetInfo,
	SheetInspectInfo,
	SlicerCacheInfo,
	SlicerInfo,
	SpillFormulaReferenceInfo,
	StructuredFormulaReferenceInfo,
	TableInfo,
	TableRowInfo,
	TableWindowInfo,
	TimelineCacheInfo,
	TimelineInfo,
	TimelineRangeInfo,
	TimelineStateInfo,
	TraceNodeInfo,
	TraceResult,
	VisualFeatureCategory,
	VisualPackageFeatureInfo,
	WholeColumnFormulaReferenceInfo,
	WholeRowFormulaReferenceInfo,
	WorkbookInfo,
	WorkbookLoadInfo,
	WorkbookViewInfo,
	WorkbookVisualInventoryInfo,
	WritePlanInfo,
} from './types.ts'
/**
 * `Ascend` — convenience entry point: `Ascend.open()`, `Ascend.create()`, `Ascend.fromCsv()`.
 *
 * `AscendWorkbook` — full mutable workbook for apply/recalc/save workflows.
 *
 * `BatchBuilder` — fluent builder for batching cell/formula/structural operations.
 */
export { Ascend, AscendWorkbook, BatchBuilder } from './workbook.ts'
