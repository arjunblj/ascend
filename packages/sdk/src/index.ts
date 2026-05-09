export { indexToColumn, parseA1 } from '@ascend/core'
export type {
	AgentCommitOptions,
	AgentCommitResult,
	AgentPlanResult,
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
export {
	ensureOutputExtension,
	escapeDelimitedCell,
	formatDisplayCellValue,
	inferExportFormat,
	normalizeExportFormat,
	toA1Ref,
} from './format-helpers.ts'
export type { OperationJsonSchema, OperationSchema, ParseOperationsResult } from './ops.ts'
export * as ops from './ops.ts'
export { getOperationsSchema, listOperations, parseOperations } from './ops.ts'
export { WorkbookReadView } from './read-view.ts'
export type {
	CellSelector,
	CellSelectorObject,
	RangeSelector,
	RangeSelectorObject,
} from './ref-selectors.ts'
export type {
	SessionCacheOptions,
	WorkbookLoadOptions,
	WorkbookSessionOpenOptions,
} from './session.ts'
export { configureSessionCache, WorkbookDocument, WorkbookSession } from './session.ts'
export { SheetHandle } from './sheet-handle.ts'
export type { WorkbookRowStreamSource } from './stream.ts'
export { streamWorkbookRows } from './stream.ts'
export { TableHandle } from './table-handle.ts'
export type {
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
	HyperlinkSummary,
	ImplicitIntersectionFormulaReferenceInfo,
	LintResult,
	LintWarning,
	MergeRangeSummary,
	NameFormulaReferenceInfo,
	PageMetadataSummary,
	PivotCacheInfo,
	PivotTableInfo,
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
	TraceNodeInfo,
	TraceResult,
	WholeColumnFormulaReferenceInfo,
	WholeRowFormulaReferenceInfo,
	WorkbookInfo,
	WorkbookLoadInfo,
	WorkbookViewInfo,
	WritePlanInfo,
} from './types.ts'
export { Ascend, AscendWorkbook, BatchBuilder } from './workbook.ts'
