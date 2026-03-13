export type { SessionCacheOptions } from './session.ts'
export { configureSessionCache, WorkbookDocument } from './session.ts'
export { SheetHandle } from './sheet-handle.ts'
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
	CellInfo,
	CheckIssue,
	CheckResult,
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	FlatCellValue,
	FormulaInfo,
	LintResult,
	LintWarning,
	PreviewResult,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	RecalcResult,
	SheetInfo,
	SheetInspectInfo,
	TableInfo,
	TraceResult,
	WorkbookInfo,
	WorkbookLoadInfo,
} from './types.ts'
export { AscendWorkbook, AscendWorkbook as Ascend } from './workbook.ts'
