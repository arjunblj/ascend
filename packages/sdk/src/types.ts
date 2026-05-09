import type {
	AutoFilter,
	CellFormulaBinding,
	ChartPartInfo,
	RangeRef,
	SheetComment,
	SheetConditionalFormat,
	SheetDataValidation,
	SheetDrawingRefs,
	SheetFormatPr,
	SheetHeaderFooter,
	SheetHyperlink,
	SheetIgnoredError,
	SheetImageRef,
	SheetPageMargins,
	SheetPageSetup,
	SheetPrintOptions,
	SheetProtection,
	SheetState,
	SheetTabColor,
	SortState,
	TableColumn,
	TableStyleInfo,
	WorkbookView,
} from '@ascend/core'
import type { CellChange, SheetDiff, WorkbookDiff } from '@ascend/engine'
import type { FormulaNode, Token } from '@ascend/formulas'
import type { AscendError, CellValue, CompatibilityReport, FeatureReport } from '@ascend/schema'

export interface WorkbookInfo {
	readonly sheetCount: number
	readonly loadedSheetCount: number
	readonly sheets: readonly SheetInfo[]
	readonly definedNames: readonly string[]
	readonly definedNameDetails: readonly DefinedNameInfo[]
	readonly cellCount: number | null
	readonly commentCount: number | null
	readonly conditionalFormatCount: number | null
	readonly dataValidationCount: number | null
	readonly imageCount: number | null
	readonly chartCount: number
	readonly pivotTableCount: number
	readonly pivotCacheCount: number
	readonly slicerCount: number
	readonly slicerCacheCount: number
	readonly sourceFormat: string
	readonly workbookViewCount: number
	readonly externalReferenceCount: number
	readonly workbookViews: readonly WorkbookViewInfo[]
	readonly externalReferences: readonly string[]
	readonly externalReferenceDetails: readonly ExternalReferenceInfo[]
	readonly charts: readonly ChartPartInfo[]
	readonly hasWorkbookProtection: boolean
	readonly pivotTables: readonly {
		readonly partPath: string
		readonly sheetName: string
		readonly name?: string
		readonly cacheId?: number
		readonly locationRef?: string
	}[]
	readonly pivotCaches: readonly {
		readonly partPath: string
		readonly cacheId?: number
		readonly relId?: string
		readonly recordCount?: number
		readonly sourceSheet?: string
		readonly sourceRef?: string
		readonly recordsPartPath?: string
	}[]
	readonly slicerCaches: readonly {
		readonly partPath: string
		readonly name?: string
		readonly sourceName?: string
		readonly pivotCacheId?: number
		readonly pivotTableNames: readonly string[]
	}[]
	readonly slicers: readonly {
		readonly partPath: string
		readonly name?: string
		readonly cacheName?: string
		readonly caption?: string
	}[]
	readonly styleSummary: {
		readonly numFmtCount: number
		readonly fontCount: number
		readonly fillCount: number
		readonly borderCount: number
		readonly cellXfCount: number
		readonly dxfCount: number
		readonly tableStyleCount: number
	}
	readonly themeSummary: {
		readonly hasThemePart: boolean
		readonly name?: string
		readonly colorSchemeName?: string
		readonly colorCount: number
		readonly majorFontLatin?: string
		readonly minorFontLatin?: string
	}
	readonly compatibility: CompatibilityReport
	readonly load: WorkbookLoadInfo
}

export interface PivotTableInfo {
	readonly partPath: string
	readonly sheetName: string
	readonly name?: string
	readonly cacheId?: number
	readonly locationRef?: string
}

export interface ExternalReferenceInfo {
	readonly partPath: string
	readonly relId?: string
	readonly target?: string
	readonly targetMode?: string
}

export interface PivotCacheInfo {
	readonly partPath: string
	readonly cacheId?: number
	readonly relId?: string
	readonly recordCount?: number
	readonly refreshedVersion?: number
	readonly minRefreshableVersion?: number
	readonly createdVersion?: number
	readonly refreshedBy?: string
	readonly refreshedDate?: number
	readonly refreshOnLoad?: boolean
	readonly enableRefresh?: boolean
	readonly invalid?: boolean
	readonly saveData?: boolean
	readonly optimizeMemory?: boolean
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly recordsPartPath?: string
}

export interface SlicerCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
}

export interface SlicerInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface SheetInfo {
	readonly name: string
	readonly rowCount: number | null
	readonly colCount: number | null
	readonly cellCount: number | null
	readonly tableCount: number | null
	readonly commentCount: number | null
	readonly conditionalFormatCount: number | null
	readonly dataValidationCount: number | null
	readonly hasFrozenPanes: boolean | null
	readonly colWidthCount: number | null
	readonly imageCount: number | null
	readonly rowHeightCount: number | null
	readonly hyperlinkCount: number | null
	readonly ignoredErrorCount: number | null
	readonly hasAutoFilter: boolean | null
	readonly hasDrawingRefs: boolean | null
	readonly hasPageMetadata: boolean | null
	readonly hasProtection: boolean | null
	readonly cellDataLoaded: boolean
}

export interface SheetCommentInfo extends SheetComment {
	readonly ref: string
}

export interface SheetHyperlinkInfo extends SheetHyperlink {
	readonly ref: string
}

/** Summary of a conditional format rule for SDK inspection. */
export interface ConditionalFormatRuleSummary {
	readonly type: string
	readonly priority?: number
	readonly range: string
}

/** Summary of a data validation for SDK inspection. */
export interface DataValidationSummary {
	readonly type?: string
	readonly formula?: string
	readonly range: string
}

/** Summary of a comment for SDK inspection. */
export interface CommentSummary {
	readonly ref: string
	readonly author?: string
	readonly text: string
}

/** Summary of a hyperlink for SDK inspection. */
export interface HyperlinkSummary {
	readonly ref: string
	readonly target?: string
	readonly display?: string
}

/** Summary of a merge range for SDK inspection. */
export interface MergeRangeSummary {
	readonly range: string
}

/** Formula binding summary for SDK inspection. */
export type FormulaBindingSummary =
	| { readonly kind: 'normal'; readonly formula: string }
	| {
			readonly kind: 'shared-anchor'
			readonly formula: string
			readonly sharedIndex: string
			readonly range?: string
	  }
	| {
			readonly kind: 'shared-member'
			readonly sharedIndex: string
			readonly masterRef?: string
	  }
	| { readonly kind: 'array'; readonly formula: string; readonly range?: string }
	| { readonly kind: 'dynamic-array'; readonly formula: string }
	| { readonly kind: 'spill'; readonly anchorRef: string }

/** Formula cell entry with ref and binding for getFormulaCells. */
export interface FormulaCellEntry {
	readonly ref: string
	readonly binding: FormulaBindingSummary
}

/** Page setup metadata for SDK inspection. */
export interface PageMetadataSummary {
	readonly margins?: SheetPageMargins
	readonly setup?: SheetPageSetup
}

export interface TableInfo {
	readonly name: string
	readonly ref: RangeRef
	readonly rowCount: number
	readonly hasHeaders: boolean
	readonly hasTotals: boolean
	readonly autoFilter: AutoFilter | null
	readonly sortState?: SortState
	readonly styleInfo?: TableStyleInfo
	readonly columnDefs: readonly TableColumn[]
	readonly headerRow?: readonly CellValue[]
	readonly totalsRow?: readonly CellValue[]
}

export interface SheetInspectInfo extends SheetInfo {
	readonly usedRange: RangeRef | null
	readonly state: SheetState
	readonly merges: readonly RangeRef[] | null
	readonly tables: readonly TableInfo[] | null
	readonly comments: readonly SheetCommentInfo[] | null
	readonly hyperlinks: readonly SheetHyperlinkInfo[] | null
	readonly ignoredErrors: readonly SheetIgnoredError[] | null
	readonly conditionalFormats: readonly SheetConditionalFormat[] | null
	readonly dataValidations: readonly SheetDataValidation[] | null
	readonly imageRefs: readonly SheetImageRef[] | null
	readonly drawingRefs: SheetDrawingRefs | null
	readonly autoFilter: AutoFilter | null
	readonly protection: SheetProtection | null
	readonly tabColor: SheetTabColor | null
	readonly sheetFormatPr: SheetFormatPr | null
	readonly pageMargins: SheetPageMargins | null
	readonly pageSetup: SheetPageSetup | null
	readonly printOptions: SheetPrintOptions | null
	readonly headerFooter: SheetHeaderFooter | null
}

export type VisualFeatureCategory = 'chart' | 'drawing' | 'image' | 'shape-or-control'

export interface VisualPackageFeatureInfo extends FeatureReport {
	readonly category: VisualFeatureCategory
}

export interface SheetVisualInventoryInfo {
	readonly sheet: string
	readonly drawingRefs: SheetDrawingRefs | null
	readonly hasDrawing: boolean | null
	readonly hasLegacyDrawing: boolean | null
	readonly imageRefs: readonly SheetImageRef[] | null
	readonly imageCount: number | null
}

export interface WorkbookVisualInventoryInfo {
	readonly load: WorkbookLoadInfo
	readonly packageFeatures: readonly VisualPackageFeatureInfo[]
	readonly sheets: readonly SheetVisualInventoryInfo[]
	readonly sheetImageCount: number | null
	readonly charts: readonly ChartPartInfo[]
	readonly structuredChartCount: number
	readonly packageChartFeatureCount: number
	readonly packageDrawingFeatureCount: number
	readonly packageMediaFeatureCount: number
	readonly hasPreservedCharts: boolean
	readonly hasPreservedDrawings: boolean
	readonly hasPreservedMedia: boolean
	readonly notes: readonly string[]
}

export interface WorkbookLoadInfo {
	readonly mode: 'full' | 'metadata-only' | 'values' | 'formula' | 'selective'
	readonly isPartial: boolean
	readonly cellsHydrated: boolean
	readonly richSheetMetadataHydrated: boolean
	readonly hasAllSheets: boolean
	readonly sourceSheets: readonly string[]
	readonly loadedSheets: readonly string[]
}

export interface DefinedNameInfo {
	readonly name: string
	readonly formula: string
	readonly normalizedFormula: string
	readonly scope: 'workbook' | 'sheet'
	readonly sheet?: string
	readonly references: readonly FormulaReferenceInfo[]
	readonly refs: readonly string[]
	readonly functions: readonly string[]
	readonly volatile: boolean
	readonly parseError?: string
}

export interface WorkbookViewInfo extends WorkbookView {}

export interface CellInfo {
	readonly ref: string
	readonly value: CellValue
	readonly formula: string | null
	readonly formulaBinding?: CellFormulaBinding
	readonly row: number
	readonly col: number
}

export type FlatCellValue = number | string | boolean | null

export interface CompactCellInfo {
	readonly ref?: string
	readonly value: CellValue
	readonly formula: string | null
	readonly formulaBinding: CellFormulaBinding | null
	readonly row: number
	readonly col: number
}

export interface AgentReadOptions {
	readonly rowOffset?: number
	readonly rowLimit?: number
	readonly includeRefs?: boolean
	readonly omitEmpty?: boolean
	readonly flatValues?: boolean
	readonly changedSince?: string
}

export interface AgentViewOptions {
	readonly rowChunkSize?: number
	readonly sampleRowLimit?: number
	readonly sampleValueLimit?: number
}

export interface AgentFormulaPatternInfo {
	readonly pattern: string
	readonly count: number
}

export interface AgentColumnSummary {
	readonly col: number
	readonly ref: string
	readonly header: FlatCellValue | null
	readonly kind: 'empty' | 'number' | 'string' | 'boolean' | 'formula' | 'mixed'
	readonly nonEmptyCount: number
	readonly formulaCount: number
	readonly sampleValues: readonly FlatCellValue[]
}

export interface AgentSampleRow {
	readonly row: number
	readonly cells: readonly CompactCellInfo[]
}

export interface AgentViewResult {
	readonly sheet: string
	readonly range: RangeRef
	readonly rowCount: number
	readonly colCount: number
	readonly nonEmptyCount: number
	readonly formulaCount: number
	readonly distinctFunctions: readonly string[]
	readonly formulaPatterns: readonly AgentFormulaPatternInfo[]
	readonly columns: readonly AgentColumnSummary[]
	readonly samples: readonly AgentSampleRow[]
	readonly notes: readonly string[]
}

export interface RangeInfo {
	readonly ref: RangeRef
	readonly cells: readonly CellInfo[]
	readonly rowCount: number
	readonly colCount: number
}

export interface CompactRangeInfo {
	readonly ref: RangeRef
	readonly cells: readonly CompactCellInfo[]
	readonly rowCount: number
	readonly colCount: number
}

export interface RangeWindowInfo extends RangeInfo {
	readonly requestedRef: RangeRef
	readonly rowOffset: number
	readonly rowLimit: number
	readonly hasMore: boolean
	readonly nextRowOffset?: number
}

export interface CompactRangeWindowInfo extends CompactRangeInfo {
	readonly requestedRef: RangeRef
	readonly rowOffset: number
	readonly rowLimit: number
	readonly hasMore: boolean
	readonly nextRowOffset?: number
	readonly changeToken?: string
}

export interface RangeRowsInfo {
	readonly requestedRef: RangeRef
	readonly ref: RangeRef
	readonly rowCount: number
	readonly colCount: number
	readonly rowOffset: number
	readonly rowLimit: number
	readonly hasMore: boolean
	readonly nextRowOffset?: number
	readonly rows: readonly (readonly CellValue[])[]
}

export interface RangeObjectsInfo {
	readonly requestedRef: RangeRef
	readonly ref: RangeRef
	readonly rowCount: number
	readonly colCount: number
	readonly rowOffset: number
	readonly rowLimit: number
	readonly hasMore: boolean
	readonly nextRowOffset?: number
	readonly headers: readonly string[]
	readonly rows: readonly Readonly<Record<string, CellValue>>[]
}

export interface TableRowInfo {
	readonly index: number
	readonly sheetRow: number
	readonly values: Readonly<Record<string, CellValue>>
}

export interface TableWindowInfo {
	readonly rowOffset: number
	readonly rowLimit: number
	readonly returnedRows: number
	readonly totalRows: number
	readonly hasMore: boolean
	readonly nextRowOffset?: number
	readonly rows: readonly TableRowInfo[]
}

export interface ChangedCell {
	readonly ref: string
	readonly oldValue: CellValue
	readonly newValue: CellValue
}

export interface PreviewResult {
	readonly diff: WorkbookDiff
	readonly sheetDiffs: readonly SheetDiff[]
	readonly cellChanges: readonly CellChange[]
	/** Agent-friendly list of changed cells with full refs (e.g. Sheet1!A1). */
	readonly changedCells: readonly ChangedCell[]
	/** Number of cells that would be recalculated (formula dependents). */
	readonly recalcScope: number
	/** Validation warnings from apply (e.g. data validation violations). */
	readonly warnings: readonly AscendError[]
	/** True if apply + recalc would complete without errors. */
	readonly wouldSucceed: boolean
	readonly errors: readonly AscendError[]
	readonly writePlan?: WritePlanInfo
}

export interface ApplyResult {
	readonly affectedCells: readonly string[]
	readonly sheetsModified: readonly string[]
	readonly recalcRequired: boolean
	readonly errors: readonly AscendError[]
	readonly warnings?: readonly AscendError[]
}

export interface ApplyAndRecalcResult {
	readonly apply: ApplyResult
	readonly recalc: RecalcResult | null
}

export interface BatchResult {
	readonly errors: readonly AscendError[]
}

export interface RecalcResult {
	readonly changed: readonly string[]
	readonly errors: ReadonlyArray<{ ref: string; error: AscendError }>
	readonly duration: number
}

export interface WritePlanInfo {
	readonly totalParts: number
	readonly byOrigin: Readonly<{
		generated: number
		'preserved-inline': number
		'preserved-source': number
		capsule: number
	}>
	readonly byOwnerKind: Readonly<{
		package: number
		workbook: number
		sheet: number
	}>
	readonly sheetPartCounts: Readonly<Record<string, number>>
}

export interface CheckResult {
	readonly valid: boolean
	readonly issues: readonly CheckIssue[]
}

export interface CheckIssue {
	readonly severity: 'error' | 'warning' | 'info'
	readonly message: string
	readonly ref?: string
}

export interface LintResult {
	readonly clean: boolean
	readonly warnings: readonly LintWarning[]
}

export interface LintWarning {
	readonly rule: string
	readonly message: string
	readonly ref?: string
}

export interface TraceResult {
	readonly ref: string
	readonly formula: string | null
	readonly value: CellValue
	readonly precedents: readonly TraceNodeInfo[]
	readonly dependents: readonly TraceNodeInfo[]
	readonly dependsOn: readonly string[]
	readonly feedsInto: readonly string[]
}

export interface TraceNodeInfo {
	readonly ref: string
	readonly formula: string | null
	readonly value: CellValue
	readonly depth: number
}

export type FormulaReferenceScope =
	| { readonly kind: 'local' }
	| { readonly kind: 'sheet'; readonly sheet: string }
	| { readonly kind: 'sheetSpan'; readonly startSheet: string; readonly endSheet: string }
	| { readonly kind: 'external'; readonly workbook: string; readonly sheet: string }

interface FormulaReferenceBase {
	readonly text: string
	readonly scope?: FormulaReferenceScope
}

export interface CellFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'cell'
}

export interface RangeFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'range'
}

export interface WholeRowFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'wholeRow'
}

export interface WholeColumnFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'wholeColumn'
}

export interface NameFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'name'
}

export interface StructuredFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'structured'
	readonly table: string
	readonly specifiers: readonly string[]
	readonly column?: string
}

export interface SpillFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'spill'
	readonly targetText: string
	readonly target?: FormulaReferenceInfo
}

export interface ImplicitIntersectionFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'implicitIntersection'
	readonly targetText: string
	readonly target?: FormulaReferenceInfo
}

export interface CompoundFormulaReferenceInfo extends FormulaReferenceBase {
	readonly kind: 'union' | 'intersection'
	readonly members: readonly FormulaReferenceInfo[]
}

export type FormulaReferenceInfo =
	| CellFormulaReferenceInfo
	| RangeFormulaReferenceInfo
	| WholeRowFormulaReferenceInfo
	| WholeColumnFormulaReferenceInfo
	| NameFormulaReferenceInfo
	| StructuredFormulaReferenceInfo
	| SpillFormulaReferenceInfo
	| ImplicitIntersectionFormulaReferenceInfo
	| CompoundFormulaReferenceInfo

export interface FormulaInfo {
	readonly ref: string
	readonly formula: string
	readonly normalizedFormula: string
	readonly value: CellValue
	readonly binding?: CellFormulaBinding
	readonly references: readonly FormulaReferenceInfo[]
	readonly refs: readonly string[]
	readonly functions: readonly string[]
	readonly volatile: boolean
	readonly tokens: readonly Token[]
	readonly ast?: FormulaNode
	readonly parseError?: string
}
