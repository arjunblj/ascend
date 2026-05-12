import type {
	ActiveContentInfo,
	AutoFilter,
	CellFormulaBinding,
	ChartPartInfo,
	ChartSheetInfo,
	RangeRef,
	SheetAdvancedFilterInfo,
	SheetComment,
	SheetConditionalFormat,
	SheetDataValidation,
	SheetDrawingObjectRef,
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
	SheetSparklineGroupInfo,
	SheetState,
	SheetTabColor,
	SheetThreadedComment,
	SheetX14ConditionalFormatInfo,
	SheetX14DataValidationInfo,
	SortState,
	TableColumn,
	TableStyleInfo,
	WorkbookConnectionPartInfo,
	WorkbookDataModelPartInfo,
	WorkbookDocumentProperties,
	WorkbookMacroSheetInfo,
	WorkbookThemeColor,
	WorkbookView,
} from '@ascend/core'

export type {
	ActiveContentInfo,
	ActiveContentKind,
	VbaModuleInfo,
	VbaModuleKind,
	VbaProjectInfo,
	WorkbookConnectionPartInfo,
	WorkbookConnectionPartKind,
	WorkbookCoreDocumentProperties,
	WorkbookCustomDocumentProperty,
	WorkbookDataModelPartInfo,
	WorkbookDataModelPartKind,
	WorkbookDocumentProperties,
	WorkbookMacroSheetInfo,
} from '@ascend/core'

import type { CellChange, ExternalReferenceResolver, SheetDiff, WorkbookDiff } from '@ascend/engine'
import type { FormulaNode, Token } from '@ascend/formulas'
import type {
	AscendError,
	CellValue,
	CompatibilityReport,
	FeatureReport,
	Operation,
} from '@ascend/schema'
import type { CapabilityPriority, CapabilityStatus } from './capabilities.ts'

export interface WorkbookInfo {
	readonly sheetCount: number
	readonly loadedSheetCount: number
	readonly sheets: readonly SheetInfo[]
	readonly definedNames: readonly string[]
	readonly definedNameDetails: readonly DefinedNameInfo[]
	readonly cellCount: number | null
	readonly commentCount: number | null
	readonly threadedCommentCount: number | null
	readonly conditionalFormatCount: number | null
	readonly dataValidationCount: number | null
	readonly x14ConditionalFormatCount: number | null
	readonly x14DataValidationCount: number | null
	readonly imageCount: number | null
	readonly sparklineGroupCount: number | null
	readonly advancedFilterCount: number | null
	readonly chartCount: number
	readonly chartSheetCount: number
	readonly macroSheetCount: number
	readonly pivotTableCount: number
	readonly pivotCacheCount: number
	readonly pivotRefreshPlans: readonly PivotRefreshPlanInfo[]
	readonly refreshMetadata: WorkbookRefreshMetadataInfo
	readonly slicerCount: number
	readonly slicerCacheCount: number
	readonly timelineCount: number
	readonly timelineCacheCount: number
	readonly connectionPartCount: number
	readonly dataModelPartCount: number
	readonly activeContentCount: number
	readonly sourceFormat: string
	readonly workbookViewCount: number
	readonly externalReferenceCount: number
	readonly workbookViews: readonly WorkbookViewInfo[]
	readonly externalReferences: readonly string[]
	readonly externalReferenceDetails: readonly ExternalReferenceInfo[]
	readonly externalReferenceUsages: readonly ExternalReferenceUsageInfo[]
	readonly charts: readonly ChartPartInfo[]
	readonly chartSheets: readonly ChartSheetInfo[]
	readonly macroSheets: readonly WorkbookMacroSheetInfo[]
	readonly hasWorkbookProtection: boolean
	readonly pivotTables: readonly PivotTableInfo[]
	readonly pivotCaches: readonly PivotCacheInfo[]
	readonly slicerCaches: readonly {
		readonly partPath: string
		readonly name?: string
		readonly sourceName?: string
		readonly pivotCacheId?: number
		readonly pivotTableNames: readonly string[]
		readonly items?: readonly SlicerCacheItemInfo[]
	}[]
	readonly slicers: readonly {
		readonly partPath: string
		readonly name?: string
		readonly cacheName?: string
		readonly caption?: string
	}[]
	readonly timelineCaches: readonly {
		readonly partPath: string
		readonly name?: string
		readonly sourceName?: string
		readonly pivotCacheId?: number
		readonly pivotTableNames: readonly string[]
		readonly state?: TimelineStateInfo
	}[]
	readonly timelines: readonly {
		readonly partPath: string
		readonly name?: string
		readonly cacheName?: string
		readonly caption?: string
	}[]
	readonly connectionParts: readonly WorkbookConnectionPartInfo[]
	readonly dataModelParts: readonly WorkbookDataModelPartInfo[]
	readonly activeContent: readonly ActiveContentInfo[]
	readonly documentProperties: WorkbookDocumentProperties
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
		readonly colors: readonly WorkbookThemeColor[]
		readonly majorFontLatin?: string
		readonly minorFontLatin?: string
	}
	readonly capabilityWarnings: readonly CapabilityWarningInfo[]
	readonly compatibility: CompatibilityReport
	readonly load: WorkbookLoadInfo
}

export interface CapabilityWarningInfo {
	readonly capabilityId: string
	readonly label: string
	readonly family: string
	readonly status: CapabilityStatus
	readonly priority: CapabilityPriority
	readonly reason: string
	readonly nextMilestone: string
	readonly evidence: readonly string[]
}

export interface PivotTableInfo {
	readonly partPath: string
	readonly sheetName: string
	readonly name?: string
	readonly cacheId?: number
	readonly locationRef?: string
	readonly location?: PivotTableLocationInfo
	readonly options?: PivotTableOptionsInfo
	readonly style?: PivotTableStyleInfo
	readonly fields: readonly PivotFieldInfo[]
	readonly rowFields: readonly PivotFieldReference[]
	readonly columnFields: readonly PivotFieldReference[]
	readonly pageFields: readonly PivotFieldReference[]
	readonly dataFields: readonly PivotDataFieldInfo[]
	readonly rowItems?: readonly PivotAxisItemInfo[]
	readonly columnItems?: readonly PivotAxisItemInfo[]
	readonly formats?: readonly PivotFormatInfo[]
	readonly chartFormats?: readonly PivotChartFormatInfo[]
}

export interface PivotTableLocationInfo {
	readonly ref?: string
	readonly firstHeaderRow?: number
	readonly firstDataRow?: number
	readonly firstDataCol?: number
	readonly rowPageCount?: number
	readonly colPageCount?: number
}

export interface PivotTableOptionsInfo {
	readonly applyAlignmentFormats?: boolean
	readonly applyBorderFormats?: boolean
	readonly applyFontFormats?: boolean
	readonly applyNumberFormats?: boolean
	readonly applyPatternFormats?: boolean
	readonly applyWidthHeightFormats?: boolean
	readonly colGrandTotals?: boolean
	readonly rowGrandTotals?: boolean
	readonly compact?: boolean
	readonly compactData?: boolean
	readonly dataOnRows?: boolean
	readonly enableDrill?: boolean
	readonly enableEdit?: boolean
	readonly gridDropZones?: boolean
	readonly hideValuesRow?: boolean
	readonly itemPrintTitles?: boolean
	readonly multipleFieldFilters?: boolean
	readonly outline?: boolean
	readonly outlineData?: boolean
	readonly showItems?: boolean
	readonly showMemberPropertyTips?: boolean
	readonly showMultipleLabel?: boolean
	readonly useAutoFormatting?: boolean
	readonly indent?: number
	readonly createdVersion?: number
	readonly updatedVersion?: number
	readonly minRefreshableVersion?: number
	readonly dataPosition?: number
	readonly chartFormat?: number
	readonly dataCaption?: string
	readonly rowHeaderCaption?: string
	readonly colHeaderCaption?: string
	readonly fillDownLabelsDefault?: boolean
	readonly enabledSubtotalsDefault?: boolean
	readonly subtotalsOnTopDefault?: boolean
}

export interface PivotTableStyleInfo {
	readonly name?: string
	readonly showRowHeaders?: boolean
	readonly showColHeaders?: boolean
	readonly showRowStripes?: boolean
	readonly showColStripes?: boolean
	readonly showLastColumn?: boolean
}

export interface ExternalReferenceInfo {
	readonly partPath: string
	readonly relId?: string
	readonly linkRelId?: string
	readonly target?: string
	readonly targetMode?: string
}

export interface ExternalReferenceUsageInfo {
	readonly workbook: string
	readonly sheet?: string
	readonly sourceKind:
		| 'cellFormula'
		| 'definedName'
		| 'chartSeriesName'
		| 'chartSeriesCategory'
		| 'chartSeriesValue'
		| 'conditionalFormat'
		| 'dataValidation'
		| 'sparklineGroupRange'
		| 'sparklineDateAxisRange'
		| 'sparklineRange'
		| 'tableColumnFormula'
		| 'tableTotalsRowFormula'
		| 'x14ConditionalFormat'
		| 'x14DataValidation'
	readonly sourceRef?: string
	readonly name?: string
	readonly formula: string
	readonly references: readonly string[]
	readonly externalReference?: ExternalReferenceInfo
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
	readonly upgradeOnRefresh?: boolean
	readonly extensionCacheId?: number
	readonly sourceType?: string
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly sourceName?: string
	readonly recordsPartPath?: string
	readonly records?: PivotCacheRecordsInfo
	readonly fields: readonly PivotCacheFieldInfo[]
}

export interface PivotCacheRecordsInfo {
	readonly partPath: string
	readonly declaredCount?: number
	readonly parsedCount: number
	readonly preview: readonly PivotCacheRecordInfo[]
	readonly materializedRecords?: readonly PivotCacheRecordInfo[]
	readonly materializedCount?: number
	readonly materializedComplete?: boolean
	readonly valueKindCounts: readonly PivotCacheRecordValueKindCount[]
}

export interface PivotCacheRowsOptions {
	readonly cacheId?: number
	readonly partPath?: string
	readonly limit?: number
}

export interface PivotCacheMaterializedRowInfo {
	readonly partPath: string
	readonly cacheId?: number
	readonly rowIndex: number
	readonly values: readonly PivotCacheDecodedValueInfo[]
}

export interface PivotCacheDecodedValueInfo {
	readonly fieldIndex: number
	readonly fieldName?: string
	readonly rawKind: PivotCacheRecordValueKind
	readonly kind: PivotCacheRecordValueKind
	readonly value?: string
	readonly sharedItemIndex?: number
	readonly sharedItemKind?: PivotCacheSharedItemInfo['kind']
}

export type PivotOutputAuditStatus = 'passed' | 'mismatch' | 'unsupported'

export interface PivotOutputAuditInfo {
	readonly pivotTable?: string
	readonly partPath: string
	readonly sheetName: string
	readonly cacheId?: number
	readonly status: PivotOutputAuditStatus
	readonly checkedValueCount: number
	readonly mismatches: readonly PivotOutputAuditMismatchInfo[]
	readonly warnings: readonly string[]
}

export interface PivotOutputAuditMismatchInfo {
	readonly ref?: string
	readonly rowLabel: string
	readonly dataField: string
	readonly expected: number
	readonly actual?: CellValue
}

export interface PivotCacheRecordInfo {
	readonly index: number
	readonly values: readonly PivotCacheRecordValueInfo[]
}

export interface PivotCacheRecordValueInfo {
	readonly index: number
	readonly kind: PivotCacheRecordValueKind
	readonly value?: string
	readonly sharedItemIndex?: number
}

export interface PivotCacheRecordValueKindCount {
	readonly kind: PivotCacheRecordValueKind
	readonly count: number
}

export type PivotCacheRecordValueKind =
	| 'string'
	| 'number'
	| 'date'
	| 'boolean'
	| 'error'
	| 'missing'
	| 'sharedItem'
	| 'unknown'

export type PivotRefreshOutputState =
	| 'cached'
	| 'stale'
	| 'refresh-on-open'
	| 'not-saved'
	| 'unknown'

export interface PivotRefreshPlanInfo {
	readonly partPath: string
	readonly cacheId?: number
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly cacheRecords?: PivotRefreshCacheRecordsInfo
	readonly pivotTables: readonly PivotRefreshTableInfo[]
	readonly outputState: PivotRefreshOutputState
	readonly canRefreshHeadlessly: false
	readonly requiresExternalRefresh: boolean
	readonly warnings: readonly string[]
	readonly recommendedOps: readonly PivotRefreshRecommendedOp[]
}

export interface PivotRefreshCacheRecordsInfo {
	readonly partPath: string
	readonly declaredCount?: number
	readonly parsedCount: number
	readonly materializedCount?: number
	readonly materializedComplete?: boolean
	readonly valueKindCounts: readonly PivotCacheRecordValueKindCount[]
}

export interface PivotRefreshTableInfo {
	readonly partPath: string
	readonly sheetName: string
	readonly name?: string
	readonly locationRef?: string
}

export interface PivotRefreshRecommendedOp {
	readonly op: 'setPivotCache'
	readonly partPath?: string
	readonly cacheId?: number
	readonly refreshOnLoad?: boolean
	readonly invalid?: boolean
	readonly saveData?: boolean
}

export type WorkbookRefreshMetadataKind =
	| 'calcSettings'
	| 'calcChain'
	| 'pivotCache'
	| 'workbookConnection'
	| 'queryTable'

export type WorkbookRefreshMetadataState =
	| 'cached'
	| 'stale'
	| 'refresh-on-open'
	| 'not-saved'
	| 'manual-calc'
	| 'unknown'

export interface WorkbookRefreshMetadataInfo {
	readonly entries: readonly WorkbookRefreshMetadataEntry[]
	readonly refreshOnOpenCount: number
	readonly staleCacheCount: number
	readonly notSavedCount: number
	readonly unknownCount: number
}

export interface WorkbookRefreshMetadataEntry {
	readonly kind: WorkbookRefreshMetadataKind
	readonly partPath: string
	readonly state: WorkbookRefreshMetadataState
	readonly name?: string
	readonly sheetName?: string
	readonly cacheId?: number
	readonly connectionId?: number
	readonly refreshOnLoad?: boolean
	readonly saveData?: boolean
	readonly invalid?: boolean
	readonly refreshedVersion?: number
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly warnings: readonly string[]
	readonly recommendedOps: readonly unknown[]
}

export interface GetPivotDataQuery {
	readonly dataField: string
	readonly pivotTable?: string
	readonly filters?: readonly GetPivotDataFilter[]
}

export interface GetPivotDataFilter {
	readonly field: string
	readonly item: string
}

export interface GetPivotDataMatchInfo {
	readonly pivotTable: PivotTableInfo
	readonly dataField: PivotDataFieldInfo
	readonly matchedFilters: readonly GetPivotDataFilter[]
	readonly unmatchedFilters: readonly GetPivotDataFilter[]
	readonly output?: GetPivotDataOutputInfo
}

export interface GetPivotDataOutputInfo {
	readonly sheetName: string
	readonly ref: string
	readonly value: CellValue
}

export interface GetPivotDataResult {
	readonly query: GetPivotDataQuery
	readonly matches: readonly GetPivotDataMatchInfo[]
	readonly canResolveOutput: boolean
	readonly warnings: readonly string[]
}

export interface PivotCacheFieldInfo {
	readonly index: number
	readonly name?: string
	readonly databaseField?: boolean
	readonly numFmtId?: number
	readonly formula?: string
	readonly sharedItemsInfo?: PivotCacheSharedItemsInfo
	readonly sharedItems?: readonly PivotCacheSharedItemInfo[]
	readonly fieldGroup?: PivotCacheFieldGroupInfo
}

export interface PivotCacheSharedItemInfo {
	readonly index: number
	readonly kind: 'string' | 'number' | 'date' | 'boolean' | 'error' | 'missing'
	readonly value?: string
}

export interface PivotCacheSharedItemsInfo {
	readonly count?: number
	readonly containsBlank?: boolean
	readonly containsDate?: boolean
	readonly containsNonDate?: boolean
	readonly containsNumber?: boolean
	readonly containsInteger?: boolean
	readonly containsString?: boolean
	readonly containsMixedTypes?: boolean
	readonly containsSemiMixedTypes?: boolean
	readonly minValue?: number
	readonly maxValue?: number
	readonly minDate?: string
	readonly maxDate?: string
}

export interface PivotCacheFieldGroupInfo {
	readonly base?: number
	readonly parent?: number
	readonly range?: PivotCacheFieldGroupRangeInfo
	readonly discreteItems?: readonly PivotCacheFieldGroupDiscreteItemInfo[]
	readonly groupItems?: readonly PivotCacheSharedItemInfo[]
}

export interface PivotCacheFieldGroupRangeInfo {
	readonly groupBy?: string
	readonly startDate?: string
	readonly endDate?: string
	readonly startNumber?: number
	readonly endNumber?: number
	readonly groupInterval?: number
	readonly autoStart?: boolean
	readonly autoEnd?: boolean
}

export interface PivotCacheFieldGroupDiscreteItemInfo {
	readonly index: number
	readonly value?: number
}

export interface PivotFieldInfo {
	readonly index: number
	readonly axis?: string
	readonly name?: string
	readonly numFmtId?: number
	readonly hidden?: boolean
	readonly dataField?: boolean
	readonly defaultSubtotal?: boolean
	readonly showAll?: boolean
	readonly multipleItemSelectionAllowed?: boolean
	readonly compact?: boolean
	readonly outline?: boolean
	readonly subtotalTop?: boolean
	readonly dragToRow?: boolean
	readonly dragToCol?: boolean
	readonly dragToPage?: boolean
	readonly includeNewItemsInFilter?: boolean
	readonly fillDownLabels?: boolean
	readonly itemPageCount?: number
	readonly sortType?: string
	readonly items?: readonly PivotFieldItemInfo[]
}

export interface PivotFieldItemInfo {
	readonly index: number
	readonly cacheIndex?: number
	readonly itemType?: string
	readonly caption?: string
	readonly hidden?: boolean
	readonly manualFilter?: boolean
	readonly showDetails?: boolean
	readonly calculated?: boolean
	readonly missing?: boolean
	readonly childItems?: boolean
	readonly expanded?: boolean
	readonly drillAcrossAttributes?: boolean
}

export interface PivotFieldReference {
	readonly index: number
	readonly name?: string
	readonly item?: number
	readonly hierarchy?: number
	readonly caption?: string
}

export interface PivotDataFieldInfo {
	readonly fieldIndex: number
	readonly name?: string
	readonly subtotal?: string
	readonly numFmtId?: number
	readonly showDataAs?: string
	readonly baseField?: number
	readonly baseItem?: number
}

export interface PivotAxisItemInfo {
	readonly index: number
	readonly itemType?: string
	readonly repeatedItemCount?: number
	readonly dataFieldIndex?: number
	readonly fieldItems: readonly PivotAxisFieldItemInfo[]
}

export interface PivotAxisFieldItemInfo {
	readonly index: number
	readonly item?: number
}

export interface PivotFormatInfo {
	readonly index: number
	readonly dxfId?: number
	readonly action?: string
	readonly area?: PivotAreaInfo
}

export interface PivotChartFormatInfo {
	readonly index: number
	readonly chart?: number
	readonly formatId?: number
	readonly series?: boolean
	readonly area?: PivotAreaInfo
}

export interface PivotAreaInfo {
	readonly type?: string
	readonly axis?: string
	readonly field?: number
	readonly fieldPosition?: number
	readonly dataOnly?: boolean
	readonly labelOnly?: boolean
	readonly grandRow?: boolean
	readonly grandCol?: boolean
	readonly cacheIndex?: boolean
	readonly outline?: boolean
	readonly collapsedLevelsAreSubtotals?: boolean
	readonly references?: readonly PivotAreaReferenceInfo[]
}

export interface PivotAreaReferenceInfo {
	readonly index: number
	readonly field?: number
	readonly itemCount?: number
	readonly selected?: boolean
	readonly items: readonly PivotAxisFieldItemInfo[]
}

export interface SlicerCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly items?: readonly SlicerCacheItemInfo[]
}

export interface SlicerCacheItemInfo {
	readonly index: number
	readonly selected?: boolean
	readonly noData?: boolean
}

export interface SlicerInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface TimelineCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly state?: TimelineStateInfo
}

export interface TimelineInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface TimelineRangeInfo {
	readonly startDate: string
	readonly endDate: string
}

export interface TimelineStateInfo {
	readonly filterType?: string
	readonly filterId?: number
	readonly filterPivotName?: string
	readonly filterTabId?: number
	readonly lastRefreshVersion?: number
	readonly minimalRefreshVersion?: number
	readonly pivotCacheId?: number
	readonly singleRangeFilterState?: boolean
	readonly selection?: TimelineRangeInfo
	readonly bounds?: TimelineRangeInfo
}

export interface SheetInfo {
	readonly name: string
	readonly rowCount: number | null
	readonly colCount: number | null
	readonly cellCount: number | null
	readonly tableCount: number | null
	readonly commentCount: number | null
	readonly threadedCommentCount: number | null
	readonly conditionalFormatCount: number | null
	readonly dataValidationCount: number | null
	readonly x14ConditionalFormatCount: number | null
	readonly x14DataValidationCount: number | null
	readonly hasFrozenPanes: boolean | null
	readonly colWidthCount: number | null
	readonly imageCount: number | null
	readonly sparklineGroupCount: number | null
	readonly advancedFilterCount: number | null
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

export interface SheetThreadedCommentInfo extends SheetThreadedComment {}

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
	readonly location?: string
	readonly display?: string
	readonly tooltip?: string
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
	| {
			readonly kind: 'blocked-spill'
			readonly formula: string
			readonly range: string
			readonly blockingRefs: readonly string[]
	  }

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
	readonly threadedComments: readonly SheetThreadedCommentInfo[] | null
	readonly hyperlinks: readonly SheetHyperlinkInfo[] | null
	readonly ignoredErrors: readonly SheetIgnoredError[] | null
	readonly conditionalFormats: readonly SheetConditionalFormat[] | null
	readonly dataValidations: readonly SheetDataValidation[] | null
	readonly x14ConditionalFormats: readonly SheetX14ConditionalFormatInfo[] | null
	readonly x14DataValidations: readonly SheetX14DataValidationInfo[] | null
	readonly imageRefs: readonly SheetImageRef[] | null
	readonly drawingObjectRefs: readonly SheetDrawingObjectRef[] | null
	readonly sparklineGroups: readonly SheetSparklineGroupInfo[] | null
	readonly advancedFilters: readonly SheetAdvancedFilterInfo[] | null
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
	readonly drawingObjectRefs: readonly SheetDrawingObjectRef[] | null
	readonly imageCount: number | null
	readonly drawingObjectCount: number | null
}

export interface WorkbookVisualInventoryInfo {
	readonly load: WorkbookLoadInfo
	readonly packageFeatures: readonly VisualPackageFeatureInfo[]
	readonly sheets: readonly SheetVisualInventoryInfo[]
	readonly sheetImageCount: number | null
	readonly sheetDrawingObjectCount: number | null
	readonly charts: readonly ChartPartInfo[]
	readonly chartSheets: readonly ChartSheetInfo[]
	readonly structuredChartCount: number
	readonly chartSheetCount: number
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

export type PivotOutputMaterializeMode = 'missing' | 'mismatches' | 'all'

export interface PivotOutputMaterializeOptions {
	readonly pivotTable?: string
	readonly partPath?: string
	readonly mode?: PivotOutputMaterializeMode
}

export interface PivotOutputMaterializeUnsupportedInfo {
	readonly pivotTable?: string
	readonly partPath: string
	readonly sheetName: string
	readonly cacheId?: number
	readonly warning: string
}

export interface PivotOutputMaterializeOpsResult {
	readonly ops: readonly Operation[]
	readonly plannedCellCount: number
	readonly unsupported: readonly PivotOutputMaterializeUnsupportedInfo[]
}

export interface PivotOutputMaterializeResult extends PivotOutputMaterializeOpsResult {
	readonly apply: ApplyResult
}

export interface BatchResult {
	readonly errors: readonly AscendError[]
}

export type {
	ExternalCellReference,
	ExternalRangeReference,
	ExternalReferenceResolver,
} from '@ascend/engine'

export interface RecalcOptions {
	readonly range?: string
	readonly externalReferences?: ExternalReferenceResolver
}

export interface EvalOptions {
	readonly externalReferences?: ExternalReferenceResolver
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
	readonly parts: readonly WritePlanPartInfo[]
	readonly skippedCapsules: readonly string[]
}

export interface WritePlanPartInfo {
	readonly path: string
	readonly owner:
		| { readonly kind: 'package' }
		| { readonly kind: 'workbook' }
		| { readonly kind: 'sheet'; readonly sheetName: string }
	readonly origin: 'generated' | 'preserved-inline' | 'preserved-source' | 'capsule'
	readonly contentType?: string
	readonly streaming: boolean
}

export interface CheckResult {
	readonly valid: boolean
	readonly issues: readonly CheckIssue[]
}

export interface CheckIssue {
	readonly rule?: string
	readonly severity: 'error' | 'warning' | 'info'
	readonly message: string
	readonly ref?: string
	readonly refs?: readonly string[]
	readonly suggestedFix?: string
	readonly details?: Readonly<Record<string, unknown>>
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
	readonly endColumn?: string
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
