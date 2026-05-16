import type { AutoFilter, SortState } from './filter.ts'
import { createSheetId, type SheetId } from './ids.ts'
import type { RangeRef } from './refs.ts'
import { SparseGrid } from './sparse-grid.ts'
import type { CellStyle } from './style.ts'
import { cloneCellStyle } from './style-clone.ts'
import type { Table } from './table.ts'

export type SheetState = 'visible' | 'hidden' | 'veryHidden'

export interface SheetComment {
	readonly text: string
	readonly author?: string
	readonly legacyDrawing?: SheetCommentLegacyDrawing
}

export interface SheetCommentLegacyDrawing {
	readonly shapeId?: string
	readonly style?: string
	readonly anchor?: readonly [number, number, number, number, number, number, number, number]
	readonly row?: number
	readonly column?: number
	readonly visible?: boolean
	readonly moveWithCells?: boolean
	readonly sizeWithCells?: boolean
	readonly autoFill?: boolean
}

export interface SheetThreadedComment {
	readonly ref: string
	readonly text: string
	readonly id?: string
	readonly parentId?: string
	readonly personId?: string
	readonly author?: string
	readonly dateTime?: string
	readonly done?: boolean
	readonly partPath?: string
}

export interface SheetHyperlink {
	readonly target?: string
	readonly location?: string
	readonly display?: string
	readonly tooltip?: string
}

export interface SheetPageMargins {
	readonly left?: number
	readonly right?: number
	readonly top?: number
	readonly bottom?: number
	readonly header?: number
	readonly footer?: number
}

export interface SheetPageSetup {
	readonly orientation?: string
	readonly paperSize?: number
	readonly scale?: number
	readonly fitToWidth?: number
	readonly fitToHeight?: number
	readonly firstPageNumber?: number
	readonly copies?: number
	readonly horizontalDpi?: number
	readonly verticalDpi?: number
	readonly pageOrder?: string
	readonly cellComments?: string
	readonly errors?: string
	readonly blackAndWhite?: boolean
	readonly draft?: boolean
	readonly useFirstPageNumber?: boolean
	readonly usePrinterDefaults?: boolean
	readonly printerSettingsRelId?: string
}

export interface SheetPageSetupPr {
	readonly fitToPage?: boolean
	readonly autoPageBreaks?: boolean
}

export interface SheetPrintOptions {
	readonly gridLines?: boolean
	readonly gridLinesSet?: boolean
	readonly headings?: boolean
	readonly horizontalCentered?: boolean
	readonly verticalCentered?: boolean
}

export interface SheetProtection {
	readonly sheet?: boolean
	readonly objects?: boolean
	readonly scenarios?: boolean
	readonly formatCells?: boolean
	readonly formatColumns?: boolean
	readonly formatRows?: boolean
	readonly insertColumns?: boolean
	readonly insertRows?: boolean
	readonly insertHyperlinks?: boolean
	readonly deleteColumns?: boolean
	readonly deleteRows?: boolean
	readonly selectLockedCells?: boolean
	readonly sort?: boolean
	readonly autoFilter?: boolean
	readonly pivotTables?: boolean
	readonly selectUnlockedCells?: boolean
	readonly password?: string
	readonly algorithmName?: string
	readonly hashValue?: string
	readonly saltValue?: string
	readonly spinCount?: number
}

export interface SheetProtectedRange {
	readonly name?: string
	readonly sqref: string
	readonly password?: string
	readonly algorithmName?: string
	readonly hashValue?: string
	readonly saltValue?: string
	readonly spinCount?: number
	readonly securityDescriptor?: string
}

export interface SheetHeaderFooter {
	readonly differentOddEven?: boolean
	readonly differentFirst?: boolean
	readonly scaleWithDoc?: boolean
	readonly alignWithMargins?: boolean
	readonly oddHeader?: string
	readonly oddFooter?: string
	readonly evenHeader?: string
	readonly evenFooter?: string
	readonly firstHeader?: string
	readonly firstFooter?: string
}

export interface SheetColDef {
	readonly min: number
	readonly max: number
	readonly width?: number
	readonly style?: number
	readonly hidden?: boolean
	readonly bestFit?: boolean
	readonly collapsed?: boolean
	readonly outlineLevel?: number
	readonly customWidth?: boolean
}

export interface SheetRowDef {
	readonly spans?: string
	readonly style?: number
	readonly customFormat?: boolean
	readonly customHeight?: boolean
	readonly hidden?: boolean
	readonly collapsed?: boolean
	readonly outlineLevel?: number
	readonly thickTop?: boolean
	readonly thickBot?: boolean
	readonly dyDescent?: number
}

export interface SheetPhoneticPr {
	readonly fontId?: number
	readonly type?: string
	readonly alignment?: string
}

export interface SheetCellMetadataAttrs {
	readonly cm?: number
	readonly vm?: number
	readonly ph?: boolean
}

export interface SheetOutlinePr {
	readonly summaryBelow?: boolean
	readonly summaryRight?: boolean
	readonly applyStyles?: boolean
	readonly showOutlineSymbols?: boolean
}

export interface SheetPreservedXml {
	readonly partPath?: string
	readonly xml?: string
	readonly relsPath?: string
	readonly relsXml?: string
}

export interface SheetDrawingRefs {
	readonly hasDrawing: boolean
	readonly hasLegacyDrawing: boolean
}

export interface SheetAnchorMarker {
	readonly col: number
	readonly row: number
	readonly colOff?: number
	readonly rowOff?: number
}

export type SheetImageAnchor =
	| {
			readonly kind: 'oneCell'
			readonly from: SheetAnchorMarker
			readonly cx?: number
			readonly cy?: number
	  }
	| {
			readonly kind: 'twoCell'
			readonly from: SheetAnchorMarker
			readonly to: SheetAnchorMarker
			readonly editAs?: string
	  }
	| {
			readonly kind: 'absolute'
			readonly x: number
			readonly y: number
			readonly cx?: number
			readonly cy?: number
	  }

export interface SheetImageRef {
	readonly drawingPartPath: string
	readonly relId: string
	readonly targetPath: string
	readonly contentType?: string
	readonly content?: Uint8Array
	readonly anchor?: SheetImageAnchor
	readonly name?: string
	readonly description?: string
}

export type SheetDrawingObjectKind =
	| 'shape'
	| 'textBox'
	| 'connector'
	| 'groupShape'
	| 'graphicFrame'
	| 'unknown'

export interface SheetDrawingObjectRelationshipRef {
	readonly id: string
	readonly type: string
	readonly target: string
	readonly targetMode?: string
}

export interface SheetDrawingObjectRef {
	readonly drawingPartPath: string
	readonly kind: SheetDrawingObjectKind
	readonly source?: 'drawingml' | 'vml'
	readonly anchor?: SheetImageAnchor
	readonly id?: number
	readonly name?: string
	readonly description?: string
	readonly text?: string
	readonly macro?: string
	readonly style?: string
	readonly vmlShapeId?: string
	readonly vmlObjectType?: string
	readonly visible?: boolean
	readonly relIds?: readonly string[]
	readonly relationshipRefs?: readonly SheetDrawingObjectRelationshipRef[]
}

export interface SheetSparklineGroupInfo {
	readonly groupIndex: number
	readonly type?: string
	readonly manualMax?: number
	readonly manualMin?: number
	readonly lineWeight?: number
	readonly uid?: string
	readonly displayEmptyCellsAs?: string
	readonly dateAxis?: boolean
	readonly markers?: boolean
	readonly highPoint?: boolean
	readonly lowPoint?: boolean
	readonly firstPoint?: boolean
	readonly lastPoint?: boolean
	readonly negative?: boolean
	readonly displayXAxis?: boolean
	readonly displayHidden?: boolean
	readonly rightToLeft?: boolean
	readonly minAxisType?: string
	readonly maxAxisType?: string
	readonly colorSeries?: string
	readonly colorNegative?: string
	readonly colorAxis?: string
	readonly colorMarkers?: string
	readonly colorFirst?: string
	readonly colorLast?: string
	readonly colorHigh?: string
	readonly colorLow?: string
	readonly dateAxisRange?: string
	readonly range?: string
	readonly locationRange?: string
	readonly sparklines?: readonly SheetSparklineInfo[]
	readonly count: number
}

export interface SheetSparklineInfo {
	readonly range?: string
	readonly locationRange?: string
}

export interface SheetX14ConditionalFormatInfo {
	readonly index: number
	readonly sqref: string
	readonly formulas: readonly string[]
	readonly type?: string
	readonly priority?: number
	readonly id?: string
	readonly preservedRuleAttributes?: Readonly<Record<string, string>>
	readonly preservedRuleChildXml?: readonly string[]
	readonly colorScale?: SheetConditionalFormatColorScale
	readonly dataBar?: SheetX14ConditionalFormatDataBarInfo
	readonly iconSet?: SheetX14ConditionalFormatIconSetInfo
	readonly deleted?: boolean
}

export interface SheetX14ConditionalFormatDataBarInfo {
	readonly cfvo: readonly SheetConditionalFormatValueObject[]
	readonly minLength?: number
	readonly maxLength?: number
	readonly border?: boolean
	readonly showValue?: boolean
	readonly gradient?: boolean
	readonly direction?: string
	readonly axisPosition?: string
	readonly negativeBarColorSameAsPositive?: boolean
	readonly negativeBarBorderColorSameAsPositive?: boolean
	readonly fillColor?: SheetConditionalFormatColor
	readonly borderColor?: SheetConditionalFormatColor
	readonly negativeFillColor?: SheetConditionalFormatColor
	readonly negativeBorderColor?: SheetConditionalFormatColor
	readonly axisColor?: SheetConditionalFormatColor
}

export interface SheetX14ConditionalFormatIconSetInfo {
	readonly cfvo: readonly SheetConditionalFormatValueObject[]
	readonly iconSet?: string
	readonly custom?: boolean
	readonly showValue?: boolean
	readonly percent?: boolean
	readonly reverse?: boolean
	readonly icons?: readonly SheetX14ConditionalFormatIconInfo[]
}

export interface SheetX14ConditionalFormatIconInfo {
	readonly iconSet?: string
	readonly iconId?: number
}

export interface SheetX14DataValidationInfo {
	readonly index: number
	readonly sqref: string
	readonly type?: string
	readonly operator?: string
	readonly allowBlank?: boolean
	readonly showInputMessage?: boolean
	readonly showErrorMessage?: boolean
	readonly showDropDown?: boolean
	readonly promptTitle?: string
	readonly prompt?: string
	readonly errorTitle?: string
	readonly error?: string
	readonly errorStyle?: string
	readonly imeMode?: string
	readonly formula1?: string
	readonly formula2?: string
	readonly preservedAttributes?: Readonly<Record<string, string>>
	readonly preservedChildXml?: readonly string[]
	readonly deleted?: boolean
}

export interface SheetAdvancedFilterInfo {
	readonly viewName?: string
	readonly guid?: string
	readonly ref?: string
	readonly autoFilter?: AutoFilter
	readonly filterColumnCount: number
	readonly sortConditionCount: number
}

export interface SheetDataValidation {
	readonly sqref: string
	readonly source?: 'x14'
	readonly uid?: string
	readonly type?: string
	readonly operator?: string
	readonly allowBlank?: boolean
	readonly showInputMessage?: boolean
	readonly showErrorMessage?: boolean
	readonly showDropDown?: boolean
	readonly promptTitle?: string
	readonly prompt?: string
	readonly errorTitle?: string
	readonly error?: string
	readonly errorStyle?: string
	readonly imeMode?: string
	readonly formula1?: string
	readonly formula2?: string
}

export interface SheetDataValidationSettings {
	readonly disablePrompts?: boolean
	readonly xWindow?: number
	readonly yWindow?: number
}

export interface SheetConditionalFormatRule {
	readonly type: string
	readonly operator?: string
	readonly dxfId?: number
	readonly priority?: number
	readonly stopIfTrue?: boolean
	readonly formulas: readonly string[]
	readonly style?: CellStyle
	readonly rank?: number
	readonly percent?: boolean
	readonly bottom?: boolean
	readonly aboveAverage?: boolean
	readonly equalAverage?: boolean
	readonly stdDev?: number
	readonly text?: string
	readonly timePeriod?: string
	readonly preservedRuleAttributes?: Readonly<Record<string, string>>
	readonly preservedRuleChildXml?: readonly string[]
	readonly colorScale?: SheetConditionalFormatColorScale
	readonly dataBar?: SheetConditionalFormatDataBar
	readonly iconSet?: SheetConditionalFormatIconSet
}

export interface SheetConditionalFormatValueObject {
	readonly type?: string
	readonly value?: string
	readonly gte?: boolean
}

export interface SheetConditionalFormatColor {
	readonly rgb?: string
	readonly theme?: number
	readonly tint?: number
	readonly indexed?: number
	readonly auto?: boolean
}

export interface SheetConditionalFormatColorScale {
	readonly cfvo: readonly SheetConditionalFormatValueObject[]
	readonly colors: readonly SheetConditionalFormatColor[]
	readonly preservedAttributes?: Readonly<Record<string, string>>
	readonly preservedChildXml?: readonly string[]
}

export interface SheetConditionalFormatDataBar {
	readonly cfvo: readonly SheetConditionalFormatValueObject[]
	readonly color?: SheetConditionalFormatColor
	readonly minLength?: number
	readonly maxLength?: number
	readonly showValue?: boolean
}

export interface SheetConditionalFormatIconSet {
	readonly cfvo: readonly SheetConditionalFormatValueObject[]
	readonly iconSet?: string
	readonly showValue?: boolean
	readonly percent?: boolean
	readonly reverse?: boolean
}

export interface SheetConditionalFormat {
	readonly sqref: string
	readonly pivot?: boolean
	readonly rules: readonly SheetConditionalFormatRule[]
}

function cloneConditionalFormatRule(rule: SheetConditionalFormatRule): SheetConditionalFormatRule {
	return {
		...rule,
		formulas: [...rule.formulas],
		...(rule.preservedRuleAttributes
			? { preservedRuleAttributes: { ...rule.preservedRuleAttributes } }
			: {}),
		...(rule.preservedRuleChildXml
			? { preservedRuleChildXml: [...rule.preservedRuleChildXml] }
			: {}),
		...(rule.style ? { style: cloneCellStyle(rule.style) } : {}),
		...(rule.colorScale
			? {
					colorScale: {
						cfvo: rule.colorScale.cfvo.map((entry) => ({ ...entry })),
						colors: rule.colorScale.colors.map((entry) => ({ ...entry })),
						...(rule.colorScale.preservedAttributes
							? { preservedAttributes: { ...rule.colorScale.preservedAttributes } }
							: {}),
						...(rule.colorScale.preservedChildXml
							? { preservedChildXml: [...rule.colorScale.preservedChildXml] }
							: {}),
					},
				}
			: {}),
		...(rule.dataBar
			? {
					dataBar: {
						...rule.dataBar,
						cfvo: rule.dataBar.cfvo.map((entry) => ({ ...entry })),
						...(rule.dataBar.color ? { color: { ...rule.dataBar.color } } : {}),
					},
				}
			: {}),
		...(rule.iconSet
			? {
					iconSet: {
						...rule.iconSet,
						cfvo: rule.iconSet.cfvo.map((entry) => ({ ...entry })),
					},
				}
			: {}),
	}
}

export interface SheetIgnoredError {
	readonly sqref: string
	readonly numberStoredAsText?: boolean
	readonly formula?: boolean
	readonly formulaRange?: boolean
	readonly evalError?: boolean
	readonly twoDigitTextYear?: boolean
	readonly unlockedFormula?: boolean
	readonly emptyCellReference?: boolean
	readonly listDataValidation?: boolean
	readonly calculatedColumn?: boolean
}

export interface SheetTabColor {
	readonly rgb?: string
	readonly theme?: number
	readonly tint?: number
	readonly indexed?: number
}

export interface SheetFormatPr {
	readonly baseColWidth?: number
	readonly defaultRowHeight?: number
	readonly defaultColWidth?: number
	readonly outlineLevelRow?: number
	readonly outlineLevelCol?: number
	readonly customHeight?: boolean
	readonly zeroHeight?: boolean
	readonly dyDescent?: number
}

export interface SheetBreak {
	readonly id: number
	readonly min?: number
	readonly max?: number
	readonly man?: boolean
	readonly pt?: boolean
}

export type SheetViewType = 'normal' | 'pageBreakPreview' | 'pageLayout'

export interface SheetView {
	readonly zoomScale?: number
	readonly zoomScaleNormal?: number
	readonly zoomScaleSheetLayoutView?: number
	readonly showGridLines?: boolean
	readonly showFormulas?: boolean
	readonly rightToLeft?: boolean
	readonly tabSelected?: boolean
	readonly view?: SheetViewType
	readonly topLeftCell?: string
}

export type SheetViewSelection = Readonly<Record<string, string>>

export class Sheet {
	readonly id: SheetId
	name: string
	readonly cells: SparseGrid
	merges: RangeRef[]
	tables: Table[]
	state: SheetState
	colWidths: Map<number, number>
	colDefs: SheetColDef[]
	rowHeights: Map<number, number>
	rowDefs: Map<number, SheetRowDef>
	frozenRows: number
	frozenCols: number
	codeName: string | null
	filterMode: boolean | null
	enableFormatConditionsCalculation: boolean | null
	sheetView: SheetView | null
	preservedSheetViewAttributes: Record<string, string> | null
	preservedPaneAttributes: Record<string, string> | null
	preservedSheetViewSelections: SheetViewSelection[] | null
	preservedDimensionRef: string | null
	preservedBlankCells: Map<number, Map<number, string>>
	preservedCellMetadata: Map<string, SheetCellMetadataAttrs>
	comments: Map<string, SheetComment>
	threadedComments: SheetThreadedComment[]
	hyperlinks: Map<string, SheetHyperlink>
	ignoredErrors: SheetIgnoredError[]
	tabColor: SheetTabColor | null
	outlinePr: SheetOutlinePr | null
	sheetFormatPr: SheetFormatPr | null
	dataValidationSettings: SheetDataValidationSettings | null
	dataValidations: SheetDataValidation[]
	conditionalFormats: SheetConditionalFormat[]
	imageRefs: SheetImageRef[]
	drawingObjectRefs: SheetDrawingObjectRef[]
	sparklineGroups: SheetSparklineGroupInfo[]
	x14ConditionalFormats: SheetX14ConditionalFormatInfo[]
	x14DataValidations: SheetX14DataValidationInfo[]
	advancedFilters: SheetAdvancedFilterInfo[]
	drawingRefs: SheetDrawingRefs
	autoFilter: AutoFilter | null
	preservedAutoFilterSortStateAttributes: Record<string, string> | null
	sortState: SortState | null
	preservedSortStateAttributes: Record<string, string> | null
	protection: SheetProtection | null
	protectedRanges: SheetProtectedRange[]
	pageMargins: SheetPageMargins | null
	pageSetup: SheetPageSetup | null
	pageSetupPr: SheetPageSetupPr | null
	printOptions: SheetPrintOptions | null
	headerFooter: SheetHeaderFooter | null
	phoneticPr: SheetPhoneticPr | null
	rowBreaks: SheetBreak[]
	colBreaks: SheetBreak[]
	/** Exact persisted formula payloads keyed as row:col; writers validate before reuse. */
	storedFormulaText: Map<string, string>
	preservedXml: SheetPreservedXml | null
	preservedExtLst: string | null
	preservedCustomSheetViews: string | null
	preservedControlsXml: string | null
	private _shared = false

	constructor(name: string, id?: SheetId) {
		this.id = id ?? createSheetId()
		this.name = name
		this.cells = new SparseGrid()
		this.merges = []
		this.tables = []
		this.state = 'visible'
		this.colWidths = new Map()
		this.colDefs = []
		this.rowHeights = new Map()
		this.rowDefs = new Map()
		this.frozenRows = 0
		this.frozenCols = 0
		this.codeName = null
		this.filterMode = null
		this.enableFormatConditionsCalculation = null
		this.sheetView = null
		this.preservedSheetViewAttributes = null
		this.preservedPaneAttributes = null
		this.preservedSheetViewSelections = null
		this.preservedDimensionRef = null
		this.preservedBlankCells = new Map()
		this.preservedCellMetadata = new Map()
		this.comments = new Map()
		this.threadedComments = []
		this.hyperlinks = new Map()
		this.ignoredErrors = []
		this.tabColor = null
		this.outlinePr = null
		this.sheetFormatPr = null
		this.dataValidationSettings = null
		this.dataValidations = []
		this.conditionalFormats = []
		this.imageRefs = []
		this.drawingObjectRefs = []
		this.sparklineGroups = []
		this.x14ConditionalFormats = []
		this.x14DataValidations = []
		this.advancedFilters = []
		this.drawingRefs = { hasDrawing: false, hasLegacyDrawing: false }
		this.autoFilter = null
		this.preservedAutoFilterSortStateAttributes = null
		this.sortState = null
		this.preservedSortStateAttributes = null
		this.protection = null
		this.protectedRanges = []
		this.pageMargins = null
		this.pageSetup = null
		this.pageSetupPr = null
		this.printOptions = null
		this.headerFooter = null
		this.phoneticPr = null
		this.rowBreaks = []
		this.colBreaks = []
		this.storedFormulaText = new Map()
		this.preservedXml = null
		this.preservedExtLst = null
		this.preservedCustomSheetViews = null
		this.preservedControlsXml = null
	}

	ensureWritable(): void {
		if (!this._shared) return
		this.merges = this.merges.map(cloneRangeRef)
		this.tables = this.tables.map(cloneTable)
		this.colWidths = new Map(this.colWidths)
		this.colDefs = this.colDefs.map((d) => ({ ...d }))
		this.rowHeights = new Map(this.rowHeights)
		this.rowDefs = new Map([...this.rowDefs.entries()].map(([row, def]) => [row, { ...def }]))
		this.preservedSheetViewAttributes = this.preservedSheetViewAttributes
			? { ...this.preservedSheetViewAttributes }
			: null
		this.preservedPaneAttributes = this.preservedPaneAttributes
			? { ...this.preservedPaneAttributes }
			: null
		this.preservedSheetViewSelections = this.preservedSheetViewSelections
			? this.preservedSheetViewSelections.map((selection) => ({ ...selection }))
			: null
		this.preservedBlankCells = new Map(
			[...this.preservedBlankCells.entries()].map(([row, cells]) => [row, new Map(cells)]),
		)
		this.preservedCellMetadata = new Map(
			[...this.preservedCellMetadata.entries()].map(([key, attrs]) => [key, { ...attrs }]),
		)
		this.comments = new Map(
			[...this.comments.entries()].map(([ref, comment]) => [ref, cloneSheetComment(comment)]),
		)
		this.threadedComments = this.threadedComments.map((comment) => ({ ...comment }))
		this.hyperlinks = new Map(this.hyperlinks)
		this.ignoredErrors = this.ignoredErrors.map((e) => ({ ...e }))
		this.dataValidationSettings = this.dataValidationSettings
			? { ...this.dataValidationSettings }
			: null
		this.dataValidations = this.dataValidations.map((d) => ({ ...d }))
		this.rowBreaks = this.rowBreaks.map((b) => ({ ...b }))
		this.colBreaks = this.colBreaks.map((b) => ({ ...b }))
		this.storedFormulaText = new Map(this.storedFormulaText)
		this.conditionalFormats = this.conditionalFormats.map((cf) => ({
			...cf,
			rules: cf.rules.map(cloneConditionalFormatRule),
		}))
		this.imageRefs = this.imageRefs.map(cloneImageRef)
		this.drawingObjectRefs = this.drawingObjectRefs.map(cloneDrawingObjectRef)
		this.sparklineGroups = this.sparklineGroups.map((group) => ({
			...group,
			...(group.sparklines ? { sparklines: group.sparklines.map((entry) => ({ ...entry })) } : {}),
		}))
		this.x14ConditionalFormats = this.x14ConditionalFormats.map(cloneX14ConditionalFormatInfo)
		this.x14DataValidations = this.x14DataValidations.map(cloneX14DataValidationInfo)
		this.advancedFilters = this.advancedFilters.map(cloneAdvancedFilterInfo)
		this.autoFilter = this.autoFilter ? cloneAutoFilter(this.autoFilter) : null
		this.preservedAutoFilterSortStateAttributes = this.preservedAutoFilterSortStateAttributes
			? { ...this.preservedAutoFilterSortStateAttributes }
			: null
		this.sortState = this.sortState ? cloneSortState(this.sortState) : null
		this.preservedSortStateAttributes = this.preservedSortStateAttributes
			? { ...this.preservedSortStateAttributes }
			: null
		this.protection = this.protection ? { ...this.protection } : null
		this.protectedRanges = this.protectedRanges.map((range) => ({ ...range }))
		this.pageSetupPr = this.pageSetupPr ? { ...this.pageSetupPr } : null
		this.phoneticPr = this.phoneticPr ? { ...this.phoneticPr } : null
		this._shared = false
	}

	clone(): Sheet {
		const s = new Sheet(this.name, this.id)
		s.cells.copyFrom(this.cells)
		s.merges = this.merges
		s.tables = this.tables
		s.state = this.state
		s.colWidths = this.colWidths
		s.colDefs = this.colDefs
		s.rowHeights = this.rowHeights
		s.rowDefs = this.rowDefs
		s.frozenRows = this.frozenRows
		s.frozenCols = this.frozenCols
		s.codeName = this.codeName
		s.filterMode = this.filterMode
		s.enableFormatConditionsCalculation = this.enableFormatConditionsCalculation
		s.sheetView = this.sheetView
		s.preservedSheetViewAttributes = this.preservedSheetViewAttributes
		s.preservedPaneAttributes = this.preservedPaneAttributes
		s.preservedSheetViewSelections = this.preservedSheetViewSelections
		s.preservedDimensionRef = this.preservedDimensionRef
		s.preservedBlankCells = this.preservedBlankCells
		s.preservedCellMetadata = this.preservedCellMetadata
		s.comments = this.comments
		s.threadedComments = this.threadedComments
		s.hyperlinks = this.hyperlinks
		s.ignoredErrors = this.ignoredErrors
		s.tabColor = this.tabColor
		s.outlinePr = this.outlinePr
		s.sheetFormatPr = this.sheetFormatPr
		s.dataValidationSettings = this.dataValidationSettings
		s.dataValidations = this.dataValidations
		s.conditionalFormats = this.conditionalFormats
		s.imageRefs = this.imageRefs
		s.drawingObjectRefs = this.drawingObjectRefs
		s.sparklineGroups = this.sparklineGroups
		s.x14ConditionalFormats = this.x14ConditionalFormats
		s.x14DataValidations = this.x14DataValidations
		s.advancedFilters = this.advancedFilters
		s.drawingRefs = this.drawingRefs
		s.autoFilter = this.autoFilter
		s.preservedAutoFilterSortStateAttributes = this.preservedAutoFilterSortStateAttributes
		s.sortState = this.sortState
		s.preservedSortStateAttributes = this.preservedSortStateAttributes
		s.protection = this.protection
		s.protectedRanges = this.protectedRanges
		s.pageMargins = this.pageMargins
		s.pageSetup = this.pageSetup
		s.pageSetupPr = this.pageSetupPr
		s.printOptions = this.printOptions
		s.headerFooter = this.headerFooter
		s.phoneticPr = this.phoneticPr
		s.rowBreaks = this.rowBreaks
		s.colBreaks = this.colBreaks
		s.storedFormulaText = this.storedFormulaText
		s.preservedXml = this.preservedXml
		s.preservedExtLst = this.preservedExtLst
		s.preservedCustomSheetViews = this.preservedCustomSheetViews
		s.preservedControlsXml = this.preservedControlsXml
		this._shared = true
		s._shared = true
		return s
	}
}

export function cloneX14ConditionalFormatInfo(
	format: SheetX14ConditionalFormatInfo,
): SheetX14ConditionalFormatInfo {
	return {
		...format,
		formulas: [...format.formulas],
		...(format.preservedRuleAttributes
			? { preservedRuleAttributes: { ...format.preservedRuleAttributes } }
			: {}),
		...(format.preservedRuleChildXml
			? { preservedRuleChildXml: [...format.preservedRuleChildXml] }
			: {}),
		...(format.colorScale
			? {
					colorScale: {
						cfvo: format.colorScale.cfvo.map((entry) => ({ ...entry })),
						colors: format.colorScale.colors.map((entry) => ({ ...entry })),
						...(format.colorScale.preservedAttributes
							? { preservedAttributes: { ...format.colorScale.preservedAttributes } }
							: {}),
						...(format.colorScale.preservedChildXml
							? { preservedChildXml: [...format.colorScale.preservedChildXml] }
							: {}),
					},
				}
			: {}),
		...(format.dataBar
			? {
					dataBar: {
						...format.dataBar,
						cfvo: format.dataBar.cfvo.map((entry) => ({ ...entry })),
						...(format.dataBar.fillColor ? { fillColor: { ...format.dataBar.fillColor } } : {}),
						...(format.dataBar.borderColor
							? { borderColor: { ...format.dataBar.borderColor } }
							: {}),
						...(format.dataBar.negativeFillColor
							? { negativeFillColor: { ...format.dataBar.negativeFillColor } }
							: {}),
						...(format.dataBar.negativeBorderColor
							? { negativeBorderColor: { ...format.dataBar.negativeBorderColor } }
							: {}),
						...(format.dataBar.axisColor ? { axisColor: { ...format.dataBar.axisColor } } : {}),
					},
				}
			: {}),
		...(format.iconSet
			? {
					iconSet: {
						...format.iconSet,
						cfvo: format.iconSet.cfvo.map((entry) => ({ ...entry })),
						...(format.iconSet.icons
							? { icons: format.iconSet.icons.map((entry) => ({ ...entry })) }
							: {}),
					},
				}
			: {}),
	}
}

export function cloneX14DataValidationInfo(
	validation: SheetX14DataValidationInfo,
): SheetX14DataValidationInfo {
	return {
		...validation,
		...(validation.preservedAttributes
			? { preservedAttributes: { ...validation.preservedAttributes } }
			: {}),
		...(validation.preservedChildXml
			? { preservedChildXml: [...validation.preservedChildXml] }
			: {}),
	}
}

export function createSheet(name: string, id?: SheetId): Sheet {
	return new Sheet(name, id)
}

function cloneRangeRef(range: RangeRef): RangeRef {
	return {
		start: { ...range.start },
		end: { ...range.end },
	}
}

function cloneSheetComment(comment: SheetComment): SheetComment {
	return {
		...comment,
		...(comment.legacyDrawing
			? {
					legacyDrawing: {
						...comment.legacyDrawing,
						...(comment.legacyDrawing.anchor ? { anchor: [...comment.legacyDrawing.anchor] } : {}),
					},
				}
			: {}),
	}
}

function cloneAutoFilter(filter: AutoFilter): AutoFilter {
	return {
		...filter,
		columns: filter.columns.map((column) => ({
			...column,
			...(column.values ? { values: [...column.values] } : {}),
			...(column.dateGroupItems
				? { dateGroupItems: column.dateGroupItems.map((item) => ({ ...item })) }
				: {}),
			...(column.customFilters
				? { customFilters: column.customFilters.map((entry) => ({ ...entry })) }
				: {}),
		})),
		...(filter.sortState ? { sortState: cloneSortState(filter.sortState) } : {}),
	}
}

function cloneSortState(sortState: SortState): SortState {
	return {
		...sortState,
		conditions: sortState.conditions.map((condition) => ({ ...condition })),
	}
}

function cloneAdvancedFilterInfo(filter: SheetAdvancedFilterInfo): SheetAdvancedFilterInfo {
	return {
		...filter,
		...(filter.autoFilter ? { autoFilter: cloneAutoFilter(filter.autoFilter) } : {}),
	}
}

function cloneTable(table: Table): Table {
	return {
		...table,
		ref: cloneRangeRef(table.ref),
		columns: table.columns.map((column) => ({ ...column })),
		...(table.autoFilter ? { autoFilter: cloneAutoFilter(table.autoFilter) } : {}),
		...(table.sortState ? { sortState: cloneSortState(table.sortState) } : {}),
		...(table.tableStyleInfo ? { tableStyleInfo: { ...table.tableStyleInfo } } : {}),
	}
}

function cloneImageRef(imageRef: SheetImageRef): SheetImageRef {
	return {
		...imageRef,
		...(imageRef.content ? { content: new Uint8Array(imageRef.content) } : {}),
		...(imageRef.anchor ? { anchor: cloneImageAnchor(imageRef.anchor) } : {}),
	}
}

function cloneDrawingObjectRef(drawingObjectRef: SheetDrawingObjectRef): SheetDrawingObjectRef {
	return {
		...drawingObjectRef,
		...(drawingObjectRef.anchor ? { anchor: cloneImageAnchor(drawingObjectRef.anchor) } : {}),
		...(drawingObjectRef.relIds ? { relIds: [...drawingObjectRef.relIds] } : {}),
		...(drawingObjectRef.relationshipRefs
			? { relationshipRefs: drawingObjectRef.relationshipRefs.map((rel) => ({ ...rel })) }
			: {}),
	}
}

function cloneImageAnchor(anchor: SheetImageAnchor): SheetImageAnchor {
	switch (anchor.kind) {
		case 'oneCell':
			return { ...anchor, from: { ...anchor.from } }
		case 'twoCell':
			return { ...anchor, from: { ...anchor.from }, to: { ...anchor.to } }
		case 'absolute':
			return { ...anchor }
	}
}
