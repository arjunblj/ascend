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
}

export interface SheetPrintOptions {
	readonly gridLines?: boolean
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

export interface SheetHeaderFooter {
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
	readonly hidden?: boolean
	readonly collapsed?: boolean
	readonly outlineLevel?: number
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

export interface SheetDrawingObjectRef {
	readonly drawingPartPath: string
	readonly kind: SheetDrawingObjectKind
	readonly anchor?: SheetImageAnchor
	readonly id?: number
	readonly name?: string
	readonly description?: string
	readonly text?: string
	readonly relIds?: readonly string[]
}

export interface SheetSparklineGroupInfo {
	readonly groupIndex: number
	readonly type?: string
	readonly displayEmptyCellsAs?: string
	readonly dateAxis?: boolean
	readonly markers?: boolean
	readonly highPoint?: boolean
	readonly lowPoint?: boolean
	readonly firstPoint?: boolean
	readonly lastPoint?: boolean
	readonly negative?: boolean
	readonly displayXAxis?: boolean
	readonly colorSeries?: string
	readonly range?: string
	readonly locationRange?: string
	readonly count: number
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
	readonly timePeriod?: string
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
	readonly rules: readonly SheetConditionalFormatRule[]
}

function cloneConditionalFormatRule(rule: SheetConditionalFormatRule): SheetConditionalFormatRule {
	return {
		...rule,
		formulas: [...rule.formulas],
		...(rule.style ? { style: cloneCellStyle(rule.style) } : {}),
		...(rule.colorScale
			? {
					colorScale: {
						cfvo: rule.colorScale.cfvo.map((entry) => ({ ...entry })),
						colors: rule.colorScale.colors.map((entry) => ({ ...entry })),
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
	readonly defaultRowHeight?: number
	readonly defaultColWidth?: number
	readonly outlineLevelRow?: number
	readonly outlineLevelCol?: number
	readonly customHeight?: boolean
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
	sheetView: SheetView | null
	comments: Map<string, SheetComment>
	threadedComments: SheetThreadedComment[]
	hyperlinks: Map<string, SheetHyperlink>
	ignoredErrors: SheetIgnoredError[]
	tabColor: SheetTabColor | null
	outlinePr: SheetOutlinePr | null
	sheetFormatPr: SheetFormatPr | null
	dataValidations: SheetDataValidation[]
	conditionalFormats: SheetConditionalFormat[]
	imageRefs: SheetImageRef[]
	drawingObjectRefs: SheetDrawingObjectRef[]
	sparklineGroups: SheetSparklineGroupInfo[]
	advancedFilters: SheetAdvancedFilterInfo[]
	drawingRefs: SheetDrawingRefs
	autoFilter: AutoFilter | null
	sortState: SortState | null
	protection: SheetProtection | null
	pageMargins: SheetPageMargins | null
	pageSetup: SheetPageSetup | null
	printOptions: SheetPrintOptions | null
	headerFooter: SheetHeaderFooter | null
	rowBreaks: SheetBreak[]
	colBreaks: SheetBreak[]
	/** Exact persisted formula payloads keyed as row:col; writers validate before reuse. */
	storedFormulaText: Map<string, string>
	preservedXml: SheetPreservedXml | null
	preservedExtLst: string | null
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
		this.sheetView = null
		this.comments = new Map()
		this.threadedComments = []
		this.hyperlinks = new Map()
		this.ignoredErrors = []
		this.tabColor = null
		this.outlinePr = null
		this.sheetFormatPr = null
		this.dataValidations = []
		this.conditionalFormats = []
		this.imageRefs = []
		this.drawingObjectRefs = []
		this.sparklineGroups = []
		this.advancedFilters = []
		this.drawingRefs = { hasDrawing: false, hasLegacyDrawing: false }
		this.autoFilter = null
		this.sortState = null
		this.protection = null
		this.pageMargins = null
		this.pageSetup = null
		this.printOptions = null
		this.headerFooter = null
		this.rowBreaks = []
		this.colBreaks = []
		this.storedFormulaText = new Map()
		this.preservedXml = null
		this.preservedExtLst = null
	}

	ensureWritable(): void {
		if (!this._shared) return
		this.merges = this.merges.map(cloneRangeRef)
		this.tables = this.tables.map(cloneTable)
		this.colWidths = new Map(this.colWidths)
		this.colDefs = this.colDefs.map((d) => ({ ...d }))
		this.rowHeights = new Map(this.rowHeights)
		this.rowDefs = new Map([...this.rowDefs.entries()].map(([row, def]) => [row, { ...def }]))
		this.comments = new Map(this.comments)
		this.threadedComments = this.threadedComments.map((comment) => ({ ...comment }))
		this.hyperlinks = new Map(this.hyperlinks)
		this.ignoredErrors = this.ignoredErrors.map((e) => ({ ...e }))
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
		this.sparklineGroups = this.sparklineGroups.map((group) => ({ ...group }))
		this.advancedFilters = this.advancedFilters.map(cloneAdvancedFilterInfo)
		this.autoFilter = this.autoFilter ? cloneAutoFilter(this.autoFilter) : null
		this.sortState = this.sortState ? cloneSortState(this.sortState) : null
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
		s.sheetView = this.sheetView
		s.comments = this.comments
		s.threadedComments = this.threadedComments
		s.hyperlinks = this.hyperlinks
		s.ignoredErrors = this.ignoredErrors
		s.tabColor = this.tabColor
		s.outlinePr = this.outlinePr
		s.sheetFormatPr = this.sheetFormatPr
		s.dataValidations = this.dataValidations
		s.conditionalFormats = this.conditionalFormats
		s.imageRefs = this.imageRefs
		s.drawingObjectRefs = this.drawingObjectRefs
		s.sparklineGroups = this.sparklineGroups
		s.advancedFilters = this.advancedFilters
		s.drawingRefs = this.drawingRefs
		s.autoFilter = this.autoFilter
		s.sortState = this.sortState
		s.protection = this.protection
		s.pageMargins = this.pageMargins
		s.pageSetup = this.pageSetup
		s.printOptions = this.printOptions
		s.headerFooter = this.headerFooter
		s.rowBreaks = this.rowBreaks
		s.colBreaks = this.colBreaks
		s.storedFormulaText = this.storedFormulaText
		s.preservedXml = this.preservedXml
		s.preservedExtLst = this.preservedExtLst
		this._shared = true
		s._shared = true
		return s
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
