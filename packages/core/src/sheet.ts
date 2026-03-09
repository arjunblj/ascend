import type { AutoFilter } from './filter.ts'
import { createSheetId, type SheetId } from './ids.ts'
import type { RangeRef } from './refs.ts'
import { SparseGrid } from './sparse-grid.ts'
import type { CellStyle } from './style.ts'
import type { Table } from './table.ts'

export type SheetState = 'visible' | 'hidden' | 'veryHidden'

export interface SheetComment {
	readonly text: string
	readonly author?: string
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

export interface SheetPreservedXml {
	readonly xml: string
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
	readonly anchor?: SheetImageAnchor
	readonly name?: string
	readonly description?: string
}

export interface SheetDataValidation {
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
}

export interface SheetConditionalFormat {
	readonly sqref: string
	readonly rules: readonly SheetConditionalFormatRule[]
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

export class Sheet {
	readonly id: SheetId
	name: string
	readonly cells: SparseGrid
	readonly merges: RangeRef[]
	readonly tables: Table[]
	state: SheetState
	readonly colWidths: Map<number, number>
	readonly colDefs: SheetColDef[]
	readonly rowHeights: Map<number, number>
	frozenRows: number
	frozenCols: number
	readonly comments: Map<string, SheetComment>
	readonly hyperlinks: Map<string, SheetHyperlink>
	readonly ignoredErrors: SheetIgnoredError[]
	tabColor: SheetTabColor | null
	sheetFormatPr: SheetFormatPr | null
	readonly dataValidations: SheetDataValidation[]
	readonly conditionalFormats: SheetConditionalFormat[]
	readonly imageRefs: SheetImageRef[]
	drawingRefs: SheetDrawingRefs
	autoFilter: AutoFilter | null
	protection: SheetProtection | null
	pageMargins: SheetPageMargins | null
	pageSetup: SheetPageSetup | null
	printOptions: SheetPrintOptions | null
	headerFooter: SheetHeaderFooter | null
	preservedXml: SheetPreservedXml | null
	preservedExtLst: string | null

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
		this.frozenRows = 0
		this.frozenCols = 0
		this.comments = new Map()
		this.hyperlinks = new Map()
		this.ignoredErrors = []
		this.tabColor = null
		this.sheetFormatPr = null
		this.dataValidations = []
		this.conditionalFormats = []
		this.imageRefs = []
		this.drawingRefs = { hasDrawing: false, hasLegacyDrawing: false }
		this.autoFilter = null
		this.protection = null
		this.pageMargins = null
		this.pageSetup = null
		this.printOptions = null
		this.headerFooter = null
		this.preservedXml = null
		this.preservedExtLst = null
	}
}

export function createSheet(name: string, id?: SheetId): Sheet {
	return new Sheet(name, id)
}
