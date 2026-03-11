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

	clone(): Sheet {
		const clone = new Sheet(this.name, this.id)
		clone.cells.copyFrom(this.cells)
		clone.merges.push(...this.merges.map(cloneRangeRef))
		clone.tables.push(...this.tables.map(cloneTable))
		clone.state = this.state
		clone.colWidths.clear()
		for (const [key, value] of this.colWidths) clone.colWidths.set(key, value)
		clone.colDefs.push(...this.colDefs.map((colDef) => ({ ...colDef })))
		clone.rowHeights.clear()
		for (const [key, value] of this.rowHeights) clone.rowHeights.set(key, value)
		clone.frozenRows = this.frozenRows
		clone.frozenCols = this.frozenCols
		for (const [key, value] of this.comments) clone.comments.set(key, { ...value })
		for (const [key, value] of this.hyperlinks) clone.hyperlinks.set(key, { ...value })
		clone.ignoredErrors.push(...this.ignoredErrors.map((ignoredError) => ({ ...ignoredError })))
		clone.tabColor = this.tabColor ? { ...this.tabColor } : null
		clone.sheetFormatPr = this.sheetFormatPr ? { ...this.sheetFormatPr } : null
		clone.dataValidations.push(
			...this.dataValidations.map((dataValidation) => ({ ...dataValidation })),
		)
		clone.conditionalFormats.push(
			...this.conditionalFormats.map((conditionalFormat) => ({
				...conditionalFormat,
				rules: conditionalFormat.rules.map((rule) => ({
					...rule,
					formulas: [...rule.formulas],
					...(rule.style ? { style: cloneCellStyle(rule.style) } : {}),
				})),
			})),
		)
		clone.imageRefs.push(...this.imageRefs.map(cloneImageRef))
		clone.drawingRefs = { ...this.drawingRefs }
		clone.autoFilter = this.autoFilter ? cloneAutoFilter(this.autoFilter) : null
		clone.protection = this.protection ? { ...this.protection } : null
		clone.pageMargins = this.pageMargins ? { ...this.pageMargins } : null
		clone.pageSetup = this.pageSetup ? { ...this.pageSetup } : null
		clone.printOptions = this.printOptions ? { ...this.printOptions } : null
		clone.headerFooter = this.headerFooter ? { ...this.headerFooter } : null
		clone.preservedXml = this.preservedXml ? { ...this.preservedXml } : null
		clone.preservedExtLst = this.preservedExtLst
		return clone
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

function cloneCellStyle(style: CellStyle): CellStyle {
	return {
		...(style.font
			? { font: { ...style.font, ...(style.font.color ? { color: { ...style.font.color } } : {}) } }
			: {}),
		...(style.fill
			? {
					fill: {
						...style.fill,
						...(style.fill.fgColor ? { fgColor: { ...style.fill.fgColor } } : {}),
						...(style.fill.bgColor ? { bgColor: { ...style.fill.bgColor } } : {}),
					},
				}
			: {}),
		...(style.border
			? {
					border: {
						...style.border,
						...(style.border.top
							? {
									top: {
										...style.border.top,
										...(style.border.top.color ? { color: { ...style.border.top.color } } : {}),
									},
								}
							: {}),
						...(style.border.bottom
							? {
									bottom: {
										...style.border.bottom,
										...(style.border.bottom.color
											? { color: { ...style.border.bottom.color } }
											: {}),
									},
								}
							: {}),
						...(style.border.left
							? {
									left: {
										...style.border.left,
										...(style.border.left.color ? { color: { ...style.border.left.color } } : {}),
									},
								}
							: {}),
						...(style.border.right
							? {
									right: {
										...style.border.right,
										...(style.border.right.color ? { color: { ...style.border.right.color } } : {}),
									},
								}
							: {}),
						...(style.border.diagonal
							? {
									diagonal: {
										...style.border.diagonal,
										...(style.border.diagonal.color
											? { color: { ...style.border.diagonal.color } }
											: {}),
									},
								}
							: {}),
					},
				}
			: {}),
		...(style.alignment ? { alignment: { ...style.alignment } } : {}),
		...(style.numberFormat ? { numberFormat: style.numberFormat } : {}),
		...(style.protection ? { protection: { ...style.protection } } : {}),
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
		...(filter.sortState
			? {
					sortState: {
						...filter.sortState,
						conditions: filter.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
	}
}

function cloneTable(table: Table): Table {
	return {
		...table,
		ref: cloneRangeRef(table.ref),
		columns: table.columns.map((column) => ({ ...column })),
		...(table.autoFilter ? { autoFilter: cloneAutoFilter(table.autoFilter) } : {}),
		...(table.sortState
			? {
					sortState: {
						...table.sortState,
						conditions: table.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
		...(table.tableStyleInfo ? { tableStyleInfo: { ...table.tableStyleInfo } } : {}),
	}
}

function cloneImageRef(imageRef: SheetImageRef): SheetImageRef {
	return {
		...imageRef,
		...(imageRef.anchor ? { anchor: cloneImageAnchor(imageRef.anchor) } : {}),
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
