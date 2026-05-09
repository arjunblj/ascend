import type { InputValue } from './values.ts'

export interface CellUpdate {
	readonly ref: string
	readonly value: InputValue
}

export interface SortSpec {
	readonly column: string | number
	readonly descending?: boolean
}

export type StyleColorInput =
	| { readonly kind: 'theme'; readonly theme: number; readonly tint?: number }
	| { readonly kind: 'rgb'; readonly rgb: string }
	| { readonly kind: 'indexed'; readonly index: number }
	| { readonly kind: 'auto' }

export type StyleHorizontalAlignInput =
	| 'general'
	| 'left'
	| 'center'
	| 'right'
	| 'fill'
	| 'justify'
	| 'centerContinuous'
	| 'distributed'

export type StyleVerticalAlignInput = 'top' | 'center' | 'bottom' | 'justify' | 'distributed'

export type StyleFillPatternInput =
	| 'none'
	| 'solid'
	| 'darkGray'
	| 'mediumGray'
	| 'lightGray'
	| 'gray125'
	| 'gray0625'
	| 'darkHorizontal'
	| 'darkVertical'
	| 'darkDown'
	| 'darkUp'
	| 'darkGrid'
	| 'darkTrellis'
	| 'lightHorizontal'
	| 'lightVertical'
	| 'lightDown'
	| 'lightUp'
	| 'lightGrid'
	| 'lightTrellis'

export type StyleBorderLineInput =
	| 'none'
	| 'thin'
	| 'medium'
	| 'dashed'
	| 'dotted'
	| 'thick'
	| 'double'
	| 'hair'
	| 'mediumDashed'
	| 'dashDot'
	| 'mediumDashDot'
	| 'dashDotDot'
	| 'mediumDashDotDot'
	| 'slantDashDot'

export interface StyleFontInput {
	readonly name?: string
	readonly size?: number
	readonly bold?: boolean
	readonly italic?: boolean
	readonly underline?: boolean | 'single' | 'double'
	readonly strikethrough?: boolean
	readonly color?: StyleColorInput
}

export interface StyleFillInput {
	readonly pattern?: StyleFillPatternInput
	readonly fgColor?: StyleColorInput
	readonly bgColor?: StyleColorInput
}

export interface StyleBorderEdgeInput {
	readonly style?: StyleBorderLineInput
	readonly color?: StyleColorInput
}

export interface StyleBorderInput {
	readonly top?: StyleBorderEdgeInput
	readonly bottom?: StyleBorderEdgeInput
	readonly left?: StyleBorderEdgeInput
	readonly right?: StyleBorderEdgeInput
	readonly diagonal?: StyleBorderEdgeInput
	readonly diagonalUp?: boolean
	readonly diagonalDown?: boolean
}

export interface StyleAlignmentInput {
	readonly horizontal?: StyleHorizontalAlignInput
	readonly vertical?: StyleVerticalAlignInput
	readonly wrapText?: boolean
	readonly shrinkToFit?: boolean
	readonly textRotation?: number
	readonly indent?: number
	readonly readingOrder?: number
}

export interface StyleInput {
	readonly font?: StyleFontInput
	readonly fill?: StyleFillInput
	readonly border?: StyleBorderInput
	readonly alignment?: StyleAlignmentInput
	readonly numberFormat?: string
	readonly protection?: {
		readonly locked?: boolean
		readonly hidden?: boolean
	}
}

export type Operation =
	| { readonly op: 'setCells'; readonly sheet: string; readonly updates: readonly CellUpdate[] }
	| {
			readonly op: 'setFormula'
			readonly sheet: string
			readonly ref: string
			readonly formula: string
	  }
	| {
			readonly op: 'fillFormula'
			readonly sheet: string
			readonly range: string
			readonly formula: string
	  }
	| {
			readonly op: 'clearRange'
			readonly sheet: string
			readonly range: string
			readonly what: 'values' | 'formulas' | 'styles' | 'all'
	  }
	| {
			readonly op: 'insertRows'
			readonly sheet: string
			readonly at: number
			readonly count: number
	  }
	| {
			readonly op: 'deleteRows'
			readonly sheet: string
			readonly at: number
			readonly count: number
	  }
	| {
			readonly op: 'insertCols'
			readonly sheet: string
			readonly at: number
			readonly count: number
	  }
	| {
			readonly op: 'deleteCols'
			readonly sheet: string
			readonly at: number
			readonly count: number
	  }
	| { readonly op: 'addSheet'; readonly name: string; readonly position?: number }
	| { readonly op: 'deleteSheet'; readonly sheet: string }
	| { readonly op: 'renameSheet'; readonly sheet: string; readonly newName: string }
	| { readonly op: 'moveSheet'; readonly sheet: string; readonly position: number }
	| {
			readonly op: 'createTable'
			readonly sheet: string
			readonly ref: string
			readonly name: string
			readonly hasHeaders: boolean
	  }
	| {
			readonly op: 'appendRows'
			readonly table: string
			readonly rows: readonly (readonly InputValue[])[]
	  }
	| {
			readonly op: 'sortRange'
			readonly sheet: string
			readonly range: string
			readonly by: readonly SortSpec[]
	  }
	| { readonly op: 'mergeCells'; readonly sheet: string; readonly range: string }
	| { readonly op: 'unmergeCells'; readonly sheet: string; readonly range: string }
	| {
			readonly op: 'setColWidth'
			readonly sheet: string
			readonly col: number
			readonly width: number
	  }
	| {
			readonly op: 'setRowHeight'
			readonly sheet: string
			readonly row: number
			readonly height: number
	  }
	| {
			readonly op: 'setComment'
			readonly sheet: string
			readonly ref: string
			readonly text: string
			readonly author?: string
	  }
	| {
			readonly op: 'setHyperlink'
			readonly sheet: string
			readonly ref: string
			readonly url: string
			readonly display?: string
	  }
	| {
			readonly op: 'setNumberFormat'
			readonly sheet: string
			readonly range: string
			readonly format: string
	  }
	| {
			readonly op: 'setDefinedName'
			readonly name: string
			readonly ref: string
			readonly scope?: string
	  }
	| { readonly op: 'deleteDefinedName'; readonly name: string; readonly scope?: string }
	| {
			readonly op: 'setStyle'
			readonly sheet: string
			readonly range: string
			readonly style: StyleInput
	  }
	| {
			readonly op: 'freezePane'
			readonly sheet: string
			readonly row: number
			readonly col: number
	  }
	| {
			readonly op: 'deleteComment'
			readonly sheet: string
			readonly ref: string
	  }
	| {
			readonly op: 'deleteHyperlink'
			readonly sheet: string
			readonly ref: string
	  }
	| {
			readonly op: 'setDataValidation'
			readonly sheet: string
			readonly range: string
			readonly rule: DataValidationRule
	  }
	| {
			readonly op: 'deleteDataValidation'
			readonly sheet: string
			readonly range: string
	  }
	| {
			readonly op: 'setAutoFilter'
			readonly sheet: string
			readonly range: string
	  }
	| { readonly op: 'clearAutoFilter'; readonly sheet: string }
	| {
			readonly op: 'setSheetProtection'
			readonly sheet: string
			readonly password?: string
			readonly options?: SheetProtectionOptions
	  }
	| {
			readonly op: 'setTabColor'
			readonly sheet: string
			readonly color: string
	  }
	| {
			readonly op: 'hideSheet'
			readonly sheet: string
			readonly hidden?: boolean
	  }
	| {
			readonly op: 'hideRows'
			readonly sheet: string
			readonly at: number
			readonly count: number
			readonly hidden?: boolean
	  }
	| {
			readonly op: 'hideCols'
			readonly sheet: string
			readonly at: number
			readonly count: number
			readonly hidden?: boolean
	  }
	| {
			readonly op: 'copySheet'
			readonly sheet: string
			readonly newName: string
			readonly position?: number
	  }
	| {
			readonly op: 'setConditionalFormat'
			readonly sheet: string
			readonly range: string
			readonly rule: ConditionalFormatRule
	  }
	| {
			readonly op: 'deleteConditionalFormat'
			readonly sheet: string
			readonly range: string
	  }
	| {
			readonly op: 'setPageSetup'
			readonly sheet: string
			readonly setup: PageSetupInput
	  }
	| {
			readonly op: 'setPrintArea'
			readonly sheet: string
			readonly range: string
	  }
	| {
			readonly op: 'copyRange'
			readonly sheet: string
			readonly source: string
			readonly target: string
	  }
	| {
			readonly op: 'moveRange'
			readonly sheet: string
			readonly source: string
			readonly target: string
	  }
	| {
			readonly op: 'groupRows'
			readonly sheet: string
			readonly from: number
			readonly to: number
			readonly collapsed?: boolean
			readonly summaryBelow?: boolean
	  }
	| {
			readonly op: 'groupCols'
			readonly sheet: string
			readonly from: number
			readonly to: number
			readonly collapsed?: boolean
			readonly summaryRight?: boolean
	  }
	| {
			readonly op: 'setRichText'
			readonly sheet: string
			readonly ref: string
			readonly runs: readonly {
				readonly text: string
				readonly bold?: boolean
				readonly italic?: boolean
				readonly underline?: boolean
				readonly color?: string
				readonly size?: number
			}[]
	  }
	| { readonly op: 'setWorkbookProtection'; readonly protection: WorkbookProtectionInput }
	| { readonly op: 'deleteTable'; readonly table: string }
	| { readonly op: 'renameTable'; readonly table: string; readonly newName: string }
	| { readonly op: 'resizeTable'; readonly table: string; readonly ref: string }
	| {
			readonly op: 'replaceImage'
			readonly sheet: string
			readonly contentBase64: string
			readonly contentType: string
			readonly targetPath?: string
			readonly relId?: string
			readonly name?: string
			readonly imageIndex?: number
	  }
	| {
			readonly op: 'insertImage'
			readonly sheet: string
			readonly contentBase64: string
			readonly contentType: string
			readonly targetPath?: string
			readonly drawingPartPath?: string
			readonly relId?: string
			readonly name?: string
			readonly description?: string
			readonly anchor?: ImageAnchorInput
	  }
	| {
			readonly op: 'deleteImage'
			readonly sheet: string
			readonly targetPath?: string
			readonly relId?: string
			readonly name?: string
			readonly imageIndex?: number
	  }
	| {
			readonly op: 'setChartSeriesSource'
			readonly seriesIndex: number
			readonly partPath?: string
			readonly sheet?: string
			readonly chartIndex?: number
			readonly nameRef?: string
			readonly categoryRef?: string
			readonly valueRef?: string
	  }
	| {
			readonly op: 'setPivotCache'
			readonly cacheId?: number
			readonly partPath?: string
			readonly pivotTable?: string
			readonly sourceSheet?: string
			readonly sourceRef?: string
			readonly refreshOnLoad?: boolean
			readonly enableRefresh?: boolean
			readonly invalid?: boolean
			readonly saveData?: boolean
	  }
	| {
			readonly op: 'rewriteExternalLink'
			readonly partPath?: string
			readonly relId?: string
			readonly linkRelId?: string
			readonly target?: string
			readonly newTarget: string
			readonly targetMode?: string
	  }

export type ImageAnchorInput =
	| {
			readonly kind: 'oneCell'
			readonly from: ImageAnchorMarkerInput
			readonly cx?: number
			readonly cy?: number
	  }
	| {
			readonly kind: 'twoCell'
			readonly from: ImageAnchorMarkerInput
			readonly to: ImageAnchorMarkerInput
			readonly editAs?: string
	  }
	| {
			readonly kind: 'absolute'
			readonly x: number
			readonly y: number
			readonly cx?: number
			readonly cy?: number
	  }

export interface ImageAnchorMarkerInput {
	readonly col: number
	readonly row: number
	readonly colOff?: number
	readonly rowOff?: number
}

export interface WorkbookProtectionInput {
	readonly lockStructure?: boolean
	readonly lockWindows?: boolean
	readonly lockRevision?: boolean
	readonly workbookPassword?: string
	readonly revisionsPassword?: string
	readonly workbookAlgorithmName?: string
	readonly workbookHashValue?: string
	readonly workbookSaltValue?: string
	readonly workbookSpinCount?: number
	readonly revisionsAlgorithmName?: string
	readonly revisionsHashValue?: string
	readonly revisionsSaltValue?: string
	readonly revisionsSpinCount?: number
}

export interface ConditionalFormatRule {
	readonly type:
		| 'cellIs'
		| 'expression'
		| 'colorScale'
		| 'dataBar'
		| 'iconSet'
		| 'top10'
		| 'aboveAverage'
		| 'duplicateValues'
		| 'containsText'
	readonly operator?:
		| 'greaterThan'
		| 'lessThan'
		| 'equal'
		| 'between'
		| 'greaterThanOrEqual'
		| 'lessThanOrEqual'
		| 'notEqual'
		| 'notBetween'
	readonly formula?: string
	readonly formula2?: string
	readonly priority?: number
	readonly stopIfTrue?: boolean
	readonly style?: StyleInput
}

export interface PageSetupInput {
	readonly orientation?: 'portrait' | 'landscape'
	readonly paperSize?: number
	readonly scale?: number
	readonly fitToWidth?: number
	readonly fitToHeight?: number
	readonly margins?: {
		readonly left?: number
		readonly right?: number
		readonly top?: number
		readonly bottom?: number
		readonly header?: number
		readonly footer?: number
	}
}

export interface DataValidationRule {
	readonly type: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom'
	readonly formula1?: string
	readonly formula2?: string
	readonly operator?:
		| 'between'
		| 'notBetween'
		| 'equal'
		| 'notEqual'
		| 'greaterThan'
		| 'lessThan'
		| 'greaterThanOrEqual'
		| 'lessThanOrEqual'
	readonly allowBlank?: boolean
	readonly showErrorMessage?: boolean
	readonly errorTitle?: string
	readonly errorMessage?: string
	readonly showInputMessage?: boolean
	readonly promptTitle?: string
	readonly prompt?: string
}

export interface SheetProtectionOptions {
	readonly formatCells?: boolean
	readonly formatColumns?: boolean
	readonly formatRows?: boolean
	readonly insertColumns?: boolean
	readonly insertRows?: boolean
	readonly deleteColumns?: boolean
	readonly deleteRows?: boolean
	readonly sort?: boolean
	readonly autoFilter?: boolean
}
