import type { InputValue } from './values.ts'

export interface CellUpdate {
	readonly ref: string
	readonly value: InputValue
}

export interface SortSpec {
	readonly column: string | number
	readonly descending?: boolean
}

export interface StyleInput {
	readonly font?: Record<string, unknown>
	readonly fill?: Record<string, unknown>
	readonly border?: Record<string, unknown>
	readonly alignment?: Record<string, unknown>
	readonly numberFormat?: string
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
