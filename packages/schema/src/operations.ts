import type { InputValue } from './values.ts'

export interface CellUpdate {
	readonly ref: string
	readonly value: InputValue
}

export interface SortSpec {
	readonly column: string | number
	readonly descending?: boolean
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
	| { readonly op: 'deleteDefinedName'; readonly name: string }
	| {
			readonly op: 'freezePane'
			readonly sheet: string
			readonly row: number
			readonly col: number
	  }
