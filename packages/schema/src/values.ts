export type ExcelError =
	| '#NULL!'
	| '#DIV/0!'
	| '#VALUE!'
	| '#REF!'
	| '#NAME?'
	| '#NUM!'
	| '#N/A'
	| '#GETTING_DATA'
	| '#SPILL!'
	| '#CALC!'

export interface RichTextRun {
	readonly text: string
	readonly bold?: boolean
	readonly italic?: boolean
	readonly underline?: boolean
	readonly strikethrough?: boolean
	readonly fontName?: string
	readonly fontSize?: number
	readonly color?: string
}

export type ScalarCellValue =
	| { readonly kind: 'empty' }
	| { readonly kind: 'number'; readonly value: number }
	| { readonly kind: 'string'; readonly value: string }
	| { readonly kind: 'boolean'; readonly value: boolean }
	| { readonly kind: 'error'; readonly value: ExcelError }
	| { readonly kind: 'date'; readonly serial: number }
	| { readonly kind: 'richText'; readonly runs: readonly RichTextRun[] }

export interface ArrayValue {
	readonly kind: 'array'
	readonly rows: readonly (readonly ScalarCellValue[])[]
}

export type CellValue = ScalarCellValue | ArrayValue

export type InputValue = string | number | boolean | null | Date

export const EMPTY: CellValue = { kind: 'empty' } as const

export function numberValue(v: number): CellValue {
	return { kind: 'number', value: v }
}

export function stringValue(v: string): CellValue {
	return { kind: 'string', value: v }
}

export function booleanValue(v: boolean): CellValue {
	return { kind: 'boolean', value: v }
}

export function errorValue(v: ExcelError): CellValue {
	return { kind: 'error', value: v }
}

export function arrayValue(rows: readonly (readonly ScalarCellValue[])[]): CellValue {
	return { kind: 'array', rows }
}

export function isError(v: CellValue): v is CellValue & { kind: 'error' } {
	return v.kind === 'error'
}

export function isEmpty(v: CellValue): v is CellValue & { kind: 'empty' } {
	return v.kind === 'empty'
}

export function isArrayValue(v: CellValue): v is ArrayValue {
	return v.kind === 'array'
}

export function topLeftScalar(v: CellValue): ScalarCellValue {
	if (v.kind !== 'array') return v
	const firstRow = v.rows[0]
	return firstRow?.[0] ?? (EMPTY as ScalarCellValue)
}
