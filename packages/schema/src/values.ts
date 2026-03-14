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

const TRUE_VALUE: CellValue = { kind: 'boolean', value: true } as const
const FALSE_VALUE: CellValue = { kind: 'boolean', value: false } as const
const ZERO_VALUE: CellValue = { kind: 'number', value: 0 } as const
const ONE_VALUE: CellValue = { kind: 'number', value: 1 } as const
const EMPTY_STRING_VALUE: CellValue = { kind: 'string', value: '' } as const

const ERROR_CACHE: Record<string, CellValue> = {
	'#NULL!': { kind: 'error', value: '#NULL!' } as const,
	'#DIV/0!': { kind: 'error', value: '#DIV/0!' } as const,
	'#VALUE!': { kind: 'error', value: '#VALUE!' } as const,
	'#REF!': { kind: 'error', value: '#REF!' } as const,
	'#NAME?': { kind: 'error', value: '#NAME?' } as const,
	'#NUM!': { kind: 'error', value: '#NUM!' } as const,
	'#N/A': { kind: 'error', value: '#N/A' } as const,
	'#GETTING_DATA': { kind: 'error', value: '#GETTING_DATA' } as const,
	'#SPILL!': { kind: 'error', value: '#SPILL!' } as const,
	'#CALC!': { kind: 'error', value: '#CALC!' } as const,
}

const SMALL_INT_CACHE_MIN = -128
const SMALL_INT_CACHE_MAX = 512
const SMALL_INT_CACHE: CellValue[] = new Array(SMALL_INT_CACHE_MAX - SMALL_INT_CACHE_MIN + 1)
for (let i = SMALL_INT_CACHE_MIN; i <= SMALL_INT_CACHE_MAX; i++) {
	SMALL_INT_CACHE[i - SMALL_INT_CACHE_MIN] = { kind: 'number', value: i } as const
}

export function numberValue(v: number): CellValue {
	if (v === 0) return ZERO_VALUE
	if (v === 1) return ONE_VALUE
	if (v === (v | 0) && v >= SMALL_INT_CACHE_MIN && v <= SMALL_INT_CACHE_MAX) {
		return SMALL_INT_CACHE[v - SMALL_INT_CACHE_MIN] as CellValue
	}
	return { kind: 'number', value: v }
}

export function stringValue(v: string): CellValue {
	if (v === '') return EMPTY_STRING_VALUE
	return { kind: 'string', value: v }
}

export function booleanValue(v: boolean): CellValue {
	return v ? TRUE_VALUE : FALSE_VALUE
}

export function errorValue(v: ExcelError): CellValue {
	return ERROR_CACHE[v] ?? { kind: 'error', value: v }
}

export function dateValue(serial: number): CellValue {
	return { kind: 'date', serial }
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

export function coerceCellValueToString(v: CellValue): string {
	if (v.kind === 'array') v = topLeftScalar(v)
	switch (v.kind) {
		case 'number':
			return String(v.value)
		case 'string':
			return v.value
		case 'boolean':
			return v.value ? 'TRUE' : 'FALSE'
		case 'empty':
			return ''
		case 'date':
			return String(v.serial)
		case 'error':
			return v.value
		case 'richText':
			return v.runs.map((r) => r.text).join('')
	}
}
