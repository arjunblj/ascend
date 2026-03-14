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

export type RichTextColor =
	| { readonly kind: 'rgb'; readonly rgb: string }
	| { readonly kind: 'theme'; readonly theme: number; readonly tint?: number }
	| { readonly kind: 'indexed'; readonly index: number }

export interface RichTextRun {
	readonly text: string
	readonly bold?: boolean
	readonly italic?: boolean
	readonly underline?: boolean
	readonly strikethrough?: boolean
	readonly fontName?: string
	readonly fontSize?: number
	readonly color?: string | RichTextColor
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

const dateCache = new Map<number, CellValue>()
const DATE_CACHE_MAX = 256

export function dateValue(serial: number): CellValue {
	let cached = dateCache.get(serial)
	if (cached) return cached
	cached = { kind: 'date', serial }
	if (dateCache.size >= DATE_CACHE_MAX) {
		const firstKey = dateCache.keys().next().value
		if (firstKey !== undefined) dateCache.delete(firstKey)
	}
	dateCache.set(serial, cached)
	return cached
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

function richTextColorsEqual(
	a: string | RichTextColor | undefined,
	b: string | RichTextColor | undefined,
): boolean {
	if (a === b) return true
	if (a === undefined || b === undefined) return false
	if (typeof a === 'string' || typeof b === 'string') return a === b
	if (a.kind !== b.kind) return false
	if (a.kind === 'rgb' && b.kind === 'rgb') return a.rgb === b.rgb
	if (a.kind === 'theme' && b.kind === 'theme') return a.theme === b.theme && a.tint === b.tint
	if (a.kind === 'indexed' && b.kind === 'indexed') return a.index === b.index
	return false
}

export function valuesEqual(a: CellValue, b: CellValue): boolean {
	if (a === b) return true
	a = topLeftScalar(a)
	b = topLeftScalar(b)
	if (a.kind !== b.kind) return false
	switch (a.kind) {
		case 'empty':
			return true
		case 'number':
			return a.value === (b as typeof a).value
		case 'string':
			return a.value === (b as typeof a).value
		case 'boolean':
			return a.value === (b as typeof a).value
		case 'error':
			return a.value === (b as typeof a).value
		case 'date':
			return a.serial === (b as typeof a).serial
		case 'richText': {
			const runsA = a.runs
			const runsB = (b as typeof a).runs
			if (runsA.length !== runsB.length) return false
			for (let index = 0; index < runsA.length; index++) {
				const left = runsA[index]
				const right = runsB[index]
				if (
					left?.text !== right?.text ||
					left?.bold !== right?.bold ||
					left?.italic !== right?.italic ||
					left?.underline !== right?.underline ||
					left?.strikethrough !== right?.strikethrough ||
					left?.fontName !== right?.fontName ||
					left?.fontSize !== right?.fontSize ||
					!richTextColorsEqual(left?.color, right?.color)
				) {
					return false
				}
			}
			return true
		}
	}
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
