import { type CellStyle, indexToColumn } from '@ascend/core'
import { serialToDate } from '@ascend/formulas'
import type { CellValue } from '@ascend/schema'

export type ExportFormat = 'csv' | 'tsv' | 'json' | 'xlsx' | 'xlsm'

/** Convert row/col indices to A1 ref (e.g. row=0, col=0 → "A1"). */
export function toA1Ref(row: number, col: number): string {
	return `${indexToColumn(col)}${row + 1}`
}

export function normalizeExportFormat(format: string): ExportFormat | null {
	switch (format.toLowerCase()) {
		case 'csv':
		case 'tsv':
		case 'json':
		case 'xlsx':
		case 'xlsm':
			return format.toLowerCase() as ExportFormat
		default:
			return null
	}
}

export function inferExportFormat(path: string): ExportFormat | null {
	const ext = path.split('.').pop()?.toLowerCase() ?? ''
	return normalizeExportFormat(ext)
}

export function ensureOutputExtension(output: string, format: ExportFormat): string {
	return output.endsWith(`.${format}`) ? output : `${output.replace(/\.[^.]+$/, '')}.${format}`
}

export interface FormatDisplayOptions {
	dateSystem?: '1900' | '1904'
	locale?: string
}

export function formatDisplayCellValue(value: CellValue, options?: FormatDisplayOptions): string {
	switch (value.kind) {
		case 'empty':
			return ''
		case 'number':
			return String(value.value)
		case 'string':
			return value.value
		case 'boolean':
			return value.value ? 'TRUE' : 'FALSE'
		case 'error':
			return value.value
		case 'date': {
			const dateSystem = options?.dateSystem ?? '1900'
			const parts = serialToDate(Math.floor(value.serial), dateSystem)
			if (!parts) return `[date:${value.serial}]`
			return `${String(parts.year).padStart(4, '0')}-${String(parts.month).padStart(2, '0')}-${String(parts.day).padStart(2, '0')}`
		}
		case 'richText':
			return value.runs.map((run) => run.text).join('')
		case 'array':
			return value.rows
				.map((row) => row.map((v) => formatDisplayCellValue(v, options)).join(','))
				.join(';')
	}
}

export function formatStyledDisplayCellValue(
	value: CellValue,
	style?: CellStyle,
	options?: FormatDisplayOptions,
): string {
	const format = style?.numberFormat
	if (!format || format.toLowerCase() === 'general') return formatDisplayCellValue(value, options)
	switch (value.kind) {
		case 'number':
			return formatNumberForDisplay(value.value, format, options)
		case 'date':
			return formatSerialForDisplay(value.serial, format, options)
		case 'array':
			return value.rows
				.map((row) => row.map((v) => formatStyledDisplayCellValue(v, style, options)).join(','))
				.join(';')
		default:
			return formatDisplayCellValue(value, options)
	}
}

function formatNumberForDisplay(
	value: number,
	format: string,
	options?: FormatDisplayOptions,
): string {
	if (isDateFormat(format)) return formatSerialForDisplay(value, format, options)
	const section = formatSection(format, value)
	const cleaned = cleanNumberFormatSection(section)
	if (cleaned === '@') return String(value)
	if (cleaned.includes('%')) return formatPercent(value, cleaned, options)
	if (!/[0#]/.test(cleaned)) return String(value)
	return formatDecimal(value, cleaned, options)
}

function formatSerialForDisplay(
	serial: number,
	format: string,
	options?: FormatDisplayOptions,
): string {
	const parts = serialToDate(Math.floor(serial), options?.dateSystem ?? '1900')
	if (!parts) return `[date:${serial}]`
	const cleaned = cleanDateFormatSection(formatSection(format, serial))
	if (/^d{1,2}\/m{1,2}\/y{2,4}$/i.test(cleaned)) {
		return `${padDate(parts.day, cleaned.startsWith('dd') ? 2 : 1)}/${padDate(
			parts.month,
			cleaned.includes('/mm/') ? 2 : 1,
		)}/${formatYear(parts.year, cleaned.endsWith('yyyy') ? 4 : 2)}`
	}
	if (/^m{1,2}\/d{1,2}\/y{2,4}$/i.test(cleaned)) {
		return `${padDate(parts.month, cleaned.startsWith('mm') ? 2 : 1)}/${padDate(
			parts.day,
			cleaned.includes('/dd/') ? 2 : 1,
		)}/${formatYear(parts.year, cleaned.endsWith('yyyy') ? 4 : 2)}`
	}
	if (/^y{2,4}-m{1,2}-d{1,2}$/i.test(cleaned)) {
		return `${formatYear(parts.year, cleaned.startsWith('yyyy') ? 4 : 2)}-${padDate(
			parts.month,
			cleaned.includes('-mm-') ? 2 : 1,
		)}-${padDate(parts.day, cleaned.endsWith('dd') ? 2 : 1)}`
	}
	return formatDisplayCellValue({ kind: 'date', serial }, options)
}

function formatPercent(value: number, section: string, options?: FormatDisplayOptions): string {
	const decimals = decimalPlaces(section.replace(/%.*$/, ''))
	return `${formatLocaleNumber(value * 100, decimals, section, options)}%`
}

function formatDecimal(value: number, section: string, options?: FormatDisplayOptions): string {
	const negative = value < 0
	const abs = Math.abs(value)
	const decimals = decimalPlaces(section)
	const firstPlaceholder = section.search(/[0#]/)
	const lastPlaceholder = Math.max(section.lastIndexOf('0'), section.lastIndexOf('#'))
	const prefix = firstPlaceholder >= 0 ? normalizeAffix(section.slice(0, firstPlaceholder)) : ''
	const suffix = lastPlaceholder >= 0 ? normalizeAffix(section.slice(lastPlaceholder + 1)) : ''
	const number = formatLocaleNumber(abs, decimals, section, options)
	const sign =
		negative &&
		!hasExplicitNegativeSection(section) &&
		!prefix.includes('(') &&
		!suffix.includes(')')
			? '-'
			: ''
	return `${sign}${prefix}${number}${suffix}`
}

function formatLocaleNumber(
	value: number,
	decimals: number,
	section: string,
	options?: FormatDisplayOptions,
): string {
	return value.toLocaleString(options?.locale ?? 'en-US', {
		useGrouping: section.includes(','),
		minimumFractionDigits: decimals,
		maximumFractionDigits: decimals,
	})
}

function formatSection(format: string, value: number): string {
	const sections = format.split(';')
	if (value < 0 && sections[1]) return sections[1]
	if (value === 0 && sections[2]) return sections[2]
	return sections[0] ?? format
}

function isDateFormat(format: string): boolean {
	const cleaned = cleanDateFormatSection(formatSection(format, 1))
	return /[ymd]/i.test(cleaned) && !cleaned.includes('%')
}

function cleanNumberFormatSection(section: string): string {
	return section
		.replace(/\[[^\]]+\]/g, '')
		.replace(/"([^"]*)"/g, '$1')
		.replace(/_./g, '')
		.replace(/\*./g, '')
		.replace(/\\/g, '')
		.trim()
}

function cleanDateFormatSection(section: string): string {
	return cleanNumberFormatSection(section).toLowerCase()
}

function decimalPlaces(section: string): number {
	const decimal = section.indexOf('.')
	if (decimal < 0) return 0
	const tail = section.slice(decimal + 1)
	const match = /^[0#]+/.exec(tail)
	return match?.[0].length ?? 0
}

function normalizeAffix(affix: string): string {
	return affix.replace(/[#,0.]+/g, '').replace(/\s+/g, '')
}

function hasExplicitNegativeSection(section: string): boolean {
	return section.includes('-') || section.includes('(') || section.includes(')')
}

function padDate(value: number, width: number): string {
	return width === 2 ? String(value).padStart(2, '0') : String(value)
}

function formatYear(year: number, width: 2 | 4): string {
	return width === 2 ? String(year % 100).padStart(2, '0') : String(year).padStart(4, '0')
}

export function escapeDelimitedCell(value: string, delimiter: string): string {
	if (!value.includes(delimiter) && !value.includes('\n') && !value.includes('"')) return value
	return `"${value.replaceAll('"', '""')}"`
}
