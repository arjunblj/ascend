import { indexToColumn } from '@ascend/core'
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

export function escapeDelimitedCell(value: string, delimiter: string): string {
	if (!value.includes(delimiter) && !value.includes('\n') && !value.includes('"')) return value
	return `"${value.replaceAll('"', '""')}"`
}
