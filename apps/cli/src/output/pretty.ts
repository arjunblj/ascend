import type { CellValue } from '@ascend/schema'

export function formatCellValue(
	v: CellValue,
	options?: { display?: boolean; dateSystem?: '1900' | '1904' },
): string {
	const display = options?.display ?? false
	const dateSystem = options?.dateSystem ?? '1900'
	switch (v.kind) {
		case 'empty':
			return ''
		case 'number':
			return String(v.value)
		case 'string':
			return v.value
		case 'boolean':
			return v.value ? 'TRUE' : 'FALSE'
		case 'error':
			return v.value
		case 'date':
			if (!display) return `[date:${v.serial}]`
			return formatDisplayDate(v.serial, dateSystem)
		case 'richText':
			return v.runs.map((r: { text: string }) => r.text).join('')
		default:
			return ''
	}
}

function formatDisplayDate(serial: number, dateSystem: '1900' | '1904'): string {
	const parts = serialToDateParts(Math.floor(serial), dateSystem)
	if (!parts) return `[date:${serial}]`
	const year = String(parts.year).padStart(4, '0')
	const month = String(parts.month).padStart(2, '0')
	const day = String(parts.day).padStart(2, '0')
	return `${year}-${month}-${day}`
}

function serialToDateParts(
	serial: number,
	dateSystem: '1900' | '1904',
): { year: number; month: number; day: number } | null {
	const msPerDay = 86_400_000
	const makeUtc = (year: number, month: number, day: number) => {
		const date = new Date(Date.UTC(year, month - 1, day))
		if (year >= 0 && year < 100) date.setUTCFullYear(year)
		return date
	}
	if (dateSystem === '1904') {
		if (serial < 0) return null
		const epoch = makeUtc(1904, 1, 1).getTime()
		const date = new Date(epoch + serial * msPerDay)
		return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
	}
	if (serial < 1) return null
	if (serial === 60) return { year: 1900, month: 2, day: 29 }
	const epoch = makeUtc(1900, 1, 1).getTime()
	const days = serial < 60 ? serial - 1 : serial - 2
	const date = new Date(epoch + days * msPerDay)
	return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
}

export function table(headers: string[], rows: string[][]): string {
	const widths = headers.map((h, i) => {
		let max = h.length
		for (const row of rows) {
			const cell = row[i] ?? ''
			if (cell.length > max) max = cell.length
		}
		return max
	})

	const pad = (s: string, w: number) => s + ' '.repeat(Math.max(0, w - s.length))
	const line = (cells: string[]) => cells.map((c, i) => pad(c, widths[i] ?? 0)).join('  ')

	const parts: string[] = []
	parts.push(line(headers))
	parts.push(widths.map((w) => '─'.repeat(w)).join('──'))
	for (const row of rows) {
		parts.push(line(row))
	}
	return parts.join('\n')
}

export function heading(text: string): string {
	return `\n${text}\n${'─'.repeat(text.length)}`
}

export function bullet(label: string, value: string | number): string {
	return `  ${label}: ${value}`
}
