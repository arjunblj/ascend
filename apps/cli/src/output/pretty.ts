import type { CellValue } from '@ascend/schema'
import { formatDisplayCellValue } from '@ascend/sdk'

export function formatCellValue(
	v: CellValue,
	options?: { display?: boolean; dateSystem?: '1900' | '1904' },
): string {
	const display = options?.display ?? false
	const dateSystem = options?.dateSystem ?? '1900'
	if (v.kind === 'date') {
		if (!display) return `[date:${v.serial}]`
		return formatDisplayCellValue(v, { dateSystem })
	}
	return formatDisplayCellValue(v, { dateSystem })
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
