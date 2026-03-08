import type { CellValue } from '@ascend/schema'

export function formatCellValue(v: CellValue): string {
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
			return `[date:${v.serial}]`
		case 'richText':
			return v.runs.map((r: { text: string }) => r.text).join('')
		default:
			return ''
	}
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
