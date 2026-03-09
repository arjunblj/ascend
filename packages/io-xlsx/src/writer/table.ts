import type { Table } from '@ascend/core'
import { escapeXml } from '../xml.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

export function buildTableXml(table: Table, tableNumber: number): string {
	const ref = `${toCellRef(table.ref.start.row, table.ref.start.col)}:${toCellRef(
		table.ref.end.row,
		table.ref.end.col,
	)}`
	const attrs = [
		`xmlns="${NS}"`,
		`id="${tableNumber}"`,
		`name="${escapeXml(table.name)}"`,
		`displayName="${escapeXml(table.name)}"`,
		`ref="${ref}"`,
		`headerRowCount="${table.hasHeaders ? 1 : 0}"`,
		`totalsRowCount="${table.hasTotals ? 1 : 0}"`,
	]

	const parts: string[] = [XML_HEADER, `<table ${attrs.join(' ')}>`]
	if (table.hasHeaders) {
		parts.push(`<autoFilter ref="${ref}"/>`)
	}
	parts.push(`<tableColumns count="${table.columns.length}">`)
	for (let index = 0; index < table.columns.length; index++) {
		const column = table.columns[index]
		if (!column) continue
		parts.push(`<tableColumn id="${index + 1}" name="${escapeXml(column.name)}">`)
		if (column.formula) {
			parts.push(`<calculatedColumnFormula>${escapeXml(column.formula)}</calculatedColumnFormula>`)
		}
		parts.push('</tableColumn>')
	}
	parts.push('</tableColumns>')
	parts.push('</table>')
	return parts.join('')
}

function toCellRef(row: number, col: number): string {
	return `${toColumnLabel(col)}${row + 1}`
}

function toColumnLabel(colIndex: number): string {
	let n = colIndex + 1
	let label = ''
	while (n > 0) {
		const rem = (n - 1) % 26
		label = String.fromCharCode(65 + rem) + label
		n = Math.floor((n - 1) / 26)
	}
	return label
}
