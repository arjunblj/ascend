import type { Table } from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { autoFilterXml, sortStateXml } from './filtering.ts'

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
	if (table.dxfId !== undefined) attrs.push(`dxfId="${table.dxfId}"`)
	if (table.headerRowDxfId !== undefined) attrs.push(`headerRowDxfId="${table.headerRowDxfId}"`)
	if (table.dataDxfId !== undefined) attrs.push(`dataDxfId="${table.dataDxfId}"`)
	if (table.totalsRowDxfId !== undefined) attrs.push(`totalsRowDxfId="${table.totalsRowDxfId}"`)
	if (table.headerRowBorderDxfId !== undefined) {
		attrs.push(`headerRowBorderDxfId="${table.headerRowBorderDxfId}"`)
	}

	const parts: string[] = [XML_HEADER, `<table ${attrs.join(' ')}>`]
	if (table.autoFilter) {
		parts.push(autoFilterXml(table.autoFilter))
	} else if (table.hasHeaders) {
		parts.push(`<autoFilter ref="${ref}"/>`)
	}
	if (table.sortState) {
		parts.push(sortStateXml(table.sortState))
	}
	parts.push(`<tableColumns count="${table.columns.length}">`)
	for (let index = 0; index < table.columns.length; index++) {
		const column = table.columns[index]
		if (!column) continue
		const columnAttrs = [`id="${column.id ?? index + 1}"`, `name="${escapeXml(column.name)}"`]
		if (column.totalsRowFunction) {
			columnAttrs.push(`totalsRowFunction="${escapeXml(column.totalsRowFunction)}"`)
		}
		if (column.totalsRowLabel) {
			columnAttrs.push(`totalsRowLabel="${escapeXml(column.totalsRowLabel)}"`)
		}
		if (column.dataDxfId !== undefined) columnAttrs.push(`dataDxfId="${column.dataDxfId}"`)
		if (column.headerRowDxfId !== undefined) {
			columnAttrs.push(`headerRowDxfId="${column.headerRowDxfId}"`)
		}
		if (column.totalsRowDxfId !== undefined) {
			columnAttrs.push(`totalsRowDxfId="${column.totalsRowDxfId}"`)
		}
		parts.push(`<tableColumn ${columnAttrs.join(' ')}>`)
		if (column.formula) {
			parts.push(`<calculatedColumnFormula>${escapeXml(column.formula)}</calculatedColumnFormula>`)
		}
		if (column.totalsRowFormula) {
			parts.push(`<totalsRowFormula>${escapeXml(column.totalsRowFormula)}</totalsRowFormula>`)
		}
		parts.push('</tableColumn>')
	}
	parts.push('</tableColumns>')
	if (table.tableStyleInfo) {
		const styleAttrs: string[] = []
		if (table.tableStyleInfo.name) styleAttrs.push(`name="${escapeXml(table.tableStyleInfo.name)}"`)
		if (table.tableStyleInfo.showFirstColumn !== undefined) {
			styleAttrs.push(`showFirstColumn="${table.tableStyleInfo.showFirstColumn ? '1' : '0'}"`)
		}
		if (table.tableStyleInfo.showLastColumn !== undefined) {
			styleAttrs.push(`showLastColumn="${table.tableStyleInfo.showLastColumn ? '1' : '0'}"`)
		}
		if (table.tableStyleInfo.showRowStripes !== undefined) {
			styleAttrs.push(`showRowStripes="${table.tableStyleInfo.showRowStripes ? '1' : '0'}"`)
		}
		if (table.tableStyleInfo.showColumnStripes !== undefined) {
			styleAttrs.push(`showColumnStripes="${table.tableStyleInfo.showColumnStripes ? '1' : '0'}"`)
		}
		parts.push(`<tableStyleInfo ${styleAttrs.join(' ')}/>`)
	}
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
