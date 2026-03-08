import type { Cell, Sheet } from '@ascend/core'
import { indexToColumn } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import { escapeXml } from '../xml.ts'
import type { SharedStringTable } from './shared-strings.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

export function buildSheetXml(
	sheet: Sheet,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
): string {
	const parts: string[] = [XML_HEADER, `<worksheet xmlns="${NS}">`]

	parts.push('<sheetData>')

	const rows = new Map<number, Array<{ col: number; cell: Cell }>>()
	for (const [row, col, cell] of sheet.cells.iterate()) {
		let rowCells = rows.get(row)
		if (!rowCells) {
			rowCells = []
			rows.set(row, rowCells)
		}
		rowCells.push({ col, cell })
	}

	const sortedRows = [...rows.entries()].sort((a, b) => a[0] - b[0])

	for (const [row, cells] of sortedRows) {
		cells.sort((a, b) => a.col - b.col)
		parts.push(`<row r="${row + 1}">`)
		for (const { col, cell } of cells) {
			const ref = `${indexToColumn(col)}${row + 1}`
			parts.push(cellXml(ref, cell, ssTable, xfMap))
		}
		parts.push('</row>')
	}

	parts.push('</sheetData>')

	if (sheet.merges.length > 0) {
		parts.push(`<mergeCells count="${sheet.merges.length}">`)
		for (const merge of sheet.merges) {
			const s = `${indexToColumn(merge.start.col)}${merge.start.row + 1}`
			const e = `${indexToColumn(merge.end.col)}${merge.end.row + 1}`
			parts.push(`<mergeCell ref="${s}:${e}"/>`)
		}
		parts.push('</mergeCells>')
	}

	parts.push('</worksheet>')
	return parts.join('')
}

function cellXml(
	ref: string,
	cell: Cell,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
): string {
	const xfIdx = xfMap.get(cell.styleId as number) ?? 0

	if (cell.formula) {
		return formulaCellXml(ref, cell, xfIdx)
	}
	return regularCellXml(ref, cell, ssTable, xfIdx)
}

function formulaCellXml(ref: string, cell: Cell, xfIdx: number): string {
	const attrs: string[] = [`r="${ref}"`]
	if (xfIdx !== 0) attrs.push(`s="${xfIdx}"`)

	const { typeAttr, valueStr } = formulaValueAttrs(cell.value)
	if (typeAttr) attrs.push(`t="${typeAttr}"`)

	const parts: string[] = [`<c ${attrs.join(' ')}>`]
	parts.push(`<f>${escapeXml(cell.formula ?? '')}</f>`)
	if (valueStr !== undefined) parts.push(`<v>${valueStr}</v>`)
	parts.push('</c>')
	return parts.join('')
}

function regularCellXml(
	ref: string,
	cell: Cell,
	ssTable: SharedStringTable,
	xfIdx: number,
): string {
	const attrs: string[] = [`r="${ref}"`]
	if (xfIdx !== 0) attrs.push(`s="${xfIdx}"`)

	const { typeAttr, valueStr } = regularValueAttrs(cell.value, ssTable)
	if (typeAttr) attrs.push(`t="${typeAttr}"`)

	if (valueStr === undefined) return `<c ${attrs.join(' ')}/>`
	return `<c ${attrs.join(' ')}><v>${valueStr}</v></c>`
}

function formulaValueAttrs(value: CellValue): {
	typeAttr: string | undefined
	valueStr: string | undefined
} {
	switch (value.kind) {
		case 'string':
			return { typeAttr: 'str', valueStr: escapeXml(value.value) }
		case 'number':
			return { typeAttr: undefined, valueStr: String(value.value) }
		case 'boolean':
			return { typeAttr: 'b', valueStr: value.value ? '1' : '0' }
		case 'error':
			return { typeAttr: 'e', valueStr: escapeXml(value.value) }
		case 'date':
			return { typeAttr: undefined, valueStr: String(value.serial) }
		case 'empty':
			return { typeAttr: undefined, valueStr: undefined }
		case 'richText':
			return {
				typeAttr: 'str',
				valueStr: escapeXml(value.runs.map((r) => r.text).join('')),
			}
	}
}

function regularValueAttrs(
	value: CellValue,
	ssTable: SharedStringTable,
): { typeAttr: string | undefined; valueStr: string | undefined } {
	switch (value.kind) {
		case 'string':
		case 'richText': {
			const idx = ssTable.getIndex(value)
			return { typeAttr: 's', valueStr: idx !== undefined ? String(idx) : '0' }
		}
		case 'number':
			return { typeAttr: undefined, valueStr: String(value.value) }
		case 'boolean':
			return { typeAttr: 'b', valueStr: value.value ? '1' : '0' }
		case 'error':
			return { typeAttr: 'e', valueStr: escapeXml(value.value) }
		case 'date':
			return { typeAttr: undefined, valueStr: String(value.serial) }
		case 'empty':
			return { typeAttr: undefined, valueStr: undefined }
	}
}
