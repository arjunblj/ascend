import type { Cell, RangeRef, StyleId } from '@ascend/core'
import { columnToIndex, parseRange, Sheet } from '@ascend/core'
import type { CellValue, ExcelError } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'

const CELL_REF_RE = /^([A-Za-z]+)(\d+)$/

export interface SheetParseContext {
	readonly sharedStrings: CellValue[]
	readonly styleIds: StyleId[]
	readonly isDateFormat: boolean[]
}

export function parseSheet(name: string, xml: string, ctx: SheetParseContext): Sheet {
	const doc = parseXml(xml)
	const ws = doc.worksheet as XmlNode | undefined
	const sheet = new Sheet(name)

	if (!ws) return sheet

	parseSheetData(ws, sheet, ctx)
	parseMergeCells(ws, sheet)

	return sheet
}

function parseSheetData(ws: XmlNode, sheet: Sheet, ctx: SheetParseContext): void {
	const sd = ws.sheetData as XmlNode | undefined
	if (!sd) return

	for (const row of asArray<XmlNode>(sd.row as XmlNode | XmlNode[])) {
		for (const c of asArray<XmlNode>(row.c as XmlNode | XmlNode[])) {
			const ref = attr(c, 'r')
			if (!ref) continue

			const pos = parseCellRef(ref)
			if (!pos) continue

			const cell = parseCellValue(c, ctx)
			if (cell) {
				sheet.cells.set(pos.row, pos.col, cell)
			}
		}
	}
}

function parseCellRef(ref: string): { row: number; col: number } | undefined {
	const m = CELL_REF_RE.exec(ref)
	const colStr = m?.[1]
	const rowStr = m?.[2]
	if (!colStr || !rowStr) return undefined
	return {
		row: Number.parseInt(rowStr, 10) - 1,
		col: columnToIndex(colStr.toUpperCase()),
	}
}

function parseCellValue(c: XmlNode, ctx: SheetParseContext): Cell | undefined {
	const type = attr(c, 't')
	const styleIdx = numAttr(c, 's') ?? 0
	const styleId = ctx.styleIds[styleIdx] ?? (0 as StyleId)
	const formula = c.f !== undefined ? String(c.f) : null
	const rawValue = c.v

	let value: CellValue

	if (type === 's') {
		const idx = typeof rawValue === 'number' ? rawValue : Number(rawValue)
		const entry = ctx.sharedStrings[idx]
		value = entry ?? stringValue('')
	} else if (type === 'b') {
		const bv = rawValue === 1 || rawValue === true || rawValue === '1'
		value = booleanValue(bv)
	} else if (type === 'e') {
		value = errorValue(String(rawValue ?? '#VALUE!') as ExcelError)
	} else if (type === 'str') {
		value = stringValue(String(rawValue ?? ''))
	} else if (type === 'inlineStr') {
		value = parseInlineString(c)
	} else if (rawValue !== undefined && rawValue !== null && rawValue !== '') {
		const num = Number(rawValue)
		if (Number.isNaN(num)) {
			value = stringValue(String(rawValue))
		} else if (ctx.isDateFormat[styleIdx]) {
			value = { kind: 'date', serial: num }
		} else {
			value = numberValue(num)
		}
	} else if (formula) {
		value = EMPTY
	} else {
		return undefined
	}

	return { value, formula, styleId }
}

function parseInlineString(c: XmlNode): CellValue {
	const is = c.is as XmlNode | undefined
	if (!is || typeof is !== 'object') return stringValue('')

	if (is.t !== undefined) {
		return stringValue(String(is.t))
	}

	const runs = asArray<XmlNode>(is.r as XmlNode | XmlNode[])
	const text = runs.map((r) => (r.t !== undefined ? String(r.t) : '')).join('')
	return stringValue(text)
}

function parseMergeCells(ws: XmlNode, sheet: Sheet): void {
	const mc = ws.mergeCells as XmlNode | undefined
	if (!mc) return

	for (const m of asArray<XmlNode>(mc.mergeCell as XmlNode | XmlNode[])) {
		const ref = attr(m, 'ref')
		if (!ref) continue
		try {
			const range: RangeRef = parseRange(ref)
			sheet.merges.push(range)
		} catch {
			// skip invalid merge refs
		}
	}
}
