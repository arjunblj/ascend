import type { Cell, RangeRef, SheetColDef, StyleId } from '@ascend/core'
import { columnToIndex, parseRange, Sheet } from '@ascend/core'
import type { FormulaCellRef } from '@ascend/formulas'
import { parseFormula, printFormula, rewriteRefs } from '@ascend/formulas'
import type { CellValue, ExcelError } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import type { Relationship } from './relationships.ts'

const CELL_REF_RE = /^([A-Za-z]+)(\d+)$/

export interface SheetParseContext {
	readonly sharedStrings: CellValue[]
	readonly styleIds: StyleId[]
	readonly isDateFormat: boolean[]
	readonly relationships?: readonly Relationship[]
}

export function parseSheet(name: string, xml: string, ctx: SheetParseContext): Sheet {
	const doc = parseXml(xml)
	const ws = doc.worksheet as XmlNode | undefined
	const sheet = new Sheet(name)

	if (!ws) return sheet

	parseSheetViews(ws, sheet)
	parseCols(ws, sheet)
	parseSheetData(ws, sheet, ctx)
	parseMergeCells(ws, sheet)
	parseAutoFilter(ws, sheet)
	parsePageMargins(ws, sheet)
	parsePageSetup(ws, sheet)
	parsePrintOptions(ws, sheet)
	parseHeaderFooter(ws, sheet)
	parseIgnoredErrors(ws, sheet)
	parseHyperlinks(ws, sheet, ctx.relationships ?? [])

	return sheet
}

function parseSheetData(ws: XmlNode, sheet: Sheet, ctx: SheetParseContext): void {
	const sd = ws.sheetData as XmlNode | undefined
	if (!sd) return
	const sharedFormulaMasters = new Map<string, { formula: string; row: number; col: number }>()

	for (const row of asArray<XmlNode>(sd.row as XmlNode | XmlNode[])) {
		const rowIndex = numAttr(row, 'r')
		const rowHeight = numAttr(row, 'ht')
		if (rowIndex !== undefined && rowHeight !== undefined && attr(row, 'customHeight') === '1') {
			sheet.rowHeights.set(rowIndex - 1, rowHeight)
		}
		for (const c of asArray<XmlNode>(row.c as XmlNode | XmlNode[])) {
			const ref = attr(c, 'r')
			if (!ref) continue

			const pos = parseCellRef(ref)
			if (!pos) continue

			const cell = parseCellValue(c, ctx, pos.row, pos.col, sharedFormulaMasters)
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

function parseCellValue(
	c: XmlNode,
	ctx: SheetParseContext,
	row: number,
	col: number,
	sharedFormulaMasters: Map<string, { formula: string; row: number; col: number }>,
): Cell | undefined {
	const type = attr(c, 't')
	const styleIdx = numAttr(c, 's') ?? 0
	const styleId = ctx.styleIds[styleIdx] ?? (0 as StyleId)
	const formula = parseFormulaText(c.f, row, col, sharedFormulaMasters)
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

function parseFormulaText(
	formulaNode: unknown,
	row: number,
	col: number,
	sharedFormulaMasters: Map<string, { formula: string; row: number; col: number }>,
): string | null {
	if (formulaNode === undefined || formulaNode === null) return null
	if (
		typeof formulaNode === 'string' ||
		typeof formulaNode === 'number' ||
		typeof formulaNode === 'boolean'
	) {
		return String(formulaNode)
	}
	if (typeof formulaNode === 'object') {
		const node = formulaNode as XmlNode
		const sharedIndex = attr(node, 'si')
		const formulaType = attr(node, 't')
		const text = node['#text']
		if (formulaType === 'shared' && sharedIndex) {
			if (text !== undefined && text !== null) {
				const formula = String(text)
				sharedFormulaMasters.set(sharedIndex, { formula, row, col })
				return formula
			}
			const master = sharedFormulaMasters.get(sharedIndex)
			if (!master) return null
			return translateSharedFormula(master.formula, master.row, master.col, row, col)
		}
		return text !== undefined && text !== null ? String(text) : null
	}
	return null
}

function translateSharedFormula(
	formula: string,
	masterRow: number,
	masterCol: number,
	row: number,
	col: number,
): string | null {
	const parsed = parseFormula(formula)
	if (!parsed.ok) return null
	const rowDelta = row - masterRow
	const colDelta = col - masterCol
	const rewritten = rewriteRefs(parsed.value, (ref: FormulaCellRef) => ({
		...ref,
		row: ref.rowAbsolute ? ref.row : ref.row + rowDelta,
		col: ref.colAbsolute ? ref.col : ref.col + colDelta,
	}))
	return printFormula(rewritten)
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

function parseSheetViews(ws: XmlNode, sheet: Sheet): void {
	const viewsNode = ws.sheetViews as XmlNode | undefined
	if (!viewsNode) return
	const firstView = asArray<XmlNode>(viewsNode.sheetView as XmlNode | XmlNode[])[0]
	if (!firstView) return
	const pane = firstView.pane as XmlNode | undefined
	if (!pane) return

	const ySplit = numAttr(pane, 'ySplit')
	if (ySplit !== undefined) sheet.frozenRows = Math.trunc(ySplit)
	const xSplit = numAttr(pane, 'xSplit')
	if (xSplit !== undefined) sheet.frozenCols = Math.trunc(xSplit)
}

function parseCols(ws: XmlNode, sheet: Sheet): void {
	const colsNode = ws.cols as XmlNode | undefined
	if (!colsNode) return

	for (const col of asArray<XmlNode>(colsNode.col as XmlNode | XmlNode[])) {
		const min = numAttr(col, 'min')
		const max = numAttr(col, 'max')
		const width = numAttr(col, 'width')
		if (min === undefined || max === undefined) continue

		const parsed: {
			min: number
			max: number
			width?: number
			style?: number
			hidden?: boolean
			bestFit?: boolean
			collapsed?: boolean
			outlineLevel?: number
			customWidth?: boolean
		} = {
			min: min - 1,
			max: max - 1,
		}
		if (width !== undefined) {
			parsed.width = width
			for (let idx = min; idx <= max; idx++) {
				sheet.colWidths.set(idx - 1, width)
			}
		}
		const style = numAttr(col, 'style')
		if (style !== undefined) parsed.style = style
		const hidden = readBoolAttribute(col, 'hidden')
		if (hidden !== undefined) parsed.hidden = hidden
		const bestFit = readBoolAttribute(col, 'bestFit')
		if (bestFit !== undefined) parsed.bestFit = bestFit
		const collapsed = readBoolAttribute(col, 'collapsed')
		if (collapsed !== undefined) parsed.collapsed = collapsed
		const customWidth = readBoolAttribute(col, 'customWidth')
		if (customWidth !== undefined) parsed.customWidth = customWidth
		const outlineLevel = numAttr(col, 'outlineLevel')
		if (outlineLevel !== undefined) parsed.outlineLevel = outlineLevel
		sheet.colDefs.push(parsed as SheetColDef)
	}
}

function parseAutoFilter(ws: XmlNode, sheet: Sheet): void {
	const autoFilter = ws.autoFilter as XmlNode | undefined
	if (!autoFilter) return
	const ref = attr(autoFilter, 'ref')
	if (ref) sheet.autoFilter = ref
}

function parsePageMargins(ws: XmlNode, sheet: Sheet): void {
	const margins = ws.pageMargins as XmlNode | undefined
	if (!margins) return

	const parsed: Record<string, number> = {}
	setIfDefined(parsed, 'left', numAttr(margins, 'left'))
	setIfDefined(parsed, 'right', numAttr(margins, 'right'))
	setIfDefined(parsed, 'top', numAttr(margins, 'top'))
	setIfDefined(parsed, 'bottom', numAttr(margins, 'bottom'))
	setIfDefined(parsed, 'header', numAttr(margins, 'header'))
	setIfDefined(parsed, 'footer', numAttr(margins, 'footer'))
	sheet.pageMargins = parsed
}

function parsePageSetup(ws: XmlNode, sheet: Sheet): void {
	const pageSetup = ws.pageSetup as XmlNode | undefined
	if (!pageSetup) return

	const parsed: Record<string, unknown> = {}
	const orientation = attr(pageSetup, 'orientation')
	if (orientation) parsed.orientation = orientation
	const paperSize = numAttr(pageSetup, 'paperSize')
	if (paperSize !== undefined) parsed.paperSize = paperSize
	const scale = numAttr(pageSetup, 'scale')
	if (scale !== undefined) parsed.scale = scale
	const fitToWidth = numAttr(pageSetup, 'fitToWidth')
	if (fitToWidth !== undefined) parsed.fitToWidth = fitToWidth
	const fitToHeight = numAttr(pageSetup, 'fitToHeight')
	if (fitToHeight !== undefined) parsed.fitToHeight = fitToHeight
	sheet.pageSetup = parsed
}

function parsePrintOptions(ws: XmlNode, sheet: Sheet): void {
	const printOptions = ws.printOptions as XmlNode | undefined
	if (!printOptions) return

	const parsed: Record<string, boolean> = {}
	setIfDefined(parsed, 'gridLines', readBoolAttribute(printOptions, 'gridLines'))
	setIfDefined(parsed, 'headings', readBoolAttribute(printOptions, 'headings'))
	setIfDefined(parsed, 'horizontalCentered', readBoolAttribute(printOptions, 'horizontalCentered'))
	setIfDefined(parsed, 'verticalCentered', readBoolAttribute(printOptions, 'verticalCentered'))
	sheet.printOptions = parsed
}

function parseHeaderFooter(ws: XmlNode, sheet: Sheet): void {
	const headerFooter = ws.headerFooter as XmlNode | undefined
	if (!headerFooter) return

	const parsed: Record<string, string> = {}
	setIfDefined(parsed, 'oddHeader', readNodeText(headerFooter.oddHeader))
	setIfDefined(parsed, 'oddFooter', readNodeText(headerFooter.oddFooter))
	setIfDefined(parsed, 'evenHeader', readNodeText(headerFooter.evenHeader))
	setIfDefined(parsed, 'evenFooter', readNodeText(headerFooter.evenFooter))
	setIfDefined(parsed, 'firstHeader', readNodeText(headerFooter.firstHeader))
	setIfDefined(parsed, 'firstFooter', readNodeText(headerFooter.firstFooter))
	sheet.headerFooter = parsed
}

function parseIgnoredErrors(ws: XmlNode, sheet: Sheet): void {
	const ignoredErrors = ws.ignoredErrors as XmlNode | undefined
	if (!ignoredErrors) return

	for (const ignoredError of asArray<XmlNode>(ignoredErrors.ignoredError as XmlNode | XmlNode[])) {
		const sqref = attr(ignoredError, 'sqref')
		if (sqref) sheet.ignoredErrors.push(sqref)
	}
}

function parseHyperlinks(ws: XmlNode, sheet: Sheet, relationships: readonly Relationship[]): void {
	const hyperlinks = ws.hyperlinks as XmlNode | undefined
	if (!hyperlinks) return
	const relMap = new Map(relationships.map((rel) => [rel.id, rel]))

	for (const hyperlink of asArray<XmlNode>(hyperlinks.hyperlink as XmlNode | XmlNode[])) {
		const ref = attr(hyperlink, 'ref')
		if (!ref) continue
		const relId = attr(hyperlink, 'r:id') ?? attr(hyperlink, 'id')
		const rel = relId ? relMap.get(relId) : undefined
		const parsed: Record<string, string> = {}
		setIfDefined(parsed, 'target', rel?.target)
		setIfDefined(parsed, 'location', attr(hyperlink, 'location'))
		setIfDefined(parsed, 'display', attr(hyperlink, 'display'))
		setIfDefined(parsed, 'tooltip', attr(hyperlink, 'tooltip'))
		sheet.hyperlinks.set(ref, parsed)
	}
}

function readNodeText(node: unknown): string | undefined {
	if (node === undefined || node === null) return undefined
	if (typeof node === 'string' || typeof node === 'number' || typeof node === 'boolean') {
		return String(node)
	}
	if (typeof node === 'object') {
		const text = (node as XmlNode)['#text']
		return text !== undefined && text !== null ? String(text) : undefined
	}
	return undefined
}

function readBoolAttribute(node: XmlNode, name: string): boolean | undefined {
	const value = attr(node, name)
	if (value === undefined) return undefined
	return value === '1' || value === 'true'
}

function setIfDefined<T>(target: Record<string, T>, key: string, value: T | undefined): void {
	if (value !== undefined) target[key] = value
}
