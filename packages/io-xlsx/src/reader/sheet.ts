import type {
	AutoFilter,
	Cell,
	CellStyle,
	RangeRef,
	SheetColDef,
	SheetConditionalFormat,
	SheetConditionalFormatRule,
	SheetDataValidation,
	StyleId,
} from '@ascend/core'
import { columnToIndex, parseRange, Sheet } from '@ascend/core'
import type { FormulaCellRef } from '@ascend/formulas'
import { parseFormula, printFormula, rewriteRefs } from '@ascend/formulas'
import type { CellValue, ExcelError } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { parseAutoFilterNode } from './filtering.ts'
import type { Relationship } from './relationships.ts'
import type { SharedStringResolver } from './shared-strings.ts'

const CELL_REF_RE = /^([A-Za-z]+)(\d+)$/
const SMALL_NUMBER_RANGE_START = -128
const SMALL_NUMBER_RANGE_END = 512
const SHEET_DATA_RE = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>|<sheetData\b[^>]*\/>/
const ROW_RE = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g
const CELL_RE = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g
const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g
const TEXT_NODE_RE =
	/<([A-Za-z_][\w:.-]*)\b([^>]*)>([\s\S]*?)<\/\1>|<([A-Za-z_][\w:.-]*)\b([^>]*)\/>/g

export interface SheetParseContext {
	readonly sharedStrings: SharedStringResolver
	readonly styleIds: StyleId[]
	readonly isDateFormat: boolean[]
	readonly differentialStyles?: readonly CellStyle[]
	readonly relationships?: readonly Relationship[]
	readonly valuePool?: ValueInternPool
	readonly valuesOnly?: boolean
	readonly formulaOnly?: boolean
}

export class ValueInternPool {
	private readonly strings = new Map<string, string>()
	private readonly stringValues = new Map<string, CellValue>()
	private readonly booleans = new Map<boolean, CellValue>([
		[true, booleanValue(true)],
		[false, booleanValue(false)],
	])
	private readonly errors = new Map<ExcelError, CellValue>()
	private readonly smallNumbers = buildSmallNumberCache()

	internString(value: string): string {
		const cached = this.strings.get(value)
		if (cached !== undefined) return cached
		this.strings.set(value, value)
		return value
	}

	internValue(value: CellValue): CellValue {
		switch (value.kind) {
			case 'empty':
				return EMPTY
			case 'boolean': {
				const cached = this.booleans.get(value.value)
				if (cached) return cached
				return value
			}
			case 'error': {
				const cached = this.errors.get(value.value)
				if (cached) return cached
				this.errors.set(value.value, value)
				return value
			}
			case 'number': {
				const cached = this.smallNumbers.get(value.value)
				if (cached) return cached
				return value
			}
			case 'string': {
				const text = this.internString(value.value)
				const cached = this.stringValues.get(text)
				if (cached) return cached
				const interned: CellValue = { kind: 'string', value: text }
				this.stringValues.set(text, interned)
				return interned
			}
			case 'richText':
				return {
					kind: 'richText',
					runs: value.runs.map((run) => ({
						...run,
						text: this.internString(run.text),
						...(run.fontName ? { fontName: this.internString(run.fontName) } : {}),
						...(run.color ? { color: this.internString(run.color) } : {}),
					})),
				}
			default:
				return value
		}
	}
}

function buildSmallNumberCache(): Map<number, CellValue> {
	const cache = new Map<number, CellValue>()
	for (let value = SMALL_NUMBER_RANGE_START; value <= SMALL_NUMBER_RANGE_END; value++) {
		cache.set(value, numberValue(value))
	}
	return cache
}

export function parseSheet(name: string, xml: string, ctx: SheetParseContext): Sheet {
	const doc = parseXml(stripSheetDataForDom(xml))
	const ws = doc.worksheet as XmlNode | undefined
	const sheet = new Sheet(name)

	if (!ws) return sheet

	parseSheetPr(ws, sheet)
	parseSheetFormatPr(ws, sheet)
	parseSheetViews(ws, sheet)
	parseCols(ws, sheet)
	parseSheetDataXml(xml, sheet, ctx)
	parseMergeCells(ws, sheet)
	parseDrawingRefs(ws, sheet)
	parseAutoFilter(ws, sheet)
	parseSheetProtection(ws, sheet)
	parsePageMargins(ws, sheet)
	parsePageSetup(ws, sheet)
	parsePrintOptions(ws, sheet)
	parseHeaderFooter(ws, sheet)
	parseIgnoredErrors(ws, sheet)
	if (!ctx.valuesOnly && !ctx.formulaOnly) {
		parseHyperlinks(ws, sheet, ctx.relationships ?? [], ctx.valuePool)
		parseConditionalFormatting(ws, sheet, ctx.differentialStyles ?? [], ctx.valuePool)
		parseDataValidations(ws, sheet, ctx.valuePool)
		extractExtLst(xml, sheet)
	}

	return sheet
}

function stripSheetDataForDom(xml: string): string {
	return xml.replace(SHEET_DATA_RE, '<sheetData/>')
}

function parseSheetDataXml(xml: string, sheet: Sheet, ctx: SheetParseContext): void {
	const sheetData = extractSheetDataContent(xml)
	if (!sheetData) return
	const sharedFormulaMasters = new Map<
		string,
		{ formula: string; row: number; col: number; ref?: string }
	>()

	for (const rowMatch of sheetData.matchAll(ROW_RE)) {
		const rowAttrs = parseRawAttributes(rowMatch[1] ?? rowMatch[3] ?? '')
		const rowIndex = numAttr(rowAttrs, 'r')
		const rowHeight = numAttr(rowAttrs, 'ht')
		if (
			rowIndex !== undefined &&
			rowHeight !== undefined &&
			attr(rowAttrs, 'customHeight') === '1'
		) {
			sheet.rowHeights.set(rowIndex - 1, rowHeight)
		}

		const rowInner = rowMatch[2] ?? ''
		for (const cellMatch of rowInner.matchAll(CELL_RE)) {
			const rawAttrs = cellMatch[1] ?? cellMatch[3] ?? ''
			const cellNode = buildCellNode(rawAttrs, cellMatch[2] ?? '')
			const ref = attr(cellNode, 'r')
			if (!ref) continue
			const pos = parseCellRef(ref)
			if (!pos) continue
			const cell = parseCellValue(cellNode, ctx, pos.row, pos.col, sharedFormulaMasters)
			if (cell) {
				sheet.cells.setResolved(
					pos.row,
					pos.col,
					cell.value,
					cell.formula,
					cell.styleId,
					cell.formulaInfo,
				)
			}
		}
	}
}

function extractSheetDataContent(xml: string): string | null {
	const match = SHEET_DATA_RE.exec(xml)
	return match?.[1] ?? null
}

function _parseSheetData(ws: XmlNode, sheet: Sheet, ctx: SheetParseContext): void {
	const sd = ws.sheetData as XmlNode | undefined
	if (!sd) return
	const sharedFormulaMasters = new Map<
		string,
		{ formula: string; row: number; col: number; ref?: string }
	>()

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
				sheet.cells.setResolved(
					pos.row,
					pos.col,
					cell.value,
					cell.formula,
					cell.styleId,
					cell.formulaInfo,
				)
			}
		}
	}
}

function buildCellNode(rawAttrs: string, innerXml: string): XmlNode {
	const node = parseRawAttributes(rawAttrs)
	const valueNode = extractTextNode(innerXml, 'v')
	if (valueNode) node.v = valueNode.text
	const formulaNode = extractTextNode(innerXml, 'f')
	if (formulaNode) {
		const fNode = parseRawAttributes(formulaNode.attrs)
		if (formulaNode.text !== undefined) fNode['#text'] = formulaNode.text
		node.f = fNode
	}
	const inlineStringNode = extractTextNode(innerXml, 'is')
	if (inlineStringNode) {
		node.is = buildInlineStringNode(inlineStringNode.text ?? '')
	}
	return node
}

function buildInlineStringNode(innerXml: string): XmlNode {
	const runs: XmlNode[] = []
	for (const run of extractNodes(innerXml, 'r')) {
		const text = extractTextNode(run.text ?? '', 't')
		runs.push({ t: text?.text ?? '' })
	}
	if (runs.length > 0) return { r: runs }
	const directText = extractTextNode(innerXml, 't')
	return directText ? { t: directText.text ?? '' } : {}
}

function parseRawAttributes(rawAttrs: string): XmlNode {
	const node: XmlNode = {}
	for (const match of rawAttrs.matchAll(ATTR_RE)) {
		const key = match[1]
		const value = match[2]
		if (!key || value === undefined) continue
		node[`@_${key}`] = decodeXmlText(value)
	}
	return node
}

function extractTextNode(
	xml: string,
	tagName: string,
): { attrs: string; text?: string } | undefined {
	for (const node of extractNodes(xml, tagName)) {
		return {
			attrs: node.attrs,
			...(node.text !== undefined ? { text: node.text } : {}),
		}
	}
	return undefined
}

function extractNodes(xml: string, tagName: string): Array<{ attrs: string; text?: string }> {
	const nodes: Array<{ attrs: string; text?: string }> = []
	for (const match of xml.matchAll(TEXT_NODE_RE)) {
		const openTag = match[1] ?? match[4]
		if (openTag !== tagName) continue
		const attrs = match[2] ?? match[5] ?? ''
		const text = match[3]
		nodes.push({
			attrs,
			...(text !== undefined ? { text: decodeXmlText(text) } : {}),
		})
	}
	return nodes
}

function decodeXmlText(text: string): string {
	return text
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&')
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
	const pool = ctx.valuePool
	const type = attr(c, 't')
	const styleIdx = numAttr(c, 's') ?? 0
	const rawValue = c.v
	const styleId = ctx.valuesOnly ? (0 as StyleId) : (ctx.styleIds[styleIdx] ?? (0 as StyleId))
	const formulaSpec =
		ctx.valuesOnly && rawValue !== undefined && rawValue !== null && rawValue !== ''
			? { text: null, info: undefined }
			: parseFormulaText(c.f, row, col, sharedFormulaMasters, pool)
	const formula = formulaSpec.text

	let value: CellValue

	if (type === 's') {
		const idx = typeof rawValue === 'number' ? rawValue : Number(rawValue)
		const entry = ctx.sharedStrings.get(idx)
		value = entry ?? (pool ? pool.internValue(stringValue('')) : stringValue(''))
	} else if (type === 'b') {
		const bv = rawValue === 1 || rawValue === true || rawValue === '1'
		value = pool ? pool.internValue(booleanValue(bv)) : booleanValue(bv)
	} else if (type === 'e') {
		value = pool
			? pool.internValue(errorValue(String(rawValue ?? '#VALUE!') as ExcelError))
			: errorValue(String(rawValue ?? '#VALUE!') as ExcelError)
	} else if (type === 'str') {
		value = pool
			? pool.internValue(stringValue(pool.internString(String(rawValue ?? ''))))
			: stringValue(String(rawValue ?? ''))
	} else if (type === 'inlineStr') {
		value = parseInlineString(c, pool)
	} else if (rawValue !== undefined && rawValue !== null && rawValue !== '') {
		const num = Number(rawValue)
		if (Number.isNaN(num)) {
			value = pool
				? pool.internValue(stringValue(pool.internString(String(rawValue))))
				: stringValue(String(rawValue))
		} else if (ctx.isDateFormat[styleIdx]) {
			value = { kind: 'date', serial: num }
		} else {
			value = pool ? pool.internValue(numberValue(num)) : numberValue(num)
		}
	} else if (formula) {
		value = EMPTY
	} else {
		return undefined
	}

	return {
		value,
		formula,
		styleId,
		...(formulaSpec.info ? { formulaInfo: formulaSpec.info } : {}),
	}
}

function parseFormulaText(
	formulaNode: unknown,
	row: number,
	col: number,
	sharedFormulaMasters: Map<string, { formula: string; row: number; col: number; ref?: string }>,
	pool?: ValueInternPool,
): { text: string | null; info?: Cell['formulaInfo'] } {
	if (formulaNode === undefined || formulaNode === null) return { text: null }
	if (
		typeof formulaNode === 'string' ||
		typeof formulaNode === 'number' ||
		typeof formulaNode === 'boolean'
	) {
		const text = String(formulaNode)
		return { text: pool ? pool.internString(text) : text }
	}
	if (typeof formulaNode === 'object') {
		const node = formulaNode as XmlNode
		const sharedIndex = attr(node, 'si')
		const formulaType = attr(node, 't')
		const text = node['#text']
		if (formulaType === 'shared' && sharedIndex) {
			const ref = attr(node, 'ref')
			if (text !== undefined && text !== null) {
				const formula = pool ? pool.internString(String(text)) : String(text)
				sharedFormulaMasters.set(sharedIndex, { formula, row, col, ...(ref ? { ref } : {}) })
				return {
					text: formula,
					info: { kind: 'shared', sharedIndex, isMaster: true, ...(ref ? { ref } : {}) },
				}
			}
			const master = sharedFormulaMasters.get(sharedIndex)
			if (!master) return { text: null }
			const translated = translateSharedFormula(master.formula, master.row, master.col, row, col)
			return {
				text: translated && pool ? pool.internString(translated) : translated,
				info: { kind: 'shared', sharedIndex, isMaster: false },
			}
		}
		if (formulaType === 'array') {
			const ref = attr(node, 'ref')
			if (text === undefined || text === null) {
				return {
					text: null,
					info: { kind: 'array', ...(ref ? { ref } : {}) },
				}
			}
			const formula = String(text)
			return {
				text: pool ? pool.internString(formula) : formula,
				info: { kind: 'array', ...(ref ? { ref } : {}) },
			}
		}
		if (text === undefined || text === null) return { text: null }
		const formula = String(text)
		return { text: pool ? pool.internString(formula) : formula }
	}
	return { text: null }
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

function parseInlineString(c: XmlNode, pool?: ValueInternPool): CellValue {
	const is = c.is as XmlNode | undefined
	if (!is || typeof is !== 'object') {
		return pool ? pool.internValue(stringValue('')) : stringValue('')
	}

	if (is.t !== undefined) {
		const text = String(is.t)
		return pool ? pool.internValue(stringValue(pool.internString(text))) : stringValue(text)
	}

	const runs = asArray<XmlNode>(is.r as XmlNode | XmlNode[])
	const text = runs.map((r) => (r.t !== undefined ? String(r.t) : '')).join('')
	return pool ? pool.internValue(stringValue(pool.internString(text))) : stringValue(text)
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

function parseDrawingRefs(ws: XmlNode, sheet: Sheet): void {
	const drawing = ws.drawing as XmlNode | undefined
	const legacyDrawing = ws.legacyDrawing as XmlNode | undefined
	sheet.drawingRefs = {
		hasDrawing: drawing !== undefined,
		hasLegacyDrawing: legacyDrawing !== undefined,
	}
}

const EXTLST_RE = /<extLst[\s>][\s\S]*?<\/extLst>/

function extractExtLst(xml: string, sheet: Sheet): void {
	const m = EXTLST_RE.exec(xml)
	if (m) sheet.preservedExtLst = m[0]
}

function parseSheetPr(ws: XmlNode, sheet: Sheet): void {
	const pr = ws.sheetPr as XmlNode | undefined
	if (!pr) return
	const tc = pr.tabColor as XmlNode | undefined
	if (tc) {
		const color: Record<string, string | number> = {}
		const rgb = attr(tc, 'rgb')
		if (rgb) color.rgb = rgb
		const theme = numAttr(tc, 'theme')
		if (theme !== undefined) color.theme = theme
		const tint = numAttr(tc, 'tint')
		if (tint !== undefined) color.tint = tint
		const indexed = numAttr(tc, 'indexed')
		if (indexed !== undefined) color.indexed = indexed
		sheet.tabColor = color as import('@ascend/core').SheetTabColor
	}
}

function parseSheetFormatPr(ws: XmlNode, sheet: Sheet): void {
	const fmt = ws.sheetFormatPr as XmlNode | undefined
	if (!fmt) return
	const fp: Record<string, number | boolean> = {}
	const drh = numAttr(fmt, 'defaultRowHeight')
	if (drh !== undefined) fp.defaultRowHeight = drh
	const dcw = numAttr(fmt, 'defaultColWidth')
	if (dcw !== undefined) fp.defaultColWidth = dcw
	const olr = numAttr(fmt, 'outlineLevelRow')
	if (olr !== undefined) fp.outlineLevelRow = olr
	const olc = numAttr(fmt, 'outlineLevelCol')
	if (olc !== undefined) fp.outlineLevelCol = olc
	const ch = boolAttr(fmt, 'customHeight')
	if (ch !== undefined) fp.customHeight = ch
	sheet.sheetFormatPr = fp as import('@ascend/core').SheetFormatPr
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
	const autoFilter = parseAutoFilterNode(ws.autoFilter as XmlNode | undefined)
	sheet.autoFilter = autoFilter as AutoFilter | null
}

function parseSheetProtection(ws: XmlNode, sheet: Sheet): void {
	const protection = ws.sheetProtection as XmlNode | undefined
	if (!protection) return
	const parsed: Record<string, string | number | boolean> = {}
	setIfDefined(parsed, 'sheet', readBoolAttribute(protection, 'sheet'))
	setIfDefined(parsed, 'objects', readBoolAttribute(protection, 'objects'))
	setIfDefined(parsed, 'scenarios', readBoolAttribute(protection, 'scenarios'))
	setIfDefined(parsed, 'formatCells', readBoolAttribute(protection, 'formatCells'))
	setIfDefined(parsed, 'formatColumns', readBoolAttribute(protection, 'formatColumns'))
	setIfDefined(parsed, 'formatRows', readBoolAttribute(protection, 'formatRows'))
	setIfDefined(parsed, 'insertColumns', readBoolAttribute(protection, 'insertColumns'))
	setIfDefined(parsed, 'insertRows', readBoolAttribute(protection, 'insertRows'))
	setIfDefined(parsed, 'insertHyperlinks', readBoolAttribute(protection, 'insertHyperlinks'))
	setIfDefined(parsed, 'deleteColumns', readBoolAttribute(protection, 'deleteColumns'))
	setIfDefined(parsed, 'deleteRows', readBoolAttribute(protection, 'deleteRows'))
	setIfDefined(parsed, 'selectLockedCells', readBoolAttribute(protection, 'selectLockedCells'))
	setIfDefined(parsed, 'sort', readBoolAttribute(protection, 'sort'))
	setIfDefined(parsed, 'autoFilter', readBoolAttribute(protection, 'autoFilter'))
	setIfDefined(parsed, 'pivotTables', readBoolAttribute(protection, 'pivotTables'))
	setIfDefined(parsed, 'selectUnlockedCells', readBoolAttribute(protection, 'selectUnlockedCells'))
	setIfDefined(parsed, 'password', attr(protection, 'password'))
	setIfDefined(parsed, 'algorithmName', attr(protection, 'algorithmName'))
	setIfDefined(parsed, 'hashValue', attr(protection, 'hashValue'))
	setIfDefined(parsed, 'saltValue', attr(protection, 'saltValue'))
	setIfDefined(parsed, 'spinCount', numAttr(protection, 'spinCount'))
	sheet.protection = parsed
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

	for (const ie of asArray<XmlNode>(ignoredErrors.ignoredError as XmlNode | XmlNode[])) {
		const sqref = attr(ie, 'sqref')
		if (!sqref) continue
		sheet.ignoredErrors.push({
			sqref,
			...(boolAttr(ie, 'numberStoredAsText') ? { numberStoredAsText: true } : {}),
			...(boolAttr(ie, 'formula') ? { formula: true } : {}),
			...(boolAttr(ie, 'formulaRange') ? { formulaRange: true } : {}),
			...(boolAttr(ie, 'evalError') ? { evalError: true } : {}),
			...(boolAttr(ie, 'twoDigitTextYear') ? { twoDigitTextYear: true } : {}),
			...(boolAttr(ie, 'unlockedFormula') ? { unlockedFormula: true } : {}),
			...(boolAttr(ie, 'emptyCellReference') ? { emptyCellReference: true } : {}),
			...(boolAttr(ie, 'listDataValidation') ? { listDataValidation: true } : {}),
			...(boolAttr(ie, 'calculatedColumn') ? { calculatedColumn: true } : {}),
		})
	}
}

function parseHyperlinks(
	ws: XmlNode,
	sheet: Sheet,
	relationships: readonly Relationship[],
	pool?: ValueInternPool,
): void {
	const hyperlinks = ws.hyperlinks as XmlNode | undefined
	if (!hyperlinks) return
	const relMap = new Map(relationships.map((rel) => [rel.id, rel]))

	for (const hyperlink of asArray<XmlNode>(hyperlinks.hyperlink as XmlNode | XmlNode[])) {
		const ref = attr(hyperlink, 'ref')
		if (!ref) continue
		const relId = attr(hyperlink, 'r:id') ?? attr(hyperlink, 'id')
		const rel = relId ? relMap.get(relId) : undefined
		const parsed: Record<string, string> = {}
		setIfDefined(
			parsed,
			'target',
			rel?.target ? (pool ? pool.internString(rel.target) : rel.target) : undefined,
		)
		setIfDefined(
			parsed,
			'location',
			attr(hyperlink, 'location')
				? pool
					? pool.internString(attr(hyperlink, 'location') as string)
					: (attr(hyperlink, 'location') as string)
				: undefined,
		)
		setIfDefined(
			parsed,
			'display',
			attr(hyperlink, 'display')
				? pool
					? pool.internString(attr(hyperlink, 'display') as string)
					: (attr(hyperlink, 'display') as string)
				: undefined,
		)
		setIfDefined(
			parsed,
			'tooltip',
			attr(hyperlink, 'tooltip')
				? pool
					? pool.internString(attr(hyperlink, 'tooltip') as string)
					: (attr(hyperlink, 'tooltip') as string)
				: undefined,
		)
		sheet.hyperlinks.set(ref, parsed)
	}
}

function parseConditionalFormatting(
	ws: XmlNode,
	sheet: Sheet,
	differentialStyles: readonly CellStyle[],
	pool?: ValueInternPool,
): void {
	for (const conditionalFormatting of asArray<XmlNode>(
		ws.conditionalFormatting as XmlNode | XmlNode[] | undefined,
	)) {
		const sqref = attr(conditionalFormatting, 'sqref')
		if (!sqref) continue
		const rules: SheetConditionalFormatRule[] = []
		for (const rule of asArray<XmlNode>(conditionalFormatting.cfRule as XmlNode | XmlNode[])) {
			const type = attr(rule, 'type')
			if (!type) continue
			const formulas = asArray<XmlNode | string | number | boolean>(
				rule.formula as XmlNode | XmlNode[] | string | string[] | undefined,
			).map((formula) => {
				const text = readNodeText(formula) ?? ''
				return pool ? pool.internString(text) : text
			})
			const dxfId = numAttr(rule, 'dxfId')
			const parsedRule: {
				type: string
				operator?: string
				dxfId?: number
				priority?: number
				stopIfTrue?: boolean
				formulas: readonly string[]
				style?: CellStyle
			} = {
				type,
				formulas,
			}
			const operator = attr(rule, 'operator')
			if (operator) parsedRule.operator = operator
			const priority = numAttr(rule, 'priority')
			if (priority !== undefined) parsedRule.priority = priority
			const stopIfTrue = readBoolAttribute(rule, 'stopIfTrue')
			if (stopIfTrue !== undefined) parsedRule.stopIfTrue = stopIfTrue
			if (dxfId !== undefined) {
				parsedRule.dxfId = dxfId
				const style = differentialStyles[dxfId]
				if (style) parsedRule.style = style
			}
			rules.push(parsedRule as SheetConditionalFormatRule)
		}
		if (rules.length > 0) {
			sheet.conditionalFormats.push({ sqref, rules } satisfies SheetConditionalFormat)
		}
	}
}

function parseDataValidations(ws: XmlNode, sheet: Sheet, pool?: ValueInternPool): void {
	const container = ws.dataValidations as XmlNode | undefined
	if (!container) return
	for (const validation of asArray<XmlNode>(container.dataValidation as XmlNode | XmlNode[])) {
		const sqref = attr(validation, 'sqref')
		if (!sqref) continue
		const parsed: {
			sqref: string
			type?: string
			operator?: string
			allowBlank?: boolean
			showInputMessage?: boolean
			showErrorMessage?: boolean
			showDropDown?: boolean
			promptTitle?: string
			prompt?: string
			errorTitle?: string
			error?: string
			errorStyle?: string
			formula1?: string
			formula2?: string
		} = { sqref }
		const type = attr(validation, 'type')
		if (type) parsed.type = type
		const operator = attr(validation, 'operator')
		if (operator) parsed.operator = operator
		const errorStyle = attr(validation, 'errorStyle')
		if (errorStyle) parsed.errorStyle = errorStyle
		const allowBlank = readBoolAttribute(validation, 'allowBlank')
		if (allowBlank !== undefined) parsed.allowBlank = allowBlank
		const showInputMessage = readBoolAttribute(validation, 'showInputMessage')
		if (showInputMessage !== undefined) parsed.showInputMessage = showInputMessage
		const showErrorMessage = readBoolAttribute(validation, 'showErrorMessage')
		if (showErrorMessage !== undefined) parsed.showErrorMessage = showErrorMessage
		const showDropDown = readBoolAttribute(validation, 'showDropDown')
		if (showDropDown !== undefined) parsed.showDropDown = showDropDown
		const promptTitle = attr(validation, 'promptTitle')
		if (promptTitle) parsed.promptTitle = promptTitle
		const prompt = attr(validation, 'prompt')
		if (prompt) parsed.prompt = prompt
		const errorTitle = attr(validation, 'errorTitle')
		if (errorTitle) parsed.errorTitle = errorTitle
		const error = attr(validation, 'error')
		if (error) parsed.error = error
		const formula1 = readNodeText(validation.formula1)
		if (formula1) parsed.formula1 = formula1
		const formula2 = readNodeText(validation.formula2)
		if (formula2) parsed.formula2 = formula2
		if (pool) {
			if (parsed.sqref) parsed.sqref = pool.internString(parsed.sqref)
			if (parsed.type) parsed.type = pool.internString(parsed.type)
			if (parsed.operator) parsed.operator = pool.internString(parsed.operator)
			if (parsed.promptTitle) parsed.promptTitle = pool.internString(parsed.promptTitle)
			if (parsed.prompt) parsed.prompt = pool.internString(parsed.prompt)
			if (parsed.errorTitle) parsed.errorTitle = pool.internString(parsed.errorTitle)
			if (parsed.error) parsed.error = pool.internString(parsed.error)
			if (parsed.errorStyle) parsed.errorStyle = pool.internString(parsed.errorStyle)
			if (parsed.formula1) parsed.formula1 = pool.internString(parsed.formula1)
			if (parsed.formula2) parsed.formula2 = pool.internString(parsed.formula2)
		}
		sheet.dataValidations.push(parsed as SheetDataValidation)
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
