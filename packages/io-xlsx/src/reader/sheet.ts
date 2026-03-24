import type {
	AutoFilter,
	Cell,
	CellStyle,
	DynamicArrayFormulaInfo,
	RangeRef,
	SheetBreak,
	SheetColDef,
	SheetConditionalFormat,
	SheetConditionalFormatRule,
	SheetDataValidation,
	StyleId,
} from '@ascend/core'
import {
	DEFAULT_STYLE_ID,
	indexToColumn,
	parseRange,
	Sheet,
	SPARSE_TO_DENSE_THRESHOLD,
} from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { parseFormula } from '@ascend/formulas'
import type { CellValue, ExcelError } from '@ascend/schema'
import {
	booleanValue,
	dateValue,
	EMPTY,
	errorValue,
	numberValue,
	richTextValue,
	stringValue,
} from '@ascend/schema'
import { normalizeStoredFormulaText } from '../formula-storage.ts'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { parseAutoFilterNode } from './filtering.ts'
import type { ParsedMetadataPart } from './metadata.ts'
import type { Relationship } from './relationships.ts'
import type { SharedStringResolver } from './shared-strings.ts'
import { decodeXmlText, findTagEnd, isSelfClosingTag } from './xml-utils.ts'

const SMALL_NUMBER_RANGE_START = -128
const SMALL_NUMBER_RANGE_END = 512
const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g
const TEXT_NODE_RE =
	/<([A-Za-z_][\w:.-]*)\b([^>]*)>([\s\S]*?)<\/\1>|<([A-Za-z_][\w:.-]*)\b([^>]*)\/>/g

const NO_FORMULA: { text: null; info: undefined } = { text: null, info: undefined }
const NULL_FORMULA_TEXT: { text: null } = { text: null }

interface CellPosition {
	readonly row: number
	readonly col: number
}

interface SharedFormulaMaster {
	readonly formula: string
	readonly row: number
	readonly col: number
	readonly ref?: string
	readonly masterRef?: string
	readonly parsed?: FormulaNode
}

type SharedFormulaMasterMap = Map<string, SharedFormulaMaster>

export interface SheetParseContext {
	readonly sharedStrings: SharedStringResolver
	readonly styleIds: StyleId[]
	readonly isDateFormat: boolean[]
	readonly differentialStyles?: readonly CellStyle[]
	readonly relationships?: readonly Relationship[]
	readonly valuePool?: ValueInternPool
	readonly valuesOnly?: boolean
	readonly formulaOnly?: boolean
	readonly formulaFeatures?: SheetFormulaFeatures
	readonly metadata?: ParsedMetadataPart
	readonly maxRows?: number
}

export interface SheetFormulaFeatures {
	hasSharedFormula: boolean
	hasArrayFormula: boolean
	hasDynamicArrayFormula: boolean
}

export interface StreamedSheetRow {
	readonly row: number
	readonly cells: readonly (readonly [number, Cell])[]
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
				const interned = stringValue(text)
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
						...(run.color
							? { color: typeof run.color === 'string' ? this.internString(run.color) : run.color }
							: {}),
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

const CHUNK_SIZE = 64

function parseDimensionRef(xml: string): string | null {
	const m = /<dimension\s+ref="([^"]+)"/.exec(xml)
	return m?.[1] ?? null
}

function applyDensityHintFromDimension(sheet: Sheet, xml: string): void {
	const ref = parseDimensionRef(xml)
	if (!ref) return
	try {
		const range = parseRange(ref)
		const rows = range.end.row - range.start.row + 1
		const cols = range.end.col - range.start.col + 1
		if (rows <= 0 || cols <= 0) return
		const numChunks = Math.ceil(rows / CHUNK_SIZE) * Math.ceil(cols / CHUNK_SIZE)
		const cellsPerChunk = (rows * cols) / numChunks
		if (cellsPerChunk >= SPARSE_TO_DENSE_THRESHOLD) {
			sheet.cells.setExpectedDensity('dense')
		}
	} catch {
		// ignore invalid dimension ref
	}
}

export function parseSheet(
	name: string,
	xml: string,
	ctx: SheetParseContext,
	sheetId?: Sheet['id'],
): Sheet {
	const sheet = new Sheet(name, sheetId)
	applyDensityHintFromDimension(sheet, xml)
	const sheetDataLoc = locateSheetData(xml)
	if (sheetDataLoc) parseSheetDataFromLoc(xml, sheetDataLoc, sheet, ctx)
	const strippedXml = sheetDataLoc
		? `${xml.slice(0, sheetDataLoc.tagStart)}<sheetData/>${xml.slice(sheetDataLoc.closeEnd)}`
		: xml
	const doc = parseXml(strippedXml)
	const ws = doc.worksheet as XmlNode | undefined
	if (!ws) return sheet

	parseSheetPr(ws, sheet)
	parseSheetFormatPr(ws, sheet)
	parseSheetViews(ws, sheet)
	parseCols(ws, sheet)
	parseMergeCells(ws, sheet)
	parseDrawingRefs(ws, sheet)
	parseAutoFilter(ws, sheet)
	parseSheetProtection(ws, sheet)
	parsePageMargins(ws, sheet)
	parsePageSetup(ws, sheet)
	parsePrintOptions(ws, sheet)
	parseHeaderFooter(ws, sheet)
	parsePageBreaks(ws, sheet)
	parseIgnoredErrors(ws, sheet)
	if (!ctx.valuesOnly && !ctx.formulaOnly) {
		parseHyperlinks(ws, sheet, ctx.relationships ?? [], ctx.valuePool)
		parseConditionalFormatting(ws, sheet, ctx.differentialStyles ?? [], ctx.valuePool)
		parseDataValidations(ws, sheet, ctx.valuePool)
		extractExtLst(xml, sheet)
	}
	return sheet
}

function parseSheetDataFromLoc(
	xml: string,
	sheetData: SheetDataLocation,
	sheet: Sheet,
	ctx: SheetParseContext,
): void {
	const sharedFormulaMasters: SharedFormulaMasterMap = new Map()
	let rowCursor = sheetData.contentStart
	let currentRow = -1
	const fallbackPos = { row: 0, col: 0 }
	const cellOut = { row: 0, col: 0 }

	while (true) {
		const rowOpen = xml.indexOf('<row', rowCursor)
		if (rowOpen === -1 || rowOpen >= sheetData.contentEnd) return
		const rowTagEnd = findTagEnd(xml, rowOpen)
		if (rowTagEnd === -1 || rowTagEnd >= sheetData.contentEnd) return
		const rowAttrsRaw = xml.slice(rowOpen + 4, rowTagEnd)
		const explicitRowIndex = rawNumAttr(rowAttrsRaw, 'r')
		const row = explicitRowIndex !== undefined ? explicitRowIndex - 1 : currentRow + 1
		currentRow = row
		if (ctx.maxRows !== undefined && row >= ctx.maxRows) return
		const rowHeight = rawNumAttr(rowAttrsRaw, 'ht')
		if (rowHeight !== undefined && rawAttr(rowAttrsRaw, 'customHeight') === '1') {
			sheet.rowHeights.set(row, rowHeight)
		}
		const hidden = rawBoolAttr(rowAttrsRaw, 'hidden')
		const collapsed = rawBoolAttr(rowAttrsRaw, 'collapsed')
		const outlineLevel = rawNumAttr(rowAttrsRaw, 'outlineLevel')
		if (hidden !== undefined || collapsed !== undefined || outlineLevel !== undefined) {
			const rowDef: Record<string, boolean | number> = {}
			if (hidden !== undefined) rowDef.hidden = hidden
			if (collapsed !== undefined) rowDef.collapsed = collapsed
			if (outlineLevel !== undefined) rowDef.outlineLevel = outlineLevel
			sheet.rowDefs.set(row, rowDef as import('@ascend/core').SheetRowDef)
		}
		if (isSelfClosingTag(xml, rowOpen, rowTagEnd)) {
			rowCursor = rowTagEnd + 1
			continue
		}

		const rowClose = xml.indexOf('</row>', rowTagEnd + 1)
		if (rowClose === -1 || rowClose > sheetData.contentEnd) return
		let cellCursor = rowTagEnd + 1
		let nextCol = 0
		while (true) {
			const cellOpen = xml.indexOf('<c', cellCursor)
			if (cellOpen === -1 || cellOpen >= rowClose) break
			const cellTagEnd = findTagEnd(xml, cellOpen)
			if (cellTagEnd === -1 || cellTagEnd > rowClose) break
			const rawAttrs = xml.slice(cellOpen + 2, cellTagEnd)
			const selfClosing = isSelfClosingTag(xml, cellOpen, cellTagEnd)
			const cellClose = selfClosing ? -1 : xml.indexOf('</c>', cellTagEnd + 1)
			const innerXml =
				!selfClosing && cellClose !== -1 && cellClose <= rowClose
					? xml.slice(cellTagEnd + 1, cellClose)
					: ''
			fallbackPos.row = row
			fallbackPos.col = nextCol
			const ok = parseFastCell(
				rawAttrs,
				innerXml,
				ctx,
				sharedFormulaMasters,
				fallbackPos,
				sheet,
				cellOut,
			)
			cellCursor =
				selfClosing || cellClose === -1 || cellClose > rowClose ? cellTagEnd + 1 : cellClose + 4
			if (!ok) continue
			nextCol = cellOut.col + 1
		}
		rowCursor = rowClose + 6
	}
}

export function* streamSheetRowsXml(
	name: string,
	xml: string,
	ctx: SheetParseContext,
): Generator<StreamedSheetRow> {
	const sheetData = locateSheetData(xml)
	if (!sheetData) return
	const sharedFormulaMasters: SharedFormulaMasterMap = new Map()
	let rowCursor = sheetData.contentStart
	let currentRow = -1
	const fallbackPos = { row: 0, col: 0 }
	const cellOut = { row: 0, col: 0 }
	const rowSheet = new Sheet(name)

	while (true) {
		const rowOpen = xml.indexOf('<row', rowCursor)
		if (rowOpen === -1 || rowOpen >= sheetData.contentEnd) return
		const rowTagEnd = findTagEnd(xml, rowOpen)
		if (rowTagEnd === -1 || rowTagEnd >= sheetData.contentEnd) return
		const rowAttrsRaw = xml.slice(rowOpen + 4, rowTagEnd)
		const explicitRowIndex = rawNumAttr(rowAttrsRaw, 'r')
		const row = explicitRowIndex !== undefined ? explicitRowIndex - 1 : currentRow + 1
		currentRow = row
		rowSheet.cells.clear()
		rowSheet.rowDefs.clear()
		const hidden = rawBoolAttr(rowAttrsRaw, 'hidden')
		const collapsed = rawBoolAttr(rowAttrsRaw, 'collapsed')
		const outlineLevel = rawNumAttr(rowAttrsRaw, 'outlineLevel')
		if (hidden !== undefined || collapsed !== undefined || outlineLevel !== undefined) {
			const rowDef: Record<string, boolean | number> = {}
			if (hidden !== undefined) rowDef.hidden = hidden
			if (collapsed !== undefined) rowDef.collapsed = collapsed
			if (outlineLevel !== undefined) rowDef.outlineLevel = outlineLevel
			rowSheet.rowDefs.set(row, rowDef as import('@ascend/core').SheetRowDef)
		}
		if (isSelfClosingTag(xml, rowOpen, rowTagEnd)) {
			yield { row, cells: [] }
			rowCursor = rowTagEnd + 1
			continue
		}

		const rowClose = xml.indexOf('</row>', rowTagEnd + 1)
		if (rowClose === -1 || rowClose > sheetData.contentEnd) return
		let cellCursor = rowTagEnd + 1
		let nextCol = 0
		while (true) {
			const cellOpen = xml.indexOf('<c', cellCursor)
			if (cellOpen === -1 || cellOpen >= rowClose) break
			const cellTagEnd = findTagEnd(xml, cellOpen)
			if (cellTagEnd === -1 || cellTagEnd > rowClose) break
			const rawAttrs = xml.slice(cellOpen + 2, cellTagEnd)
			const selfClosing = isSelfClosingTag(xml, cellOpen, cellTagEnd)
			const cellClose = selfClosing ? -1 : xml.indexOf('</c>', cellTagEnd + 1)
			const innerXml =
				!selfClosing && cellClose !== -1 && cellClose <= rowClose
					? xml.slice(cellTagEnd + 1, cellClose)
					: ''
			fallbackPos.row = row
			fallbackPos.col = nextCol
			const ok = parseFastCell(
				rawAttrs,
				innerXml,
				ctx,
				sharedFormulaMasters,
				fallbackPos,
				rowSheet,
				cellOut,
			)
			cellCursor =
				selfClosing || cellClose === -1 || cellClose > rowClose ? cellTagEnd + 1 : cellClose + 4
			if (!ok) continue
			nextCol = cellOut.col + 1
		}
		const first = rowSheet.cells.iterateRows().next()
		yield { row, cells: first.done ? [] : first.value[1] }
		rowCursor = rowClose + 6
	}
}

function parseSlowCell(
	rawAttrs: string,
	innerXml: string,
	ctx: SheetParseContext,
	sharedFormulaMasters: SharedFormulaMasterMap,
	fallbackPosition: CellPosition | undefined,
	sheet: Sheet,
	out: { row: number; col: number },
): boolean {
	const cellNode = buildCellNode(rawAttrs, innerXml)
	const ref = attr(cellNode, 'r')
	const pos = ref ? parseCellRef(ref) : fallbackPosition
	if (!pos) return false
	out.row = pos.row
	out.col = pos.col
	return resolveCellToSheet(cellNode, ctx, pos.row, pos.col, sharedFormulaMasters, sheet)
}

function parseFastCell(
	rawAttrs: string,
	innerXml: string,
	ctx: SheetParseContext,
	sharedFormulaMasters: SharedFormulaMasterMap,
	fallbackPosition: CellPosition | undefined,
	sheet: Sheet,
	out: { row: number; col: number },
): boolean {
	const pos = resolveCellPosition(rawAttrs, fallbackPosition)
	if (!pos) return false
	const type = rawAttr(rawAttrs, 't')
	if (type === 'inlineStr' || innerXml.includes('<is')) {
		return parseSlowCell(rawAttrs, innerXml, ctx, sharedFormulaMasters, pos, sheet, out)
	}

	const pool = ctx.valuePool
	const styleIdx = rawNumAttr(rawAttrs, 's') ?? 0
	const rawValue = extractVTagText(innerXml)
	const styleId = ctx.valuesOnly ? DEFAULT_STYLE_ID : (ctx.styleIds[styleIdx] ?? DEFAULT_STYLE_ID)
	const metadataIndex = rawNumAttr(rawAttrs, 'cm')
	const formulaSpec = ctx.valuesOnly
		? NO_FORMULA
		: parseFormulaText(
				extractRawFormulaNode(innerXml),
				pos.row,
				pos.col,
				sharedFormulaMasters,
				pool,
				ctx.formulaFeatures,
			)
	const binding = attachDynamicArrayBinding(
		formulaSpec.info,
		formulaSpec.text,
		metadataIndex,
		ctx.metadata,
		ctx.formulaFeatures,
	)

	let value: CellValue
	if (type === 's') {
		const idx = rawValue !== undefined ? fastParseNonNegInt(rawValue) : -1
		if (idx < 0) {
			value = pool ? pool.internValue(stringValue('')) : stringValue('')
		} else {
			const entry = ctx.sharedStrings.get(idx)
			value = entry ?? (pool ? pool.internValue(stringValue('')) : stringValue(''))
		}
	} else if (type === 'b') {
		value = booleanValue(rawValue === '1')
	} else if (type === 'e') {
		value = errorValue((rawValue ?? '#VALUE!') as ExcelError)
	} else if (type === 'str') {
		const text = rawValue ?? ''
		value = pool ? pool.internValue(stringValue(pool.internString(text))) : stringValue(text)
	} else if (rawValue !== undefined && rawValue !== '') {
		const num = Number(rawValue)
		if (Number.isNaN(num)) {
			value = pool
				? pool.internValue(stringValue(pool.internString(rawValue)))
				: stringValue(rawValue)
		} else if (ctx.isDateFormat[styleIdx]) {
			value = dateValue(num)
		} else {
			value = pool ? pool.internValue(numberValue(num)) : numberValue(num)
		}
	} else if (formulaSpec.text) {
		value = EMPTY
	} else if (ctx.valuesOnly && innerXml.includes('<f')) {
		value = EMPTY
	} else if (type) {
		value = type === 'n' ? (pool ? pool.internValue(numberValue(0)) : numberValue(0)) : EMPTY
	} else {
		return false
	}

	out.row = pos.row
	out.col = pos.col
	sheet.cells.setResolved(pos.row, pos.col, value, formulaSpec.text, styleId, binding)
	return true
}

interface SheetDataLocation {
	readonly contentStart: number
	readonly contentEnd: number
	readonly tagStart: number
	readonly closeEnd: number
}

function locateSheetData(xml: string): SheetDataLocation | null {
	const open = xml.indexOf('<sheetData')
	if (open === -1) return null
	const tagEnd = findTagEnd(xml, open)
	if (tagEnd === -1) return null
	if (isSelfClosingTag(xml, open, tagEnd)) return null
	const close = xml.indexOf('</sheetData>', tagEnd + 1)
	if (close === -1) return null
	return { contentStart: tagEnd + 1, contentEnd: close, tagStart: open, closeEnd: close + 12 }
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
		runs.push(buildInlineRunNode(run.text ?? ''))
	}
	if (runs.length > 0) return { r: runs }
	const directText = extractTextNode(innerXml, 't')
	return directText ? { t: directText.text ?? '' } : {}
}

function buildInlineRunNode(runInnerXml: string): XmlNode {
	const textNode = extractTextNode(runInnerXml, 't')
	const result: XmlNode = { t: textNode?.text ?? '' }
	const rPrNode = extractTextNode(runInnerXml, 'rPr')
	if (!rPrNode?.text) return result
	const rPr: XmlNode = {}
	if (runInnerXml.includes('<b')) rPr.b = {}
	if (runInnerXml.includes('<i')) rPr.i = {}
	if (runInnerXml.includes('<u')) rPr.u = {}
	if (runInnerXml.includes('<strike')) rPr.strike = {}
	const rFont = extractTextNode(runInnerXml, 'rFont') ?? extractTextNode(runInnerXml, 'font')
	if (rFont?.attrs) {
		const valMatch = /val="([^"]*)"/.exec(rFont.attrs)
		if (valMatch) rPr.rFont = { '@_val': valMatch[1] }
	}
	const sz = extractTextNode(runInnerXml, 'sz')
	if (sz?.attrs) {
		const valMatch = /val="([^"]*)"/.exec(sz.attrs)
		if (valMatch) rPr.sz = { '@_val': Number(valMatch[1]) }
	}
	const color = extractTextNode(runInnerXml, 'color')
	if (color?.attrs) {
		const rgbMatch = /rgb="([^"]*)"/.exec(color.attrs)
		if (rgbMatch) rPr.color = { '@_rgb': rgbMatch[1] }
	}
	if (Object.keys(rPr).length > 0) result.rPr = rPr
	return result
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

const ATTR_NEEDLES: Record<string, string> = {
	r: 'r="',
	t: 't="',
	s: 's="',
	cm: 'cm="',
	ref: 'ref="',
	si: 'si="',
}

function rawAttr(rawAttrs: string, name: string): string | undefined {
	const needle = ATTR_NEEDLES[name] ?? `${name}="`
	const start = rawAttrs.indexOf(needle)
	if (start === -1) return undefined
	const valueStart = start + needle.length
	const valueEnd = rawAttrs.indexOf('"', valueStart)
	if (valueEnd === -1) return undefined
	return decodeXmlText(rawAttrs.slice(valueStart, valueEnd))
}

function rawNumAttr(rawAttrs: string, name: string): number | undefined {
	const needle = ATTR_NEEDLES[name] ?? `${name}="`
	const start = rawAttrs.indexOf(needle)
	if (start === -1) return undefined
	const valueStart = start + needle.length
	const valueEnd = rawAttrs.indexOf('"', valueStart)
	if (valueEnd === -1) return undefined
	const parsed = Number(rawAttrs.slice(valueStart, valueEnd))
	return Number.isNaN(parsed) ? undefined : parsed
}

function rawBoolAttr(rawAttrs: string, name: string): boolean | undefined {
	const value = rawAttr(rawAttrs, name)
	if (value === undefined) return undefined
	if (value === '1' || value.toLowerCase() === 'true') return true
	if (value === '0' || value.toLowerCase() === 'false') return false
	return undefined
}

function fastParseNonNegInt(s: string): number {
	let result = 0
	for (let i = 0; i < s.length; i++) {
		const d = s.charCodeAt(i) - 48
		if (d < 0 || d > 9) return -1
		result = result * 10 + d
	}
	return s.length > 0 ? result : -1
}

function resolveCellPosition(
	rawAttrs: string,
	fallbackPosition?: CellPosition,
): CellPosition | undefined {
	const ref = rawAttr(rawAttrs, 'r')
	if (ref) return parseCellRef(ref)
	return fallbackPosition
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

function extractVTagText(xml: string): string | undefined {
	const open = xml.indexOf('<v')
	if (open === -1) return undefined
	const contentStart = xml.indexOf('>', open)
	if (contentStart === -1) return undefined
	const close = xml.indexOf('</v>', contentStart + 1)
	if (close === -1) return undefined
	const slice = xml.slice(contentStart + 1, close)
	return slice.includes('&') ? decodeXmlText(slice) : slice
}

interface RawFormulaNode {
	readonly rawAttrs: string
	readonly text?: string
}

function extractRawFormulaNode(innerXml: string): RawFormulaNode | undefined {
	const open = innerXml.indexOf('<f')
	if (open === -1) return undefined
	const tagEnd = findTagEnd(innerXml, open)
	if (tagEnd === -1) return undefined
	const rawAttrs = innerXml.slice(open + 2, tagEnd)
	if (isSelfClosingTag(innerXml, open, tagEnd)) return { rawAttrs }
	const close = innerXml.indexOf('</f>', tagEnd + 1)
	if (close === -1) return { rawAttrs }
	return { rawAttrs, text: decodeXmlText(innerXml.slice(tagEnd + 1, close)) }
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

function parseCellRef(ref: string): { row: number; col: number } | undefined {
	let index = 0
	let col = 0
	while (index < ref.length) {
		const code = ref.charCodeAt(index)
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) {
			col = col * 26 + (code - 64)
		} else if (code >= 97 && code <= 122) {
			col = col * 26 + (code - 96)
		} else {
			return undefined
		}
		index += 1
	}
	if (index === 0 || index >= ref.length) return undefined
	let row = 0
	while (index < ref.length) {
		const code = ref.charCodeAt(index)
		if (code < 48 || code > 57) return undefined
		row = row * 10 + (code - 48)
		index += 1
	}
	if (row <= 0 || col <= 0) return undefined
	return { row: row - 1, col: col - 1 }
}

function resolveCellToSheet(
	c: XmlNode,
	ctx: SheetParseContext,
	row: number,
	col: number,
	sharedFormulaMasters: SharedFormulaMasterMap,
	sheet: Sheet,
): boolean {
	const pool = ctx.valuePool
	const type = attr(c, 't')
	const styleIdx = numAttr(c, 's') ?? 0
	const rawValue = c.v
	const styleId = ctx.valuesOnly ? DEFAULT_STYLE_ID : (ctx.styleIds[styleIdx] ?? DEFAULT_STYLE_ID)
	const metadataIndex = numAttr(c, 'cm')
	const formulaSpec = ctx.valuesOnly
		? NO_FORMULA
		: parseFormulaText(c.f, row, col, sharedFormulaMasters, pool, ctx.formulaFeatures)
	const formula = formulaSpec.text
	const binding = attachDynamicArrayBinding(
		formulaSpec.info,
		formula,
		metadataIndex,
		ctx.metadata,
		ctx.formulaFeatures,
	)

	let value: CellValue

	if (type === 's') {
		const idx = typeof rawValue === 'number' ? rawValue : Number(rawValue)
		const entry = ctx.sharedStrings.get(idx)
		value = entry ?? stringValue('')
	} else if (type === 'b') {
		value = booleanValue(rawValue === 1 || rawValue === true || rawValue === '1')
	} else if (type === 'e') {
		value = errorValue((rawValue != null ? String(rawValue) : '#VALUE!') as ExcelError)
	} else if (type === 'str') {
		const text = rawValue != null ? String(rawValue) : ''
		value = pool ? pool.internValue(stringValue(pool.internString(text))) : stringValue(text)
	} else if (type === 'inlineStr') {
		value = parseInlineString(c, pool)
	} else if (rawValue !== undefined && rawValue !== null && rawValue !== '') {
		const num = Number(rawValue)
		if (Number.isNaN(num)) {
			value = pool
				? pool.internValue(stringValue(pool.internString(String(rawValue))))
				: stringValue(String(rawValue))
		} else if (ctx.isDateFormat[styleIdx]) {
			value = dateValue(num)
		} else {
			value = pool ? pool.internValue(numberValue(num)) : numberValue(num)
		}
	} else if (formula) {
		value = EMPTY
	} else if (ctx.valuesOnly && c.f !== undefined && c.f !== null) {
		value = EMPTY
	} else if (type) {
		value = type === 'n' ? (pool ? pool.internValue(numberValue(0)) : numberValue(0)) : EMPTY
	} else {
		return false
	}

	sheet.cells.setResolved(row, col, value, ctx.valuesOnly ? null : formula, styleId, binding)
	return true
}

function parseFormulaText(
	formulaNode: unknown,
	row: number,
	col: number,
	sharedFormulaMasters: SharedFormulaMasterMap,
	pool?: ValueInternPool,
	formulaFeatures?: SheetFormulaFeatures,
): { text: string | null; info?: Cell['formulaInfo'] } {
	if (formulaNode === undefined || formulaNode === null) return NULL_FORMULA_TEXT
	if (
		typeof formulaNode === 'string' ||
		typeof formulaNode === 'number' ||
		typeof formulaNode === 'boolean'
	) {
		const text = String(formulaNode)
		return { text: pool ? pool.internString(text) : text }
	}
	if (isRawFormulaNode(formulaNode)) {
		const sharedIndex = rawAttr(formulaNode.rawAttrs, 'si')
		const formulaType = rawAttr(formulaNode.rawAttrs, 't')
		const text = formulaNode.text
		return parseResolvedFormulaText(
			formulaType,
			sharedIndex,
			rawAttr(formulaNode.rawAttrs, 'ref'),
			text,
			row,
			col,
			sharedFormulaMasters,
			pool,
			formulaFeatures,
		)
	}
	if (typeof formulaNode === 'object') {
		const node = formulaNode as XmlNode
		const sharedIndex = attr(node, 'si')
		const formulaType = attr(node, 't')
		const text = node['#text']
		return parseResolvedFormulaText(
			formulaType,
			sharedIndex,
			attr(node, 'ref'),
			text,
			row,
			col,
			sharedFormulaMasters,
			pool,
			formulaFeatures,
		)
	}
	return NULL_FORMULA_TEXT
}

function parseResolvedFormulaText(
	formulaType: string | undefined,
	sharedIndex: string | undefined,
	ref: string | undefined,
	text: unknown,
	row: number,
	col: number,
	sharedFormulaMasters: SharedFormulaMasterMap,
	pool?: ValueInternPool,
	formulaFeatures?: SheetFormulaFeatures,
): { text: string | null; info?: Cell['formulaInfo'] } {
	if (formulaType === 'shared' && sharedIndex) {
		if (formulaFeatures) formulaFeatures.hasSharedFormula = true
		if (text !== undefined && text !== null) {
			const normalized = normalizeStoredFormulaText(String(text))
			const formula = pool ? pool.internString(normalized) : normalized
			const parsed = parseFormula(formula)
			sharedFormulaMasters.set(sharedIndex, {
				formula,
				row,
				col,
				masterRef: toCellRef(row, col),
				...(ref ? { ref } : {}),
				...(parsed.ok ? { parsed: parsed.value } : {}),
			})
			return {
				text: formula,
				info: {
					kind: 'shared',
					sharedIndex,
					isMaster: true,
					masterRef: toCellRef(row, col),
					...(ref ? { ref } : {}),
				},
			}
		}
		const master = sharedFormulaMasters.get(sharedIndex)
		if (!master) return NULL_FORMULA_TEXT
		return {
			text: null,
			info: {
				kind: 'shared',
				sharedIndex,
				isMaster: false,
				...(master.masterRef ? { masterRef: master.masterRef } : {}),
			},
		}
	}
	if (formulaType === 'array') {
		if (formulaFeatures) formulaFeatures.hasArrayFormula = true
		if (text === undefined || text === null) {
			return {
				text: null,
				info: { kind: 'array', ...(ref ? { ref } : {}) },
			}
		}
		const formula = normalizeStoredFormulaText(String(text))
		return {
			text: pool ? pool.internString(formula) : formula,
			info: { kind: 'array', ...(ref ? { ref } : {}) },
		}
	}
	if (text === undefined || text === null) return NULL_FORMULA_TEXT
	const formula = normalizeStoredFormulaText(String(text))
	return { text: pool ? pool.internString(formula) : formula }
}

function attachDynamicArrayBinding(
	existing: Cell['formulaInfo'] | undefined,
	formulaText: string | null,
	metadataIndex: number | undefined,
	metadata: ParsedMetadataPart | undefined,
	formulaFeatures: SheetFormulaFeatures | undefined,
): Cell['formulaInfo'] | undefined {
	if (existing || !formulaText || metadataIndex === undefined || metadataIndex <= 0) return existing
	const record = metadata?.dynamicArrayByCellMetadataIndex.get(metadataIndex)
	if (!record) return existing
	if (formulaFeatures) formulaFeatures.hasDynamicArrayFormula = true
	const binding: DynamicArrayFormulaInfo = {
		kind: 'dynamicArray',
		metadataIndex,
		...(record.collapsed !== undefined ? { collapsed: record.collapsed } : {}),
	}
	return binding
}

function isRawFormulaNode(value: unknown): value is RawFormulaNode {
	return typeof value === 'object' && value !== null && 'rawAttrs' in value
}

function toCellRef(row: number, col: number): string {
	return `${indexToColumn(col)}${row + 1}`
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

	const rawRuns = asArray<XmlNode>(is.r as XmlNode | XmlNode[])
	if (rawRuns.length === 0) return stringValue('')

	const runs: import('@ascend/schema').RichTextRun[] = []
	for (const r of rawRuns) {
		const text = r.t !== undefined ? String(r.t) : ''
		const rPr = r.rPr as XmlNode | undefined
		if (rPr && typeof rPr === 'object') {
			const run: import('@ascend/schema').RichTextRun = {
				text: pool ? pool.internString(text) : text,
				...(rPr.b !== undefined ? { bold: true } : {}),
				...(rPr.i !== undefined ? { italic: true } : {}),
				...(rPr.u !== undefined ? { underline: true } : {}),
				...(rPr.strike !== undefined ? { strikethrough: true } : {}),
				...parseInlineRunFontProps(rPr),
			}
			runs.push(run)
		} else {
			runs.push({ text: pool ? pool.internString(text) : text })
		}
	}

	const first = runs[0]
	if (
		runs.length === 1 &&
		first &&
		!first.bold &&
		!first.italic &&
		!first.underline &&
		!first.strikethrough &&
		!first.fontName &&
		!first.fontSize &&
		!first.color
	) {
		return pool ? pool.internValue(stringValue(first.text)) : stringValue(first.text)
	}

	return richTextValue(runs)
}

function parseInlineRunFontProps(
	rPr: XmlNode,
): Pick<import('@ascend/schema').RichTextRun, 'fontName' | 'fontSize' | 'color'> {
	const result: Record<string, unknown> = {}
	const rFont = rPr.rFont
	if (typeof rFont === 'object' && rFont !== null) {
		const name = attr(rFont as XmlNode, 'val')
		if (name) result.fontName = name
	}
	const sz = rPr.sz
	if (typeof sz === 'object' && sz !== null) {
		const size = numAttr(sz as XmlNode, 'val')
		if (size !== undefined) result.fontSize = size
	}
	const color = rPr.color
	if (typeof color === 'object' && color !== null) {
		const colorNode = color as XmlNode
		const rgb = attr(colorNode, 'rgb')
		if (rgb) result.color = rgb
	}
	return result as Pick<import('@ascend/schema').RichTextRun, 'fontName' | 'fontSize' | 'color'>
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
	const outlinePr = pr.outlinePr as XmlNode | undefined
	if (outlinePr) {
		const parsed: Record<string, boolean> = {}
		const summaryBelow = readBoolAttribute(outlinePr, 'summaryBelow')
		if (summaryBelow !== undefined) parsed.summaryBelow = summaryBelow
		const summaryRight = readBoolAttribute(outlinePr, 'summaryRight')
		if (summaryRight !== undefined) parsed.summaryRight = summaryRight
		const applyStyles = readBoolAttribute(outlinePr, 'applyStyles')
		if (applyStyles !== undefined) parsed.applyStyles = applyStyles
		const showOutlineSymbols = readBoolAttribute(outlinePr, 'showOutlineSymbols')
		if (showOutlineSymbols !== undefined) parsed.showOutlineSymbols = showOutlineSymbols
		sheet.outlinePr = parsed as import('@ascend/core').SheetOutlinePr
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
	if (pane) {
		const ySplit = numAttr(pane, 'ySplit')
		if (ySplit !== undefined) sheet.frozenRows = Math.trunc(ySplit)
		const xSplit = numAttr(pane, 'xSplit')
		if (xSplit !== undefined) sheet.frozenCols = Math.trunc(xSplit)
	}

	const viewAttrs: Record<string, number | boolean | string> = {}
	const zoomScale = numAttr(firstView, 'zoomScale')
	if (zoomScale !== undefined) viewAttrs.zoomScale = zoomScale
	const showGridLines = readBoolAttribute(firstView, 'showGridLines')
	if (showGridLines !== undefined) viewAttrs.showGridLines = showGridLines
	const showFormulas = readBoolAttribute(firstView, 'showFormulas')
	if (showFormulas !== undefined) viewAttrs.showFormulas = showFormulas
	const rightToLeft = readBoolAttribute(firstView, 'rightToLeft')
	if (rightToLeft !== undefined) viewAttrs.rightToLeft = rightToLeft
	const tabSelected = readBoolAttribute(firstView, 'tabSelected')
	if (tabSelected !== undefined) viewAttrs.tabSelected = tabSelected
	const viewVal = attr(firstView, 'view')
	if (viewVal === 'normal' || viewVal === 'pageBreakPreview' || viewVal === 'pageLayout') {
		viewAttrs.view = viewVal
	}
	if (Object.keys(viewAttrs).length > 0) {
		sheet.sheetView = viewAttrs as import('@ascend/core').SheetView
	}
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

function parsePageBreaks(ws: XmlNode, sheet: Sheet): void {
	parseBreakCollection(ws.rowBreaks as XmlNode | undefined, sheet.rowBreaks)
	parseBreakCollection(ws.colBreaks as XmlNode | undefined, sheet.colBreaks)
}

function parseBreakCollection(node: XmlNode | undefined, target: SheetBreak[]): void {
	if (!node) return
	for (const brk of asArray<XmlNode>(node.brk as XmlNode | XmlNode[] | undefined)) {
		const id = numAttr(brk, 'id')
		if (id === undefined) continue
		const min = numAttr(brk, 'min')
		const max = numAttr(brk, 'max')
		const man = boolAttr(brk, 'man')
		const pt = boolAttr(brk, 'pt')
		const parsed: SheetBreak = {
			id,
			...(min !== undefined ? { min } : {}),
			...(max !== undefined ? { max } : {}),
			...(man !== undefined ? { man } : {}),
			...(pt !== undefined ? { pt } : {}),
		}
		target.push(parsed)
	}
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
				rank?: number
				percent?: boolean
				bottom?: boolean
				aboveAverage?: boolean
				equalAverage?: boolean
				timePeriod?: string
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
			const rank = numAttr(rule, 'rank')
			if (rank !== undefined) parsedRule.rank = rank
			const percent = readBoolAttribute(rule, 'percent')
			if (percent !== undefined) parsedRule.percent = percent
			const bottom = readBoolAttribute(rule, 'bottom')
			if (bottom !== undefined) parsedRule.bottom = bottom
			const cfAboveAverage = readBoolAttribute(rule, 'aboveAverage')
			if (cfAboveAverage !== undefined) parsedRule.aboveAverage = cfAboveAverage
			const equalAverage = readBoolAttribute(rule, 'equalAverage')
			if (equalAverage !== undefined) parsedRule.equalAverage = equalAverage
			const timePeriod = attr(rule, 'timePeriod')
			if (timePeriod) parsedRule.timePeriod = timePeriod
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
