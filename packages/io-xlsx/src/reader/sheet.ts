import type {
	AutoFilter,
	Cell,
	CellStyle,
	DynamicArrayFormulaInfo,
	RangeRef,
	SheetAdvancedFilterInfo,
	SheetBreak,
	SheetColDef,
	SheetConditionalFormat,
	SheetConditionalFormatColor,
	SheetConditionalFormatRule,
	SheetConditionalFormatValueObject,
	SheetDataValidation,
	SheetSparklineGroupInfo,
	SheetX14ConditionalFormatDataBarInfo,
	SheetX14ConditionalFormatIconInfo,
	SheetX14ConditionalFormatIconSetInfo,
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
import { parseAutoFilterNode, parseSortStateNode } from './filtering.ts'
import type { ParsedMetadataPart } from './metadata.ts'
import type { Relationship } from './relationships.ts'
import type { SharedStringResolver } from './shared-strings.ts'
import {
	decodeXmlText,
	findTagEnd,
	isSelfClosingTag,
	normalizeMainSpreadsheetNamespacePrefix,
} from './xml-utils.ts'

const SMALL_NUMBER_RANGE_START = -128
const SMALL_NUMBER_RANGE_END = 512
const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g
const BYTE_XML_DECODER = new TextDecoder('utf-8')
const TEXT_NODE_RE =
	/<([A-Za-z_][\w:.-]*)\b([^>]*)>([\s\S]*?)<\/\1>|<([A-Za-z_][\w:.-]*)\b([^>]*)\/>/g
const X14_DATA_VALIDATION_RE =
	/<([A-Za-z_][\w.-]*):dataValidation\b([^>]*)>([\s\S]*?)<\/\1:dataValidation>/gi
const X14_SELF_CLOSING_DATA_VALIDATION_RE = /<([A-Za-z_][\w.-]*):dataValidation\b([^>]*)\/>/gi
const X14_CONDITIONAL_FORMATTING_RE =
	/<([A-Za-z_][\w.-]*):conditionalFormatting\b([^>]*)>([\s\S]*?)<\/\1:conditionalFormatting>/gi

interface ParsedFormulaText {
	readonly text: string | null
	readonly storedText?: string
	readonly info?: Cell['formulaInfo']
}

const NO_FORMULA: ParsedFormulaText = { text: null, info: undefined }
const NULL_FORMULA_TEXT: ParsedFormulaText = { text: null }

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
	readonly hasDateStyles?: boolean
	readonly differentialStyles?: readonly CellStyle[]
	readonly relationships?: readonly Relationship[]
	readonly valuePool?: ValueInternPool
	readonly valuesOnly?: boolean
	readonly formulaOnly?: boolean
	readonly richMetadata?: boolean
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
		if (rows >= 1_048_576 && cols >= 16_384) return
		const numChunks = Math.ceil(rows / CHUNK_SIZE) * Math.ceil(cols / CHUNK_SIZE)
		const cellsPerChunk = (rows * cols) / numChunks
		if (cellsPerChunk >= SPARSE_TO_DENSE_THRESHOLD) {
			sheet.cells.setExpectedDensity('dense')
		}
	} catch {
		// Ignore invalid dimension refs; the parser will still hydrate cells normally.
	}
}

export function parseSheet(
	name: string,
	xml: string,
	ctx: SheetParseContext,
	sheetId?: Sheet['id'],
): Sheet {
	xml = normalizeMainSpreadsheetNamespacePrefix(xml)
	const sheet = new Sheet(name, sheetId)
	applyDensityHintFromDimension(sheet, xml)
	const sheetDataLoc = locateSheetData(xml)
	if (sheetDataLoc) parseSheetDataFromLoc(xml, sheetDataLoc, sheet, ctx)
	const richMetadata = ctx.richMetadata === true
	if (ctx.formulaOnly && !richMetadata) return sheet
	if (ctx.valuesOnly && !richMetadata && !hasValuesModeSheetMetadata(xml, sheetDataLoc)) {
		return sheet
	}
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
	parseSortState(ws, sheet)
	parseSheetProtection(ws, sheet)
	parsePageMargins(ws, sheet)
	parsePageSetup(ws, sheet)
	parsePrintOptions(ws, sheet)
	parseHeaderFooter(ws, sheet)
	parsePageBreaks(ws, sheet)
	parseIgnoredErrors(ws, sheet)
	if (richMetadata) {
		parseHyperlinks(ws, sheet, ctx.relationships ?? [], ctx.valuePool)
		parseConditionalFormatting(ws, sheet, ctx.differentialStyles ?? [], ctx.valuePool)
		parseDataValidations(ws, sheet, ctx.valuePool)
		parseX14ConditionalFormats(strippedXml, sheet, ctx.valuePool)
		parseX14DataValidations(strippedXml, sheet, ctx.valuePool)
		parseAdvancedFilters(ws, sheet)
		parseSparklineGroups(ws, sheet)
		extractCustomSheetViews(xml, sheet)
		extractExtLst(xml, sheet)
		extractControls(xml, sheet)
	}
	return sheet
}

export function parseSheetValuesOnlyBytes(
	name: string,
	bytes: Uint8Array,
	ctx: SheetParseContext,
	sheetId?: Sheet['id'],
): Sheet | null {
	if (!ctx.valuesOnly || ctx.richMetadata || ctx.formulaOnly) return null
	const sheetDataLoc = locateSheetDataBytes(bytes)
	if (!sheetDataLoc) return null
	if (hasPrefixedElementNameBytes(bytes)) return null
	if (hasUnsupportedValuesOnlyOuterTagsBytes(bytes, sheetDataLoc)) return null

	const sheet = new Sheet(name, sheetId)
	applyDensityHintFromDimensionBytes(sheet, bytes)
	return parseSheetDataBytes(bytes, sheetDataLoc, sheet, ctx) ? sheet : null
}

function hasValuesModeSheetMetadata(xml: string, sheetDataLoc: SheetDataLocation | null): boolean {
	const beforeEnd = sheetDataLoc?.tagStart ?? xml.length
	const afterStart = sheetDataLoc?.closeEnd ?? xml.length
	return (
		hasAnyTagInRange(xml, 0, beforeEnd, VALUES_MODE_SHEET_METADATA_TAGS) ||
		hasAnyTagInRange(xml, afterStart, xml.length, VALUES_MODE_SHEET_METADATA_TAGS)
	)
}

const VALUES_MODE_SHEET_METADATA_TAGS = [
	'sheetPr',
	'sheetFormatPr',
	'sheetViews',
	'cols',
	'mergeCells',
	'drawing',
	'legacyDrawing',
	'autoFilter',
	'sortState',
	'sheetProtection',
	'pageMargins',
	'pageSetup',
	'printOptions',
	'headerFooter',
	'rowBreaks',
	'colBreaks',
	'ignoredErrors',
] as const

function hasAnyTagInRange(
	xml: string,
	start: number,
	end: number,
	tags: readonly string[],
): boolean {
	for (const tag of tags) {
		if (hasWorksheetTagInRange(xml, start, end, tag)) return true
	}
	return false
}

function hasWorksheetTagInRange(xml: string, start: number, end: number, tagName: string): boolean {
	if (end <= start) return false
	const segment = xml.slice(start, end)
	let cursor = 0
	while (cursor < segment.length) {
		const index = segment.indexOf(`<${tagName}`, cursor)
		if (index === -1) return false
		const next = segment.charCodeAt(index + tagName.length + 1)
		if (next === 9 || next === 10 || next === 13 || next === 32 || next === 47 || next === 62) {
			return true
		}
		cursor = index + tagName.length + 1
	}
	return false
}

function hasRowPresentationMetadata(rawAttrs: string): boolean {
	return (
		rawAttrs.includes('ht="') ||
		rawAttrs.includes('customHeight="') ||
		rawAttrs.includes('hidden="') ||
		rawAttrs.includes('collapsed="') ||
		rawAttrs.includes('outlineLevel="')
	)
}

function hasRowPresentationMetadataInRange(xml: string, start: number, end: number): boolean {
	return (
		rawAttrValueStartInRange(xml, start, end, 'ht') !== -1 ||
		rawAttrValueStartInRange(xml, start, end, 'customHeight') !== -1 ||
		rawAttrValueStartInRange(xml, start, end, 'hidden') !== -1 ||
		rawAttrValueStartInRange(xml, start, end, 'collapsed') !== -1 ||
		rawAttrValueStartInRange(xml, start, end, 'outlineLevel') !== -1
	)
}

const BYTE_LT = 60
const BYTE_SLASH = 47
const BYTE_COLON = 58
const BYTE_QUESTION = 63
const BYTE_BANG = 33
const BYTE_QUOTE = 34
const BYTE_AMP = 38
const BYTE_SPACE = 32
const BYTE_TAB = 9
const BYTE_LF = 10
const BYTE_CR = 13
const BYTES_WORKSHEET_OPEN = bytesLiteral('<worksheet')
const BYTES_DIMENSION_OPEN = bytesLiteral('<dimension')
const BYTES_SHEET_DATA_OPEN = bytesLiteral('<sheetData')
const BYTES_SHEET_DATA_CLOSE = bytesLiteral('</sheetData>')
const BYTES_ROW_OPEN = bytesLiteral('<row')
const BYTES_ROW_CLOSE = bytesLiteral('</row>')
const BYTES_CELL_OPEN = bytesLiteral('<c')
const BYTES_CELL_CLOSE = bytesLiteral('</c>')
const BYTES_V_OPEN = bytesLiteral('<v')
const BYTES_V_CLOSE = bytesLiteral('</v>')
const BYTES_F_OPEN = bytesLiteral('<f')
const BYTES_IS_OPEN = bytesLiteral('<is')
const BYTES_T_OPEN = bytesLiteral('<t')
const BYTES_T_CLOSE = bytesLiteral('</t>')
const BYTES_IS_CLOSE = bytesLiteral('</is>')

const VALUES_MODE_ALLOWED_OUTER_TAGS = new Set(['worksheet', 'dimension', 'sheetData'])

function bytesLiteral(value: string): Uint8Array {
	const out = new Uint8Array(value.length)
	for (let index = 0; index < value.length; index++) out[index] = value.charCodeAt(index)
	return out
}

function parseSheetDataBytes(
	bytes: Uint8Array,
	sheetData: SheetDataLocation,
	sheet: Sheet,
	ctx: SheetParseContext,
): boolean {
	let rowCursor = sheetData.contentStart
	let currentRow = -1
	const cellOut = { row: 0, col: 0 }

	while (true) {
		const rowOpen = indexOfElementOpenBytes(bytes, BYTES_ROW_OPEN, rowCursor, sheetData.contentEnd)
		if (rowOpen === -1) return true
		const rowTagEnd = findTagEndBytes(bytes, rowOpen)
		if (rowTagEnd === -1 || rowTagEnd >= sheetData.contentEnd) return false
		const rowAttrStart = rowOpen + BYTES_ROW_OPEN.length
		const explicitRowIndex = rawPositiveIntAttrInBytes(bytes, rowAttrStart, rowTagEnd, 'r')
		const row = explicitRowIndex !== undefined ? explicitRowIndex - 1 : currentRow + 1
		currentRow = row
		if (ctx.maxRows !== undefined && row >= ctx.maxRows) return true
		parseRowPresentationMetadataBytes(bytes, rowAttrStart, rowTagEnd, row, sheet)
		if (isSelfClosingTagBytes(bytes, rowOpen, rowTagEnd)) {
			rowCursor = rowTagEnd + 1
			continue
		}

		const rowClose = indexOfBytes(bytes, BYTES_ROW_CLOSE, rowTagEnd + 1, sheetData.contentEnd)
		if (rowClose === -1) return false
		let cellCursor = rowTagEnd + 1
		let nextCol = 0
		while (true) {
			const cellOpen = indexOfElementOpenBytes(bytes, BYTES_CELL_OPEN, cellCursor, rowClose)
			if (cellOpen === -1) break
			const cellTagEnd = findTagEndBytes(bytes, cellOpen)
			if (cellTagEnd === -1 || cellTagEnd > rowClose) return false
			const selfClosing = isSelfClosingTagBytes(bytes, cellOpen, cellTagEnd)
			let cellClose = -1
			if (!selfClosing) {
				cellClose = indexOfBytes(bytes, BYTES_CELL_CLOSE, cellTagEnd + 1, rowClose)
				if (cellClose === -1) return false
			}
			const parsedCell = parseValuesOnlyCellBytes(
				bytes,
				cellOpen + BYTES_CELL_OPEN.length,
				cellTagEnd,
				cellTagEnd + 1,
				selfClosing ? cellTagEnd + 1 : cellClose,
				selfClosing,
				row,
				nextCol,
				ctx,
				sheet,
				cellOut,
			)
			if (parsedCell === false) {
				return false
			}
			nextCol = cellOut.col + 1
			cellCursor = selfClosing ? cellTagEnd + 1 : cellClose + BYTES_CELL_CLOSE.length
		}
		rowCursor = rowClose + BYTES_ROW_CLOSE.length
	}
}

function parseValuesOnlyCellBytes(
	bytes: Uint8Array,
	attrStart: number,
	attrEnd: number,
	bodyStart: number,
	bodyEnd: number,
	selfClosing: boolean,
	fallbackRow: number,
	fallbackCol: number,
	ctx: SheetParseContext,
	sheet: Sheet,
	out: { row: number; col: number },
): 'set' | 'skip' | false {
	if (!resolveCellPositionBytes(bytes, attrStart, attrEnd, fallbackRow, fallbackCol, out)) {
		return false
	}
	const type = rawAttrAsciiBytes(bytes, attrStart, attrEnd, 't')
	const styleIdx =
		ctx.hasDateStyles && type !== 's' ? (rawNumAttrInBytes(bytes, attrStart, attrEnd, 's') ?? 0) : 0
	if (selfClosing) {
		if (type !== 'n') return 'skip'
		sheet.cells.setPlainNumber(out.row, out.col, 0)
		return 'set'
	}

	if (type === 'inlineStr' || startsWithInlineStringBytes(bytes, bodyStart, bodyEnd)) {
		const text = parseInlineStringTextBytes(bytes, bodyStart, bodyEnd)
		if (text === undefined) return false
		sheet.cells.setStringResolved(out.row, out.col, text, null, DEFAULT_STYLE_ID)
		return 'set'
	}

	const rawValue = extractVTagTextBytes(bytes, bodyStart, bodyEnd)
	if (type === 's') {
		const idx = rawValue !== undefined ? fastParseNonNegInt(rawValue) : -1
		if (idx < 0) {
			sheet.cells.setStringResolved(out.row, out.col, '', null, DEFAULT_STYLE_ID)
			return 'set'
		}
		const text = ctx.sharedStrings.getString?.(idx)
		if (text !== undefined) {
			sheet.cells.setStringResolved(out.row, out.col, text, null, DEFAULT_STYLE_ID)
			return 'set'
		}
		const entry = ctx.sharedStrings.get(idx)
		sheet.cells.setResolved(out.row, out.col, entry ?? stringValue(''), null, DEFAULT_STYLE_ID)
		return 'set'
	}
	if (type === 'b') {
		sheet.cells.setResolved(
			out.row,
			out.col,
			internValue(ctx, booleanValue(rawValue === '1')),
			null,
			DEFAULT_STYLE_ID,
		)
		return 'set'
	}
	if (type === 'e') {
		sheet.cells.setResolved(
			out.row,
			out.col,
			internValue(ctx, errorValue((rawValue ?? '#VALUE!') as ExcelError)),
			null,
			DEFAULT_STYLE_ID,
		)
		return 'set'
	}
	if (type === 'str') {
		sheet.cells.setStringResolved(out.row, out.col, rawValue ?? '', null, DEFAULT_STYLE_ID)
		return 'set'
	}
	if (rawValue !== undefined && rawValue !== '') {
		const num = Number(rawValue)
		if (Number.isNaN(num)) {
			sheet.cells.setStringResolved(out.row, out.col, rawValue, null, DEFAULT_STYLE_ID)
		} else if (ctx.isDateFormat[styleIdx]) {
			sheet.cells.setResolved(
				out.row,
				out.col,
				internValue(ctx, dateValue(num)),
				null,
				DEFAULT_STYLE_ID,
			)
		} else {
			sheet.cells.setPlainNumber(out.row, out.col, num)
		}
		return 'set'
	}
	if (hasElementOpenBytes(bytes, BYTES_F_OPEN, bodyStart, bodyEnd)) {
		sheet.cells.setResolved(out.row, out.col, EMPTY, null, DEFAULT_STYLE_ID)
		return 'set'
	}
	if (type) {
		return setValuesOnlyTypedEmptyCell(sheet, out.row, out.col, type) ? 'set' : 'skip'
	}
	return 'skip'
}

function setValuesOnlyTypedEmptyCell(
	sheet: Sheet,
	row: number,
	col: number,
	type: string | undefined,
): boolean {
	if (type === undefined) return false
	if (type === 'n') {
		sheet.cells.setPlainNumber(row, col, 0)
	} else {
		sheet.cells.setResolved(row, col, EMPTY, null, DEFAULT_STYLE_ID)
	}
	return true
}

function parseRowPresentationMetadataBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	row: number,
	sheet: Sheet,
): void {
	const rowHeight = rawNumAttrInBytes(bytes, start, end, 'ht')
	if (rowHeight !== undefined && rawAttrAsciiBytes(bytes, start, end, 'customHeight') === '1') {
		sheet.rowHeights.set(row, rowHeight)
	}
	const hidden = rawBoolAttrInBytes(bytes, start, end, 'hidden')
	const collapsed = rawBoolAttrInBytes(bytes, start, end, 'collapsed')
	const outlineLevel = rawNumAttrInBytes(bytes, start, end, 'outlineLevel')
	if (hidden === undefined && collapsed === undefined && outlineLevel === undefined) return
	const rowDef: Record<string, boolean | number> = {}
	if (hidden !== undefined) rowDef.hidden = hidden
	if (collapsed !== undefined) rowDef.collapsed = collapsed
	if (outlineLevel !== undefined) rowDef.outlineLevel = outlineLevel
	sheet.rowDefs.set(row, rowDef as import('@ascend/core').SheetRowDef)
}

function parseSheetDataFromLoc(
	xml: string,
	sheetData: SheetDataLocation,
	sheet: Sheet,
	ctx: SheetParseContext,
): void {
	const sharedFormulaMasters: SharedFormulaMasterMap = new Map()
	const cellCtx =
		ctx.formulaOnly && !sheetDataHasFormula(xml, sheetData) ? { ...ctx, valuesOnly: true } : ctx
	let rowCursor = sheetData.contentStart
	let currentRow = -1
	const fallbackPos = { row: 0, col: 0 }
	const cellOut = { row: 0, col: 0 }

	while (true) {
		const rowOpen = xml.indexOf('<row', rowCursor)
		if (rowOpen === -1 || rowOpen >= sheetData.contentEnd) return
		const rowTagEnd = findTagEnd(xml, rowOpen)
		if (rowTagEnd === -1 || rowTagEnd >= sheetData.contentEnd) return
		const rowAttrStart = rowOpen + 4
		const explicitRowIndex = rawPositiveIntAttrInRange(xml, rowAttrStart, rowTagEnd, 'r')
		const row = explicitRowIndex !== undefined ? explicitRowIndex - 1 : currentRow + 1
		currentRow = row
		if (ctx.maxRows !== undefined && row >= ctx.maxRows) return
		if (hasRowPresentationMetadataInRange(xml, rowAttrStart, rowTagEnd)) {
			const rowAttrsRaw = xml.slice(rowAttrStart, rowTagEnd)
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
		}
		if (isSelfClosingTag(xml, rowOpen, rowTagEnd)) {
			rowCursor = rowTagEnd + 1
			continue
		}

		const rowClose = xml.indexOf('</row>', rowTagEnd + 1)
		if (rowClose === -1 || rowClose > sheetData.contentEnd) return
		if (parseSimpleValuesRow(xml, rowTagEnd + 1, rowClose, row, cellCtx, sheet)) {
			rowCursor = rowClose + 6
			continue
		}
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
			fallbackPos.row = row
			fallbackPos.col = nextCol
			if (cellCtx.valuesOnly && !selfClosing && cellClose !== -1 && cellClose <= rowClose) {
				if (
					parseSimpleValuesNumberCell(
						rawAttrs,
						xml,
						cellTagEnd + 1,
						cellClose,
						cellCtx,
						fallbackPos,
						sheet,
						cellOut,
					)
				) {
					cellCursor = cellClose + 4
					nextCol = cellOut.col + 1
					continue
				}
				if (
					parseSimpleValuesSharedStringCell(
						rawAttrs,
						xml,
						cellTagEnd + 1,
						cellClose,
						cellCtx,
						fallbackPos,
						sheet,
						cellOut,
					)
				) {
					cellCursor = cellClose + 4
					nextCol = cellOut.col + 1
					continue
				}
				if (
					parseSimpleValuesPlainStringCell(
						rawAttrs,
						xml,
						cellTagEnd + 1,
						cellClose,
						cellCtx,
						fallbackPos,
						sheet,
						cellOut,
					)
				) {
					cellCursor = cellClose + 4
					nextCol = cellOut.col + 1
					continue
				}
				if (
					parseSimpleValuesInlineStringCell(
						rawAttrs,
						xml,
						cellTagEnd + 1,
						cellClose,
						cellCtx,
						fallbackPos,
						sheet,
						cellOut,
					)
				) {
					cellCursor = cellClose + 4
					nextCol = cellOut.col + 1
					continue
				}
			}
			const innerXml =
				!selfClosing && cellClose !== -1 && cellClose <= rowClose
					? xml.slice(cellTagEnd + 1, cellClose)
					: ''
			const ok = parseFastCell(
				rawAttrs,
				innerXml,
				cellCtx,
				sharedFormulaMasters,
				fallbackPos,
				sheet,
				cellOut,
			)
			cellCursor =
				selfClosing || cellClose === -1 || cellClose > rowClose ? cellTagEnd + 1 : cellClose + 4
			if (!ok) {
				if (resolveCellPositionInto(rawAttrs, fallbackPos, cellOut)) nextCol = cellOut.col + 1
				continue
			}
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
		const rowCloseForSlice = xml.indexOf('</row>', rowTagEnd + 1)
		const rowEndForSlice = isSelfClosingTag(xml, rowOpen, rowTagEnd)
			? rowTagEnd + 1
			: rowCloseForSlice === -1 || rowCloseForSlice > sheetData.contentEnd
				? -1
				: rowCloseForSlice + 6
		if (rowEndForSlice === -1) return
		const directParsed = parseStreamedValuesRowXml(xml, rowOpen, rowEndForSlice, currentRow, ctx)
		if (directParsed) {
			currentRow = directParsed.row
			yield directParsed
			rowCursor = rowEndForSlice
			continue
		}
		const parsed = parseStreamedSheetRowXml(
			xml.slice(rowOpen, rowEndForSlice),
			currentRow,
			ctx,
			sharedFormulaMasters,
			rowSheet,
			fallbackPos,
			cellOut,
		)
		if (!parsed) return
		currentRow = parsed.row
		yield parsed
		rowCursor = rowEndForSlice
	}
}

export function* streamSheetRowsTextChunks(
	name: string,
	chunks: Iterable<string>,
	ctx: SheetParseContext,
): Generator<StreamedSheetRow> {
	const sharedFormulaMasters: SharedFormulaMasterMap = new Map()
	let currentRow = -1
	let buffer = ''
	let inSheetData = false
	const fallbackPos = { row: 0, col: 0 }
	const cellOut = { row: 0, col: 0 }
	const rowSheet = new Sheet(name)
	for (const chunk of chunks) {
		buffer += chunk
		while (true) {
			if (!inSheetData) {
				const sheetDataOpen = buffer.indexOf('<sheetData')
				if (sheetDataOpen === -1) {
					buffer = buffer.slice(Math.max(0, buffer.length - '<sheetData'.length))
					break
				}
				const sheetDataTagEnd = findTagEnd(buffer, sheetDataOpen)
				if (sheetDataTagEnd === -1) {
					buffer = buffer.slice(sheetDataOpen)
					break
				}
				if (isSelfClosingTag(buffer, sheetDataOpen, sheetDataTagEnd)) return
				buffer = buffer.slice(sheetDataTagEnd + 1)
				inSheetData = true
			}

			const sheetDataClose = buffer.indexOf('</sheetData>')
			const rowOpen = buffer.indexOf('<row')
			if (rowOpen === -1 || (sheetDataClose !== -1 && sheetDataClose < rowOpen)) return
			if (rowOpen > 0) {
				buffer = buffer.slice(rowOpen)
			}
			const rowTagEnd = findTagEnd(buffer, 0)
			if (rowTagEnd === -1) break
			let rowEnd: number
			if (isSelfClosingTag(buffer, 0, rowTagEnd)) {
				rowEnd = rowTagEnd + 1
			} else {
				const rowClose = buffer.indexOf('</row>', rowTagEnd + 1)
				if (rowClose === -1) break
				rowEnd = rowClose + 6
			}
			const directParsed = parseStreamedValuesRowXml(buffer, 0, rowEnd, currentRow, ctx)
			if (directParsed) {
				buffer = buffer.slice(rowEnd)
				currentRow = directParsed.row
				yield directParsed
				continue
			}
			const parsed = parseStreamedSheetRowXml(
				buffer.slice(0, rowEnd),
				currentRow,
				ctx,
				sharedFormulaMasters,
				rowSheet,
				fallbackPos,
				cellOut,
			)
			buffer = buffer.slice(rowEnd)
			if (!parsed) continue
			currentRow = parsed.row
			yield parsed
		}
	}
}

function parseStreamedValuesRowXml(
	xml: string,
	rowOpen: number,
	rowEnd: number,
	currentRow: number,
	ctx: SheetParseContext,
): StreamedSheetRow | null {
	if (!ctx.valuesOnly || ctx.formulaOnly) return null
	const rowTagEnd = findTagEnd(xml, rowOpen)
	if (rowTagEnd === -1 || rowTagEnd > rowEnd) return null
	const rowAttrsRaw = xml.slice(rowOpen + 4, rowTagEnd)
	const explicitRowIndex = rawNumAttr(rowAttrsRaw, 'r')
	const row = explicitRowIndex !== undefined ? explicitRowIndex - 1 : currentRow + 1
	if (ctx.maxRows !== undefined && row >= ctx.maxRows) return null
	if (isSelfClosingTag(xml, rowOpen, rowTagEnd)) return { row, cells: [] }

	const rowClose = rowEnd >= 6 && xml.startsWith('</row>', rowEnd - 6) ? rowEnd - 6 : -1
	if (rowClose === -1) return null
	const canonical = parseCanonicalStreamedValuesRow(xml, rowTagEnd + 1, rowClose, row, ctx)
	if (canonical) return canonical
	let cellCursor = rowTagEnd + 1
	let nextCol = 0
	const cells: [number, Cell][] = []
	const fallbackPos = { row, col: 0 }
	const out = { row, col: 0 }
	while (true) {
		const cellOpen = xml.indexOf('<c', cellCursor)
		if (cellOpen === -1 || cellOpen >= rowClose) break
		const cellTagEnd = findTagEnd(xml, cellOpen)
		if (cellTagEnd === -1 || cellTagEnd > rowClose) return null
		const rawAttrs = xml.slice(cellOpen + 2, cellTagEnd)
		const selfClosing = isSelfClosingTag(xml, cellOpen, cellTagEnd)
		const cellClose = selfClosing ? -1 : xml.indexOf('</c>', cellTagEnd + 1)
		if (!selfClosing && (cellClose === -1 || cellClose > rowClose)) return null
		fallbackPos.col = nextCol
		const parsed = parseDirectValuesCell(
			rawAttrs,
			xml,
			cellTagEnd + 1,
			selfClosing ? cellTagEnd + 1 : cellClose,
			selfClosing,
			ctx,
			fallbackPos,
			out,
		)
		if (parsed === undefined) return null
		if (parsed) {
			cells.push([out.col, parsed])
		}
		nextCol = out.col + 1
		cellCursor = selfClosing ? cellTagEnd + 1 : cellClose + 4
	}
	return { row, cells }
}

function parseCanonicalStreamedValuesRow(
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	row: number,
	ctx: SheetParseContext,
): StreamedSheetRow | null {
	if (!ctx.valuesOnly || ctx.hasDateStyles) return null
	let cursor = bodyStart
	let nextCol = 0
	const rowText = String(row + 1)
	const cells: [number, Cell][] = []
	const out: {
		row: number
		col: number
		numberValue: number | undefined
		stringValue: string | undefined
	} = {
		row,
		col: 0,
		numberValue: undefined,
		stringValue: undefined,
	}
	while (true) {
		cursor = skipXmlWhitespace(xml, cursor, bodyEnd)
		if (cursor >= bodyEnd) return { row, cells }
		const canonicalNext = parseCanonicalValuesCell(xml, cursor, bodyEnd, rowText, row, nextCol, out)
		if (canonicalNext === -1) return null
		const value =
			out.numberValue !== undefined
				? internValue(ctx, numberValue(out.numberValue))
				: out.stringValue !== undefined
					? internValue(ctx, stringValue(out.stringValue))
					: undefined
		if (value === undefined) return null
		cells.push([out.col, { value, formula: null, styleId: DEFAULT_STYLE_ID }])
		nextCol = out.col + 1
		cursor = canonicalNext
	}
}

function parseDirectValuesCell(
	rawAttrs: string,
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	selfClosing: boolean,
	ctx: SheetParseContext,
	fallbackPosition: CellPosition,
	out: { row: number; col: number },
): Cell | null | undefined {
	if (rawAttrs.includes('cm="') || rawAttrs.includes('vm="')) return undefined
	if (!resolveCellPositionInto(rawAttrs, fallbackPosition, out)) return undefined
	const type = rawAttr(rawAttrs, 't')
	const styleIdx = ctx.hasDateStyles && type !== 's' ? (rawNumAttr(rawAttrs, 's') ?? 0) : 0
	const styleId = DEFAULT_STYLE_ID
	if (selfClosing) {
		if (type === 'n') return { value: internValue(ctx, numberValue(0)), formula: null, styleId }
		return null
	}

	if (type === 'inlineStr' || xml.startsWith('<is', skipXmlWhitespace(xml, bodyStart, bodyEnd))) {
		const text = parseSimpleInlineStringText(xml, bodyStart, bodyEnd)
		if (text === undefined) return undefined
		return { value: internValue(ctx, stringValue(text)), formula: null, styleId }
	}

	const rawValue = extractVTagTextInRange(xml, bodyStart, bodyEnd)
	if (type === 's') {
		const idx = rawValue !== undefined ? fastParseNonNegInt(rawValue) : -1
		if (idx < 0) return { value: internValue(ctx, stringValue('')), formula: null, styleId }
		const entry = ctx.sharedStrings.get(idx)
		if (entry) return { value: internValue(ctx, entry), formula: null, styleId }
		return { value: internValue(ctx, stringValue('')), formula: null, styleId }
	}
	if (type === 'b') {
		return { value: internValue(ctx, booleanValue(rawValue === '1')), formula: null, styleId }
	}
	if (type === 'e') {
		return {
			value: internValue(ctx, errorValue((rawValue ?? '#VALUE!') as ExcelError)),
			formula: null,
			styleId,
		}
	}
	if (type === 'str') {
		return { value: internValue(ctx, stringValue(rawValue ?? '')), formula: null, styleId }
	}
	if (rawValue !== undefined && rawValue !== '') {
		const num = Number(rawValue)
		if (Number.isNaN(num)) {
			return { value: internValue(ctx, stringValue(rawValue)), formula: null, styleId }
		}
		if (ctx.isDateFormat[styleIdx]) {
			return { value: internValue(ctx, dateValue(num)), formula: null, styleId }
		}
		return { value: internValue(ctx, numberValue(num)), formula: null, styleId }
	}
	if (hasFormulaTagInRange(xml, bodyStart, bodyEnd)) {
		return { value: EMPTY, formula: null, styleId }
	}
	if (type === 'n') return { value: internValue(ctx, numberValue(0)), formula: null, styleId }
	if (type) return { value: EMPTY, formula: null, styleId }
	return null
}

function internValue(ctx: SheetParseContext, value: CellValue): CellValue {
	return ctx.valuePool ? ctx.valuePool.internValue(value) : value
}

function extractVTagTextInRange(xml: string, start: number, end: number): string | undefined {
	const open = xml.indexOf('<v', start)
	if (open === -1 || open >= end) return undefined
	const contentStart = xml.indexOf('>', open)
	if (contentStart === -1 || contentStart >= end) return undefined
	const close = xml.indexOf('</v>', contentStart + 1)
	if (close === -1 || close > end) return undefined
	const slice = xml.slice(contentStart + 1, close)
	return slice.includes('&') ? decodeXmlText(slice) : slice
}

function parseSimpleInlineStringText(xml: string, start: number, end: number): string | undefined {
	let cursor = skipXmlWhitespace(xml, start, end)
	if (xml.startsWith('<f', cursor)) {
		cursor = skipSimpleElement(xml, cursor, end, 'f')
		if (cursor === -1) return undefined
		cursor = skipXmlWhitespace(xml, cursor, end)
	}
	if (!xml.startsWith('<is', cursor)) return undefined
	const isTagEnd = findTagEnd(xml, cursor)
	if (isTagEnd === -1 || isTagEnd >= end || isSelfClosingTag(xml, cursor, isTagEnd)) {
		return undefined
	}
	cursor = skipXmlWhitespace(xml, isTagEnd + 1, end)
	if (!xml.startsWith('<t', cursor)) return undefined
	const textTagEnd = findTagEnd(xml, cursor)
	if (textTagEnd === -1 || textTagEnd >= end || isSelfClosingTag(xml, cursor, textTagEnd)) {
		return undefined
	}
	const textClose = xml.indexOf('</t>', textTagEnd + 1)
	if (textClose === -1 || textClose > end) return undefined
	const rawText = xml.slice(textTagEnd + 1, textClose)
	if (rawText.includes('<')) return undefined
	cursor = skipXmlWhitespace(xml, textClose + 4, end)
	if (!xml.startsWith('</is>', cursor)) return undefined
	cursor = skipXmlWhitespace(xml, cursor + 5, end)
	if (cursor !== end) return undefined
	return rawText.includes('&') ? decodeXmlText(rawText) : rawText
}

function hasFormulaTagInRange(xml: string, start: number, end: number): boolean {
	const formulaOpen = xml.indexOf('<f', start)
	return formulaOpen !== -1 && formulaOpen < end
}

function parseStreamedSheetRowXml(
	rowXml: string,
	currentRow: number,
	ctx: SheetParseContext,
	sharedFormulaMasters: SharedFormulaMasterMap,
	rowSheet: Sheet,
	fallbackPos: { row: number; col: number },
	cellOut: { row: number; col: number },
): StreamedSheetRow | null {
	const rowOpen = rowXml.indexOf('<row')
	if (rowOpen === -1) return null
	const rowTagEnd = findTagEnd(rowXml, rowOpen)
	if (rowTagEnd === -1) return null
	const rowAttrsRaw = rowXml.slice(rowOpen + 4, rowTagEnd)
	const explicitRowIndex = rawNumAttr(rowAttrsRaw, 'r')
	const row = explicitRowIndex !== undefined ? explicitRowIndex - 1 : currentRow + 1
	if (ctx.maxRows !== undefined && row >= ctx.maxRows) return null
	rowSheet.cells.clear()
	rowSheet.rowDefs.clear()
	if (hasRowPresentationMetadata(rowAttrsRaw)) {
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
	}
	if (isSelfClosingTag(rowXml, rowOpen, rowTagEnd)) return { row, cells: [] }

	const rowClose = rowXml.indexOf('</row>', rowTagEnd + 1)
	if (rowClose === -1) return null
	let cellCursor = rowTagEnd + 1
	let nextCol = 0
	while (true) {
		const cellOpen = rowXml.indexOf('<c', cellCursor)
		if (cellOpen === -1 || cellOpen >= rowClose) break
		const cellTagEnd = findTagEnd(rowXml, cellOpen)
		if (cellTagEnd === -1 || cellTagEnd > rowClose) break
		const rawAttrs = rowXml.slice(cellOpen + 2, cellTagEnd)
		const selfClosing = isSelfClosingTag(rowXml, cellOpen, cellTagEnd)
		const cellClose = selfClosing ? -1 : rowXml.indexOf('</c>', cellTagEnd + 1)
		fallbackPos.row = row
		fallbackPos.col = nextCol
		if (ctx.valuesOnly && !selfClosing && cellClose !== -1 && cellClose <= rowClose) {
			if (
				parseSimpleValuesNumberCell(
					rawAttrs,
					rowXml,
					cellTagEnd + 1,
					cellClose,
					ctx,
					fallbackPos,
					rowSheet,
					cellOut,
				)
			) {
				cellCursor = cellClose + 4
				nextCol = cellOut.col + 1
				continue
			}
			if (
				parseSimpleValuesSharedStringCell(
					rawAttrs,
					rowXml,
					cellTagEnd + 1,
					cellClose,
					ctx,
					fallbackPos,
					rowSheet,
					cellOut,
				)
			) {
				cellCursor = cellClose + 4
				nextCol = cellOut.col + 1
				continue
			}
			if (
				parseSimpleValuesPlainStringCell(
					rawAttrs,
					rowXml,
					cellTagEnd + 1,
					cellClose,
					ctx,
					fallbackPos,
					rowSheet,
					cellOut,
				)
			) {
				cellCursor = cellClose + 4
				nextCol = cellOut.col + 1
				continue
			}
			if (
				parseSimpleValuesInlineStringCell(
					rawAttrs,
					rowXml,
					cellTagEnd + 1,
					cellClose,
					ctx,
					fallbackPos,
					rowSheet,
					cellOut,
				)
			) {
				cellCursor = cellClose + 4
				nextCol = cellOut.col + 1
				continue
			}
		}
		const innerXml =
			!selfClosing && cellClose !== -1 && cellClose <= rowClose
				? rowXml.slice(cellTagEnd + 1, cellClose)
				: ''
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
		if (!ok) {
			if (resolveCellPositionInto(rawAttrs, fallbackPos, cellOut)) nextCol = cellOut.col + 1
			continue
		}
		nextCol = cellOut.col + 1
	}
	const first = rowSheet.cells.iterateRows().next()
	return { row, cells: first.done ? [] : first.value[1] }
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
	if (fallbackPosition) {
		if (!resolveCellPositionInto(rawAttrs, fallbackPosition, out)) return false
	} else {
		const ref = rawAttr(rawAttrs, 'r')
		const pos = ref ? parseCellRef(ref) : undefined
		if (!pos) return false
		out.row = pos.row
		out.col = pos.col
	}
	const row = out.row
	const col = out.col
	const type = rawAttr(rawAttrs, 't')
	if (type === 'inlineStr' || innerXml.includes('<is')) {
		return parseSlowCell(rawAttrs, innerXml, ctx, sharedFormulaMasters, { row, col }, sheet, out)
	}

	const pool = ctx.valuePool
	const rawValue = extractVTagText(innerXml)
	const styleIdx =
		!ctx.valuesOnly ||
		(ctx.hasDateStyles && type !== 's' && rawValue !== undefined && rawValue !== '')
			? (rawNumAttr(rawAttrs, 's') ?? 0)
			: 0
	const styleId = ctx.valuesOnly ? DEFAULT_STYLE_ID : (ctx.styleIds[styleIdx] ?? DEFAULT_STYLE_ID)
	const formulaSpec = ctx.valuesOnly
		? NO_FORMULA
		: parseFormulaText(
				extractRawFormulaNode(innerXml),
				row,
				col,
				sharedFormulaMasters,
				pool,
				ctx.formulaFeatures,
			)
	const binding = ctx.valuesOnly
		? undefined
		: attachDynamicArrayBinding(
				formulaSpec.info,
				formulaSpec.text,
				rawNumAttr(rawAttrs, 'cm'),
				ctx.metadata,
				ctx.formulaFeatures,
			)

	let value: CellValue
	if (type === 's') {
		const idx = rawValue !== undefined ? fastParseNonNegInt(rawValue) : -1
		if (idx < 0) {
			if (ctx.valuesOnly) {
				sheet.cells.setStringResolved(row, col, '', null, DEFAULT_STYLE_ID)
				return true
			}
			value = pool ? pool.internValue(stringValue('')) : stringValue('')
		} else {
			if (ctx.valuesOnly) {
				const text = ctx.sharedStrings.getString?.(idx)
				if (text !== undefined) {
					sheet.cells.setStringResolved(row, col, text, null, DEFAULT_STYLE_ID)
					return true
				}
			}
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

	sheet.cells.setResolved(row, col, value, formulaSpec.text, styleId, binding)
	if (!ctx.valuesOnly && formulaSpec.text && formulaSpec.storedText !== undefined) {
		sheet.storedFormulaText.set(formulaStorageKey(row, col), formulaSpec.storedText)
	}
	return true
}

function parseSimpleValuesNumberCell(
	rawAttrs: string,
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	ctx: SheetParseContext,
	fallbackPosition: CellPosition,
	sheet: Sheet,
	out: { row: number; col: number },
): boolean {
	if (!ctx.valuesOnly || ctx.hasDateStyles) return false
	if (rawAttrs.includes('cm="') || rawAttrs.includes('vm="')) return false
	const typeStart = rawAttrValueStart(rawAttrs, 't')
	if (typeStart !== -1 && !rawAttrEquals(rawAttrs, 't', 'n')) return false
	let cursor = skipXmlWhitespace(xml, bodyStart, bodyEnd)
	if (!xml.startsWith('<v>', cursor)) return false
	cursor += 3
	const valueStart = cursor
	while (cursor < bodyEnd && !xml.startsWith('</v>', cursor)) cursor += 1
	if (cursor === bodyEnd || cursor === valueStart) return false
	const valueText = xml.slice(valueStart, cursor)
	if (valueText.includes('&')) return false
	const value = Number(valueText)
	if (Number.isNaN(value)) return false
	cursor = skipXmlWhitespace(xml, cursor + 4, bodyEnd)
	if (cursor !== bodyEnd) return false
	if (!resolveCellPositionInto(rawAttrs, fallbackPosition, out)) return false
	sheet.cells.setPlainNumber(out.row, out.col, value)
	return true
}

function parseSimpleValuesRow(
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	row: number,
	ctx: SheetParseContext,
	sheet: Sheet,
): boolean {
	if (!ctx.valuesOnly || ctx.hasDateStyles) return false
	let cursor = bodyStart
	let nextCol = 0
	const rowText = String(row + 1)
	const out: {
		row: number
		col: number
		numberValue: number | undefined
		stringValue: string | undefined
	} = {
		row,
		col: 0,
		numberValue: undefined,
		stringValue: undefined,
	}
	while (true) {
		cursor = skipXmlWhitespace(xml, cursor, bodyEnd)
		if (cursor >= bodyEnd) return true
		const canonicalNext = parseCanonicalValuesCell(xml, cursor, bodyEnd, rowText, row, nextCol, out)
		if (canonicalNext !== -1) {
			if (out.numberValue !== undefined) {
				sheet.cells.setPlainNumber(out.row, out.col, out.numberValue)
			} else if (out.stringValue !== undefined) {
				sheet.cells.setPlainString(out.row, out.col, out.stringValue)
			} else return false
			nextCol = out.col + 1
			cursor = canonicalNext
			continue
		}
		if (!xml.startsWith('<c', cursor)) return false
		const tagEnd = xml.indexOf('>', cursor + 2)
		if (tagEnd === -1 || tagEnd >= bodyEnd) return false
		if (!resolveSimpleNumberCellOpen(xml, cursor + 2, tagEnd, row, nextCol, out)) {
			return false
		}
		cursor = skipXmlWhitespace(xml, tagEnd + 1, bodyEnd)
		if (!xml.startsWith('<v>', cursor)) return false
		cursor += 3
		const valueStart = cursor
		const valueEnd = xml.indexOf('</v>', valueStart)
		if (valueEnd === -1 || valueEnd > bodyEnd || valueEnd === valueStart) return false
		const value = parseSimpleXmlNumber(xml, valueStart, valueEnd)
		if (value === undefined) return false
		cursor = skipXmlWhitespace(xml, valueEnd + 4, bodyEnd)
		if (!xml.startsWith('</c>', cursor)) return false
		sheet.cells.setPlainNumber(out.row, out.col, value)
		nextCol = out.col + 1
		cursor += 4
	}
}

function parseCanonicalValuesCell(
	xml: string,
	cursor: number,
	bodyEnd: number,
	fallbackRowText: string,
	fallbackRow: number,
	fallbackCol: number,
	out: {
		row: number
		col: number
		numberValue: number | undefined
		stringValue: string | undefined
	},
): number {
	if (!xml.startsWith('<c r="', cursor)) return -1
	let index = cursor + 6
	let row = fallbackRow
	let col = fallbackCol
	const expectedRefEnd = consumeExpectedCellRef(xml, index, bodyEnd, fallbackRowText, fallbackCol)
	if (expectedRefEnd === -1) {
		const parsed = parseCellRefInXml(xml, index, bodyEnd)
		if (!parsed) return -1
		index = parsed.end
		row = parsed.row
		col = parsed.col
	} else {
		index = expectedRefEnd
	}
	out.row = row
	out.col = col
	out.numberValue = undefined
	out.stringValue = undefined

	const inlineValueStart = parseCanonicalInlineStringValueStart(xml, index)
	if (inlineValueStart !== -1) {
		let valueEnd = inlineValueStart
		let hasEntity = false
		while (valueEnd < bodyEnd) {
			const code = xml.charCodeAt(valueEnd)
			if (code === 60) break
			if (code === 38) hasEntity = true
			valueEnd += 1
		}
		if (!xml.startsWith('</t></is></c>', valueEnd)) return -1
		const rawText = xml.slice(inlineValueStart, valueEnd)
		out.stringValue = hasEntity ? decodeXmlText(rawText) : rawText
		return valueEnd + 13
	}

	const contentStart = resolveCanonicalNumberContentStart(xml, index, bodyEnd)
	if (contentStart === -1) return -1
	const valueStart = resolveCanonicalValueStart(xml, contentStart, bodyEnd)
	if (valueStart === -1) return -1
	const parsedInt = parseCanonicalIntegerValue(xml, valueStart, bodyEnd)
	if (parsedInt) {
		out.numberValue = parsedInt.value
		return parsedInt.next
	}
	const valueEnd = xml.indexOf('</v></c>', valueStart)
	if (valueEnd === -1 || valueEnd > bodyEnd) return -1
	const value = parseSimpleXmlNumber(xml, valueStart, valueEnd)
	if (value === undefined) return -1
	out.numberValue = value
	return valueEnd + 8
}

function parseCanonicalInlineStringValueStart(xml: string, refEnd: number): number {
	const prefix = '" t="inlineStr"><is><t>'
	return xml.startsWith(prefix, refEnd) ? refEnd + prefix.length : -1
}

function resolveCanonicalNumberContentStart(xml: string, refEnd: number, bodyEnd: number): number {
	if (xml.startsWith('">', refEnd)) return refEnd + 2
	if (!xml.startsWith('" s="', refEnd)) return -1
	const singleDigitStyleEnd = refEnd + 7
	if (
		singleDigitStyleEnd < bodyEnd &&
		xml.charCodeAt(refEnd + 5) >= 48 &&
		xml.charCodeAt(refEnd + 5) <= 57 &&
		xml.charCodeAt(refEnd + 6) === 34 &&
		xml.charCodeAt(singleDigitStyleEnd) === 62
	) {
		return singleDigitStyleEnd + 1
	}
	const styleEnd = xml.indexOf('"', refEnd + 5)
	if (styleEnd === -1 || styleEnd >= bodyEnd || xml.charCodeAt(styleEnd + 1) !== 62) return -1
	return styleEnd + 2
}

function resolveCanonicalValueStart(xml: string, contentStart: number, bodyEnd: number): number {
	if (xml.startsWith('<v>', contentStart)) return contentStart + 3
	if (!xml.startsWith('<f', contentStart)) return -1
	if (xml.startsWith('<f>', contentStart)) {
		const formulaClose = xml.indexOf('</f>', contentStart + 3)
		const valueOpen = formulaClose + 4
		return formulaClose !== -1 && valueOpen < bodyEnd && xml.startsWith('<v>', valueOpen)
			? valueOpen + 3
			: -1
	}
	const formulaTagEnd = findTagEnd(xml, contentStart)
	if (formulaTagEnd === -1 || formulaTagEnd >= bodyEnd) return -1
	const formulaEnd = isSelfClosingTag(xml, contentStart, formulaTagEnd)
		? formulaTagEnd + 1
		: xml.indexOf('</f>', formulaTagEnd + 1)
	if (formulaEnd === -1 || formulaEnd >= bodyEnd) return -1
	const valueOpen = isSelfClosingTag(xml, contentStart, formulaTagEnd) ? formulaEnd : formulaEnd + 4
	return xml.startsWith('<v>', valueOpen) ? valueOpen + 3 : -1
}

function parseCanonicalIntegerValue(
	xml: string,
	start: number,
	bodyEnd: number,
): { value: number; next: number } | undefined {
	let cursor = start
	let sign = 1
	if (xml.charCodeAt(cursor) === 45) {
		sign = -1
		cursor += 1
	}
	let value = 0
	const digitStart = cursor
	while (cursor < bodyEnd) {
		const code = xml.charCodeAt(cursor)
		if (code < 48 || code > 57) break
		value = value * 10 + (code - 48)
		cursor += 1
	}
	if (cursor === digitStart || !xml.startsWith('</v></c>', cursor)) return undefined
	return { value: sign * value, next: cursor + 8 }
}

function consumeExpectedCellRef(
	xml: string,
	start: number,
	end: number,
	rowText: string,
	col: number,
): number {
	if (col < 0 || col > 25) return -1
	if (start >= end || xml.charCodeAt(start) !== 65 + col) return -1
	let cursor = start + 1
	for (let index = 0; index < rowText.length; index++) {
		if (cursor >= end) return -1
		if (xml.charCodeAt(cursor) !== rowText.charCodeAt(index)) return -1
		cursor += 1
	}
	return cursor < end && xml.charCodeAt(cursor) === 34 ? cursor : -1
}

function parseCellRefInXml(
	xml: string,
	start: number,
	end: number,
): { row: number; col: number; end: number } | undefined {
	let index = start
	let col = 0
	while (index < end) {
		const code = xml.charCodeAt(index)
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
	if (index === start || index >= end) return undefined
	let row = 0
	while (index < end) {
		const code = xml.charCodeAt(index)
		if (code < 48 || code > 57) break
		row = row * 10 + (code - 48)
		index += 1
	}
	if (row <= 0 || col <= 0 || index >= end || xml.charCodeAt(index) !== 34) return undefined
	return { row: row - 1, col: col - 1, end: index }
}

function parseSimpleXmlNumber(xml: string, start: number, end: number): number | undefined {
	if (start >= end) return undefined
	let cursor = start
	let sign = 1
	const first = xml.charCodeAt(cursor)
	if (first === 45) {
		sign = -1
		cursor += 1
		if (cursor >= end) return undefined
	}
	let value = 0
	const digitStart = cursor
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code < 48 || code > 57) break
		value = value * 10 + (code - 48)
		cursor += 1
	}
	if (cursor === end && cursor > digitStart) return sign * value
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code === 38 || code === 60) return undefined
		cursor += 1
	}
	const parsed = Number(xml.slice(start, end))
	return Number.isNaN(parsed) ? undefined : parsed
}

function resolveSimpleNumberCellOpen(
	xml: string,
	attrStart: number,
	attrEnd: number,
	fallbackRow: number,
	fallbackCol: number,
	out: { row: number; col: number },
): boolean {
	let cursor = attrStart
	let refStart = -1
	let refEnd = -1
	while (cursor < attrEnd) {
		cursor = skipXmlWhitespace(xml, cursor, attrEnd)
		if (cursor >= attrEnd) break
		if (xml.charCodeAt(cursor) === 47) return false
		const nameStart = cursor
		while (cursor < attrEnd) {
			const code = xml.charCodeAt(cursor)
			if (code === 61 || code === 9 || code === 10 || code === 13 || code === 32) break
			cursor += 1
		}
		const nameEnd = cursor
		cursor = skipXmlWhitespace(xml, cursor, attrEnd)
		if (xml.charCodeAt(cursor) !== 61) return false
		cursor = skipXmlWhitespace(xml, cursor + 1, attrEnd)
		if (xml.charCodeAt(cursor) !== 34) return false
		const valueStart = cursor + 1
		const valueEnd = xml.indexOf('"', valueStart)
		if (valueEnd === -1 || valueEnd > attrEnd) return false
		const nameLength = nameEnd - nameStart
		if (nameLength === 1) {
			const name = xml.charCodeAt(nameStart)
			if (name === 114) {
				refStart = valueStart
				refEnd = valueEnd
			} else if (name === 116) {
				if (valueEnd !== valueStart + 1 || xml.charCodeAt(valueStart) !== 110) return false
			}
		} else if (
			nameLength === 2 &&
			((xml.charCodeAt(nameStart) === 99 && xml.charCodeAt(nameStart + 1) === 109) ||
				(xml.charCodeAt(nameStart) === 118 && xml.charCodeAt(nameStart + 1) === 109))
		) {
			return false
		}
		cursor = valueEnd + 1
	}
	if (refStart === -1) {
		out.row = fallbackRow
		out.col = fallbackCol
		return true
	}
	return resolveCellPositionInXml(xml, refStart, refEnd, out)
}

function resolveCellPositionInXml(
	xml: string,
	valueStart: number,
	valueEnd: number,
	out: { row: number; col: number },
): boolean {
	let index = valueStart
	let col = 0
	while (index < valueEnd) {
		const code = xml.charCodeAt(index)
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) {
			col = col * 26 + (code - 64)
		} else if (code >= 97 && code <= 122) {
			col = col * 26 + (code - 96)
		} else {
			return false
		}
		index += 1
	}
	if (index === valueStart || index >= valueEnd) return false
	let row = 0
	while (index < valueEnd) {
		const code = xml.charCodeAt(index)
		if (code < 48 || code > 57) return false
		row = row * 10 + (code - 48)
		index += 1
	}
	if (row <= 0 || col <= 0) return false
	out.row = row - 1
	out.col = col - 1
	return true
}

function parseSimpleValuesSharedStringCell(
	rawAttrs: string,
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	ctx: SheetParseContext,
	fallbackPosition: CellPosition,
	sheet: Sheet,
	out: { row: number; col: number },
): boolean {
	if (!ctx.valuesOnly || !rawAttrEquals(rawAttrs, 't', 's')) return false
	if (rawAttrs.includes('cm="') || rawAttrs.includes('vm="')) return false
	let cursor = skipXmlWhitespace(xml, bodyStart, bodyEnd)
	if (!xml.startsWith('<v>', cursor)) return false
	cursor += 3
	let value = 0
	const valueStart = cursor
	while (cursor < bodyEnd) {
		const code = xml.charCodeAt(cursor)
		if (code < 48 || code > 57) break
		value = value * 10 + (code - 48)
		cursor += 1
	}
	if (cursor === valueStart || !xml.startsWith('</v>', cursor)) return false
	cursor = skipXmlWhitespace(xml, cursor + 4, bodyEnd)
	if (cursor !== bodyEnd) return false
	const text = ctx.sharedStrings.getString?.(value)
	if (text === undefined) return false
	if (!resolveCellPositionInto(rawAttrs, fallbackPosition, out)) return false
	sheet.cells.setStringResolved(out.row, out.col, text, null, DEFAULT_STYLE_ID)
	return true
}

function parseSimpleValuesPlainStringCell(
	rawAttrs: string,
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	ctx: SheetParseContext,
	fallbackPosition: CellPosition,
	sheet: Sheet,
	out: { row: number; col: number },
): boolean {
	if (!ctx.valuesOnly || !rawAttrEquals(rawAttrs, 't', 'str')) return false
	if (rawAttrs.includes('cm="') || rawAttrs.includes('vm="')) return false
	let cursor = skipXmlWhitespace(xml, bodyStart, bodyEnd)
	if (xml.startsWith('<f', cursor)) {
		cursor = skipSimpleElement(xml, cursor, bodyEnd, 'f')
		if (cursor === -1) return false
		cursor = skipXmlWhitespace(xml, cursor, bodyEnd)
	}
	if (!xml.startsWith('<v>', cursor)) return false
	cursor += 3
	const valueStart = cursor
	const valueEnd = xml.indexOf('</v>', valueStart)
	if (valueEnd === -1 || valueEnd > bodyEnd) return false
	const rawText = xml.slice(valueStart, valueEnd)
	if (rawText.includes('<')) return false
	cursor = skipXmlWhitespace(xml, valueEnd + 4, bodyEnd)
	if (cursor !== bodyEnd) return false
	if (!resolveCellPositionInto(rawAttrs, fallbackPosition, out)) return false
	sheet.cells.setStringResolved(out.row, out.col, decodeXmlText(rawText), null, DEFAULT_STYLE_ID)
	return true
}

function parseSimpleValuesInlineStringCell(
	rawAttrs: string,
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	ctx: SheetParseContext,
	fallbackPosition: CellPosition,
	sheet: Sheet,
	out: { row: number; col: number },
): boolean {
	if (!ctx.valuesOnly || !rawAttrEquals(rawAttrs, 't', 'inlineStr')) return false
	if (rawAttrs.includes('cm="') || rawAttrs.includes('vm="')) return false
	let cursor = skipXmlWhitespace(xml, bodyStart, bodyEnd)
	if (xml.startsWith('<f', cursor)) {
		cursor = skipSimpleElement(xml, cursor, bodyEnd, 'f')
		if (cursor === -1) return false
		cursor = skipXmlWhitespace(xml, cursor, bodyEnd)
	}
	if (!xml.startsWith('<is', cursor)) return false
	const isTagEnd = findTagEnd(xml, cursor)
	if (isTagEnd === -1 || isTagEnd >= bodyEnd || isSelfClosingTag(xml, cursor, isTagEnd)) {
		return false
	}
	cursor = skipXmlWhitespace(xml, isTagEnd + 1, bodyEnd)
	if (!xml.startsWith('<t', cursor)) return false
	const textTagEnd = findTagEnd(xml, cursor)
	if (textTagEnd === -1 || textTagEnd >= bodyEnd || isSelfClosingTag(xml, cursor, textTagEnd)) {
		return false
	}
	const textClose = xml.indexOf('</t>', textTagEnd + 1)
	if (textClose === -1 || textClose > bodyEnd) return false
	const rawText = xml.slice(textTagEnd + 1, textClose)
	if (rawText.includes('<')) return false
	cursor = skipXmlWhitespace(xml, textClose + 4, bodyEnd)
	if (!xml.startsWith('</is>', cursor)) return false
	cursor = skipXmlWhitespace(xml, cursor + 5, bodyEnd)
	if (cursor !== bodyEnd) return false
	if (!resolveCellPositionInto(rawAttrs, fallbackPosition, out)) return false
	sheet.cells.setStringResolved(out.row, out.col, decodeXmlText(rawText), null, DEFAULT_STYLE_ID)
	return true
}

function skipSimpleElement(xml: string, start: number, end: number, tagName: string): number {
	const tagEnd = findTagEnd(xml, start)
	if (tagEnd === -1 || tagEnd > end) return -1
	if (isSelfClosingTag(xml, start, tagEnd)) return tagEnd + 1
	const closeNeedle = `</${tagName}>`
	const close = xml.indexOf(closeNeedle, tagEnd + 1)
	if (close === -1 || close + closeNeedle.length > end) return -1
	return close + closeNeedle.length
}

function skipXmlWhitespace(xml: string, cursor: number, end: number): number {
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code !== 9 && code !== 10 && code !== 13 && code !== 32) break
		cursor += 1
	}
	return cursor
}

function rawAttrEquals(rawAttrs: string, name: string, expected: string): boolean {
	const valueStart = rawAttrValueStart(rawAttrs, name)
	if (valueStart === -1) return false
	const valueEnd = rawAttrs.indexOf('"', valueStart)
	return valueEnd === valueStart + expected.length && rawAttrs.startsWith(expected, valueStart)
}

function resolveCellPositionInto(
	rawAttrs: string,
	fallbackPosition: CellPosition,
	out: { row: number; col: number },
): boolean {
	const valueStart = rawAttrValueStart(rawAttrs, 'r')
	if (valueStart === -1) {
		out.row = fallbackPosition.row
		out.col = fallbackPosition.col
		return true
	}
	const valueEnd = rawAttrs.indexOf('"', valueStart)
	if (valueEnd === -1) return false
	let index = valueStart
	let col = 0
	while (index < valueEnd) {
		const code = rawAttrs.charCodeAt(index)
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) {
			col = col * 26 + (code - 64)
		} else if (code >= 97 && code <= 122) {
			col = col * 26 + (code - 96)
		} else {
			return false
		}
		index += 1
	}
	if (index === valueStart || index >= valueEnd) return false
	let row = 0
	while (index < valueEnd) {
		const code = rawAttrs.charCodeAt(index)
		if (code < 48 || code > 57) return false
		row = row * 10 + (code - 48)
		index += 1
	}
	if (row <= 0 || col <= 0) return false
	out.row = row - 1
	out.col = col - 1
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

function locateSheetDataBytes(bytes: Uint8Array): SheetDataLocation | null {
	if (indexOfElementOpenBytes(bytes, BYTES_WORKSHEET_OPEN, 0, bytes.length) === -1) return null
	const open = indexOfElementOpenBytes(bytes, BYTES_SHEET_DATA_OPEN, 0, bytes.length)
	if (open === -1) return null
	const tagEnd = findTagEndBytes(bytes, open)
	if (tagEnd === -1) return null
	if (isSelfClosingTagBytes(bytes, open, tagEnd)) {
		return {
			contentStart: tagEnd + 1,
			contentEnd: tagEnd + 1,
			tagStart: open,
			closeEnd: tagEnd + 1,
		}
	}
	const close = indexOfBytes(bytes, BYTES_SHEET_DATA_CLOSE, tagEnd + 1, bytes.length)
	if (close === -1) return null
	return {
		contentStart: tagEnd + 1,
		contentEnd: close,
		tagStart: open,
		closeEnd: close + BYTES_SHEET_DATA_CLOSE.length,
	}
}

function applyDensityHintFromDimensionBytes(sheet: Sheet, bytes: Uint8Array): void {
	const open = indexOfElementOpenBytes(bytes, BYTES_DIMENSION_OPEN, 0, bytes.length)
	if (open === -1) return
	const tagEnd = findTagEndBytes(bytes, open)
	if (tagEnd === -1) return
	const ref = rawAttrDecodedBytes(bytes, open + BYTES_DIMENSION_OPEN.length, tagEnd, 'ref')
	if (!ref) return
	try {
		const range = parseRange(ref)
		const rows = range.end.row - range.start.row + 1
		const cols = range.end.col - range.start.col + 1
		if (rows <= 0 || cols <= 0) return
		if (rows >= 1_048_576 && cols >= 16_384) return
		const numChunks = Math.ceil(rows / CHUNK_SIZE) * Math.ceil(cols / CHUNK_SIZE)
		const cellsPerChunk = (rows * cols) / numChunks
		if (cellsPerChunk >= SPARSE_TO_DENSE_THRESHOLD) {
			sheet.cells.setExpectedDensity('dense')
		}
	} catch {
		// Match the string parser: invalid dimension refs are only a skipped density hint.
	}
}

function hasPrefixedElementNameBytes(bytes: Uint8Array): boolean {
	let cursor = 0
	while (cursor < bytes.length) {
		if (bytes[cursor] !== BYTE_LT) {
			cursor += 1
			continue
		}
		cursor += 1
		const first = bytes[cursor]
		if (first === BYTE_QUESTION || first === BYTE_BANG) {
			const tagEnd = findTagEndBytes(bytes, cursor - 1)
			if (tagEnd === -1) return true
			cursor = tagEnd + 1
			continue
		}
		if (first === BYTE_SLASH) cursor += 1
		while (cursor < bytes.length) {
			const code = bytes[cursor]
			if (code === BYTE_COLON) return true
			if (isXmlNameDelimiterByte(code)) break
			cursor += 1
		}
	}
	return false
}

function hasUnsupportedValuesOnlyOuterTagsBytes(
	bytes: Uint8Array,
	sheetData: SheetDataLocation,
): boolean {
	return (
		hasUnsupportedValuesOnlyOuterTagsInRangeBytes(bytes, 0, sheetData.tagStart) ||
		hasUnsupportedValuesOnlyOuterTagsInRangeBytes(bytes, sheetData.closeEnd, bytes.length)
	)
}

function hasUnsupportedValuesOnlyOuterTagsInRangeBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
): boolean {
	let cursor = start
	while (cursor < end) {
		if (bytes[cursor] !== BYTE_LT) {
			cursor += 1
			continue
		}
		const marker = bytes[cursor + 1]
		if (marker === BYTE_QUESTION || marker === BYTE_BANG) {
			const tagEnd = findTagEndBytes(bytes, cursor)
			if (tagEnd === -1 || tagEnd > end) return true
			cursor = tagEnd + 1
			continue
		}
		const closing = marker === BYTE_SLASH
		const nameStart = cursor + (closing ? 2 : 1)
		const nameEnd = elementNameEndBytes(bytes, nameStart, end)
		if (nameEnd === -1) return true
		const name = asciiSlice(bytes, nameStart, nameEnd)
		if (!VALUES_MODE_ALLOWED_OUTER_TAGS.has(name)) return true
		cursor = nameEnd + 1
	}
	return false
}

function sheetDataHasFormula(xml: string, sheetData: SheetDataLocation): boolean {
	const formulaOpen = xml.indexOf('<f', sheetData.contentStart)
	return formulaOpen !== -1 && formulaOpen < sheetData.contentEnd
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
	const valueStart = rawAttrValueStart(rawAttrs, name)
	if (valueStart === -1) return undefined
	const valueEnd = rawAttrs.indexOf('"', valueStart)
	if (valueEnd === -1) return undefined
	return decodeXmlText(rawAttrs.slice(valueStart, valueEnd))
}

function rawNumAttr(rawAttrs: string, name: string): number | undefined {
	const valueStart = rawAttrValueStart(rawAttrs, name)
	if (valueStart === -1) return undefined
	const valueEnd = rawAttrs.indexOf('"', valueStart)
	if (valueEnd === -1) return undefined
	const parsed = Number(rawAttrs.slice(valueStart, valueEnd))
	return Number.isNaN(parsed) ? undefined : parsed
}

function rawAttrValueStart(rawAttrs: string, name: string): number {
	const needle = ATTR_NEEDLES[name] ?? `${name}="`
	const start = rawAttrs.indexOf(needle)
	return start === -1 ? -1 : start + needle.length
}

function rawAttrValueStartInRange(xml: string, start: number, end: number, name: string): number {
	const needle = ATTR_NEEDLES[name] ?? `${name}="`
	const last = end - needle.length
	for (let index = start; index <= last; index++) {
		if (xml.charCodeAt(index) === needle.charCodeAt(0) && xml.startsWith(needle, index)) {
			return index + needle.length
		}
	}
	return -1
}

function rawPositiveIntAttrInRange(
	xml: string,
	start: number,
	end: number,
	name: string,
): number | undefined {
	let cursor = rawAttrValueStartInRange(xml, start, end, name)
	if (cursor === -1) return undefined
	let value = 0
	const digitStart = cursor
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code === 34) return cursor === digitStart ? undefined : value
		if (code < 48 || code > 57) return undefined
		value = value * 10 + (code - 48)
		cursor += 1
	}
	return undefined
}

function findTagEndBytes(bytes: Uint8Array, start: number): number {
	for (let index = start + 1; index < bytes.length; index++) {
		if (bytes[index] === 62) return index
	}
	return -1
}

function isSelfClosingTagBytes(bytes: Uint8Array, tagStart: number, tagEnd: number): boolean {
	for (let index = tagEnd - 1; index > tagStart; index--) {
		const code = bytes[index]
		if (code === BYTE_SLASH) return true
		if (!isXmlWhitespaceByte(code)) return false
	}
	return false
}

function indexOfBytes(bytes: Uint8Array, needle: Uint8Array, start: number, end: number): number {
	if (needle.length === 0) return start
	const last = end - needle.length
	const first = needle[0]
	for (let index = start; index <= last; index++) {
		if (bytes[index] !== first) continue
		let matched = true
		for (let offset = 1; offset < needle.length; offset++) {
			if (bytes[index + offset] !== needle[offset]) {
				matched = false
				break
			}
		}
		if (matched) return index
	}
	return -1
}

function indexOfElementOpenBytes(
	bytes: Uint8Array,
	needle: Uint8Array,
	start: number,
	end: number,
): number {
	let cursor = start
	while (cursor < end) {
		const found = indexOfBytes(bytes, needle, cursor, end)
		if (found === -1) return -1
		const next = bytes[found + needle.length]
		if (isXmlNameDelimiterByte(next)) return found
		cursor = found + needle.length
	}
	return -1
}

function hasElementOpenBytes(
	bytes: Uint8Array,
	needle: Uint8Array,
	start: number,
	end: number,
): boolean {
	return indexOfElementOpenBytes(bytes, needle, start, end) !== -1
}

function elementNameEndBytes(bytes: Uint8Array, start: number, end: number): number {
	let cursor = start
	while (cursor < end) {
		const code = bytes[cursor]
		if (isXmlNameDelimiterByte(code)) return cursor
		cursor += 1
	}
	return -1
}

function isXmlNameDelimiterByte(code: number | undefined): boolean {
	return (
		code === undefined ||
		code === BYTE_SPACE ||
		code === BYTE_TAB ||
		code === BYTE_LF ||
		code === BYTE_CR ||
		code === BYTE_SLASH ||
		code === 62
	)
}

function isXmlWhitespaceByte(code: number | undefined): boolean {
	return code === BYTE_SPACE || code === BYTE_TAB || code === BYTE_LF || code === BYTE_CR
}

function skipXmlWhitespaceBytes(bytes: Uint8Array, cursor: number, end: number): number {
	while (cursor < end && isXmlWhitespaceByte(bytes[cursor])) cursor += 1
	return cursor
}

function rawAttrRangeBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: string,
): { start: number; end: number } | undefined {
	let cursor = start
	while (cursor < end) {
		cursor = skipXmlWhitespaceBytes(bytes, cursor, end)
		if (cursor >= end || bytes[cursor] === BYTE_SLASH) break
		const nameStart = cursor
		while (cursor < end) {
			const code = bytes[cursor]
			if (code === 61 || isXmlWhitespaceByte(code)) break
			cursor += 1
		}
		const nameEnd = cursor
		cursor = skipXmlWhitespaceBytes(bytes, cursor, end)
		if (bytes[cursor] !== 61) return undefined
		cursor = skipXmlWhitespaceBytes(bytes, cursor + 1, end)
		if (bytes[cursor] !== BYTE_QUOTE) return undefined
		const valueStart = cursor + 1
		cursor = valueStart
		while (cursor < end && bytes[cursor] !== BYTE_QUOTE) cursor += 1
		if (cursor >= end) return undefined
		if (asciiEquals(bytes, nameStart, nameEnd, name)) return { start: valueStart, end: cursor }
		cursor += 1
	}
	return undefined
}

function rawAttrAsciiBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: string,
): string | undefined {
	const range = rawAttrRangeBytes(bytes, start, end, name)
	return range ? asciiSlice(bytes, range.start, range.end) : undefined
}

function rawAttrDecodedBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: string,
): string | undefined {
	const range = rawAttrRangeBytes(bytes, start, end, name)
	return range ? decodeXmlBytesText(bytes, range.start, range.end) : undefined
}

function rawNumAttrInBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: string,
): number | undefined {
	const range = rawAttrRangeBytes(bytes, start, end, name)
	if (!range) return undefined
	const parsed = Number(asciiSlice(bytes, range.start, range.end))
	return Number.isNaN(parsed) ? undefined : parsed
}

function rawPositiveIntAttrInBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: string,
): number | undefined {
	const range = rawAttrRangeBytes(bytes, start, end, name)
	if (!range) return undefined
	let value = 0
	for (let cursor = range.start; cursor < range.end; cursor++) {
		const code = bytes[cursor] ?? -1
		if (code < 48 || code > 57) return undefined
		value = value * 10 + (code - 48)
	}
	return range.end > range.start ? value : undefined
}

function rawBoolAttrInBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: string,
): boolean | undefined {
	const value = rawAttrAsciiBytes(bytes, start, end, name)
	if (value === undefined) return undefined
	if (value === '1' || value.toLowerCase() === 'true') return true
	if (value === '0' || value.toLowerCase() === 'false') return false
	return undefined
}

function resolveCellPositionBytes(
	bytes: Uint8Array,
	attrStart: number,
	attrEnd: number,
	fallbackRow: number,
	fallbackCol: number,
	out: { row: number; col: number },
): boolean {
	const range = rawAttrRangeBytes(bytes, attrStart, attrEnd, 'r')
	if (!range) {
		out.row = fallbackRow
		out.col = fallbackCol
		return true
	}
	let cursor = range.start
	let col = 0
	while (cursor < range.end) {
		const code = bytes[cursor] ?? -1
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) {
			col = col * 26 + (code - 64)
		} else if (code >= 97 && code <= 122) {
			col = col * 26 + (code - 96)
		} else {
			return false
		}
		cursor += 1
	}
	if (cursor === range.start || cursor >= range.end) return false
	let row = 0
	while (cursor < range.end) {
		const code = bytes[cursor] ?? -1
		if (code < 48 || code > 57) return false
		row = row * 10 + (code - 48)
		cursor += 1
	}
	if (row <= 0 || col <= 0) return false
	out.row = row - 1
	out.col = col - 1
	return true
}

function extractVTagTextBytes(bytes: Uint8Array, start: number, end: number): string | undefined {
	const open = indexOfElementOpenBytes(bytes, BYTES_V_OPEN, start, end)
	if (open === -1) return undefined
	const contentStart = findTagEndBytes(bytes, open)
	if (contentStart === -1 || contentStart >= end) return undefined
	const close = indexOfBytes(bytes, BYTES_V_CLOSE, contentStart + 1, end)
	if (close === -1) return undefined
	return decodeXmlBytesText(bytes, contentStart + 1, close)
}

function startsWithInlineStringBytes(bytes: Uint8Array, start: number, end: number): boolean {
	let cursor = skipXmlWhitespaceBytes(bytes, start, end)
	if (startsWithElementOpenAtBytes(bytes, BYTES_F_OPEN, cursor, end)) {
		cursor = skipSimpleElementBytes(bytes, cursor, end, 'f')
		if (cursor === -1) return false
		cursor = skipXmlWhitespaceBytes(bytes, cursor, end)
	}
	return startsWithElementOpenAtBytes(bytes, BYTES_IS_OPEN, cursor, end)
}

function parseInlineStringTextBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
): string | undefined {
	let cursor = skipXmlWhitespaceBytes(bytes, start, end)
	if (startsWithElementOpenAtBytes(bytes, BYTES_F_OPEN, cursor, end)) {
		cursor = skipSimpleElementBytes(bytes, cursor, end, 'f')
		if (cursor === -1) return undefined
		cursor = skipXmlWhitespaceBytes(bytes, cursor, end)
	}
	if (!startsWithElementOpenAtBytes(bytes, BYTES_IS_OPEN, cursor, end)) return undefined
	const isTagEnd = findTagEndBytes(bytes, cursor)
	if (isTagEnd === -1 || isTagEnd >= end || isSelfClosingTagBytes(bytes, cursor, isTagEnd)) {
		return undefined
	}
	cursor = skipXmlWhitespaceBytes(bytes, isTagEnd + 1, end)
	if (!startsWithElementOpenAtBytes(bytes, BYTES_T_OPEN, cursor, end)) return undefined
	const textTagEnd = findTagEndBytes(bytes, cursor)
	if (textTagEnd === -1 || textTagEnd >= end || isSelfClosingTagBytes(bytes, cursor, textTagEnd)) {
		return undefined
	}
	const textClose = indexOfBytes(bytes, BYTES_T_CLOSE, textTagEnd + 1, end)
	if (textClose === -1) return undefined
	for (let index = textTagEnd + 1; index < textClose; index++) {
		if (bytes[index] === BYTE_LT) return undefined
	}
	cursor = skipXmlWhitespaceBytes(bytes, textClose + BYTES_T_CLOSE.length, end)
	if (indexOfBytes(bytes, BYTES_IS_CLOSE, cursor, cursor + BYTES_IS_CLOSE.length) !== cursor) {
		return undefined
	}
	cursor = skipXmlWhitespaceBytes(bytes, cursor + BYTES_IS_CLOSE.length, end)
	if (cursor !== end) return undefined
	return decodeXmlBytesText(bytes, textTagEnd + 1, textClose)
}

function skipSimpleElementBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	tagName: string,
): number {
	const tagEnd = findTagEndBytes(bytes, start)
	if (tagEnd === -1 || tagEnd > end) return -1
	if (isSelfClosingTagBytes(bytes, start, tagEnd)) return tagEnd + 1
	const close = bytesLiteral(`</${tagName}>`)
	const closeStart = indexOfBytes(bytes, close, tagEnd + 1, end)
	return closeStart === -1 ? -1 : closeStart + close.length
}

function startsWithElementOpenAtBytes(
	bytes: Uint8Array,
	needle: Uint8Array,
	start: number,
	end: number,
): boolean {
	if (start + needle.length > end) return false
	for (let offset = 0; offset < needle.length; offset++) {
		if (bytes[start + offset] !== needle[offset]) return false
	}
	return isXmlNameDelimiterByte(bytes[start + needle.length])
}

function asciiEquals(bytes: Uint8Array, start: number, end: number, expected: string): boolean {
	if (end - start !== expected.length) return false
	for (let index = 0; index < expected.length; index++) {
		if (bytes[start + index] !== expected.charCodeAt(index)) return false
	}
	return true
}

function asciiSlice(bytes: Uint8Array, start: number, end: number): string {
	let out = ''
	for (let index = start; index < end; index++) out += String.fromCharCode(bytes[index] ?? 0)
	return out
}

function decodeXmlBytesText(bytes: Uint8Array, start: number, end: number): string {
	let hasEntity = false
	for (let index = start; index < end; index++) {
		if (bytes[index] === BYTE_AMP) {
			hasEntity = true
			break
		}
	}
	const text = BYTE_XML_DECODER.decode(bytes.subarray(start, end))
	return hasEntity ? decodeXmlText(text) : text
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
		if (c.is !== undefined && c.is !== null) {
			value = parseInlineString(c, pool)
		} else if (formula || (ctx.valuesOnly && c.f !== undefined && c.f !== null)) {
			value = EMPTY
		} else {
			return false
		}
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
	if (!ctx.valuesOnly && formula && formulaSpec.storedText !== undefined) {
		sheet.storedFormulaText.set(formulaStorageKey(row, col), formulaSpec.storedText)
	}
	return true
}

function parseFormulaText(
	formulaNode: unknown,
	row: number,
	col: number,
	sharedFormulaMasters: SharedFormulaMasterMap,
	pool?: ValueInternPool,
	formulaFeatures?: SheetFormulaFeatures,
): ParsedFormulaText {
	if (formulaNode === undefined || formulaNode === null) return NULL_FORMULA_TEXT
	if (
		typeof formulaNode === 'string' ||
		typeof formulaNode === 'number' ||
		typeof formulaNode === 'boolean'
	) {
		const text = String(formulaNode)
		if (text === '') return NULL_FORMULA_TEXT
		return parseResolvedFormulaText(
			undefined,
			undefined,
			undefined,
			text,
			row,
			col,
			sharedFormulaMasters,
			pool,
			formulaFeatures,
		)
	}
	if (isRawFormulaNode(formulaNode)) {
		const sharedIndex = rawAttr(formulaNode.rawAttrs, 'si')
		const formulaType = rawAttr(formulaNode.rawAttrs, 't')
		if (formulaType === 'dataTable') return parseDataTableFormulaInfoFromRaw(formulaNode.rawAttrs)
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
		if (formulaType === 'dataTable') return parseDataTableFormulaInfoFromNode(node)
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

function parseDataTableFormulaInfoFromRaw(rawAttrs: string): ParsedFormulaText {
	const ref = rawAttr(rawAttrs, 'ref')
	const dt2D = rawBoolAttr(rawAttrs, 'dt2D')
	const dtr = rawBoolAttr(rawAttrs, 'dtr')
	const r1 = rawAttr(rawAttrs, 'r1')
	const r2 = rawAttr(rawAttrs, 'r2')
	const del1 = rawBoolAttr(rawAttrs, 'del1')
	const del2 = rawBoolAttr(rawAttrs, 'del2')
	return {
		text: null,
		info: {
			kind: 'dataTable',
			...(ref !== undefined ? { ref } : {}),
			...(dt2D !== undefined ? { dt2D } : {}),
			...(dtr !== undefined ? { dtr } : {}),
			...(r1 !== undefined ? { r1 } : {}),
			...(r2 !== undefined ? { r2 } : {}),
			...(del1 !== undefined ? { del1 } : {}),
			...(del2 !== undefined ? { del2 } : {}),
		},
	}
}

function parseDataTableFormulaInfoFromNode(node: XmlNode): ParsedFormulaText {
	const ref = attr(node, 'ref')
	const dt2D = boolAttr(node, 'dt2D')
	const dtr = boolAttr(node, 'dtr')
	const r1 = attr(node, 'r1')
	const r2 = attr(node, 'r2')
	const del1 = boolAttr(node, 'del1')
	const del2 = boolAttr(node, 'del2')
	return {
		text: null,
		info: {
			kind: 'dataTable',
			...(ref !== undefined ? { ref } : {}),
			...(dt2D !== undefined ? { dt2D } : {}),
			...(dtr !== undefined ? { dtr } : {}),
			...(r1 !== undefined ? { r1 } : {}),
			...(r2 !== undefined ? { r2 } : {}),
			...(del1 !== undefined ? { del1 } : {}),
			...(del2 !== undefined ? { del2 } : {}),
		},
	}
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
): ParsedFormulaText {
	if (formulaType === 'shared' && sharedIndex) {
		if (formulaFeatures) formulaFeatures.hasSharedFormula = true
		if (text !== undefined && text !== null && text !== '') {
			const storedText = String(text)
			const normalized = normalizeStoredFormulaText(storedText)
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
				storedText,
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
		const storedText = String(text)
		const formula = normalizeStoredFormulaText(storedText)
		return {
			text: pool ? pool.internString(formula) : formula,
			storedText,
			info: { kind: 'array', ...(ref ? { ref } : {}) },
		}
	}
	if (text === undefined || text === null) return NULL_FORMULA_TEXT
	const storedText = String(text)
	const formula = normalizeStoredFormulaText(storedText)
	if (formula === '') return NULL_FORMULA_TEXT
	return { text: pool ? pool.internString(formula) : formula, storedText }
}

function formulaStorageKey(row: number, col: number): string {
	return `${row}:${col}`
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

const WORKSHEET_EXTLST_RE =
	/<(?:(?<prefix>[A-Za-z_][\w.-]*):)?extLst\b(?<attrs>[^>]*)>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?extLst>/g
const CUSTOM_SHEET_VIEWS_RE =
	/<(?:(?<prefix>[A-Za-z_][\w.-]*):)?customSheetViews\b(?<attrs>[^>]*)>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?customSheetViews>/g
const WORKSHEET_CONTROLS_RE =
	/<(?:(?<prefix>[A-Za-z_][\w.-]*):)?controls\b(?<attrs>[^>]*)>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?controls>/g
const SPREADSHEETML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const SPREADSHEET_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

function extractExtLst(xml: string, sheet: Sheet): void {
	let preserved: string | undefined
	for (const match of xml.matchAll(WORKSHEET_EXTLST_RE)) {
		preserved = withSpreadsheetmlNamespace(
			match[0],
			'extLst',
			match.groups?.prefix,
			match.groups?.attrs ?? '',
		)
	}
	if (preserved) sheet.preservedExtLst = preserved
}

function extractCustomSheetViews(xml: string, sheet: Sheet): void {
	let preserved: string | undefined
	for (const match of xml.matchAll(CUSTOM_SHEET_VIEWS_RE)) {
		preserved = withSpreadsheetmlNamespace(
			match[0],
			'customSheetViews',
			match.groups?.prefix,
			match.groups?.attrs ?? '',
		)
	}
	if (preserved) sheet.preservedCustomSheetViews = preserved
}

function extractControls(xml: string, sheet: Sheet): void {
	let preserved: string | undefined
	for (const match of xml.matchAll(WORKSHEET_CONTROLS_RE)) {
		preserved = withSpreadsheetmlNamespace(
			withControlNamespaces(match[0], match.groups?.attrs ?? ''),
			'controls',
			match.groups?.prefix,
			match.groups?.attrs ?? '',
		)
	}
	if (preserved) sheet.preservedControlsXml = preserved
}

function withControlNamespaces(xml: string, attrs: string): string {
	if (!xml.includes('xdr:') || /\sxmlns:xdr=/.test(attrs)) return xml
	return xml.replace(
		/^<(?:(?<prefix>[A-Za-z_][\w.-]*):)?controls\b/,
		(match) => `${match} xmlns:xdr="${SPREADSHEET_DRAWING_NS}"`,
	)
}

function withSpreadsheetmlNamespace(
	xml: string,
	tag: string,
	prefix: string | undefined,
	attrs: string,
): string {
	if (!prefix || new RegExp(`\\sxmlns:${prefix}=`).test(attrs)) return xml
	return xml.replace(
		new RegExp(`^<${prefix}:${tag}\\b`),
		`<${prefix}:${tag} xmlns:${prefix}="${SPREADSHEETML_NS}"`,
	)
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
	const zoomScaleNormal = numAttr(firstView, 'zoomScaleNormal')
	if (zoomScaleNormal !== undefined) viewAttrs.zoomScaleNormal = zoomScaleNormal
	const zoomScaleSheetLayoutView = numAttr(firstView, 'zoomScaleSheetLayoutView')
	if (zoomScaleSheetLayoutView !== undefined)
		viewAttrs.zoomScaleSheetLayoutView = zoomScaleSheetLayoutView
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
	const topLeftCell = attr(firstView, 'topLeftCell')
	if (topLeftCell !== undefined) viewAttrs.topLeftCell = topLeftCell
	if (
		Object.keys(viewAttrs).length > 0 ||
		pane ||
		attr(firstView, 'workbookViewId') !== undefined
	) {
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

function parseSortState(ws: XmlNode, sheet: Sheet): void {
	sheet.sortState = parseSortStateNode(ws.sortState as XmlNode | undefined)
}

function parseAdvancedFilters(ws: XmlNode, sheet: Sheet): void {
	const customSheetViews = childNode(ws, 'customSheetViews')
	if (!customSheetViews) return
	for (const view of childNodes(customSheetViews, 'customSheetView')) {
		const autoFilter = parseAutoFilterNode(childNode(view, 'autoFilter')) as AutoFilter | null
		if (!autoFilter) continue
		const viewName = attr(view, 'name')
		const guid = attr(view, 'guid')
		const parsed: SheetAdvancedFilterInfo = {
			ref: autoFilter.ref,
			autoFilter,
			filterColumnCount: autoFilter.columns.length,
			sortConditionCount: autoFilter.sortState?.conditions.length ?? 0,
		}
		if (viewName) Object.assign(parsed, { viewName })
		if (guid) Object.assign(parsed, { guid })
		sheet.advancedFilters.push(parsed)
	}
}

function parseSparklineGroups(ws: XmlNode, sheet: Sheet): void {
	const groups = findDescendantNodes(ws, 'sparklineGroup')
	for (const [groupIndex, group] of groups.entries()) {
		const sparklines = findDescendantNodes(group, 'sparkline')
		const sparklineRefs = sparklines.map(parseSparklineRef)
		const firstSparkline = sparklineRefs[0]
		const range = firstSparkline?.range
		const locationRange = firstSparkline?.locationRange
		const colorSeries = readSparklineColor(group, 'colorSeries')
		const colorNegative = readSparklineColor(group, 'colorNegative')
		const colorAxis = readSparklineColor(group, 'colorAxis')
		const colorMarkers = readSparklineColor(group, 'colorMarkers')
		const colorFirst = readSparklineColor(group, 'colorFirst')
		const colorLast = readSparklineColor(group, 'colorLast')
		const colorHigh = readSparklineColor(group, 'colorHigh')
		const colorLow = readSparklineColor(group, 'colorLow')
		const type = attr(group, 'type')
		const manualMax = numAttr(group, 'manualMax')
		const manualMin = numAttr(group, 'manualMin')
		const lineWeight = numAttr(group, 'lineWeight')
		const displayEmptyCellsAs = attr(group, 'displayEmptyCellsAs')
		const minAxisType = attr(group, 'minAxisType')
		const maxAxisType = attr(group, 'maxAxisType')
		const uid = attr(group, 'xr2:uid')
		const dateAxisRange = childText(group, 'f')
		const dateAxis = readBoolAttribute(group, 'dateAxis')
		const markers = readBoolAttribute(group, 'markers')
		const highPoint = readBoolAttribute(group, 'high')
		const lowPoint = readBoolAttribute(group, 'low')
		const firstPoint = readBoolAttribute(group, 'first')
		const lastPoint = readBoolAttribute(group, 'last')
		const negative = readBoolAttribute(group, 'negative')
		const displayXAxis = readBoolAttribute(group, 'displayXAxis')
		const displayHidden = readBoolAttribute(group, 'displayHidden')
		const rightToLeft = readBoolAttribute(group, 'rightToLeft')
		const parsed: SheetSparklineGroupInfo = {
			groupIndex,
			count: sparklines.length,
		}
		if (type) Object.assign(parsed, { type })
		if (manualMax !== undefined) Object.assign(parsed, { manualMax })
		if (manualMin !== undefined) Object.assign(parsed, { manualMin })
		if (lineWeight !== undefined) Object.assign(parsed, { lineWeight })
		if (displayEmptyCellsAs) Object.assign(parsed, { displayEmptyCellsAs })
		if (minAxisType) Object.assign(parsed, { minAxisType })
		if (maxAxisType) Object.assign(parsed, { maxAxisType })
		if (uid) Object.assign(parsed, { uid })
		if (dateAxis !== undefined) Object.assign(parsed, { dateAxis })
		if (markers !== undefined) Object.assign(parsed, { markers })
		if (highPoint !== undefined) Object.assign(parsed, { highPoint })
		if (lowPoint !== undefined) Object.assign(parsed, { lowPoint })
		if (firstPoint !== undefined) Object.assign(parsed, { firstPoint })
		if (lastPoint !== undefined) Object.assign(parsed, { lastPoint })
		if (negative !== undefined) Object.assign(parsed, { negative })
		if (displayXAxis !== undefined) Object.assign(parsed, { displayXAxis })
		if (displayHidden !== undefined) Object.assign(parsed, { displayHidden })
		if (rightToLeft !== undefined) Object.assign(parsed, { rightToLeft })
		if (colorSeries) Object.assign(parsed, { colorSeries })
		if (colorNegative) Object.assign(parsed, { colorNegative })
		if (colorAxis) Object.assign(parsed, { colorAxis })
		if (colorMarkers) Object.assign(parsed, { colorMarkers })
		if (colorFirst) Object.assign(parsed, { colorFirst })
		if (colorLast) Object.assign(parsed, { colorLast })
		if (colorHigh) Object.assign(parsed, { colorHigh })
		if (colorLow) Object.assign(parsed, { colorLow })
		if (dateAxisRange) Object.assign(parsed, { dateAxisRange })
		if (range) Object.assign(parsed, { range })
		if (locationRange) Object.assign(parsed, { locationRange })
		if (sparklineRefs.length > 0) Object.assign(parsed, { sparklines: sparklineRefs })
		sheet.sparklineGroups.push(parsed)
	}
}

function childNode(node: XmlNode | undefined, localName: string): XmlNode | undefined {
	if (!node) return undefined
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || localPart(key) !== localName) continue
		if (Array.isArray(value)) return value[0] as XmlNode | undefined
		return value as XmlNode | undefined
	}
	return undefined
}

function childNodes(node: XmlNode | undefined, localName: string): XmlNode[] {
	if (!node) return []
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || localPart(key) !== localName) continue
		return asArray<XmlNode>(value as XmlNode | XmlNode[] | undefined)
	}
	return []
}

function childText(node: XmlNode | undefined, localName: string): string | undefined {
	const child = childNode(node, localName)
	if (!child) return undefined
	const text = child['#text']
	if (text !== undefined && text !== null) return String(text)
	if (typeof child === 'string' || typeof child === 'number' || typeof child === 'boolean') {
		return String(child)
	}
	return undefined
}

function findDescendantNodes(node: XmlNode | undefined, localName: string): XmlNode[] {
	if (!node) return []
	const matches: XmlNode[] = []
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_')) continue
		for (const child of asArray<XmlNode>(value as XmlNode | XmlNode[] | undefined)) {
			if (!child || typeof child !== 'object') continue
			if (localPart(key) === localName) matches.push(child)
			matches.push(...findDescendantNodes(child, localName))
		}
	}
	return matches
}

function parseSparklineRef(node: XmlNode): {
	readonly range?: string
	readonly locationRange?: string
} {
	const range = childText(node, 'f')
	const locationRange = childText(node, 'sqref')
	return {
		...(range ? { range } : {}),
		...(locationRange ? { locationRange } : {}),
	}
}

function readSparklineColor(group: XmlNode, localName: string): string | undefined {
	const color = childNode(group, localName)
	if (!color) return undefined
	return attr(color, 'rgb') ?? attr(color, 'theme') ?? attr(color, 'indexed')
}

function localPart(name: string): string {
	const colon = name.indexOf(':')
	return colon >= 0 ? name.slice(colon + 1) : name
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
				colorScale?: SheetConditionalFormatRule['colorScale']
				dataBar?: SheetConditionalFormatRule['dataBar']
				iconSet?: SheetConditionalFormatRule['iconSet']
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
			const colorScale = parseConditionalColorScale(rule.colorScale as XmlNode | undefined)
			if (colorScale) parsedRule.colorScale = colorScale
			const dataBar = parseConditionalDataBar(rule.dataBar as XmlNode | undefined)
			if (dataBar) parsedRule.dataBar = dataBar
			const iconSet = parseConditionalIconSet(rule.iconSet as XmlNode | undefined)
			if (iconSet) parsedRule.iconSet = iconSet
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

function parseConditionalColorScale(
	node: XmlNode | undefined,
): SheetConditionalFormatRule['colorScale'] | undefined {
	if (!node) return undefined
	return {
		cfvo: parseConditionalCfvo(node),
		colors: asArray<XmlNode>(node.color as XmlNode | XmlNode[] | undefined).map(parseCfColor),
	}
}

function parseConditionalDataBar(
	node: XmlNode | undefined,
): SheetConditionalFormatRule['dataBar'] | undefined {
	if (!node) return undefined
	const parsed: {
		cfvo: NonNullable<SheetConditionalFormatRule['dataBar']>['cfvo']
		color?: NonNullable<NonNullable<SheetConditionalFormatRule['dataBar']>['color']>
		minLength?: number
		maxLength?: number
		showValue?: boolean
	} = {
		cfvo: parseConditionalCfvo(node),
	}
	const colorNode = node.color as XmlNode | undefined
	if (colorNode) parsed.color = parseCfColor(colorNode)
	const minLength = numAttr(node, 'minLength')
	if (minLength !== undefined) parsed.minLength = minLength
	const maxLength = numAttr(node, 'maxLength')
	if (maxLength !== undefined) parsed.maxLength = maxLength
	const showValue = readBoolAttribute(node, 'showValue')
	if (showValue !== undefined) parsed.showValue = showValue
	return parsed
}

function parseConditionalIconSet(
	node: XmlNode | undefined,
): SheetConditionalFormatRule['iconSet'] | undefined {
	if (!node) return undefined
	const parsed: {
		cfvo: NonNullable<SheetConditionalFormatRule['iconSet']>['cfvo']
		iconSet?: string
		showValue?: boolean
		percent?: boolean
		reverse?: boolean
	} = {
		cfvo: parseConditionalCfvo(node),
	}
	const iconSet = attr(node, 'iconSet')
	if (iconSet) parsed.iconSet = iconSet
	const showValue = readBoolAttribute(node, 'showValue')
	if (showValue !== undefined) parsed.showValue = showValue
	const percent = readBoolAttribute(node, 'percent')
	if (percent !== undefined) parsed.percent = percent
	const reverse = readBoolAttribute(node, 'reverse')
	if (reverse !== undefined) parsed.reverse = reverse
	return parsed
}

function parseConditionalCfvo(
	node: XmlNode,
): NonNullable<SheetConditionalFormatRule['colorScale']>['cfvo'] {
	return asArray<XmlNode>(node.cfvo as XmlNode | XmlNode[] | undefined).map((entry) => {
		const parsed: { type?: string; value?: string; gte?: boolean } = {}
		const type = attr(entry, 'type')
		if (type) parsed.type = type
		const value = attr(entry, 'val')
		if (value) parsed.value = value
		const gte = readBoolAttribute(entry, 'gte')
		if (gte !== undefined) parsed.gte = gte
		return parsed
	})
}

function parseCfColor(
	node: XmlNode,
): NonNullable<SheetConditionalFormatRule['colorScale']>['colors'][number] {
	const parsed: { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean } =
		{}
	const rgb = attr(node, 'rgb')
	if (rgb) parsed.rgb = rgb
	const theme = numAttr(node, 'theme')
	if (theme !== undefined) parsed.theme = theme
	const tint = numAttr(node, 'tint')
	if (tint !== undefined) parsed.tint = tint
	const indexed = numAttr(node, 'indexed')
	if (indexed !== undefined) parsed.indexed = indexed
	const auto = readBoolAttribute(node, 'auto')
	if (auto !== undefined) parsed.auto = auto
	return parsed
}

function parseDataValidations(ws: XmlNode, sheet: Sheet, pool?: ValueInternPool): void {
	const container = ws.dataValidations as XmlNode | undefined
	if (!container) return
	for (const validation of asArray<XmlNode>(container.dataValidation as XmlNode | XmlNode[])) {
		const sqref = attr(validation, 'sqref')
		if (!sqref) continue
		const parsed = parseDataValidationAttributes(validation, sqref)
		const formula1 = readNodeText(validation.formula1)
		if (formula1) parsed.formula1 = formula1
		const formula2 = readNodeText(validation.formula2)
		if (formula2) parsed.formula2 = formula2
		pushDataValidation(sheet, parsed, pool)
	}
}

function parseX14ConditionalFormats(xml: string, sheet: Sheet, pool?: ValueInternPool): void {
	let index = 0
	for (const match of xml.matchAll(X14_CONDITIONAL_FORMATTING_RE)) {
		const body = match[3] ?? ''
		const sqref = readFirstXmlElementText(body, 'sqref')
		if (!sqref) {
			index += 1
			continue
		}
		const formulas = readXmlElementTexts(body, 'f')
		const rule = readFirstXmlElement(body, 'cfRule')
		const type = rule ? attr(rule.attrs, 'type') : undefined
		const priority = rule ? numAttr(rule.attrs, 'priority') : undefined
		const id = rule ? attr(rule.attrs, 'id') : undefined
		const dataBar = rule ? parseX14DataBar(rule.body) : undefined
		const iconSet = rule ? parseX14IconSet(rule.body) : undefined
		sheet.x14ConditionalFormats.push({
			index,
			sqref: pool ? pool.internString(sqref) : sqref,
			formulas: pool ? formulas.map((formula) => pool.internString(formula)) : formulas,
			...(type ? { type: pool ? pool.internString(type) : type } : {}),
			...(priority !== undefined ? { priority } : {}),
			...(id ? { id: pool ? pool.internString(id) : id } : {}),
			...(dataBar ? { dataBar: internX14DataBar(dataBar, pool) } : {}),
			...(iconSet ? { iconSet: internX14IconSet(iconSet, pool) } : {}),
		})
		index += 1
	}
}

function parseX14DataBar(xml: string): SheetX14ConditionalFormatDataBarInfo | undefined {
	const dataBar = readFirstXmlElement(xml, 'dataBar')
	if (!dataBar) return undefined
	const parsed: SheetX14ConditionalFormatDataBarInfo = {
		cfvo: readX14Cfvos(dataBar.body),
	}
	const minLength = numAttr(dataBar.attrs, 'minLength')
	if (minLength !== undefined) Object.assign(parsed, { minLength })
	const maxLength = numAttr(dataBar.attrs, 'maxLength')
	if (maxLength !== undefined) Object.assign(parsed, { maxLength })
	const border = readBoolAttribute(dataBar.attrs, 'border')
	if (border !== undefined) Object.assign(parsed, { border })
	const negativeBarBorderColorSameAsPositive = readBoolAttribute(
		dataBar.attrs,
		'negativeBarBorderColorSameAsPositive',
	)
	if (negativeBarBorderColorSameAsPositive !== undefined) {
		Object.assign(parsed, { negativeBarBorderColorSameAsPositive })
	}
	const borderColor = readX14Color(dataBar.body, 'borderColor')
	if (borderColor) Object.assign(parsed, { borderColor })
	const negativeFillColor = readX14Color(dataBar.body, 'negativeFillColor')
	if (negativeFillColor) Object.assign(parsed, { negativeFillColor })
	const negativeBorderColor = readX14Color(dataBar.body, 'negativeBorderColor')
	if (negativeBorderColor) Object.assign(parsed, { negativeBorderColor })
	const axisColor = readX14Color(dataBar.body, 'axisColor')
	if (axisColor) Object.assign(parsed, { axisColor })
	return parsed
}

function parseX14IconSet(xml: string): SheetX14ConditionalFormatIconSetInfo | undefined {
	const iconSet = readFirstXmlElement(xml, 'iconSet')
	if (!iconSet) return undefined
	const parsed: SheetX14ConditionalFormatIconSetInfo = {
		cfvo: readX14Cfvos(iconSet.body),
	}
	const iconSetName = attr(iconSet.attrs, 'iconSet')
	if (iconSetName) Object.assign(parsed, { iconSet: iconSetName })
	const custom = readBoolAttribute(iconSet.attrs, 'custom')
	if (custom !== undefined) Object.assign(parsed, { custom })
	const showValue = readBoolAttribute(iconSet.attrs, 'showValue')
	if (showValue !== undefined) Object.assign(parsed, { showValue })
	const percent = readBoolAttribute(iconSet.attrs, 'percent')
	if (percent !== undefined) Object.assign(parsed, { percent })
	const reverse = readBoolAttribute(iconSet.attrs, 'reverse')
	if (reverse !== undefined) Object.assign(parsed, { reverse })
	const icons = readX14CfIcons(iconSet.body)
	if (icons.length > 0) Object.assign(parsed, { icons })
	return parsed
}

function readX14Cfvos(xml: string): SheetConditionalFormatValueObject[] {
	const cfvos: SheetConditionalFormatValueObject[] = []
	for (const entry of readXmlElements(xml, 'cfvo')) {
		const parsed: SheetConditionalFormatValueObject = {}
		const type = attr(entry.attrs, 'type')
		if (type) Object.assign(parsed, { type })
		const value = attr(entry.attrs, 'val') ?? readFirstXmlElementText(entry.body, 'f')
		if (value !== undefined) Object.assign(parsed, { value })
		const gte = readBoolAttribute(entry.attrs, 'gte')
		if (gte !== undefined) Object.assign(parsed, { gte })
		cfvos.push(parsed)
	}
	return cfvos
}

function readX14CfIcons(xml: string): SheetX14ConditionalFormatIconInfo[] {
	const icons: SheetX14ConditionalFormatIconInfo[] = []
	for (const entry of readXmlElements(xml, 'cfIcon')) {
		const parsed: SheetX14ConditionalFormatIconInfo = {}
		const iconSet = attr(entry.attrs, 'iconSet')
		if (iconSet) Object.assign(parsed, { iconSet })
		const iconId = numAttr(entry.attrs, 'iconId')
		if (iconId !== undefined) Object.assign(parsed, { iconId })
		if (Object.keys(parsed).length > 0) icons.push(parsed)
	}
	return icons
}

function readX14Color(xml: string, localName: string): SheetConditionalFormatColor | undefined {
	const color = readFirstXmlElement(xml, localName)
	if (!color) return undefined
	const parsed: SheetConditionalFormatColor = {}
	const rgb = attr(color.attrs, 'rgb')
	if (rgb) Object.assign(parsed, { rgb })
	const theme = numAttr(color.attrs, 'theme')
	if (theme !== undefined) Object.assign(parsed, { theme })
	const tint = numAttr(color.attrs, 'tint')
	if (tint !== undefined) Object.assign(parsed, { tint })
	const indexed = numAttr(color.attrs, 'indexed')
	if (indexed !== undefined) Object.assign(parsed, { indexed })
	const auto = readBoolAttribute(color.attrs, 'auto')
	if (auto !== undefined) Object.assign(parsed, { auto })
	return Object.keys(parsed).length > 0 ? parsed : undefined
}

function internX14DataBar(
	dataBar: SheetX14ConditionalFormatDataBarInfo,
	pool: ValueInternPool | undefined,
): SheetX14ConditionalFormatDataBarInfo {
	if (!pool) return dataBar
	return {
		...dataBar,
		cfvo: dataBar.cfvo.map((entry) => internConditionalFormatValueObject(entry, pool)),
	}
}

function internX14IconSet(
	iconSet: SheetX14ConditionalFormatIconSetInfo,
	pool: ValueInternPool | undefined,
): SheetX14ConditionalFormatIconSetInfo {
	if (!pool) return iconSet
	return {
		...iconSet,
		...(iconSet.iconSet ? { iconSet: pool.internString(iconSet.iconSet) } : {}),
		cfvo: iconSet.cfvo.map((entry) => internConditionalFormatValueObject(entry, pool)),
		...(iconSet.icons
			? {
					icons: iconSet.icons.map((icon) => ({
						...icon,
						...(icon.iconSet ? { iconSet: pool.internString(icon.iconSet) } : {}),
					})),
				}
			: {}),
	}
}

function internConditionalFormatValueObject(
	cfvo: SheetConditionalFormatValueObject,
	pool: ValueInternPool,
): SheetConditionalFormatValueObject {
	return {
		...cfvo,
		...(cfvo.type ? { type: pool.internString(cfvo.type) } : {}),
		...(cfvo.value ? { value: pool.internString(cfvo.value) } : {}),
	}
}

type MutableDataValidation = {
	sqref: string
	source?: 'x14'
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
	imeMode?: string
	formula1?: string
	formula2?: string
}

function parseX14DataValidations(xml: string, sheet: Sheet, pool?: ValueInternPool): void {
	let index = 0
	for (const match of xml.matchAll(X14_DATA_VALIDATION_RE)) {
		const rawAttrs = match[2] ?? ''
		const body = match[3] ?? ''
		const attrs = parseRawAttributes(rawAttrs)
		const sqref = attr(attrs, 'sqref') ?? readFirstXmlElementText(body, 'sqref')
		if (!sqref) {
			index += 1
			continue
		}
		const parsed = parseDataValidationAttributes(attrs, sqref)
		const formula1 = readX14DataValidationFormula(body, 'formula1')
		if (formula1) parsed.formula1 = formula1
		const formula2 = readX14DataValidationFormula(body, 'formula2')
		if (formula2) parsed.formula2 = formula2
		pushX14DataValidationInfo(sheet, index, parsed, pool)
		parsed.source = 'x14'
		pushDataValidation(sheet, parsed, pool)
		index += 1
	}
	for (const match of xml.matchAll(X14_SELF_CLOSING_DATA_VALIDATION_RE)) {
		const attrs = parseRawAttributes(match[2] ?? '')
		const sqref = attr(attrs, 'sqref')
		if (!sqref) {
			index += 1
			continue
		}
		const parsed = parseDataValidationAttributes(attrs, sqref)
		pushX14DataValidationInfo(sheet, index, parsed, pool)
		parsed.source = 'x14'
		pushDataValidation(sheet, parsed, pool)
		index += 1
	}
}

function pushX14DataValidationInfo(
	sheet: Sheet,
	index: number,
	parsed: MutableDataValidation,
	pool?: ValueInternPool,
): void {
	sheet.x14DataValidations.push({
		index,
		sqref: pool ? pool.internString(parsed.sqref) : parsed.sqref,
		...(parsed.type ? { type: pool ? pool.internString(parsed.type) : parsed.type } : {}),
		...(parsed.operator
			? { operator: pool ? pool.internString(parsed.operator) : parsed.operator }
			: {}),
		...(parsed.allowBlank !== undefined ? { allowBlank: parsed.allowBlank } : {}),
		...(parsed.showInputMessage !== undefined ? { showInputMessage: parsed.showInputMessage } : {}),
		...(parsed.showErrorMessage !== undefined ? { showErrorMessage: parsed.showErrorMessage } : {}),
		...(parsed.showDropDown !== undefined ? { showDropDown: parsed.showDropDown } : {}),
		...(parsed.promptTitle
			? { promptTitle: pool ? pool.internString(parsed.promptTitle) : parsed.promptTitle }
			: {}),
		...(parsed.prompt ? { prompt: pool ? pool.internString(parsed.prompt) : parsed.prompt } : {}),
		...(parsed.errorTitle
			? { errorTitle: pool ? pool.internString(parsed.errorTitle) : parsed.errorTitle }
			: {}),
		...(parsed.error ? { error: pool ? pool.internString(parsed.error) : parsed.error } : {}),
		...(parsed.errorStyle
			? { errorStyle: pool ? pool.internString(parsed.errorStyle) : parsed.errorStyle }
			: {}),
		...(parsed.imeMode
			? { imeMode: pool ? pool.internString(parsed.imeMode) : parsed.imeMode }
			: {}),
		...(parsed.formula1
			? { formula1: pool ? pool.internString(parsed.formula1) : parsed.formula1 }
			: {}),
		...(parsed.formula2
			? { formula2: pool ? pool.internString(parsed.formula2) : parsed.formula2 }
			: {}),
	})
}

function parseDataValidationAttributes(node: XmlNode, sqref: string): MutableDataValidation {
	const parsed: MutableDataValidation = { sqref }
	const type = attr(node, 'type')
	if (type) parsed.type = type
	const operator = attr(node, 'operator')
	if (operator) parsed.operator = operator
	const errorStyle = attr(node, 'errorStyle')
	if (errorStyle) parsed.errorStyle = errorStyle
	const imeMode = attr(node, 'imeMode')
	if (imeMode) parsed.imeMode = imeMode
	const allowBlank = readBoolAttribute(node, 'allowBlank')
	if (allowBlank !== undefined) parsed.allowBlank = allowBlank
	const showInputMessage = readBoolAttribute(node, 'showInputMessage')
	if (showInputMessage !== undefined) parsed.showInputMessage = showInputMessage
	const showErrorMessage = readBoolAttribute(node, 'showErrorMessage')
	if (showErrorMessage !== undefined) parsed.showErrorMessage = showErrorMessage
	const showDropDown = readBoolAttribute(node, 'showDropDown')
	if (showDropDown !== undefined) parsed.showDropDown = showDropDown
	const promptTitle = attr(node, 'promptTitle')
	if (promptTitle) parsed.promptTitle = promptTitle
	const prompt = attr(node, 'prompt')
	if (prompt) parsed.prompt = prompt
	const errorTitle = attr(node, 'errorTitle')
	if (errorTitle) parsed.errorTitle = errorTitle
	const error = attr(node, 'error')
	if (error) parsed.error = error
	return parsed
}

function pushDataValidation(
	sheet: Sheet,
	parsed: MutableDataValidation,
	pool?: ValueInternPool,
): void {
	if (sheet.dataValidations.some((existing) => isSameDataValidation(existing, parsed))) return
	if (pool) {
		parsed.sqref = pool.internString(parsed.sqref)
		if (parsed.type) parsed.type = pool.internString(parsed.type)
		if (parsed.operator) parsed.operator = pool.internString(parsed.operator)
		if (parsed.promptTitle) parsed.promptTitle = pool.internString(parsed.promptTitle)
		if (parsed.prompt) parsed.prompt = pool.internString(parsed.prompt)
		if (parsed.errorTitle) parsed.errorTitle = pool.internString(parsed.errorTitle)
		if (parsed.error) parsed.error = pool.internString(parsed.error)
		if (parsed.errorStyle) parsed.errorStyle = pool.internString(parsed.errorStyle)
		if (parsed.imeMode) parsed.imeMode = pool.internString(parsed.imeMode)
		if (parsed.formula1) parsed.formula1 = pool.internString(parsed.formula1)
		if (parsed.formula2) parsed.formula2 = pool.internString(parsed.formula2)
	}
	sheet.dataValidations.push(parsed as SheetDataValidation)
}

function isSameDataValidation(left: SheetDataValidation, right: MutableDataValidation): boolean {
	return (
		left.sqref === right.sqref &&
		left.type === right.type &&
		left.operator === right.operator &&
		left.allowBlank === right.allowBlank &&
		left.showInputMessage === right.showInputMessage &&
		left.showErrorMessage === right.showErrorMessage &&
		left.showDropDown === right.showDropDown &&
		left.promptTitle === right.promptTitle &&
		left.prompt === right.prompt &&
		left.errorTitle === right.errorTitle &&
		left.error === right.error &&
		left.errorStyle === right.errorStyle &&
		left.imeMode === right.imeMode &&
		left.formula1 === right.formula1 &&
		left.formula2 === right.formula2
	)
}

function readX14DataValidationFormula(
	xml: string,
	tagName: 'formula1' | 'formula2',
): string | undefined {
	const formulaBody = readFirstXmlElementBody(xml, tagName)
	if (formulaBody === undefined) return undefined
	return readFirstXmlElementText(formulaBody, 'f') ?? readXmlTextContent(formulaBody)
}

function readFirstXmlElementText(xml: string, localName: string): string | undefined {
	const body = readFirstXmlElementBody(xml, localName)
	if (body === undefined) return undefined
	return readXmlTextContent(body)
}

function readXmlElementTexts(xml: string, localName: string): string[] {
	const escapedName = escapeRegExp(localName)
	const pattern = new RegExp(
		`<(?:[A-Za-z_][\\w.-]*:)?${escapedName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${escapedName}>`,
		'gi',
	)
	const values: string[] = []
	for (const match of xml.matchAll(pattern)) {
		const text = readXmlTextContent(match[1] ?? '')
		if (text !== undefined) values.push(text)
	}
	return values
}

function readXmlElements(
	xml: string,
	localName: string,
): { readonly attrs: XmlNode; readonly body: string }[] {
	const escapedName = escapeRegExp(localName)
	const pattern = new RegExp(
		`<(?:[A-Za-z_][\\w.-]*:)?${escapedName}\\b([^>]*)>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${escapedName}>|<(?:[A-Za-z_][\\w.-]*:)?${escapedName}\\b([^>]*)/>`,
		'gi',
	)
	const elements: { attrs: XmlNode; body: string }[] = []
	for (const match of xml.matchAll(pattern)) {
		elements.push({
			attrs: parseRawAttributes(match[1] ?? match[3] ?? ''),
			body: match[2] ?? '',
		})
	}
	return elements
}

function readFirstXmlElement(
	xml: string,
	localName: string,
): { readonly attrs: XmlNode; readonly body: string } | undefined {
	return readXmlElements(xml, localName)[0]
}

function readFirstXmlElementBody(xml: string, localName: string): string | undefined {
	return readFirstXmlElement(xml, localName)?.body
}

function readXmlTextContent(xml: string): string | undefined {
	const text = decodeXmlText(xml.replace(/<[^>]*>/g, '').trim())
	return text.length > 0 ? text : undefined
}

function escapeRegExp(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
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
