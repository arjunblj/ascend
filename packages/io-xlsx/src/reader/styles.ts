import type {
	AlignmentStyle,
	BorderEdge,
	BorderLineStyle,
	BorderStyle,
	CellStyle,
	Color,
	FillPattern,
	FillStyle,
	FontStyle,
	HorizontalAlign,
	NamedStyleInfo,
	VerticalAlign,
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

type WritablePartial<T> = { -readonly [K in keyof T]?: T[K] }

const BUILTIN_DATE_FMT_IDS = new Set([
	14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
	52, 53, 54, 55, 56, 57, 58,
])

const BUILTIN_NUM_FMTS = new Map<number, string>([
	[0, 'General'],
	[1, '0'],
	[2, '0.00'],
	[3, '#,##0'],
	[4, '#,##0.00'],
	[9, '0%'],
	[10, '0.00%'],
	[11, '0.00E+00'],
	[12, '# ?/?'],
	[13, '# ??/??'],
	// Builtin ID 14 is locale-dependent in Excel; keep a neutral canonical form.
	[14, 'm/d/yy'],
	[15, 'd-mmm-yy'],
	[16, 'd-mmm'],
	[17, 'mmm-yy'],
	[18, 'h:mm AM/PM'],
	[19, 'h:mm:ss AM/PM'],
	[20, 'h:mm'],
	[21, 'h:mm:ss'],
	[22, 'm/d/yy h:mm'],
	[37, '#,##0 ;(#,##0)'],
	[38, '#,##0 ;[Red](#,##0)'],
	[39, '#,##0.00;(#,##0.00)'],
	[40, '#,##0.00;[Red](#,##0.00)'],
	[45, 'mm:ss'],
	[46, '[h]:mm:ss'],
	[47, 'mmss.0'],
	[48, '##0.0E+0'],
	[49, '@'],
])

export interface ParsedStyles {
	readonly cellStyles: CellStyle[]
	readonly differentialStyles: readonly CellStyle[]
	readonly isDateFormat: boolean[]
	readonly metadata: {
		readonly numFmtCount: number
		readonly fontCount: number
		readonly fillCount: number
		readonly borderCount: number
		readonly cellXfCount: number
		readonly dxfCount: number
		readonly tableStyleCount: number
	}
}

export interface ParsedStylesLite {
	readonly isDateFormat: boolean[]
	readonly metadata: ParsedStyles['metadata']
}

export function parseStyles(xml: string): ParsedStyles {
	const fast = parseStylesFast(xml)
	if (fast) return fast

	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
	const ss = doc.styleSheet as XmlNode | undefined
	if (!ss) {
		return {
			cellStyles: [{}],
			differentialStyles: [],
			isDateFormat: [false],
			metadata: {
				numFmtCount: 0,
				fontCount: 0,
				fillCount: 0,
				borderCount: 0,
				cellXfCount: 1,
				dxfCount: 0,
				tableStyleCount: 0,
			},
		}
	}

	const numFmts = parseNumFmts(ss)
	const fonts = parseFonts(ss)
	const fills = parseFills(ss)
	const borders = parseBorders(ss)
	const cellXfCount = countNodes(ss.cellXfs, 'xf')
	const dxfCount = countNodes(ss.dxfs, 'dxf')
	const tableStyleCount = countNodes(ss.tableStyles, 'tableStyle')
	const differentialStyles = parseDxfs(ss, fonts, fills, borders, numFmts)

	const namedStyles = parseCellStyleNames(ss)
	const built = buildCellStyles(ss, fonts, fills, borders, numFmts)
	return {
		...built,
		differentialStyles,
		metadata: {
			numFmtCount: Math.max(0, numFmts.size - BUILTIN_NUM_FMTS.size),
			fontCount: Math.max(0, fonts.length),
			fillCount: Math.max(0, fills.length),
			borderCount: Math.max(0, borders.length),
			cellXfCount: Math.max(0, cellXfCount),
			dxfCount: Math.max(0, dxfCount),
			tableStyleCount: Math.max(0, tableStyleCount),
			...(namedStyles.length > 0 ? { namedStyles } : {}),
		},
	}
}

function parseStylesFast(xml: string): ParsedStyles | null {
	if (!xml.includes('<styleSheet')) return null
	if (
		/<(?:dxfs|cellStyles|tableStyles|extLst)\b/i.test(xml) ||
		/<(?:alignment|protection|gradientFill)\b/i.test(xml)
	) {
		return null
	}
	const numFmts = parseNumFmtsLite(xml)
	const fonts = parseFontsFast(extractXmlSection(xml, 'fonts'))
	const fills = parseFillsFast(extractXmlSection(xml, 'fills'))
	const borders = parseBordersFast(extractXmlSection(xml, 'borders'))
	if (!fonts || !fills || !borders) return null
	const built = buildCellStylesFast(
		extractXmlSection(xml, 'cellXfs'),
		fonts,
		fills,
		borders,
		numFmts,
	)
	if (!built) return null
	return {
		...built,
		differentialStyles: [],
		metadata: {
			numFmtCount: Math.max(0, numFmts.size - BUILTIN_NUM_FMTS.size),
			fontCount: fonts.length,
			fillCount: fills.length,
			borderCount: borders.length,
			cellXfCount: built.cellStyles.length,
			dxfCount: 0,
			tableStyleCount: 0,
		},
	}
}

function parseFontsFast(section: string | null): FontStyle[] | null {
	if (!section) return [{}]
	const fonts: FontStyle[] = []
	for (const fontXml of matchXmlChildren(section, 'font')) {
		if (hasUnsupportedFastStyleXml(fontXml)) return null
		const props: WritablePartial<FontStyle> = {}
		const name = rawXmlAttr(firstChildAttrs(fontXml, 'name') ?? '', 'val')
		if (name) props.name = decodeXmlAttribute(name)
		const size = rawXmlNumberAttr(firstChildAttrs(fontXml, 'sz') ?? '', 'val')
		if (size !== undefined) props.size = size
		const bold = readFastBooleanElement(fontXml, 'b')
		if (bold !== undefined) props.bold = bold
		const italic = readFastBooleanElement(fontXml, 'i')
		if (italic !== undefined) props.italic = italic
		const strike = readFastBooleanElement(fontXml, 'strike')
		if (strike !== undefined) props.strikethrough = strike
		const underline = parseFastUnderline(fontXml)
		if (underline !== undefined) props.underline = underline
		const colorAttrs = firstChildAttrs(fontXml, 'color')
		const color = colorAttrs ? parseColorAttrsFast(colorAttrs) : undefined
		if (color) props.color = color
		fonts.push(props as FontStyle)
	}
	return fonts.length > 0 ? fonts : [{}]
}

function parseFillsFast(section: string | null): FillStyle[] | null {
	if (!section) return [{}]
	const fills: FillStyle[] = []
	for (const fillXml of matchXmlChildren(section, 'fill')) {
		if (fillXml.includes('<gradientFill')) return null
		const patternFillXml = firstChildXml(fillXml, 'patternFill')
		if (!patternFillXml) {
			fills.push({})
			continue
		}
		const attrs = tagAttrs(patternFillXml) ?? ''
		const props: WritablePartial<FillStyle> = {}
		const pattern = rawXmlAttr(attrs, 'patternType')
		if (pattern) props.pattern = pattern as FillPattern
		const fgAttrs = firstChildAttrs(patternFillXml, 'fgColor')
		const fg = fgAttrs ? parseColorAttrsFast(fgAttrs) : undefined
		if (fg) props.fgColor = fg
		const bgAttrs = firstChildAttrs(patternFillXml, 'bgColor')
		const bg = bgAttrs ? parseColorAttrsFast(bgAttrs) : undefined
		if (bg) props.bgColor = bg
		fills.push(props as FillStyle)
	}
	return fills.length > 0 ? fills : [{}]
}

function parseBordersFast(section: string | null): BorderStyle[] | null {
	if (!section) return [{}]
	const borders: BorderStyle[] = []
	for (const borderXml of matchXmlChildren(section, 'border')) {
		if (hasUnsupportedFastStyleXml(borderXml)) return null
		const attrs = tagAttrs(borderXml) ?? ''
		const props: WritablePartial<BorderStyle> = {}
		for (const edgeName of ['top', 'bottom', 'left', 'right', 'diagonal'] as const) {
			const edge = parseBorderEdgeFast(firstChildXml(borderXml, edgeName))
			if (edge) props[edgeName] = edge
		}
		const diagUp = rawXmlBoolAttr(attrs, 'diagonalUp')
		if (diagUp !== undefined) props.diagonalUp = diagUp
		const diagDown = rawXmlBoolAttr(attrs, 'diagonalDown')
		if (diagDown !== undefined) props.diagonalDown = diagDown
		borders.push(props as BorderStyle)
	}
	return borders.length > 0 ? borders : [{}]
}

function buildCellStylesFast(
	cellXfsXml: string | null,
	fonts: readonly FontStyle[],
	fills: readonly FillStyle[],
	borders: readonly BorderStyle[],
	numFmts: Map<number, string>,
): Pick<ParsedStyles, 'cellStyles' | 'isDateFormat'> | null {
	if (!cellXfsXml) return { cellStyles: [{}], isDateFormat: [false] }
	const cellStyles: CellStyle[] = []
	const isDateFormat: boolean[] = []
	for (const xfXml of matchXmlChildren(cellXfsXml, 'xf')) {
		if (!isSelfClosingXml(xfXml)) return null
		const attrs = tagAttrs(xfXml) ?? ''
		const fontId = rawXmlNumberAttr(attrs, 'fontId') ?? 0
		const fillId = rawXmlNumberAttr(attrs, 'fillId') ?? 0
		const borderId = rawXmlNumberAttr(attrs, 'borderId') ?? 0
		const numFmtId = rawXmlNumberAttr(attrs, 'numFmtId') ?? 0
		const font = fonts[fontId]
		const fill = fills[fillId]
		const border = borders[borderId]
		const formatCode = numFmts.get(numFmtId)
		if (isDefaultCellXf(fontId, fillId, borderId, numFmtId, formatCode)) {
			cellStyles.push({})
			isDateFormat.push(false)
			continue
		}
		const style: WritablePartial<CellStyle> = {}
		if (font && hasProps(font)) style.font = font
		if (fill && hasProps(fill)) style.fill = fill
		if (border && hasProps(border)) style.border = border
		if (formatCode && formatCode !== 'General') style.numberFormat = formatCode
		cellStyles.push(style as CellStyle)
		isDateFormat.push(checkDateFormat(numFmtId, formatCode))
	}
	if (cellStyles.length === 0) {
		cellStyles.push({})
		isDateFormat.push(false)
	}
	return { cellStyles, isDateFormat }
}

export function parseStylesLite(xml: string): ParsedStylesLite {
	const normalized = normalizeMainSpreadsheetNamespacePrefix(xml)
	const numFmts = parseNumFmtsLite(normalized)
	const cellXfsXml = extractXmlSection(normalized, 'cellXfs')
	const cellXfCount = countXmlChildren(cellXfsXml, 'xf')
	const isDateFormat = buildDateFormatFlagsLite(cellXfsXml, numFmts)
	return {
		isDateFormat,
		metadata: {
			numFmtCount: Math.max(0, numFmts.size - BUILTIN_NUM_FMTS.size),
			fontCount: countXmlChildren(extractXmlSection(normalized, 'fonts'), 'font'),
			fillCount: countXmlChildren(extractXmlSection(normalized, 'fills'), 'fill'),
			borderCount: countXmlChildren(extractXmlSection(normalized, 'borders'), 'border'),
			cellXfCount: Math.max(0, cellXfCount),
			dxfCount: countXmlChildren(extractXmlSection(normalized, 'dxfs'), 'dxf'),
			tableStyleCount: countXmlChildren(extractXmlSection(normalized, 'tableStyles'), 'tableStyle'),
		},
	}
}

function parseNumFmtsLite(xml: string): Map<number, string> {
	const fmts = new Map(BUILTIN_NUM_FMTS)
	const section = extractXmlSection(xml, 'numFmts')
	if (!section) return fmts
	for (const match of section.matchAll(/<numFmt\b([^>]*)\/?>/g)) {
		const attrs = match[1] ?? ''
		const id = rawXmlNumberAttr(attrs, 'numFmtId')
		const code = rawXmlAttr(attrs, 'formatCode')
		if (id !== undefined && code) fmts.set(id, decodeXmlAttribute(code))
	}
	return fmts
}

function buildDateFormatFlagsLite(
	cellXfsXml: string | null,
	numFmts: Map<number, string>,
): boolean[] {
	if (!cellXfsXml) return [false]
	const isDateFormat: boolean[] = []
	for (const match of cellXfsXml.matchAll(/<xf\b([^>]*)\/?>/g)) {
		const numFmtId = rawXmlNumberAttr(match[1] ?? '', 'numFmtId') ?? 0
		isDateFormat.push(checkDateFormat(numFmtId, numFmts.get(numFmtId)))
	}
	return isDateFormat.length > 0 ? isDateFormat : [false]
}

function extractXmlSection(xml: string, tag: string): string | null {
	const open = new RegExp(`<${tag}\\b[^>]*>`, 'i').exec(xml)
	if (!open || open.index === undefined) return null
	const contentStart = open.index + open[0].length
	const close = new RegExp(`</${tag}>`, 'i').exec(xml.slice(contentStart))
	if (!close || close.index === undefined) return null
	return xml.slice(contentStart, contentStart + close.index)
}

function countXmlChildren(section: string | null, tag: string): number {
	if (!section) return 0
	let count = 0
	for (const _match of section.matchAll(new RegExp(`<${tag}\\b`, 'g'))) count++
	return count
}

function matchXmlChildren(section: string, tag: string): string[] {
	const children: string[] = []
	let index = 0
	const closeTag = `</${tag}>`
	while (index < section.length) {
		const start = findTagStartFrom(section, tag, index)
		if (start < 0) break
		const openEnd = section.indexOf('>', start)
		if (openEnd < 0) break
		if (section.charCodeAt(openEnd - 1) === 47) {
			children.push(section.slice(start, openEnd + 1))
			index = openEnd + 1
			continue
		}
		const closeStart = section.indexOf(closeTag, openEnd + 1)
		if (closeStart < 0) break
		const end = closeStart + closeTag.length
		children.push(section.slice(start, end))
		index = end
	}
	return children
}

function firstChildXml(xml: string, tag: string): string | undefined {
	const start = findTagStart(xml, tag)
	if (start < 0) return undefined
	const openEnd = xml.indexOf('>', start)
	if (openEnd < 0) return undefined
	if (xml.charCodeAt(openEnd - 1) === 47) return xml.slice(start, openEnd + 1)
	const closeTag = `</${tag}>`
	const closeStart = xml.indexOf(closeTag, openEnd + 1)
	if (closeStart < 0) return undefined
	return xml.slice(start, closeStart + closeTag.length)
}

function firstChildAttrs(xml: string, tag: string): string | undefined {
	const start = findTagStart(xml, tag)
	if (start < 0) return undefined
	return attrsFromOpenTag(xml, start, tag.length)
}

function tagAttrs(xml: string): string | undefined {
	if (!xml.startsWith('<') || xml.startsWith('</')) return undefined
	let nameEnd = 1
	while (nameEnd < xml.length && !isXmlNameTerminator(xml.charCodeAt(nameEnd))) nameEnd++
	return attrsFromOpenTag(xml, 0, nameEnd - 1)
}

function isSelfClosingXml(xml: string): boolean {
	return /\/>\s*$/.test(xml)
}

function findTagStart(xml: string, tag: string): number {
	return findTagStartFrom(xml, tag, 0)
}

function findTagStartFrom(xml: string, tag: string, startIndex: number): number {
	let index = startIndex
	const needle = `<${tag}`
	while (index < xml.length) {
		const found = xml.indexOf(needle, index)
		if (found < 0) return -1
		const next = xml.charCodeAt(found + needle.length)
		if (isXmlNameTerminator(next)) return found
		index = found + needle.length
	}
	return -1
}

function attrsFromOpenTag(
	xml: string,
	tagStart: number,
	tagNameLength: number,
): string | undefined {
	const openEnd = xml.indexOf('>', tagStart)
	if (openEnd < 0) return undefined
	let attrEnd = openEnd
	while (attrEnd > tagStart && /\s/.test(xml.charAt(attrEnd - 1))) attrEnd--
	if (xml.charCodeAt(attrEnd - 1) === 47) attrEnd--
	return xml.slice(tagStart + tagNameLength + 1, attrEnd)
}

function isXmlNameTerminator(code: number): boolean {
	return code === 32 || code === 9 || code === 10 || code === 13 || code === 47 || code === 62
}

function rawXmlAttr(attrs: string, name: string): string | undefined {
	let index = 0
	while (index < attrs.length) {
		const found = attrs.indexOf(name, index)
		if (found < 0) return undefined
		const before = found === 0 ? 32 : attrs.charCodeAt(found - 1)
		const eq = found + name.length
		if (isXmlNameTerminator(before) && attrs.charCodeAt(eq) === 61) {
			const quote = attrs.charCodeAt(eq + 1)
			if (quote === 34 || quote === 39) {
				const end = attrs.indexOf(String.fromCharCode(quote), eq + 2)
				if (end >= 0) return attrs.slice(eq + 2, end)
			}
		}
		index = found + name.length
	}
	return undefined
}

function rawXmlNumberAttr(attrs: string, name: string): number | undefined {
	const raw = rawXmlAttr(attrs, name)
	if (raw === undefined) return undefined
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) ? value : undefined
}

function rawXmlBoolAttr(attrs: string, name: string): boolean | undefined {
	const raw = rawXmlAttr(attrs, name)
	if (raw === undefined) return undefined
	if (raw === '1' || raw === 'true') return true
	if (raw === '0' || raw === 'false') return false
	return undefined
}

function readFastBooleanElement(xml: string, tag: string): boolean | undefined {
	const attrs = firstChildAttrs(xml, tag)
	if (attrs === undefined) return undefined
	const value = rawXmlAttr(attrs, 'val')
	if (value === '0' || value === 'false') return false
	return true
}

function parseFastUnderline(xml: string): boolean | 'single' | 'double' | undefined {
	const attrs = firstChildAttrs(xml, 'u')
	if (attrs === undefined) return undefined
	const value = rawXmlAttr(attrs, 'val')
	if (value === 'none') return undefined
	if (value === 'double') return 'double'
	if (value === 'single') return 'single'
	return true
}

function parseColorAttrsFast(attrs: string): Color | undefined {
	const rgb = rawXmlAttr(attrs, 'rgb')
	if (rgb) return { kind: 'rgb', rgb: decodeXmlAttribute(rgb) }
	const theme = rawXmlNumberAttr(attrs, 'theme')
	if (theme !== undefined) {
		const tintRaw = rawXmlAttr(attrs, 'tint')
		const tint = tintRaw !== undefined ? Number(tintRaw) : undefined
		return tint !== undefined && Number.isFinite(tint)
			? { kind: 'theme', theme, tint }
			: { kind: 'theme', theme }
	}
	const indexed = rawXmlNumberAttr(attrs, 'indexed')
	if (indexed !== undefined) return { kind: 'indexed', index: indexed }
	if (rawXmlAttr(attrs, 'auto')) return { kind: 'auto' }
	return undefined
}

function parseBorderEdgeFast(edgeXml: string | undefined): BorderEdge | undefined {
	if (!edgeXml) return undefined
	const attrs = tagAttrs(edgeXml) ?? ''
	const style = rawXmlAttr(attrs, 'style')
	if (!style || style === 'none') return undefined
	const colorAttrs = firstChildAttrs(edgeXml, 'color')
	const color = colorAttrs ? parseColorAttrsFast(colorAttrs) : undefined
	return color ? { style: style as BorderLineStyle, color } : { style: style as BorderLineStyle }
}

function hasUnsupportedFastStyleXml(xml: string): boolean {
	return /<(?:charset|family|scheme|vertAlign|outline|shadow|condense|extend|rFont|decorative|script|swiss|modern)\b/i.test(
		xml,
	)
}

function decodeXmlAttribute(value: string): string {
	return value
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&')
}

function parseDxfs(
	ss: XmlNode,
	fonts: FontStyle[],
	fills: FillStyle[],
	borders: BorderStyle[],
	numFmts: Map<number, string>,
): readonly CellStyle[] {
	const dxfsNode = ss.dxfs as XmlNode | undefined
	if (!dxfsNode) return []
	return asArray<XmlNode>(dxfsNode.dxf as XmlNode | XmlNode[]).map((dxf) =>
		parseDxfStyle(dxf, fonts, fills, borders, numFmts),
	)
}

function parseDxfStyle(
	dxf: XmlNode,
	_fonts: FontStyle[],
	_fills: FillStyle[],
	_borders: BorderStyle[],
	numFmts: Map<number, string>,
): CellStyle {
	const style: WritablePartial<CellStyle> = {}

	const fontNode = dxf.font as XmlNode | undefined
	if (fontNode) {
		const font = parseFont(fontNode)
		if (hasProps(font)) style.font = font
	}

	const fillNode = dxf.fill as XmlNode | undefined
	if (fillNode) {
		const fill = parseFill(fillNode)
		if (hasProps(fill)) style.fill = fill
	}

	const borderNode = dxf.border as XmlNode | undefined
	if (borderNode) {
		const border = parseBorder(borderNode)
		if (hasProps(border)) style.border = border
	}

	const numFmtNode = dxf.numFmt as XmlNode | undefined
	if (numFmtNode) {
		const numFmtId = numAttr(numFmtNode, 'numFmtId')
		const formatCode =
			attr(numFmtNode, 'formatCode') ?? (numFmtId !== undefined ? numFmts.get(numFmtId) : undefined)
		if (formatCode) style.numberFormat = formatCode
	}

	const alignment = parseAlignment(dxf)
	if (alignment) style.alignment = alignment
	const protection = parseProtection(dxf)
	if (protection) style.protection = protection
	return style as CellStyle
}

function parseNumFmts(ss: XmlNode): Map<number, string> {
	const fmts = new Map(BUILTIN_NUM_FMTS)
	const node = ss.numFmts as XmlNode | undefined
	if (!node) return fmts

	for (const nf of asArray<XmlNode>(node.numFmt as XmlNode | XmlNode[])) {
		const id = numAttr(nf, 'numFmtId')
		const code = attr(nf, 'formatCode')
		if (id !== undefined && code) fmts.set(id, code)
	}
	return fmts
}

function parseFonts(ss: XmlNode): FontStyle[] {
	const node = ss.fonts as XmlNode | undefined
	if (!node) return [{}]

	return asArray<XmlNode>(node.font as XmlNode | XmlNode[]).map(parseFont)
}

function parseFont(f: XmlNode): FontStyle {
	const props: WritablePartial<FontStyle> = {}

	const nameNode = f.name
	if (typeof nameNode === 'object' && nameNode !== null) {
		const n = attr(nameNode as XmlNode, 'val')
		if (n) props.name = n
	}

	const szNode = f.sz
	if (typeof szNode === 'object' && szNode !== null) {
		const s = numAttr(szNode as XmlNode, 'val')
		if (s !== undefined) props.size = s
	}

	if (f.b !== undefined) props.bold = !isFalseElement(f.b)
	if (f.i !== undefined) props.italic = !isFalseElement(f.i)
	if (f.strike !== undefined) props.strikethrough = !isFalseElement(f.strike)

	const underline = parseUnderline(f)
	if (underline !== undefined) props.underline = underline

	const color = parseColor(f.color)
	if (color) props.color = color

	return props as FontStyle
}

function parseUnderline(node: XmlNode): boolean | 'single' | 'double' | undefined {
	const el = node.u
	if (el === undefined) return undefined
	if (typeof el === 'object' && el !== null) {
		const val = attr(el as XmlNode, 'val')
		if (val === 'double') return 'double'
		if (val === 'single') return 'single'
		if (val === 'none') return undefined
	}
	return true
}

function isFalseElement(el: unknown): boolean {
	if (typeof el === 'object' && el !== null) {
		const val = attr(el as XmlNode, 'val')
		return val === '0' || val === 'false'
	}
	return false
}

function parseFills(ss: XmlNode): FillStyle[] {
	const node = ss.fills as XmlNode | undefined
	if (!node) return [{}]

	return asArray<XmlNode>(node.fill as XmlNode | XmlNode[]).map(parseFill)
}

function parseFill(f: XmlNode): FillStyle {
	const gradientFill = f.gradientFill as XmlNode | undefined
	if (gradientFill && typeof gradientFill === 'object') {
		const gradient = parseGradientFill(gradientFill)
		if (gradient) return { gradient }
	}
	const pf = f.patternFill as XmlNode | undefined
	if (!pf || typeof pf !== 'object') return {}

	const props: WritablePartial<FillStyle> = {}
	const pattern = attr(pf, 'patternType')
	if (pattern) props.pattern = pattern as FillPattern

	const fg = parseColor(pf.fgColor)
	if (fg) props.fgColor = fg

	const bg = parseColor(pf.bgColor)
	if (bg) props.bgColor = bg

	return props as FillStyle
}

function parseGradientFill(node: XmlNode): FillStyle['gradient'] | undefined {
	const stops = asArray<XmlNode>(node.stop as XmlNode | XmlNode[])
		.map((stop) => {
			const position = numAttr(stop, 'position')
			const color = parseColor(stop.color)
			if (position === undefined || !color) return null
			return { position, color }
		})
		.filter((stop): stop is NonNullable<typeof stop> => stop !== null)
	if (stops.length === 0) return undefined
	const type = attr(node, 'type')
	const degree = numAttr(node, 'degree')
	const left = numAttr(node, 'left')
	const right = numAttr(node, 'right')
	const top = numAttr(node, 'top')
	const bottom = numAttr(node, 'bottom')
	return {
		...(type ? { type: type as 'linear' | 'path' } : {}),
		...(degree !== undefined ? { degree } : {}),
		...(left !== undefined ? { left } : {}),
		...(right !== undefined ? { right } : {}),
		...(top !== undefined ? { top } : {}),
		...(bottom !== undefined ? { bottom } : {}),
		stops,
	}
}

function parseBorders(ss: XmlNode): BorderStyle[] {
	const node = ss.borders as XmlNode | undefined
	if (!node) return [{}]

	return asArray<XmlNode>(node.border as XmlNode | XmlNode[]).map(parseBorder)
}

function parseBorder(b: XmlNode): BorderStyle {
	const props: WritablePartial<BorderStyle> = {}

	const top = parseBorderEdge(b.top)
	if (top) props.top = top
	const bottom = parseBorderEdge(b.bottom)
	if (bottom) props.bottom = bottom
	const left = parseBorderEdge(b.left)
	if (left) props.left = left
	const right = parseBorderEdge(b.right)
	if (right) props.right = right
	const diagonal = parseBorderEdge(b.diagonal)
	if (diagonal) props.diagonal = diagonal

	const diagUp = boolAttr(b, 'diagonalUp')
	if (diagUp !== undefined) props.diagonalUp = diagUp
	const diagDown = boolAttr(b, 'diagonalDown')
	if (diagDown !== undefined) props.diagonalDown = diagDown

	return props as BorderStyle
}

function parseBorderEdge(el: unknown): BorderEdge | undefined {
	if (typeof el !== 'object' || el === null) return undefined
	const node = el as XmlNode
	const style = attr(node, 'style')
	if (!style || style === 'none') return undefined

	const color = parseColor(node.color)
	const props: WritablePartial<BorderEdge> = { style: style as BorderLineStyle }
	if (color) props.color = color
	return props as BorderEdge
}

function parseColor(el: unknown): Color | undefined {
	if (typeof el !== 'object' || el === null) return undefined
	const node = el as XmlNode

	const rgb = attr(node, 'rgb')
	if (rgb) return { kind: 'rgb', rgb }

	const theme = numAttr(node, 'theme')
	if (theme !== undefined) {
		const tint = numAttr(node, 'tint')
		return tint !== undefined ? { kind: 'theme', theme, tint } : { kind: 'theme', theme }
	}

	const indexed = numAttr(node, 'indexed')
	if (indexed !== undefined) return { kind: 'indexed', index: indexed }

	if (attr(node, 'auto')) return { kind: 'auto' }
	return undefined
}

function buildCellStyles(
	ss: XmlNode,
	fonts: FontStyle[],
	fills: FillStyle[],
	borders: BorderStyle[],
	numFmts: Map<number, string>,
): Pick<ParsedStyles, 'cellStyles' | 'isDateFormat'> {
	const xfsNode = ss.cellXfs as XmlNode | undefined
	if (!xfsNode) return { cellStyles: [{}], isDateFormat: [false] }

	const cellStyles: CellStyle[] = []
	const isDateFormat: boolean[] = []

	for (const xf of asArray<XmlNode>(xfsNode.xf as XmlNode | XmlNode[])) {
		const fontId = numAttr(xf, 'fontId') ?? 0
		const fillId = numAttr(xf, 'fillId') ?? 0
		const borderId = numAttr(xf, 'borderId') ?? 0
		const numFmtId = numAttr(xf, 'numFmtId') ?? 0

		const font = fonts[fontId]
		const fill = fills[fillId]
		const border = borders[borderId]
		const formatCode = numFmts.get(numFmtId)
		const alignment = parseAlignment(xf)
		const protection = parseProtection(xf)
		if (
			isDefaultCellXf(fontId, fillId, borderId, numFmtId, formatCode) &&
			!alignment &&
			!protection
		) {
			cellStyles.push({})
			isDateFormat.push(false)
			continue
		}

		const style: WritablePartial<CellStyle> = {}
		if (font && hasProps(font)) style.font = font
		if (fill && hasProps(fill)) style.fill = fill
		if (border && hasProps(border)) style.border = border
		if (formatCode && formatCode !== 'General') {
			style.numberFormat = formatCode
		}

		if (alignment) style.alignment = alignment

		if (protection) style.protection = protection

		cellStyles.push(style as CellStyle)
		isDateFormat.push(checkDateFormat(numFmtId, formatCode))
	}

	if (cellStyles.length === 0) {
		cellStyles.push({})
		isDateFormat.push(false)
	}

	return { cellStyles, isDateFormat }
}

function isDefaultCellXf(
	fontId: number,
	fillId: number,
	borderId: number,
	numFmtId: number,
	formatCode: string | undefined,
): boolean {
	return (
		fontId === 0 &&
		fillId === 0 &&
		borderId === 0 &&
		numFmtId === 0 &&
		(formatCode === undefined || formatCode === 'General')
	)
}

function parseAlignment(xf: XmlNode): AlignmentStyle | undefined {
	const el = xf.alignment
	if (typeof el !== 'object' || el === null) return undefined
	const a = el as XmlNode

	const props: WritablePartial<AlignmentStyle> = {}

	const h = attr(a, 'horizontal')
	if (h) props.horizontal = h as HorizontalAlign
	const v = attr(a, 'vertical')
	if (v) props.vertical = v as VerticalAlign

	const wrap = boolAttr(a, 'wrapText')
	if (wrap !== undefined) props.wrapText = wrap
	const shrink = boolAttr(a, 'shrinkToFit')
	if (shrink !== undefined) props.shrinkToFit = shrink

	const rotation = numAttr(a, 'textRotation')
	if (rotation !== undefined) props.textRotation = rotation
	const indent = numAttr(a, 'indent')
	if (indent !== undefined) props.indent = indent
	const reading = numAttr(a, 'readingOrder')
	if (reading !== undefined) props.readingOrder = reading

	return hasProps(props) ? (props as AlignmentStyle) : undefined
}

function parseProtection(xf: XmlNode): { locked?: boolean; hidden?: boolean } | undefined {
	const el = xf.protection
	if (typeof el !== 'object' || el === null) return undefined
	const p = el as XmlNode

	const locked = boolAttr(p, 'locked')
	const hidden = boolAttr(p, 'hidden')
	if (locked === undefined && hidden === undefined) return undefined

	const props: { locked?: boolean; hidden?: boolean } = {}
	if (locked !== undefined) props.locked = locked
	if (hidden !== undefined) props.hidden = hidden
	return props
}

function checkDateFormat(numFmtId: number, formatCode: string | undefined): boolean {
	if (BUILTIN_DATE_FMT_IDS.has(numFmtId)) return true
	if (!formatCode) return false
	return hasDateTimeFormatToken(formatCode)
}

function hasDateTimeFormatToken(formatCode: string): boolean {
	for (let i = 0; i < formatCode.length; i++) {
		const ch = formatCode.charAt(i)
		if (ch === '"') {
			i = skipUntil(formatCode, i + 1, '"')
			continue
		}
		if (ch === '\\' || ch === '_' || ch === '*') {
			i++
			continue
		}
		if (ch === '[') {
			const end = skipUntil(formatCode, i + 1, ']')
			const content = formatCode.slice(i + 1, end).trim()
			if (/^[hmsHMS]+$/.test(content)) return true
			i = end
			continue
		}
		if (/[ymdhsYMDHS]/.test(ch)) return true
	}
	return false
}

function skipUntil(text: string, start: number, target: string): number {
	const index = text.indexOf(target, start)
	return index >= 0 ? index : text.length
}

function hasProps(obj: object): boolean {
	for (const _ in obj) return true
	return false
}

function parseCellStyleNames(ss: XmlNode): NamedStyleInfo[] {
	const node = ss.cellStyles as XmlNode | undefined
	if (!node) return []
	const result: NamedStyleInfo[] = []
	for (const cs of asArray<XmlNode>(node.cellStyle as XmlNode | XmlNode[])) {
		const name = attr(cs, 'name')
		if (!name) continue
		const builtinId = numAttr(cs, 'builtinId')
		const hidden = attr(cs, 'hidden') === '1' ? true : undefined
		result.push({
			name,
			...(builtinId !== undefined ? { builtinId } : {}),
			...(hidden !== undefined ? { hidden } : {}),
		})
	}
	return result
}

function countNodes(node: unknown, key: string): number {
	if (typeof node !== 'object' || node === null) return 0
	return asArray<XmlNode>((node as XmlNode)[key] as XmlNode | XmlNode[]).length
}
