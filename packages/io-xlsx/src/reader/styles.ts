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
	VerticalAlign,
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'

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
	[14, 'mm-dd-yy'],
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
	readonly isDateFormat: boolean[]
}

export function parseStyles(xml: string): ParsedStyles {
	const doc = parseXml(xml)
	const ss = doc.styleSheet as XmlNode | undefined
	if (!ss) return { cellStyles: [{}], isDateFormat: [false] }

	const numFmts = parseNumFmts(ss)
	const fonts = parseFonts(ss)
	const fills = parseFills(ss)
	const borders = parseBorders(ss)

	return buildCellStyles(ss, fonts, fills, borders, numFmts)
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
	const props: Record<string, unknown> = {}

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
	const pf = f.patternFill as XmlNode | undefined
	if (!pf || typeof pf !== 'object') return {}

	const props: Record<string, unknown> = {}
	const pattern = attr(pf, 'patternType')
	if (pattern) props.pattern = pattern as FillPattern

	const fg = parseColor(pf.fgColor)
	if (fg) props.fgColor = fg

	const bg = parseColor(pf.bgColor)
	if (bg) props.bgColor = bg

	return props as FillStyle
}

function parseBorders(ss: XmlNode): BorderStyle[] {
	const node = ss.borders as XmlNode | undefined
	if (!node) return [{}]

	return asArray<XmlNode>(node.border as XmlNode | XmlNode[]).map(parseBorder)
}

function parseBorder(b: XmlNode): BorderStyle {
	const props: Record<string, unknown> = {}

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
	const props: Record<string, unknown> = { style: style as BorderLineStyle }
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
): ParsedStyles {
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

		const style: Record<string, unknown> = {}
		if (font && hasProps(font)) style.font = font
		if (fill && hasProps(fill)) style.fill = fill
		if (border && hasProps(border)) style.border = border
		if (formatCode && formatCode !== 'General') {
			style.numberFormat = formatCode
		}

		const alignment = parseAlignment(xf)
		if (alignment) style.alignment = alignment

		const protection = parseProtection(xf)
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

function parseAlignment(xf: XmlNode): AlignmentStyle | undefined {
	const el = xf.alignment
	if (typeof el !== 'object' || el === null) return undefined
	const a = el as XmlNode

	const props: Record<string, unknown> = {}

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

	const props: Record<string, unknown> = {}
	if (locked !== undefined) props.locked = locked
	if (hidden !== undefined) props.hidden = hidden
	return props as { locked?: boolean; hidden?: boolean }
}

function checkDateFormat(numFmtId: number, formatCode: string | undefined): boolean {
	if (BUILTIN_DATE_FMT_IDS.has(numFmtId)) return true
	if (!formatCode) return false
	const clean = formatCode.replace(/"[^"]*"/g, '').replace(/\\./g, '')
	return /[ymdhsYMDHS]/.test(clean) && !/[#0?]/.test(clean)
}

function hasProps(obj: object): boolean {
	for (const _ in obj) return true
	return false
}
