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
	const doc = parseXml(xml)
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

export function parseStylesLite(xml: string): ParsedStylesLite {
	const doc = parseXml(xml)
	const ss = doc.styleSheet as XmlNode | undefined
	if (!ss) {
		return {
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
	const cellXfCount = countNodes(ss.cellXfs, 'xf')
	const isDateFormat = buildDateFormatFlags(ss, numFmts)
	return {
		isDateFormat,
		metadata: {
			numFmtCount: Math.max(0, numFmts.size - BUILTIN_NUM_FMTS.size),
			fontCount: Math.max(0, countNodes(ss.fonts, 'font')),
			fillCount: Math.max(0, countNodes(ss.fills, 'fill')),
			borderCount: Math.max(0, countNodes(ss.borders, 'border')),
			cellXfCount: Math.max(0, cellXfCount),
			dxfCount: Math.max(0, countNodes(ss.dxfs, 'dxf')),
			tableStyleCount: Math.max(0, countNodes(ss.tableStyles, 'tableStyle')),
		},
	}
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
	const style: Record<string, unknown> = {}

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
	const gradientFill = f.gradientFill as XmlNode | undefined
	if (gradientFill && typeof gradientFill === 'object') {
		const gradient = parseGradientFill(gradientFill)
		if (gradient) return { gradient }
	}
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

function buildDateFormatFlags(ss: XmlNode, numFmts: Map<number, string>): boolean[] {
	const xfsNode = ss.cellXfs as XmlNode | undefined
	if (!xfsNode) return [false]
	const isDateFormat: boolean[] = []
	for (const xf of asArray<XmlNode>(xfsNode.xf as XmlNode | XmlNode[])) {
		const numFmtId = numAttr(xf, 'numFmtId') ?? 0
		isDateFormat.push(checkDateFormat(numFmtId, numFmts.get(numFmtId)))
	}
	return isDateFormat.length > 0 ? isDateFormat : [false]
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
