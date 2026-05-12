import type {
	AlignmentStyle,
	BorderEdge,
	BorderStyle,
	Color,
	FillStyle,
	FontStyle,
	StyleId,
	StyleRegistry,
} from '@ascend/core'

function colorKey(c: Color | undefined): string {
	if (!c) return ''
	if (c.kind === 'theme') return `theme:${c.theme}:${c.tint ?? ''}`
	if (c.kind === 'rgb') return `rgb:${c.rgb}`
	if (c.kind === 'indexed') return `idx:${c.index}`
	return 'auto'
}

function borderEdgeKey(e: BorderEdge | undefined): string {
	if (!e) return ''
	return `${e.style ?? ''}:${colorKey(e.color)}`
}

function fontKey(f: FontStyle | undefined): string {
	if (!f || !hasProps(f)) return '{}'
	return [
		f.name ?? '',
		String(f.size ?? ''),
		String(f.bold ?? ''),
		String(f.italic ?? ''),
		String(f.underline ?? ''),
		String(f.strikethrough ?? ''),
		colorKey(f.color),
	].join('|')
}

function fillKey(f: FillStyle | undefined): string {
	if (!f || !hasProps(f)) return ''
	if (f.gradient) {
		return [
			'gradient',
			f.gradient.type ?? '',
			String(f.gradient.degree ?? ''),
			String(f.gradient.left ?? ''),
			String(f.gradient.right ?? ''),
			String(f.gradient.top ?? ''),
			String(f.gradient.bottom ?? ''),
			...f.gradient.stops.flatMap((stop) => [String(stop.position), colorKey(stop.color)]),
		].join('|')
	}
	return [f.pattern ?? 'none', colorKey(f.fgColor), colorKey(f.bgColor)].join('|')
}

function borderKey(b: BorderStyle | undefined): string {
	if (!b || !hasProps(b)) return '{}'
	return [
		borderEdgeKey(b.top),
		borderEdgeKey(b.bottom),
		borderEdgeKey(b.left),
		borderEdgeKey(b.right),
		borderEdgeKey(b.diagonal),
		String(b.diagonalUp ?? ''),
		String(b.diagonalDown ?? ''),
	].join('|')
}

import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'
import { readNumberXmlAttr, readXmlAttr, removeXmlAttr, setXmlAttr } from './xml-attrs.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

const BUILTIN_FMT_CODES = new Map<string, number>([
	['General', 0],
	['0', 1],
	['0.00', 2],
	['#,##0', 3],
	['#,##0.00', 4],
	['0%', 9],
	['0.00%', 10],
	['0.00E+00', 11],
	['# ?/?', 12],
	['# ??/??', 13],
	['mm-dd-yy', 14],
	['m/d/yy', 14],
	['d-mmm-yy', 15],
	['d-mmm', 16],
	['mmm-yy', 17],
	['h:mm AM/PM', 18],
	['h:mm:ss AM/PM', 19],
	['h:mm', 20],
	['h:mm:ss', 21],
	['m/d/yy h:mm', 22],
	['#,##0 ;(#,##0)', 37],
	['#,##0 ;[Red](#,##0)', 38],
	['#,##0.00;(#,##0.00)', 39],
	['#,##0.00;[Red](#,##0.00)', 40],
	['mm:ss', 45],
	['[h]:mm:ss', 46],
	['mmss.0', 47],
	['##0.0E+0', 48],
	['@', 49],
])

interface XfEntry {
	fontId: number
	fillId: number
	borderId: number
	numFmtId: number
	alignment?: AlignmentStyle
	protection?: { locked?: boolean; hidden?: boolean }
}

export interface StylesResult {
	xml: string
	xfMap: Map<number, number>
}

export function buildPreservedStylesXml(
	xml: string,
	preserved: import('@ascend/core').WorkbookPreservedStyles,
	registry: StyleRegistry,
): StylesResult | undefined {
	const parsed = parsePreservedStylesXml(xml)
	const xfMap = new Map<number, number>()
	for (let i = 0; i < registry.size; i++) {
		const xfIndex = preserved.xfByStyleId[i]
		if (xfIndex !== undefined) {
			xfMap.set(i, xfIndex)
			continue
		}
		const baseStyleId = preserved.baseStyleIdByStyleId?.[i]
		if (baseStyleId === undefined) return undefined
		const baseXfIndex = xfMap.get(baseStyleId) ?? preserved.xfByStyleId[baseStyleId]
		if (baseXfIndex === undefined) return undefined
		const style = registry.get(i as StyleId) ?? {}
		const baseStyle = registry.get(baseStyleId as StyleId) ?? {}
		if (!stylesDifferOnlyByNumberFormat(style, baseStyle)) return undefined
		const nextNumFmtId = resolvePreservedNumFmtId(style.numberFormat, parsed)
		const baseXfXml = parsed.xfEntries[baseXfIndex]
		if (!baseXfXml) return undefined
		parsed.appendedXfEntries.push(patchXfNumFmt(baseXfXml, nextNumFmtId))
		const appendedIndex = parsed.xfEntries.length + parsed.appendedXfEntries.length - 1
		xfMap.set(i, appendedIndex)
	}
	return { xml: serializePatchedStylesXml(parsed), xfMap }
}

export function buildStylesXml(
	registry: StyleRegistry,
	differentialStyles: readonly import('@ascend/core').CellStyle[] = [],
): StylesResult {
	const fonts: FontStyle[] = [{}]
	const fontKeys = new Map<string, number>([[fontKey(undefined), 0]])

	const fills: FillStyle[] = [{ pattern: 'none' }, { pattern: 'gray125' }]
	const fillKeys = new Map<string, number>([
		[fillKey({ pattern: 'none' }), 0],
		[fillKey({ pattern: 'gray125' }), 1],
	])

	const borders: BorderStyle[] = [{}]
	const borderKeys = new Map<string, number>([[borderKey(undefined), 0]])

	const customNumFmts = new Map<string, number>()
	let nextNumFmtId = 164

	const xfEntries: XfEntry[] = []
	const xfMap = new Map<number, number>()

	for (let i = 0; i < registry.size; i++) {
		const style = registry.get(i as StyleId) ?? {}
		const fontId = lookupOrAdd(style.font, fonts, fontKeys, fontKey)
		const fillId = lookupOrAddFill(style.fill, fills, fillKeys)
		const borderId = lookupOrAdd(style.border, borders, borderKeys, borderKey)
		const numFmtId = resolveNumFmtId(style.numberFormat, customNumFmts, nextNumFmtId)
		if (numFmtId >= nextNumFmtId) nextNumFmtId = numFmtId + 1

		xfMap.set(i, xfEntries.length)
		const entry: XfEntry = { fontId, fillId, borderId, numFmtId }
		if (style.alignment) entry.alignment = style.alignment
		if (style.protection) entry.protection = style.protection
		xfEntries.push(entry)
	}

	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<styleSheet xmlns="${NS}">`)

	if (customNumFmts.size > 0) {
		out.push(`<numFmts count="${customNumFmts.size}">`)
		for (const [code, id] of customNumFmts) {
			out.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`)
		}
		out.push('</numFmts>')
	}

	out.push(`<fonts count="${fonts.length}">`)
	for (const font of fonts) pushFontXml(out, font)
	out.push('</fonts>')

	out.push(`<fills count="${fills.length}">`)
	for (const fill of fills) pushFillXml(out, fill)
	out.push('</fills>')

	out.push(`<borders count="${borders.length}">`)
	for (const border of borders) pushBorderXml(out, border)
	out.push('</borders>')

	out.push(`<cellXfs count="${xfEntries.length}">`)
	for (const xf of xfEntries) pushXfXml(out, xf)
	out.push('</cellXfs>')

	if (differentialStyles.length > 0) {
		out.push(`<dxfs count="${differentialStyles.length}">`)
		for (const style of differentialStyles) {
			pushDxfXml(out, style)
		}
		out.push('</dxfs>')
	}

	out.push('</styleSheet>')

	return { xml: out.toString(), xfMap }
}

function lookupOrAdd<T extends object>(
	value: T | undefined,
	table: T[],
	keys: Map<string, number>,
	keyFn: (v: T) => string,
): number {
	const v = value && hasProps(value) ? value : ({} as T)
	const key = keyFn(v)
	let idx = keys.get(key)
	if (idx === undefined) {
		idx = table.length
		table.push(v)
		keys.set(key, idx)
	}
	return idx
}

function lookupOrAddFill(
	fill: FillStyle | undefined,
	table: FillStyle[],
	keys: Map<string, number>,
): number {
	const f: FillStyle = fill && hasProps(fill) ? fill : { pattern: 'none' }
	const key = fillKey(f)
	let idx = keys.get(key)
	if (idx === undefined) {
		idx = table.length
		table.push(f)
		keys.set(key, idx)
	}
	return idx
}

function resolveNumFmtId(
	fmt: string | undefined,
	custom: Map<string, number>,
	nextId: number,
): number {
	if (!fmt || fmt === 'General') return 0
	const builtin = BUILTIN_FMT_CODES.get(fmt)
	if (builtin !== undefined) return builtin
	let id = custom.get(fmt)
	if (id === undefined) {
		id = nextId
		custom.set(fmt, id)
	}
	return id
}

interface ParsedPreservedStyles {
	xml: string
	customNumFmtIds: Map<string, number>
	nextNumFmtId: number
	xfEntries: string[]
	appendedXfEntries: string[]
	appendedNumFmts: Array<{ id: number; code: string }>
}

function parsePreservedStylesXml(xml: string): ParsedPreservedStyles {
	const customNumFmtIds = new Map<string, number>()
	let nextNumFmtId = 164
	for (const match of xml.matchAll(/<numFmt\b([^>]*)\/>/g)) {
		const attrs = match[1] ?? ''
		const id = readNumberXmlAttr(attrs, 'numFmtId')
		const rawCode = readXmlAttr(attrs, 'formatCode')
		const code = rawCode !== undefined ? decodeXmlAttr(rawCode) : undefined
		if (id === undefined || !code) continue
		if (!BUILTIN_FMT_CODES.has(code)) customNumFmtIds.set(code, id)
		if (id >= nextNumFmtId) nextNumFmtId = id + 1
	}
	const cellXfsMatch = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)
	const xfEntries = cellXfsMatch?.[1]?.match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) ?? []
	return {
		xml,
		customNumFmtIds,
		nextNumFmtId,
		xfEntries,
		appendedXfEntries: [],
		appendedNumFmts: [],
	}
}

function resolvePreservedNumFmtId(
	formatCode: string | undefined,
	parsed: ParsedPreservedStyles,
): number {
	if (!formatCode || formatCode === 'General') return 0
	const builtin = BUILTIN_FMT_CODES.get(formatCode)
	if (builtin !== undefined) return builtin
	const existing = parsed.customNumFmtIds.get(formatCode)
	if (existing !== undefined) return existing
	const id = parsed.nextNumFmtId++
	parsed.customNumFmtIds.set(formatCode, id)
	parsed.appendedNumFmts.push({ id, code: formatCode })
	return id
}

function shallowEqual(a: object, b: object): boolean {
	const aEntries = Object.entries(a)
	const bKeys = Object.keys(b)
	if (aEntries.length !== bKeys.length) return false
	for (const [k, v] of aEntries) {
		if ((b as Record<string, unknown>)[k] !== v) return false
	}
	return true
}

function optionalFieldEqual(a: object | undefined, b: object | undefined): boolean {
	if (a === b) return true
	if (!a || !b) return false
	return shallowEqual(a, b)
}

function stylesDifferOnlyByNumberFormat(
	style: import('@ascend/core').CellStyle,
	baseStyle: import('@ascend/core').CellStyle,
): boolean {
	return (
		optionalFieldEqual(style.font, baseStyle.font) &&
		optionalFieldEqual(style.fill, baseStyle.fill) &&
		optionalFieldEqual(style.border, baseStyle.border) &&
		optionalFieldEqual(style.alignment, baseStyle.alignment) &&
		optionalFieldEqual(style.protection, baseStyle.protection)
	)
}

function patchXfNumFmt(xfXmlSource: string, numFmtId: number): string {
	let out = setXmlAttr(xfXmlSource, 'numFmtId', numFmtId)
	if (!/numFmtId=/.test(out)) {
		out = out.replace('<xf', `<xf numFmtId="${numFmtId}"`)
	}
	if (numFmtId === 0) {
		out = removeXmlAttr(out, 'applyNumberFormat')
		return out
	}
	if (/applyNumberFormat=/.test(out)) {
		return setXmlAttr(out, 'applyNumberFormat', '1')
	}
	return out.replace(/<xf\b/, '<xf applyNumberFormat="1"')
}

function serializePatchedStylesXml(parsed: ParsedPreservedStyles): string {
	let xml = parsed.xml
	if (parsed.appendedNumFmts.length > 0) {
		const appended = parsed.appendedNumFmts
			.map((entry) => `<numFmt numFmtId="${entry.id}" formatCode="${escapeXml(entry.code)}"/>`)
			.join('')
		if (/<numFmts\b[^>]*>/.test(xml)) {
			xml = xml.replace(
				/<numFmts\b([^>]*)>([\s\S]*?)<\/numFmts>/,
				(_match, attrs: string, body: string) => {
					const count = readNumberXmlAttr(attrs, 'count') ?? 0
					const updatedAttrs = setXmlAttr(attrs, 'count', count + parsed.appendedNumFmts.length)
					return `<numFmts${updatedAttrs}>${body}${appended}</numFmts>`
				},
			)
		} else {
			xml = xml.replace(
				/<styleSheet\b[^>]*>/,
				(match) =>
					`${match}<numFmts count="${parsed.appendedNumFmts.length}">${appended}</numFmts>`,
			)
		}
	}
	if (parsed.appendedXfEntries.length > 0) {
		const appended = parsed.appendedXfEntries.join('')
		xml = xml.replace(
			/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/,
			(_match, attrs: string, body: string) => {
				const count = readNumberXmlAttr(attrs, 'count') ?? 0
				const updatedAttrs = setXmlAttr(attrs, 'count', count + parsed.appendedXfEntries.length)
				return `<cellXfs${updatedAttrs}>${body}${appended}</cellXfs>`
			},
		)
	}
	return xml
}

function decodeXmlAttr(value: string): string {
	return value
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&')
}

function pushFontXml(out: ChunkedStringBuilder, font: FontStyle): void {
	out.push('<font>')
	if (font.bold) out.push('<b/>')
	if (font.italic) out.push('<i/>')
	if (font.strikethrough) out.push('<strike/>')
	if (font.underline !== undefined) {
		if (font.underline === 'double') out.push('<u val="double"/>')
		else if (font.underline === 'single') out.push('<u val="single"/>')
		else if (font.underline) out.push('<u/>')
	}
	if (font.size !== undefined) out.push(`<sz val="${font.size}"/>`)
	if (font.color) out.push(colorXml(font.color, 'color'))
	if (font.name) out.push(`<name val="${escapeXml(font.name)}"/>`)
	out.push('</font>')
}

function pushFillXml(out: ChunkedStringBuilder, fill: FillStyle): void {
	if (fill.gradient) {
		const attrs: string[] = []
		if (fill.gradient.type) attrs.push(`type="${fill.gradient.type}"`)
		if (fill.gradient.degree !== undefined) attrs.push(`degree="${fill.gradient.degree}"`)
		if (fill.gradient.left !== undefined) attrs.push(`left="${fill.gradient.left}"`)
		if (fill.gradient.right !== undefined) attrs.push(`right="${fill.gradient.right}"`)
		if (fill.gradient.top !== undefined) attrs.push(`top="${fill.gradient.top}"`)
		if (fill.gradient.bottom !== undefined) attrs.push(`bottom="${fill.gradient.bottom}"`)
		out.push('<fill>')
		out.push(`<gradientFill${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}>`)
		for (const stop of fill.gradient.stops) {
			out.push(`<stop position="${stop.position}">`)
			out.push(colorXml(stop.color, 'color'))
			out.push('</stop>')
		}
		out.push('</gradientFill>')
		out.push('</fill>')
		return
	}
	const pattern = fill.pattern ?? 'none'
	if (!fill.fgColor && !fill.bgColor) {
		out.push(`<fill><patternFill patternType="${pattern}"/></fill>`)
		return
	}
	out.push('<fill>')
	out.push(`<patternFill patternType="${pattern}">`)
	if (fill.fgColor) out.push(colorXml(fill.fgColor, 'fgColor'))
	if (fill.bgColor) out.push(colorXml(fill.bgColor, 'bgColor'))
	out.push('</patternFill>')
	out.push('</fill>')
}

function pushBorderXml(out: ChunkedStringBuilder, border: BorderStyle): void {
	const attrs: string[] = []
	if (border.diagonalUp) attrs.push(' diagonalUp="1"')
	if (border.diagonalDown) attrs.push(' diagonalDown="1"')

	out.push(`<border${attrs.join('')}>`)
	out.push(edgeXml(border.left, 'left'))
	out.push(edgeXml(border.right, 'right'))
	out.push(edgeXml(border.top, 'top'))
	out.push(edgeXml(border.bottom, 'bottom'))
	out.push(edgeXml(border.diagonal, 'diagonal'))
	out.push('</border>')
}

function edgeXml(edge: BorderEdge | undefined, tag: string): string {
	if (!edge?.style || edge.style === 'none') return `<${tag}/>`
	if (edge.color) {
		return `<${tag} style="${edge.style}">${colorXml(edge.color, 'color')}</${tag}>`
	}
	return `<${tag} style="${edge.style}"/>`
}

function colorXml(color: Color, tag: string): string {
	switch (color.kind) {
		case 'rgb':
			return `<${tag} rgb="${escapeXml(color.rgb)}"/>`
		case 'theme':
			return color.tint !== undefined
				? `<${tag} theme="${color.theme}" tint="${color.tint}"/>`
				: `<${tag} theme="${color.theme}"/>`
		case 'indexed':
			return `<${tag} indexed="${color.index}"/>`
		case 'auto':
			return `<${tag} auto="1"/>`
	}
}

function pushXfXml(out: ChunkedStringBuilder, xf: XfEntry): void {
	const attrs: string[] = [
		`numFmtId="${xf.numFmtId}"`,
		`fontId="${xf.fontId}"`,
		`fillId="${xf.fillId}"`,
		`borderId="${xf.borderId}"`,
	]

	if (xf.numFmtId !== 0) attrs.push('applyNumberFormat="1"')
	if (xf.fontId !== 0) attrs.push('applyFont="1"')
	if (xf.fillId !== 0) attrs.push('applyFill="1"')
	if (xf.borderId !== 0) attrs.push('applyBorder="1"')
	if (xf.alignment) attrs.push('applyAlignment="1"')
	if (xf.protection) attrs.push('applyProtection="1"')

	if (!xf.alignment && !xf.protection) {
		out.push(`<xf ${attrs.join(' ')}/>`)
		return
	}

	out.push(`<xf ${attrs.join(' ')}>`)
	if (xf.alignment) out.push(alignmentXml(xf.alignment))
	if (xf.protection) out.push(protectionXml(xf.protection))
	out.push('</xf>')
}

function pushDxfXml(out: ChunkedStringBuilder, style: import('@ascend/core').CellStyle): void {
	out.push('<dxf>')
	if (style.font && hasProps(style.font)) pushFontXml(out, style.font)
	if (style.fill && hasProps(style.fill)) pushFillXml(out, style.fill)
	if (style.border && hasProps(style.border)) pushBorderXml(out, style.border)
	if (style.numberFormat && style.numberFormat !== 'General') {
		const builtin = BUILTIN_FMT_CODES.get(style.numberFormat)
		const numFmtId = builtin ?? 164
		out.push(`<numFmt numFmtId="${numFmtId}" formatCode="${escapeXml(style.numberFormat)}"/>`)
	}
	if (style.alignment) out.push(alignmentXml(style.alignment))
	if (style.protection) out.push(protectionXml(style.protection))
	out.push('</dxf>')
}

function alignmentXml(a: AlignmentStyle): string {
	const attrs: string[] = []
	if (a.horizontal) attrs.push(`horizontal="${a.horizontal}"`)
	if (a.vertical) attrs.push(`vertical="${a.vertical}"`)
	if (a.wrapText) attrs.push('wrapText="1"')
	if (a.shrinkToFit) attrs.push('shrinkToFit="1"')
	if (a.textRotation !== undefined) attrs.push(`textRotation="${a.textRotation}"`)
	if (a.indent !== undefined) attrs.push(`indent="${a.indent}"`)
	if (a.readingOrder !== undefined) attrs.push(`readingOrder="${a.readingOrder}"`)
	return `<alignment ${attrs.join(' ')}/>`
}

function protectionXml(p: { locked?: boolean; hidden?: boolean }): string {
	const attrs: string[] = []
	if (p.locked !== undefined) attrs.push(`locked="${p.locked ? '1' : '0'}"`)
	if (p.hidden !== undefined) attrs.push(`hidden="${p.hidden ? '1' : '0'}"`)
	return `<protection ${attrs.join(' ')}/>`
}

function hasProps(obj: object): boolean {
	for (const _ in obj) return true
	return false
}
