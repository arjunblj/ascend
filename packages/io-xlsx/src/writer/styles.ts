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
import { escapeXml } from '../xml.ts'

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
	preserved: import('@ascend/core').WorkbookPreservedStyles,
	registry: StyleRegistry,
): StylesResult | undefined {
	const xfMap = new Map<number, number>()
	for (let i = 0; i < registry.size; i++) {
		const xfIndex = preserved.xfByStyleId[i]
		if (xfIndex === undefined) return undefined
		xfMap.set(i, xfIndex)
	}
	return { xml: preserved.xml, xfMap }
}

export function buildStylesXml(
	registry: StyleRegistry,
	differentialStyles: readonly import('@ascend/core').CellStyle[] = [],
): StylesResult {
	const fonts: FontStyle[] = [{}]
	const fontKeys = new Map<string, number>([['{}', 0]])

	const fills: FillStyle[] = [{ pattern: 'none' }, { pattern: 'gray125' }]
	const fillKeys = new Map<string, number>([
		[JSON.stringify({ pattern: 'none' }), 0],
		[JSON.stringify({ pattern: 'gray125' }), 1],
	])

	const borders: BorderStyle[] = [{}]
	const borderKeys = new Map<string, number>([['{}', 0]])

	const customNumFmts = new Map<string, number>()
	let nextNumFmtId = 164

	const xfEntries: XfEntry[] = []
	const xfMap = new Map<number, number>()

	for (let i = 0; i < registry.size; i++) {
		const style = registry.get(i as StyleId) ?? {}
		const fontId = lookupOrAdd(style.font, fonts, fontKeys)
		const fillId = lookupOrAddFill(style.fill, fills, fillKeys)
		const borderId = lookupOrAdd(style.border, borders, borderKeys)
		const numFmtId = resolveNumFmtId(style.numberFormat, customNumFmts, nextNumFmtId)
		if (numFmtId >= nextNumFmtId) nextNumFmtId = numFmtId + 1

		xfMap.set(i, xfEntries.length)
		const entry: XfEntry = { fontId, fillId, borderId, numFmtId }
		if (style.alignment) entry.alignment = style.alignment
		if (style.protection) entry.protection = style.protection
		xfEntries.push(entry)
	}

	const parts: string[] = [XML_HEADER, `<styleSheet xmlns="${NS}">`]

	if (customNumFmts.size > 0) {
		parts.push(`<numFmts count="${customNumFmts.size}">`)
		for (const [code, id] of customNumFmts) {
			parts.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`)
		}
		parts.push('</numFmts>')
	}

	parts.push(`<fonts count="${fonts.length}">`)
	for (const font of fonts) parts.push(fontXml(font))
	parts.push('</fonts>')

	parts.push(`<fills count="${fills.length}">`)
	for (const fill of fills) parts.push(fillXml(fill))
	parts.push('</fills>')

	parts.push(`<borders count="${borders.length}">`)
	for (const border of borders) parts.push(borderXml(border))
	parts.push('</borders>')

	parts.push(`<cellXfs count="${xfEntries.length}">`)
	for (const xf of xfEntries) parts.push(xfXml(xf))
	parts.push('</cellXfs>')

	if (differentialStyles.length > 0) {
		parts.push(`<dxfs count="${differentialStyles.length}">`)
		for (const style of differentialStyles) {
			parts.push(dxfXml(style))
		}
		parts.push('</dxfs>')
	}

	parts.push('</styleSheet>')

	return { xml: parts.join(''), xfMap }
}

function lookupOrAdd<T extends object>(
	value: T | undefined,
	table: T[],
	keys: Map<string, number>,
): number {
	const v = value && hasProps(value) ? value : ({} as T)
	const key = JSON.stringify(v)
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
	const key = JSON.stringify(f)
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

function fontXml(font: FontStyle): string {
	const parts: string[] = ['<font>']
	if (font.bold) parts.push('<b/>')
	if (font.italic) parts.push('<i/>')
	if (font.strikethrough) parts.push('<strike/>')
	if (font.underline !== undefined) {
		if (font.underline === 'double') parts.push('<u val="double"/>')
		else if (font.underline === 'single') parts.push('<u val="single"/>')
		else if (font.underline) parts.push('<u/>')
	}
	if (font.size !== undefined) parts.push(`<sz val="${font.size}"/>`)
	if (font.color) parts.push(colorXml(font.color, 'color'))
	if (font.name) parts.push(`<name val="${escapeXml(font.name)}"/>`)
	parts.push('</font>')
	return parts.join('')
}

function fillXml(fill: FillStyle): string {
	const pattern = fill.pattern ?? 'none'
	if (!fill.fgColor && !fill.bgColor) {
		return `<fill><patternFill patternType="${pattern}"/></fill>`
	}
	const parts: string[] = ['<fill>', `<patternFill patternType="${pattern}">`]
	if (fill.fgColor) parts.push(colorXml(fill.fgColor, 'fgColor'))
	if (fill.bgColor) parts.push(colorXml(fill.bgColor, 'bgColor'))
	parts.push('</patternFill>', '</fill>')
	return parts.join('')
}

function borderXml(border: BorderStyle): string {
	const attrs: string[] = []
	if (border.diagonalUp) attrs.push(' diagonalUp="1"')
	if (border.diagonalDown) attrs.push(' diagonalDown="1"')

	const parts: string[] = [`<border${attrs.join('')}>`]
	parts.push(edgeXml(border.left, 'left'))
	parts.push(edgeXml(border.right, 'right'))
	parts.push(edgeXml(border.top, 'top'))
	parts.push(edgeXml(border.bottom, 'bottom'))
	parts.push(edgeXml(border.diagonal, 'diagonal'))
	parts.push('</border>')
	return parts.join('')
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

function xfXml(xf: XfEntry): string {
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
		return `<xf ${attrs.join(' ')}/>`
	}

	const parts: string[] = [`<xf ${attrs.join(' ')}>`]
	if (xf.alignment) parts.push(alignmentXml(xf.alignment))
	if (xf.protection) parts.push(protectionXml(xf.protection))
	parts.push('</xf>')
	return parts.join('')
}

function dxfXml(style: import('@ascend/core').CellStyle): string {
	const parts: string[] = ['<dxf>']
	if (style.font && hasProps(style.font)) parts.push(fontXml(style.font))
	if (style.fill && hasProps(style.fill)) parts.push(fillXml(style.fill))
	if (style.border && hasProps(style.border)) parts.push(borderXml(style.border))
	if (style.numberFormat && style.numberFormat !== 'General') {
		const builtin = BUILTIN_FMT_CODES.get(style.numberFormat)
		const numFmtId = builtin ?? 164
		parts.push(`<numFmt numFmtId="${numFmtId}" formatCode="${escapeXml(style.numberFormat)}"/>`)
	}
	if (style.alignment) parts.push(alignmentXml(style.alignment))
	if (style.protection) parts.push(protectionXml(style.protection))
	parts.push('</dxf>')
	return parts.join('')
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
