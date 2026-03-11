import type { CellValue, RichTextRun } from '@ascend/schema'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { decodeXmlText, findTagEnd, isSelfClosingTag } from './xml-utils.ts'

export interface SharedStringResolver {
	readonly count: number
	get(index: number): CellValue | undefined
}

export function parseSharedStrings(
	xml: string,
	options: {
		readonly normalize?: (value: CellValue) => CellValue
	} = {},
): SharedStringResolver {
	return createEagerSharedStrings(xml, options.normalize)
}

export function emptySharedStrings(): SharedStringResolver {
	return {
		count: 0,
		get(): CellValue | undefined {
			return undefined
		},
	}
}

function createEagerSharedStrings(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
): SharedStringResolver {
	const entries = parseSharedStringEntries(xml, normalize)
	if (entries.length === 0 && xml.includes('<si')) {
		const fallback = parseSharedStringEntriesWithDom(xml, normalize)
		if (fallback.length > 0) {
			return {
				count: fallback.length,
				get(index: number): CellValue | undefined {
					return fallback[index]
				},
			}
		}
	}

	return {
		count: entries.length,
		get(index: number): CellValue | undefined {
			return entries[index]
		},
	}
}

function parseSharedStringEntries(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
): CellValue[] {
	const entries: CellValue[] = []
	let cursor = 0
	while (true) {
		const open = xml.indexOf('<si', cursor)
		if (open === -1) break
		const tagEnd = findTagEnd(xml, open)
		if (tagEnd === -1) break
		if (isSelfClosingTag(xml, open, tagEnd)) {
			const parsed: CellValue = { kind: 'string', value: '' }
			entries.push(normalize ? normalize(parsed) : parsed)
			cursor = tagEnd + 1
			continue
		}
		const close = xml.indexOf('</si>', tagEnd + 1)
		if (close === -1) break
		const parsed = parseSharedStringChunk(xml.slice(tagEnd + 1, close))
		entries.push(normalize ? normalize(parsed) : parsed)
		cursor = close + 5
	}
	return entries
}

function parseSharedStringEntriesWithDom(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
): CellValue[] {
	const entries: CellValue[] = []
	const doc = parseXml(xml)
	const sst = doc.sst as XmlNode | undefined
	if (!sst) return entries
	for (const si of asArray<XmlNode>(sst.si as XmlNode | XmlNode[])) {
		const parsed = parseSharedStringNode(si)
		entries.push(normalize ? normalize(parsed) : parsed)
	}
	return entries
}

function parseSharedStringChunk(chunk: string): CellValue {
	if (!chunk.includes('<r')) {
		const text = extractTextContent(chunk)
		if (text !== undefined) return { kind: 'string', value: text }
	}

	const runs: RichTextRun[] = []
	let cursor = 0
	while (true) {
		const runOpen = chunk.indexOf('<r', cursor)
		if (runOpen === -1) break
		const runTagEnd = findTagEnd(chunk, runOpen)
		if (runTagEnd === -1) break
		const runClose = chunk.indexOf('</r>', runTagEnd + 1)
		if (runClose === -1) break
		const runBody = chunk.slice(runTagEnd + 1, runClose)
		runs.push(parseRunChunk(runBody))
		cursor = runClose + 4
	}
	if (runs.length === 0) return { kind: 'string', value: '' }
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
		return { kind: 'string', value: first.text }
	}
	return { kind: 'richText', runs }
}

function parseSharedStringNode(si: XmlNode): CellValue {
	if (si.t !== undefined) {
		return { kind: 'string', value: String(si.t) }
	}
	if (si.r !== undefined) {
		return parseRichText(si)
	}
	return { kind: 'string', value: '' }
}

function parseRichText(si: XmlNode): CellValue {
	const runs: RichTextRun[] = []

	for (const r of asArray<XmlNode>(si.r as XmlNode | XmlNode[])) {
		const text = r.t !== undefined ? String(r.t) : ''
		const rPr = r.rPr as XmlNode | undefined

		if (rPr && typeof rPr === 'object') {
			const run: RichTextRun = {
				text,
				...(rPr.b !== undefined ? { bold: true } : {}),
				...(rPr.i !== undefined ? { italic: true } : {}),
				...(rPr.u !== undefined ? { underline: true } : {}),
				...(rPr.strike !== undefined ? { strikethrough: true } : {}),
				...parseFontProps(rPr),
			}
			runs.push(run)
		} else {
			runs.push({ text })
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
		return { kind: 'string', value: first.text }
	}

	return { kind: 'richText', runs }
}

function parseRunChunk(chunk: string): RichTextRun {
	const text = extractTextContent(chunk) ?? ''
	const runProps = extractSectionContent(chunk, 'rPr')
	if (!runProps) return { text }
	const run: RichTextRun = {
		text,
		...(runProps.includes('<b') ? { bold: true } : {}),
		...(runProps.includes('<i') ? { italic: true } : {}),
		...(runProps.includes('<u') ? { underline: true } : {}),
		...(runProps.includes('<strike') ? { strikethrough: true } : {}),
		...parseFontPropsChunk(runProps),
	}
	return run
}

function parseFontProps(rPr: XmlNode): Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> {
	const result: Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> = {}

	const rFont = rPr.rFont
	if (typeof rFont === 'object' && rFont !== null) {
		const name = attr(rFont as XmlNode, 'val')
		if (name) (result as Record<string, unknown>).fontName = name
	}

	const sz = rPr.sz
	if (typeof sz === 'object' && sz !== null) {
		const size = numAttr(sz as XmlNode, 'val')
		if (size !== undefined) (result as Record<string, unknown>).fontSize = size
	}

	const color = rPr.color
	if (typeof color === 'object' && color !== null) {
		const rgb = attr(color as XmlNode, 'rgb')
		if (rgb) (result as Record<string, unknown>).color = rgb
	}

	return result
}

function parseFontPropsChunk(chunk: string): Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> {
	const result: Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> = {}
	const fontName = extractAttributeValue(chunk, 'rFont', 'val')
	if (fontName) (result as Record<string, unknown>).fontName = fontName
	const fontSize = extractAttributeValue(chunk, 'sz', 'val')
	if (fontSize !== undefined) {
		const parsed = Number(fontSize)
		if (!Number.isNaN(parsed)) (result as Record<string, unknown>).fontSize = parsed
	}
	const color = extractAttributeValue(chunk, 'color', 'rgb')
	if (color) (result as Record<string, unknown>).color = color
	return result
}

function extractTextContent(chunk: string): string | undefined {
	const textOpen = chunk.indexOf('<t')
	if (textOpen === -1) return undefined
	const textTagEnd = findTagEnd(chunk, textOpen)
	if (textTagEnd === -1) return undefined
	if (isSelfClosingTag(chunk, textOpen, textTagEnd)) return ''
	const textClose = chunk.indexOf('</t>', textTagEnd + 1)
	if (textClose === -1) return undefined
	return decodeXmlText(chunk.slice(textTagEnd + 1, textClose))
}

function extractSectionContent(chunk: string, tagName: string): string | undefined {
	const open = chunk.indexOf(`<${tagName}`)
	if (open === -1) return undefined
	const tagEnd = findTagEnd(chunk, open)
	if (tagEnd === -1) return undefined
	if (isSelfClosingTag(chunk, open, tagEnd)) return ''
	const close = chunk.indexOf(`</${tagName}>`, tagEnd + 1)
	if (close === -1) return undefined
	return chunk.slice(tagEnd + 1, close)
}

function extractAttributeValue(
	chunk: string,
	tagName: string,
	attribute: string,
): string | undefined {
	const open = chunk.indexOf(`<${tagName}`)
	if (open === -1) return undefined
	const tagEnd = findTagEnd(chunk, open)
	if (tagEnd === -1) return undefined
	const attrs = chunk.slice(open + tagName.length + 1, tagEnd)
	const needle = `${attribute}="`
	const start = attrs.indexOf(needle)
	if (start === -1) return undefined
	const valueStart = start + needle.length
	const valueEnd = attrs.indexOf('"', valueStart)
	if (valueEnd === -1) return undefined
	const value = attrs.slice(valueStart, valueEnd)
	return value ? decodeXmlText(value) : undefined
}
