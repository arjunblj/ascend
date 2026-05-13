import type { CellValue, RichTextRun } from '@ascend/schema'
import { stringValue } from '@ascend/schema'
import { XMLParser } from 'fast-xml-parser'
import { asArray, attr, numAttr, type XmlNode } from '../xml.ts'
import {
	decodeXmlText,
	findTagEnd,
	isSelfClosingTag,
	normalizeMainSpreadsheetNamespacePrefix,
} from './xml-utils.ts'

type WritablePartial<T> = { -readonly [K in keyof T]?: T[K] }

const preserveTextParser = new XMLParser({
	attributeNamePrefix: '@_',
	ignoreAttributes: false,
	parseTagValue: true,
	trimValues: false,
	processEntities: true,
})

export interface SharedStringResolver {
	readonly count: number
	get(index: number): CellValue | undefined
	getString?(index: number): string | undefined
}

export function parseSharedStrings(
	xml: string,
	options: {
		readonly normalize?: (value: CellValue) => CellValue
		readonly normalizeString?: (value: string) => CellValue
		readonly lazy?: boolean
	} = {},
): SharedStringResolver {
	xml = normalizeMainSpreadsheetNamespacePrefix(xml)
	if (options.lazy) return createLazySharedStrings(xml, options.normalize, options.normalizeString)
	return createEagerSharedStrings(xml, options.normalize, options.normalizeString)
}

export function emptySharedStrings(): SharedStringResolver {
	return {
		count: 0,
		get(): CellValue | undefined {
			return undefined
		},
		getString(): string | undefined {
			return undefined
		},
	}
}

function createEagerSharedStrings(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
	normalizeString?: (value: string) => CellValue,
): SharedStringResolver {
	const entries = parseSharedStringEntries(xml, normalize, normalizeString)
	if (entries.length === 0 && xml.includes('<si')) {
		const fallback = parseSharedStringEntriesWithDom(xml, normalize)
		if (fallback.length > 0) {
			return {
				count: fallback.length,
				get(index: number): CellValue | undefined {
					return fallback[index]
				},
				getString(index: number): string | undefined {
					const entry = fallback[index]
					return entry?.kind === 'string' ? entry.value : undefined
				},
			}
		}
	}

	return {
		count: entries.length,
		get(index: number): CellValue | undefined {
			return entries[index]
		},
		getString(index: number): string | undefined {
			const entry = entries[index]
			return entry?.kind === 'string' ? entry.value : undefined
		},
	}
}

function createLazySharedStrings(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
	normalizeString?: (value: string) => CellValue,
): SharedStringResolver {
	const offsets: number[] = []
	let cursor = 0
	while (true) {
		const pos = xml.indexOf('<si', cursor)
		if (pos === -1) break
		offsets.push(pos)
		cursor = pos + 3
	}

	if (offsets.length === 0 && xml.includes('<si')) {
		return createEagerSharedStrings(xml, normalize)
	}

	const entries = new Array<CellValue | undefined>(offsets.length)
	const resolved = new Uint8Array(offsets.length)
	const plainTextEntries = new Array<string | null | undefined>(offsets.length)
	const plainTextResolved = new Uint8Array(offsets.length)

	return {
		count: offsets.length,
		getString(index: number): string | undefined {
			if (index < 0 || index >= offsets.length) return undefined
			if (plainTextResolved[index]) {
				const text = plainTextEntries[index]
				return text === null ? undefined : (text ?? '')
			}
			if (resolved[index]) {
				const entry = entries[index]
				return entry?.kind === 'string' ? entry.value : undefined
			}

			const start = offsets[index] as number
			const fastText = parseSimplePlainSharedStringText(xml, start)
			if (fastText !== undefined) {
				plainTextResolved[index] = 1
				plainTextEntries[index] = fastText
				return fastText
			}
			const tagEnd = findTagEnd(xml, start)
			if (tagEnd === -1) {
				plainTextResolved[index] = 1
				return undefined
			}

			let text: string | undefined
			if (isSelfClosingTag(xml, start, tagEnd)) {
				text = ''
			} else {
				const close = xml.indexOf('</si>', tagEnd + 1)
				text = close === -1 ? undefined : parsePlainSharedStringEntry(xml, tagEnd + 1, close)
			}
			plainTextResolved[index] = 1
			if (text === undefined) {
				plainTextEntries[index] = null
				return undefined
			}
			plainTextEntries[index] = text
			return text
		},
		get(index: number): CellValue | undefined {
			if (index < 0 || index >= offsets.length) return undefined
			if (resolved[index]) return entries[index]
			if (
				plainTextResolved[index] &&
				plainTextEntries[index] !== undefined &&
				plainTextEntries[index] !== null
			) {
				const text = plainTextEntries[index] ?? ''
				if (normalizeString) {
					const result = normalizeString(text)
					entries[index] = result
					resolved[index] = 1
					return result
				}
				const value = stringValue(text)
				const result = normalize ? normalize(value) : value
				entries[index] = result
				resolved[index] = 1
				return result
			}

			const start = offsets[index] as number
			const tagEnd = findTagEnd(xml, start)
			if (tagEnd === -1) {
				resolved[index] = 1
				return undefined
			}

			let value: CellValue
			if (isSelfClosingTag(xml, start, tagEnd)) {
				value = stringValue('')
			} else {
				const close = xml.indexOf('</si>', tagEnd + 1)
				if (close === -1) {
					resolved[index] = 1
					return undefined
				}
				value = parseSharedStringEntry(xml, tagEnd + 1, close)
			}

			const result = normalize ? normalize(value) : value
			entries[index] = result
			resolved[index] = 1
			return result
		},
	}
}

function parseSimplePlainSharedStringEntry(
	xml: string,
	start: number,
): { readonly text: string; readonly next: number } | undefined {
	if (!startsSimplePlainSharedString(xml, start)) return undefined
	const valueStart = start + 7
	const valueEnd = xml.indexOf('</t>', valueStart)
	if (valueEnd === -1) return undefined
	if (!endsSimplePlainSharedString(xml, valueEnd + 4)) return undefined
	const text = xml.slice(valueStart, valueEnd)
	return { text: text.includes('&') ? decodeXmlText(text) : text, next: valueEnd + 9 }
}

function parseSimplePlainSharedStringText(xml: string, start: number): string | undefined {
	if (!startsSimplePlainSharedString(xml, start)) return undefined
	const valueStart = start + 7
	const valueEnd = xml.indexOf('</t>', valueStart)
	if (valueEnd === -1) return undefined
	if (!endsSimplePlainSharedString(xml, valueEnd + 4)) return undefined
	const text = xml.slice(valueStart, valueEnd)
	return text.includes('&') ? decodeXmlText(text) : text
}

function startsSimplePlainSharedString(xml: string, start: number): boolean {
	return (
		start + 6 < xml.length &&
		xml.charCodeAt(start) === 60 &&
		xml.charCodeAt(start + 1) === 115 &&
		xml.charCodeAt(start + 2) === 105 &&
		xml.charCodeAt(start + 3) === 62 &&
		xml.charCodeAt(start + 4) === 60 &&
		xml.charCodeAt(start + 5) === 116 &&
		xml.charCodeAt(start + 6) === 62
	)
}

function endsSimplePlainSharedString(xml: string, start: number): boolean {
	return (
		start + 4 < xml.length &&
		xml.charCodeAt(start) === 60 &&
		xml.charCodeAt(start + 1) === 47 &&
		xml.charCodeAt(start + 2) === 115 &&
		xml.charCodeAt(start + 3) === 105 &&
		xml.charCodeAt(start + 4) === 62
	)
}

function parseSharedStringEntries(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
	normalizeString?: (value: string) => CellValue,
): CellValue[] {
	const entries: CellValue[] = []
	let cursor = 0
	while (true) {
		const open = xml.indexOf('<si', cursor)
		if (open === -1) break
		const fastPlain = parseSimplePlainSharedStringEntry(xml, open)
		if (fastPlain !== undefined) {
			entries.push(
				normalizeString
					? normalizeString(fastPlain.text)
					: normalizeCellValue(stringValue(fastPlain.text), normalize),
			)
			cursor = fastPlain.next
			continue
		}
		const tagEnd = findTagEnd(xml, open)
		if (tagEnd === -1) break
		if (isSelfClosingTag(xml, open, tagEnd)) {
			entries.push(
				normalizeString ? normalizeString('') : normalizeCellValue(stringValue(''), normalize),
			)
			cursor = tagEnd + 1
			continue
		}
		const close = xml.indexOf('</si>', tagEnd + 1)
		if (close === -1) break
		const parsed = parseSharedStringEntry(xml, tagEnd + 1, close)
		entries.push(normalizeSharedStringValue(parsed, normalize, normalizeString))
		cursor = close + 5
	}
	return entries
}

function normalizeSharedStringValue(
	value: CellValue,
	normalize?: (value: CellValue) => CellValue,
	normalizeString?: (value: string) => CellValue,
): CellValue {
	return value.kind === 'string' && normalizeString
		? normalizeString(value.value)
		: normalizeCellValue(value, normalize)
}

function normalizeCellValue(
	value: CellValue,
	normalize?: (value: CellValue) => CellValue,
): CellValue {
	return normalize ? normalize(value) : value
}

function parseSharedStringEntry(xml: string, start: number, end: number): CellValue {
	if (!hasTagInRange(xml, start, end, 'r')) {
		const text = extractTextContentRange(xml, start, end)
		if (text !== undefined) return stringValue(text)
	}
	return parseSharedStringChunk(xml.slice(start, end))
}

function parsePlainSharedStringEntry(xml: string, start: number, end: number): string | undefined {
	if (hasTagInRange(xml, start, end, 'r')) return undefined
	return extractTextContentRange(xml, start, end)
}

function hasTagInRange(xml: string, start: number, end: number, tagName: string): boolean {
	let cursor = start
	while (true) {
		const open = xml.indexOf('<', cursor)
		if (open === -1 || open >= end) return false
		const nameStart = open + 1
		if (
			xml.startsWith(tagName, nameStart) &&
			isXmlNameTerminator(xml.charCodeAt(nameStart + tagName.length))
		) {
			return true
		}
		cursor = nameStart
	}
}

function isXmlNameTerminator(code: number): boolean {
	return (
		code === 0x20 ||
		code === 0x2f ||
		code === 0x3e ||
		code === 0x09 ||
		code === 0x0a ||
		code === 0x0d
	)
}

function parseSharedStringEntriesWithDom(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
): CellValue[] {
	const entries: CellValue[] = []
	const doc = preserveTextParser.parse(xml) as XmlNode
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
		if (text !== undefined) return stringValue(text)
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
	if (runs.length === 0) return stringValue('')
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
		return stringValue(first.text)
	}
	return { kind: 'richText', runs }
}

function parseSharedStringNode(si: XmlNode): CellValue {
	if (si.t !== undefined) {
		return stringValue(getXmlTextValue(si.t))
	}
	if (si.r !== undefined) {
		return parseRichText(si)
	}
	return stringValue('')
}

function parseRichText(si: XmlNode): CellValue {
	const runs: RichTextRun[] = []

	for (const r of asArray<XmlNode>(si.r as XmlNode | XmlNode[])) {
		const text = r.t !== undefined ? getXmlTextValue(r.t) : ''
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
		return stringValue(first.text)
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
	const result: WritablePartial<Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'>> = {}

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
		const theme = numAttr(colorNode, 'theme')
		const tint = numAttr(colorNode, 'tint')
		const indexed = numAttr(colorNode, 'indexed')
		if (theme !== undefined) {
			result.color =
				tint !== undefined
					? { kind: 'theme' as const, theme, tint }
					: { kind: 'theme' as const, theme }
		} else if (indexed !== undefined) {
			result.color = { kind: 'indexed' as const, index: indexed }
		} else if (rgb) {
			result.color = rgb
		}
	}

	return result
}

function getXmlTextValue(node: unknown): string {
	if (node === undefined || node === null) return ''
	if (typeof node === 'string' || typeof node === 'number' || typeof node === 'boolean') {
		return String(node)
	}
	if (typeof node === 'object') {
		const text = (node as Record<string, unknown>)['#text']
		if (typeof text === 'string' || typeof text === 'number' || typeof text === 'boolean') {
			return String(text)
		}
	}
	return ''
}

function parseFontPropsChunk(chunk: string): Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> {
	const result: WritablePartial<Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'>> = {}
	const fontName = extractAttributeValue(chunk, 'rFont', 'val')
	if (fontName) result.fontName = fontName
	const fontSize = extractAttributeValue(chunk, 'sz', 'val')
	if (fontSize !== undefined) {
		const parsed = Number(fontSize)
		if (!Number.isNaN(parsed)) result.fontSize = parsed
	}
	const themeStr = extractAttributeValue(chunk, 'color', 'theme')
	const tintStr = extractAttributeValue(chunk, 'color', 'tint')
	const indexedStr = extractAttributeValue(chunk, 'color', 'indexed')
	const rgbStr = extractAttributeValue(chunk, 'color', 'rgb')
	if (themeStr !== undefined) {
		const theme = Number(themeStr)
		const tint = tintStr !== undefined ? Number(tintStr) : undefined
		if (!Number.isNaN(theme)) {
			result.color =
				tint !== undefined && !Number.isNaN(tint)
					? { kind: 'theme' as const, theme, tint }
					: { kind: 'theme' as const, theme }
		}
	} else if (indexedStr !== undefined) {
		const index = Number(indexedStr)
		if (!Number.isNaN(index)) {
			result.color = { kind: 'indexed' as const, index }
		}
	} else if (rgbStr) {
		result.color = rgbStr
	}
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

function extractTextContentRange(xml: string, start: number, end: number): string | undefined {
	const textOpen = xml.indexOf('<t', start)
	if (textOpen === -1 || textOpen >= end) return undefined
	const textTagEnd = findTagEnd(xml, textOpen)
	if (textTagEnd === -1 || textTagEnd >= end) return undefined
	if (isSelfClosingTag(xml, textOpen, textTagEnd)) return ''
	const textClose = xml.indexOf('</t>', textTagEnd + 1)
	if (textClose === -1 || textClose > end) return undefined
	const text = xml.slice(textTagEnd + 1, textClose)
	return text.includes('&') ? decodeXmlText(text) : text
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
