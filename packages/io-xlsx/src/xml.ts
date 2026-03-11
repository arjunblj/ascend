import { XMLParser } from 'fast-xml-parser'

export type XmlNode = Record<string, unknown>

const ALTERNATE_CONTENT_RE = /<mc:AlternateContent\b[\s\S]*?>([\s\S]*?)<\/mc:AlternateContent>/g
const FALLBACK_RE = /<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/
const CHOICE_RE = /<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/

const parser = new XMLParser({
	attributeNamePrefix: '@_',
	ignoreAttributes: false,
	parseTagValue: true,
	trimValues: true,
	processEntities: true,
})

export function parseXml(content: string): XmlNode {
	return parser.parse(normalizeMarkupCompatibility(content)) as XmlNode
}

export function asArray<T>(val: T | T[] | undefined | null): T[] {
	if (val === undefined || val === null) return []
	return Array.isArray(val) ? val : [val]
}

export function attr(node: XmlNode, name: string): string | undefined {
	const val = node[`@_${name}`]
	if (val === undefined || val === null) return undefined
	return String(val)
}

export function numAttr(node: XmlNode, name: string): number | undefined {
	const val = node[`@_${name}`]
	if (val === undefined || val === null) return undefined
	const n = Number(val)
	return Number.isNaN(n) ? undefined : n
}

export function boolAttr(node: XmlNode, name: string): boolean | undefined {
	const val = attr(node, name)
	if (val === undefined) return undefined
	return val === '1' || val === 'true'
}

const XML_ESCAPE_RE = /[&<>"]/g

function escapeXmlChar(ch: string): string {
	switch (ch) {
		case '&':
			return '&amp;'
		case '<':
			return '&lt;'
		case '>':
			return '&gt;'
		case '"':
			return '&quot;'
		default:
			return ch
	}
}

export function escapeXml(s: string): string {
	if (s.length === 0) return s
	XML_ESCAPE_RE.lastIndex = 0
	if (!XML_ESCAPE_RE.test(s)) return s
	XML_ESCAPE_RE.lastIndex = 0
	return s.replace(XML_ESCAPE_RE, escapeXmlChar)
}

function normalizeMarkupCompatibility(content: string): string {
	return content.replace(ALTERNATE_CONTENT_RE, (_full, inner: string) => {
		const fallback = inner.match(FALLBACK_RE)?.[1]
		if (fallback !== undefined) return fallback
		return inner.match(CHOICE_RE)?.[1] ?? ''
	})
}
