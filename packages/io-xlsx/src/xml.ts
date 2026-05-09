import { XMLParser } from 'fast-xml-parser'

export type XmlNode = Record<string, unknown>

const ALTERNATE_CONTENT_RE = /<mc:AlternateContent\b[\s\S]*?>([\s\S]*?)<\/mc:AlternateContent>/g
const FALLBACK_RE = /<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/
const CHOICE_RE = /<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/
const XML_ESCAPE_RE = /[&<>"]/
const bunEscapeHtml = getBunEscapeHtml()

type BunEscapeHtml = (value: string) => string

const parser = new XMLParser({
	attributeNamePrefix: '@_',
	ignoreAttributes: false,
	parseTagValue: true,
	trimValues: true,
	processEntities: true,
})

export function parseXml(content: string): XmlNode {
	const normalized = content.includes('mc:AlternateContent')
		? normalizeMarkupCompatibility(content)
		: content
	return parser.parse(normalized) as XmlNode
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
	const firstEscape = s.search(XML_ESCAPE_RE)
	if (firstEscape === -1) return s
	if (bunEscapeHtml) return bunEscapeHtml(s)
	let escaped = s.slice(0, firstEscape)
	for (let i = firstEscape; i < s.length; i++) {
		const ch = s[i] as string
		escaped += escapeXmlChar(ch)
	}
	return escaped
}

function getBunEscapeHtml(): BunEscapeHtml | undefined {
	const maybeBun = (globalThis as { readonly Bun?: { readonly escapeHTML?: BunEscapeHtml } }).Bun
	return typeof maybeBun?.escapeHTML === 'function' ? maybeBun.escapeHTML.bind(maybeBun) : undefined
}

function normalizeMarkupCompatibility(content: string): string {
	return content.replace(ALTERNATE_CONTENT_RE, (_full, inner: string) => {
		const fallback = inner.match(FALLBACK_RE)?.[1]
		if (fallback !== undefined) return fallback
		return inner.match(CHOICE_RE)?.[1] ?? ''
	})
}
