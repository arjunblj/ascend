const XML_ENTITY_RE = /&(?:lt|gt|quot|apos|amp|#x[0-9a-fA-F]+|#\d+);/g
const XML_ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const XMLNS_PREFIX_RE = /\sxmlns:([A-Za-z_][\w.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const MAIN_SPREADSHEET_NAMESPACES = new Set([
	'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
	'http://purl.oclc.org/ooxml/spreadsheetml/main',
])
const XML_ENTITY_MAP: Record<string, string> = {
	'&lt;': '<',
	'&gt;': '>',
	'&quot;': '"',
	'&apos;': "'",
	'&amp;': '&',
}

function resolveXmlEntity(m: string): string {
	const mapped = XML_ENTITY_MAP[m]
	if (mapped) return mapped
	if (m.startsWith('&#x')) return String.fromCodePoint(Number.parseInt(m.slice(3, -1), 16))
	if (m.startsWith('&#')) return String.fromCodePoint(Number.parseInt(m.slice(2, -1), 10))
	return m
}

export function decodeXmlText(text: string): string {
	if (!text.includes('&')) return text
	return text.replace(XML_ENTITY_RE, resolveXmlEntity)
}

export function parseXmlAttributes(rawAttrs: string): Map<string, string> {
	const attrs = new Map<string, string>()
	XML_ATTR_RE.lastIndex = 0
	for (const match of rawAttrs.matchAll(XML_ATTR_RE)) {
		const key = match[1]
		const value = match[2] ?? match[3]
		if (!key || value === undefined) continue
		attrs.set(key, decodeXmlText(value))
	}
	return attrs
}

export function normalizeMainSpreadsheetNamespacePrefix(xml: string): string {
	let normalized = xml
	XMLNS_PREFIX_RE.lastIndex = 0
	for (const match of xml.matchAll(XMLNS_PREFIX_RE)) {
		const prefix = match[1]
		const namespace = match[2] ?? match[3]
		if (!prefix || namespace === undefined) continue
		if (!MAIN_SPREADSHEET_NAMESPACES.has(decodeXmlText(namespace))) continue
		normalized = normalized.replace(new RegExp(`(<\\/?)${escapeRegExp(prefix)}:`, 'g'), '$1')
	}
	return normalized
}

function escapeRegExp(value: string): string {
	return value.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&')
}

export function findTagEnd(xml: string, start: number): number {
	return xml.indexOf('>', start + 1)
}

export function isSelfClosingTag(xml: string, tagStart: number, tagEnd: number): boolean {
	for (let idx = tagEnd - 1; idx > tagStart; idx--) {
		const ch = xml[idx]
		if (!ch) break
		if (ch === '/') return true
		if (ch !== ' ' && ch !== '\n' && ch !== '\r' && ch !== '\t') return false
	}
	return false
}
