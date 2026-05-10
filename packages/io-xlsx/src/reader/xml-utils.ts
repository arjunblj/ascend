const XML_ENTITY_RE = /&(?:lt|gt|quot|apos|amp|#x[0-9a-fA-F]+|#\d+);/g
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

export function normalizeMainSpreadsheetNamespacePrefix(xml: string): string {
	const match =
		/\sxmlns:([A-Za-z_][\w.-]*)="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/.exec(
			xml,
		)
	const prefix = match?.[1]
	if (!prefix) return xml
	return xml.replace(new RegExp(`(<\\/?)${prefix}:`, 'g'), '$1')
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
