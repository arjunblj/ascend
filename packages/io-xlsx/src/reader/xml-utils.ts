const XML_ENTITY_RE = /&(?:lt|gt|quot|apos|amp);/g
const XML_ENTITY_MAP: Record<string, string> = {
	'&lt;': '<',
	'&gt;': '>',
	'&quot;': '"',
	'&apos;': "'",
	'&amp;': '&',
}

export function decodeXmlText(text: string): string {
	if (!text.includes('&')) return text
	return text.replace(XML_ENTITY_RE, (m) => XML_ENTITY_MAP[m] ?? m)
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
