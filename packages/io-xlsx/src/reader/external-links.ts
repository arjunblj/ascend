import { decodeXmlText } from './xml-utils.ts'

const NS_PREFIX = String.raw`(?:[A-Za-z_][\w.-]*:)?`
const EXTERNAL_BOOK_RE = new RegExp(`<${NS_PREFIX}externalBook\\b([^>]*)/?>`)
const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g

export function parseExternalBookRelationshipId(xml: string): string | undefined {
	const match = EXTERNAL_BOOK_RE.exec(xml)
	const rawAttrs = match?.[1]
	if (!rawAttrs) return undefined
	for (const attrMatch of rawAttrs.matchAll(ATTR_RE)) {
		const name = attrMatch[1]
		const value = attrMatch[2]
		if (!name || value === undefined) continue
		if (name === 'id' || name.endsWith(':id')) return decodeXmlText(value)
	}
	return undefined
}
