import { parseXmlAttributes } from './xml-utils.ts'

const NS_PREFIX = String.raw`(?:[A-Za-z_][\w.-]*:)?`
const EXTERNAL_BOOK_RE = new RegExp(`<${NS_PREFIX}externalBook\\b([^>]*)/?>`)

export function parseExternalBookRelationshipId(xml: string): string | undefined {
	const match = EXTERNAL_BOOK_RE.exec(xml)
	const rawAttrs = match?.[1]
	if (!rawAttrs) return undefined
	for (const [name, value] of parseXmlAttributes(rawAttrs)) {
		if (name === 'id' || name.endsWith(':id')) return value
	}
	return undefined
}
