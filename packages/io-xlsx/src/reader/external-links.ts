import { parseXmlAttributes } from './xml-utils.ts'

const NS_PREFIX = String.raw`(?:[A-Za-z_][\w.-]*:)?`
const EXTERNAL_LINK_SOURCE_RE = new RegExp(
	`<${NS_PREFIX}(externalBook|ddeLink|oleLink)\\b([^>]*)/?>`,
)
const RELATIONSHIP_NAMESPACES = new Set([
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	'http://purl.oclc.org/ooxml/officeDocument/relationships',
])

export type ExternalLinkKind = 'externalBook' | 'ddeLink' | 'oleLink'

export interface ExternalLinkInfo {
	readonly kind: ExternalLinkKind
	readonly relationshipId?: string
	readonly ddeService?: string
	readonly ddeTopic?: string
}

export function parseExternalBookRelationshipId(xml: string): string | undefined {
	const info = parseExternalLinkInfo(xml)
	return info?.kind === 'externalBook' ? info.relationshipId : undefined
}

export function parseExternalLinkInfo(xml: string): ExternalLinkInfo | undefined {
	const match = EXTERNAL_LINK_SOURCE_RE.exec(xml)
	const kind = match?.[1] as ExternalLinkKind | undefined
	const rawAttrs = match?.[2]
	if (!kind) return undefined
	if (!rawAttrs) return undefined
	const namespaces = relationshipNamespacePrefixes(xml, match.index)
	const attrs = parseXmlAttributes(rawAttrs)
	const relationshipId = readRelationshipId(attrs, namespaces)
	return {
		kind,
		...(relationshipId ? { relationshipId } : {}),
		...(kind === 'ddeLink' && attrs.get('ddeService')
			? { ddeService: attrs.get('ddeService') as string }
			: {}),
		...(kind === 'ddeLink' && attrs.get('ddeTopic')
			? { ddeTopic: attrs.get('ddeTopic') as string }
			: {}),
	}
}

function readRelationshipId(
	attrs: ReadonlyMap<string, string>,
	namespaces: ReadonlyMap<string, string>,
): string | undefined {
	for (const [name, value] of attrs) {
		if (name === 'id') return value
		const separator = name.indexOf(':')
		if (separator <= 0 || name.slice(separator + 1) !== 'id') continue
		const namespace = namespaces.get(name.slice(0, separator))
		if (namespace && RELATIONSHIP_NAMESPACES.has(namespace)) return value
	}
	return undefined
}

function relationshipNamespacePrefixes(
	xml: string,
	externalBookIndex: number,
): Map<string, string> {
	const namespaces = new Map<string, string>()
	let cursor = 0
	while (cursor <= externalBookIndex) {
		const start = xml.indexOf('<', cursor)
		if (start < 0 || start > externalBookIndex) break
		if (xml[start + 1] === '/' || xml.startsWith('<!--', start) || xml.startsWith('<?', start)) {
			cursor = start + 1
			continue
		}
		const end = xml.indexOf('>', start + 1)
		if (end < 0) break
		for (const [name, value] of parseXmlAttributes(xml.slice(start + 1, end))) {
			if (name.startsWith('xmlns:')) namespaces.set(name.slice('xmlns:'.length), value)
		}
		cursor = end + 1
	}
	return namespaces
}
