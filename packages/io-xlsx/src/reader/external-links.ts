import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
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
	readonly externalBookSheetNames?: readonly string[]
	readonly externalBookDefinedNames?: readonly ExternalBookDefinedNameInfo[]
	readonly ddeService?: string
	readonly ddeTopic?: string
	readonly ddeItems?: readonly ExternalLinkDdeItemInfo[]
}

export interface ExternalBookDefinedNameInfo {
	readonly name: string
	readonly refersTo?: string
	readonly sheetId?: number
}

export interface ExternalLinkDdeItemInfo {
	readonly name: string
	readonly advise?: boolean
	readonly preferPicture?: boolean
	readonly ole?: boolean
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
	const externalBookMetadata = kind === 'externalBook' ? parseExternalBookMetadata(xml) : {}
	const ddeMetadata = kind === 'ddeLink' ? parseDdeLinkMetadata(xml) : {}
	return {
		kind,
		...(relationshipId ? { relationshipId } : {}),
		...externalBookMetadata,
		...ddeMetadata,
		...(kind === 'ddeLink' && attrs.get('ddeService')
			? { ddeService: attrs.get('ddeService') as string }
			: {}),
		...(kind === 'ddeLink' && attrs.get('ddeTopic')
			? { ddeTopic: attrs.get('ddeTopic') as string }
			: {}),
	}
}

function parseDdeLinkMetadata(xml: string): Pick<ExternalLinkInfo, 'ddeItems'> {
	let doc: XmlNode
	try {
		doc = parseXml(xml)
	} catch {
		return {}
	}
	const externalLink = firstElement(doc, 'externalLink')
	const ddeLink = childNode(externalLink, 'ddeLink') ?? firstElement(doc, 'ddeLink')
	if (!ddeLink) return {}
	const ddeItems = childNodes(childNode(ddeLink, 'ddeItems'), 'ddeItem')
		.map((node) => {
			const name = attr(node, 'name')
			if (!name) return undefined
			const advise = boolAttr(node, 'advise')
			const preferPicture = boolAttr(node, 'preferPic')
			const ole = boolAttr(node, 'ole')
			return {
				name,
				...(advise !== undefined ? { advise } : {}),
				...(preferPicture !== undefined ? { preferPicture } : {}),
				...(ole !== undefined ? { ole } : {}),
			}
		})
		.filter((value): value is ExternalLinkDdeItemInfo => value !== undefined)
	return ddeItems.length > 0 ? { ddeItems } : {}
}

function parseExternalBookMetadata(
	xml: string,
): Pick<ExternalLinkInfo, 'externalBookSheetNames' | 'externalBookDefinedNames'> {
	let doc: XmlNode
	try {
		doc = parseXml(xml)
	} catch {
		return {}
	}
	const externalLink = firstElement(doc, 'externalLink')
	const externalBook = childNode(externalLink, 'externalBook') ?? firstElement(doc, 'externalBook')
	if (!externalBook) return {}
	const sheetNames = childNodes(childNode(externalBook, 'sheetNames'), 'sheetName')
		.map((node) => attr(node, 'val'))
		.filter((value): value is string => value !== undefined)
	const definedNames = childNodes(childNode(externalBook, 'definedNames'), 'definedName')
		.map((node) => {
			const name = attr(node, 'name')
			if (!name) return undefined
			const refersTo = attr(node, 'refersTo')
			const sheetId = numAttr(node, 'sheetId')
			return {
				name,
				...(refersTo ? { refersTo } : {}),
				...(sheetId !== undefined ? { sheetId } : {}),
			}
		})
		.filter((value): value is ExternalBookDefinedNameInfo => value !== undefined)
	return {
		...(sheetNames.length > 0 ? { externalBookSheetNames: sheetNames } : {}),
		...(definedNames.length > 0 ? { externalBookDefinedNames: definedNames } : {}),
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

function firstElement(doc: XmlNode, localName: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(doc)) {
		if (key.startsWith('@_') || localPart(key) !== localName || !isXmlNode(value)) continue
		return value
	}
	return undefined
}

function childNode(node: XmlNode | undefined, localName: string): XmlNode | undefined {
	return childNodes(node, localName)[0]
}

function childNodes(node: XmlNode | undefined, localName: string): XmlNode[] {
	if (!node) return []
	const matches: XmlNode[] = []
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || localPart(key) !== localName) continue
		for (const item of asArray(value as XmlNode | XmlNode[] | undefined)) {
			if (isXmlNode(item)) matches.push(item)
		}
	}
	return matches
}

function isXmlNode(value: unknown): value is XmlNode {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function localPart(name: string): string {
	return name.includes(':') ? (name.split(':').pop() ?? name) : name
}
