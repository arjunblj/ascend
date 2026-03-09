import { asArray, attr, parseXml, type XmlNode } from '../xml.ts'

export interface Relationship {
	readonly id: string
	readonly type: string
	readonly target: string
}

export const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
export const REL_WORKSHEET =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
export const REL_STYLES =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
export const REL_SHARED_STRINGS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
export const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
export const REL_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'

export function parseRelationships(xml: string): Relationship[] {
	const doc = parseXml(xml)
	const root = doc.Relationships as XmlNode | undefined
	if (!root) return []

	const rels: Relationship[] = []
	for (const entry of asArray<XmlNode>(root.Relationship as XmlNode | XmlNode[])) {
		const id = attr(entry, 'Id')
		const type = attr(entry, 'Type')
		const target = attr(entry, 'Target')
		if (id && type && target) {
			rels.push({ id, type, target })
		}
	}
	return rels
}

export function getRelsPath(partPath: string): string {
	const lastSlash = partPath.lastIndexOf('/')
	const dir = lastSlash >= 0 ? partPath.substring(0, lastSlash + 1) : ''
	const name = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath
	return `${dir}_rels/${name}.rels`
}

export function resolvePath(basePart: string, target: string): string {
	if (target.startsWith('/')) return target.substring(1)
	const baseDir = basePart.substring(0, basePart.lastIndexOf('/') + 1)
	const segments = [...baseDir.split('/').filter(Boolean), ...target.split('/')]
	const resolved: string[] = []
	for (const segment of segments) {
		if (segment === '' || segment === '.') continue
		if (segment === '..') {
			resolved.pop()
			continue
		}
		resolved.push(segment)
	}
	return resolved.join('/')
}
