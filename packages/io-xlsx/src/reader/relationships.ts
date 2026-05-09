export interface Relationship {
	readonly id: string
	readonly type: string
	readonly target: string
	readonly targetMode?: string
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
export const REL_SHEET_METADATA =
	'http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata'
export const REL_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
export const REL_COMMENTS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
export const REL_DRAWING =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
export const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
export const REL_PIVOT_CACHE_DEFINITION =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
export const REL_PIVOT_TABLE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
export const REL_CALC_CHAIN =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'
export const REL_SLICER_CACHE = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
export const REL_VML_DRAWING =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
export const REL_THREADED_COMMENT =
	'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment'

const STRICT_REL_PREFIX = 'http://purl.oclc.org/ooxml/officeDocument/relationships/'
const TRANSITIONAL_REL_PREFIX =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

const RELATIONSHIP_RE = /<Relationship\b([^>]*)\/>/g
const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g

export function parseRelationships(xml: string): Relationship[] {
	const rels: Relationship[] = []
	for (const match of xml.matchAll(RELATIONSHIP_RE)) {
		const rawAttrs = match[1]
		if (!rawAttrs) continue
		let id: string | undefined
		let type: string | undefined
		let target: string | undefined
		let targetMode: string | undefined
		for (const attrMatch of rawAttrs.matchAll(ATTR_RE)) {
			const key = attrMatch[1]
			const value = attrMatch[2]
			if (!key || value === undefined) continue
			if (key === 'Id') id = value
			else if (key === 'Type') type = value
			else if (key === 'Target') target = value
			else if (key === 'TargetMode') targetMode = value
		}
		if (id && type && target) {
			rels.push({
				id,
				type: normalizeRelationshipType(type),
				target,
				...(targetMode ? { targetMode } : {}),
			})
		}
	}
	return rels
}

function normalizeRelationshipType(type: string): string {
	if (!type.startsWith(STRICT_REL_PREFIX)) return type
	const suffix = type.slice(STRICT_REL_PREFIX.length)
	if (suffix === 'sheetMetadata') return type
	return `${TRANSITIONAL_REL_PREFIX}${suffix}`
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
