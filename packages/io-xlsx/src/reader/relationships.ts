import { parseXmlAttributes } from './xml-utils.ts'

export interface Relationship {
	readonly id: string
	readonly type: string
	readonly rawType?: string
	readonly target: string
	readonly targetMode?: string
	readonly extraAttributes?: readonly RelationshipAttribute[]
}

export interface RelationshipAttribute {
	readonly name: string
	readonly value: string
}

export const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
export const REL_CORE_PROPERTIES =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
export const REL_EXTENDED_PROPERTIES =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
export const REL_CUSTOM_PROPERTIES =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
export const REL_THUMBNAIL =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
export const REL_DIGITAL_SIGNATURE_ORIGIN =
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'
export const REL_DIGITAL_SIGNATURE =
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature'
export const REL_WORKSHEET =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
export const REL_CHARTSHEET =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet'
export const REL_MACROSHEET = 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet'
export const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
export const REL_CHART_STYLE = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
export const REL_CHART_COLOR_STYLE =
	'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
export const REL_STYLES =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
export const REL_SHARED_STRINGS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
export const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
export const REL_SHEET_METADATA =
	'http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata'
export const REL_XML_MAPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps'
export const REL_CUSTOM_PROPERTY =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty'
export const REL_DIAGRAM_DATA =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData'
export const REL_REVISION_HEADERS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders'
export const REL_REVISION_LOG =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog'
export const REL_USER_NAMES =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames'
export const REL_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
export const REL_QUERY_TABLE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable'
export const REL_HYPERLINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
export const REL_CUSTOM_XML =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
export const REL_CUSTOM_XML_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
export const REL_CUSTOM_UI =
	'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
export const REL_CUSTOM_UI_2007 =
	'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
export const REL_COMMENTS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
export const REL_DRAWING =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
export const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
export const REL_PIVOT_CACHE_DEFINITION =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
export const REL_PIVOT_CACHE_RECORDS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords'
export const REL_PIVOT_TABLE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
export const REL_CALC_CHAIN =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'
export const REL_SLICER = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
export const REL_SLICER_CACHE = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
export const REL_TIMELINE = 'http://schemas.microsoft.com/office/2011/relationships/timeline'
export const REL_TIMELINE_CACHE =
	'http://schemas.microsoft.com/office/2011/relationships/timelineCache'
export const REL_VML_DRAWING =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
export const REL_THREADED_COMMENT =
	'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment'
export const REL_EXTERNAL_LINK_PATH =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'
export const REL_EXTERNAL_LINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
export const REL_EXTERNAL_LINK_STARTUP_PATH =
	'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup'
export const REL_EXTERNAL_LINK_ALTERNATE_STARTUP_PATH =
	'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup'
export const REL_EXTERNAL_LINK_LIBRARY_PATH =
	'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary'
export const REL_EXTERNAL_LINK_MISSING_PATH =
	'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing'
export const REL_CONNECTIONS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
export const REL_DATA_MODEL = 'http://schemas.microsoft.com/office/2011/relationships/model'
export const REL_DATA_MODEL_TABLE =
	'http://schemas.microsoft.com/office/2011/relationships/modelTable'
export const REL_POWER_QUERY_MASHUP =
	'http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup'
export const REL_PRINTER_SETTINGS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
export const REL_OLE_OBJECT =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'
export const REL_CONTROL_PROP =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
export const REL_ACTIVE_X_CONTROL =
	'http://schemas.microsoft.com/office/2006/relationships/activeXControl'
export const REL_ACTIVE_X_CONTROL_BINARY =
	'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary'
export const REL_VBA_PROJECT = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
export const REL_VBA_PROJECT_SIGNATURE =
	'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature'

export type ExternalLinkPathRelationshipKind =
	| 'externalLinkPath'
	| 'xlStartup'
	| 'xlAlternateStartup'
	| 'xlLibrary'
	| 'xlPathMissing'
	| 'unknown'

const STRICT_REL_PREFIX = 'http://purl.oclc.org/ooxml/officeDocument/relationships/'
const TRANSITIONAL_REL_PREFIX =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const EXTERNAL_LINK_PATH_REL_TYPES = new Set([
	REL_EXTERNAL_LINK_PATH,
	REL_EXTERNAL_LINK_STARTUP_PATH,
	REL_EXTERNAL_LINK_ALTERNATE_STARTUP_PATH,
	REL_EXTERNAL_LINK_LIBRARY_PATH,
	REL_EXTERNAL_LINK_MISSING_PATH,
])

const RELATIONSHIP_RE =
	/<(?:[A-Za-z_][\w.-]*:)?Relationship\b([^>]*?)(?:\/>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Relationship>)/g

export function parseRelationships(xml: string): Relationship[] {
	const rels: Relationship[] = []
	for (const match of xml.matchAll(RELATIONSHIP_RE)) {
		const rawAttrs = match[1]
		if (!rawAttrs) continue
		const attrs = parseXmlAttributes(rawAttrs)
		const id = attrs.get('Id')
		const type = attrs.get('Type')
		const target = attrs.get('Target')
		const targetMode = attrs.get('TargetMode')
		if (id && type && target) {
			const normalizedType = normalizeRelationshipType(type)
			rels.push({
				id,
				type: normalizedType,
				...(normalizedType !== type ? { rawType: type } : {}),
				target,
				...(targetMode ? { targetMode } : {}),
				...relationshipExtraAttributes(attrs),
			})
		}
	}
	return rels
}

function relationshipExtraAttributes(
	attrs: ReadonlyMap<string, string>,
): { readonly extraAttributes: readonly RelationshipAttribute[] } | Record<string, never> {
	const extraAttributes: RelationshipAttribute[] = []
	for (const [name, value] of attrs) {
		if (isKnownRelationshipAttribute(name)) continue
		extraAttributes.push({ name, value })
	}
	return extraAttributes.length > 0 ? { extraAttributes } : {}
}

function isKnownRelationshipAttribute(name: string): boolean {
	return name === 'Id' || name === 'Type' || name === 'Target' || name === 'TargetMode'
}

export function isExternalLinkPathRelationshipType(type: string): boolean {
	return EXTERNAL_LINK_PATH_REL_TYPES.has(normalizeRelationshipType(type))
}

export function externalLinkPathRelationshipKind(type: string): ExternalLinkPathRelationshipKind
export function externalLinkPathRelationshipKind(type: undefined): undefined
export function externalLinkPathRelationshipKind(
	type: string | undefined,
): ExternalLinkPathRelationshipKind | undefined {
	if (!type) return undefined
	switch (normalizeRelationshipType(type)) {
		case REL_EXTERNAL_LINK_PATH:
			return 'externalLinkPath'
		case REL_EXTERNAL_LINK_STARTUP_PATH:
			return 'xlStartup'
		case REL_EXTERNAL_LINK_ALTERNATE_STARTUP_PATH:
			return 'xlAlternateStartup'
		case REL_EXTERNAL_LINK_LIBRARY_PATH:
			return 'xlLibrary'
		case REL_EXTERNAL_LINK_MISSING_PATH:
			return 'xlPathMissing'
		default:
			return 'unknown'
	}
}

function normalizeRelationshipType(type: string): string {
	if (!type.startsWith(STRICT_REL_PREFIX)) return type
	const suffix = type.slice(STRICT_REL_PREFIX.length)
	if (suffix === 'sheetMetadata') return type
	if (suffix === 'extendedProperties') return `${TRANSITIONAL_REL_PREFIX}extended-properties`
	return `${TRANSITIONAL_REL_PREFIX}${suffix}`
}

export function getRelsPath(partPath: string): string {
	const lastSlash = partPath.lastIndexOf('/')
	const dir = lastSlash >= 0 ? partPath.substring(0, lastSlash + 1) : ''
	const name = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath
	return `${dir}_rels/${name}.rels`
}

export function resolvePath(basePart: string, target: string): string {
	const normalizedTarget = target.replace(/\\/g, '/')
	if (normalizedTarget.startsWith('/')) return decodePackagePath(normalizedTarget.substring(1))
	const baseDir = basePart.substring(0, basePart.lastIndexOf('/') + 1)
	const segments = [...baseDir.split('/').filter(Boolean), ...normalizedTarget.split('/')]
	const resolved: string[] = []
	for (const segment of segments) {
		if (segment === '' || segment === '.') continue
		if (segment === '..') {
			resolved.pop()
			continue
		}
		resolved.push(decodePackagePathSegment(segment))
	}
	return resolved.join('/')
}

function decodePackagePath(path: string): string {
	return path
		.split('/')
		.map((segment) => decodePackagePathSegment(segment))
		.join('/')
}

function decodePackagePathSegment(segment: string): string {
	if (!segment.includes('%')) return segment
	const protectedSeparators = segment.replace(/%(2f|5c)/gi, (match) => `%25${match.slice(1)}`)
	try {
		return decodeURIComponent(protectedSeparators)
	} catch {
		return segment
	}
}
