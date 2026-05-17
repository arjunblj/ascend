import {
	parseThreadedCommentPersonEntriesXml,
	type ThreadedCommentPersonEntry,
} from './reader/comments.ts'
import { parseContentTypes } from './reader/content-types.ts'
import { maybeDecryptOoxmlPackage } from './reader/encryption.ts'
import {
	isExternalLinkPathRelationshipType,
	parseRelationships,
	REL_ACTIVE_X_CONTROL,
	REL_ACTIVE_X_CONTROL_BINARY,
	REL_CALC_CHAIN,
	REL_CHART,
	REL_CHART_COLOR_STYLE,
	REL_CHART_STYLE,
	REL_CHARTSHEET,
	REL_COMMENTS,
	REL_CONNECTIONS,
	REL_CONTROL_PROP,
	REL_CORE_PROPERTIES,
	REL_CUSTOM_PROPERTIES,
	REL_CUSTOM_PROPERTY,
	REL_CUSTOM_UI,
	REL_CUSTOM_UI_2007,
	REL_CUSTOM_XML,
	REL_CUSTOM_XML_PROPS,
	REL_DATA_MODEL,
	REL_DATA_MODEL_TABLE,
	REL_DIAGRAM_DATA,
	REL_DIGITAL_SIGNATURE,
	REL_DIGITAL_SIGNATURE_ORIGIN,
	REL_DRAWING,
	REL_EXTENDED_PROPERTIES,
	REL_EXTERNAL_LINK,
	REL_HYPERLINK,
	REL_IMAGE,
	REL_MACROSHEET,
	REL_OFFICE_DOC,
	REL_OLE_OBJECT,
	REL_PIVOT_CACHE_DEFINITION,
	REL_PIVOT_CACHE_RECORDS,
	REL_PIVOT_TABLE,
	REL_POWER_QUERY_MASHUP,
	REL_PRINTER_SETTINGS,
	REL_QUERY_TABLE,
	REL_REVISION_HEADERS,
	REL_REVISION_LOG,
	REL_SHARED_STRINGS,
	REL_SHEET_METADATA,
	REL_SLICER,
	REL_SLICER_CACHE,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_THREADED_COMMENT,
	REL_THUMBNAIL,
	REL_TIMELINE,
	REL_TIMELINE_CACHE,
	REL_USER_NAMES,
	REL_VBA_PROJECT,
	REL_VBA_PROJECT_SIGNATURE,
	REL_VML_DRAWING,
	REL_WORKSHEET,
	REL_XML_MAPS,
	type Relationship,
	resolvePath,
} from './reader/relationships.ts'
import { extractZip, type ZipArchive } from './reader/zip.ts'

export type XlsxPackageContentTypeSource = 'override' | 'default' | 'fallback' | 'package'

export type XlsxPackageOwnerScope =
	| 'package'
	| 'relationship-part'
	| 'workbook'
	| 'worksheet'
	| 'chartsheet'
	| 'macrosheet'
	| 'drawing'
	| 'chart'
	| 'pivot'
	| 'slicer'
	| 'timeline'
	| 'external-link'
	| 'custom-xml'
	| 'analytics'
	| 'active-content'
	| 'security'
	| 'document-properties'
	| 'metadata'
	| 'unknown'

export type XlsxPackageLossPolicy =
	| 'generated'
	| 'preserve-exact'
	| 'discard-on-recalc'
	| 'invalidate-on-edit'
	| 'inspect-only'
	| 'unknown-review-required'

export type XlsxPackagePreservationMode =
	| 'generated'
	| 'preserve-exact'
	| 'discarded-for-recalc'
	| 'invalidated-on-edit'
	| 'inspect-only'
	| 'review-required'

export interface XlsxPackageGraph {
	readonly parts: readonly XlsxPackageGraphPart[]
	readonly relationships: readonly XlsxPackageGraphRelationship[]
	readonly contentTypeDefaults: readonly XlsxPackageContentTypeDefault[]
	readonly contentTypeOverrides: readonly XlsxPackageContentTypeOverride[]
}

export interface XlsxPackageGraphPart {
	readonly path: string
	readonly contentType: string
	readonly contentTypeSource: XlsxPackageContentTypeSource
	readonly ownerScope: XlsxPackageOwnerScope
	readonly sourceRelationshipPart?: string
	readonly sourceRelationshipId?: string
	readonly sourceRelationshipType?: string
	readonly sourceRelationshipRawType?: string
	readonly sourceRelationshipRawTarget?: string
	readonly sourceRelationshipResolvedTarget?: string
	readonly sourceRelationshipTargetMode?: string
	readonly featureFamily: string
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly bytePreservationExpected: boolean
	readonly threadedCommentPersons?: readonly ThreadedCommentPersonEntry[]
}

export interface XlsxPackageGraphRelationship {
	readonly sourcePartPath: string
	readonly relationshipPartPath: string
	readonly id: string
	readonly type: string
	readonly rawType?: string
	readonly rawTarget: string
	readonly resolvedTarget?: string
	readonly targetMode?: string
	readonly featureFamily: string
}

export interface XlsxPackageContentTypeDefault {
	readonly extension: string
	readonly contentType: string
}

export interface XlsxPackageContentTypeOverride {
	readonly partPath: string
	readonly contentType: string
}

export interface InspectXlsxPackageGraphOptions {
	readonly password?: string
}

const CT_PACKAGE_CONTENT_TYPES = 'application/vnd.openxmlformats-package.content-types+xml'
const CT_RELS = 'application/vnd.openxmlformats-package.relationships+xml'
const CT_FALLBACK = 'application/octet-stream'

export function inspectXlsxPackageGraph(
	bytes: Uint8Array,
	options: InspectXlsxPackageGraphOptions = {},
): XlsxPackageGraph {
	let archive: ZipArchive
	try {
		archive = extractZip(bytes)
	} catch (error) {
		const decrypted = maybeDecryptOoxmlPackage(bytes, options.password)
		if (!decrypted.ok) throw new Error(decrypted.error.message)
		if (!decrypted.value) throw error
		archive = extractZip(decrypted.value)
	}
	const contentTypesXml = archive.readText('[Content_Types].xml') ?? ''
	const contentTypes = parseContentTypes(contentTypesXml)
	const relationships = collectPackageRelationships(archive)
	const incoming = mapIncomingRelationships(relationships)
	const parts = [...archive.entries()]
		.filter((entry) => !isIgnorablePackageEntry(entry.path))
		.map((entry) => {
			const incomingRelationships = incoming.get(entry.path) ?? []
			const primaryRelationship = pickPrimaryRelationship(incomingRelationships)
			const contentType = resolvePackageContentType(entry.path, contentTypes)
			const featureFamily = classifyPackageFeatureFamily(
				entry.path,
				contentType.value,
				primaryRelationship?.type,
			)
			const preservationPolicy = packageFeatureLossPolicy(featureFamily)
			return {
				path: entry.path,
				contentType: contentType.value,
				contentTypeSource: contentType.source,
				ownerScope: classifyOwnerScope(entry.path, incomingRelationships),
				...(primaryRelationship
					? {
							sourceRelationshipPart: primaryRelationship.relationshipPartPath,
							sourceRelationshipId: primaryRelationship.id,
							sourceRelationshipType: primaryRelationship.type,
							...(primaryRelationship.rawType
								? { sourceRelationshipRawType: primaryRelationship.rawType }
								: {}),
							sourceRelationshipRawTarget: primaryRelationship.rawTarget,
							...(primaryRelationship.resolvedTarget
								? { sourceRelationshipResolvedTarget: primaryRelationship.resolvedTarget }
								: {}),
							...(primaryRelationship.targetMode
								? { sourceRelationshipTargetMode: primaryRelationship.targetMode }
								: {}),
						}
					: {}),
				featureFamily,
				preservationPolicy,
				bytePreservationExpected:
					preservationPolicy === 'preserve-exact' || preservationPolicy === 'inspect-only',
				...(isThreadedCommentPersonsPart(entry.path)
					? {
							threadedCommentPersons: parseThreadedCommentPersonEntriesXml(
								archive.readText(entry.path) ?? '',
							),
						}
					: {}),
			}
		})
		.sort((left, right) => left.path.localeCompare(right.path))
	return {
		parts,
		relationships,
		contentTypeDefaults: [...contentTypes.defaults]
			.map(([extension, contentType]) => ({ extension, contentType }))
			.sort((left, right) => left.extension.localeCompare(right.extension)),
		contentTypeOverrides: [...contentTypes.overrides]
			.map(([partPath, contentType]) => ({ partPath, contentType }))
			.sort((left, right) => left.partPath.localeCompare(right.partPath)),
	}
}

function isThreadedCommentPersonsPart(path: string): boolean {
	return /^xl\/persons\/[^/]+\.xml$/i.test(path)
}

export function classifyPackageFeatureFamily(
	partPath: string,
	contentType = '',
	relType = '',
): string {
	const path = partPath.replace(/\\/g, '/')
	const lowerPath = path.toLowerCase()
	const lowerContentType = contentType.toLowerCase()
	const lowerRelType = relType.toLowerCase()
	if (path === '[Content_Types].xml') return 'packageContentTypes'
	if (path === '_rels/.rels' || path.endsWith('.rels')) return 'packageRelationships'
	if (isVendorSecurityPart(lowerPath, lowerRelType)) return 'preservedVendorSecurity'
	if (
		lowerRelType.endsWith('/relationships/metadata/core-properties') ||
		lowerRelType.endsWith('/relationships/metadata/thumbnail') ||
		lowerRelType.endsWith('/relationships/extended-properties') ||
		lowerRelType.endsWith('/relationships/custom-properties')
	) {
		return 'preservedDocumentProperties'
	}
	if (lowerRelType.includes('/relationships/digital-signature/')) return 'preservedSignature'
	if (path.startsWith('docProps/')) return 'preservedDocumentProperties'
	if (path === 'xl/workbook.xml') return 'workbook'
	if (/(^|\/)sharedStrings\.xml$/i.test(path)) return 'sharedStrings'
	if (/(^|\/)worksheets\/sheet\d+\.xml$/i.test(path)) return 'worksheet'
	if (/^xl\/worksheets\/sheet\d+_[^/]+\.xml$/i.test(path)) return 'preservedWorksheetSidecar'
	if (lowerRelType.endsWith('/relationships/vbaprojectsignature')) return 'preservedSignature'
	if (lowerRelType.endsWith('/relationships/vbaproject')) return 'preservedMacro'
	if (lowerRelType.includes('/relationships/activexcontrol')) return 'preservedActiveX'
	if (lowerRelType.endsWith('/relationships/ctrlprop')) return 'preservedControl'
	if (lowerRelType.endsWith('/relationships/oleobject')) return 'preservedEmbedding'
	if (lowerRelType.endsWith('/relationships/customxml')) return 'preservedCustomXml'
	if (lowerRelType.endsWith('/relationships/customxmlprops')) return 'preservedCustomXml'
	if (lowerRelType.endsWith('/relationships/xmlmaps')) return 'preservedCustomXml'
	if (lowerRelType.endsWith('/relationships/customproperty')) return 'preservedMetadata'
	if (lowerRelType.includes('/relationships/diagram')) return 'preservedDrawing'
	if (
		lowerRelType.endsWith('/relationships/revisionheaders') ||
		lowerRelType.endsWith('/relationships/revisionlog') ||
		lowerRelType.endsWith('/relationships/usernames')
	) {
		return 'preservedRevision'
	}
	if (lowerRelType.endsWith('/relationships/ui/extensibility')) return 'preservedCustomUi'
	if (lowerRelType.endsWith('/relationships/calcchain')) return 'preservedCalcChain'
	if (lowerRelType.endsWith('/relationships/sheetmetadata')) return 'preservedMetadata'
	if (path.includes('/chartsheets/')) return 'preservedChartSheet'
	if (path.includes('/macrosheets/')) return 'preservedMacroSheet'
	if (
		/(^|\/)charts\/style\d+\.xml$/i.test(path) ||
		lowerRelType.endsWith('/relationships/chartstyle')
	) {
		return 'preservedChartStyle'
	}
	if (
		/(^|\/)charts\/colors\d+\.xml$/i.test(path) ||
		lowerRelType.endsWith('/relationships/chartcolorstyle')
	) {
		return 'preservedChartColor'
	}
	if (path.includes('/charts/') || path.includes('/chartEx/')) return 'preservedChart'
	if (path.includes('/drawings/') && path.endsWith('.vml')) return 'preservedVml'
	if (path.includes('/drawings/')) return 'preservedDrawing'
	if (path.includes('/media/')) return 'preservedMedia'
	if (
		path.includes('/model/') ||
		lowerRelType.endsWith('/relationships/model') ||
		lowerRelType.endsWith('/relationships/modeltable')
	) {
		return 'preservedDataModel'
	}
	if (path.includes('/tables/')) return 'preservedTable'
	if (path.includes('/queryTables/') || lowerRelType.endsWith('/relationships/querytable')) {
		return 'preservedQueryTable'
	}
	if (lowerRelType.endsWith('/relationships/hyperlink')) return 'preservedHyperlink'
	if (lowerRelType.endsWith('/relationships/comments')) return 'preservedComments'
	if (lowerRelType.endsWith('/relationships/threadedcomment')) {
		return 'preservedThreadedComments'
	}
	if (/\/comments\d+\.xml$/i.test(path)) return 'preservedComments'
	if (path.includes('/threadedComments/') || path.includes('/persons/')) {
		return 'preservedThreadedComments'
	}
	if (lowerPath.includes('/richdata/')) return 'preservedMetadata'
	if (
		lowerPath.includes('/customproperty') ||
		lowerContentType.includes('spreadsheetml.customproperty') ||
		lowerRelType.endsWith('/relationships/customproperty')
	) {
		return 'preservedMetadata'
	}
	if (lowerPath.includes('/volatiledependencies/')) return 'preservedMetadata'
	if (/(^|\/)externalLinks\//.test(path) || lowerRelType.endsWith('/relationships/externallink')) {
		return 'preservedExternalLink'
	}
	if (/(^|\/)pivotTables\//.test(path) || /(^|\/)pivotCache\//.test(path)) {
		return 'preservedPivot'
	}
	if (
		path.includes('/slicers/') ||
		path.includes('/slicerCaches/') ||
		lowerRelType.endsWith('/relationships/slicer') ||
		lowerRelType.endsWith('/relationships/slicercache')
	) {
		return 'preservedSlicer'
	}
	if (
		path.includes('/timelines/') ||
		path.includes('/timelineCaches/') ||
		lowerRelType.endsWith('/relationships/timeline') ||
		lowerRelType.endsWith('/relationships/timelinecache')
	) {
		return 'preservedTimeline'
	}
	if (path.endsWith('/connections.xml') || lowerRelType.endsWith('/relationships/connections')) {
		return 'preservedConnection'
	}
	if (path.includes('/customData/') || lowerRelType.endsWith('/relationships/powerquerymashup')) {
		return 'preservedPowerQuery'
	}
	if (path.includes('/theme/')) return 'preservedTheme'
	if (path.includes('/styles.xml')) return 'preservedStyles'
	if (path.includes('/metadata')) return 'preservedMetadata'
	if (path.includes('/diagrams/') || lowerRelType.includes('/relationships/diagram')) {
		return 'preservedDrawing'
	}
	if (path === 'xl/xmlMaps.xml' || lowerRelType.endsWith('/relationships/xmlmaps')) {
		return 'preservedCustomXml'
	}
	if (/(^|\/)customXml\//i.test(path) || lowerRelType.endsWith('/relationships/customxml')) {
		return 'preservedCustomXml'
	}
	if (
		path.includes('/revisions/') ||
		lowerRelType.endsWith('/relationships/revisionheaders') ||
		lowerRelType.endsWith('/relationships/revisionlog') ||
		lowerRelType.endsWith('/relationships/usernames')
	) {
		return 'preservedRevision'
	}
	if (path.endsWith('/calcChain.xml')) return 'preservedCalcChain'
	if (
		path.includes('/vbaProjectSignature') ||
		path.startsWith('_xmlsignatures/') ||
		lowerRelType.endsWith('/relationships/vbaprojectsignature')
	) {
		return 'preservedSignature'
	}
	if (path.includes('/vbaProject') || lowerRelType.endsWith('/relationships/vbaproject')) {
		return 'preservedMacro'
	}
	if (path.includes('/activeX/') || lowerRelType.includes('/relationships/activexcontrol')) {
		return 'preservedActiveX'
	}
	if (path.includes('/ctrlProps/') || lowerRelType.endsWith('/relationships/ctrlprop')) {
		return 'preservedControl'
	}
	if (
		path.includes('/printerSettings/') ||
		lowerRelType.endsWith('/relationships/printersettings')
	) {
		return 'preservedPrinterSettings'
	}
	if (path.includes('/embeddings/') || lowerRelType.endsWith('/relationships/oleobject')) {
		return 'preservedEmbedding'
	}
	if (
		lowerPath.startsWith('customui/') ||
		lowerContentType.includes('customui') ||
		lowerRelType.endsWith('/relationships/ui/extensibility')
	) {
		return 'preservedCustomUi'
	}
	if (lowerContentType.includes('macro') || lowerRelType.includes('macro')) return 'preservedMacro'
	if (lowerContentType.includes('activex') || lowerRelType.includes('activex')) {
		return 'preservedActiveX'
	}
	return 'preservedOther'
}

function collectPackageRelationships(
	archive: ReturnType<typeof extractZip>,
): XlsxPackageGraphRelationship[] {
	const relationships: XlsxPackageGraphRelationship[] = []
	for (const entry of archive.entries()) {
		if (!entry.path.endsWith('.rels')) continue
		const sourcePartPath = sourcePartFromRelsPath(entry.path)
		if (sourcePartPath === null) continue
		const xml = archive.readText(entry.path)
		if (!xml) continue
		for (const relationship of parseRelationships(xml)) {
			const targetMode = relationship.targetMode
			const resolvedTarget =
				targetMode?.toLowerCase() === 'external'
					? undefined
					: resolvePath(sourcePartPath, relationship.target)
			relationships.push({
				sourcePartPath,
				relationshipPartPath: entry.path,
				id: relationship.id,
				type: relationship.type,
				...(relationship.rawType ? { rawType: relationship.rawType } : {}),
				rawTarget: relationship.target,
				...(resolvedTarget ? { resolvedTarget } : {}),
				...(targetMode ? { targetMode } : {}),
				featureFamily: classifyRelationshipFeatureFamily(relationship, resolvedTarget),
			})
		}
	}
	return relationships.sort((left, right) =>
		`${left.relationshipPartPath}\u0000${left.id}`.localeCompare(
			`${right.relationshipPartPath}\u0000${right.id}`,
		),
	)
}

function sourcePartFromRelsPath(path: string): string | null {
	if (path === '_rels/.rels') return ''
	const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(path)
	if (!match) return null
	const fileName = match[2]
	if (!fileName) return null
	return match[1] ? `${match[1]}/${fileName}` : fileName
}

function mapIncomingRelationships(
	relationships: readonly XlsxPackageGraphRelationship[],
): Map<string, XlsxPackageGraphRelationship[]> {
	const incoming = new Map<string, XlsxPackageGraphRelationship[]>()
	for (const relationship of relationships) {
		if (!relationship.resolvedTarget) continue
		const list = incoming.get(relationship.resolvedTarget)
		if (list) list.push(relationship)
		else incoming.set(relationship.resolvedTarget, [relationship])
	}
	for (const list of incoming.values()) {
		list.sort((left, right) => relationshipPriority(left) - relationshipPriority(right))
	}
	return incoming
}

function pickPrimaryRelationship(
	relationships: readonly XlsxPackageGraphRelationship[],
): XlsxPackageGraphRelationship | undefined {
	return relationships[0]
}

function relationshipPriority(relationship: XlsxPackageGraphRelationship): number {
	if (relationship.sourcePartPath === '') return 0
	if (relationship.sourcePartPath === 'xl/workbook.xml') return 1
	if (relationship.type === REL_WORKSHEET || relationship.type === REL_CHARTSHEET) return 2
	return 3
}

function resolvePackageContentType(
	partPath: string,
	contentTypes: ReturnType<typeof parseContentTypes>,
): { value: string; source: XlsxPackageContentTypeSource } {
	if (partPath === '[Content_Types].xml') {
		return { value: CT_PACKAGE_CONTENT_TYPES, source: 'package' }
	}
	const override = contentTypes.overrides.get(partPath)
	if (override) return { value: override, source: 'override' }
	const extension = partPath.split('.').pop()
	if (extension) {
		const defaultType = contentTypes.defaults.get(extension)
		if (defaultType) return { value: defaultType, source: 'default' }
	}
	if (partPath.endsWith('.rels')) return { value: CT_RELS, source: 'fallback' }
	return { value: CT_FALLBACK, source: 'fallback' }
}

function classifyOwnerScope(
	partPath: string,
	incomingRelationships: readonly XlsxPackageGraphRelationship[],
): XlsxPackageOwnerScope {
	if (partPath === '[Content_Types].xml' || partPath === '_rels/.rels') return 'package'
	if (partPath.endsWith('.rels')) return 'relationship-part'
	const primary = pickPrimaryRelationship(incomingRelationships)
	if (isVendorSecurityPart(partPath.toLowerCase(), primary?.type.toLowerCase() ?? '')) {
		return 'security'
	}
	if (partPath.startsWith('docProps/')) return 'document-properties'
	if (
		primary?.type === REL_CORE_PROPERTIES ||
		primary?.type === REL_THUMBNAIL ||
		primary?.type === REL_EXTENDED_PROPERTIES ||
		primary?.type === REL_CUSTOM_PROPERTIES
	) {
		return 'document-properties'
	}
	if (primary?.type === REL_OFFICE_DOC || partPath === 'xl/workbook.xml') return 'workbook'
	if (primary?.type === REL_WORKSHEET) return 'worksheet'
	if (primary?.type === REL_CHARTSHEET) return 'chartsheet'
	if (primary?.type === REL_MACROSHEET) return 'macrosheet'
	if (primary?.type === REL_DRAWING || primary?.type === REL_VML_DRAWING) return 'drawing'
	if (primary?.type === REL_CHART) return 'chart'
	if (
		primary?.type === REL_PIVOT_TABLE ||
		primary?.type === REL_PIVOT_CACHE_DEFINITION ||
		primary?.type === REL_PIVOT_CACHE_RECORDS
	) {
		return 'pivot'
	}
	if (
		primary?.type === REL_SLICER ||
		primary?.type === REL_SLICER_CACHE ||
		/(^|\/)(slicerCaches|slicers)\//i.test(partPath)
	) {
		return 'slicer'
	}
	if (
		primary?.type === REL_TIMELINE ||
		primary?.type === REL_TIMELINE_CACHE ||
		/(^|\/)(timelineCaches|timelines)\//i.test(partPath)
	) {
		return 'timeline'
	}
	if (primary?.type === REL_EXTERNAL_LINK) return 'external-link'
	if (/(^|\/)externalLinks\//.test(partPath)) return 'external-link'
	if (primary?.type === REL_CUSTOM_XML) return 'custom-xml'
	if (primary?.type === REL_CUSTOM_XML_PROPS) return 'custom-xml'
	if (primary?.type === REL_XML_MAPS) return 'custom-xml'
	if (/(^|\/)customXml\//i.test(partPath)) return 'custom-xml'
	if (
		primary?.type === REL_CONNECTIONS ||
		primary?.type === REL_DATA_MODEL ||
		primary?.type === REL_DATA_MODEL_TABLE ||
		primary?.type === REL_POWER_QUERY_MASHUP ||
		primary?.type === REL_QUERY_TABLE
	) {
		return 'analytics'
	}
	if (/(^|\/)(model|customData|queryTables)\//i.test(partPath)) return 'analytics'
	if (primary?.type === REL_VBA_PROJECT_SIGNATURE) return 'security'
	if (primary?.type === REL_DIGITAL_SIGNATURE_ORIGIN || primary?.type === REL_DIGITAL_SIGNATURE) {
		return 'security'
	}
	if (
		primary?.type === REL_ACTIVE_X_CONTROL ||
		primary?.type === REL_ACTIVE_X_CONTROL_BINARY ||
		primary?.type === REL_CONTROL_PROP ||
		primary?.type === REL_OLE_OBJECT ||
		primary?.type === REL_VBA_PROJECT ||
		primary?.type === REL_CUSTOM_UI ||
		primary?.type === REL_CUSTOM_UI_2007
	) {
		return 'active-content'
	}
	if (/(^|\/)(activeX|ctrlProps|embeddings|customUI)\//i.test(partPath)) {
		return 'active-content'
	}
	if (primary?.type === REL_CALC_CHAIN || primary?.type === REL_SHEET_METADATA) return 'metadata'
	if (primary?.type === REL_CUSTOM_PROPERTY) return 'metadata'
	if (primary?.type === REL_DIAGRAM_DATA) return 'drawing'
	if (
		primary?.type === REL_REVISION_HEADERS ||
		primary?.type === REL_REVISION_LOG ||
		primary?.type === REL_USER_NAMES
	) {
		return 'metadata'
	}
	if (/(^|\/)(metadata|calcChain)\.xml$/i.test(partPath)) return 'metadata'
	if (primary?.sourcePartPath.includes('/worksheets/')) return 'worksheet'
	if (primary?.sourcePartPath.includes('/tables/')) return 'worksheet'
	if (primary?.sourcePartPath.includes('/drawings/')) return 'drawing'
	if (primary?.sourcePartPath.includes('/charts/')) return 'chart'
	return 'unknown'
}

function classifyRelationshipFeatureFamily(
	relationship: Relationship,
	resolvedTarget: string | undefined,
): string {
	if (relationship.type === REL_OFFICE_DOC) return 'workbook'
	if (
		relationship.type === REL_CORE_PROPERTIES ||
		relationship.type === REL_THUMBNAIL ||
		relationship.type === REL_EXTENDED_PROPERTIES ||
		relationship.type === REL_CUSTOM_PROPERTIES
	) {
		return 'preservedDocumentProperties'
	}
	if (
		relationship.type === REL_DIGITAL_SIGNATURE_ORIGIN ||
		relationship.type === REL_DIGITAL_SIGNATURE
	) {
		return 'preservedSignature'
	}
	if (relationship.type === REL_WORKSHEET) return 'worksheet'
	if (relationship.type === REL_CHARTSHEET) return 'preservedChartSheet'
	if (relationship.type === REL_MACROSHEET) return 'preservedMacroSheet'
	if (relationship.type === REL_SHARED_STRINGS) return 'sharedStrings'
	if (relationship.type === REL_STYLES) return 'preservedStyles'
	if (relationship.type === REL_THEME) return 'preservedTheme'
	if (relationship.type === REL_CALC_CHAIN) return 'preservedCalcChain'
	if (relationship.type === REL_SHEET_METADATA) return 'preservedMetadata'
	if (relationship.type === REL_TABLE) return 'preservedTable'
	if (relationship.type === REL_QUERY_TABLE) return 'preservedQueryTable'
	if (relationship.type === REL_CONNECTIONS) return 'preservedConnection'
	if (relationship.type === REL_DATA_MODEL) return 'preservedDataModel'
	if (relationship.type === REL_DATA_MODEL_TABLE) return 'preservedDataModel'
	if (relationship.type === REL_POWER_QUERY_MASHUP) return 'preservedPowerQuery'
	if (relationship.type === REL_HYPERLINK) return 'preservedHyperlink'
	if (relationship.type === REL_CUSTOM_XML) return 'preservedCustomXml'
	if (relationship.type === REL_CUSTOM_XML_PROPS) return 'preservedCustomXml'
	if (relationship.type === REL_XML_MAPS) return 'preservedCustomXml'
	if (relationship.type === REL_CUSTOM_PROPERTY) return 'preservedMetadata'
	if (relationship.type === REL_DIAGRAM_DATA) return 'preservedDrawing'
	if (
		relationship.type === REL_REVISION_HEADERS ||
		relationship.type === REL_REVISION_LOG ||
		relationship.type === REL_USER_NAMES
	) {
		return 'preservedRevision'
	}
	if (relationship.type === REL_CUSTOM_UI || relationship.type === REL_CUSTOM_UI_2007) {
		return 'preservedCustomUi'
	}
	if (relationship.type === REL_COMMENTS) return 'preservedComments'
	if (relationship.type === REL_THREADED_COMMENT) return 'preservedThreadedComments'
	if (relationship.type === REL_DRAWING) return 'preservedDrawing'
	if (relationship.type === REL_VML_DRAWING) return 'preservedVml'
	if (relationship.type === REL_CHART) return 'preservedChart'
	if (relationship.type === REL_CHART_STYLE) return 'preservedChartStyle'
	if (relationship.type === REL_CHART_COLOR_STYLE) return 'preservedChartColor'
	if (relationship.type === REL_IMAGE) return 'preservedMedia'
	if (relationship.type === REL_EXTERNAL_LINK) return 'preservedExternalLink'
	if (isExternalLinkPathRelationshipType(relationship.type)) return 'preservedExternalLink'
	if (
		relationship.type === REL_PIVOT_TABLE ||
		relationship.type === REL_PIVOT_CACHE_DEFINITION ||
		relationship.type === REL_PIVOT_CACHE_RECORDS
	) {
		return 'preservedPivot'
	}
	if (relationship.type === REL_SLICER || relationship.type === REL_SLICER_CACHE) {
		return 'preservedSlicer'
	}
	if (relationship.type === REL_TIMELINE || relationship.type === REL_TIMELINE_CACHE) {
		return 'preservedTimeline'
	}
	if (relationship.type === REL_PRINTER_SETTINGS) return 'preservedPrinterSettings'
	if (relationship.type === REL_OLE_OBJECT) return 'preservedEmbedding'
	if (relationship.type === REL_CONTROL_PROP) return 'preservedControl'
	if (
		relationship.type === REL_ACTIVE_X_CONTROL ||
		relationship.type === REL_ACTIVE_X_CONTROL_BINARY
	) {
		return 'preservedActiveX'
	}
	if (relationship.type === REL_VBA_PROJECT_SIGNATURE) return 'preservedSignature'
	if (relationship.type === REL_VBA_PROJECT) return 'preservedMacro'
	return resolvedTarget
		? classifyPackageFeatureFamily(resolvedTarget, '', relationship.type)
		: classifyPackageFeatureFamily(relationship.target, '', relationship.type)
}

function isVendorSecurityPart(lowerPath: string, lowerRelType: string): boolean {
	return lowerPath.startsWith('ddp/') || lowerRelType.includes('schemas.dell.com/ddp/')
}

export function packageFeatureLossPolicy(featureFamily: string): XlsxPackageLossPolicy {
	if (
		featureFamily === 'packageContentTypes' ||
		featureFamily === 'packageRelationships' ||
		featureFamily === 'workbook' ||
		featureFamily === 'worksheet' ||
		featureFamily === 'sharedStrings'
	) {
		return 'generated'
	}
	if (
		featureFamily === 'preservedDataModel' ||
		featureFamily === 'preservedPowerQuery' ||
		featureFamily === 'preservedRevision'
	) {
		return 'inspect-only'
	}
	if (featureFamily === 'preservedCalcChain') return 'discard-on-recalc'
	if (featureFamily === 'preservedSignature') return 'invalidate-on-edit'
	if (featureFamily === 'preservedOther') return 'unknown-review-required'
	return 'preserve-exact'
}

export function xlsxPackagePreservationModeForPolicy(
	policy: XlsxPackageLossPolicy,
): XlsxPackagePreservationMode {
	switch (policy) {
		case 'generated':
			return 'generated'
		case 'preserve-exact':
			return 'preserve-exact'
		case 'discard-on-recalc':
			return 'discarded-for-recalc'
		case 'invalidate-on-edit':
			return 'invalidated-on-edit'
		case 'inspect-only':
			return 'inspect-only'
		case 'unknown-review-required':
			return 'review-required'
	}
}

function isIgnorablePackageEntry(partPath: string): boolean {
	return (
		partPath.endsWith('/') ||
		partPath === '.DS_Store' ||
		partPath.endsWith('/.DS_Store') ||
		partPath.startsWith('__MACOSX/')
	)
}
