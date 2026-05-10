import { unzipSync } from 'fflate'

export interface OoxmlPackageProbe {
	readonly paths: readonly string[]
	readonly counts: Record<string, number>
	readonly features: Record<string, boolean>
	readonly analytics: OoxmlAnalyticsProbe
}

export interface OoxmlRelationshipProbe {
	readonly sourcePartPath: string
	readonly relsPartPath: string
	readonly id: string
	readonly type: string
	readonly target: string
	readonly targetMode?: string
	readonly targetPartPath?: string
}

export interface OoxmlWorkbookPivotCacheProbe {
	readonly cacheId: number
	readonly relId: string
}

export interface OoxmlPivotTableProbe {
	readonly partPath: string
	readonly name?: string
	readonly cacheId?: number
	readonly locationRef?: string
}

export interface OoxmlPivotCacheProbe {
	readonly partPath: string
	readonly cacheId?: number
	readonly relId?: string
	readonly recordsPartPath?: string
}

export interface OoxmlLinkedCacheProbe {
	readonly partPath: string
	readonly name?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
}

export interface OoxmlLinkedUiProbe {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
}

export interface OoxmlAnalyticsProbe {
	readonly relationships: readonly OoxmlRelationshipProbe[]
	readonly workbookPivotCaches: readonly OoxmlWorkbookPivotCacheProbe[]
	readonly pivotTableRelationships: readonly OoxmlRelationshipProbe[]
	readonly pivotCacheRelationships: readonly OoxmlRelationshipProbe[]
	readonly pivotCacheRecordRelationships: readonly OoxmlRelationshipProbe[]
	readonly slicerCacheRelationships: readonly OoxmlRelationshipProbe[]
	readonly pivotTables: readonly OoxmlPivotTableProbe[]
	readonly pivotCaches: readonly OoxmlPivotCacheProbe[]
	readonly pivotCacheRecords: readonly string[]
	readonly slicerCaches: readonly OoxmlLinkedCacheProbe[]
	readonly slicers: readonly OoxmlLinkedUiProbe[]
	readonly timelineCaches: readonly OoxmlLinkedCacheProbe[]
	readonly timelines: readonly OoxmlLinkedUiProbe[]
}

const decoder = new TextDecoder()
const PIVOT_CACHE_DEFINITION_REL =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
const PIVOT_CACHE_RECORDS_REL =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords'
const PIVOT_TABLE_REL =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
const SLICER_CACHE_REL = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
const STRICT_REL_PREFIX = 'http://purl.oclc.org/ooxml/officeDocument/relationships/'
const TRANSITIONAL_REL_PREFIX =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

export function inspectOoxmlPackageFeatures(bytes: Uint8Array): OoxmlPackageProbe {
	const files = unzipSync(bytes)
	const paths = Object.keys(files).sort((a, b) => a.localeCompare(b))
	const worksheetXml = paths
		.filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path))
		.map((path) => decoder.decode(files[path]))
	const workbookXml = decodePart(files, 'xl/workbook.xml')
	const contentTypesXml = decodePart(files, '[Content_Types].xml')
	const counts = {
		worksheets: worksheetXml.length,
		formulas: countXmlTags(worksheetXml, 'f'),
		charts: countPaths(paths, /^xl\/(?:charts|chartEx)\//i),
		tables: countPaths(paths, /^xl\/tables\/table\d+\.xml$/i),
		drawings: countPaths(paths, /^xl\/drawings\/drawing\d+\.xml$/i),
		pivot_tables: countPaths(paths, /^xl\/pivotTables\/pivotTable\d+\.xml$/i),
		pivot_caches: countPaths(paths, /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i),
		comments: countPaths(paths, /^xl\/comments\d+\.xml$/i),
		threaded_comments: countPaths(paths, /^xl\/threadedComments\//i),
		media: countPaths(paths, /^xl\/media\//i),
		slicers: countPaths(paths, /^xl\/slicers\//i),
		slicer_caches: countPaths(paths, /^xl\/slicerCaches\//i),
		timelines: countPaths(paths, /^xl\/timelines\//i),
		timeline_caches: countPaths(paths, /^xl\/timelineCaches\//i),
		sparklines: countXmlTags(worksheetXml, 'sparklineGroup'),
		macros: countPaths(paths, /^xl\/vbaProject\.bin$/i),
		active_content:
			countPaths(paths, /^xl\/activeX\//i) +
			countPaths(paths, /^xl\/ctrlProps\//i) +
			countPaths(paths, /^xl\/embeddings\//i),
		custom_xml: countPaths(paths, /^customXml\//i),
		external_links: countPaths(paths, /^xl\/externalLinks\//i),
		connections:
			countPaths(paths, /^xl\/connections\.xml$/i) + countPaths(paths, /^xl\/queryTables\//i),
		calc_chain: countPaths(paths, /^xl\/calcChain\.xml$/i),
	}
	const relationships = parseRelationships(files)
	const pivotCacheRelationships = relationships.filter(
		(relationship) => relationship.type === PIVOT_CACHE_DEFINITION_REL,
	)
	const pivotCacheRecordRelationships = relationships.filter(
		(relationship) => relationship.type === PIVOT_CACHE_RECORDS_REL,
	)
	const analytics: OoxmlAnalyticsProbe = {
		relationships,
		workbookPivotCaches: parseWorkbookPivotCaches(workbookXml),
		pivotTableRelationships: relationships.filter(
			(relationship) => relationship.type === PIVOT_TABLE_REL,
		),
		pivotCacheRelationships,
		pivotCacheRecordRelationships,
		slicerCacheRelationships: relationships.filter(
			(relationship) => relationship.type === SLICER_CACHE_REL,
		),
		pivotTables: paths
			.filter((path) => /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(path))
			.map((path) => parsePivotTable(path, decodePart(files, path))),
		pivotCaches: paths
			.filter((path) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(path))
			.map((path) =>
				parsePivotCache(
					path,
					decodePart(files, path),
					pivotCacheRelationships,
					pivotCacheRecordRelationships,
				),
			),
		pivotCacheRecords: paths.filter((path) =>
			/^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/i.test(path),
		),
		slicerCaches: paths
			.filter((path) => /^xl\/slicerCaches\/slicerCache\d+\.xml$/i.test(path))
			.map((path) => parseLinkedCache(path, decodePart(files, path))),
		slicers: paths
			.filter((path) => /^xl\/slicers\/slicer\d+\.xml$/i.test(path))
			.flatMap((path) => parseLinkedUi(path, decodePart(files, path), 'slicer')),
		timelineCaches: paths
			.filter((path) => /^xl\/timelineCaches\/timelineCache\d+\.xml$/i.test(path))
			.map((path) => parseLinkedCache(path, decodePart(files, path))),
		timelines: paths
			.filter((path) => /^xl\/timelines\/timeline\d+\.xml$/i.test(path))
			.flatMap((path) => parseLinkedUi(path, decodePart(files, path), 'timeline')),
	}
	const features = {
		macros: counts.macros > 0,
		charts:
			counts.charts > 0 ||
			countPaths(paths, /^xl\/chartsheets\//i) > 0 ||
			contentTypesXml.includes('chartsheet+xml'),
		pivot_tables: counts.pivot_tables > 0,
		tables: counts.tables > 0,
		drawings: counts.drawings > 0,
		comments: counts.comments > 0,
		threaded_comments: counts.threaded_comments > 0,
		conditional_formatting: hasAnyXmlTag(worksheetXml, 'conditionalFormatting'),
		data_validations: hasAnyXmlTag(worksheetXml, 'dataValidation'),
		merged_cells: hasAnyXmlTag(worksheetXml, 'mergeCell'),
		hyperlinks: hasAnyXmlTag(worksheetXml, 'hyperlink'),
		defined_names: hasXmlTag(workbookXml, 'definedName'),
		external_links: counts.external_links > 0,
		connections: counts.connections > 0,
		slicers: counts.slicers > 0 || counts.slicer_caches > 0,
		timelines: counts.timelines > 0 || counts.timeline_caches > 0,
		sparklines: counts.sparklines > 0,
		images_or_media: counts.media > 0,
		custom_xml: counts.custom_xml > 0,
		calc_chain: counts.calc_chain > 0,
		active_content: counts.active_content > 0,
	}
	return { paths, counts, features, analytics }
}

function decodePart(files: Record<string, Uint8Array>, path: string): string {
	const part = files[path]
	return part ? decoder.decode(part) : ''
}

function countPaths(paths: readonly string[], pattern: RegExp): number {
	return paths.filter((path) => pattern.test(path)).length
}

function countXmlTags(xmlParts: readonly string[], tag: string): number {
	const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${tag}(?:\\s|/|>)`, 'gi')
	return xmlParts.reduce((sum, xml) => sum + (xml.match(pattern)?.length ?? 0), 0)
}

function hasAnyXmlTag(xmlParts: readonly string[], tag: string): boolean {
	return xmlParts.some((xml) => hasXmlTag(xml, tag))
}

function hasXmlTag(xml: string, tag: string): boolean {
	return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${tag}(?:\\s|/|>)`, 'i').test(xml)
}

function parseRelationships(files: Record<string, Uint8Array>): OoxmlRelationshipProbe[] {
	const relationships: OoxmlRelationshipProbe[] = []
	for (const path of Object.keys(files)) {
		if (!path.endsWith('.rels')) continue
		const sourcePartPath = sourcePartFromRelsPath(path)
		if (!sourcePartPath) continue
		const xml = decodePart(files, path)
		for (const attrs of matchTagAttributes(xml, 'Relationship')) {
			const id = readAttribute(attrs, 'Id')
			const rawType = readAttribute(attrs, 'Type')
			const target = readAttribute(attrs, 'Target')
			const type = rawType ? normalizeRelationshipType(rawType) : undefined
			if (!id || !type || !target) continue
			const targetMode = readAttribute(attrs, 'TargetMode')
			const parsed: {
				sourcePartPath: string
				relsPartPath: string
				id: string
				type: string
				target: string
				targetMode?: string
				targetPartPath?: string
			} = { sourcePartPath, relsPartPath: path, id, type, target }
			if (targetMode) parsed.targetMode = targetMode
			if (targetMode?.toLowerCase() !== 'external') {
				parsed.targetPartPath = normalizeRelationshipTarget(sourcePartPath, target)
			}
			relationships.push(parsed)
		}
	}
	return relationships.sort((left, right) =>
		`${left.relsPartPath}:${left.id}`.localeCompare(`${right.relsPartPath}:${right.id}`),
	)
}

function sourcePartFromRelsPath(path: string): string | null {
	if (path === '_rels/.rels') return ''
	const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(path)
	if (!match) return null
	return match[1] ? `${match[1]}/${match[2]}` : match[2]
}

function normalizeRelationshipTarget(sourcePartPath: string, target: string): string {
	const normalizedTarget = target.replaceAll('\\', '/')
	if (normalizedTarget.startsWith('/')) return normalizePath(normalizedTarget.slice(1))
	const sourceDir = sourcePartPath.includes('/')
		? sourcePartPath.slice(0, sourcePartPath.lastIndexOf('/'))
		: ''
	return normalizePath(sourceDir ? `${sourceDir}/${normalizedTarget}` : normalizedTarget)
}

function normalizePath(path: string): string {
	const parts: string[] = []
	for (const part of path.split('/')) {
		if (!part || part === '.') continue
		if (part === '..') parts.pop()
		else parts.push(part)
	}
	return parts.join('/')
}

function parseWorkbookPivotCaches(xml: string): OoxmlWorkbookPivotCacheProbe[] {
	return matchTagAttributes(xml, 'pivotCache')
		.map((attrs) => {
			const cacheId = readNumberAttribute(attrs, 'cacheId')
			const relId = readAttribute(attrs, 'r:id') ?? readAttribute(attrs, 'id')
			if (cacheId === undefined || !relId) return null
			return { cacheId, relId }
		})
		.filter((entry): entry is OoxmlWorkbookPivotCacheProbe => entry !== null)
}

function parsePivotTable(partPath: string, xml: string): OoxmlPivotTableProbe {
	const attrs = matchTagAttributes(xml, 'pivotTableDefinition')[0] ?? ''
	const locationAttrs = matchTagAttributes(xml, 'location')[0] ?? ''
	const parsed: {
		partPath: string
		name?: string
		cacheId?: number
		locationRef?: string
	} = { partPath }
	setString(parsed, 'name', readAttribute(attrs, 'name'))
	setNumber(parsed, 'cacheId', readNumberAttribute(attrs, 'cacheId'))
	setString(parsed, 'locationRef', readAttribute(locationAttrs, 'ref'))
	return parsed
}

function parsePivotCache(
	partPath: string,
	xml: string,
	pivotCacheRelationships: readonly OoxmlRelationshipProbe[],
	pivotCacheRecordRelationships: readonly OoxmlRelationshipProbe[],
): OoxmlPivotCacheProbe {
	const attrs = matchTagAttributes(xml, 'pivotCacheDefinition')[0] ?? ''
	const rel = pivotCacheRelationships.find(
		(relationship) => relationship.targetPartPath === partPath,
	)
	const recordRelId = readAttribute(attrs, 'r:id') ?? readAttribute(attrs, 'id')
	const recordRel = recordRelId
		? pivotCacheRecordRelationships.find(
				(relationship) =>
					relationship.sourcePartPath === partPath && relationship.id === recordRelId,
			)
		: undefined
	const parsed: {
		partPath: string
		cacheId?: number
		relId?: string
		recordsPartPath?: string
	} = { partPath }
	const workbookPivotCache = rel?.id
	setString(parsed, 'relId', workbookPivotCache)
	setString(parsed, 'recordsPartPath', recordRel?.targetPartPath)
	return parsed
}

function parseLinkedCache(partPath: string, xml: string): OoxmlLinkedCacheProbe {
	const attrs =
		matchTagAttributes(xml, 'slicerCacheDefinition')[0] ??
		matchTagAttributes(xml, 'timelineCacheDefinition')[0] ??
		''
	const tabularAttrs = matchTagAttributes(xml, 'tabular')[0] ?? ''
	const parsed: {
		partPath: string
		name?: string
		pivotCacheId?: number
		pivotTableNames: readonly string[]
	} = {
		partPath,
		pivotTableNames: matchTagAttributes(xml, 'pivotTable')
			.map((pivotTableAttrs) => readAttribute(pivotTableAttrs, 'name'))
			.filter((name): name is string => Boolean(name)),
	}
	setString(parsed, 'name', readAttribute(attrs, 'name'))
	setNumber(parsed, 'pivotCacheId', readNumberAttribute(tabularAttrs, 'pivotCacheId'))
	return parsed
}

function parseLinkedUi(
	partPath: string,
	xml: string,
	tag: 'slicer' | 'timeline',
): OoxmlLinkedUiProbe[] {
	return matchTagAttributes(xml, tag).map((attrs) => {
		const parsed: {
			partPath: string
			name?: string
			cacheName?: string
		} = { partPath }
		setString(parsed, 'name', readAttribute(attrs, 'name'))
		setString(parsed, 'cacheName', readAttribute(attrs, 'cache'))
		return parsed
	})
}

function matchTagAttributes(xml: string, tag: string): string[] {
	const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${tag}\\b([^>]*)>`, 'gi')
	return [...xml.matchAll(pattern)].map((match) => match[1] ?? '')
}

function readAttribute(attrs: string, name: string): string | undefined {
	const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
	const match = new RegExp(`(?:^|\\s)${escapedName}=(["'])(.*?)\\1`, 'i').exec(attrs)
	return match?.[2] ? decodeXmlAttribute(match[2]) : undefined
}

function readNumberAttribute(attrs: string, name: string): number | undefined {
	const value = readAttribute(attrs, name)
	if (value === undefined) return undefined
	const parsed = Number(value)
	return Number.isFinite(parsed) ? parsed : undefined
}

function normalizeRelationshipType(type: string): string {
	if (!type.startsWith(STRICT_REL_PREFIX)) return type
	const suffix = type.slice(STRICT_REL_PREFIX.length)
	if (suffix === 'sheetMetadata') return type
	if (suffix === 'extendedProperties') return `${TRANSITIONAL_REL_PREFIX}extended-properties`
	return `${TRANSITIONAL_REL_PREFIX}${suffix}`
}

function decodeXmlAttribute(value: string): string {
	return value
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&')
}

function setString<T extends Record<string, unknown>>(
	target: T,
	key: keyof T,
	value: string | undefined,
) {
	if (value !== undefined) target[key] = value as T[keyof T]
}

function setNumber<T extends Record<string, unknown>>(
	target: T,
	key: keyof T,
	value: number | undefined,
) {
	if (value !== undefined) target[key] = value as T[keyof T]
}
