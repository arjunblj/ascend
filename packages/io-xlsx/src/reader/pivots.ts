import type { PivotCacheInfo, PivotTableInfo, SlicerCacheInfo, SlicerInfo } from '@ascend/core'
import { attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import type { Relationship } from './relationships.ts'
import { resolvePath } from './relationships.ts'

export interface PivotCacheEntry {
	readonly cacheId: number
	readonly relId: string
}

export function parsePivotCacheDefinitionXml(
	xml: string,
	partPath: string,
	cacheId: number | undefined,
	relId: string | undefined,
	relationships: readonly Relationship[],
): PivotCacheInfo | null {
	const doc = parseXml(xml)
	const root = doc.pivotCacheDefinition as XmlNode | undefined
	if (!root) return null
	const worksheetSource = ((root.cacheSource as XmlNode | undefined)?.worksheetSource ??
		undefined) as XmlNode | undefined
	const recordCount = numAttr(root, 'recordCount')
	const recordsRelId = attr(root, 'r:id') ?? attr(root, 'id')
	const recordsRel = recordsRelId
		? relationships.find((relationship) => relationship.id === recordsRelId)
		: undefined
	const parsed: {
		partPath: string
		cacheId?: number
		relId?: string
		recordCount?: number
		refreshedVersion?: number
		minRefreshableVersion?: number
		createdVersion?: number
		refreshedBy?: string
		refreshedDate?: number
		refreshOnLoad?: boolean
		enableRefresh?: boolean
		invalid?: boolean
		saveData?: boolean
		optimizeMemory?: boolean
		sourceSheet?: string
		sourceRef?: string
		recordsPartPath?: string
	} = { partPath }
	if (cacheId !== undefined) parsed.cacheId = cacheId
	if (relId) parsed.relId = relId
	if (recordCount !== undefined) parsed.recordCount = recordCount
	setNumberIfDefined(parsed, 'refreshedVersion', numAttr(root, 'refreshedVersion'))
	setNumberIfDefined(parsed, 'minRefreshableVersion', numAttr(root, 'minRefreshableVersion'))
	setNumberIfDefined(parsed, 'createdVersion', numAttr(root, 'createdVersion'))
	setStringIfDefined(parsed, 'refreshedBy', attr(root, 'refreshedBy'))
	setNumberIfDefined(parsed, 'refreshedDate', numAttr(root, 'refreshedDate'))
	setBoolIfDefined(parsed, 'refreshOnLoad', boolAttr(root, 'refreshOnLoad'))
	setBoolIfDefined(parsed, 'enableRefresh', boolAttr(root, 'enableRefresh'))
	setBoolIfDefined(parsed, 'invalid', boolAttr(root, 'invalid'))
	setBoolIfDefined(parsed, 'saveData', boolAttr(root, 'saveData'))
	setBoolIfDefined(parsed, 'optimizeMemory', boolAttr(root, 'optimizeMemory'))
	const sourceSheet = worksheetSource ? attr(worksheetSource, 'sheet') : undefined
	if (sourceSheet) parsed.sourceSheet = sourceSheet
	const sourceRef = worksheetSource ? attr(worksheetSource, 'ref') : undefined
	if (sourceRef) parsed.sourceRef = sourceRef
	if (recordsRel) parsed.recordsPartPath = resolvePath(partPath, recordsRel.target)
	return parsed as PivotCacheInfo
}

function setNumberIfDefined(
	target: Record<string, unknown>,
	key: string,
	value: number | undefined,
): void {
	if (value !== undefined) target[key] = value
}

function setStringIfDefined(
	target: Record<string, unknown>,
	key: string,
	value: string | undefined,
): void {
	if (value !== undefined) target[key] = value
}

function setBoolIfDefined(
	target: Record<string, unknown>,
	key: string,
	value: boolean | undefined,
): void {
	if (value !== undefined) target[key] = value
}

export function parsePivotTableXml(
	xml: string,
	partPath: string,
	sheetName: string,
): PivotTableInfo | null {
	const doc = parseXml(xml)
	const root = doc.pivotTableDefinition as XmlNode | undefined
	if (!root) return null
	const location = root.location as XmlNode | undefined
	const parsed: {
		partPath: string
		sheetName: string
		name?: string
		cacheId?: number
		locationRef?: string
	} = {
		partPath,
		sheetName,
	}
	const name = attr(root, 'name')
	if (name) parsed.name = name
	const cacheId = numAttr(root, 'cacheId')
	if (cacheId !== undefined) parsed.cacheId = cacheId
	const locationRef = location ? attr(location, 'ref') : undefined
	if (locationRef) parsed.locationRef = locationRef
	return parsed as PivotTableInfo
}

export function parseSlicerCacheXml(xml: string, partPath: string): SlicerCacheInfo | null {
	const doc = parseXml(xml)
	const root = (doc.slicerCacheDefinition ??
		doc['x14:slicerCacheDefinition'] ??
		doc['s:slicerCacheDefinition']) as XmlNode | undefined
	if (!root) return null
	const pivotTablesNode = root.pivotTables as XmlNode | undefined
	const pivotTableNodes = pivotTablesNode
		? ((pivotTablesNode.pivotTable as XmlNode | XmlNode[] | undefined) ?? [])
		: []
	const dataNode = root.data as XmlNode | undefined
	const tabular = dataNode ? ((dataNode.tabular as XmlNode | undefined) ?? undefined) : undefined
	const parsed: {
		partPath: string
		name?: string
		sourceName?: string
		pivotCacheId?: number
		pivotTableNames: readonly string[]
	} = {
		partPath,
		pivotTableNames: (Array.isArray(pivotTableNodes) ? pivotTableNodes : [pivotTableNodes])
			.filter(Boolean)
			.map((pivotTable) => attr(pivotTable as XmlNode, 'name') ?? '')
			.filter(Boolean),
	}
	const name = attr(root, 'name')
	if (name) parsed.name = name
	const sourceName = attr(root, 'sourceName')
	if (sourceName) parsed.sourceName = sourceName
	const pivotCacheId = tabular ? numAttr(tabular, 'pivotCacheId') : undefined
	if (pivotCacheId !== undefined) parsed.pivotCacheId = pivotCacheId
	return parsed as SlicerCacheInfo
}

export function parseSlicerXml(xml: string, partPath: string): readonly SlicerInfo[] {
	const doc = parseXml(xml)
	const root = (doc.slicers ?? doc['x14:slicers']) as XmlNode | undefined
	if (!root) return []
	const slicerNodes = root.slicer as XmlNode | XmlNode[] | undefined
	const nodes = Array.isArray(slicerNodes) ? slicerNodes : slicerNodes ? [slicerNodes] : []
	return nodes.map((node) => {
		const parsed: {
			partPath: string
			name?: string
			cacheName?: string
			caption?: string
		} = { partPath }
		const name = attr(node, 'name')
		if (name) parsed.name = name
		const cacheName = attr(node, 'cache')
		if (cacheName) parsed.cacheName = cacheName
		const caption = attr(node, 'caption')
		if (caption) parsed.caption = caption
		return parsed as SlicerInfo
	})
}
