import type {
	PivotCacheFieldGroupInfo,
	PivotCacheFieldInfo,
	PivotCacheInfo,
	PivotCacheSharedItemInfo,
	PivotCacheSharedItemsInfo,
	PivotDataFieldInfo,
	PivotFieldInfo,
	PivotFieldItemInfo,
	PivotFieldReference,
	PivotTableInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TimelineCacheInfo,
	TimelineInfo,
	TimelineRangeInfo,
	TimelineStateInfo,
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import type { Relationship } from './relationships.ts'
import { resolvePath } from './relationships.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

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
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
	const root = childNode(doc, 'pivotCacheDefinition')
	if (!root) return null
	const worksheetSource = childNode(childNode(root, 'cacheSource'), 'worksheetSource')
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
		fields: readonly PivotCacheFieldInfo[]
	} = { partPath, fields: parseCacheFields(root) }
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
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
	const root = childNode(doc, 'pivotTableDefinition')
	if (!root) return null
	const location = childNode(root, 'location')
	const parsed: {
		partPath: string
		sheetName: string
		name?: string
		cacheId?: number
		locationRef?: string
		fields: readonly PivotFieldInfo[]
		rowFields: readonly PivotFieldReference[]
		columnFields: readonly PivotFieldReference[]
		pageFields: readonly PivotFieldReference[]
		dataFields: readonly PivotDataFieldInfo[]
	} = {
		partPath,
		sheetName,
		fields: parsePivotFields(root),
		rowFields: parseFieldReferences(childNode(root, 'rowFields'), 'field', 'x'),
		columnFields: parseFieldReferences(childNode(root, 'colFields'), 'field', 'x'),
		pageFields: parseFieldReferences(childNode(root, 'pageFields'), 'pageField', 'fld'),
		dataFields: parseDataFields(root),
	}
	const name = attr(root, 'name')
	if (name) parsed.name = name
	const cacheId = numAttr(root, 'cacheId')
	if (cacheId !== undefined) parsed.cacheId = cacheId
	const locationRef = location ? attr(location, 'ref') : undefined
	if (locationRef) parsed.locationRef = locationRef
	return parsed as PivotTableInfo
}

function parseCacheFields(root: XmlNode): PivotCacheFieldInfo[] {
	const cacheFields = childNode(root, 'cacheFields')
	return childNodes(cacheFields, 'cacheField').map((node, index) => {
		const parsed: {
			index: number
			name?: string
			databaseField?: boolean
			numFmtId?: number
			formula?: string
			sharedItemsInfo?: PivotCacheSharedItemsInfo
			sharedItems?: readonly PivotCacheSharedItemInfo[]
			fieldGroup?: PivotCacheFieldGroupInfo
		} = { index }
		setStringIfDefined(parsed, 'name', attr(node, 'name'))
		setBoolIfDefined(parsed, 'databaseField', boolAttr(node, 'databaseField'))
		setNumberIfDefined(parsed, 'numFmtId', numAttr(node, 'numFmtId'))
		setStringIfDefined(parsed, 'formula', attr(node, 'formula'))
		const sharedItemsNode = childNode(node, 'sharedItems')
		const sharedItemsInfo = parseCacheSharedItemsInfo(sharedItemsNode)
		if (sharedItemsInfo) parsed.sharedItemsInfo = sharedItemsInfo
		const sharedItems = parseCacheSharedItems(sharedItemsNode)
		if (sharedItems.length > 0) parsed.sharedItems = sharedItems
		const fieldGroup = parseCacheFieldGroup(childNode(node, 'fieldGroup'))
		if (fieldGroup) parsed.fieldGroup = fieldGroup
		return parsed
	})
}

function parseCacheSharedItemsInfo(
	sharedItems: XmlNode | undefined,
): PivotCacheSharedItemsInfo | undefined {
	if (!sharedItems) return undefined
	const parsed: {
		count?: number
		containsBlank?: boolean
		containsDate?: boolean
		containsNonDate?: boolean
		containsNumber?: boolean
		containsInteger?: boolean
		containsString?: boolean
		containsMixedTypes?: boolean
		containsSemiMixedTypes?: boolean
		minValue?: number
		maxValue?: number
		minDate?: string
		maxDate?: string
	} = {}
	setNumberIfDefined(parsed, 'count', numAttr(sharedItems, 'count'))
	setBoolIfDefined(parsed, 'containsBlank', boolAttr(sharedItems, 'containsBlank'))
	setBoolIfDefined(parsed, 'containsDate', boolAttr(sharedItems, 'containsDate'))
	setBoolIfDefined(parsed, 'containsNonDate', boolAttr(sharedItems, 'containsNonDate'))
	setBoolIfDefined(parsed, 'containsNumber', boolAttr(sharedItems, 'containsNumber'))
	setBoolIfDefined(parsed, 'containsInteger', boolAttr(sharedItems, 'containsInteger'))
	setBoolIfDefined(parsed, 'containsString', boolAttr(sharedItems, 'containsString'))
	setBoolIfDefined(parsed, 'containsMixedTypes', boolAttr(sharedItems, 'containsMixedTypes'))
	setBoolIfDefined(
		parsed,
		'containsSemiMixedTypes',
		boolAttr(sharedItems, 'containsSemiMixedTypes'),
	)
	setNumberIfDefined(parsed, 'minValue', numAttr(sharedItems, 'minValue'))
	setNumberIfDefined(parsed, 'maxValue', numAttr(sharedItems, 'maxValue'))
	setStringIfDefined(parsed, 'minDate', attr(sharedItems, 'minDate'))
	setStringIfDefined(parsed, 'maxDate', attr(sharedItems, 'maxDate'))
	return Object.keys(parsed).length > 0 ? (parsed as PivotCacheSharedItemsInfo) : undefined
}

function parseCacheSharedItems(sharedItems: XmlNode | undefined): PivotCacheSharedItemInfo[] {
	if (!sharedItems) return []
	const items: PivotCacheSharedItemInfo[] = []
	for (const [key, value] of Object.entries(sharedItems)) {
		if (key.startsWith('@_')) continue
		const kind = sharedItemKind(localPart(key))
		if (!kind) continue
		for (const node of asArray(value as XmlNode | XmlNode[])) {
			const parsed: {
				index: number
				kind: PivotCacheSharedItemInfo['kind']
				value?: string
			} = { index: items.length, kind }
			setStringIfDefined(parsed, 'value', attr(node, 'v'))
			items.push(parsed)
		}
	}
	return items
}

function parseCacheFieldGroup(
	fieldGroup: XmlNode | undefined,
): PivotCacheFieldGroupInfo | undefined {
	if (!fieldGroup) return undefined
	const parsed: {
		base?: number
		parent?: number
		discreteItems?: readonly { readonly index: number; readonly value?: number }[]
		groupItems?: readonly PivotCacheSharedItemInfo[]
	} = {}
	setNumberIfDefined(parsed, 'base', numAttr(fieldGroup, 'base'))
	setNumberIfDefined(parsed, 'parent', numAttr(fieldGroup, 'par'))
	const discreteItems = parseDiscreteGroupItems(childNode(fieldGroup, 'discretePr'))
	if (discreteItems.length > 0) parsed.discreteItems = discreteItems
	const groupItems = parseCacheSharedItems(childNode(fieldGroup, 'groupItems'))
	if (groupItems.length > 0) parsed.groupItems = groupItems
	return Object.keys(parsed).length > 0 ? (parsed as PivotCacheFieldGroupInfo) : undefined
}

function parseDiscreteGroupItems(
	discretePr: XmlNode | undefined,
): readonly { readonly index: number; readonly value?: number }[] {
	return childNodes(discretePr, 'x').map((node, index) => {
		const parsed: { index: number; value?: number } = { index }
		setNumberIfDefined(parsed, 'value', numAttr(node, 'v'))
		return parsed
	})
}

function sharedItemKind(localName: string): PivotCacheSharedItemInfo['kind'] | null {
	switch (localName) {
		case 's':
			return 'string'
		case 'n':
			return 'number'
		case 'd':
			return 'date'
		case 'b':
			return 'boolean'
		case 'e':
			return 'error'
		case 'm':
			return 'missing'
		default:
			return null
	}
}

function parsePivotFields(root: XmlNode): PivotFieldInfo[] {
	const pivotFields = childNode(root, 'pivotFields')
	return childNodes(pivotFields, 'pivotField').map((node, index) => {
		const parsed: {
			index: number
			axis?: string
			name?: string
			numFmtId?: number
			hidden?: boolean
			dataField?: boolean
			defaultSubtotal?: boolean
			showAll?: boolean
			multipleItemSelectionAllowed?: boolean
			items?: readonly PivotFieldItemInfo[]
		} = { index }
		setStringIfDefined(parsed, 'axis', attr(node, 'axis'))
		setStringIfDefined(parsed, 'name', attr(node, 'name'))
		setNumberIfDefined(parsed, 'numFmtId', numAttr(node, 'numFmtId'))
		setBoolIfDefined(parsed, 'hidden', boolAttr(node, 'hidden'))
		setBoolIfDefined(parsed, 'dataField', boolAttr(node, 'dataField'))
		setBoolIfDefined(parsed, 'defaultSubtotal', boolAttr(node, 'defaultSubtotal'))
		setBoolIfDefined(parsed, 'showAll', boolAttr(node, 'showAll'))
		setBoolIfDefined(
			parsed,
			'multipleItemSelectionAllowed',
			boolAttr(node, 'multipleItemSelectionAllowed'),
		)
		const items = parsePivotFieldItems(node)
		if (items.length > 0) parsed.items = items
		return parsed
	})
}

function parsePivotFieldItems(pivotField: XmlNode): PivotFieldItemInfo[] {
	const items = childNode(pivotField, 'items')
	return childNodes(items, 'item').map((node, index) => {
		const parsed: {
			index: number
			cacheIndex?: number
			itemType?: string
			caption?: string
			hidden?: boolean
			manualFilter?: boolean
			showDetails?: boolean
			calculated?: boolean
			missing?: boolean
			childItems?: boolean
			expanded?: boolean
			drillAcrossAttributes?: boolean
		} = { index }
		setNumberIfDefined(parsed, 'cacheIndex', numAttr(node, 'x'))
		setStringIfDefined(parsed, 'itemType', attr(node, 't'))
		setStringIfDefined(parsed, 'caption', attr(node, 'n'))
		setBoolIfDefined(parsed, 'hidden', boolAttr(node, 'h'))
		setBoolIfDefined(parsed, 'manualFilter', boolAttr(node, 's'))
		setBoolIfDefined(parsed, 'showDetails', boolAttr(node, 'sd'))
		setBoolIfDefined(parsed, 'calculated', boolAttr(node, 'f'))
		setBoolIfDefined(parsed, 'missing', boolAttr(node, 'm'))
		setBoolIfDefined(parsed, 'childItems', boolAttr(node, 'c'))
		setBoolIfDefined(parsed, 'expanded', boolAttr(node, 'd'))
		setBoolIfDefined(parsed, 'drillAcrossAttributes', boolAttr(node, 'e'))
		return parsed
	})
}

function parseFieldReferences(
	parent: XmlNode | undefined,
	childName: string,
	attrName: string,
): PivotFieldReference[] {
	if (!parent) return []
	return childNodes(parent, childName)
		.map((node) => {
			const index = numAttr(node, attrName)
			if (index === undefined) return null
			const parsed: {
				index: number
				name?: string
				item?: number
				hierarchy?: number
				caption?: string
			} = { index }
			setStringIfDefined(parsed, 'name', attr(node, 'name'))
			setNumberIfDefined(parsed, 'item', numAttr(node, 'item'))
			setNumberIfDefined(parsed, 'hierarchy', numAttr(node, 'hier'))
			setStringIfDefined(parsed, 'caption', attr(node, 'cap'))
			return parsed
		})
		.filter((entry): entry is PivotFieldReference => entry !== null)
}

function parseDataFields(root: XmlNode): PivotDataFieldInfo[] {
	const dataFields = childNode(root, 'dataFields')
	return childNodes(dataFields, 'dataField')
		.map((node) => {
			const fieldIndex = numAttr(node, 'fld')
			if (fieldIndex === undefined) return null
			const parsed: {
				fieldIndex: number
				name?: string
				subtotal?: string
				numFmtId?: number
			} = { fieldIndex }
			setStringIfDefined(parsed, 'name', attr(node, 'name'))
			setStringIfDefined(parsed, 'subtotal', attr(node, 'subtotal'))
			setNumberIfDefined(parsed, 'numFmtId', numAttr(node, 'numFmtId'))
			return parsed
		})
		.filter((entry): entry is PivotDataFieldInfo => entry !== null)
}

function childNode(node: XmlNode | undefined, localName: string): XmlNode | undefined {
	return childNodes(node, localName)[0]
}

function childNodes(node: XmlNode | undefined, localName: string): XmlNode[] {
	if (!node) return []
	const matches: XmlNode[] = []
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || localPart(key) !== localName) continue
		matches.push(...asArray(value as XmlNode | XmlNode[]))
	}
	return matches
}

function localPart(name: string): string {
	const colon = name.indexOf(':')
	return colon === -1 ? name : name.slice(colon + 1)
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
		items?: readonly {
			readonly index: number
			readonly selected?: boolean
			readonly noData?: boolean
		}[]
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
	const items = tabular ? parseSlicerCacheItems(tabular) : []
	if (items.length > 0) parsed.items = items
	return parsed as SlicerCacheInfo
}

function parseSlicerCacheItems(tabular: XmlNode): {
	readonly index: number
	readonly selected?: boolean
	readonly noData?: boolean
}[] {
	const itemsNode = tabular.items as XmlNode | undefined
	return asArray<XmlNode>(itemsNode?.i as XmlNode | XmlNode[] | undefined)
		.map((node) => {
			const index = numAttr(node, 'x')
			if (index === undefined) return null
			const parsed: { index: number; selected?: boolean; noData?: boolean } = { index }
			setBoolIfDefined(parsed, 'selected', boolAttr(node, 's'))
			setBoolIfDefined(parsed, 'noData', boolAttr(node, 'nd'))
			return parsed
		})
		.filter(
			(entry): entry is { index: number; selected?: boolean; noData?: boolean } => entry !== null,
		)
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

export function parseTimelineCacheXml(xml: string, partPath: string): TimelineCacheInfo | null {
	const doc = parseXml(xml)
	const root = (doc.timelineCacheDefinition ??
		doc['x15:timelineCacheDefinition'] ??
		doc['x14:timelineCacheDefinition']) as XmlNode | undefined
	if (!root) return null
	const pivotTablesNode = root.pivotTables as XmlNode | undefined
	const pivotTableNodes = pivotTablesNode
		? ((pivotTablesNode.pivotTable as XmlNode | XmlNode[] | undefined) ?? [])
		: []
	const dataNode = root.data as XmlNode | undefined
	const tabular = dataNode ? ((dataNode.tabular as XmlNode | undefined) ?? undefined) : undefined
	const stateNode = childNode(root, 'state')
	const parsed: {
		partPath: string
		name?: string
		sourceName?: string
		pivotCacheId?: number
		pivotTableNames: readonly string[]
		state?: TimelineStateInfo
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
	const state = stateNode ? parseTimelineState(stateNode) : undefined
	if (state) parsed.state = state
	return parsed as TimelineCacheInfo
}

function parseTimelineState(state: XmlNode): TimelineStateInfo | undefined {
	const parsed: {
		filterType?: string
		filterId?: number
		filterPivotName?: string
		filterTabId?: number
		lastRefreshVersion?: number
		minimalRefreshVersion?: number
		pivotCacheId?: number
		singleRangeFilterState?: boolean
		selection?: TimelineRangeInfo
		bounds?: TimelineRangeInfo
	} = {}
	setStringIfDefined(parsed, 'filterType', attr(state, 'filterType'))
	setStringIfDefined(parsed, 'filterPivotName', attr(state, 'filterPivotName'))
	setNumberIfDefined(parsed, 'filterId', numAttr(state, 'filterId'))
	setNumberIfDefined(parsed, 'filterTabId', numAttr(state, 'filterTabId'))
	setNumberIfDefined(parsed, 'lastRefreshVersion', numAttr(state, 'lastRefreshVersion'))
	setNumberIfDefined(parsed, 'minimalRefreshVersion', numAttr(state, 'minimalRefreshVersion'))
	setNumberIfDefined(parsed, 'pivotCacheId', numAttr(state, 'pivotCacheId'))
	setBoolIfDefined(parsed, 'singleRangeFilterState', boolAttr(state, 'singleRangeFilterState'))
	const selection = parseTimelineRange(childNode(state, 'selection'))
	if (selection) parsed.selection = selection
	const bounds = parseTimelineRange(childNode(state, 'bounds'))
	if (bounds) parsed.bounds = bounds
	return Object.keys(parsed).length > 0 ? (parsed as TimelineStateInfo) : undefined
}

function parseTimelineRange(node: XmlNode | undefined): TimelineRangeInfo | undefined {
	if (!node) return undefined
	const startDate = attr(node, 'startDate')
	const endDate = attr(node, 'endDate')
	return startDate && endDate ? { startDate, endDate } : undefined
}

export function parseTimelineXml(xml: string, partPath: string): readonly TimelineInfo[] {
	const doc = parseXml(xml)
	const root = (doc.timelines ?? doc['x15:timelines'] ?? doc['x14:timelines']) as
		| XmlNode
		| undefined
	if (!root) return []
	const timelineNodes = root.timeline as XmlNode | XmlNode[] | undefined
	const nodes = Array.isArray(timelineNodes) ? timelineNodes : timelineNodes ? [timelineNodes] : []
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
		return parsed as TimelineInfo
	})
}
