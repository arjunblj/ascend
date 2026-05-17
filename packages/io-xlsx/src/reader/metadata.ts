import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

export interface DynamicArrayMetadataRecord {
	readonly metadataIndex: number
	readonly collapsed?: boolean
}

export interface ParsedMetadataPart {
	readonly dynamicArrayByCellMetadataIndex: ReadonlyMap<number, DynamicArrayMetadataRecord>
}

export function parseMetadataXml(xml: string): ParsedMetadataPart {
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
	const metadata = doc.metadata as XmlNode | undefined
	if (!metadata) {
		return { dynamicArrayByCellMetadataIndex: new Map() }
	}

	const metadataTypes = childNodes(metadata, 'metadataType', childNode(metadata, 'metadataTypes'))
	const dynamicArrayTypeIndex = metadataTypes.findIndex(
		(node) => attr(node, 'name')?.toUpperCase() === 'XLDAPR',
	)
	if (dynamicArrayTypeIndex === -1) {
		return { dynamicArrayByCellMetadataIndex: new Map() }
	}

	const collapsedByValueIndex = new Map<number, boolean>()
	for (const futureMetadata of childNodes(metadata, 'futureMetadata')) {
		if (attr(futureMetadata, 'name')?.toUpperCase() !== 'XLDAPR') continue
		const records = childNodes(futureMetadata, 'bk')
		for (let valueIndex = 0; valueIndex < records.length; valueIndex++) {
			const record = records[valueIndex]
			if (!record) continue
			const properties = findNodeBySuffix(record, 'dynamicArrayProperties')
			if (!properties) continue
			const collapsed = boolAttr(properties, 'fCollapsed')
			if (collapsed !== undefined) collapsedByValueIndex.set(valueIndex, collapsed)
		}
	}

	const dynamicArrayByCellMetadataIndex = new Map<number, DynamicArrayMetadataRecord>()
	const cellMetadataRecords = childNodes(metadata, 'bk', childNode(metadata, 'cellMetadata'))
	for (let index = 0; index < cellMetadataRecords.length; index++) {
		const record = cellMetadataRecords[index]
		if (!record) continue
		const rc = childNode(record, 'rc')
		if (!rc) continue
		const typeIndex = numAttr(rc, 't')
		if (typeIndex !== dynamicArrayTypeIndex + 1) continue
		const valueIndex = numAttr(rc, 'v') ?? 0
		const metadataIndex = index + 1
		const collapsed = collapsedByValueIndex.get(valueIndex)
		dynamicArrayByCellMetadataIndex.set(
			metadataIndex,
			collapsed !== undefined ? { metadataIndex, collapsed } : { metadataIndex },
		)
	}

	return { dynamicArrayByCellMetadataIndex }
}

function childNode(node: XmlNode, tagName: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(node)) {
		if (key === tagName || key.endsWith(`:${tagName}`)) {
			return Array.isArray(value)
				? (value[0] as XmlNode | undefined)
				: (value as XmlNode | undefined)
		}
	}
	return undefined
}

function childNodes(node: XmlNode, tagName: string, scope?: XmlNode): XmlNode[] {
	const target = scope ?? node
	for (const [key, value] of Object.entries(target)) {
		if (key === tagName || key.endsWith(`:${tagName}`)) {
			return asArray<XmlNode>(value as XmlNode | XmlNode[])
		}
	}
	return []
}

function findNodeBySuffix(node: XmlNode, suffix: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(node)) {
		if (key === suffix || key.endsWith(`:${suffix}`)) {
			return Array.isArray(value)
				? (value[0] as XmlNode | undefined)
				: (value as XmlNode | undefined)
		}
		if (value && typeof value === 'object' && !Array.isArray(value)) {
			const nested = findNodeBySuffix(value as XmlNode, suffix)
			if (nested) return nested
		}
		if (Array.isArray(value)) {
			for (const entry of value) {
				if (!entry || typeof entry !== 'object') continue
				const nested = findNodeBySuffix(entry as XmlNode, suffix)
				if (nested) return nested
			}
		}
	}
	return undefined
}
