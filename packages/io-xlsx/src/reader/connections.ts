import type { WorkbookConnectionPartInfo, WorkbookConnectionPartKind } from '@ascend/core'
import type { PreservationCapsule } from '../preserve.ts'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'

interface ParsedConnectionAttrs {
	name?: string
	connectionId?: number
	refreshOnLoad?: boolean
	saveData?: boolean
	refreshedVersion?: number
}

export function parseConnectionPartInfos(
	capsule: PreservationCapsule,
	xml: string | undefined,
): WorkbookConnectionPartInfo[] {
	const kind = classifyConnectionPart(capsule)
	if (!kind) return []
	const base = {
		kind,
		partPath: capsule.partPath,
		contentType: capsule.contentType,
		...(capsule.relType ? { relType: capsule.relType } : {}),
		...(capsule.anchor.kind === 'sheet' ? { sheetName: capsule.anchor.sheetName } : {}),
		relationshipCount: capsule.relationships.length,
	}
	if (!xml) return [base]
	if (kind === 'queryTable') {
		const root = parseXml(xml).queryTable as XmlNode | undefined
		return [{ ...base, ...readConnectionAttrs(root, { queryTable: true }) }]
	}
	if (kind === 'connection') {
		const doc = parseXml(xml)
		const root = doc.connections as XmlNode | undefined
		const connections = asArray(root?.connection as XmlNode | XmlNode[] | undefined)
		if (connections.length === 0) return [base]
		return connections.map((connection) => ({
			...base,
			...readConnectionAttrs(connection),
		}))
	}
	return [base]
}

function classifyConnectionPart(capsule: PreservationCapsule): WorkbookConnectionPartKind | null {
	const path = capsule.partPath.toLowerCase()
	const contentType = capsule.contentType.toLowerCase()
	const relType = capsule.relType?.toLowerCase() ?? ''
	if (
		path.includes('/querytables/') ||
		contentType.includes('querytable+xml') ||
		relType.includes('/querytable')
	) {
		return 'queryTable'
	}
	if (
		path.includes('/customdata/') ||
		contentType.includes('customdata') ||
		contentType.includes('mashup') ||
		relType.includes('powerquery') ||
		relType.includes('mashup')
	) {
		return 'powerQueryMashup'
	}
	if (
		path.endsWith('/connections.xml') ||
		contentType.includes('connections+xml') ||
		relType.includes('/connections')
	) {
		return 'connection'
	}
	return null
}

function readConnectionAttrs(
	node: XmlNode | undefined,
	options: { queryTable?: boolean } = {},
): Partial<WorkbookConnectionPartInfo> {
	if (!node) return {}
	const parsed: ParsedConnectionAttrs = {}
	const name = attr(node, 'name')
	if (name) parsed.name = name
	const id = numAttr(node, options.queryTable ? 'connectionId' : 'id')
	if (id !== undefined) parsed.connectionId = id
	const refreshOnLoad = boolAttr(node, 'refreshOnLoad')
	if (refreshOnLoad !== undefined) parsed.refreshOnLoad = refreshOnLoad
	const saveData = boolAttr(node, 'saveData')
	if (saveData !== undefined) parsed.saveData = saveData
	const removeDataOnSave = boolAttr(node, 'removeDataOnSave')
	if (options.queryTable && saveData === undefined && removeDataOnSave !== undefined) {
		parsed.saveData = !removeDataOnSave
	}
	const refreshedVersion = numAttr(node, 'refreshedVersion')
	if (refreshedVersion !== undefined) parsed.refreshedVersion = refreshedVersion
	return parsed
}
