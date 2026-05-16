import type { WorkbookConnectionPartInfo, WorkbookConnectionPartKind } from '@ascend/core'
import type { PreservationCapsule } from '../preserve.ts'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'

interface ParsedConnectionAttrs {
	name?: string
	connectionId?: number
	connectionType?: number
	description?: string
	deleted?: boolean
	backgroundRefresh?: boolean
	keepAlive?: boolean
	refreshInterval?: number
	refreshOnLoad?: boolean
	saveData?: boolean
	savePassword?: boolean
	refreshedVersion?: number
	refreshedDateIso?: string
	minRefreshableVersion?: number
	credentials?: string
	singleSignOnId?: string
	sourceFile?: string
	odcFile?: string
	onlyUseConnectionFile?: boolean
	command?: string
	hasConnectionString?: boolean
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
	const connectionType = numAttr(node, 'type')
	if (connectionType !== undefined) parsed.connectionType = connectionType
	const description = attr(node, 'description')
	if (description) parsed.description = description
	const deleted = boolAttr(node, 'deleted')
	if (deleted !== undefined) parsed.deleted = deleted
	const backgroundRefresh = boolAttr(node, 'background')
	if (backgroundRefresh !== undefined) parsed.backgroundRefresh = backgroundRefresh
	const keepAlive = boolAttr(node, 'keepAlive')
	if (keepAlive !== undefined) parsed.keepAlive = keepAlive
	const refreshInterval = numAttr(node, 'interval')
	if (refreshInterval !== undefined) parsed.refreshInterval = refreshInterval
	const refreshOnLoad = boolAttr(node, 'refreshOnLoad')
	if (refreshOnLoad !== undefined) parsed.refreshOnLoad = refreshOnLoad
	const saveData = boolAttr(node, 'saveData')
	if (saveData !== undefined) parsed.saveData = saveData
	const savePassword = boolAttr(node, 'savePassword')
	if (savePassword !== undefined) parsed.savePassword = savePassword
	const removeDataOnSave = boolAttr(node, 'removeDataOnSave')
	if (options.queryTable && saveData === undefined && removeDataOnSave !== undefined) {
		parsed.saveData = !removeDataOnSave
	}
	const refreshedVersion = numAttr(node, 'refreshedVersion')
	if (refreshedVersion !== undefined) parsed.refreshedVersion = refreshedVersion
	const refreshedDateIso = attr(node, 'refreshedDateIso')
	if (refreshedDateIso) parsed.refreshedDateIso = refreshedDateIso
	const minRefreshableVersion = numAttr(node, 'minRefreshableVersion')
	if (minRefreshableVersion !== undefined) parsed.minRefreshableVersion = minRefreshableVersion
	const credentials = attr(node, 'credentials')
	if (credentials) parsed.credentials = credentials
	const singleSignOnId = attr(node, 'singleSignOnId')
	if (singleSignOnId) parsed.singleSignOnId = singleSignOnId
	const textPr = node.textPr as XmlNode | undefined
	const dbPr = node.dbPr as XmlNode | undefined
	const sourceFile = attr(node, 'sourceFile') ?? (textPr ? attr(textPr, 'sourceFile') : undefined)
	if (sourceFile) parsed.sourceFile = sourceFile
	const odcFile = attr(node, 'odcFile')
	if (odcFile) parsed.odcFile = odcFile
	const onlyUseConnectionFile = boolAttr(node, 'onlyUseConnectionFile')
	if (onlyUseConnectionFile !== undefined) parsed.onlyUseConnectionFile = onlyUseConnectionFile
	const command = dbPr ? attr(dbPr, 'command') : undefined
	if (command) parsed.command = command
	const connectionString = dbPr ? attr(dbPr, 'connection') : undefined
	if (connectionString) parsed.hasConnectionString = true
	return parsed
}
