import type {
	QueryTableFieldInfo,
	WorkbookConnectionPartInfo,
	WorkbookConnectionPartKind,
} from '@ascend/core'
import type { PreservationCapsule } from '../preserve.ts'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'

interface ParsedConnectionAttrs {
	name?: string
	connectionId?: number
	connectionType?: number
	description?: string
	deleted?: boolean
	newConnection?: boolean
	backgroundRefresh?: boolean
	firstBackgroundRefresh?: boolean
	keepAlive?: boolean
	refreshInterval?: number
	refreshOnLoad?: boolean
	reconnectionMethod?: number
	saveData?: boolean
	preserveFormatting?: boolean
	adjustColumnWidth?: boolean
	fillFormulas?: boolean
	disableEdit?: boolean
	disableRefresh?: boolean
	headers?: boolean
	rowNumbers?: boolean
	autoFormatId?: number
	applyNumberFormats?: boolean
	applyBorderFormats?: boolean
	applyFontFormats?: boolean
	applyPatternFormats?: boolean
	applyAlignmentFormats?: boolean
	applyWidthHeightFormats?: boolean
	queryTableRefreshNextId?: number
	queryTableFields?: readonly QueryTableFieldInfo[]
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
	commandType?: number
	serverCommand?: boolean
	webUrl?: string
	webHtmlTables?: boolean
	webXml?: boolean
	webSourceData?: boolean
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
		const root = firstElement(parseXml(xml), 'queryTable')
		return [{ ...base, ...readConnectionAttrs(root, { queryTable: true }) }]
	}
	if (kind === 'connection') {
		const doc = parseXml(xml)
		const root = firstElement(doc, 'connections')
		const connections = childNodes(root, 'connection')
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
	const newConnection = boolAttr(node, 'new')
	if (!options.queryTable && newConnection !== undefined) parsed.newConnection = newConnection
	const backgroundRefresh = boolAttr(node, options.queryTable ? 'backgroundRefresh' : 'background')
	if (backgroundRefresh !== undefined) parsed.backgroundRefresh = backgroundRefresh
	const firstBackgroundRefresh = boolAttr(node, 'firstBackgroundRefresh')
	if (options.queryTable && firstBackgroundRefresh !== undefined) {
		parsed.firstBackgroundRefresh = firstBackgroundRefresh
	}
	const keepAlive = boolAttr(node, 'keepAlive')
	if (keepAlive !== undefined) parsed.keepAlive = keepAlive
	const refreshInterval = numAttr(node, 'interval')
	if (refreshInterval !== undefined) parsed.refreshInterval = refreshInterval
	const refreshOnLoad = boolAttr(node, 'refreshOnLoad')
	if (refreshOnLoad !== undefined) parsed.refreshOnLoad = refreshOnLoad
	const reconnectionMethod = numAttr(node, 'reconnectionMethod')
	if (!options.queryTable && reconnectionMethod !== undefined) {
		parsed.reconnectionMethod = reconnectionMethod
	}
	const saveData = boolAttr(node, 'saveData')
	if (saveData !== undefined) parsed.saveData = saveData
	const savePassword = boolAttr(node, 'savePassword')
	if (savePassword !== undefined) parsed.savePassword = savePassword
	const removeDataOnSave = boolAttr(node, 'removeDataOnSave')
	if (options.queryTable && saveData === undefined && removeDataOnSave !== undefined) {
		parsed.saveData = !removeDataOnSave
	}
	if (options.queryTable) {
		const preserveFormatting = boolAttr(node, 'preserveFormatting')
		if (preserveFormatting !== undefined) parsed.preserveFormatting = preserveFormatting
		const adjustColumnWidth = boolAttr(node, 'adjustColumnWidth')
		if (adjustColumnWidth !== undefined) parsed.adjustColumnWidth = adjustColumnWidth
		const fillFormulas = boolAttr(node, 'fillFormulas')
		if (fillFormulas !== undefined) parsed.fillFormulas = fillFormulas
		const disableEdit = boolAttr(node, 'disableEdit')
		if (disableEdit !== undefined) parsed.disableEdit = disableEdit
		const disableRefresh = boolAttr(node, 'disableRefresh')
		if (disableRefresh !== undefined) parsed.disableRefresh = disableRefresh
		const headers = boolAttr(node, 'headers')
		if (headers !== undefined) parsed.headers = headers
		const rowNumbers = boolAttr(node, 'rowNumbers')
		if (rowNumbers !== undefined) parsed.rowNumbers = rowNumbers
		const autoFormatId = numAttr(node, 'autoFormatId')
		if (autoFormatId !== undefined) parsed.autoFormatId = autoFormatId
		const applyNumberFormats = boolAttr(node, 'applyNumberFormats')
		if (applyNumberFormats !== undefined) parsed.applyNumberFormats = applyNumberFormats
		const applyBorderFormats = boolAttr(node, 'applyBorderFormats')
		if (applyBorderFormats !== undefined) parsed.applyBorderFormats = applyBorderFormats
		const applyFontFormats = boolAttr(node, 'applyFontFormats')
		if (applyFontFormats !== undefined) parsed.applyFontFormats = applyFontFormats
		const applyPatternFormats = boolAttr(node, 'applyPatternFormats')
		if (applyPatternFormats !== undefined) parsed.applyPatternFormats = applyPatternFormats
		const applyAlignmentFormats = boolAttr(node, 'applyAlignmentFormats')
		if (applyAlignmentFormats !== undefined) parsed.applyAlignmentFormats = applyAlignmentFormats
		const applyWidthHeightFormats = boolAttr(node, 'applyWidthHeightFormats')
		if (applyWidthHeightFormats !== undefined) {
			parsed.applyWidthHeightFormats = applyWidthHeightFormats
		}
		const queryTableRefresh = childNode(node, 'queryTableRefresh')
		const queryTableRefreshNextId = queryTableRefresh
			? numAttr(queryTableRefresh, 'nextId')
			: undefined
		if (queryTableRefreshNextId !== undefined) {
			parsed.queryTableRefreshNextId = queryTableRefreshNextId
		}
		const queryTableFields = parseQueryTableFields(queryTableRefresh)
		if (queryTableFields.length > 0) parsed.queryTableFields = queryTableFields
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
	const textPr = childNode(node, 'textPr')
	const dbPr = childNode(node, 'dbPr')
	const webPr = childNode(node, 'webPr')
	const sourceFile = attr(node, 'sourceFile') ?? (textPr ? attr(textPr, 'sourceFile') : undefined)
	if (sourceFile) parsed.sourceFile = sourceFile
	const odcFile = attr(node, 'odcFile')
	if (odcFile) parsed.odcFile = odcFile
	const onlyUseConnectionFile = boolAttr(node, 'onlyUseConnectionFile')
	if (onlyUseConnectionFile !== undefined) parsed.onlyUseConnectionFile = onlyUseConnectionFile
	const command = dbPr ? attr(dbPr, 'command') : undefined
	if (command) parsed.command = command
	const commandType = dbPr ? numAttr(dbPr, 'commandType') : undefined
	if (commandType !== undefined) parsed.commandType = commandType
	const serverCommand = dbPr ? boolAttr(dbPr, 'serverCommand') : undefined
	if (serverCommand !== undefined) parsed.serverCommand = serverCommand
	const webUrl = webPr ? attr(webPr, 'url') : undefined
	if (webUrl) parsed.webUrl = webUrl
	const webHtmlTables = webPr ? boolAttr(webPr, 'htmlTables') : undefined
	if (webHtmlTables !== undefined) parsed.webHtmlTables = webHtmlTables
	const webXml = webPr ? boolAttr(webPr, 'xml') : undefined
	if (webXml !== undefined) parsed.webXml = webXml
	const webSourceData = webPr ? boolAttr(webPr, 'sourceData') : undefined
	if (webSourceData !== undefined) parsed.webSourceData = webSourceData
	const connectionString = dbPr ? attr(dbPr, 'connection') : undefined
	if (connectionString) parsed.hasConnectionString = true
	return parsed
}

function parseQueryTableFields(queryTableRefresh: XmlNode | undefined): QueryTableFieldInfo[] {
	const fieldsRoot = childNode(queryTableRefresh, 'queryTableFields')
	const fields: QueryTableFieldInfo[] = []
	for (const field of childNodes(fieldsRoot, 'queryTableField')) {
		const id = numAttr(field, 'id')
		if (id === undefined) continue
		const name = attr(field, 'name')
		const tableColumnId = numAttr(field, 'tableColumnId')
		fields.push({
			id,
			...(name ? { name } : {}),
			...(tableColumnId !== undefined ? { tableColumnId } : {}),
		})
	}
	return fields
}

function firstElement(doc: XmlNode, localName: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(doc)) {
		if (key.startsWith('@_') || localPart(key) !== localName || !isXmlNode(value)) continue
		return value
	}
	return undefined
}

function childNode(node: XmlNode | undefined, localName: string): XmlNode | undefined {
	return childNodes(node, localName)[0]
}

function childNodes(node: XmlNode | undefined, localName: string): XmlNode[] {
	if (!node) return []
	const matches: XmlNode[] = []
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || localPart(key) !== localName) continue
		for (const item of asArray(value as XmlNode | XmlNode[] | undefined)) {
			if (isXmlNode(item)) matches.push(item)
		}
	}
	return matches
}

function isXmlNode(value: unknown): value is XmlNode {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function localPart(name: string): string {
	return name.includes(':') ? (name.split(':').pop() ?? name) : name
}
