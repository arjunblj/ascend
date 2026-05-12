import {
	createTableId,
	parseRange,
	type RangeRef,
	type SheetId,
	type SortState,
	type Table,
	type TableColumn,
	type TableQueryTableRef,
	type TableStyleInfo,
	type TableXmlColumnPr,
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { parseAutoFilterNode, parseSortStateNode } from './filtering.ts'
import type { Relationship } from './relationships.ts'
import { REL_QUERY_TABLE, resolvePath } from './relationships.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'

export interface ParseTableOptions {
	readonly tablePath?: string
	readonly contentType?: string
	readonly contentTypeSource?: 'override' | 'default' | 'fallback'
	readonly sourcePartPath?: string
	readonly sourceRelationshipPart?: string
	readonly sourceRelationship?: Relationship
	readonly sourceRelationshipResolvedTarget?: string
	readonly relationships?: readonly Relationship[]
}

export function parseTable(
	xml: string,
	sheetId: SheetId,
	options: ParseTableOptions = {},
): Table | null {
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml), { preserveWhitespace: true })
	const table = doc.table as XmlNode | undefined
	if (!table) return null

	const displayName = attr(table, 'displayName')
	const nameAttribute = attr(table, 'name')
	const name = displayName ?? nameAttribute
	const refText = attr(table, 'ref')
	if (!name || !refText) return null

	let ref: RangeRef
	try {
		ref = parseRange(refText)
	} catch {
		return null
	}

	const headerRowCount = numAttr(table, 'headerRowCount')
	const totalsRowCount =
		numAttr(table, 'totalsRowCount') ?? (boolAttr(table, 'totalsRowShown') ? 1 : 0)
	const uid = attr(table, 'xr:uid')
	const columns = parseTableColumns(table)
	const autoFilter = parseAutoFilterNode(table.autoFilter as XmlNode | undefined)
	const sortState = parseSortStateNode(table.sortState as XmlNode | undefined, {
		preserveAttributes: true,
	})
	const tableStyleInfo = parseTableStyleInfo(table.tableStyleInfo as XmlNode | undefined)
	const tableType = attr(table, 'tableType')
	const insertRow = boolAttr(table, 'insertRow')
	const insertRowShift = boolAttr(table, 'insertRowShift')
	const altTextInfo = parseTableAltText(table)
	const dxfId = numAttr(table, 'dxfId')
	const dataCellStyle = attr(table, 'dataCellStyle')
	const headerRowDxfId = numAttr(table, 'headerRowDxfId')
	const headerRowCellStyle = attr(table, 'headerRowCellStyle')
	const dataDxfId = numAttr(table, 'dataDxfId')
	const totalsRowDxfId = numAttr(table, 'totalsRowDxfId')
	const headerRowBorderDxfId = numAttr(table, 'headerRowBorderDxfId')
	const tableBorderDxfId = numAttr(table, 'tableBorderDxfId')
	const queryTable = parseQueryTableRef(options.tablePath, options.relationships)

	const parsed: {
		id: Table['id']
		name: string
		nameAttribute?: string | null
		sheetId: SheetId
		partPath?: string
		contentType?: string
		contentTypeSource?: 'override' | 'default' | 'fallback'
		sourcePartPath?: string
		sourceRelationshipPart?: string
		sourceRelationshipId?: string
		sourceRelationshipType?: string
		sourceRelationshipRawType?: string
		sourceRelationshipRawTarget?: string
		sourceRelationshipResolvedTarget?: string
		sourceRelationshipTargetMode?: string
		uid?: string
		ref: RangeRef
		tableType?: string
		insertRow?: boolean
		insertRowShift?: boolean
		columns: readonly TableColumn[]
		hasHeaders: boolean
		hasTotals: boolean
		altText?: string
		altTextSummary?: string
		autoFilter?: Table['autoFilter']
		sortState?: SortState
		dxfId?: number
		dataCellStyle?: string
		headerRowDxfId?: number
		headerRowCellStyle?: string
		dataDxfId?: number
		totalsRowDxfId?: number
		headerRowBorderDxfId?: number
		tableBorderDxfId?: number
		tableStyleInfo?: TableStyleInfo
		queryTable?: TableQueryTableRef
	} = {
		id: createTableId(),
		name,
		sheetId,
		ref,
		columns,
		hasHeaders: (headerRowCount ?? 1) !== 0,
		hasTotals: totalsRowCount > 0,
	}
	if (tableType) parsed.tableType = tableType
	if (nameAttribute !== undefined) parsed.nameAttribute = nameAttribute
	else parsed.nameAttribute = null
	if (options.tablePath) parsed.partPath = options.tablePath
	if (options.contentType) parsed.contentType = options.contentType
	if (options.contentTypeSource) parsed.contentTypeSource = options.contentTypeSource
	if (options.sourcePartPath) parsed.sourcePartPath = options.sourcePartPath
	if (options.sourceRelationshipPart) parsed.sourceRelationshipPart = options.sourceRelationshipPart
	if (options.sourceRelationship) {
		parsed.sourceRelationshipId = options.sourceRelationship.id
		parsed.sourceRelationshipType = options.sourceRelationship.type
		if (options.sourceRelationship.rawType) {
			parsed.sourceRelationshipRawType = options.sourceRelationship.rawType
		}
		parsed.sourceRelationshipRawTarget = options.sourceRelationship.target
		if (options.sourceRelationship.targetMode) {
			parsed.sourceRelationshipTargetMode = options.sourceRelationship.targetMode
		}
	}
	if (options.sourceRelationshipResolvedTarget) {
		parsed.sourceRelationshipResolvedTarget = options.sourceRelationshipResolvedTarget
	}
	if (uid) parsed.uid = uid
	if (insertRow !== undefined) parsed.insertRow = insertRow
	if (insertRowShift !== undefined) parsed.insertRowShift = insertRowShift
	if (altTextInfo.altText) parsed.altText = altTextInfo.altText
	if (altTextInfo.altTextSummary) parsed.altTextSummary = altTextInfo.altTextSummary
	if (autoFilter) parsed.autoFilter = autoFilter
	if (sortState) parsed.sortState = sortState
	if (dxfId !== undefined) parsed.dxfId = dxfId
	if (dataCellStyle) parsed.dataCellStyle = dataCellStyle
	if (headerRowDxfId !== undefined) parsed.headerRowDxfId = headerRowDxfId
	if (headerRowCellStyle) parsed.headerRowCellStyle = headerRowCellStyle
	if (dataDxfId !== undefined) parsed.dataDxfId = dataDxfId
	if (totalsRowDxfId !== undefined) parsed.totalsRowDxfId = totalsRowDxfId
	if (headerRowBorderDxfId !== undefined) parsed.headerRowBorderDxfId = headerRowBorderDxfId
	if (tableBorderDxfId !== undefined) parsed.tableBorderDxfId = tableBorderDxfId
	if (tableStyleInfo) parsed.tableStyleInfo = tableStyleInfo
	if (queryTable) parsed.queryTable = queryTable
	return parsed as Table
}

function parseTableAltText(table: XmlNode): { altText?: string; altTextSummary?: string } {
	let altText = attr(table, 'altText')
	let altTextSummary = attr(table, 'altTextSummary')
	for (const ext of asArray<XmlNode>(
		(table.extLst as XmlNode | undefined)?.ext as XmlNode | XmlNode[],
	)) {
		const x14Table = findX14TableNode(ext)
		if (!x14Table) continue
		altText = attr(x14Table, 'altText') ?? altText
		altTextSummary = attr(x14Table, 'altTextSummary') ?? altTextSummary
	}
	return { ...(altText ? { altText } : {}), ...(altTextSummary ? { altTextSummary } : {}) }
}

function findX14TableNode(ext: XmlNode): XmlNode | undefined {
	for (const [key, value] of Object.entries(ext)) {
		const separator = key.indexOf(':')
		if (separator <= 0 || key.slice(separator + 1) !== 'table') continue
		const prefix = key.slice(0, separator)
		if (attr(ext, `xmlns:${prefix}`) !== X14_NS) continue
		return value as XmlNode
	}
	return undefined
}

function parseTableColumns(table: XmlNode): readonly TableColumn[] {
	const tableColumnsNode = table.tableColumns as XmlNode | undefined
	if (!tableColumnsNode) return []
	const columns: TableColumn[] = []
	for (const column of asArray<XmlNode>(tableColumnsNode.tableColumn as XmlNode | XmlNode[])) {
		const name = attr(column, 'name')
		if (!name) continue
		const formulaNode = column.calculatedColumnFormula
		const formula =
			formulaNode !== undefined && formulaNode !== null
				? extractFormulaText(formulaNode)
				: undefined
		const totalsRowFormulaNode = column.totalsRowFormula
		const totalsRowFormula =
			totalsRowFormulaNode !== undefined && totalsRowFormulaNode !== null
				? extractFormulaText(totalsRowFormulaNode)
				: undefined
		const parsed: {
			id?: number
			uid?: string
			uniqueName?: string
			name: string
			formula?: string
			formulaIsArray?: boolean
			xmlColumnPr?: TableXmlColumnPr
			totalsRowFunction?: string
			totalsRowFormula?: string
			totalsRowLabel?: string
			queryTableFieldId?: number
			dataCellStyle?: string
			dataDxfId?: number
			headerRowDxfId?: number
			totalsRowDxfId?: number
		} = { name }
		const id = numAttr(column, 'id')
		if (id !== undefined) parsed.id = id
		const uid = attr(column, 'xr3:uid')
		if (uid) parsed.uid = uid
		const uniqueName = attr(column, 'uniqueName')
		if (uniqueName) parsed.uniqueName = uniqueName
		if (formula) parsed.formula = formula
		if (formula) {
			const formulaIsArray = boolAttr(formulaNode as XmlNode, 'array')
			if (formulaIsArray !== undefined) parsed.formulaIsArray = formulaIsArray
		}
		const xmlColumnPr = parseXmlColumnPr(column.xmlColumnPr as XmlNode | undefined)
		if (xmlColumnPr) parsed.xmlColumnPr = xmlColumnPr
		const totalsRowFunction = attr(column, 'totalsRowFunction')
		if (totalsRowFunction) parsed.totalsRowFunction = totalsRowFunction
		if (totalsRowFormula) parsed.totalsRowFormula = totalsRowFormula
		const totalsRowLabel = attr(column, 'totalsRowLabel')
		if (totalsRowLabel) parsed.totalsRowLabel = totalsRowLabel
		const queryTableFieldId = numAttr(column, 'queryTableFieldId')
		if (queryTableFieldId !== undefined) parsed.queryTableFieldId = queryTableFieldId
		const dataCellStyle = attr(column, 'dataCellStyle')
		if (dataCellStyle) parsed.dataCellStyle = dataCellStyle
		const dataDxfId = numAttr(column, 'dataDxfId')
		if (dataDxfId !== undefined) parsed.dataDxfId = dataDxfId
		const headerRowDxfId = numAttr(column, 'headerRowDxfId')
		if (headerRowDxfId !== undefined) parsed.headerRowDxfId = headerRowDxfId
		const totalsRowDxfId = numAttr(column, 'totalsRowDxfId')
		if (totalsRowDxfId !== undefined) parsed.totalsRowDxfId = totalsRowDxfId
		columns.push(parsed)
	}
	return columns
}

function parseXmlColumnPr(node: XmlNode | undefined): TableXmlColumnPr | undefined {
	if (!node) return undefined
	const parsed: {
		mapId?: number
		xpath?: string
		xmlDataType?: string
	} = {}
	const mapId = numAttr(node, 'mapId')
	if (mapId !== undefined) parsed.mapId = mapId
	const xpath = attr(node, 'xpath')
	if (xpath) parsed.xpath = xpath
	const xmlDataType = attr(node, 'xmlDataType')
	if (xmlDataType) parsed.xmlDataType = xmlDataType
	return Object.keys(parsed).length > 0 ? parsed : undefined
}

function parseQueryTableRef(
	tablePath: string | undefined,
	relationships: readonly Relationship[] | undefined,
): TableQueryTableRef | undefined {
	if (!tablePath || !relationships) return undefined
	const rel = relationships.find((entry) => entry.type === REL_QUERY_TABLE)
	if (!rel) return undefined
	return {
		relationshipId: rel.id,
		partPath: resolvePath(tablePath, rel.target),
		relationshipType: rel.type,
		target: rel.target,
		...(rel.targetMode ? { targetMode: rel.targetMode } : {}),
	}
}

function parseTableStyleInfo(node: XmlNode | undefined): TableStyleInfo | undefined {
	if (!node) return undefined
	const parsed: {
		name?: string
		showFirstColumn?: boolean
		showLastColumn?: boolean
		showRowStripes?: boolean
		showColumnStripes?: boolean
	} = {}
	const name = attr(node, 'name')
	if (name) parsed.name = name
	const showFirstColumn = boolAttr(node, 'showFirstColumn')
	if (showFirstColumn !== undefined) parsed.showFirstColumn = showFirstColumn
	const showLastColumn = boolAttr(node, 'showLastColumn')
	if (showLastColumn !== undefined) parsed.showLastColumn = showLastColumn
	const showRowStripes = boolAttr(node, 'showRowStripes')
	if (showRowStripes !== undefined) parsed.showRowStripes = showRowStripes
	const showColumnStripes = boolAttr(node, 'showColumnStripes')
	if (showColumnStripes !== undefined) parsed.showColumnStripes = showColumnStripes
	return Object.keys(parsed).length > 0 ? parsed : undefined
}

function extractFormulaText(node: unknown): string | undefined {
	if (typeof node === 'string' || typeof node === 'number' || typeof node === 'boolean') {
		return String(node)
	}
	if (typeof node === 'object' && node !== null) {
		const text = (node as XmlNode)['#text']
		if (text !== undefined && text !== null) return String(text)
	}
	return undefined
}
