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
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { parseAutoFilterNode, parseSortStateNode } from './filtering.ts'
import type { Relationship } from './relationships.ts'
import { REL_QUERY_TABLE, resolvePath } from './relationships.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

export interface ParseTableOptions {
	readonly tablePath?: string
	readonly relationships?: readonly Relationship[]
}

export function parseTable(
	xml: string,
	sheetId: SheetId,
	options: ParseTableOptions = {},
): Table | null {
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
	const table = doc.table as XmlNode | undefined
	if (!table) return null

	const displayName = attr(table, 'displayName')
	const name = displayName ?? attr(table, 'name')
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
	const columns = parseTableColumns(table)
	const autoFilter = parseAutoFilterNode(table.autoFilter as XmlNode | undefined)
	const sortState = parseSortStateNode(table.sortState as XmlNode | undefined)
	const tableStyleInfo = parseTableStyleInfo(table.tableStyleInfo as XmlNode | undefined)
	const tableType = attr(table, 'tableType')
	const insertRow = boolAttr(table, 'insertRow')
	const insertRowShift = boolAttr(table, 'insertRowShift')
	const dxfId = numAttr(table, 'dxfId')
	const headerRowDxfId = numAttr(table, 'headerRowDxfId')
	const dataDxfId = numAttr(table, 'dataDxfId')
	const totalsRowDxfId = numAttr(table, 'totalsRowDxfId')
	const headerRowBorderDxfId = numAttr(table, 'headerRowBorderDxfId')
	const queryTable = parseQueryTableRef(options.tablePath, options.relationships)

	const parsed: {
		id: Table['id']
		name: string
		sheetId: SheetId
		ref: RangeRef
		tableType?: string
		insertRow?: boolean
		insertRowShift?: boolean
		columns: readonly TableColumn[]
		hasHeaders: boolean
		hasTotals: boolean
		autoFilter?: Table['autoFilter']
		sortState?: SortState
		dxfId?: number
		headerRowDxfId?: number
		dataDxfId?: number
		totalsRowDxfId?: number
		headerRowBorderDxfId?: number
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
	if (insertRow !== undefined) parsed.insertRow = insertRow
	if (insertRowShift !== undefined) parsed.insertRowShift = insertRowShift
	if (autoFilter) parsed.autoFilter = autoFilter
	if (sortState) parsed.sortState = sortState
	if (dxfId !== undefined) parsed.dxfId = dxfId
	if (headerRowDxfId !== undefined) parsed.headerRowDxfId = headerRowDxfId
	if (dataDxfId !== undefined) parsed.dataDxfId = dataDxfId
	if (totalsRowDxfId !== undefined) parsed.totalsRowDxfId = totalsRowDxfId
	if (headerRowBorderDxfId !== undefined) parsed.headerRowBorderDxfId = headerRowBorderDxfId
	if (tableStyleInfo) parsed.tableStyleInfo = tableStyleInfo
	if (queryTable) parsed.queryTable = queryTable
	return parsed as Table
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
			uniqueName?: string
			name: string
			formula?: string
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
		const uniqueName = attr(column, 'uniqueName')
		if (uniqueName) parsed.uniqueName = uniqueName
		if (formula) parsed.formula = formula
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
