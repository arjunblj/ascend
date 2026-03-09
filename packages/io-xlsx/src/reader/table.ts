import {
	createTableId,
	parseRange,
	type RangeRef,
	type SheetId,
	type Table,
	type TableColumn,
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import { parseAutoFilterNode } from './filtering.ts'

export function parseTable(xml: string, sheetId: SheetId): Table | null {
	const doc = parseXml(xml)
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

	return {
		id: createTableId(),
		name,
		sheetId,
		ref,
		columns,
		hasHeaders: (headerRowCount ?? 1) !== 0,
		hasTotals: totalsRowCount > 0,
		...(autoFilter ? { autoFilter } : {}),
	}
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
		if (formula) {
			columns.push({ name, formula })
		} else {
			columns.push({ name })
		}
	}
	return columns
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
