import type { Table, Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'

export interface StructuredRefRange {
	readonly sheetIndex: number
	readonly startRow: number
	readonly endRow: number
	readonly startCol: number
	readonly endCol: number
}

export function resolveStructuredRefRange(
	workbook: Workbook,
	node: Extract<FormulaNode, { type: 'structuredRef' }>,
	sheetIndex: number,
	row: number,
	col: number,
): StructuredRefRange | null {
	const tableRef = resolveTableRef(workbook, node.table, sheetIndex, row, col)
	if (!tableRef) return null

	const { table } = tableRef
	const bodyStartRow = table.ref.start.row + (table.hasHeaders ? 1 : 0)
	const bodyEndRow = table.ref.end.row - (table.hasTotals ? 1 : 0)
	const headerRow = table.ref.start.row
	const totalsRow = table.ref.end.row

	let startRow = bodyStartRow
	let endRow = bodyEndRow
	if (node.specifiers.includes('@')) {
		if (row < bodyStartRow || row > bodyEndRow) return null
		startRow = row
		endRow = row
	} else if (node.specifiers.includes('#Headers')) {
		startRow = headerRow
		endRow = headerRow
	} else if (node.specifiers.includes('#Totals')) {
		if (!table.hasTotals) return null
		startRow = totalsRow
		endRow = totalsRow
	} else if (node.specifiers.includes('#All')) {
		startRow = table.ref.start.row
		endRow = table.ref.end.row
	} else if (node.specifiers.includes('#Data')) {
		startRow = bodyStartRow
		endRow = bodyEndRow
	}

	let startCol = table.ref.start.col
	let endCol = table.ref.end.col
	if (node.column) {
		const columnIndex = table.columns.findIndex(
			(column) => column.name.toLowerCase() === node.column?.toLowerCase(),
		)
		if (columnIndex < 0) return null
		startCol = table.ref.start.col + columnIndex
		endCol = startCol
	}

	return {
		sheetIndex: tableRef.sheetIndex,
		startRow,
		endRow,
		startCol,
		endCol,
	}
}

function resolveTableRef(
	workbook: Workbook,
	tableName: string,
	sheetIndex: number,
	row: number,
	col: number,
): { sheetIndex: number; table: Table } | null {
	const tableNameLower = tableName.toLowerCase()
	if (tableName.length > 0) {
		const currentSheet = workbook.sheets[sheetIndex]
		if (currentSheet) {
			const table = currentSheet.tables.find(
				(candidate) => candidate.name.toLowerCase() === tableNameLower,
			)
			if (table) return { sheetIndex, table }
		}
		for (let i = 0; i < workbook.sheets.length; i++) {
			if (i === sheetIndex) continue
			const sheet = workbook.sheets[i]
			if (!sheet) continue
			const table = sheet.tables.find(
				(candidate) => candidate.name.toLowerCase() === tableNameLower,
			)
			if (table) return { sheetIndex: i, table }
		}
		return null
	}

	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return null
	const table = sheet.tables.find(
		(candidate) =>
			row >= candidate.ref.start.row &&
			row <= candidate.ref.end.row &&
			col >= candidate.ref.start.col &&
			col <= candidate.ref.end.col,
	)
	return table ? { sheetIndex, table } : null
}
