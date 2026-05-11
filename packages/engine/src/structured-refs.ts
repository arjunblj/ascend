import type { Table, Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'

export interface StructuredRefRange {
	readonly sheetIndex: number
	readonly startRow: number
	readonly endRow: number
	readonly startCol: number
	readonly endCol: number
}

export interface StructuredRefResolver {
	resolve(
		node: Extract<FormulaNode, { type: 'structuredRef' }>,
		sheetIndex: number,
		row: number,
		col: number,
	): StructuredRefRange | null
}

interface ResolvedTableRef {
	readonly sheetIndex: number
	readonly table: Table
}

export function createStructuredRefResolver(workbook: Workbook): StructuredRefResolver {
	const firstTableByName = new Map<string, ResolvedTableRef>()
	const tableBySheetAndName = new Map<number, Map<string, ResolvedTableRef>>()
	const tablesBySheet = new Map<number, ResolvedTableRef[]>()
	const columnIndexByTable = new Map<Table, Map<string, number>>()

	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet || sheet.tables.length === 0) continue
		const sheetTables: ResolvedTableRef[] = []
		const sheetByName = new Map<string, ResolvedTableRef>()
		for (const table of sheet.tables) {
			const ref = { sheetIndex, table }
			sheetTables.push(ref)
			const name = table.name.toLowerCase()
			if (!firstTableByName.has(name)) firstTableByName.set(name, ref)
			sheetByName.set(name, ref)
		}
		tablesBySheet.set(sheetIndex, sheetTables)
		tableBySheetAndName.set(sheetIndex, sheetByName)
	}

	const columnIndex = (table: Table, name: string): number => {
		let indexByName = columnIndexByTable.get(table)
		if (!indexByName) {
			indexByName = new Map()
			for (let index = 0; index < table.columns.length; index++) {
				indexByName.set(table.columns[index]?.name.toLowerCase() ?? '', index)
			}
			columnIndexByTable.set(table, indexByName)
		}
		return indexByName.get(name.toLowerCase()) ?? -1
	}

	const resolveTable = (
		tableName: string,
		sheetIndex: number,
		row: number,
		col: number,
	): ResolvedTableRef | null => {
		const tableNameLower = tableName.toLowerCase()
		if (tableName.length > 0) {
			return (
				tableBySheetAndName.get(sheetIndex)?.get(tableNameLower) ??
				firstTableByName.get(tableNameLower) ??
				null
			)
		}

		const sheetTables = tablesBySheet.get(sheetIndex)
		if (!sheetTables) return null
		for (const ref of sheetTables) {
			const { table } = ref
			if (
				row >= table.ref.start.row &&
				row <= table.ref.end.row &&
				col >= table.ref.start.col &&
				col <= table.ref.end.col
			) {
				return ref
			}
		}
		return null
	}

	return {
		resolve(node, sheetIndex, row, col) {
			return resolveStructuredRefRangeWithResolver(
				node,
				sheetIndex,
				row,
				col,
				resolveTable,
				columnIndex,
			)
		},
	}
}

export function resolveStructuredRefRange(
	workbook: Workbook,
	node: Extract<FormulaNode, { type: 'structuredRef' }>,
	sheetIndex: number,
	row: number,
	col: number,
): StructuredRefRange | null {
	return createStructuredRefResolver(workbook).resolve(node, sheetIndex, row, col)
}

function resolveStructuredRefRangeWithResolver(
	node: Extract<FormulaNode, { type: 'structuredRef' }>,
	sheetIndex: number,
	row: number,
	col: number,
	resolveTable: (
		tableName: string,
		sheetIndex: number,
		row: number,
		col: number,
	) => ResolvedTableRef | null,
	columnIndex: (table: Table, name: string) => number,
): StructuredRefRange | null {
	const tableRef = resolveTable(node.table, sheetIndex, row, col)
	if (!tableRef) return null

	const { table } = tableRef
	const bodyStartRow = table.ref.start.row + (table.hasHeaders ? 1 : 0)
	const bodyEndRow = table.ref.end.row - (table.hasTotals ? 1 : 0)
	const headerRow = table.ref.start.row
	const totalsRow = table.ref.end.row

	let startRow = bodyStartRow
	let endRow = bodyEndRow
	if (node.specifiers.includes('@') || node.specifiers.includes('#This Row')) {
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
		const startColumnIndex = columnIndex(table, node.column)
		if (startColumnIndex < 0) return null
		startCol = table.ref.start.col + startColumnIndex
		if (node.endColumn) {
			const endColumnIndex = columnIndex(table, node.endColumn)
			if (endColumnIndex < startColumnIndex) return null
			endCol = table.ref.start.col + endColumnIndex
		} else {
			endCol = startCol
		}
	}

	return {
		sheetIndex: tableRef.sheetIndex,
		startRow,
		endRow,
		startCol,
		endCol,
	}
}
