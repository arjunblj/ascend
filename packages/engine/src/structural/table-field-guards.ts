import type {
	RangeRef,
	Sheet,
	SheetConditionalFormatValueObject,
	Table,
	Workbook,
} from '@ascend/core'
import { toA1 } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { cachedParseFormula } from '@ascend/formulas'

export interface DeletedTableColumnSet {
	readonly table: Table
	readonly tableNameLower: string
	readonly deletedIndexes: ReadonlySet<number>
	readonly deletedNames: ReadonlySet<string>
	readonly indexByName: ReadonlyMap<string, number>
}

export interface DeletedTableColumnReference {
	readonly sourceKind: string
	readonly sourceRef: string
	readonly tableName: string
	readonly columnName: string
}

export interface DeletedTableColumnReferenceOptions {
	readonly skipDeletedCells?: {
		readonly sheet: Sheet
		readonly startCol: number
		readonly endColExclusive: number
	}
	readonly skipDeletedTableColumnFormulas?: boolean
}

export function collectDeletedTableColumns(
	sheet: Sheet,
	at: number,
	count: number,
): DeletedTableColumnSet[] {
	const deleted: DeletedTableColumnSet[] = []
	const deleteEnd = at + count - 1
	for (const table of sheet.tables) {
		const deletedIndexes = new Set<number>()
		const deletedNames = new Set<string>()
		const indexByName = new Map<string, number>()
		for (const [index, column] of table.columns.entries()) {
			const absoluteCol = table.ref.start.col + index
			const lowerName = column.name.toLowerCase()
			indexByName.set(lowerName, index)
			if (absoluteCol < at || absoluteCol > deleteEnd) continue
			deletedIndexes.add(index)
			deletedNames.add(lowerName)
		}
		if (deletedIndexes.size === 0) continue
		deleted.push({
			table,
			tableNameLower: table.name.toLowerCase(),
			deletedIndexes,
			deletedNames,
			indexByName,
		})
	}
	return deleted
}

export function collectDroppedTableColumnsForResize(
	table: Table,
	nextRef: RangeRef,
): DeletedTableColumnSet[] {
	const deletedIndexes = new Set<number>()
	const deletedNames = new Set<string>()
	const indexByName = new Map<string, number>()
	for (const [index, column] of table.columns.entries()) {
		const lowerName = column.name.toLowerCase()
		indexByName.set(lowerName, index)
		const absoluteCol = table.ref.start.col + index
		if (absoluteCol >= nextRef.start.col && absoluteCol <= nextRef.end.col) continue
		deletedIndexes.add(index)
		deletedNames.add(lowerName)
	}
	return deletedIndexes.size === 0
		? []
		: [
				{
					table,
					tableNameLower: table.name.toLowerCase(),
					deletedIndexes,
					deletedNames,
					indexByName,
				},
			]
}

export function findDeletedTableColumnReference(
	workbook: Workbook,
	deletedSets: readonly DeletedTableColumnSet[],
	options: DeletedTableColumnReferenceOptions = {},
): DeletedTableColumnReference | null {
	if (deletedSets.length === 0) return null

	for (const sheet of workbook.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (cell.formula === null) continue
			if (
				options.skipDeletedCells &&
				sheet === options.skipDeletedCells.sheet &&
				col >= options.skipDeletedCells.startCol &&
				col < options.skipDeletedCells.endColExclusive
			) {
				continue
			}
			const parsed = cachedParseFormula(cell.formula)
			if (!parsed.ok) continue
			const localTableName = tableContainingCell(sheet, row, col)?.name
			const match = findDeletedTableColumnInFormula(parsed.value, deletedSets, localTableName)
			if (!match) continue
			return {
				sourceKind: 'cell formula',
				sourceRef: `${sheet.name}!${toA1({ row, col })}`,
				tableName: match.table.name,
				columnName: match.columnName,
			}
		}
	}

	for (const entry of workbook.definedNames.list()) {
		const parsed = cachedParseFormula(entry.formula)
		if (!parsed.ok) continue
		const match = findDeletedTableColumnInFormula(parsed.value, deletedSets)
		if (!match) continue
		return {
			sourceKind: 'defined name',
			sourceRef: entry.name,
			tableName: match.table.name,
			columnName: match.columnName,
		}
	}

	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			const localTableName = table.name
			const deletedTableColumns = deletedSets.find((set) => set.table === table)
			for (const [index, column] of table.columns.entries()) {
				if (
					options.skipDeletedTableColumnFormulas &&
					deletedTableColumns?.deletedIndexes.has(index)
				) {
					continue
				}
				const formulas = [
					['table column formula', column.formula],
					['table totals formula', column.totalsRowFormula],
				] as const
				for (const [sourceKind, formula] of formulas) {
					const match = findDeletedTableColumnInFormulaText(formula, deletedSets, localTableName)
					if (!match) continue
					return {
						sourceKind,
						sourceRef: `${sheet.name}:${table.name}[${column.name}]`,
						tableName: match.table.name,
						columnName: match.columnName,
					}
				}
			}
		}
	}

	for (const sheet of workbook.sheets) {
		const match = findDeletedTableColumnInSheetMetadata(sheet, deletedSets)
		if (match) return match
	}

	return null
}

function findDeletedTableColumnInSheetMetadata(
	sheet: Sheet,
	deletedSets: readonly DeletedTableColumnSet[],
): DeletedTableColumnReference | null {
	for (const validation of sheet.dataValidations) {
		for (const [label, formula] of [
			['data validation formula1', validation.formula1],
			['data validation formula2', validation.formula2],
		] as const) {
			const match = findDeletedTableColumnInFormulaText(formula, deletedSets)
			if (!match) continue
			return deletedTableMetadataReference(label, sheet, validation.sqref, match)
		}
	}
	for (const validation of sheet.x14DataValidations) {
		for (const [label, formula] of [
			['x14 data validation formula1', validation.formula1],
			['x14 data validation formula2', validation.formula2],
		] as const) {
			const match = findDeletedTableColumnInFormulaText(formula, deletedSets)
			if (!match) continue
			return deletedTableMetadataReference(label, sheet, validation.sqref, match)
		}
	}
	for (const format of sheet.conditionalFormats) {
		for (const rule of format.rules) {
			for (const formula of rule.formulas) {
				const match = findDeletedTableColumnInFormulaText(formula, deletedSets)
				if (!match) continue
				return deletedTableMetadataReference(
					'conditional format formula',
					sheet,
					format.sqref,
					match,
				)
			}
			const match =
				findDeletedTableColumnInValueObjects(rule.colorScale?.cfvo, deletedSets) ??
				findDeletedTableColumnInValueObjects(rule.dataBar?.cfvo, deletedSets) ??
				findDeletedTableColumnInValueObjects(rule.iconSet?.cfvo, deletedSets)
			if (!match) continue
			return deletedTableMetadataReference(
				'conditional format value object',
				sheet,
				format.sqref,
				match,
			)
		}
	}
	for (const format of sheet.x14ConditionalFormats) {
		for (const formula of format.formulas ?? []) {
			const match = findDeletedTableColumnInFormulaText(formula, deletedSets)
			if (!match) continue
			return deletedTableMetadataReference(
				'x14 conditional format formula',
				sheet,
				format.sqref,
				match,
			)
		}
		const match =
			findDeletedTableColumnInValueObjects(format.dataBar?.cfvo, deletedSets) ??
			findDeletedTableColumnInValueObjects(format.iconSet?.cfvo, deletedSets)
		if (!match) continue
		return deletedTableMetadataReference(
			'x14 conditional format value object',
			sheet,
			format.sqref,
			match,
		)
	}
	return null
}

function findDeletedTableColumnInValueObjects(
	cfvo: readonly SheetConditionalFormatValueObject[] | undefined,
	deletedSets: readonly DeletedTableColumnSet[],
): DeletedTableColumnMatch | null {
	if (!cfvo) return null
	for (const entry of cfvo) {
		const match = findDeletedTableColumnInFormulaText(entry.value, deletedSets)
		if (match) return match
	}
	return null
}

function deletedTableMetadataReference(
	sourceKind: string,
	sheet: Sheet,
	ref: string,
	match: DeletedTableColumnMatch,
): DeletedTableColumnReference {
	return {
		sourceKind,
		sourceRef: `${sheet.name}!${ref}`,
		tableName: match.table.name,
		columnName: match.columnName,
	}
}

interface DeletedTableColumnMatch {
	readonly table: Table
	readonly columnName: string
}

function findDeletedTableColumnInFormulaText(
	formula: string | undefined,
	deletedSets: readonly DeletedTableColumnSet[],
	localTableName?: string,
): DeletedTableColumnMatch | null {
	if (!formula) return null
	const parsed = cachedParseFormula(formula)
	return parsed.ok
		? findDeletedTableColumnInFormula(parsed.value, deletedSets, localTableName)
		: null
}

function findDeletedTableColumnInFormula(
	node: FormulaNode,
	deletedSets: readonly DeletedTableColumnSet[],
	localTableName?: string,
): DeletedTableColumnMatch | null {
	switch (node.type) {
		case 'structuredRef':
			return findDeletedStructuredReference(node, deletedSets, localTableName)
		case 'binary':
			return (
				findDeletedTableColumnInFormula(node.left, deletedSets, localTableName) ??
				findDeletedTableColumnInFormula(node.right, deletedSets, localTableName)
			)
		case 'dynamicRangeRef':
			return (
				findDeletedTableColumnInFormula(node.start, deletedSets, localTableName) ??
				findDeletedTableColumnInFormula(node.end, deletedSets, localTableName)
			)
		case 'unary':
			return findDeletedTableColumnInFormula(node.operand, deletedSets, localTableName)
		case 'function':
			for (const arg of node.args) {
				const match = findDeletedTableColumnInFormula(arg, deletedSets, localTableName)
				if (match) return match
			}
			return null
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) {
					const match = findDeletedTableColumnInFormula(cell, deletedSets, localTableName)
					if (match) return match
				}
			}
			return null
		case 'spillRef':
			return findDeletedTableColumnInFormula(node.target, deletedSets, localTableName)
		case 'sheetSpanRef':
			return findDeletedTableColumnInFormula(node.target, deletedSets, localTableName)
		default:
			return null
	}
}

function findDeletedStructuredReference(
	node: Extract<FormulaNode, { type: 'structuredRef' }>,
	deletedSets: readonly DeletedTableColumnSet[],
	localTableName?: string,
): DeletedTableColumnMatch | null {
	const tableName = node.table || localTableName
	if (!tableName) return null
	const deleted = deletedSets.find((set) => set.tableNameLower === tableName.toLowerCase())
	if (!deleted) return null
	if (!node.column && !node.endColumn) {
		return { table: deleted.table, columnName: firstDeletedColumnName(deleted) }
	}
	const start = node.column
	const end = node.endColumn
	if (!start || !end) {
		const column = start ?? end
		if (column && deleted.deletedNames.has(column.toLowerCase())) {
			return { table: deleted.table, columnName: column }
		}
		return null
	}
	const startIndex = deleted.indexByName.get(start.toLowerCase())
	const endIndex = deleted.indexByName.get(end.toLowerCase())
	if (startIndex === undefined || endIndex === undefined) {
		if (deleted.deletedNames.has(start.toLowerCase()))
			return { table: deleted.table, columnName: start }
		if (deleted.deletedNames.has(end.toLowerCase()))
			return { table: deleted.table, columnName: end }
		return null
	}
	const min = Math.min(startIndex, endIndex)
	const max = Math.max(startIndex, endIndex)
	for (const index of deleted.deletedIndexes) {
		if (index >= min && index <= max) {
			return { table: deleted.table, columnName: deleted.table.columns[index]?.name ?? start }
		}
	}
	return null
}

function firstDeletedColumnName(deleted: DeletedTableColumnSet): string {
	const [index] = [...deleted.deletedIndexes].sort((a, b) => a - b)
	return deleted.table.columns[index ?? 0]?.name ?? deleted.table.name
}

function tableContainingCell(sheet: Sheet, row: number, col: number): Table | undefined {
	return sheet.tables.find(
		(table) =>
			row >= table.ref.start.row &&
			row <= table.ref.end.row &&
			col >= table.ref.start.col &&
			col <= table.ref.end.col,
	)
}
