import type { AutoFilter, RangeRef, Sheet, Table, TableColumn, TableStyleInfo } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import type { TableWindowInfo } from './types.ts'

export class TableHandle {
	private readonly tableName: string
	private readonly resolveTable: () => { readonly table: Table; readonly sheet: Sheet } | undefined

	constructor(
		tableName: string,
		resolveTable: () => { readonly table: Table; readonly sheet: Sheet } | undefined,
	) {
		this.tableName = tableName
		this.resolveTable = resolveTable
	}

	get name(): string {
		return this.tableName
	}

	get columns(): readonly string[] {
		return this.requireTable().table.columns.map((c) => c.name)
	}

	get rowCount(): number {
		const { table } = this.requireTable()
		const headerOffset = table.hasHeaders ? 1 : 0
		const totalOffset = table.hasTotals ? 1 : 0
		return table.ref.end.row - table.ref.start.row + 1 - headerOffset - totalOffset
	}

	get ref(): RangeRef {
		return this.requireTable().table.ref
	}

	get hasHeaders(): boolean {
		return this.requireTable().table.hasHeaders
	}

	get hasTotals(): boolean {
		return this.requireTable().table.hasTotals
	}

	get styleInfo(): TableStyleInfo | undefined {
		return this.requireTable().table.tableStyleInfo
	}

	get autoFilter(): AutoFilter | null {
		return this.requireTable().table.autoFilter ?? null
	}

	get sortState() {
		return this.requireTable().table.sortState ?? null
	}

	get columnDefs(): readonly TableColumn[] {
		return this.requireTable().table.columns
	}

	rows(opts?: { offset?: number; limit?: number }): readonly Record<string, CellValue>[] {
		return this.readRows(opts).rows.map((row) => row.values)
	}

	readRows(opts?: { offset?: number; limit?: number }): TableWindowInfo {
		const { table, sheet } = this.requireTable()
		const headerOffset = table.hasHeaders ? 1 : 0
		const dataStartRow = table.ref.start.row + headerOffset
		const totalOffset = table.hasTotals ? 1 : 0
		const dataEndRow = table.ref.end.row - totalOffset
		const colNames = table.columns.map((c) => c.name)

		const limit = opts?.limit ?? Number.POSITIVE_INFINITY
		const offset = Math.max(0, opts?.offset ?? 0)
		const totalRows = Math.max(0, dataEndRow - dataStartRow + 1)
		const startRow = Math.min(dataEndRow + 1, dataStartRow + offset)
		const endRow = Math.min(dataEndRow, startRow + limit - 1)
		const rows: Array<{ index: number; sheetRow: number; values: Record<string, CellValue> }> = []

		for (let r = startRow; r <= endRow; r++) {
			const row: Record<string, CellValue> = {}
			for (let c = 0; c < colNames.length; c++) {
				const colName = colNames[c]
				if (!colName) continue
				const cell = sheet.cells.get(r, table.ref.start.col + c)
				row[colName] = cell?.value ?? { kind: 'empty' }
			}
			rows.push({
				index: r - dataStartRow,
				sheetRow: r,
				values: row,
			})
		}

		const returnedRows = rows.length
		const nextRowOffset = offset + returnedRows
		return {
			rowOffset: offset,
			rowLimit: Number.isFinite(limit) ? limit : totalRows,
			returnedRows,
			totalRows,
			hasMore: nextRowOffset < totalRows,
			...(nextRowOffset < totalRows ? { nextRowOffset } : {}),
			rows,
		}
	}

	headerRow(): readonly CellValue[] | null {
		const { table, sheet } = this.requireTable()
		if (!table.hasHeaders) return null
		const values: CellValue[] = []
		for (let col = table.ref.start.col; col <= table.ref.end.col; col++) {
			values.push(sheet.cells.get(table.ref.start.row, col)?.value ?? { kind: 'empty' })
		}
		return values
	}

	totalsRow(): readonly CellValue[] | null {
		const { table, sheet } = this.requireTable()
		if (!table.hasTotals) return null
		const rowIndex = table.ref.end.row
		const values: CellValue[] = []
		for (let col = table.ref.start.col; col <= table.ref.end.col; col++) {
			values.push(sheet.cells.get(rowIndex, col)?.value ?? { kind: 'empty' })
		}
		return values
	}

	private requireTable(): { readonly table: Table; readonly sheet: Sheet } {
		const resolved = this.resolveTable()
		if (resolved) return resolved
		throw new Error(
			`Table "${this.tableName}" is no longer available in the current workbook view.`,
		)
	}
}
