import type { AutoFilter, RangeRef, Sheet, Table, TableColumn, TableStyleInfo } from '@ascend/core'
import type { CellValue } from '@ascend/schema'

export class TableHandle {
	private readonly table: Table
	private readonly sheet: Sheet

	constructor(table: Table, sheet: Sheet) {
		this.table = table
		this.sheet = sheet
	}

	get name(): string {
		return this.table.name
	}

	get columns(): readonly string[] {
		return this.table.columns.map((c) => c.name)
	}

	get rowCount(): number {
		const headerOffset = this.table.hasHeaders ? 1 : 0
		const totalOffset = this.table.hasTotals ? 1 : 0
		return this.table.ref.end.row - this.table.ref.start.row + 1 - headerOffset - totalOffset
	}

	get ref(): RangeRef {
		return this.table.ref
	}

	get hasHeaders(): boolean {
		return this.table.hasHeaders
	}

	get hasTotals(): boolean {
		return this.table.hasTotals
	}

	get styleInfo(): TableStyleInfo | undefined {
		return this.table.tableStyleInfo
	}

	get autoFilter(): AutoFilter | null {
		return this.table.autoFilter ?? null
	}

	get sortState() {
		return this.table.sortState ?? null
	}

	get columnDefs(): readonly TableColumn[] {
		return this.table.columns
	}

	rows(opts?: { offset?: number; limit?: number }): readonly Record<string, CellValue>[] {
		const headerOffset = this.table.hasHeaders ? 1 : 0
		const dataStartRow = this.table.ref.start.row + headerOffset
		const totalOffset = this.table.hasTotals ? 1 : 0
		const dataEndRow = this.table.ref.end.row - totalOffset
		const colNames = this.table.columns.map((c) => c.name)

		const result: Record<string, CellValue>[] = []
		const limit = opts?.limit ?? Number.POSITIVE_INFINITY
		const offset = Math.max(0, opts?.offset ?? 0)
		let seen = 0

		for (let r = dataStartRow; r <= dataEndRow && result.length < limit; r++) {
			if (seen++ < offset) continue
			const row: Record<string, CellValue> = {}
			for (let c = 0; c < colNames.length; c++) {
				const colName = colNames[c]
				if (!colName) continue
				const cell = this.sheet.cells.get(r, this.table.ref.start.col + c)
				row[colName] = cell?.value ?? { kind: 'empty' }
			}
			result.push(row)
		}

		return result
	}

	headerRow(): readonly CellValue[] | null {
		if (!this.table.hasHeaders) return null
		const values: CellValue[] = []
		for (let col = this.table.ref.start.col; col <= this.table.ref.end.col; col++) {
			values.push(this.sheet.cells.get(this.table.ref.start.row, col)?.value ?? { kind: 'empty' })
		}
		return values
	}

	totalsRow(): readonly CellValue[] | null {
		if (!this.table.hasTotals) return null
		const rowIndex = this.table.ref.end.row
		const values: CellValue[] = []
		for (let col = this.table.ref.start.col; col <= this.table.ref.end.col; col++) {
			values.push(this.sheet.cells.get(rowIndex, col)?.value ?? { kind: 'empty' })
		}
		return values
	}
}
