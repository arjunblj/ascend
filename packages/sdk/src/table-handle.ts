import type { Sheet, Table } from '@ascend/core'
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

	rows(opts?: { limit?: number }): readonly Record<string, CellValue>[] {
		const headerOffset = this.table.hasHeaders ? 1 : 0
		const dataStartRow = this.table.ref.start.row + headerOffset
		const totalOffset = this.table.hasTotals ? 1 : 0
		const dataEndRow = this.table.ref.end.row - totalOffset
		const colNames = this.table.columns.map((c) => c.name)

		const result: Record<string, CellValue>[] = []
		const limit = opts?.limit ?? Number.POSITIVE_INFINITY

		for (let r = dataStartRow; r <= dataEndRow && result.length < limit; r++) {
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
}
