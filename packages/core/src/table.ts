import type { AutoFilter } from './filter.ts'
import type { SheetId, TableId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface TableColumn {
	readonly name: string
	readonly formula?: string
}

export interface Table {
	readonly id: TableId
	readonly name: string
	readonly sheetId: SheetId
	readonly ref: RangeRef
	readonly columns: readonly TableColumn[]
	readonly hasHeaders: boolean
	readonly hasTotals: boolean
	readonly autoFilter?: AutoFilter
}
