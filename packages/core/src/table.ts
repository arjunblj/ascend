import type { AutoFilter } from './filter.ts'
import type { SheetId, TableId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface TableColumn {
	readonly id?: number
	readonly name: string
	readonly formula?: string
	readonly totalsRowFunction?: string
	readonly totalsRowLabel?: string
	readonly dataDxfId?: number
	readonly headerRowDxfId?: number
	readonly totalsRowDxfId?: number
}

export interface TableStyleInfo {
	readonly name?: string
	readonly showFirstColumn?: boolean
	readonly showLastColumn?: boolean
	readonly showRowStripes?: boolean
	readonly showColumnStripes?: boolean
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
	readonly headerRowDxfId?: number
	readonly dataDxfId?: number
	readonly totalsRowDxfId?: number
	readonly headerRowBorderDxfId?: number
	readonly tableStyleInfo?: TableStyleInfo
}
