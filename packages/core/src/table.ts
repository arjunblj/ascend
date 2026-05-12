import type { AutoFilter, SortState } from './filter.ts'
import type { SheetId, TableId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface TableColumn {
	readonly id?: number
	readonly uid?: string
	readonly uniqueName?: string
	readonly name: string
	readonly formula?: string
	readonly formulaIsArray?: boolean
	readonly xmlColumnPr?: TableXmlColumnPr
	readonly totalsRowFunction?: string
	readonly totalsRowFormula?: string
	readonly totalsRowLabel?: string
	readonly queryTableFieldId?: number
	readonly dataCellStyle?: string
	readonly dataDxfId?: number
	readonly headerRowDxfId?: number
	readonly totalsRowDxfId?: number
}

export interface TableXmlColumnPr {
	readonly mapId?: number
	readonly xpath?: string
	readonly xmlDataType?: string
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
	readonly nameAttribute?: string | null
	readonly sheetId: SheetId
	readonly partPath?: string
	readonly contentType?: string
	readonly contentTypeSource?: 'override' | 'default' | 'fallback'
	readonly sourcePartPath?: string
	readonly sourceRelationshipPart?: string
	readonly sourceRelationshipId?: string
	readonly sourceRelationshipType?: string
	readonly sourceRelationshipRawType?: string
	readonly sourceRelationshipRawTarget?: string
	readonly sourceRelationshipResolvedTarget?: string
	readonly sourceRelationshipTargetMode?: string
	readonly uid?: string
	readonly ref: RangeRef
	readonly tableType?: string
	readonly insertRow?: boolean
	readonly insertRowShift?: boolean
	readonly columns: readonly TableColumn[]
	readonly hasHeaders: boolean
	readonly hasTotals: boolean
	readonly altText?: string
	readonly altTextSummary?: string
	readonly autoFilter?: AutoFilter
	readonly sortState?: SortState
	readonly dxfId?: number
	readonly dataCellStyle?: string
	readonly headerRowDxfId?: number
	readonly headerRowCellStyle?: string
	readonly dataDxfId?: number
	readonly totalsRowDxfId?: number
	readonly headerRowBorderDxfId?: number
	readonly tableBorderDxfId?: number
	readonly tableStyleInfo?: TableStyleInfo
	readonly queryTable?: TableQueryTableRef
}

export interface TableQueryTableRef {
	readonly relationshipId: string
	readonly partPath: string
	readonly relationshipType: string
	readonly target: string
	readonly targetMode?: string
}
