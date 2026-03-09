import { createSheetId, type SheetId } from './ids.ts'
import type { RangeRef } from './refs.ts'
import { SparseGrid } from './sparse-grid.ts'
import type { Table } from './table.ts'

export type SheetState = 'visible' | 'hidden' | 'veryHidden'

export interface SheetComment {
	readonly text: string
	readonly author?: string
}

export interface SheetHyperlink {
	readonly target?: string
	readonly location?: string
	readonly display?: string
	readonly tooltip?: string
}

export interface SheetPageMargins {
	readonly left?: number
	readonly right?: number
	readonly top?: number
	readonly bottom?: number
	readonly header?: number
	readonly footer?: number
}

export interface SheetPageSetup {
	readonly orientation?: string
	readonly paperSize?: number
	readonly scale?: number
	readonly fitToWidth?: number
	readonly fitToHeight?: number
}

export interface SheetPrintOptions {
	readonly gridLines?: boolean
	readonly headings?: boolean
	readonly horizontalCentered?: boolean
	readonly verticalCentered?: boolean
}

export interface SheetHeaderFooter {
	readonly oddHeader?: string
	readonly oddFooter?: string
	readonly evenHeader?: string
	readonly evenFooter?: string
	readonly firstHeader?: string
	readonly firstFooter?: string
}

export interface SheetColDef {
	readonly min: number
	readonly max: number
	readonly width?: number
	readonly style?: number
	readonly hidden?: boolean
	readonly bestFit?: boolean
	readonly collapsed?: boolean
	readonly outlineLevel?: number
	readonly customWidth?: boolean
}

export interface SheetPreservedXml {
	readonly xml: string
	readonly relsXml?: string
}

export class Sheet {
	readonly id: SheetId
	name: string
	readonly cells: SparseGrid
	readonly merges: RangeRef[]
	readonly tables: Table[]
	state: SheetState
	readonly colWidths: Map<number, number>
	readonly colDefs: SheetColDef[]
	readonly rowHeights: Map<number, number>
	frozenRows: number
	frozenCols: number
	readonly comments: Map<string, SheetComment>
	readonly hyperlinks: Map<string, SheetHyperlink>
	readonly ignoredErrors: string[]
	autoFilter: string | null
	pageMargins: SheetPageMargins | null
	pageSetup: SheetPageSetup | null
	printOptions: SheetPrintOptions | null
	headerFooter: SheetHeaderFooter | null
	preservedXml: SheetPreservedXml | null

	constructor(name: string, id?: SheetId) {
		this.id = id ?? createSheetId()
		this.name = name
		this.cells = new SparseGrid()
		this.merges = []
		this.tables = []
		this.state = 'visible'
		this.colWidths = new Map()
		this.colDefs = []
		this.rowHeights = new Map()
		this.frozenRows = 0
		this.frozenCols = 0
		this.comments = new Map()
		this.hyperlinks = new Map()
		this.ignoredErrors = []
		this.autoFilter = null
		this.pageMargins = null
		this.pageSetup = null
		this.printOptions = null
		this.headerFooter = null
		this.preservedXml = null
	}
}

export function createSheet(name: string, id?: SheetId): Sheet {
	return new Sheet(name, id)
}
