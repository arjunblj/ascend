import { createSheetId, type SheetId } from './ids.ts'
import type { RangeRef } from './refs.ts'
import { SparseGrid } from './sparse-grid.ts'
import type { Table } from './table.ts'

export type SheetState = 'visible' | 'hidden' | 'veryHidden'

export interface SheetComment {
	readonly text: string
	readonly author?: string
}

export class Sheet {
	readonly id: SheetId
	name: string
	readonly cells: SparseGrid
	readonly merges: RangeRef[]
	readonly tables: Table[]
	state: SheetState
	readonly colWidths: Map<number, number>
	readonly rowHeights: Map<number, number>
	frozenRows: number
	frozenCols: number
	readonly comments: Map<string, SheetComment>

	constructor(name: string, id?: SheetId) {
		this.id = id ?? createSheetId()
		this.name = name
		this.cells = new SparseGrid()
		this.merges = []
		this.tables = []
		this.state = 'visible'
		this.colWidths = new Map()
		this.rowHeights = new Map()
		this.frozenRows = 0
		this.frozenCols = 0
		this.comments = new Map()
	}
}

export function createSheet(name: string, id?: SheetId): Sheet {
	return new Sheet(name, id)
}
