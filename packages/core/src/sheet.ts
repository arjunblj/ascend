import { createSheetId, type SheetId } from './ids.ts'
import type { RangeRef } from './refs.ts'
import { SparseGrid } from './sparse-grid.ts'
import type { Table } from './table.ts'

export type SheetState = 'visible' | 'hidden' | 'veryHidden'

export class Sheet {
	readonly id: SheetId
	name: string
	readonly cells: SparseGrid
	readonly merges: RangeRef[]
	readonly tables: Table[]
	state: SheetState

	constructor(name: string, id?: SheetId) {
		this.id = id ?? createSheetId()
		this.name = name
		this.cells = new SparseGrid()
		this.merges = []
		this.tables = []
		this.state = 'visible'
	}
}

export function createSheet(name: string, id?: SheetId): Sheet {
	return new Sheet(name, id)
}
