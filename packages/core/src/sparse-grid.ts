import type { CellValue } from '@ascend/schema'
import type { StyleId } from './ids.ts'
import type { RangeRef } from './refs.ts'

export interface Cell {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
}

const PACK_FACTOR = 1 << 20

function packKey(row: number, col: number): number {
	return row * PACK_FACTOR + col
}

function unpackKey(key: number): readonly [number, number] {
	const col = key % PACK_FACTOR
	const row = (key - col) / PACK_FACTOR
	return [row, col] as const
}

export class SparseGrid {
	private readonly data = new Map<number, Cell>()

	get(row: number, col: number): Cell | undefined {
		return this.data.get(packKey(row, col))
	}

	set(row: number, col: number, cell: Cell): void {
		this.data.set(packKey(row, col), cell)
	}

	delete(row: number, col: number): boolean {
		return this.data.delete(packKey(row, col))
	}

	*getRange(range: RangeRef): Generator<readonly [number, number, Cell]> {
		for (let r = range.start.row; r <= range.end.row; r++) {
			for (let c = range.start.col; c <= range.end.col; c++) {
				const cell = this.data.get(packKey(r, c))
				if (cell) yield [r, c, cell] as const
			}
		}
	}

	usedRange(): RangeRef | null {
		if (this.data.size === 0) return null

		let minRow = Number.POSITIVE_INFINITY
		let maxRow = Number.NEGATIVE_INFINITY
		let minCol = Number.POSITIVE_INFINITY
		let maxCol = Number.NEGATIVE_INFINITY

		for (const key of this.data.keys()) {
			const [row, col] = unpackKey(key)
			if (row < minRow) minRow = row
			if (row > maxRow) maxRow = row
			if (col < minCol) minCol = col
			if (col > maxCol) maxCol = col
		}

		return {
			start: { row: minRow, col: minCol },
			end: { row: maxRow, col: maxCol },
		}
	}

	*iterate(): Generator<readonly [number, number, Cell]> {
		for (const [key, cell] of this.data) {
			const [row, col] = unpackKey(key)
			yield [row, col, cell] as const
		}
	}

	cellCount(): number {
		return this.data.size
	}

	clear(): void {
		this.data.clear()
	}
}
