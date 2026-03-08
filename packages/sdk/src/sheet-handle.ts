import type { RangeRef, Sheet } from '@ascend/core'
import { parseA1, parseRange, toA1 } from '@ascend/core'
import type { CellInfo, RangeInfo } from './types.ts'

export class SheetHandle {
	private readonly sheet: Sheet

	constructor(sheet: Sheet) {
		this.sheet = sheet
	}

	get name(): string {
		return this.sheet.name
	}

	get rowCount(): number {
		const used = this.sheet.cells.usedRange()
		return used ? used.end.row + 1 : 0
	}

	get colCount(): number {
		const used = this.sheet.cells.usedRange()
		return used ? used.end.col + 1 : 0
	}

	cell(ref: string): CellInfo | undefined {
		const parsed = parseA1(ref)
		const cell = this.sheet.cells.get(parsed.row, parsed.col)
		if (!cell) return undefined
		return {
			ref,
			value: cell.value,
			formula: cell.formula,
			row: parsed.row,
			col: parsed.col,
		}
	}

	range(rangeRef: string): RangeInfo {
		const parsed = parseRange(rangeRef)
		const cells: CellInfo[] = []

		for (let r = parsed.start.row; r <= parsed.end.row; r++) {
			for (let c = parsed.start.col; c <= parsed.end.col; c++) {
				const cell = this.sheet.cells.get(r, c)
				if (cell) {
					cells.push({
						ref: toA1({ row: r, col: c }),
						value: cell.value,
						formula: cell.formula,
						row: r,
						col: c,
					})
				}
			}
		}

		return {
			ref: parsed,
			cells,
			rowCount: parsed.end.row - parsed.start.row + 1,
			colCount: parsed.end.col - parsed.start.col + 1,
		}
	}

	*streamRange(rangeRef: string): Generator<readonly CellInfo[]> {
		const parsed = parseRange(rangeRef)
		for (let r = parsed.start.row; r <= parsed.end.row; r++) {
			const row: CellInfo[] = []
			for (let c = parsed.start.col; c <= parsed.end.col; c++) {
				const cell = this.sheet.cells.get(r, c)
				if (!cell) continue
				row.push({
					ref: toA1({ row: r, col: c }),
					value: cell.value,
					formula: cell.formula,
					row: r,
					col: c,
				})
			}
			yield row
		}
	}

	usedRange(): RangeRef | null {
		return this.sheet.cells.usedRange()
	}
}
