import type { RangeRef, Sheet } from '@ascend/core'
import { parseA1, parseRange, toA1 } from '@ascend/core'
import type { CellInfo, RangeInfo, RangeWindowInfo } from './types.ts'

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
		const cells = collectCells(this.sheet, parsed)
		return {
			ref: parsed,
			cells,
			rowCount: parsed.end.row - parsed.start.row + 1,
			colCount: parsed.end.col - parsed.start.col + 1,
		}
	}

	readWindow(rangeRef: string, opts?: { rowOffset?: number; rowLimit?: number }): RangeWindowInfo {
		const requestedRef = parseRange(rangeRef)
		const rowOffset = Math.max(0, opts?.rowOffset ?? 0)
		const totalRows = requestedRef.end.row - requestedRef.start.row + 1
		const defaultLimit = totalRows
		const rowLimit = Math.max(1, opts?.rowLimit ?? defaultLimit)
		const startRow = requestedRef.start.row + rowOffset
		const endRow = Math.min(requestedRef.end.row, startRow + rowLimit - 1)
		const windowRef: RangeRef = {
			...requestedRef,
			start: { ...requestedRef.start, row: Math.min(startRow, requestedRef.end.row) },
			end: {
				...requestedRef.end,
				row: Math.max(Math.min(endRow, requestedRef.end.row), requestedRef.start.row),
			},
		}
		const cells = collectCells(this.sheet, windowRef)
		const consumedRows = Math.max(0, endRow - requestedRef.start.row + 1)
		const hasMore = requestedRef.start.row + rowOffset + rowLimit - 1 < requestedRef.end.row
		return {
			requestedRef,
			ref: windowRef,
			cells,
			rowCount: Math.max(0, windowRef.end.row - windowRef.start.row + 1),
			colCount: requestedRef.end.col - requestedRef.start.col + 1,
			rowOffset,
			rowLimit,
			hasMore,
			...(hasMore ? { nextRowOffset: consumedRows } : {}),
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

function collectCells(sheet: Sheet, range: RangeRef): CellInfo[] {
	const cells: CellInfo[] = []
	for (let r = range.start.row; r <= range.end.row; r++) {
		for (let c = range.start.col; c <= range.end.col; c++) {
			const cell = sheet.cells.get(r, c)
			if (!cell) continue
			cells.push({
				ref: toA1({ row: r, col: c }),
				value: cell.value,
				formula: cell.formula,
				row: r,
				col: c,
			})
		}
	}
	return cells
}
