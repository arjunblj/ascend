import type {
	AutoFilter,
	RangeRef,
	Sheet,
	SheetComment,
	SheetConditionalFormat,
	SheetDataValidation,
	SheetFormatPr,
	SheetHyperlink,
	SheetImageRef,
	SheetProtection,
	SheetTabColor,
} from '@ascend/core'
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
			...(cell.formulaInfo ? { formulaBinding: cell.formulaInfo } : {}),
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
		const rowMap = new Map<number, CellInfo[]>()
		for (const [row, rowCells] of this.sheet.cells.iterateRowsInRange(parsed)) {
			rowMap.set(
				row,
				rowCells.map(([col, cell]) => ({
					ref: toA1({ row, col }),
					value: cell.value,
					formula: cell.formula,
					...(cell.formulaInfo ? { formulaBinding: cell.formulaInfo } : {}),
					row,
					col,
				})),
			)
		}
		for (let row = parsed.start.row; row <= parsed.end.row; row++) {
			yield rowMap.get(row) ?? []
		}
	}

	usedRange(): RangeRef | null {
		return this.sheet.cells.usedRange()
	}

	get state(): string {
		return this.sheet.state
	}

	get tabColor(): SheetTabColor | null {
		return this.sheet.tabColor
	}

	get sheetFormatPr(): SheetFormatPr | null {
		return this.sheet.sheetFormatPr
	}

	get frozenRows(): number {
		return this.sheet.frozenRows
	}

	get frozenCols(): number {
		return this.sheet.frozenCols
	}

	get merges(): readonly RangeRef[] {
		return this.sheet.merges
	}

	get autoFilter(): AutoFilter | null {
		return this.sheet.autoFilter
	}

	get protection(): SheetProtection | null {
		return this.sheet.protection
	}

	get conditionalFormats(): readonly SheetConditionalFormat[] {
		return this.sheet.conditionalFormats
	}

	get dataValidations(): readonly SheetDataValidation[] {
		return this.sheet.dataValidations
	}

	get imageRefs(): readonly SheetImageRef[] {
		return this.sheet.imageRefs
	}

	comments(): ReadonlyMap<string, SheetComment> {
		return this.sheet.comments
	}

	hyperlinks(): ReadonlyMap<string, SheetHyperlink> {
		return this.sheet.hyperlinks
	}

	comment(ref: string): SheetComment | undefined {
		return this.sheet.comments.get(ref)
	}

	hyperlink(ref: string): SheetHyperlink | undefined {
		return this.sheet.hyperlinks.get(ref)
	}
}

function collectCells(sheet: Sheet, range: RangeRef): CellInfo[] {
	const cells: CellInfo[] = []
	for (const [row, rowCells] of sheet.cells.iterateRowsInRange(range)) {
		for (const [col, cell] of rowCells) {
			cells.push({
				ref: toA1({ row, col }),
				value: cell.value,
				formula: cell.formula,
				...(cell.formulaInfo ? { formulaBinding: cell.formulaInfo } : {}),
				row,
				col,
			})
		}
	}
	return cells
}
