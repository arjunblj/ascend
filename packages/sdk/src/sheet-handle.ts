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
import type { CellValue } from '@ascend/schema'
import type {
	CellInfo,
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
} from './types.ts'

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
		const compact = this.cellCompact(ref)
		return compact ? toCellInfo(compact, ref) : undefined
	}

	cellCompact(ref: string): CompactCellInfo | undefined {
		const parsed = parseA1(ref)
		const cell = this.sheet.cells.get(parsed.row, parsed.col)
		if (!cell) return undefined
		return makeCompactCellInfo(parsed.row, parsed.col, cell, ref)
	}

	range(rangeRef: string): RangeInfo {
		const compact = this.rangeCompact(rangeRef, { includeRefs: true })
		return {
			ref: compact.ref,
			cells: compact.cells.map((cell) => toCellInfo(cell)),
			rowCount: compact.rowCount,
			colCount: compact.colCount,
		}
	}

	rangeCompact(rangeRef: string, opts?: { includeRefs?: boolean }): CompactRangeInfo {
		const parsed = parseRange(rangeRef)
		const cells = collectCellsCompact(this.sheet, parsed, opts)
		return {
			ref: parsed,
			cells,
			rowCount: parsed.end.row - parsed.start.row + 1,
			colCount: parsed.end.col - parsed.start.col + 1,
		}
	}

	readWindow(rangeRef: string, opts?: { rowOffset?: number; rowLimit?: number }): RangeWindowInfo {
		const compact = this.readWindowCompact(rangeRef, { ...opts, includeRefs: true })
		return {
			requestedRef: compact.requestedRef,
			ref: compact.ref,
			cells: compact.cells.map((cell) => toCellInfo(cell)),
			rowCount: compact.rowCount,
			colCount: compact.colCount,
			rowOffset: compact.rowOffset,
			rowLimit: compact.rowLimit,
			hasMore: compact.hasMore,
			...(compact.nextRowOffset !== undefined ? { nextRowOffset: compact.nextRowOffset } : {}),
		}
	}

	readWindowCompact(
		rangeRef: string,
		opts?: { rowOffset?: number; rowLimit?: number; includeRefs?: boolean },
	): CompactRangeWindowInfo {
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
		const cells = collectCellsCompact(this.sheet, windowRef, opts)
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

	readRows(rangeRef: string, opts?: { rowOffset?: number; rowLimit?: number }): RangeRowsInfo {
		const window = this.readWindowCompact(rangeRef, {
			...(opts?.rowOffset !== undefined ? { rowOffset: opts.rowOffset } : {}),
			...(opts?.rowLimit !== undefined ? { rowLimit: opts.rowLimit } : {}),
			includeRefs: false,
		})
		return {
			requestedRef: window.requestedRef,
			ref: window.ref,
			rowCount: window.rowCount,
			colCount: window.colCount,
			rowOffset: window.rowOffset,
			rowLimit: window.rowLimit,
			hasMore: window.hasMore,
			...(window.nextRowOffset !== undefined ? { nextRowOffset: window.nextRowOffset } : {}),
			rows: buildValueRows(
				window.cells,
				window.rowCount,
				window.colCount,
				window.ref.start.row,
				window.ref.start.col,
			),
		}
	}

	readObjects(
		rangeRef: string,
		opts?: { rowOffset?: number; rowLimit?: number; headers?: readonly string[] | 'first-row' },
	): RangeObjectsInfo {
		const rowsInfo = this.readRows(rangeRef, opts)
		const useFirstRow = opts?.headers === undefined || opts.headers === 'first-row'
		const sourceRows = rowsInfo.rows
		const headerValues = useFirstRow ? (sourceRows[0] ?? []) : undefined
		const headers = useFirstRow
			? Array.from({ length: rowsInfo.colCount }, (_, index) =>
					normalizeObjectHeader(headerValues?.[index], index),
				)
			: [...(opts?.headers ?? [])]
		const dataRows = useFirstRow ? sourceRows.slice(1) : sourceRows
		return {
			requestedRef: rowsInfo.requestedRef,
			ref: rowsInfo.ref,
			rowCount: useFirstRow ? Math.max(0, rowsInfo.rowCount - 1) : rowsInfo.rowCount,
			colCount: rowsInfo.colCount,
			rowOffset: rowsInfo.rowOffset,
			rowLimit: rowsInfo.rowLimit,
			hasMore: rowsInfo.hasMore,
			...(rowsInfo.nextRowOffset !== undefined ? { nextRowOffset: rowsInfo.nextRowOffset } : {}),
			headers,
			rows: dataRows.map((row) => {
				const objectRow: Record<string, import('@ascend/schema').CellValue> = {}
				for (let index = 0; index < headers.length; index++) {
					const header = headers[index]
					if (!header) continue
					objectRow[header] = row[index] ?? { kind: 'empty' }
				}
				return objectRow
			}),
		}
	}

	*streamRange(rangeRef: string): Generator<readonly CellInfo[]> {
		for (const row of this.streamRangeCompact(rangeRef, { includeRefs: true })) {
			yield row.map((cell) => toCellInfo(cell))
		}
	}

	*streamRangeCompact(
		rangeRef: string,
		opts?: { includeRefs?: boolean },
	): Generator<readonly CompactCellInfo[]> {
		const parsed = parseRange(rangeRef)
		const rows = this.sheet.cells.iterateRowsInRange(parsed)
		let next = rows.next()
		for (let row = parsed.start.row; row <= parsed.end.row; row++) {
			if (!next.done && next.value[0] === row) {
				yield next.value[1].map(([col, cell]) =>
					makeCompactCellInfo(
						row,
						col,
						cell,
						opts?.includeRefs === false ? undefined : toA1({ row, col }),
					),
				)
				next = rows.next()
			} else {
				yield []
			}
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

function collectCellsCompact(
	sheet: Sheet,
	range: RangeRef,
	opts?: { includeRefs?: boolean },
): CompactCellInfo[] {
	const cells: CompactCellInfo[] = []
	for (const [row, rowCells] of sheet.cells.iterateRowsInRange(range)) {
		for (const [col, cell] of rowCells) {
			cells.push(
				makeCompactCellInfo(
					row,
					col,
					cell,
					opts?.includeRefs === false ? undefined : toA1({ row, col }),
				),
			)
		}
	}
	return cells
}

function makeCompactCellInfo(
	row: number,
	col: number,
	cell: NonNullable<ReturnType<Sheet['cells']['get']>>,
	ref?: string,
): CompactCellInfo {
	return {
		...(ref ? { ref } : {}),
		value: cell.value,
		formula: cell.formula,
		formulaBinding: cell.formulaInfo ?? null,
		row,
		col,
	}
}

function toCellInfo(cell: CompactCellInfo, explicitRef?: string): CellInfo {
	const ref = explicitRef ?? cell.ref
	if (!ref) {
		throw new Error('CellInfo conversion requires a reference')
	}
	return {
		ref,
		value: cell.value,
		formula: cell.formula,
		...(cell.formulaBinding ? { formulaBinding: cell.formulaBinding } : {}),
		row: cell.row,
		col: cell.col,
	}
}

function buildValueRows(
	cells: readonly CompactCellInfo[],
	rowCount: number,
	colCount: number,
	startRow: number,
	startCol: number,
): CellValue[][] {
	const rows: CellValue[][] = []
	let index = 0
	for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
		const rowIndex = startRow + rowOffset
		const values: CellValue[] = []
		for (let colOffset = 0; colOffset < colCount; colOffset++) {
			const colIndex = startCol + colOffset
			const cell = cells[index]
			if (cell && cell.row === rowIndex && cell.col === colIndex) {
				values.push(cell.value)
				index += 1
			} else {
				values.push({ kind: 'empty' })
			}
		}
		rows.push(values)
	}
	return rows
}

function normalizeObjectHeader(value: CellValue | undefined, index: number): string {
	if (!value) return `Column${index + 1}`
	switch (value.kind) {
		case 'string':
			return value.value || `Column${index + 1}`
		case 'number':
			return String(value.value)
		case 'boolean':
			return value.value ? 'TRUE' : 'FALSE'
		case 'date':
			return String(value.serial)
		case 'error':
			return value.value
		case 'richText': {
			const text = value.runs.map((run: { text: string }) => run.text).join('')
			return text || `Column${index + 1}`
		}
		case 'empty':
			return `Column${index + 1}`
	}
	return `Column${index + 1}`
}
