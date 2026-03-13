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
import { AscendException, ascendError, type CellValue } from '@ascend/schema'
import type {
	AgentReadOptions,
	CellInfo,
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	FlatCellValue,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
} from './types.ts'

export class SheetHandle {
	private readonly sheetName: string
	private readonly resolveSheet: () => Sheet | undefined
	private readonly resolveFormula: (
		row: number,
		col: number,
		cell: NonNullable<ReturnType<Sheet['cells']['get']>>,
	) => string | null
	private _changeVersion = 0
	private readonly _changeSnapshots = new Map<
		string,
		{ token: string; cells: Map<string, CompactCellInfo> }
	>()

	constructor(
		sheetName: string,
		resolveSheet: () => Sheet | undefined,
		resolveFormula: (
			row: number,
			col: number,
			cell: NonNullable<ReturnType<Sheet['cells']['get']>>,
		) => string | null,
	) {
		this.sheetName = sheetName
		this.resolveSheet = resolveSheet
		this.resolveFormula = resolveFormula
	}

	get name(): string {
		return this.sheetName
	}

	get rowCount(): number {
		const used = this.requireSheet().cells.usedRange()
		return used ? used.end.row + 1 : 0
	}

	get colCount(): number {
		const used = this.requireSheet().cells.usedRange()
		return used ? used.end.col + 1 : 0
	}

	cell(ref: string): CellInfo | undefined {
		const compact = this.cellCompact(ref)
		return compact ? toCellInfo(compact, ref) : undefined
	}

	cellCompact(ref: string): CompactCellInfo | undefined {
		const parsed = parseA1(ref)
		const cell = this.requireSheet().cells.get(parsed.row, parsed.col)
		if (!cell) return undefined
		return makeCompactCellInfo(
			parsed.row,
			parsed.col,
			cell,
			this.resolveFormula(parsed.row, parsed.col, cell),
			ref,
		)
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

	rangeCompact(
		rangeRef: string,
		opts?: { includeRefs?: boolean; omitEmpty?: boolean; flatValues?: boolean },
	): CompactRangeInfo {
		const parsed = parseRange(rangeRef)
		let cells = collectCellsCompact(this.requireSheet(), parsed, this.resolveFormula, opts)
		if (opts?.omitEmpty) {
			cells = cells.filter((c) => c.value.kind !== 'empty')
		}
		if (opts?.flatValues) {
			cells = cells.map((c) => ({
				...c,
				value: flattenCellValue(c.value) as unknown as CellValue,
			}))
		}
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

	readWindowCompact(rangeRef: string, opts?: AgentReadOptions): CompactRangeWindowInfo {
		const requestedRef = parseRange(rangeRef)
		const sheet = this.requireSheet()
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
		let cells = collectCellsCompact(sheet, windowRef, this.resolveFormula, opts)
		if (opts?.omitEmpty) {
			cells = cells.filter((c) => c.value.kind !== 'empty')
		}
		if (opts?.flatValues) {
			cells = cells.map((c) => ({
				...c,
				value: flattenCellValue(c.value) as unknown as CellValue,
			}))
		}
		const consumedRows = Math.max(0, endRow - requestedRef.start.row + 1)
		const hasMore = requestedRef.start.row + rowOffset + rowLimit - 1 < requestedRef.end.row
		const snapshotKey = opts?.changedSince !== undefined ? `${this.sheetName}:${rangeRef}` : null
		let changeToken: string | undefined
		if (snapshotKey) {
			const version = this._changeVersion++
			changeToken = `${version}`
			const previous = this._changeSnapshots.get(snapshotKey)
			if (previous && opts?.changedSince === previous.token) {
				const currentMap = buildCellMap(cells)
				const changed = diffCellMaps(previous.cells, currentMap)
				this._changeSnapshots.set(snapshotKey, { token: changeToken, cells: currentMap })
				cells = changed
			} else {
				this._changeSnapshots.set(snapshotKey, {
					token: changeToken,
					cells: buildCellMap(cells),
				})
			}
		}
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
			...(changeToken !== undefined ? { changeToken } : {}),
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
		const rows = this.requireSheet().cells.iterateRowsInRange(parsed)
		let next = rows.next()
		for (let row = parsed.start.row; row <= parsed.end.row; row++) {
			if (!next.done && next.value[0] === row) {
				yield next.value[1].map(([col, cell]) =>
					makeCompactCellInfo(
						row,
						col,
						cell,
						this.resolveFormula(row, col, cell),
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
		return this.requireSheet().cells.usedRange()
	}

	get state(): string {
		return this.requireSheet().state
	}

	get tabColor(): SheetTabColor | null {
		return this.requireSheet().tabColor
	}

	get sheetFormatPr(): SheetFormatPr | null {
		return this.requireSheet().sheetFormatPr
	}

	get frozenRows(): number {
		return this.requireSheet().frozenRows
	}

	get frozenCols(): number {
		return this.requireSheet().frozenCols
	}

	get merges(): readonly RangeRef[] {
		return this.requireSheet().merges
	}

	get autoFilter(): AutoFilter | null {
		return this.requireSheet().autoFilter
	}

	get protection(): SheetProtection | null {
		return this.requireSheet().protection
	}

	get conditionalFormats(): readonly SheetConditionalFormat[] {
		return this.requireSheet().conditionalFormats
	}

	get dataValidations(): readonly SheetDataValidation[] {
		return this.requireSheet().dataValidations
	}

	get imageRefs(): readonly SheetImageRef[] {
		return this.requireSheet().imageRefs
	}

	comments(): ReadonlyMap<string, SheetComment> {
		return this.requireSheet().comments
	}

	hyperlinks(): ReadonlyMap<string, SheetHyperlink> {
		return this.requireSheet().hyperlinks
	}

	comment(ref: string): SheetComment | undefined {
		return this.requireSheet().comments.get(ref)
	}

	hyperlink(ref: string): SheetHyperlink | undefined {
		return this.requireSheet().hyperlinks.get(ref)
	}

	private requireSheet(): Sheet {
		const sheet = this.resolveSheet()
		if (sheet) return sheet
		throw new AscendException(
			ascendError(
				'SHEET_NOT_FOUND',
				`Sheet "${this.sheetName}" is no longer available in the current workbook view.`,
				{
					refs: [this.sheetName],
				},
			),
		)
	}
}

function collectCellsCompact(
	sheet: Sheet,
	range: RangeRef,
	resolveFormula: (
		row: number,
		col: number,
		cell: NonNullable<ReturnType<Sheet['cells']['get']>>,
	) => string | null,
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
					resolveFormula(row, col, cell),
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
	formula: string | null,
	ref?: string,
): CompactCellInfo {
	return {
		...(ref ? { ref } : {}),
		value: cell.value,
		formula,
		formulaBinding: cell.formulaInfo ?? null,
		row,
		col,
	}
}

function toCellInfo(cell: CompactCellInfo, explicitRef?: string): CellInfo {
	const ref = explicitRef ?? cell.ref
	if (!ref) {
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', 'CellInfo conversion requires a reference'),
		)
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

function flattenCellValue(value: CellValue): FlatCellValue {
	switch (value.kind) {
		case 'number':
			return value.value
		case 'string':
			return value.value
		case 'boolean':
			return value.value
		case 'empty':
			return null
		case 'error':
			return value.value
		case 'date':
			return value.serial
		case 'richText':
			return value.runs.map((run: { text: string }) => run.text).join('')
		default:
			return null
	}
}

function buildCellMap(cells: readonly CompactCellInfo[]): Map<string, CompactCellInfo> {
	const map = new Map<string, CompactCellInfo>()
	for (const cell of cells) {
		map.set(`${cell.row},${cell.col}`, cell)
	}
	return map
}

function cellInfoEqual(a: CompactCellInfo, b: CompactCellInfo): boolean {
	if (a.formula !== b.formula) return false
	const av = a.value
	const bv = b.value
	if (av.kind !== bv.kind) return false
	if (av.kind === 'empty') return true
	if (av.kind === 'number' && bv.kind === 'number') return av.value === bv.value
	if (av.kind === 'string' && bv.kind === 'string') return av.value === bv.value
	if (av.kind === 'boolean' && bv.kind === 'boolean') return av.value === bv.value
	return JSON.stringify(av) === JSON.stringify(bv)
}

function diffCellMaps(
	previous: Map<string, CompactCellInfo>,
	current: Map<string, CompactCellInfo>,
): CompactCellInfo[] {
	const changed: CompactCellInfo[] = []
	for (const [key, cell] of current) {
		const prev = previous.get(key)
		if (!prev || !cellInfoEqual(prev, cell)) changed.push(cell)
	}
	return changed
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
