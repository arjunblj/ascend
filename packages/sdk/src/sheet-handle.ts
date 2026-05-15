import type {
	AutoFilter,
	CellFormulaBinding,
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
	StyleId,
} from '@ascend/core'
import { parseA1, parseRange, toA1, toRangeString } from '@ascend/core'
import { AscendException, ascendError, type CellValue, topLeftScalar } from '@ascend/schema'
import type { FormatDisplayOptions } from './format-helpers.ts'
import {
	type CellSelector,
	parseLocalCellSelector,
	parseLocalRangeSelector,
	type RangeSelector,
} from './ref-selectors.ts'
import type {
	AgentReadOptions,
	CellInfo,
	CommentSummary,
	CompactCellInfo,
	CompactRangeChangeInvalidation,
	CompactRangeChangeInvalidationReason,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	ConditionalFormatRuleSummary,
	DataValidationSummary,
	FlatCellValue,
	FormulaBindingSummary,
	FormulaCellEntry,
	HyperlinkSummary,
	MergeRangeSummary,
	PageMetadataSummary,
	RangeAggregateInfo,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	WorkbookReadSnapshotInfo,
} from './types.ts'

export interface SheetMetadataQueryOptions {
	readonly range?: RangeSelector
}

const COMPACT_CHANGE_SNAPSHOT_RETENTION = 8
const COMPACT_CELL_KEY_COLUMN_STRIDE = 0x4000

interface CompactChangeSnapshot {
	readonly token: string
	readonly cells: Map<CompactCellKey, CompactCellInfo>
}

interface CompactChangeSnapshotLedger {
	readonly snapshots: CompactChangeSnapshot[]
}

type CompactCellKey = number | string

export class SheetHandle {
	private readonly sheetName: string
	private readonly resolveSheet: () => Sheet | undefined
	private readonly resolveFormula: (
		row: number,
		col: number,
		cell: CellFormulaSource,
	) => string | null
	private readonly resolveDisplay: (
		row: number,
		col: number,
		cell: CellDisplaySource,
		options?: FormatDisplayOptions,
	) => string
	private readonly resolveSnapshot: () => WorkbookReadSnapshotInfo
	private _changeVersion = 0
	private readonly _changeSnapshots = new Map<string, CompactChangeSnapshotLedger>()

	constructor(
		sheetName: string,
		resolveSheet: () => Sheet | undefined,
		resolveFormula: (row: number, col: number, cell: CellFormulaSource) => string | null,
		resolveDisplay: (
			row: number,
			col: number,
			cell: CellDisplaySource,
			options?: FormatDisplayOptions,
		) => string,
		resolveSnapshot?: () => WorkbookReadSnapshotInfo,
	) {
		this.sheetName = sheetName
		this.resolveSheet = resolveSheet
		this.resolveFormula = resolveFormula
		this.resolveDisplay = resolveDisplay
		this.resolveSnapshot = resolveSnapshot ?? fallbackReadSnapshot
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

	cell(ref: CellSelector): CellInfo | undefined {
		const compact = this.cellCompact(ref)
		const refText = typeof ref === 'string' ? ref : parseLocalCellSelector(ref).ref
		return compact ? toCellInfo(compact, refText) : undefined
	}

	cellCompact(ref: CellSelector): CompactCellInfo | undefined {
		const { ref: refText, cell: parsed } = parseLocalCellSelector(ref)
		const cell = this.requireSheet().cells.get(parsed.row, parsed.col)
		if (!cell) return undefined
		return makeCompactCellInfo(
			parsed.row,
			parsed.col,
			cell.value,
			cell.formulaInfo,
			this.resolveFormula(parsed.row, parsed.col, cell),
			refText,
		)
	}

	formatCellForDisplay(ref: CellSelector, options?: FormatDisplayOptions): string | undefined {
		const { cell: parsed } = parseLocalCellSelector(ref)
		const cell = this.requireSheet().cells.get(parsed.row, parsed.col)
		if (!cell) return undefined
		return this.resolveDisplay(parsed.row, parsed.col, cell, options)
	}

	range(rangeRef: RangeSelector): RangeInfo {
		const compact = this.rangeCompact(rangeRef, { includeRefs: true })
		return {
			snapshot: compact.snapshot,
			ref: compact.ref,
			cells: compact.cells.map((cell) => toCellInfo(cell)),
			rowCount: compact.rowCount,
			colCount: compact.colCount,
		}
	}

	rangeCompact(
		rangeRef: RangeSelector,
		opts?: { includeRefs?: boolean; omitEmpty?: boolean; flatValues?: boolean },
	): CompactRangeInfo {
		const { ref: parsed } = parseLocalRangeSelector(rangeRef)
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
			snapshot: this.resolveSnapshot(),
			ref: parsed,
			cells,
			rowCount: parsed.end.row - parsed.start.row + 1,
			colCount: parsed.end.col - parsed.start.col + 1,
		}
	}

	readWindow(rangeRef: string, opts?: { rowOffset?: number; rowLimit?: number }): RangeWindowInfo {
		const compact = this.readWindowCompact(rangeRef, { ...opts, includeRefs: true })
		return {
			snapshot: compact.snapshot,
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
		let cells =
			opts?.includeRefs === false && opts.flatValues
				? collectFlatCellsCompact(sheet, windowRef, this.resolveFormula, opts)
				: collectCellsCompact(sheet, windowRef, this.resolveFormula, opts)
		if (opts?.omitEmpty && !(opts.includeRefs === false && opts.flatValues)) {
			cells = cells.filter((c) => c.value.kind !== 'empty')
		}
		if (opts?.flatValues && !(opts.includeRefs === false && opts.flatValues)) {
			cells = cells.map((c) => ({
				...c,
				value: flattenCellValue(c.value) as unknown as CellValue,
			}))
		}
		const consumedRows = Math.max(0, endRow - requestedRef.start.row + 1)
		const hasMore = requestedRef.start.row + rowOffset + rowLimit - 1 < requestedRef.end.row
		const snapshotKey =
			opts?.changedSince !== undefined
				? compactChangeSnapshotKey(this.sheetName, windowRef, opts)
				: null
		let changeToken: string | undefined
		let changeInvalidation: CompactRangeChangeInvalidation | undefined
		if (snapshotKey) {
			const version = this._changeVersion++
			changeToken = `${version}`
			const baseToken = opts?.changedSince
			const ledger = this._changeSnapshots.get(snapshotKey)
			const currentMap = buildCellMap(cells)
			if (baseToken && !isCompactChangeToken(baseToken)) {
				changeInvalidation = {
					baseToken,
					changeToken,
					reason: 'base-token-invalid',
					requiredAction: 'use-returned-window',
				}
				this.rememberCompactChangeSnapshot(snapshotKey, changeToken, currentMap)
			} else if (baseToken && ledger) {
				const previous = findCompactChangeSnapshot(ledger, baseToken)
				if (previous) {
					cells = diffCellMaps(previous.cells, currentMap)
				} else {
					changeInvalidation = {
						baseToken,
						changeToken,
						reason: compactChangeInvalidationReason(baseToken, ledger),
						requiredAction: 'use-returned-window',
					}
				}
				this.rememberCompactChangeSnapshot(snapshotKey, changeToken, currentMap)
			} else {
				if (baseToken) {
					changeInvalidation = {
						baseToken,
						changeToken,
						reason: 'base-snapshot-missing',
						requiredAction: 'use-returned-window',
					}
				}
				this.rememberCompactChangeSnapshot(snapshotKey, changeToken, currentMap)
			}
		}
		return {
			snapshot: this.resolveSnapshot(),
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
			...(changeInvalidation !== undefined ? { changeInvalidation } : {}),
		}
	}

	private rememberCompactChangeSnapshot(
		snapshotKey: string,
		token: string,
		cells: Map<CompactCellKey, CompactCellInfo>,
	): void {
		const ledger = this._changeSnapshots.get(snapshotKey) ?? { snapshots: [] }
		ledger.snapshots.push({ token, cells })
		while (ledger.snapshots.length > COMPACT_CHANGE_SNAPSHOT_RETENTION) {
			ledger.snapshots.shift()
		}
		this._changeSnapshots.set(snapshotKey, ledger)
	}

	readRows(rangeRef: string, opts?: { rowOffset?: number; rowLimit?: number }): RangeRowsInfo {
		const window = this.readWindowCompact(rangeRef, {
			...(opts?.rowOffset !== undefined ? { rowOffset: opts.rowOffset } : {}),
			...(opts?.rowLimit !== undefined ? { rowLimit: opts.rowLimit } : {}),
			includeRefs: false,
		})
		return {
			snapshot: window.snapshot,
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
			snapshot: rowsInfo.snapshot,
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

	aggregateRange(rangeRef: string): RangeAggregateInfo {
		const ref = parseRange(rangeRef)
		const sheet = this.requireSheet()
		const rowCount = ref.end.row - ref.start.row + 1
		const colCount = ref.end.col - ref.start.col + 1
		const totalCells = rowCount * colCount
		let populatedCells = 0
		let numericCount = 0
		let textCount = 0
		let booleanCount = 0
		let errorCount = 0
		let sum = 0
		let min = Number.POSITIVE_INFINITY
		let max = Number.NEGATIVE_INFINITY
		let firstError: string | undefined

		sheet.cells.forEachCellInRange(ref, (_row, _col, cell) => {
			const value = topLeftScalar(cell.value)
			if (value.kind === 'empty') return
			populatedCells += 1
			switch (value.kind) {
				case 'number':
				case 'date':
					{
						const numericValue = value.kind === 'number' ? value.value : value.serial
						numericCount += 1
						sum += numericValue
						min = Math.min(min, numericValue)
						max = Math.max(max, numericValue)
					}
					break
				case 'string':
				case 'richText':
					textCount += 1
					break
				case 'boolean':
					booleanCount += 1
					break
				case 'error':
					errorCount += 1
					firstError ??= value.value
					break
			}
		})

		return {
			ref,
			totalCells,
			populatedCells,
			blankCount: totalCells - populatedCells,
			numericCount,
			textCount,
			booleanCount,
			errorCount,
			sum,
			average: numericCount > 0 ? sum / numericCount : null,
			min: numericCount > 0 ? min : null,
			max: numericCount > 0 ? max : null,
			...(firstError ? { firstError } : {}),
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
						cell.value,
						cell.formulaInfo,
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

	/**
	 * Return an array of conditional format rule summaries (type, priority, range).
	 */
	getConditionalFormats(): readonly ConditionalFormatRuleSummary[] {
		const sheet = this.requireSheet()
		const summaries: ConditionalFormatRuleSummary[] = []
		for (const cf of sheet.conditionalFormats) {
			const range = cf.sqref
			for (const rule of cf.rules) {
				summaries.push({
					type: rule.type,
					...(rule.priority !== undefined ? { priority: rule.priority } : {}),
					range,
				})
			}
		}
		return summaries
	}

	/**
	 * Return an array of data validation summaries (type, formula, range).
	 */
	getDataValidations(): readonly DataValidationSummary[] {
		const sheet = this.requireSheet()
		return sheet.dataValidations.map((dv) => ({
			...(dv.type ? { type: dv.type } : {}),
			...(dv.formula1 ? { formula: dv.formula1 } : {}),
			range: dv.sqref,
		}))
	}

	/**
	 * Return an array of comment summaries (ref, author, text).
	 */
	getComments(options: SheetMetadataQueryOptions = {}): readonly CommentSummary[] {
		const sheet = this.requireSheet()
		const viewport = options.range ? parseLocalRangeSelector(options.range).ref : null
		return [...sheet.comments.entries()]
			.filter(([ref]) => !viewport || refInRange(ref, viewport))
			.map(([ref, c]) => ({
				ref,
				...(c.author !== undefined ? { author: c.author } : {}),
				text: c.text,
			}))
	}

	/**
	 * Return an array of hyperlink summaries.
	 */
	getHyperlinks(options: SheetMetadataQueryOptions = {}): readonly HyperlinkSummary[] {
		const sheet = this.requireSheet()
		const viewport = options.range ? parseLocalRangeSelector(options.range).ref : null
		return [...sheet.hyperlinks.entries()]
			.filter(([ref]) => !viewport || refInRange(ref, viewport))
			.map(([ref, h]) => ({
				ref,
				...(h.target !== undefined ? { target: h.target } : {}),
				...(h.location !== undefined ? { location: h.location } : {}),
				...(h.display !== undefined ? { display: h.display } : {}),
				...(h.tooltip !== undefined ? { tooltip: h.tooltip } : {}),
			}))
	}

	/**
	 * Return an array of merge range summaries.
	 */
	getMerges(): readonly MergeRangeSummary[] {
		const sheet = this.requireSheet()
		return sheet.merges.map((m) => ({ range: toRangeString(m) }))
	}

	/**
	 * Return page setup info if present (margins, orientation, etc.).
	 */
	getPageMetadata(): PageMetadataSummary | null {
		const sheet = this.requireSheet()
		if (!sheet.pageMargins && !sheet.pageSetup) return null
		return {
			...(sheet.pageMargins ? { margins: sheet.pageMargins } : {}),
			...(sheet.pageSetup ? { setup: sheet.pageSetup } : {}),
		}
	}

	/**
	 * Return true if the sheet has an auto-filter applied.
	 */
	hasAutoFilter(): boolean {
		return this.requireSheet().autoFilter !== null
	}

	/**
	 * Return true if the sheet has drawings (charts, images, shapes).
	 */
	hasDrawings(): boolean {
		const refs = this.requireSheet().drawingRefs
		return refs.hasDrawing || refs.hasLegacyDrawing
	}

	/**
	 * Return a summary of the formula binding at the given cell.
	 * Returns null for non-formula cells.
	 */
	getFormulaBinding(ref: string): FormulaBindingSummary | null {
		const parsed = parseA1(ref)
		const cell = this.requireSheet().cells.get(parsed.row, parsed.col)
		if (!cell) return null
		const formula = this.resolveFormula(parsed.row, parsed.col, cell)
		const binding = toFormulaBindingSummary(formula, cell.formulaInfo)
		if (!binding) return null
		return binding
	}

	/**
	 * Return all formula cells in the sheet with their binding info.
	 * Optionally filter by kind (e.g. 'normal', 'shared-anchor', 'shared-member', 'array', 'dynamic-array', 'spill').
	 */
	getFormulaCells(options?: { kind?: string }): readonly FormulaCellEntry[] {
		const sheet = this.requireSheet()
		const entries: FormulaCellEntry[] = []
		for (const [row, rowCells] of sheet.cells.iterateRows()) {
			for (const [col, cell] of rowCells) {
				const formula = this.resolveFormula(row, col, cell)
				const binding = toFormulaBindingSummary(formula, cell.formulaInfo)
				if (!binding) continue
				if (options?.kind && binding.kind !== options.kind) continue
				entries.push({
					ref: toA1({ row, col }),
					binding,
				})
			}
		}
		return entries
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

function toFormulaBindingSummary(
	formula: string | null,
	info: CellFormulaBinding | undefined,
): FormulaBindingSummary | null {
	if (info?.kind === 'spill') {
		return { kind: 'spill', anchorRef: info.anchorRef }
	}
	if (formula === null) return null
	if (!info) return { kind: 'normal', formula }
	switch (info.kind) {
		case 'shared':
			if (info.isMaster) {
				return {
					kind: 'shared-anchor',
					formula,
					sharedIndex: info.sharedIndex,
					...(info.ref ? { range: info.ref } : {}),
				}
			}
			return {
				kind: 'shared-member',
				sharedIndex: info.sharedIndex,
				...(info.masterRef ? { masterRef: info.masterRef } : {}),
			}
		case 'array':
			return {
				kind: 'array',
				formula,
				...(info.ref ? { range: info.ref } : {}),
			}
		case 'dynamicArray':
			return { kind: 'dynamic-array', formula }
		case 'blockedSpill':
			return {
				kind: 'blocked-spill',
				formula,
				range: info.ref,
				...(info.reason ? { cause: info.reason } : {}),
				blockingRefs: info.blockingRefs,
			}
		default:
			return { kind: 'normal', formula }
	}
}

function collectCellsCompact(
	sheet: Sheet,
	range: RangeRef,
	resolveFormula: (row: number, col: number, cell: CellFormulaSource) => string | null,
	opts?: { includeRefs?: boolean },
): CompactCellInfo[] {
	const cells: CompactCellInfo[] = []
	const includeRefs = opts?.includeRefs !== false
	sheet.cells.forEachCellContentInRange(range, (row, col, value, formula, formulaInfo) => {
		const formulaSource = { formula, formulaInfo }
		const resolvedFormula = resolveFormula(row, col, formulaSource)
		if (includeRefs) {
			cells.push(
				makeCompactCellInfo(row, col, value, formulaInfo, resolvedFormula, toA1({ row, col })),
			)
			return
		}
		cells.push({
			value,
			formula: resolvedFormula,
			formulaBinding: formulaInfo ?? null,
			row,
			col,
		})
	})
	return cells
}

function collectFlatCellsCompact(
	sheet: Sheet,
	range: RangeRef,
	resolveFormula: (row: number, col: number, cell: CellFormulaSource) => string | null,
	opts?: { omitEmpty?: boolean },
): CompactCellInfo[] {
	const cells: CompactCellInfo[] = []
	const omitEmpty = opts?.omitEmpty === true
	sheet.cells.forEachCellContentInRange(range, (row, col, value, formula, formulaInfo) => {
		if (omitEmpty && value.kind === 'empty') return
		const formulaSource = { formula, formulaInfo }
		cells.push({
			value: flattenCellValue(value) as unknown as CellValue,
			formula: resolveFormula(row, col, formulaSource),
			formulaBinding: formulaInfo ?? null,
			row,
			col,
		})
	})
	return cells
}

function makeCompactCellInfo(
	row: number,
	col: number,
	value: CellValue,
	formulaInfo: CellFormulaBinding | undefined,
	formula: string | null,
	ref?: string,
): CompactCellInfo {
	return {
		...(ref ? { ref } : {}),
		value,
		formula,
		formulaBinding: formulaInfo ?? null,
		row,
		col,
	}
}

function fallbackReadSnapshot(): WorkbookReadSnapshotInfo {
	return {
		token: 'w0:m0:f0:s0:full|cells|rich|all-sheets|rows:all|loaded:',
		generations: { workbook: 0, sheetMetadata: 0, formulas: 0, styles: 0 },
		load: {
			mode: 'full',
			isPartial: false,
			cellsHydrated: true,
			richSheetMetadataHydrated: true,
			hasAllSheets: true,
			partialReasons: [],
			sourceSheets: [],
			loadedSheets: [],
		},
	}
}

interface CellFormulaSource {
	readonly formula: string | null
	readonly formulaInfo?: CellFormulaBinding | undefined
}

interface CellDisplaySource {
	readonly value: CellValue
	readonly formula: string | null
	readonly styleId: StyleId
	readonly formulaInfo?: CellFormulaBinding | undefined
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

function buildCellMap(cells: readonly CompactCellInfo[]): Map<CompactCellKey, CompactCellInfo> {
	const map = new Map<CompactCellKey, CompactCellInfo>()
	for (const cell of cells) {
		map.set(compactCellKey(cell.row, cell.col), cell)
	}
	return map
}

function compactCellKey(row: number, col: number): CompactCellKey {
	if (row >= 0 && col >= 0 && col < COMPACT_CELL_KEY_COLUMN_STRIDE) {
		const key = row * COMPACT_CELL_KEY_COLUMN_STRIDE + col
		if (Number.isSafeInteger(key)) return key
	}
	return `${row},${col}`
}

function compactChangeSnapshotKey(
	sheetName: string,
	windowRef: RangeRef,
	opts: AgentReadOptions,
): string {
	return JSON.stringify({
		sheetName,
		ref: toRangeString(windowRef),
		includeRefs: opts.includeRefs !== false,
		omitEmpty: opts.omitEmpty === true,
		flatValues: opts.flatValues === true,
		projection: opts.changeProjectionKey ?? '',
	})
}

function isCompactChangeToken(token: string): boolean {
	return /^\d+$/.test(token)
}

function findCompactChangeSnapshot(
	ledger: CompactChangeSnapshotLedger,
	token: string,
): CompactChangeSnapshot | undefined {
	return ledger.snapshots.find((snapshot) => snapshot.token === token)
}

function compactChangeInvalidationReason(
	baseToken: string,
	ledger: CompactChangeSnapshotLedger,
): CompactRangeChangeInvalidationReason {
	const baseVersion = Number(baseToken)
	const oldest = ledger.snapshots[0]
	const oldestVersion = oldest ? Number(oldest.token) : Number.NaN
	if (
		Number.isFinite(baseVersion) &&
		Number.isFinite(oldestVersion) &&
		baseVersion < oldestVersion
	) {
		return 'base-token-expired'
	}
	return 'base-token-stale'
}

function refInRange(ref: string, range: RangeRef): boolean {
	const cell = parseA1(ref)
	return (
		cell.row >= range.start.row &&
		cell.row <= range.end.row &&
		cell.col >= range.start.col &&
		cell.col <= range.end.col
	)
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
	previous: Map<CompactCellKey, CompactCellInfo>,
	current: Map<CompactCellKey, CompactCellInfo>,
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
