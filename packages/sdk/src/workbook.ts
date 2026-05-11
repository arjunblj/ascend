import { writeFile } from 'node:fs/promises'
import {
	createWorkbook,
	indexToColumn,
	parseA1,
	parseRange,
	type RangeRef,
	type Workbook,
} from '@ascend/core'
import {
	applyOperations,
	applyWithTransaction,
	type CalcContext,
	defaultCalcContext,
	cellValuesEqual as diffCellValuesEqual,
	diffWorkbooks,
	type EvalContext,
	evaluate,
	type PatchResult,
	recalculate,
} from '@ascend/engine'
import { cachedParseFormula, normalizeFormulaInput } from '@ascend/formulas'
import { readCsv, writeCsv } from '@ascend/io-csv'
import {
	extractZip,
	type PreservationCapsule,
	summarizePlannedWrite,
	writeXlsx,
	type ZipArchive,
} from '@ascend/io-xlsx'
import {
	AscendException,
	ascendError,
	type CellValue,
	type CompatibilityReport,
	type CsvDialect,
	coerceCellValueToString,
	EMPTY,
	emptyReport,
	type InputValue,
	type Operation,
} from '@ascend/schema'
import { check as verifyCheck, lint as verifyLint } from '@ascend/verify'
import { getCapability, listCapabilities, summarizeCapabilities } from './capabilities.ts'
import { partialDependencyCheckIssue, sdkCheckIssueFromVerify } from './check-issues.ts'
import { buildWorkbookLoadInfo, openWorkbookSource } from './load.ts'
import { getOperationsSchema, listOperations, parseOperations } from './ops.ts'
import { WorkbookReadView } from './read-view.ts'
import {
	type CellSelector,
	normalizeCellSelector,
	normalizeRangeSelector,
	type RangeSelector,
} from './ref-selectors.ts'
import type {
	ApplyAndRecalcResult,
	ApplyResult,
	BatchResult,
	CheckIssue,
	CheckResult,
	EvalOptions,
	LintResult,
	LintWarning,
	RecalcOptions,
	RecalcResult,
	WritePlanInfo,
} from './types.ts'

function cloneWorkbook(source: Workbook): Workbook {
	return source.clone()
}

function stringMatches(
	haystack: string,
	needle: string,
	mode: 'exact' | 'contains' | 'startsWith' | 'endsWith',
): boolean {
	switch (mode) {
		case 'exact':
			return haystack === needle
		case 'contains':
			return haystack.includes(needle)
		case 'startsWith':
			return haystack.startsWith(needle)
		case 'endsWith':
			return haystack.endsWith(needle)
	}
}

/**
 * Full mutable workbook. Use for apply, recalc, save, export, and any editing.
 * Opens the full workbook into memory. Use `WorkbookDocument` for read-only
 * operations (inspect, read, check, lint, trace) when you don't need to modify.
 */
export class AscendWorkbook extends WorkbookReadView {
	private readonly caps: PreservationCapsule[]
	private originalBytes: Uint8Array | null
	private sourceArchive: ZipArchive | undefined
	private dirty: boolean
	private readonly dirtySheets = new Set<string>()
	private pendingDirtyRefs: string[] = []
	private pendingFullRecalc = false
	private _batchMode = false
	private workbookMetaDirty = false
	private documentPropertiesDirty = false
	private calcStateDirty = false
	private calcChainDirty = false
	private sharedStringsDirty = false
	private stylesDirty = false

	private constructor(
		workbook: Workbook,
		capsules: PreservationCapsule[],
		report: CompatibilityReport,
		loadInfo: import('./types.ts').WorkbookLoadInfo,
		originalBytes: Uint8Array | null,
	) {
		super(workbook, report, loadInfo)
		this.caps = capsules
		this.originalBytes = originalBytes
		this.dirty = false
	}

	/**
	 * Open a workbook from a file path or bytes.
	 * @example
	 * const wb = await AscendWorkbook.open('./data.xlsx')
	 * const wb2 = await AscendWorkbook.open(bytes, { mode: 'values', sheets: ['Sheet1'] })
	 * const wb3 = await AscendWorkbook.open(bytes, { mode: 'values', richMetadata: true })
	 */
	static async open(
		pathOrBytes: string | Uint8Array,
		options?: {
			mode?: 'full' | 'metadata-only' | 'values' | 'formula'
			sheets?: readonly string[]
			richMetadata?: boolean
			password?: string
			pivotCacheRecordMaterializeLimit?: number | 'all'
		},
	): Promise<AscendWorkbook> {
		const loaded = await openWorkbookSource(pathOrBytes, options)
		return new AscendWorkbook(
			loaded.workbook,
			[...loaded.capsules],
			loaded.report,
			loaded.loadInfo,
			loaded.originalBytes,
		)
	}

	/**
	 * Create a new empty workbook with a default Sheet1.
	 * @example
	 * const wb = AscendWorkbook.create()
	 */
	static create(): AscendWorkbook {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		return new AscendWorkbook(
			wb,
			[],
			emptyReport('ascend'),
			buildWorkbookLoadInfo({
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				richSheetMetadataHydrated: true,
				hasAllSheets: true,
				sourceSheetNames: ['Sheet1'],
				loadedSheetNames: ['Sheet1'],
			}),
			null,
		)
	}

	/**
	 * Create a workbook from CSV content.
	 * @example
	 * const wb = AscendWorkbook.fromCsv('a,b,c\n1,2,3')
	 */
	static fromCsv(content: string, dialect?: Partial<CsvDialect>): AscendWorkbook {
		const result = readCsv(content, dialect)
		if (!result.ok) throw new AscendException(result.error)
		return new AscendWorkbook(
			result.value,
			[],
			emptyReport('csv'),
			buildWorkbookLoadInfo({
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				richSheetMetadataHydrated: true,
				hasAllSheets: true,
				sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
				loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
			}),
			null,
		)
	}

	// --- Mutation ---

	preview(ops: readonly Operation[]): import('./types.ts').PreviewResult {
		if (this.loadInfo.isPartial) {
			return {
				diff: {
					sheets: [],
					namesAdded: [],
					namesRemoved: [],
					namesChanged: [],
					workbookProtectionChanged: false,
					sheetFeatures: [],
				},
				sheetDiffs: [],
				cellChanges: [],
				changedCells: [],
				recalcScope: 0,
				warnings: [],
				wouldSucceed: false,
				errors: [partialWorkbookEditError()],
			}
		}
		const clone = cloneWorkbook(this.wb)
		const errors: import('@ascend/schema').AscendError[] = []

		const result = applyOperations(clone, ops)
		if (!result.ok) {
			errors.push(...('errors' in result.error ? result.error.errors : [result.error]))
			return {
				diff: {
					sheets: [],
					namesAdded: [],
					namesRemoved: [],
					namesChanged: [],
					workbookProtectionChanged: false,
					sheetFeatures: [],
				},
				sheetDiffs: [],
				cellChanges: [],
				changedCells: [],
				recalcScope: 0,
				warnings: [],
				wouldSucceed: false,
				errors,
			}
		}

		const warnings = [...(result.value.warnings ?? [])]

		let recalcResult: import('./types.ts').RecalcResult | undefined
		if (result.value.recalcRequired) {
			const recalcTargets = deriveRecalcTargets(ops)
			const calcResult = recalculate(
				clone,
				defaultCalcContext({
					dateSystem: clone.calcSettings.dateSystem,
					iterativeCalc: clone.calcSettings.iterativeCalc,
				}),
				resolveRecalcOptions(recalcTargets.fullRecalc, recalcTargets.dirtyRefs),
			)
			recalcResult = {
				changed: calcResult.changed,
				errors: calcResult.errors,
				duration: calcResult.duration,
			}
			for (const issue of calcResult.errors) {
				errors.push({
					...issue.error,
					...(issue.error.refs ? {} : { refs: [issue.ref] }),
				})
			}
		}

		const fastDiff = buildFastPreviewDiff(this.wb, clone, ops, result.value, recalcResult)
		const diff = fastDiff ?? diffWorkbooks(this.wb, clone)
		const cellChanges = diff.sheets.flatMap((s) => s.cellsChanged)
		const changedCells = diff.sheets.flatMap((s) =>
			s.cellsChanged.map((c) => ({
				ref: `${s.name}!${c.ref}`,
				oldValue: c.before,
				newValue: c.after,
			})),
		)
		const recalcScope = recalcResult?.changed.length ?? 0
		let cachedWritePlan: WritePlanInfo | undefined
		let writePlanComputed = false
		let cachedDirtyFlags:
			| {
					workbookMetaDirty: boolean
					documentPropertiesDirty: boolean
					calcChainDirty: boolean
					sharedStringsDirty: boolean
					stylesDirty: boolean
			  }
			| undefined
		const previewResult: import('./types.ts').PreviewResult = {
			diff,
			sheetDiffs: diff.sheets,
			cellChanges,
			changedCells,
			recalcScope,
			warnings,
			wouldSucceed: errors.length === 0,
			errors,
		}
		Object.defineProperty(previewResult, 'writePlan', {
			configurable: true,
			enumerable: true,
			get: () => {
				if (!writePlanComputed) {
					writePlanComputed = true
					cachedDirtyFlags ??= this.deriveDirtyFlags(ops)
					const sourceArchive = this.getSourceArchive(false)
					const plan = summarizePlannedWrite(clone, this.caps.length > 0 ? this.caps : undefined, {
						dirtySheetNames: result.value.sheetsModified,
						workbookMetaDirty: cachedDirtyFlags.workbookMetaDirty,
						documentPropertiesDirty: cachedDirtyFlags.documentPropertiesDirty,
						calcStateDirty: cachedDirtyFlags.workbookMetaDirty || result.value.recalcRequired,
						calcChainDirty: cachedDirtyFlags.calcChainDirty,
						sharedStringsDirty: cachedDirtyFlags.sharedStringsDirty,
						stylesDirty: cachedDirtyFlags.stylesDirty,
						...(sourceArchive ? { sourceArchive } : {}),
					})
					cachedWritePlan = plan.ok ? plan.value : undefined
				}
				return cachedWritePlan
			},
		})
		return previewResult
	}

	/**
	 * Apply operations to the workbook. Does not recalculate formulas.
	 * When transaction is true, the batch is atomic: all operations succeed or none apply.
	 * @example
	 * wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] }])
	 * wb.apply(ops, { transaction: true })
	 */
	apply(
		ops: readonly Operation[],
		options?: { collectAllErrors?: boolean; transaction?: boolean },
	): ApplyResult {
		if (this.loadInfo.isPartial) {
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				errors: [partialWorkbookEditError()],
			}
		}
		const dirtyFlags = this.deriveDirtyFlags(ops)
		const nextWorkbook = cloneWorkbook(this.wb)
		const applyFn = options?.transaction ? applyWithTransaction : applyOperations
		const result = applyFn(nextWorkbook, ops, options)
		if (!result.ok) {
			const errors = 'errors' in result.error ? result.error.errors : [result.error]
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				errors,
			}
		}

		this.wb = nextWorkbook
		this.markDirty()
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.documentPropertiesDirty ||= dirtyFlags.documentPropertiesDirty
		this.calcStateDirty ||= dirtyFlags.workbookMetaDirty || result.value.recalcRequired
		this.calcChainDirty ||= dirtyFlags.calcChainDirty
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		this.stylesDirty ||= dirtyFlags.stylesDirty
		if (result.value.recalcRequired) this.mergePendingRecalcTargets(ops)
		if (!this._batchMode) this.clearReadCaches()
		return {
			affectedCells: result.value.affectedCells,
			sheetsModified: result.value.sheetsModified,
			recalcRequired: result.value.recalcRequired,
			errors: [],
			...(result.value.warnings && result.value.warnings.length > 0
				? { warnings: result.value.warnings }
				: {}),
		}
	}

	/**
	 * Apply operations and recalculate affected formulas.
	 * @example
	 * wb.applyAndRecalc([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 10 }] }])
	 */
	applyAndRecalc(ops: readonly Operation[], opts?: { range?: string }): ApplyAndRecalcResult {
		const apply = this.apply(ops)
		if (apply.errors.length > 0 || !apply.recalcRequired) {
			return { apply, recalc: null }
		}
		return {
			apply,
			recalc: this.recalc(opts),
		}
	}

	/**
	 * Apply operations atomically and recalculate once, or run a function with deferred recalc.
	 * @example
	 * wb.batch([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }, { ref: 'B1', value: 2 }] }])
	 * wb.batch(() => { wb.set('A1', 1); wb.set('B1', 2) })
	 */
	batch(ops: readonly Operation[]): BatchResult
	batch(fn: () => void): void
	batch(opsOrFn: readonly Operation[] | (() => void)): BatchResult | undefined {
		if (typeof opsOrFn === 'function') {
			if (this.loadInfo.isPartial) {
				throw new AscendException(
					ascendError(
						'VALIDATION_ERROR',
						'Cannot run batch in a partial workbook view. Reopen with a full load before editing.',
					),
				)
			}
			this._batchMode = true
			let completed = false
			try {
				opsOrFn()
				completed = true
			} finally {
				this._batchMode = false
				if (completed) this.recalc()
			}
			return
		}
		if (this.loadInfo.isPartial) {
			return { errors: [partialWorkbookEditError()] }
		}
		const dirtyFlags = this.deriveDirtyFlags(opsOrFn)
		const nextWorkbook = cloneWorkbook(this.wb)
		const result = applyOperations(nextWorkbook, opsOrFn)
		if (!result.ok) {
			const errs = 'errors' in result.error ? result.error.errors : [result.error]
			return { errors: errs }
		}

		this.wb = nextWorkbook
		this.markDirty()
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.documentPropertiesDirty ||= dirtyFlags.documentPropertiesDirty
		this.calcStateDirty ||= dirtyFlags.workbookMetaDirty || result.value.recalcRequired
		this.calcChainDirty ||= dirtyFlags.calcChainDirty
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		this.stylesDirty ||= dirtyFlags.stylesDirty
		if (result.value.recalcRequired) this.mergePendingRecalcTargets(opsOrFn)
		this.clearReadCaches()
		return { errors: [] }
	}

	/**
	 * Recalculate formulas. Optionally limit to a range.
	 * @example
	 * wb.recalc()
	 * wb.recalc({ range: 'Sheet1!A1:C10' })
	 */
	recalc(opts?: RecalcOptions): RecalcResult {
		if (this.loadInfo.isPartial) {
			return {
				changed: [],
				errors: [{ ref: '', error: partialWorkbookEditError() }],
				duration: 0,
			}
		}
		const ctx: CalcContext = defaultCalcContext({
			dateSystem: this.wb.calcSettings.dateSystem,
			...(opts?.externalReferences ? { externalReferences: opts.externalReferences } : {}),
			iterativeCalc: this.wb.calcSettings.iterativeCalc,
		})
		let rangeRef: RangeRef | undefined
		if (opts?.range) {
			rangeRef = parseRange(opts.range)
		}
		const result = recalculate(
			this.wb,
			ctx,
			rangeRef
				? { range: rangeRef }
				: resolveRecalcOptions(this.pendingFullRecalc, this.pendingDirtyRefs),
		)
		if (result.changed.length > 0 || result.errors.length > 0) {
			this.markDirty()
			this.calcStateDirty = result.errors.length > 0
			this.sharedStringsDirty = true
			for (const ref of result.changed) {
				const bang = ref.indexOf('!')
				if (bang !== -1) this.dirtySheets.add(ref.slice(0, bang))
			}
		}
		if (result.errors.length === 0) {
			this.calcStateDirty = false
			this.wb.calcSettings = {
				...this.wb.calcSettings,
				fullCalcOnLoad: false,
				calcCompleted: true,
				calcOnSave: true,
				forceFullCalc: false,
			}
		}
		this.clearReadCaches()
		this.pendingDirtyRefs = []
		this.pendingFullRecalc = false
		return {
			changed: result.changed,
			errors: result.errors,
			duration: result.duration,
		}
	}

	/**
	 * Evaluate a formula against the current workbook state without writing to a cell.
	 * Throws on parse errors. Returns the computed CellValue (which may be an error
	 * value like #REF! if the formula references invalid cells).
	 * @example
	 * const result = wb.eval('SUM(A1:A10)')
	 */
	eval(formula: string, opts?: EvalOptions): CellValue {
		const normalized = normalizeFormulaInput(formula)
		const parsed = cachedParseFormula(normalized)
		if (!parsed.ok) {
			throw new AscendException(
				ascendError('INVALID_ARGUMENT', `Failed to parse formula: ${formula}`),
			)
		}
		const ctx: EvalContext = {
			workbook: this.wb,
			calcContext: defaultCalcContext({
				dateSystem: this.wb.calcSettings.dateSystem,
				...(opts?.externalReferences ? { externalReferences: opts.externalReferences } : {}),
				iterativeCalc: this.wb.calcSettings.iterativeCalc,
			}),
			sheetIndex: 0,
			row: 0,
			col: 0,
		}
		return evaluate(parsed.value, ctx)
	}

	// --- Verification ---

	check(): CheckResult {
		const issue = this.dependencyVerificationIssue()
		if (issue) {
			return {
				valid: false,
				issues: [partialDependencyCheckIssue(issue)],
			}
		}
		const result = verifyCheck(this.wb, {
			formulas: this.formulaAnalysis(),
			dependencies: this.dependencyAnalysis(),
		})
		const issues: CheckIssue[] = result.issues.map(sdkCheckIssueFromVerify)
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const result = verifyLint(this.wb, this.formulaAnalysis())
		const warnings: LintWarning[] = result.violations.map((violation) => ({
			rule: violation.rule,
			message: violation.message,
			ref: violation.ref,
		}))
		return { clean: warnings.length === 0, warnings }
	}

	/**
	 * Find cells matching a value or formula in a sheet.
	 * @example
	 * const matches = wb.find('Sheet1', { value: 'hello', match: 'contains' })
	 * const formulas = wb.find('Sheet1', { value: 'SUM', match: 'contains', in: 'formula' })
	 */
	find(
		sheet: string,
		options: {
			value: string | number | boolean
			match?: 'exact' | 'contains' | 'startsWith' | 'endsWith'
			in?: 'value' | 'formula' | 'both'
		},
	): Array<{ ref: string; value: CellValue; formula?: string }> {
		const s = this.wb.getSheet(sheet)
		if (!s) return []
		const results: Array<{ ref: string; value: CellValue; formula?: string }> = []
		const { value: searchValue, match = 'exact', in: searchIn = 'value' } = options

		for (const [row, col, cell] of s.cells.iterate()) {
			const cellValue = cell.value ?? EMPTY
			const ref = indexToColumn(col) + (row + 1)
			const formula = s.cells.readFormula(row, col)

			if (typeof searchValue === 'number' || typeof searchValue === 'boolean') {
				if (searchIn === 'value' || searchIn === 'both') {
					const scalar = cellValue.kind === 'array' ? (cellValue.rows[0]?.[0] ?? EMPTY) : cellValue
					if (scalar.kind === 'number' && searchValue === scalar.value) {
						results.push({ ref, value: cellValue, ...(formula ? { formula } : {}) })
						continue
					}
					if (scalar.kind === 'boolean' && searchValue === scalar.value) {
						results.push({ ref, value: cellValue, ...(formula ? { formula } : {}) })
					}
				}
			} else {
				const searchLower = searchValue.toLowerCase()
				let matched = false

				if (searchIn === 'value' || searchIn === 'both') {
					const cellStr = coerceCellValueToString(cellValue)
					matched = stringMatches(cellStr.toLowerCase(), searchLower, match)
				}

				if (!matched && (searchIn === 'formula' || searchIn === 'both') && formula) {
					matched = stringMatches(formula.toLowerCase(), searchLower, match)
				}

				if (matched) {
					results.push({ ref, value: cellValue, ...(formula ? { formula } : {}) })
				}
			}
		}
		return results
	}

	/**
	 * Set a cell formula.
	 * @example
	 * wb.setFormula('Sheet1!B1', '=A1*2')
	 */
	setFormula(cellRef: CellSelector, formula: string): ApplyResult {
		const { sheetName, ref } = normalizeCellSelector(cellRef, this.wb)
		return this.apply([
			{ op: 'setFormula', sheet: sheetName, ref, formula: normalizeFormulaInput(formula) },
		])
	}

	/**
	 * Set a cell value.
	 * @example
	 * wb.set('Sheet1!A1', 42)
	 */
	set(cellRef: CellSelector, value: InputValue): ApplyResult {
		const { sheetName, ref } = normalizeCellSelector(cellRef, this.wb)
		return this.apply([{ op: 'setCells', sheet: sheetName, updates: [{ ref, value }] }])
	}

	/**
	 * Get a cell's value.
	 * @example
	 * const value = wb.get('Sheet1!A1')
	 */
	get(cellRef: CellSelector): CellValue {
		const { sheetName, ref } = normalizeCellSelector(cellRef, this.wb)
		const handle = this.sheet(sheetName)
		if (!handle) return EMPTY
		const cell = handle.cell(ref)
		return cell?.value ?? EMPTY
	}

	fillFormula(rangeRef: RangeSelector, formula: string): ApplyResult {
		const { sheetName, ref } = normalizeRangeSelector(rangeRef, this.wb)
		return this.apply([
			{ op: 'fillFormula', sheet: sheetName, range: ref, formula: normalizeFormulaInput(formula) },
		])
	}

	/**
	 * Add a new sheet.
	 * @example
	 * wb.addSheet('Data')
	 */
	addSheet(name: string): ApplyResult {
		return this.apply([{ op: 'addSheet', name }])
	}

	/**
	 * Return a builder for batched operations.
	 * @example
	 * wb.builder().set('A1', 1).set('B1', 2).commit()
	 */
	builder(): BatchBuilder {
		return new BatchBuilder(this)
	}

	/**
	 * Delete a sheet.
	 * @example
	 * wb.deleteSheet('Sheet2')
	 */
	deleteSheet(sheet: string): ApplyResult {
		return this.apply([{ op: 'deleteSheet', sheet }])
	}

	renameSheet(sheet: string, newName: string): ApplyResult {
		return this.apply([{ op: 'renameSheet', sheet, newName }])
	}

	insertRows(sheet: string, at: number, count: number): ApplyResult {
		return this.apply([{ op: 'insertRows', sheet, at, count }])
	}

	deleteRows(sheet: string, at: number, count: number): ApplyResult {
		return this.apply([{ op: 'deleteRows', sheet, at, count }])
	}

	insertCols(sheet: string, at: number, count: number): ApplyResult {
		return this.apply([{ op: 'insertCols', sheet, at, count }])
	}

	deleteCols(sheet: string, at: number, count: number): ApplyResult {
		return this.apply([{ op: 'deleteCols', sheet, at, count }])
	}

	// --- Export ---

	/**
	 * Save the workbook to a file. Supports .xlsx, .xlsm, .csv, .tsv.
	 * @example
	 * await wb.save('./output.xlsx')
	 */
	async save(path: string): Promise<void> {
		this.assertWritable()
		const ext = path.split('.').pop()?.toLowerCase() ?? ''

		if (ext === 'csv' || ext === 'tsv') {
			const result =
				ext === 'tsv' ? writeCsv(this.wb, { dialect: { delimiter: '\t' } }) : writeCsv(this.wb)
			if (!result.ok) throw new AscendException(result.error)
			await writeFile(path, result.value, 'utf-8')
			return
		}

		const bytes = this.toBytes()
		await writeFile(path, bytes)
	}

	/**
	 * Serialize the workbook to XLSX bytes.
	 * @example
	 * const bytes = wb.toBytes()
	 */
	toBytes(): Uint8Array {
		this.assertWritable()
		if (this.originalBytes && !this.dirty) return this.originalBytes
		const sourceArchive = this.getSourceArchive()
		const writeOptions: import('@ascend/io-xlsx').WriteXlsxOptions = {
			dirtySheetNames: [...this.dirtySheets],
			workbookMetaDirty: this.workbookMetaDirty,
			documentPropertiesDirty: this.documentPropertiesDirty,
			calcStateDirty: this.calcStateDirty,
			calcChainDirty: this.calcChainDirty,
			sharedStringsDirty: this.sharedStringsDirty,
			stylesDirty: this.stylesDirty,
			...(sourceArchive ? { sourceArchive } : {}),
		}
		const result = writeXlsx(this.wb, this.caps.length > 0 ? this.caps : undefined, writeOptions)
		if (!result.ok) throw new AscendException(result.error)
		this.captureSerializedState(result.value)
		return result.value
	}

	async *toStream(): AsyncGenerator<Uint8Array> {
		yield this.toBytes()
	}

	/**
	 * Export the workbook or a sheet/range to CSV.
	 * @example
	 * const csv = wb.toCsv()
	 * const csv2 = wb.toCsv({ sheet: 'Data', dialect: { delimiter: ';' } })
	 */
	toCsv(opts?: { sheet?: string; range?: string; dialect?: Partial<CsvDialect> }): string {
		this.assertWritable()
		const result = writeCsv(this.wb, opts)
		if (!result.ok) throw new AscendException(result.error)
		return result.value
	}

	writePlanSummary(): WritePlanInfo {
		this.assertWritable()
		const sourceArchive = this.getSourceArchive(false)
		const writeOptions: import('@ascend/io-xlsx').WriteXlsxOptions = {
			dirtySheetNames: [...this.dirtySheets],
			workbookMetaDirty: this.workbookMetaDirty,
			documentPropertiesDirty: this.documentPropertiesDirty,
			calcStateDirty: this.calcStateDirty,
			calcChainDirty: this.calcChainDirty,
			sharedStringsDirty: this.sharedStringsDirty,
			stylesDirty: this.stylesDirty,
			...(sourceArchive ? { sourceArchive } : {}),
		}
		const result = summarizePlannedWrite(
			this.wb,
			this.caps.length > 0 ? this.caps : undefined,
			writeOptions,
		)
		if (!result.ok) throw new AscendException(result.error)
		return result.value
	}

	private assertWritable(): void {
		if (!this.loadInfo.isPartial) return
		throw new AscendException(
			ascendError(
				'EXPORT_ERROR',
				'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
			),
		)
	}

	private deriveDirtyFlags(ops: readonly Operation[]): {
		workbookMetaDirty: boolean
		documentPropertiesDirty: boolean
		calcChainDirty: boolean
		sharedStringsDirty: boolean
		stylesDirty: boolean
	} {
		let workbookMetaDirty = false
		let documentPropertiesDirty = false
		let calcChainDirty = false
		let sharedStringsDirty = false
		let stylesDirty = false
		let sharedStringKeys: Set<string> | null = null
		const getSharedStringKeys = (): Set<string> => {
			if (sharedStringKeys) return sharedStringKeys
			sharedStringKeys = collectSharedStringKeys(this.wb)
			return sharedStringKeys
		}
		for (const op of ops) {
			switch (op.op) {
				case 'addSheet':
				case 'deleteSheet':
				case 'renameSheet':
				case 'moveSheet':
				case 'setDefinedName':
				case 'deleteDefinedName':
				case 'setPivotCache':
				case 'setPivotFieldItem':
				case 'setConnectionRefresh':
				case 'setTimelineRange':
				case 'rewriteExternalLink':
					workbookMetaDirty = true
					calcChainDirty = true
					break
				case 'setDocumentProperties':
					workbookMetaDirty = true
					documentPropertiesDirty = true
					break
				case 'setTheme':
					workbookMetaDirty = true
					break
				case 'setFormula':
				case 'fillFormula':
				case 'insertRows':
				case 'deleteRows':
				case 'insertCols':
				case 'deleteCols':
				case 'createTable':
				case 'appendRows':
				case 'sortRange':
				case 'copySheet':
				case 'copyRange':
				case 'moveRange':
				case 'deleteTable':
				case 'renameTable':
				case 'resizeTable':
				case 'setTableColumn':
					calcChainDirty = true
					if (op.op === 'setFormula' || op.op === 'fillFormula') sharedStringsDirty = true
					if (op.op !== 'appendRows') break
					if (
						op.rows.some((row) =>
							row.some(
								(value) =>
									typeof value === 'string' &&
									!getSharedStringKeys().has(makePlainSharedStringKey(value)),
							),
						)
					) {
						sharedStringsDirty = true
					}
					break
				case 'setRichText':
					sharedStringsDirty = true
					break
				case 'setNumberFormat':
				case 'setStyle':
					stylesDirty = true
					break
				case 'setCells':
					if (
						op.updates.some((update) => {
							if (typeof update.value !== 'string') return false
							return !getSharedStringKeys().has(makePlainSharedStringKey(update.value))
						})
					) {
						sharedStringsDirty = true
					}
					break
				case 'clearRange':
					if (op.what === 'formulas' || op.what === 'all') calcChainDirty = true
					if (op.what === 'styles' || op.what === 'all') stylesDirty = true
					break
			}
		}
		return {
			workbookMetaDirty,
			documentPropertiesDirty,
			calcChainDirty,
			sharedStringsDirty,
			stylesDirty,
		}
	}

	private markDirty(): void {
		if (!this.dirty) this.originalBytes = null
		this.dirty = true
	}

	private captureSerializedState(bytes: Uint8Array): void {
		this.originalBytes = bytes
		this.wb.sourceArchiveBytes = bytes
		this.sourceArchive = undefined
		this.dirty = false
		this.dirtySheets.clear()
		this.workbookMetaDirty = false
		this.documentPropertiesDirty = false
		this.calcStateDirty = false
		this.calcChainDirty = false
		this.sharedStringsDirty = false
		this.stylesDirty = false
		this.pendingDirtyRefs = []
		this.pendingFullRecalc = false
	}

	private mergePendingRecalcTargets(ops: readonly Operation[]): void {
		const refs = deriveRecalcTargets(ops)
		if (refs.fullRecalc) {
			this.pendingFullRecalc = true
			this.pendingDirtyRefs = []
			return
		}
		if (!this.pendingFullRecalc) {
			this.pendingDirtyRefs = [...new Set([...this.pendingDirtyRefs, ...refs.dirtyRefs])]
		}
	}

	private getSourceArchive(cache = true): ZipArchive | undefined {
		if (!this.wb.sourceArchiveBytes) return undefined
		if (!cache) return extractZip(this.wb.sourceArchiveBytes)
		if (this.sourceArchive) return this.sourceArchive
		this.sourceArchive = extractZip(this.wb.sourceArchiveBytes)
		return this.sourceArchive
	}
}

/** Convenient entry point: `import { Ascend } from '@ascend/sdk'` then `Ascend.create()`, `Ascend.open(bytes)`, `Ascend.fromCsv(csv)`. */
export const Ascend = {
	open: AscendWorkbook.open,
	create: AscendWorkbook.create,
	fromCsv: AscendWorkbook.fromCsv,
	listOperations,
	getOperationsSchema,
	parseOperations,
	listCapabilities,
	getCapability,
	summarizeCapabilities,
}

export class BatchBuilder {
	private ops: Operation[] = []

	constructor(private wb: AscendWorkbook) {}

	set(cellRef: CellSelector, value: InputValue): this {
		const { sheet, ref } = parseCellRef(cellRef, this.wb.getWorkbookModel())
		const last = this.ops[this.ops.length - 1]
		if (last?.op === 'setCells' && last.sheet === sheet) {
			this.ops[this.ops.length - 1] = {
				op: 'setCells',
				sheet,
				updates: [...last.updates, { ref, value }],
			}
			return this
		}
		this.ops.push({ op: 'setCells', sheet, updates: [{ ref, value }] })
		return this
	}

	formula(cellRef: CellSelector, formula: string): this {
		const { sheet, ref } = parseCellRef(cellRef, this.wb.getWorkbookModel())
		this.ops.push({ op: 'setFormula', sheet, ref, formula: normalizeFormulaInput(formula) })
		return this
	}

	addSheet(name: string): this {
		this.ops.push({ op: 'addSheet', name })
		return this
	}

	deleteSheet(sheet: string): this {
		this.ops.push({ op: 'deleteSheet', sheet })
		return this
	}

	insertRows(sheet: string, at: number, count: number): this {
		this.ops.push({ op: 'insertRows', sheet, at, count })
		return this
	}

	deleteRows(sheet: string, at: number, count: number): this {
		this.ops.push({ op: 'deleteRows', sheet, at, count })
		return this
	}

	insertCols(sheet: string, at: number, count: number): this {
		this.ops.push({ op: 'insertCols', sheet, at, count })
		return this
	}

	deleteCols(sheet: string, at: number, count: number): this {
		this.ops.push({ op: 'deleteCols', sheet, at, count })
		return this
	}

	clearRange(sheet: string, range: string): this {
		this.ops.push({ op: 'clearRange', sheet, range, what: 'all' })
		return this
	}

	mergeCells(sheet: string, range: string): this {
		this.ops.push({ op: 'mergeCells', sheet, range })
		return this
	}

	unmergeCells(sheet: string, range: string): this {
		this.ops.push({ op: 'unmergeCells', sheet, range })
		return this
	}

	commit(): ApplyResult {
		return this.wb.apply(this.ops)
	}

	commitAndRecalc(opts?: { range?: string }): ApplyAndRecalcResult {
		return this.wb.applyAndRecalc(this.ops, opts)
	}
}

function parseCellRef(cellRef: CellSelector, workbook: Workbook): { sheet: string; ref: string } {
	const parsed = normalizeCellSelector(cellRef, workbook)
	return { sheet: parsed.sheetName, ref: parsed.ref }
}

function partialWorkbookEditError() {
	return ascendError(
		'VALIDATION_ERROR',
		'Cannot modify a partial workbook view. Reopen the workbook with a full load before applying edits or recalculation.',
		{
			suggestedFix:
				'Open the workbook with mode "full" and all sheets loaded before editing or recalculating.',
		},
	)
}

function buildFastPreviewDiff(
	before: Workbook,
	after: Workbook,
	ops: readonly Operation[],
	applyResult: PatchResult,
	recalcResult: RecalcResult | undefined,
): import('@ascend/engine').WorkbookDiff | undefined {
	if (!ops.every(isFastPreviewOp)) return undefined
	const candidateRefs = collectFastPreviewRefs(ops, applyResult, recalcResult)
	const bySheet = new Map<string, import('@ascend/engine').CellChange[]>()
	for (const fullRef of candidateRefs) {
		const parsed = splitFullRef(fullRef, after)
		if (!parsed) continue
		const beforeCell = before.getSheet(parsed.sheet)?.cells.get(parsed.row, parsed.col)
		const afterCell = after.getSheet(parsed.sheet)?.cells.get(parsed.row, parsed.col)
		if (!beforeCell && !afterCell) continue
		const beforeValue = beforeCell?.value ?? EMPTY
		const afterValue = afterCell?.value ?? EMPTY
		const beforeFormula = beforeCell?.formula ?? null
		const afterFormula = afterCell?.formula ?? null
		if (beforeFormula === afterFormula && diffCellValuesEqual(beforeValue, afterValue)) continue
		const change: import('@ascend/engine').CellChange = {
			ref: parsed.ref,
			before: beforeValue,
			after: afterValue,
			formulaBefore: beforeFormula,
			formulaAfter: afterFormula,
		}
		const sheetChanges = bySheet.get(parsed.sheet)
		if (sheetChanges) sheetChanges.push(change)
		else bySheet.set(parsed.sheet, [change])
	}
	return {
		sheets: [...bySheet.entries()].map(([name, cellsChanged]) => ({
			name,
			cellsAdded: [],
			cellsRemoved: [],
			cellsChanged,
		})),
		namesAdded: [],
		namesRemoved: [],
		namesChanged: [],
		workbookProtectionChanged: false,
		sheetFeatures: [],
	}
}

function isFastPreviewOp(op: Operation): boolean {
	switch (op.op) {
		case 'setCells':
		case 'setFormula':
		case 'setComment':
		case 'setHyperlink':
			return true
		case 'setNumberFormat':
		case 'setStyle':
			return !op.range.includes(':') || isSingleCellRef(op.range)
		default:
			return false
	}
}

function isSingleCellRef(ref: string): boolean {
	try {
		const range = parseRange(ref)
		return range.start.row === range.end.row && range.start.col === range.end.col
	} catch {
		return false
	}
}

function collectFastPreviewRefs(
	ops: readonly Operation[],
	applyResult: PatchResult,
	recalcResult: RecalcResult | undefined,
): Set<string> {
	const refs = new Set<string>()
	for (const op of ops) {
		switch (op.op) {
			case 'setCells':
				for (const update of op.updates) refs.add(`${op.sheet}!${update.ref}`)
				break
			case 'setFormula':
			case 'setComment':
			case 'setHyperlink':
				refs.add(`${op.sheet}!${op.ref}`)
				break
			case 'setNumberFormat':
			case 'setStyle':
				if (isSingleCellRef(op.range)) refs.add(`${op.sheet}!${op.range}`)
				break
		}
	}
	for (const ref of applyResult.affectedCells) {
		if (ref.includes('!')) refs.add(ref)
		else if (applyResult.sheetsModified.length === 1)
			refs.add(`${applyResult.sheetsModified[0]}!${ref}`)
	}
	for (const ref of recalcResult?.changed ?? []) refs.add(ref)
	return refs
}

function splitFullRef(
	fullRef: string,
	workbook: Workbook,
): { sheet: string; ref: string; row: number; col: number } | undefined {
	const bang = fullRef.lastIndexOf('!')
	const sheet =
		bang === -1 ? workbook.sheets[0]?.name : fullRef.slice(0, bang).replace(/^'|'$/g, '')
	const ref = bang === -1 ? fullRef : fullRef.slice(bang + 1)
	if (!sheet || !ref) return undefined
	try {
		const parsed = parseA1(ref)
		return { sheet, ref, row: parsed.row, col: parsed.col }
	} catch {
		const range = parseRange(ref)
		if (range.start.row !== range.end.row || range.start.col !== range.end.col) return undefined
		const singleRef = `${indexToColumn(range.start.col)}${range.start.row + 1}`
		return { sheet, ref: singleRef, row: range.start.row, col: range.start.col }
	}
}

function deriveRecalcTargets(ops: readonly Operation[]): {
	fullRecalc: boolean
	dirtyRefs: readonly string[]
} {
	const refs = deriveDirtyRefsFromOps(ops)
	if (refs === null) {
		return { fullRecalc: true, dirtyRefs: [] }
	}
	return { fullRecalc: false, dirtyRefs: refs }
}

function collectSharedStringKeys(workbook: Workbook): Set<string> {
	const keys = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const [, , cell] of sheet.cells.iterate()) {
			const key = makeSharedStringKey(cell.value)
			if (key) keys.add(key)
		}
	}
	return keys
}

function makePlainSharedStringKey(value: string): string {
	return `s:${value}`
}

function makeSharedStringKey(value: import('@ascend/schema').CellValue | string): string | null {
	if (typeof value === 'string') return makePlainSharedStringKey(value)
	if (value.kind === 'string') return `s:${value.value}`
	if (value.kind === 'richText') return `r:${JSON.stringify(value.runs)}`
	return null
}

function resolveRecalcOptions(
	fullRecalc: boolean,
	dirtyRefs: readonly string[],
): { dirtyOnly?: boolean; dirtyRefs?: readonly string[] } | undefined {
	if (fullRecalc) return undefined
	if (dirtyRefs.length === 0) return undefined
	return { dirtyOnly: true, dirtyRefs }
}

function deriveDirtyRefsFromOps(ops: readonly Operation[]): string[] | null {
	const refs: string[] = []
	for (const op of ops) {
		switch (op.op) {
			case 'setCells':
				refs.push(...op.updates.map((update) => `${op.sheet}!${update.ref}`))
				break
			case 'setFormula':
				refs.push(`${op.sheet}!${op.ref}`)
				break
			case 'fillFormula':
				refs.push(`${op.sheet}!${op.range}`)
				break
			case 'clearRange':
				if (op.what !== 'styles') refs.push(`${op.sheet}!${op.range}`)
				break
			case 'appendRows':
				return null
			case 'insertRows':
			case 'deleteRows':
			case 'insertCols':
			case 'deleteCols':
			case 'addSheet':
			case 'deleteSheet':
			case 'renameSheet':
			case 'moveSheet':
			case 'createTable':
			case 'sortRange':
			case 'mergeCells':
			case 'unmergeCells':
			case 'setDefinedName':
			case 'deleteDefinedName':
				return null
			default:
				break
		}
	}
	return refs
}
