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
	type CalcContext,
	defaultCalcContext,
	cellValuesEqual as diffCellValuesEqual,
	diffWorkbooks,
	type PatchResult,
	recalculate,
} from '@ascend/engine'
import { normalizeFormulaInput } from '@ascend/formulas'
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
	type CompatibilityReport,
	type CsvDialect,
	EMPTY,
	emptyReport,
	type Operation,
} from '@ascend/schema'
import { check as verifyCheck, lint as verifyLint } from '@ascend/verify'
import { buildWorkbookLoadInfo, openWorkbookSource } from './load.ts'
import { parseFullRef, WorkbookReadView } from './read-view.ts'
import type {
	ApplyAndRecalcResult,
	ApplyResult,
	BatchResult,
	CheckIssue,
	CheckResult,
	LintResult,
	LintWarning,
	RecalcResult,
	WritePlanInfo,
} from './types.ts'

function cloneWorkbook(source: Workbook): Workbook {
	return source.clone()
}

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
	private calcStateDirty = false
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

	static async open(
		pathOrBytes: string | Uint8Array,
		options?: {
			mode?: 'full' | 'metadata-only' | 'values' | 'formula'
			sheets?: readonly string[]
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
				diff: { sheets: [], namesAdded: [], namesRemoved: [], namesChanged: [] },
				sheetDiffs: [],
				cellChanges: [],
				errors: [partialWorkbookEditError()],
			}
		}
		const clone = cloneWorkbook(this.wb)
		const errors: import('@ascend/schema').AscendError[] = []

		const result = applyOperations(clone, ops)
		if (!result.ok) {
			errors.push(result.error)
			return {
				diff: { sheets: [], namesAdded: [], namesRemoved: [], namesChanged: [] },
				sheetDiffs: [],
				cellChanges: [],
				errors,
			}
		}

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
		let cachedWritePlan: WritePlanInfo | undefined
		let writePlanComputed = false
		let cachedDirtyFlags:
			| {
					workbookMetaDirty: boolean
					sharedStringsDirty: boolean
					stylesDirty: boolean
			  }
			| undefined
		const previewResult: import('./types.ts').PreviewResult = {
			diff,
			sheetDiffs: diff.sheets,
			cellChanges,
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
						calcStateDirty: cachedDirtyFlags.workbookMetaDirty || result.value.recalcRequired,
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

	apply(ops: readonly Operation[]): ApplyResult {
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
		const result = applyOperations(nextWorkbook, ops)
		if (!result.ok) {
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				errors: [result.error],
			}
		}

		this.wb = nextWorkbook
		this.markDirty()
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.calcStateDirty ||= dirtyFlags.workbookMetaDirty || result.value.recalcRequired
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		this.stylesDirty ||= dirtyFlags.stylesDirty
		if (result.value.recalcRequired) this.mergePendingRecalcTargets(ops)
		if (!this._batchMode) this.clearReadCaches()
		return {
			affectedCells: result.value.affectedCells,
			sheetsModified: result.value.sheetsModified,
			recalcRequired: result.value.recalcRequired,
			errors: [],
		}
	}

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

	/** Apply an array of operations atomically and recalculate. */
	batch(ops: readonly Operation[]): BatchResult
	/** Execute `fn` (which may call `apply` multiple times) and defer recalculation until `fn` completes. */
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
			return { errors: [result.error] }
		}

		this.wb = nextWorkbook
		this.markDirty()
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.calcStateDirty ||= dirtyFlags.workbookMetaDirty || result.value.recalcRequired
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		this.stylesDirty ||= dirtyFlags.stylesDirty
		if (result.value.recalcRequired) this.mergePendingRecalcTargets(opsOrFn)
		this.clearReadCaches()
		return { errors: [] }
	}

	recalc(opts?: { range?: string }): RecalcResult {
		if (this.loadInfo.isPartial) {
			return {
				changed: [],
				errors: [{ ref: '', error: partialWorkbookEditError() }],
				duration: 0,
			}
		}
		const ctx: CalcContext = defaultCalcContext({
			dateSystem: this.wb.calcSettings.dateSystem,
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

	// --- Verification ---

	check(): CheckResult {
		const issue = this.dependencyVerificationIssue()
		if (issue) {
			return {
				valid: false,
				issues: [{ severity: 'warning', message: issue }],
			}
		}
		const result = verifyCheck(this.wb, {
			formulas: this.formulaAnalysis(),
			dependencies: this.dependencyAnalysis(),
		})
		const issues: CheckIssue[] = result.issues.map((issue) =>
			issue.refs?.[0]
				? {
						severity: issue.severity === 'info' ? 'warning' : issue.severity,
						message: issue.message,
						ref: issue.refs[0],
					}
				: {
						severity: issue.severity === 'info' ? 'warning' : issue.severity,
						message: issue.message,
					},
		)
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

	setFormula(cellRef: string, formula: string): ApplyResult {
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		return this.apply([
			{ op: 'setFormula', sheet: sheetName, ref, formula: normalizeFormulaInput(formula) },
		])
	}

	fillFormula(rangeRef: string, formula: string): ApplyResult {
		const { sheetName, ref } = parseFullRef(rangeRef, this.wb)
		return this.apply([
			{ op: 'fillFormula', sheet: sheetName, range: ref, formula: normalizeFormulaInput(formula) },
		])
	}

	// --- Export ---

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

	toBytes(): Uint8Array {
		this.assertWritable()
		if (this.originalBytes && !this.dirty) return this.originalBytes
		const sourceArchive = this.getSourceArchive()
		const writeOptions: import('@ascend/io-xlsx').WriteXlsxOptions = {
			dirtySheetNames: [...this.dirtySheets],
			workbookMetaDirty: this.workbookMetaDirty,
			calcStateDirty: this.calcStateDirty,
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
			calcStateDirty: this.calcStateDirty,
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
		sharedStringsDirty: boolean
		stylesDirty: boolean
	} {
		let workbookMetaDirty = false
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
					workbookMetaDirty = true
					break
				case 'setFormula':
				case 'fillFormula':
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
				case 'appendRows':
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
				case 'clearRange':
					if (op.what === 'styles' || op.what === 'all') stylesDirty = true
					break
			}
		}
		return { workbookMetaDirty, sharedStringsDirty, stylesDirty }
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
		this.calcStateDirty = false
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
