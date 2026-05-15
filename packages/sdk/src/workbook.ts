import { readFile, writeFile } from 'node:fs/promises'
import {
	type Cell,
	createWorkbook,
	indexToColumn,
	parseA1,
	parseA1Safe,
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
	inspectXlsxPackageGraph,
	type PreservationCapsule,
	summarizePlannedWrite,
	writeXlsx,
	type XlsxPackageGraph,
	type ZipArchive,
	type ZipCompressionProfile,
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
import {
	partialDependencyCheckIssue,
	partialDependencyLintWarning,
	sdkCheckIssueFromVerify,
} from './check-issues.ts'
import {
	buildMutationJournal,
	failedMutationJournal,
	type MutationJournal,
	unavailableMutationJournal,
} from './journal.ts'
import {
	buildWorkbookLoadInfo,
	type LoadedWorkbookSource,
	type OpenWorkbookSourceOptions,
	openWorkbookSource,
} from './load.ts'
import { getOperationsSchema, listOperations, parseOperations } from './ops.ts'
import { compilePathMutations } from './path-mutations.ts'
import { inspectRawPackagePart } from './raw-package.ts'
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
	DirtyRegion,
	DumpBatchOptions,
	DumpBatchResult,
	DumpBatchUnsupportedCell,
	EvalOptions,
	LintResult,
	LintWarning,
	PathMutation,
	PathMutationResult,
	PivotOutputMaterializeOptions,
	PivotOutputMaterializeResult,
	RawPackagePartInfo,
	RawPackagePartOptions,
	RecalcOptions,
	RecalcResult,
	TemplateMergeOptions,
	TemplateMergePlaceholder,
	TemplateMergeResult,
	TemplateMergeUnsupportedCell,
	TemplateMergeValue,
	WorkbookGenerationInfo,
	WorkbookLoadInfo,
	WritePlanInfo,
} from './types.ts'
import type { WorkbookTrustReport, WorkbookTrustReportOptions } from './workbook-trust.ts'

export interface PreviewOptions {
	readonly journal?: boolean
}

export interface ApplyOptions {
	readonly collectAllErrors?: boolean
	readonly transaction?: boolean
	readonly journal?: boolean
}

export interface WorkbookBytesOptions {
	readonly compressionProfile?: ZipCompressionProfile
}

interface WorkbookMutationRollbackSnapshot {
	readonly workbook: Workbook
	readonly originalBytes: Uint8Array | null
	readonly sourceArchive: ZipArchive | undefined
	readonly packageGraphCache:
		| { readonly bytes: Uint8Array; readonly graph: XlsxPackageGraph }
		| undefined
	readonly dirty: boolean
	readonly dirtySheets: readonly string[]
	readonly pendingDirtyCellRefs: readonly (readonly [string, readonly string[]])[]
	readonly pendingDirtyRefs: readonly string[]
	readonly pendingFullRecalc: boolean
	readonly workbookMetaDirty: boolean
	readonly documentPropertiesDirty: boolean
	readonly calcStateDirty: boolean
	readonly calcChainDirty: boolean
	readonly sharedStringsDirty: boolean
	readonly stylesDirty: boolean
	readonly workbookGeneration: number
	readonly sheetMetadataGeneration: number
	readonly formulaGeneration: number
	readonly styleGeneration: number
}

function cloneWorkbook(source: Workbook): Workbook {
	return source.clone()
}

function maybeBuildMutationJournal(
	workbook: Workbook,
	ops: readonly Operation[],
	enabled: boolean | undefined,
) {
	if (!enabled) return undefined
	try {
		return buildMutationJournal(workbook, ops)
	} catch (error) {
		return failedMutationJournal(error)
	}
}

function partialWorkbookMutationJournal(loadInfo: WorkbookLoadInfo) {
	return unavailableMutationJournal(
		`Mutation journal is unavailable because the workbook is partially loaded in ${loadInfo.mode} mode. Reopen the workbook with a full load before applying edits.`,
		undefined,
		{ reason: 'partial-workbook' },
	)
}

function failedApplyMutationJournal(
	journal: MutationJournal | undefined,
): MutationJournal | undefined {
	if (!journal) return undefined
	if (journal.issues.some((issue) => issue.code === 'JOURNAL_BUILD_FAILED')) return journal
	return unavailableMutationJournal(
		'Mutation journal is unavailable because the requested operations did not apply successfully. Fix the apply errors before using rollback journal.',
	)
}

interface CellRollbackSnapshot {
	readonly sheet: string
	readonly row: number
	readonly col: number
	readonly cell: Cell | undefined
}

function collectSetCellsRollbackSnapshots(
	workbook: Workbook,
	ops: readonly Operation[],
): CellRollbackSnapshot[] | null {
	const snapshots: CellRollbackSnapshot[] = []
	const seen = new Set<string>()
	for (const op of ops) {
		if (op.op !== 'setCells') return null
		const sheet = workbook.getSheet(op.sheet)
		if (!sheet) continue
		for (const update of op.updates) {
			const ref = parseA1Safe(update.ref)
			if (!ref) return null
			const key = `${sheet.id}:${ref.row}:${ref.col}`
			if (seen.has(key)) continue
			seen.add(key)
			snapshots.push({
				sheet: op.sheet,
				row: ref.row,
				col: ref.col,
				cell: sheet.cells.get(ref.row, ref.col),
			})
		}
	}
	return snapshots
}

function restoreCellRollbackSnapshots(
	workbook: Workbook,
	snapshots: readonly CellRollbackSnapshot[],
): void {
	for (let i = snapshots.length - 1; i >= 0; i--) {
		const snapshot = snapshots[i]
		if (!snapshot) continue
		const sheet = workbook.getSheet(snapshot.sheet)
		if (!sheet) continue
		if (snapshot.cell) sheet.cells.set(snapshot.row, snapshot.col, snapshot.cell)
		else sheet.cells.delete(snapshot.row, snapshot.col)
	}
}

function applySetCellsTransactionInPlace(
	workbook: Workbook,
	ops: readonly Operation[],
	options?: ApplyOptions,
): ReturnType<typeof applyOperations> | null {
	if (options?.collectAllErrors) return null
	const snapshots = collectSetCellsRollbackSnapshots(workbook, ops)
	if (!snapshots) return null
	const result = applyOperations(workbook, ops, options)
	if (!result.ok) restoreCellRollbackSnapshots(workbook, snapshots)
	return result
}

function isApplyPatchNoOp(ops: readonly Operation[], result: PatchResult): boolean {
	return (
		ops.every(
			(op) =>
				op.op === 'setCells' ||
				op.op === 'setComment' ||
				op.op === 'setHyperlink' ||
				op.op === 'setNumberFormat' ||
				op.op === 'setStyle' ||
				op.op === 'clearRange',
		) &&
		result.affectedCells.length === 0 &&
		result.sheetsModified.length === 0 &&
		!result.recalcRequired &&
		(!result.warnings || result.warnings.length === 0)
	)
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

function dumpableInputValue(
	value: CellValue,
): { ok: true; value: InputValue } | { ok: false; reason: string } {
	switch (value.kind) {
		case 'empty':
			return { ok: true, value: null }
		case 'number':
			return { ok: true, value: value.value }
		case 'string':
			return { ok: true, value: value.value }
		case 'boolean':
			return { ok: true, value: value.value }
		case 'date':
			return {
				ok: false,
				reason: 'Date cells require style-aware serialization before replay.',
			}
		case 'error':
			return {
				ok: false,
				reason: 'Error values require formula/error-value operation support before replay.',
			}
		case 'richText':
			return {
				ok: false,
				reason: 'Rich text requires setRichText operation support before replay.',
			}
		case 'array':
			return {
				ok: false,
				reason: 'Array values require spill-aware formula serialization before replay.',
			}
	}
}

interface TemplateDelimiters {
	readonly open: string
	readonly close: string
}

interface TemplatePlaceholderMatch {
	readonly key: string
	readonly placeholder: string
	readonly start: number
	readonly end: number
}

function normalizeTemplateDelimiters(options: TemplateMergeOptions): TemplateDelimiters {
	return {
		open: options.delimiters?.open ?? '{{',
		close: options.delimiters?.close ?? '}}',
	}
}

function findTemplatePlaceholders(
	text: string,
	delimiters: TemplateDelimiters,
): TemplatePlaceholderMatch[] {
	if (delimiters.open.length === 0 || delimiters.close.length === 0) return []
	const matches: TemplatePlaceholderMatch[] = []
	let cursor = 0
	while (cursor < text.length) {
		const start = text.indexOf(delimiters.open, cursor)
		if (start === -1) break
		const contentStart = start + delimiters.open.length
		const close = text.indexOf(delimiters.close, contentStart)
		if (close === -1) break
		const end = close + delimiters.close.length
		const key = text.slice(contentStart, close).trim()
		if (key.length > 0) {
			matches.push({
				key,
				placeholder: text.slice(start, end),
				start,
				end,
			})
		}
		cursor = end
	}
	return matches
}

function missingTemplatePlaceholders(
	sheet: string,
	ref: string,
	source: 'value' | 'formula',
	matches: readonly TemplatePlaceholderMatch[],
	data: ReadonlyMap<string, TemplateMergeValue>,
): TemplateMergePlaceholder[] {
	const missing: TemplateMergePlaceholder[] = []
	for (const match of matches) {
		if (!data.has(match.key)) {
			missing.push({
				sheet,
				ref,
				source,
				placeholder: match.placeholder,
				key: match.key,
			})
		}
	}
	return missing
}

function replaceTemplatePlaceholders(
	text: string,
	matches: readonly TemplatePlaceholderMatch[],
	valueFor: (value: TemplateMergeValue) => string,
	data: ReadonlyMap<string, TemplateMergeValue>,
): string {
	let output = ''
	let cursor = 0
	for (const match of matches) {
		output += text.slice(cursor, match.start)
		output += valueFor(data.get(match.key) ?? null)
		cursor = match.end
	}
	output += text.slice(cursor)
	return output
}

function templateTextValue(value: TemplateMergeValue): string {
	return value === null ? '' : String(value)
}

function templateFormulaValue(value: TemplateMergeValue): string {
	if (value === null) return '""'
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
	if (typeof value === 'number') return String(value)
	return `"${value.replaceAll('"', '""')}"`
}

function richTextPlainText(value: Extract<CellValue, { kind: 'richText' }>): string {
	return value.runs.map((run) => run.text).join('')
}

function normalizeTemplateMergeData(
	data: Readonly<Record<string, unknown>>,
): ReadonlyMap<string, TemplateMergeValue> {
	const values = new Map<string, TemplateMergeValue>()
	for (const [key, value] of Object.entries(data)) {
		if (
			value === null ||
			typeof value === 'string' ||
			typeof value === 'number' ||
			typeof value === 'boolean'
		) {
			values.set(key, value)
			continue
		}
		throw new AscendException(
			ascendError(
				'INVALID_ARGUMENT',
				`Template value "${key}" must be string, number, boolean, or null`,
				{
					details: {
						key,
						valueType: Array.isArray(value) ? 'array' : typeof value,
					},
					suggestedFix: 'Pass only scalar JSON values for template merge data.',
				},
			),
		)
	}
	return values
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
	private packageGraphCache:
		| { readonly bytes: Uint8Array; readonly graph: XlsxPackageGraph }
		| undefined
	private dirty: boolean
	private readonly dirtySheets = new Set<string>()
	private readonly pendingDirtyCellRefs = new Map<string, Set<string>>()
	private pendingDirtyRefs: string[] = []
	private pendingFullRecalc = false
	private _batchMode = false
	private workbookMetaDirty = false
	private documentPropertiesDirty = false
	private calcStateDirty = false
	private calcChainDirty = false
	private sharedStringsDirty = false
	private stylesDirty = false
	private workbookGeneration = 0
	private sheetMetadataGeneration = 0
	private formulaGeneration = 0
	private styleGeneration = 0

	private constructor(
		workbook: Workbook,
		capsules: PreservationCapsule[],
		report: CompatibilityReport,
		loadInfo: import('./types.ts').WorkbookLoadInfo,
		originalBytes: Uint8Array | null,
		sourceArchive?: ZipArchive,
	) {
		super(workbook, report, loadInfo)
		this.caps = capsules
		this.originalBytes = originalBytes
		this.sourceArchive = sourceArchive
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
		options?: OpenWorkbookSourceOptions,
	): Promise<AscendWorkbook> {
		const loaded = await openWorkbookSource(pathOrBytes, options)
		return AscendWorkbook.fromLoadedSource(loaded)
	}

	static async openSourceBytes(
		path: string,
		options?: Omit<OpenWorkbookSourceOptions, 'sourceExtension'>,
	): Promise<{ readonly workbook: AscendWorkbook; readonly sourceBytes: Uint8Array }> {
		const sourceBytes =
			typeof Bun !== 'undefined'
				? await Bun.file(path).bytes()
				: new Uint8Array(await readFile(path))
		const sourceExtension = path.split('.').pop()?.toLowerCase() ?? ''
		const loaded = await openWorkbookSource(sourceBytes, { ...options, sourceExtension })
		return { workbook: AscendWorkbook.fromLoadedSource(loaded), sourceBytes }
	}

	static fromLoadedSource(loaded: LoadedWorkbookSource): AscendWorkbook {
		return new AscendWorkbook(
			loaded.workbook,
			[...loaded.capsules],
			loaded.report,
			loaded.loadInfo,
			loaded.originalBytes,
			loaded.sourceArchive,
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
				partialReasons: [],
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
				partialReasons: [],
				sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
				loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
			}),
			null,
		)
	}

	// --- Mutation ---

	preview(ops: readonly Operation[], options?: PreviewOptions): import('./types.ts').PreviewResult {
		if (this.loadInfo.isPartial) {
			const journal = options?.journal ? partialWorkbookMutationJournal(this.loadInfo) : undefined
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
				errors: [partialWorkbookEditError(this.loadInfo)],
				...(journal ? { journal } : {}),
			}
		}
		const journal = maybeBuildMutationJournal(this.wb, ops, options?.journal)
		const clone = cloneWorkbook(this.wb)
		const errors: import('@ascend/schema').AscendError[] = []

		const result = applyOperations(clone, ops)
		if (!result.ok) {
			errors.push(...('errors' in result.error ? result.error.errors : [result.error]))
			const failureJournal = failedApplyMutationJournal(journal)
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
				...(failureJournal ? { journal: failureJournal } : {}),
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
				dirtyRegions: dirtyRegionsFromRefs(calcResult.changed, [], clone),
				generations: this.currentGenerations(),
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
					calcStateDirty: boolean
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
			...(journal ? { journal } : {}),
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
						calcStateDirty: cachedDirtyFlags.calcStateDirty || result.value.recalcRequired,
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
	apply(ops: readonly Operation[], options?: ApplyOptions): ApplyResult {
		if (ops.length === 0) {
			const journal = maybeBuildMutationJournal(this.wb, ops, options?.journal)
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				dirtyRegions: [],
				generations: this.currentGenerations(),
				errors: [],
				...(journal ? { journal } : {}),
			}
		}
		if (this.loadInfo.isPartial) {
			const journal = options?.journal ? partialWorkbookMutationJournal(this.loadInfo) : undefined
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				dirtyRegions: [],
				generations: this.currentGenerations(),
				errors: [partialWorkbookEditError(this.loadInfo)],
				...(journal ? { journal } : {}),
			}
		}
		const journal = maybeBuildMutationJournal(this.wb, ops, options?.journal)
		const dirtyFlags = this.deriveDirtyFlags(ops)
		const nextWorkbook = options?.transaction ? this.wb : cloneWorkbook(this.wb)
		const result = options?.transaction
			? (applySetCellsTransactionInPlace(nextWorkbook, ops, options) ??
				applyWithTransaction(nextWorkbook, ops, options))
			: applyOperations(nextWorkbook, ops, options)
		if (!result.ok) {
			const errors = 'errors' in result.error ? result.error.errors : [result.error]
			const failureJournal = failedApplyMutationJournal(journal)
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				dirtyRegions: [],
				generations: this.currentGenerations(),
				errors,
				...(failureJournal ? { journal: failureJournal } : {}),
			}
		}
		if (isApplyPatchNoOp(ops, result.value)) {
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				dirtyRegions: [],
				generations: this.currentGenerations(),
				errors: [],
				...(journal ? { journal } : {}),
			}
		}
		if (!options?.transaction) this.wb = nextWorkbook
		this.advanceApplyGenerations(ops, dirtyFlags, result.value)
		this.markDirty()
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.mergePendingDirtyCellRefs(ops, result.value)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.documentPropertiesDirty ||= dirtyFlags.documentPropertiesDirty
		this.calcStateDirty ||= dirtyFlags.calcStateDirty || result.value.recalcRequired
		this.calcChainDirty ||= dirtyFlags.calcChainDirty
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		this.stylesDirty ||= dirtyFlags.stylesDirty
		if (result.value.recalcRequired) this.mergePendingRecalcTargets(ops)
		if (!this._batchMode) this.clearReadCaches()
		return {
			affectedCells: result.value.affectedCells,
			sheetsModified: result.value.sheetsModified,
			recalcRequired: result.value.recalcRequired,
			dirtyRegions: dirtyRegionsFromRefs(
				result.value.affectedCells,
				result.value.sheetsModified,
				this.wb,
			),
			generations: this.currentGenerations(),
			errors: [],
			...(journal ? { journal } : {}),
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
	 * Materialize supported PivotTable output cells through ordinary cell writes.
	 * Pivot metadata and preserved PivotTable package parts are left untouched.
	 */
	materializePivotOutputs(
		options: PivotOutputMaterializeOptions = {},
	): PivotOutputMaterializeResult {
		const planned = this.pivotOutputMaterializeOps(options)
		if (planned.ops.length === 0) {
			return {
				...planned,
				apply: {
					affectedCells: [],
					sheetsModified: [],
					recalcRequired: false,
					dirtyRegions: [],
					generations: this.currentGenerations(),
					errors: [],
				},
			}
		}
		return { ...planned, apply: this.apply(planned.ops, { transaction: true }) }
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
					partialWorkbookEditError(this.loadInfo, {
						message:
							'Cannot run batch in a partial workbook view. Reopen with a full load before editing.',
					}),
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
			return { errors: [partialWorkbookEditError(this.loadInfo)] }
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
		this.mergePendingDirtyCellRefs(opsOrFn, result.value)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.documentPropertiesDirty ||= dirtyFlags.documentPropertiesDirty
		this.calcStateDirty ||= dirtyFlags.calcStateDirty || result.value.recalcRequired
		this.calcChainDirty ||= dirtyFlags.calcChainDirty
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		this.stylesDirty ||= dirtyFlags.stylesDirty
		if (result.value.recalcRequired) this.mergePendingRecalcTargets(opsOrFn)
		this.clearReadCaches()
		return { errors: [] }
	}

	/**
	 * Serialize supported cell values and formulas into a deterministic operation batch.
	 * Unsupported value kinds are reported instead of silently lossy-converted.
	 */
	dumpBatch(options: DumpBatchOptions = {}): DumpBatchResult {
		if (this.loadInfo.isPartial) {
			return {
				ops: [],
				sheetCount: 0,
				cellCount: 0,
				formulaCount: 0,
				unsupported: [],
				replayable: false,
				blocked: partialReplayBlockedInfo(
					this.loadInfo,
					'Cannot dump replay operations from a partial workbook view. Reopen with a full load before producing dump batches.',
				),
			}
		}
		const includeValues = options.includeValues ?? true
		const includeFormulas = options.includeFormulas ?? true
		const selectedSheets = options.sheets ? new Set(options.sheets) : null
		const ops: Operation[] = []
		const unsupported: DumpBatchUnsupportedCell[] = []
		let sheetCount = 0
		let cellCount = 0
		let formulaCount = 0

		for (const sheet of this.wb.sheets) {
			if (selectedSheets && !selectedSheets.has(sheet.name)) continue
			sheetCount += 1
			const updates: Array<{ ref: string; value: InputValue }> = []
			const formulaOps: Operation[] = []
			for (const [row, col, cell] of sheet.cells.iterate()) {
				const ref = `${indexToColumn(col)}${row + 1}`
				const formula = sheet.cells.readFormula(row, col)
				if (formula) {
					cellCount += 1
					formulaCount += 1
					if (includeFormulas)
						formulaOps.push({ op: 'setFormula', sheet: sheet.name, ref, formula })
					continue
				}
				if (!includeValues) continue
				const input = dumpableInputValue(cell.value ?? EMPTY)
				if (input.ok) {
					cellCount += 1
					updates.push({ ref, value: input.value })
				} else {
					unsupported.push({
						sheet: sheet.name,
						ref,
						valueKind: (cell.value ?? EMPTY).kind,
						reason: input.reason,
					})
				}
			}
			if (updates.length > 0) ops.push({ op: 'setCells', sheet: sheet.name, updates })
			ops.push(...formulaOps)
		}

		return {
			ops,
			sheetCount,
			cellCount,
			formulaCount,
			unsupported,
			replayable: unsupported.length === 0,
		}
	}

	/**
	 * Compile {{key}} template replacements into replayable operations.
	 * Missing keys and unsupported rich-text placeholders are reported without mutating.
	 */
	templateMerge(
		data: Readonly<Record<string, unknown>>,
		options: TemplateMergeOptions = {},
	): TemplateMergeResult {
		if (this.loadInfo.isPartial) {
			return {
				ops: [],
				sheetCount: 0,
				cellCount: 0,
				formulaCount: 0,
				replacementCount: 0,
				unresolved: [],
				unsupported: [],
				replayable: false,
				blocked: partialReplayBlockedInfo(
					this.loadInfo,
					'Cannot compile template merge replay operations from a partial workbook view. Reopen with a full load before producing template merge batches.',
				),
			}
		}
		const includeValues = options.includeValues ?? true
		const includeFormulas = options.includeFormulas ?? true
		const selectedSheets = options.sheets ? new Set(options.sheets) : null
		const delimiters = normalizeTemplateDelimiters(options)
		const values = normalizeTemplateMergeData(data)
		const ops: Operation[] = []
		const unresolved: TemplateMergePlaceholder[] = []
		const unsupported: TemplateMergeUnsupportedCell[] = []
		let sheetCount = 0
		let cellCount = 0
		let formulaCount = 0
		let replacementCount = 0

		for (const sheet of this.wb.sheets) {
			if (selectedSheets && !selectedSheets.has(sheet.name)) continue
			sheetCount += 1
			const updates: Array<{ ref: string; value: InputValue }> = []
			const formulaOps: Operation[] = []

			for (const [row, col, cell] of sheet.cells.iterate()) {
				const ref = `${indexToColumn(col)}${row + 1}`
				const formula = sheet.cells.readFormula(row, col)
				if (formula) {
					if (!includeFormulas) continue
					const matches = findTemplatePlaceholders(formula, delimiters)
					if (matches.length === 0) continue
					formulaCount += 1
					cellCount += 1
					const missing = missingTemplatePlaceholders(sheet.name, ref, 'formula', matches, values)
					unresolved.push(...missing)
					if (missing.length > 0) continue
					formulaOps.push({
						op: 'setFormula',
						sheet: sheet.name,
						ref,
						formula: replaceTemplatePlaceholders(formula, matches, templateFormulaValue, values),
					})
					replacementCount += matches.length
					continue
				}

				if (!includeValues) continue
				const value = cell.value ?? EMPTY
				const text =
					value.kind === 'string'
						? value.value
						: value.kind === 'richText'
							? richTextPlainText(value)
							: null
				if (text === null) continue
				const matches = findTemplatePlaceholders(text, delimiters)
				if (matches.length === 0) continue
				cellCount += 1
				if (value.kind === 'richText') {
					unsupported.push({
						sheet: sheet.name,
						ref,
						source: 'value',
						valueKind: value.kind,
						reason: 'Rich text placeholders require run-preserving merge support before replay.',
					})
					continue
				}

				const missing = missingTemplatePlaceholders(sheet.name, ref, 'value', matches, values)
				unresolved.push(...missing)
				if (missing.length > 0) continue
				const firstMatch = matches[0]
				const wholeCellPlaceholder =
					matches.length === 1 &&
					firstMatch !== undefined &&
					firstMatch.start === 0 &&
					firstMatch.end === text.length
				const mergedValue = wholeCellPlaceholder
					? (values.get(firstMatch.key) ?? null)
					: replaceTemplatePlaceholders(text, matches, templateTextValue, values)
				updates.push({ ref, value: mergedValue })
				replacementCount += matches.length
			}
			if (updates.length > 0) ops.push({ op: 'setCells', sheet: sheet.name, updates })
			ops.push(...formulaOps)
		}

		return {
			ops,
			sheetCount,
			cellCount,
			formulaCount,
			replacementCount,
			unresolved,
			unsupported,
			replayable: unresolved.length === 0 && unsupported.length === 0,
		}
	}

	/**
	 * Compile stable path-addressed mutations into canonical operations for plan/commit.
	 */
	compilePathMutations(mutations: readonly PathMutation[]): PathMutationResult {
		if (this.loadInfo.isPartial && mutations.length > 0) {
			return partialWorkbookPathMutationResult(this.loadInfo, mutations)
		}
		return compilePathMutations(this.wb, mutations)
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
				dirtyRegions: [],
				generations: this.currentGenerations(),
				errors: [{ ref: '', error: partialWorkbookEditError(this.loadInfo) }],
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
		const rangeLimited = rangeRef !== undefined
		const hadGlobalCalcStaleness =
			this.calcStateDirty ||
			this.calcChainDirty ||
			this.pendingFullRecalc ||
			this.pendingDirtyRefs.length > 0 ||
			calcSettingsNeedClean(this.wb.calcSettings)
		const result = recalculate(
			this.wb,
			ctx,
			rangeRef
				? { range: rangeRef }
				: resolveRecalcOptions(this.pendingFullRecalc, this.pendingDirtyRefs),
		)
		const formulaValuesChangedOrErrored = result.changed.length > 0 || result.errors.length > 0
		if (formulaValuesChangedOrErrored) {
			this.workbookGeneration += 1
			this.formulaGeneration += 1
			this.markDirty()
			this.calcStateDirty = result.errors.length > 0 || (rangeLimited && hadGlobalCalcStaleness)
			this.sharedStringsDirty = true
			for (const ref of result.changed) {
				const bang = ref.indexOf('!')
				if (bang !== -1) {
					const sheetName = ref.slice(0, bang)
					this.dirtySheets.add(sheetName)
					this.pendingDirtyCellRefs.delete(sheetName)
				}
			}
		}
		const keepGlobalCalcStale = rangeLimited && hadGlobalCalcStaleness
		if (result.errors.length === 0 && !keepGlobalCalcStale) {
			const sourceCalcSettings = this.wb.calcSettings
			const cleanCalcSettings = {
				...sourceCalcSettings,
				fullCalcOnLoad: false,
				...(sourceCalcSettings.calcCompleted !== undefined ? { calcCompleted: true } : {}),
				...(sourceCalcSettings.calcOnSave !== undefined ? { calcOnSave: true } : {}),
				...(sourceCalcSettings.forceFullCalc !== undefined ? { forceFullCalc: false } : {}),
			}
			const calcSettingsChanged =
				sourceCalcSettings.fullCalcOnLoad !== cleanCalcSettings.fullCalcOnLoad ||
				sourceCalcSettings.calcCompleted !== cleanCalcSettings.calcCompleted ||
				sourceCalcSettings.calcOnSave !== cleanCalcSettings.calcOnSave ||
				sourceCalcSettings.forceFullCalc !== cleanCalcSettings.forceFullCalc
			if (calcSettingsChanged) {
				if (!formulaValuesChangedOrErrored) {
					this.workbookGeneration += 1
					this.formulaGeneration += 1
				}
				this.markDirty()
				this.workbookMetaDirty = true
			}
			this.calcStateDirty = false
			this.wb.calcSettings = cleanCalcSettings
		}
		this.clearReadCaches()
		if (!keepGlobalCalcStale) {
			this.pendingDirtyRefs = []
			this.pendingFullRecalc = false
		}
		return {
			changed: result.changed,
			dirtyRegions: dirtyRegionsFromRefs(result.changed, [], this.wb),
			generations: this.currentGenerations(),
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
			...(this.wb.sourceArchiveBytes ? { packageGraph: this.packageGraph() } : {}),
		})
		const issues: CheckIssue[] = result.issues.map(sdkCheckIssueFromVerify)
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const issue = this.dependencyVerificationIssue()
		const result = verifyLint(this.wb, this.formulaAnalysis())
		const warnings: LintWarning[] = result.violations.map((violation) => ({
			rule: violation.rule,
			severity: violation.severity,
			message: violation.message,
			ref: violation.ref,
		}))
		if (issue) warnings.unshift(partialDependencyLintWarning(issue))
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
	toBytes(options: WorkbookBytesOptions = {}): Uint8Array {
		this.assertWritable()
		if (this.originalBytes && !this.dirty && !options.compressionProfile) return this.originalBytes
		const sourceArchive = this.getSourceArchive()
		const dirtyCellPatches = this.dirtyCellPatchOptions()
		const writeOptions: import('@ascend/io-xlsx').WriteXlsxOptions = {
			dirtySheetNames: [...this.dirtySheets],
			...(dirtyCellPatches ? { dirtyCellPatches } : {}),
			workbookMetaDirty: this.workbookMetaDirty,
			documentPropertiesDirty: this.documentPropertiesDirty,
			calcStateDirty: this.calcStateDirty,
			calcChainDirty: this.calcChainDirty,
			sharedStringsDirty: this.sharedStringsDirty,
			stylesDirty: this.stylesDirty,
			...(options.compressionProfile ? { compressionProfile: options.compressionProfile } : {}),
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
		const dirtyCellPatches = this.dirtyCellPatchOptions()
		const writeOptions: import('@ascend/io-xlsx').WriteXlsxOptions = {
			dirtySheetNames: [...this.dirtySheets],
			...(dirtyCellPatches ? { dirtyCellPatches } : {}),
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

	packageGraph(): XlsxPackageGraph {
		const bytes = this.originalBytes && !this.dirty ? this.originalBytes : this.toBytes()
		if (this.packageGraphCache?.bytes === bytes) return this.packageGraphCache.graph
		const graph = inspectXlsxPackageGraph(bytes)
		this.packageGraphCache = { bytes, graph }
		return graph
	}

	override trustReport(options: WorkbookTrustReportOptions = {}): WorkbookTrustReport {
		const info = this.inspect()
		const packageGraph =
			options.packageGraph ?? (info.sourceFormat === 'xlsx' ? this.packageGraph() : undefined)
		return super.trustReport({ ...options, ...(packageGraph ? { packageGraph } : {}) })
	}

	rawPackagePart(options: RawPackagePartOptions): RawPackagePartInfo {
		if (this.originalBytes && !this.dirty) {
			return {
				...inspectRawPackagePart(this.originalBytes, { ...options, origin: 'source' }),
				load: this.loadInfo,
			}
		}
		return {
			...inspectRawPackagePart(this.toBytes(), {
				...options,
				origin: 'serialized-current',
			}),
			load: this.loadInfo,
		}
	}

	private assertWritable(): void {
		if (!this.loadInfo.isPartial) return
		throw new AscendException(
			ascendError(
				'EXPORT_ERROR',
				'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
				{
					details: partialLoadErrorDetails(this.loadInfo),
					suggestedFix:
						'Open the workbook with mode "full" and all sheets loaded before saving or exporting.',
				},
			),
		)
	}

	private currentGenerations(): WorkbookGenerationInfo {
		return {
			workbook: this.workbookGeneration,
			sheetMetadata: this.sheetMetadataGeneration,
			formulas: this.formulaGeneration,
			styles: this.styleGeneration,
		}
	}

	protected override currentReadGenerations(): WorkbookGenerationInfo {
		return this.currentGenerations()
	}

	private advanceApplyGenerations(
		ops: readonly Operation[],
		dirtyFlags: ReturnType<AscendWorkbook['deriveDirtyFlags']>,
		applyResult: PatchResult,
	): void {
		this.workbookGeneration += 1
		if (dirtyFlags.workbookMetaDirty || ops.some((op) => operationChangesSheetMetadata(op))) {
			this.sheetMetadataGeneration += 1
		}
		if (dirtyFlags.calcStateDirty || dirtyFlags.calcChainDirty || applyResult.recalcRequired) {
			this.formulaGeneration += 1
		}
		if (dirtyFlags.stylesDirty) {
			this.styleGeneration += 1
		}
	}

	private deriveDirtyFlags(ops: readonly Operation[]): {
		workbookMetaDirty: boolean
		documentPropertiesDirty: boolean
		calcStateDirty: boolean
		calcChainDirty: boolean
		sharedStringsDirty: boolean
		stylesDirty: boolean
	} {
		let workbookMetaDirty = false
		let documentPropertiesDirty = false
		let calcStateDirty = false
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
				case 'setPivotCache':
				case 'setPivotFieldItem':
				case 'setConnectionRefresh':
				case 'setTimelineRange':
				case 'rewriteExternalLink':
					workbookMetaDirty = true
					calcStateDirty = true
					calcChainDirty = true
					break
				case 'setDocumentProperties':
					workbookMetaDirty = true
					documentPropertiesDirty = true
					calcStateDirty = true
					break
				case 'setDefinedName':
				case 'deleteDefinedName':
					workbookMetaDirty = true
					if (op.name !== '_xlnm.Print_Area') {
						calcStateDirty = true
						calcChainDirty = true
					}
					break
				case 'setPrintArea':
					workbookMetaDirty = true
					break
				case 'setWorkbookProperties':
				case 'setWorkbookView':
				case 'setCalcSettings':
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
					{
						const sheet = this.wb.getSheet(op.sheet)
						if (
							sheet &&
							op.updates.some((update) => {
								const ref = parseA1Safe(update.ref)
								if (!ref) return false
								const formula = sheet.cells.readFormula(ref.row, ref.col)
								return (
									(formula !== undefined && formula !== null) ||
									sheet.cells.readFormulaInfo(ref.row, ref.col) !== undefined
								)
							})
						) {
							calcChainDirty = true
						}
					}
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
			calcStateDirty,
			calcChainDirty,
			sharedStringsDirty,
			stylesDirty,
		}
	}

	private mergePendingDirtyCellRefs(ops: readonly Operation[], applyResult?: PatchResult): void {
		const refsBySheet = new Map<string, Set<string>>()
		for (const op of ops) {
			if (op.op !== 'setCells') {
				this.pendingDirtyCellRefs.clear()
				return
			}
			let refs = refsBySheet.get(op.sheet)
			if (!refs) {
				refs = new Set<string>()
				refsBySheet.set(op.sheet, refs)
			}
			for (const update of op.updates) {
				const ref = parseA1Safe(update.ref)
				if (!ref) {
					this.pendingDirtyCellRefs.clear()
					return
				}
				refs.add(`${indexToColumn(ref.col)}${ref.row + 1}`)
			}
		}
		if (applyResult) {
			for (const affected of applyResult.affectedCells) {
				if (affected.includes('!')) {
					const parsed = splitFullRef(affected, this.wb)
					if (!parsed) {
						this.pendingDirtyCellRefs.clear()
						return
					}
					let refs = refsBySheet.get(parsed.sheet)
					if (!refs) {
						refs = new Set<string>()
						refsBySheet.set(parsed.sheet, refs)
					}
					refs.add(`${indexToColumn(parsed.col)}${parsed.row + 1}`)
					continue
				}
				if (applyResult.sheetsModified.length !== 1) {
					this.pendingDirtyCellRefs.clear()
					return
				}
				const ref = parseA1Safe(affected)
				if (!ref) {
					this.pendingDirtyCellRefs.clear()
					return
				}
				const sheetName = applyResult.sheetsModified[0]
				if (!sheetName) {
					this.pendingDirtyCellRefs.clear()
					return
				}
				let refs = refsBySheet.get(sheetName)
				if (!refs) {
					refs = new Set<string>()
					refsBySheet.set(sheetName, refs)
				}
				refs.add(`${indexToColumn(ref.col)}${ref.row + 1}`)
			}
		}
		for (const [sheetName, refs] of refsBySheet) {
			let pending = this.pendingDirtyCellRefs.get(sheetName)
			if (!pending) {
				pending = new Set<string>()
				this.pendingDirtyCellRefs.set(sheetName, pending)
			}
			for (const ref of refs) pending.add(ref)
		}
	}

	private dirtyCellPatchOptions(): import('@ascend/io-xlsx').DirtyCellPatch[] | undefined {
		const patches: import('@ascend/io-xlsx').DirtyCellPatch[] = []
		for (const sheetName of this.dirtySheets) {
			const refs = this.pendingDirtyCellRefs.get(sheetName)
			if (!refs || refs.size === 0) continue
			patches.push({ sheetName, refs: [...refs] })
		}
		return patches.length > 0 ? patches : undefined
	}

	private markDirty(): void {
		if (!this.dirty) this.originalBytes = null
		this.packageGraphCache = undefined
		this.dirty = true
	}

	createMutationRollbackSnapshot(): WorkbookMutationRollbackSnapshot {
		return {
			workbook: this.wb.clone(),
			originalBytes: this.originalBytes ? new Uint8Array(this.originalBytes) : null,
			sourceArchive: this.sourceArchive,
			packageGraphCache: this.packageGraphCache
				? {
						bytes: new Uint8Array(this.packageGraphCache.bytes),
						graph: this.packageGraphCache.graph,
					}
				: undefined,
			dirty: this.dirty,
			dirtySheets: [...this.dirtySheets],
			pendingDirtyCellRefs: [...this.pendingDirtyCellRefs.entries()].map(
				([sheet, refs]) => [sheet, [...refs]] as const,
			),
			pendingDirtyRefs: [...this.pendingDirtyRefs],
			pendingFullRecalc: this.pendingFullRecalc,
			workbookMetaDirty: this.workbookMetaDirty,
			documentPropertiesDirty: this.documentPropertiesDirty,
			calcStateDirty: this.calcStateDirty,
			calcChainDirty: this.calcChainDirty,
			sharedStringsDirty: this.sharedStringsDirty,
			stylesDirty: this.stylesDirty,
			workbookGeneration: this.workbookGeneration,
			sheetMetadataGeneration: this.sheetMetadataGeneration,
			formulaGeneration: this.formulaGeneration,
			styleGeneration: this.styleGeneration,
		}
	}

	restoreMutationRollbackSnapshot(snapshot: WorkbookMutationRollbackSnapshot): void {
		this.wb = snapshot.workbook
		this.originalBytes = snapshot.originalBytes ? new Uint8Array(snapshot.originalBytes) : null
		this.sourceArchive = snapshot.sourceArchive
		this.packageGraphCache = snapshot.packageGraphCache
			? {
					bytes: new Uint8Array(snapshot.packageGraphCache.bytes),
					graph: snapshot.packageGraphCache.graph,
				}
			: undefined
		this.dirty = snapshot.dirty
		this.dirtySheets.clear()
		for (const sheet of snapshot.dirtySheets) this.dirtySheets.add(sheet)
		this.pendingDirtyCellRefs.clear()
		for (const [sheet, refs] of snapshot.pendingDirtyCellRefs) {
			this.pendingDirtyCellRefs.set(sheet, new Set(refs))
		}
		this.pendingDirtyRefs = [...snapshot.pendingDirtyRefs]
		this.pendingFullRecalc = snapshot.pendingFullRecalc
		this.workbookMetaDirty = snapshot.workbookMetaDirty
		this.documentPropertiesDirty = snapshot.documentPropertiesDirty
		this.calcStateDirty = snapshot.calcStateDirty
		this.calcChainDirty = snapshot.calcChainDirty
		this.sharedStringsDirty = snapshot.sharedStringsDirty
		this.stylesDirty = snapshot.stylesDirty
		this.workbookGeneration = snapshot.workbookGeneration
		this.sheetMetadataGeneration = snapshot.sheetMetadataGeneration
		this.formulaGeneration = snapshot.formulaGeneration
		this.styleGeneration = snapshot.styleGeneration
		this.clearReadCaches()
	}

	private captureSerializedState(bytes: Uint8Array): void {
		this.originalBytes = bytes
		this.wb.sourceArchiveBytes = bytes
		this.sourceArchive = undefined
		this.packageGraphCache = undefined
		this.dirty = false
		this.dirtySheets.clear()
		this.pendingDirtyCellRefs.clear()
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
		if (this.sourceArchive) return this.sourceArchive
		if (!cache) return extractZip(this.wb.sourceArchiveBytes)
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

function partialWorkbookEditError(
	loadInfo: WorkbookLoadInfo,
	options: { readonly message?: string } = {},
) {
	return ascendError(
		'VALIDATION_ERROR',
		options.message ??
			'Cannot modify a partial workbook view. Reopen the workbook with a full load before applying edits or recalculation.',
		{
			details: partialLoadErrorDetails(loadInfo),
			suggestedFix:
				'Open the workbook with mode "full" and all sheets loaded before editing or recalculating.',
		},
	)
}

function calcSettingsNeedClean(calcSettings: Workbook['calcSettings']): boolean {
	return (
		calcSettings.fullCalcOnLoad ||
		calcSettings.calcCompleted === false ||
		calcSettings.calcOnSave === false ||
		calcSettings.forceFullCalc === true
	)
}

function partialWorkbookPathMutationResult(
	loadInfo: WorkbookLoadInfo,
	mutations: readonly PathMutation[],
): PathMutationResult {
	const details = partialLoadErrorDetails(loadInfo)
	return {
		ops: [],
		mutationCount: mutations.length,
		issueCount: mutations.length,
		issues: mutations.map((mutation) => ({
			path: mutation.path,
			code: 'partial_workbook_view',
			message:
				'Cannot compile path mutations from a partial workbook view. Reopen with a full load before planning or committing edits.',
			details,
		})),
		replayable: false,
	}
}

function partialReplayBlockedInfo(
	loadInfo: WorkbookLoadInfo,
	message: string,
): import('./types.ts').PartialReplayBlockedInfo {
	return {
		code: 'partial_workbook_view',
		message,
		load: loadInfo,
	}
}

function partialLoadErrorDetails(loadInfo: WorkbookLoadInfo): Record<string, unknown> {
	return {
		partialWorkbookView: true,
		mode: loadInfo.mode,
		isPartial: loadInfo.isPartial,
		cellsHydrated: loadInfo.cellsHydrated,
		richSheetMetadataHydrated: loadInfo.richSheetMetadataHydrated,
		hasAllSheets: loadInfo.hasAllSheets,
		...(loadInfo.maxRows !== undefined ? { maxRows: loadInfo.maxRows } : {}),
		partialReasons: loadInfo.partialReasons,
		sourceSheets: loadInfo.sourceSheets,
		loadedSheets: loadInfo.loadedSheets,
		requiredLoad: {
			mode: 'full',
			allSheets: true,
			maxRows: null,
		},
	}
}

function operationChangesSheetMetadata(op: Operation): boolean {
	switch (op.op) {
		case 'addSheet':
		case 'deleteSheet':
		case 'renameSheet':
		case 'moveSheet':
		case 'copySheet':
		case 'hideSheet':
		case 'setTabColor':
		case 'setSheetProtection':
		case 'freezePane':
		case 'setColWidth':
		case 'setRowHeight':
		case 'hideRows':
		case 'hideCols':
		case 'createTable':
		case 'appendRows':
		case 'deleteTable':
		case 'renameTable':
		case 'resizeTable':
		case 'setTableColumn':
		case 'setTableStyle':
		case 'setComment':
		case 'setThreadedComment':
		case 'deleteComment':
		case 'setHyperlink':
		case 'deleteHyperlink':
		case 'setDataValidation':
		case 'deleteDataValidation':
		case 'setConditionalFormat':
		case 'deleteConditionalFormat':
		case 'setAutoFilter':
		case 'clearAutoFilter':
		case 'setAdvancedFilter':
		case 'setPrintArea':
		case 'setPageSetup':
		case 'mergeCells':
		case 'unmergeCells':
		case 'insertRows':
		case 'deleteRows':
		case 'insertCols':
		case 'deleteCols':
		case 'copyRange':
		case 'moveRange':
		case 'insertImage':
		case 'deleteImage':
		case 'replaceImage':
		case 'setDrawingText':
		case 'setSparklineGroup':
			return true
		default:
			return false
	}
}

function dirtyRegionsFromRefs(
	refs: readonly string[],
	fallbackSheets: readonly string[],
	workbook: Workbook,
): DirtyRegion[] {
	const bySheet = new Map<
		string,
		{ minRow: number; minCol: number; maxRow: number; maxCol: number; refs: string[] }
	>()
	for (const ref of refs) {
		const parsed = parseDirtyRef(ref, fallbackSheets, workbook)
		if (!parsed) continue
		const current = bySheet.get(parsed.sheet)
		const fullRef = `${parsed.sheet}!${parsed.ref}`
		if (current) {
			current.minRow = Math.min(current.minRow, parsed.start.row)
			current.minCol = Math.min(current.minCol, parsed.start.col)
			current.maxRow = Math.max(current.maxRow, parsed.end.row)
			current.maxCol = Math.max(current.maxCol, parsed.end.col)
			current.refs.push(fullRef)
		} else {
			bySheet.set(parsed.sheet, {
				minRow: parsed.start.row,
				minCol: parsed.start.col,
				maxRow: parsed.end.row,
				maxCol: parsed.end.col,
				refs: [fullRef],
			})
		}
	}
	return [...bySheet.entries()].map(([sheet, region]) => ({
		sheet,
		range: `${indexToColumn(region.minCol)}${region.minRow + 1}:${indexToColumn(region.maxCol)}${region.maxRow + 1}`,
		refs: region.refs,
	}))
}

function parseDirtyRef(
	ref: string,
	fallbackSheets: readonly string[],
	workbook: Workbook,
):
	| {
			sheet: string
			ref: string
			start: { row: number; col: number }
			end: { row: number; col: number }
	  }
	| undefined {
	const bang = ref.lastIndexOf('!')
	const sheet =
		bang === -1
			? fallbackSheets.length === 1
				? fallbackSheets[0]
				: workbook.sheets[0]?.name
			: ref.slice(0, bang).replace(/^'|'$/g, '')
	const body = bang === -1 ? ref : ref.slice(bang + 1)
	if (!sheet || !body) return undefined
	try {
		const range = parseRange(body)
		return {
			sheet,
			ref: body,
			start: range.start,
			end: range.end,
		}
	} catch {
		return undefined
	}
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
			case 'deleteTable':
			case 'renameTable':
			case 'resizeTable':
			case 'setTableColumn':
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
