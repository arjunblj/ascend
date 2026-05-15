import { createHash } from 'node:crypto'
import { readFileSync, statSync } from 'node:fs'
import { readFile, stat } from 'node:fs/promises'
import { extname, resolve } from 'node:path'
import {
	DEFAULT_STYLE_ID,
	parseA1Safe,
	parseRange,
	RangeIndex,
	type RangeIndexEntry,
	type RangeRef,
	rangeIntersection,
	type SheetConditionalFormat,
	type SheetDataValidation,
	sqrefIntersects,
	toA1,
	type Workbook,
} from '@ascend/core'
import { resolveCellFormulaText } from '@ascend/engine'
import {
	extractZip,
	inspectXlsxPackageGraph,
	summarizePlannedWrite,
	type XlsxPackageGraph,
} from '@ascend/io-xlsx'
import { AscendException, ascendError, type CellValue, type Operation } from '@ascend/schema'
import { check as verifyCheck, lint as verifyLint } from '@ascend/verify'
import {
	partialDependencyCheckIssue,
	partialDependencyLintWarning,
	sdkCheckIssueFromVerify,
} from './check-issues.ts'
import { formatStyledDisplayCellValue } from './format-helpers.ts'
import { emptyMutationJournal, unavailableMutationJournal } from './journal.ts'
import { type LoadedWorkbookSource, openWorkbookSource } from './load.ts'
import { inspectRawPackagePart } from './raw-package.ts'
import { WorkbookReadView } from './read-view.ts'
import type { CellSelector } from './ref-selectors.ts'
import type { SheetHandle } from './sheet-handle.ts'
import type { TableHandle } from './table-handle.ts'
import type {
	ActiveContentInfo,
	AgentViewOptions,
	AgentViewResult,
	CheckIssue,
	CheckResult,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	ExternalReferenceUsageInfo,
	FormulaInfo,
	LintResult,
	LintWarning,
	PivotCacheInfo,
	PivotCacheMaterializedRowInfo,
	PivotCacheRowsOptions,
	PivotOutputAuditInfo,
	PivotOutputMaterializeOpsResult,
	PivotOutputMaterializeOptions,
	PivotRefreshPlanInfo,
	PivotTableInfo,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	RawPackagePartInfo,
	RawPackagePartOptions,
	SheetInspectInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TableInfo,
	TimelineCacheInfo,
	TimelineInfo,
	TraceResult,
	WorkbookInfo,
	WorkbookLoadInfo,
	WorkbookRefreshMetadataInfo,
	WorkbookVisualInventoryInfo,
	WritePlanInfo,
} from './types.ts'
import { type ApplyOptions, AscendWorkbook, type WorkbookBytesOptions } from './workbook.ts'
import type { WorkbookTrustReport, WorkbookTrustReportOptions } from './workbook-trust.ts'

export interface WorkbookLoadOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
	readonly sheets?: readonly string[]
	readonly maxRows?: number
	readonly richMetadata?: boolean
	readonly formulaModeHydrateValues?: boolean
	readonly password?: string
	readonly pivotCacheRecordMaterializeLimit?: number | 'all'
}

export interface WorkbookFirstWindowOptions
	extends Omit<WorkbookLoadOptions, 'mode' | 'sheets' | 'maxRows'> {
	readonly sheet?: string
	readonly range: string
	readonly rowOffset?: number
	readonly rowLimit?: number
	readonly includeRefs?: boolean
	readonly omitEmpty?: boolean
	readonly flatValues?: boolean
}

export interface WorkbookFirstWindowResult {
	readonly document: WorkbookDocument
	readonly info: WorkbookInfo
	readonly sheet: string
	readonly window: CompactRangeWindowInfo
	readonly load: WorkbookLoadInfo
}

export interface WorkbookSessionFirstWindowResult extends WorkbookFirstWindowResult {
	readonly session: WorkbookSession
}

export interface AscendSessionOpenOptions extends Omit<WorkbookLoadOptions, 'mode'> {
	readonly mode?: WorkbookLoadOptions['mode'] | 'interactive'
	readonly prepareEdits?: boolean
}

export interface InteractiveViewportRequest {
	readonly sheet: string
	readonly topRow: number
	readonly leftCol: number
	readonly rowCount: number
	readonly colCount: number
	readonly overscanRows?: number
	readonly overscanCols?: number
	readonly changedSince?: string
}

export interface InteractiveViewportCell {
	readonly row: number
	readonly col: number
	readonly ref: string
	readonly value: CellValue
	readonly flatValue: number | string | boolean | null
	readonly displayText: string
	readonly formula: string | null
	readonly formulaBinding: import('@ascend/core').CellFormulaBinding | null
	readonly styleId: number
	readonly flags: {
		readonly formula: boolean
		readonly comment: boolean
		readonly hyperlink: boolean
		readonly merged: boolean
		readonly validation: boolean
		readonly conditionalFormat: boolean
		readonly table: boolean
	}
}

export interface InteractiveViewportLayoutEntry {
	readonly index: number
	readonly size?: number
	readonly hidden?: boolean
	readonly outlineLevel?: number
	readonly collapsed?: boolean
}

export interface InteractiveViewportPatch {
	readonly baseToken: string
	readonly changeToken: string
	readonly changedCells: readonly InteractiveViewportCell[]
	readonly removedRefs: readonly string[]
	readonly byteLength: number
}

export interface InteractiveViewportPatchResult {
	readonly patch: InteractiveViewportPatch | null
	readonly patchInvalidation?: InteractiveViewportPatchInvalidation
}

export interface InteractiveUiRangeRequest {
	readonly sheet: string
	readonly range: string
}

export interface InteractiveUiRangeResult {
	readonly sheet: string
	readonly requested: RangeRef
	readonly range: RangeRef
	readonly generation: {
		readonly session: number
		readonly workbook: number
		readonly sheetMetadata: number
		readonly formulas: number
		readonly styles: number
	}
	readonly load: WorkbookLoadInfo
	readonly rowCount: number
	readonly colCount: number
	readonly cells: readonly InteractiveViewportCell[]
	readonly flatValues: readonly (number | string | boolean | null)[]
	readonly displayText: readonly string[]
}

export type InteractiveViewportPatchInvalidationReason =
	| 'base-snapshot-missing'
	| 'base-token-stale'
	| 'base-token-invalid'
	| 'base-token-expired'
	| 'viewport-invalidated'

export interface InteractiveViewportPatchInvalidation {
	readonly baseToken: string
	readonly changeToken: string
	readonly reason: InteractiveViewportPatchInvalidationReason
	readonly requiredAction: 'use-returned-snapshot'
}

export interface InteractiveViewportResult {
	readonly sheet: string
	readonly requested: RangeRef
	readonly viewport: RangeRef
	readonly generation: {
		readonly session: number
		readonly workbook: number
		readonly sheetMetadata: number
		readonly formulas: number
		readonly styles: number
	}
	readonly changeToken: string
	readonly load: WorkbookLoadInfo
	readonly rowCount: number
	readonly colCount: number
	readonly cells: readonly InteractiveViewportCell[]
	readonly flatValues: readonly (number | string | boolean | null)[]
	readonly displayText: readonly string[]
	readonly rowLayout: readonly InteractiveViewportLayoutEntry[]
	readonly colLayout: readonly InteractiveViewportLayoutEntry[]
	readonly frozen: { readonly rows: number; readonly cols: number }
	readonly merges: readonly RangeRef[]
	readonly comments: readonly import('./types.ts').SheetCommentInfo[]
	readonly hyperlinks: readonly import('./types.ts').SheetHyperlinkInfo[]
	readonly dataValidations: readonly import('@ascend/core').SheetDataValidation[]
	readonly conditionalFormats: readonly import('@ascend/core').SheetConditionalFormat[]
	readonly tables: readonly TableInfo[]
	readonly autoFilter: import('@ascend/core').AutoFilter | null
	readonly patch?: InteractiveViewportPatch
	readonly patchInvalidation?: InteractiveViewportPatchInvalidation
}

export interface AscendSessionApplyOptions extends Pick<ApplyOptions, 'journal'> {
	readonly recalc?: boolean
}

export interface AscendSessionApplyTimings {
	readonly inspectReadMs: number
	readonly ensureMutableWorkbookMs: number
	readonly mutableWorkbookCached?: boolean
	readonly mutableWorkbookReusedReadModel?: boolean
	readonly mutableWorkbookOpenMs?: number
	readonly rebaseViewportSnapshotsMs?: number
	readonly applyMs: number
	readonly recalcMs: number
	readonly generationSnapshotMs: number
	readonly inspectWriteMs: number
	readonly totalMs: number
}

export interface AscendSessionApplyResult {
	readonly apply: import('./types.ts').ApplyResult
	readonly recalc: import('./types.ts').RecalcResult | null
	readonly load: {
		readonly read: WorkbookLoadInfo
		readonly write: WorkbookLoadInfo
		readonly promotedToFull: boolean
	}
	readonly generation: {
		readonly session: number
		readonly workbook: number
		readonly sheetMetadata: number
		readonly formulas: number
		readonly styles: number
	}
	readonly timings: AscendSessionApplyTimings
}

export interface AscendSessionPrepareEditsResult {
	readonly load: {
		readonly read: WorkbookLoadInfo
		readonly write: WorkbookLoadInfo
		readonly promotedToFull: boolean
	}
	readonly timings: {
		readonly ensureMutableWorkbookMs: number
		readonly mutableWorkbookCached?: boolean
		readonly mutableWorkbookReusedReadModel?: boolean
		readonly mutableWorkbookOpenMs?: number
		readonly rebaseViewportSnapshotsMs?: number
		readonly inspectWriteMs: number
		readonly totalMs: number
	}
}

export interface AscendSessionEditReadiness {
	readonly ready: boolean
	readonly preparing: boolean
	readonly generation: number
	readonly read: WorkbookLoadInfo
	readonly write: WorkbookLoadInfo | null
	readonly promotedToFull: boolean
	readonly timings: {
		readonly mutableWorkbookCached: boolean
		readonly mutableWorkbookReusedReadModel: boolean
		readonly mutableWorkbookOpenMs: number
		readonly rebaseViewportSnapshotsMs: number
	} | null
}

interface SessionFileIdentity {
	readonly path: string
	readonly size: number
	readonly mtimeMs: number
	readonly ctimeMs: number
	readonly sha256: string
}

interface SessionBytesIdentity {
	readonly key: string
	readonly size: number
}

type SessionIdentity = SessionFileIdentity | SessionBytesIdentity

interface SessionCacheEntry {
	readonly key: string
	readonly identity: SessionIdentity
	readonly document: WorkbookDocument
	readonly sizeBytes: number
	accessedAt: number
}

interface PathSnapshot {
	readonly identity: SessionFileIdentity
	readonly bytes: Uint8Array
	readonly sourceExtension?: string
}

interface InteractiveChangeEntry {
	readonly generation: number
	readonly refs: ReadonlySet<string> | null
}

interface InteractiveViewportSnapshot {
	readonly token: string
	readonly request: InteractiveViewportRequest
	readonly cells: Map<string, InteractiveViewportCell>
}

interface InteractiveViewportSnapshotLedger {
	readonly snapshots: InteractiveViewportSnapshot[]
}

interface MutableWorkbookEnsureTimings {
	readonly cached: boolean
	readonly reusedReadModel: boolean
	readonly openMs: number
	readonly rebaseViewportSnapshotsMs: number
}

type InteractiveSheetModel = ReturnType<WorkbookReadView['getWorkbookModel']>['sheets'][number]

interface InteractiveViewportOverlayIndexes {
	readonly sheet: InteractiveSheetModel
	readonly sheetMetadataGeneration: number
	readonly mergeIndex: RangeIndex<RangeRef>
	readonly validationIndex: RangeIndex<SheetDataValidation>
	readonly conditionalFormatIndex: RangeIndex<SheetConditionalFormat>
	readonly tableIndex: RangeIndex<TableInfo>
}

const interactiveViewportOverlayIndexCache = new WeakMap<
	WorkbookReadView,
	Map<string, InteractiveViewportOverlayIndexes>
>()

export interface SessionCacheOptions {
	readonly maxCacheSize?: number
	readonly maxCacheAge?: number
	readonly maxCacheBytes?: number
}

const cacheConfig = {
	maxCacheSize: 50,
	maxCacheAge: 5 * 60 * 1000,
	maxCacheBytes: 64 * 1024 * 1024,
}

const DEFAULT_FIRST_WINDOW_ROWS = 500
const INTERACTIVE_VIEWPORT_SNAPSHOT_RETENTION = 8

const sessionCache = new Map<string, SessionCacheEntry>()

export function configureSessionCache(opts: SessionCacheOptions): void {
	if (opts.maxCacheSize !== undefined) cacheConfig.maxCacheSize = opts.maxCacheSize
	if (opts.maxCacheAge !== undefined) cacheConfig.maxCacheAge = opts.maxCacheAge
	if (opts.maxCacheBytes !== undefined) cacheConfig.maxCacheBytes = opts.maxCacheBytes
}

function isEntryExpired(entry: SessionCacheEntry): boolean {
	return Date.now() - entry.accessedAt > cacheConfig.maxCacheAge
}

/**
 * Read-only document view of a workbook. Use for inspect, read, check, lint, trace,
 * agent-view, and other operations that do not modify the file. Supports session
 * caching for repeated opens of the same file. Use `AscendWorkbook` when you need
 * to apply operations, recalculate, or save.
 */
export class WorkbookDocument {
	private readonly cacheKey: string
	private readonly identity: SessionIdentity
	private readonly source: string | Uint8Array
	private readonly options: WorkbookLoadOptions
	private readonly view: WorkbookReadView
	private readonly capsules: LoadedWorkbookSource['capsules']
	private readonly originalBytes: Uint8Array | null

	private constructor(
		cacheKey: string,
		source: string | Uint8Array,
		identity: SessionIdentity,
		options: WorkbookLoadOptions,
		view: WorkbookReadView,
		capsules: LoadedWorkbookSource['capsules'],
		originalBytes: Uint8Array | null,
	) {
		this.cacheKey = cacheKey
		this.source = source
		this.identity = identity
		this.options = options
		this.view = view
		this.capsules = capsules
		this.originalBytes = originalBytes
	}

	static async open(
		source: string | Uint8Array,
		options: WorkbookLoadOptions = {},
	): Promise<WorkbookDocument> {
		if (typeof source === 'string') {
			const cachedPathDocument = await readCachedPathDocument(source, options)
			if (cachedPathDocument) return cachedPathDocument
		}
		const hashPathSnapshot = shouldHashPathSnapshot(options)
		const pathSnapshot =
			typeof source === 'string' && hashPathSnapshot
				? await readPathSnapshot(source, true)
				: undefined
		const pathIdentity =
			typeof source === 'string' && !pathSnapshot ? await readIdentity(source, false) : undefined
		const identity =
			pathSnapshot?.identity ?? pathIdentity ?? readBytesIdentity(source as Uint8Array)
		const key = makeSessionKey(identity, options)
		const cached = sessionCache.get(key)
		if (cached && isIdentityEqual(cached.identity, identity)) {
			if (isEntryExpired(cached)) {
				sessionCache.delete(key)
			} else {
				touchCacheEntry(cached)
				return cached.document
			}
		}

		const loaded = await openWorkbookSource(
			pathSnapshot?.bytes ?? source,
			pathSnapshot?.sourceExtension
				? { ...options, sourceExtension: pathSnapshot.sourceExtension }
				: options,
		)
		if (pathIdentity) {
			const after = await readIdentity(pathIdentity.path, false)
			if (!isIdentityEqual(pathIdentity, after)) {
				throw new AscendException(unstablePathDocumentError(pathIdentity.path, null))
			}
		}
		const document = new WorkbookDocument(
			key,
			pathSnapshot?.identity.path ?? source,
			identity,
			normalizeOptions(options),
			new WorkbookReadView(loaded.workbook, loaded.report, loaded.loadInfo),
			loaded.capsules,
			loaded.originalBytes,
		)
		document.refreshCacheFootprint('base')
		return document
	}

	static async openPathSnapshot(
		file: string,
		bytes: Uint8Array,
		identity: SessionFileIdentity,
		options: WorkbookLoadOptions = {},
	): Promise<WorkbookDocument> {
		const key = makeSessionKey(identity, options)
		sessionCache.delete(key)

		const sourceExtension = extname(file).replace(/^\./, '').toLowerCase()
		const loaded = await openWorkbookSource(bytes, {
			...options,
			...(sourceExtension ? { sourceExtension } : {}),
		})
		const document = new WorkbookDocument(
			key,
			identity.path,
			identity,
			normalizeOptions(options),
			new WorkbookReadView(loaded.workbook, loaded.report, loaded.loadInfo),
			loaded.capsules,
			loaded.originalBytes,
		)
		document.refreshCacheFootprint('base')
		return document
	}

	static async openFirstWindow(
		source: string | Uint8Array,
		options: WorkbookFirstWindowOptions,
	): Promise<WorkbookFirstWindowResult> {
		const document = await WorkbookDocument.open(source, firstWindowLoadOptions(options))
		return readFirstWindow(document, options)
	}

	static clearCache(): void {
		sessionCache.clear()
	}

	static drop(file: string, options: WorkbookLoadOptions = {}): void {
		const key = makeSessionKey(
			{ path: resolve(file), size: 0, mtimeMs: 0, ctimeMs: 0, sha256: '' },
			options,
		)
		sessionCache.delete(key)
	}

	get file(): string {
		return 'path' in this.identity ? this.identity.path : this.identity.key
	}

	get sheets(): readonly string[] {
		return this.view.sheets
	}

	get report() {
		return this.view.report
	}

	get loadOptions(): WorkbookLoadOptions {
		return this.options
	}

	async withLoad(options: WorkbookLoadOptions): Promise<WorkbookDocument> {
		const nextOptions = mergeOpenOptions(this.options, options)
		if (sameOpenOptions(this.options, nextOptions)) return this
		return WorkbookDocument.open(this.source, nextOptions)
	}

	async withSheet(
		sheetName: string,
		options?: { mode?: 'values' | 'formula' | 'full' },
	): Promise<WorkbookDocument> {
		return this.withLoad({
			...(options?.mode ? { mode: options.mode } : {}),
			sheets: [sheetName],
		})
	}

	async withSheets(
		sheetNames: readonly string[],
		options?: { mode?: 'values' | 'formula' | 'full' },
	): Promise<WorkbookDocument> {
		return this.withLoad({
			...(options?.mode ? { mode: options.mode } : {}),
			sheets: sheetNames,
		})
	}

	toMutableWorkbook(): AscendWorkbook | null {
		const load = this.inspect().load
		if (load.isPartial) return null
		return AscendWorkbook.fromLoadedSource({
			workbook: this.view.getWorkbookModel().clone(),
			capsules: this.capsules,
			report: this.report,
			loadInfo: load,
			originalBytes: this.originalBytes,
		})
	}

	inspect(): WorkbookInfo {
		return this.view.inspect()
	}

	async trustReport(options: WorkbookTrustReportOptions = {}): Promise<WorkbookTrustReport> {
		const info = this.inspect()
		const packageGraph =
			options.packageGraph ?? (info.sourceFormat === 'xlsx' ? await this.packageGraph() : undefined)
		return this.view.trustReport({ ...options, ...(packageGraph ? { packageGraph } : {}) })
	}

	visualInventory(): WorkbookVisualInventoryInfo {
		return this.view.visualInventory()
	}

	getWorkbookModel(): Workbook {
		return this.view.getWorkbookModel()
	}

	writePlanSummary(): WritePlanInfo {
		const workbook = this.view.getWorkbookModel()
		const sourceArchiveBytes = workbook.sourceArchiveBytes ?? this.originalBytes
		const result = summarizePlannedWrite(
			workbook,
			this.capsules.length > 0 ? [...this.capsules] : undefined,
			sourceArchiveBytes ? { sourceArchive: extractZip(sourceArchiveBytes) } : {},
		)
		if (!result.ok) throw new AscendException(result.error)
		return result.value
	}

	async packageGraph(): Promise<XlsxPackageGraph> {
		return inspectXlsxPackageGraph(await this.readSourceBytes())
	}

	async rawPackagePart(options: RawPackagePartOptions): Promise<RawPackagePartInfo> {
		return {
			...inspectRawPackagePart(await this.readSourceBytes(), { ...options, origin: 'source' }),
			load: this.inspect().load,
		}
	}

	inspectSheet(name: string): SheetInspectInfo | undefined {
		return this.view.inspectSheet(name)
	}

	sheet(name: string): SheetHandle | undefined {
		return this.view.sheet(name)
	}

	readRange(sheetName: string, range: string): RangeInfo | undefined {
		return this.view.readRange(sheetName, range)
	}

	readRangeCompact(
		sheetName: string,
		range: string,
		opts?: { includeRefs?: boolean; omitEmpty?: boolean; flatValues?: boolean },
	): CompactRangeInfo | undefined {
		return this.view.readRangeCompact(sheetName, range, opts)
	}

	readWindow(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeWindowInfo | undefined {
		return this.view.readWindow(sheetName, range, opts)
	}

	readWindowCompact(
		sheetName: string,
		range: string,
		opts?: import('./types.ts').AgentReadOptions,
	): CompactRangeWindowInfo | undefined {
		return this.view.readWindowCompact(sheetName, range, opts)
	}

	readViewport(request: InteractiveViewportRequest): InteractiveViewportResult {
		return readInteractiveViewport(this.view, request, 0, '0')
	}

	readUiRange(request: InteractiveUiRangeRequest): InteractiveUiRangeResult {
		return readInteractiveUiRange(this.view, request, 0)
	}

	readSnapshotInfo(): import('./types.ts').WorkbookReadSnapshotInfo {
		return this.view.readSnapshotInfo()
	}

	agentView(
		sheetName: string,
		range: string,
		opts?: AgentViewOptions,
	): AgentViewResult | undefined {
		return this.view.agentView(sheetName, range, opts)
	}

	private async readSourceBytes(): Promise<Uint8Array> {
		if (this.source instanceof Uint8Array) return this.source
		return typeof Bun !== 'undefined'
			? Bun.file(this.source).bytes()
			: new Uint8Array(await readFile(this.source))
	}

	readRows(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeRowsInfo | undefined {
		return this.view.readRows(sheetName, range, opts)
	}

	readObjects(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number; headers?: readonly string[] | 'first-row' },
	): RangeObjectsInfo | undefined {
		return this.view.readObjects(sheetName, range, opts)
	}

	*streamRange(
		sheetName: string,
		range: string,
	): Generator<readonly import('./types.ts').CellInfo[]> {
		yield* this.view.streamRange(sheetName, range)
	}

	*streamRangeCompact(
		sheetName: string,
		range: string,
		opts?: { includeRefs?: boolean },
	): Generator<readonly import('./types.ts').CompactCellInfo[]> {
		yield* this.view.streamRangeCompact(sheetName, range, opts)
	}

	*streamWindows(
		sheetName: string,
		range: string,
		opts?: { rowLimit?: number },
	): Generator<RangeWindowInfo> {
		yield* this.view.streamWindows(sheetName, range, opts)
	}

	*streamWindowsCompact(
		sheetName: string,
		range: string,
		opts?: { rowLimit?: number; includeRefs?: boolean },
	): Generator<CompactRangeWindowInfo> {
		yield* this.view.streamWindowsCompact(sheetName, range, opts)
	}

	trace(cellRef: CellSelector, opts?: { maxDepth?: number }): TraceResult | undefined {
		this.refreshCacheFootprint('verify')
		return this.view.trace(cellRef, opts)
	}

	traceIssue(cellRef: CellSelector): CheckIssue | undefined {
		this.refreshCacheFootprint('verify')
		return this.view.traceIssue(cellRef)
	}

	formula(cellRef: CellSelector): FormulaInfo | undefined {
		return this.view.formula(cellRef)
	}

	check(): CheckResult {
		const issue = this.view.dependencyVerificationIssue()
		if (issue) {
			return {
				valid: false,
				issues: [partialDependencyCheckIssue(issue)],
			}
		}
		const workbook = this.view.getWorkbookModel()
		const result = verifyCheck(workbook, {
			formulas: this.view.formulaAnalysis(),
			dependencies: this.view.dependencyAnalysis(),
			...(workbook.sourceArchiveBytes
				? { packageGraph: inspectXlsxPackageGraph(workbook.sourceArchiveBytes) }
				: {}),
		})
		this.refreshCacheFootprint('verify')
		const issues = result.issues.map(sdkCheckIssueFromVerify)
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const issue = this.view.dependencyVerificationIssue()
		const result = verifyLint(this.view.getWorkbookModel(), this.view.formulaAnalysis())
		this.refreshCacheFootprint('verify')
		const warnings: LintWarning[] = result.violations.map((violation) => ({
			rule: violation.rule,
			severity: violation.severity,
			message: violation.message,
			ref: violation.ref,
		}))
		if (issue) warnings.unshift(partialDependencyLintWarning(issue))
		return { clean: warnings.length === 0, warnings }
	}

	definedName(name: string, sheetName?: string): DefinedNameInfo | undefined {
		return this.view.definedName(name, sheetName)
	}

	definedNames(sheetName?: string): readonly DefinedNameInfo[] {
		return this.view.definedNames(sheetName)
	}

	table(name: string): TableHandle | undefined {
		return this.view.table(name)
	}

	pivotTables(sheetName?: string): readonly PivotTableInfo[] {
		return this.view.pivotTables(sheetName)
	}

	pivotCaches(): readonly PivotCacheInfo[] {
		return this.view.pivotCaches()
	}

	pivotCacheRows(options?: PivotCacheRowsOptions): readonly PivotCacheMaterializedRowInfo[] {
		return this.view.pivotCacheRows(options)
	}

	pivotOutputAudits(): readonly PivotOutputAuditInfo[] {
		return this.view.pivotOutputAudits()
	}

	pivotOutputMaterializeOps(
		options: PivotOutputMaterializeOptions = {},
	): PivotOutputMaterializeOpsResult {
		return this.view.pivotOutputMaterializeOps(options)
	}

	pivotRefreshPlans(): readonly PivotRefreshPlanInfo[] {
		return this.view.pivotRefreshPlans()
	}

	refreshMetadata(): WorkbookRefreshMetadataInfo {
		return this.view.refreshMetadata()
	}

	slicerCaches(): readonly SlicerCacheInfo[] {
		return this.view.slicerCaches()
	}

	slicers(): readonly SlicerInfo[] {
		return this.view.slicers()
	}

	timelineCaches(): readonly TimelineCacheInfo[] {
		return this.view.timelineCaches()
	}

	timelines(): readonly TimelineInfo[] {
		return this.view.timelines()
	}

	activeContent(): readonly ActiveContentInfo[] {
		return this.view.activeContent()
	}

	workbookViews(): readonly import('./types.ts').WorkbookViewInfo[] {
		return this.view.workbookViews()
	}

	externalReferenceUsages(): readonly ExternalReferenceUsageInfo[] {
		return this.view.externalReferenceUsages()
	}

	externalReferences(): readonly string[] {
		return this.view.externalReferences()
	}

	private refreshCacheFootprint(usage: 'base' | 'verify'): void {
		replaceCacheEntry(this.cacheKey, {
			key: this.cacheKey,
			identity: this.identity,
			document: this,
			sizeBytes: sessionSizeBytes(this.identity, this.view, usage),
			accessedAt: Date.now(),
		})
	}
}

export interface WorkbookSessionOpenOptions extends WorkbookLoadOptions {}

/**
 * Session that holds a loaded workbook and reuses parsed state (analysis, dependency graph)
 * across repeated inspect, read, and trace calls. Faster than opening independently for
 * each operation. Use for read-only workflows that perform multiple operations on the same file.
 */
export class WorkbookSession {
	private document: WorkbookDocument
	private source: { path?: string; bytes?: Uint8Array } | null
	private readonly options: WorkbookLoadOptions
	private fileIdentity: SessionFileIdentity | null
	private closed = false

	private constructor(
		document: WorkbookDocument,
		source: { path?: string; bytes?: Uint8Array },
		options: WorkbookLoadOptions,
		fileIdentity: SessionFileIdentity | null,
	) {
		this.document = document
		this.source = source
		this.options = options
		this.fileIdentity = fileIdentity
	}

	static async open(
		pathOrBytes: string | Uint8Array,
		options: WorkbookSessionOpenOptions = {},
	): Promise<WorkbookSession> {
		const opts = normalizeOptions(options)
		const source =
			typeof pathOrBytes === 'string' ? { path: resolve(pathOrBytes) } : { bytes: pathOrBytes }
		const opened =
			source.path !== undefined
				? await openStablePathDocument(source.path, opts)
				: { document: await WorkbookDocument.open(pathOrBytes, opts), identity: null }
		const document = opened.document
		const fileIdentity = opened.identity
		return new WorkbookSession(document, source, opts, fileIdentity)
	}

	static async openFirstWindow(
		pathOrBytes: string | Uint8Array,
		options: WorkbookFirstWindowOptions,
	): Promise<WorkbookSessionFirstWindowResult> {
		const session = await WorkbookSession.open(pathOrBytes, firstWindowLoadOptions(options))
		const result = readFirstWindow(session.workbook(), options)
		return { ...result, session }
	}

	inspect(): WorkbookInfo {
		this.assertOpen()
		return this.document.inspect()
	}

	read(
		range: string,
		opts?: { includeRefs?: boolean; omitEmpty?: boolean; flatValues?: boolean },
	): RangeInfo | CompactRangeInfo | undefined {
		this.assertOpen()
		const defaultSheet = this.document.sheets[0] ?? 'Sheet1'
		const { sheetName, ref } = parseRangeRef(range, defaultSheet)
		const sheet = this.document.sheet(sheetName)
		if (!sheet) return undefined
		if (opts && (opts.flatValues || opts.omitEmpty || opts.includeRefs)) {
			return this.document.readRangeCompact(sheetName, ref, opts)
		}
		return this.document.readRange(sheetName, ref)
	}

	trace(ref: CellSelector, opts?: { maxDepth?: number }): TraceResult | undefined {
		this.assertOpen()
		return this.document.trace(ref, opts)
	}

	traceIssue(ref: CellSelector): CheckIssue | undefined {
		this.assertOpen()
		return this.document.traceIssue(ref)
	}

	pivotCacheRows(options?: PivotCacheRowsOptions): readonly PivotCacheMaterializedRowInfo[] {
		this.assertOpen()
		return this.document.pivotCacheRows(options)
	}

	pivotOutputAudits(): readonly PivotOutputAuditInfo[] {
		this.assertOpen()
		return this.document.pivotOutputAudits()
	}

	pivotOutputMaterializeOps(
		options: PivotOutputMaterializeOptions = {},
	): PivotOutputMaterializeOpsResult {
		this.assertOpen()
		return this.document.pivotOutputMaterializeOps(options)
	}

	isStale(): boolean {
		if (this.closed) return false
		if (this.source?.bytes || !this.fileIdentity) return false
		try {
			const info = statSync(this.fileIdentity.path)
			if (
				info.size !== this.fileIdentity.size ||
				info.mtimeMs !== this.fileIdentity.mtimeMs ||
				info.ctimeMs !== this.fileIdentity.ctimeMs
			) {
				return true
			}
			if (!this.fileIdentity.sha256) return false
			const sha256 = createHash('sha256')
				.update(new Uint8Array(readFileSync(this.fileIdentity.path)))
				.digest('hex')
			return sha256 !== this.fileIdentity.sha256
		} catch {
			return true
		}
	}

	async refresh(): Promise<void> {
		this.assertOpen()
		const path = this.source?.path
		if (!path) return
		WorkbookDocument.drop(path, this.options)
		const opened = await openStablePathDocument(path, this.options)
		this.document = opened.document
		this.fileIdentity = opened.identity
	}

	workbook(): WorkbookDocument {
		this.assertOpen()
		return this.document
	}

	close(): void {
		this.closed = true
		this.document = null as unknown as WorkbookDocument
		this.source = null
		this.fileIdentity = null
	}

	private assertOpen(): void {
		if (this.closed) throw new Error('WorkbookSession is closed')
	}
}

/**
 * Experimental interactive session facade for grid-like clients. It keeps a
 * retained WorkbookSession and returns viewport-shaped payloads with generation
 * and change tokens so callers do not need to stitch together range reads and
 * broad sheet metadata scans.
 */
export class AscendSession {
	private readonly session: WorkbookSession
	private source: string | Uint8Array | null
	private readonly options: AscendSessionOpenOptions
	private mutableWorkbook: AscendWorkbook | null = null
	private mutableWorkbookPromise: Promise<AscendWorkbook> | null = null
	private mutableWorkbookReadyTimings: MutableWorkbookEnsureTimings | null = null
	private documentGeneration = 0
	private changeVersion = 0
	private readonly recentChanges: InteractiveChangeEntry[] = []
	private readonly viewportSnapshots = new Map<string, InteractiveViewportSnapshotLedger>()

	private constructor(
		session: WorkbookSession,
		source: string | Uint8Array,
		options: AscendSessionOpenOptions,
	) {
		this.session = session
		this.source = source
		this.options = options
	}

	static async open(
		pathOrBytes: string | Uint8Array,
		options: AscendSessionOpenOptions = {},
	): Promise<AscendSession> {
		const loadOptions = stripPrepareEditsOption(options)
		const session = await WorkbookSession.open(pathOrBytes, interactiveInitialOpenOptions(options))
		const ascendSession = new AscendSession(session, pathOrBytes, loadOptions)
		if (options.prepareEdits) await ascendSession.prepareEdits()
		return ascendSession
	}

	inspect(): WorkbookInfo {
		return this.session.inspect()
	}

	readViewport(request: InteractiveViewportRequest): InteractiveViewportResult {
		const changeToken = `${this.documentGeneration}:${this.changeVersion++}`
		const base = this.mutableWorkbook
			? readInteractiveViewport(this.mutableWorkbook, request, this.documentGeneration, changeToken)
			: this.session.workbook().readViewport(request)
		const result = {
			...base,
			generation: { ...base.generation, session: this.documentGeneration },
			changeToken,
		}
		const snapshotKey = interactiveViewportSnapshotKey(request)
		const currentCells = interactiveCellMap(result.cells)
		const ledger = this.viewportSnapshots.get(snapshotKey)
		if (!request.changedSince) {
			this.rememberViewportSnapshot(snapshotKey, {
				token: changeToken,
				request,
				cells: currentCells,
			})
			return result
		}
		const baseGeneration = interactiveTokenGeneration(request.changedSince)
		if (baseGeneration === null) {
			this.rememberViewportSnapshot(snapshotKey, {
				token: changeToken,
				request,
				cells: currentCells,
			})
			return {
				...result,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					changeToken,
					'base-token-invalid',
				),
			}
		}
		if (!ledger) {
			this.rememberViewportSnapshot(snapshotKey, {
				token: changeToken,
				request,
				cells: currentCells,
			})
			return {
				...result,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					changeToken,
					'base-snapshot-missing',
				),
			}
		}
		const previous = findInteractiveViewportSnapshot(ledger, request.changedSince)
		if (!previous) {
			this.rememberViewportSnapshot(snapshotKey, {
				token: changeToken,
				request,
				cells: currentCells,
			})
			return {
				...result,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					changeToken,
					interactiveViewportSnapshotInvalidationReason(request.changedSince, ledger),
				),
			}
		}
		const changes = this.changedRefsSince(baseGeneration)
		if (changes.kind !== 'refs') {
			this.rememberViewportSnapshot(snapshotKey, {
				token: changeToken,
				request,
				cells: currentCells,
			})
			return {
				...result,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					changeToken,
					changes.kind === 'expired' ? 'base-token-expired' : 'viewport-invalidated',
				),
			}
		}
		const patch = diffInteractiveViewportCells(
			previous.token,
			changeToken,
			previous.cells,
			currentCells,
		)
		this.rememberViewportSnapshot(snapshotKey, { token: changeToken, request, cells: currentCells })
		return { ...result, patch }
	}

	readViewportPatch(request: InteractiveViewportRequest): InteractiveViewportPatch | null {
		return this.readViewportPatchResult(request).patch
	}

	readViewportPatchResult(request: InteractiveViewportRequest): InteractiveViewportPatchResult {
		if (!request.changedSince) return { patch: null }
		const snapshotKey = interactiveViewportSnapshotKey(request)
		const ledger = this.viewportSnapshots.get(snapshotKey)
		const baseGeneration = interactiveTokenGeneration(request.changedSince)
		if (baseGeneration === null) {
			return {
				patch: null,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					`${this.documentGeneration}:${this.changeVersion++}`,
					'base-token-invalid',
				),
			}
		}
		if (!ledger) {
			return {
				patch: null,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					`${this.documentGeneration}:${this.changeVersion++}`,
					'base-snapshot-missing',
				),
			}
		}
		const previous = findInteractiveViewportSnapshot(ledger, request.changedSince)
		if (!previous) {
			return {
				patch: null,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					`${this.documentGeneration}:${this.changeVersion++}`,
					interactiveViewportSnapshotInvalidationReason(request.changedSince, ledger),
				),
			}
		}
		const changes = this.changedRefsSince(baseGeneration)
		if (changes.kind !== 'refs') {
			return {
				patch: null,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					`${this.documentGeneration}:${this.changeVersion++}`,
					changes.kind === 'expired' ? 'base-token-expired' : 'viewport-invalidated',
				),
			}
		}
		const refs = changes.refs
		const changeToken = `${this.documentGeneration}:${this.changeVersion++}`
		if (!this.mutableWorkbook && refs.size > 0) {
			return {
				patch: null,
				patchInvalidation: interactivePatchInvalidation(
					request.changedSince,
					changeToken,
					'viewport-invalidated',
				),
			}
		}
		const changedCells: InteractiveViewportCell[] = []
		const removedRefs: string[] = []
		const nextCells = new Map(previous.cells)
		if (this.mutableWorkbook) {
			const context = createInteractiveViewportContext(this.mutableWorkbook, request)
			for (const fullRef of refs) {
				const parsed = splitInteractiveFullRef(fullRef, context.workbook)
				if (!parsed || parsed.sheet !== request.sheet) continue
				if (!cellInRange(parsed.row, parsed.col, context.viewport)) continue
				const cell = context.sheet.cells.get(parsed.row, parsed.col)
				if (!cell) {
					if (nextCells.delete(parsed.ref)) removedRefs.push(parsed.ref)
					continue
				}
				const current = interactiveViewportCellFromContent(
					context,
					parsed.row,
					parsed.col,
					cell.value,
					cell.formula,
					cell.formulaInfo,
				)
				const before = previous.cells.get(parsed.ref)
				nextCells.set(parsed.ref, current)
				if (!before || !interactiveViewportCellsEqual(before, current)) changedCells.push(current)
			}
		}
		this.rememberViewportSnapshot(snapshotKey, {
			token: changeToken,
			request: previous.request,
			cells: nextCells,
		})
		return {
			patch: {
				baseToken: request.changedSince,
				changeToken,
				changedCells,
				removedRefs,
				byteLength: JSON.stringify({ changedCells, removedRefs }).length,
			},
		}
	}

	readUiRange(request: InteractiveUiRangeRequest): InteractiveUiRangeResult {
		if (this.mutableWorkbook) {
			return readInteractiveUiRange(this.mutableWorkbook, request, this.documentGeneration)
		}
		const base = this.session.workbook().readUiRange(request)
		return {
			...base,
			generation: { ...base.generation, session: this.documentGeneration },
		}
	}

	async prepareEdits(): Promise<AscendSessionPrepareEditsResult> {
		const totalStart = performance.now()
		const readLoad = this.session.inspect().load
		if (this.session.isStale()) throw new AscendException(staleInteractiveSessionError())
		const ensureStart = performance.now()
		const ensured = await this.ensureMutableWorkbook()
		const workbook = ensured.workbook
		const ensureMutableWorkbookMs = performance.now() - ensureStart
		const inspectWriteStart = performance.now()
		const writeLoad = workbook.inspect().load
		const inspectWriteMs = performance.now() - inspectWriteStart
		return {
			load: {
				read: readLoad,
				write: writeLoad,
				promotedToFull: readLoad.isPartial && !writeLoad.isPartial,
			},
			timings: {
				ensureMutableWorkbookMs,
				mutableWorkbookCached: ensured.timings.cached,
				mutableWorkbookReusedReadModel: ensured.timings.reusedReadModel,
				mutableWorkbookOpenMs: ensured.timings.openMs,
				rebaseViewportSnapshotsMs: ensured.timings.rebaseViewportSnapshotsMs,
				inspectWriteMs,
				totalMs: performance.now() - totalStart,
			},
		}
	}

	editReadiness(): AscendSessionEditReadiness {
		const read = this.session.inspect().load
		const write = this.mutableWorkbook?.inspect().load ?? null
		const timings = this.mutableWorkbookReadyTimings
		return {
			ready: write !== null && !write.isPartial,
			preparing: this.mutableWorkbook === null && this.mutableWorkbookPromise !== null,
			generation: this.documentGeneration,
			read,
			write,
			promotedToFull: write !== null && read.isPartial && !write.isPartial,
			timings: timings
				? {
						mutableWorkbookCached: timings.cached,
						mutableWorkbookReusedReadModel: timings.reusedReadModel,
						mutableWorkbookOpenMs: timings.openMs,
						rebaseViewportSnapshotsMs: timings.rebaseViewportSnapshotsMs,
					}
				: null,
		}
	}

	async apply(
		ops: readonly Operation[],
		options: AscendSessionApplyOptions = {},
	): Promise<AscendSessionApplyResult> {
		const totalStart = performance.now()
		const inspectReadStart = performance.now()
		const readLoad = this.session.inspect().load
		const inspectReadMs = performance.now() - inspectReadStart
		if (ops.length === 0 && options.recalc !== true) {
			const generations = (this.mutableWorkbook ?? this.session.workbook()).readSnapshotInfo()
				.generations
			const journal = options.journal ? emptyMutationJournal() : undefined
			return {
				apply: {
					affectedCells: [],
					sheetsModified: [],
					recalcRequired: false,
					dirtyRegions: [],
					generations,
					errors: [],
					...(journal ? { journal } : {}),
				},
				recalc: null,
				load: {
					read: readLoad,
					write: readLoad,
					promotedToFull: false,
				},
				generation: { session: this.documentGeneration, ...generations },
				timings: {
					inspectReadMs,
					ensureMutableWorkbookMs: 0,
					applyMs: 0,
					recalcMs: 0,
					generationSnapshotMs: 0,
					inspectWriteMs: 0,
					totalMs: performance.now() - totalStart,
				},
			}
		}
		if (this.session.isStale()) {
			const generations = (this.mutableWorkbook ?? this.session.workbook()).readSnapshotInfo()
				.generations
			const journal = options.journal
				? unavailableMutationJournal(
						'Mutation journal is unavailable because the interactive session is stale. Refresh the session and reread before applying edits.',
					)
				: undefined
			return {
				apply: {
					affectedCells: [],
					sheetsModified: [],
					recalcRequired: false,
					dirtyRegions: [],
					generations,
					errors: [staleInteractiveSessionError()],
					...(journal ? { journal } : {}),
				},
				recalc: null,
				load: {
					read: readLoad,
					write: readLoad,
					promotedToFull: false,
				},
				generation: { session: this.documentGeneration, ...generations },
				timings: {
					inspectReadMs,
					ensureMutableWorkbookMs: 0,
					applyMs: 0,
					recalcMs: 0,
					generationSnapshotMs: 0,
					inspectWriteMs: 0,
					totalMs: performance.now() - totalStart,
				},
			}
		}
		const ensureStart = performance.now()
		const ensured = await this.ensureMutableWorkbook()
		const workbook = ensured.workbook
		const ensureMutableWorkbookMs = performance.now() - ensureStart
		const beforeGenerations = workbook.readSnapshotInfo().generations
		const applyStart = performance.now()
		const apply = workbook.apply(ops, {
			transaction: true,
			...(options.journal !== undefined ? { journal: options.journal } : {}),
		})
		const applyMs = performance.now() - applyStart
		let recalc: import('./types.ts').RecalcResult | null = null
		let recalcMs = 0
		if (
			apply.errors.length === 0 &&
			((apply.recalcRequired && options.recalc !== false) || options.recalc === true)
		) {
			const recalcStart = performance.now()
			recalc = workbook.recalc()
			recalcMs = performance.now() - recalcStart
		}
		const generationStart = performance.now()
		const generations = workbook.readSnapshotInfo().generations
		const generationSnapshotMs = performance.now() - generationStart
		if (
			apply.errors.length === 0 &&
			sessionApplyChanged(apply, recalc, beforeGenerations, generations)
		) {
			this.documentGeneration += 1
			this.recordInteractiveChanges(ops, apply, recalc)
		}
		const inspectWriteStart = performance.now()
		const writeLoad = workbook.inspect().load
		const inspectWriteMs = performance.now() - inspectWriteStart
		return {
			apply,
			recalc,
			load: {
				read: readLoad,
				write: writeLoad,
				promotedToFull: readLoad.isPartial && !writeLoad.isPartial,
			},
			generation: { session: this.documentGeneration, ...generations },
			timings: {
				inspectReadMs,
				ensureMutableWorkbookMs,
				mutableWorkbookCached: ensured.timings.cached,
				mutableWorkbookReusedReadModel: ensured.timings.reusedReadModel,
				mutableWorkbookOpenMs: ensured.timings.openMs,
				rebaseViewportSnapshotsMs: ensured.timings.rebaseViewportSnapshotsMs,
				applyMs,
				recalcMs,
				generationSnapshotMs,
				inspectWriteMs,
				totalMs: performance.now() - totalStart,
			},
		}
	}

	async toBytes(options: WorkbookBytesOptions = {}): Promise<Uint8Array> {
		const ensured = await this.ensureMutableWorkbook()
		return ensured.workbook.toBytes(options)
	}

	isStale(): boolean {
		return this.session.isStale()
	}

	async refresh(): Promise<void> {
		await this.session.refresh()
		this.mutableWorkbook = null
		this.mutableWorkbookPromise = null
		this.mutableWorkbookReadyTimings = null
		this.documentGeneration += 1
		this.viewportSnapshots.clear()
	}

	workbook(): WorkbookDocument {
		return this.session.workbook()
	}

	close(): void {
		this.session.close()
		this.source = null
		this.mutableWorkbook = null
		this.mutableWorkbookPromise = null
		this.mutableWorkbookReadyTimings = null
		this.viewportSnapshots.clear()
	}

	private rememberViewportSnapshot(
		snapshotKey: string,
		snapshot: InteractiveViewportSnapshot,
	): void {
		const ledger = this.viewportSnapshots.get(snapshotKey) ?? { snapshots: [] }
		ledger.snapshots.push(snapshot)
		while (ledger.snapshots.length > INTERACTIVE_VIEWPORT_SNAPSHOT_RETENTION) {
			ledger.snapshots.shift()
		}
		this.viewportSnapshots.set(snapshotKey, ledger)
	}

	private async ensureMutableWorkbook(): Promise<{
		readonly workbook: AscendWorkbook
		readonly timings: MutableWorkbookEnsureTimings
	}> {
		if (this.mutableWorkbook) {
			return {
				workbook: this.mutableWorkbook,
				timings: {
					cached: true,
					reusedReadModel: false,
					openMs: 0,
					rebaseViewportSnapshotsMs: 0,
				},
			}
		}
		if (!this.mutableWorkbookPromise) {
			this.mutableWorkbookPromise = (async () => {
				const openStart = performance.now()
				let reusedReadModel = false
				let workbook = this.session.workbook().toMutableWorkbook()
				if (workbook) {
					reusedReadModel = true
				} else {
					const source = this.source
					if (source === null) throw new Error('AscendSession is closed')
					workbook = await AscendWorkbook.open(source, writableOpenOptions(this.options))
				}
				const openMs = performance.now() - openStart
				const rebaseStart = performance.now()
				if (!reusedReadModel) this.rebaseViewportSnapshots(workbook)
				const rebaseViewportSnapshotsMs = performance.now() - rebaseStart
				this.mutableWorkbook = workbook
				this.mutableWorkbookReadyTimings = {
					cached: false,
					reusedReadModel,
					openMs,
					rebaseViewportSnapshotsMs,
				}
				return workbook
			})()
			this.mutableWorkbookPromise.catch(() => {
				this.mutableWorkbookPromise = null
				this.mutableWorkbookReadyTimings = null
			})
		}
		const workbook = await this.mutableWorkbookPromise
		return {
			workbook,
			timings: this.mutableWorkbookReadyTimings ?? {
				cached: false,
				reusedReadModel: false,
				openMs: 0,
				rebaseViewportSnapshotsMs: 0,
			},
		}
	}

	private rebaseViewportSnapshots(workbook: AscendWorkbook): void {
		const rebaseRefs = new Set<string>()
		for (const ledger of this.viewportSnapshots.values()) {
			for (const snapshot of ledger.snapshots) {
				const rebased = readInteractiveViewport(
					workbook,
					snapshot.request,
					this.documentGeneration,
					snapshot.token,
				)
				const cells = interactiveCellMap(rebased.cells)
				collectInteractiveCellMapDiffRefs(snapshot.request.sheet, snapshot.cells, cells, rebaseRefs)
			}
		}
		if (rebaseRefs.size > 0) {
			this.documentGeneration += 1
			this.recentChanges.push({ generation: this.documentGeneration, refs: null })
			if (this.recentChanges.length > 128)
				this.recentChanges.splice(0, this.recentChanges.length - 128)
		}
	}

	private recordInteractiveChanges(
		ops: readonly Operation[],
		apply: import('./types.ts').ApplyResult,
		recalc: import('./types.ts').RecalcResult | null,
	): void {
		this.recentChanges.push({
			generation: this.documentGeneration,
			refs: collectInteractiveChangeRefs(ops, apply, recalc),
		})
		if (this.recentChanges.length > 128)
			this.recentChanges.splice(0, this.recentChanges.length - 128)
	}

	private changedRefsSince(
		baseGeneration: number,
	):
		| { readonly kind: 'refs'; readonly refs: Set<string> }
		| { readonly kind: 'expired' | 'invalidated' } {
		const oldest = this.recentChanges[0]
		if (oldest && oldest.generation > baseGeneration + 1) return { kind: 'expired' }
		const refs = new Set<string>()
		for (const entry of this.recentChanges) {
			if (entry.generation <= baseGeneration) continue
			if (!entry.refs) return { kind: 'invalidated' }
			for (const ref of entry.refs) refs.add(ref)
		}
		return { kind: 'refs', refs }
	}
}

function interactiveOpenOptions(options: AscendSessionOpenOptions): WorkbookLoadOptions {
	const loadOptions = stripPrepareEditsOption(options)
	const { mode, ...rest } = loadOptions
	if (mode === 'interactive' || mode === undefined) {
		if (rest.maxRows === undefined && rest.sheets === undefined) {
			return normalizeOptions({ ...rest, mode: 'full', richMetadata: true })
		}
		return normalizeOptions({ ...rest, mode: 'formula', richMetadata: true })
	}
	return normalizeOptions({ ...rest, mode })
}

function interactiveInitialOpenOptions(options: AscendSessionOpenOptions): WorkbookLoadOptions {
	const loadOptions = stripPrepareEditsOption(options)
	if (
		options.prepareEdits === true &&
		loadOptions.maxRows === undefined &&
		loadOptions.sheets === undefined &&
		(loadOptions.mode === undefined ||
			loadOptions.mode === 'interactive' ||
			loadOptions.mode === 'full')
	) {
		const { mode: _mode, ...rest } = loadOptions
		return normalizeOptions({ ...rest, mode: 'full', richMetadata: true })
	}
	return interactiveOpenOptions(loadOptions)
}

function stripPrepareEditsOption(options: AscendSessionOpenOptions): AscendSessionOpenOptions {
	const { prepareEdits: _prepareEdits, ...loadOptions } = options
	return loadOptions
}

function sessionApplyChanged(
	apply: import('./types.ts').ApplyResult,
	recalc: import('./types.ts').RecalcResult | null,
	beforeGenerations?: import('./types.ts').WorkbookGenerationInfo,
	afterGenerations?: import('./types.ts').WorkbookGenerationInfo,
): boolean {
	return (
		apply.affectedCells.length > 0 ||
		apply.sheetsModified.length > 0 ||
		apply.dirtyRegions.length > 0 ||
		apply.recalcRequired ||
		(recalc?.changed.length ?? 0) > 0 ||
		(beforeGenerations !== undefined &&
			afterGenerations !== undefined &&
			workbookGenerationsChanged(beforeGenerations, afterGenerations))
	)
}

function workbookGenerationsChanged(
	before: import('./types.ts').WorkbookGenerationInfo,
	after: import('./types.ts').WorkbookGenerationInfo,
): boolean {
	return (
		before.workbook !== after.workbook ||
		before.sheetMetadata !== after.sheetMetadata ||
		before.formulas !== after.formulas ||
		before.styles !== after.styles
	)
}

function writableOpenOptions(options: AscendSessionOpenOptions): WorkbookLoadOptions {
	const loadOptions = stripPrepareEditsOption(options)
	if (loadOptions.maxRows !== undefined || loadOptions.sheets !== undefined) {
		return interactiveOpenOptions(loadOptions)
	}
	const { mode: _mode, ...rest } = loadOptions
	return normalizeOptions({ ...rest, mode: 'full', richMetadata: true })
}

function readInteractiveViewport(
	view: WorkbookReadView,
	request: InteractiveViewportRequest,
	sessionGeneration: number,
	changeToken: string,
): InteractiveViewportResult {
	const context = createInteractiveViewportContext(view, request)
	const payload = readInteractiveCellPayload(context)
	return {
		sheet: request.sheet,
		requested: context.requested,
		viewport: context.viewport,
		generation: {
			session: sessionGeneration,
			workbook: context.snapshot.generations.workbook,
			sheetMetadata: context.snapshot.generations.sheetMetadata,
			formulas: context.snapshot.generations.formulas,
			styles: context.snapshot.generations.styles,
		},
		changeToken,
		load: context.snapshot.load,
		rowCount: context.viewport.end.row - context.viewport.start.row + 1,
		colCount: context.viewport.end.col - context.viewport.start.col + 1,
		cells: payload.cells,
		flatValues: payload.flatValues,
		displayText: payload.displayText,
		rowLayout: rowLayout(context.sheet, context.viewport),
		colLayout: colLayout(context.sheet, context.viewport),
		frozen: { rows: context.sheet.frozenRows, cols: context.sheet.frozenCols },
		merges: context.merges,
		comments: [...context.sheet.comments.entries()]
			.filter(([ref]) => cellRefInRange(ref, context.viewport))
			.map(([ref, comment]) => ({ ref, ...comment })),
		hyperlinks: [...context.sheet.hyperlinks.entries()]
			.filter(([ref]) => cellRefInRange(ref, context.viewport))
			.map(([ref, hyperlink]) => ({ ref, ...hyperlink })),
		dataValidations: context.dataValidations,
		conditionalFormats: context.conditionalFormats,
		tables: context.tables,
		autoFilter: context.autoFilter,
	}
}

function readInteractiveUiRange(
	view: WorkbookReadView,
	request: InteractiveUiRangeRequest,
	sessionGeneration: number,
): InteractiveUiRangeResult {
	const parsed = parseRange(request.range)
	const range: RangeRef = { ...parsed, sheet: parsed.sheet ?? request.sheet }
	if (range.sheet !== request.sheet) {
		throw new Error(
			`Range sheet "${range.sheet}" does not match requested sheet "${request.sheet}"`,
		)
	}
	const context = createInteractiveRangeContext(view, request.sheet, range, range)
	const payload = readInteractiveCellPayload(context)
	return {
		sheet: request.sheet,
		requested: context.requested,
		range: context.viewport,
		generation: {
			session: sessionGeneration,
			workbook: context.snapshot.generations.workbook,
			sheetMetadata: context.snapshot.generations.sheetMetadata,
			formulas: context.snapshot.generations.formulas,
			styles: context.snapshot.generations.styles,
		},
		load: context.snapshot.load,
		rowCount: context.viewport.end.row - context.viewport.start.row + 1,
		colCount: context.viewport.end.col - context.viewport.start.col + 1,
		cells: payload.cells,
		flatValues: payload.flatValues,
		displayText: payload.displayText,
	}
}

function readInteractiveCellPayload(context: ReturnType<typeof createInteractiveRangeContext>): {
	readonly cells: InteractiveViewportCell[]
	readonly flatValues: readonly (number | string | boolean | null)[]
	readonly displayText: readonly string[]
} {
	const cells: InteractiveViewportCell[] = []
	const cellCount = viewportCellCount(context.viewport)
	const flatValues = new Array<number | string | boolean | null>(cellCount).fill(null)
	const displayText = new Array<string>(cellCount).fill('')
	context.sheet.cells.forEachCellContentInRange(
		context.viewport,
		(row, col, value, formula, formulaBinding) => {
			const cell = interactiveViewportCellFromContent(
				context,
				row,
				col,
				value,
				formula,
				formulaBinding,
			)
			const index = viewportFlatIndex(context.viewport, row, col)
			flatValues[index] = cell.flatValue
			displayText[index] = cell.displayText
			cells.push(cell)
		},
	)
	return { cells, flatValues, displayText }
}

function createInteractiveViewportContext(
	view: WorkbookReadView,
	request: InteractiveViewportRequest,
) {
	const requested = viewportRange(
		request.sheet,
		request.topRow,
		request.leftCol,
		request.rowCount,
		request.colCount,
	)
	const viewport = viewportRange(
		request.sheet,
		Math.max(0, request.topRow - Math.max(0, request.overscanRows ?? 0)),
		Math.max(0, request.leftCol - Math.max(0, request.overscanCols ?? 0)),
		request.rowCount + Math.max(0, request.overscanRows ?? 0) * 2,
		request.colCount + Math.max(0, request.overscanCols ?? 0) * 2,
	)
	return createInteractiveRangeContext(view, request.sheet, requested, viewport)
}

function createInteractiveRangeContext(
	view: WorkbookReadView,
	sheetName: string,
	requested: RangeRef,
	viewport: RangeRef,
) {
	const snapshot = view.readSnapshotInfo()
	const workbook = view.getWorkbookModel()
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) throw new Error(`Sheet "${sheetName}" not found`)
	const sheetIndex = workbook.sheets.findIndex((candidate) => candidate.name === sheetName)
	const overlayIndexes = interactiveViewportOverlayIndexes(
		view,
		sheetName,
		sheet,
		snapshot.generations.sheetMetadata,
	)
	const mergeEntries = overlayIndexes.mergeIndex.intersectingEntries(viewport)
	const validationEntries = overlayIndexes.validationIndex.intersectingEntries(viewport)
	const conditionalFormatEntries =
		overlayIndexes.conditionalFormatIndex.intersectingEntries(viewport)
	const tableEntries = overlayIndexes.tableIndex.intersectingEntries(viewport)
	return {
		snapshot,
		workbook,
		sheet,
		requested,
		viewport,
		sheetIndex,
		comments: sheet.comments,
		hyperlinks: sheet.hyperlinks,
		merges: mergeEntries.map((entry) => entry.value),
		dataValidations: uniqueRangeIndexValues(validationEntries),
		conditionalFormats: uniqueRangeIndexValues(conditionalFormatEntries),
		tables: uniqueRangeIndexValues(tableEntries),
		autoFilter:
			sheet.autoFilter && sqrefIntersects(sheet.autoFilter.ref, viewport) ? sheet.autoFilter : null,
		mergeMask: viewportRangeMask(
			mergeEntries.map((entry) => entry.range),
			viewport,
		),
		validationMask: viewportRangeMask(
			validationEntries.map((entry) => entry.range),
			viewport,
		),
		conditionalFormatMask: viewportRangeMask(
			conditionalFormatEntries.map((entry) => entry.range),
			viewport,
		),
		tableMask: viewportRangeMask(
			tableEntries.map((entry) => entry.range),
			viewport,
		),
	}
}

interface ViewportRangeMask {
	has(offset: number): boolean
}

const EMPTY_VIEWPORT_RANGE_MASK: ViewportRangeMask = {
	has: () => false,
}

function viewportRangeMask(ranges: readonly RangeRef[], viewport: RangeRef): ViewportRangeMask {
	if (ranges.length === 0) return EMPTY_VIEWPORT_RANGE_MASK
	const rowCount = viewport.end.row - viewport.start.row + 1
	const colCount = viewport.end.col - viewport.start.col + 1
	const mask = new Uint8Array(rowCount * colCount)
	for (const range of ranges) {
		const intersection = rangeIntersection(range, viewport)
		if (!intersection) continue
		for (let row = intersection.start.row; row <= intersection.end.row; row++) {
			const rowOffset = (row - viewport.start.row) * colCount
			for (let col = intersection.start.col; col <= intersection.end.col; col++) {
				mask[rowOffset + col - viewport.start.col] = 1
			}
		}
	}
	return {
		has: (offset: number) => mask[offset] === 1,
	}
}

function interactiveViewportOverlayIndexes(
	view: WorkbookReadView,
	sheetName: string,
	sheet: InteractiveSheetModel,
	sheetMetadataGeneration: number,
): InteractiveViewportOverlayIndexes {
	let bySheet = interactiveViewportOverlayIndexCache.get(view)
	if (!bySheet) {
		bySheet = new Map()
		interactiveViewportOverlayIndexCache.set(view, bySheet)
	}
	const cached = bySheet.get(sheetName)
	if (
		cached &&
		cached.sheet === sheet &&
		cached.sheetMetadataGeneration === sheetMetadataGeneration
	) {
		return cached
	}
	const indexes: InteractiveViewportOverlayIndexes = {
		sheet,
		sheetMetadataGeneration,
		mergeIndex: RangeIndex.fromRanges(sheet.merges, (merge) => merge),
		validationIndex: RangeIndex.fromSqrefs(sheet.dataValidations, (validation) => validation.sqref),
		conditionalFormatIndex: RangeIndex.fromSqrefs(
			sheet.conditionalFormats,
			(format) => format.sqref,
		),
		tableIndex: RangeIndex.fromRanges(
			view.inspectSheet(sheetName)?.tables ?? [],
			(table) => table.ref,
		),
	}
	bySheet.set(sheetName, indexes)
	return indexes
}

function interactiveViewportCellFromContent(
	context: ReturnType<typeof createInteractiveViewportContext>,
	row: number,
	col: number,
	value: CellValue,
	formula: string | null,
	formulaBinding: import('@ascend/core').CellFormulaBinding | undefined,
): InteractiveViewportCell {
	const ref = toA1({ row, col })
	const styleId = context.sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
	const formulaText =
		context.sheetIndex >= 0
			? resolveCellFormulaText(context.workbook, context.sheetIndex, row, col, {
					formula,
					formulaInfo: formulaBinding,
				})
			: formula
	const flatValue = flattenViewportValue(value)
	const display = formatStyledDisplayCellValue(value, context.workbook.styles.get(styleId), {
		dateSystem: context.workbook.calcSettings.dateSystem,
	})
	const index = viewportFlatIndex(context.viewport, row, col)
	return {
		row,
		col,
		ref,
		value,
		flatValue,
		displayText: display,
		formula: formulaText,
		formulaBinding: formulaBinding ?? null,
		styleId,
		flags: {
			formula: formulaText !== null,
			comment: context.comments.has(ref),
			hyperlink: context.hyperlinks.has(ref),
			merged: context.mergeMask.has(index),
			validation: context.validationMask.has(index),
			conditionalFormat: context.conditionalFormatMask.has(index),
			table: context.tableMask.has(index),
		},
	}
}

function viewportRange(
	sheet: string,
	topRow: number,
	leftCol: number,
	rowCount: number,
	colCount: number,
): RangeRef {
	const safeRowCount = Math.max(1, Math.floor(rowCount))
	const safeColCount = Math.max(1, Math.floor(colCount))
	const start = {
		row: Math.max(0, Math.floor(topRow)),
		col: Math.max(0, Math.floor(leftCol)),
	}
	return {
		sheet,
		start,
		end: {
			row: start.row + safeRowCount - 1,
			col: start.col + safeColCount - 1,
		},
	}
}

function viewportCellCount(range: RangeRef): number {
	return (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1)
}

function viewportFlatIndex(range: RangeRef, row: number, col: number): number {
	return (row - range.start.row) * (range.end.col - range.start.col + 1) + (col - range.start.col)
}

function flattenViewportValue(value: CellValue): number | string | boolean | null {
	switch (value.kind) {
		case 'number':
		case 'string':
		case 'boolean':
			return value.value
		default:
			return null
	}
}

function uniqueRangeIndexValues<T>(entries: readonly RangeIndexEntry<T>[]): T[] {
	const values: T[] = []
	const seen = new Set<number>()
	for (const entry of entries) {
		if (seen.has(entry.sourceIndex)) continue
		seen.add(entry.sourceIndex)
		values.push(entry.value)
	}
	return values
}

function rowLayout(
	sheet: ReturnType<WorkbookReadView['getWorkbookModel']>['sheets'][number],
	range: RangeRef,
): InteractiveViewportLayoutEntry[] {
	const rows: InteractiveViewportLayoutEntry[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		const def = sheet.rowDefs.get(row)
		const size = sheet.rowHeights.get(row)
		if (def || size !== undefined) {
			rows.push({
				index: row,
				...(size !== undefined ? { size } : {}),
				...(def?.hidden !== undefined ? { hidden: def.hidden } : {}),
				...(def?.outlineLevel !== undefined ? { outlineLevel: def.outlineLevel } : {}),
				...(def?.collapsed !== undefined ? { collapsed: def.collapsed } : {}),
			})
		}
	}
	return rows
}

function colLayout(
	sheet: ReturnType<WorkbookReadView['getWorkbookModel']>['sheets'][number],
	range: RangeRef,
): InteractiveViewportLayoutEntry[] {
	const cols: InteractiveViewportLayoutEntry[] = []
	for (let col = range.start.col; col <= range.end.col; col++) {
		const def = sheet.colDefs.find(
			(candidate) =>
				(col >= candidate.min && col <= candidate.max) ||
				(col + 1 >= candidate.min && col + 1 <= candidate.max),
		)
		const size = sheet.colWidths.get(col)
		if (def || size !== undefined) {
			cols.push({
				index: col,
				...(size !== undefined ? { size } : {}),
				...(def?.hidden !== undefined ? { hidden: def.hidden } : {}),
				...(def?.outlineLevel !== undefined ? { outlineLevel: def.outlineLevel } : {}),
				...(def?.collapsed !== undefined ? { collapsed: def.collapsed } : {}),
			})
		}
	}
	return cols
}

function cellInRange(row: number, col: number, range: RangeRef): boolean {
	return (
		row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col
	)
}

function cellRefInRange(ref: string, range: RangeRef): boolean {
	const parsed = parseRange(ref)
	return cellInRange(parsed.start.row, parsed.start.col, range)
}

export function interactiveViewportSnapshotKey(request: InteractiveViewportRequest): string {
	return [
		request.sheet,
		request.topRow,
		request.leftCol,
		request.rowCount,
		request.colCount,
		request.overscanRows ?? 0,
		request.overscanCols ?? 0,
	].join(':')
}

export function mergeInteractiveViewportPatch(
	cells: readonly InteractiveViewportCell[],
	patch: InteractiveViewportPatch,
): InteractiveViewportCell[] {
	const next = new Map(cells.map((cell) => [cell.ref, cell]))
	for (const ref of patch.removedRefs) next.delete(ref)
	for (const cell of patch.changedCells) next.set(cell.ref, cell)
	return [...next.values()].sort((left, right) => left.row - right.row || left.col - right.col)
}

function interactiveCellMap(
	cells: readonly InteractiveViewportCell[],
): Map<string, InteractiveViewportCell> {
	return new Map(cells.map((cell) => [cell.ref, cell]))
}

function collectInteractiveCellMapDiffRefs(
	sheet: string,
	left: ReadonlyMap<string, InteractiveViewportCell>,
	right: ReadonlyMap<string, InteractiveViewportCell>,
	refs: Set<string>,
): void {
	for (const [ref, cell] of left) {
		const other = right.get(ref)
		if (!other || !interactiveViewportCellsEqual(other, cell)) refs.add(`${sheet}!${ref}`)
	}
	for (const ref of right.keys()) {
		if (!left.has(ref)) refs.add(`${sheet}!${ref}`)
	}
}

function diffInteractiveViewportCells(
	baseToken: string,
	changeToken: string,
	previous: ReadonlyMap<string, InteractiveViewportCell>,
	current: ReadonlyMap<string, InteractiveViewportCell>,
): InteractiveViewportPatch {
	const changedCells: InteractiveViewportCell[] = []
	const removedRefs: string[] = []
	for (const [ref, cell] of current) {
		const before = previous.get(ref)
		if (!before || !interactiveViewportCellsEqual(before, cell)) changedCells.push(cell)
	}
	for (const ref of previous.keys()) {
		if (!current.has(ref)) removedRefs.push(ref)
	}
	return {
		baseToken,
		changeToken,
		changedCells,
		removedRefs,
		byteLength: JSON.stringify({ changedCells, removedRefs }).length,
	}
}

function interactivePatchInvalidation(
	baseToken: string,
	changeToken: string,
	reason: InteractiveViewportPatchInvalidationReason,
): InteractiveViewportPatchInvalidation {
	return {
		baseToken,
		changeToken,
		reason,
		requiredAction: 'use-returned-snapshot',
	}
}

function interactiveViewportCellsEqual(
	left: InteractiveViewportCell,
	right: InteractiveViewportCell,
): boolean {
	return (
		left.row === right.row &&
		left.col === right.col &&
		left.ref === right.ref &&
		left.flatValue === right.flatValue &&
		left.displayText === right.displayText &&
		left.formula === right.formula &&
		left.styleId === right.styleId &&
		interactiveCellValuesEqual(left.value, right.value) &&
		interactiveFormulaBindingsEqual(left.formulaBinding, right.formulaBinding) &&
		left.flags.formula === right.flags.formula &&
		left.flags.comment === right.flags.comment &&
		left.flags.hyperlink === right.flags.hyperlink &&
		left.flags.merged === right.flags.merged &&
		left.flags.validation === right.flags.validation &&
		left.flags.conditionalFormat === right.flags.conditionalFormat &&
		left.flags.table === right.flags.table
	)
}

function interactiveCellValuesEqual(left: CellValue, right: CellValue): boolean {
	if (left === right) return true
	if (left.kind !== right.kind) return false
	switch (left.kind) {
		case 'empty':
			return true
		case 'number':
		case 'string':
		case 'boolean':
		case 'error':
			return left.value === (right as typeof left).value
		case 'date':
			return left.serial === (right as typeof left).serial
		case 'richText':
		case 'array':
			return JSON.stringify(left) === JSON.stringify(right)
	}
}

function interactiveFormulaBindingsEqual(
	left: InteractiveViewportCell['formulaBinding'],
	right: InteractiveViewportCell['formulaBinding'],
): boolean {
	if (left === right) return true
	if (!left || !right || left.kind !== right.kind) return false
	return JSON.stringify(left) === JSON.stringify(right)
}

function interactiveTokenGeneration(token: string): number | null {
	return interactiveTokenParts(token)?.generation ?? null
}

function interactiveTokenParts(token: string): { generation: number; version: number } | null {
	const sep = token.indexOf(':')
	if (sep < 0) return null
	const generationText = token.slice(0, sep)
	const versionText = token.slice(sep + 1)
	if (!/^\d+$/.test(generationText) || !/^\d+$/.test(versionText)) return null
	const generation = Number(generationText)
	const version = Number(versionText)
	if (!Number.isSafeInteger(generation) || !Number.isSafeInteger(version)) return null
	return { generation, version }
}

function findInteractiveViewportSnapshot(
	ledger: InteractiveViewportSnapshotLedger,
	token: string,
): InteractiveViewportSnapshot | undefined {
	return ledger.snapshots.find((snapshot) => snapshot.token === token)
}

function interactiveViewportSnapshotInvalidationReason(
	baseToken: string,
	ledger: InteractiveViewportSnapshotLedger,
): InteractiveViewportPatchInvalidationReason {
	const base = interactiveTokenParts(baseToken)
	const oldest = ledger.snapshots[0] ? interactiveTokenParts(ledger.snapshots[0].token) : null
	if (
		base &&
		oldest &&
		(base.generation < oldest.generation ||
			(base.generation === oldest.generation && base.version < oldest.version))
	) {
		return 'base-token-expired'
	}
	return 'base-token-stale'
}

function collectInteractiveChangeRefs(
	ops: readonly Operation[],
	apply: import('./types.ts').ApplyResult,
	recalc: import('./types.ts').RecalcResult | null,
): Set<string> | null {
	if (ops.some((op) => !isCellLocalViewportPatchOperation(op))) return null
	const refs = new Set<string>()
	for (const region of apply.dirtyRegions) {
		for (const ref of region.refs) refs.add(ref)
	}
	for (const region of recalc?.dirtyRegions ?? []) {
		for (const ref of region.refs) refs.add(ref)
	}
	for (const ref of recalc?.changed ?? []) refs.add(ref)
	return refs
}

function staleInteractiveSessionError() {
	return ascendError(
		'VALIDATION_ERROR',
		'Cannot promote a stale interactive session. Refresh the session and reread before applying edits.',
		{
			details: {
				rule: 'stale-interactive-session',
				staleSession: true,
				requiredAction: 'refresh',
			},
			suggestedFix:
				'Call refresh(), reread the viewport, and reapply the edit against the fresh session.',
		},
	)
}

function isCellLocalViewportPatchOperation(op: Operation): boolean {
	switch (op.op) {
		case 'setCells':
		case 'setFormula':
		case 'fillFormula':
		case 'clearRange':
		case 'setNumberFormat':
		case 'setStyle':
			return true
		default:
			return false
	}
}

function splitInteractiveFullRef(
	fullRef: string,
	workbook: ReturnType<WorkbookReadView['getWorkbookModel']>,
): { sheet: string; ref: string; row: number; col: number } | null {
	const bang = fullRef.lastIndexOf('!')
	const sheet = bang >= 0 ? fullRef.slice(0, bang).replace(/^'|'$/g, '') : workbook.sheets[0]?.name
	const ref = bang >= 0 ? fullRef.slice(bang + 1) : fullRef
	if (!sheet || !workbook.getSheet(sheet)) return null
	const parsed = parseA1Safe(ref)
	if (!parsed) return null
	return { sheet, ref, row: parsed.row, col: parsed.col }
}

function parseRangeRef(range: string, defaultSheet: string): { sheetName: string; ref: string } {
	const bang = range.indexOf('!')
	if (bang !== -1) {
		return {
			sheetName: range.substring(0, bang).replace(/^'|'$/g, ''),
			ref: range.substring(bang + 1),
		}
	}
	return { sheetName: defaultSheet, ref: range }
}

function sheetFilterFromFirstWindowOptions(
	options: WorkbookFirstWindowOptions,
): string | undefined {
	if (options.sheet) return options.sheet
	const bang = options.range.indexOf('!')
	return bang === -1 ? undefined : options.range.substring(0, bang).replace(/^'|'$/g, '')
}

function firstWindowLoadOptions(options: WorkbookFirstWindowOptions): WorkbookLoadOptions {
	const rowOffset = Math.max(0, options.rowOffset ?? 0)
	const rowLimit = Math.max(1, options.rowLimit ?? DEFAULT_FIRST_WINDOW_ROWS)
	const sheet = sheetFilterFromFirstWindowOptions(options)
	return normalizeOptions({
		mode: 'values',
		maxRows: rowOffset + rowLimit,
		...(sheet ? { sheets: [sheet] } : {}),
		...(options.richMetadata ? { richMetadata: true } : {}),
		...(options.password !== undefined ? { password: options.password } : {}),
		...(options.pivotCacheRecordMaterializeLimit !== undefined
			? { pivotCacheRecordMaterializeLimit: options.pivotCacheRecordMaterializeLimit }
			: {}),
	})
}

function readFirstWindow(
	document: WorkbookDocument,
	options: WorkbookFirstWindowOptions,
): WorkbookFirstWindowResult {
	const info = document.inspect()
	const defaultSheet = options.sheet ?? document.sheets[0]
	if (!defaultSheet) throw new Error('No sheets in workbook')
	const { sheetName, ref } = parseRangeRef(options.range, defaultSheet)
	const sheet = document.sheet(sheetName)
	if (!sheet) throw new Error(`Sheet "${sheetName}" not found`)
	const rowLimit = Math.max(1, options.rowLimit ?? DEFAULT_FIRST_WINDOW_ROWS)
	const window = sheet.readWindowCompact(ref, {
		...(options.rowOffset !== undefined ? { rowOffset: options.rowOffset } : {}),
		rowLimit,
		includeRefs: options.includeRefs ?? false,
		omitEmpty: options.omitEmpty ?? true,
		flatValues: options.flatValues ?? true,
	})
	return {
		document,
		info,
		sheet: sheetName,
		window,
		load: info.load,
	}
}

function normalizeOptions(options: WorkbookLoadOptions): WorkbookLoadOptions {
	return {
		...(options.mode ? { mode: options.mode } : {}),
		...(options.sheets ? { sheets: [...options.sheets].sort((a, b) => a.localeCompare(b)) } : {}),
		...(options.maxRows !== undefined ? { maxRows: options.maxRows } : {}),
		...(options.richMetadata ? { richMetadata: true } : {}),
		...(options.formulaModeHydrateValues !== undefined
			? { formulaModeHydrateValues: options.formulaModeHydrateValues }
			: {}),
		...(options.password !== undefined ? { password: options.password } : {}),
		...(options.pivotCacheRecordMaterializeLimit !== undefined
			? { pivotCacheRecordMaterializeLimit: options.pivotCacheRecordMaterializeLimit }
			: {}),
	}
}

function mergeOpenOptions(
	current: WorkbookLoadOptions,
	next: WorkbookLoadOptions,
): WorkbookLoadOptions {
	const mode = strongerMode(current.mode, next.mode)
	const currentSheets = current.sheets ?? []
	const nextSheets = next.sheets ?? []
	const mergedSheets =
		currentSheets.length === 0 && nextSheets.length === 0
			? undefined
			: [...new Set([...currentSheets, ...nextSheets])].sort((a, b) => a.localeCompare(b))
	const pivotCacheRecordMaterializeLimit = strongerPivotCacheRecordMaterializeLimit(
		current.pivotCacheRecordMaterializeLimit,
		next.pivotCacheRecordMaterializeLimit,
	)
	const maxRows =
		next.maxRows !== undefined
			? current.maxRows === undefined
				? undefined
				: Math.max(current.maxRows, next.maxRows)
			: next.mode === 'full'
				? undefined
				: current.maxRows
	return normalizeOptions({
		...(mode ? { mode } : {}),
		...(mergedSheets ? { sheets: mergedSheets } : {}),
		...(maxRows !== undefined ? { maxRows } : {}),
		...(current.richMetadata || next.richMetadata ? { richMetadata: true } : {}),
		...(next.formulaModeHydrateValues !== undefined ||
		current.formulaModeHydrateValues !== undefined
			? {
					formulaModeHydrateValues:
						next.formulaModeHydrateValues ?? current.formulaModeHydrateValues,
				}
			: {}),
		...(next.password !== undefined || current.password !== undefined
			? { password: next.password ?? current.password }
			: {}),
		...(pivotCacheRecordMaterializeLimit !== undefined ? { pivotCacheRecordMaterializeLimit } : {}),
	})
}

function sameOpenOptions(left: WorkbookLoadOptions, right: WorkbookLoadOptions): boolean {
	const normalizedLeft = normalizeOptions(left)
	const normalizedRight = normalizeOptions(right)
	if ((normalizedLeft.mode ?? 'full') !== (normalizedRight.mode ?? 'full')) return false
	if (normalizedLeft.maxRows !== normalizedRight.maxRows) return false
	if (
		normalizedLeft.pivotCacheRecordMaterializeLimit !==
		normalizedRight.pivotCacheRecordMaterializeLimit
	) {
		return false
	}
	if (normalizedLeft.password !== normalizedRight.password) return false
	if (normalizedLeft.formulaModeHydrateValues !== normalizedRight.formulaModeHydrateValues) {
		return false
	}
	const leftSheets = normalizedLeft.sheets ?? []
	const rightSheets = normalizedRight.sheets ?? []
	if (leftSheets.length !== rightSheets.length) return false
	return (
		(normalizedLeft.richMetadata ?? false) === (normalizedRight.richMetadata ?? false) &&
		leftSheets.every((sheet, index) => sheet === rightSheets[index])
	)
}

function strongerMode(
	current: WorkbookLoadOptions['mode'],
	next: WorkbookLoadOptions['mode'],
): WorkbookLoadOptions['mode'] {
	const rank: Record<NonNullable<WorkbookLoadOptions['mode']>, number> = {
		'metadata-only': 0,
		values: 1,
		formula: 2,
		full: 3,
	}
	if (!current) return next
	if (!next) return current
	return (rank[next] ?? 0) > (rank[current] ?? 0) ? next : current
}

function strongerPivotCacheRecordMaterializeLimit(
	current: WorkbookLoadOptions['pivotCacheRecordMaterializeLimit'],
	next: WorkbookLoadOptions['pivotCacheRecordMaterializeLimit'],
): WorkbookLoadOptions['pivotCacheRecordMaterializeLimit'] {
	if (next === undefined) return current
	const currentValue = current === undefined ? 2048 : current === 'all' ? Infinity : current
	const nextValue = next === 'all' ? Infinity : next
	if (nextValue <= currentValue) return current
	return next
}

async function readIdentity(file: string, hashContent = true): Promise<SessionFileIdentity> {
	const path = resolve(file)
	const info = await stat(path)
	return {
		path,
		size: info.size,
		mtimeMs: info.mtimeMs,
		ctimeMs: info.ctimeMs,
		sha256: hashContent
			? createHash('sha256')
					.update(await readPathBytes(path))
					.digest('hex')
			: '',
	}
}

async function readCachedPathDocument(
	file: string,
	options: WorkbookLoadOptions,
): Promise<WorkbookDocument | undefined> {
	const path = resolve(file)
	const key = makeSessionKey({ path, size: 0, mtimeMs: 0, ctimeMs: 0, sha256: '' }, options)
	const cached = sessionCache.get(key)
	if (!cached || !('path' in cached.identity)) return undefined
	if (isEntryExpired(cached)) {
		sessionCache.delete(key)
		return undefined
	}
	let info: Awaited<ReturnType<typeof stat>>
	try {
		info = await stat(path)
	} catch {
		sessionCache.delete(key)
		return undefined
	}
	const identity = cached.identity
	if (
		info.size !== identity.size ||
		info.mtimeMs !== identity.mtimeMs ||
		info.ctimeMs !== identity.ctimeMs
	) {
		sessionCache.delete(key)
		return undefined
	}
	touchCacheEntry(cached)
	return cached.document
}

function shouldHashPathSnapshot(options: WorkbookLoadOptions): boolean {
	return options.maxRows === undefined
}

async function readPathSnapshot(file: string, hashContent = true): Promise<PathSnapshot> {
	const path = resolve(file)
	for (let attempt = 0; attempt < 3; attempt++) {
		const before = await stat(path)
		const bytes = await readPathBytes(path)
		const after = await stat(path)
		if (
			before.size !== after.size ||
			before.mtimeMs !== after.mtimeMs ||
			before.ctimeMs !== after.ctimeMs
		) {
			continue
		}
		return {
			identity: {
				path,
				size: after.size,
				mtimeMs: after.mtimeMs,
				ctimeMs: after.ctimeMs,
				sha256: hashContent ? createHash('sha256').update(bytes).digest('hex') : '',
			},
			bytes,
			sourceExtension: extname(path).replace(/^\./, '').toLowerCase(),
		}
	}
	throw new AscendException(unstablePathDocumentError(path, null))
}

async function readPathBytes(path: string): Promise<Uint8Array> {
	if (typeof Bun !== 'undefined') return Bun.file(path).bytes()
	const bytes = await readFile(path)
	return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

async function openStablePathDocument(
	path: string,
	options: WorkbookLoadOptions,
): Promise<{ document: WorkbookDocument; identity: SessionFileIdentity }> {
	const hashContent = shouldHashPathSnapshot(options)
	let lastMismatch: {
		readonly before: SessionFileIdentity
		readonly after: SessionFileIdentity
		readonly snapshotSha256: string
	} | null = null
	for (let attempt = 0; attempt < 3; attempt++) {
		const before = await readIdentity(path, hashContent)
		const bytes = new Uint8Array(await readFile(path))
		const after = await readIdentity(path, hashContent)
		const snapshotSha256 = hashContent ? createHash('sha256').update(bytes).digest('hex') : ''
		if (isIdentityEqual(before, after) && (!hashContent || snapshotSha256 === after.sha256)) {
			const document = await WorkbookDocument.openPathSnapshot(path, bytes, after, options)
			return { document, identity: after }
		}
		lastMismatch = { before, after, snapshotSha256 }
		WorkbookDocument.drop(path, options)
	}
	throw new AscendException(unstablePathDocumentError(path, lastMismatch))
}

function unstablePathDocumentError(
	path: string,
	mismatch: {
		readonly before: SessionFileIdentity
		readonly after: SessionFileIdentity
		readonly snapshotSha256: string
	} | null,
) {
	return ascendError('VALIDATION_ERROR', 'Workbook file changed while opening session', {
		details: {
			rule: 'unstable-workbook-open',
			path: resolve(path),
			...(mismatch
				? {
						before: fileIdentityDetails(mismatch.before),
						after: fileIdentityDetails(mismatch.after),
						snapshotSha256: mismatch.snapshotSha256,
					}
				: {}),
		},
		suggestedFix: 'Retry after the file writer has finished and the workbook path is stable.',
	})
}

function fileIdentityDetails(identity: SessionFileIdentity): Record<string, string | number> {
	return {
		path: identity.path,
		size: identity.size,
		mtimeMs: identity.mtimeMs,
		ctimeMs: identity.ctimeMs,
		sha256: identity.sha256,
	}
}

function readBytesIdentity(bytes: Uint8Array): SessionBytesIdentity {
	const hash = createHash('sha256').update(bytes).digest('hex')
	return {
		key: `bytes:${hash}`,
		size: bytes.byteLength,
	}
}

function makeSessionKey(identity: SessionIdentity, options: WorkbookLoadOptions): string {
	const normalized = normalizeOptions(options)
	return JSON.stringify({
		source: 'path' in identity ? identity.path : identity.key,
		mode: normalized.mode ?? 'full',
		sheets: normalized.sheets ?? [],
		maxRows: normalized.maxRows ?? null,
		richMetadata: normalized.richMetadata ?? false,
		password: normalized.password ?? null,
		pivotCacheRecordMaterializeLimit: normalized.pivotCacheRecordMaterializeLimit ?? null,
	})
}

function isIdentityEqual(left: SessionIdentity, right: SessionIdentity): boolean {
	if ('path' in left && 'path' in right) {
		return (
			left.path === right.path &&
			left.size === right.size &&
			left.mtimeMs === right.mtimeMs &&
			left.ctimeMs === right.ctimeMs &&
			left.sha256 === right.sha256
		)
	}
	if (!('path' in left) && !('path' in right)) {
		return left.key === right.key && left.size === right.size
	}
	return false
}

function touchCacheEntry(entry: SessionCacheEntry): void {
	sessionCache.delete(entry.key)
	entry.accessedAt = Date.now()
	sessionCache.set(entry.key, entry)
}

function evictExpiredEntries(): void {
	for (const [key, entry] of sessionCache) {
		if (isEntryExpired(entry)) sessionCache.delete(key)
	}
}

function setCacheEntry(entry: SessionCacheEntry): void {
	entry.accessedAt = Date.now()
	sessionCache.set(entry.key, entry)
	evictExpiredEntries()
	while (
		sessionCache.size > cacheConfig.maxCacheSize ||
		totalCachedSessionBytes() > cacheConfig.maxCacheBytes
	) {
		const oldest = sessionCache.keys().next().value
		if (!oldest) break
		sessionCache.delete(oldest)
	}
}

function replaceCacheEntry(previousKey: string, entry: SessionCacheEntry): void {
	sessionCache.delete(previousKey)
	setCacheEntry(entry)
}

function totalCachedSessionBytes(): number {
	let total = 0
	for (const entry of sessionCache.values()) total += entry.sizeBytes
	return total
}

function sessionSizeBytes(
	identity: SessionIdentity,
	view: WorkbookReadView,
	usage: 'base' | 'verify',
): number {
	const workbook = view.inspect()
	const sourceBytes = workbook.load.isPartial ? 0 : identity.size
	const loadedSheets = workbook.loadedSheetCount
	const cells = workbook.cellCount ?? 0
	const metadataUnits =
		(workbook.commentCount ?? 0) +
		(workbook.conditionalFormatCount ?? 0) +
		(workbook.dataValidationCount ?? 0) +
		(workbook.imageCount ?? 0)
	const styleUnits =
		workbook.styleSummary.numFmtCount +
		workbook.styleSummary.fontCount +
		workbook.styleSummary.fillCount +
		workbook.styleSummary.borderCount +
		workbook.styleSummary.cellXfCount +
		workbook.styleSummary.dxfCount +
		workbook.styleSummary.tableStyleCount
	const baseEstimate =
		sourceBytes +
		loadedSheets * 64 * 1024 +
		cells * (view.report.sourceFormat === 'csv' ? 24 : 40) +
		metadataUnits * 256 +
		styleUnits * 128
	if (usage === 'base') return baseEstimate
	const formulaCount = view.formulaAnalysis().formulas.size
	return baseEstimate + formulaCount * 192
}
