import { createHash } from 'node:crypto'
import { statSync } from 'node:fs'
import { readFile, stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import {
	DEFAULT_STYLE_ID,
	parseRange,
	RangeIndex,
	type RangeIndexEntry,
	type RangeRef,
	rangeMaskOffsets,
	sqrefIntersects,
	toA1,
} from '@ascend/core'
import { resolveCellFormulaText } from '@ascend/engine'
import { inspectXlsxPackageGraph, type XlsxPackageGraph } from '@ascend/io-xlsx'
import type { CellValue, Operation } from '@ascend/schema'
import { check as verifyCheck, lint as verifyLint } from '@ascend/verify'
import {
	partialDependencyCheckIssue,
	partialDependencyLintWarning,
	sdkCheckIssueFromVerify,
} from './check-issues.ts'
import { formatStyledDisplayCellValue } from './format-helpers.ts'
import { openWorkbookSource } from './load.ts'
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
} from './types.ts'
import { type ApplyOptions, AscendWorkbook } from './workbook.ts'

export interface WorkbookLoadOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
	readonly sheets?: readonly string[]
	readonly maxRows?: number
	readonly richMetadata?: boolean
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
	readonly changedCells: readonly InteractiveViewportCell[]
	readonly removedRefs: readonly string[]
	readonly byteLength: number
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
}

export interface AscendSessionApplyOptions extends Pick<ApplyOptions, 'journal'> {
	readonly recalc?: boolean
}

export interface AscendSessionApplyResult {
	readonly apply: import('./types.ts').ApplyResult
	readonly recalc: import('./types.ts').RecalcResult | null
	readonly generation: {
		readonly session: number
		readonly workbook: number
		readonly sheetMetadata: number
		readonly formulas: number
		readonly styles: number
	}
}

interface SessionFileIdentity {
	readonly path: string
	readonly size: number
	readonly mtimeMs: number
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

export interface SessionCacheOptions {
	readonly maxCacheSize?: number
	readonly maxCacheAge?: number
	readonly maxCacheBytes?: number
}

const cacheConfig = {
	maxCacheSize: 50,
	maxCacheAge: 5 * 60 * 1000,
	maxCacheBytes: 32 * 1024 * 1024,
}

const DEFAULT_FIRST_WINDOW_ROWS = 500

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

	private constructor(
		cacheKey: string,
		source: string | Uint8Array,
		identity: SessionIdentity,
		options: WorkbookLoadOptions,
		view: WorkbookReadView,
	) {
		this.cacheKey = cacheKey
		this.source = source
		this.identity = identity
		this.options = options
		this.view = view
	}

	static async open(
		source: string | Uint8Array,
		options: WorkbookLoadOptions = {},
	): Promise<WorkbookDocument> {
		const identity =
			typeof source === 'string' ? await readIdentity(source) : readBytesIdentity(source)
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

		const loaded = await openWorkbookSource(source, options)
		const document = new WorkbookDocument(
			key,
			source,
			identity,
			normalizeOptions(options),
			new WorkbookReadView(loaded.workbook, loaded.report, loaded.loadInfo),
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
		const key = makeSessionKey({ path: resolve(file), size: 0, mtimeMs: 0 }, options)
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

	inspect(): WorkbookInfo {
		return this.view.inspect()
	}

	visualInventory(): WorkbookVisualInventoryInfo {
		return this.view.visualInventory()
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
	private readonly source: { path?: string; bytes?: Uint8Array }
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
		if (this.source.bytes || !this.fileIdentity) return false
		try {
			const info = statSync(this.fileIdentity.path)
			return info.size !== this.fileIdentity.size || info.mtimeMs !== this.fileIdentity.mtimeMs
		} catch {
			return true
		}
	}

	async refresh(): Promise<void> {
		this.assertOpen()
		const path = this.source.path
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
	private readonly source: string | Uint8Array
	private readonly options: AscendSessionOpenOptions
	private mutableWorkbook: AscendWorkbook | null = null
	private documentGeneration = 0
	private changeVersion = 0
	private readonly viewportSnapshots = new Map<
		string,
		{
			readonly token: string
			readonly request: InteractiveViewportRequest
			readonly cells: Map<string, InteractiveViewportCell>
		}
	>()

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
		const session = await WorkbookSession.open(pathOrBytes, interactiveOpenOptions(options))
		return new AscendSession(session, pathOrBytes, options)
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
		const previous = this.viewportSnapshots.get(snapshotKey)
		this.viewportSnapshots.set(snapshotKey, { token: changeToken, request, cells: currentCells })
		if (!previous || previous.token !== request.changedSince) return result
		const patch = diffInteractiveViewportCells(previous.token, previous.cells, currentCells)
		return { ...result, patch }
	}

	async apply(
		ops: readonly Operation[],
		options: AscendSessionApplyOptions = {},
	): Promise<AscendSessionApplyResult> {
		const workbook = await this.ensureMutableWorkbook()
		const apply = workbook.apply(ops, {
			transaction: true,
			...(options.journal !== undefined ? { journal: options.journal } : {}),
		})
		let recalc: import('./types.ts').RecalcResult | null = null
		if (apply.errors.length === 0 && apply.recalcRequired && options.recalc !== false) {
			recalc = workbook.recalc()
		}
		if (apply.errors.length === 0) {
			this.documentGeneration += 1
		}
		const generations = workbook.readSnapshotInfo().generations
		return {
			apply,
			recalc,
			generation: { session: this.documentGeneration, ...generations },
		}
	}

	isStale(): boolean {
		return this.session.isStale()
	}

	async refresh(): Promise<void> {
		await this.session.refresh()
		this.mutableWorkbook = null
		this.documentGeneration += 1
		this.viewportSnapshots.clear()
	}

	workbook(): WorkbookDocument {
		return this.session.workbook()
	}

	close(): void {
		this.session.close()
		this.mutableWorkbook = null
		this.viewportSnapshots.clear()
	}

	private async ensureMutableWorkbook(): Promise<AscendWorkbook> {
		if (this.mutableWorkbook) return this.mutableWorkbook
		this.mutableWorkbook = await AscendWorkbook.open(this.source, writableOpenOptions(this.options))
		this.rebaseViewportSnapshots(this.mutableWorkbook)
		return this.mutableWorkbook
	}

	private rebaseViewportSnapshots(workbook: AscendWorkbook): void {
		for (const [key, snapshot] of this.viewportSnapshots) {
			const rebased = readInteractiveViewport(
				workbook,
				snapshot.request,
				this.documentGeneration,
				snapshot.token,
			)
			this.viewportSnapshots.set(key, {
				...snapshot,
				cells: interactiveCellMap(rebased.cells),
			})
		}
	}
}

function interactiveOpenOptions(options: AscendSessionOpenOptions): WorkbookLoadOptions {
	const { mode, ...rest } = options
	if (mode === 'interactive' || mode === undefined) {
		return normalizeOptions({ ...rest, mode: 'formula', richMetadata: true })
	}
	return normalizeOptions({ ...rest, mode })
}

function writableOpenOptions(options: AscendSessionOpenOptions): WorkbookLoadOptions {
	if (options.maxRows !== undefined || options.sheets !== undefined) {
		return interactiveOpenOptions(options)
	}
	const { mode: _mode, ...rest } = options
	return normalizeOptions({ ...rest, mode: 'full', richMetadata: true })
}

function readInteractiveViewport(
	view: WorkbookReadView,
	request: InteractiveViewportRequest,
	sessionGeneration: number,
	changeToken: string,
): InteractiveViewportResult {
	const snapshot = view.readSnapshotInfo()
	const workbook = view.getWorkbookModel()
	const sheet = workbook.getSheet(request.sheet)
	if (!sheet) throw new Error(`Sheet "${request.sheet}" not found`)
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
	const sheetIndex = workbook.sheets.findIndex((candidate) => candidate.name === request.sheet)
	const comments = sheet.comments
	const hyperlinks = sheet.hyperlinks
	const mergeEntries = RangeIndex.fromRanges(sheet.merges, (merge) => merge).intersectingEntries(
		viewport,
	)
	const validationEntries = RangeIndex.fromSqrefs(
		sheet.dataValidations,
		(validation) => validation.sqref,
	).intersectingEntries(viewport)
	const conditionalFormatEntries = RangeIndex.fromSqrefs(
		sheet.conditionalFormats,
		(format) => format.sqref,
	).intersectingEntries(viewport)
	const tableEntries = RangeIndex.fromRanges(
		view.inspectSheet(request.sheet)?.tables ?? [],
		(table) => table.ref,
	).intersectingEntries(viewport)
	const merges = mergeEntries.map((entry) => entry.value)
	const dataValidations = uniqueRangeIndexValues(validationEntries)
	const conditionalFormats = uniqueRangeIndexValues(conditionalFormatEntries)
	const tables = uniqueRangeIndexValues(tableEntries)
	const mergeMask = rangeMaskOffsets(
		mergeEntries.map((entry) => entry.range),
		viewport,
	)
	const validationMask = rangeMaskOffsets(
		validationEntries.map((entry) => entry.range),
		viewport,
	)
	const conditionalFormatMask = rangeMaskOffsets(
		conditionalFormatEntries.map((entry) => entry.range),
		viewport,
	)
	const tableMask = rangeMaskOffsets(
		tableEntries.map((entry) => entry.range),
		viewport,
	)
	const autoFilter =
		sheet.autoFilter && sqrefIntersects(sheet.autoFilter.ref, viewport) ? sheet.autoFilter : null
	const cells: InteractiveViewportCell[] = []
	const flatValues = Array.from({ length: viewportCellCount(viewport) }, () => null) as Array<
		number | string | boolean | null
	>
	const displayText = Array.from({ length: viewportCellCount(viewport) }, () => '')
	sheet.cells.forEachCellContentInRange(viewport, (row, col, value, formula, formulaBinding) => {
		const ref = toA1({ row, col })
		const styleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
		const formulaText =
			sheetIndex >= 0
				? resolveCellFormulaText(workbook, sheetIndex, row, col, {
						formula,
						formulaInfo: formulaBinding,
					})
				: formula
		const flatValue = flattenViewportValue(value)
		const display = formatStyledDisplayCellValue(value, workbook.styles.get(styleId), {
			dateSystem: workbook.calcSettings.dateSystem,
		})
		const index = viewportFlatIndex(viewport, row, col)
		flatValues[index] = flatValue
		displayText[index] = display
		cells.push({
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
				comment: comments.has(ref),
				hyperlink: hyperlinks.has(ref),
				merged: mergeMask.has(index),
				validation: validationMask.has(index),
				conditionalFormat: conditionalFormatMask.has(index),
				table: tableMask.has(index),
			},
		})
	})
	return {
		sheet: request.sheet,
		requested,
		viewport,
		generation: {
			session: sessionGeneration,
			workbook: snapshot.generations.workbook,
			sheetMetadata: snapshot.generations.sheetMetadata,
			formulas: snapshot.generations.formulas,
			styles: snapshot.generations.styles,
		},
		changeToken,
		load: snapshot.load,
		rowCount: viewport.end.row - viewport.start.row + 1,
		colCount: viewport.end.col - viewport.start.col + 1,
		cells,
		flatValues,
		displayText,
		rowLayout: rowLayout(sheet, viewport),
		colLayout: colLayout(sheet, viewport),
		frozen: { rows: sheet.frozenRows, cols: sheet.frozenCols },
		merges,
		comments: [...sheet.comments.entries()]
			.filter(([ref]) => cellRefInRange(ref, viewport))
			.map(([ref, comment]) => ({ ref, ...comment })),
		hyperlinks: [...sheet.hyperlinks.entries()]
			.filter(([ref]) => cellRefInRange(ref, viewport))
			.map(([ref, hyperlink]) => ({ ref, ...hyperlink })),
		dataValidations,
		conditionalFormats,
		tables,
		autoFilter,
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

function interactiveViewportSnapshotKey(request: InteractiveViewportRequest): string {
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

function interactiveCellMap(
	cells: readonly InteractiveViewportCell[],
): Map<string, InteractiveViewportCell> {
	return new Map(cells.map((cell) => [cell.ref, cell]))
}

function diffInteractiveViewportCells(
	baseToken: string,
	previous: ReadonlyMap<string, InteractiveViewportCell>,
	current: ReadonlyMap<string, InteractiveViewportCell>,
): InteractiveViewportPatch {
	const changedCells: InteractiveViewportCell[] = []
	const removedRefs: string[] = []
	for (const [ref, cell] of current) {
		const before = previous.get(ref)
		if (!before || JSON.stringify(before) !== JSON.stringify(cell)) changedCells.push(cell)
	}
	for (const ref of previous.keys()) {
		if (!current.has(ref)) removedRefs.push(ref)
	}
	return {
		baseToken,
		changedCells,
		removedRefs,
		byteLength: JSON.stringify({ changedCells, removedRefs }).length,
	}
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

async function readIdentity(file: string): Promise<SessionFileIdentity> {
	const path = resolve(file)
	const info = await stat(path)
	return {
		path,
		size: info.size,
		mtimeMs: info.mtimeMs,
	}
}

async function openStablePathDocument(
	path: string,
	options: WorkbookLoadOptions,
): Promise<{ document: WorkbookDocument; identity: SessionFileIdentity }> {
	let lastDocument: WorkbookDocument | null = null
	for (let attempt = 0; attempt < 2; attempt++) {
		const before = await readIdentity(path)
		const document = await WorkbookDocument.open(path, options)
		const after = await readIdentity(path)
		if (isIdentityEqual(before, after)) return { document, identity: after }
		lastDocument = document
		WorkbookDocument.drop(path, options)
	}
	if (!lastDocument) {
		return {
			document: await WorkbookDocument.open(path, options),
			identity: await readIdentity(path),
		}
	}
	return { document: lastDocument, identity: await readIdentity(path) }
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
		return left.path === right.path && left.size === right.size && left.mtimeMs === right.mtimeMs
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
	const sourceBytes = identity.size
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
