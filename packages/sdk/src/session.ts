import { createHash } from 'node:crypto'
import { statSync } from 'node:fs'
import { stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import { check as verifyCheck, lint as verifyLint } from '@ascend/verify'
import { partialDependencyCheckIssue, sdkCheckIssueFromVerify } from './check-issues.ts'
import { openWorkbookSource } from './load.ts'
import { WorkbookReadView } from './read-view.ts'
import type { CellSelector } from './ref-selectors.ts'
import type { SheetHandle } from './sheet-handle.ts'
import type { TableHandle } from './table-handle.ts'
import type {
	ActiveContentInfo,
	AgentViewOptions,
	AgentViewResult,
	CheckResult,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	ExternalReferenceUsageInfo,
	FormulaInfo,
	LintResult,
	PivotCacheInfo,
	PivotCacheMaterializedRowInfo,
	PivotCacheRowsOptions,
	PivotOutputAuditInfo,
	PivotRefreshPlanInfo,
	PivotTableInfo,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	SheetInspectInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TimelineCacheInfo,
	TimelineInfo,
	TraceResult,
	WorkbookInfo,
	WorkbookRefreshMetadataInfo,
	WorkbookVisualInventoryInfo,
} from './types.ts'

export interface WorkbookLoadOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
	readonly sheets?: readonly string[]
	readonly richMetadata?: boolean
	readonly password?: string
	readonly pivotCacheRecordMaterializeLimit?: number | 'all'
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

	agentView(
		sheetName: string,
		range: string,
		opts?: AgentViewOptions,
	): AgentViewResult | undefined {
		return this.view.agentView(sheetName, range, opts)
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
		const result = verifyCheck(this.view.getWorkbookModel(), {
			formulas: this.view.formulaAnalysis(),
			dependencies: this.view.dependencyAnalysis(),
		})
		this.refreshCacheFootprint('verify')
		const issues = result.issues.map(sdkCheckIssueFromVerify)
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const result = verifyLint(this.view.getWorkbookModel(), this.view.formulaAnalysis())
		this.refreshCacheFootprint('verify')
		return {
			clean: result.violations.length === 0,
			warnings: result.violations.map((violation) => ({
				rule: violation.rule,
				message: violation.message,
				ref: violation.ref,
			})),
		}
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

	pivotCacheRows(options?: PivotCacheRowsOptions): readonly PivotCacheMaterializedRowInfo[] {
		this.assertOpen()
		return this.document.pivotCacheRows(options)
	}

	pivotOutputAudits(): readonly PivotOutputAuditInfo[] {
		this.assertOpen()
		return this.document.pivotOutputAudits()
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

function normalizeOptions(options: WorkbookLoadOptions): WorkbookLoadOptions {
	return {
		...(options.mode ? { mode: options.mode } : {}),
		...(options.sheets ? { sheets: [...options.sheets].sort((a, b) => a.localeCompare(b)) } : {}),
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
	return normalizeOptions({
		...(mode ? { mode } : {}),
		...(mergedSheets ? { sheets: mergedSheets } : {}),
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
