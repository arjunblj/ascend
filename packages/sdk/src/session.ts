import { createHash } from 'node:crypto'
import { stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import { check as verifyCheck, lint as verifyLint, trace as verifyTrace } from '@ascend/verify'
import { openWorkbookSource } from './load.ts'
import { WorkbookReadView } from './read-view.ts'
import type { SheetHandle } from './sheet-handle.ts'
import type { TableHandle } from './table-handle.ts'
import type {
	CheckResult,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	FormulaInfo,
	LintResult,
	PivotCacheInfo,
	PivotTableInfo,
	RangeInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	SheetInspectInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TraceResult,
	WorkbookInfo,
} from './types.ts'

export interface WorkbookSessionOpenOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
	readonly sheets?: readonly string[]
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
	readonly session: WorkbookSession
	readonly sizeBytes: number
}

const MAX_CACHED_SESSIONS = 16
const MAX_CACHED_SESSION_BYTES = 32 * 1024 * 1024
const sessionCache = new Map<string, SessionCacheEntry>()

export class WorkbookSession {
	private readonly identity: SessionIdentity
	private readonly source: string | Uint8Array
	private options: WorkbookSessionOpenOptions
	private view: WorkbookReadView

	private constructor(
		source: string | Uint8Array,
		identity: SessionIdentity,
		options: WorkbookSessionOpenOptions,
		view: WorkbookReadView,
	) {
		this.source = source
		this.identity = identity
		this.options = options
		this.view = view
	}

	static async open(
		source: string | Uint8Array,
		options: WorkbookSessionOpenOptions = {},
	): Promise<WorkbookSession> {
		const identity =
			typeof source === 'string' ? await readIdentity(source) : readBytesIdentity(source)
		const key = makeSessionKey(identity, options)
		const cached = sessionCache.get(key)
		if (cached && isIdentityEqual(cached.identity, identity)) {
			touchCacheEntry(cached)
			return cached.session
		}

		const loaded = await openWorkbookSource(source, options)
		const session = new WorkbookSession(
			source,
			identity,
			normalizeOptions(options),
			new WorkbookReadView(loaded.workbook, loaded.report, loaded.loadInfo),
		)
		setCacheEntry({ key, identity, session, sizeBytes: sessionSizeBytes(identity) })
		return session
	}

	static clearCache(): void {
		sessionCache.clear()
	}

	static drop(file: string, options: WorkbookSessionOpenOptions = {}): void {
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

	get openOptions(): WorkbookSessionOpenOptions {
		return this.options
	}

	async upgrade(options: WorkbookSessionOpenOptions): Promise<WorkbookSession> {
		const nextOptions = mergeOpenOptions(this.options, options)
		if (sameOpenOptions(this.options, nextOptions)) return this
		const loaded = await openWorkbookSource(this.source, nextOptions)
		this.options = nextOptions
		this.view.replaceWorkbook(loaded.workbook, loaded.report, loaded.loadInfo)
		return this
	}

	async hydrateSheet(
		sheetName: string,
		options?: { mode?: 'values' | 'formula' | 'full' },
	): Promise<WorkbookSession> {
		return this.upgrade({
			...(options?.mode ? { mode: options.mode } : {}),
			sheets: [sheetName],
		})
	}

	async hydrateSheets(
		sheetNames: readonly string[],
		options?: { mode?: 'values' | 'formula' | 'full' },
	): Promise<WorkbookSession> {
		return this.upgrade({
			...(options?.mode ? { mode: options.mode } : {}),
			sheets: sheetNames,
		})
	}

	inspect(): WorkbookInfo {
		return this.view.inspect()
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
		opts?: { includeRefs?: boolean },
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
		opts?: { rowOffset?: number; rowLimit?: number; includeRefs?: boolean },
	): CompactRangeWindowInfo | undefined {
		return this.view.readWindowCompact(sheetName, range, opts)
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

	trace(cellRef: string, opts?: { maxDepth?: number }): TraceResult | undefined {
		const bang = cellRef.indexOf('!')
		const sheetName = bang >= 0 ? cellRef.slice(0, bang).replace(/^'|'$/g, '') : this.sheets[0]
		const ref = bang >= 0 ? cellRef.slice(bang + 1) : cellRef
		if (!sheetName) return undefined
		const result = verifyTrace(
			this.view.getWorkbookModel(),
			sheetName,
			ref,
			opts,
			this.view.analysis(),
		)
		return result.ok
			? {
					ref: `${sheetName}!${ref}`,
					formula: result.value.formula,
					value: result.value.value,
					precedents: result.value.precedents.map((node) => ({
						ref: `${node.sheet}!${node.ref}`,
						formula: node.formula,
						value: node.value,
						depth: node.depth,
					})),
					dependents: result.value.dependents.map((node) => ({
						ref: `${node.sheet}!${node.ref}`,
						formula: node.formula,
						value: node.value,
						depth: node.depth,
					})),
					dependsOn: result.value.precedents.map((node) => `${node.sheet}!${node.ref}`),
					feedsInto: result.value.dependents.map((node) => `${node.sheet}!${node.ref}`),
				}
			: undefined
	}

	formula(cellRef: string): FormulaInfo | undefined {
		return this.view.formula(cellRef)
	}

	check(): CheckResult {
		const result = verifyCheck(this.view.getWorkbookModel(), this.view.analysis())
		const issues = result.issues.map((issue) => ({
			severity: issue.severity === 'info' ? 'warning' : issue.severity,
			message: issue.message,
			...(issue.refs?.[0] ? { ref: issue.refs[0] } : {}),
		}))
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const result = verifyLint(this.view.getWorkbookModel(), this.view.analysis())
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

	slicerCaches(): readonly SlicerCacheInfo[] {
		return this.view.slicerCaches()
	}

	slicers(): readonly SlicerInfo[] {
		return this.view.slicers()
	}

	workbookViews(): readonly import('./types.ts').WorkbookViewInfo[] {
		return this.view.workbookViews()
	}

	externalReferences(): readonly string[] {
		return this.view.externalReferences()
	}
}

function normalizeOptions(options: WorkbookSessionOpenOptions): WorkbookSessionOpenOptions {
	return {
		...(options.mode ? { mode: options.mode } : {}),
		...(options.sheets ? { sheets: [...options.sheets].sort((a, b) => a.localeCompare(b)) } : {}),
	}
}

function mergeOpenOptions(
	current: WorkbookSessionOpenOptions,
	next: WorkbookSessionOpenOptions,
): WorkbookSessionOpenOptions {
	const mode = strongerMode(current.mode, next.mode)
	const currentSheets = current.sheets ?? []
	const nextSheets = next.sheets ?? []
	const mergedSheets =
		currentSheets.length === 0 && nextSheets.length === 0
			? undefined
			: [...new Set([...currentSheets, ...nextSheets])].sort((a, b) => a.localeCompare(b))
	return normalizeOptions({
		...(mode ? { mode } : {}),
		...(mergedSheets ? { sheets: mergedSheets } : {}),
	})
}

function sameOpenOptions(
	left: WorkbookSessionOpenOptions,
	right: WorkbookSessionOpenOptions,
): boolean {
	const normalizedLeft = normalizeOptions(left)
	const normalizedRight = normalizeOptions(right)
	if ((normalizedLeft.mode ?? 'full') !== (normalizedRight.mode ?? 'full')) return false
	const leftSheets = normalizedLeft.sheets ?? []
	const rightSheets = normalizedRight.sheets ?? []
	if (leftSheets.length !== rightSheets.length) return false
	return leftSheets.every((sheet, index) => sheet === rightSheets[index])
}

function strongerMode(
	current: WorkbookSessionOpenOptions['mode'],
	next: WorkbookSessionOpenOptions['mode'],
): WorkbookSessionOpenOptions['mode'] {
	const rank: Record<NonNullable<WorkbookSessionOpenOptions['mode']>, number> = {
		'metadata-only': 0,
		values: 1,
		formula: 2,
		full: 3,
	}
	if (!current) return next
	if (!next) return current
	return (rank[next] ?? 0) > (rank[current] ?? 0) ? next : current
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

function readBytesIdentity(bytes: Uint8Array): SessionBytesIdentity {
	const hash = createHash('sha256').update(bytes).digest('hex')
	return {
		key: `bytes:${hash}`,
		size: bytes.byteLength,
	}
}

function makeSessionKey(identity: SessionIdentity, options: WorkbookSessionOpenOptions): string {
	const normalized = normalizeOptions(options)
	return JSON.stringify({
		source: 'path' in identity ? identity.path : identity.key,
		mode: normalized.mode ?? 'full',
		sheets: normalized.sheets ?? [],
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
	sessionCache.set(entry.key, entry)
}

function setCacheEntry(entry: SessionCacheEntry): void {
	sessionCache.set(entry.key, entry)
	while (
		sessionCache.size > MAX_CACHED_SESSIONS ||
		totalCachedSessionBytes() > MAX_CACHED_SESSION_BYTES
	) {
		const oldest = sessionCache.keys().next().value
		if (!oldest) break
		sessionCache.delete(oldest)
	}
}

function totalCachedSessionBytes(): number {
	let total = 0
	for (const entry of sessionCache.values()) total += entry.sizeBytes
	return total
}

function sessionSizeBytes(identity: SessionIdentity): number {
	return identity.size
}
