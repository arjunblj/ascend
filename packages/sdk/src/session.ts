import { stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import { analyzeWorkbook, type WorkbookAnalysis } from '@ascend/engine'
import { check as verifyCheck, lint as verifyLint, trace as verifyTrace } from '@ascend/verify'
import { openWorkbookSource } from './load.ts'
import { WorkbookReadView } from './read-view.ts'
import type { SheetHandle } from './sheet-handle.ts'
import type { TableHandle } from './table-handle.ts'
import type {
	CheckResult,
	DefinedNameInfo,
	FormulaInfo,
	LintResult,
	RangeInfo,
	RangeWindowInfo,
	SheetInspectInfo,
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

interface SessionCacheEntry {
	readonly key: string
	readonly identity: SessionFileIdentity
	readonly session: WorkbookSession
}

const MAX_CACHED_SESSIONS = 16
const sessionCache = new Map<string, SessionCacheEntry>()

export class WorkbookSession {
	private readonly identity: SessionFileIdentity
	private readonly options: WorkbookSessionOpenOptions
	private readonly view: WorkbookReadView
	private analysis?: WorkbookAnalysis

	private constructor(
		identity: SessionFileIdentity,
		options: WorkbookSessionOpenOptions,
		view: WorkbookReadView,
	) {
		this.identity = identity
		this.options = options
		this.view = view
	}

	static async open(
		file: string,
		options: WorkbookSessionOpenOptions = {},
	): Promise<WorkbookSession> {
		const identity = await readIdentity(file)
		const key = makeSessionKey(identity.path, options)
		const cached = sessionCache.get(key)
		if (cached && isIdentityEqual(cached.identity, identity)) {
			touchCacheEntry(cached)
			return cached.session
		}

		const loaded = await openWorkbookSource(identity.path, options)
		const session = new WorkbookSession(
			identity,
			normalizeOptions(options),
			new WorkbookReadView(loaded.workbook, loaded.report, loaded.loadInfo),
		)
		setCacheEntry({ key, identity, session })
		return session
	}

	static clearCache(): void {
		sessionCache.clear()
	}

	static drop(file: string, options: WorkbookSessionOpenOptions = {}): void {
		const key = makeSessionKey(resolve(file), options)
		sessionCache.delete(key)
	}

	get file(): string {
		return this.identity.path
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

	readWindow(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeWindowInfo | undefined {
		return this.view.readWindow(sheetName, range, opts)
	}

	*streamRange(
		sheetName: string,
		range: string,
	): Generator<readonly import('./types.ts').CellInfo[]> {
		yield* this.view.streamRange(sheetName, range)
	}

	*streamWindows(
		sheetName: string,
		range: string,
		opts?: { rowLimit?: number },
	): Generator<RangeWindowInfo> {
		yield* this.view.streamWindows(sheetName, range, opts)
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
			this.getAnalysis(),
		)
		return result.ok
			? {
					ref: `${sheetName}!${ref}`,
					formula: result.value.formula,
					dependsOn: result.value.precedents.map((node) => `${node.sheet}!${node.ref}`),
					feedsInto: result.value.dependents.map((node) => `${node.sheet}!${node.ref}`),
				}
			: undefined
	}

	formula(cellRef: string): FormulaInfo | undefined {
		return this.view.formula(cellRef)
	}

	check(): CheckResult {
		const result = verifyCheck(this.view.getWorkbookModel(), this.getAnalysis())
		const issues = result.issues.map((issue) => ({
			severity: issue.severity === 'info' ? 'warning' : issue.severity,
			message: issue.message,
			...(issue.refs?.[0] ? { ref: issue.refs[0] } : {}),
		}))
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const result = verifyLint(this.view.getWorkbookModel(), this.getAnalysis())
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

	table(name: string): TableHandle | undefined {
		return this.view.table(name)
	}

	private getAnalysis(): WorkbookAnalysis {
		if (!this.analysis) {
			this.analysis = analyzeWorkbook(this.view.getWorkbookModel())
		}
		return this.analysis
	}
}

function normalizeOptions(options: WorkbookSessionOpenOptions): WorkbookSessionOpenOptions {
	return {
		...(options.mode ? { mode: options.mode } : {}),
		...(options.sheets ? { sheets: [...options.sheets].sort((a, b) => a.localeCompare(b)) } : {}),
	}
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

function makeSessionKey(path: string, options: WorkbookSessionOpenOptions): string {
	const normalized = normalizeOptions(options)
	return JSON.stringify({
		path,
		mode: normalized.mode ?? 'full',
		sheets: normalized.sheets ?? [],
	})
}

function isIdentityEqual(left: SessionFileIdentity, right: SessionFileIdentity): boolean {
	return left.path === right.path && left.size === right.size && left.mtimeMs === right.mtimeMs
}

function touchCacheEntry(entry: SessionCacheEntry): void {
	sessionCache.delete(entry.key)
	sessionCache.set(entry.key, entry)
}

function setCacheEntry(entry: SessionCacheEntry): void {
	sessionCache.set(entry.key, entry)
	while (sessionCache.size > MAX_CACHED_SESSIONS) {
		const oldest = sessionCache.keys().next().value
		if (!oldest) break
		sessionCache.delete(oldest)
	}
}
