import { stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import type { SheetHandle } from './sheet-handle.ts'
import type { TableHandle } from './table-handle.ts'
import type {
	DefinedNameInfo,
	FormulaInfo,
	RangeInfo,
	RangeWindowInfo,
	SheetInspectInfo,
	TraceResult,
	WorkbookInfo,
} from './types.ts'
import { AscendWorkbook } from './workbook.ts'

export interface WorkbookSessionOpenOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values'
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
	private readonly workbook: AscendWorkbook

	private constructor(
		identity: SessionFileIdentity,
		options: WorkbookSessionOpenOptions,
		workbook: AscendWorkbook,
	) {
		this.identity = identity
		this.options = options
		this.workbook = workbook
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

		const workbook = await AscendWorkbook.open(identity.path, options)
		const session = new WorkbookSession(identity, normalizeOptions(options), workbook)
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
		return this.workbook.sheets
	}

	get report() {
		return this.workbook.report
	}

	get openOptions(): WorkbookSessionOpenOptions {
		return this.options
	}

	inspect(): WorkbookInfo {
		return this.workbook.inspect()
	}

	inspectSheet(name: string): SheetInspectInfo | undefined {
		return this.workbook.inspectSheet(name)
	}

	sheet(name: string): SheetHandle | undefined {
		return this.workbook.sheet(name)
	}

	readRange(sheetName: string, range: string): RangeInfo | undefined {
		return this.workbook.readRange(sheetName, range)
	}

	readWindow(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeWindowInfo | undefined {
		return this.workbook.readWindow(sheetName, range, opts)
	}

	*streamRange(
		sheetName: string,
		range: string,
	): Generator<readonly import('./types.ts').CellInfo[]> {
		yield* this.workbook.streamRange(sheetName, range)
	}

	*streamWindows(
		sheetName: string,
		range: string,
		opts?: { rowLimit?: number },
	): Generator<RangeWindowInfo> {
		yield* this.workbook.streamWindows(sheetName, range, opts)
	}

	trace(cellRef: string, opts?: { maxDepth?: number }): TraceResult | undefined {
		return this.workbook.trace(cellRef, opts)
	}

	formula(cellRef: string): FormulaInfo | undefined {
		return this.workbook.formula(cellRef)
	}

	definedName(name: string, sheetName?: string): DefinedNameInfo | undefined {
		return this.workbook.definedName(name, sheetName)
	}

	table(name: string): TableHandle | undefined {
		return this.workbook.table(name)
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
