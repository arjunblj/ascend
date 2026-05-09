import { Database } from 'bun:sqlite'
import { mkdirSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'
import type { FileHubEntry } from '../runtime/types.ts'

export interface RecentWorkbookRecord {
	readonly path: string
	readonly label: string
	readonly pinned: boolean
	readonly openedAt: number
}

export class RecentWorkbookStore {
	private readonly db: Database

	constructor(path = defaultRecentStorePath()) {
		mkdirSync(dirname(path), { recursive: true })
		this.db = new Database(path, { create: true })
		this.db.run('pragma journal_mode = WAL')
		this.db.run(
			`create table if not exists recent_workbooks (
				path text primary key,
				label text not null,
				pinned integer not null default 0,
				opened_at integer not null
			)`,
		)
	}

	record(path: string): void {
		const normalized = path.trim()
		if (!normalized) return
		const label = normalized.split(/[\\/]/).pop() ?? normalized
		this.db
			.query(
				`insert into recent_workbooks (path, label, pinned, opened_at)
				 values ($path, $label, coalesce((select pinned from recent_workbooks where path = $path), 0), $openedAt)
				 on conflict(path) do update set label = excluded.label, opened_at = excluded.opened_at`,
			)
			.run({ $path: normalized, $label: label, $openedAt: Date.now() })
	}

	pin(path: string, pinned: boolean): void {
		this.db
			.query('update recent_workbooks set pinned = $pinned where path = $path')
			.run({ $path: path, $pinned: pinned ? 1 : 0 })
	}

	list(limit = 12): readonly RecentWorkbookRecord[] {
		return this.db
			.query(
				`select path, label, pinned, opened_at as openedAt
				 from recent_workbooks
				 order by pinned desc, opened_at desc
				 limit $limit`,
			)
			.all({ $limit: limit })
			.map((row) => normalizeRow(row))
	}

	entries(limit = 12): readonly FileHubEntry[] {
		return this.list(limit).map((record) => ({
			label: record.label,
			path: record.path,
			detail: record.path,
			pinned: record.pinned,
		}))
	}

	close(): void {
		this.db.close()
	}
}

function normalizeRow(row: unknown): RecentWorkbookRecord {
	const record = row as {
		readonly path?: unknown
		readonly label?: unknown
		readonly pinned?: unknown
		readonly openedAt?: unknown
	}
	const path = typeof record.path === 'string' ? record.path : ''
	const label = typeof record.label === 'string' ? record.label : path
	return {
		path,
		label,
		pinned: record.pinned === 1 || record.pinned === true,
		openedAt: typeof record.openedAt === 'number' ? record.openedAt : 0,
	}
}

function defaultRecentStorePath(): string {
	return join(
		process.env.ASCEND_TUI_STATE_DIR ?? join(homedir(), '.ascend', 'tui'),
		'recent.sqlite',
	)
}
