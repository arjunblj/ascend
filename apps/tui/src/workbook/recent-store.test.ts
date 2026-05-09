import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { RecentWorkbookStore } from './recent-store.ts'

describe('RecentWorkbookStore', () => {
	test('persists recent workbooks in pinned-recency order', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-recent-'))
		try {
			const path = join(dir, 'recent.sqlite')
			const store = new RecentWorkbookStore(path)
			store.record('/tmp/first.xlsx')
			store.record('/tmp/second.xlsx')
			store.pin('/tmp/first.xlsx', true)
			store.close()

			const reopened = new RecentWorkbookStore(path)
			expect(reopened.list().map((entry) => entry.path)).toEqual([
				'/tmp/first.xlsx',
				'/tmp/second.xlsx',
			])
			expect(reopened.entries()[0]).toMatchObject({
				label: 'first.xlsx',
				path: '/tmp/first.xlsx',
				pinned: true,
			})
			reopened.close()
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})
})
