import { describe, expect, test } from 'bun:test'
import { readFile, stat, unlink, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendSession, AscendWorkbook } from './index.ts'

describe('interactive edit readiness', () => {
	test('open can return a fully prepared edit-ready session', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'interactive',
			prepareEdits: true,
		})
		try {
			const readiness = session.editReadiness()
			expect(readiness).toMatchObject({
				ready: true,
				preparing: false,
				promotedToFull: false,
				read: { mode: 'full', isPartial: false },
				write: { mode: 'full', isPartial: false },
				timings: { mutableWorkbookReusedReadModel: true },
			})
			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			])
			expect(edit.apply.errors).toEqual([])
			expect(edit.timings.mutableWorkbookCached).toBe(true)
			expect(
				session.readViewport({
					sheet: 'Sheet1',
					topRow: 0,
					leftCol: 0,
					rowCount: 1,
					colCount: 1,
				}).cells[0]?.flatValue,
			).toBe(2)
		} finally {
			session.close()
		}
	})

	test('prepareEdits option stays read-only for capped sessions', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
		])
		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'interactive',
			maxRows: 1,
			prepareEdits: true,
		})
		try {
			expect(session.editReadiness()).toMatchObject({
				ready: false,
				preparing: false,
				promotedToFull: false,
				read: { isPartial: true, maxRows: 1 },
				write: { isPartial: true, maxRows: 1 },
			})
			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 3 }] },
			])
			expect(edit.apply.errors[0]?.message).toContain('partial workbook')
			expect(
				session.readViewport({
					sheet: 'Sheet1',
					topRow: 0,
					leftCol: 0,
					rowCount: 1,
					colCount: 1,
				}).cells[0]?.flatValue,
			).toBe(1)
		} finally {
			session.close()
		}
	})

	test('sessions can serialize applied edits for grid clients', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'interactive',
			prepareEdits: true,
		})
		try {
			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] },
			])
			expect(edit.apply.errors).toEqual([])

			const bytes = await session.toBytes({ compressionProfile: 'store' })
			const reopened = await AscendWorkbook.open(bytes)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'number',
				value: 42,
			})
		} finally {
			session.close()
		}
	})

	test('edit-ready path sessions reject stale source writes', async () => {
		const input = join(tmpdir(), `ascend-edit-ready-stale-${Date.now()}-${process.pid}.xlsx`)
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'original' }] }])
		await wb.save(input)

		const session = await AscendSession.open(input, {
			mode: 'interactive',
			prepareEdits: true,
		})
		try {
			expect(session.editReadiness().ready).toBe(true)
			const changed = AscendWorkbook.create()
			changed.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'changed elsewhere' }] },
			])
			await changed.save(input)
			expect(session.isStale()).toBe(true)

			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'agent edit' }] },
			])
			expect(edit.apply.errors[0]).toMatchObject({
				code: 'VALIDATION_ERROR',
				details: {
					rule: 'stale-interactive-session',
					staleSession: true,
					requiredAction: 'refresh',
				},
			})
			expect(session.editReadiness().ready).toBe(true)

			const reopened = await AscendWorkbook.open(input)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'changed elsewhere',
			})
		} finally {
			session.close()
			await unlink(input).catch(() => {})
		}
	})

	test('edit-ready path sessions reject same-size same-mtime source rewrites', async () => {
		const input = join(tmpdir(), `ascend-edit-ready-hash-stale-${Date.now()}-${process.pid}.xlsx`)
		const stableTime = new Date('2026-05-14T00:00:00.000Z')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'original' }] }])
		await wb.save(input)
		await utimes(input, stableTime, stableTime)

		const session = await AscendSession.open(input, {
			mode: 'interactive',
			prepareEdits: true,
		})
		try {
			const before = await stat(input)
			const bytes = new Uint8Array(await readFile(input))
			bytes[bytes.byteLength - 1] ^= 1
			await writeFile(input, bytes)
			await utimes(input, stableTime, stableTime)
			const after = await stat(input)
			expect(after.size).toBe(before.size)
			expect(after.mtimeMs).toBe(before.mtimeMs)
			expect(session.isStale()).toBe(true)

			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'agent edit' }] },
			])
			expect(edit.apply.errors[0]).toMatchObject({
				code: 'VALIDATION_ERROR',
				details: {
					rule: 'stale-interactive-session',
					staleSession: true,
					requiredAction: 'refresh',
				},
			})
		} finally {
			session.close()
			await unlink(input).catch(() => {})
		}
	})
})
