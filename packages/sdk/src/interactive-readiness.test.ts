import { describe, expect, test } from 'bun:test'
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
})
