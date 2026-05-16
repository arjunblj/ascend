import { describe, expect, test } from 'bun:test'
import { Workbook } from '@ascend/core'
import { applyOperations } from '../../../engine/src/index.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('image lifecycle operations', () => {
	test('insertImage writes generated drawing relationships and media', () => {
		const wb = new Workbook()
		wb.addSheet('Images')
		const applied = applyOperations(wb, [
			{
				op: 'insertImage',
				sheet: 'Images',
				contentBase64: 'iVBORw0KGgo=',
				contentType: 'image/png',
				name: 'Logo',
				description: 'Brand logo',
				anchor: { kind: 'oneCell', from: { row: 1, col: 1 }, cx: 320000, cy: 240000 },
			},
		])
		expectOk(applied)

		const written = writeXlsx(wb)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)

		expect(reopened.value.workbook.sheets[0]?.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			name: 'Logo',
			description: 'Brand logo',
			anchor: { kind: 'oneCell', from: { row: 1, col: 1 }, cx: 320000, cy: 240000 },
		})
	})

	test('structural edits shift image anchors through write and reopen', () => {
		const wb = new Workbook()
		wb.addSheet('Images')
		const applied = applyOperations(wb, [
			{
				op: 'insertImage',
				sheet: 'Images',
				contentBase64: 'iVBORw0KGgo=',
				contentType: 'image/png',
				name: 'Logo',
				anchor: {
					kind: 'twoCell',
					from: { row: 1, col: 1 },
					to: { row: 4, col: 3 },
					editAs: 'oneCell',
				},
			},
			{ op: 'insertRows', sheet: 'Images', at: 1, count: 2 },
			{ op: 'insertCols', sheet: 'Images', at: 2, count: 1 },
		])
		expectOk(applied)

		const written = writeXlsx(wb)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)

		expect(reopened.value.workbook.sheets[0]?.imageRefs[0]).toMatchObject({
			name: 'Logo',
			anchor: {
				kind: 'twoCell',
				from: { row: 3, col: 1 },
				to: { row: 6, col: 4 },
				editAs: 'oneCell',
			},
		})
	})

	test('deleteImage removes generated drawing when last image is deleted', () => {
		const wb = new Workbook()
		wb.addSheet('Images')
		const inserted = applyOperations(wb, [
			{
				op: 'insertImage',
				sheet: 'Images',
				contentBase64: 'iVBORw0KGgo=',
				contentType: 'image/png',
				name: 'Logo',
			},
		])
		expectOk(inserted)
		const deleted = applyOperations(wb, [{ op: 'deleteImage', sheet: 'Images', name: 'Logo' }])
		expectOk(deleted)

		const written = writeXlsx(wb)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)

		expect(reopened.value.workbook.sheets[0]?.drawingRefs.hasDrawing).toBe(false)
		expect(reopened.value.workbook.sheets[0]?.imageRefs).toEqual([])
	})
})
