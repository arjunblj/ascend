import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('interactive client contract', () => {
	test('apply returns dirty regions and monotonic generation tokens', () => {
		const wb = AscendWorkbook.create()

		const valueEdit = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 10 },
					{ ref: 'D4', value: 20 },
				],
			},
		])
		expect(valueEdit.errors).toEqual([])
		expect(valueEdit.dirtyRegions).toEqual([
			{
				sheet: 'Sheet1',
				range: 'B2:D4',
				refs: ['Sheet1!B2', 'Sheet1!D4'],
			},
		])
		expect(valueEdit.generations).toEqual({
			workbook: 1,
			sheetMetadata: 0,
			formulas: 1,
			styles: 0,
		})

		const styleEdit = wb.apply([
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'B2:B2', format: '0.00' },
		])
		expect(styleEdit.errors).toEqual([])
		expect(styleEdit.dirtyRegions).toEqual([
			{ sheet: 'Sheet1', range: 'B2:B2', refs: ['Sheet1!B2'] },
		])
		expect(styleEdit.generations).toEqual({
			workbook: 2,
			sheetMetadata: 0,
			formulas: 1,
			styles: 1,
		})

		const structuralEdit = wb.apply([{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 }])
		expect(structuralEdit.errors).toEqual([])
		expect(structuralEdit.generations).toEqual({
			workbook: 3,
			sheetMetadata: 1,
			formulas: 2,
			styles: 1,
		})
	})

	test('recalc returns changed dirty regions and advances formula generation once', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 1 }],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		const before = wb.recalc()
		expect(before.errors).toEqual([])

		const apply = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		const recalc = wb.recalc()

		expect(apply.dirtyRegions).toEqual([{ sheet: 'Sheet1', range: 'A1:A1', refs: ['Sheet1!A1'] }])
		expect(recalc.changed).toEqual(['Sheet1!B1'])
		expect(recalc.dirtyRegions).toEqual([{ sheet: 'Sheet1', range: 'B1:B1', refs: ['Sheet1!B1'] }])
		expect(recalc.generations.formulas).toBeGreaterThan(apply.generations.formulas)
		expect(recalc.generations.workbook).toBeGreaterThan(apply.generations.workbook)
	})

	test('apply can return a reversible journal for supported interactive edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const apply = wb.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
				{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review', author: 'agent' },
			],
			{ journal: true },
		)

		expect(apply.errors).toEqual([])
		expect(apply.journal?.supported).toBe(true)
		expect(apply.journal?.exact).toBe(true)
		expect(apply.journal?.inverseOps).toEqual([
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B2' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])

		const undo = wb.apply(apply.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 1 })
		expect(wb.sheet('Sheet1')?.comment('B2')).toBeUndefined()
	})

	test('preview journals expose lossy formula preimages without mutating the workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' }])

		const preview = wb.preview(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] }],
			{ journal: true },
		)

		expect(preview.wouldSucceed).toBe(true)
		expect(preview.journal?.supported).toBe(true)
		expect(preview.journal?.exact).toBe(false)
		expect(preview.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Formula cache for Sheet1!A1 cannot be restored with public operations',
				refs: ['Sheet1!A1'],
			},
		])
		expect(preview.journal?.inverseOps).toEqual([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' },
		])
		expect(wb.sheet('Sheet1')?.cell('A1')?.formula).toBe('1+1')
	})

	test('journal inverse ops preserve internal order when restoring cleared cell state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C3', value: 12 }] }])
		wb.apply([{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C3:C3', format: '0.00' }])

		const cleared = wb.apply([{ op: 'clearRange', sheet: 'Sheet1', range: 'C3:C3', what: 'all' }], {
			journal: true,
		})
		expect(cleared.journal?.exact).toBe(true)

		const undo = wb.apply(cleared.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('C3')?.value).toEqual({ kind: 'number', value: 12 })
		expect(wb.cellStyle('Sheet1!C3')?.numberFormat).toBe('0.00')
	})
})
