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

	test('journal inverse ops restore merge and unmerge metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' }])

		const unmerged = wb.apply([{ op: 'unmergeCells', sheet: 'Sheet1', range: 'A1:B2' }], {
			journal: true,
		})

		expect(unmerged.journal?.exact).toBe(true)
		expect(unmerged.journal?.inverseOps).toEqual([
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' },
		])

		const undoUnmerge = wb.apply(unmerged.journal?.inverseOps ?? [], { transaction: true })
		expect(undoUnmerge.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.merges).toEqual([
			{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
		])

		const wb2 = AscendWorkbook.create()
		const merged = wb2.apply([{ op: 'mergeCells', sheet: 'Sheet1', range: 'C1:D1' }], {
			journal: true,
		})
		expect(merged.journal?.inverseOps).toEqual([
			{ op: 'unmergeCells', sheet: 'Sheet1', range: 'C1:D1' },
		])
	})

	test('journal inverse ops restore data validation changes', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A10',
				rule: {
					type: 'list',
					formula1: '"Open,Closed"',
					allowBlank: true,
					showErrorMessage: true,
				},
			},
		])

		const changed = wb.apply(
			[
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'A1:A10',
					rule: { type: 'whole', operator: 'greaterThan', formula1: '0' },
				},
			],
			{ journal: true },
		)

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A10',
				rule: {
					type: 'list',
					formula1: '"Open,Closed"',
					allowBlank: true,
					showErrorMessage: true,
				},
			},
		])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getDataValidations()).toEqual([
			{ type: 'list', formula: '"Open,Closed"', range: 'A1:A10' },
		])
	})

	test('journal inverse ops restore simple worksheet auto filters', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10', column: 1, values: ['Open'] },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				sortRef: 'A2:C10',
				sortBy: 'B2:B10',
				descending: true,
			},
		])

		const cleared = wb.apply([{ op: 'clearAutoFilter', sheet: 'Sheet1' }], { journal: true })

		expect(cleared.journal?.exact).toBe(true)
		expect(cleared.journal?.inverseOps).toEqual([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10' },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				column: 1,
				values: ['Open'],
			},
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				sortRef: 'A2:C10',
				sortBy: 'B2:B10',
				descending: true,
			},
		])

		const undo = wb.apply(cleared.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.autoFilter).toEqual({
			ref: 'A1:C10',
			columns: [{ colId: 1, kind: 'filters', values: ['Open'] }],
			sortState: {
				ref: 'A2:C10',
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		})

		const wb2 = AscendWorkbook.create()
		const created = wb2.apply([{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'D1:E5' }], {
			journal: true,
		})
		expect(created.journal?.inverseOps).toEqual([{ op: 'clearAutoFilter', sheet: 'Sheet1' }])
	})

	test('journal inverse ops restore conditional format replacements', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B1:B10',
				rule: {
					type: 'cellIs',
					operator: 'greaterThan',
					formula: '10',
					priority: 3,
					style: { font: { bold: true } },
				},
			},
		])

		const changed = wb.apply(
			[
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'B1:B10',
					rule: { type: 'expression', formula: 'B1<0', priority: 1 },
				},
			],
			{ journal: true },
		)

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'B1:B10' },
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B1:B10',
				rule: {
					type: 'cellIs',
					operator: 'greaterThan',
					formula: '10',
					priority: 3,
					style: { font: { bold: true } },
				},
				mode: 'replace',
			},
		])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getConditionalFormats()).toEqual([
			{ type: 'cellIs', priority: 3, range: 'B1:B10' },
		])
	})

	test('journal inverse ops restore defined name edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' }])

		const changed = wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!B1:B10' }], {
			journal: true,
		})

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' },
		])

		const undoChange = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undoChange.errors).toEqual([])
		expect(wb.definedName('Budget')?.formula).toBe('Sheet1!A1:A10')

		const deleted = wb.apply([{ op: 'deleteDefinedName', name: 'Budget' }], { journal: true })
		expect(deleted.journal?.inverseOps).toEqual([
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' },
		])

		const undoDelete = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDelete.errors).toEqual([])
		expect(wb.definedName('Budget')?.formula).toBe('Sheet1!A1:A10')

		const created = wb.apply([{ op: 'setDefinedName', name: 'Scratch', ref: 'Sheet1!C1:C2' }], {
			journal: true,
		})
		expect(created.journal?.inverseOps).toEqual([{ op: 'deleteDefinedName', name: 'Scratch' }])
	})

	test('journal inverse ops restore table renames and column metadata edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C2', value: 5 },
					{ ref: 'A3', value: 'Total' },
					{ ref: 'B3', value: 2 },
					{ ref: 'C3', value: 5 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C3', name: 'Sales', hasHeaders: true },
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Amount',
				formula: '=[@Qty]*10',
				totalsRowFunction: 'sum',
			},
		])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<{
					tables: Array<{
						hasTotals?: boolean
						ref: { start: { row: number; col: number }; end: { row: number; col: number } }
					}>
				}>
			}
		}
		const tableModel = internal.wb.sheets[0]?.tables[0]
		if (tableModel) {
			tableModel.hasTotals = true
			tableModel.ref = { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
		}

		const columnEdit = wb.apply(
			[
				{
					op: 'setTableColumn',
					table: 'Sales',
					column: 'Amount',
					newName: 'Revenue',
					formula: '=[@Qty]*12',
					totalsRowFunction: 'average',
				},
			],
			{ journal: true },
		)

		expect(columnEdit.journal?.exact).toBe(true)
		expect(columnEdit.journal?.inverseOps).toEqual([
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Revenue',
				newName: 'Amount',
				formula: '[@Qty]*10',
				totalsRowFunction: 'sum',
			},
		])

		const undoColumn = wb.apply(columnEdit.journal?.inverseOps ?? [], { transaction: true })
		expect(undoColumn.errors).toEqual([])
		expect(wb.table('Sales')?.columns).toEqual(['Region', 'Qty', 'Amount'])
		expect(wb.table('Sales')?.columnDefs[2]).toMatchObject({
			name: 'Amount',
			formula: '[@Qty]*10',
			totalsRowFunction: 'sum',
		})

		const renamed = wb.apply([{ op: 'renameTable', table: 'Sales', newName: 'RevenueTable' }], {
			journal: true,
		})
		expect(renamed.journal?.inverseOps).toEqual([
			{ op: 'renameTable', table: 'RevenueTable', newName: 'Sales' },
		])

		const undoRename = wb.apply(renamed.journal?.inverseOps ?? [], { transaction: true })
		expect(undoRename.errors).toEqual([])
		expect(wb.table('Sales')?.name).toBe('Sales')
	})

	test('journal inverse ops restore table lifecycle metadata edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Product' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'Widget' },
					{ ref: 'B2', value: 2 },
					{ ref: 'A3', value: 'Gadget' },
					{ ref: 'B3', value: 3 },
				],
			},
		])

		const created = wb.apply(
			[{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Sales', hasHeaders: true }],
			{ journal: true },
		)
		expect(created.journal?.exact).toBe(true)
		expect(created.journal?.inverseOps).toEqual([{ op: 'deleteTable', table: 'Sales' }])

		wb.apply([
			{ op: 'setTableColumn', table: 'Sales', column: 'Qty', formula: '=1+1' },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
		])

		const resized = wb.apply([{ op: 'resizeTable', table: 'Sales', ref: 'A1:B2' }], {
			journal: true,
		})
		expect(resized.journal?.exact).toBe(true)
		expect(resized.journal?.inverseOps).toEqual([
			{ op: 'resizeTable', table: 'Sales', ref: 'A1:B3' },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 1, formula: '1+1' },
		])

		const undoResize = wb.apply(resized.journal?.inverseOps ?? [], { transaction: true })
		expect(undoResize.errors).toEqual([])
		expect(wb.table('Sales')?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 1 },
		})
		expect(wb.table('Sales')?.columnDefs[1]?.formula).toBe('1+1')

		const deleted = wb.apply([{ op: 'deleteTable', table: 'Sales' }], { journal: true })
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.inverseOps).toEqual([
			{
				op: 'createTable',
				sheet: 'Sheet1',
				ref: 'A1:B3',
				name: 'Sales',
				hasHeaders: true,
			},
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 0, newName: 'Product' },
			{ op: 'setTableColumn', table: 'Sales', column: 1, newName: 'Qty', formula: '1+1' },
		])

		const undoDelete = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDelete.errors).toEqual([])
		expect(wb.table('Sales')?.columns).toEqual(['Product', 'Qty'])
		expect(wb.table('Sales')?.columnDefs[1]?.formula).toBe('1+1')
		expect(wb.table('Sales')?.styleInfo).toEqual({ name: 'TableStyleMedium2' })
	})
})
