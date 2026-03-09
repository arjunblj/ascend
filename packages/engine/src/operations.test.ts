import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import { applyOperation, applyOperations } from './operations.ts'

const sid = 0 as StyleId

function cell(value: ReturnType<typeof numberValue>, formula: string | null = null) {
	return { value, formula, styleId: sid }
}

function setup() {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	sheet.cells.set(0, 0, cell(numberValue(10)))
	sheet.cells.set(1, 0, cell(numberValue(20)))
	sheet.cells.set(2, 0, cell(numberValue(30)))
	sheet.cells.set(0, 1, cell(stringValue('hello')))
	return wb
}

describe('applyOperation', () => {
	test('setCells sets values on existing sheet', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 99 },
				{ ref: 'C1', value: 'new' },
			],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.affectedCells).toEqual(['A1', 'C1'])
		expect(result.value.recalcRequired).toBe(true)

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.cells.get(0, 0)?.value).toEqual(numberValue(99))
		expect(s.cells.get(0, 2)?.value).toEqual(stringValue('new'))
	})

	test('setFormula sets formula on cell', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'A1',
			formula: 'SUM(A2:A3)',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.affectedCells).toEqual(['A1'])
		const c = wb.getSheet('Sheet1')?.cells.get(0, 0)
		expect(c?.formula).toBe('SUM(A2:A3)')
		expect(c?.value).toEqual(numberValue(10))
	})

	test('fillFormula translates references across a range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(1)))
		sheet.cells.set(1, 0, cell(numberValue(2)))

		const result = applyOperation(wb, {
			op: 'fillFormula',
			sheet: 'Sheet1',
			range: 'B1:B2',
			formula: '=A1*2',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.cells.get(0, 1)?.formula).toBe('A1*2')
		expect(sheet.cells.get(1, 1)?.formula).toBe('A2*2')
	})

	test('addSheet creates a new sheet', () => {
		const wb = setup()
		const result = applyOperation(wb, { op: 'addSheet', name: 'Sheet2' })
		expect(result.ok).toBe(true)
		expect(wb.getSheet('Sheet2')).toBeDefined()
		expect(wb.sheets).toHaveLength(2)
	})

	test('addSheet rejects duplicate name', () => {
		const wb = setup()
		const result = applyOperation(wb, { op: 'addSheet', name: 'Sheet1' })
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect(result.error.code).toBe('NAME_CONFLICT')
	})

	test('deleteSheet removes sheet', () => {
		const wb = setup()
		wb.addSheet('Sheet2')
		const result = applyOperation(wb, { op: 'deleteSheet', sheet: 'Sheet2' })
		expect(result.ok).toBe(true)
		expect(wb.getSheet('Sheet2')).toBeUndefined()
		expect(wb.sheets).toHaveLength(1)
	})

	test('insertRows shifts cells down', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'insertRows',
			sheet: 'Sheet1',
			at: 1,
			count: 2,
		})
		expect(result.ok).toBe(true)

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(s.cells.get(1, 0)).toBeUndefined()
		expect(s.cells.get(2, 0)).toBeUndefined()
		expect(s.cells.get(3, 0)?.value).toEqual(numberValue(20))
		expect(s.cells.get(4, 0)?.value).toEqual(numberValue(30))
	})

	test('deleteRows shifts cells up', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'deleteRows',
			sheet: 'Sheet1',
			at: 0,
			count: 1,
		})
		expect(result.ok).toBe(true)

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.cells.get(0, 0)?.value).toEqual(numberValue(20))
		expect(s.cells.get(1, 0)?.value).toEqual(numberValue(30))
		expect(s.cells.get(2, 0)).toBeUndefined()
	})

	test('insertRows rewrites formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(1, 0, cell(numberValue(2)))
		s.cells.set(2, 0, cell(EMPTY, 'SUM(A1:A2)'))

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 })

		const formulaCell = s.cells.get(3, 0)
		expect(formulaCell?.formula).toBe('SUM(A1:A3)')
	})

	test('renameSheet updates sheet name', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'renameSheet',
			sheet: 'Sheet1',
			newName: 'Data',
		})
		expect(result.ok).toBe(true)
		expect(wb.getSheet('Data')).toBeDefined()
		expect(wb.getSheet('Sheet1')).toBeUndefined()
	})

	test('clearRange removes cell data', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A1:A3',
			what: 'all',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.cells.get(0, 0)).toBeUndefined()
		expect(s.cells.get(1, 0)).toBeUndefined()
		expect(s.cells.get(2, 0)).toBeUndefined()
		expect(s.cells.get(0, 1)?.value).toEqual(stringValue('hello'))
	})

	test('operation on non-existent sheet returns error', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'NoSuchSheet',
			updates: [{ ref: 'A1', value: 1 }],
		})
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect(result.error.code).toBe('SHEET_NOT_FOUND')
	})

	test('mergeCells adds merge to sheet', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'mergeCells',
			sheet: 'Sheet1',
			range: 'A1:B2',
		})
		expect(result.ok).toBe(true)

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.merges).toHaveLength(1)
		expect(s.merges[0]?.start).toEqual({ row: 0, col: 0 })
		expect(s.merges[0]?.end).toEqual({ row: 1, col: 1 })
	})

	test('setDefinedName stores a named range', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setDefinedName',
			name: 'MyRange',
			ref: 'Sheet1!A1:A3',
		})
		expect(result.ok).toBe(true)
		expect(wb.definedNames.get('MyRange')).toBe('Sheet1!A1:A3')
	})

	test('setDefinedName can target a sheet scope', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setDefinedName',
			name: 'Budget',
			ref: 'Sheet1!A1',
			scope: 'Sheet1',
		})
		expect(result.ok).toBe(true)
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		expect(wb.definedNames.resolve('Budget', sheet.id)?.scope.kind).toBe('sheet')
	})

	test('setHyperlink stores hyperlink metadata on the sheet', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'B1',
			url: 'https://example.com/report',
			display: 'Report',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(wb.getSheet('Sheet1')?.hyperlinks.get('B1')).toEqual({
			target: 'https://example.com/report',
			display: 'Report',
		})
	})

	test('setNumberFormat applies styles across a range', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setNumberFormat',
			sheet: 'Sheet1',
			range: 'A1:A2',
			format: '0.00%',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		const styleA1 = wb.styles.get(sheet.cells.get(0, 0)?.styleId ?? sid)
		const styleA2 = wb.styles.get(sheet.cells.get(1, 0)?.styleId ?? sid)
		expect(styleA1?.numberFormat).toBe('0.00%')
		expect(styleA2?.numberFormat).toBe('0.00%')
	})

	test('sortRange sorts a block by header name and moves metadata with rows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('B'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('A'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.hyperlinks.set('A2', { target: 'https://example.com/b' })
		sheet.comments.set('B3', { text: 'lowest' })

		const result = applyOperation(wb, {
			op: 'sortRange',
			sheet: 'Sheet1',
			range: 'A1:B3',
			by: [{ column: 'Score' }],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('A'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('B'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(2))
		expect(sheet.hyperlinks.get('A3')).toEqual({ target: 'https://example.com/b' })
		expect(sheet.comments.get('B2')).toEqual({ text: 'lowest' })
	})
})

describe('applyOperations', () => {
	test('applies multiple operations in sequence', () => {
		const wb = createWorkbook()
		const result = applyOperations(wb, [
			{ op: 'addSheet', name: 'Sheet1' },
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A3', formula: 'SUM(A1:A2)' },
		])
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(s.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(s.cells.get(2, 0)?.formula).toBe('SUM(A1:A2)')
		expect(result.value.recalcRequired).toBe(true)
	})

	test('stops on first error', () => {
		const wb = createWorkbook()
		const result = applyOperations(wb, [
			{ op: 'addSheet', name: 'Sheet1' },
			{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
			{ op: 'addSheet', name: 'Sheet2' },
		])
		expect(result.ok).toBe(false)
		expect(wb.getSheet('Sheet2')).toBeUndefined()
	})
})
