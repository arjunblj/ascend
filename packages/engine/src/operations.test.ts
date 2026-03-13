import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, createWorkbook } from '@ascend/core'
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

	test('setCells serializes Date inputs using workbook date system', () => {
		const wb = setup()
		wb.calcSettings = { ...wb.calcSettings, dateSystem: '1904' }
		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'D1', value: new Date(Date.UTC(1904, 0, 2)) }],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(wb.getSheet('Sheet1')?.cells.get(0, 3)?.value).toEqual({
			kind: 'date',
			serial: 1,
		})
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

	test('setRichText writes rich text runs to a cell', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setRichText',
			sheet: 'Sheet1',
			ref: 'B2',
			runs: [
				{ text: 'Hello', bold: true },
				{ text: ' World', italic: true },
			],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(wb.getSheet('Sheet1')?.cells.get(1, 1)?.value).toEqual({
			kind: 'richText',
			runs: [
				{ text: 'Hello', bold: true },
				{ text: ' World', italic: true },
			],
		})
	})

	test('setConditionalFormat stores conditional formatting rules', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A3',
			rule: {
				type: 'cellIs',
				operator: 'greaterThan',
				formula: '10',
				priority: 1,
			},
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(wb.getSheet('Sheet1')?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A3',
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						formulas: ['10'],
						priority: 1,
					},
				],
			},
		])
	})

	test('setPageSetup and setPrintArea write print metadata', () => {
		const wb = setup()
		const result1 = applyOperation(wb, {
			op: 'setPageSetup',
			sheet: 'Sheet1',
			setup: {
				orientation: 'landscape',
				scale: 80,
				margins: { left: 0.5, right: 0.5 },
			},
		})
		expect(result1.ok).toBe(true)
		if (!result1.ok) return

		const result2 = applyOperation(wb, {
			op: 'setPrintArea',
			sheet: 'Sheet1',
			range: 'A1:B5',
		})
		expect(result2.ok).toBe(true)
		if (!result2.ok) return

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.pageSetup).toEqual({ orientation: 'landscape', scale: 80 })
		expect(sheet?.pageMargins).toEqual({ left: 0.5, right: 0.5 })
		expect(wb.definedNames.resolve('_xlnm.Print_Area', sheet?.id, sheet?.id)?.formula).toBe(
			"'Sheet1'!A1:B5",
		)
	})

	test('setDataValidation stores validation metadata', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'A1:A3',
			rule: {
				type: 'list',
				formula1: '"Yes,No"',
				allowBlank: false,
			},
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(wb.getSheet('Sheet1')?.dataValidations).toEqual([
			{
				sqref: 'A1:A3',
				type: 'list',
				formula1: '"Yes,No"',
				allowBlank: false,
				showErrorMessage: true,
			},
		])
	})

	test('copyRange copies values and translates relative formulas', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) return
		sheet.cells.set(0, 2, { value: numberValue(0), formula: 'A1+B1', styleId: sid })

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1:C1',
			target: 'A3',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('hello'))
		expect(sheet.cells.get(2, 2)?.formula).toBe('A3+B3')
	})

	test('moveRange relocates source cells and clears original range', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) return

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:A2',
			target: 'C1',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(0, 0)).toBeUndefined()
		expect(sheet.cells.get(1, 0)).toBeUndefined()
	})

	test('hideSheet and hideCols update sheet visibility metadata', () => {
		const wb = setup()
		const result1 = applyOperation(wb, {
			op: 'hideSheet',
			sheet: 'Sheet1',
			hidden: true,
		})
		expect(result1.ok).toBe(true)
		if (!result1.ok) return

		const result2 = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: true,
		})
		expect(result2.ok).toBe(true)
		if (!result2.ok) return

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.state).toBe('hidden')
		expect(sheet?.colDefs).toContainEqual({ min: 2, max: 2, hidden: true })
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

	test('deleteSheet removes sheet-scoped names and pivot metadata for the deleted sheet', () => {
		const wb = setup()
		const sheet2 = wb.addSheet('Sheet2')
		wb.definedNames.set('LocalBudget', 'Sheet2!A1', { kind: 'sheet', sheetId: sheet2.id })
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet2',
			name: 'PivotTable1',
			cacheId: 4,
			locationRef: 'A1',
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_PivotTable1',
			pivotTableNames: ['PivotTable1'],
		})

		const result = applyOperation(wb, { op: 'deleteSheet', sheet: 'Sheet2' })
		expect(result.ok).toBe(true)
		expect(wb.definedNames.list().some((entry) => entry.name === 'LocalBudget')).toBe(false)
		expect(wb.pivotTables).toHaveLength(0)
		expect(wb.slicerCaches).toHaveLength(0)
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

	test('deleteRows shrinks overlapping table, filter, and validation ranges', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		s.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		s.cells.set(2, 0, { value: stringValue('Debt'), formula: null, styleId: sid })
		s.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: sid })
		s.cells.set(3, 0, { value: stringValue('Equity'), formula: null, styleId: sid })
		s.cells.set(3, 1, { value: numberValue(30), formula: null, styleId: sid })
		s.autoFilter = { ref: 'A1:B4', columns: [] }
		s.dataValidations.push({ sqref: 'A2:B4', type: 'list', formula1: 'A2' })
		s.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: { ref: 'A1:B4', columns: [] },
		})

		applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 })

		expect(s.autoFilter?.ref).toBe('A1:B3')
		expect(s.dataValidations[0]?.sqref).toBe('A2:B3')
		expect(s.tables[0]?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 1 },
		})
		expect(s.tables[0]?.autoFilter?.ref).toBe('A1:B3')
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

	test('insertRows rewrites whole-row references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(1, 0, cell(numberValue(2)))
		s.cells.set(2, 0, cell(EMPTY, 'SUM(1:2)'))

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 })

		expect(s.cells.get(3, 0)?.formula).toBe('SUM(1:3)')
	})

	test('insertRows shifts comments, hyperlinks, validations, ignored errors, and row heights', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.comments.set('A2', { text: 'note' })
		s.hyperlinks.set('B2', { target: 'https://example.com', location: 'Sheet1!A2' })
		s.dataValidations.push({ sqref: 'A2:B2', type: 'list', formula1: 'A2' })
		s.conditionalFormats.push({
			sqref: 'A2',
			rules: [{ type: 'expression', formulas: ['A2>0'] }],
		})
		s.ignoredErrors.push({ sqref: 'A2', formula: true })
		s.rowHeights.set(1, 24)

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 2 })

		expect(s.comments.get('A4')).toEqual({ text: 'note' })
		expect(s.hyperlinks.get('B4')).toEqual({
			target: 'https://example.com',
			location: 'Sheet1!A4',
		})
		expect(s.dataValidations[0]?.sqref).toBe('A4:B4')
		expect(s.dataValidations[0]?.formula1).toBe('A4')
		expect(s.conditionalFormats[0]?.sqref).toBe('A4')
		expect(s.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('A4>0')
		expect(s.ignoredErrors[0]?.sqref).toBe('A4')
		expect(s.rowHeights.get(3)).toBe(24)
	})

	test('insertCols shifts tables, filters, comments, and hyperlinks', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		s.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		s.comments.set('A1', { text: 'header' })
		s.hyperlinks.set('B2', { target: 'https://example.com/value' })
		s.autoFilter = { ref: 'A1:B2', columns: [] }
		s.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: { ref: 'A1:B2', columns: [] },
		})

		applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 0, count: 1 })

		expect(s.comments.get('B1')).toEqual({ text: 'header' })
		expect(s.hyperlinks.get('C2')).toEqual({ target: 'https://example.com/value' })
		expect(s.autoFilter?.ref).toBe('B1:C2')
		expect(s.tables[0]?.ref).toEqual({
			start: { row: 0, col: 1 },
			end: { row: 1, col: 2 },
		})
		expect(s.tables[0]?.autoFilter?.ref).toBe('B1:C2')
	})

	test('deleteCols rewrites formulas and whole-column references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(0, 1, cell(numberValue(2)))
		s.cells.set(0, 2, cell(numberValue(3)))
		s.cells.set(1, 0, cell(EMPTY, 'SUM(A1:C1)'))
		s.cells.set(2, 0, cell(EMPTY, 'SUM(A:C)'))

		applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expect(s.cells.get(1, 0)?.formula).toBe('SUM(A1:B1)')
		expect(s.cells.get(2, 0)?.formula).toBe('SUM(A:B)')
	})

	test('renameSheet updates sheet name', () => {
		const wb = setup()
		wb.definedNames.set('Budget', 'Sheet1!A1')
		const result = applyOperation(wb, {
			op: 'renameSheet',
			sheet: 'Sheet1',
			newName: 'Data',
		})
		expect(result.ok).toBe(true)
		expect(wb.getSheet('Data')).toBeDefined()
		expect(wb.getSheet('Sheet1')).toBeUndefined()
		expect(wb.definedNames.get('Budget')).toBe('Data!A1')
	})

	test('renameSheet updates whole-column references in formulas and defined names', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(EMPTY, 'SUM(Sheet1!A:A)'))
		wb.definedNames.set('AllA', 'Sheet1!A:A')

		applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet1', newName: 'Data' })

		expect(s.cells.get(0, 0)?.formula).toBe('SUM(Data!A:A)')
		expect(wb.definedNames.get('AllA')).toBe('Data!A:A')
	})

	test('renameSheet updates 3D sheet-span endpoints in formulas', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		wb.addSheet('Sheet2')
		wb.addSheet('Sheet3')
		s1.cells.set(0, 0, cell(EMPTY, 'SUM(Sheet1:Sheet3!A1)'))

		applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet3', newName: 'Summary' })

		expect(s1.cells.get(0, 0)?.formula).toBe('SUM(Sheet1:Summary!A1)')
	})

	test('renameSheet updates validation, conditional-format, and table formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.dataValidations.push({ sqref: 'A1', type: 'list', formula1: 'Sheet1!A:A' })
		s.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'expression', formulas: ['SUM(Sheet1!A:A)>0'] }],
		})
		s.hyperlinks.set('A1', { location: 'Sheet1!A1', display: 'jump' })
		s.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [
				{ name: 'Name', formula: 'Sheet1!A:A', totalsRowFormula: 'SUM(Sheet1!A:A)' },
				{ name: 'Value' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet1', newName: 'Data' })

		expect(s.dataValidations[0]?.formula1).toBe('Data!A:A')
		expect(s.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('SUM(Data!A:A)>0')
		expect(s.hyperlinks.get('A1')?.location).toBe('Data!A1')
		expect(s.tables[0]?.columns[0]?.formula).toBe('Data!A:A')
		expect(s.tables[0]?.columns[0]?.totalsRowFormula).toBe('SUM(Data!A:A)')
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

	test('deleteDefinedName can target a sheet scope without removing workbook scope', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('expected sheet')
		wb.definedNames.set('Budget', '10')
		wb.definedNames.set('Budget', '20', { kind: 'sheet', sheetId: sheet.id })

		const result = applyOperation(wb, { op: 'deleteDefinedName', name: 'Budget', scope: 'Sheet1' })
		expect(result.ok).toBe(true)
		expect(wb.definedNames.get('Budget')).toBe('10')
		expect(wb.definedNames.resolve('Budget', sheet.id)?.formula).toBe('10')
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
		sheet.dataValidations.push({ sqref: 'A2:B2', type: 'list', formula1: '"A,B"' })
		sheet.conditionalFormats.push({ sqref: 'B3', rules: [] })
		sheet.ignoredErrors.push({ sqref: 'A2', formula: true })

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
		expect(sheet.dataValidations[0]?.sqref).toBe('A3:B3')
		expect(sheet.conditionalFormats[0]?.sqref).toBe('B2')
		expect(sheet.ignoredErrors[0]?.sqref).toBe('A3')
	})

	test('sortRange only clears formula metadata on the affected sheet', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Source')
		const other = wb.addSheet('Other')
		source.cells.set(0, 0, { value: stringValue('b'), formula: null, styleId: sid })
		source.cells.set(1, 0, { value: stringValue('a'), formula: null, styleId: sid })
		other.cells.set(0, 0, {
			value: numberValue(3),
			formula: 'SUM(B1:B2)',
			styleId: sid,
			formulaInfo: { kind: 'array', ref: 'A1:A2' },
		})

		const result = applyOperation(wb, {
			op: 'sortRange',
			sheet: 'Source',
			range: 'A1:A2',
			by: [{ column: 'A' }],
		})
		expect(result.ok).toBe(true)
		expect(other.cells.get(0, 0)?.formulaInfo).toEqual({ kind: 'array', ref: 'A1:A2' })
	})

	test('createTable infers columns from the header row', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })

		const result = applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'BalanceTable',
			hasHeaders: true,
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.tables).toHaveLength(1)
		expect(sheet.tables[0]?.columns).toEqual([{ name: 'Name' }, { name: 'Value' }])
		expect(sheet.autoFilter).toEqual({
			ref: 'A1:B2',
			columns: [],
		})
	})

	test('appendRows expands a table and writes new values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'BalanceTable',
			hasHeaders: true,
		})

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'BalanceTable',
			rows: [['Debt', 20]],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.tables[0]?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 1 },
		})
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('Debt'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(20))
	})

	test('appendRows serializes Date values using workbook date system', () => {
		const wb = createWorkbook()
		wb.calcSettings = { ...wb.calcSettings, dateSystem: '1904' }
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Date'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A1',
			name: 'DateTable',
			hasHeaders: true,
		})

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'DateTable',
			rows: [[new Date(Date.UTC(1904, 0, 2))]],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.cells.get(1, 0)?.value).toEqual({
			kind: 'date',
			serial: 1,
		})
	})

	test('appendRows expands table filter and sort metadata refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'BalanceTable',
			hasHeaders: true,
		})
		const table = sheet.tables[0]
		if (!table) throw new Error('expected table')
		sheet.tables[0] = {
			...table,
			autoFilter: {
				ref: 'A1:B2',
				columns: [],
				sortState: { ref: 'A1:B2', conditions: [{ ref: 'B2:B2' }] },
			},
			sortState: { ref: 'A1:B2', conditions: [{ ref: 'A1:A2' }] },
		}

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'BalanceTable',
			rows: [['Debt', 20]],
		})
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(sheet.tables[0]?.autoFilter?.ref).toBe('A1:B3')
		expect(sheet.tables[0]?.autoFilter?.sortState?.ref).toBe('A1:B3')
		expect(sheet.tables[0]?.sortState?.ref).toBe('A1:B3')
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
