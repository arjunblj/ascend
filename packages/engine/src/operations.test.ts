import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, createWorkbook } from '@ascend/core'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import { recalculate } from './calc.ts'
import { defaultCalcContext } from './calc-context.ts'
import { applyOperation, applyOperations, applyWithTransaction } from './operations.ts'

const sid = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function expectErr<T, E>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is {
	ok: false
	error: E
} {
	expect(result.ok).toBe(false)
	if (result.ok) throw new Error('Expected operation to fail')
}

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
		expectOk(result)

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
		expectOk(result)

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
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1'])
		const c = wb.getSheet('Sheet1')?.cells.get(0, 0)
		expect(c?.formula).toBe('SUM(A2:A3)')
		expect(c?.value).toEqual(numberValue(10))
	})

	test('setCells replaces formula content with a literal value', () => {
		const wb = setup()
		expectOk(
			applyOperation(wb, {
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A1',
				formula: 'A2+A3',
			}),
		)
		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A1', value: 42 }],
		})
		expectOk(result)

		const c = wb.getSheet('Sheet1')?.cells.get(0, 0)
		expect(c?.value).toEqual(numberValue(42))
		expect(c?.formula).toBeNull()
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
		expectOk(result)

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
		expectOk(result)

		expect(wb.getSheet('Sheet1')?.cells.get(1, 1)?.value).toEqual({
			kind: 'richText',
			runs: [
				{ text: 'Hello', bold: true },
				{ text: ' World', italic: true },
			],
		})
	})

	test('replaceImage swaps media bytes while preserving anchor metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
			anchor: {
				kind: 'oneCell',
				from: { row: 1, col: 1 },
				cx: 320000,
				cy: 240000,
			},
			name: 'Logo',
			description: 'Brand logo',
		})

		const result = applyOperation(wb, {
			op: 'replaceImage',
			sheet: 'Sheet1',
			name: 'Logo',
			contentBase64: 'BAUG',
			contentType: 'image/png',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual([])
		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(sheet.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			name: 'Logo',
			description: 'Brand logo',
			anchor: {
				kind: 'oneCell',
				from: { row: 1, col: 1 },
			},
		})
		expect(Array.from(sheet.imageRefs[0]?.content ?? [])).toEqual([4, 5, 6])
	})

	test('setDrawingText updates a selected text-bearing drawing object', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 2,
			name: 'Callout',
			text: 'Revenue up',
			anchor: {
				kind: 'twoCell',
				from: { row: 4, col: 1 },
				to: { row: 6, col: 4 },
			},
		})

		const result = applyOperation(wb, {
			op: 'setDrawingText',
			sheet: 'Sheet1',
			drawingPartPath: 'xl/drawings/drawing1.xml',
			id: 2,
			text: 'Revenue flat',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual([])
		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(sheet.drawingObjectRefs[0]).toMatchObject({
			id: 2,
			name: 'Callout',
			text: 'Revenue flat',
			anchor: {
				kind: 'twoCell',
				from: { row: 4, col: 1 },
				to: { row: 6, col: 4 },
			},
		})
	})

	test('setDrawingText validates selectors and text-bearing objects', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.drawingObjectRefs.push(
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				kind: 'textBox',
				id: 2,
				name: 'Duplicate',
				text: 'First',
			},
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				kind: 'shape',
				id: 3,
				name: 'Duplicate',
			},
		)

		const missingSelector = applyOperation(wb, {
			op: 'setDrawingText',
			sheet: 'Sheet1',
			text: 'Updated',
		})
		expectErr(missingSelector)
		expect(missingSelector.error.message).toContain('requires drawingPartPath')

		const ambiguous = applyOperation(wb, {
			op: 'setDrawingText',
			sheet: 'Sheet1',
			name: 'Duplicate',
			text: 'Updated',
		})
		expectErr(ambiguous)
		expect(ambiguous.error.message).toContain('matched 2 drawing objects')

		const noText = applyOperation(wb, {
			op: 'setDrawingText',
			sheet: 'Sheet1',
			id: 3,
			text: 'Updated',
		})
		expectErr(noText)
		expect(noText.error.message).toContain('no editable text body')
	})

	test('setThreadedComment updates existing text while preserving thread metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.threadedComments.push(
			{
				ref: 'A1',
				text: 'Please review',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc1',
				personId: '0',
				author: 'Ada Lovelace',
				dateTime: '2024-01-01T00:00:00.000',
			},
			{
				ref: 'A1',
				text: 'Reviewed',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc2',
				parentId: 'tc1',
				personId: '1',
				author: 'Grace Hopper',
				dateTime: '2024-01-02T00:00:00.000',
				done: true,
			},
		)

		const result = applyOperation(wb, {
			op: 'setThreadedComment',
			sheet: 'Sheet1',
			threadedCommentId: 'tc2',
			text: 'Reviewed and approved',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['Sheet1!A1'])
		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(sheet.threadedComments[1]).toEqual({
			ref: 'A1',
			text: 'Reviewed and approved',
			partPath: 'xl/threadedComments/threadedComment1.xml',
			id: 'tc2',
			parentId: 'tc1',
			personId: '1',
			author: 'Grace Hopper',
			dateTime: '2024-01-02T00:00:00.000',
			done: true,
		})
	})

	test('setThreadedComment validates selectors and ambiguous refs', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.threadedComments.push(
			{
				ref: 'A1',
				text: 'First',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc1',
			},
			{
				ref: 'A1',
				text: 'Second',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc2',
				parentId: 'tc1',
			},
		)

		const missingSelector = applyOperation(wb, {
			op: 'setThreadedComment',
			sheet: 'Sheet1',
			text: 'Updated',
		})
		expectErr(missingSelector)
		expect(missingSelector.error.message).toContain('requires partPath')

		const ambiguous = applyOperation(wb, {
			op: 'setThreadedComment',
			sheet: 'Sheet1',
			ref: 'A1',
			text: 'Updated',
		})
		expectErr(ambiguous)
		expect(ambiguous.error.message).toContain('matched 2 comments')

		const badIndex = applyOperation(wb, {
			op: 'setThreadedComment',
			sheet: 'Sheet1',
			commentIndex: -1,
			text: 'Updated',
		})
		expectErr(badIndex)
		expect(badIndex.error.message).toContain('commentIndex')
	})

	test('insertImage allocates image identity and anchor metadata', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'insertImage',
			sheet: 'Sheet1',
			contentBase64: 'BAUG',
			contentType: 'image/png',
			name: 'Logo',
			description: 'Brand logo',
			anchor: { kind: 'oneCell', from: { row: 1, col: 1 }, cx: 320000, cy: 240000 },
		})
		expectOk(result)

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.drawingRefs.hasDrawing).toBe(true)
		expect(sheet?.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			name: 'Logo',
			description: 'Brand logo',
			anchor: { kind: 'oneCell', from: { row: 1, col: 1 }, cx: 320000, cy: 240000 },
		})
		expect(Array.from(sheet?.imageRefs[0]?.content ?? [])).toEqual([4, 5, 6])
	})

	test('deleteImage removes a selected image ref', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: false }
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
			name: 'Logo',
		})

		const result = applyOperation(wb, {
			op: 'deleteImage',
			sheet: 'Sheet1',
			name: 'Logo',
		})
		expectOk(result)

		expect(sheet.imageRefs).toEqual([])
		expect(sheet.drawingRefs.hasDrawing).toBe(false)
	})

	test('setChartSeriesSource updates parsed chart source refs', () => {
		const wb = setup()
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'barChart',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
			],
		})

		const result = applyOperation(wb, {
			op: 'setChartSeriesSource',
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
			categoryRef: 'Sheet1!$A$2:$A$10',
			valueRef: 'Sheet1!$C$2:$C$10',
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(wb.chartParts[0]?.series[0]).toMatchObject({
			nameRef: 'Sheet1!$B$1',
			categoryRef: 'Sheet1!$A$2:$A$10',
			valueRef: 'Sheet1!$C$2:$C$10',
		})
	})

	test('setSparklineGroup updates source ranges and display flags', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.sparklineGroups.push({
			groupIndex: 0,
			type: 'line',
			markers: true,
			highPoint: true,
			displayXAxis: true,
			range: 'Sheet1!B2:B4',
			locationRange: 'D2:D4',
			count: 1,
		})

		const result = applyOperation(wb, {
			op: 'setSparklineGroup',
			sheet: 'Sheet1',
			groupIndex: 0,
			range: 'Sheet1!C2:C4',
			locationRange: 'E2:E4',
			type: 'column',
			markers: false,
			highPoint: false,
			displayXAxis: false,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(sheet.sparklineGroups[0]).toMatchObject({
			type: 'column',
			markers: false,
			highPoint: false,
			displayXAxis: false,
			range: 'Sheet1!C2:C4',
			locationRange: 'E2:E4',
		})
	})

	test('setSparklineGroup validates existing groups and editable fields', () => {
		const wb = setup()
		expect(
			applyOperation(wb, {
				op: 'setSparklineGroup',
				sheet: 'Sheet1',
				groupIndex: 0,
				range: 'Sheet1!C2:C4',
			}).ok,
		).toBe(false)
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.sparklineGroups.push({ groupIndex: 0, count: 1 })
		expect(
			applyOperation(wb, {
				op: 'setSparklineGroup',
				sheet: 'Sheet1',
				groupIndex: -1,
				range: 'Sheet1!C2:C4',
			}).ok,
		).toBe(false)
		expect(applyOperation(wb, { op: 'setSparklineGroup', sheet: 'Sheet1', groupIndex: 0 }).ok).toBe(
			false,
		)
	})

	test('setAdvancedFilter updates custom sheet view criteria and sort metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.advancedFilters.push({
			viewName: 'WestOnly',
			guid: '{11111111-1111-1111-1111-111111111111}',
			ref: 'A1:C20',
			filterColumnCount: 1,
			sortConditionCount: 1,
			autoFilter: {
				ref: 'A1:C20',
				columns: [{ colId: 0, kind: 'filters', values: ['West'] }],
				sortState: {
					ref: 'A2:C20',
					conditions: [{ ref: 'C2:C20', descending: true }],
				},
			},
		})

		const result = applyOperation(wb, {
			op: 'setAdvancedFilter',
			sheet: 'Sheet1',
			filterIndex: 0,
			range: 'A1:D20',
			column: 1,
			values: ['East', 'North'],
			sortRef: 'A2:D20',
			sortBy: 'B2:B20',
			descending: false,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(sheet.advancedFilters[0]).toMatchObject({
			viewName: 'WestOnly',
			guid: '{11111111-1111-1111-1111-111111111111}',
			ref: 'A1:D20',
			filterColumnCount: 2,
			sortConditionCount: 1,
			autoFilter: {
				ref: 'A1:D20',
				columns: [
					{ colId: 0, kind: 'filters', values: ['West'] },
					{ colId: 1, kind: 'filters', values: ['East', 'North'] },
				],
				sortState: {
					ref: 'A2:D20',
					conditions: [{ ref: 'B2:B20', descending: false }],
				},
			},
		})
	})

	test('setAdvancedFilter validates selectors and update fields', () => {
		const wb = setup()
		expect(
			applyOperation(wb, {
				op: 'setAdvancedFilter',
				sheet: 'Sheet1',
				filterIndex: 0,
				column: 0,
				values: ['East'],
			}).ok,
		).toBe(false)

		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.advancedFilters.push({
			ref: 'A1:C20',
			filterColumnCount: 0,
			sortConditionCount: 0,
			autoFilter: { ref: 'A1:C20', columns: [] },
		})
		expect(
			applyOperation(wb, {
				op: 'setAdvancedFilter',
				sheet: 'Sheet1',
				filterIndex: -1,
				column: 0,
				values: ['East'],
			}).ok,
		).toBe(false)
		expect(
			applyOperation(wb, {
				op: 'setAdvancedFilter',
				sheet: 'Sheet1',
				filterIndex: 0,
				values: ['East'],
			}).ok,
		).toBe(false)
		expect(
			applyOperation(wb, {
				op: 'setAdvancedFilter',
				sheet: 'Sheet1',
				filterIndex: 0,
				column: 0,
			}).ok,
		).toBe(false)
		expect(
			applyOperation(wb, { op: 'setAdvancedFilter', sheet: 'Sheet1', filterIndex: 0 }).ok,
		).toBe(false)
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
		expectOk(result)

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

	test('setConditionalFormat can append and reassign rule priorities', () => {
		const wb = setup()
		applyOperation(wb, {
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A3',
			rule: { type: 'expression', formula: 'A1>0', priority: 9 },
		})
		const result = applyOperation(wb, {
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A3',
			mode: 'append',
			reassignPriorities: true,
			rule: { type: 'cellIs', operator: 'lessThan', formula: '100', priority: 4 },
		})
		expectOk(result)

		expect(wb.getSheet('Sheet1')?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A3',
				rules: [
					{ type: 'expression', formulas: ['A1>0'], priority: 1 },
					{
						type: 'cellIs',
						operator: 'lessThan',
						formulas: ['100'],
						priority: 2,
					},
				],
			},
		])
	})

	test('deleteConditionalFormat can remove a single rule by priority', () => {
		const wb = setup()
		applyOperation(wb, {
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A3',
			rule: { type: 'expression', formula: 'A1>0', priority: 1 },
		})
		applyOperation(wb, {
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A3',
			mode: 'append',
			rule: { type: 'expression', formula: 'A1<100', priority: 2 },
		})

		const result = applyOperation(wb, {
			op: 'deleteConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A3',
			priority: 1,
		})
		expectOk(result)

		expect(wb.getSheet('Sheet1')?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A3',
				rules: [{ type: 'expression', formulas: ['A1<100'], priority: 2 }],
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
		expectOk(result1)

		const result2 = applyOperation(wb, {
			op: 'setPrintArea',
			sheet: 'Sheet1',
			range: 'A1:B5',
		})
		expectOk(result2)

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
		expectOk(result)

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
		expectOk(result)

		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('hello'))
		expect(sheet.cells.get(2, 2)?.formula).toBe('A3+B3')
	})

	test('copyRange can paste values without carrying formulas or formats', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		const targetStyle = wb.styles.register({ numberFormat: '0.0%' })
		sheet.cells.set(0, 0, { value: numberValue(12), formula: 'B1*2', styleId: sourceStyle })
		sheet.cells.set(2, 2, { value: numberValue(99), formula: 'Z1', styleId: targetStyle })

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1',
			target: 'C3',
			mode: 'values',
		})
		expectOk(result)

		expect(sheet.cells.get(2, 2)).toEqual({
			value: numberValue(12),
			formula: null,
			styleId: targetStyle,
		})
		expect(result.value.recalcRequired).toBe(true)
	})

	test('copyRange can paste formats without changing values or formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		sheet.cells.set(0, 0, { value: numberValue(12), formula: null, styleId: sourceStyle })
		sheet.cells.set(2, 2, { value: numberValue(99), formula: 'A1+1', styleId: sid })

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1',
			target: 'C3',
			mode: 'formats',
		})
		expectOk(result)

		expect(sheet.cells.get(2, 2)).toEqual({
			value: numberValue(99),
			formula: 'A1+1',
			styleId: sourceStyle,
		})
		expect(result.value.recalcRequired).toBe(false)
	})

	test('copyRange exposes comments, hyperlinks, and validation paste modes', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(1)))
		sheet.cells.set(0, 1, cell(numberValue(2)))
		sheet.comments.set('A1', { text: 'Review', author: 'Ascend' })
		sheet.hyperlinks.set('B1', { target: 'https://example.com', display: 'Example' })
		sheet.dataValidations.push({ sqref: 'A1:B1', type: 'whole', formula1: 'A1' })

		expectOk(
			applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'A1',
				target: 'C1',
				mode: 'comments',
			}),
		)
		expectOk(
			applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'B1',
				target: 'D1',
				mode: 'hyperlinks',
			}),
		)
		expectOk(
			applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'A1:B1',
				target: 'E1',
				mode: 'validations',
			}),
		)

		expect(sheet.comments.get('C1')).toEqual({ text: 'Review', author: 'Ascend' })
		expect(sheet.hyperlinks.get('D1')).toEqual({
			target: 'https://example.com',
			display: 'Example',
		})
		expect(sheet.dataValidations.at(-1)).toEqual({
			sqref: 'E1:F1',
			type: 'whole',
			formula1: 'E1',
		})
		expect(sheet.cells.get(0, 2)).toBeUndefined()
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
		expectOk(result)

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
		expectOk(result1)

		const result2 = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: true,
		})
		expectOk(result2)

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.state).toBe('hidden')
		expect(sheet?.colDefs).toContainEqual({ min: 2, max: 2, hidden: true })
	})

	test('groupRows assigns outline metadata and collapsed boundary row', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'groupRows',
			sheet: 'Sheet1',
			from: 1,
			to: 3,
			collapsed: true,
		})
		expectOk(result)

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.outlinePr).toEqual({ summaryBelow: true })
		expect(sheet?.rowDefs.get(1)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(sheet?.rowDefs.get(3)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(sheet?.rowDefs.get(4)).toEqual({ collapsed: true })
		expect(sheet?.sheetFormatPr?.outlineLevelRow).toBe(1)
	})

	test('groupCols assigns outline metadata and collapsed boundary column', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'groupCols',
			sheet: 'Sheet1',
			from: 0,
			to: 1,
			collapsed: true,
		})
		expectOk(result)

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.outlinePr).toEqual({ summaryRight: true })
		expect(sheet?.colDefs).toContainEqual({ min: 0, max: 0, hidden: true, outlineLevel: 1 })
		expect(sheet?.colDefs).toContainEqual({ min: 1, max: 1, hidden: true, outlineLevel: 1 })
		expect(sheet?.colDefs).toContainEqual({ min: 2, max: 2, collapsed: true })
		expect(sheet?.sheetFormatPr?.outlineLevelCol).toBe(1)
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
		expectErr(result)
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
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
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

	test('insertRows rewrites local spill references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(EMPTY, 'SEQUENCE(2)'))
		s.cells.set(0, 1, cell(EMPTY, 'SUM(A1#)'))

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 0, count: 1 })

		expect(s.cells.get(1, 1)?.formula).toBe('SUM(A2#)')
	})

	test('insertRows rewrites sheet-qualified spill references from other sheets', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.cells.set(0, 0, cell(EMPTY, 'SEQUENCE(2)'))
		s2.cells.set(0, 0, cell(EMPTY, 'SUM(Sheet1!A1#)'))

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 0, count: 1 })

		expect(s2.cells.get(0, 0)?.formula).toBe('SUM(Sheet1!A2#)')
	})

	test('insertCols rewrites spill references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(EMPTY, 'SEQUENCE(2)'))
		s.cells.set(0, 1, cell(EMPTY, 'SUM(A1#)'))

		applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 0, count: 1 })

		expect(s.cells.get(0, 2)?.formula).toBe('SUM(B1#)')
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

	test('insertRows within formula range expands range end', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		for (let r = 0; r < 10; r++) {
			s.cells.set(r, 0, cell(numberValue(r + 1)))
		}
		s.cells.set(10, 0, cell(EMPTY, 'SUM(A1:A10)'))

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 4, count: 3 })

		expect(s.cells.get(13, 0)?.formula).toBe('SUM(A1:A13)')
	})

	test('deleteRows within formula range shrinks range end', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		for (let r = 0; r < 10; r++) {
			s.cells.set(r, 0, cell(numberValue(r + 1)))
		}
		s.cells.set(10, 0, cell(EMPTY, 'SUM(A1:A10)'))

		applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 2, count: 3 })

		expect(s.cells.get(7, 0)?.formula).toBe('SUM(A1:A7)')
	})

	test('insertCols shifts formula references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(0, 1, cell(numberValue(2)))
		s.cells.set(0, 2, cell(numberValue(3)))
		s.cells.set(1, 0, cell(EMPTY, 'B1+C1'))

		applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 })

		expect(s.cells.get(1, 0)?.formula).toBe('C1+D1')
	})

	test('copyRange translates relative references when copying to new columns', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(10)))
		s.cells.set(1, 0, cell(numberValue(20)))
		s.cells.set(0, 1, cell(EMPTY, 'A1+A2'))

		applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			target: 'C1',
		})

		expect(s.cells.get(0, 3)?.formula).toBe('C1+C2')
	})

	test('moveRange preserves absolute references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(100)))
		s.cells.set(2, 0, cell(EMPTY, '$A$1+1'))

		applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A3',
			target: 'B3',
		})

		expect(s.cells.get(2, 1)?.formula).toBe('$A$1+1')
	})

	test('deleteSheet referenced by formula yields #REF! on recalc', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s2.cells.set(0, 0, cell(numberValue(42)))
		s1.cells.set(0, 0, cell(EMPTY, 'Sheet2!A1'))

		recalculate(wb, defaultCalcContext())
		expect(s1.cells.get(0, 0)?.value).toEqual(numberValue(42))

		applyOperation(wb, { op: 'deleteSheet', sheet: 'Sheet2' })
		recalculate(wb, defaultCalcContext())
		expect(s1.cells.get(0, 0)?.value).toEqual({ kind: 'error', value: '#REF!' })
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
		expectOk(result)

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
		expectErr(result)
		expect(result.error.code).toBe('SHEET_NOT_FOUND')
	})

	test('errors include suggestedFix for self-correction', () => {
		const wb = setup()
		wb.addSheet('Sheet2')

		const sheetNotFound = applyOperation(wb, {
			op: 'setCells',
			sheet: 'NoSuchSheet',
			updates: [{ ref: 'A1', value: 1 }],
		})
		expectErr(sheetNotFound)
		expect(sheetNotFound.error.suggestedFix).toContain('Available sheets:')
		expect(sheetNotFound.error.suggestedFix).toContain('Sheet1')
		expect(sheetNotFound.error.suggestedFix).toContain('Sheet2')

		const invalidRange = applyOperation(wb, {
			op: 'fillFormula',
			sheet: 'Sheet1',
			range: 'not-a-range',
			formula: '=1',
		})
		expectErr(invalidRange)
		expect(invalidRange.error.code).toBe('INVALID_RANGE')
		expect(invalidRange.error.suggestedFix).toContain('A1')

		const nameConflict = applyOperation(wb, { op: 'addSheet', name: 'Sheet1' })
		expectErr(nameConflict)
		expect(nameConflict.error.code).toBe('NAME_CONFLICT')
		expect(nameConflict.error.suggestedFix).toBeDefined()
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
		expectOk(result)

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
		expectOk(result)

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
		expectOk(result)

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
		expectOk(result)

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
		expectOk(result)

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
		expectOk(result)

		expect(sheet.cells.get(1, 0)?.value).toEqual({
			kind: 'date',
			serial: 1,
		})
	})

	test('appendRows inserts before totals row when hasTotals', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A2',
			name: 'BalanceTable',
			hasHeaders: true,
		})
		const table = sheet.tables[0]
		if (!table) throw new Error('expected table')
		sheet.tables[0] = { ...table, hasTotals: true }

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'BalanceTable',
			rows: [['Debt']],
		})
		expectOk(result)
		expect(sheet.tables[0]?.ref.end.row).toBe(2)
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('Debt'))
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('Cash'))
	})

	test('setTableColumn applies calculated-column formulas and totals metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(7), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:C3',
			name: 'Sales',
			hasHeaders: true,
		})

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Total',
			formula: '=[@Qty]*[@Price]',
			totalsRowFunction: 'sum',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['C2', 'C3'])
		expect(result.value.recalcRequired).toBe(true)
		expect(sheet.tables[0]?.columns[2]).toMatchObject({
			name: 'Total',
			formula: '[@Qty]*[@Price]',
			totalsRowFunction: 'sum',
		})
		expect(sheet.cells.get(1, 2)?.formula).toBe('[@Qty]*[@Price]')
		expect(sheet.cells.get(2, 2)?.formula).toBe('[@Qty]*[@Price]')

		const appended = applyOperation(wb, {
			op: 'appendRows',
			table: 'Sales',
			rows: [[4, 8]],
		})
		expectOk(appended)
		expect(sheet.cells.get(3, 2)?.formula).toBe('[@Qty]*[@Price]')
	})

	test('table management operations rename, resize, and delete table metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: stringValue('West'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'Sales',
			hasHeaders: true,
		})
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'SUM(Sales[Value])', styleId: sid })
		wb.definedNames.set('SalesValues', 'SUM(Sales[Value])')

		const renamed = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'Revenue',
		})
		expectOk(renamed)
		expect(renamed.value.recalcRequired).toBe(true)
		expect(sheet.tables[0]?.name).toBe('Revenue')
		expect(sheet.cells.get(3, 0)?.formula).toBe('SUM(Revenue[Value])')
		expect(wb.definedNames.get('SalesValues')).toBe('SUM(Revenue[Value])')

		const resized = applyOperation(wb, { op: 'resizeTable', table: 'Revenue', ref: 'A1:C2' })
		expectOk(resized)
		expect(resized.value.recalcRequired).toBe(true)
		expect(sheet.tables[0]?.ref.end.col).toBe(2)
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Name',
			'Value',
			'Region',
		])

		const deleted = applyOperation(wb, { op: 'deleteTable', table: 'Revenue' })
		expectOk(deleted)
		expect(deleted.value.recalcRequired).toBe(true)
		expect(sheet.tables).toHaveLength(0)
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(10))
	})

	test('setWorkbookProtection updates workbook-level protection metadata', () => {
		const wb = createWorkbook()
		const result = applyOperation(wb, {
			op: 'setWorkbookProtection',
			protection: {
				lockStructure: true,
				workbookAlgorithmName: 'SHA-512',
				workbookSpinCount: 100000,
			},
		})
		expectOk(result)
		expect(result.value.recalcRequired).toBe(false)
		expect(wb.workbookProtection).toEqual({
			lockStructure: true,
			workbookAlgorithmName: 'SHA-512',
			workbookSpinCount: 100000,
		})
	})

	test('setPivotCache updates source and refresh metadata by pivot table', () => {
		const wb = setup()
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			fields: [],
		})

		const result = applyOperation(wb, {
			op: 'setPivotCache',
			pivotTable: 'PivotTable1',
			sourceSheet: 'RawData',
			sourceRef: 'A1:E20',
			refreshOnLoad: true,
			invalid: true,
			saveData: false,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(result.value.warnings?.[0]?.message).toContain('Pivot cache source changed')
		expect(result.value.warnings?.[0]?.details).toMatchObject({
			cacheId: 34,
			refreshOnLoad: true,
			invalid: true,
		})
		expect(wb.pivotCaches[0]).toMatchObject({
			sourceSheet: 'RawData',
			sourceRef: 'A1:E20',
			refreshOnLoad: true,
			invalid: true,
			saveData: false,
		})
	})

	test('setConnectionRefresh updates query-table refresh metadata', () => {
		const wb = setup()
		wb.connectionParts.push({
			kind: 'queryTable',
			partPath: 'xl/queryTables/queryTable1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
			relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			sheetName: 'Sheet1',
			relationshipCount: 0,
			name: 'SalesQuery',
			connectionId: 1,
			refreshOnLoad: false,
			saveData: true,
		})

		const result = applyOperation(wb, {
			op: 'setConnectionRefresh',
			partPath: 'xl/queryTables/queryTable1.xml',
			connectionId: 1,
			refreshOnLoad: true,
			saveData: false,
			refreshedVersion: 8,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(result.value.warnings?.map((warning) => warning.message)).toEqual([
			'Connection is marked refresh-on-open; external data may change when Excel opens the workbook.',
			'Connection cache data is not saved; refresh is required before cached external output can be trusted.',
		])
		expect(wb.connectionParts[0]).toMatchObject({
			refreshOnLoad: true,
			saveData: false,
			refreshedVersion: 8,
		})
	})

	test('setConnectionRefresh validates selectors and editable fields', () => {
		const wb = setup()
		wb.connectionParts.push({
			kind: 'connection',
			partPath: 'xl/connections.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
			relationshipCount: 0,
			name: 'SalesConnection',
			connectionId: 1,
		})

		const missingSelector = applyOperation(wb, {
			op: 'setConnectionRefresh',
			refreshOnLoad: true,
		})
		expectErr(missingSelector)
		expect(missingSelector.error.message).toContain('requires partPath')

		const missingUpdate = applyOperation(wb, {
			op: 'setConnectionRefresh',
			partPath: 'xl/connections.xml',
		})
		expectErr(missingUpdate)
		expect(missingUpdate.error.message).toContain('requires refreshOnLoad')
	})

	test('setPivotFieldItem updates item and page filter state with refresh warning', () => {
		const wb = setup()
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					items: [
						{ index: 0, cacheIndex: 0, hidden: true },
						{ index: 1, cacheIndex: 1, showDetails: false },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0 }],
			dataFields: [],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			fields: [],
		})

		const result = applyOperation(wb, {
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 0,
			hidden: null,
			manualFilter: true,
			selectedPageItem: 1,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(result.value.warnings?.[0]?.message).toContain('Pivot field item state changed')
		expect(result.value.warnings?.[0]?.details).toMatchObject({
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 0,
			refreshOnLoad: true,
			invalid: true,
		})
		expect(wb.pivotTables[0]?.fields[0]?.items).toEqual([
			{ index: 0, cacheIndex: 0, manualFilter: true },
			{ index: 1, cacheIndex: 1, showDetails: false },
		])
		expect(wb.pivotTables[0]?.pageFields).toEqual([{ index: 0, item: 1 }])
		expect(wb.pivotCaches[0]).toMatchObject({ refreshOnLoad: true, invalid: true })
	})

	test('setPivotFieldItem validates selectors and existing inventory indexes', () => {
		const wb = setup()
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			fields: [{ index: 0, items: [{ index: 0, cacheIndex: 0 }] }],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})

		const missingSelector = applyOperation(wb, {
			op: 'setPivotFieldItem',
			fieldIndex: 0,
			itemIndex: 0,
			hidden: true,
		})
		expectErr(missingSelector)
		expect(missingSelector.error.message).toContain('requires pivotTable, partPath, or sheet')

		const missingUpdate = applyOperation(wb, {
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 0,
		})
		expectErr(missingUpdate)
		expect(missingUpdate.error.message).toContain('requires an item or page filter update')

		const missingItem = applyOperation(wb, {
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 2,
			hidden: true,
		})
		expectErr(missingItem)
		expect(missingItem.error.message).toContain('Pivot field item 2 was not found')

		const missingPageField = applyOperation(wb, {
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 0,
			selectedPageItem: 0,
		})
		expectErr(missingPageField)
		expect(missingPageField.error.message).toContain('Pivot field 0 is not a page field')

		const pivot = wb.pivotTables[0]
		if (!pivot) throw new Error('Expected pivot test fixture')
		wb.pivotTables[0] = { ...pivot, pageFields: [{ index: 0 }] }
		const missingPageItem = applyOperation(wb, {
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 0,
			selectedPageItem: 2,
		})
		expectErr(missingPageItem)
		expect(missingPageItem.error.message).toContain('Pivot page-field item 2 was not found')
	})

	test('setSlicerCacheItem updates tabular item state with refresh warning', () => {
		const wb = setup()
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			sourceName: 'State',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
			items: [{ index: 0, selected: true }, { index: 1 }],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			fields: [],
		})

		const result = applyOperation(wb, {
			op: 'setSlicerCacheItem',
			slicerCache: 'Slicer_State',
			item: 0,
			selected: null,
			noData: true,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(result.value.warnings?.[0]?.message).toContain('Slicer cache item state changed')
		expect(result.value.warnings?.[0]?.details).toMatchObject({
			slicerCache: 'Slicer_State',
			item: 0,
			pivotTables: ['PivotTable1'],
			cacheIds: [34],
			cachePartPaths: ['xl/pivotCache/pivotCacheDefinition1.xml'],
		})
		expect(wb.slicerCaches[0]?.items).toEqual([{ index: 0, noData: true }, { index: 1 }])
		expect(wb.pivotCaches[0]).toMatchObject({ refreshOnLoad: true, invalid: true })
	})

	test('setSlicerCacheItem validates selectors and editable flags', () => {
		const wb = setup()
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			pivotTableNames: [],
		})

		const missingSelector = applyOperation(wb, {
			op: 'setSlicerCacheItem',
			item: 0,
			selected: true,
		})
		expectErr(missingSelector)
		expect(missingSelector.error.message).toContain('requires slicerCache or partPath')

		const missingUpdate = applyOperation(wb, {
			op: 'setSlicerCacheItem',
			slicerCache: 'Slicer_State',
			item: 0,
		})
		expectErr(missingUpdate)
		expect(missingUpdate.error.message).toContain('requires selected or noData')
	})

	test('setTimelineRange updates selected date range with refresh warning', () => {
		const wb = setup()
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.timelineCaches.push({
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			name: 'Timeline_Order_Date',
			sourceName: 'Order Date',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
			state: {
				singleRangeFilterState: true,
				selection: {
					startDate: '2024-01-01T00:00:00',
					endDate: '2024-03-31T00:00:00',
				},
			},
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			fields: [],
		})

		const result = applyOperation(wb, {
			op: 'setTimelineRange',
			timelineCache: 'Timeline_Order_Date',
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-06-30T00:00:00',
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(result.value.warnings?.[0]?.message).toContain('Timeline range changed')
		expect(result.value.warnings?.[0]?.details).toMatchObject({
			timelineCache: 'Timeline_Order_Date',
			pivotTables: ['PivotTable1'],
			cacheIds: [34],
			cachePartPaths: ['xl/pivotCache/pivotCacheDefinition1.xml'],
		})
		expect(wb.timelineCaches[0]?.state?.selection).toEqual({
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-06-30T00:00:00',
		})
		expect(wb.pivotCaches[0]).toMatchObject({ refreshOnLoad: true, invalid: true })
	})

	test('setTimelineRange validates selectors and date range order', () => {
		const wb = setup()
		wb.timelineCaches.push({
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			name: 'Timeline_Order_Date',
			pivotTableNames: [],
		})

		const missingSelector = applyOperation(wb, {
			op: 'setTimelineRange',
			startDate: '2024-01-01T00:00:00',
			endDate: '2024-03-31T00:00:00',
		})
		expectErr(missingSelector)
		expect(missingSelector.error.message).toContain('requires timelineCache or partPath')

		const reversedRange = applyOperation(wb, {
			op: 'setTimelineRange',
			timelineCache: 'Timeline_Order_Date',
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-03-31T00:00:00',
		})
		expectErr(reversedRange)
		expect(reversedRange.error.message).toContain('startDate must be <= endDate')
	})

	test('rewriteExternalLink updates selected external workbook target metadata', () => {
		const wb = setup()
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			linkRelId: 'rIdExt',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const result = applyOperation(wb, {
			op: 'rewriteExternalLink',
			partPath: 'xl/externalLinks/externalLink1.xml',
			linkRelId: 'rIdExt',
			newTarget: '../sources/reforecast.xlsx',
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual([])
		expect(result.value.recalcRequired).toBe(false)
		expect(wb.externalReferenceDetails[0]).toMatchObject({
			target: '../sources/reforecast.xlsx',
			targetMode: 'External',
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
		expectOk(result)

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
		expectOk(result)

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

	test('collectAllErrors returns all validation errors', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const result = applyOperations(
			wb,
			[
				{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
				{ op: 'addSheet', name: 'Sheet1' },
				{ op: 'deleteSheet', sheet: 'NonExistent' },
			],
			{ collectAllErrors: true },
		)
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect('errors' in result.error).toBe(true)
		const errors = result.error.errors
		expect(errors).toHaveLength(3)
		expect(errors[0]?.code).toBe('SHEET_NOT_FOUND')
		expect(errors[0]?.message).toContain('Missing')
		expect(errors[1]?.code).toBe('NAME_CONFLICT')
		expect(errors[1]?.message).toContain('Sheet1')
		expect(errors[2]?.code).toBe('SHEET_NOT_FOUND')
		expect(errors[2]?.message).toContain('NonExistent')
	})
})

describe('applyWithTransaction', () => {
	test('batch of 3 operations where all succeed - workbook is modified', () => {
		const wb = createWorkbook()
		const result = applyWithTransaction(wb, [
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
		expectOk(result)

		const s = wb.getSheet('Sheet1')
		expect(s).toBeDefined()
		if (!s) return
		expect(s.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(s.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(s.cells.get(2, 0)?.formula).toBe('SUM(A1:A2)')
		expect(result.value.recalcRequired).toBe(true)
	})

	test('batch of 3 operations where the 3rd fails - workbook is NOT modified (rolled back)', () => {
		const wb = createWorkbook()
		const result = applyWithTransaction(wb, [
			{ op: 'addSheet', name: 'Sheet1' },
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 99 }],
			},
			{ op: 'deleteSheet', sheet: 'NonExistent' },
		])
		expectErr(result)

		expect(wb.getSheet('Sheet1')).toBeUndefined()
		expect(wb.sheets).toHaveLength(0)
	})
})
