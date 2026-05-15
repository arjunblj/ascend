import { describe, expect, test } from 'bun:test'
import type { Sheet, StyleId, Workbook } from '@ascend/core'
import { createTableId, createWorkbook } from '@ascend/core'
import { EMPTY, errorValue, numberValue, type Operation, stringValue } from '@ascend/schema'
import { analyzeWorkbook } from './analysis.ts'
import { recalculate } from './calc.ts'
import { defaultCalcContext } from './calc-context.ts'
import { cellKey } from './dep-graph.ts'
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

function normalizeSheet1Affected(refs: readonly string[]) {
	return refs.map((ref) => ref.replace(/^Sheet1!/, '')).sort()
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

function setupFormulaAnalysisWorkbook() {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	sheet.cells.set(0, 0, { value: numberValue(2), formula: 'B1*2', styleId: sid })
	sheet.cells.set(0, 1, cell(numberValue(1)))
	sheet.cells.set(0, 2, { value: numberValue(3), formula: 'A1+B1', styleId: sid })
	sheet.cells.set(1, 0, { value: numberValue(3), formula: 'A1+1', styleId: sid })
	sheet.cells.set(1, 1, cell(numberValue(4)))
	sheet.cells.set(1, 2, { value: numberValue(7), formula: 'SUM(A1:B2)', styleId: sid })
	return wb
}

function formulaAnalysisSnapshot(wb: Workbook) {
	const analysis = analyzeWorkbook(wb)
	return {
		formulas: [...analysis.formulas.entries()]
			.map(([key, formula]) => ({
				key,
				sheetName: formula.sheetName,
				row: formula.row,
				col: formula.col,
				formula: formula.formula,
				refs: formula.refs,
				deps: [...formula.deps].sort(),
				rangeDeps: formula.rangeDeps,
				parseError: formula.parseError,
			}))
			.sort((a, b) => String(a.key).localeCompare(String(b.key))),
		sharedFormulaGroups: [...analysis.sharedFormulaGroups.entries()]
			.map(([key, members]) => ({ key, members: [...members].sort() }))
			.sort((a, b) => a.key.localeCompare(b.key)),
	}
}

function expectCachedFormulaAnalysisMatchesFullRecompute(wb: Workbook) {
	expect(formulaAnalysisSnapshot(wb)).toEqual(formulaAnalysisSnapshot(wb.clone()))
}

function addSharedFormulaGroup(sheet: Sheet, startRow = 0, col = 0) {
	const masterRef = `${String.fromCharCode(65 + col)}${startRow + 1}`
	sheet.cells.set(startRow, col, {
		value: numberValue(20),
		formula: 'B1*2',
		styleId: sid,
		formulaInfo: {
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef,
			ref: `${masterRef}:${String.fromCharCode(65 + col)}${startRow + 2}`,
		},
	})
	sheet.cells.set(startRow + 1, col, {
		value: numberValue(40),
		formula: null,
		styleId: sid,
		formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef },
	})
}

function addSpillGroup(sheet: Sheet, startRow = 0, col = 0) {
	const column = String.fromCharCode(65 + col)
	const anchorRef = `Sheet1!${column}${startRow + 1}`
	const ref = `${column}${startRow + 1}:${column}${startRow + 3}`
	const anchorInfo = { kind: 'spill' as const, anchorRef, ref, isAnchor: true }
	const memberInfo = { kind: 'spill' as const, anchorRef, ref, isAnchor: false }
	sheet.cells.set(startRow, col, {
		value: numberValue(1),
		formula: 'SEQUENCE(3)',
		styleId: sid,
		formulaInfo: anchorInfo,
	})
	sheet.cells.set(startRow + 1, col, {
		value: numberValue(2),
		formula: null,
		styleId: sid,
		formulaInfo: memberInfo,
	})
	sheet.cells.set(startRow + 2, col, {
		value: numberValue(3),
		formula: null,
		styleId: sid,
		formulaInfo: memberInfo,
	})
}

function addDynamicArrayAnchor(sheet: Sheet, row = 0, col = 0) {
	sheet.cells.set(row, col, {
		value: numberValue(1),
		formula: 'SEQUENCE(3)',
		styleId: sid,
		formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
	})
}

function addDynamicArrayAnchorWithStaleSpillFootprint(sheet: Sheet, anchorRef = 'Sheet1!A1') {
	const spillRef = 'A1:A3'
	addDynamicArrayAnchor(sheet)
	sheet.cells.set(1, 0, {
		value: numberValue(2),
		formula: null,
		styleId: sid,
		formulaInfo: { kind: 'spill', anchorRef, ref: spillRef, isAnchor: false },
	})
	sheet.cells.set(2, 0, {
		value: numberValue(3),
		formula: null,
		styleId: sid,
		formulaInfo: { kind: 'spill', anchorRef, ref: spillRef, isAnchor: false },
	})
}

function addBlockedSpillFormula(sheet: Sheet) {
	sheet.cells.set(0, 0, {
		value: errorValue('#SPILL!'),
		formula: 'SEQUENCE(3)',
		styleId: sid,
		formulaInfo: {
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			blockingRefs: ['A2'],
		},
	})
	sheet.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: sid })
}

function addDataTableFormula(sheet: Sheet) {
	sheet.cells.set(2, 2, {
		value: numberValue(10),
		formula: null,
		styleId: sid,
		formulaInfo: {
			kind: 'dataTable',
			ref: 'C3:C5',
			dt2D: false,
			dtr: true,
			r1: 'A1',
		},
	})
	sheet.cells.set(3, 2, cell(numberValue(20)))
	sheet.cells.set(4, 2, cell(numberValue(30)))
}

function setupSalesRepTable() {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	sheet.tables.push({
		id: createTableId(),
		name: 'Sales',
		sheetId: sheet.id,
		ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
		columns: [
			{ id: 1, name: 'Region' },
			{ id: 2, name: 'Rep' },
			{ id: 3, name: 'Amount' },
		],
		hasHeaders: true,
		hasTotals: false,
	})
	return { wb, sheet }
}

function setupQueryTableBackedSalesTable() {
	const { wb, sheet } = setupSalesRepTable()
	const table = sheet.tables[0]
	if (!table) throw new Error('Expected Sales table')
	sheet.tables[0] = {
		...table,
		partPath: 'xl/tables/table1.xml',
		columns: table.columns.map((column, index) => ({
			...column,
			queryTableFieldId: index + 1,
		})),
		queryTable: {
			relationshipId: 'rIdQueryTable1',
			partPath: 'xl/queryTables/queryTable1.xml',
			relationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			relationshipRawType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			target: '../queryTables/queryTable1.xml',
		},
	}
	wb.connectionParts.push({
		kind: 'queryTable',
		partPath: 'xl/queryTables/queryTable1.xml',
		contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
		relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
		sheetName: 'Sheet1',
		relationshipCount: 0,
		name: 'SalesQuery',
		connectionId: 1,
	})
	return { wb, sheet }
}

function setupDuplicateSalesTables() {
	const wb = createWorkbook()
	const sheet1 = wb.addSheet('Sheet1')
	const sheet2 = wb.addSheet('Sheet2')
	sheet1.tables.push({
		id: createTableId(),
		name: 'Sales',
		sheetId: sheet1.id,
		partPath: 'xl/tables/table1.xml',
		ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
		columns: [{ name: 'Region' }, { name: 'Amount' }],
		hasHeaders: true,
		hasTotals: false,
	})
	sheet2.tables.push({
		id: createTableId(),
		name: 'Sales',
		sheetId: sheet2.id,
		partPath: 'xl/tables/table2.xml',
		ref: { start: { row: 4, col: 3 }, end: { row: 6, col: 4 } },
		columns: [{ name: 'Region' }, { name: 'Amount' }],
		hasHeaders: true,
		hasTotals: false,
	})
	return { wb, sheet1, sheet2 }
}

const TABLE_FIELD_METADATA_BLOCKERS: readonly {
	readonly label: string
	readonly sourceKind: string
	readonly sourceRef: string
	readonly add: (sheet: Sheet) => void
}[] = [
	{
		label: 'normal data validation formula2',
		sourceKind: 'data validation formula2',
		sourceRef: 'Sheet1!E1:E4',
		add: (sheet) =>
			sheet.dataValidations.push({
				sqref: 'E1:E4',
				type: 'decimal',
				formula1: '0',
				formula2: 'MAX(Sales[Rep])',
			}),
	},
	{
		label: 'normal conditional format formula',
		sourceKind: 'conditional format formula',
		sourceRef: 'Sheet1!F1:F4',
		add: (sheet) =>
			sheet.conditionalFormats.push({
				sqref: 'F1:F4',
				rules: [{ type: 'expression', formulas: ['COUNTA(Sales[Rep])>0'] }],
			}),
	},
	{
		label: 'normal conditional format value object',
		sourceKind: 'conditional format value object',
		sourceRef: 'Sheet1!G1:G4',
		add: (sheet) =>
			sheet.conditionalFormats.push({
				sqref: 'G1:G4',
				rules: [
					{
						type: 'expression',
						formulas: [],
						colorScale: {
							cfvo: [{ type: 'formula', value: 'COUNTA(Sales[Rep])' }],
							colors: [{ rgb: 'FFFF0000' }],
						},
					},
				],
			}),
	},
	{
		label: 'x14 data validation formula2',
		sourceKind: 'x14 data validation formula2',
		sourceRef: 'Sheet1!H1:H4',
		add: (sheet) =>
			sheet.x14DataValidations.push({
				index: 0,
				sqref: 'H1:H4',
				type: 'decimal',
				formula1: '0',
				formula2: 'MAX(Sales[Rep])',
			}),
	},
	{
		label: 'x14 conditional format formula',
		sourceKind: 'x14 conditional format formula',
		sourceRef: 'Sheet1!I1:I4',
		add: (sheet) =>
			sheet.x14ConditionalFormats.push({
				index: 0,
				sqref: 'I1:I4',
				formulas: ['COUNTA(Sales[Rep])>0'],
			}),
	},
	{
		label: 'x14 conditional format value object',
		sourceKind: 'x14 conditional format value object',
		sourceRef: 'Sheet1!J1:J4',
		add: (sheet) =>
			sheet.x14ConditionalFormats.push({
				index: 0,
				sqref: 'J1:J4',
				dataBar: { cfvo: [{ type: 'formula', value: 'COUNTA(Sales[Rep])' }] },
			}),
	},
]

const LOCAL_TABLE_FIELD_METADATA_BLOCKERS: readonly {
	readonly label: string
	readonly sourceKind: string
	readonly sourceRef: string
	readonly add: (sheet: Sheet) => void
}[] = [
	{
		label: 'local data validation formula2',
		sourceKind: 'data validation formula2',
		sourceRef: 'Sheet1!B2:B4',
		add: (sheet) =>
			sheet.dataValidations.push({
				sqref: 'B2:B4',
				type: 'decimal',
				formula1: '0',
				formula2: 'MAX([@Rep])',
			}),
	},
	{
		label: 'local conditional format formula',
		sourceKind: 'conditional format formula',
		sourceRef: 'Sheet1!B2:B4',
		add: (sheet) =>
			sheet.conditionalFormats.push({
				sqref: 'B2:B4',
				rules: [{ type: 'expression', formulas: ['COUNTA([@Rep])>0'] }],
			}),
	},
	{
		label: 'local conditional format value object',
		sourceKind: 'conditional format value object',
		sourceRef: 'Sheet1!B2:B4',
		add: (sheet) =>
			sheet.conditionalFormats.push({
				sqref: 'B2:B4',
				rules: [
					{
						type: 'expression',
						formulas: [],
						colorScale: {
							cfvo: [{ type: 'formula', value: 'COUNTA([@Rep])' }],
							colors: [{ rgb: 'FFFF0000' }],
						},
					},
				],
			}),
	},
	{
		label: 'local x14 data validation formula2',
		sourceKind: 'x14 data validation formula2',
		sourceRef: 'Sheet1!B2:B4',
		add: (sheet) =>
			sheet.x14DataValidations.push({
				index: 0,
				sqref: 'B2:B4',
				type: 'decimal',
				formula1: '0',
				formula2: 'MAX([@Rep])',
			}),
	},
	{
		label: 'local x14 conditional format formula',
		sourceKind: 'x14 conditional format formula',
		sourceRef: 'Sheet1!B2:B4',
		add: (sheet) =>
			sheet.x14ConditionalFormats.push({
				index: 0,
				sqref: 'B2:B4',
				formulas: ['COUNTA([@Rep])>0'],
			}),
	},
	{
		label: 'local x14 conditional format value object',
		sourceKind: 'x14 conditional format value object',
		sourceRef: 'Sheet1!B2:B4',
		add: (sheet) =>
			sheet.x14ConditionalFormats.push({
				index: 0,
				sqref: 'B2:B4',
				dataBar: { cfvo: [{ type: 'formula', value: 'COUNTA([@Rep])' }] },
			}),
	},
]

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

	test('setCells skips literal no-ops without dirtying the patch', () => {
		const wb = setup()
		expectOk(
			applyOperation(wb, {
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A1',
				rule: { type: 'whole', formula1: '1', formula2: '5', operator: 'between' },
			}),
		)
		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 10 },
				{ ref: 'D1', value: null },
			],
		})
		expectOk(result)

		expect(result.value).toEqual({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
		})
		expect(wb.getSheet('Sheet1')?.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(wb.getSheet('Sheet1')?.cells.get(0, 3)).toBeUndefined()
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

	test('setFormula materializes imported shared formula groups before rewriting a member', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: numberValue(20),
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'A1',
				ref: 'A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(40),
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})

		const result = applyOperation(wb, {
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'A2',
			formula: 'B2*3',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formula).toBe('B2*3')
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
	})

	test('setFormula materializes spill groups before rewriting a member', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const anchorInfo = {
			kind: 'spill' as const,
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: true,
		}
		const memberInfo = {
			kind: 'spill' as const,
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: false,
		}
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(3)',
			styleId: sid,
			formulaInfo: anchorInfo,
		})
		sheet.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: memberInfo,
		})
		sheet.cells.set(2, 0, {
			value: numberValue(3),
			formula: null,
			styleId: sid,
			formulaInfo: memberInfo,
		})

		const result = applyOperation(wb, {
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'A2',
			formula: '10+1',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formula).toBe('10+1')
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formula).toBeNull()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()

		expect(recalculate(wb, defaultCalcContext()).errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(11))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
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

	test('setCells materializes imported shared formula groups before literal replacement', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSharedFormulaGroup(sheet)

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formula).toBeNull()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells materializes shared formula groups with case-insensitive sheet-qualified master refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: numberValue(20),
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'Sheet1!A1',
				ref: 'A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(40),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: false,
				masterRef: 'sheet1!$A$1',
			},
		})

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formula).toBeNull()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells scopes shared formula materialization by master when shared indexes collide', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSharedFormulaGroup(sheet)
		sheet.cells.set(0, 2, {
			value: numberValue(30),
			formula: 'D1*2',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'C1',
				ref: 'C1:C2',
			},
		})
		sheet.cells.set(1, 2, {
			value: numberValue(60),
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'C1' },
		})

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(0, 2)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'C1',
			ref: 'C1:C2',
		})
		expect(sheet.cells.get(1, 2)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'C1',
		})
	})

	test('setCells materializes spill groups before literal replacement', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSpillGroup(sheet)

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells detaches stale dynamic-array spill footprints before replacing the anchor', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchorWithStaleSpillFootprint(sheet)

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A1', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells detaches dynamic-array anchors before replacing a spill member', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchorWithStaleSpillFootprint(sheet)

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells detaches dynamic-array anchors with absolute spill anchor refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchorWithStaleSpillFootprint(sheet, 'Sheet1!$A$1')

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells detaches dynamic-array anchors with case-insensitive spill anchor refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchorWithStaleSpillFootprint(sheet, 'sheet1!$A$1')

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setCells detaches only the edited dynamic-array group when metadata indexes collide', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchorWithStaleSpillFootprint(sheet)
		sheet.cells.set(0, 2, {
			value: numberValue(10),
			formula: 'SEQUENCE(2)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(1, 2, {
			value: numberValue(11),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!C1',
				ref: 'C1:C2',
				isAnchor: false,
			},
		})

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A1', value: 9 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(0, 2)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(sheet.cells.get(1, 2)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!C1',
			ref: 'C1:C2',
			isAnchor: false,
		})
	})

	test('setCells detaches dynamic-array spill footprints with escaped anchor refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet("Bob's Budget")
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(3)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: "'Bob''s Budget'!A1",
				ref: 'A1:A3',
				isAnchor: false,
			},
		})
		sheet.cells.set(2, 0, {
			value: numberValue(3),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: "'Bob''s Budget'!A1",
				ref: 'A1:A3',
				isAnchor: false,
			},
		})

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: "Bob's Budget",
			updates: [{ ref: 'A1', value: 99 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('destructive edits against dynamic-array spill members detach the anchor and footprint', () => {
		const cases: readonly {
			readonly name: string
			readonly op: Parameters<typeof applyOperation>[1]
			readonly seed?: (sheet: Sheet) => void
			readonly affected?: readonly string[]
		}[] = [
			{
				name: 'setCells',
				op: { op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 9 }] },
			},
			{
				name: 'setRichText',
				op: { op: 'setRichText', sheet: 'Sheet1', ref: 'A2', runs: [{ text: 'override' }] },
			},
			{
				name: 'fillFormula',
				op: { op: 'fillFormula', sheet: 'Sheet1', range: 'A2', formula: 'B2*3' },
			},
			{
				name: 'clearRange all',
				op: { op: 'clearRange', sheet: 'Sheet1', range: 'A2', what: 'all' },
			},
			{
				name: 'copyRange values',
				op: { op: 'copyRange', sheet: 'Sheet1', source: 'C1', target: 'A2', mode: 'values' },
				seed: (sheet) => sheet.cells.set(0, 2, cell(numberValue(9))),
			},
			{
				name: 'moveRange',
				op: { op: 'moveRange', sheet: 'Sheet1', source: 'A2', target: 'C1', mode: 'all' },
				affected: ['A1', 'A2', 'A3', 'C1'],
			},
		]

		for (const entry of cases) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			addDynamicArrayAnchorWithStaleSpillFootprint(sheet)
			entry.seed?.(sheet)

			const result = applyOperation(wb, entry.op)
			expectOk(result)

			expect(normalizeSheet1Affected(result.value.affectedCells), entry.name).toEqual(
				entry.affected ?? ['A1', 'A2', 'A3'],
			)
			expect(sheet.cells.get(0, 0)?.formulaInfo, entry.name).toBeUndefined()
			expect(sheet.cells.get(1, 0)?.formulaInfo, entry.name).toBeUndefined()
			expect(sheet.cells.get(2, 0)?.formulaInfo, entry.name).toBeUndefined()
		}
	})

	test('style-only edits against dynamic-array spill members preserve formula bindings', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchorWithStaleSpillFootprint(sheet)

		const result = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'styles',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A2'])
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(sheet.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: false,
		})
		expect(sheet.cells.get(2, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: false,
		})
	})

	test('setCells detaches blocked-spill metadata when replacing a blocker', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addBlockedSpillFormula(sheet)

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: null }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(EMPTY)
	})

	test('setCells detaches data-table metadata when replacing a data-table member', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDataTableFormula(sheet)

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'C4', value: 99 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['C3', 'C4'])
		expect(sheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(99))
		expect(sheet.cells.get(4, 2)?.value).toEqual(numberValue(30))
	})

	test('setCells detaches blocked-spill metadata with case-insensitive sheet-qualified ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: errorValue('#SPILL!'),
			formula: 'SEQUENCE(3)',
			styleId: sid,
			formulaInfo: {
				kind: 'blockedSpill',
				anchorRef: 'Sheet1!A1',
				ref: 'sheet1!A1:A3',
				blockingRefs: ['sheet1!A2'],
			},
		})
		sheet.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: sid })

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A2', value: null }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(EMPTY)
	})

	test('setCells detaches data-table metadata with case-insensitive sheet-qualified ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(2, 2, {
			value: numberValue(10),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'dataTable',
				ref: 'sheet1!C3:C5',
				dt2D: false,
				dtr: true,
				r1: 'A1',
			},
		})
		sheet.cells.set(3, 2, cell(numberValue(20)))
		sheet.cells.set(4, 2, cell(numberValue(30)))

		const result = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'C4', value: 99 }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['C3', 'C4'])
		expect(sheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(99))
		expect(sheet.cells.get(4, 2)?.value).toEqual(numberValue(30))
	})

	test('fillFormula materializes imported shared formula groups before replacement', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSharedFormulaGroup(sheet)

		const result = applyOperation(wb, {
			op: 'fillFormula',
			sheet: 'Sheet1',
			range: 'A2:A2',
			formula: 'B2*3',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formula).toBe('B2*3')
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
	})

	test('fillFormula materializes dynamic-array metadata before replacement', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDynamicArrayAnchor(sheet)

		const result = applyOperation(wb, {
			op: 'fillFormula',
			sheet: 'Sheet1',
			range: 'A1',
			formula: '1+1',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('1+1')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
	})

	test('clearRange values detaches data-table metadata but style-only edits preserve it', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addDataTableFormula(sheet)

		const styles = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'C4',
			what: 'styles',
		})
		expectOk(styles)
		expect(styles.value.affectedCells).toEqual(['C4'])
		expect(sheet.cells.get(2, 2)?.formulaInfo).toEqual({
			kind: 'dataTable',
			ref: 'C3:C5',
			dt2D: false,
			dtr: true,
			r1: 'A1',
		})

		const values = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'C4',
			what: 'values',
		})
		expectOk(values)
		expect(values.value.affectedCells).toEqual(['C3', 'C4'])
		expect(sheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(3, 2)?.value).toEqual(EMPTY)
	})

	test('clearRange values detaches blocked-spill metadata when clearing a blocker but style clears preserve it', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addBlockedSpillFormula(sheet)

		const styles = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'styles',
		})
		expectOk(styles)
		expect(styles.value.affectedCells).toEqual(['A2'])
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			blockingRefs: ['A2'],
		})

		const formulas = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'formulas',
		})
		expectOk(formulas)
		expect(formulas.value).toEqual({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
		})
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			blockingRefs: ['A2'],
		})
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('blocker'))

		const values = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'values',
		})
		expectOk(values)
		expect(values.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(EMPTY)
	})

	test('clearRange formulas detaches blocked-spill metadata when clearing the anchor', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addBlockedSpillFormula(sheet)

		const result = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A1',
			what: 'formulas',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('blocker'))
	})

	test('clearRange styles preserves legacy array formula metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		const formulaInfo = { kind: 'array' as const, ref: 'A1:A2' }
		sheet.cells.set(0, 0, {
			value: numberValue(10),
			formula: 'B1:B2*2',
			styleId: sourceStyle,
			formulaInfo,
		})
		sheet.cells.set(1, 0, {
			value: numberValue(20),
			formula: null,
			styleId: sourceStyle,
			formulaInfo,
		})

		const result = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A1:A2',
			what: 'styles',
		})
		expectOk(result)

		expect(result.value).toEqual({
			affectedCells: ['A1', 'A2'],
			sheetsModified: ['Sheet1'],
			recalcRequired: false,
		})
		expect(sheet.cells.get(0, 0)).toEqual({
			value: numberValue(10),
			formula: 'B1:B2*2',
			styleId: sid,
			formulaInfo,
		})
		expect(sheet.cells.get(1, 0)).toEqual({
			value: numberValue(20),
			formula: null,
			styleId: sid,
			formulaInfo,
		})
	})

	test('setRichText materializes spill groups before replacing a spill member', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSpillGroup(sheet)

		const result = applyOperation(wb, {
			op: 'setRichText',
			sheet: 'Sheet1',
			ref: 'A2',
			runs: [{ text: 'manual', bold: true }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'manual', bold: true }],
		})
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setRichText detaches non-spill formula metadata before replacement', () => {
		const sharedWb = createWorkbook()
		const shared = sharedWb.addSheet('Sheet1')
		addSharedFormulaGroup(shared)

		const sharedResult = applyOperation(sharedWb, {
			op: 'setRichText',
			sheet: 'Sheet1',
			ref: 'A2',
			runs: [{ text: 'manual shared member' }],
		})
		expectOk(sharedResult)
		expect(sharedResult.value.affectedCells).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(shared.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(shared.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(shared.cells.get(1, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'manual shared member' }],
		})
		expect(shared.cells.get(1, 0)?.formulaInfo).toBeUndefined()

		const blockedWb = createWorkbook()
		const blocked = blockedWb.addSheet('Sheet1')
		addBlockedSpillFormula(blocked)
		const blockedResult = applyOperation(blockedWb, {
			op: 'setRichText',
			sheet: 'Sheet1',
			ref: 'A2',
			runs: [{ text: 'manual blocker' }],
		})
		expectOk(blockedResult)
		expect(blockedResult.value.affectedCells).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(blocked.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(blocked.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(blocked.cells.get(1, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'manual blocker' }],
		})

		const tableWb = createWorkbook()
		const tableSheet = tableWb.addSheet('Sheet1')
		addDataTableFormula(tableSheet)
		const tableResult = applyOperation(tableWb, {
			op: 'setRichText',
			sheet: 'Sheet1',
			ref: 'C4',
			runs: [{ text: 'manual table member' }],
		})
		expectOk(tableResult)
		expect(tableResult.value.affectedCells).toEqual(['Sheet1!C3', 'Sheet1!C4'])
		expect(tableSheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
		expect(tableSheet.cells.get(3, 2)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'manual table member' }],
		})
		expect(tableSheet.cells.get(3, 2)?.formulaInfo).toBeUndefined()
	})

	test('clearRange materializes spill groups before deleting a member', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSpillGroup(sheet)

		const result = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'all',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2', 'A3'])
		expect(sheet.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('clearRange materializes formula bindings for formula clears but preserves them for value and style clears', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSharedFormulaGroup(sheet)
		addSpillGroup(sheet, 0, 2)

		const styles = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'C2',
			what: 'styles',
		})
		expectOk(styles)
		expect(sheet.cells.get(1, 2)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!C1',
			ref: 'C1:C3',
			isAnchor: false,
		})

		const values = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'values',
		})
		expectOk(values)
		expect(values.value.affectedCells).toEqual(['A2'])
		expect(sheet.cells.get(1, 0)?.value).toEqual(EMPTY)
		expect(sheet.cells.get(1, 0)?.formula).toBeNull()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A1',
		})

		const formulas = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A2',
			what: 'formulas',
		})
		expectOk(formulas)
		expect(formulas.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(sheet.cells.get(1, 0)?.formula).toBeNull()
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
	})

	test('clearRange values preserves spill formulas and metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSpillGroup(sheet)

		const result = applyOperation(wb, {
			op: 'clearRange',
			sheet: 'Sheet1',
			range: 'A1',
			what: 'values',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1'])
		expect(sheet.cells.get(0, 0)?.value).toEqual(EMPTY)
		expect(sheet.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: true,
		})
		expect(sheet.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: false,
		})
	})

	test('fillFormula patches materialized formula bindings in cached analysis', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSpillGroup(sheet)
		expect(analyzeWorkbook(wb).formulas.get(cellKey(0, 0, 0))?.formula).toBe('SEQUENCE(3)')

		const result = applyOperation(wb, {
			op: 'fillFormula',
			sheet: 'Sheet1',
			range: 'A2',
			formula: 'B2*3',
		})
		expectOk(result)

		const analysis = analyzeWorkbook(wb)
		expect(analysis.formulas.get(cellKey(0, 0, 0))).toBeUndefined()
		expect(analysis.formulas.get(cellKey(0, 1, 0))?.formula).toBe('B2*3')
		expect(analysis.formulas.get(cellKey(0, 2, 0))).toBeUndefined()
	})

	test('incremental formula analysis patches match full recomputation for representative edits', () => {
		const cases: readonly {
			readonly name: string
			readonly op: Operation
		}[] = [
			{
				name: 'setFormula',
				op: { op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: 'B1*5' },
			},
			{
				name: 'fillFormula',
				op: { op: 'fillFormula', sheet: 'Sheet1', range: 'C1:C2', formula: 'A1+B1' },
			},
			{
				name: 'same-sheet copyRange',
				op: { op: 'copyRange', sheet: 'Sheet1', source: 'A1:C1', target: 'A3' },
			},
			{
				name: 'insertRows',
				op: { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 },
			},
			{
				name: 'deleteRows',
				op: { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 },
			},
			{
				name: 'insertCols',
				op: { op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 },
			},
			{
				name: 'deleteCols',
				op: { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 },
			},
		]

		for (const entry of cases) {
			const wb = setupFormulaAnalysisWorkbook()
			expectCachedFormulaAnalysisMatchesFullRecompute(wb)

			const result = applyOperation(wb, entry.op)
			expectOk(result)
			expectCachedFormulaAnalysisMatchesFullRecompute(wb)
		}
	})

	test('cell edits reject partial legacy array formula ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const formulaInfo = { kind: 'array' as const, ref: 'A1:B2' }
		sheet.cells.set(0, 0, { value: numberValue(1), formula: 'A3:B4', styleId: sid, formulaInfo })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid, formulaInfo })

		expectErr(
			applyOperation(wb, {
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'B2', value: 9 }],
			}),
		)
		expectErr(applyOperation(wb, { op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' }))
		expectErr(
			applyOperation(wb, { op: 'fillFormula', sheet: 'Sheet1', range: 'B2:C2', formula: '1+1' }),
		)
		expectErr(
			applyOperation(wb, { op: 'clearRange', sheet: 'Sheet1', range: 'A1:B1', what: 'values' }),
		)
		expectErr(
			applyOperation(wb, {
				op: 'setRichText',
				sheet: 'Sheet1',
				ref: 'B2',
				runs: [{ text: 'blocked' }],
			}),
		)
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual(formulaInfo)
	})

	test('range transfers reject legacy array formula intersections', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const formulaInfo = { kind: 'array' as const, ref: 'A1:B2' }
		sheet.cells.set(0, 0, { value: numberValue(1), formula: 'A3:B4', styleId: sid, formulaInfo })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(4, 4, cell(numberValue(99)))

		expectErr(applyOperation(wb, { op: 'copyRange', sheet: 'Sheet1', source: 'E5', target: 'B2' }))
		expectErr(applyOperation(wb, { op: 'moveRange', sheet: 'Sheet1', source: 'A1', target: 'E6' }))

		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual(formulaInfo)
		expect(sheet.cells.get(4, 4)?.value).toEqual(numberValue(99))
		expect(sheet.cells.get(5, 4)).toBeUndefined()
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
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) throw new Error('expected sheet')
		const styleId = wb.styles.register({ numberFormat: '@' })
		sheet.cells.set(1, 1, { value: stringValue('plain'), formula: null, styleId })

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

		const cell = sheet.cells.get(1, 1)
		expect(cell?.value).toEqual({
			kind: 'richText',
			runs: [
				{ text: 'Hello', bold: true },
				{ text: ' World', italic: true },
			],
		})
		expect(cell?.styleId).toBe(styleId)
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

	test('setDrawingText preserves drawing object identity and relationship metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 7,
			name: 'Chart Callout',
			description: 'Revenue callout',
			text: 'Original',
			anchor: {
				kind: 'absolute',
				x: 1000,
				y: 2000,
				cx: 3000,
				cy: 4000,
			},
			relIds: ['rIdChart1'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})

		const result = applyOperation(wb, {
			op: 'setDrawingText',
			sheet: 'Sheet1',
			drawingObjectIndex: 0,
			text: 'Updated',
		})
		expectOk(result)

		expect(sheet.drawingObjectRefs[0]).toEqual({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 7,
			name: 'Chart Callout',
			description: 'Revenue callout',
			text: 'Updated',
			anchor: {
				kind: 'absolute',
				x: 1000,
				y: 2000,
				cx: 3000,
				cy: 4000,
			},
			relIds: ['rIdChart1'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
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

	test('setComment edits text while preserving legacy drawing metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.comments.set('A1', {
			text: 'Original',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 0,
				column: 0,
				anchor: [0, 15, 0, 2, 2, 15, 3, 2],
				visible: true,
			},
		})

		const result = applyOperation(wb, {
			op: 'setComment',
			sheet: 'Sheet1',
			ref: 'a1',
			text: 'Updated',
		})
		expectOk(result)

		expect(sheet.comments.get('A1')).toEqual({
			text: 'Updated',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 0,
				column: 0,
				anchor: [0, 15, 0, 2, 2, 15, 3, 2],
				visible: true,
			},
		})
	})

	test('deleteComment removes legacy and threaded comments at the cell ref', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.comments.set('A1', { text: 'Legacy' })
		sheet.threadedComments.push(
			{ ref: 'A1', text: 'Root', id: 'tc1' },
			{ ref: 'A1', text: 'Reply', id: 'tc2', parentId: 'tc1' },
			{ ref: 'B1', text: 'Keep', id: 'tc3' },
		)

		const result = applyOperation(wb, { op: 'deleteComment', sheet: 'Sheet1', ref: 'a1' })
		expectOk(result)

		expect(sheet.comments.has('A1')).toBe(false)
		expect(sheet.threadedComments).toEqual([{ ref: 'B1', text: 'Keep', id: 'tc3' }])
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

	test('insertImage reuses existing drawing object part without colliding with chart relationships', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing7.xml',
			kind: 'graphicFrame',
			id: 5,
			name: 'Revenue Chart',
			relationshipRefs: [
				{
					id: 'rId1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})

		const result = applyOperation(wb, {
			op: 'insertImage',
			sheet: 'Sheet1',
			contentBase64: 'BAUG',
			contentType: 'image/png',
			anchor: {
				kind: 'twoCell',
				from: { row: 2, col: 2, rowOff: 10 },
				to: { row: 6, col: 5, colOff: 20 },
				editAs: 'oneCell',
			},
		})
		expectOk(result)

		expect(sheet.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing7.xml',
			relId: 'rIdImage1',
			anchor: {
				kind: 'twoCell',
				from: { row: 2, col: 2, rowOff: 10 },
				to: { row: 6, col: 5, colOff: 20 },
				editAs: 'oneCell',
			},
		})
		expect(sheet.drawingObjectRefs[0]?.relationshipRefs?.[0]?.id).toBe('rId1')
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

	test('deleteImage preserves drawing state when non-image drawing objects remain', () => {
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
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 2,
			name: 'Callout',
			text: 'Keep me',
		})

		const result = applyOperation(wb, {
			op: 'deleteImage',
			sheet: 'Sheet1',
			name: 'Logo',
		})
		expectOk(result)

		expect(sheet.imageRefs).toEqual([])
		expect(sheet.drawingObjectRefs).toHaveLength(1)
		expect(sheet.drawingRefs.hasDrawing).toBe(true)
	})

	test('insertImage rejects workbook-wide targetPath collisions', () => {
		const wb = setup()
		const sheet2 = wb.addSheet('Sheet2')
		sheet2.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing2.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
		})

		const result = applyOperation(wb, {
			op: 'insertImage',
			sheet: 'Sheet1',
			targetPath: 'xl/media/image1.png',
			contentBase64: 'BAUG',
			contentType: 'image/png',
		})

		expectErr(result)
		expect(result.error.message).toContain('targetPath already exists')
	})

	test('insertImage rejects relationship id collisions with drawing objects in the same drawing part', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 5,
			name: 'Revenue Chart',
			relIds: ['rId2'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})

		const relIdsResult = applyOperation(wb, {
			op: 'insertImage',
			sheet: 'Sheet1',
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId2',
			contentBase64: 'BAUG',
			contentType: 'image/png',
		})
		expectErr(relIdsResult)
		expect(relIdsResult.error.message).toContain('relId already exists')

		const relationshipRefsResult = applyOperation(wb, {
			op: 'insertImage',
			sheet: 'Sheet1',
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdChart1',
			contentBase64: 'BAUG',
			contentType: 'image/png',
		})
		expectErr(relationshipRefsResult)
		expect(relationshipRefsResult.error.message).toContain('relId already exists')
		expect(sheet.imageRefs).toEqual([])
	})

	test('setChartSeriesSource updates parsed chart source refs', () => {
		const wb = setup()
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'sheet1',
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
			sheet: 'Sheet1',
			seriesIndex: 0,
			categoryRef: 'Sheet1!$A$2:$A$10',
			valueRef: 'Sheet1!$C$2:$C$10',
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(wb.chartParts[0]?.series[0]).toMatchObject({
			nameRef: 'Sheet1!$B$1',
			categoryRef: 'Sheet1!$A$2:$A$10',
			valueRef: 'Sheet1!$C$2:$C$10',
		})
	})

	test('setChartSeriesSource preserves chart identity and untouched series metadata', () => {
		const wb = setup()
		wb.chartParts.push(
			{
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				chartType: 'lineChart',
				title: 'Revenue',
				series: [
					{
						nameText: 'Actual',
						categoryRef: 'Sheet1!$A$2:$A$4',
						valueRef: 'Sheet1!$B$2:$B$4',
					},
					{
						nameText: 'Plan',
						categoryRef: 'Sheet1!$A$2:$A$4',
						valueRef: 'Sheet1!$C$2:$C$4',
					},
				],
			},
			{
				partPath: 'xl/charts/chart2.xml',
				sheetName: 'Sheet2',
				chartType: 'barChart',
				series: [{ valueRef: 'Sheet2!$B$2:$B$4' }],
			},
		)

		const result = applyOperation(wb, {
			op: 'setChartSeriesSource',
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
			valueRef: 'Sheet1!$D$2:$D$4',
		})
		expectOk(result)

		expect(wb.chartParts).toEqual([
			{
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				chartType: 'lineChart',
				title: 'Revenue',
				series: [
					{
						nameText: 'Actual',
						categoryRef: 'Sheet1!$A$2:$A$4',
						valueRef: 'Sheet1!$D$2:$D$4',
					},
					{
						nameText: 'Plan',
						categoryRef: 'Sheet1!$A$2:$A$4',
						valueRef: 'Sheet1!$C$2:$C$4',
					},
				],
			},
			{
				partPath: 'xl/charts/chart2.xml',
				sheetName: 'Sheet2',
				chartType: 'barChart',
				series: [{ valueRef: 'Sheet2!$B$2:$B$4' }],
			},
		])
	})

	test('setChartSeriesSource rejects ambiguous chart selectors', () => {
		const wb = setup()
		wb.chartParts.push(
			{
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				series: [{ valueRef: 'Sheet1!$B$2:$B$4' }],
			},
			{
				partPath: 'xl/charts/chart2.xml',
				sheetName: 'Sheet1',
				series: [{ valueRef: 'Sheet1!$C$2:$C$4' }],
			},
		)

		const result = applyOperation(wb, {
			op: 'setChartSeriesSource',
			sheet: 'Sheet1',
			seriesIndex: 0,
			valueRef: 'Sheet1!$D$2:$D$4',
		})

		expectErr(result)
		expect(result.error.message).toContain('matched 2 charts')
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

	test('setPageSetup rejects invalid public metadata before mutation', () => {
		const cases: readonly Operation[] = [
			{
				op: 'setPageSetup',
				sheet: 'Sheet1',
				setup: { orientation: 'sideways' },
			} as unknown as Operation,
			{ op: 'setPageSetup', sheet: 'Sheet1', setup: { paperSize: -1 } },
			{ op: 'setPageSetup', sheet: 'Sheet1', setup: { scale: 0 } },
			{ op: 'setPageSetup', sheet: 'Sheet1', setup: { fitToWidth: 1.5 } },
			{ op: 'setPageSetup', sheet: 'Sheet1', setup: { margins: { left: -0.1 } } },
		]

		for (const op of cases) {
			const wb = setup()
			const result = applyOperation(wb, op)
			expectErr(result)
			expect(result.error.code, JSON.stringify(op)).toBe('VALIDATION_ERROR')
			const sheet = wb.getSheet('Sheet1')
			expect(sheet?.pageSetup, JSON.stringify(op)).toBeNull()
			expect(sheet?.pageMargins, JSON.stringify(op)).toBeNull()
		}
	})

	test('setPrintArea escapes quoted sheet names in defined-name metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet("Bob's Budget")

		const result = applyOperation(wb, {
			op: 'setPrintArea',
			sheet: "Bob's Budget",
			range: 'A1:B2',
		})
		expectOk(result)

		expect(wb.definedNames.resolve('_xlnm.Print_Area', sheet.id, sheet.id)?.formula).toBe(
			"'Bob''s Budget'!A1:B2",
		)
	})

	test('setPrintArea rejects invalid range metadata', () => {
		const wb = setup()

		const result = applyOperation(wb, {
			op: 'setPrintArea',
			sheet: 'Sheet1',
			range: 'not a range',
		})

		expectErr(result)
		expect(result.error.code).toBe('INVALID_RANGE')
		expect(wb.definedNames.has('_xlnm.Print_Area')).toBe(false)
	})

	test('setPageSetup preserves imported print metadata on partial updates', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.pageSetup = { orientation: 'portrait', paperSize: 1, firstPageNumber: 3 }
		sheet.pageMargins = { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 }

		const result = applyOperation(wb, {
			op: 'setPageSetup',
			sheet: 'Sheet1',
			setup: {
				orientation: 'landscape',
				margins: { left: 0.25 },
			},
		})
		expectOk(result)

		expect(sheet.pageSetup).toEqual({
			orientation: 'landscape',
			paperSize: 1,
			firstPageNumber: 3,
		})
		expect(sheet.pageMargins).toEqual({ left: 0.25, right: 0.5, top: 0.75, bottom: 0.75 })
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

	test('sheet metadata range operations reject invalid ranges before mutation', () => {
		const cases: readonly Operation[] = [
			{ op: 'setDataValidation', sheet: 'Sheet1', range: 'not a range', rule: { type: 'whole' } },
			{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'not a range' },
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'not a range',
				rule: { type: 'expression', formula: 'TRUE' },
			},
			{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'not a range' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'not a range' },
		]

		for (const op of cases) {
			const wb = setup()
			const result = applyOperation(wb, op)
			expectErr(result)
			expect(result.error.code, op.op).toBe('INVALID_RANGE')
			const sheet = wb.getSheet('Sheet1')
			expect(sheet?.dataValidations, op.op).toEqual([])
			expect(sheet?.conditionalFormats, op.op).toEqual([])
			expect(sheet?.autoFilter, op.op).toBeNull()
		}
	})

	test('setAutoFilter preserves existing criteria when updating the range', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) return
		sheet.autoFilter = {
			ref: 'A1:E22',
			columns: [{ colId: 0, kind: 'filters', values: ['1'] }],
			sortState: {
				ref: 'A1:E22',
				conditions: [{ ref: 'A2:A22', descending: true }],
			},
		}

		const result = applyOperation(wb, {
			op: 'setAutoFilter',
			sheet: 'Sheet1',
			range: 'A1:E30',
		})
		expectOk(result)

		expect(sheet.autoFilter).toEqual({
			ref: 'A1:E30',
			columns: [{ colId: 0, kind: 'filters', values: ['1'] }],
			sortState: {
				ref: 'A1:E22',
				conditions: [{ ref: 'A2:A22', descending: true }],
			},
		})
	})

	test('setAutoFilter edits value-list criteria and sort metadata', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setAutoFilter',
			sheet: 'Sheet1',
			range: 'A1:E22',
			column: 4,
			values: ['3'],
			sortRef: 'A1:E22',
			sortBy: 'E2:E22',
			descending: true,
		})
		expectOk(result)

		expect(wb.getSheet('Sheet1')?.autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [{ colId: 4, kind: 'filters', values: ['3'] }],
			sortState: {
				ref: 'A1:E22',
				conditions: [{ ref: 'E2:E22', descending: true }],
			},
		})
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
		sheet.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'expression', formulas: ['A1>0'] }],
		})

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
		expect(sheet.conditionalFormats.at(-1)).toEqual({
			sqref: 'C3',
			rules: [{ type: 'expression', formulas: ['C3>0'] }],
		})
		expect(result.value.recalcRequired).toBe(false)
	})

	test('copyRange format-only paste preserves formula binding metadata on target members', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		sheet.cells.set(0, 1, { value: numberValue(12), formula: null, styleId: sourceStyle })
		addSpillGroup(sheet)

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'B1',
			target: 'A2',
			mode: 'formats',
		})
		expectOk(result)

		expect(sheet.cells.get(1, 0)).toEqual({
			value: numberValue(2),
			formula: null,
			styleId: sourceStyle,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!A1',
				ref: 'A1:A3',
				isAnchor: false,
			},
		})
		expect(result.value.recalcRequired).toBe(false)
	})

	test('copyRange style-only paste preserves non-spill formula binding metadata', () => {
		for (const mode of ['formats', 'styles'] as const) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
			sheet.cells.set(0, 3, { value: numberValue(12), formula: null, styleId: sourceStyle })
			addSharedFormulaGroup(sheet)
			sheet.cells.set(0, 1, cell(numberValue(10)))
			sheet.cells.set(1, 1, cell(numberValue(20)))

			const shared = applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'D1',
				target: 'A2',
				mode,
			})
			expectOk(shared)

			expect(shared.value.recalcRequired, mode).toBe(false)
			expect(shared.value.affectedCells, mode).toEqual(['A2'])
			expect(sheet.cells.get(1, 0)).toEqual({
				value: numberValue(40),
				formula: null,
				styleId: sourceStyle,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})

			addDataTableFormula(sheet)
			const table = applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'D1',
				target: 'C4',
				mode,
			})
			expectOk(table)

			expect(table.value.recalcRequired, mode).toBe(false)
			expect(table.value.affectedCells, mode).toEqual(['C4'])
			expect(sheet.cells.get(2, 2)?.formulaInfo).toEqual({
				kind: 'dataTable',
				ref: 'C3:C5',
				dt2D: false,
				dtr: true,
				r1: 'A1',
			})
			expect(sheet.cells.get(3, 2)?.styleId).toBe(sourceStyle)
		}
	})

	test('copyRange value paste materializes target formula bindings before overwriting a member', () => {
		const cases: readonly {
			readonly name: string
			readonly setup: (sheet: Sheet) => void
			readonly target: string
			readonly affectedCells: readonly string[]
			readonly assert: (sheet: Sheet) => void
		}[] = [
			{
				name: 'shared formula member',
				setup: (sheet) => {
					addSharedFormulaGroup(sheet)
					sheet.cells.set(0, 1, cell(numberValue(10)))
					sheet.cells.set(1, 1, cell(numberValue(20)))
				},
				target: 'A2',
				affectedCells: ['A1', 'A2'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
					expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(99))
					expect(sheet.cells.get(1, 0)?.formula).toBeNull()
					expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
				},
			},
			{
				name: 'dynamic spill member',
				setup: addDynamicArrayAnchorWithStaleSpillFootprint,
				target: 'A2',
				affectedCells: ['A1', 'A2', 'A3'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
					expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(99))
					expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
				},
			},
		]

		for (const entry of cases) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			entry.setup(sheet)
			sheet.cells.set(0, 3, cell(numberValue(99)))

			const result = applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'D1',
				target: entry.target,
				mode: 'values',
			})
			expectOk(result)

			expect(result.value.affectedCells, entry.name).toEqual(entry.affectedCells)
			entry.assert(sheet)
		}
	})

	test('copyRange formula paste modes materialize target formula bindings before overwriting a member', () => {
		for (const mode of ['formulas', 'all'] as const) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			addSharedFormulaGroup(sheet)
			sheet.cells.set(0, 1, cell(numberValue(10)))
			sheet.cells.set(1, 1, cell(numberValue(20)))
			sheet.cells.set(0, 3, {
				value: numberValue(11),
				formula: '10+1',
				styleId: sid,
			})

			const result = applyOperation(wb, {
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'D1',
				target: 'A2',
				mode,
			})
			expectOk(result)

			expect(result.value.affectedCells, mode).toEqual(['A1', 'A2'])
			expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
			expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
			expect(sheet.cells.get(1, 0)?.formula).toBe('10+1')
			expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		}
	})

	test('copyRange value paste detaches data-table metadata on overwritten target members', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		sheet.cells.set(0, 0, cell(numberValue(7)))
		sheet.cells.set(0, 1, { value: numberValue(12), formula: null, styleId: sourceStyle })
		addDataTableFormula(sheet)

		const formats = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'B1',
			target: 'C4',
			mode: 'formats',
		})
		expectOk(formats)
		expect(sheet.cells.get(2, 2)?.formulaInfo).toEqual({
			kind: 'dataTable',
			ref: 'C3:C5',
			dt2D: false,
			dtr: true,
			r1: 'A1',
		})

		const values = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1',
			target: 'C4',
			mode: 'values',
		})
		expectOk(values)

		expect(values.value.affectedCells).toEqual(['C3', 'C4'])
		expect(values.value.recalcRequired).toBe(true)
		expect(sheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(7))
	})

	test('copyRange value paste detaches blocked-spill metadata on overwritten blockers', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addBlockedSpillFormula(sheet)
		sheet.cells.set(0, 1, { value: numberValue(99), formula: null, styleId: sid })

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'B1',
			target: 'A2',
			mode: 'values',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(99))
	})

	test('moveRange materializes imported formula bindings before moving one member', () => {
		const cases: readonly {
			readonly name: string
			readonly setup: (sheet: Sheet) => void
			readonly op: Extract<Operation, { op: 'moveRange' }>
			readonly affectedCells: readonly string[]
			readonly assert: (sheet: Sheet) => void
		}[] = [
			{
				name: 'shared formula member',
				setup: (sheet) => {
					addSharedFormulaGroup(sheet)
					sheet.cells.set(0, 1, cell(numberValue(10)))
					sheet.cells.set(1, 1, cell(numberValue(20)))
				},
				op: { op: 'moveRange', sheet: 'Sheet1', source: 'A2', target: 'C2', mode: 'all' },
				affectedCells: ['A1', 'A2', 'C2'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
					expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 0)).toBeUndefined()
					expect(sheet.cells.get(1, 2)?.formula).toBe('D2*2')
					expect(sheet.cells.get(1, 2)?.formulaInfo).toBeUndefined()
				},
			},
			{
				name: 'dynamic spill member',
				setup: addDynamicArrayAnchorWithStaleSpillFootprint,
				op: { op: 'moveRange', sheet: 'Sheet1', source: 'A2', target: 'C2', mode: 'all' },
				affectedCells: ['A1', 'A2', 'A3', 'C2'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 0)).toBeUndefined()
					expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
					expect(sheet.cells.get(1, 2)?.formulaInfo).toBeUndefined()
				},
			},
			{
				name: 'data-table member',
				setup: addDataTableFormula,
				op: { op: 'moveRange', sheet: 'Sheet1', source: 'C4', target: 'E4', mode: 'all' },
				affectedCells: ['C3', 'E4', 'C4'],
				assert: (sheet) => {
					expect(sheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(3, 2)).toBeUndefined()
					expect(sheet.cells.get(3, 4)?.value).toEqual(numberValue(20))
					expect(sheet.cells.get(4, 2)?.value).toEqual(numberValue(30))
				},
			},
		]

		for (const entry of cases) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			entry.setup(sheet)

			const result = applyOperation(wb, entry.op)
			expectOk(result)

			expect(result.value.affectedCells, entry.name).toEqual(entry.affectedCells)
			entry.assert(sheet)
		}
	})

	test('moveRange materializes target formula bindings before overwriting a member', () => {
		const cases: readonly {
			readonly name: string
			readonly setup: (sheet: Sheet) => void
			readonly affectedCells: readonly string[]
			readonly assert: (sheet: Sheet) => void
		}[] = [
			{
				name: 'shared formula member',
				setup: (sheet) => {
					addSharedFormulaGroup(sheet)
					sheet.cells.set(0, 1, cell(numberValue(10)))
					sheet.cells.set(1, 1, cell(numberValue(20)))
				},
				affectedCells: ['A1', 'A2', 'D1'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
					expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(99))
					expect(sheet.cells.get(1, 0)?.formula).toBeNull()
					expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(0, 3)).toBeUndefined()
				},
			},
			{
				name: 'dynamic spill member',
				setup: addDynamicArrayAnchorWithStaleSpillFootprint,
				affectedCells: ['A1', 'A2', 'A3', 'D1'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
					expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(99))
					expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(0, 3)).toBeUndefined()
				},
			},
		]

		for (const entry of cases) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			entry.setup(sheet)
			sheet.cells.set(0, 3, cell(numberValue(99)))

			const result = applyOperation(wb, {
				op: 'moveRange',
				sheet: 'Sheet1',
				source: 'D1',
				target: 'A2',
				mode: 'all',
			})
			expectOk(result)

			expect(result.value.affectedCells, entry.name).toEqual(entry.affectedCells)
			entry.assert(sheet)
		}
	})

	test('moveRange formula paste modes materialize target formula bindings before overwriting a member', () => {
		for (const mode of ['values', 'formulas'] as const) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			addSharedFormulaGroup(sheet)
			sheet.cells.set(0, 1, cell(numberValue(10)))
			sheet.cells.set(1, 1, cell(numberValue(20)))
			sheet.cells.set(0, 3, {
				value: numberValue(11),
				formula: '10+1',
				styleId: sid,
			})

			const result = applyOperation(wb, {
				op: 'moveRange',
				sheet: 'Sheet1',
				source: 'D1',
				target: 'A2',
				mode,
			})
			expectOk(result)

			expect(result.value.affectedCells, mode).toEqual(['A1', 'A2', 'D1'])
			expect(sheet.cells.get(0, 0)?.formula).toBe('B1*2')
			expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
			expect(sheet.cells.get(1, 0)?.formula).toBe(mode === 'formulas' ? '10+1' : null)
			expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
			expect(sheet.cells.get(0, 3)).toBeUndefined()
		}
	})

	test('copyRange exposes comments, hyperlinks, and validation paste modes', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(1)))
		sheet.cells.set(0, 1, cell(numberValue(2)))
		sheet.comments.set('A1', {
			text: 'Review',
			author: 'Ascend',
			legacyDrawing: { shapeId: '_x0000_s1025', row: 0, column: 0 },
		})
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

		expect(sheet.comments.get('C1')).toEqual({
			text: 'Review',
			author: 'Ascend',
			legacyDrawing: { shapeId: '_x0000_s1025', row: 0, column: 2 },
		})
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

	test('copyRange comments mode retargets legacy VML metadata and clones threaded comments', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.comments.set('A1', {
			text: 'Review',
			author: 'Ascend',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 0,
				column: 0,
				anchor: [0, 0, 0, 0, 1, 0, 3, 0],
			},
		})
		sheet.threadedComments.push(
			{ ref: 'A1', text: 'Root', id: 'tc1', author: 'Ada' },
			{ ref: 'A1', text: 'Reply', id: 'tc2', parentId: 'tc1', author: 'Grace' },
		)

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1',
			target: 'C3',
			mode: 'comments',
		})
		expectOk(result)

		expect(sheet.comments.get('C3')).toEqual({
			text: 'Review',
			author: 'Ascend',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 2,
				column: 2,
				anchor: [2, 0, 2, 0, 3, 0, 5, 0],
			},
		})
		expect(sheet.threadedComments).toEqual([
			{ ref: 'A1', text: 'Root', id: 'tc1', author: 'Ada' },
			{ ref: 'A1', text: 'Reply', id: 'tc2', parentId: 'tc1', author: 'Grace' },
			{ ref: 'C3', text: 'Root', id: 'tc1-copy', author: 'Ada' },
			{ ref: 'C3', text: 'Reply', id: 'tc2-copy', parentId: 'tc1-copy', author: 'Grace' },
		])
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

	test('moveRange format-only moves styles without deleting source formulas or bindings', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		const targetStyle = wb.styles.register({ numberFormat: '0.0%' })
		addSpillGroup(sheet)
		for (const row of [0, 1, 2]) {
			const existing = sheet.cells.get(row, 0)
			if (!existing) throw new Error('missing spill source')
			sheet.cells.set(row, 0, { ...existing, styleId: sourceStyle })
			sheet.cells.set(row, 2, {
				value: numberValue(99 + row),
				formula: row === 0 ? 'A1+1' : null,
				styleId: targetStyle,
			})
		}
		sheet.cells.set(5, 5, {
			value: numberValue(3),
			formula: 'SUM(A1:A2)',
			styleId: sid,
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:A3',
			target: 'C1',
			mode: 'formats',
		})
		expectOk(result)

		expect(result.value.recalcRequired).toBe(false)
		expect([...result.value.affectedCells].sort()).toEqual(['A1', 'A2', 'A3', 'C1', 'C2', 'C3'])
		expect(sheet.cells.get(0, 0)).toEqual({
			value: numberValue(1),
			formula: 'SEQUENCE(3)',
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!A1',
				ref: 'A1:A3',
				isAnchor: true,
			},
		})
		expect(sheet.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: false,
		})
		expect(sheet.cells.get(0, 2)).toEqual({
			value: numberValue(99),
			formula: 'A1+1',
			styleId: sourceStyle,
		})
		expect(sheet.cells.get(5, 5)?.formula).toBe('SUM(A1:A2)')
	})

	test('moveRange style-only modes preserve non-spill formula binding metadata', () => {
		for (const mode of ['formats', 'styles'] as const) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
			addSharedFormulaGroup(sheet)
			sheet.cells.set(0, 1, cell(numberValue(10)))
			sheet.cells.set(1, 1, cell(numberValue(20)))
			const source = sheet.cells.get(1, 0)
			if (!source) throw new Error('missing shared source')
			sheet.cells.set(1, 0, { ...source, styleId: sourceStyle })
			addDataTableFormula(sheet)

			const result = applyOperation(wb, {
				op: 'moveRange',
				sheet: 'Sheet1',
				source: 'A2',
				target: 'C4',
				mode,
			})
			expectOk(result)

			expect(result.value.recalcRequired, mode).toBe(false)
			expect(result.value.affectedCells, mode).toEqual(['C4', 'A2'])
			expect(sheet.cells.get(1, 0)).toEqual({
				value: numberValue(40),
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})
			expect(sheet.cells.get(2, 2)?.formulaInfo).toEqual({
				kind: 'dataTable',
				ref: 'C3:C5',
				dt2D: false,
				dtr: true,
				r1: 'A1',
			})
			expect(sheet.cells.get(3, 2)?.styleId).toBe(sourceStyle)
		}
	})

	test('moveRange detaches blocked-spill metadata when moving a blocker', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addBlockedSpillFormula(sheet)

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A2',
			target: 'C1',
			mode: 'values',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'C1', 'A2'])
		expect(sheet.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)).toBeUndefined()
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('blocker'))
	})

	test('moveRange rewrites formulas and names that reference the moved range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(10)))
		sheet.cells.set(1, 0, cell(numberValue(20)))
		sheet.cells.set(0, 1, cell(EMPTY, 'A1*2'))
		sheet.cells.set(1, 1, cell(EMPTY, 'SUM(A1:A2)'))
		wb.definedNames.set('Input', 'Sheet1!A1')

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:A2',
			target: 'C1',
		})
		expectOk(result)

		expect(sheet.cells.get(0, 1)?.formula).toBe('C1*2')
		expect(sheet.cells.get(1, 1)?.formula).toBe('SUM(C1:C2)')
		expect(wb.definedNames.get('Input')).toBe('Sheet1!C1')
		expect(result.value.affectedCells).toContain('B1')
		expect(result.value.affectedCells).toContain('B2')
	})

	test('moveRange reports every shared formula member when rewriting the shared master', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(10)))
		sheet.cells.set(1, 0, cell(numberValue(20)))
		sheet.cells.set(0, 1, {
			value: numberValue(20),
			formula: 'A1*2',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'move-rewrite',
				isMaster: true,
				masterRef: 'Sheet1!B1',
				ref: 'B1:B2',
			},
		})
		sheet.cells.set(1, 1, {
			value: numberValue(40),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'move-rewrite',
				isMaster: false,
				masterRef: 'sheet1!$B$1',
			},
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:A2',
			target: 'C1',
		})
		expectOk(result)

		expect(sheet.cells.get(0, 1)?.formula).toBe('C1*2')
		expect(sheet.cells.get(1, 1)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: 'move-rewrite',
			isMaster: false,
			masterRef: 'sheet1!$B$1',
		})
		expect(result.value.affectedCells).toContain('B1')
		expect(result.value.affectedCells).toContain('B2')
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('moveRange reports every dynamic spill member when rewriting the anchor', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(3)))
		sheet.cells.set(0, 1, {
			value: numberValue(1),
			formula: 'SEQUENCE(A1)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		for (let row = 1; row <= 2; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(row + 1),
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'spill', anchorRef: 'Sheet1!B1', ref: 'B1:B3', isAnchor: false },
			})
		}

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1',
			target: 'C1',
		})
		expectOk(result)

		expect(sheet.cells.get(0, 1)?.formula).toBe('SEQUENCE(C1)')
		expect(result.value.affectedCells).toContain('B1')
		expect(result.value.affectedCells).toContain('B2')
		expect(result.value.affectedCells).toContain('B3')
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('moveRange rewrites cross-sheet formulas and names to the target sheet', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const target = wb.addSheet('Sheet2')
		const summary = wb.addSheet('Summary')
		source.cells.set(0, 0, cell(numberValue(10)))
		summary.cells.set(0, 0, cell(EMPTY, 'sheet1!A1+1'))
		target.cells.set(4, 4, cell(EMPTY, 'Sheet1!A1*2'))
		wb.definedNames.set('Input', 'Sheet1!A1')

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1',
			targetSheet: 'Sheet2',
			target: 'B2',
		})
		expectOk(result)

		expect(summary.cells.get(0, 0)?.formula).toBe('Sheet2!B2+1')
		expect(target.cells.get(4, 4)?.formula).toBe('Sheet2!B2*2')
		expect(wb.definedNames.get('Input')).toBe('Sheet2!B2')
		expect(result.value.affectedCells).toContain('Summary!A1')
		expect(result.value.affectedCells).toContain('Sheet2!E5')
	})

	test('moveRange rewrites worksheet metadata formulas that reference the moved range', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		wb.addSheet('Sheet2')
		const summary = wb.addSheet('Summary')
		source.cells.set(0, 0, cell(numberValue(10)))
		summary.dataValidations.push({ sqref: 'A1', type: 'list', formula1: 'Sheet1!A1' })
		summary.conditionalFormats.push({
			sqref: 'B1',
			rules: [
				{
					type: 'expression',
					formulas: ['Sheet1!A1>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'Sheet1!A1' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
				},
			],
		})
		summary.x14DataValidations.push({
			index: 0,
			sqref: 'C1',
			type: 'list',
			formula1: 'Sheet1!A1',
		})
		summary.x14ConditionalFormats.push({
			index: 0,
			sqref: 'D1',
			formulas: ['Sheet1!A1>0'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'Sheet1!A1' }],
				colors: [{ rgb: 'FF63BE7B' }],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'Sheet1!A1' }] },
		})
		summary.tables.push({
			id: createTableId(),
			name: 'SummaryTable',
			sheetId: summary.id,
			ref: { start: { row: 0, col: 5 }, end: { row: 1, col: 6 } },
			columns: [
				{
					name: 'Metric',
					formula: 'Sheet1!A1',
					totalsRowFormula: 'SUM(Sheet1!A1)',
				},
				{ name: 'Value' },
			],
			hasHeaders: true,
			hasTotals: true,
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1',
			targetSheet: 'Sheet2',
			target: 'B2',
		})
		expectOk(result)

		expect(summary.dataValidations[0]?.formula1).toBe('Sheet2!B2')
		expect(summary.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('Sheet2!B2>0')
		expect(summary.conditionalFormats[0]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe('Sheet2!B2')
		expect(summary.x14DataValidations[0]?.formula1).toBe('Sheet2!B2')
		expect(summary.x14ConditionalFormats[0]?.formulas[0]).toBe('Sheet2!B2>0')
		expect(summary.x14ConditionalFormats[0]?.colorScale?.cfvo[0]?.value).toBe('Sheet2!B2')
		expect(summary.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('Sheet2!B2')
		expect(summary.tables[0]?.columns[0]?.formula).toBe('Sheet2!B2')
		expect(summary.tables[0]?.columns[0]?.totalsRowFormula).toBe('SUM(Sheet2!B2)')
	})

	test('moveRange rejects cell formulas with partially moved range references before mutation', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(1)))
		sheet.cells.set(1, 0, cell(numberValue(2)))
		sheet.cells.set(2, 0, cell(numberValue(3)))
		sheet.cells.set(0, 3, cell(EMPTY, 'SUM(A1:A3)'))

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A2',
			target: 'C2',
		})

		expectErr(result)
		expect(result.error.message).toContain('partially overlaps the moved cells')
		expect(result.error.details).toMatchObject({
			kind: 'partial-move-formula-reference',
			ownerKind: 'cell-formula',
			owner: 'Sheet1!D1',
			reference: 'A1:A3',
			source: 'A2',
		})
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 2)).toBeUndefined()
		expect(sheet.cells.get(0, 3)?.formula).toBe('SUM(A1:A3)')
	})

	test('moveRange rejects cross-sheet formulas with partially moved range references', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const target = wb.addSheet('Sheet2')
		const summary = wb.addSheet('Summary')
		source.cells.set(0, 0, cell(numberValue(1)))
		source.cells.set(1, 0, cell(numberValue(2)))
		source.cells.set(2, 0, cell(numberValue(3)))
		summary.cells.set(0, 0, cell(EMPTY, 'SUM(Sheet1!A1:A3)'))

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A2',
			targetSheet: 'Sheet2',
			target: 'C2',
		})

		expectErr(result)
		expect(result.error.details).toMatchObject({
			kind: 'partial-move-formula-reference',
			owner: 'Summary!A1',
			reference: 'Sheet1!A1:A3',
		})
		expect(source.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(target.cells.get(1, 2)).toBeUndefined()
		expect(summary.cells.get(0, 0)?.formula).toBe('SUM(Sheet1!A1:A3)')
	})

	test('moveRange rejects hyperlink locations with partially moved range references', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const summary = wb.addSheet('Summary')
		source.cells.set(0, 0, cell(numberValue(1)))
		source.cells.set(1, 0, cell(numberValue(2)))
		source.cells.set(2, 0, cell(numberValue(3)))
		summary.hyperlinks.set('B1', {
			location: 'Sheet1!A1:A3',
			display: 'Jump',
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A2',
			target: 'C2',
		})

		expectErr(result)
		expect(result.error.details).toMatchObject({
			kind: 'partial-move-formula-reference',
			ownerKind: 'hyperlink-location',
			owner: 'Summary!hyperlink(B1).location',
			reference: 'Sheet1!A1:A3',
			source: 'A2',
		})
		expect(source.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(source.cells.get(1, 2)).toBeUndefined()
		expect(summary.hyperlinks.get('B1')?.location).toBe('Sheet1!A1:A3')
	})

	test('moveRange rejects defined names and worksheet metadata formulas with partial moved range refs', () => {
		const cases: readonly {
			readonly label: string
			readonly setup: (wb: Workbook, source: Sheet, summary: Sheet) => void
			readonly ownerKind: string
			readonly ownerIncludes: string
		}[] = [
			{
				label: 'defined name',
				setup: (wb) => wb.definedNames.set('Totals', 'Sheet1!A1:A3'),
				ownerKind: 'defined-name',
				ownerIncludes: 'Totals',
			},
			{
				label: 'data validation formula',
				setup: (_wb, _source, summary) =>
					summary.dataValidations.push({
						sqref: 'A1',
						type: 'list',
						formula1: 'SUM(Sheet1!A1:A3)',
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'dataValidation',
			},
			{
				label: 'conditional format formula',
				setup: (_wb, _source, summary) =>
					summary.conditionalFormats.push({
						sqref: 'B1',
						rules: [{ type: 'expression', formulas: ['SUM(Sheet1!A1:A3)>0'] }],
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'conditionalFormat',
			},
			{
				label: 'conditional format cfvo',
				setup: (_wb, _source, summary) =>
					summary.conditionalFormats.push({
						sqref: 'C1',
						rules: [
							{
								type: 'colorScale',
								formulas: [],
								colorScale: {
									cfvo: [{ type: 'formula', value: 'SUM(Sheet1!A1:A3)' }],
									colors: [{ rgb: 'FFFF0000' }],
								},
							},
						],
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'colorScale.cfvo',
			},
			{
				label: 'x14 data validation formula',
				setup: (_wb, _source, summary) =>
					summary.x14DataValidations.push({
						index: 0,
						sqref: 'D1',
						type: 'list',
						formula1: 'SUM(Sheet1!A1:A3)',
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'x14DataValidation',
			},
			{
				label: 'x14 conditional format formula',
				setup: (_wb, _source, summary) =>
					summary.x14ConditionalFormats.push({
						index: 0,
						sqref: 'E1',
						formulas: ['SUM(Sheet1!A1:A3)>0'],
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'x14ConditionalFormat',
			},
			{
				label: 'x14 conditional format cfvo',
				setup: (_wb, _source, summary) =>
					summary.x14ConditionalFormats.push({
						index: 0,
						sqref: 'F1',
						formulas: [],
						dataBar: { cfvo: [{ type: 'formula', value: 'SUM(Sheet1!A1:A3)' }] },
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'dataBar.cfvo',
			},
			{
				label: 'table column formula',
				setup: (_wb, _source, summary) =>
					summary.tables.push({
						id: createTableId(),
						name: 'SummaryTable',
						sheetId: summary.id,
						ref: { start: { row: 0, col: 7 }, end: { row: 2, col: 8 } },
						columns: [{ name: 'Name' }, { name: 'Calc', formula: 'SUM(Sheet1!A1:A3)' }],
						hasHeaders: true,
						hasTotals: false,
					}),
				ownerKind: 'worksheet-metadata',
				ownerIncludes: 'table(SummaryTable)',
			},
		]

		for (const scenario of cases) {
			const wb = createWorkbook()
			const source = wb.addSheet('Sheet1')
			const summary = wb.addSheet('Summary')
			source.cells.set(0, 0, cell(numberValue(1)))
			source.cells.set(1, 0, cell(numberValue(2)))
			source.cells.set(2, 0, cell(numberValue(3)))
			scenario.setup(wb, source, summary)

			const result = applyOperation(wb, {
				op: 'moveRange',
				sheet: 'Sheet1',
				source: 'A2',
				target: 'C2',
			})

			expectErr(result)
			expect(result.error.details).toMatchObject({
				kind: 'partial-move-formula-reference',
				ownerKind: scenario.ownerKind,
				reference: 'Sheet1!A1:A3',
				source: 'A2',
			})
			expect(String(result.error.details?.owner)).toContain(scenario.ownerIncludes)
			expect(source.cells.get(1, 0)?.value, scenario.label).toEqual(numberValue(2))
			expect(source.cells.get(1, 2), scenario.label).toBeUndefined()
		}
	})

	test('copyRange copies merged-cell layout and replaces covered target merges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(1)))
		sheet.cells.set(0, 1, cell(numberValue(2)))
		sheet.merges.push({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 1 },
		})
		sheet.merges.push({
			start: { row: 2, col: 2 },
			end: { row: 2, col: 3 },
		})

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			target: 'C3',
		})
		expectOk(result)

		expect(sheet.merges).toEqual([
			{
				start: { row: 0, col: 0 },
				end: { row: 0, col: 1 },
			},
			{
				start: { row: 2, col: 2 },
				end: { row: 2, col: 3 },
			},
		])
	})

	test('moveRange relocates merged-cell layout and removes source merges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(numberValue(1)))
		sheet.cells.set(0, 1, cell(numberValue(2)))
		sheet.merges.push({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 1 },
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			target: 'A3',
		})
		expectOk(result)

		expect(sheet.merges).toEqual([
			{
				start: { row: 2, col: 0 },
				end: { row: 2, col: 1 },
			},
		])
	})

	test('copyRange rejects partial merged-cell sources and target overlaps', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.merges.push({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 1 },
		})
		sheet.merges.push({
			start: { row: 2, col: 2 },
			end: { row: 3, col: 3 },
		})

		const partialSource = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1',
			target: 'E1',
		})
		expectErr(partialSource)
		expect(partialSource.error.message).toContain('part of a merged range')

		const partialTarget = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			target: 'D4',
		})
		expectErr(partialTarget)
		expect(partialTarget.error.message).toContain('partially overlaps')
	})

	test('copyRange can copy cells and layout metadata to another sheet', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const target = wb.addSheet('Sheet2')
		const sourceStyle = wb.styles.register({ numberFormat: '$#,##0.00' })
		source.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sourceStyle })
		source.cells.set(0, 1, { value: numberValue(20), formula: 'A1*2', styleId: sourceStyle })
		source.comments.set('A1', {
			text: 'Review',
			author: 'Ascend',
			legacyDrawing: { shapeId: '_x0000_s1025', row: 0, column: 0 },
		})
		source.hyperlinks.set('B1', { target: 'https://example.com', display: 'Example' })
		source.dataValidations.push({ sqref: 'A1:B1', type: 'whole', formula1: 'A1' })
		source.conditionalFormats.push({
			sqref: 'A1:B1',
			rules: [
				{
					type: 'expression',
					formulas: ['A1>0'],
					colorScale: {
						cfvo: [{ type: 'min' }, { type: 'formula', value: 'B1' }],
						colors: [{ rgb: 'FFFF0000' }, { rgb: 'FF00FF00' }],
					},
				},
			],
		})
		source.merges.push({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 1 },
		})
		target.merges.push({
			start: { row: 2, col: 2 },
			end: { row: 2, col: 3 },
		})

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			targetSheet: 'Sheet2',
			target: 'C3',
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['Sheet1', 'Sheet2'])
		expect(result.value.affectedCells).toContain('Sheet2!C3')
		expect(target.cells.get(2, 2)?.value).toEqual(numberValue(10))
		expect(target.cells.get(2, 3)?.formula).toBe('C3*2')
		expect(target.cells.get(2, 3)?.styleId).toBe(sourceStyle)
		expect(target.comments.get('C3')).toEqual({
			text: 'Review',
			author: 'Ascend',
			legacyDrawing: { shapeId: '_x0000_s1025', row: 2, column: 2 },
		})
		expect(target.hyperlinks.get('D3')).toEqual({
			target: 'https://example.com',
			display: 'Example',
		})
		expect(target.dataValidations.at(-1)).toEqual({
			sqref: 'C3:D3',
			type: 'whole',
			formula1: 'C3',
		})
		expect(target.conditionalFormats.at(-1)).toEqual({
			sqref: 'C3:D3',
			rules: [
				{
					type: 'expression',
					formulas: ['C3>0'],
					colorScale: {
						cfvo: [{ type: 'min' }, { type: 'formula', value: 'D3' }],
						colors: [{ rgb: 'FFFF0000' }, { rgb: 'FF00FF00' }],
					},
				},
			],
		})
		expect(target.merges).toEqual([
			{
				start: { row: 2, col: 2 },
				end: { row: 2, col: 3 },
			},
		])
		expect(source.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(source.merges).toHaveLength(1)
	})

	test('moveRange can move cells and layout metadata to another sheet', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const target = wb.addSheet('Sheet2')
		source.cells.set(0, 0, cell(numberValue(10)))
		source.cells.set(0, 1, cell(numberValue(20)))
		source.comments.set('A1', {
			text: 'Review',
			legacyDrawing: { shapeId: '_x0000_s1026', row: 0, column: 0 },
		})
		source.threadedComments.push({ ref: 'A1', text: 'Thread', id: 'tc1' })
		source.hyperlinks.set('B1', { location: 'Sheet1!A1', display: 'Jump' })
		source.dataValidations.push({ sqref: 'A1:B1', type: 'whole', formula1: 'A1' })
		source.conditionalFormats.push({
			sqref: 'A1:B1',
			rules: [{ type: 'expression', formulas: ['A1>0'] }],
		})
		source.merges.push({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 1 },
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			targetSheet: 'Sheet2',
			target: 'A2',
		})
		expectOk(result)

		expect(source.cells.get(0, 0)).toBeUndefined()
		expect(source.cells.get(0, 1)).toBeUndefined()
		expect(source.comments.size).toBe(0)
		expect(source.threadedComments).toEqual([])
		expect(source.hyperlinks.size).toBe(0)
		expect(source.dataValidations).toEqual([])
		expect(source.conditionalFormats).toEqual([])
		expect(source.merges).toEqual([])
		expect(target.cells.get(1, 0)?.value).toEqual(numberValue(10))
		expect(target.cells.get(1, 1)?.value).toEqual(numberValue(20))
		expect(target.comments.get('A2')).toEqual({
			text: 'Review',
			legacyDrawing: { shapeId: '_x0000_s1026', row: 1, column: 0 },
		})
		expect(target.threadedComments).toEqual([{ ref: 'A2', text: 'Thread', id: 'tc1' }])
		expect(target.hyperlinks.get('B2')).toEqual({ location: 'Sheet2!A2', display: 'Jump' })
		expect(target.dataValidations).toEqual([{ sqref: 'A2:B2', type: 'whole', formula1: 'A2' }])
		expect(target.conditionalFormats).toEqual([
			{ sqref: 'A2:B2', rules: [{ type: 'expression', formulas: ['A2>0'] }] },
		])
		expect(target.merges).toEqual([
			{
				start: { row: 1, col: 0 },
				end: { row: 1, col: 1 },
			},
		])
		expect(result.value.affectedCells).toContain('Sheet1!A1')
		expect(result.value.affectedCells).toContain('Sheet2!A2')
	})

	test('moveRange rewrites internal hyperlink locations that point into moved cells', () => {
		const wb = createWorkbook()
		const source = wb.addSheet("Bob's Budget")
		wb.addSheet('Review!')
		source.cells.set(0, 0, cell(numberValue(10)))
		source.hyperlinks.set('C1', {
			location: "'Bob''s Budget'!A1",
			display: 'Jump',
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: "Bob's Budget",
			source: 'A1',
			targetSheet: 'Review!',
			target: 'B2',
		})
		expectOk(result)

		expect(source.hyperlinks.get('C1')).toEqual({
			location: "'Review!'!B2",
			display: 'Jump',
		})
		expect(result.value.affectedCells).toContain("Bob's Budget!C1")
	})

	test('moveRange retargets explicit source-sheet metadata formulas moved to another sheet', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const target = wb.addSheet('Sheet2')
		source.cells.set(0, 0, cell(numberValue(10)))
		source.cells.set(0, 1, cell(numberValue(20)))
		source.dataValidations.push({
			sqref: 'A1:B1',
			type: 'whole',
			formula1: 'Sheet1!A1',
			formula2: 'Sheet1!B1',
		})
		source.conditionalFormats.push({
			sqref: 'A1:B1',
			rules: [
				{
					type: 'expression',
					formulas: ['Sheet1!A1>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'Sheet1!A1' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: { cfvo: [{ type: 'formula', value: 'Sheet1!A1' }] },
					iconSet: { cfvo: [{ type: 'formula', value: 'Sheet1!B1' }] },
				},
			],
		})
		source.x14DataValidations.push({
			index: 0,
			sqref: 'A1:B1',
			type: 'whole',
			formula1: 'Sheet1!A1',
			formula2: 'Sheet1!B1',
		})
		source.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A1:B1',
			formulas: ['Sheet1!A1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'Sheet1!A1' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'Sheet1!B1' }] },
		})

		const result = applyOperation(wb, {
			op: 'moveRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			targetSheet: 'Sheet2',
			target: 'A2',
		})
		expectOk(result)

		expect(target.dataValidations[0]).toMatchObject({
			sqref: 'A2:B2',
			formula1: 'Sheet2!A2',
			formula2: 'Sheet2!B2',
		})
		const rule = target.conditionalFormats[0]?.rules[0]
		expect(target.conditionalFormats[0]?.sqref).toBe('A2:B2')
		expect(rule?.formulas[0]).toBe('Sheet2!A2>0')
		expect(rule?.colorScale?.cfvo[0]?.value).toBe('Sheet2!A2')
		expect(rule?.dataBar?.cfvo[0]?.value).toBe('Sheet2!A2')
		expect(rule?.iconSet?.cfvo[0]?.value).toBe('Sheet2!B2')
		expect(source.x14DataValidations[0]?.deleted).toBe(true)
		expect(source.x14ConditionalFormats[0]?.deleted).toBe(true)
		expect(target.x14DataValidations[0]).toMatchObject({
			sqref: 'A2:B2',
			formula1: 'Sheet2!A2',
			formula2: 'Sheet2!B2',
		})
		expect(target.x14ConditionalFormats[0]?.sqref).toBe('A2:B2')
		expect(target.x14ConditionalFormats[0]?.formulas[0]).toBe('Sheet2!A2>0')
		expect(target.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('Sheet2!A2')
		expect(target.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('Sheet2!B2')
	})

	test('copyRange copies x14 validation and conditional format formulas', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		const target = wb.addSheet('Sheet2')
		source.cells.set(0, 0, cell(numberValue(10)))
		source.cells.set(0, 1, cell(numberValue(20)))
		source.x14DataValidations.push({
			index: 3,
			sqref: 'A1:B1',
			type: 'whole',
			formula1: 'A1',
			formula2: 'B1',
		})
		source.x14ConditionalFormats.push({
			index: 4,
			sqref: 'A1:B1',
			formulas: ['A1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'A1' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'B1' }] },
		})

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A1:B1',
			targetSheet: 'Sheet2',
			target: 'C3',
		})
		expectOk(result)

		expect(source.x14DataValidations).toHaveLength(1)
		expect(source.x14ConditionalFormats).toHaveLength(1)
		expect(target.x14DataValidations).toEqual([
			{
				index: 0,
				sqref: 'C3:D3',
				type: 'whole',
				formula1: 'C3',
				formula2: 'D3',
			},
		])
		expect(target.x14ConditionalFormats).toEqual([
			{
				index: 0,
				sqref: 'C3:D3',
				formulas: ['C3>0'],
				dataBar: { cfvo: [{ type: 'formula', value: 'C3' }] },
				iconSet: { cfvo: [{ type: 'formula', value: 'D3' }] },
			},
		])
	})

	test('hideSheet hideRows and hideCols update sheet visibility metadata', () => {
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
		expect(wb.getSheet('Sheet1')?.colDefs).toContainEqual({ min: 1, max: 1, hidden: true })

		const result2b = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: false,
		})
		expectOk(result2b)

		const result3 = applyOperation(wb, {
			op: 'hideRows',
			sheet: 'Sheet1',
			at: 2,
			count: 1,
			hidden: true,
		})
		expectOk(result3)
		expect(wb.getSheet('Sheet1')?.rowDefs.get(2)).toEqual({ hidden: true })
		expect(wb.getSheet('Sheet1')?.rowHeights.get(2)).toBeUndefined()

		const result4 = applyOperation(wb, {
			op: 'hideRows',
			sheet: 'Sheet1',
			at: 2,
			count: 1,
			hidden: false,
		})
		expectOk(result4)

		const sheet = wb.getSheet('Sheet1')
		expect(sheet?.state).toBe('hidden')
		expect(sheet?.colDefs.find((def) => def.min === 1 && def.max === 1)).toBeUndefined()
		expect(sheet?.rowHeights.get(2)).toBeUndefined()
		expect(sheet?.rowDefs.get(2)).toBeUndefined()
	})

	test('hideCols splits and rejoins imported column definition ranges', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.colDefs.push({ min: 0, max: 2, width: 18, customWidth: true })

		const hidden = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: true,
		})
		expectOk(hidden)
		expect(sheet.colDefs).toEqual([
			{ min: 0, max: 0, width: 18, customWidth: true },
			{ min: 1, max: 1, width: 18, customWidth: true, hidden: true },
			{ min: 2, max: 2, width: 18, customWidth: true },
		])

		const unhidden = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: false,
		})
		expectOk(unhidden)
		expect(sheet.colDefs).toEqual([{ min: 0, max: 2, width: 18, customWidth: true }])
	})

	test('hideCols preserves public column widths when creating column definitions', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')

		const resized = applyOperation(wb, {
			op: 'setColWidth',
			sheet: 'Sheet1',
			col: 1,
			width: 12,
		})
		expectOk(resized)

		const hidden = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: true,
		})
		expectOk(hidden)
		expect(sheet.colWidths.get(1)).toBe(12)
		expect(sheet.colDefs).toEqual([{ min: 1, max: 1, width: 12, customWidth: true, hidden: true }])

		const unhidden = applyOperation(wb, {
			op: 'hideCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
			hidden: false,
		})
		expectOk(unhidden)
		expect(sheet.colWidths.get(1)).toBe(12)
		expect(sheet.colDefs).toEqual([{ min: 1, max: 1, width: 12, customWidth: true }])
	})

	test('setColWidth updates imported column definition ranges used by the writer', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.colDefs.push({ min: 0, max: 2, width: 18, customWidth: true })
		sheet.colWidths.set(0, 18)
		sheet.colWidths.set(1, 18)
		sheet.colWidths.set(2, 18)

		const resized = applyOperation(wb, {
			op: 'setColWidth',
			sheet: 'Sheet1',
			col: 1,
			width: 20,
		})
		expectOk(resized)
		expect(sheet.colWidths.get(1)).toBe(20)
		expect(sheet.colDefs).toEqual([
			{ min: 0, max: 0, width: 18, customWidth: true },
			{ min: 1, max: 1, width: 20, customWidth: true },
			{ min: 2, max: 2, width: 18, customWidth: true },
		])

		const restored = applyOperation(wb, {
			op: 'setColWidth',
			sheet: 'Sheet1',
			col: 1,
			width: 18,
		})
		expectOk(restored)
		expect(sheet.colDefs).toEqual([{ min: 0, max: 2, width: 18, customWidth: true }])
	})

	test('setRowHeight marks imported row height metadata as custom', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.rowDefs.set(2, { hidden: true, customHeight: false })

		const resized = applyOperation(wb, {
			op: 'setRowHeight',
			sheet: 'Sheet1',
			row: 2,
			height: 24,
		})
		expectOk(resized)
		expect(sheet.rowHeights.get(2)).toBe(24)
		expect(sheet.rowDefs.get(2)).toEqual({ hidden: true, customHeight: true })
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

	test('row and column layout operations reject invalid coordinates without metadata writes', () => {
		const cases: readonly Operation[] = [
			{ op: 'insertRows', sheet: 'Sheet1', at: -1, count: 1 },
			{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 0 },
			{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 0 },
			{ op: 'insertCols', sheet: 'Sheet1', at: -1, count: 1 },
			{ op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 0 },
			{ op: 'hideRows', sheet: 'Sheet1', at: -1, count: 1 },
			{ op: 'hideCols', sheet: 'Sheet1', at: -1, count: 1 },
			{ op: 'setRowHeight', sheet: 'Sheet1', row: -1, height: 12 },
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: -1 },
			{ op: 'setColWidth', sheet: 'Sheet1', col: -1, width: 12 },
			{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: -1 },
			{ op: 'groupRows', sheet: 'Sheet1', from: 0.5, to: 2 },
			{ op: 'groupCols', sheet: 'Sheet1', from: 0, to: 1.5 },
		]

		for (const op of cases) {
			const wb = setup()
			const result = applyOperation(wb, op)
			expectErr(result)
			expect(result.error.code, op.op).toBe('VALIDATION_ERROR')
			const sheet = wb.getSheet('Sheet1')
			expect(sheet?.rowDefs.size, op.op).toBe(0)
			expect(sheet?.colDefs, op.op).toEqual([])
			expect(sheet?.rowHeights.size, op.op).toBe(0)
			expect(sheet?.colWidths.size, op.op).toBe(0)
		}
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

	test('addSheet rejects Excel-invalid names', () => {
		const wb = setup()
		const result = applyOperation(wb, { op: 'addSheet', name: 'Bad/Name' })
		expectErr(result)
		expect(result.error.code).toBe('VALIDATION_ERROR')
		expect(result.error.message).toContain('invalid characters')
	})

	test('sheet layout operations reject invalid positions and panes before mutation', () => {
		const cases: readonly Operation[] = [
			{ op: 'addSheet', name: 'Sheet2', position: -1 },
			{ op: 'copySheet', sheet: 'Sheet1', newName: 'Copy', position: -1 },
			{ op: 'moveSheet', sheet: 'Sheet1', position: -1 },
			{ op: 'setTabColor', sheet: 'Sheet1', color: 'not-a-color' },
			{ op: 'setTabColor', sheet: 'Sheet1', color: '#FF0000' },
			{ op: 'freezePane', sheet: 'Sheet1', row: -1, col: 0 },
			{ op: 'freezePane', sheet: 'Sheet1', row: 0.5, col: 0 },
			{ op: 'freezePane', sheet: 'Sheet1', row: 0, col: -1 },
		]

		for (const op of cases) {
			const wb = setup()
			const beforeSheets = wb.sheets.map((sheet) => sheet.name)
			const result = applyOperation(wb, op)
			expectErr(result)
			expect(result.error.code, op.op).toBe('VALIDATION_ERROR')
			expect(
				wb.sheets.map((sheet) => sheet.name),
				op.op,
			).toEqual(beforeSheets)
			const sheet = wb.getSheet('Sheet1')
			expect(sheet?.frozenRows, op.op).toBe(0)
			expect(sheet?.frozenCols, op.op).toBe(0)
			expect(sheet?.tabColor, op.op).toBeNull()
		}
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
			sheetName: 'sheet2',
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

	test('structural row and column edits on a cloned workbook do not mutate source metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.merges.push({ start: { row: 1, col: 2 }, end: { row: 1, col: 3 } })
		sheet.dataValidations.push({ sqref: 'F2:F2', type: 'list', formula1: '"yes,no"' })

		const clone = wb.clone()
		expectOk(applyOperation(clone, { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }))
		expectOk(applyOperation(clone, { op: 'insertCols', sheet: 'Sheet1', at: 0, count: 1 }))

		expect(wb.getSheet('Sheet1')?.merges).toEqual([
			{ start: { row: 1, col: 2 }, end: { row: 1, col: 3 } },
		])
		expect(wb.getSheet('Sheet1')?.dataValidations).toEqual([
			{ sqref: 'F2:F2', type: 'list', formula1: '"yes,no"' },
		])
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

	test('deleteRows tombstones fully removed x14 metadata without rewriting formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push({
			index: 0,
			sqref: 'A2',
			type: 'list',
			formula1: 'A3',
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B2',
			formulas: ['A3>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'A3' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'B3' }] },
		})

		expectOk(applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.x14DataValidations[0]).toMatchObject({
			sqref: '',
			deleted: true,
			formula1: 'A3',
		})
		expect(s.x14ConditionalFormats[0]).toMatchObject({
			sqref: '',
			deleted: true,
			formulas: ['A3>0'],
		})
		expect(s.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('A3')
		expect(s.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('B3')
	})

	test('deleteRows rewrites deleted formula refs to #REF! across formula surfaces', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		const summary = wb.addSheet('Summary')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(1, 0, cell(numberValue(2)))
		s.cells.set(2, 0, cell(numberValue(3)))
		s.cells.set(0, 3, cell(EMPTY, 'A2'))
		s.cells.set(0, 4, cell(EMPTY, 'SUM(A2)'))
		s.cells.set(0, 5, cell(EMPTY, 'SUM(A1:A3)'))
		summary.cells.set(0, 0, cell(EMPTY, 'sheet1!A2'))
		wb.definedNames.set('DeletedInput', 'Sheet1!A2')
		s.dataValidations.push({ sqref: 'G1', type: 'list', formula1: 'A2' })
		s.conditionalFormats.push({
			sqref: 'H1',
			rules: [
				{
					type: 'expression',
					formulas: ['A2>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'A2' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
				},
			],
		})
		s.x14DataValidations.push({ index: 0, sqref: 'I1', type: 'list', formula1: 'A2' })
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'J1',
			formulas: ['A2>0'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'A2' }],
				colors: [{ rgb: 'FF63BE7B' }],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'A2' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'A2' }] },
		})
		s.tables.push({
			id: createTableId(),
			name: 'FormulaTable',
			sheetId: s.id,
			ref: { start: { row: 4, col: 0 }, end: { row: 5, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Calc', formula: 'A2', totalsRowFormula: 'SUM(A1:A3)' }],
			hasHeaders: true,
			hasTotals: false,
		})

		expectOk(applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.cells.get(0, 3)?.formula).toBe('#REF!')
		expect(s.cells.get(0, 4)?.formula).toBe('SUM(#REF!)')
		expect(s.cells.get(0, 5)?.formula).toBe('SUM(A1:A2)')
		expect(summary.cells.get(0, 0)?.formula).toBe('#REF!')
		expect(wb.definedNames.get('DeletedInput')).toBe('#REF!')
		expect(s.dataValidations[0]?.formula1).toBe('#REF!')
		expect(s.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['#REF!>0'])
		expect(s.conditionalFormats[0]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe('#REF!')
		expect(s.x14DataValidations[0]?.formula1).toBe('#REF!')
		expect(s.x14ConditionalFormats[0]?.formulas).toEqual(['#REF!>0'])
		expect(s.x14ConditionalFormats[0]?.colorScale?.cfvo[0]?.value).toBe('#REF!')
		expect(s.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('#REF!')
		expect(s.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('#REF!')
		expect(s.tables[0]?.columns[1]?.formula).toBe('#REF!')
		expect(s.tables[0]?.columns[1]?.totalsRowFormula).toBe('SUM(A1:A2)')
	})

	test('insertRows rewrites cross-sheet worksheet metadata formulas', () => {
		const wb = createWorkbook()
		const input = wb.addSheet('Input')
		const summary = wb.addSheet('Summary')
		input.cells.set(1, 0, cell(numberValue(1)))
		input.cells.set(1, 1, cell(numberValue(2)))
		summary.dataValidations.push({
			sqref: 'A1',
			type: 'list',
			formula1: 'Input!A2',
			formula2: 'A2',
		})
		summary.conditionalFormats.push({
			sqref: 'A2',
			rules: [
				{
					type: 'expression',
					formulas: ['Input!A2>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'Input!B2' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: { cfvo: [{ type: 'formula', value: 'Input!A2' }] },
					iconSet: { cfvo: [{ type: 'formula', value: 'Input!B2' }] },
				},
			],
		})
		summary.x14DataValidations.push({
			index: 0,
			sqref: 'B1',
			type: 'list',
			formula1: 'Input!A2',
			formula2: 'A2',
		})
		summary.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B2',
			formulas: ['Input!A2>0'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'Input!B2' }],
				colors: [{ rgb: 'FF63BE7B' }],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'Input!A2' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'Input!B2' }] },
		})
		summary.tables.push({
			id: createTableId(),
			name: 'SummaryTable',
			sheetId: summary.id,
			ref: { start: { row: 4, col: 0 }, end: { row: 5, col: 1 } },
			columns: [
				{ name: 'Name' },
				{ name: 'Calc', formula: 'Input!A2', totalsRowFormula: 'SUM(Input!A2:A3)' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		expectOk(applyOperation(wb, { op: 'insertRows', sheet: 'Input', at: 0, count: 1 }))

		expect(summary.dataValidations[0]?.formula1).toBe('Input!A3')
		expect(summary.dataValidations[0]?.formula2).toBe('A2')
		expect(summary.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['Input!A3>0'])
		expect(summary.conditionalFormats[0]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe('Input!B3')
		expect(summary.conditionalFormats[0]?.rules[0]?.dataBar?.cfvo[0]?.value).toBe('Input!A3')
		expect(summary.conditionalFormats[0]?.rules[0]?.iconSet?.cfvo[0]?.value).toBe('Input!B3')
		expect(summary.x14DataValidations[0]?.formula1).toBe('Input!A3')
		expect(summary.x14DataValidations[0]?.formula2).toBe('A2')
		expect(summary.x14ConditionalFormats[0]?.formulas).toEqual(['Input!A3>0'])
		expect(summary.x14ConditionalFormats[0]?.colorScale?.cfvo[0]?.value).toBe('Input!B3')
		expect(summary.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('Input!A3')
		expect(summary.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('Input!B3')
		expect(summary.tables[0]?.columns[1]?.formula).toBe('Input!A3')
		expect(summary.tables[0]?.columns[1]?.totalsRowFormula).toBe('SUM(Input!A3:A4)')
	})

	test('deleteRows rejects partial table header row deletion', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'HeaderedTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: { ref: 'A1:B4', columns: [] },
		})

		const result = applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 0, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('header row')
		expect(s.tables[0]?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 3, col: 1 },
		})
	})

	test('deleteRows rejects partial table totals row deletion', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'TotalsTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 4, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value', totalsRowFormula: 'SUM([Value])' }],
			hasHeaders: true,
			hasTotals: true,
			autoFilter: { ref: 'A1:B4', columns: [] },
		})

		const result = applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 4, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('totals row')
		expect(s.tables[0]?.hasTotals).toBe(true)
		expect(s.tables[0]?.ref.end.row).toBe(4)
	})

	test('deleteRows removes unreferenced table metadata when the full table row span is deleted', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'DeletedRowsTable',
			sheetId: s.id,
			ref: { start: { row: 1, col: 0 }, end: { row: 3, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: { ref: 'A2:B4', columns: [] },
		})

		expectOk(applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 3 }))

		expect(s.tables).toHaveLength(0)
	})

	test('row and column shifts reject ambiguous overlapping table ownership', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(1, 0, { value: stringValue('kept'), formula: null, styleId: sid })
		s.tables.push(
			{
				id: createTableId(),
				name: 'Sales',
				sheetId: s.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
				columns: [{ name: 'Region' }, { name: 'Amount' }],
				hasHeaders: true,
				hasTotals: false,
			},
			{
				id: createTableId(),
				name: 'Forecast',
				sheetId: s.id,
				ref: { start: { row: 1, col: 1 }, end: { row: 3, col: 2 } },
				columns: [{ name: 'Scenario' }, { name: 'Value' }],
				hasHeaders: true,
				hasTotals: false,
			},
		)

		const result = applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('would overlap table Forecast')
		expect(result.error.details).toMatchObject({
			kind: 'overlapping-table-ranges',
			left: { tableName: 'Sales', ref: 'A1:B4' },
			right: { tableName: 'Forecast', ref: 'B3:C5' },
		})
		expect(s.cells.get(1, 0)?.value).toEqual(stringValue('kept'))
		expect(s.tables[0]?.ref.end.row).toBe(2)
		expect(s.tables[1]?.ref.start.row).toBe(1)

		const colWb = createWorkbook()
		const colSheet = colWb.addSheet('Sheet1')
		colSheet.cells.set(0, 1, { value: stringValue('kept'), formula: null, styleId: sid })
		colSheet.tables.push(
			{
				id: createTableId(),
				name: 'Actuals',
				sheetId: colSheet.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
				columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Owner' }],
				hasHeaders: true,
				hasTotals: false,
			},
			{
				id: createTableId(),
				name: 'Targets',
				sheetId: colSheet.id,
				ref: { start: { row: 1, col: 1 }, end: { row: 2, col: 3 } },
				columns: [{ name: 'Scenario' }, { name: 'Value' }, { name: 'Owner' }],
				hasHeaders: true,
				hasTotals: false,
			},
		)

		const colResult = applyOperation(colWb, {
			op: 'insertCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
		})

		expectErr(colResult)
		expect(colResult.error.message).toContain('would overlap table Targets')
		expect(colResult.error.details).toMatchObject({
			kind: 'overlapping-table-ranges',
			left: { tableName: 'Actuals', ref: 'A1:D2' },
			right: { tableName: 'Targets', ref: 'C2:E3' },
		})
		expect(colSheet.cells.get(0, 1)?.value).toEqual(stringValue('kept'))
		expect(colSheet.tables[0]?.ref.end.col).toBe(2)
		expect(colSheet.tables[1]?.ref.start.col).toBe(1)
	})

	test('deleteCols prunes deleted sortCondition refs across sheet, autofilter, and table sort states', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.autoFilter = {
			ref: 'A1:D4',
			columns: [],
			sortState: {
				ref: 'A1:D4',
				conditions: [{ ref: 'B1:B4', descending: true }, { ref: 'D1:D4' }],
			},
		}
		s.sortState = {
			ref: 'A1:D4',
			conditions: [{ ref: 'B1:B4', descending: true }, { ref: 'D1:D4' }],
		}
		s.tables.push({
			id: createTableId(),
			name: 'SortedTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } },
			columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }, { name: 'D' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:D4',
				columns: [],
				sortState: {
					ref: 'A1:D4',
					conditions: [{ ref: 'B1:B4', descending: true }, { ref: 'D1:D4' }],
				},
			},
			sortState: {
				ref: 'A1:D4',
				conditions: [{ ref: 'B1:B4', descending: true }, { ref: 'D1:D4' }],
			},
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.autoFilter?.sortState).toEqual({
			ref: 'A1:C4',
			conditions: [{ ref: 'C1:C4' }],
		})
		expect(s.sortState).toEqual({
			ref: 'A1:C4',
			conditions: [{ ref: 'C1:C4' }],
		})
		expect(s.tables[0]?.autoFilter?.sortState).toEqual({
			ref: 'A1:C4',
			conditions: [{ ref: 'C1:C4' }],
		})
		expect(s.tables[0]?.sortState).toEqual({
			ref: 'A1:C4',
			conditions: [{ ref: 'C1:C4' }],
		})
	})

	test('deleteCols removes empty sort states when every sortCondition ref is deleted', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.autoFilter = {
			ref: 'A1:C4',
			columns: [],
			sortState: { ref: 'A1:C4', conditions: [{ ref: 'B1:B4' }] },
		}
		s.preservedAutoFilterSortStateAttributes = { ref: 'A1:C4' }
		s.sortState = { ref: 'A1:C4', conditions: [{ ref: 'B1:B4' }] }
		s.preservedSortStateAttributes = { ref: 'A1:C4' }
		s.tables.push({
			id: createTableId(),
			name: 'EmptySortedTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:C4',
				columns: [],
				sortState: { ref: 'A1:C4', conditions: [{ ref: 'B1:B4' }] },
			},
			sortState: { ref: 'A1:C4', conditions: [{ ref: 'B1:B4' }] },
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.autoFilter).toEqual({ ref: 'A1:B4', columns: [] })
		expect(s.preservedAutoFilterSortStateAttributes).toBeNull()
		expect(s.sortState).toBeNull()
		expect(s.preservedSortStateAttributes).toBeNull()
		expect(s.tables[0]?.autoFilter).toEqual({ ref: 'A1:B4', columns: [] })
		expect(s.tables[0]?.sortState).toBeUndefined()
	})

	test('deleteCols remaps autoFilter filterColumn ids and removes criteria for deleted columns', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		const autoFilter = {
			ref: 'A1:D4',
			columns: [
				{ colId: 1, kind: 'filters' as const, values: ['West'] },
				{ colId: 3, kind: 'filters' as const, values: ['Open'] },
			],
		}
		s.autoFilter = autoFilter
		s.advancedFilters.push({
			viewName: 'SavedView',
			ref: 'A1:D4',
			autoFilter,
			filterColumnCount: 2,
			sortConditionCount: 0,
		})
		s.tables.push({
			id: createTableId(),
			name: 'FilteredTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } },
			columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }, { name: 'D' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter,
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.autoFilter).toEqual({
			ref: 'A1:C4',
			columns: [{ colId: 2, kind: 'filters', values: ['Open'] }],
		})
		expect(s.advancedFilters[0]).toMatchObject({
			ref: 'A1:C4',
			filterColumnCount: 1,
			sortConditionCount: 0,
			autoFilter: {
				ref: 'A1:C4',
				columns: [{ colId: 2, kind: 'filters', values: ['Open'] }],
			},
		})
		expect(s.tables[0]?.autoFilter).toEqual({
			ref: 'A1:C4',
			columns: [{ colId: 2, kind: 'filters', values: ['Open'] }],
		})
	})

	test('insertCols remaps filterColumn ids relative to expanded autoFilter ranges', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		const autoFilter = {
			ref: 'A1:C4',
			columns: [{ colId: 1, kind: 'filters' as const, values: ['West'] }],
		}
		s.autoFilter = autoFilter
		s.advancedFilters.push({
			viewName: 'SavedView',
			ref: 'A1:C4',
			autoFilter,
			filterColumnCount: 1,
			sortConditionCount: 0,
		})
		s.tables.push({
			id: createTableId(),
			name: 'InsertedFilterTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter,
		})

		expectOk(applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.autoFilter).toEqual({
			ref: 'A1:D4',
			columns: [{ colId: 2, kind: 'filters', values: ['West'] }],
		})
		expect(s.advancedFilters[0]?.autoFilter).toEqual({
			ref: 'A1:D4',
			columns: [{ colId: 2, kind: 'filters', values: ['West'] }],
		})
		expect(s.tables[0]?.autoFilter).toEqual({
			ref: 'A1:D4',
			columns: [{ colId: 2, kind: 'filters', values: ['West'] }],
		})
	})

	test('deleteCols removes deleted tableColumn metadata and keeps surviving columns aligned with table ref', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'ColumnDeleteTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } },
			columns: [
				{ id: 10, name: 'Region' },
				{ id: 11, name: 'Rep', totalsRowLabel: 'Total' },
				{ id: 12, name: 'Amount', formula: 'SUM(A2:A4)' },
				{ id: 13, name: 'Status' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:D4',
				columns: [
					{ colId: 1, kind: 'filters', values: ['Ada'] },
					{ colId: 3, kind: 'filters', values: ['Open'] },
				],
			},
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.tables[0]?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 3, col: 2 },
		})
		expect(s.tables[0]?.columns).toEqual([
			{ id: 10, name: 'Region' },
			{ id: 12, name: 'Amount', formula: 'SUM(A2:A4)' },
			{ id: 13, name: 'Status' },
		])
		expect(s.tables[0]?.autoFilter).toEqual({
			ref: 'A1:C4',
			columns: [{ colId: 2, kind: 'filters', values: ['Open'] }],
		})
	})

	test('deleteCols rejects table field deletion when cell formulas still reference the field', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		s.cells.set(0, 4, cell(EMPTY, 'COUNTA(Sales[Rep])'))

		const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(s.tables[0]?.columns.map((column) => column.name)).toEqual(['Region', 'Rep', 'Amount'])
	})

	test('deleteCols rejects table field deletion when defined names still reference the field', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		wb.definedNames.set('SalesReps', 'COUNTA(Sales[Rep])')

		const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(wb.definedNames.get('SalesReps')).toBe('COUNTA(Sales[Rep])')
	})

	test('deleteCols rejects table field deletion when worksheet metadata still references the field', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		s.dataValidations.push({ sqref: 'E1:E4', type: 'list', formula1: 'Sales[Rep]' })

		const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(s.dataValidations[0]?.formula1).toBe('Sales[Rep]')
	})

	test('deleteCols rejects table field deletion across worksheet metadata formula surfaces', () => {
		for (const scenario of TABLE_FIELD_METADATA_BLOCKERS) {
			const { wb, sheet } = setupSalesRepTable()
			scenario.add(sheet)

			const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

			expectErr(result)
			const diagnostic = `${scenario.label}: ${result.error.message}`
			expect(diagnostic).toContain(`${scenario.sourceRef} ${scenario.sourceKind}`)
			expect(diagnostic).toContain('Sales[Rep]')
			expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
			expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
				'Region',
				'Rep',
				'Amount',
			])
		}
	})

	test('deleteCols rejects table field deletion across local worksheet metadata formula surfaces', () => {
		for (const scenario of LOCAL_TABLE_FIELD_METADATA_BLOCKERS) {
			const { wb, sheet } = setupSalesRepTable()
			scenario.add(sheet)

			const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

			expectErr(result)
			const diagnostic = `${scenario.label}: ${result.error.message}`
			expect(diagnostic).toContain(`${scenario.sourceRef} ${scenario.sourceKind}`)
			expect(diagnostic).toContain('Sales[Rep]')
			expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
			expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
				'Region',
				'Rep',
				'Amount',
			])
		}
	})

	test('deleteCols rejects table field deletion when table-scoped worksheet metadata uses local structured refs', () => {
		const { wb, sheet } = setupSalesRepTable()
		sheet.dataValidations.push({
			sqref: 'A2:C4',
			type: 'list',
			formula1: 'COUNTA([@Rep])',
		})

		const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(result.error.message).toContain('Sheet1!A2:C4 data validation formula1')
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Region',
			'Rep',
			'Amount',
		])
		expect(sheet.dataValidations[0]?.formula1).toBe('COUNTA([@Rep])')
	})

	test('deleteCols rejects table field deletion when surviving calculated columns use local structured refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount', formula: 'COUNTA([Rep])' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(s.tables[0]?.columns[2]?.formula).toBe('COUNTA([Rep])')
	})

	test('deleteCols removes unreferenced table metadata when the full table range is deleted', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'SingleColumn',
			sheetId: s.id,
			ref: { start: { row: 0, col: 1 }, end: { row: 3, col: 1 } },
			columns: [{ id: 1, name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'B1:B4',
				columns: [{ colId: 0, kind: 'filters', values: ['10'] }],
			},
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.tables).toHaveLength(0)
	})

	test('deleteCols removes queryTable sidecar metadata when the full table range is deleted', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 0, count: 3 }))

		expect(sheet.tables).toHaveLength(0)
		expect(wb.connectionParts).toEqual([])
	})

	test('deleteRows removes queryTable sidecar metadata when the full table range is deleted', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()

		expectOk(applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 0, count: 4 }))

		expect(sheet.tables).toHaveLength(0)
		expect(wb.connectionParts).toEqual([])
	})

	test('deleteCols rejects partial queryTable-backed table field deletion', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()

		const result = applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('queryTable-backed table "Sales"')
		expect(result.error.message).toContain('field bindings ambiguous')
		expect(result.error.details).toMatchObject({
			kind: 'query-table-column-structural-edit',
			tableName: 'Sales',
			currentRef: 'A1:C4',
			shiftedRef: 'A1:B4',
			queryTablePartPath: 'xl/queryTables/queryTable1.xml',
		})
		expect(sheet.tables[0]?.columns.map((column) => column.queryTableFieldId)).toEqual([1, 2, 3])
		expect(wb.connectionParts).toHaveLength(1)
	})

	test('insertCols adds generated tableColumn metadata for inserted columns inside a table', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'ColumnInsertTable',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Column2' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:C4',
				columns: [{ colId: 1, kind: 'filters', values: ['West'] }],
			},
		})

		expectOk(applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.tables[0]?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 3, col: 3 },
		})
		expect(s.tables[0]?.columns).toEqual([
			{ id: 1, name: 'Region' },
			{ id: 4, name: 'Column2_2' },
			{ id: 2, name: 'Column2' },
			{ id: 3, name: 'Amount' },
		])
		expect(s.tables[0]?.autoFilter).toEqual({
			ref: 'A1:D4',
			columns: [{ colId: 2, kind: 'filters', values: ['West'] }],
		})
	})

	test('insertCols rejects generated columns inside queryTable-backed tables', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()

		const result = applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 })

		expectErr(result)
		expect(result.error.message).toContain('queryTable-backed table "Sales"')
		expect(result.error.message).toContain('field bindings ambiguous')
		expect(result.error.details).toMatchObject({
			kind: 'query-table-column-structural-edit',
			tableName: 'Sales',
			currentRef: 'A1:C4',
			shiftedRef: 'A1:D4',
			queryTablePartPath: 'xl/queryTables/queryTable1.xml',
		})
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Region',
			'Rep',
			'Amount',
		])
	})

	test('insertCols before queryTable-backed tables shifts ownership without remapping fields', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()

		expectOk(applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 0, count: 1 }))

		expect(sheet.tables[0]?.ref).toEqual({
			start: { row: 0, col: 1 },
			end: { row: 3, col: 3 },
		})
		expect(sheet.tables[0]?.columns.map((column) => column.queryTableFieldId)).toEqual([1, 2, 3])
		expect(wb.connectionParts).toHaveLength(1)
	})

	test('setTableColumn rejects renaming queryTable-backed columns', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()
		sheet.cells.set(0, 1, { value: stringValue('Rep'), formula: null, styleId: sid })

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Rep',
			newName: 'Representative',
		})

		expectErr(result)
		expect(result.error.message).toContain('Cannot rename queryTable-backed column "Rep"')
		expect(result.error.details).toMatchObject({
			kind: 'query-table-column-rename',
			tableName: 'Sales',
			columnName: 'Rep',
			queryTableFieldId: 2,
			queryTablePartPath: 'xl/queryTables/queryTable1.xml',
		})
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Region',
			'Rep',
			'Amount',
		])
		expect(sheet.tables[0]?.columns.map((column) => column.queryTableFieldId)).toEqual([1, 2, 3])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('Rep'))
		expect(wb.connectionParts).toHaveLength(1)
	})

	test('setTableColumn keeps formula-only edits available on queryTable-backed columns', () => {
		const { wb, sheet } = setupQueryTableBackedSalesTable()

		expectOk(
			applyOperation(wb, {
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Amount',
				formula: '=[@Region]&[@Rep]',
			}),
		)

		expect(sheet.tables[0]?.columns[2]).toMatchObject({
			name: 'Amount',
			queryTableFieldId: 3,
			formula: '[@Region]&[@Rep]',
		})
		expect(wb.connectionParts).toHaveLength(1)
	})

	test('insertRows rewrites formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(1, 0, cell(numberValue(2)))
		s.cells.set(2, 0, cell(EMPTY, 'SUM(A1:A2)'))

		const result = applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 })
		expectOk(result)

		const formulaCell = s.cells.get(3, 0)
		expect(formulaCell?.formula).toBe('SUM(A1:A3)')
		expect(result.value.affectedCells).toEqual(['A4'])
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
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

		const result = applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 0, count: 1 })
		expectOk(result)

		expect(s2.cells.get(0, 0)?.formula).toBe('SUM(Sheet1!A2#)')
		expect(result.value.affectedCells).toEqual(['Sheet2!A1'])
		expect(result.value.sheetsModified).toEqual(['Sheet1', 'Sheet2'])
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
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
		s.comments.set('A2', {
			text: 'note',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 1,
				column: 0,
				anchor: [0, 15, 1, 2, 2, 15, 4, 16],
				visible: true,
			},
		})
		s.threadedComments.push({
			ref: 'B2',
			text: 'thread',
			id: 'tc1',
			partPath: 'xl/threadedComments/threadedComment1.xml',
			personId: '0',
			author: 'Ada',
		})
		s.hyperlinks.set('B2', { target: 'https://example.com', location: 'sheet1!A2' })
		s.dataValidations.push({ sqref: 'A2:B2', type: 'list', formula1: 'A2' })
		s.conditionalFormats.push({
			sqref: 'A2',
			rules: [
				{
					type: 'expression',
					formulas: ['A2>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'B2' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: { cfvo: [{ type: 'formula', value: 'C2' }] },
					iconSet: { cfvo: [{ type: 'formula', value: 'D2' }] },
				},
			],
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C2',
			formulas: ['A2>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'B2' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'A2' }] },
		})
		s.x14DataValidations.push({
			index: 0,
			sqref: 'D2',
			type: 'list',
			allowBlank: true,
			formula1: 'A2',
		})
		s.ignoredErrors.push({ sqref: 'A2', formula: true })
		s.rowHeights.set(1, 24)

		applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 2 })

		expect(s.comments.get('A4')).toEqual({
			text: 'note',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 3,
				column: 0,
				anchor: [0, 15, 3, 2, 2, 15, 6, 16],
				visible: true,
			},
		})
		expect(s.threadedComments).toEqual([
			{
				ref: 'B4',
				text: 'thread',
				id: 'tc1',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				personId: '0',
				author: 'Ada',
			},
		])
		expect(s.hyperlinks.get('B4')).toEqual({
			target: 'https://example.com',
			location: 'sheet1!A4',
		})
		expect(s.dataValidations[0]?.sqref).toBe('A4:B4')
		expect(s.dataValidations[0]?.formula1).toBe('A4')
		expect(s.conditionalFormats[0]?.sqref).toBe('A4')
		expect(s.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('A4>0')
		expect(s.conditionalFormats[0]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe('B4')
		expect(s.conditionalFormats[0]?.rules[0]?.dataBar?.cfvo[0]?.value).toBe('C4')
		expect(s.conditionalFormats[0]?.rules[0]?.iconSet?.cfvo[0]?.value).toBe('D4')
		expect(s.x14ConditionalFormats[0]?.sqref).toBe('C4')
		expect(s.x14ConditionalFormats[0]?.formulas[0]).toBe('A4>0')
		expect(s.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('B4')
		expect(s.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('A4')
		expect(s.x14DataValidations[0]).toMatchObject({
			sqref: 'D4',
			type: 'list',
			allowBlank: true,
			formula1: 'A4',
		})
		expect(s.ignoredErrors[0]?.sqref).toBe('A4')
		expect(s.rowHeights.get(3)).toBe(24)
	})

	test('insertRows shifts hyperlink locations with escaped sheet names', () => {
		const wb = createWorkbook()
		const s = wb.addSheet("Bob's Budget")
		s.hyperlinks.set('A1', { location: "'Bob''s Budget'!A2", display: 'jump' })

		const result = applyOperation(wb, {
			op: 'insertRows',
			sheet: "Bob's Budget",
			at: 1,
			count: 1,
		})
		expectOk(result)

		expect(s.hyperlinks.get('A1')?.location).toBe("'Bob''s Budget'!A3")
	})

	test('row and column shifts keep comment refs and VML metadata coherent', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A2', {
			text: 'note',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 1,
				column: 0,
				anchor: [0, 15, 1, 2, 2, 15, 4, 2],
			},
		})
		s.threadedComments.push({
			ref: 'B2',
			text: 'thread',
			id: 'tc1',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		expectOk(applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 2 }))
		expectOk(applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 0, count: 1 }))

		expect(s.comments.get('B4')).toEqual({
			text: 'note',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 3,
				column: 1,
				anchor: [1, 15, 3, 2, 3, 15, 6, 2],
			},
		})
		expect(s.threadedComments).toEqual([
			{
				ref: 'C4',
				text: 'thread',
				id: 'tc1',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		])
	})

	test('row and column deletes remove comments whose refs are deleted', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A2', { text: 'delete me' })
		s.comments.set('C2', { text: 'keep me' })
		s.threadedComments.push(
			{ ref: 'B2', text: 'delete thread', id: 'tc1' },
			{ ref: 'D2', text: 'keep thread', id: 'tc2' },
		)

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 0, count: 2 }))

		expect([...s.comments.entries()]).toEqual([['A2', { text: 'keep me' }]])
		expect(s.threadedComments).toEqual([{ ref: 'B2', text: 'keep thread', id: 'tc2' }])
	})

	test('deleteRows removes deleted comments and retargets surviving VML metadata', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('B2', {
			text: 'delete me',
			legacyDrawing: {
				shapeId: '_x0000_s1027',
				row: 1,
				column: 1,
				anchor: [1, 15, 1, 2, 3, 15, 4, 16],
			},
		})
		s.comments.set('C4', {
			text: 'keep me',
			legacyDrawing: {
				shapeId: '_x0000_s1028',
				row: 3,
				column: 2,
				anchor: [2, 15, 3, 2, 4, 15, 6, 16],
			},
		})
		s.threadedComments.push(
			{ ref: 'D2', text: 'delete thread', id: 'tc-delete' },
			{ ref: 'E4', text: 'keep thread', id: 'tc-keep' },
		)

		expectOk(applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 2 }))

		expect([...s.comments.entries()]).toEqual([
			[
				'C2',
				{
					text: 'keep me',
					legacyDrawing: {
						shapeId: '_x0000_s1028',
						row: 1,
						column: 2,
						anchor: [2, 15, 1, 2, 4, 15, 4, 16],
					},
				},
			],
		])
		expect(s.threadedComments).toEqual([{ ref: 'E2', text: 'keep thread', id: 'tc-keep' }])
	})

	test('deleteCols removes deleted comments and retargets surviving VML metadata', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('B2', {
			text: 'delete me',
			legacyDrawing: {
				shapeId: '_x0000_s1029',
				row: 1,
				column: 1,
				anchor: [1, 15, 1, 2, 3, 15, 4, 16],
			},
		})
		s.comments.set('E3', {
			text: 'keep me',
			legacyDrawing: {
				shapeId: '_x0000_s1030',
				row: 2,
				column: 4,
				anchor: [4, 15, 2, 2, 6, 15, 5, 16],
			},
		})
		s.threadedComments.push(
			{ ref: 'C4', text: 'delete thread', id: 'tc-delete' },
			{ ref: 'F4', text: 'keep thread', id: 'tc-keep' },
		)

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 2 }))

		expect([...s.comments.entries()]).toEqual([
			[
				'C3',
				{
					text: 'keep me',
					legacyDrawing: {
						shapeId: '_x0000_s1030',
						row: 2,
						column: 2,
						anchor: [2, 15, 2, 2, 4, 15, 5, 16],
					},
				},
			],
		])
		expect(s.threadedComments).toEqual([{ ref: 'D4', text: 'keep thread', id: 'tc-keep' }])
	})

	test('insertCols shifts tables, filters, comments, and hyperlinks', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		s.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		s.comments.set('A1', {
			text: 'header',
			legacyDrawing: {
				shapeId: '_x0000_s1026',
				row: 0,
				column: 0,
				anchor: [0, 15, 0, 2, 2, 15, 3, 16],
			},
		})
		s.threadedComments.push({
			ref: 'A2',
			text: 'threaded header',
			id: 'tc2',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})
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

		expect(s.comments.get('B1')).toEqual({
			text: 'header',
			legacyDrawing: {
				shapeId: '_x0000_s1026',
				row: 0,
				column: 1,
				anchor: [1, 15, 0, 2, 3, 15, 3, 16],
			},
		})
		expect(s.threadedComments[0]).toMatchObject({
			ref: 'B2',
			text: 'threaded header',
			id: 'tc2',
		})
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

	test('deleteCols rewrites deleted formula refs to #REF! across formula surfaces', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))
		s.cells.set(0, 1, cell(numberValue(2)))
		s.cells.set(0, 2, cell(numberValue(3)))
		s.cells.set(1, 0, cell(EMPTY, 'B1'))
		s.cells.set(2, 0, cell(EMPTY, 'SUM(A1:C1)'))
		wb.definedNames.set('DeletedColumnInput', 'Sheet1!B1')
		s.dataValidations.push({ sqref: 'A4', type: 'list', formula1: 'B1' })
		s.conditionalFormats.push({
			sqref: 'A5',
			rules: [{ type: 'expression', formulas: ['B1>0'] }],
		})
		s.x14DataValidations.push({ index: 0, sqref: 'A6', type: 'list', formula1: 'B1' })
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A7',
			formulas: ['B1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'B1' }] },
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }))

		expect(s.cells.get(1, 0)?.formula).toBe('#REF!')
		expect(s.cells.get(2, 0)?.formula).toBe('SUM(A1:B1)')
		expect(wb.definedNames.get('DeletedColumnInput')).toBe('#REF!')
		expect(s.dataValidations[0]?.formula1).toBe('#REF!')
		expect(s.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['#REF!>0'])
		expect(s.x14DataValidations[0]?.formula1).toBe('#REF!')
		expect(s.x14ConditionalFormats[0]?.formulas).toEqual(['#REF!>0'])
		expect(s.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('#REF!')
	})

	test('deleteCols rewrites cross-sheet worksheet metadata formulas to #REF!', () => {
		const wb = createWorkbook()
		const input = wb.addSheet('Input')
		const summary = wb.addSheet('Summary')
		input.cells.set(0, 0, cell(numberValue(1)))
		input.cells.set(0, 1, cell(numberValue(2)))
		input.cells.set(0, 2, cell(numberValue(3)))
		summary.dataValidations.push({ sqref: 'A1', type: 'list', formula1: 'Input!B1' })
		summary.conditionalFormats.push({
			sqref: 'A2',
			rules: [
				{
					type: 'expression',
					formulas: ['Input!B1>0'],
					dataBar: { cfvo: [{ type: 'formula', value: 'Input!B1' }] },
				},
			],
		})
		summary.x14DataValidations.push({
			index: 0,
			sqref: 'B1',
			type: 'list',
			formula1: 'Input!B1',
		})
		summary.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B2',
			formulas: ['Input!B1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'Input!B1' }] },
		})
		summary.tables.push({
			id: createTableId(),
			name: 'SummaryTable',
			sheetId: summary.id,
			ref: { start: { row: 4, col: 0 }, end: { row: 5, col: 1 } },
			columns: [
				{ name: 'Name' },
				{ name: 'Calc', formula: 'Input!B1', totalsRowFormula: 'SUM(Input!A1:C1)' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		expectOk(applyOperation(wb, { op: 'deleteCols', sheet: 'Input', at: 1, count: 1 }))

		expect(summary.dataValidations[0]?.formula1).toBe('#REF!')
		expect(summary.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['#REF!>0'])
		expect(summary.conditionalFormats[0]?.rules[0]?.dataBar?.cfvo[0]?.value).toBe('#REF!')
		expect(summary.x14DataValidations[0]?.formula1).toBe('#REF!')
		expect(summary.x14ConditionalFormats[0]?.formulas).toEqual(['#REF!>0'])
		expect(summary.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('#REF!')
		expect(summary.tables[0]?.columns[1]?.formula).toBe('#REF!')
		expect(summary.tables[0]?.columns[1]?.totalsRowFormula).toBe('SUM(Input!A1:B1)')
	})

	test('row and column shifts reject legacy array formula impact', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		const formulaInfo = { kind: 'array' as const, ref: 'B2:C3' }
		s.cells.set(1, 1, { value: numberValue(1), formula: 'B5:C6', styleId: sid, formulaInfo })
		s.cells.set(1, 2, { value: numberValue(2), formula: null, styleId: sid, formulaInfo })
		s.cells.set(2, 1, { value: numberValue(3), formula: null, styleId: sid, formulaInfo })
		s.cells.set(2, 2, { value: numberValue(4), formula: null, styleId: sid, formulaInfo })

		expectErr(applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 }))
		expectErr(applyOperation(wb, { op: 'deleteRows', sheet: 'Sheet1', at: 0, count: 1 }))
		expectErr(applyOperation(wb, { op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 }))
		expectErr(applyOperation(wb, { op: 'deleteCols', sheet: 'Sheet1', at: 0, count: 1 }))

		expect(s.cells.get(1, 1)?.formulaInfo).toEqual(formulaInfo)
		expect(s.cells.get(2, 2)?.value).toEqual(numberValue(4))

		const other = wb.addSheet('Other')
		other.cells.set(9, 0, cell(numberValue(10)))
		expectErr(applyOperation(wb, { op: 'insertRows', sheet: 'Other', at: 20, count: 1 }))
		expect(other.cells.get(9, 0)?.value).toEqual(numberValue(10))
	})

	test('row and column shifts reject imported formula bindings before mutation', () => {
		const formulaBindings = [
			{
				formula: 'B1*2',
				formulaInfo: { kind: 'shared' as const, sharedIndex: '0', isMaster: true, masterRef: 'A1' },
			},
			{
				formula: 'SEQUENCE(2)',
				formulaInfo: {
					kind: 'spill' as const,
					anchorRef: 'Shared!A1',
					ref: 'A1:A2',
					isAnchor: true,
				},
			},
			{
				formula: 'SEQUENCE(2)',
				formulaInfo: { kind: 'dynamicArray' as const, metadataIndex: 1, collapsed: false },
			},
			{
				formula: 'SEQUENCE(2)',
				formulaInfo: {
					kind: 'blockedSpill' as const,
					anchorRef: 'Shared!A1',
					ref: 'A1:A2',
					blockingRefs: ['A2'],
				},
			},
			{
				formula: null,
				formulaInfo: { kind: 'dataTable' as const, ref: 'A1:A2', dtr: true, r1: 'B1' },
			},
		]
		const operations = [
			{ op: 'insertRows' as const, sheet: 'Other', at: 10, count: 1 },
			{ op: 'deleteRows' as const, sheet: 'Other', at: 10, count: 1 },
			{ op: 'insertCols' as const, sheet: 'Other', at: 10, count: 1 },
			{ op: 'deleteCols' as const, sheet: 'Other', at: 10, count: 1 },
		]

		for (const binding of formulaBindings) {
			for (const operation of operations) {
				const wb = createWorkbook()
				const source = wb.addSheet('Shared')
				const other = wb.addSheet('Other')
				source.cells.set(0, 0, {
					value: numberValue(20),
					formula: binding.formula,
					styleId: sid,
					formulaInfo: binding.formulaInfo,
				})
				other.cells.set(4, 4, cell(numberValue(5)))

				const result = applyOperation(wb, operation)

				expectErr(result)
				expect(result.error.message).toContain(
					`imported ${binding.formulaInfo.kind} formula metadata`,
				)
				expect(source.cells.get(0, 0)?.formulaInfo).toEqual(binding.formulaInfo)
				expect(other.cells.get(4, 4)?.value).toEqual(numberValue(5))
			}
		}
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

	test('copyRange materializes effective formulas from shared formula members', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, {
			value: numberValue(20),
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		s.cells.set(1, 0, {
			value: numberValue(40),
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})
		s.cells.set(0, 1, cell(numberValue(10)))
		s.cells.set(1, 1, cell(numberValue(20)))

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'A2',
			target: 'C2',
		})
		expectOk(result)

		expect(s.cells.get(1, 2)?.formula).toBe('D2*2')
		expect(s.cells.get(1, 2)?.formulaInfo).toBeUndefined()
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
		const summary = wb.addSheet('Summary')
		summary.cells.set(0, 0, cell(EMPTY, 'sheet1!A1+1'))
		wb.definedNames.set('Budget', 'sheet1!A1')
		const result = applyOperation(wb, {
			op: 'renameSheet',
			sheet: 'Sheet1',
			newName: 'Data',
		})
		expect(result.ok).toBe(true)
		expect(wb.getSheet('Data')).toBeDefined()
		expect(wb.getSheet('Sheet1')).toBeUndefined()
		expect(wb.definedNames.get('Budget')).toBe('Data!A1')
		expect(summary.cells.get(0, 0)?.formula).toBe('Data!A1+1')
	})

	test('renameSheet materializes imported formula bindings before sheet-reference rewrites', () => {
		const wb = createWorkbook()
		const renamed = wb.addSheet('Sheet1')
		const other = wb.addSheet('Other')
		renamed.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'Other!A1',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'rename-local',
				isMaster: true,
				masterRef: 'A1',
				ref: 'A1:A2',
			},
		})
		renamed.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'rename-local',
				isMaster: false,
				masterRef: 'A1',
			},
		})
		other.cells.set(0, 2, {
			value: numberValue(3),
			formula: 'Sheet1!A1*2',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'rename-remote',
				isMaster: true,
				masterRef: 'C1',
				ref: 'C1:C2',
			},
		})
		other.cells.set(1, 2, {
			value: numberValue(4),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'rename-remote',
				isMaster: false,
				masterRef: 'C1',
			},
		})

		const result = applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet1', newName: 'Data' })
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['Data!A1', 'Data!A2', 'Other!C1', 'Other!C2'])
		expect(result.value.sheetsModified).toEqual(['Data', 'Other'])
		expect(renamed.cells.get(0, 0)?.formula).toBe('Other!A1')
		expect(renamed.cells.get(1, 0)?.formula).toBe('Other!A2')
		expect(other.cells.get(0, 2)?.formula).toBe('Data!A1*2')
		expect(other.cells.get(1, 2)?.formula).toBe('Data!A2*2')
		expect(renamed.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(renamed.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(other.cells.get(0, 2)?.formulaInfo).toBeUndefined()
		expect(other.cells.get(1, 2)?.formulaInfo).toBeUndefined()
	})

	test('renameSheet retargets dynamic spill metadata without dropping the anchor formula', () => {
		const wb = createWorkbook()
		const renamed = wb.addSheet('Sheet1')
		const other = wb.addSheet('Other')
		renamed.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(2)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		renamed.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
				isAnchor: false,
			},
		})
		other.cells.set(0, 0, {
			value: numberValue(2),
			formula: 'Sheet1!A1*2',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 2, collapsed: false },
		})
		other.cells.set(1, 0, {
			value: numberValue(4),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Other!A1',
				ref: 'Other!A1:A2',
				isAnchor: false,
			},
		})

		const result = applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet1', newName: 'Data' })
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['Data!A2', 'Other!A1', 'Other!A2'])
		expect(renamed.cells.get(0, 0)?.formula).toBe('SEQUENCE(2)')
		expect(renamed.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(renamed.cells.get(1, 0)?.formula).toBeNull()
		expect(renamed.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Data!A1',
			ref: 'Data!A1:A2',
			isAnchor: false,
		})
		expect(other.cells.get(0, 0)?.formula).toBe('Data!A1*2')
		expect(other.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 2,
			collapsed: false,
		})
		expect(other.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Other!A1',
			ref: 'Other!A1:A2',
			isAnchor: false,
		})
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('renameSheet rejects Excel-invalid target names before mutating workbook', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'renameSheet',
			sheet: 'Sheet1',
			newName: 'Bad/Name',
		})
		expectErr(result)
		expect(result.error.code).toBe('VALIDATION_ERROR')
		expect(result.error.message).toContain('invalid characters')
		expect(wb.getSheet('Sheet1')).toBeDefined()
		expect(wb.getSheet('Bad/Name')).toBeUndefined()
	})

	test('renameSheet updates whole-column references in formulas and defined names', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		const other = wb.addSheet('Other')
		s.cells.set(0, 0, cell(EMPTY, 'SUM(Sheet1!A:A)'))
		other.cells.set(0, 0, cell(EMPTY, 'Sheet1!A1*2'))
		wb.definedNames.set('AllA', 'Sheet1!A:A')

		const result = applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet1', newName: 'Data' })
		expectOk(result)

		expect(s.cells.get(0, 0)?.formula).toBe('SUM(Data!A:A)')
		expect(other.cells.get(0, 0)?.formula).toBe('Data!A1*2')
		expect(wb.definedNames.get('AllA')).toBe('Data!A:A')
		expect(result.value.affectedCells).toContain('Data!A1')
		expect(result.value.affectedCells).toContain('Other!A1')
		expect(result.value.sheetsModified).toEqual(['Data', 'Other'])
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
			rules: [
				{
					type: 'expression',
					formulas: ['SUM(Sheet1!A:A)>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'Sheet1!B:B' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: { cfvo: [{ type: 'formula', value: 'Sheet1!C:C' }] },
					iconSet: { cfvo: [{ type: 'formula', value: 'Sheet1!D:D' }] },
				},
			],
		})
		s.x14DataValidations.push({
			index: 0,
			sqref: 'A2',
			type: 'list',
			formula1: 'Sheet1!A:A',
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A2',
			formulas: ['SUM(Sheet1!A:A)>0'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'Sheet1!B:B' }],
				colors: [{ rgb: 'FF63BE7B' }],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'Sheet1!C:C' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'Sheet1!D:D' }] },
		})
		s.hyperlinks.set('A1', { location: 'sheet1!A1', display: 'jump' })
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
		expect(s.conditionalFormats[0]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe('Data!B:B')
		expect(s.conditionalFormats[0]?.rules[0]?.dataBar?.cfvo[0]?.value).toBe('Data!C:C')
		expect(s.conditionalFormats[0]?.rules[0]?.iconSet?.cfvo[0]?.value).toBe('Data!D:D')
		expect(s.x14DataValidations[0]?.formula1).toBe('Data!A:A')
		expect(s.x14ConditionalFormats[0]?.formulas[0]).toBe('SUM(Data!A:A)>0')
		expect(s.x14ConditionalFormats[0]?.colorScale?.cfvo[0]?.value).toBe('Data!B:B')
		expect(s.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('Data!C:C')
		expect(s.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('Data!D:D')
		expect(s.hyperlinks.get('A1')?.location).toBe('Data!A1')
		expect(s.tables[0]?.columns[0]?.formula).toBe('Data!A:A')
		expect(s.tables[0]?.columns[0]?.totalsRowFormula).toBe('SUM(Data!A:A)')
	})

	test('renameSheet updates hyperlink locations with escaped sheet names', () => {
		const wb = createWorkbook()
		const s = wb.addSheet("Bob's Budget")
		s.hyperlinks.set('A1', { location: "'Bob''s Budget'!A1", display: 'jump' })

		const result = applyOperation(wb, {
			op: 'renameSheet',
			sheet: "Bob's Budget",
			newName: "Bob's Actuals",
		})
		expectOk(result)

		expect(s.hyperlinks.get('A1')?.location).toBe("'Bob''s Actuals'!A1")
	})

	test('renameSheet updates chart ownership and series source refs', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'sheet1',
			chartType: 'lineChart',
			series: [
				{
					nameRef: 'sheet1!$B$1',
					categoryRef: 'sheet1!$A$2:$A$4',
					valueRef: 'sheet1!$B$2:$B$4',
				},
			],
		})

		expectOk(applyOperation(wb, { op: 'renameSheet', sheet: 'Sheet1', newName: 'Data' }))

		expect(wb.chartParts[0]).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Data',
			series: [
				{
					nameRef: 'Data!$B$1',
					categoryRef: 'Data!$A$2:$A$4',
					valueRef: 'Data!$B$2:$B$4',
				},
			],
		})
	})

	test('deleteSheet removes chart parts owned by the deleted sheet', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.addSheet('Sheet2')
		wb.chartParts.push(
			{
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				series: [{ valueRef: 'Sheet1!$A$1:$A$3' }],
			},
			{
				partPath: 'xl/charts/chart2.xml',
				sheetName: 'sheet2',
				series: [{ valueRef: 'Sheet2!$A$1:$A$3' }],
			},
		)

		expectOk(applyOperation(wb, { op: 'deleteSheet', sheet: 'Sheet2' }))

		expect(wb.chartParts).toEqual([
			{
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				series: [{ valueRef: 'Sheet1!$A$1:$A$3' }],
			},
		])
	})

	test('copySheet duplicates visual metadata and chart source refs independently', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		source.preservedXml = {
			partPath: 'xl/worksheets/sheet1.xml',
			relsPath: 'xl/worksheets/_rels/sheet1.xml.rels',
			xml: '<worksheet><sheetData/></worksheet>',
			relsXml: '<Relationships/>',
		}
		source.drawingRefs = { hasDrawing: true, hasLegacyDrawing: false }
		source.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
			name: 'Logo',
		})
		source.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 4,
			name: 'Chart 1',
			relationshipRefs: [
				{
					id: 'rId2',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'sheet1',
			chartType: 'barChart',
			series: [
				{
					nameRef: 'sheet1!$B$1',
					categoryRef: 'sheet1!$A$2:$A$4',
					valueRef: 'sheet1!$B$2:$B$4',
				},
			],
		})

		expectOk(applyOperation(wb, { op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }))
		const copy = wb.getSheet('Copy')
		expect(copy).toBeDefined()
		if (!copy) return

		expect(copy.id).not.toBe(source.id)
		expect(copy.preservedXml).toBeNull()
		expect(copy.imageRefs[0]).not.toBe(source.imageRefs[0])
		expect(copy.drawingObjectRefs[0]).not.toBe(source.drawingObjectRefs[0])
		expect(copy.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing2.xml',
			targetPath: 'xl/media/image2.png',
		})
		expect(copy.drawingObjectRefs[0]?.drawingPartPath).toBe('xl/drawings/drawing2.xml')
		expect(copy.drawingObjectRefs[0]?.relationshipRefs?.[0]?.target).toBe('../charts/chart2.xml')
		expect(wb.chartParts).toHaveLength(2)
		expect(wb.chartParts[1]).toMatchObject({
			partPath: 'xl/charts/chart2.xml',
			sheetName: 'Copy',
			series: [
				{
					nameRef: 'Copy!$B$1',
					categoryRef: 'Copy!$A$2:$A$4',
					valueRef: 'Copy!$B$2:$B$4',
				},
			],
		})

		expectOk(
			applyOperation(wb, {
				op: 'replaceImage',
				sheet: 'Copy',
				name: 'Logo',
				contentBase64: 'BAUG',
				contentType: 'image/png',
			}),
		)
		expect(Array.from(source.imageRefs[0]?.content ?? [])).toEqual([1, 2, 3])
		expect(Array.from(copy.imageRefs[0]?.content ?? [])).toEqual([4, 5, 6])
	})

	test('copySheet retargets copied spill formula binding anchors to the new sheet', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		source.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(2)',
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!A1',
				ref: 'A1:A2',
				isAnchor: true,
			},
		})
		source.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!A1',
				ref: 'A1:A2',
				isAnchor: false,
			},
		})
		source.cells.set(2, 0, {
			value: { kind: 'error', value: '#SPILL!' },
			formula: 'SEQUENCE(2)',
			styleId: sid,
			formulaInfo: {
				kind: 'blockedSpill',
				anchorRef: 'Sheet1!A3',
				ref: 'A3:A4',
				blockingRefs: ['A4'],
			},
		})

		expectOk(applyOperation(wb, { op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }))

		const copy = wb.getSheet('Copy')
		expect(copy?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Copy!A1',
			ref: 'A1:A2',
			isAnchor: true,
		})
		expect(copy?.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Copy!A1',
			ref: 'A1:A2',
			isAnchor: false,
		})
		expect(copy?.cells.get(2, 0)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Copy!A3',
			ref: 'A3:A4',
			blockingRefs: ['A4'],
		})
		expect(source.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A2',
			isAnchor: true,
		})
	})

	test('copySheet retargets copied sheet-qualified formula binding refs', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		source.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'Sheet1!B1*2',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'copy-shared',
				isMaster: true,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		source.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'copy-shared',
				isMaster: false,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		source.cells.set(3, 0, {
			value: numberValue(4),
			formula: 'SUM(Sheet1!C4:Sheet1!D5)',
			styleId: sid,
			formulaInfo: { kind: 'array', ref: 'Sheet1!A4:B5' },
		})
		source.cells.set(6, 0, {
			value: numberValue(7),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'dataTable',
				ref: 'Sheet1!A7:B8',
				dtr: true,
				r1: 'Sheet1!C1',
				r2: 'Sheet1!D1',
			},
		})

		expectOk(applyOperation(wb, { op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }))

		const copy = wb.getSheet('Copy')
		expect(copy?.cells.get(0, 0)?.formula).toBe('Copy!B1*2')
		expect(copy?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: 'copy-shared',
			isMaster: true,
			masterRef: 'Copy!A1',
			ref: 'Copy!A1:A2',
		})
		expect(copy?.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: 'copy-shared',
			isMaster: false,
			masterRef: 'Copy!A1',
			ref: 'Copy!A1:A2',
		})
		expect(copy?.cells.get(3, 0)?.formula).toBe('SUM(Copy!C4:D5)')
		expect(copy?.cells.get(3, 0)?.formulaInfo).toEqual({ kind: 'array', ref: 'Copy!A4:B5' })
		expect(copy?.cells.get(6, 0)?.formulaInfo).toEqual({
			kind: 'dataTable',
			ref: 'Copy!A7:B8',
			dtr: true,
			r1: 'Copy!C1',
			r2: 'Copy!D1',
		})
		expect(source.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: 'copy-shared',
			isMaster: true,
			masterRef: 'Sheet1!A1',
			ref: 'Sheet1!A1:A2',
		})
		expect(source.cells.get(0, 0)?.formula).toBe('Sheet1!B1*2')
		expect(source.cells.get(3, 0)?.formula).toBe('SUM(Sheet1!C4:Sheet1!D5)')
	})

	test('copySheet retargets copied worksheet metadata formulas', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Data')
		source.dataValidations.push({
			sqref: 'A1',
			type: 'list',
			formula1: 'data!B1:B3',
			formula2: 'C1:C3',
		})
		source.conditionalFormats.push({
			sqref: 'A1:A3',
			rules: [
				{
					type: 'expression',
					formulas: ['data!B1>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'data!C1' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: { cfvo: [{ type: 'formula', value: 'data!D1' }] },
					iconSet: { cfvo: [{ type: 'formula', value: 'data!E1' }] },
				},
			],
		})
		source.x14DataValidations.push({
			index: 0,
			sqref: 'B1',
			type: 'list',
			formula1: 'data!B1:B3',
			formula2: 'C1:C3',
		})
		source.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B1:B3',
			formulas: ['data!B1>0'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'data!C1' }],
				colors: [{ rgb: 'FF63BE7B' }],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'data!D1' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'data!E1' }] },
		})

		expectOk(applyOperation(wb, { op: 'copySheet', sheet: 'Data', newName: 'Copy' }))

		const copy = wb.getSheet('Copy')
		expect(copy?.dataValidations[0]).toMatchObject({
			formula1: 'Copy!B1:B3',
			formula2: 'C1:C3',
		})
		expect(copy?.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['Copy!B1>0'])
		expect(copy?.conditionalFormats[0]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe('Copy!C1')
		expect(copy?.conditionalFormats[0]?.rules[0]?.dataBar?.cfvo[0]?.value).toBe('Copy!D1')
		expect(copy?.conditionalFormats[0]?.rules[0]?.iconSet?.cfvo[0]?.value).toBe('Copy!E1')
		expect(copy?.x14DataValidations[0]).toMatchObject({
			formula1: 'Copy!B1:B3',
			formula2: 'C1:C3',
		})
		expect(copy?.x14ConditionalFormats[0]?.formulas).toEqual(['Copy!B1>0'])
		expect(copy?.x14ConditionalFormats[0]?.colorScale?.cfvo[0]?.value).toBe('Copy!C1')
		expect(copy?.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('Copy!D1')
		expect(copy?.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('Copy!E1')

		expect(source.dataValidations[0]?.formula1).toBe('data!B1:B3')
		expect(source.x14ConditionalFormats[0]?.formulas).toEqual(['data!B1>0'])
	})

	test('copySheet rejects duplicate and Excel-invalid target names before mutating workbook', () => {
		const wb = setup()
		wb.addSheet('Existing')

		const duplicate = applyOperation(wb, {
			op: 'copySheet',
			sheet: 'Sheet1',
			newName: 'Existing',
		})
		expectErr(duplicate)
		expect(duplicate.error.code).toBe('NAME_CONFLICT')

		const invalid = applyOperation(wb, {
			op: 'copySheet',
			sheet: 'Sheet1',
			newName: 'Bad/Name',
		})
		expectErr(invalid)
		expect(invalid.error.code).toBe('VALIDATION_ERROR')
		expect(wb.getSheet('Bad/Name')).toBeUndefined()
		expect(wb.sheets).toHaveLength(2)
	})

	test('moveSheet preserves visual metadata and chart source refs', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Sheet1')
		wb.addSheet('Sheet2')
		source.drawingRefs = { hasDrawing: true, hasLegacyDrawing: false }
		source.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
			name: 'Logo',
		})
		source.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 2,
			name: 'Callout',
			text: 'Keep me',
		})
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ valueRef: 'Sheet1!$B$2:$B$4' }],
		})

		expectOk(applyOperation(wb, { op: 'moveSheet', sheet: 'Sheet1', position: 1 }))

		expect(wb.sheets.map((sheet) => sheet.name)).toEqual(['Sheet2', 'Sheet1'])
		expect(source.imageRefs[0]?.name).toBe('Logo')
		expect(source.drawingObjectRefs[0]?.text).toBe('Keep me')
		expect(wb.chartParts[0]).toEqual({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ valueRef: 'Sheet1!$B$2:$B$4' }],
		})
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

	test('mergeCells on a cloned workbook does not mutate the source sheet', () => {
		const wb = setup()
		const clone = wb.clone()
		expectOk(applyOperation(clone, { op: 'mergeCells', sheet: 'Sheet1', range: 'C1:D1' }))

		expect(wb.getSheet('Sheet1')?.merges).toEqual([])
		expect(clone.getSheet('Sheet1')?.merges).toHaveLength(1)
	})

	test('mergeCells rejects overlapping merge ranges', () => {
		const wb = setup()
		expectOk(applyOperation(wb, { op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' }))
		const result = applyOperation(wb, { op: 'mergeCells', sheet: 'Sheet1', range: 'B2:C3' })
		expectErr(result)
		expect(result.error.message).toContain('cannot overlap')
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

	test('setDefinedName rejects Excel-invalid public names', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setDefinedName',
			name: 'A1',
			ref: 'Sheet1!A1:A3',
		})

		expectErr(result)
		expect(result.error.code).toBe('VALIDATION_ERROR')
		expect(wb.definedNames.has('A1')).toBe(false)
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
		const external = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'B1',
			url: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})
		expectOk(external)
		const internal = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'C1',
			location: 'Sheet1!A1',
			display: 'Jump',
		})
		expectOk(internal)

		expect(wb.getSheet('Sheet1')?.hyperlinks.get('B1')).toEqual({
			target: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})
		expect(wb.getSheet('Sheet1')?.hyperlinks.get('C1')).toEqual({
			location: 'Sheet1!A1',
			display: 'Jump',
		})
	})

	test('comment and hyperlink setters skip semantic no-ops without dirtying metadata', () => {
		const wb = setup()
		const sheet = wb.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.comments.set('A1', {
			text: 'Review',
			author: 'Ada',
			legacyDrawing: { shapeId: '_x0000_s1025', row: 0, column: 0 },
		})
		sheet.hyperlinks.set('B1', {
			target: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})

		const comment = applyOperation(wb, {
			op: 'setComment',
			sheet: 'Sheet1',
			ref: 'A1',
			text: 'Review',
			author: 'Ada',
		})
		expectOk(comment)
		expect(comment.value).toEqual({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
		})

		const hyperlink = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'B1',
			url: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})
		expectOk(hyperlink)
		expect(hyperlink.value).toEqual({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
		})
	})

	test('setHyperlink ignores blank destination fields when another destination is valid', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'B1',
			url: '   ',
			location: 'Sheet1!A1',
			display: 'Jump',
		})
		expectOk(result)
		expect(wb.getSheet('Sheet1')?.hyperlinks.get('B1')).toEqual({
			location: 'Sheet1!A1',
			display: 'Jump',
		})
	})

	test('hyperlink operations canonicalize refs and delete case-insensitively', () => {
		const wb = setup()
		const lower = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'b1',
			url: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})
		expectOk(lower)
		expect(wb.getSheet('Sheet1')?.hyperlinks.has('b1')).toBe(false)
		expect(wb.getSheet('Sheet1')?.hyperlinks.get('B1')).toEqual({
			target: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})

		const overwrite = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'B1',
			location: 'Sheet1!A1',
		})
		expectOk(overwrite)
		expect(wb.getSheet('Sheet1')?.hyperlinks.get('B1')).toEqual({
			location: 'Sheet1!A1',
		})

		const deleted = applyOperation(wb, { op: 'deleteHyperlink', sheet: 'Sheet1', ref: 'b1' })
		expectOk(deleted)
		expect(wb.getSheet('Sheet1')?.hyperlinks.size).toBe(0)
	})

	test('setHyperlink rejects links without a url or location', () => {
		const wb = setup()
		const result = applyOperation(wb, {
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: 'B1',
		})
		expect(result.ok).toBe(false)
		if (!result.ok) expect(result.error.code).toBe('VALIDATION_ERROR')
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

	test('style setters skip semantic no-ops without dirtying styles', () => {
		const wb = setup()
		expectOk(
			applyOperation(wb, {
				op: 'setNumberFormat',
				sheet: 'Sheet1',
				range: 'A1:A1',
				format: '0.00%',
			}),
		)
		expectOk(
			applyOperation(wb, {
				op: 'setStyle',
				sheet: 'Sheet1',
				range: 'B1:B1',
				style: { font: { bold: true }, numberFormat: '$0.00' },
			}),
		)

		const numberFormat = applyOperation(wb, {
			op: 'setNumberFormat',
			sheet: 'Sheet1',
			range: 'A1:A1',
			format: '0.00%',
		})
		expectOk(numberFormat)
		expect(numberFormat.value).toEqual({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
		})

		const style = applyOperation(wb, {
			op: 'setStyle',
			sheet: 'Sheet1',
			range: 'B1:B1',
			style: { font: { bold: true }, numberFormat: '$0.00' },
		})
		expectOk(style)
		expect(style.value).toEqual({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
		})
	})

	test('sortRange sorts a block by header name and moves metadata with rows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Calc'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('B'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(20), formula: 'B2*10', styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('A'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(10), formula: 'B3*10', styleId: sid })
		sheet.hyperlinks.set('A2', { target: 'https://example.com/b' })
		sheet.comments.set('B3', {
			text: 'lowest',
			legacyDrawing: {
				shapeId: '_x0000_s1027',
				row: 2,
				column: 1,
				anchor: [1, 15, 2, 2, 3, 15, 5, 16],
			},
		})
		sheet.threadedComments.push({ ref: 'B3', text: 'lowest thread', id: 'tc-lowest' })
		sheet.dataValidations.push({ sqref: 'A2:B2', type: 'whole', formula1: 'A2' })
		sheet.conditionalFormats.push({
			sqref: 'B3',
			rules: [
				{
					type: 'expression',
					formulas: ['A3>0'],
					dataBar: { cfvo: [{ type: 'formula', value: 'A3' }] },
					iconSet: { cfvo: [{ type: 'formula', value: 'B3' }] },
				},
			],
		})
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'A2:B2',
			type: 'whole',
			formula1: 'A2',
			formula2: 'B2',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B3',
			formulas: ['A3>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'A3' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'B3' }] },
		})
		sheet.ignoredErrors.push({ sqref: 'A2', formula: true })
		sheet.rowHeights.set(1, 24)
		sheet.rowDefs.set(1, { hidden: true, customHeight: true })
		sheet.rowDefs.set(2, { outlineLevel: 2, collapsed: true })

		const result = applyOperation(wb, {
			op: 'sortRange',
			sheet: 'Sheet1',
			range: 'A1:C3',
			by: [{ column: 'Score' }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A2', 'B2', 'C2', 'A3', 'B3', 'C3'])
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('A'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 2)?.formula).toBe('B2*10')
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('B'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 2)?.formula).toBe('B3*10')
		expect(sheet.hyperlinks.get('A3')).toEqual({ target: 'https://example.com/b' })
		expect(sheet.comments.get('B2')).toEqual({
			text: 'lowest',
			legacyDrawing: {
				shapeId: '_x0000_s1027',
				row: 1,
				column: 1,
				anchor: [1, 15, 1, 2, 3, 15, 4, 16],
			},
		})
		expect(sheet.threadedComments).toEqual([{ ref: 'B2', text: 'lowest thread', id: 'tc-lowest' }])
		expect(sheet.dataValidations[0]).toMatchObject({ sqref: 'A3:B3', formula1: 'A3' })
		expect(sheet.conditionalFormats[0]?.sqref).toBe('B2')
		expect(sheet.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['A2>0'])
		expect(sheet.conditionalFormats[0]?.rules[0]?.dataBar?.cfvo[0]?.value).toBe('A2')
		expect(sheet.conditionalFormats[0]?.rules[0]?.iconSet?.cfvo[0]?.value).toBe('B2')
		expect(sheet.x14DataValidations[0]).toMatchObject({
			index: 0,
			sqref: 'A3:B3',
			formula1: 'A3',
			formula2: 'B3',
		})
		expect(sheet.x14ConditionalFormats[0]?.sqref).toBe('B2')
		expect(sheet.x14ConditionalFormats[0]?.formulas).toEqual(['A2>0'])
		expect(sheet.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('A2')
		expect(sheet.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('B2')
		expect(sheet.ignoredErrors[0]?.sqref).toBe('A3')
		expect(sheet.rowHeights.get(2)).toBe(24)
		expect(sheet.rowDefs.get(1)).toEqual({ outlineLevel: 2, collapsed: true })
		expect(sheet.rowDefs.get(2)).toEqual({ hidden: true, customHeight: true })
	})

	test('sortRange reports first-row affected cells when sorting by column letter', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('B'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('A'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(1), formula: null, styleId: sid })

		const result = applyOperation(wb, {
			op: 'sortRange',
			sheet: 'Sheet1',
			range: 'A1:B2',
			by: [{ column: 'A' }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'B1', 'A2', 'B2'])
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('A'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('B'))
	})

	test('sortRange materializes imported shared formulas before sorting', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
		addSharedFormulaGroup(sheet, 0, 3)

		const result = applyOperation(wb, {
			op: 'sortRange',
			sheet: 'Sheet1',
			range: 'A1:D2',
			by: [{ column: 'A' }],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['D1', 'D2', 'A2', 'B2', 'C2'])
		expect(sheet.cells.get(0, 3)?.formula).toBe('B1*2')
		expect(sheet.cells.get(1, 3)?.formula).toBe('B2*2')
		expect(sheet.cells.get(0, 3)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 3)?.formulaInfo).toBeUndefined()
	})

	test('sortRange materializes non-shared formula metadata before sorting', () => {
		const cases: readonly {
			readonly name: string
			readonly setup: (sheet: Sheet) => void
			readonly range: string
			readonly affectedCells: readonly string[]
			readonly assert: (sheet: Sheet) => void
		}[] = [
			{
				name: 'dynamic spill range',
				setup: (sheet) => {
					sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
					sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
					sheet.cells.set(2, 0, { value: stringValue('c'), formula: null, styleId: sid })
					addDynamicArrayAnchorWithStaleSpillFootprint(sheet, 'Sheet1!D1')
				},
				range: 'A1:D3',
				affectedCells: ['A1', 'A2', 'A3', 'B1', 'C1', 'D1', 'B2', 'C2', 'D2', 'B3', 'C3', 'D3'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 3)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(1, 3)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(2, 3)?.formulaInfo).toBeUndefined()
				},
			},
			{
				name: 'data-table range',
				setup: (sheet) => {
					for (let row = 0; row < 5; row++) {
						sheet.cells.set(row, 0, {
							value: stringValue(String.fromCharCode(97 + row)),
							formula: null,
							styleId: sid,
						})
					}
					addDataTableFormula(sheet)
				},
				range: 'A1:C5',
				affectedCells: ['C3', 'A2', 'B2', 'C2', 'A3', 'B3', 'A4', 'B4', 'C4', 'A5', 'B5', 'C5'],
				assert: (sheet) => {
					expect(sheet.cells.get(2, 2)?.formulaInfo).toBeUndefined()
					expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(20))
					expect(sheet.cells.get(4, 2)?.value).toEqual(numberValue(30))
				},
			},
			{
				name: 'blocked spill anchor',
				setup: addBlockedSpillFormula,
				range: 'A1:A2',
				affectedCells: ['A1', 'A2'],
				assert: (sheet) => {
					expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('blocker'))
					expect(sheet.cells.get(1, 0)?.formula).toBe('SEQUENCE(3)')
					expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
				},
			},
		]

		for (const entry of cases) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			entry.setup(sheet)

			const result = applyOperation(wb, {
				op: 'sortRange',
				sheet: 'Sheet1',
				range: entry.range,
				by: [{ column: 'A' }],
			})
			expectOk(result)

			expect(result.value.affectedCells, entry.name).toEqual(entry.affectedCells)
			entry.assert(sheet)
		}
	})

	test('sortRange preserves formula metadata outside the sorted rows', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Source')
		const other = wb.addSheet('Other')
		source.cells.set(0, 0, { value: stringValue('b'), formula: null, styleId: sid })
		source.cells.set(1, 0, { value: stringValue('a'), formula: null, styleId: sid })
		addSharedFormulaGroup(source, 0, 3)
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
		expect(source.cells.get(0, 3)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'D1',
			ref: 'D1:D2',
		})
		expect(source.cells.get(1, 3)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'D1',
		})
		expect(other.cells.get(0, 0)?.formulaInfo).toEqual({ kind: 'array', ref: 'A1:A2' })
	})

	test('sortRange rejects legacy array formula intersections', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const formulaInfo = { kind: 'array' as const, ref: 'A1:B2' }
		sheet.cells.set(0, 0, { value: numberValue(1), formula: 'A4:B5', styleId: sid, formulaInfo })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid, formulaInfo })

		expectErr(
			applyOperation(wb, {
				op: 'sortRange',
				sheet: 'Sheet1',
				range: 'A1:B2',
				by: [{ column: 'A' }],
			}),
		)

		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual(formulaInfo)
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(4))
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
		expect(result.value.recalcRequired).toBe(true)
	})

	test('createTable invalidates structured references that target the new table', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(3, 0, {
			value: EMPTY,
			formula: 'SUM(BalanceTable[Value])',
			styleId: sid,
		})

		expect(recalculate(wb, defaultCalcContext()).errors).toEqual([])

		const result = applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'BalanceTable',
			hasHeaders: true,
		})
		expectOk(result)
		expect(result.value.recalcRequired).toBe(true)

		expect(recalculate(wb, defaultCalcContext()).errors).toEqual([])
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(10))
	})

	test('createTable rejects workbook-scoped duplicate table names', () => {
		const wb = createWorkbook()
		const sheet1 = wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		sheet1.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet1.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet2',
			ref: 'A1:B2',
			name: 'sales',
			hasHeaders: true,
		})

		expectErr(result)
		expect(result.error.message).toContain('already exists')
		expect(sheet2.tables).toHaveLength(0)
	})

	test('createTable rejects Excel-invalid table names before mutating topology', () => {
		const invalidNames = [
			'',
			'1Sales',
			'Sales Data',
			'A1',
			'XFD1048576',
			'R1C1',
			'C',
			'R',
			`T${'x'.repeat(255)}`,
		]
		for (const name of invalidNames) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')

			const result = applyOperation(wb, {
				op: 'createTable',
				sheet: 'Sheet1',
				ref: 'A1:B2',
				name,
				hasHeaders: true,
			})

			expectErr(result)
			expect(result.error.code).toBe('VALIDATION_ERROR')
			expect(result.error.suggestedFix).toContain('A1-style')
			expect(sheet.tables).toHaveLength(0)
			expect(sheet.autoFilter).toBeNull()
		}
	})

	test('createTable rejects ranges that overlap existing tables', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
			columns: [{ name: 'Region' }, { name: 'Rep' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'C3:D5',
			name: 'Forecast',
			hasHeaders: true,
		})

		expectErr(result)
		expect(result.error.message).toContain('overlaps table "Sales"')
		expect(result.error.details).toMatchObject({
			kind: 'overlapping-table-ranges',
			operation: 'create',
			tableName: 'Forecast',
			ref: 'C3:D5',
			overlappingTable: { tableName: 'Sales', ref: 'A1:C3' },
		})
		expect(sheet.tables.map((table) => table.name)).toEqual(['Sales'])
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

	test('appendRows detaches imported formula bindings under expanded table rows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Value'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A1',
			name: 'Values',
			hasHeaders: true,
		})
		addSharedFormulaGroup(sheet, 1, 0)

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'Values',
			rows: [[9]],
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A2', 'A3'])
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(1, 0)?.formula).toBeNull()
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formula).toBe('B2*2')
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
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

	test('appendRows validates row width before shifting totals rows', () => {
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
			rows: [['Debt', 20]],
		})
		expectErr(result)

		expect(sheet.tables[0]?.ref.end.row).toBe(1)
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('Cash'))
		expect(sheet.cells.get(2, 0)).toBeUndefined()
	})

	test('appendRows rejects expansion into another table range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push(
			{
				id: createTableId(),
				name: 'Sales',
				sheetId: sheet.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
				columns: [{ name: 'Name' }, { name: 'Amount' }],
				hasHeaders: true,
				hasTotals: false,
			},
			{
				id: createTableId(),
				name: 'Forecast',
				sheetId: sheet.id,
				ref: { start: { row: 2, col: 0 }, end: { row: 3, col: 1 } },
				columns: [{ name: 'Name' }, { name: 'Amount' }],
				hasHeaders: true,
				hasTotals: false,
			},
		)

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'Sales',
			rows: [['Debt', 20]],
		})

		expectErr(result)
		expect(result.error.message).toContain('overlaps table "Forecast"')
		expect(result.error.details).toMatchObject({
			kind: 'overlapping-table-ranges',
			operation: 'append',
			tableName: 'Sales',
			ref: 'A1:B3',
			overlappingTable: { tableName: 'Forecast', ref: 'A3:B4' },
		})
		expect(sheet.tables[0]?.ref.end.row).toBe(1)
		expect(sheet.cells.get(2, 0)).toBeUndefined()
	})

	test('appendRows rejects totals-row insertion that would shift another table', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push(
			{
				id: createTableId(),
				name: 'Sales',
				sheetId: sheet.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
				columns: [{ name: 'Name' }, { name: 'Amount' }],
				hasHeaders: true,
				hasTotals: true,
			},
			{
				id: createTableId(),
				name: 'Forecast',
				sheetId: sheet.id,
				ref: { start: { row: 4, col: 3 }, end: { row: 5, col: 4 } },
				columns: [{ name: 'Scenario' }, { name: 'Value' }],
				hasHeaders: true,
				hasTotals: false,
			},
		)

		const result = applyOperation(wb, {
			op: 'appendRows',
			table: 'Sales',
			rows: [['Debt', 20]],
		})

		expectErr(result)
		expect(result.error.message).toContain('would shift table "Forecast"')
		expect(result.error.details).toMatchObject({
			kind: 'table-totals-append-would-shift-table',
			tableName: 'Sales',
			ref: 'A1:B3',
			shiftedTable: { tableName: 'Forecast', ref: 'D5:E6' },
		})
		expect(sheet.tables[0]?.ref.end.row).toBe(2)
		expect(sheet.tables[1]?.ref.start.row).toBe(4)
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
		const table = sheet.tables[0]
		const totalColumn = table?.columns[2]
		if (table && totalColumn) {
			const columns = [...table.columns]
			columns[2] = { ...totalColumn, formula: 'OLD()', formulaIsArray: true }
			sheet.tables[0] = { ...table, columns }
		}

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
		expect(sheet.tables[0]?.columns[2]?.formulaIsArray).toBeUndefined()
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

	test('setTableColumn formula materializes imported formula bindings under the body column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Amount'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A2',
			name: 'Sales',
			hasHeaders: true,
		})
		addSharedFormulaGroup(sheet, 1, 0)

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Amount',
			formula: '=1+1',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A2', 'A3'])
		expect(sheet.cells.get(1, 0)?.formula).toBe('1+1')
		expect(sheet.cells.get(1, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(2, 0)?.formula).toBe('B2*2')
		expect(sheet.cells.get(2, 0)?.formulaInfo).toBeUndefined()
	})

	test('setTableColumn rejects legacy array intersections before mutating table metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Amount'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A2',
			name: 'Sales',
			hasHeaders: true,
		})
		const formulaInfo = { kind: 'array' as const, ref: 'A2:A3' }
		sheet.cells.set(1, 0, {
			value: numberValue(2),
			formula: 'SEQUENCE(2)',
			styleId: sid,
			formulaInfo,
		})
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid, formulaInfo })

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Amount',
			formula: '=1+1',
		})
		expectErr(result)

		expect(sheet.tables[0]?.columns[0]?.formula).toBeUndefined()
		expect(sheet.cells.get(1, 0)?.formula).toBe('SEQUENCE(2)')
		expect(sheet.cells.get(1, 0)?.formulaInfo).toEqual(formulaInfo)
		expect(sheet.cells.get(2, 0)?.formulaInfo).toEqual(formulaInfo)
	})

	test('setTableColumn rename materializes imported shared formulas before structured ref rewrites', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'Sales',
			hasHeaders: true,
		})
		sheet.cells.set(4, 0, {
			value: numberValue(2),
			formula: 'SUM(Sales[Qty])',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'rename-shared',
				isMaster: true,
				masterRef: 'A5',
				ref: 'A5:A6',
			},
		})
		sheet.cells.set(5, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'rename-shared',
				isMaster: false,
				masterRef: 'A5',
			},
		})

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Qty',
			newName: 'Units',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A5', 'A6', 'A1'])
		expect(sheet.cells.get(4, 0)?.formula).toBe('SUM(Sales[Units])')
		expect(sheet.cells.get(5, 0)?.formula).toBe('SUM(Sales[Units])')
		expect(sheet.cells.get(4, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(5, 0)?.formulaInfo).toBeUndefined()
	})

	test('setTableColumn rename preserves dynamic spill metadata and reports rewritten formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'Sales',
			hasHeaders: true,
		})
		sheet.cells.set(0, 3, {
			value: numberValue(2),
			formula: 'FILTER(Sales[Qty],Sales[Qty]>0)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(1, 3, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!D1',
				ref: 'D1:D2',
				isAnchor: false,
			},
		})

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Qty',
			newName: 'Units',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A1', 'D1', 'D2'])
		expect(sheet.cells.get(0, 3)?.formula).toBe('FILTER(Sales[Units],Sales[Units]>0)')
		expect(sheet.cells.get(0, 3)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(sheet.cells.get(1, 3)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!D1',
			ref: 'D1:D2',
			isAnchor: false,
		})
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('setTableColumn renames columns and rewrites structured references', () => {
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
		applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Total',
			formula: '=[@Qty]*[@Price]',
			totalsRowFormula: '=SUM(Sales[Qty])',
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Sales[[Qty]:[Price]])', styleId: sid })
		wb.definedNames.set('SalesQty', 'SUM(Sales[Qty])')
		sheet.dataValidations.push({
			sqref: 'D2:D3',
			type: 'list',
			formula1: 'SUM(Sales[Qty])',
		})
		sheet.conditionalFormats.push({
			sqref: 'E2:E3',
			rules: [
				{
					type: 'expression',
					formulas: ['SUM(Sales[Qty])>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
						color: { rgb: 'FF00AA00' },
					},
					iconSet: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
					},
				},
			],
		})
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'F2:F3',
			type: 'list',
			formula1: 'SUM(Sales[Qty])',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'G2:G3',
			formulas: ['SUM(Sales[Qty])>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }] },
		})
		sheet.dataValidations.push({
			sqref: 'A2:A3',
			type: 'list',
			formula1: '[@Qty]',
		})
		sheet.dataValidations.push({
			sqref: 'H2:H3',
			type: 'list',
			formula1: '[@Qty]',
		})
		sheet.conditionalFormats.push({
			sqref: 'A2:C3',
			rules: [
				{
					type: 'expression',
					formulas: ['[@Qty]>0'],
					dataBar: { cfvo: [{ type: 'formula', value: '[@Qty]' }] },
				},
			],
		})
		sheet.x14DataValidations.push({
			index: 1,
			sqref: 'B2:B3',
			type: 'list',
			formula1: '[@Qty]',
		})
		sheet.x14ConditionalFormats.push({
			index: 1,
			sqref: 'C2:C3',
			formulas: ['[@Qty]>0'],
			dataBar: { cfvo: [{ type: 'formula', value: '[@Qty]' }] },
		})

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Qty',
			newName: 'Units',
		})
		expectOk(result)
		expect(result.value.affectedCells).toEqual(['A1', 'C2', 'C3', 'A5'])
		expect(result.value.recalcRequired).toBe(true)
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Units',
			'Price',
			'Total',
		])
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('Units'))
		expect(sheet.cells.get(1, 2)?.formula).toBe('[@Units]*[@Price]')
		expect(sheet.tables[0]?.columns[2]?.formula).toBe('[@Units]*[@Price]')
		expect(sheet.tables[0]?.columns[2]?.totalsRowFormula).toBe('SUM(Sales[Units])')
		expect(sheet.cells.get(4, 0)?.formula).toBe('SUM(Sales[[Units]:[Price]])')
		expect(wb.definedNames.get('SalesQty')).toBe('SUM(Sales[Units])')
		expect(sheet.dataValidations[0]?.formula1).toBe('SUM(Sales[Units])')
		const rule = sheet.conditionalFormats[0]?.rules[0]
		expect(rule?.formulas[0]).toBe('SUM(Sales[Units])>0')
		expect(rule?.colorScale?.cfvo[0]?.value).toBe('SUM(Sales[Units])')
		expect(rule?.dataBar?.cfvo[0]?.value).toBe('SUM(Sales[Units])')
		expect(rule?.iconSet?.cfvo[0]?.value).toBe('SUM(Sales[Units])')
		expect(sheet.x14DataValidations[0]?.formula1).toBe('SUM(Sales[Units])')
		expect(sheet.x14ConditionalFormats[0]?.formulas[0]).toBe('SUM(Sales[Units])>0')
		expect(sheet.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('SUM(Sales[Units])')
		expect(sheet.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('SUM(Sales[Units])')
		expect(sheet.dataValidations[1]?.formula1).toBe('[@Units]')
		expect(sheet.dataValidations[2]?.formula1).toBe('[@Qty]')
		const localRule = sheet.conditionalFormats[1]?.rules[0]
		expect(localRule?.formulas[0]).toBe('[@Units]>0')
		expect(localRule?.dataBar?.cfvo[0]?.value).toBe('[@Units]')
		expect(sheet.x14DataValidations[1]?.formula1).toBe('[@Units]')
		expect(sheet.x14ConditionalFormats[1]?.formulas[0]).toBe('[@Units]>0')
		expect(sheet.x14ConditionalFormats[1]?.dataBar?.cfvo[0]?.value).toBe('[@Units]')

		const duplicate = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Units',
			newName: 'Price',
		})
		expectErr(duplicate)
	})

	test('setTableColumn on a cloned workbook does not mutate the source table', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, cell(stringValue('Qty')))
		sheet.cells.set(0, 1, cell(stringValue('Amount')))
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Qty' }, { name: 'Amount', formula: '[@Qty]*10' }],
			hasHeaders: true,
			hasTotals: false,
		})
		const clone = wb.clone()

		expectOk(
			applyOperation(clone, {
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Amount',
				newName: 'Revenue',
				formula: '=[@Qty]*12',
			}),
		)

		expect(wb.getSheet('Sheet1')?.tables[0]?.columns[1]).toMatchObject({
			name: 'Amount',
			formula: '[@Qty]*10',
		})
		expect(clone.getSheet('Sheet1')?.tables[0]?.columns[1]).toMatchObject({
			name: 'Revenue',
			formula: '[@Qty]*12',
		})
	})

	test('setTableColumn rewrites table-scoped worksheet metadata local structured refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Qty' },
				{ id: 2, name: 'Price' },
				{ id: 3, name: 'Total' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.dataValidations.push({
			sqref: 'A2:C4',
			type: 'list',
			formula1: 'SUM([@Qty])',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A2:C4',
			formulas: ['[@Qty]>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'SUM([@Qty])' }] },
		})
		sheet.conditionalFormats.push({
			sqref: 'E2:E4',
			rules: [{ type: 'expression', formulas: ['[@Qty]>0'] }],
		})

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Qty',
			newName: 'Units',
		})

		expectOk(result)
		expect(sheet.dataValidations[0]?.formula1).toBe('SUM([@Units])')
		expect(sheet.x14ConditionalFormats[0]?.formulas[0]).toBe('[@Units]>0')
		expect(sheet.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('SUM([@Units])')
		expect(sheet.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('[@Qty]>0')
	})

	test('setTableColumn escapes renamed structured reference columns with Excel special characters', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:C2',
			name: 'Sales',
			hasHeaders: true,
		})
		applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Total',
			formula: '=[@Qty]*[@Price]',
			totalsRowFormula: '=SUM(Sales[Qty])',
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Sales[[Qty]:[Price]])', styleId: sid })

		const result = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Qty',
			newName: 'Units#',
		})
		expectOk(result)

		expect(sheet.cells.get(1, 2)?.formula).toBe("[@Units'#]*[@Price]")
		expect(sheet.tables[0]?.columns[2]?.formula).toBe("[@Units'#]*[@Price]")
		expect(sheet.tables[0]?.columns[2]?.totalsRowFormula).toBe("SUM(Sales[Units'#])")
		expect(sheet.cells.get(4, 0)?.formula).toBe("SUM(Sales[[Units'#]:[Price]])")
	})

	test('setTableColumn materializes totals-row cells from metadata edits', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B3',
			name: 'Sales',
			hasHeaders: true,
		})
		const table = sheet.tables[0]
		if (!table) throw new Error('expected table')
		sheet.tables[0] = { ...table, hasTotals: true }

		const label = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Name',
			totalsRowLabel: 'Total',
		})
		const sum = applyOperation(wb, {
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Amount',
			totalsRowFunction: 'sum',
		})
		expectOk(label)
		expectOk(sum)
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('Total'))
		expect(sheet.cells.get(2, 1)?.formula).toBe('SUBTOTAL(109,Sales[Amount])')
		expect(sum.value.recalcRequired).toBe(true)

		const appended = applyOperation(wb, {
			op: 'appendRows',
			table: 'Sales',
			rows: [['Debt', 20]],
		})
		expectOk(appended)
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('Debt'))
		expect(sheet.cells.get(3, 0)?.value).toEqual(stringValue('Total'))
		expect(sheet.cells.get(3, 1)?.formula).toBe('SUBTOTAL(109,Sales[Amount])')
	})

	test('setTableStyle edits table style metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B1',
			name: 'Sales',
			hasHeaders: true,
		})

		const styled = applyOperation(wb, {
			op: 'setTableStyle',
			table: 'Sales',
			styleName: 'TableStyleMedium2',
			showFirstColumn: false,
			showLastColumn: true,
			showRowStripes: true,
			showColumnStripes: false,
		})
		expectOk(styled)
		expect(styled.value.recalcRequired).toBe(false)
		expect(sheet.tables[0]?.tableStyleInfo).toEqual({
			name: 'TableStyleMedium2',
			showFirstColumn: false,
			showLastColumn: true,
			showRowStripes: true,
			showColumnStripes: false,
		})

		const clearedName = applyOperation(wb, {
			op: 'setTableStyle',
			table: 'Sales',
			styleName: null,
		})
		expectOk(clearedName)
		expect(sheet.tables[0]?.tableStyleInfo).toEqual({
			showFirstColumn: false,
			showLastColumn: true,
			showRowStripes: true,
			showColumnStripes: false,
		})

		const missingField = applyOperation(wb, { op: 'setTableStyle', table: 'Sales' })
		expectErr(missingField)
	})

	test('resizeTable preserves overlapping tableColumn metadata when shrinking columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 10, name: 'Region', dataCellStyle: 'Input' },
				{ id: 11, name: 'Amount', formula: '[@Units]*[@Price]', totalsRowFunction: 'sum' },
				{ id: 12, name: 'Notes', totalsRowLabel: 'Memo' },
			],
			hasHeaders: true,
			hasTotals: true,
		})

		expectOk(applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:B4' }))

		expect(sheet.tables[0]?.columns).toEqual([
			{ id: 10, name: 'Region', dataCellStyle: 'Input' },
			{ id: 11, name: 'Amount', formula: '[@Units]*[@Price]', totalsRowFunction: 'sum' },
		])
	})

	test('resizeTable preserves overlapping tableColumn metadata when shifting the range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 3, { value: stringValue('Forecast'), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount', formula: '[@Units]*[@Price]' },
				{ id: 3, name: 'Status', totalsRowLabel: 'Open' },
			],
			hasHeaders: true,
			hasTotals: true,
		})

		expectOk(applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'B1:D4' }))

		expect(sheet.tables[0]?.columns).toEqual([
			{ id: 2, name: 'Amount', formula: '[@Units]*[@Price]' },
			{ id: 3, name: 'Status', totalsRowLabel: 'Open' },
			{ id: 4, name: 'Forecast' },
		])
	})

	test('resizeTable remaps table filter criteria and prunes dropped sort conditions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 3, { value: stringValue('Forecast'), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
				{ id: 3, name: 'Status' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:C4',
				columns: [
					{ colId: 0, kind: 'filters', values: ['West'] },
					{ colId: 2, kind: 'filters', values: ['Open'] },
				],
				sortState: {
					ref: 'A1:C4',
					conditions: [{ ref: 'A2:A4' }, { ref: 'C2:C4', descending: true }],
				},
			},
			sortState: {
				ref: 'A1:C4',
				conditions: [{ ref: 'A2:A4' }, { ref: 'C2:C4', descending: true }],
			},
		})

		expectOk(applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'B1:D4' }))

		expect(sheet.tables[0]?.autoFilter).toEqual({
			ref: 'B1:D4',
			columns: [{ colId: 1, kind: 'filters', values: ['Open'] }],
			sortState: {
				ref: 'B1:D4',
				conditions: [{ ref: 'C2:C4', descending: true }],
			},
		})
		expect(sheet.tables[0]?.sortState).toEqual({
			ref: 'B1:D4',
			conditions: [{ ref: 'C2:C4', descending: true }],
		})
	})

	test('resizeTable expands table sort condition row spans when the table grows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:B2',
				columns: [{ colId: 1, kind: 'filters', values: ['10'] }],
				sortState: { ref: 'A1:B2', conditions: [{ ref: 'B2:B2', descending: true }] },
			},
			sortState: { ref: 'A1:B2', conditions: [{ ref: 'A1:A2' }] },
		})

		expectOk(applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:B4' }))

		expect(sheet.tables[0]?.autoFilter).toEqual({
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'filters', values: ['10'] }],
			sortState: { ref: 'A1:B4', conditions: [{ ref: 'B2:B4', descending: true }] },
		})
		expect(sheet.tables[0]?.sortState).toEqual({
			ref: 'A1:B4',
			conditions: [{ ref: 'A1:A4' }],
		})
	})

	test('resizeTable removes table sort states when all conditions target dropped columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
				{ id: 3, name: 'Status' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:C4',
				columns: [{ colId: 2, kind: 'filters', values: ['Open'] }],
				sortState: { ref: 'A1:C4', conditions: [{ ref: 'C2:C4', descending: true }] },
			},
			sortState: { ref: 'A1:C4', conditions: [{ ref: 'C2:C4', descending: true }] },
		})

		expectOk(applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:B4' }))

		expect(sheet.tables[0]?.autoFilter).toEqual({
			ref: 'A1:B4',
			columns: [],
		})
		expect(sheet.tables[0]?.sortState).toBeUndefined()
	})

	test('resizeTable preserves existing tableColumn metadata when expanding columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 2, { value: stringValue('Forecast'), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } },
			columns: [
				{ id: 1, name: 'Region', queryTableFieldId: 7 },
				{ id: 2, name: 'Amount', formula: 'SUM(A2:A4)', totalsRowFormula: 'SUM([Amount])' },
			],
			hasHeaders: true,
			hasTotals: true,
		})

		expectOk(applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:C4' }))

		expect(sheet.tables[0]?.columns).toEqual([
			{ id: 1, name: 'Region', queryTableFieldId: 7 },
			{ id: 2, name: 'Amount', formula: 'SUM(A2:A4)', totalsRowFormula: 'SUM([Amount])' },
			{ id: 3, name: 'Forecast' },
		])
	})

	test('resizeTable rejects ranges that overlap another table', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push(
			{
				id: createTableId(),
				name: 'Sales',
				sheetId: sheet.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } },
				columns: [
					{ id: 1, name: 'Region' },
					{ id: 2, name: 'Amount' },
				],
				hasHeaders: true,
				hasTotals: false,
			},
			{
				id: createTableId(),
				name: 'Forecast',
				sheetId: sheet.id,
				ref: { start: { row: 0, col: 3 }, end: { row: 3, col: 4 } },
				columns: [
					{ id: 3, name: 'Scenario' },
					{ id: 4, name: 'Value' },
				],
				hasHeaders: true,
				hasTotals: false,
			},
		)

		const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'B1:D4' })

		expectErr(result)
		expect(result.error.message).toContain('overlaps table "Forecast"')
		expect(result.error.details).toMatchObject({
			kind: 'overlapping-table-ranges',
			operation: 'resize',
			tableName: 'Sales',
			ref: 'B1:D4',
			overlappingTable: { tableName: 'Forecast', ref: 'D1:E4' },
		})
		expect(sheet.tables[0]?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 3, col: 1 },
		})
	})

	test('resizeTable rejects dropping table fields referenced by cell formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'COUNTA(Sales[Rep])', styleId: sid })

		const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:A4' })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Region',
			'Rep',
			'Amount',
		])
	})

	test('resizeTable rejects dropping table fields referenced by defined names', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		wb.definedNames.set('SalesReps', 'COUNTA(Sales[Rep])')

		const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:A4' })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(wb.definedNames.get('SalesReps')).toBe('COUNTA(Sales[Rep])')
	})

	test('resizeTable rejects dropping table fields referenced by worksheet metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.dataValidations.push({ sqref: 'E1:E4', type: 'list', formula1: 'Sales[Rep]' })

		const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:A4' })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Rep]')
		expect(sheet.dataValidations[0]?.formula1).toBe('Sales[Rep]')
	})

	test('resizeTable rejects dropped table fields across worksheet metadata formula surfaces', () => {
		for (const scenario of TABLE_FIELD_METADATA_BLOCKERS) {
			const { wb, sheet } = setupSalesRepTable()
			scenario.add(sheet)

			const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:A4' })

			expectErr(result)
			const diagnostic = `${scenario.label}: ${result.error.message}`
			expect(diagnostic).toContain(`${scenario.sourceRef} ${scenario.sourceKind}`)
			expect(diagnostic).toContain('Sales[Rep]')
			expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
			expect(sheet.tables[0]?.ref).toEqual({
				start: { row: 0, col: 0 },
				end: { row: 3, col: 2 },
			})
			expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
				'Region',
				'Rep',
				'Amount',
			])
		}
	})

	test('resizeTable rejects dropped table fields across local worksheet metadata formula surfaces', () => {
		for (const scenario of LOCAL_TABLE_FIELD_METADATA_BLOCKERS) {
			const { wb, sheet } = setupSalesRepTable()
			scenario.add(sheet)

			const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:A4' })

			expectErr(result)
			const diagnostic = `${scenario.label}: ${result.error.message}`
			expect(diagnostic).toContain(`${scenario.sourceRef} ${scenario.sourceKind}`)
			expect(diagnostic).toContain('Sales[Rep]')
			expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
			expect(sheet.tables[0]?.ref).toEqual({
				start: { row: 0, col: 0 },
				end: { row: 3, col: 2 },
			})
			expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
				'Region',
				'Rep',
				'Amount',
			])
		}
	})

	test('resizeTable rejects shifted ranges dropping referenced left-edge fields', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 3, { value: stringValue('Forecast'), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
				{ id: 3, name: 'Status' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'COUNTA(Sales[Region])', styleId: sid })

		const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'B1:D4' })

		expectErr(result)
		expect(result.error.message).toContain('Sales[Region]')
		expect(sheet.tables[0]?.ref.start.col).toBe(0)
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Region',
			'Amount',
			'Status',
		])
	})

	test('resizeTable materializes imported shared formulas and reports detached cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:B2',
			name: 'Sales',
			hasHeaders: true,
		})
		addSharedFormulaGroup(sheet, 0, 3)

		const result = applyOperation(wb, { op: 'resizeTable', table: 'Sales', ref: 'A1:C2' })
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['D1', 'D2'])
		expect(sheet.cells.get(0, 3)?.formula).toBe('B1*2')
		expect(sheet.cells.get(1, 3)?.formula).toBe('B2*2')
		expect(sheet.cells.get(0, 3)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 3)?.formulaInfo).toBeUndefined()
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('deleteTable rejects structured references that would outlive the table', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Rep' },
				{ id: 3, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'COUNTA(Sales[Rep])', styleId: sid })

		const result = applyOperation(wb, { op: 'deleteTable', table: 'Sales' })

		expectErr(result)
		expect(result.error.message).toContain('Sheet1!A6 cell formula')
		expect(result.error.message).toContain('Sales[Rep]')
		expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
		expect(sheet.tables).toHaveLength(1)
		expect(sheet.cells.get(5, 0)?.formula).toBe('COUNTA(Sales[Rep])')
	})

	test('deleteTable materializes imported shared formulas and reports detached cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A2',
			name: 'Sales',
			hasHeaders: true,
		})
		addSharedFormulaGroup(sheet, 0, 3)

		const result = applyOperation(wb, { op: 'deleteTable', table: 'Sales' })
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['D1', 'D2'])
		expect(sheet.tables).toEqual([])
		expect(sheet.cells.get(0, 3)?.formula).toBe('B1*2')
		expect(sheet.cells.get(1, 3)?.formula).toBe('B2*2')
		expect(sheet.cells.get(0, 3)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 3)?.formulaInfo).toBeUndefined()
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('deleteTable rejects defined names and worksheet metadata structured references', () => {
		{
			const { wb, sheet } = setupSalesRepTable()
			wb.definedNames.set('SalesReps', 'COUNTA(Sales[Rep])')

			const result = applyOperation(wb, { op: 'deleteTable', table: 'Sales' })

			expectErr(result)
			expect(result.error.message).toContain('SalesReps defined name')
			expect(result.error.message).toContain('Sales[Rep]')
			expect(sheet.tables).toHaveLength(1)
			expect(wb.definedNames.get('SalesReps')).toBe('COUNTA(Sales[Rep])')
		}

		for (const scenario of TABLE_FIELD_METADATA_BLOCKERS) {
			const { wb, sheet } = setupSalesRepTable()
			scenario.add(sheet)

			const result = applyOperation(wb, { op: 'deleteTable', table: 'Sales' })

			expectErr(result)
			const diagnostic = `${scenario.label}: ${result.error.message}`
			expect(diagnostic).toContain(`${scenario.sourceRef} ${scenario.sourceKind}`)
			expect(diagnostic).toContain('Sales[Rep]')
			expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
			expect(sheet.tables).toHaveLength(1)
		}

		for (const scenario of LOCAL_TABLE_FIELD_METADATA_BLOCKERS) {
			const { wb, sheet } = setupSalesRepTable()
			scenario.add(sheet)

			const result = applyOperation(wb, { op: 'deleteTable', table: 'Sales' })

			expectErr(result)
			const diagnostic = `${scenario.label}: ${result.error.message}`
			expect(diagnostic).toContain(`${scenario.sourceRef} ${scenario.sourceKind}`)
			expect(diagnostic).toContain('Sales[Rep]')
			expect(result.error.suggestedFix).toContain('Rewrite or remove structured references')
			expect(sheet.tables).toHaveLength(1)
		}
	})

	test('table edit operations reject duplicate imported table names before mutation', () => {
		const operations = [
			{
				op: 'appendRows',
				table: 'Sales',
				rows: [['North', 10]],
			},
			{ op: 'deleteTable', table: 'Sales' },
			{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
			{ op: 'resizeTable', table: 'sales', ref: 'A1:B4' },
			{ op: 'setTableColumn', table: 'Sales', column: 'Amount', newName: 'Revenue' },
			{ op: 'setTableStyle', table: 'Sales', showRowStripes: false },
		] as const

		for (const op of operations) {
			const { wb, sheet1, sheet2 } = setupDuplicateSalesTables()
			const result = applyOperation(wb, op)

			expectErr(result)
			expect(result.error.code).toBe('VALIDATION_ERROR')
			expect(result.error.message).toContain('2 table parts use that name')
			expect(result.error.details).toMatchObject({
				kind: 'duplicate-table-name-operation',
				operation: op.op,
				matches: [
					{ sheetName: 'Sheet1', ref: 'A1:B3', partPath: 'xl/tables/table1.xml' },
					{ sheetName: 'Sheet2', ref: 'D5:E7', partPath: 'xl/tables/table2.xml' },
				],
			})
			expect(sheet1.tables[0]?.name).toBe('Sales')
			expect(sheet1.tables[0]?.ref.end.row).toBe(2)
			expect(sheet2.tables[0]?.name).toBe('Sales')
			expect(sheet2.tables[0]?.ref.end.row).toBe(6)
		}
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
		sheet.tables.push({
			id: createTableId(),
			name: 'Archive',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 7 }, end: { row: 1, col: 7 } },
			columns: [{ name: 'Legacy' }],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'SUM(Sales[Value])', styleId: sid })
		wb.definedNames.set('SalesValues', 'SUM(Sales[Value])')
		sheet.dataValidations.push({
			sqref: 'D2:D3',
			type: 'list',
			formula1: 'SUM(Sales[Value])',
		})
		sheet.conditionalFormats.push({
			sqref: 'E2:E3',
			rules: [
				{
					type: 'expression',
					formulas: ['SUM(Sales[Value])>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Value])' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Value])' }],
						color: { rgb: 'FF00AA00' },
					},
					iconSet: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Value])' }],
					},
				},
			],
		})
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'F2:F3',
			type: 'list',
			formula1: 'SUM(Sales[Value])',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'G2:G3',
			formulas: ['SUM(Sales[Value])>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Value])' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Value])' }] },
		})

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
		expect(sheet.dataValidations[0]?.formula1).toBe('SUM(Revenue[Value])')
		const rule = sheet.conditionalFormats[0]?.rules[0]
		expect(rule?.formulas[0]).toBe('SUM(Revenue[Value])>0')
		expect(rule?.colorScale?.cfvo[0]?.value).toBe('SUM(Revenue[Value])')
		expect(rule?.dataBar?.cfvo[0]?.value).toBe('SUM(Revenue[Value])')
		expect(rule?.iconSet?.cfvo[0]?.value).toBe('SUM(Revenue[Value])')
		expect(sheet.x14DataValidations[0]?.formula1).toBe('SUM(Revenue[Value])')
		expect(sheet.x14ConditionalFormats[0]?.formulas[0]).toBe('SUM(Revenue[Value])>0')
		expect(sheet.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('SUM(Revenue[Value])')
		expect(sheet.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('SUM(Revenue[Value])')

		const revenue = sheet.tables[0]
		if (!revenue) throw new Error('expected renamed table')
		sheet.tables[0] = {
			...revenue,
			autoFilter: {
				ref: 'A1:B2',
				columns: [{ colId: 1, kind: 'filters', values: ['10'] }],
				sortState: { ref: 'A1:B2', conditions: [{ ref: 'B2:B2', descending: true }] },
			},
			sortState: { ref: 'A1:B2', conditions: [{ ref: 'A1:A2' }] },
		}

		const resized = applyOperation(wb, { op: 'resizeTable', table: 'Revenue', ref: 'A1:C2' })
		expectOk(resized)
		expect(resized.value.recalcRequired).toBe(true)
		expect(sheet.tables[0]?.ref.end.col).toBe(2)
		expect(sheet.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Name',
			'Value',
			'Region',
		])
		expect(sheet.tables[0]?.autoFilter).toEqual({
			ref: 'A1:C2',
			columns: [{ colId: 1, kind: 'filters', values: ['10'] }],
			sortState: { ref: 'A1:C2', conditions: [{ ref: 'B2:B2', descending: true }] },
		})
		expect(sheet.tables[0]?.sortState).toEqual({
			ref: 'A1:C2',
			conditions: [{ ref: 'A1:A2' }],
		})

		const deleted = applyOperation(wb, { op: 'deleteTable', table: 'Archive' })
		expectOk(deleted)
		expect(deleted.value.recalcRequired).toBe(true)
		expect(sheet.tables.map((table) => table.name)).toEqual(['Revenue'])
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(10))
	})

	test('renameTable materializes imported shared formulas before table name rewrites', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A2',
			name: 'Sales',
			hasHeaders: true,
		})
		sheet.cells.set(4, 0, {
			value: numberValue(2),
			formula: 'SUM(Sales[Qty])',
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'table-rename-shared',
				isMaster: true,
				masterRef: 'A5',
				ref: 'A5:A6',
			},
		})
		sheet.cells.set(5, 0, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'table-rename-shared',
				isMaster: false,
				masterRef: 'A5',
			},
		})

		const result = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'Revenue',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['A5', 'A6'])
		expect(sheet.cells.get(4, 0)?.formula).toBe('SUM(Revenue[Qty])')
		expect(sheet.cells.get(5, 0)?.formula).toBe('SUM(Revenue[Qty])')
		expect(sheet.cells.get(4, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(5, 0)?.formulaInfo).toBeUndefined()
	})

	test('renameTable preserves dynamic spill metadata while rewriting the anchor formula', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		applyOperation(wb, {
			op: 'createTable',
			sheet: 'Sheet1',
			ref: 'A1:A2',
			name: 'Sales',
			hasHeaders: true,
		})
		sheet.cells.set(0, 3, {
			value: numberValue(2),
			formula: 'FILTER(Sales[Qty],Sales[Qty]>0)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(1, 3, {
			value: numberValue(2),
			formula: null,
			styleId: sid,
			formulaInfo: {
				kind: 'spill',
				anchorRef: 'Sheet1!D1',
				ref: 'D1:D2',
				isAnchor: false,
			},
		})

		const result = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'Revenue',
		})
		expectOk(result)

		expect(result.value.affectedCells).toEqual(['D1', 'D2'])
		expect(sheet.cells.get(0, 3)?.formula).toBe('FILTER(Revenue[Qty],Revenue[Qty]>0)')
		expect(sheet.cells.get(0, 3)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(sheet.cells.get(1, 3)?.formula).toBeNull()
		expect(sheet.cells.get(1, 3)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!D1',
			ref: 'D1:D2',
			isAnchor: false,
		})
		expectCachedFormulaAnalysisMatchesFullRecompute(wb)
	})

	test('renameTable rejects workbook-scoped case-insensitive duplicate table names', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const otherSheet = wb.addSheet('Sheet2')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		otherSheet.tables.push({
			id: createTableId(),
			name: 'FORECAST',
			sheetId: otherSheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Sales[Amount])', styleId: sid })
		wb.definedNames.set('SalesAmount', 'SUM(Sales[Amount])')

		const duplicate = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'forecast',
		})

		expectErr(duplicate)
		expect(duplicate.error.code).toBe('NAME_CONFLICT')
		expect(duplicate.error.suggestedFix).toContain('workbook-unique')
		expect(sheet.tables[0]?.name).toBe('Sales')
		expect(otherSheet.tables[0]?.name).toBe('FORECAST')
		expect(sheet.cells.get(4, 0)?.formula).toBe('SUM(Sales[Amount])')
		expect(wb.definedNames.get('SalesAmount')).toBe('SUM(Sales[Amount])')

		const caseOnly = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'SALES',
		})
		expectOk(caseOnly)
		expect(sheet.tables[0]?.name).toBe('SALES')
	})

	test('renameTable rewrites case-insensitive structured references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(sales[Amount])', styleId: sid })
		wb.definedNames.set('SalesAmount', 'SUM(SALES[Amount])')
		sheet.dataValidations.push({
			sqref: 'C2:C3',
			type: 'custom',
			formula1: 'SUM(Sales[Amount])>0',
		})

		const renamed = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'Revenue',
		})

		expectOk(renamed)
		expect(sheet.tables[0]?.name).toBe('Revenue')
		expect(sheet.cells.get(4, 0)?.formula).toBe('SUM(Revenue[Amount])')
		expect(wb.definedNames.get('SalesAmount')).toBe('SUM(Revenue[Amount])')
		expect(sheet.dataValidations[0]?.formula1).toBe('SUM(Revenue[Amount])>0')
	})

	test('renameTable preserves absent table name attribute state', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		const renamed = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'Revenue',
		})

		expectOk(renamed)
		expect(sheet.tables[0]?.name).toBe('Revenue')
		expect(Object.hasOwn(sheet.tables[0] ?? {}, 'nameAttribute')).toBe(false)
	})

	test('renameTable rejects Excel-invalid names before rewriting structured references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Sales[Amount])', styleId: sid })
		wb.definedNames.set('SalesAmount', 'SUM(Sales[Amount])')

		const result = applyOperation(wb, {
			op: 'renameTable',
			table: 'Sales',
			newName: 'R1C1',
		})

		expectErr(result)
		expect(result.error.code).toBe('VALIDATION_ERROR')
		expect(result.error.message).toContain('cell reference')
		expect(sheet.tables[0]?.name).toBe('Sales')
		expect(sheet.cells.get(4, 0)?.formula).toBe('SUM(Sales[Amount])')
		expect(wb.definedNames.get('SalesAmount')).toBe('SUM(Sales[Amount])')
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
			sheetName: 'sheet1',
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
			relId: 'rIdPivotCache1',
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
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

		expect(result.value.sheetsModified).toEqual(['sheet1'])
		expect(result.value.recalcRequired).toBe(false)
		expect(result.value.warnings?.[0]?.message).toContain('Pivot cache source changed')
		expect(result.value.warnings?.[0]?.details).toMatchObject({
			cacheId: 34,
			refreshOnLoad: true,
			invalid: true,
		})
		expect(wb.pivotCaches[0]).toMatchObject({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			relId: 'rIdPivotCache1',
			sourceSheet: 'RawData',
			sourceRef: 'A1:E20',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			refreshOnLoad: true,
			invalid: true,
			saveData: false,
		})
	})

	test('setPivotCache rejects invalid source ranges and mismatched selectors without mutation', () => {
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
			relId: 'rIdPivotCache1',
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			fields: [],
		})

		const beforeInvalidRange = JSON.stringify(wb.pivotCaches)
		const invalidRange = applyOperation(wb, {
			op: 'setPivotCache',
			cacheId: 34,
			sourceRef: 'D10:A1',
		})
		expectErr(invalidRange)
		expect(invalidRange.error.message).toContain('sourceRef must be an ordered A1 range')
		expect(JSON.stringify(wb.pivotCaches)).toBe(beforeInvalidRange)

		const beforeMismatch = JSON.stringify(wb.pivotCaches)
		const mismatchedSelectors = applyOperation(wb, {
			op: 'setPivotCache',
			cacheId: 99,
			pivotTable: 'PivotTable1',
			sourceRef: 'A1:E20',
		})
		expectErr(mismatchedSelectors)
		expect(mismatchedSelectors.error.message).toContain('cacheId does not match pivotTable')
		expect(JSON.stringify(wb.pivotCaches)).toBe(beforeMismatch)

		const unbound = setup()
		unbound.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		unbound.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			fields: [],
		})
		const beforeUnbound = JSON.stringify(unbound.pivotCaches)
		const unboundPivot = applyOperation(unbound, {
			op: 'setPivotCache',
			pivotTable: 'PivotTable1',
			sourceRef: 'A1:E20',
		})
		expectErr(unboundPivot)
		expect(unboundPivot.error.message).toContain('has no cacheId')
		expect(JSON.stringify(unbound.pivotCaches)).toBe(beforeUnbound)
	})

	test('setConnectionRefresh updates query-table refresh metadata', () => {
		const wb = setup()
		wb.connectionParts.push({
			kind: 'queryTable',
			partPath: 'xl/queryTables/queryTable1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
			relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			sheetName: 'sheet1',
			relationshipCount: 0,
			name: 'SalesQuery',
			connectionId: 1,
			refreshOnLoad: false,
			saveData: true,
		})

		const result = applyOperation(wb, {
			op: 'setConnectionRefresh',
			sheet: 'Sheet1',
			refreshOnLoad: true,
			saveData: false,
			refreshedVersion: 8,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['sheet1'])
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
			sheetName: 'sheet1',
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
			sheet: 'Sheet1',
			fieldIndex: 0,
			itemIndex: 0,
			hidden: null,
			manualFilter: true,
			selectedPageItem: 1,
		})
		expectOk(result)

		expect(result.value.sheetsModified).toEqual(['sheet1'])
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

	test('setPivotFieldItem targets explicit inventory indexes and preserves item metadata', () => {
		const wb = setup()
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{
					index: 2,
					axis: 'axisPage',
					items: [
						{ index: 4, cacheIndex: 9, caption: 'Q1', calculated: true },
						{ index: 7, cacheIndex: 12, missing: true },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 2, item: 4, name: 'Quarter' }],
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
			fieldIndex: 2,
			itemIndex: 4,
			showDetails: true,
			selectedPageItem: 7,
		})
		expectOk(result)

		expect(wb.pivotTables[0]?.fields[0]?.items).toEqual([
			{ index: 4, cacheIndex: 9, caption: 'Q1', calculated: true, showDetails: true },
			{ index: 7, cacheIndex: 12, missing: true },
		])
		expect(wb.pivotTables[0]?.pageFields).toEqual([{ index: 2, item: 7, name: 'Quarter' }])

		const beforeMissingItem = JSON.stringify(wb.pivotTables)
		const missingItem = applyOperation(wb, {
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 2,
			itemIndex: 5,
			hidden: true,
		})
		expectErr(missingItem)
		expect(missingItem.error.message).toContain('Pivot field item 5 was not found')
		expect(JSON.stringify(wb.pivotTables)).toBe(beforeMissingItem)
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
		expect(wb.slicerCaches[0]).toMatchObject({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			sourceName: 'State',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
		})
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

	test('setSlicerCacheItem rejects missing item indexes and duplicate names without mutation', () => {
		const wb = setup()
		wb.slicerCaches.push(
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				name: 'Slicer_State',
				pivotTableNames: [],
				items: [{ index: 0, selected: true }],
			},
			{
				partPath: 'xl/slicerCaches/slicerCache2.xml',
				name: 'Slicer_State',
				pivotTableNames: [],
				items: [{ index: 0, selected: false }],
			},
		)

		const beforeMissingItem = JSON.stringify(wb.slicerCaches)
		const missingItem = applyOperation(wb, {
			op: 'setSlicerCacheItem',
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			item: 2,
			selected: false,
		})
		expectErr(missingItem)
		expect(missingItem.error.message).toContain('Slicer cache item 2 was not found')
		expect(JSON.stringify(wb.slicerCaches)).toBe(beforeMissingItem)

		const beforeDuplicateName = JSON.stringify(wb.slicerCaches)
		const duplicateName = applyOperation(wb, {
			op: 'setSlicerCacheItem',
			slicerCache: 'Slicer_State',
			item: 0,
			selected: false,
		})
		expectErr(duplicateName)
		expect(duplicateName.error.message).toContain('matched 2 caches')
		expect(JSON.stringify(wb.slicerCaches)).toBe(beforeDuplicateName)
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
				filterId: 7,
				filterPivotName: 'PivotTable1',
				singleRangeFilterState: true,
				selection: {
					startDate: '2024-01-01T00:00:00',
					endDate: '2024-03-31T00:00:00',
				},
				bounds: {
					startDate: '2023-01-01T00:00:00',
					endDate: '2024-12-31T00:00:00',
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
		expect(wb.timelineCaches[0]).toMatchObject({
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			name: 'Timeline_Order_Date',
			sourceName: 'Order Date',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
			state: {
				filterId: 7,
				filterPivotName: 'PivotTable1',
				bounds: {
					startDate: '2023-01-01T00:00:00',
					endDate: '2024-12-31T00:00:00',
				},
			},
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
		expect(wb.timelineCaches[0]?.state).toBeUndefined()
	})

	test('setTimelineRange rejects impossible dates and duplicate names without mutation', () => {
		const wb = setup()
		wb.timelineCaches.push(
			{
				partPath: 'xl/timelineCaches/timelineCache1.xml',
				name: 'Timeline_Order_Date',
				pivotTableNames: [],
				state: {
					selection: {
						startDate: '2024-01-01T00:00:00',
						endDate: '2024-03-31T00:00:00',
					},
				},
			},
			{
				partPath: 'xl/timelineCaches/timelineCache2.xml',
				name: 'Timeline_Order_Date',
				pivotTableNames: [],
				state: {
					selection: {
						startDate: '2024-01-01T00:00:00',
						endDate: '2024-03-31T00:00:00',
					},
				},
			},
		)

		const beforeImpossibleDate = JSON.stringify(wb.timelineCaches)
		const impossibleDate = applyOperation(wb, {
			op: 'setTimelineRange',
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			startDate: '2024-02-31T00:00:00',
			endDate: '2024-03-31T00:00:00',
		})
		expectErr(impossibleDate)
		expect(impossibleDate.error.message).toContain('requires ISO-like startDate and endDate')
		expect(JSON.stringify(wb.timelineCaches)).toBe(beforeImpossibleDate)

		const beforeDuplicateName = JSON.stringify(wb.timelineCaches)
		const duplicateName = applyOperation(wb, {
			op: 'setTimelineRange',
			timelineCache: 'Timeline_Order_Date',
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-06-30T00:00:00',
		})
		expectErr(duplicateName)
		expect(duplicateName.error.message).toContain('matched 2 caches')
		expect(JSON.stringify(wb.timelineCaches)).toBe(beforeDuplicateName)
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

	test('rewriteExternalLink rejects ambiguous duplicate target selectors', () => {
		const wb = setup()
		wb.externalReferenceDetails.push(
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				relId: 'rId2',
				linkRelId: 'rIdExt1',
				target: '../sources/shared.xlsx',
				targetMode: 'External',
			},
			{
				partPath: 'xl/externalLinks/externalLink2.xml',
				relId: 'rId3',
				linkRelId: 'rIdExt2',
				target: '../sources/shared.xlsx',
				targetMode: 'External',
			},
		)

		const ambiguous = applyOperation(wb, {
			op: 'rewriteExternalLink',
			target: '../sources/shared.xlsx',
			newTarget: '../sources/reforecast.xlsx',
		})
		expectErr(ambiguous)
		expect(ambiguous.error.message).toContain('matched 2 links')
		expect(wb.externalReferenceDetails.map((entry) => entry.target)).toEqual([
			'../sources/shared.xlsx',
			'../sources/shared.xlsx',
		])

		const disambiguated = applyOperation(wb, {
			op: 'rewriteExternalLink',
			target: '../sources/shared.xlsx',
			linkRelId: 'rIdExt2',
			newTarget: '../sources/reforecast.xlsx',
		})
		expectOk(disambiguated)
		expect(wb.externalReferenceDetails.map((entry) => entry.target)).toEqual([
			'../sources/shared.xlsx',
			'../sources/reforecast.xlsx',
		])
	})

	test('rewriteExternalLink selects missing path bindings by intended externalBook relationship id', () => {
		const wb = setup()
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			externalBookRelId: 'rIdMissing',
			linkBindingStatus: 'missingPathRelationship',
		})

		const result = applyOperation(wb, {
			op: 'rewriteExternalLink',
			linkRelId: 'rIdMissing',
			newTarget: '../sources/repaired.xlsx',
			targetMode: 'External',
		})
		expectOk(result)

		expect(wb.externalReferenceDetails[0]).toMatchObject({
			externalBookRelId: 'rIdMissing',
			linkBindingStatus: 'missingPathRelationship',
			target: '../sources/repaired.xlsx',
			targetMode: 'External',
		})
		expect(wb.externalReferenceDetails[0]?.linkRelId).toBeUndefined()
	})

	test('rewriteExternalLink rejects blank replacement targets without mutation', () => {
		const wb = setup()
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			linkRelId: 'rIdExt',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const result = applyOperation(wb, {
			op: 'rewriteExternalLink',
			linkRelId: 'rIdExt',
			newTarget: ' \t ',
		})
		expectErr(result)
		expect(result.error.message).toContain('newTarget must be a non-empty external workbook target')
		expect(wb.externalReferenceDetails[0]?.target).toBe('../sources/source.xlsx')
	})

	test('rewriteExternalLink preserves symbolic external formula and defined name references', () => {
		const wb = setup()
		const sheet = wb.sheets[0]
		if (!sheet) throw new Error('expected sheet')
		sheet.cells.set(0, 2, {
			value: numberValue(0),
			formula: 'SUM([1]Sheet1!B2:B10)+[Budget.xlsx]FY26!A1',
			styleId: sid,
		})
		wb.definedNames.set('ExternalSource', '[1]Sheet1!A1:D10')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			linkRelId: 'rIdExt',
			linkRelationshipKind: 'externalLinkPath',
			linkBindingStatus: 'externalBookRelId',
			target: '../sources/Budget.xlsx',
			targetMode: 'External',
		})

		const result = applyOperation(wb, {
			op: 'rewriteExternalLink',
			linkRelId: 'rIdExt',
			newTarget: '../sources/Reforecast.xlsx',
		})
		expectOk(result)

		expect(result.value.recalcRequired).toBe(false)
		expect(sheet.cells.get(0, 2)?.formula).toBe('SUM([1]Sheet1!B2:B10)+[Budget.xlsx]FY26!A1')
		expect(wb.definedNames.get('ExternalSource')).toBe('[1]Sheet1!A1:D10')
		expect(wb.externalReferenceDetails[0]?.target).toBe('../sources/Reforecast.xlsx')
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
		expect(sheet.tables[0]?.autoFilter?.sortState?.conditions).toEqual([{ ref: 'B2:B3' }])
		expect(sheet.tables[0]?.sortState?.ref).toBe('A1:B3')
		expect(sheet.tables[0]?.sortState?.conditions).toEqual([{ ref: 'A1:A3' }])
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
