import { describe, expect, test } from 'bun:test'
import type { AutoFilter, CellStyle, Sheet, StyleId } from '@ascend/core'
import { createTableId, createWorkbook, parseRange } from '@ascend/core'
import { dateToSerial } from '@ascend/formulas'
import {
	booleanValue,
	dateValue,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
} from '@ascend/schema'
import { recalculate } from './calc.ts'
import type { CalcContext } from './calc-context.ts'
import { defaultCalcContext } from './calc-context.ts'

const sid = 0 as StyleId

function makeCtx(overrides?: Partial<CalcContext>): CalcContext {
	return { ...defaultCalcContext(), ...overrides }
}

function addSalesTable(sheet: Sheet, name: string, autoFilter?: AutoFilter): void {
	const tableRef = autoFilter ? parseRange(autoFilter.ref) : parseRange('A1:B4')
	sheet.tables.push({
		id: createTableId(),
		name,
		sheetId: sheet.id,
		ref: tableRef,
		columns: [
			{ id: 1, name: 'Region' },
			{ id: 2, name: 'Sales' },
		],
		hasHeaders: true,
		hasTotals: false,
		...(autoFilter ? { autoFilter } : {}),
	})
}

function populateRegionSalesRows(sheet: Sheet): void {
	sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
	sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
	sheet.cells.set(1, 0, { value: stringValue('West'), formula: null, styleId: sid })
	sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
	sheet.cells.set(2, 0, { value: stringValue('East'), formula: null, styleId: sid })
	sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: sid })
	sheet.cells.set(3, 0, { value: stringValue('North'), formula: null, styleId: sid })
	sheet.cells.set(3, 1, { value: numberValue(40), formula: null, styleId: sid })
}

describe('recalculate', () => {
	test('simple SUM formula', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		const cell = sheet.cells.get(2, 0)
		expect(cell?.value).toEqual(numberValue(3))
	})

	test('dirty refs split at the final bang for sheet names containing bangs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet("Q1's Data!")
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'A1+B1', styleId: sid })

		const initial = recalculate(wb, makeCtx())
		expect(initial.errors).toEqual([])
		expect(initial.changed).toEqual(["Q1's Data!!C1"])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(3))

		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		const dirty = recalculate(wb, makeCtx(), { dirtyRefs: ["Q1's Data!!A1"] })
		expect(dirty.errors).toEqual([])
		expect(dirty.changed).toEqual(["Q1's Data!!C1"])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(7))
	})

	test('external workbook references resolve through calculation hook', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '[Budget.xlsx]Inputs!B2', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'SUM([Budget.xlsx]Inputs!A1:A3)', styleId: sid })

		const result = recalculate(
			wb,
			makeCtx({
				externalReferences: {
					resolveCell: ({ workbook, sheet: sheetName, row, col }) => {
						if (workbook !== 'Budget.xlsx' || sheetName !== 'Inputs') return undefined
						if (row === 1 && col === 1) return numberValue(42)
						if (col === 0) return numberValue(row + 1)
						return undefined
					},
				},
			}),
		)

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(42))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(6))
	})

	test('path-qualified external workbook references resolve through calculation hook', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: "'C:/tmp/[Budget.xlsx]Inputs'!B2",
			styleId: sid,
		})
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: "SUM('C:/tmp/[Budget.xlsx]Inputs'!A1:A3)",
			styleId: sid,
		})

		const result = recalculate(
			wb,
			makeCtx({
				externalReferences: {
					resolveCell: ({ workbook, sheet: sheetName, row, col }) => {
						if (workbook !== 'C:/tmp/Budget.xlsx' || sheetName !== 'Inputs') return undefined
						if (row === 1 && col === 1) return numberValue(42)
						if (col === 0) return numberValue(row + 1)
						return undefined
					},
				},
			}),
		)

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(42))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(6))
	})

	test('external workbook references remain #REF! without a resolver', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '[Budget.xlsx]Inputs!B2', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('SUM supports INDEX as a dynamic range endpoint', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:INDEX(A:A,5))', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(15))
	})

	test('INDEX selects from multi-area references with area_num', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(40), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'INDEX((A1:A2,C1:C2),2,1,2)', styleId: sid })
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'INDEX((A1:A2,C1:C2),2,1)', styleId: sid })
		sheet.cells.set(2, 3, { value: EMPTY, formula: 'INDEX((A1:A2,C1:C2),1,1,3)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(40))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 3)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('INDEX returns selected multi-area references to aggregate functions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(40), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'SUM(INDEX((A1:A2,C1:C2),0,1,2))',
			styleId: sid,
		})
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'SUM(INDEX((A1:A2,C1:C2),0,1))', styleId: sid })
		sheet.cells.set(2, 3, {
			value: EMPTY,
			formula: 'SUM(INDEX((A1:A2,C1:C2),0,1,3))',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(70))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(2, 3)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('SUM supports OFFSET as a dynamic range endpoint', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 4; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 2), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:OFFSET(A1,3,0))', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(14))
	})

	test('reference functions can form both dynamic range endpoints', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'SUM(INDEX(A:A,2):INDEX(A:A,4))',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(9))
	})

	test('full recalc fast-paths simple shared relative binary formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 4; row++) {
			sheet.cells.set(row, 1, { value: numberValue(row + 1), formula: null, styleId: sid })
			sheet.cells.set(row, 2, { value: numberValue((row + 1) * 10), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1+C1',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		for (let row = 1; row < 4; row++) {
			sheet.cells.set(row, 0, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3', 'Sheet1!A4'])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(11))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(44))
		expect(sheet.cells.get(3, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A1',
		})
	})

	test('shared relative binary fast path handles division by zero like arithmetic eval', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1/C1',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('growing range aggregate optimization evaluates COUNT, AVERAGE, MIN, and MAX', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const inputs = [numberValue(4), stringValue('ignored'), numberValue(2), numberValue(8)]
		for (let row = 0; row < inputs.length; row++) {
			sheet.cells.set(row, 0, { value: inputs[row] ?? EMPTY, formula: null, styleId: sid })
		}
		const cases = [
			{ fn: 'COUNT', col: 1, expected: [1, 1, 2, 3] },
			{ fn: 'AVERAGE', col: 2, expected: [4, 4, 3, 14 / 3] },
			{ fn: 'MIN', col: 3, expected: [4, 4, 2, 2] },
			{ fn: 'MAX', col: 4, expected: [4, 4, 4, 8] },
		] as const
		for (const c of cases) {
			for (let row = 0; row < inputs.length; row++) {
				sheet.cells.set(row, c.col, {
					value: EMPTY,
					formula: `${c.fn}(A$1:A${row + 1})`,
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (const c of cases) {
			for (let row = 0; row < inputs.length; row++) {
				expect(sheet.cells.get(row, c.col)?.value).toEqual(numberValue(c.expected[row] ?? 0))
			}
		}
	})

	test('dirty recalc updates a tail prefix aggregate from the indexed previous prefix', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: sid })
			sheet.cells.set(row, 1, {
				value: EMPTY,
				formula: `SUM(A$1:A${row + 1})`,
				styleId: sid,
			})
		}

		expect(recalculate(wb, makeCtx()).errors).toEqual([])
		sheet.cells.set(4, 0, { value: numberValue(100), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!A5'] })

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Sheet1!B5'])
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(110))
	})

	test('dirty recalc updates tail AVERAGE from cached prefix state without counting blanks', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(4), formula: null, styleId: sid })
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 1, {
				value: EMPTY,
				formula: `AVERAGE(A$1:A${row + 1})`,
				styleId: sid,
			})
		}

		expect(recalculate(wb, makeCtx()).errors).toEqual([])
		sheet.cells.set(2, 0, { value: numberValue(8), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!A3'] })

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Sheet1!B3'])
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(5))
	})

	test('dirty recalc delta-updates head edits across cached prefix SUM formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: sid })
			sheet.cells.set(row, 1, {
				value: EMPTY,
				formula: `SUM(A$1:A${row + 1})`,
				styleId: sid,
			})
		}

		expect(recalculate(wb, makeCtx()).errors).toEqual([])
		sheet.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!A1'] })

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual([
			'Sheet1!B1',
			'Sheet1!B2',
			'Sheet1!B3',
			'Sheet1!B4',
			'Sheet1!B5',
		])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(100))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(114))
	})

	test('chain of dependent formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1*2', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+5', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(25))
	})

	test('full recalc fast-paths previous-row additive formula chains', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+3', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+3', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'A3-1', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Sheet1!A2', 'Sheet1!A3', 'Sheet1!A4'])
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(8))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(3, 0)?.formula).toBe('A3-1')
	})

	test('recalculation after value change reports changed cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+10', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(11))

		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(15))
		expect(result.changed.length).toBeGreaterThan(0)
	})

	test('dirty refs limit recalculation scope to affected dependency subgraph', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'C1+1', styleId: sid })

		recalculate(wb, makeCtx())
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(11))
		expect(result.changed).toContain('Sheet1!B1')
		expect(result.changed).not.toContain('Sheet1!D1')
	})

	test('dirty recalc only reports parse errors inside the affected subgraph', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SUM((', styleId: sid })

		recalculate(wb, makeCtx())
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
		expect(result.errors).toEqual([])
	})

	test('dirty recalc invalidates XLOOKUP exact index cache after lookup range edits', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('c'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'XLOOKUP(D1,A1:A3,B1:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(20))

		sheet.cells.set(1, 0, { value: stringValue('x'), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A2'] })

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 4)?.value).toEqual(errorValue('#N/A'))
	})

	test('dirty recalc invalidates XLOOKUP return vector cache after return range edits', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('c'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'XLOOKUP(D1,A1:A3,B1:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(20))

		sheet.cells.set(1, 1, { value: numberValue(99), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!B2'] })

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(99))
	})

	test('dirty INDEX/MATCH return cache stays correct after lookup key edits', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('c'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'INDEX(B1:B3,MATCH(D1,A1:A3,0))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(20))

		sheet.cells.set(0, 3, { value: stringValue('c'), formula: null, styleId: sid })
		const keyResult = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!D1'] })
		expect(keyResult.errors).toEqual([])
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(30))

		sheet.cells.set(1, 1, { value: numberValue(99), formula: null, styleId: sid })
		const oldReturnResult = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!B2'] })
		expect(oldReturnResult.errors).toEqual([])
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(30))

		sheet.cells.set(2, 1, { value: numberValue(77), formula: null, styleId: sid })
		const newReturnResult = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!B3'] })
		expect(newReturnResult.errors).toEqual([])
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(77))
	})

	test('Salsa backdating: when formula output unchanged, downstream dependents are not re-evaluated', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'IF(A1>0,"yes","no")', styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'B1&"!"', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('yes'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('yes!'))

		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('yes'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('yes!'))
		expect(result.changed).toEqual([])
	})

	test('multi-sheet reference', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Data')
		const s2 = wb.addSheet('Summary')
		s1.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: EMPTY, formula: 'Data!A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(s2.cells.get(0, 0)?.value).toEqual(numberValue(42))
	})

	test('sheet-scoped defined name shadows workbook-scoped name', () => {
		const wb = createWorkbook()
		const sheet1 = wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		wb.definedNames.set('Rate', '0.1')
		wb.definedNames.set('Rate', '0.2', { kind: 'sheet', sheetId: sheet2.id })
		sheet1.cells.set(0, 0, { value: EMPTY, formula: 'Rate*100', styleId: sid })
		sheet2.cells.set(0, 0, { value: EMPTY, formula: 'Rate*100', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet1.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(sheet2.cells.get(0, 0)?.value).toEqual(numberValue(20))
	})

	test('qualified local defined names can be referenced from another sheet', () => {
		const wb = createWorkbook()
		const sheet1 = wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		const calc = wb.addSheet('Calc')
		wb.definedNames.set('Budget', '10', { kind: 'sheet', sheetId: sheet1.id })
		wb.definedNames.set('Budget', '20', { kind: 'sheet', sheetId: sheet2.id })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'Sheet1!Budget+Sheet2!Budget', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(calc.cells.get(0, 0)?.value).toEqual(numberValue(30))
	})

	test('qualified local defined name ranges on other sheets do not self-reference by coordinate', () => {
		const wb = createWorkbook()
		const sheet1 = wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		sheet1.cells.set(0, 0, { value: numberValue(7), formula: null, styleId: sid })
		wb.definedNames.set('LocalCell', 'A1', { kind: 'sheet', sheetId: sheet1.id })
		sheet2.cells.set(0, 0, { value: EMPTY, formula: 'Sheet1!LocalCell', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet2.cells.get(0, 0)?.value).toEqual(numberValue(7))
	})

	test('workbook-index-qualified defined names resolve workbook scope', () => {
		const wb = createWorkbook()
		const data = wb.addSheet('Data')
		const calc = wb.addSheet('Calc')
		data.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		data.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		wb.definedNames.set('col1_', 'Data!A1:A2')
		wb.definedNames.set('col1_', 'Data!A2:A2', { kind: 'sheet', sheetId: calc.id })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'SUM([0]!col1_)', styleId: sid })
		calc.cells.set(0, 1, { value: EMPTY, formula: 'SUM([0]!missing_name)', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(calc.cells.get(0, 0)?.value).toEqual(numberValue(5))
		expect(calc.cells.get(0, 1)?.value).toEqual(errorValue('#NAME?'))
	})

	test('structured references can sum a table column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Player'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Mina'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Noah'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(12), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Scores',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Player' }, { name: 'Score' }],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Scores[Score])', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(22))
	})

	test('dirty recalc updates same-sheet defined name backed by a structured reference', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		populateRegionSalesRows(sheet)
		addSalesTable(sheet, 'Sales')
		wb.definedNames.set('SalesValues', 'Sales[Sales]')
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SUM(SalesValues)', styleId: sid })

		expect(recalculate(wb, makeCtx()).errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(90))

		sheet.cells.set(2, 1, { value: numberValue(300), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyRefs: ['Sheet1!B3'] })

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Sheet1!D1'])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(360))
	})

	test('dirty recalc updates cross-sheet defined name backed by a structured reference', () => {
		const wb = createWorkbook()
		const data = wb.addSheet('Data')
		const summary = wb.addSheet('Summary')
		populateRegionSalesRows(data)
		addSalesTable(data, 'Sales')
		wb.definedNames.set('SalesValues', 'Sales[Sales]')
		summary.cells.set(0, 0, { value: EMPTY, formula: 'SUM(SalesValues)', styleId: sid })

		expect(recalculate(wb, makeCtx()).errors).toEqual([])
		expect(summary.cells.get(0, 0)?.value).toEqual(numberValue(90))

		data.cells.set(2, 1, { value: numberValue(300), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyRefs: ['Data!B3'] })

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Summary!A1'])
		expect(summary.cells.get(0, 0)?.value).toEqual(numberValue(360))
	})

	test('whole-column references can aggregate used cells in a column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'SUM(A:A)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(10))
	})

	test('dirty recalc updates formulas using whole-column references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'SUM(A:A)', styleId: sid })

		recalculate(wb, makeCtx())
		sheet.cells.set(2, 0, { value: numberValue(10), formula: null, styleId: sid })

		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A3'] })
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(14))
	})

	test('whole-row references can aggregate used cells in a row', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: EMPTY, formula: 'SUM(2:2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(10))
	})

	test('ROWS and COLUMNS treat whole-dimension refs as full-sheet ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'ROWS(A:A)', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'COLUMNS(1:1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1_048_576))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(16_384))
	})

	test('bare contiguous ranges spill under array semantics', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1:A3', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(3))
	})

	test('binary arithmetic spills range operands elementwise', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'A1:A3*2',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1 },
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(6))
	})

	test('binary arithmetic spills array constants elementwise', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '{1,2,3}+{10,20,30}', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(11))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(22))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(33))
	})

	test('binary comparison spills boolean masks', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'A1:A3>1',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1 },
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(booleanValue(false))
		expect(sheet.cells.get(1, 1)?.value).toEqual(booleanValue(true))
		expect(sheet.cells.get(2, 1)?.value).toEqual(booleanValue(true))
	})

	test('binary comparisons use Excel type ordering without cross-type numeric coercion', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '"1.1"=1.1', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: '1<"1"', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'FALSE=0', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'FALSE>"TRUE"', styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: '"TRUE"<>TRUE', styleId: sid })
		sheet.cells.set(5, 0, { value: EMPTY, formula: '""=B1', styleId: sid })
		sheet.cells.set(6, 0, { value: EMPTY, formula: 'FALSE=B1', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(booleanValue(false))
		expect(sheet.cells.get(1, 0)?.value).toEqual(booleanValue(true))
		expect(sheet.cells.get(2, 0)?.value).toEqual(booleanValue(false))
		expect(sheet.cells.get(3, 0)?.value).toEqual(booleanValue(true))
		expect(sheet.cells.get(4, 0)?.value).toEqual(booleanValue(true))
		expect(sheet.cells.get(5, 0)?.value).toEqual(booleanValue(true))
		expect(sheet.cells.get(6, 0)?.value).toEqual(booleanValue(true))
	})

	test('IF supports array conditions and array branch values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('x'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'IF(ISNUMBER(A1:A2),A1:A2,0)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(0))
	})

	test('NOT spills boolean masks for range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: booleanValue(true), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: booleanValue(false), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'IF(NOT(A1:A4),"no","yes")', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('yes'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('no'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('no'))
		expect(sheet.cells.get(3, 1)?.value).toEqual(stringValue('yes'))
	})

	test('IFERROR and IFNA replace matching errors inside arrays', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: errorValue('#N/A'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'IFERROR(A1:A3,0)', styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'IFNA(A1:A3,"missing")', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(errorValue('#DIV/0!'))
		expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('missing'))
	})

	test('error predicates map over arrays inside IF conditions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: errorValue('#N/A'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'IF(ISERROR(A1:A3),0,A1:A3)', styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'IF(ISNA(A1:A3),"missing",A1:A3)',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'IF(ISERR(A1:A3),"non-na",A1:A3)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(errorValue('#DIV/0!'))
		expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('missing'))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('non-na'))
		expect(sheet.cells.get(2, 3)?.value).toEqual(errorValue('#N/A'))
	})

	test('common IS predicates spill boolean masks for range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('north'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: booleanValue(true), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'IF(ISBLANK(A1:A5),"blank","filled")',
			styleId: sid,
		})
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'IF(ISTEXT(A1:A5),A1:A5,"")',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'IF(ISLOGICAL(A1:A5),"logical","other")',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'IF(ISNONTEXT(A1:A5),"not-text","text")',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'IF(ISEVEN(A1:A5),"even","odd")',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'IF(ISODD(A1:A5),"odd","even")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('blank'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('filled'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue(''))
		expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('north'))
		expect(sheet.cells.get(2, 3)?.value).toEqual(stringValue('logical'))
		expect(sheet.cells.get(3, 3)?.value).toEqual(stringValue('other'))
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('not-text'))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue('text'))
		expect(sheet.cells.get(0, 5)?.value).toEqual(stringValue('even'))
		expect(sheet.cells.get(1, 5)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(2, 5)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(3, 5)?.value).toEqual(stringValue('even'))
		expect(sheet.cells.get(4, 5)?.value).toEqual(stringValue('odd'))
		expect(sheet.cells.get(3, 6)?.value).toEqual(stringValue('even'))
		expect(sheet.cells.get(4, 6)?.value).toEqual(stringValue('odd'))
	})

	test('ISFORMULA spills formula masks for range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '1+1', styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'NA()', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'IF(ISFORMULA(A1:A4),"formula","value")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('formula'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('value'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('formula'))
		expect(sheet.cells.get(3, 1)?.value).toEqual(stringValue('value'))
	})

	test('ERROR.TYPE spills error codes for range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: errorValue('#N/A'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: errorValue('#VALUE!'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'IFERROR(ERROR.TYPE(A1:A4),0)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(3))
	})

	test('common text scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue(' north '), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('East'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('No'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'IF(LEN(A1:A3)>3,UPPER(TRIM(A1:A3)),LOWER(A1:A3))',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('NORTH'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('EAST'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('no'))
	})

	test('text extraction functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('AX-100'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('BY-205'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('CZ-330'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'LEFT(A1:A3,2)&MID(A1:A3,4,1)&RIGHT(A1:A3,1)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('AX10'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('BY25'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('CZ30'))
	})

	test('FIND and SEARCH spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('East Region'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('west'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('southeast'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'IF(ISNUMBER(SEARCH("east",A1:A3)),"match","miss")',
			styleId: sid,
		})
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'IFERROR(FIND("E",A1:A3),0)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('match'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('miss'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('match'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(0))
	})

	test('SUBSTITUTE and REPLACE spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Q1-2024-East'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Q2-2024-West'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Q3-2024-North'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'SUBSTITUTE(A1:A3,"-","/")&""',
			styleId: sid,
		})
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SUBSTITUTE(A1:A3,"-","/",2)&""',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'REPLACE(A1:A3,4,4,"FY25")&""',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('Q1/2024/East'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('Q2/2024/West'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('Q3/2024/North'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('Q1-2024/East'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('Q2-2024/West'))
		expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('Q3-2024/North'))
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Q1-FY25-East'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Q2-FY25-West'))
		expect(sheet.cells.get(2, 3)?.value).toEqual(stringValue('Q3-FY25-North'))
	})

	test('EXACT, PROPER, and VALUE spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('sku-01'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('SKU-02'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('sku-03'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('sku-01'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('sku-02'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('SKU-03'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: stringValue('$1,200'), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: stringValue('75%'), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: stringValue('40'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: stringValue('north region'), formula: null, styleId: sid })
		sheet.cells.set(1, 4, { value: stringValue('east region'), formula: null, styleId: sid })
		sheet.cells.set(2, 4, { value: stringValue('south region'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'IF(EXACT(A1:A3,B1:B3),"same","diff")',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'IF(VALUE(D1:D3)>100,PROPER(E1:E3),"small")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('same'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('diff'))
		expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('diff'))
		expect(sheet.cells.get(0, 5)?.value).toEqual(stringValue('North Region'))
		expect(sheet.cells.get(1, 5)?.value).toEqual(stringValue('small'))
		expect(sheet.cells.get(2, 5)?.value).toEqual(stringValue('small'))
	})

	test('character code and repeat text functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(65), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(66), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(67), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('A\u0001B'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('\u0002CD'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('EF\u0003'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('x'), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: stringValue('ab'), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: stringValue('Q'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: stringValue('A'), formula: null, styleId: sid })
		sheet.cells.set(1, 4, { value: stringValue('BC'), formula: null, styleId: sid })
		sheet.cells.set(2, 4, { value: stringValue('DEF'), formula: null, styleId: sid })
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'CHAR(A1:A3)&UNICHAR(A1:A3)',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'CODE(E1:E3)+UNICODE(E1:E3)',
			styleId: sid,
		})
		sheet.cells.set(0, 7, {
			value: EMPTY,
			formula: 'CLEAN(B1:B3)&""',
			styleId: sid,
		})
		sheet.cells.set(0, 8, {
			value: EMPTY,
			formula: 'REPT(C1:C3,D1:D3)&""',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 5)?.value).toEqual(stringValue('AA'))
		expect(sheet.cells.get(1, 5)?.value).toEqual(stringValue('BB'))
		expect(sheet.cells.get(2, 5)?.value).toEqual(stringValue('CC'))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(130))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(132))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(136))
		expect(sheet.cells.get(0, 7)?.value).toEqual(stringValue('AB'))
		expect(sheet.cells.get(1, 7)?.value).toEqual(stringValue('CD'))
		expect(sheet.cells.get(2, 7)?.value).toEqual(stringValue('EF'))
		expect(sheet.cells.get(0, 8)?.value).toEqual(stringValue('x'))
		expect(sheet.cells.get(1, 8)?.value).toEqual(stringValue('abab'))
		expect(sheet.cells.get(2, 8)?.value).toEqual(stringValue('QQQ'))
	})

	test('numeric text formatting functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0.125), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(0.5), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(1.25), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(1234.56), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(-1234.56), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(0.125), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'TEXT(A1:A3,"0.0%")&""',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'DOLLAR(C1:C3,D1:D3)&""',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'FIXED(C1:C3,D1:D3,FALSE)&""',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('12.5%'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('50.0%'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('125.0%'))
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('$1,234.6'))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue('($1,235)'))
		expect(sheet.cells.get(2, 4)?.value).toEqual(stringValue('$0.13'))
		expect(sheet.cells.get(0, 5)?.value).toEqual(stringValue('1,234.6'))
		expect(sheet.cells.get(1, 5)?.value).toEqual(stringValue('-1,235'))
		expect(sheet.cells.get(2, 5)?.value).toEqual(stringValue('0.13'))
	})

	test('common math scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(-1.25), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2.55), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(-4.41), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(-5), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(7), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'ROUND(ABS(A1:A3),1)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'MOD(D1:D3,2)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'POWER(B1:B3,2)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'IF(SIGN(A1:A3)<0,INT(A1:A3),ROUNDUP(A1:A3,0))',
			styleId: sid,
		})
		sheet.cells.set(0, 7, {
			value: EMPTY,
			formula: 'ROUNDDOWN(A1:A3,0)+TRUNC(A1:A3)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1.3))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2.6))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(4.4))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(16))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(-2))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(-5))
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(-2))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(2, 7)?.value).toEqual(numberValue(-8))
	})

	test('math conversion scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(16), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(255), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(8), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(16), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: stringValue('101'), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: stringValue('20'), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: stringValue('FF'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: stringValue('V'), formula: null, styleId: sid })
		sheet.cells.set(1, 4, { value: stringValue('XVI'), formula: null, styleId: sid })
		sheet.cells.set(2, 4, { value: stringValue('MCMXCIX'), formula: null, styleId: sid })
		sheet.cells.set(0, 5, { value: EMPTY, formula: 'QUOTIENT(A1:A3,C1:C3)+0', styleId: sid })
		sheet.cells.set(0, 6, { value: EMPTY, formula: 'BASE(A1:A3,B1:B3)&""', styleId: sid })
		sheet.cells.set(0, 7, { value: EMPTY, formula: 'DECIMAL(D1:D3,B1:B3)+0', styleId: sid })
		sheet.cells.set(0, 8, { value: EMPTY, formula: 'ROMAN(A1:A3)&""', styleId: sid })
		sheet.cells.set(0, 9, { value: EMPTY, formula: 'ARABIC(E1:E3)+0', styleId: sid })
		sheet.cells.set(0, 10, { value: EMPTY, formula: 'SQRTPI(A1:A3)+0', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(8))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(63))
		expect(sheet.cells.get(0, 6)?.value).toEqual(stringValue('101'))
		expect(sheet.cells.get(1, 6)?.value).toEqual(stringValue('20'))
		expect(sheet.cells.get(2, 6)?.value).toEqual(stringValue('FF'))
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(16))
		expect(sheet.cells.get(2, 7)?.value).toEqual(numberValue(255))
		expect(sheet.cells.get(0, 8)?.value).toEqual(stringValue('V'))
		expect(sheet.cells.get(1, 8)?.value).toEqual(stringValue('XVI'))
		expect(sheet.cells.get(2, 8)?.value).toEqual(stringValue('CCLV'))
		expect(sheet.cells.get(0, 9)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(1, 9)?.value).toEqual(numberValue(16))
		expect(sheet.cells.get(2, 9)?.value).toEqual(numberValue(1999))
		expect(sheet.cells.get(0, 10)?.value).toEqual(numberValue(Math.sqrt(5 * Math.PI)))
		expect(sheet.cells.get(1, 10)?.value).toEqual(numberValue(Math.sqrt(16 * Math.PI)))
		expect(sheet.cells.get(2, 10)?.value).toEqual(numberValue(Math.sqrt(255 * Math.PI)))
	})

	test('engineering bitwise scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(13), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(16), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(25), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(-1), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(3, 2, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(14), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(3, 3, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'BITAND(A1:A4,B1:B4)+0', styleId: sid })
		sheet.cells.set(0, 5, { value: EMPTY, formula: 'BITOR(A1:A4,B1:B4)+0', styleId: sid })
		sheet.cells.set(0, 6, { value: EMPTY, formula: 'BITXOR(A1:A4,B1:B4)+0', styleId: sid })
		sheet.cells.set(0, 7, { value: EMPTY, formula: 'BITLSHIFT(A1:A4,C1:C4)+0', styleId: sid })
		sheet.cells.set(0, 8, { value: EMPTY, formula: 'BITRSHIFT(A1:A4,C1:C4)+0', styleId: sid })
		sheet.cells.set(0, 9, { value: EMPTY, formula: 'DELTA(A1:A4,D1:D4)+0', styleId: sid })
		sheet.cells.set(0, 10, { value: EMPTY, formula: 'GESTEP(A1:A4,D1:D4)+0', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(3, 4)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(29))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(3, 5)?.value).toEqual(numberValue(17))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(3, 6)?.value).toEqual(numberValue(17))
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(2, 7)?.value).toEqual(numberValue(32))
		expect(sheet.cells.get(3, 7)?.value).toEqual(numberValue(16))
		expect(sheet.cells.get(0, 8)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(1, 8)?.value).toEqual(numberValue(26))
		expect(sheet.cells.get(2, 8)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(3, 8)?.value).toEqual(numberValue(16))
		expect(sheet.cells.get(0, 9)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 9)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 9)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(3, 9)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(0, 10)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 10)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(2, 10)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(3, 10)?.value).toEqual(numberValue(0))
	})

	test('engineering radix conversion functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const decimals = [9, 64, 255]
		const places = [4, 7, 8]
		const binaries = ['1001', '1000000', '11111111']
		const hexes = ['9', '40', 'FF']
		const octals = ['11', '100', '377']
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(decimals[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(places[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: stringValue(binaries[row] as string),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 3, {
				value: stringValue(hexes[row] as string),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 4, {
				value: stringValue(octals[row] as string),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 5, { value: EMPTY, formula: 'BIN2DEC(C1:C3)+0', styleId: sid })
		sheet.cells.set(0, 6, { value: EMPTY, formula: 'DEC2BIN(A1:A3,B1:B3)&""', styleId: sid })
		sheet.cells.set(0, 7, { value: EMPTY, formula: 'HEX2DEC(D1:D3)+0', styleId: sid })
		sheet.cells.set(0, 8, { value: EMPTY, formula: 'DEC2HEX(A1:A3,2)&""', styleId: sid })
		sheet.cells.set(0, 9, { value: EMPTY, formula: 'OCT2DEC(E1:E3)+0', styleId: sid })
		sheet.cells.set(0, 10, { value: EMPTY, formula: 'DEC2OCT(A1:A3)&""', styleId: sid })
		sheet.cells.set(0, 11, { value: EMPTY, formula: 'BIN2HEX(C1:C3)&""', styleId: sid })
		sheet.cells.set(0, 12, { value: EMPTY, formula: 'BIN2OCT(C1:C3)&""', styleId: sid })
		sheet.cells.set(0, 13, { value: EMPTY, formula: 'HEX2BIN(D1:D3,B1:B3)&""', styleId: sid })
		sheet.cells.set(0, 14, { value: EMPTY, formula: 'HEX2OCT(D1:D3)&""', styleId: sid })
		sheet.cells.set(0, 15, { value: EMPTY, formula: 'OCT2BIN(E1:E3,B1:B3)&""', styleId: sid })
		sheet.cells.set(0, 16, { value: EMPTY, formula: 'OCT2HEX(E1:E3)&""', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			expect(sheet.cells.get(row, 5)?.value).toEqual(numberValue(decimals[row] as number))
			expect(sheet.cells.get(row, 6)?.value).toEqual(stringValue(binaries[row] as string))
			expect(sheet.cells.get(row, 7)?.value).toEqual(numberValue(decimals[row] as number))
			expect(sheet.cells.get(row, 9)?.value).toEqual(numberValue(decimals[row] as number))
			expect(sheet.cells.get(row, 10)?.value).toEqual(stringValue(octals[row] as string))
			expect(sheet.cells.get(row, 11)?.value).toEqual(stringValue(hexes[row] as string))
			expect(sheet.cells.get(row, 12)?.value).toEqual(stringValue(octals[row] as string))
			expect(sheet.cells.get(row, 13)?.value).toEqual(stringValue(binaries[row] as string))
			expect(sheet.cells.get(row, 14)?.value).toEqual(stringValue(octals[row] as string))
			expect(sheet.cells.get(row, 15)?.value).toEqual(stringValue(binaries[row] as string))
			expect(sheet.cells.get(row, 16)?.value).toEqual(stringValue(hexes[row] as string))
		}
		expect(sheet.cells.get(0, 8)?.value).toEqual(stringValue('09'))
		expect(sheet.cells.get(1, 8)?.value).toEqual(stringValue('40'))
		expect(sheet.cells.get(2, 8)?.value).toEqual(stringValue('FF'))
	})

	test('engineering special and complex scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const xs = [0.5, 1, 1.5]
		const orders = [0, 1, 2]
		const positiveXs = [1, 2, 3]
		const imaginaries = [4, -1, 2]
		const complexValues = ['3+4i', '2-1i', '1+2i']
		const divisorValues = ['1+i', '2+3i', '4-2i']
		for (let row = 0; row < 3; row++) {
			for (const [col, value] of [
				[0, xs[row]],
				[1, orders[row]],
				[2, positiveXs[row]],
				[3, imaginaries[row]],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(value as number),
					formula: null,
					styleId: sid,
				})
			}
			sheet.cells.set(row, 4, {
				value: stringValue(complexValues[row] as string),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 5, {
				value: stringValue(divisorValues[row] as string),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = [
			'CONVERT(A1:A3,"m","ft")+0',
			'ERF(A1:A3)+0',
			'ERF.PRECISE(A1:A3)+0',
			'ERFC(A1:A3)+0',
			'ERFC.PRECISE(A1:A3)+0',
			'BESSELI(A1:A3,B1:B3)+0',
			'BESSELJ(A1:A3,B1:B3)+0',
			'BESSELK(C1:C3,B1:B3)+0',
			'BESSELY(C1:C3,B1:B3)+0',
			'COMPLEX(A1:A3,D1:D3)&""',
			'IMREAL(E1:E3)+0',
			'IMAGINARY(E1:E3)+0',
			'IMABS(E1:E3)+0',
			'IMARGUMENT(E1:E3)+0',
			'IMCONJUGATE(E1:E3)&""',
			'IMSUB(E1:E3,F1:F3)&""',
			'IMDIV(E1:E3,F1:F3)&""',
			'IMPOWER(E1:E3,B1:B3)&""',
			'IMSQRT(E1:E3)&""',
			'IMEXP(E1:E3)&""',
			'IMLN(E1:E3)&""',
			'IMSIN(E1:E3)&""',
			'IMCOS(E1:E3)&""',
			'IMLOG10(E1:E3)&""',
			'IMLOG2(E1:E3)&""',
			'IMTAN(E1:E3)&""',
			'IMSINH(E1:E3)&""',
			'IMCOSH(E1:E3)&""',
			'IMSEC(E1:E3)&""',
			'IMCSC(E1:E3)&""',
			'IMCOT(E1:E3)&""',
			'IMSECH(E1:E3)&""',
			'IMCSCH(E1:E3)&""',
		]
		const scalarFormulas = [
			'CONVERT(A{r},"m","ft")+0',
			'ERF(A{r})+0',
			'ERF.PRECISE(A{r})+0',
			'ERFC(A{r})+0',
			'ERFC.PRECISE(A{r})+0',
			'BESSELI(A{r},B{r})+0',
			'BESSELJ(A{r},B{r})+0',
			'BESSELK(C{r},B{r})+0',
			'BESSELY(C{r},B{r})+0',
			'COMPLEX(A{r},D{r})&""',
			'IMREAL(E{r})+0',
			'IMAGINARY(E{r})+0',
			'IMABS(E{r})+0',
			'IMARGUMENT(E{r})+0',
			'IMCONJUGATE(E{r})&""',
			'IMSUB(E{r},F{r})&""',
			'IMDIV(E{r},F{r})&""',
			'IMPOWER(E{r},B{r})&""',
			'IMSQRT(E{r})&""',
			'IMEXP(E{r})&""',
			'IMLN(E{r})&""',
			'IMSIN(E{r})&""',
			'IMCOS(E{r})&""',
			'IMLOG10(E{r})&""',
			'IMLOG2(E{r})&""',
			'IMTAN(E{r})&""',
			'IMSINH(E{r})&""',
			'IMCOSH(E{r})&""',
			'IMSEC(E{r})&""',
			'IMCSC(E{r})&""',
			'IMCOT(E{r})&""',
			'IMSECH(E{r})&""',
			'IMCSCH(E{r})&""',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 6 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 45 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 6 + col)?.value).toEqual(sheet.cells.get(row, 45 + col)?.value)
			}
		}
	})

	test('engineering special and complex scalar functions implicitly intersect range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: sid })
			sheet.cells.set(row, 1, { value: numberValue(row), formula: null, styleId: sid })
			sheet.cells.set(row, 2, {
				value: stringValue(row === 1 ? '3+4i' : '1+i'),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 3, {
				value: stringValue(row === 1 ? '1+i' : '2+i'),
				formula: null,
				styleId: sid,
			})
		}
		const rangeFormulas = [
			'CONVERT(A1:A3,"m","ft")',
			'ERF(A1:A3)',
			'BESSELJ(A1:A3,B1:B3)',
			'COMPLEX(A1:A3,B1:B3)',
			'IMABS(C1:C3)',
			'IMSUB(C1:C3,D1:D3)',
		]
		const scalarFormulas = [
			'CONVERT(A2,"m","ft")',
			'ERF(A2)',
			'BESSELJ(A2,B2)',
			'COMPLEX(A2,B2)',
			'IMABS(C2)',
			'IMSUB(C2,D2)',
		]
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 4 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 12 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 4 + col)?.value).toEqual(sheet.cells.get(1, 12 + col)?.value)
		}
	})

	test('rounding scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2.5), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(-2.5), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6.3), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(-6.7), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(-1), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: numberValue(-1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'CEILING(A1:A4,1)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'FLOOR(A1:A4,1)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'CEILING.MATH(A1:A4)+FLOOR.MATH(A1:A4)',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'CEILING.PRECISE(A1:A4,2)+FLOOR.PRECISE(A1:A4,2)',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'EVEN(A1:A4)+ODD(A1:A4)',
			styleId: sid,
		})
		sheet.cells.set(0, 7, {
			value: EMPTY,
			formula: 'MROUND(A1:A4,B1:B4)+0',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(-2))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(-6))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(-3))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(3, 3)?.value).toEqual(numberValue(-7))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(-5))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(13))
		expect(sheet.cells.get(3, 4)?.value).toEqual(numberValue(-13))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(-6))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(14))
		expect(sheet.cells.get(3, 5)?.value).toEqual(numberValue(-14))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(-7))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(15))
		expect(sheet.cells.get(3, 6)?.value).toEqual(numberValue(-15))
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(-3))
		expect(sheet.cells.get(2, 7)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(3, 7)?.value).toEqual(numberValue(-7))
	})

	test('logarithmic scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(100), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'LOG10(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'LOG(A1:A3,10)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'LN(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'EXP(B1:B3)+0',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(Math.LN10))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(Math.log(100)))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(Math.E))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(Math.exp(2)))
	})

	test('trigonometric scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(45), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(60), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(0.5), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SIN(RADIANS(A1:A4))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'COS(RADIANS(A1:A4))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'TAN(RADIANS(A1:A4))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'DEGREES(ASIN(B1:B3))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'DEGREES(ACOS(B1:B3))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 7, {
			value: EMPTY,
			formula: 'DEGREES(ATAN(B1:B3))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 8, {
			value: EMPTY,
			formula: 'DEGREES(ATAN2(1,B1:B3))+0',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 4; row++) {
			const degrees = sheet.cells.get(row, 0)?.value
			expect(degrees?.kind).toBe('number')
			if (degrees?.kind !== 'number') continue
			const radians = (degrees.value * Math.PI) / 180
			const sin = sheet.cells.get(row, 2)?.value
			const cos = sheet.cells.get(row, 3)?.value
			const tan = sheet.cells.get(row, 4)?.value
			expect(sin?.kind).toBe('number')
			expect(cos?.kind).toBe('number')
			expect(tan?.kind).toBe('number')
			if (sin?.kind === 'number') expect(sin.value).toBeCloseTo(Math.sin(radians), 12)
			if (cos?.kind === 'number') expect(cos.value).toBeCloseTo(Math.cos(radians), 12)
			if (tan?.kind === 'number') expect(tan.value).toBeCloseTo(Math.tan(radians), 12)
		}
		for (let row = 0; row < 3; row++) {
			const input = sheet.cells.get(row, 1)?.value
			expect(input?.kind).toBe('number')
			if (input?.kind !== 'number') continue
			const asin = sheet.cells.get(row, 5)?.value
			const acos = sheet.cells.get(row, 6)?.value
			const atan = sheet.cells.get(row, 7)?.value
			const atan2 = sheet.cells.get(row, 8)?.value
			expect(asin?.kind).toBe('number')
			expect(acos?.kind).toBe('number')
			expect(atan?.kind).toBe('number')
			expect(atan2?.kind).toBe('number')
			if (asin?.kind === 'number')
				expect(asin.value).toBeCloseTo((Math.asin(input.value) * 180) / Math.PI, 12)
			if (acos?.kind === 'number')
				expect(acos.value).toBeCloseTo((Math.acos(input.value) * 180) / Math.PI, 12)
			if (atan?.kind === 'number')
				expect(atan.value).toBeCloseTo((Math.atan(input.value) * 180) / Math.PI, 12)
			if (atan2?.kind === 'number')
				expect(atan2.value).toBeCloseTo((Math.atan2(input.value, 1) * 180) / Math.PI, 12)
		}
	})

	test('hyperbolic and reciprocal trig scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0.5), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SINH(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'COSH(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'TANH(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'COT(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'SEC(A1:A3)+CSC(A1:A3)',
			styleId: sid,
		})
		sheet.cells.set(0, 7, {
			value: EMPTY,
			formula: 'COTH(A1:A3)+CSCH(A1:A3)+SECH(A1:A3)',
			styleId: sid,
		})
		sheet.cells.set(0, 8, {
			value: EMPTY,
			formula: 'ACOT(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 9, {
			value: EMPTY,
			formula: 'ACOSH(B1:B3)+ASINH(A1:A3)+ATANH(A1:A3/4)+ACOTH(B1:B3)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			const input = sheet.cells.get(row, 0)?.value
			const acoshInput = sheet.cells.get(row, 1)?.value
			expect(input?.kind).toBe('number')
			expect(acoshInput?.kind).toBe('number')
			if (input?.kind !== 'number' || acoshInput?.kind !== 'number') continue
			const sinh = sheet.cells.get(row, 2)?.value
			const cosh = sheet.cells.get(row, 3)?.value
			const tanh = sheet.cells.get(row, 4)?.value
			const cot = sheet.cells.get(row, 5)?.value
			const reciprocal = sheet.cells.get(row, 6)?.value
			const hyperbolicReciprocal = sheet.cells.get(row, 7)?.value
			const acot = sheet.cells.get(row, 8)?.value
			const inverseHyperbolic = sheet.cells.get(row, 9)?.value
			expect(sinh?.kind).toBe('number')
			expect(cosh?.kind).toBe('number')
			expect(tanh?.kind).toBe('number')
			expect(cot?.kind).toBe('number')
			expect(reciprocal?.kind).toBe('number')
			expect(hyperbolicReciprocal?.kind).toBe('number')
			expect(acot?.kind).toBe('number')
			expect(inverseHyperbolic?.kind).toBe('number')
			if (sinh?.kind === 'number') expect(sinh.value).toBeCloseTo(Math.sinh(input.value), 12)
			if (cosh?.kind === 'number') expect(cosh.value).toBeCloseTo(Math.cosh(input.value), 12)
			if (tanh?.kind === 'number') expect(tanh.value).toBeCloseTo(Math.tanh(input.value), 12)
			if (cot?.kind === 'number')
				expect(cot.value).toBeCloseTo(Math.cos(input.value) / Math.sin(input.value), 12)
			if (reciprocal?.kind === 'number') {
				expect(reciprocal.value).toBeCloseTo(
					1 / Math.cos(input.value) + 1 / Math.sin(input.value),
					12,
				)
			}
			if (hyperbolicReciprocal?.kind === 'number') {
				expect(hyperbolicReciprocal.value).toBeCloseTo(
					Math.cosh(input.value) / Math.sinh(input.value) +
						1 / Math.sinh(input.value) +
						1 / Math.cosh(input.value),
					12,
				)
			}
			if (acot?.kind === 'number')
				expect(acot.value).toBeCloseTo(Math.PI / 2 - Math.atan(input.value), 12)
			if (inverseHyperbolic?.kind === 'number') {
				expect(inverseHyperbolic.value).toBeCloseTo(
					Math.acosh(acoshInput.value) +
						Math.asinh(input.value) +
						Math.atanh(input.value / 4) +
						0.5 * Math.log((acoshInput.value + 1) / (acoshInput.value - 1)),
					12,
				)
			}
		}
	})

	test('combinatorics scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1.9), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(8), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'FACT(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'FACTDOUBLE(A1:A3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 5, {
			value: EMPTY,
			formula: 'COMBIN(B1:B3,C1:C3)+0',
			styleId: sid,
		})
		sheet.cells.set(0, 6, {
			value: EMPTY,
			formula: 'COMBINA(B1:B3,C1:C3)+0',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(120))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(15))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(28))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(36))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(56))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(35))
	})

	test('statistical distribution scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const xValues = [0.5, 1.25, 2.5]
		const probabilities = [0.2, 0.5, 0.8]
		const means = [0, 1, 2]
		const standardDeviations = [1, 1.5, 2]
		const trials = [10, 20, 30]
		const successes = [2, 5, 12]
		const successProbabilities = [0.2, 0.3, 0.4]
		const upperSuccesses = [4, 8, 15]
		const degrees1 = [5, 10, 15]
		const degrees2 = [6, 12, 18]
		const alphas = [2, 3, 4]
		const betas = [3, 4, 5]
		const populations = [50, 60, 70]
		for (let row = 0; row < 3; row++) {
			for (const [col, values] of [
				[0, xValues],
				[1, probabilities],
				[2, means],
				[3, standardDeviations],
				[4, trials],
				[5, successes],
				[6, successProbabilities],
				[7, upperSuccesses],
				[8, degrees1],
				[9, degrees2],
				[10, alphas],
				[11, betas],
				[12, populations],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(values[row] as number),
					formula: null,
					styleId: sid,
				})
			}
		}
		const arrayFormulas = [
			'NORM.DIST(A1:A3,C1:C3,D1:D3,TRUE)+0',
			'NORM.INV(B1:B3,C1:C3,D1:D3)+0',
			'NORM.S.DIST(A1:A3,TRUE)+0',
			'NORM.S.INV(B1:B3)+0',
			'BINOM.DIST(F1:F3,E1:E3,G1:G3,TRUE)+0',
			'BINOM.INV(E1:E3,G1:G3,B1:B3)+0',
			'BINOM.DIST.RANGE(E1:E3,G1:G3,F1:F3,H1:H3)+0',
			'POISSON.DIST(F1:F3,K1:K3,TRUE)+0',
			'EXPON.DIST(A1:A3,K1:K3,TRUE)+0',
			'WEIBULL.DIST(A1:A3,K1:K3,L1:L3,TRUE)+0',
			'GAMMA.DIST(A1:A3,K1:K3,L1:L3,TRUE)+0',
			'GAMMA.INV(B1:B3,K1:K3,L1:L3)+0',
			'LOGNORM.DIST(A1:A3,C1:C3,D1:D3,TRUE)+0',
			'LOGNORM.INV(B1:B3,C1:C3,D1:D3)+0',
			'NEGBINOM.DIST(F1:F3,E1:E3,G1:G3,TRUE)+0',
			'HYPGEOM.DIST(F1:F3,E1:E3,H1:H3,M1:M3,TRUE)+0',
			'BETA.DIST(B1:B3,K1:K3,L1:L3,TRUE)+0',
			'BETA.INV(B1:B3,K1:K3,L1:L3)+0',
			'CHISQ.DIST(A1:A3,I1:I3,TRUE)+0',
			'CHISQ.DIST.RT(A1:A3,I1:I3)+0',
			'CHISQ.INV(B1:B3,I1:I3)+0',
			'CHISQ.INV.RT(B1:B3,I1:I3)+0',
			'T.DIST(A1:A3,I1:I3,TRUE)+0',
			'T.DIST.2T(A1:A3,I1:I3)+0',
			'T.DIST.RT(A1:A3,I1:I3)+0',
			'T.INV(B1:B3,I1:I3)+0',
			'T.INV.2T(B1:B3,I1:I3)+0',
			'F.DIST(A1:A3,I1:I3,J1:J3,TRUE)+0',
			'F.DIST.RT(A1:A3,I1:I3,J1:J3)+0',
			'F.INV(B1:B3,I1:I3,J1:J3)+0',
			'F.INV.RT(B1:B3,I1:I3,J1:J3)+0',
			'FISHER(B1:B3/2)+0',
			'FISHERINV(A1:A3)+0',
			'STANDARDIZE(A1:A3,C1:C3,D1:D3)+0',
			'PHI(A1:A3)+0',
			'GAUSS(A1:A3)+0',
			'GAMMA(A1:A3)+0',
			'GAMMALN(A1:A3)+0',
			'GAMMALN.PRECISE(A1:A3)+0',
			'CONFIDENCE.NORM(B1:B3,D1:D3,E1:E3)+0',
			'CONFIDENCE.T(B1:B3,D1:D3,E1:E3)+0',
		]
		const scalarFormulas = [
			'NORM.DIST(A{r},C{r},D{r},TRUE)+0',
			'NORM.INV(B{r},C{r},D{r})+0',
			'NORM.S.DIST(A{r},TRUE)+0',
			'NORM.S.INV(B{r})+0',
			'BINOM.DIST(F{r},E{r},G{r},TRUE)+0',
			'BINOM.INV(E{r},G{r},B{r})+0',
			'BINOM.DIST.RANGE(E{r},G{r},F{r},H{r})+0',
			'POISSON.DIST(F{r},K{r},TRUE)+0',
			'EXPON.DIST(A{r},K{r},TRUE)+0',
			'WEIBULL.DIST(A{r},K{r},L{r},TRUE)+0',
			'GAMMA.DIST(A{r},K{r},L{r},TRUE)+0',
			'GAMMA.INV(B{r},K{r},L{r})+0',
			'LOGNORM.DIST(A{r},C{r},D{r},TRUE)+0',
			'LOGNORM.INV(B{r},C{r},D{r})+0',
			'NEGBINOM.DIST(F{r},E{r},G{r},TRUE)+0',
			'HYPGEOM.DIST(F{r},E{r},H{r},M{r},TRUE)+0',
			'BETA.DIST(B{r},K{r},L{r},TRUE)+0',
			'BETA.INV(B{r},K{r},L{r})+0',
			'CHISQ.DIST(A{r},I{r},TRUE)+0',
			'CHISQ.DIST.RT(A{r},I{r})+0',
			'CHISQ.INV(B{r},I{r})+0',
			'CHISQ.INV.RT(B{r},I{r})+0',
			'T.DIST(A{r},I{r},TRUE)+0',
			'T.DIST.2T(A{r},I{r})+0',
			'T.DIST.RT(A{r},I{r})+0',
			'T.INV(B{r},I{r})+0',
			'T.INV.2T(B{r},I{r})+0',
			'F.DIST(A{r},I{r},J{r},TRUE)+0',
			'F.DIST.RT(A{r},I{r},J{r})+0',
			'F.INV(B{r},I{r},J{r})+0',
			'F.INV.RT(B{r},I{r},J{r})+0',
			'FISHER(B{r}/2)+0',
			'FISHERINV(A{r})+0',
			'STANDARDIZE(A{r},C{r},D{r})+0',
			'PHI(A{r})+0',
			'GAUSS(A{r})+0',
			'GAMMA(A{r})+0',
			'GAMMALN(A{r})+0',
			'GAMMALN.PRECISE(A{r})+0',
			'CONFIDENCE.NORM(B{r},D{r},E{r})+0',
			'CONFIDENCE.T(B{r},D{r},E{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 13 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 60 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				const actual = sheet.cells.get(row, 13 + col)?.value
				const expected = sheet.cells.get(row, 60 + col)?.value
				expect(actual?.kind).toBe('number')
				expect(expected?.kind).toBe('number')
				if (actual?.kind === 'number' && expected?.kind === 'number') {
					expect(actual.value).toBeCloseTo(expected.value, 10)
				}
			}
		}
	})

	test('order statistic functions spill over array k arguments while preserving data ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [9, 4, 7, 1, 5]
		const ks = [1, 2, 4]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < ks.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(ks[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'LARGE(A1:A5,B1:B3)+0', styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SMALL(A1:A5,B1:B3)+0', styleId: sid })
		for (let row = 0; row < ks.length; row++) {
			sheet.cells.set(row, 4, {
				value: EMPTY,
				formula: `LARGE(A1:A5,B${row + 1})+0`,
				styleId: sid,
			})
			sheet.cells.set(row, 5, {
				value: EMPTY,
				formula: `SMALL(A1:A5,B${row + 1})+0`,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < ks.length; row++) {
			expect(sheet.cells.get(row, 2)?.value).toEqual(sheet.cells.get(row, 4)?.value)
			expect(sheet.cells.get(row, 3)?.value).toEqual(sheet.cells.get(row, 5)?.value)
		}
	})

	test('top-level order statistic functions implicitly intersect range k arguments', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [9, 4, 7, 1, 5]
		const ks = [1, 2, 4]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < ks.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(ks[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(1, 2, { value: EMPTY, formula: 'LARGE(A1:A5,B1:B3)', styleId: sid })
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'LARGE(A1:A5,B2)', styleId: sid })
		sheet.cells.set(1, 4, { value: EMPTY, formula: 'SMALL(A1:A5,B1:B3)', styleId: sid })
		sheet.cells.set(1, 5, { value: EMPTY, formula: 'SMALL(A1:A5,B2)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 2)?.value).toEqual(sheet.cells.get(1, 3)?.value)
		expect(sheet.cells.get(1, 4)?.value).toEqual(sheet.cells.get(1, 5)?.value)
	})

	test('rank functions spill over number and order arguments while preserving reference ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [9, 4, 7, 7, 5]
		const numbers = [9, 7, 4]
		const orders = [0, 0, 1]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < numbers.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(numbers[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(orders[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = [
			'RANK(B1:B3,A1:A5,C1:C3)+0',
			'RANK.EQ(B1:B3,A1:A5,C1:C3)+0',
			'RANK.AVG(B1:B3,A1:A5,C1:C3)+0',
		]
		const scalarFormulas = [
			'RANK(B{r},A1:A5,C{r})+0',
			'RANK.EQ(B{r},A1:A5,C{r})+0',
			'RANK.AVG(B{r},A1:A5,C{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 3 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < numbers.length; row++) {
				sheet.cells.set(row, 8 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < numbers.length; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 3 + col)?.value).toEqual(sheet.cells.get(row, 8 + col)?.value)
			}
		}
	})

	test('top-level rank functions implicitly intersect number and order ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [9, 4, 7, 7, 5]
		const numbers = [9, 7, 4]
		const orders = [0, 0, 1]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < numbers.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(numbers[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(orders[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const rangeFormulas = [
			'RANK(B1:B3,A1:A5,C1:C3)',
			'RANK.EQ(B1:B3,A1:A5,C1:C3)',
			'RANK.AVG(B1:B3,A1:A5,C1:C3)',
		]
		const scalarFormulas = ['RANK(B2,A1:A5,C2)', 'RANK.EQ(B2,A1:A5,C2)', 'RANK.AVG(B2,A1:A5,C2)']
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 3 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 8 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 3 + col)?.value).toEqual(sheet.cells.get(1, 8 + col)?.value)
		}
	})

	test('percentile and quartile functions spill over selector arguments while preserving data ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 3, 2, 4, 8, 10]
		const percentiles = [0.25, 0.5, 0.75]
		const quartiles = [1, 2, 3]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < percentiles.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(percentiles[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(quartiles[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = [
			'PERCENTILE(A1:A6,B1:B3)+0',
			'PERCENTILE.INC(A1:A6,B1:B3)+0',
			'PERCENTILE.EXC(A1:A6,B1:B3)+0',
			'QUARTILE(A1:A6,C1:C3)+0',
			'QUARTILE.INC(A1:A6,C1:C3)+0',
			'QUARTILE.EXC(A1:A6,C1:C3)+0',
		]
		const scalarFormulas = [
			'PERCENTILE(A1:A6,B{r})+0',
			'PERCENTILE.INC(A1:A6,B{r})+0',
			'PERCENTILE.EXC(A1:A6,B{r})+0',
			'QUARTILE(A1:A6,C{r})+0',
			'QUARTILE.INC(A1:A6,C{r})+0',
			'QUARTILE.EXC(A1:A6,C{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 3 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < percentiles.length; row++) {
				sheet.cells.set(row, 12 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < percentiles.length; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 3 + col)?.value).toEqual(sheet.cells.get(row, 12 + col)?.value)
			}
		}
	})

	test('top-level percentile and quartile functions implicitly intersect selector ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 3, 2, 4, 8, 10]
		const percentiles = [0.25, 0.5, 0.75]
		const quartiles = [1, 2, 3]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < percentiles.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(percentiles[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(quartiles[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const rangeFormulas = [
			'PERCENTILE(A1:A6,B1:B3)',
			'PERCENTILE.INC(A1:A6,B1:B3)',
			'PERCENTILE.EXC(A1:A6,B1:B3)',
			'QUARTILE(A1:A6,C1:C3)',
			'QUARTILE.INC(A1:A6,C1:C3)',
			'QUARTILE.EXC(A1:A6,C1:C3)',
		]
		const scalarFormulas = [
			'PERCENTILE(A1:A6,B2)',
			'PERCENTILE.INC(A1:A6,B2)',
			'PERCENTILE.EXC(A1:A6,B2)',
			'QUARTILE(A1:A6,C2)',
			'QUARTILE.INC(A1:A6,C2)',
			'QUARTILE.EXC(A1:A6,C2)',
		]
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 3 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 12 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 3 + col)?.value).toEqual(sheet.cells.get(1, 12 + col)?.value)
		}
	})

	test('percentrank functions spill over lookup and significance arguments while preserving data ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 3, 2, 4, 8, 10]
		const lookups = [2, 4, 8]
		const significance = [3, 4, 2]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < lookups.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(lookups[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(significance[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = [
			'PERCENTRANK(A1:A6,B1:B3,C1:C3)+0',
			'PERCENTRANK.INC(A1:A6,B1:B3,C1:C3)+0',
			'PERCENTRANK.EXC(A1:A6,B1:B3,C1:C3)+0',
		]
		const scalarFormulas = [
			'PERCENTRANK(A1:A6,B{r},C{r})+0',
			'PERCENTRANK.INC(A1:A6,B{r},C{r})+0',
			'PERCENTRANK.EXC(A1:A6,B{r},C{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 3 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < lookups.length; row++) {
				sheet.cells.set(row, 8 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < lookups.length; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 3 + col)?.value).toEqual(sheet.cells.get(row, 8 + col)?.value)
			}
		}
	})

	test('top-level percentrank functions implicitly intersect lookup and significance ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 3, 2, 4, 8, 10]
		const lookups = [2, 4, 8]
		const significance = [3, 4, 2]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < lookups.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(lookups[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(significance[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const rangeFormulas = [
			'PERCENTRANK(A1:A6,B1:B3,C1:C3)',
			'PERCENTRANK.INC(A1:A6,B1:B3,C1:C3)',
			'PERCENTRANK.EXC(A1:A6,B1:B3,C1:C3)',
		]
		const scalarFormulas = [
			'PERCENTRANK(A1:A6,B2,C2)',
			'PERCENTRANK.INC(A1:A6,B2,C2)',
			'PERCENTRANK.EXC(A1:A6,B2,C2)',
		]
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 3 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 8 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 3 + col)?.value).toEqual(sheet.cells.get(1, 8 + col)?.value)
		}
	})

	test('forecast functions spill over target x values while preserving known ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const knownXs = [1, 2, 3, 4, 5]
		const knownYs = [3, 5, 7, 9, 11]
		const targetXs = [6, 7, 8]
		for (let row = 0; row < knownXs.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(knownXs[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(knownYs[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < targetXs.length; row++) {
			sheet.cells.set(row, 2, {
				value: numberValue(targetXs[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = ['FORECAST(C1:C3,B1:B5,A1:A5)+0', 'FORECAST.LINEAR(C1:C3,B1:B5,A1:A5)+0']
		const scalarFormulas = ['FORECAST(C{r},B1:B5,A1:A5)+0', 'FORECAST.LINEAR(C{r},B1:B5,A1:A5)+0']
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 3 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < targetXs.length; row++) {
				sheet.cells.set(row, 8 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < targetXs.length; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 3 + col)?.value).toEqual(sheet.cells.get(row, 8 + col)?.value)
			}
		}
	})

	test('top-level forecast functions implicitly intersect target x ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const knownXs = [1, 2, 3, 4, 5]
		const knownYs = [3, 5, 7, 9, 11]
		const targetXs = [6, 7, 8]
		for (let row = 0; row < knownXs.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(knownXs[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(knownYs[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < targetXs.length; row++) {
			sheet.cells.set(row, 2, {
				value: numberValue(targetXs[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const rangeFormulas = ['FORECAST(C1:C3,B1:B5,A1:A5)', 'FORECAST.LINEAR(C1:C3,B1:B5,A1:A5)']
		const scalarFormulas = ['FORECAST(C2,B1:B5,A1:A5)', 'FORECAST.LINEAR(C2,B1:B5,A1:A5)']
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 3 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 8 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 3 + col)?.value).toEqual(sheet.cells.get(1, 8 + col)?.value)
		}
	})

	test('TRIMMEAN spills over percent arguments while preserving the data range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 2, 3, 4, 5, 6, 30, 40]
		const percents = [0, 0.25, 0.5]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < percents.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(percents[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'TRIMMEAN(A1:A8,B1:B3)+0',
			styleId: sid,
		})
		for (let row = 0; row < percents.length; row++) {
			sheet.cells.set(row, 4, {
				value: EMPTY,
				formula: `TRIMMEAN(A1:A8,B${row + 1})+0`,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < percents.length; row++) {
			expect(sheet.cells.get(row, 2)?.value).toEqual(sheet.cells.get(row, 4)?.value)
		}
	})

	test('top-level TRIMMEAN implicitly intersects percent ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 2, 3, 4, 5, 6, 30, 40]
		const percents = [0, 0.25, 0.5]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < percents.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(percents[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(1, 2, {
			value: EMPTY,
			formula: 'TRIMMEAN(A1:A8,B1:B3)',
			styleId: sid,
		})
		sheet.cells.set(1, 4, {
			value: EMPTY,
			formula: 'TRIMMEAN(A1:A8,B2)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 2)?.value).toEqual(sheet.cells.get(1, 4)?.value)
	})

	test('AGGREGATE selector functions spill over k while preserving the data range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 3, 2, 4, 8, 10]
		const indexes = [1, 2, 3]
		const percentiles = [0.25, 0.5, 0.75]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < indexes.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(indexes[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(percentiles[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = [
			'AGGREGATE(14,0,A1:A6,B1:B3)+0',
			'AGGREGATE(15,0,A1:A6,B1:B3)+0',
			'AGGREGATE(16,0,A1:A6,C1:C3)+0',
			'AGGREGATE(17,0,A1:A6,B1:B3)+0',
			'AGGREGATE(18,0,A1:A6,C1:C3)+0',
			'AGGREGATE(19,0,A1:A6,B1:B3)+0',
		]
		const scalarFormulas = [
			'AGGREGATE(14,0,A1:A6,B{r})+0',
			'AGGREGATE(15,0,A1:A6,B{r})+0',
			'AGGREGATE(16,0,A1:A6,C{r})+0',
			'AGGREGATE(17,0,A1:A6,B{r})+0',
			'AGGREGATE(18,0,A1:A6,C{r})+0',
			'AGGREGATE(19,0,A1:A6,B{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 3 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < indexes.length; row++) {
				sheet.cells.set(row, 12 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < indexes.length; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 3 + col)?.value).toEqual(sheet.cells.get(row, 12 + col)?.value)
			}
		}
	})

	test('top-level AGGREGATE selector functions implicitly intersect k ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const values = [1, 3, 2, 4, 8, 10]
		const indexes = [1, 2, 3]
		const percentiles = [0.25, 0.5, 0.75]
		for (let row = 0; row < values.length; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(values[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < indexes.length; row++) {
			sheet.cells.set(row, 1, {
				value: numberValue(indexes[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(percentiles[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const rangeFormulas = [
			'AGGREGATE(14,0,A1:A6,B1:B3)',
			'AGGREGATE(16,0,A1:A6,C1:C3)',
			'AGGREGATE(19,0,A1:A6,B1:B3)',
		]
		const scalarFormulas = [
			'AGGREGATE(14,0,A1:A6,B2)',
			'AGGREGATE(16,0,A1:A6,C2)',
			'AGGREGATE(19,0,A1:A6,B2)',
		]
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 3 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 8 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 3 + col)?.value).toEqual(sheet.cells.get(1, 8 + col)?.value)
		}
	})

	test('AGGREGATE reference form preserves multiple range arguments instead of spilling them', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(row + 1),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(row + 4),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'AGGREGATE(9,4,A1:A3,B1:B3)+0',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(21))
		expect(sheet.cells.get(1, 2)).toBeUndefined()
	})

	test('legacy statistical compatibility functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const xValues = [0.5, 1.25, 2.5]
		const probabilities = [0.2, 0.5, 0.8]
		const means = [0, 1, 2]
		const standardDeviations = [1, 1.5, 2]
		const trials = [10, 20, 30]
		const successes = [2, 5, 12]
		const successProbabilities = [0.2, 0.3, 0.4]
		const upperSuccesses = [4, 8, 15]
		const degrees1 = [5, 10, 15]
		const degrees2 = [6, 12, 18]
		const alphas = [2, 3, 4]
		const betas = [3, 4, 5]
		const populations = [50, 60, 70]
		const tails = [1, 2, 1]
		for (let row = 0; row < 3; row++) {
			for (const [col, values] of [
				[0, xValues],
				[1, probabilities],
				[2, means],
				[3, standardDeviations],
				[4, trials],
				[5, successes],
				[6, successProbabilities],
				[7, upperSuccesses],
				[8, degrees1],
				[9, degrees2],
				[10, alphas],
				[11, betas],
				[12, populations],
				[13, tails],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(values[row] as number),
					formula: null,
					styleId: sid,
				})
			}
		}
		const arrayFormulas = [
			'BETADIST(B1:B3,K1:K3,L1:L3)+0',
			'BETAINV(B1:B3,K1:K3,L1:L3)+0',
			'BINOMDIST(F1:F3,E1:E3,G1:G3,TRUE)+0',
			'CRITBINOM(E1:E3,G1:G3,B1:B3)+0',
			'CHIDIST(A1:A3,I1:I3)+0',
			'CHIINV(B1:B3,I1:I3)+0',
			'CONFIDENCE(B1:B3,D1:D3,E1:E3)+0',
			'EXPONDIST(A1:A3,K1:K3,TRUE)+0',
			'FDIST(A1:A3,I1:I3,J1:J3)+0',
			'FINV(B1:B3,I1:I3,J1:J3)+0',
			'GAMMADIST(A1:A3,K1:K3,L1:L3,TRUE)+0',
			'GAMMAINV(B1:B3,K1:K3,L1:L3)+0',
			'HYPGEOMDIST(F1:F3,E1:E3,H1:H3,M1:M3)+0',
			'LOGINV(B1:B3,C1:C3,D1:D3)+0',
			'LOGNORMDIST(A1:A3,C1:C3,D1:D3)+0',
			'NEGBINOMDIST(F1:F3,E1:E3,G1:G3)+0',
			'NORMDIST(A1:A3,C1:C3,D1:D3,TRUE)+0',
			'NORMINV(B1:B3,C1:C3,D1:D3)+0',
			'NORMSDIST(A1:A3)+0',
			'NORMSINV(B1:B3)+0',
			'POISSON(F1:F3,K1:K3,TRUE)+0',
			'TDIST(A1:A3,I1:I3,N1:N3)+0',
			'TINV(B1:B3,I1:I3)+0',
			'WEIBULL(A1:A3,K1:K3,L1:L3,TRUE)+0',
		]
		const scalarFormulas = [
			'BETADIST(B{r},K{r},L{r})+0',
			'BETAINV(B{r},K{r},L{r})+0',
			'BINOMDIST(F{r},E{r},G{r},TRUE)+0',
			'CRITBINOM(E{r},G{r},B{r})+0',
			'CHIDIST(A{r},I{r})+0',
			'CHIINV(B{r},I{r})+0',
			'CONFIDENCE(B{r},D{r},E{r})+0',
			'EXPONDIST(A{r},K{r},TRUE)+0',
			'FDIST(A{r},I{r},J{r})+0',
			'FINV(B{r},I{r},J{r})+0',
			'GAMMADIST(A{r},K{r},L{r},TRUE)+0',
			'GAMMAINV(B{r},K{r},L{r})+0',
			'HYPGEOMDIST(F{r},E{r},H{r},M{r})+0',
			'LOGINV(B{r},C{r},D{r})+0',
			'LOGNORMDIST(A{r},C{r},D{r})+0',
			'NEGBINOMDIST(F{r},E{r},G{r})+0',
			'NORMDIST(A{r},C{r},D{r},TRUE)+0',
			'NORMINV(B{r},C{r},D{r})+0',
			'NORMSDIST(A{r})+0',
			'NORMSINV(B{r})+0',
			'POISSON(F{r},K{r},TRUE)+0',
			'TDIST(A{r},I{r},N{r})+0',
			'TINV(B{r},I{r})+0',
			'WEIBULL(A{r},K{r},L{r},TRUE)+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 14 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 40 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				const actual = sheet.cells.get(row, 14 + col)?.value
				const expected = sheet.cells.get(row, 40 + col)?.value
				expect(actual?.kind).toBe('number')
				expect(expected?.kind).toBe('number')
				if (actual?.kind === 'number' && expected?.kind === 'number') {
					expect(actual.value).toBeCloseTo(expected.value, 10)
				}
			}
		}
	})

	test('loan financial scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const rates = [0.005, 0.004, 0.006]
		const periods = [12, 24, 36]
		const presentValues = [1000, 5000, 8000]
		const payments = [-86.06642970708324, -218.9093792032636, -246.27134933277662]
		const futureValues = [0, 0, 0]
		const types = [0, 0, 1]
		const paymentPeriods = [1, 6, 12]
		for (let row = 0; row < 3; row++) {
			for (const [col, values] of [
				[0, rates],
				[1, periods],
				[2, presentValues],
				[3, payments],
				[4, futureValues],
				[5, types],
				[6, paymentPeriods],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(values[row] as number),
					formula: null,
					styleId: sid,
				})
			}
		}
		const arrayFormulas = [
			'PMT(A1:A3,B1:B3,C1:C3,E1:E3,F1:F3)+0',
			'FV(A1:A3,B1:B3,D1:D3,C1:C3,F1:F3)+0',
			'PV(A1:A3,B1:B3,D1:D3,E1:E3,F1:F3)+0',
			'NPER(A1:A3,D1:D3,C1:C3,E1:E3,F1:F3)+0',
			'RATE(B1:B3,D1:D3,C1:C3,E1:E3,F1:F3)+0',
			'IPMT(A1:A3,G1:G3,B1:B3,C1:C3,E1:E3,F1:F3)+0',
			'PPMT(A1:A3,G1:G3,B1:B3,C1:C3,E1:E3,F1:F3)+0',
		]
		const scalarFormulas = [
			'PMT(A{r},B{r},C{r},E{r},F{r})+0',
			'FV(A{r},B{r},D{r},C{r},F{r})+0',
			'PV(A{r},B{r},D{r},E{r},F{r})+0',
			'NPER(A{r},D{r},C{r},E{r},F{r})+0',
			'RATE(B{r},D{r},C{r},E{r},F{r})+0',
			'IPMT(A{r},G{r},B{r},C{r},E{r},F{r})+0',
			'PPMT(A{r},G{r},B{r},C{r},E{r},F{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 7 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 14 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				const actual = sheet.cells.get(row, 7 + col)?.value
				const expected = sheet.cells.get(row, 14 + col)?.value
				expect(actual?.kind).toBe('number')
				expect(expected?.kind).toBe('number')
				if (actual?.kind === 'number' && expected?.kind === 'number') {
					expect(actual.value).toBeCloseTo(expected.value, 10)
				}
			}
		}
	})

	test('financial helper scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const rates = [0.1, 0.08, 0.12]
		const periods = [4, 6, 8]
		const presentValues = [1000, 2500, 5000]
		const startPeriods = [1, 2, 3]
		const endPeriods = [1, 4, 6]
		const types = [0, 1, 0]
		const paymentPeriods = [1, 3, 5]
		const compoundsPerYear = [4, 12, 2]
		const futureValues = [2000, 3200, 9000]
		const fractionalDollars = [1.02, 12.08, 100.16]
		const fractions = [16, 8, 32]
		const decimalDollars = [1.125, 12.5, 100.75]
		for (let row = 0; row < 3; row++) {
			for (const [col, values] of [
				[0, rates],
				[1, periods],
				[2, presentValues],
				[3, startPeriods],
				[4, endPeriods],
				[5, types],
				[6, paymentPeriods],
				[7, compoundsPerYear],
				[8, futureValues],
				[9, fractionalDollars],
				[10, fractions],
				[11, decimalDollars],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(values[row] as number),
					formula: null,
					styleId: sid,
				})
			}
		}
		const arrayFormulas = [
			'ISPMT(A1:A3,G1:G3,B1:B3,C1:C3)+0',
			'CUMIPMT(A1:A3,B1:B3,C1:C3,D1:D3,E1:E3,F1:F3)+0',
			'CUMPRINC(A1:A3,B1:B3,C1:C3,D1:D3,E1:E3,F1:F3)+0',
			'EFFECT(A1:A3,H1:H3)+0',
			'NOMINAL(A1:A3,H1:H3)+0',
			'PDURATION(A1:A3,C1:C3,I1:I3)+0',
			'RRI(B1:B3,C1:C3,I1:I3)+0',
			'DOLLARDE(J1:J3,K1:K3)+0',
			'DOLLARFR(L1:L3,K1:K3)+0',
		]
		const scalarFormulas = [
			'ISPMT(A{r},G{r},B{r},C{r})+0',
			'CUMIPMT(A{r},B{r},C{r},D{r},E{r},F{r})+0',
			'CUMPRINC(A{r},B{r},C{r},D{r},E{r},F{r})+0',
			'EFFECT(A{r},H{r})+0',
			'NOMINAL(A{r},H{r})+0',
			'PDURATION(A{r},C{r},I{r})+0',
			'RRI(B{r},C{r},I{r})+0',
			'DOLLARDE(J{r},K{r})+0',
			'DOLLARFR(L{r},K{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 12 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 30 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				const actual = sheet.cells.get(row, 12 + col)?.value
				const expected = sheet.cells.get(row, 30 + col)?.value
				expect(actual?.kind).toBe('number')
				expect(expected?.kind).toBe('number')
				if (actual?.kind === 'number' && expected?.kind === 'number') {
					expect(actual.value).toBeCloseTo(expected.value, 10)
				}
			}
		}
	})

	test('depreciation financial scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const costs = [1000, 2400, 5000]
		const salvageValues = [100, 300, 500]
		const lives = [6, 10, 5]
		const periods = [1, 2, 3]
		const months = [7, 12, 12]
		const startPeriods = [0, 1, 2]
		const endPeriods = [1, 2.5, 3]
		const factors = [2, 1.5, 2]
		const noSwitches = [0, 1, 0]
		for (let row = 0; row < 3; row++) {
			for (const [col, values] of [
				[0, costs],
				[1, salvageValues],
				[2, lives],
				[3, periods],
				[4, months],
				[5, startPeriods],
				[6, endPeriods],
				[7, factors],
				[8, noSwitches],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(values[row] as number),
					formula: null,
					styleId: sid,
				})
			}
		}
		const arrayFormulas = [
			'SLN(A1:A3,B1:B3,C1:C3)+0',
			'SYD(A1:A3,B1:B3,C1:C3,D1:D3)+0',
			'DDB(A1:A3,B1:B3,C1:C3,D1:D3,H1:H3)+0',
			'DB(A1:A3,B1:B3,C1:C3,D1:D3,E1:E3)+0',
			'VDB(A1:A3,B1:B3,C1:C3,F1:F3,G1:G3,H1:H3,I1:I3)+0',
		]
		const scalarFormulas = [
			'SLN(A{r},B{r},C{r})+0',
			'SYD(A{r},B{r},C{r},D{r})+0',
			'DDB(A{r},B{r},C{r},D{r},H{r})+0',
			'DB(A{r},B{r},C{r},D{r},E{r})+0',
			'VDB(A{r},B{r},C{r},F{r},G{r},H{r},I{r})+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 9 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 14 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				const actual = sheet.cells.get(row, 9 + col)?.value
				const expected = sheet.cells.get(row, 14 + col)?.value
				expect(actual?.kind).toBe('number')
				expect(expected?.kind).toBe('number')
				if (actual?.kind === 'number' && expected?.kind === 'number') {
					expect(actual.value).toBeCloseTo(expected.value, 10)
				}
			}
		}
	})

	test('FREQUENCY formulas can count unique filtered numeric ids from array expressions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(11), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula:
				'SUM(--(FREQUENCY((A1:A3=C1)*(IF(ISNUMBER(B1:B3),B1:B3,0)),(A1:A3=C1)*(IF(ISNUMBER(B1:B3),B1:B3,0)))>0))-1',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(2))
	})

	test('ROW and N preserve arrays for unique-id criteria expressions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('EventID'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(1, 1, {
			value: EMPTY,
			formula: 'N(IF(MATCH(A2:A4,A2:A4,0)=ROW(A2:A4)-ROW(A1),A2:A4,0))',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(20))
	})

	test('legacy top-level ROW and COLUMN ranges return top-left scalar', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'ROW(A2:A4)', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'COLUMN(C1:E1)', styleId: sid })
		sheet.cells.set(2, 1, {
			value: EMPTY,
			formula: 'ROW(A2:A4)',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1 },
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(4))
	})

	test('ROW and COLUMN without references use the formula cell position', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(4, 2, { value: EMPTY, formula: 'ROW()', styleId: sid })
		sheet.cells.set(4, 3, { value: EMPTY, formula: 'COLUMN()', styleId: sid })
		sheet.cells.set(4, 4, { value: EMPTY, formula: 'ADDRESS(ROW(),COLUMN())', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 2)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(4, 3)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(4, 4)?.value).toEqual(stringValue('$E$5'))
	})

	test('ADDRESS spills over range row and column arguments in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const rows = [1, 12, 104]
		const cols = [1, 27, 703]
		const modes = [1, 4, 2]
		const useA1 = [true, true, false]
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(rows[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(cols[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(modes[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 3, {
				value: booleanValue(useA1[row] as boolean),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'ADDRESS(A1:A3,B1:B3,C1:C3,D1:D3)&""',
			styleId: sid,
		})
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 5, {
				value: EMPTY,
				formula: `ADDRESS(A${row + 1},B${row + 1},C${row + 1},D${row + 1})&""`,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			expect(sheet.cells.get(row, 4)?.value).toEqual(sheet.cells.get(row, 5)?.value)
		}
	})

	test('top-level ADDRESS implicitly intersects range row and column arguments', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const rows = [1, 12, 104]
		const cols = [1, 27, 703]
		const modes = [1, 4, 2]
		const useA1 = [true, true, false]
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(rows[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(cols[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(modes[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 3, {
				value: booleanValue(useA1[row] as boolean),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(1, 4, {
			value: EMPTY,
			formula: 'ADDRESS(A1:A3,B1:B3,C1:C3,D1:D3)',
			styleId: sid,
		})
		sheet.cells.set(1, 5, { value: EMPTY, formula: 'ADDRESS(A2,B2,C2,D2)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 4)?.value).toEqual(sheet.cells.get(1, 5)?.value)
	})

	test('conditional aggregates return arrays for array criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('EventID'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'SUM(--(COUNTIFS(A2:A4,N(IF(MATCH(A2:A4,A2:A4,0)=ROW(A2:A4)-ROW(A1),A2:A4,0)))>0))',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
	})

	test('binary operators implicitly intersect direct whole-column operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 3, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 4, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 4, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(1, 14, { value: EMPTY, formula: 'E:E/D:D', styleId: sid })
		sheet.cells.set(2, 14, { value: EMPTY, formula: 'E:E/$D:$D', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 14)?.value).toEqual(errorValue('#DIV/0!'))
		expect(sheet.cells.get(2, 14)?.value).toEqual(numberValue(2))
	})

	test('legacy top-level unary and binary range operators implicitly intersect', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: stringValue(''), formula: null, styleId: sid })
		sheet.cells.set(0, 5, { value: stringValue('col'), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: '-A1:C1', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'A1:C1*2', styleId: sid })
		sheet.cells.set(2, 3, { value: EMPTY, formula: 'A1:C1+D1', styleId: sid })
		sheet.cells.set(1, 4, { value: EMPTY, formula: '+E1', styleId: sid })
		sheet.cells.set(1, 5, { value: EMPTY, formula: '+E1:G1', styleId: sid })
		sheet.cells.set(1, 7, { value: EMPTY, formula: '+H1', styleId: sid })
		sheet.cells.set(5, 5, { value: EMPTY, formula: 'A1:B2&D1&F1', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(-2))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(2, 3)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue(''))
		expect(sheet.cells.get(1, 5)?.value).toEqual(stringValue('col'))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(5, 5)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('binary operators implicitly intersect whole-column concat operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: stringValue('north'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('-'), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: stringValue('east'), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'A:A&B:B&C:C', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('north-east'))
	})

	test('binary operators implicitly intersect direct whole-row operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: EMPTY, formula: '1:1+2:2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(8))
	})

	test('scalar text functions implicitly intersect whole-column operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(2, 0, { value: stringValue('1000 to 4999'), formula: null, styleId: sid })
		sheet.cells.set(2, 14, { value: EMPTY, formula: 'RIGHT($A:$A,FIND(" ",$A:$A))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 14)?.value).toEqual(stringValue(' 4999'))
	})

	test('scalar functions implicitly intersect normal range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: errorValue('#VALUE!'), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: numberValue(345), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('abc'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('xy'), formula: null, styleId: sid })
		sheet.cells.set(0, 5, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 6, { value: numberValue(5.1), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'TANH(A1:A3)', styleId: sid })
		sheet.cells.set(1, 6, { value: EMPTY, formula: 'CEILING(F1:G1,A1:A3)', styleId: sid })
		sheet.cells.set(4, 5, { value: EMPTY, formula: 'CEILING(B1:C2,1)', styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'LEN(B1:C1)', styleId: sid })
		sheet.cells.set(1, 4, { value: EMPTY, formula: 'FACT(A1:A3)', styleId: sid })
		sheet.cells.set(4, 4, { value: EMPTY, formula: 'TANH(A1:B2)', styleId: sid })
		sheet.cells.set(4, 7, { value: EMPTY, formula: 'ABS(A4:A5)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(0.964027580075817))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(4, 4)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(4, 5)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(4, 7)?.value).toEqual(numberValue(345))
	})

	test('legacy top-level CONCATENATE implicitly intersects range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 7, { value: stringValue('left'), formula: null, styleId: sid })
		sheet.cells.set(0, 8, { value: stringValue('mid'), formula: null, styleId: sid })
		sheet.cells.set(0, 9, { value: stringValue('right'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('one'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('two'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('three'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('block'), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: stringValue('range'), formula: null, styleId: sid })
		sheet.cells.set(2, 8, { value: EMPTY, formula: 'CONCATENATE(H1:J1,A2:A4)', styleId: sid })
		sheet.cells.set(4, 4, { value: EMPTY, formula: 'CONCATENATE(B2:C3,"x")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 8)?.value).toEqual(stringValue('midtwo'))
		expect(sheet.cells.get(4, 4)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('top-level T evaluates range operands from their top-left value', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 7, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 8, { value: stringValue('mid'), formula: null, styleId: sid })
		sheet.cells.set(0, 9, { value: stringValue('right'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('text'), formula: null, styleId: sid })
		sheet.cells.set(2, 9, { value: EMPTY, formula: 'T(H1:J1)', styleId: sid })
		sheet.cells.set(3, 9, { value: EMPTY, formula: 'T(B2:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 9)?.value).toEqual(stringValue(''))
		expect(sheet.cells.get(3, 9)?.value).toEqual(stringValue(''))
	})

	test('compiled IF formulas preserve whole-column implicit intersection', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 3, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(1000), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(5000), formula: null, styleId: sid })
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 4, {
				value: EMPTY,
				formula: 'IF(D:D<1000,"Less Than 1000",IF(D:D<4999,"1000 to 4999","5000+"))',
				styleId: sid,
			})
		}

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('Less Than 1000'))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue('1000 to 4999'))
		expect(sheet.cells.get(2, 4)?.value).toEqual(stringValue('5000+'))
	})

	test('compiled root IF preserves boolean branch cell values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: booleanValue(true), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: booleanValue(false), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'IF(A1>0,B1,C1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(booleanValue(false))
	})

	test('shared formulas shift relative whole-column refs and preserve absolute whole-column refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 3, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 5, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(0, 7, {
			value: EMPTY,
			formula: 'D:D/$E:$E',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'H1' },
		})
		sheet.cells.set(0, 8, {
			value: EMPTY,
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'H1' },
		})
		sheet.cells.set(0, 9, {
			value: EMPTY,
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'H1' },
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(0, 8)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 9)?.value).toEqual(numberValue(3))
	})

	test('GETPIVOTDATA reads cached visible pivot output by row field', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(3, 17, { value: stringValue('Segment'), formula: null, styleId: sid })
		sheet.cells.set(3, 18, { value: stringValue('Sum of Revenue'), formula: null, styleId: sid })
		sheet.cells.set(3, 19, { value: stringValue('Sum of Costs'), formula: null, styleId: sid })
		sheet.cells.set(4, 17, { value: stringValue('Strategic'), formula: null, styleId: sid })
		sheet.cells.set(4, 18, { value: numberValue(49_478_096), formula: null, styleId: sid })
		sheet.cells.set(4, 19, { value: numberValue(32_125_000), formula: null, styleId: sid })
		sheet.cells.set(5, 17, { value: stringValue('SMB'), formula: null, styleId: sid })
		sheet.cells.set(5, 18, { value: numberValue(49_521_904), formula: null, styleId: sid })
		sheet.cells.set(5, 19, { value: numberValue(31_875_000), formula: null, styleId: sid })
		sheet.cells.set(6, 17, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(6, 18, { value: numberValue(99_000_000), formula: null, styleId: sid })
		sheet.cells.set(16, 17, { value: stringValue('Strategic'), formula: null, styleId: sid })
		sheet.cells.set(16, 18, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sum of Revenue",$R$4,"Segment",$R17)',
			styleId: sid,
		})
		sheet.cells.set(16, 19, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sum of Costs",$R$4,"Segment","Strategic")',
			styleId: sid,
		})
		sheet.cells.set(16, 20, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sum of Revenue",$R$4)',
			styleId: sid,
		})
		sheet.cells.set(16, 21, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Revenue",$R$4,"Segment","Strategic")',
			styleId: sid,
		})
		sheet.cells.set(16, 22, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Revenue",$R$4:$T$6,"Segment","Strategic")',
			styleId: sid,
		})
		sheet.cells.set(16, 23, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Revenue",PivotAnchor,"Segment","Strategic")',
			styleId: sid,
		})
		sheet.cells.set(16, 24, {
			value: EMPTY,
			formula: 'SUM(GETPIVOTDATA("Revenue",$R$4,"Segment",{"Strategic","SMB"}))',
			styleId: sid,
		})
		wb.definedNames.set('PivotAnchor', "'Pivot Tables'!R4:T7")
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'R4:T7',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 0, name: 'Sum of Revenue', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(16, 18)?.value).toEqual(numberValue(49_478_096))
		expect(sheet.cells.get(16, 19)?.value).toEqual(numberValue(32_125_000))
		expect(sheet.cells.get(16, 20)?.value).toEqual(numberValue(99_000_000))
		expect(sheet.cells.get(16, 21)?.value).toEqual(numberValue(49_478_096))
		expect(sheet.cells.get(16, 22)?.value).toEqual(numberValue(49_478_096))
		expect(sheet.cells.get(16, 23)?.value).toEqual(numberValue(49_478_096))
		expect(sheet.cells.get(16, 24)?.value).toEqual(numberValue(99_000_000))
	})

	test('GETPIVOTDATA range anchors prefer the most recently listed overlapping pivot', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(0, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sum of Sales'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: stringValue('Sum of Sales'), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 4, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(3, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1:$E$2)',
			styleId: sid,
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'A1:B2',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 0, name: 'Sum of Sales', subtotal: 'sum' }],
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable2.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable2',
			cacheId: 2,
			locationRef: 'D1:E2',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 0, name: 'Sum of Sales', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(20))
	})

	test('GETPIVOTDATA reads no-filter grand total from column-labeled pivot output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(0, 0, { value: stringValue('Pivot Table'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Sum of Qux'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Column Labels'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(57), formula: null, styleId: sid })
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Qux",$A$1)',
			styleId: sid,
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'A1:C3',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 0, name: 'Sum of Qux', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(57))
	})

	test('GETPIVOTDATA preserves blank grand total cells from saved pivot output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(0, 0, { value: stringValue('Pivot Table'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Sum of Blank'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Column Labels'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(4, 0, {
			value: stringValue('stale'),
			formula: 'GETPIVOTDATA("Blank",$A$1)',
			styleId: sid,
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'A1:C3',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 0, name: 'Sum of Blank', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(EMPTY)
	})

	test('GETPIVOTDATA resolves column field filters from visible pivot output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(0, 0, { value: stringValue('Sum of Sales'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Product'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('North'), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: stringValue('South'), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('Widget'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('Gadget'), formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(3, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(3, 3, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(4, 1, { value: numberValue(11), formula: null, styleId: sid })
		sheet.cells.set(4, 2, { value: numberValue(22), formula: null, styleId: sid })
		sheet.cells.set(4, 3, { value: numberValue(33), formula: null, styleId: sid })
		sheet.cells.set(6, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Region","South")',
			styleId: sid,
		})
		sheet.cells.set(7, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Product","Widget","Region","South")',
			styleId: sid,
		})
		sheet.cells.set(8, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Region","South","Product","Widget")',
			styleId: sid,
		})
		sheet.cells.set(9, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Region","West")',
			styleId: sid,
		})
		sheet.cells.set(10, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Region","2")',
			styleId: sid,
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'A1:D5',
			fields: [
				{ index: 0, axis: 'axisRow', name: 'Product' },
				{ index: 1, axis: 'axisCol', name: 'Region' },
				{ index: 2, dataField: true, name: 'Sales' },
			],
			rowFields: [{ index: 0 }],
			columnFields: [{ index: 1 }],
			pageFields: [],
			dataFields: [{ fieldIndex: 2, name: 'Sum of Sales', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(6, 0)?.value).toEqual(numberValue(22))
		expect(sheet.cells.get(7, 0)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(8, 0)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(9, 0)?.value).toEqual(errorValue('#REF!'))
		expect(sheet.cells.get(10, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('GETPIVOTDATA returns #REF! for missing visible pivot items', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(3, 29, { value: stringValue('Product_Category'), formula: null, styleId: sid })
		sheet.cells.set(3, 30, { value: stringValue('Sum of Revenue'), formula: null, styleId: sid })
		sheet.cells.set(4, 29, { value: stringValue('Services'), formula: null, styleId: sid })
		sheet.cells.set(4, 30, { value: numberValue(98_854_675), formula: null, styleId: sid })
		sheet.cells.set(16, 29, {
			value: EMPTY,
			formula: 'IFERROR(GETPIVOTDATA("Sum of Revenue",$AD$4,"Product_Category","Hardware"),0)',
			styleId: sid,
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable2.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable2',
			cacheId: 1,
			locationRef: 'AD4:AE5',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 0, name: 'Sum of Revenue', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(16, 29)?.value).toEqual(numberValue(0))
	})

	test('GETPIVOTDATA validates report filter arguments against selected page fields', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Pivot Tables')
		sheet.cells.set(0, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sum of Sales'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Grand Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(100), formula: null, styleId: sid })
		sheet.cells.set(3, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Region","North")',
			styleId: sid,
		})
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Region","South")',
			styleId: sid,
		})
		sheet.cells.set(5, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Close Date",DATE(1999,3,5))',
			styleId: sid,
		})
		sheet.cells.set(6, 0, {
			value: EMPTY,
			formula: 'GETPIVOTDATA("Sales",$A$1,"Close Date",DATE(1999,3,6))',
			styleId: sid,
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			fields: [
				{
					index: 0,
					name: 'Region',
					sharedItems: [
						{ index: 0, kind: 'string', value: 'North' },
						{ index: 1, kind: 'string', value: 'South' },
					],
				},
				{ index: 1, name: 'Sales' },
				{
					index: 2,
					name: 'Close Date',
					sharedItems: [
						{ index: 0, kind: 'date', value: '1999-03-05T00:00:00' },
						{ index: 1, kind: 'date', value: '1999-03-06T00:00:00' },
					],
				},
			],
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Pivot Tables',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'A1:B2',
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					name: 'Region',
					items: [
						{ index: 0, cacheIndex: 0 },
						{ index: 1, cacheIndex: 1 },
					],
				},
				{ index: 1, dataField: true, name: 'Sales' },
				{
					index: 2,
					axis: 'axisPage',
					name: 'Close Date',
					items: [
						{ index: 0, cacheIndex: 0 },
						{ index: 1, cacheIndex: 1 },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [
				{ index: 0, item: 0 },
				{ index: 2, item: 0 },
			],
			dataFields: [{ fieldIndex: 1, name: 'Sum of Sales', subtotal: 'sum' }],
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(100))
		expect(sheet.cells.get(4, 0)?.value).toEqual(errorValue('#REF!'))
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(100))
		expect(sheet.cells.get(6, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('binary arithmetic broadcasts row and column arrays', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'A1:A3*B1:C1',
			styleId: sid,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1 },
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(40))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(60))
	})

	test('implicit intersection uses the formula row for single-column ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: '@A:A', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
	})

	test('implicit intersection on whole-column refs can return empty cells outside usedRange', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(9, 1, { value: EMPTY, formula: '@A:A', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(9, 1)?.value).toEqual(EMPTY)
	})

	test('implicit intersection uses the formula column for single-row ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: '@1:1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
	})

	test('SUM supports union references inside a parenthesized single argument', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SUM((A1:A2,C1:C2))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
	})

	test('SUM supports intersection references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'SUM(A:B 2:2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(10))
	})

	test('non-overlapping intersection returns #NULL!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'A1:A2 C1:C2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#NULL!'))
	})

	test('overlapping self-intersection produces circular reference error', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SUM(A1:A2 A1:B1)', styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx())
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('COUNTA, COUNTBLANK, and PRODUCT support union references', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'COUNTA((A1:A2,C1:C2))', styleId: sid })
		sheet.cells.set(1, 4, { value: EMPTY, formula: 'COUNTBLANK((A1:A2,C1:C2))', styleId: sid })
		sheet.cells.set(2, 4, { value: EMPTY, formula: 'PRODUCT((A1:A2,C1:C2))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(24))
	})

	test('SUMPRODUCT returns #VALUE! for mismatched range shapes', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUMPRODUCT(A1:A2,B1:B1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('SUMIFS returns #VALUE! for mismatched criteria shapes', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUMIFS(A1:A2,B1:B1,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('conditional aggregate cache stays isolated and refreshes between recalculations', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const categories = ['East', 'West', 'East', 'West']
		const values = [10, 20, 30, 40]
		for (let row = 0; row < categories.length; row++) {
			sheet.cells.set(row, 0, {
				value: stringValue(categories[row] ?? ''),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(values[row] ?? 0),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'SUMIF(A$1:A$4,"East",B$1:B$4)',
			styleId: sid,
		})
		sheet.cells.set(1, 3, {
			value: EMPTY,
			formula: 'AVERAGEIF(A$1:A$4,"East",B$1:B$4)',
			styleId: sid,
		})
		sheet.cells.set(2, 3, {
			value: EMPTY,
			formula: 'SUMIFS(B$1:B$4,A$1:A$4,"East")',
			styleId: sid,
		})
		sheet.cells.set(3, 3, {
			value: EMPTY,
			formula: 'SUMIFS(B$1:B$4,A$1:A$4,"East")',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(40))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(40))
		expect(sheet.cells.get(3, 3)?.value).toEqual(numberValue(40))

		sheet.cells.set(0, 1, { value: numberValue(100), formula: null, styleId: sid })
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(130))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(65))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(130))
		expect(sheet.cells.get(3, 3)?.value).toEqual(numberValue(130))
	})

	test('order statistic cache refreshes between recalculations', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(row + 1),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'LARGE(A$1:A$5,2)', styleId: sid })
		sheet.cells.set(1, 2, { value: EMPTY, formula: 'LARGE(A$1:A$5,2)', styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'SMALL(A$1:A$5,2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(2))

		sheet.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: sid })
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(3))
	})

	test('current-row structured references resolve within a table body', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: EMPTY, formula: '[@Qty]*[@Price]', styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: '[@Qty]*[@Price]', styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
			columns: [{ name: 'Qty' }, { name: 'Price' }, { name: 'Total' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(20))
	})

	test('legacy #This Row structured references resolve within a table body', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(7), formula: null, styleId: sid })
		sheet.cells.set(1, 2, {
			value: EMPTY,
			formula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]',
			styleId: sid,
		})
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
			columns: [{ name: 'Qty' }, { name: 'Price' }, { name: 'Total' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(42))
	})

	test('structured references support #Headers and #All specifiers', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Qty' }, { name: 'Price' }],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'COLUMNS(Sales[#Headers])', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'ROWS(Sales[#All])', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(3))
	})

	describe('structured table references (Table1 with Name, Revenue, Quantity)', () => {
		function makeTable1Workbook() {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: sid })
			sheet.cells.set(0, 1, { value: stringValue('Revenue'), formula: null, styleId: sid })
			sheet.cells.set(0, 2, { value: stringValue('Quantity'), formula: null, styleId: sid })
			sheet.cells.set(1, 0, { value: stringValue('A'), formula: null, styleId: sid })
			sheet.cells.set(1, 1, { value: numberValue(100), formula: null, styleId: sid })
			sheet.cells.set(1, 2, { value: numberValue(10), formula: null, styleId: sid })
			sheet.cells.set(2, 0, { value: stringValue('B'), formula: null, styleId: sid })
			sheet.cells.set(2, 1, { value: numberValue(200), formula: null, styleId: sid })
			sheet.cells.set(2, 2, { value: numberValue(20), formula: null, styleId: sid })
			sheet.cells.set(3, 0, { value: stringValue('C'), formula: null, styleId: sid })
			sheet.cells.set(3, 1, { value: numberValue(300), formula: null, styleId: sid })
			sheet.cells.set(3, 2, { value: numberValue(30), formula: null, styleId: sid })
			sheet.tables.push({
				id: createTableId(),
				name: 'Table1',
				sheetId: sheet.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
				columns: [{ name: 'Name' }, { name: 'Revenue' }, { name: 'Quantity' }],
				hasHeaders: true,
				hasTotals: false,
			})
			return { wb, sheet }
		}

		test('Table1[Revenue] resolves to data column (used in SUM)', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUM(Table1[Revenue])', styleId: sid })

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(600))
		})

		test('Table1 column ranges resolve all columns inclusively', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, {
				value: EMPTY,
				formula: 'SUM(Table1[[Revenue]:[Quantity]])',
				styleId: sid,
			})
			sheet.cells.set(5, 1, {
				value: EMPTY,
				formula: 'COLUMNS(Table1[[Revenue]:[Quantity]])',
				styleId: sid,
			})

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(660))
			expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(2))
		})

		test('current-row structured reference column ranges resolve same-row cells', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(2, 4, {
				value: EMPTY,
				formula: 'SUM(Table1[@[Revenue]:[Quantity]])',
				styleId: sid,
			})

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(220))
		})

		test('Table1[@Revenue] resolves to same-row value (implicit intersection)', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(1, 4, { value: EMPTY, formula: 'Table1[@Revenue]', styleId: sid })
			sheet.cells.set(2, 4, { value: EMPTY, formula: 'Table1[@Revenue]', styleId: sid })
			sheet.cells.set(3, 4, { value: EMPTY, formula: 'Table1[@Revenue]', styleId: sid })

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(100))
			expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(200))
			expect(sheet.cells.get(3, 4)?.value).toEqual(numberValue(300))
		})

		test('direct structured column references implicitly intersect by formula row', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(1, 4, { value: EMPTY, formula: 'Table1[Revenue]', styleId: sid })
			sheet.cells.set(2, 4, { value: EMPTY, formula: 'Table1[Revenue]', styleId: sid })

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(100))
			expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(200))
		})

		test('structured columns stay array-valued inside binary lookup expressions', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, {
				value: EMPTY,
				formula: 'INDEX(Table1[],MATCH(200,Table1[Revenue]*(Table1[Name]="B"),0),3)',
				styleId: sid,
			})

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(20))
		})

		test('LET bindings remain scalar-valued inside binary expressions', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, {
				value: EMPTY,
				formula: 'LET(_xlpm.SourceCount,2,IF(_xlpm.SourceCount>0,"Includes ",""))',
				styleId: sid,
			})

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(stringValue('Includes '))
		})

		test('Table1[#Headers] references header row', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, { value: EMPTY, formula: 'COLUMNS(Table1[#Headers])', styleId: sid })
			sheet.cells.set(5, 1, { value: EMPTY, formula: 'ROWS(Table1[#Headers])', styleId: sid })

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(3))
			expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(1))
		})

		test('Table1[#All] references all rows including headers', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, { value: EMPTY, formula: 'ROWS(Table1[#All])', styleId: sid })
			sheet.cells.set(5, 1, { value: EMPTY, formula: 'COLUMNS(Table1[#All])', styleId: sid })

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(4))
			expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(3))
		})

		test('Table1[#Data] references data rows only (excludes headers)', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUM(Table1[#Data])', styleId: sid })

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(660))
		})

		test('combined structured reference row specifiers include each contiguous table band', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(4, 0, { value: stringValue('Total'), formula: null, styleId: sid })
			sheet.cells.set(4, 1, { value: numberValue(600), formula: null, styleId: sid })
			sheet.cells.set(4, 2, { value: numberValue(60), formula: null, styleId: sid })
			const table = sheet.tables[0]
			if (!table) throw new Error('expected test table')
			sheet.tables[0] = {
				...table,
				ref: { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } },
				hasTotals: true,
			}
			sheet.cells.set(6, 0, {
				value: EMPTY,
				formula: 'ROWS(Table1[[#Headers],[#Data],[Revenue]])',
				styleId: sid,
			})
			sheet.cells.set(6, 1, {
				value: EMPTY,
				formula: 'SUM(Table1[[#Data],[#Totals],[Revenue]])',
				styleId: sid,
			})

			const result = recalculate(wb, makeCtx())
			expect(result.errors).toEqual([])
			expect(sheet.cells.get(6, 0)?.value).toEqual(numberValue(4))
			expect(sheet.cells.get(6, 1)?.value).toEqual(numberValue(1200))
		})

		test('Table1[#Totals] returns #REF! when table has no totals row', () => {
			const { wb, sheet } = makeTable1Workbook()
			sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUM(Table1[#Totals])', styleId: sid })

			recalculate(wb, makeCtx())
			expect(sheet.cells.get(5, 0)?.value).toEqual(errorValue('#REF!'))
		})
	})

	test('COUNTIF supports wildcard criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('North'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('Northeast'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('South'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('Northwest'), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'COUNTIF(A1:A4,"North*")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(3))
	})

	test('COUNTIF matches error criteria produced by expressions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'COUNTIF(A1:A2,1/0)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(1))
	})

	test('COUNTIF implicitly intersects range criteria by formula column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 5, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 6, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 7, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'COUNTIF(F1:H1,B1:D1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(2))
	})

	test('SUMIF treats non-intersecting range criteria as #VALUE! criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: errorValue('#VALUE!'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: errorValue('#VALUE!'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(2, 3, { value: numberValue(7), formula: null, styleId: sid })
		sheet.cells.set(4, 4, { value: EMPTY, formula: 'SUMIF(A1:A3,B1:C1,D1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(4, 4)?.value).toEqual(numberValue(12))
	})

	test('COUNTIF blank and nonblank criteria match Excel semantics', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue(''), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('value'), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'COUNTIF(A1:A3,"")', styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'COUNTIF(A1:A3,"<>")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(1))
	})

	test('deterministic NOW via CalcContext', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'TODAY()', styleId: sid })

		const fixedDate = new Date(2025, 0, 15)
		const ctx = makeCtx({ today: fixedDate, now: fixedDate })
		recalculate(wb, ctx)

		const cell = sheet.cells.get(0, 0)
		const expectedSerial = dateToSerial(2025, 1, 15)
		expect(cell?.value).toEqual(numberValue(expectedSerial))
	})

	test('volatile cells with unchanged value do not propagate to dependents: TODAY() same date skips dependents', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'TODAY()', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const fixedDate = new Date(2025, 0, 15)
		const ctx = makeCtx({ today: fixedDate, now: fixedDate })

		recalculate(wb, ctx)
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(dateToSerial(2025, 1, 15)))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(dateToSerial(2025, 1, 15) + 1))

		const result = recalculate(wb, ctx, { dirtyOnly: true })
		expect(result.changed).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(dateToSerial(2025, 1, 15)))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(dateToSerial(2025, 1, 15) + 1))
	})

	test('DATE respects 1904 date system in CalcContext', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'DATE(1904,1,2)', styleId: sid })

		recalculate(wb, makeCtx({ dateSystem: '1904' }))

		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
	})

	test('common date and time scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const dates = [dateToSerial(2024, 1, 31), dateToSerial(2024, 2, 29), dateToSerial(2025, 12, 15)]
		const times = [
			(12 * 3600) / 86_400,
			(6 * 3600 + 30 * 60 + 15) / 86_400,
			(23 * 3600 + 59 * 60 + 59) / 86_400,
		]
		const offsets = [1, -1, 2]
		const dateText = ['2024-01-31', '2024-02-29', '12/15/2025']
		const timeText = ['12:00 PM', '6:30:15', '11:59:59 PM']
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(dates[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(offsets[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(times[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 3, {
				value: stringValue(dateText[row] as string),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 4, {
				value: stringValue(timeText[row] as string),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 5, { value: EMPTY, formula: 'YEAR(A1:A3)+0', styleId: sid })
		sheet.cells.set(0, 6, { value: EMPTY, formula: 'MONTH(A1:A3)*100+DAY(A1:A3)', styleId: sid })
		sheet.cells.set(0, 7, { value: EMPTY, formula: 'EOMONTH(A1:A3,B1:B3)+0', styleId: sid })
		sheet.cells.set(0, 8, { value: EMPTY, formula: 'EDATE(A1:A3,B1:B3)+0', styleId: sid })
		sheet.cells.set(0, 9, {
			value: EMPTY,
			formula: 'HOUR(C1:C3)*10000+MINUTE(C1:C3)*100+SECOND(C1:C3)',
			styleId: sid,
		})
		sheet.cells.set(0, 10, { value: EMPTY, formula: 'DATEVALUE(D1:D3)+0', styleId: sid })
		sheet.cells.set(0, 11, { value: EMPTY, formula: 'TIMEVALUE(E1:E3)+0', styleId: sid })
		sheet.cells.set(0, 12, {
			value: EMPTY,
			formula: 'DATE(YEAR(A1:A3),MONTH(A1:A3),DAY(A1:A3))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 13, {
			value: EMPTY,
			formula: 'TIME(HOUR(C1:C3),MINUTE(C1:C3),SECOND(C1:C3))+0',
			styleId: sid,
		})
		sheet.cells.set(0, 14, { value: EMPTY, formula: 'DAYS(A1:A3,DATE(2024,1,1))+0', styleId: sid })
		sheet.cells.set(0, 15, { value: EMPTY, formula: 'WEEKDAY(A1:A3,2)+0', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(2024))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(2024))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(2025))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(131))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(229))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(1215))
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(dateToSerial(2024, 2, 29)))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(dateToSerial(2024, 1, 31)))
		expect(sheet.cells.get(2, 7)?.value).toEqual(numberValue(dateToSerial(2026, 2, 28)))
		expect(sheet.cells.get(0, 8)?.value).toEqual(numberValue(dateToSerial(2024, 2, 29)))
		expect(sheet.cells.get(1, 8)?.value).toEqual(numberValue(dateToSerial(2024, 1, 29)))
		expect(sheet.cells.get(2, 8)?.value).toEqual(numberValue(dateToSerial(2026, 2, 15)))
		expect(sheet.cells.get(0, 9)?.value).toEqual(numberValue(120000))
		expect(sheet.cells.get(1, 9)?.value).toEqual(numberValue(63015))
		expect(sheet.cells.get(2, 9)?.value).toEqual(numberValue(235959))
		for (let row = 0; row < 3; row++) {
			expect(sheet.cells.get(row, 10)?.value).toEqual(numberValue(dates[row] as number))
			expect(sheet.cells.get(row, 11)?.value).toEqual(numberValue(times[row] as number))
			expect(sheet.cells.get(row, 12)?.value).toEqual(numberValue(dates[row] as number))
			expect(sheet.cells.get(row, 13)?.value).toEqual(numberValue(times[row] as number))
			expect(sheet.cells.get(row, 14)?.value).toEqual(
				numberValue((dates[row] as number) - dateToSerial(2024, 1, 1)),
			)
		}
		expect(sheet.cells.get(0, 15)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 15)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(2, 15)?.value).toEqual(numberValue(1))
	})

	test('date interval scalar functions spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const starts = [dateToSerial(2024, 1, 1), dateToSerial(2024, 2, 29), dateToSerial(2024, 12, 31)]
		const ends = [dateToSerial(2024, 1, 31), dateToSerial(2025, 2, 28), dateToSerial(2025, 12, 31)]
		const units = ['D', 'M', 'Y']
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(starts[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(ends[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: stringValue(units[row] as string),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'DATEDIF(A1:A3,B1:B3,C1:C3)+0', styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'DAYS360(A1:A3,B1:B3)+0', styleId: sid })
		sheet.cells.set(0, 5, { value: EMPTY, formula: 'YEARFRAC(A1:A3,B1:B3,0)+0', styleId: sid })
		sheet.cells.set(0, 6, { value: EMPTY, formula: 'WEEKNUM(A1:A3,2)+0', styleId: sid })
		sheet.cells.set(0, 7, { value: EMPTY, formula: 'ISOWEEKNUM(A1:A3)+0', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(11))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(360))
		expect(sheet.cells.get(2, 4)?.value).toEqual(numberValue(360))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(30 / 360))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(2, 5)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 6)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 6)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(2, 6)?.value).toEqual(numberValue(53))
		expect(sheet.cells.get(0, 7)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 7)?.value).toEqual(numberValue(9))
		expect(sheet.cells.get(2, 7)?.value).toEqual(numberValue(1))
	})

	test('workday scalar functions spill while preserving holiday range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const starts = [dateToSerial(2024, 1, 2), dateToSerial(2024, 1, 10), dateToSerial(2024, 2, 1)]
		const ends = [dateToSerial(2024, 1, 12), dateToSerial(2024, 1, 24), dateToSerial(2024, 2, 16)]
		const offsets = [5, -3, 10]
		const weekendCodes = [1, 11, 2]
		const holidays = [dateToSerial(2024, 1, 15), dateToSerial(2024, 2, 9)]
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 0, {
				value: numberValue(starts[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 1, {
				value: numberValue(ends[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 2, {
				value: numberValue(offsets[row] as number),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 3, {
				value: numberValue(weekendCodes[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		for (let row = 0; row < holidays.length; row++) {
			sheet.cells.set(row, 7, {
				value: numberValue(holidays[row] as number),
				formula: null,
				styleId: sid,
			})
		}
		const arrayFormulas = [
			'NETWORKDAYS(A1:A3,B1:B3,H1:H2)+0',
			'WORKDAY(A1:A3,C1:C3,H1:H2)+0',
			'NETWORKDAYS.INTL(A1:A3,B1:B3,D1:D3,H1:H2)+0',
			'WORKDAY.INTL(A1:A3,C1:C3,D1:D3,H1:H2)+0',
		]
		const scalarFormulas = [
			'NETWORKDAYS(A{r},B{r},H1:H2)+0',
			'WORKDAY(A{r},C{r},H1:H2)+0',
			'NETWORKDAYS.INTL(A{r},B{r},D{r},H1:H2)+0',
			'WORKDAY.INTL(A{r},C{r},D{r},H1:H2)+0',
		]
		for (let col = 0; col < arrayFormulas.length; col++) {
			sheet.cells.set(0, 8 + col, {
				value: EMPTY,
				formula: arrayFormulas[col] as string,
				styleId: sid,
			})
			for (let row = 0; row < 3; row++) {
				sheet.cells.set(row, 13 + col, {
					value: EMPTY,
					formula: (scalarFormulas[col] as string).replaceAll('{r}', String(row + 1)),
					styleId: sid,
				})
			}
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let row = 0; row < 3; row++) {
			for (let col = 0; col < arrayFormulas.length; col++) {
				expect(sheet.cells.get(row, 8 + col)?.value).toEqual(sheet.cells.get(row, 13 + col)?.value)
			}
		}
	})

	test('date and workday scalar functions implicitly intersect date arguments and preserve holiday ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const years = [2024, 2025, 2026]
		const months = [1, 2, 3]
		const days = [8, 10, 12]
		const starts = [dateToSerial(2024, 1, 8), dateToSerial(2024, 1, 10), dateToSerial(2024, 2, 1)]
		const ends = [dateToSerial(2024, 1, 12), dateToSerial(2024, 1, 16), dateToSerial(2024, 2, 16)]
		const offsets = [1, 3, 5]
		const weekendCodes = [1, 1, 1]
		const dateText = ['2024-01-08', '2024-01-10', '2024-02-01']
		const timeText = ['08:15:30', '09:45:15', '17:05:10']
		for (let row = 0; row < 3; row++) {
			for (const [col, value] of [
				[0, years[row]],
				[1, months[row]],
				[2, days[row]],
				[3, starts[row]],
				[4, ends[row]],
				[5, offsets[row]],
				[6, weekendCodes[row]],
			] as const) {
				sheet.cells.set(row, col, {
					value: numberValue(value as number),
					formula: null,
					styleId: sid,
				})
			}
			sheet.cells.set(row, 7, {
				value: stringValue(dateText[row] as string),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(row, 8, {
				value: stringValue(timeText[row] as string),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 9, {
			value: numberValue(dateToSerial(2024, 1, 15)),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(1, 9, {
			value: numberValue(dateToSerial(2024, 2, 9)),
			formula: null,
			styleId: sid,
		})
		const rangeFormulas = [
			'DATE(A1:A3,B1:B3,C1:C3)',
			'DATEVALUE(H1:H3)',
			'YEAR(D1:D3)',
			'MONTH(D1:D3)',
			'DAY(D1:D3)',
			'EDATE(D1:D3,F1:F3)',
			'EOMONTH(D1:D3,F1:F3)',
			'DAYS(E1:E3,D1:D3)',
			'HOUR(I1:I3)',
			'MINUTE(I1:I3)',
			'SECOND(I1:I3)',
			'TIME(A1:A3,B1:B3,C1:C3)',
			'TIMEVALUE(I1:I3)',
			'WEEKDAY(D1:D3,2)',
			'NETWORKDAYS(D1:D3,E1:E3,J1:J2)',
			'WORKDAY(D1:D3,F1:F3,J1:J2)',
			'NETWORKDAYS.INTL(D1:D3,E1:E3,G1:G3,J1:J2)',
			'WORKDAY.INTL(D1:D3,F1:F3,G1:G3,J1:J2)',
		]
		const scalarFormulas = [
			'DATE(A2,B2,C2)',
			'DATEVALUE(H2)',
			'YEAR(D2)',
			'MONTH(D2)',
			'DAY(D2)',
			'EDATE(D2,F2)',
			'EOMONTH(D2,F2)',
			'DAYS(E2,D2)',
			'HOUR(I2)',
			'MINUTE(I2)',
			'SECOND(I2)',
			'TIME(A2,B2,C2)',
			'TIMEVALUE(I2)',
			'WEEKDAY(D2,2)',
			'NETWORKDAYS(D2,E2,J1:J2)',
			'WORKDAY(D2,F2,J1:J2)',
			'NETWORKDAYS.INTL(D2,E2,G2,J1:J2)',
			'WORKDAY.INTL(D2,F2,G2,J1:J2)',
		]
		for (let col = 0; col < rangeFormulas.length; col++) {
			sheet.cells.set(1, 10 + col, {
				value: EMPTY,
				formula: rangeFormulas[col] as string,
				styleId: sid,
			})
			sheet.cells.set(1, 30 + col, {
				value: EMPTY,
				formula: scalarFormulas[col] as string,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		for (let col = 0; col < rangeFormulas.length; col++) {
			expect(sheet.cells.get(1, 10 + col)?.value).toEqual(sheet.cells.get(1, 30 + col)?.value)
		}
	})

	test('RAND is deterministic per cell and seed', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'RAND()', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'RAND()', styleId: sid })

		const ctx = makeCtx({ randomSeed: 7 })
		recalculate(wb, ctx)
		const a1 = sheet.cells.get(0, 0)?.value
		const a2 = sheet.cells.get(1, 0)?.value
		recalculate(wb, ctx)
		expect(sheet.cells.get(0, 0)?.value).toEqual(a1)
		expect(sheet.cells.get(1, 0)?.value).toEqual(a2)
		expect(a1).not.toEqual(a2)
	})

	test('wrong function arity returns #VALUE!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'ABS()', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'ABS(1,2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('exponentiation precedence matches spreadsheet semantics', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '2^3^2', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: '-2^2', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: '-(2^2)', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: '(-2)^2', styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: '2^-2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(512))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(-4))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(0.25))
	})

	test('invalid real exponentiation returns #NUM! instead of NaN', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '(-1)^0.5', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: '0^-1', styleId: sid })

		recalculate(wb, makeCtx())

		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#NUM!'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('INDIRECT resolves A1-style ranges inside aggregate functions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUM(INDIRECT("A1:A2"))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(5))
	})

	test('dirty recalc re-evaluates INDIRECT formulas through the volatile path', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'INDIRECT("A1")*2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(4))

		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
	})

	test('INDIRECT resolves R1C1-style references when requested', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(11), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'INDIRECT("R[1]C[-1]",FALSE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(11))
	})

	test('OFFSET returns shifted ranges to aggregators', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1,1,0,2,1))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(11))
	})

	test('OFFSET preserves error identity from offset arguments', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'OFFSET(A1,#N/A,0)', styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'OFFSET(A1,0,#DIV/0!)', styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'OFFSET(A1,0,0,#NUM!,1)', styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'OFFSET(A1,0,0,1,#NAME?)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#N/A'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#DIV/0!'))
		expect(sheet.cells.get(0, 3)?.value).toEqual(errorValue('#NUM!'))
		expect(sheet.cells.get(0, 4)?.value).toEqual(errorValue('#NAME?'))
	})

	test('dirty recalc re-evaluates OFFSET formulas through the volatile path', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1,1,0,1,1))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(5))

		sheet.cells.set(1, 0, { value: numberValue(9), formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A2'] })
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(9))
	})

	test('INDIRECT resolves cross-sheet references', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const sheet2 = wb.addSheet('Sheet2')
		sheet2.cells.set(0, 0, { value: numberValue(88), formula: null, styleId: sid })
		const sheet1 = wb.sheets[0]
		if (!sheet1) throw new Error('missing sheet')
		sheet1.cells.set(0, 0, { value: EMPTY, formula: 'INDIRECT("Sheet2!A1")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet1.cells.get(0, 0)?.value).toEqual(numberValue(88))
	})

	test('INDIRECT returns #REF! for nonexistent sheet', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'INDIRECT("NoSheet!A1")', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('INDIRECT resolves named ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(200), formula: null, styleId: sid })
		wb.definedNames.set('MyRange', 'Sheet1!A1:A2')
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(INDIRECT("MyRange"))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(300))
	})

	test('top-level INDIRECT intersects named ranges at the current row', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('aaa'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('bbb'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('ccc'), formula: null, styleId: sid })
		wb.definedNames.set('Names', 'Sheet1!A1:A3')
		for (let row = 0; row < 3; row++) {
			sheet.cells.set(row, 1, { value: EMPTY, formula: 'INDIRECT("Names")', styleId: sid })
		}

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('aaa'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('bbb'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(stringValue('ccc'))
	})

	test('OFFSET with height and width returns a 2x2 range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'SUM(OFFSET(A1,0,0,2,2))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(10))
	})

	test('OFFSET defaults height/width to the base reference dimensions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1:A2,0,0))', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1:A2,0,0,))', styleId: sid })
		sheet.cells.set(2, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1:A2,0,0,1,))', styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1:A2,0,0,,))', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'SUM(OFFSET(A1:A2,0,0,0,1))', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(4, 1)?.value).toEqual(errorValue('#REF!'))
	})

	test('OFFSET scalar arguments implicitly intersect range operands', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(6, 4, { value: stringValue(''), formula: '""', styleId: sid })
		sheet.cells.set(6, 5, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(6, 6, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(6, 7, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(6, 8, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(7, 6, { value: numberValue(1.001), formula: null, styleId: sid })
		sheet.cells.set(1007, 6, {
			value: EMPTY,
			formula: 'OFFSET(F7,E7:I7,E7:I7)',
			styleId: sid,
		})
		sheet.cells.set(1008, 11, {
			value: EMPTY,
			formula: 'OFFSET(F7,E7:I7,E7:I7)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1007, 6)?.value).toEqual(numberValue(1.001))
		expect(sheet.cells.get(1008, 11)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('SEQUENCE spills vertically into neighboring cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))
	})

	test('basic spill: SEQUENCE(3,1) spills to 3 cells vertically', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))
	})

	test('blocked spill returns #SPILL! in the anchor cell', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			blockingRefs: ['A2'],
		})
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('blocker'))
	})

	test('legacy array formulas write into their imported fixed footprint', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const formulaInfo = { kind: 'array' as const, ref: 'A1:B2' }
		sheet.cells.set(0, 0, {
			value: numberValue(99),
			formula: 'SEQUENCE(2,2)',
			styleId: sid,
			formulaInfo,
		})
		sheet.cells.set(0, 1, {
			value: stringValue('cached member'),
			formula: null,
			styleId: sid,
			formulaInfo,
		})
		sheet.cells.set(1, 0, { value: EMPTY, formula: null, styleId: sid, formulaInfo })
		sheet.cells.set(1, 1, { value: EMPTY, formula: null, styleId: sid, formulaInfo })

		recalculate(wb, makeCtx())

		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual(formulaInfo)
		expect(sheet.cells.get(0, 1)?.formulaInfo).toEqual(formulaInfo)
	})

	test('legacy single-cell matrix formulas keep top-left scalar without array metadata', () => {
		const wb = createWorkbook()
		wb.sourceArchiveBytes = new Uint8Array()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'MINVERSE(A1:B2)', styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'MMULT(A1:B2,A1:B2)', styleId: sid })
		sheet.cells.set(3, 2, { value: EMPTY, formula: 'FREQUENCY(A1:B2,{2,4})', styleId: sid })
		sheet.cells.set(3, 3, { value: EMPTY, formula: 'GROWTH(A1:B2)', styleId: sid })
		sheet.cells.set(4, 0, { value: stringValue('blocker'), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(3, 0)?.value.kind).toBe('number')
		expect(sheet.cells.get(3, 0)?.value.value).toBeCloseTo(-2)
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(3, 3)?.value.kind).toBe('number')
		expect(sheet.cells.get(4, 0)?.value).toEqual(stringValue('blocker'))
	})

	test('imported legacy single-cell INDEX array result returns #VALUE! without spill metadata', () => {
		const wb = createWorkbook()
		wb.sourceArchiveBytes = new Uint8Array()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(13, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(13, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(14, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(14, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(6, 3, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(735, 4, {
			value: errorValue('#VALUE!'),
			formula: 'INDEX(B14:C15,D7,2)',
			styleId: sid,
		})
		sheet.cells.set(736, 4, { value: errorValue('#VALUE!'), formula: null, styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(735, 4)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(736, 4)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('spill conflict: non-empty cell blocks spill range produces #SPILL!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(4)', styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(99), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A4',
			blockingRefs: ['A4'],
		})
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(99))
	})

	test('dirty recalc expands spill once blocker is removed', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))

		sheet.cells.delete(1, 0)
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A2'] })

		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			isAnchor: true,
		})
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))
	})

	test('dirty recalc clears stale spill cells when a spill shrinks', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 4; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
			sheet.cells.set(r, 1, {
				value: { kind: 'boolean', value: true },
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A4,B1:B4,0)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(4))

		sheet.cells.set(2, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })

		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!B3:B4'] })
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 2)).toBeUndefined()
		expect(sheet.cells.get(3, 2)).toBeUndefined()
	})

	test('rebuilds imported stale spill metadata when recalculated dimensions shrink', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const anchorInfo = {
			kind: 'spill' as const,
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A4',
			isAnchor: true,
		}
		const memberInfo = {
			kind: 'spill' as const,
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A4',
			isAnchor: false,
		}
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(2)',
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
		sheet.cells.set(3, 0, {
			value: numberValue(4),
			formula: null,
			styleId: sid,
			formulaInfo: memberInfo,
		})

		recalculate(wb, makeCtx())

		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A2',
			isAnchor: true,
		})
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'spill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A2',
			isAnchor: false,
		})
		expect(sheet.cells.get(2, 0)).toBeUndefined()
		expect(sheet.cells.get(3, 0)).toBeUndefined()
	})

	test('spill shrink/grow: dynamic array size change updates spill range correctly', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 5; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
			sheet.cells.set(r, 1, {
				value: { kind: 'boolean', value: r < 2 },
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A5,B1:B5)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 2)).toBeUndefined()

		sheet.cells.set(2, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })

		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!B3:B4'] })
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(4, 2)).toBeUndefined()
	})

	test('dirty recalc skips unchanged spill outputs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'SEQUENCE(3,1,A1-A1+1)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(3))

		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		const result = recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })

		expect(result.changed).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(3))
	})

	test('spill operator resolves a spilled range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1#)', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'SUM(A2#)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(1, 1)?.value).toEqual(errorValue('#REF!'))
	})

	test('dirty recalc updates same-sheet spill dependents when a spill grows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 4; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
			sheet.cells.set(r, 1, {
				value: { kind: 'boolean', value: r < 2 },
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A4,B1:B4)', styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'SUM(C1#)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(3))

		sheet.cells.set(2, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!B3:B4'] })

		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
	})

	test('dirty recalc updates cross-sheet spill dependents when a spill shrinks', () => {
		const wb = createWorkbook()
		const source = wb.addSheet('Source')
		const summary = wb.addSheet('Summary')
		for (let r = 0; r < 4; r++) {
			source.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
			source.cells.set(r, 1, {
				value: { kind: 'boolean', value: true },
				formula: null,
				styleId: sid,
			})
		}
		source.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A4,B1:B4)', styleId: sid })
		summary.cells.set(0, 0, { value: EMPTY, formula: 'SUM(Source!C1#)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(summary.cells.get(0, 0)?.value).toEqual(numberValue(10))

		source.cells.set(2, 1, {
			value: { kind: 'boolean', value: false },
			formula: null,
			styleId: sid,
		})
		source.cells.set(3, 1, {
			value: { kind: 'boolean', value: false },
			formula: null,
			styleId: sid,
		})
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Source!B3:B4'] })

		expect(source.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(source.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(source.cells.get(2, 2)).toBeUndefined()
		expect(summary.cells.get(0, 0)?.value).toEqual(numberValue(3))
	})

	test('dynamic array functions spill sorted results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SORT(A1:A3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(30))
	})

	test('spill collision: dynamic array blocked by occupied cell returns #SPILL!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(5)', styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('blocker'), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))
	})

	test('spill collision: merged cells block dynamic array output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.merges.push(parseRange('A2:B2'))
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))
		const info = sheet.cells.get(0, 0)?.formulaInfo
		expect(info).toMatchObject({
			kind: 'blockedSpill',
			ref: 'A1:A3',
			blockingRefs: ['A2'],
		})
		expect(sheet.cells.get(1, 0)).toBeUndefined()
	})

	test('spill collision: table ranges block dynamic array output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'Table1',
			sheetId: sheet.id,
			ref: parseRange('A2:A3'),
			columns: [{ id: 1, name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#SPILL!'))
		const info = sheet.cells.get(0, 0)?.formulaInfo
		expect(info).toMatchObject({
			kind: 'blockedSpill',
			ref: 'A1:A3',
			blockingRefs: ['A2', 'A3'],
		})
		expect(sheet.cells.get(1, 0)).toBeUndefined()
		expect(sheet.cells.get(2, 0)).toBeUndefined()
	})

	test('spill collision: worksheet row edge blocks dynamic array output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1_048_575, 0, { value: EMPTY, formula: 'SEQUENCE(2)', styleId: sid })

		recalculate(wb, makeCtx())

		expect(sheet.cells.get(1_048_575, 0)?.value).toEqual(errorValue('#SPILL!'))
		expect(sheet.cells.get(1_048_575, 0)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1048576',
			ref: 'A1048576:A1048577',
			reason: 'sheet-edge',
			blockingRefs: [],
		})
		expect(sheet.cells.get(1_048_576, 0)).toBeUndefined()
	})

	test('spill collision: worksheet column edge blocks dynamic array output', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 16_383, { value: EMPTY, formula: 'SEQUENCE(1,2)', styleId: sid })

		recalculate(wb, makeCtx())

		expect(sheet.cells.get(0, 16_383)?.value).toEqual(errorValue('#SPILL!'))
		expect(sheet.cells.get(0, 16_383)?.formulaInfo).toEqual({
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!XFD1',
			ref: 'XFD1:XFE1',
			reason: 'sheet-edge',
			blockingRefs: [],
		})
		expect(sheet.cells.get(0, 16_384)).toBeUndefined()
	})

	test('spill cell clearing: when spill anchor is deleted, all spill cells are cleared', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))

		sheet.cells.delete(0, 0)
		recalculate(wb, makeCtx())

		expect(sheet.cells.get(0, 0)).toBeUndefined()
		expect(sheet.cells.get(1, 0)).toBeUndefined()
		expect(sheet.cells.get(2, 0)).toBeUndefined()
	})

	test('dirty recalc clears stale spill footprint when spill anchor is deleted', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1#)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))

		sheet.cells.delete(0, 0)
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })

		expect(sheet.cells.get(0, 0)).toBeUndefined()
		expect(sheet.cells.get(1, 0)).toBeUndefined()
		expect(sheet.cells.get(2, 0)).toBeUndefined()
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#REF!'))
	})

	test('dirty recalc clears stale spill cells when spill anchor becomes a value', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(3))

		sheet.cells.set(0, 0, { value: numberValue(99), formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!A1'] })

		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(99))
		expect(sheet.cells.get(0, 0)?.formulaInfo).toBeUndefined()
		expect(sheet.cells.get(1, 0)).toBeUndefined()
		expect(sheet.cells.get(2, 0)).toBeUndefined()
	})

	test('spill resize: FILTER result shrinks on recalc when fewer rows match', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 4; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
			sheet.cells.set(r, 1, {
				value: { kind: 'boolean', value: r < 2 },
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A4,B1:B4)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 2)?.value).toBeUndefined()
	})

	test('cross-sheet spill reference: formula on Sheet2 references Sheet1 spill anchor', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.cells.set(0, 0, { value: EMPTY, formula: 'SEQUENCE(3)', styleId: sid })
		s2.cells.set(0, 0, { value: EMPTY, formula: 'SUM(Sheet1!A1#)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(s2.cells.get(0, 0)?.value).toEqual(numberValue(6))
	})

	test('nested dynamic arrays: SORT(FILTER(...)) spills sorted filtered results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })
		sheet.cells.set(3, 1, { value: { kind: 'boolean', value: true }, formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SORT(FILTER(A1:A4,B1:B4))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(30))
	})

	test('FILTER with no matches returns #CALC! when no fallback', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: { kind: 'boolean', value: false }, formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'FILTER(A1:A2,B1:B2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#CALC!'))
	})

	test('TRANSPOSE spills values across columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'TRANSPOSE(A1:A2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(5))
	})

	test('TAKE spills the requested leading rows', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'TAKE(A1:A3,2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
	})

	test('DROP spills the remaining rows after dropping the requested prefix', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'DROP(A1:A3,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(3))
	})

	test('CHOOSECOLS can spill multiple selected columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'CHOOSECOLS(A1:C1,1,3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(3))
	})

	test('HSTACK spills joined ranges across columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'HSTACK(A1:A2,B1:B2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(20))
	})

	test('SORTBY supports multiple sort keys', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('a'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'SORTBY(A1:B3,A1:A3,1,B1:B3,-1)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('a'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue('b'))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 4)?.value).toEqual(stringValue('a'))
	})

	test('UNIQUE can compare columns when by_col is true', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'UNIQUE(A1:C2,TRUE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(1, 4)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 5)?.value).toEqual(numberValue(4))
	})

	test('TOCOL supports ignore blanks and scan-by-column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'TOCOL(A1:B2,1,TRUE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 3)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 3)?.value).toEqual(numberValue(3))
	})

	test('TOCOL returns #CALC! when ignore rules remove every value', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'TOCOL(A1:B2,1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#CALC!'))
	})

	test('TOROW supports ignore errors and scan-by-column', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: errorValue('#N/A'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'TOROW(A1:B2,2,TRUE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(3))
	})

	test('FILTER returns #VALUE! for include shape mismatch and propagates include errors', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'FILTER(A1:B2,{TRUE})', styleId: sid })
		sheet.cells.set(1, 3, { value: EMPTY, formula: 'FILTER(A1:A2,{#N/A;TRUE})', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(errorValue('#N/A'))
	})

	test('TAKE and DROP return #CALC! for zero extents', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'TAKE(A1:A2,0)', styleId: sid })
		sheet.cells.set(1, 2, { value: EMPTY, formula: 'DROP(A1:A2,0)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#CALC!'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(errorValue('#CALC!'))
	})

	test('CHOOSECOLS and CHOOSEROWS support negative indices from the end', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(6), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'CHOOSECOLS(A1:C2,-1,-2)', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'CHOOSEROWS(A1:C2,-1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(6))
	})

	test('XMATCH spills results for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'XMATCH({20;30},A1:A3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(3))
	})

	test('MATCH spills results for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('North'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('South'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: stringValue('West'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'MATCH({"South";"West"},A1:A3,0)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(3))
	})

	test('XLOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'XLOOKUP({2;3},A1:A3,B1:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Three'))
	})

	test('VLOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'VLOOKUP({2;3},A1:B3,2,FALSE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Three'))
	})

	test('HLOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: EMPTY, formula: 'HLOOKUP({2,3},A1:C2,2,FALSE)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(0, 5)?.value).toEqual(stringValue('Three'))
	})

	test('LOOKUP spills scalar-return matches for array lookup values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('One'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('Two'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('Three'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: EMPTY, formula: 'LOOKUP({2;3},A1:A3,B1:B3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('Two'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('Three'))
	})

	test('CHOOSE spills values selected by an array index', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'CHOOSE({1,2,1},"North","South")',
			styleId: sid,
		})
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'CHOOSE({1;3},"first",1/0,"third")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('North'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('South'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('North'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('first'))
		expect(sheet.cells.get(2, 0)?.value).toEqual(stringValue('third'))
	})

	test('SWITCH spills results for array expressions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: stringValue('b'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SWITCH(A1:A4,1,"one",2,"two","B","bee","other")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('one'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('two'))
		expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('other'))
		expect(sheet.cells.get(3, 2)?.value).toEqual(stringValue('bee'))
	})

	test('SWITCH array expressions preserve lazy branch evaluation', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SWITCH(A1:A2,1,"ok",2,1/0,"fallback")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('ok'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('ok'))
	})

	test('IFS spills results for array conditions', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(95), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(84), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(72), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: numberValue(58), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'IFS(A1:A4>=90,"A",A1:A4>=80,"B",A1:A4>=70,"C",TRUE,"F")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 2)?.value).toEqual(stringValue('A'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(stringValue('B'))
		expect(sheet.cells.get(2, 2)?.value).toEqual(stringValue('C'))
		expect(sheet.cells.get(3, 2)?.value).toEqual(stringValue('F'))
	})

	test('IFS array conditions preserve selected array results and lazy branch evaluation', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('x'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: stringValue('y'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('first'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('second'), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'IFS(A1:A2="x",B1:B2,A1:A2="z",1/0,TRUE,"fallback")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('first'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('fallback'))
	})

	test('TEXTBEFORE and TEXTAFTER support modern text slicing formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTBEFORE("alpha-beta-gamma","-",2)',
			styleId: sid,
		})
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'TEXTAFTER("alpha-beta-gamma","-",2)',
			styleId: sid,
		})
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('alpha-beta'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('gamma'))
	})

	test('TEXTBEFORE and TEXTAFTER support case-insensitive matching and if_not_found', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTBEFORE("Alpha-beta","BETA",1,1,0,"missing")',
			styleId: sid,
		})
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'TEXTAFTER("Alpha","-")', styleId: sid })
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('Alpha-'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#N/A'))
	})

	test('TEXTBEFORE and TEXTAFTER spill over range operands in array formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: stringValue('alpha-beta-gamma'),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(1, 0, {
			value: stringValue('north/east/south'),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(2, 0, { value: stringValue('one|two'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('-'), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: stringValue('/'), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: stringValue('|'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 2, { value: numberValue(-1), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'TEXTBEFORE(A1:A3,B1:B3,C1:C3)&""',
			styleId: sid,
		})
		sheet.cells.set(0, 4, {
			value: EMPTY,
			formula: 'TEXTAFTER(A1:A3,B1:B3,C1:C3)&""',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(stringValue('alpha-beta'))
		expect(sheet.cells.get(1, 3)?.value).toEqual(stringValue('north/east'))
		expect(sheet.cells.get(2, 3)?.value).toEqual(stringValue('one'))
		expect(sheet.cells.get(0, 4)?.value).toEqual(stringValue('gamma'))
		expect(sheet.cells.get(1, 4)?.value).toEqual(stringValue('south'))
		expect(sheet.cells.get(2, 4)?.value).toEqual(stringValue('two'))
	})

	test('TEXTSPLIT spills rows and columns from delimited text', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTSPLIT("a,b;c,d",",",";")',
			styleId: sid,
		})
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('a'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('b'))
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('c'))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('d'))
	})

	test('TEXTSPLIT supports ignore_empty and pad_with', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'TEXTSPLIT("a,,c",",",,TRUE,0,"pad")',
			styleId: sid,
		})
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(stringValue('a'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('c'))
		expect(sheet.cells.get(0, 2)).toBeUndefined()
	})

	test('external workbook references without cached values evaluate to #REF!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '[Book.xlsx]Sheet1!A1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('external workbook references preserve imported cached values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: numberValue(42),
			formula: '[Book.xlsx]Sheet1!A1',
			styleId: sid,
		})
		sheet.cells.set(0, 1, {
			value: stringValue('cached'),
			formula: 'IF(TRUE,[Book.xlsx]Sheet1!B1,"fallback")',
			styleId: sid,
		})
		sheet.cells.set(0, 2, {
			value: errorValue('#N/A'),
			formula: '[Book.xlsx]Sheet1!C1',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())
		expect(result.changed).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(42))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('cached'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#N/A'))
	})

	test('path-qualified external workbook references preserve imported cached values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: numberValue(42),
			formula: "'C:/tmp/[Book.xlsx]Sheet1'!A1",
			styleId: sid,
		})
		sheet.cells.set(0, 1, {
			value: stringValue('cached'),
			formula: 'IF(TRUE,\'C:/tmp/[Book.xlsx]Sheet1\'!B1,"fallback")',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())
		expect(result.changed).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(42))
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('cached'))
	})

	test('3D sheet-span references aggregate across contiguous sheets', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const s3 = wb.addSheet('Sheet3')
		const calc = wb.addSheet('Calc')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		s3.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'SUM(Sheet1:Sheet3!A1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(calc.cells.get(0, 0)?.value).toEqual(numberValue(6))
	})

	test('bare 3D sheet-span references return #VALUE! in scalar contexts', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const calc = wb.addSheet('Calc')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'Sheet1:Sheet2!A1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(calc.cells.get(0, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('error propagation through formula chain', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1/B1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#DIV/0!'))
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('top-level references to blank cells calculate as zero', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: '"x"&A1', styleId: sid })
		sheet.cells.set(2, 1, { value: EMPTY, formula: 'OFFSET(A1,0,0)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(0))
		expect(sheet.cells.get(1, 1)?.value).toEqual(stringValue('x'))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(0))
	})

	test('scalar numeric coercion distinguishes empty text from blank cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue(''), formula: '""', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'COS(A1)', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'B1+1', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(2, 0)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(1))
	})

	test('circular reference produces error', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
	})

	test('range self-reference produces circular reference error', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SUM(A1:A1)', styleId: sid })

		const result = recalculate(wb, makeCtx())
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('defined name range self-reference produces circular reference error', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		wb.definedNames.set('TotalRange', 'Sheet1!A1:A1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'SUM(TotalRange)', styleId: sid })

		const result = recalculate(wb, makeCtx())
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
	})

	test('circular references preserve imported cached values when present', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: 'B1', styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: 'A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(20))
	})

	test('recalc preserves date-typed formula results for criteria boundary precision', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const serial = 44227.99998842592
		sheet.cells.set(0, 0, { value: dateValue(serial), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: dateValue(serial), formula: 'A1+0', styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 4, { value: dateValue(serial), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'SUMIFS(D1:D1,E1:E1,"<="&B1)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(dateValue(serial))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(5))
	})

	test('conditional aggregates treat date criteria as date serial matches', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: dateValue(45000), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: dateValue(45000), formula: null, styleId: sid })
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'SUMIFS(A1:A1,B1:B1,C1)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(10))
	})

	test('structured references resolve escaped special characters in column names', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'BillingData',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } },
			columns: [{ id: 1, name: 'Check#' }],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.cells.set(0, 0, { value: stringValue('Check#'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(100), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(200), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: "SUBTOTAL(103,BillingData[Check'#])",
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
	})

	test('SUBTOTAL table column references include manually hidden rows for 1-11 variants', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'SalesTable')
		populateRegionSalesRows(sheet)
		sheet.rowDefs.set(2, { hidden: true })
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'SUBTOTAL(9,SalesTable[Sales])',
			styleId: sid,
		})
		sheet.cells.set(4, 1, {
			value: EMPTY,
			formula: 'SUBTOTAL(109,SalesTable[Sales])',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(90))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(60))
	})

	test('SUBTOTAL table column references exclude filtered rows for all variants', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'FilteredSales', {
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'filters', values: ['20', '40'] }],
		})
		populateRegionSalesRows(sheet)
		sheet.rowDefs.set(2, { hidden: true })
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'SUBTOTAL(9,FilteredSales[Sales])',
			styleId: sid,
		})
		sheet.cells.set(4, 1, {
			value: EMPTY,
			formula: 'SUBTOTAL(109,FilteredSales[Sales])',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(60))
	})

	test('SUBTOTAL table column references evaluate explicit filter criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'FilteredSales', {
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'filters', values: ['20', '40'] }],
		})
		populateRegionSalesRows(sheet)
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'SUBTOTAL(9,FilteredSales[Sales])',
			styleId: sid,
		})
		sheet.cells.set(4, 1, {
			value: EMPTY,
			formula: 'SUBTOTAL(109,FilteredSales[Sales])',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(60))
	})

	test('SUBTOTAL table column references prefer saved filter visibility over stale criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'FilteredSales', {
			ref: 'A1:B4',
			columns: [{ colId: 0, kind: 'filters', values: ['West'] }],
		})
		populateRegionSalesRows(sheet)
		sheet.rowDefs.set(2, { hidden: true })
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'SUBTOTAL(9,FilteredSales[Sales])',
			styleId: sid,
		})
		sheet.cells.set(4, 1, {
			value: EMPTY,
			formula: 'SUBTOTAL(109,FilteredSales[Sales])',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(60))
	})

	test('SUBTOTAL skips hidden rows in referenced ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.rowDefs.set(1, { hidden: true })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUBTOTAL(109,A1:A3)', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'SUBTOTAL(102,A1:A3)', styleId: sid })
		sheet.cells.set(2, 1, { value: EMPTY, formula: 'SUBTOTAL(105,A1:A3)', styleId: sid })
		sheet.cells.set(3, 1, { value: EMPTY, formula: 'SUBTOTAL(9,A1:A3)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(40))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(60))
	})

	test('AGGREGATE honors hidden-row options in referenced ranges', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(30), formula: null, styleId: sid })
		sheet.rowDefs.set(1, { hidden: true })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,A1:A3)', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'AGGREGATE(9,5,A1:A3)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(40))
	})

	test('AGGREGATE combines error-ignore and hidden-row options', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: errorValue('#DIV/0!'), formula: null, styleId: sid })
		sheet.rowDefs.set(1, { hidden: true })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'AGGREGATE(9,6,A1:A3)', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'AGGREGATE(9,7,A1:A3)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(10))
	})

	test('AGGREGATE table column references exclude filtered rows for all options', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'FilteredSales', {
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'filters', values: ['20', '40'] }],
		})
		populateRegionSalesRows(sheet)
		sheet.rowDefs.set(2, { hidden: true })
		sheet.cells.set(4, 0, {
			value: EMPTY,
			formula: 'AGGREGATE(9,4,FilteredSales[Sales])',
			styleId: sid,
		})
		sheet.cells.set(4, 1, {
			value: EMPTY,
			formula: 'AGGREGATE(9,5,FilteredSales[Sales])',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(60))
	})

	test('AGGREGATE and sheet autoFilter ranges evaluate explicit filter criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B4',
			columns: [{ colId: 0, kind: 'filters', values: ['West', 'North'] }],
		}
		populateRegionSalesRows(sheet)
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(60))
	})

	test('SUBTOTAL and AGGREGATE evaluate custom filter comparisons', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B4',
			columns: [
				{
					colId: 1,
					kind: 'customFilters',
					customFilters: [{ operator: 'greaterThan', val: '25' }],
				},
			],
		}
		populateRegionSalesRows(sheet)
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(70))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(70))
	})

	test('SUBTOTAL and AGGREGATE evaluate cell fill color filters', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const redFill: CellStyle = {
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFFFD7D7' } },
		}
		const blueFill = wb.styles.register({
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFB3CAC7' } },
		})
		const redFillId = wb.styles.register(redFill)
		wb.differentialStyles.push(redFill)
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [{ colId: 1, kind: 'colorFilter', dxfId: 0 }],
		}
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, [value, styleId]] of [
			[10, redFillId],
			[20, blueFill],
			[30, redFillId],
			[40, sid],
		].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B5)', styleId: sid })
		sheet.cells.set(5, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B5)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(40))
		expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(40))
	})

	test('color filters can target font color and conditional-format colors', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const redFont: CellStyle = { font: { color: { kind: 'rgb', rgb: 'FFFF0000' } } }
		const redFill: CellStyle = {
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
		}
		const redFontId = wb.styles.register(redFont)
		const plainId = wb.styles.register({})
		wb.differentialStyles.push(redFont, redFill)
		sheet.autoFilter = {
			ref: 'A1:C5',
			columns: [
				{ colId: 0, kind: 'colorFilter', dxfId: 0, cellColor: false },
				{ colId: 1, kind: 'colorFilter', dxfId: 1, cellColor: true },
			],
		}
		sheet.conditionalFormats.push({
			sqref: 'B2:B5',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					dxfId: 1,
					priority: 1,
					formulas: ['25'],
					style: redFill,
				},
			],
		})
		sheet.cells.set(0, 0, { value: stringValue('Flag'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Amount'), formula: null, styleId: sid })
		for (const [index, [score, amount, flagStyle]] of [
			[10, 5, redFontId],
			[30, 10, plainId],
			[40, 20, redFontId],
			[50, 30, redFontId],
		].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: flagStyle,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(score), formula: null, styleId: sid })
			sheet.cells.set(index + 1, 2, { value: numberValue(amount), formula: null, styleId: sid })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,C2:C5)', styleId: sid })
		sheet.cells.set(5, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,C2:C5)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(50))
		expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(50))
	})

	test('SUBTOTAL and AGGREGATE evaluate icon filter criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [{ colId: 1, kind: 'iconFilter', iconSet: '3TrafficLights1', iconId: 1 }],
		}
		sheet.conditionalFormats.push({
			sqref: 'B2:B5',
			rules: [
				{
					type: 'iconSet',
					priority: 1,
					formulas: [],
					iconSet: {
						iconSet: '3TrafficLights1',
						cfvo: [
							{ type: 'num', value: '0' },
							{ type: 'num', value: '50' },
							{ type: 'num', value: '80' },
						],
					},
				},
			],
		})
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 40, 70, 90].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B5)', styleId: sid })
		sheet.cells.set(5, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B5)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(70))
		expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(70))
	})

	test('table icon filters respect percent thresholds and reversed icon sets', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'Sales', {
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'iconFilter', iconSet: '3Arrows', iconId: 0 }],
		})
		sheet.conditionalFormats.push({
			sqref: 'B2:B4',
			rules: [
				{
					type: 'iconSet',
					priority: 1,
					formulas: [],
					iconSet: {
						iconSet: '3Arrows',
						reverse: true,
						cfvo: [
							{ type: 'percent', value: '0' },
							{ type: 'percent', value: '50' },
							{ type: 'percent', value: '75' },
						],
					},
				},
			],
		})
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 50, 90].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(90))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(90))
	})

	test('icon filters honor percentile thresholds and exclusive boundaries', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'iconFilter', iconSet: '3Symbols', iconId: 0 }],
		}
		sheet.conditionalFormats.push({
			sqref: 'B2:B4',
			rules: [
				{
					type: 'iconSet',
					priority: 1,
					formulas: [],
					iconSet: {
						iconSet: '3Symbols',
						cfvo: [
							{ type: 'percentile', value: '0' },
							{ type: 'percentile', value: '50', gte: false },
							{ type: 'percentile', value: '80' },
						],
					},
				},
			],
		})
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 55, 80].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(65))
	})

	test('SUBTOTAL and AGGREGATE evaluate top10 filter criteria with ties', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B6',
			columns: [{ colId: 1, kind: 'top10', top: true, val: 2 }],
		}
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 30, 30, 20, 5].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(6, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B6)', styleId: sid })
		sheet.cells.set(6, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B6)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(6, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(6, 1)?.value).toEqual(numberValue(60))
	})

	test('table references evaluate bottom percent top10 filter criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		addSalesTable(sheet, 'FilteredSales', {
			ref: 'A1:B6',
			columns: [{ colId: 1, kind: 'top10', top: false, percent: true, val: 40 }],
		})
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 30, 30, 20, 5].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(6, 0, {
			value: EMPTY,
			formula: 'SUBTOTAL(9,FilteredSales[Sales])',
			styleId: sid,
		})
		sheet.cells.set(6, 1, {
			value: EMPTY,
			formula: 'AGGREGATE(9,4,FilteredSales[Sales])',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(6, 0)?.value).toEqual(numberValue(15))
		expect(sheet.cells.get(6, 1)?.value).toEqual(numberValue(15))
	})

	test('top10 filter criteria honor saved filterVal thresholds', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B4',
			columns: [{ colId: 1, kind: 'top10', top: true, val: 1, filterVal: 25 }],
		}
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 30, 40].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(70))
	})

	test('dynamic filter criteria evaluate above-average thresholds', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [{ colId: 1, kind: 'dynamicFilter', dynamicFilterType: 'aboveAverage' }],
		}
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Sales'), formula: null, styleId: sid })
		for (const [index, value] of [10, 20, 30, 40].entries()) {
			sheet.cells.set(index + 1, 0, {
				value: stringValue(`R${index + 1}`),
				formula: null,
				styleId: sid,
			})
			sheet.cells.set(index + 1, 1, { value: numberValue(value), formula: null, styleId: sid })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B5)', styleId: sid })
		sheet.cells.set(5, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B5)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(70))
		expect(sheet.cells.get(5, 1)?.value).toEqual(numberValue(70))
	})

	test('dynamic date ranges evaluate ISO thresholds', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [
				{
					colId: 0,
					kind: 'dynamicFilter',
					dynamicFilterType: 'thisMonth',
					dynamicFilterValIso: '2026-03-01T00:00:00',
					dynamicFilterMaxValIso: '2026-04-01T00:00:00',
				},
			],
		}
		sheet.cells.set(0, 0, { value: stringValue('Date'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		for (const [index, [serial, amount]] of [
			[dateToSerial(2026, 2, 28), 5],
			[dateToSerial(2026, 3, 1), 10],
			[dateToSerial(2026, 3, 31), 20],
			[dateToSerial(2026, 4, 1), 30],
		].entries()) {
			sheet.cells.set(index + 1, 0, { value: dateValue(serial), formula: null, styleId: sid })
			sheet.cells.set(index + 1, 1, { value: numberValue(amount), formula: null, styleId: sid })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B5)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(30))
	})

	test('dynamic relative date filters use the calculation date', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [{ colId: 0, kind: 'dynamicFilter', dynamicFilterType: 'thisMonth' }],
		}
		sheet.cells.set(0, 0, { value: stringValue('Date'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		for (const [index, [serial, amount]] of [
			[dateToSerial(2026, 4, 30), 5],
			[dateToSerial(2026, 5, 1), 10],
			[dateToSerial(2026, 5, 31), 20],
			[dateToSerial(2026, 6, 1), 30],
		].entries()) {
			sheet.cells.set(index + 1, 0, { value: dateValue(serial), formula: null, styleId: sid })
			sheet.cells.set(index + 1, 1, { value: numberValue(amount), formula: null, styleId: sid })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B5)', styleId: sid })

		const result = recalculate(wb, makeCtx({ today: new Date(2026, 4, 10) }))

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(30))
	})

	test('dynamic month and quarter filters evaluate date buckets', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B5',
			columns: [
				{ colId: 0, kind: 'dynamicFilter', dynamicFilterType: 'M3' },
				{ colId: 1, kind: 'dynamicFilter', dynamicFilterType: 'Q1' },
			],
		}
		sheet.cells.set(0, 0, { value: stringValue('InvoiceDate'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('ShipDate'), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: stringValue('Amount'), formula: null, styleId: sid })
		const rows = [
			[dateToSerial(2026, 3, 5), dateToSerial(2026, 1, 15), 10],
			[dateToSerial(2026, 3, 20), dateToSerial(2026, 3, 31), 20],
			[dateToSerial(2026, 4, 1), dateToSerial(2026, 1, 5), 30],
			[dateToSerial(2026, 3, 12), dateToSerial(2026, 4, 1), 40],
		]
		for (const [index, [invoiceDate, shipDate, amount]] of rows.entries()) {
			sheet.cells.set(index + 1, 0, { value: dateValue(invoiceDate), formula: null, styleId: sid })
			sheet.cells.set(index + 1, 1, { value: dateValue(shipDate), formula: null, styleId: sid })
			sheet.cells.set(index + 1, 2, { value: numberValue(amount), formula: null, styleId: sid })
		}
		sheet.cells.set(5, 0, { value: EMPTY, formula: 'SUBTOTAL(9,C2:C5)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(5, 0)?.value).toEqual(numberValue(30))
	})

	test('SUBTOTAL and AGGREGATE evaluate date group filter criteria', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B4',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					dateGroupItems: [{ year: 2026, month: 3, dateTimeGrouping: 'month' }],
				},
			],
		}
		sheet.cells.set(0, 0, { value: stringValue('Date'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: dateValue(dateToSerial(2026, 3, 1)),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, {
			value: dateValue(dateToSerial(2026, 3, 31)),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(3, 0, {
			value: dateValue(dateToSerial(2026, 4, 1)),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(3, 1, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })
		sheet.cells.set(4, 1, { value: EMPTY, formula: 'AGGREGATE(9,4,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(30))
	})

	test('date group filter criteria respect the calculation date system', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B3',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					dateGroupItems: [{ year: 2026, month: 3, dateTimeGrouping: 'month' }],
				},
			],
		}
		sheet.cells.set(0, 0, { value: stringValue('Date'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: dateValue(dateToSerial(2026, 3, 15, '1904')),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, {
			value: dateValue(dateToSerial(2026, 4, 15, '1904')),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B3)', styleId: sid })

		const result = recalculate(wb, makeCtx({ dateSystem: '1904' }))

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(10))
	})

	test('date group filter criteria evaluate time components', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = {
			ref: 'A1:B4',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					dateGroupItems: [
						{
							year: 2026,
							month: 3,
							day: 15,
							hour: 13,
							minute: 45,
							second: 30,
							dateTimeGrouping: 'second',
						},
					],
				},
			],
		}
		const matchingTime = (13 * 3600 + 45 * 60 + 30) / 86_400
		const adjacentTime = (13 * 3600 + 45 * 60 + 31) / 86_400
		sheet.cells.set(0, 0, { value: stringValue('Timestamp'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: dateValue(dateToSerial(2026, 3, 15) + matchingTime),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(2, 0, {
			value: dateValue(dateToSerial(2026, 3, 15) + adjacentTime),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(3, 0, {
			value: dateValue(dateToSerial(2026, 3, 16) + matchingTime),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(3, 1, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(4, 0, { value: EMPTY, formula: 'SUBTOTAL(9,B2:B4)', styleId: sid })

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(10))
	})

	test('string concatenation formula', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: stringValue(' World'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1&B1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(stringValue('Hello World'))
	})

	test('string concatenation uses Excel-like general number text', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: numberValue(5.2421445218367324e-14),
			formula: null,
			styleId: sid,
		})
		sheet.cells.set(0, 1, { value: EMPTY, formula: '"x"&A1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(stringValue('x5.24214452183673E-14'))
	})

	test('formula parse error is reported and cell displays #VALUE!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '=INVALID(((', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(result.errors.length).toBeGreaterThan(0)
		expect(result.errors[0]?.error.code).toBe('FORMULA_PARSE_ERROR')
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('duration is tracked', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const result = recalculate(wb, makeCtx())
		expect(result.duration).toBeGreaterThanOrEqual(0)
	})

	test('arithmetic operators evaluate correctly', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1-B1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A1*B1', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: 'A1^2', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(7))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(100))
	})
})

describe('iterative calculation', () => {
	test('simple circular ref converges with iterative calc enabled', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '0.5*B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: '0.5*A1+1', styleId: sid })

		const ctx = makeCtx({
			iterativeCalc: {
				enabled: true,
				maxIterations: 100,
				maxChange: 0.001,
			},
		})
		recalculate(wb, ctx)

		const a1 = sheet.cells.get(0, 0)?.value
		const b1 = sheet.cells.get(0, 1)?.value
		expect(a1?.kind).toBe('number')
		expect(b1?.kind).toBe('number')
		if (a1?.kind === 'number' && b1?.kind === 'number') {
			expect(a1.value).toBeGreaterThanOrEqual(0)
			expect(a1.value).toBeLessThanOrEqual(1)
			expect(b1.value).toBeGreaterThanOrEqual(1)
			expect(b1.value).toBeLessThanOrEqual(2)
		}
	})

	test('max iterations respected', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1+1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const ctx = makeCtx({
			iterativeCalc: {
				enabled: true,
				maxIterations: 5,
				maxChange: 0.001,
			},
		})
		recalculate(wb, ctx)

		const a1 = sheet.cells.get(0, 0)?.value
		const b1 = sheet.cells.get(0, 1)?.value
		expect(a1?.kind).toBe('number')
		expect(b1?.kind).toBe('number')
		if (a1?.kind === 'number' && b1?.kind === 'number') {
			expect(a1.value).toBeLessThanOrEqual(6)
			expect(b1.value).toBeLessThanOrEqual(6)
		}
	})

	test('without iterative calc, circular refs produce #REF!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: sid })

		const result = recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#REF!'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#REF!'))
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors.length).toBeGreaterThan(0)
	})

	test('external workbook-qualified 3D ranges do not self-cycle against local cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 1, {
			value: EMPTY,
			formula: 'SUM([1]FY26:FY28!B2:B10)',
			styleId: sid,
		})

		const result = recalculate(wb, makeCtx())
		const circErrors = result.errors.filter((e) => e.error.code === 'CIRCULAR_REF')
		expect(circErrors).toEqual([])
	})
})

describe('shared formula evaluation', () => {
	test('shared formulas A1:A10 = B1*2 shifted per row evaluate correctly', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 10; r++) {
			sheet.cells.set(r, 1, { value: numberValue(r + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		for (let r = 1; r < 10; r++) {
			sheet.cells.set(r, 0, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})
		}

		recalculate(wb, makeCtx())
		for (let r = 0; r < 10; r++) {
			const expected = (r + 1) * 2
			expect(sheet.cells.get(r, 0)?.value).toEqual(numberValue(expected))
		}
	})

	test('recalc after changing source cell updates shared formula results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})
		sheet.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(6))

		sheet.cells.set(0, 1, { value: numberValue(7), formula: null, styleId: sid })
		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(14))
	})

	test('group batch execution of 100 shared formula cells with B+C matches single-cell results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 100; r++) {
			sheet.cells.set(r, 1, { value: numberValue(r + 1), formula: null, styleId: sid })
			sheet.cells.set(r, 2, { value: numberValue((r + 1) * 10), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1+C1',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		for (let r = 1; r < 100; r++) {
			sheet.cells.set(r, 0, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})
		}

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		for (let r = 0; r < 100; r++) {
			const n = r + 1
			expect(sheet.cells.get(r, 0)?.value).toEqual(numberValue(n + n * 10))
		}
	})

	test('dirty recalc propagates correctly through shared formula groups', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 10; r++) {
			sheet.cells.set(r, 1, { value: numberValue(r + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1*3',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		for (let r = 1; r < 10; r++) {
			sheet.cells.set(r, 0, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
			})
		}

		recalculate(wb, makeCtx())
		for (let r = 0; r < 10; r++) {
			expect(sheet.cells.get(r, 0)?.value).toEqual(numberValue((r + 1) * 3))
		}

		sheet.cells.set(4, 1, { value: numberValue(99), formula: null, styleId: sid })
		recalculate(wb, makeCtx(), { dirtyOnly: true, dirtyRefs: ['Sheet1!B5'] })
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(297))
		for (let r = 0; r < 10; r++) {
			if (r === 4) continue
			expect(sheet.cells.get(r, 0)?.value).toEqual(numberValue((r + 1) * 3))
		}
	})

	test('multiple shared formula groups on same sheet evaluate independently', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 5; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'A1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' },
		})
		for (let r = 1; r < 5; r++) {
			sheet.cells.set(r, 1, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' },
			})
		}
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'A1+10',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '1', isMaster: true, masterRef: 'C1' },
		})
		for (let r = 1; r < 5; r++) {
			sheet.cells.set(r, 2, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '1', isMaster: false, masterRef: 'C1' },
			})
		}

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		for (let r = 0; r < 5; r++) {
			expect(sheet.cells.get(r, 1)?.value).toEqual(numberValue((r + 1) * 2))
			expect(sheet.cells.get(r, 2)?.value).toEqual(numberValue(r + 1 + 10))
		}
	})

	test('shared formula group of 100 cells =A<row>*2 evaluates correctly (TEST-6)', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 100; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'A1*2',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' },
		})
		for (let r = 1; r < 100; r++) {
			sheet.cells.set(r, 1, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' },
			})
		}

		recalculate(wb, makeCtx())
		for (let r = 0; r < 100; r++) {
			const expected = (r + 1) * 2
			expect(sheet.cells.get(r, 1)?.value).toEqual(numberValue(expected))
		}
	})

	test('shared formula groups preserve absolute refs while shifting relative refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: sid })
		for (let r = 0; r < 5; r++) {
			sheet.cells.set(r, 1, { value: numberValue(r + 1), formula: null, styleId: sid })
		}
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: '$A$1+B1',
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'C1' },
		})
		for (let r = 1; r < 5; r++) {
			sheet.cells.set(r, 2, {
				value: EMPTY,
				formula: null,
				styleId: sid,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'C1' },
			})
		}

		recalculate(wb, makeCtx())
		for (let r = 0; r < 5; r++) {
			expect(sheet.cells.get(r, 2)?.value).toEqual(numberValue(100 + r + 1))
		}
	})
})

describe('large range correctness', () => {
	test('SUM over 10K cells verifies exact result', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		let expectedSum = 0
		for (let i = 0; i < 10_000; i++) {
			const val = i + 1
			expectedSum += val
			sheet.cells.set(Math.floor(i / 100), i % 100, {
				value: numberValue(val),
				formula: null,
				styleId: sid,
			})
		}
		sheet.cells.set(101, 0, {
			value: EMPTY,
			formula: 'SUM(A1:CV100)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(101, 0)?.value).toEqual(numberValue(expectedSum))
	})

	test('AVERAGE over 10K cells verifies correct average', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		let sum = 0
		for (let i = 0; i < 10_000; i++) {
			const val = i + 1
			sum += val
			sheet.cells.set(Math.floor(i / 100), i % 100, {
				value: numberValue(val),
				formula: null,
				styleId: sid,
			})
		}
		const expectedAvg = sum / 10_000
		sheet.cells.set(101, 0, {
			value: EMPTY,
			formula: 'AVERAGE(A1:CV100)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		const result = sheet.cells.get(101, 0)?.value
		expect(result?.kind).toBe('number')
		if (result?.kind === 'number') {
			expect(Math.abs(result.value - expectedAvg)).toBeLessThan(0.0001)
		}
	})

	test('cross-sheet SUM - 3 sheets, SUM(Sheet1!A1:A100) on Sheet2', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		let expectedSum = 0
		for (let r = 0; r < 100; r++) {
			const val = r + 1
			expectedSum += val
			s1.cells.set(r, 0, { value: numberValue(val), formula: null, styleId: sid })
		}
		s2.cells.set(0, 0, {
			value: EMPTY,
			formula: 'SUM(Sheet1!A1:A100)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(s2.cells.get(0, 0)?.value).toEqual(numberValue(expectedSum))
	})

	test('compiled IF opcode: short-circuit true branch', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IF(A1>5,A1*2,B1*3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(20))
	})

	test('compiled IF opcode: short-circuit false branch', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IF(A1>5,A1*2,B1*3)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(60))
	})

	test('full recalc fast path handles scalar IF with SUM fallback', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const aValues = [1, 0, 2, 1]
		const bValues = [10, 20, 30]
		for (let row = 0; row < aValues.length; row++) {
			sheet.cells.set(row, 0, { value: numberValue(aValues[row]), formula: null, styleId: sid })
			if (row < bValues.length) {
				sheet.cells.set(row, 1, { value: numberValue(bValues[row]), formula: null, styleId: sid })
			}
			sheet.cells.set(row, 2, {
				value: EMPTY,
				formula: `IF(A${row + 1}>0,B${row + 1},SUM(B1:B3))`,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())

		expect(result.errors).toEqual([])
		expect(result.changed).toEqual(['Sheet1!C1', 'Sheet1!C2', 'Sheet1!C3', 'Sheet1!C4'])
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(30))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(0))
	})

	test('compiled IF opcode: missing false branch defaults to FALSE', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IF(A1>0,A1+1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual({ kind: 'boolean', value: false })
	})

	test('compiled IF opcode: error condition propagates', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: '1/0', styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IF(A1>0,10,20)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('compiled IFERROR opcode: non-error passes through', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IFERROR(A1*2+3,-1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(23))
	})

	test('compiled IFERROR opcode: error returns fallback', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IFERROR(1/A1,-1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(-1))
	})

	test('compiled IFNA opcode: non-NA passes through', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'IFNA(A1+10,0)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(15))
	})

	test('compiled CELL_SHEET opcode: cross-sheet cell reference', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: EMPTY, formula: 'Sheet1!A1*2+1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(s2.cells.get(0, 0)?.value).toEqual(numberValue(85))
	})

	test('compiled IF chain: nested IF compiles to bytecode', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(75), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'IF(A1>=90,4,IF(A1>=80,3,IF(A1>=70,2,1)))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
	})

	test('many IF formulas benefit from compiled path', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 200; r++) {
			sheet.cells.set(r, 0, { value: numberValue(r), formula: null, styleId: sid })
			sheet.cells.set(r, 1, {
				value: EMPTY,
				formula: `IF(A${r + 1}>100,A${r + 1}*2,A${r + 1}+10)`,
				styleId: sid,
			})
		}

		const result = recalculate(wb, makeCtx())
		expect(result.errors).toEqual([])
		expect(sheet.cells.get(50, 1)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(150, 1)?.value).toEqual(numberValue(300))
	})

	test('constant folding: 2*3+A1 folds 2*3 to 6 at compile time', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: '2*3+A1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: '10/2-A1', styleId: sid })
		sheet.cells.set(3, 0, { value: EMPTY, formula: '2^10+A1', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(2, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(1028))
	})

	test('lazy CHOOSE only evaluates selected branch', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 3, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'CHOOSE(A1,B1,C1,D1)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(20))
	})

	test('lazy SWITCH only evaluates matched branch', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('B'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'SWITCH(A1,"A",10,"B",20,"C",30,99)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(20))
	})

	test('lazy SWITCH returns default when no match', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: stringValue('Z'), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'SWITCH(A1,"A",10,"B",20,99)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(99))
	})

	test('lazy IFS only evaluates first true condition result', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(75), formula: null, styleId: sid })
		sheet.cells.set(1, 0, {
			value: EMPTY,
			formula: 'IFS(A1>=90,4,A1>=80,3,A1>=70,2,A1>=60,1)',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(1, 0)?.value).toEqual(numberValue(2))
	})
})

describe('LAMBDA, MAP, REDUCE, SCAN', () => {
	test('standalone LAMBDA returns #CALC!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'LAMBDA(x,x+1)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#CALC!'))
	})

	test('defined name LAMBDA can be called as a function', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		wb.definedNames.set('AddOne', 'LAMBDA(x,x+1)')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'AddOne(10)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(11))
	})

	test('defined name LAMBDA with multiple parameters', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		wb.definedNames.set('Add', 'LAMBDA(a,b,a+b)')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'Add(3,7)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(10))
	})

	test('defined name LAMBDA with wrong arg count returns #VALUE!', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		wb.definedNames.set('AddOne', 'LAMBDA(x,x+1)')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'AddOne(1,2)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('MAP applies lambda to each element in a range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'MAP(A1:A3,LAMBDA(x,x*10))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(20))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(30))
	})

	test('MAP applies multi-argument lambda across matching arrays', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'MAP(A1:A3,B1:B3,LAMBDA(x,y,x+y))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(11))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(22))
		expect(sheet.cells.get(2, 2)?.value).toEqual(numberValue(33))
	})

	test('MAP returns #VALUE! for mismatched array shapes or lambda arity', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'MAP(A1:A2,B1,LAMBDA(x,y,x+y))',
			styleId: sid,
		})
		sheet.cells.set(1, 2, {
			value: EMPTY,
			formula: 'MAP(A1:A2,B1:B2,LAMBDA(x,x*2))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(errorValue('#VALUE!'))
		expect(sheet.cells.get(1, 2)?.value).toEqual(errorValue('#VALUE!'))
	})

	test('MAP with single-cell input returns scalar', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'MAP(A1,LAMBDA(x,x*2))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(10))
	})

	test('MAP with defined name lambda', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		wb.definedNames.set('Double', 'LAMBDA(x,x*2)')
		sheet.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'MAP(A1:A2,Double)', styleId: sid })

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(8))
	})

	test('REDUCE accumulates values with lambda', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
	})

	test('REDUCE with non-zero initial value', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'REDUCE(1,A1:A2,LAMBDA(acc,x,acc*x))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(6))
	})

	test('SCAN produces running accumulation results', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(6))
	})

	test('SCAN with multiplication', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'SCAN(1,A1:A3,LAMBDA(acc,x,acc*x))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(6))
		expect(sheet.cells.get(2, 1)?.value).toEqual(numberValue(24))
	})

	test('MAP propagates error from lambda', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'MAP(A1:A2,LAMBDA(x,1/x))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('REDUCE propagates error from lambda', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: sid })
		sheet.cells.set(0, 1, {
			value: EMPTY,
			formula: 'REDUCE(1,A1,LAMBDA(acc,x,acc/x))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#DIV/0!'))
	})

	test('LAMBDA body can reference cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, {
			value: EMPTY,
			formula: 'MAP(B1:B2,LAMBDA(x,x+A1))',
			styleId: sid,
		})

		recalculate(wb, makeCtx())
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(101))
		expect(sheet.cells.get(1, 2)?.value).toEqual(numberValue(102))
	})
})
