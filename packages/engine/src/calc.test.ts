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
