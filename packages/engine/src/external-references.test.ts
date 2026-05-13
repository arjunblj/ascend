import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { EMPTY, errorValue, numberValue } from '@ascend/schema'
import { recalculate } from './calc.ts'
import { defaultCalcContext, type ExternalRangeReference } from './calc-context.ts'

const sid = 0 as StyleId
const maxRow = 1_048_575
const maxCol = 16_383

function formulaCell(formula: string) {
	return { value: EMPTY, formula, styleId: sid }
}

describe('external reference evaluation', () => {
	test('whole-column and whole-row references resolve through externalReferences', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		const decoy = wb.addSheet('[Budget.xlsx]Inputs')
		decoy.cells.set(0, 0, { value: numberValue(10_000), formula: null, styleId: sid })
		decoy.cells.set(1, 1, { value: numberValue(20_000), formula: null, styleId: sid })

		const columnFormula = 'SUM([Budget.xlsx]Inputs!A:A)'
		const rowFormula = 'SUM([Budget.xlsx]Inputs!2:2)'
		sheet.cells.set(0, 0, formulaCell(columnFormula))
		sheet.cells.set(0, 1, formulaCell(rowFormula))

		const rangeCalls: ExternalRangeReference[] = []
		const result = recalculate(
			wb,
			defaultCalcContext({
				externalReferences: {
					resolveRange: (ref) => {
						rangeCalls.push(ref)
						if (ref.workbook !== 'Budget.xlsx' || ref.sheet !== 'Inputs') return undefined
						if (ref.row === 0 && ref.col === 0 && ref.endRow === maxRow && ref.endCol === 0) {
							return [[numberValue(10)], [numberValue(20)], [numberValue(30)]]
						}
						if (ref.row === 1 && ref.col === 0 && ref.endRow === 1 && ref.endCol === maxCol) {
							return [[numberValue(4), numberValue(5), numberValue(6)]]
						}
						return undefined
					},
				},
			}),
		)

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(numberValue(60))
		expect(sheet.cells.get(0, 1)?.value).toEqual(numberValue(15))
		expect(sheet.cells.get(0, 0)?.formula).toBe(columnFormula)
		expect(sheet.cells.get(0, 1)?.formula).toBe(rowFormula)
		expect(rangeCalls).toContainEqual({
			workbook: 'Budget.xlsx',
			sheet: 'Inputs',
			row: 0,
			col: 0,
			endRow: maxRow,
			endCol: 0,
		})
		expect(rangeCalls).toContainEqual({
			workbook: 'Budget.xlsx',
			sheet: 'Inputs',
			row: 1,
			col: 0,
			endRow: 1,
			endCol: maxCol,
		})
	})

	test('implicit intersection of external ranges reads external range data', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(999), formula: null, styleId: sid })
		sheet.cells.set(1, 3, { value: numberValue(888), formula: null, styleId: sid })

		const columnFormula = '@[Budget.xlsx]Inputs!A:A'
		const rowFormula = '@[Budget.xlsx]Inputs!2:2'
		const intersectionFormula = '[Budget.xlsx]Inputs!A:A [Budget.xlsx]Inputs!2:2'
		const reversedIntersectionFormula = '[Budget.xlsx]Inputs!2:2 [Budget.xlsx]Inputs!A:A'
		sheet.cells.set(1, 1, formulaCell(columnFormula))
		sheet.cells.set(0, 3, formulaCell(rowFormula))
		sheet.cells.set(0, 4, formulaCell(intersectionFormula))
		sheet.cells.set(0, 5, formulaCell(reversedIntersectionFormula))

		const result = recalculate(
			wb,
			defaultCalcContext({
				externalReferences: {
					resolveRange: (ref) => {
						if (ref.workbook !== 'Budget.xlsx' || ref.sheet !== 'Inputs') return undefined
						if (ref.row === 0 && ref.col === 0 && ref.endRow === maxRow && ref.endCol === 0) {
							return [[numberValue(10)], [numberValue(25)], [numberValue(40)]]
						}
						if (ref.row === 1 && ref.col === 0 && ref.endRow === 1 && ref.endCol === maxCol) {
							return [[numberValue(25), numberValue(30), numberValue(32), numberValue(35)]]
						}
						return undefined
					},
				},
			}),
		)

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(1, 1)?.value).toEqual(numberValue(25))
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(35))
		expect(sheet.cells.get(0, 4)?.value).toEqual(numberValue(25))
		expect(sheet.cells.get(0, 5)?.value).toEqual(numberValue(25))
		expect(sheet.cells.get(1, 1)?.formula).toBe(columnFormula)
		expect(sheet.cells.get(0, 3)?.formula).toBe(rowFormula)
		expect(sheet.cells.get(0, 4)?.formula).toBe(intersectionFormula)
		expect(sheet.cells.get(0, 5)?.formula).toBe(reversedIntersectionFormula)
	})

	test('external intersections require the same workbook and sheet identity', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(1, 0, { value: numberValue(999), formula: null, styleId: sid })

		const localIntersection = '[Budget.xlsx]Inputs!A:A A2:A2'
		const workbookMismatch = '[Budget.xlsx]Inputs!A:A [Other.xlsx]Inputs!2:2'
		const externalCellIntersection = '[Budget.xlsx]Inputs!A2 [Budget.xlsx]Inputs!A:A'
		sheet.cells.set(0, 0, formulaCell(localIntersection))
		sheet.cells.set(0, 1, formulaCell(workbookMismatch))
		sheet.cells.set(0, 2, formulaCell(externalCellIntersection))

		const result = recalculate(
			wb,
			defaultCalcContext({
				externalReferences: {
					resolveCell: (ref) => {
						if (ref.workbook === 'Budget.xlsx' && ref.sheet === 'Inputs') {
							return ref.row === 1 && ref.col === 0 ? numberValue(25) : numberValue(10)
						}
						if (ref.workbook === 'Other.xlsx' && ref.sheet === 'Inputs') {
							return numberValue(100)
						}
						return undefined
					},
					resolveRange: (ref) => {
						if (ref.workbook === 'Budget.xlsx' && ref.sheet === 'Inputs') {
							if (ref.row === 0 && ref.col === 0 && ref.endRow === maxRow && ref.endCol === 0) {
								return [[numberValue(10)], [numberValue(25)], [numberValue(40)]]
							}
						}
						if (ref.workbook === 'Other.xlsx' && ref.sheet === 'Inputs') {
							if (ref.row === 1 && ref.col === 0 && ref.endRow === 1 && ref.endCol === maxCol) {
								return [[numberValue(100), numberValue(200)]]
							}
						}
						return undefined
					},
				},
			}),
		)

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 0)?.value).toEqual(errorValue('#NULL!'))
		expect(sheet.cells.get(0, 1)?.value).toEqual(errorValue('#NULL!'))
		expect(sheet.cells.get(0, 2)?.value).toEqual(numberValue(25))
	})
})
