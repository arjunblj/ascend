import { describe, expect, test } from 'bun:test'
import { createWorkbook } from '@ascend/core'
import { EMPTY, numberValue } from '@ascend/schema'
import { analyzeWorkbook } from './analysis.ts'

describe('analyzeWorkbook', () => {
	test('builds a dependency graph and tracks volatile formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: 0 })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: 0 })
		s.cells.set(0, 2, { value: EMPTY, formula: 'NOW()', styleId: 0 })

		const analysis = analyzeWorkbook(wb)
		expect(analysis.formulas.size).toBe(2)

		const formulas = [...analysis.formulas.values()]
		expect(formulas.some((formula) => formula.volatile)).toBe(true)
		expect(analysis.dependencyGraph.getPrecedents('0:0:1')).toEqual(['0:0:0'])
	})

	test('can scope analysis to a range', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: '1+1', styleId: 0 })
		s.cells.set(5, 0, { value: EMPTY, formula: '2+2', styleId: 0 })

		const analysis = analyzeWorkbook(wb, {
			range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
		})

		expect(analysis.formulas.size).toBe(1)
		expect([...analysis.formulas.values()][0]?.row).toBe(0)
	})

	test('stores rectangular range dependencies as spans', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: 0 })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: 0 })
		s.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: 0 })

		const analysis = analyzeWorkbook(wb)
		const formula = [...analysis.formulas.values()][0]
		expect(formula?.deps).toEqual([])
		expect(formula?.rangeDeps).toEqual([
			{
				sheetIndex: 0,
				startRow: 0,
				startCol: 0,
				endRow: 1,
				endCol: 0,
			},
		])
		expect(analysis.dependencyGraph.getPrecedents('0:0:1')).toEqual(['0:0:0', '0:1:0'])
	})
})
