import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { EMPTY, numberValue } from '@ascend/schema'
import { analyzeWorkbook } from './analysis.ts'
import { cellKey } from './dep-graph.ts'

const sid = 0 as StyleId

describe('analyzeWorkbook', () => {
	test('builds a dependency graph and tracks volatile formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })
		s.cells.set(0, 2, { value: EMPTY, formula: 'NOW()', styleId: sid })

		const analysis = analyzeWorkbook(wb)
		expect(analysis.formulas.size).toBe(2)

		const formulas = [...analysis.formulas.values()]
		expect(formulas.some((formula) => formula.volatile)).toBe(true)
		expect(analysis.dependencyGraph.getPrecedents(cellKey(0, 0, 1))).toEqual([cellKey(0, 0, 0)])
	})

	test('can scope analysis to a range', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: '1+1', styleId: sid })
		s.cells.set(5, 0, { value: EMPTY, formula: '2+2', styleId: sid })

		const analysis = analyzeWorkbook(wb, {
			range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
		})

		expect(analysis.formulas.size).toBe(1)
		expect([...analysis.formulas.values()][0]?.row).toBe(0)
	})

	test('stores rectangular range dependencies as spans', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: sid })

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
		expect(analysis.dependencyGraph.getPrecedents(cellKey(0, 0, 1))).toEqual([
			cellKey(0, 0, 0),
			cellKey(0, 1, 0),
		])
		expect(analysis.dependencyGraph.getDependents(cellKey(0, 1, 0))).toEqual([cellKey(0, 0, 1)])
	})

	test('LET bindings do not get treated as defined-name dependencies', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		wb.definedNames.set('x', 'Sheet1!B1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'LET(x, A1, x+1)', styleId: sid })
		s.cells.set(0, 1, { value: numberValue(99), formula: null, styleId: sid })

		const analysis = analyzeWorkbook(wb)
		const formula = [...analysis.formulas.values()][0]
		expect(formula?.deps).toEqual([cellKey(0, 0, 0)])
		expect(formula?.rangeDeps).toEqual([])
	})

	test('formula group detection reuses AST for consecutive column cells with same shape', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		s.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1*2', styleId: sid })
		s.cells.set(1, 1, { value: EMPTY, formula: 'A2*2', styleId: sid })
		s.cells.set(2, 1, { value: EMPTY, formula: 'A3*2', styleId: sid })

		const analysis = analyzeWorkbook(wb)
		expect(analysis.formulas.size).toBe(3)
		const b1 = analysis.formulas.get(cellKey(0, 0, 1))
		const b2 = analysis.formulas.get(cellKey(0, 1, 1))
		const b3 = analysis.formulas.get(cellKey(0, 2, 1))
		expect(b1?.deps).toEqual([cellKey(0, 0, 0)])
		expect(b2?.deps).toEqual([cellKey(0, 1, 0)])
		expect(b3?.deps).toEqual([cellKey(0, 2, 0)])
	})

	test('expands 3D sheet-span references into dependencies across contiguous sheets', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const s3 = wb.addSheet('Sheet3')
		const calc = wb.addSheet('Calc')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: sid })
		s3.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
		calc.cells.set(0, 0, { value: EMPTY, formula: 'SUM(Sheet1:Sheet3!A1)', styleId: sid })

		const analysis = analyzeWorkbook(wb)
		const formula = [...analysis.formulas.values()].find((entry) => entry.sheetName === 'Calc')
		expect(formula?.deps).toEqual([cellKey(0, 0, 0), cellKey(1, 0, 0), cellKey(2, 0, 0)])
		expect(formula?.rangeDeps).toEqual([])
	})
})
