import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, createWorkbook } from '@ascend/core'
import { EMPTY, numberValue } from '@ascend/schema'
import {
	analyzeWorkbook,
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	invalidateWorkbookAnalysis,
} from './analysis.ts'
import { cellKey } from './dep-graph.ts'
import { applyOperation } from './operations.ts'

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
		const p = analysis.dependencyGraph.getPrecedents(cellKey(0, 0, 1))
		expect(p.cells).toEqual([cellKey(0, 0, 0)])
		expect(p.ranges).toEqual([])
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
		const p = analysis.dependencyGraph.getPrecedents(cellKey(0, 0, 1))
		expect(p.cells).toEqual([])
		expect(p.ranges).toEqual([
			{
				sheetIndex: 0,
				startRow: 0,
				startCol: 0,
				endRow: 1,
				endCol: 0,
			},
		])
		expect(analysis.dependencyGraph.getDependents(cellKey(0, 1, 0))).toEqual([cellKey(0, 0, 1)])
	})

	test('narrows INDEX structured table dependencies with constant columns', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(1, 2, {
			value: EMPTY,
			formula: 'INDEX(Sales[],MATCH(Sales[[#This Row],[Item]],Sales[Item],0),2)',
			styleId: sid,
		})
		sheet.cells.set(2, 0, { value: numberValue(11), formula: null, styleId: sid })
		sheet.cells.set(2, 1, { value: numberValue(21), formula: null, styleId: sid })
		sheet.cells.set(2, 2, { value: EMPTY, formula: 'C2+1', styleId: sid })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
			columns: [{ name: 'Item' }, { name: 'Price' }, { name: 'Total' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const analysis = analyzeWorkbook(wb)
		const formulaKey = cellKey(0, 1, 2)
		const formula = analysis.formulas.get(formulaKey)
		expect(formula?.deps).toEqual([])
		expect(formula?.rangeDeps).toContainEqual({
			sheetIndex: 0,
			startRow: 1,
			startCol: 1,
			endRow: 2,
			endCol: 1,
		})
		expect(formula?.rangeDeps).not.toContainEqual({
			sheetIndex: 0,
			startRow: 1,
			startCol: 0,
			endRow: 2,
			endCol: 2,
		})
		expect(analysis.cycleKeys.has(formulaKey)).toBe(false)
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

	test('LAMBDA parameters do not get treated as defined-name dependencies', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		wb.definedNames.set('x', 'Sheet1!B1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: numberValue(99), formula: null, styleId: sid })
		s.cells.set(0, 2, { value: EMPTY, formula: 'MAP(A1:A2,LAMBDA(x,x+1))', styleId: sid })

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

	test('detects growing SUM ranges for decomposition reuse', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:A1)', styleId: sid })
		s.cells.set(1, 1, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: sid })

		const analysis = analyzeWorkbook(wb)
		const second = analysis.formulas.get(cellKey(0, 1, 1))
		expect(second?.growingRangeAggregate).toEqual({
			functionName: 'SUM',
			previousKey: cellKey(0, 0, 1),
			previousSheetIndex: 0,
			previousRow: 0,
			previousCol: 1,
			appendSheetIndex: 0,
			appendStartRow: 1,
			appendStartCol: 0,
			appendEndRow: 1,
			appendEndCol: 0,
		})
	})

	test('detects growing SUM ranges across row gaps', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		s.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: sid })
		s.cells.set(3, 0, { value: numberValue(4), formula: null, styleId: sid })
		s.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:A1)', styleId: sid })
		s.cells.set(3, 1, { value: EMPTY, formula: 'SUM(A1:A4)', styleId: sid })

		const analysis = analyzeWorkbook(wb)
		const fourth = analysis.formulas.get(cellKey(0, 3, 1))
		expect(fourth?.growingRangeAggregate).toEqual({
			functionName: 'SUM',
			previousKey: cellKey(0, 0, 1),
			previousSheetIndex: 0,
			previousRow: 0,
			previousCol: 1,
			appendSheetIndex: 0,
			appendStartRow: 1,
			appendStartCol: 0,
			appendEndRow: 3,
			appendEndCol: 0,
		})
	})

	test('detects growing aggregate ranges for COUNT, AVERAGE, MIN, and MAX', () => {
		for (const functionName of ['COUNT', 'AVERAGE', 'MIN', 'MAX'] as const) {
			const wb = createWorkbook()
			const s = wb.addSheet('Sheet1')
			s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
			s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
			s.cells.set(0, 1, { value: EMPTY, formula: `${functionName}(A1:A1)`, styleId: sid })
			s.cells.set(1, 1, { value: EMPTY, formula: `${functionName}(A1:A2)`, styleId: sid })

			const analysis = analyzeWorkbook(wb)
			const second = analysis.formulas.get(cellKey(0, 1, 1))
			expect(second?.growingRangeAggregate?.functionName).toBe(functionName)
			expect(second?.growingRangeAggregate?.previousKey).toBe(cellKey(0, 0, 1))
		}
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

	test('caches cycle membership alongside dependency analysis', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1+1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const analysis = analyzeWorkbook(wb)
		expect(analysis.cycles).toHaveLength(1)
		expect(new Set(analysis.cycles[0])).toEqual(new Set([cellKey(0, 0, 0), cellKey(0, 0, 1)]))
		expect(analysis.cycleKeys.has(cellKey(0, 0, 0))).toBe(true)
		expect(analysis.cycleKeys.has(cellKey(0, 0, 1))).toBe(true)
	})

	test('dependency analysis exposes resolved formula dependencies', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const analysis = analyzeWorkbookDependencies(wb)
		const formula = analysis.resolvedFormulas.get(cellKey(0, 1, 0))
		expect(formula?.deps).toEqual([cellKey(0, 0, 0)])
		expect(formula?.rangeDeps).toEqual([])
	})

	test('insertRows incrementally shifts cached analysis keys and dependencies', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: sid })

		const before = analyzeWorkbook(wb)
		expect(before.formulas.has(cellKey(0, 2, 0))).toBe(true)

		const result = applyOperation(wb, { op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 })
		expect(result.ok).toBe(true)

		const after = analyzeWorkbook(wb)
		const shifted = after.formulas.get(cellKey(0, 3, 0))
		expect(shifted?.formula).toBe('SUM(A1:A3)')
		expect(shifted?.rangeDeps).toEqual([
			{
				sheetIndex: 0,
				startRow: 0,
				startCol: 0,
				endRow: 2,
				endCol: 0,
			},
		])
		expect(after.formulas.has(cellKey(0, 2, 0))).toBe(false)
	})

	test('copyRange incrementally patches cached analysis for translated formulas', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: EMPTY, formula: 'A1+B1', styleId: sid })

		analyzeWorkbook(wb)
		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'C1:C1',
			target: 'C2',
		})
		expect(result.ok).toBe(true)

		const after = analyzeWorkbook(wb)
		const copied = after.formulas.get(cellKey(0, 1, 2))
		expect(copied?.formula).toBe('A2+B2')
		expect(copied?.deps).toEqual([cellKey(0, 1, 0), cellKey(0, 1, 1)])
	})

	test('analyzeWorkbook returns cached result on second call (formulas parsed once, shared across tools)', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const first = analyzeWorkbook(wb)
		const second = analyzeWorkbook(wb)
		expect(second).toBe(first)

		const formulasFirst = analyzeWorkbookFormulas(wb)
		const formulasSecond = analyzeWorkbookFormulas(wb)
		expect(formulasSecond).toBe(formulasFirst)

		const depsFirst = analyzeWorkbookDependencies(wb)
		const depsSecond = analyzeWorkbookDependencies(wb)
		expect(depsSecond).toBe(depsFirst)
	})

	test('invalidateWorkbookAnalysis clears cache so next analyzeWorkbook returns fresh result', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const before = analyzeWorkbook(wb)
		invalidateWorkbookAnalysis(wb)
		const after = analyzeWorkbook(wb)

		expect(after).not.toBe(before)
		expect(after.formulas.size).toBe(1)
		expect(after.formulas.get(cellKey(0, 0, 1))?.formula).toBe('A1+1')
	})
})
