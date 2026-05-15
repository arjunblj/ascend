import { describe, expect, test } from 'bun:test'
import type { Sheet, StyleId } from '@ascend/core'
import { createTableId, createWorkbook } from '@ascend/core'
import { EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
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

	test('incremental formula patches refresh cached cycle metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1+1', styleId: sid })

		const cachedFull = analyzeWorkbook(wb)
		const cachedDeps = analyzeWorkbookDependencies(wb)
		expect(cachedFull.cycles).toHaveLength(0)
		expect(cachedDeps.cycles).toHaveLength(0)

		const createCycle = applyOperation(wb, {
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'A1',
			formula: 'B1+1',
		})
		expect(createCycle.ok).toBe(true)

		const afterCreate = analyzeWorkbook(wb)
		const depsAfterCreate = analyzeWorkbookDependencies(wb)
		expect(afterCreate).toBe(cachedFull)
		expect(depsAfterCreate).toBe(cachedDeps)
		expect(afterCreate.cycles).toHaveLength(1)
		expect(depsAfterCreate.cycles).toHaveLength(1)
		expect(afterCreate.cycleKeys.has(cellKey(0, 0, 0))).toBe(true)
		expect(afterCreate.cycleKeys.has(cellKey(0, 0, 1))).toBe(true)
		expect(depsAfterCreate.cycleKeys.has(cellKey(0, 0, 0))).toBe(true)
		expect(depsAfterCreate.cycleKeys.has(cellKey(0, 0, 1))).toBe(true)

		const breakCycle = applyOperation(wb, {
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A1', value: 1 }],
		})
		expect(breakCycle.ok).toBe(true)

		const afterBreak = analyzeWorkbook(wb)
		const depsAfterBreak = analyzeWorkbookDependencies(wb)
		expect(afterBreak.cycles).toHaveLength(0)
		expect(afterBreak.cycleKeys.has(cellKey(0, 0, 0))).toBe(false)
		expect(afterBreak.cycleKeys.has(cellKey(0, 0, 1))).toBe(false)
		expect(depsAfterBreak.cycles).toHaveLength(0)
		expect(depsAfterBreak.cycleKeys.has(cellKey(0, 0, 0))).toBe(false)
		expect(depsAfterBreak.cycleKeys.has(cellKey(0, 0, 1))).toBe(false)

		invalidateWorkbookAnalysis(wb)
		const fresh = analyzeWorkbook(wb)
		expect([...afterBreak.formulas]).toEqual([...fresh.formulas])
		expect([...afterBreak.cycleKeys]).toEqual([...fresh.cycleKeys])
	})

	test('incremental formula rewrites refresh cached shared-formula groups', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: EMPTY,
			formula: 'B1+1',
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
			value: EMPTY,
			formula: null,
			styleId: sid,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})

		const cached = analyzeWorkbook(wb)
		expect(cached.sharedFormulaGroups.get('0:0')).toEqual([cellKey(0, 0, 0), cellKey(0, 1, 0)])

		const result = applyOperation(wb, {
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'A2',
			formula: 'B2+2',
		})
		expect(result.ok).toBe(true)

		const after = analyzeWorkbook(wb)
		expect(after).toBe(cached)
		expect(after.formulas.get(cellKey(0, 1, 0))?.formula).toBe('B2+2')
		expect(after.sharedFormulaGroups.get('0:0')).toBeUndefined()

		const sharedGroups = [...after.sharedFormulaGroups]
		invalidateWorkbookAnalysis(wb)
		expect(sharedGroups).toEqual([...analyzeWorkbook(wb).sharedFormulaGroups])
	})

	test('incremental formula patches refresh cached growing aggregate metadata', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:A1)', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'SUM(A1:A2)', styleId: sid })

		const cached = analyzeWorkbook(wb)
		expect(cached.formulas.get(cellKey(0, 1, 1))?.growingRangeAggregate).toBeDefined()

		const result = applyOperation(wb, {
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'B1',
			formula: 'A1+1',
		})
		expect(result.ok).toBe(true)

		const after = analyzeWorkbook(wb)
		expect(after).toBe(cached)
		expect(after.formulas.get(cellKey(0, 1, 1))?.growingRangeAggregate).toBeUndefined()
		expect(after.growingAggregateAppendIndex.get(cellKey(0, 1, 0))).toBeUndefined()

		const afterFormulas = [...after.formulas]
		invalidateWorkbookAnalysis(wb)
		const fresh = analyzeWorkbook(wb)
		expect(afterFormulas).toEqual([...fresh.formulas])
		expect(fresh.formulas.get(cellKey(0, 1, 1))?.growingRangeAggregate).toBeUndefined()
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

	test('copyRange patches cached analysis for materialized shared formula targets', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
		sheet.cells.set(0, 2, { value: numberValue(9), formula: null, styleId: sid })
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

		const cached = analyzeWorkbook(wb)
		expect(cached.formulas.get(cellKey(0, 1, 0))?.formula).toBe('B2*2')

		const result = applyOperation(wb, {
			op: 'copyRange',
			sheet: 'Sheet1',
			source: 'C1',
			target: 'A2',
		})
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error('copyRange failed')
		expect(result.value.affectedCells).toEqual(['A1', 'A2'])

		const after = analyzeWorkbook(wb)
		expect(after).toBe(cached)
		expect(after.formulas.get(cellKey(0, 0, 0))?.formula).toBe('B1*2')
		expect(after.formulas.has(cellKey(0, 1, 0))).toBe(false)
		expect(after.sharedFormulaGroups.get('0:0')).toBeUndefined()

		const patchedFormulas = [...after.formulas]
		const patchedSharedGroups = [...after.sharedFormulaGroups]
		invalidateWorkbookAnalysis(wb)
		const fresh = analyzeWorkbook(wb)
		expect([...fresh.formulas]).toEqual(patchedFormulas)
		expect([...fresh.sharedFormulaGroups]).toEqual(patchedSharedGroups)
	})

	test('destructive formula-binding edits keep cached analysis equal to full recomputation', () => {
		const cases: readonly {
			readonly name: string
			readonly seed: (sheet: Sheet) => void
			readonly op: Parameters<typeof applyOperation>[1]
			readonly affectedCells: readonly string[]
			readonly reusesCache: boolean
		}[] = [
			{
				name: 'dynamic spill member',
				seed: (sheet) => {
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
							anchorRef: 'Sheet1!A1',
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
							anchorRef: 'Sheet1!A1',
							ref: 'A1:A3',
							isAnchor: false,
						},
					})
					sheet.cells.set(0, 2, { value: numberValue(9), formula: null, styleId: sid })
				},
				op: { op: 'copyRange', sheet: 'Sheet1', source: 'C1', target: 'A2', mode: 'values' },
				affectedCells: ['A1', 'A2', 'A3'],
				reusesCache: true,
			},
			{
				name: 'blocked spill blocker',
				seed: (sheet) => {
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
					sheet.cells.set(0, 2, { value: numberValue(9), formula: null, styleId: sid })
				},
				op: { op: 'copyRange', sheet: 'Sheet1', source: 'C1', target: 'A2', mode: 'values' },
				affectedCells: ['A1', 'A2'],
				reusesCache: true,
			},
			{
				name: 'rich text shared formula member',
				seed: (sheet) => {
					sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
					sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
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
				},
				op: {
					op: 'setRichText',
					sheet: 'Sheet1',
					ref: 'A2',
					runs: [{ text: 'manual shared member' }],
				},
				affectedCells: ['Sheet1!A1', 'Sheet1!A2'],
				reusesCache: false,
			},
			{
				name: 'rich text dynamic spill member',
				seed: (sheet) => {
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
							anchorRef: 'Sheet1!A1',
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
							anchorRef: 'Sheet1!A1',
							ref: 'A1:A3',
							isAnchor: false,
						},
					})
				},
				op: {
					op: 'setRichText',
					sheet: 'Sheet1',
					ref: 'A2',
					runs: [{ text: 'manual spill member' }],
				},
				affectedCells: ['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3'],
				reusesCache: false,
			},
			{
				name: 'rich text data table member',
				seed: (sheet) => {
					sheet.cells.set(2, 2, {
						value: numberValue(10),
						formula: null,
						styleId: sid,
						formulaInfo: { kind: 'dataTable', ref: 'C3:C5', dtr: true, r1: 'A1' },
					})
					sheet.cells.set(3, 2, { value: numberValue(20), formula: null, styleId: sid })
					sheet.cells.set(4, 2, { value: numberValue(30), formula: null, styleId: sid })
				},
				op: {
					op: 'setRichText',
					sheet: 'Sheet1',
					ref: 'C4',
					runs: [{ text: 'manual data table member' }],
				},
				affectedCells: ['Sheet1!C3', 'Sheet1!C4'],
				reusesCache: false,
			},
			{
				name: 'data table member',
				seed: (sheet) => {
					sheet.cells.set(2, 2, {
						value: numberValue(10),
						formula: null,
						styleId: sid,
						formulaInfo: { kind: 'dataTable', ref: 'C3:C5', dtr: true, r1: 'A1' },
					})
					sheet.cells.set(3, 2, { value: numberValue(20), formula: null, styleId: sid })
					sheet.cells.set(0, 5, { value: numberValue(9), formula: null, styleId: sid })
				},
				op: { op: 'copyRange', sheet: 'Sheet1', source: 'F1', target: 'C4', mode: 'values' },
				affectedCells: ['C3', 'C4'],
				reusesCache: true,
			},
			{
				name: 'moveRange target shared formula member',
				seed: (sheet) => {
					sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
					sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: sid })
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
				},
				op: { op: 'moveRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'all' },
				affectedCells: ['A1', 'A2', 'D1'],
				reusesCache: false,
			},
			{
				name: 'moveRange target data table member',
				seed: (sheet) => {
					sheet.cells.set(2, 2, {
						value: numberValue(10),
						formula: null,
						styleId: sid,
						formulaInfo: { kind: 'dataTable', ref: 'C3:C5', dtr: true, r1: 'A1' },
					})
					sheet.cells.set(3, 2, { value: numberValue(20), formula: null, styleId: sid })
					sheet.cells.set(0, 5, { value: numberValue(9), formula: null, styleId: sid })
				},
				op: { op: 'moveRange', sheet: 'Sheet1', source: 'F1', target: 'C4', mode: 'all' },
				affectedCells: ['C3', 'C4', 'F1'],
				reusesCache: false,
			},
			{
				name: 'sortRange shared formula member',
				seed: (sheet) => {
					sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
					sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
					sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: sid })
					sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: sid })
					sheet.cells.set(0, 3, {
						value: numberValue(20),
						formula: 'B1*2',
						styleId: sid,
						formulaInfo: {
							kind: 'shared',
							sharedIndex: '0',
							isMaster: true,
							masterRef: 'D1',
							ref: 'D1:D2',
						},
					})
					sheet.cells.set(1, 3, {
						value: numberValue(40),
						formula: null,
						styleId: sid,
						formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'D1' },
					})
				},
				op: { op: 'sortRange', sheet: 'Sheet1', range: 'A1:D2', by: [{ column: 'A' }] },
				affectedCells: ['D1', 'D2', 'A2', 'B2', 'C2'],
				reusesCache: false,
			},
			{
				name: 'sortRange dynamic spill member',
				seed: (sheet) => {
					sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: sid })
					sheet.cells.set(1, 0, { value: stringValue('b'), formula: null, styleId: sid })
					sheet.cells.set(2, 0, { value: stringValue('c'), formula: null, styleId: sid })
					sheet.cells.set(0, 3, {
						value: numberValue(1),
						formula: 'SEQUENCE(3)',
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
							ref: 'D1:D3',
							isAnchor: false,
						},
					})
					sheet.cells.set(2, 3, {
						value: numberValue(3),
						formula: null,
						styleId: sid,
						formulaInfo: {
							kind: 'spill',
							anchorRef: 'Sheet1!D1',
							ref: 'D1:D3',
							isAnchor: false,
						},
					})
				},
				op: { op: 'sortRange', sheet: 'Sheet1', range: 'A1:D3', by: [{ column: 'A' }] },
				affectedCells: ['D1', 'D2', 'D3', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'],
				reusesCache: false,
			},
			{
				name: 'sortRange data table member',
				seed: (sheet) => {
					for (let row = 0; row < 5; row++) {
						sheet.cells.set(row, 0, {
							value: stringValue(String.fromCharCode(97 + row)),
							formula: null,
							styleId: sid,
						})
					}
					sheet.cells.set(2, 2, {
						value: numberValue(10),
						formula: null,
						styleId: sid,
						formulaInfo: { kind: 'dataTable', ref: 'C3:C5', dtr: true, r1: 'A1' },
					})
					sheet.cells.set(3, 2, { value: numberValue(20), formula: null, styleId: sid })
					sheet.cells.set(4, 2, { value: numberValue(30), formula: null, styleId: sid })
				},
				op: { op: 'sortRange', sheet: 'Sheet1', range: 'A1:C5', by: [{ column: 'A' }] },
				affectedCells: ['C3', 'A2', 'B2', 'C2', 'A3', 'B3', 'A4', 'B4', 'C4', 'A5', 'B5', 'C5'],
				reusesCache: false,
			},
			{
				name: 'sortRange blocked spill anchor',
				seed: (sheet) => {
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
				},
				op: { op: 'sortRange', sheet: 'Sheet1', range: 'A1:A2', by: [{ column: 'A' }] },
				affectedCells: ['A1', 'A2'],
				reusesCache: false,
			},
		]

		for (const entry of cases) {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			entry.seed(sheet)

			const cached = analyzeWorkbook(wb)
			const result = applyOperation(wb, entry.op)
			expect(result.ok, entry.name).toBe(true)
			if (!result.ok) throw new Error(`${entry.name} edit failed`)
			expect(result.value.affectedCells, entry.name).toEqual(entry.affectedCells)

			const after = analyzeWorkbook(wb)
			if (entry.reusesCache) {
				expect(after, entry.name).toBe(cached)
			} else {
				expect(after, entry.name).not.toBe(cached)
			}
			const patchedFormulas = [...after.formulas]
			const patchedSharedGroups = [...after.sharedFormulaGroups]
			invalidateWorkbookAnalysis(wb)
			const fresh = analyzeWorkbook(wb)
			expect([...fresh.formulas], entry.name).toEqual(patchedFormulas)
			expect([...fresh.sharedFormulaGroups], entry.name).toEqual(patchedSharedGroups)
		}
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
