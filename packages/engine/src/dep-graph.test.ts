import { describe, expect, test } from 'bun:test'
import { cellKey, DependencyGraph, IntervalIndex, parseCellKey } from './dep-graph.ts'

describe('DependencyGraph', () => {
	test('empty graph returns no dependents', () => {
		const g = new DependencyGraph()
		expect(g.getDependents(cellKey(0, 0, 0))).toEqual([])
	})

	test('empty graph returns no precedents', () => {
		const g = new DependencyGraph()
		expect(g.getPrecedents(cellKey(0, 0, 0))).toEqual({ cells: [], ranges: [] })
	})

	test('addFormula tracks precedents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 0, 1)
		const c = cellKey(0, 0, 2)
		g.addFormula(c, [a, b], false)
		const p = g.getPrecedents(c)
		expect(p.ranges).toEqual([])
		expect(p.cells.sort()).toEqual([a, b].sort())
	})

	test('addFormula tracks dependents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const c = cellKey(0, 0, 2)
		g.addFormula(c, [a], false)
		expect(g.getDependents(a)).toEqual([c])
	})

	test('range dependencies track dependents without eager expansion', () => {
		const g = new DependencyGraph()
		const formula = cellKey(0, 0, 2)
		const range = {
			sheetIndex: 0,
			startRow: 0,
			startCol: 0,
			endRow: 2,
			endCol: 0,
		}
		g.addFormula(formula, [], false, [range])
		expect(g.getDependents(cellKey(0, 1, 0))).toEqual([formula])
		const p = g.getPrecedents(formula)
		expect(p.cells).toEqual([])
		expect(p.ranges).toEqual([range])
	})

	test('dirty propagation includes range-backed dependents', () => {
		const g = new DependencyGraph()
		const source = cellKey(0, 0, 0)
		const formula = cellKey(0, 0, 2)
		g.addFormula(formula, [], false, [
			{
				sheetIndex: 0,
				startRow: 0,
				startCol: 0,
				endRow: 2,
				endCol: 0,
			},
		])
		const dirty = g.getDirtySet([source])
		expect(dirty.has(source)).toBe(true)
		expect(dirty.has(formula)).toBe(true)
	})

	test('removeFormula cleans up dependents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const c = cellKey(0, 0, 2)
		g.addFormula(c, [a], false)
		g.removeFormula(c)
		expect(g.getDependents(a)).toEqual([])
		expect(g.getPrecedents(c)).toEqual({ cells: [], ranges: [] })
	})

	test('removing non-existent formula is a no-op', () => {
		const g = new DependencyGraph()
		g.removeFormula(cellKey(0, 99, 99))
		expect(g.getAllFormulaCells()).toEqual([])
	})

	test('dirty propagation through a chain', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 2, 0)
		g.addFormula(b, [a], false)
		g.addFormula(c, [b], false)
		const dirty = g.getDirtySet([a])
		expect(dirty.has(a)).toBe(true)
		expect(dirty.has(b)).toBe(true)
		expect(dirty.has(c)).toBe(true)
	})

	test('dirty propagation with diamond dependency', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 1, 1)
		const d = cellKey(0, 2, 0)
		g.addFormula(b, [a], false)
		g.addFormula(c, [a], false)
		g.addFormula(d, [b, c], false)
		const dirty = g.getDirtySet([a])
		expect(dirty.size).toBe(4)
	})

	test('topological sort orders precedents before dependents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 2, 0)
		g.addFormula(b, [a], false)
		g.addFormula(c, [b], false)
		const dirty = new Set([a, b, c])
		const order = g.getEvalOrder(dirty)
		expect(order.indexOf(a)).toBeLessThan(order.indexOf(b))
		expect(order.indexOf(b)).toBeLessThan(order.indexOf(c))
	})

	test('topological sort handles diamond correctly', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 1, 1)
		const d = cellKey(0, 2, 0)
		g.addFormula(b, [a], false)
		g.addFormula(c, [a], false)
		g.addFormula(d, [b, c], false)
		const dirty = new Set([a, b, c, d])
		const order = g.getEvalOrder(dirty)
		expect(order.indexOf(a)).toBeLessThan(order.indexOf(b))
		expect(order.indexOf(a)).toBeLessThan(order.indexOf(c))
		expect(order.indexOf(b)).toBeLessThan(order.indexOf(d))
		expect(order.indexOf(c)).toBeLessThan(order.indexOf(d))
	})

	test('topological sort keeps range precedents before dependent formulas', () => {
		const g = new DependencyGraph()
		const a1 = cellKey(0, 0, 0)
		const a2 = cellKey(0, 1, 0)
		const sumCell = cellKey(0, 2, 0)
		g.addFormula(a1, [], false)
		g.addFormula(a2, [a1], false)
		g.addFormula(sumCell, [], false, [
			{ sheetIndex: 0, startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
		])
		const dirty = new Set([a1, a2, sumCell])
		const order = g.getEvalOrder(dirty)
		expect(order.indexOf(a1)).toBeLessThan(order.indexOf(sumCell))
		expect(order.indexOf(a2)).toBeLessThan(order.indexOf(sumCell))
	})

	test('detect simple cycle A->B->A', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		g.addFormula(a, [b], false)
		g.addFormula(b, [a], false)
		const cycles = g.detectCycles()
		expect(cycles.length).toBe(1)
		const scc = cycles[0]
		expect(scc).toContain(a)
		expect(scc).toContain(b)
	})

	test('detect self-referencing cycle', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		g.addFormula(a, [a], false)
		const cycles = g.detectCycles()
		expect(cycles.length).toBe(1)
		expect(cycles[0]).toEqual([a])
	})

	test('range dependencies alone do not imply self-referencing cycles', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		g.addFormula(a, [], false, [{ sheetIndex: 0, startRow: 0, startCol: 0, endRow: 0, endCol: 0 }])
		const cycles = g.detectCycles()
		expect(cycles.length).toBe(0)
	})

	test('detect complex multi-node cycle', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 2, 0)
		g.addFormula(a, [c], false)
		g.addFormula(b, [a], false)
		g.addFormula(c, [b], false)
		const cycles = g.detectCycles()
		expect(cycles.length).toBe(1)
		expect(cycles[0]?.length).toBe(3)
	})

	test('no cycles in acyclic graph', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		g.addFormula(b, [a], false)
		const cycles = g.detectCycles()
		expect(cycles.length).toBe(0)
	})

	test('volatile cell tracking', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		g.addFormula(a, [], true)
		g.addFormula(b, [a], false)
		const volatiles = g.getVolatiles()
		expect(volatiles).toEqual([a])
	})

	test('volatile cells always dirty', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		g.addFormula(a, [], true)
		g.addFormula(b, [a], false)
		const dirty = g.getDirtySet(g.getVolatiles())
		expect(dirty.has(a)).toBe(true)
		expect(dirty.has(b)).toBe(true)
	})

	test('cellKey packs and unpacks correctly', () => {
		expect(parseCellKey(cellKey(0, 5, 2))).toEqual([0, 5, 2])
		expect(parseCellKey(cellKey(1, 0, 0))).toEqual([1, 0, 0])
		expect(parseCellKey(cellKey(3, 1048575, 16383))).toEqual([3, 1048575, 16383])
	})

	test('hasFormula returns correct state', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		expect(g.hasFormula(a)).toBe(false)
		g.addFormula(a, [], false)
		expect(g.hasFormula(a)).toBe(true)
	})

	test('clear removes all data', () => {
		const g = new DependencyGraph()
		g.addFormula(cellKey(0, 0, 0), [cellKey(0, 1, 0)], false)
		g.clear()
		expect(g.getAllFormulaCells()).toEqual([])
		expect(g.getDependents(cellKey(0, 1, 0))).toEqual([])
	})

	test('getDirtySet with overlapping cumulative SUM ranges', () => {
		const g = new DependencyGraph()
		const N = 200
		for (let r = 0; r < N; r++) {
			const formulaKey = cellKey(0, r, 1)
			g.addFormula(formulaKey, [], false, [
				{ sheetIndex: 0, startRow: 0, startCol: 0, endRow: r, endCol: 0 },
			])
		}

		const allDataKeys = []
		for (let r = 0; r < N; r++) allDataKeys.push(cellKey(0, r, 0))
		const dirty = g.getDirtySet(allDataKeys)

		for (let r = 0; r < N; r++) {
			expect(dirty.has(cellKey(0, r, 0))).toBe(true)
			expect(dirty.has(cellKey(0, r, 1))).toBe(true)
		}
		expect(dirty.size).toBe(N * 2)
	})

	test('getDirtySet with single cell change in overlapping ranges', () => {
		const g = new DependencyGraph()
		const N = 100
		for (let r = 0; r < N; r++) {
			g.addFormula(cellKey(0, r, 1), [], false, [
				{ sheetIndex: 0, startRow: 0, startCol: 0, endRow: r, endCol: 0 },
			])
		}

		const dirty = g.getDirtySet([cellKey(0, 50, 0)])
		expect(dirty.has(cellKey(0, 50, 0))).toBe(true)
		for (let r = 50; r < N; r++) {
			expect(dirty.has(cellKey(0, r, 1))).toBe(true)
		}
		for (let r = 0; r < 50; r++) {
			expect(dirty.has(cellKey(0, r, 1))).toBe(false)
		}
	})

	test('getDirtySet batch vs per-cell gives identical results', () => {
		const g = new DependencyGraph()
		for (let r = 0; r < 50; r++) {
			g.addFormula(cellKey(0, r, 1), [], false, [
				{ sheetIndex: 0, startRow: 0, startCol: 0, endRow: r, endCol: 0 },
			])
		}

		const singleCellResults = new Set<number>()
		for (let r = 0; r < 50; r++) {
			const dirty = g.getDirtySet([cellKey(0, r, 0)])
			for (const k of dirty) singleCellResults.add(k)
		}

		const allKeys = []
		for (let r = 0; r < 50; r++) allKeys.push(cellKey(0, r, 0))
		const batchDirty = g.getDirtySet(allKeys)

		for (const k of singleCellResults) {
			expect(batchDirty.has(k)).toBe(true)
		}
		for (const k of batchDirty) {
			if (k < cellKey(0, 0, 1)) continue
			expect(singleCellResults.has(k)).toBe(true)
		}
	})

	test('overlapping ranges: getDirtySet runs in subquadratic time', () => {
		const N = 1000
		const g = new DependencyGraph()
		for (let r = 0; r < N; r++) {
			g.addFormula(cellKey(0, r, 1), [], false, [
				{ sheetIndex: 0, startRow: 0, startCol: 0, endRow: r, endCol: 0 },
			])
		}

		const allDataKeys: number[] = []
		for (let r = 0; r < N; r++) allDataKeys.push(cellKey(0, r, 0))

		g.getDirtySet([cellKey(0, 0, 0)])
		const start = performance.now()
		const dirty = g.getDirtySet(allDataKeys)
		const elapsed = performance.now() - start

		expect(dirty.size).toBe(N * 2)
		expect(elapsed).toBeLessThan(50)
	})

	test('getIndependentSubgraphs returns single component for connected chain', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 2, 0)
		g.addFormula(b, [a], false)
		g.addFormula(c, [b], false)
		const subgraphs = g.getIndependentSubgraphs()
		expect(subgraphs.length).toBe(1)
		expect(subgraphs[0]?.length).toBe(2)
		expect(subgraphs[0]).toContain(b)
		expect(subgraphs[0]).toContain(c)
	})

	test('getIndependentSubgraphs returns multiple components for independent sheets', () => {
		const g = new DependencyGraph()
		const s1a = cellKey(0, 0, 0)
		const s1b = cellKey(0, 1, 0)
		const s2a = cellKey(1, 0, 0)
		const s2b = cellKey(1, 1, 0)
		g.addFormula(s1a, [], false)
		g.addFormula(s1b, [s1a], false)
		g.addFormula(s2a, [], false)
		g.addFormula(s2b, [s2a], false)
		const subgraphs = g.getIndependentSubgraphs()
		expect(subgraphs.length).toBe(2)
		const sizes = subgraphs.map((sg) => sg.length).sort((a, b) => a - b)
		expect(sizes).toEqual([2, 2])
	})

	test('getIndependentSubgraphs returns one component per isolated formula', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 1, 0)
		const c = cellKey(0, 2, 0)
		g.addFormula(a, [], false)
		g.addFormula(b, [], false)
		g.addFormula(c, [], false)
		const subgraphs = g.getIndependentSubgraphs()
		expect(subgraphs.length).toBe(3)
		expect(subgraphs.every((sg) => sg.length === 1)).toBe(true)
	})

	test('getIndependentSubgraphs empty graph returns empty', () => {
		const g = new DependencyGraph()
		expect(g.getIndependentSubgraphs()).toEqual([])
	})
})

describe('IntervalIndex', () => {
	test('queryBatch returns all matching formula keys', () => {
		const idx = new IntervalIndex()
		const k1 = cellKey(0, 10, 0)
		const k2 = cellKey(0, 11, 0)
		const k3 = cellKey(0, 12, 0)
		idx.insert(0, 5, 0, 0, k1)
		idx.insert(3, 8, 0, 0, k2)
		idx.insert(10, 15, 0, 0, k3)

		const cells = new Map<number, number[]>([[0, [4]]])
		const result = idx.queryBatch(cells)
		expect(result.sort()).toEqual([k1, k2].sort())
	})

	test('queryBatch with no matches returns empty', () => {
		const idx = new IntervalIndex()
		idx.insert(0, 5, 0, 0, cellKey(0, 10, 0))
		const cells = new Map<number, number[]>([[0, [6, 7, 8]]])
		expect(idx.queryBatch(cells)).toEqual([])
	})

	test('queryBatch deduplicates formula keys', () => {
		const idx = new IntervalIndex()
		const fk = cellKey(0, 10, 0)
		idx.insert(0, 5, 0, 0, fk)
		idx.insert(3, 8, 0, 0, fk)
		const cells = new Map<number, number[]>([[0, [4]]])
		expect(idx.queryBatch(cells)).toEqual([fk])
	})

	test('queryBatch handles multi-column ranges', () => {
		const idx = new IntervalIndex()
		const fk = cellKey(0, 10, 0)
		idx.insert(0, 5, 0, 2, fk)
		const noMatch = new Map<number, number[]>([[3, [3]]])
		expect(idx.queryBatch(noMatch)).toEqual([])
		const match = new Map<number, number[]>([[1, [3]]])
		expect(idx.queryBatch(match)).toEqual([fk])
	})

	test('queryBatch matches wide ranges from the last changed column', () => {
		const idx = new IntervalIndex()
		const fk = cellKey(0, 10, 0)
		idx.insert(0, 5, 0, 16_383, fk)

		const cells = new Map<number, number[]>([[16_383, [3]]])

		expect(idx.queryBatch(cells)).toEqual([fk])
	})

	test('query skips columns outside all indexed ranges', () => {
		const idx = new IntervalIndex()
		const fk = cellKey(0, 10, 0)
		idx.insert(0, 10, 2, 2, fk)

		expect(idx.query(5, 1)).toEqual([])
		expect(idx.query(5, 3)).toEqual([])
		expect(idx.query(5, 2)).toEqual([fk])
	})

	test('query column bounds update after removing a formula', () => {
		const idx = new IntervalIndex()
		const left = cellKey(0, 10, 0)
		const right = cellKey(0, 11, 0)
		idx.insert(0, 10, 0, 0, left)
		idx.insert(0, 10, 2, 2, right)

		idx.remove(left)

		expect(idx.query(5, 0)).toEqual([])
		expect(idx.query(5, 2)).toEqual([right])
	})

	test('query row bounds update after removing a formula', () => {
		const idx = new IntervalIndex()
		const top = cellKey(0, 10, 0)
		const bottom = cellKey(0, 11, 0)
		idx.insert(0, 10, 0, 0, top)
		idx.insert(20, 30, 0, 0, bottom)

		idx.remove(top)

		expect(idx.query(5, 0)).toEqual([])
		expect(idx.query(25, 0)).toEqual([bottom])
	})
})
