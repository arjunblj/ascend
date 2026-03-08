import { describe, expect, test } from 'bun:test'
import { cellKey, DependencyGraph } from './dep-graph.ts'

describe('DependencyGraph', () => {
	test('empty graph returns no dependents', () => {
		const g = new DependencyGraph()
		expect(g.getDependents('0:0:0')).toEqual([])
	})

	test('empty graph returns no precedents', () => {
		const g = new DependencyGraph()
		expect(g.getPrecedents('0:0:0')).toEqual([])
	})

	test('addFormula tracks precedents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const b = cellKey(0, 0, 1)
		const c = cellKey(0, 0, 2)
		g.addFormula(c, [a, b], false)
		expect(g.getPrecedents(c).sort()).toEqual([a, b].sort())
	})

	test('addFormula tracks dependents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const c = cellKey(0, 0, 2)
		g.addFormula(c, [a], false)
		expect(g.getDependents(a)).toEqual([c])
	})

	test('removeFormula cleans up dependents', () => {
		const g = new DependencyGraph()
		const a = cellKey(0, 0, 0)
		const c = cellKey(0, 0, 2)
		g.addFormula(c, [a], false)
		g.removeFormula(c)
		expect(g.getDependents(a)).toEqual([])
		expect(g.getPrecedents(c)).toEqual([])
	})

	test('removing non-existent formula is a no-op', () => {
		const g = new DependencyGraph()
		g.removeFormula('0:99:99')
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

	test('cellKey format', () => {
		expect(cellKey(0, 5, 2)).toBe('0:5:2')
		expect(cellKey(1, 0, 0)).toBe('1:0:0')
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
})
