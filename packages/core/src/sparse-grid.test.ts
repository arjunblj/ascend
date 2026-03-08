import { describe, expect, test } from 'bun:test'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import type { StyleId } from './ids.ts'
import type { Cell } from './sparse-grid.ts'
import { SparseGrid } from './sparse-grid.ts'

const S0 = 0 as StyleId

function makeCell(value: Cell['value'], formula: string | null = null): Cell {
	return { value, formula, styleId: S0 }
}

describe('SparseGrid', () => {
	test('get returns undefined for empty cell', () => {
		const grid = new SparseGrid()
		expect(grid.get(0, 0)).toBeUndefined()
	})

	test('set and get', () => {
		const grid = new SparseGrid()
		const cell = makeCell(numberValue(42))
		grid.set(0, 0, cell)
		expect(grid.get(0, 0)).toEqual(cell)
	})

	test('overwrite existing cell', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(0, 0, makeCell(numberValue(2)))
		expect(grid.get(0, 0)?.value).toEqual(numberValue(2))
	})

	test('delete existing cell returns true', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		expect(grid.delete(0, 0)).toBe(true)
		expect(grid.get(0, 0)).toBeUndefined()
	})

	test('delete non-existent cell returns false', () => {
		const grid = new SparseGrid()
		expect(grid.delete(0, 0)).toBe(false)
	})

	test('cellCount tracks insertions and deletions', () => {
		const grid = new SparseGrid()
		expect(grid.cellCount()).toBe(0)
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(1, 1, makeCell(numberValue(2)))
		expect(grid.cellCount()).toBe(2)
		grid.delete(0, 0)
		expect(grid.cellCount()).toBe(1)
	})

	test('usedRange returns null for empty grid', () => {
		const grid = new SparseGrid()
		expect(grid.usedRange()).toBeNull()
	})

	test('usedRange computes bounding box', () => {
		const grid = new SparseGrid()
		grid.set(1, 2, makeCell(numberValue(1)))
		grid.set(5, 0, makeCell(numberValue(2)))
		grid.set(3, 4, makeCell(numberValue(3)))
		expect(grid.usedRange()).toEqual({
			start: { row: 1, col: 0 },
			end: { row: 5, col: 4 },
		})
	})

	test('iterate yields all cells', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(2, 3, makeCell(stringValue('hello')))
		grid.set(1, 1, makeCell(EMPTY))

		const entries = [...grid.iterate()]
		expect(entries).toHaveLength(3)

		const coords = entries.map(([r, c]) => [r, c])
		expect(coords).toContainEqual([0, 0])
		expect(coords).toContainEqual([2, 3])
		expect(coords).toContainEqual([1, 1])
	})

	test('getRange yields only cells within bounds', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(0, 1, makeCell(numberValue(2)))
		grid.set(1, 0, makeCell(numberValue(3)))
		grid.set(5, 5, makeCell(numberValue(99)))

		const range = { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }
		const entries = [...grid.getRange(range)]
		expect(entries).toHaveLength(3)

		const coords = entries.map(([r, c]) => [r, c])
		expect(coords).not.toContainEqual([5, 5])
	})

	test('clear removes all cells', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(1, 1, makeCell(numberValue(2)))
		grid.clear()
		expect(grid.cellCount()).toBe(0)
		expect(grid.get(0, 0)).toBeUndefined()
	})

	test('handles high row and column indices', () => {
		const grid = new SparseGrid()
		const cell = makeCell(stringValue('far'))
		grid.set(1048575, 16383, cell)
		expect(grid.get(1048575, 16383)).toEqual(cell)
		expect(grid.cellCount()).toBe(1)
	})

	test('cells with formulas', () => {
		const grid = new SparseGrid()
		const cell = makeCell(numberValue(10), 'SUM(A1:A5)')
		grid.set(0, 5, cell)
		const retrieved = grid.get(0, 5)
		expect(retrieved?.formula).toBe('SUM(A1:A5)')
		expect(retrieved?.value).toEqual(numberValue(10))
	})
})
