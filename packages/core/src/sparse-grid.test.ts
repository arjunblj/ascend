import { describe, expect, test } from 'bun:test'
import {
	type CellValue,
	EMPTY,
	numberValue,
	type ScalarCellValue,
	stringValue,
} from '@ascend/schema'
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

	test('setStringResolved stores plain strings directly', () => {
		const grid = new SparseGrid()
		grid.setStringResolved(0, 0, 'hello', null, S0)
		expect(grid.get(0, 0)).toEqual(makeCell(stringValue('hello')))
		expect(grid.readString(0, 0)).toBe('hello')
	})

	test('plain setters store values while preserving grid bookkeeping', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, { ...makeCell(numberValue(1), 'A1+1'), formulaInfo: { kind: 'array' } })
		expect(grid.cellCount()).toBe(1)
		expect(grid.formulaCellCount()).toBe(1)
		expect(grid.formulaInfoCellCount()).toBe(1)

		grid.setPlainNumber(0, 0, 2)
		grid.setPlainString(1, 1, 'hello')

		expect(grid.cellCount()).toBe(2)
		expect(grid.formulaCellCount()).toBe(0)
		expect(grid.formulaInfoCellCount()).toBe(0)
		expect(grid.readNumber(0, 0)).toBe(2)
		expect(grid.readString(1, 1)).toBe('hello')
		expect(grid.usedRange()).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 1, col: 1 },
		})
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

	test('formulaCellCount tracks formula mutations', () => {
		const grid = new SparseGrid()
		expect(grid.formulaCellCount()).toBe(0)
		expect(grid.formulaInfoCellCount()).toBe(0)
		grid.set(0, 0, makeCell(numberValue(1), 'A1+1'))
		grid.set(0, 1, makeCell(numberValue(2)))
		expect(grid.formulaCellCount()).toBe(1)
		expect(grid.formulaInfoCellCount()).toBe(0)
		grid.set(0, 1, { ...makeCell(numberValue(2)), formulaInfo: { kind: 'array', ref: 'B1' } })
		expect(grid.formulaCellCount()).toBe(2)
		expect(grid.formulaInfoCellCount()).toBe(1)
		grid.clearFormulaInfo(0, 1)
		expect(grid.formulaCellCount()).toBe(1)
		expect(grid.formulaInfoCellCount()).toBe(0)
		grid.set(0, 0, makeCell(numberValue(3)))
		expect(grid.formulaCellCount()).toBe(0)
		expect(grid.formulaInfoCellCount()).toBe(0)
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

	test('forEachValueInRange calls fn with value, row, col for cells in range', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(0, 1, makeCell(numberValue(2)))
		grid.set(1, 0, makeCell(numberValue(3)))
		grid.set(5, 5, makeCell(numberValue(99)))

		const collected: Array<[CellValue, number, number]> = []
		grid.forEachValueInRange(0, 0, 1, 1, (value, row, col) => collected.push([value, row, col]))
		expect(collected).toHaveLength(3)
		expect(collected.map(([, r, c]) => [r, c])).not.toContainEqual([5, 5])
	})

	test('forEachRow calls fn with row and Map of col->value', () => {
		const grid = new SparseGrid()
		grid.set(2, 3, makeCell(stringValue('hello')))
		grid.set(0, 1, makeCell(numberValue(2)))
		grid.set(0, 0, makeCell(numberValue(1)))

		const rows: Array<[number, Map<number, unknown>]> = []
		grid.forEachRow((row, cells) => {
			rows.push([row, new Map(cells)])
		})
		expect(rows).toHaveLength(2)
		expect(rows[0][0]).toBe(0)
		expect(rows[0][1].get(0)).toEqual(numberValue(1))
		expect(rows[0][1].get(1)).toEqual(numberValue(2))
		expect(rows[1][0]).toBe(2)
		expect(rows[1][1].get(3)).toEqual(stringValue('hello'))
	})

	test('iterateRows yields sorted row-major cells', () => {
		const grid = new SparseGrid()
		grid.set(2, 3, makeCell(stringValue('hello')))
		grid.set(0, 1, makeCell(numberValue(2)))
		grid.set(0, 0, makeCell(numberValue(1)))

		expect([...grid.iterateRows()]).toEqual([
			[
				0,
				[
					[0, makeCell(numberValue(1))],
					[1, makeCell(numberValue(2))],
				],
			],
			[2, [[3, makeCell(stringValue('hello'))]]],
		])
	})

	test('iterateRows stays sorted after chunk insertions and deletions', () => {
		const grid = new SparseGrid()
		grid.set(0, 70, makeCell(numberValue(70)))
		grid.set(0, 1, makeCell(numberValue(1)))
		grid.set(64, 0, makeCell(numberValue(64)))
		expect([...grid.iterateRows()]).toEqual([
			[
				0,
				[
					[1, makeCell(numberValue(1))],
					[70, makeCell(numberValue(70))],
				],
			],
			[64, [[0, makeCell(numberValue(64))]]],
		])

		grid.delete(0, 1)
		grid.set(0, 130, makeCell(numberValue(130)))
		expect([...grid.iterateRows()]).toEqual([
			[
				0,
				[
					[70, makeCell(numberValue(70))],
					[130, makeCell(numberValue(130))],
				],
			],
			[64, [[0, makeCell(numberValue(64))]]],
		])
	})

	test('iterateRowsInRange yields only populated cells within row and column bounds', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(0, 5, makeCell(numberValue(2)))
		grid.set(3, 1, makeCell(numberValue(3)))
		grid.set(4, 2, makeCell(numberValue(4)))

		expect([
			...grid.iterateRowsInRange({
				start: { row: 0, col: 1 },
				end: { row: 3, col: 5 },
			}),
		]).toEqual([
			[0, [[5, makeCell(numberValue(2))]]],
			[3, [[1, makeCell(numberValue(3))]]],
		])
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

	test('set and get cell at high row number (1,048,575)', () => {
		const grid = new SparseGrid()
		const cell = makeCell(numberValue(42))
		grid.set(1_048_575, 0, cell)
		expect(grid.get(1_048_575, 0)).toEqual(cell)
		expect(grid.getValue(1_048_575, 0)).toEqual(numberValue(42))
	})

	test('set and get cell at high column number (16,383)', () => {
		const grid = new SparseGrid()
		const cell = makeCell(stringValue('edge'))
		grid.set(0, 16_383, cell)
		expect(grid.get(0, 16_383)).toEqual(cell)
		expect(grid.getValue(0, 16_383)).toEqual(stringValue('edge'))
	})

	test('usedRange with single cell', () => {
		const grid = new SparseGrid()
		grid.set(7, 13, makeCell(numberValue(1)))
		expect(grid.usedRange()).toEqual({
			start: { row: 7, col: 13 },
			end: { row: 7, col: 13 },
		})
	})

	test('empty grid operations', () => {
		const grid = new SparseGrid()
		expect(grid.usedRange()).toBeNull()
		expect(grid.get(0, 0)).toBeUndefined()
		expect(grid.getValue(0, 0)).toBeUndefined()
		expect(grid.has(0, 0)).toBe(false)
		expect(grid.delete(0, 0)).toBe(false)
		expect(grid.cellCount()).toBe(0)
		expect([...grid.iterate()]).toHaveLength(0)
		expect([...grid.iterateRows()]).toHaveLength(0)
	})

	test('cells with formulas', () => {
		const grid = new SparseGrid()
		const cell = makeCell(numberValue(10), 'SUM(A1:A5)')
		grid.set(0, 5, cell)
		const retrieved = grid.get(0, 5)
		expect(retrieved?.formula).toBe('SUM(A1:A5)')
		expect(retrieved?.value).toEqual(numberValue(10))
	})

	test('insertRows shifts populated rows without rewriting cells', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(2, 0, makeCell(numberValue(2)))
		grid.insertRows(1, 2)
		expect(grid.get(0, 0)?.value).toEqual(numberValue(1))
		expect(grid.get(4, 0)?.value).toEqual(numberValue(2))
		expect(grid.get(2, 0)).toBeUndefined()
	})

	test('deleteRows removes and compacts shifted rows', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(1, 0, makeCell(numberValue(2)))
		grid.set(3, 0, makeCell(numberValue(3)))
		grid.deleteRows(1, 2)
		expect(grid.get(0, 0)?.value).toEqual(numberValue(1))
		expect(grid.get(1, 0)?.value).toEqual(numberValue(3))
		expect(grid.cellCount()).toBe(2)
	})

	test('insertCols shifts cells within each row', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(0, 2, makeCell(numberValue(2)))
		grid.insertCols(1, 2)
		expect(grid.get(0, 0)?.value).toEqual(numberValue(1))
		expect(grid.get(0, 4)?.value).toEqual(numberValue(2))
	})

	test('deleteCols removes deleted columns and compacts survivors', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(0, 1, makeCell(numberValue(2)))
		grid.set(0, 3, makeCell(numberValue(3)))
		grid.deleteCols(1, 2)
		expect(grid.get(0, 0)?.value).toEqual(numberValue(1))
		expect(grid.get(0, 1)?.value).toEqual(numberValue(3))
		expect(grid.cellCount()).toBe(2)
	})

	test('clone preserves cells without sharing mutation path', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(stringValue('alpha')))
		grid.set(1, 1, makeCell(numberValue(2), 'A1*2'))

		const clone = grid.clone()
		clone.set(0, 0, makeCell(stringValue('beta')))

		expect(grid.get(0, 0)?.value).toEqual(stringValue('alpha'))
		expect(clone.get(0, 0)?.value).toEqual(stringValue('beta'))
		expect(clone.get(1, 1)?.formula).toBe('A1*2')
	})

	test('clone isolates mutable array cell values', () => {
		const grid = new SparseGrid()
		grid.set(
			0,
			0,
			makeCell({
				kind: 'array',
				rows: [[numberValue(1) as ScalarCellValue], [numberValue(2) as ScalarCellValue]],
			}),
		)

		const clone = grid.clone()
		const clonedValue = clone.get(0, 0)?.value
		expect(clonedValue?.kind).toBe('array')
		if (!clonedValue || clonedValue.kind !== 'array') return

		;(clonedValue.rows[0] as ScalarCellValue[])[0] = numberValue(99) as ScalarCellValue
		const originalValue = grid.get(0, 0)?.value
		expect(originalValue?.kind).toBe('array')
		if (!originalValue || originalValue.kind !== 'array') return
		expect(originalValue.rows[0]?.[0]).toEqual(numberValue(1) as ScalarCellValue)
	})

	test('copy-on-write: clear on clone does not affect original', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(1, 1, makeCell(numberValue(2)))

		const clone = grid.clone()
		clone.clear()

		expect(grid.cellCount()).toBe(2)
		expect(grid.get(0, 0)?.value).toEqual(numberValue(1))
		expect(grid.get(1, 1)?.value).toEqual(numberValue(2))
		expect(clone.cellCount()).toBe(0)
	})

	test('copy-on-write: row insert on clone does not affect original', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(2, 0, makeCell(numberValue(3)))

		const clone = grid.clone()
		clone.insertRows(1, 1)

		expect(grid.get(0, 0)?.value).toEqual(numberValue(1))
		expect(grid.get(2, 0)?.value).toEqual(numberValue(3))
		expect(clone.get(0, 0)?.value).toEqual(numberValue(1))
		expect(clone.get(3, 0)?.value).toEqual(numberValue(3))
		expect(clone.get(2, 0)).toBeUndefined()
	})

	test('copy-on-write: mutating original after clone does not affect clone', () => {
		const grid = new SparseGrid()
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(1, 0, makeCell(numberValue(2)))

		const clone = grid.clone()
		grid.delete(0, 0)
		grid.set(1, 0, makeCell(stringValue('changed')))

		expect(clone.get(0, 0)?.value).toEqual(numberValue(1))
		expect(clone.get(1, 0)?.value).toEqual(numberValue(2))
		expect(grid.get(0, 0)).toBeUndefined()
		expect(grid.get(1, 0)?.value).toEqual(stringValue('changed'))
	})

	test('setExpectedDensity dense: new chunks are DenseChunk directly', () => {
		const grid = new SparseGrid()
		grid.setExpectedDensity('dense')
		grid.setResolved(0, 0, numberValue(1), null, S0)
		grid.setResolved(100, 100, numberValue(2), null, S0)
		expect(grid.getChunkKindAt(0, 0)).toBe('dense')
		expect(grid.getChunkKindAt(100, 100)).toBe('dense')
	})

	test('setExpectedDensity sparse: new chunks start as SparseChunk', () => {
		const grid = new SparseGrid()
		grid.setExpectedDensity('sparse')
		grid.setResolved(0, 0, numberValue(1), null, S0)
		expect(grid.getChunkKindAt(0, 0)).toBe('sparse')
	})

	test('setExpectedDensity sparse: SparseChunk upgrades to DenseChunk at threshold', () => {
		const grid = new SparseGrid()
		grid.setExpectedDensity('sparse')
		// 100 cells in a 10x10 block stay inside chunk (0,0) for supported chunk sizes.
		for (let i = 0; i < 100; i++) {
			grid.setResolved(Math.floor(i / 10), i % 10, numberValue(i), null, S0)
		}
		expect(grid.getChunkKindAt(0, 0)).toBe('dense')
	})

	test('setExpectedDensity auto: switches to dense when fill ratio > 50%', () => {
		const grid = new SparseGrid()
		grid.setExpectedDensity('auto')
		// Fill 4 chunks with >8192 total cells (4*4096*0.5) to trigger dense switch
		const cellsPerChunk = 2200
		for (let chunk = 0; chunk < 4; chunk++) {
			const baseRow = chunk * 64
			for (let i = 0; i < cellsPerChunk; i++) {
				const r = baseRow + (i >> 6)
				const c = i & 63
				grid.setResolved(r, c, numberValue(chunk * 1000 + i), null, S0)
			}
		}
		// 5th chunk - should be DenseChunk due to auto switch
		grid.setResolved(256, 0, numberValue(1), null, S0)
		expect(grid.getChunkKindAt(256, 0)).toBe('dense')
	})

	test('copy-on-write: reads work before any mutation on clone', () => {
		const grid = new SparseGrid()
		for (let i = 0; i < 50; i++) {
			grid.set(i, 0, makeCell(numberValue(i)))
		}
		const clone = grid.clone()
		expect(clone.cellCount()).toBe(50)
		expect(clone.get(25, 0)?.value).toEqual(numberValue(25))
		expect(clone.getValue(49, 0)).toEqual(numberValue(49))
		expect(clone.has(0, 0)).toBe(true)
		expect([...clone.iterate()]).toHaveLength(50)
	})

	test('copy-on-write: DenseChunk clone preserves all cells after mutation', () => {
		const grid = new SparseGrid()
		grid.setExpectedDensity('dense')
		grid.set(0, 0, makeCell(numberValue(1)))
		grid.set(1, 1, makeCell(numberValue(2)))
		grid.set(2, 2, makeCell(numberValue(3)))
		const clone = grid.clone()
		clone.set(0, 0, makeCell(stringValue('mutated')))
		expect(clone.cellCount()).toBe(3)
		expect([...clone.iterate()]).toHaveLength(3)
		expect(clone.get(0, 0)?.value).toEqual(stringValue('mutated'))
		expect(clone.get(1, 1)?.value).toEqual(numberValue(2))
		expect(clone.get(2, 2)?.value).toEqual(numberValue(3))
	})
})

describe('SparseGrid property-based', () => {
	function xorshift32(state: { s: number }): number {
		let s = state.s
		s ^= s << 13
		s ^= s >> 17
		s ^= s << 5
		state.s = s >>> 0
		return (s >>> 0) / 0xffffffff
	}

	function randInt(rng: { s: number }, min: number, max: number): number {
		return min + Math.floor(xorshift32(rng) * (max - min + 1))
	}

	test('COW invariant: clone a grid, mutate the clone, original unchanged', () => {
		const rng = { s: 314159 }
		for (let trial = 0; trial < 20; trial++) {
			const grid = new SparseGrid()
			const cellCount = randInt(rng, 5, 50)
			const origValues = new Map<string, CellValue>()
			for (let i = 0; i < cellCount; i++) {
				const r = randInt(rng, 0, 30)
				const c = randInt(rng, 0, 15)
				const val = numberValue(randInt(rng, -1000, 1000))
				grid.set(r, c, makeCell(val))
				origValues.set(`${r},${c}`, val)
			}

			const clone = grid.clone()

			for (let i = 0; i < 30; i++) {
				const r = randInt(rng, 0, 30)
				const c = randInt(rng, 0, 15)
				if (xorshift32(rng) < 0.6) {
					clone.set(r, c, makeCell(stringValue(`mutated_${i}`)))
				} else {
					clone.delete(r, c)
				}
			}

			for (const [key, val] of origValues) {
				const [r, c] = key.split(',').map(Number) as [number, number]
				expect(grid.getValue(r, c)).toEqual(val)
			}
		}
	})

	test('dense cache consistency: readValue matches getValue for all cells', () => {
		const rng = { s: 271828 }
		const grid = new SparseGrid()
		const positions: Array<[number, number]> = []

		for (let i = 0; i < 200; i++) {
			const r = randInt(rng, 0, 20)
			const c = randInt(rng, 0, 10)
			grid.set(r, c, makeCell(numberValue(i)))
			positions.push([r, c])
		}

		for (let r = 0; r <= 25; r++) {
			for (let c = 0; c <= 15; c++) {
				const fromGet = grid.getValue(r, c)
				const fromRead = grid.readValue(r, c)

				if (fromGet === undefined) {
					expect(fromRead).toEqual(EMPTY)
				} else {
					expect(fromRead).toEqual(fromGet)
				}
			}
		}
	})

	test('row index consistency: after insertRows/deleteRows, getRange returns correct cells', () => {
		const grid = new SparseGrid()
		for (let r = 0; r < 10; r++) {
			for (let c = 0; c < 3; c++) {
				grid.set(r, c, makeCell(numberValue(r * 100 + c)))
			}
		}

		grid.insertRows(5, 3)

		for (let r = 0; r < 5; r++) {
			for (let c = 0; c < 3; c++) {
				expect(grid.getValue(r, c)).toEqual(numberValue(r * 100 + c))
			}
		}
		for (let r = 5; r < 8; r++) {
			for (let c = 0; c < 3; c++) {
				expect(grid.has(r, c)).toBe(false)
			}
		}
		for (let r = 8; r < 13; r++) {
			for (let c = 0; c < 3; c++) {
				expect(grid.getValue(r, c)).toEqual(numberValue((r - 3) * 100 + c))
			}
		}

		const rangeEntries = [
			...grid.getRange({
				start: { row: 0, col: 0 },
				end: { row: 12, col: 2 },
			}),
		]
		expect(rangeEntries).toHaveLength(30)

		grid.deleteRows(5, 3)
		for (let r = 0; r < 10; r++) {
			for (let c = 0; c < 3; c++) {
				expect(grid.getValue(r, c)).toEqual(numberValue(r * 100 + c))
			}
		}
	})

	test('random mutation sequence: cellCount and usedRange are consistent', () => {
		const rng = { s: 161803 }
		const grid = new SparseGrid()
		const tracker = new Map<string, boolean>()

		for (let i = 0; i < 1000; i++) {
			const op = randInt(rng, 0, 3)
			switch (op) {
				case 0: {
					const r = randInt(rng, 0, 50)
					const c = randInt(rng, 0, 20)
					grid.set(r, c, makeCell(numberValue(i)))
					tracker.set(`${r},${c}`, true)
					break
				}
				case 1: {
					const r = randInt(rng, 0, 50)
					const c = randInt(rng, 0, 20)
					grid.delete(r, c)
					tracker.delete(`${r},${c}`)
					break
				}
				case 2: {
					if (grid.cellCount() > 0 && grid.cellCount() < 200) {
						const at = randInt(rng, 0, 30)
						const count = randInt(rng, 1, 3)
						grid.insertRows(at, count)
						const shifted = new Map<string, boolean>()
						for (const key of tracker.keys()) {
							const [r, c] = key.split(',').map(Number) as [number, number]
							const nr = r >= at ? r + count : r
							shifted.set(`${nr},${c}`, true)
						}
						tracker.clear()
						for (const [k, v] of shifted) tracker.set(k, v)
					}
					break
				}
				case 3: {
					if (grid.cellCount() > 0) {
						const at = randInt(rng, 0, 30)
						const count = randInt(rng, 1, 2)
						grid.deleteRows(at, count)
						const shifted = new Map<string, boolean>()
						for (const key of tracker.keys()) {
							const [r, c] = key.split(',').map(Number) as [number, number]
							if (r >= at && r < at + count) continue
							const nr = r >= at + count ? r - count : r
							shifted.set(`${nr},${c}`, true)
						}
						tracker.clear()
						for (const [k, v] of shifted) tracker.set(k, v)
					}
					break
				}
			}
		}

		expect(grid.cellCount()).toBe(tracker.size)

		const range = grid.usedRange()
		if (tracker.size === 0) {
			expect(range).toBeNull()
		} else {
			expect(range).not.toBeNull()
			if (!range) return
			for (const key of tracker.keys()) {
				const [r, c] = key.split(',').map(Number) as [number, number]
				expect(r).toBeGreaterThanOrEqual(range.start.row)
				expect(r).toBeLessThanOrEqual(range.end.row)
				expect(c).toBeGreaterThanOrEqual(range.start.col)
				expect(c).toBeLessThanOrEqual(range.end.col)
			}
		}
	})
})
