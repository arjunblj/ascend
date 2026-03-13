import { describe, expect, test } from 'bun:test'
import { ColumnSegmentTrees, SegmentTree } from './segment-tree.ts'

describe('SegmentTree', () => {
	test('build from values and query full range', () => {
		const tree = SegmentTree.fromValues([1, 2, 3, 4, 5])
		expect(tree.querySum(0, 4)).toBe(15)
		expect(tree.queryMin(0, 4)).toBe(1)
		expect(tree.queryMax(0, 4)).toBe(5)
	})

	test('query partial range', () => {
		const tree = SegmentTree.fromValues([10, 20, 30, 40, 50])
		expect(tree.querySum(1, 3)).toBe(90)
		expect(tree.queryMin(1, 3)).toBe(20)
		expect(tree.queryMax(1, 3)).toBe(40)
	})

	test('single element', () => {
		const tree = SegmentTree.fromValues([42])
		expect(tree.querySum(0, 0)).toBe(42)
		expect(tree.queryMin(0, 0)).toBe(42)
		expect(tree.queryMax(0, 0)).toBe(42)
	})

	test('point update', () => {
		const tree = SegmentTree.fromValues([1, 2, 3, 4, 5])
		tree.update(2, 100)
		expect(tree.querySum(0, 4)).toBe(112)
		expect(tree.queryMax(0, 4)).toBe(100)
		expect(tree.queryMin(0, 4)).toBe(1)
	})

	test('update then query partial', () => {
		const tree = SegmentTree.fromValues([5, 10, 15, 20])
		tree.update(0, 100)
		expect(tree.querySum(0, 1)).toBe(110)
		expect(tree.queryMin(0, 1)).toBe(10)
		expect(tree.queryMax(0, 1)).toBe(100)
	})

	test('empty range query returns identity', () => {
		const tree = SegmentTree.fromValues([1, 2, 3])
		expect(tree.querySum(5, 10)).toBe(0)
		expect(tree.queryMin(5, 10)).toBe(Infinity)
		expect(tree.queryMax(5, 10)).toBe(-Infinity)
	})

	test('out of bounds indices clamped', () => {
		const tree = SegmentTree.fromValues([10, 20, 30])
		expect(tree.querySum(-1, 1)).toBe(30)
		expect(tree.querySum(1, 100)).toBe(50)
	})

	test('update out of bounds is no-op', () => {
		const tree = SegmentTree.fromValues([1, 2, 3])
		tree.update(-1, 999)
		tree.update(10, 999)
		expect(tree.querySum(0, 2)).toBe(6)
	})

	test('size property', () => {
		const tree = SegmentTree.fromValues([1, 2, 3, 4])
		expect(tree.size).toBe(4)
	})

	test('constructor creates zero-initialized tree', () => {
		const tree = new SegmentTree(4)
		expect(tree.querySum(0, 3)).toBe(0)
		tree.update(1, 10)
		expect(tree.querySum(0, 3)).toBe(10)
		expect(tree.queryMax(0, 3)).toBe(10)
	})

	test('negative values', () => {
		const tree = SegmentTree.fromValues([-5, -3, -1, 2, 4])
		expect(tree.querySum(0, 4)).toBe(-3)
		expect(tree.queryMin(0, 4)).toBe(-5)
		expect(tree.queryMax(0, 4)).toBe(4)
	})

	test('performance: 100k elements random updates and queries', () => {
		const n = 100_000
		const values = new Float64Array(n)
		for (let i = 0; i < n; i++) values[i] = Math.random() * 1000 - 500

		const tree = SegmentTree.fromValues(values)

		const bruteSum = (lo: number, hi: number) => {
			let s = 0
			for (let i = lo; i <= hi; i++) s += values[i] as number
			return s
		}

		for (let t = 0; t < 100; t++) {
			const lo = Math.floor(Math.random() * n)
			const hi = Math.min(lo + Math.floor(Math.random() * 1000), n - 1)
			expect(Math.abs(tree.querySum(lo, hi) - bruteSum(lo, hi))).toBeLessThan(1e-6)
		}

		for (let t = 0; t < 100; t++) {
			const idx = Math.floor(Math.random() * n)
			const newVal = Math.random() * 1000
			values[idx] = newVal
			tree.update(idx, newVal)
		}

		for (let t = 0; t < 50; t++) {
			const lo = Math.floor(Math.random() * n)
			const hi = Math.min(lo + Math.floor(Math.random() * 1000), n - 1)
			expect(Math.abs(tree.querySum(lo, hi) - bruteSum(lo, hi))).toBeLessThan(1e-6)
		}
	})

	test('performance: many point queries on 100k tree', () => {
		const n = 100_000
		const values = new Float64Array(n)
		for (let i = 0; i < n; i++) values[i] = i

		const tree = SegmentTree.fromValues(values)

		const start = performance.now()
		for (let i = 0; i < 10_000; i++) {
			const lo = Math.floor(Math.random() * n)
			const hi = Math.min(lo + 100, n - 1)
			tree.querySum(lo, hi)
			tree.queryMin(lo, hi)
			tree.queryMax(lo, hi)
		}
		const elapsed = performance.now() - start
		expect(elapsed).toBeLessThan(1000)
	})
})

describe('ColumnSegmentTrees', () => {
	test('update and query single column', () => {
		const trees = new ColumnSegmentTrees(10)
		trees.update(0, 0, 100)
		trees.update(1, 0, 200)
		trees.update(2, 0, 300)

		expect(trees.querySum(0, 0, 2)).toBe(600)
		expect(trees.queryMin(0, 0, 2)).toBe(100)
		expect(trees.queryMax(0, 0, 2)).toBe(300)
	})

	test('multi-column range query', () => {
		const trees = new ColumnSegmentTrees(5)
		trees.update(0, 0, 1)
		trees.update(0, 1, 2)
		trees.update(1, 0, 3)
		trees.update(1, 1, 4)

		expect(trees.querySumRange(0, 0, 1, 1)).toBe(10)
		expect(trees.queryMinRange(0, 0, 1, 1)).toBe(1)
		expect(trees.queryMaxRange(0, 0, 1, 1)).toBe(4)
	})

	test('query non-existent column returns identity', () => {
		const trees = new ColumnSegmentTrees(5)
		expect(trees.querySum(99, 0, 4)).toBe(0)
		expect(trees.queryMin(99, 0, 4)).toBe(Infinity)
		expect(trees.queryMax(99, 0, 4)).toBe(-Infinity)
	})

	test('has and clear', () => {
		const trees = new ColumnSegmentTrees(5)
		expect(trees.has(0)).toBe(false)
		trees.update(0, 0, 10)
		expect(trees.has(0)).toBe(true)
		trees.clear()
		expect(trees.has(0)).toBe(false)
	})

	test('getOrCreate returns same tree for same column', () => {
		const trees = new ColumnSegmentTrees(5)
		const t1 = trees.getOrCreate(3)
		const t2 = trees.getOrCreate(3)
		expect(t1).toBe(t2)
	})
})
