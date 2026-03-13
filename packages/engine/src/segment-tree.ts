export class SegmentTree {
	private readonly n: number
	private readonly sumTree: Float64Array
	private readonly minTree: Float64Array
	private readonly maxTree: Float64Array

	constructor(capacity: number) {
		this.n = capacity
		const size = 4 * capacity
		this.sumTree = new Float64Array(size)
		this.minTree = new Float64Array(size).fill(Infinity)
		this.maxTree = new Float64Array(size).fill(-Infinity)
	}

	static fromValues(values: ArrayLike<number>): SegmentTree {
		const tree = new SegmentTree(values.length)
		if (values.length > 0) {
			tree._build(values, 1, 0, values.length - 1)
		}
		return tree
	}

	private _build(values: ArrayLike<number>, node: number, start: number, end: number): void {
		if (start === end) {
			const v = values[start] as number
			this.sumTree[node] = v
			this.minTree[node] = v
			this.maxTree[node] = v
			return
		}
		const mid = (start + end) >> 1
		const left = node << 1
		const right = left | 1
		this._build(values, left, start, mid)
		this._build(values, right, mid + 1, end)
		this.sumTree[node] = (this.sumTree[left] as number) + (this.sumTree[right] as number)
		this.minTree[node] = Math.min(this.minTree[left] as number, this.minTree[right] as number)
		this.maxTree[node] = Math.max(this.maxTree[left] as number, this.maxTree[right] as number)
	}

	update(index: number, value: number): void {
		if (index < 0 || index >= this.n) return
		this._update(1, 0, this.n - 1, index, value)
	}

	private _update(node: number, start: number, end: number, index: number, value: number): void {
		if (start === end) {
			this.sumTree[node] = value
			this.minTree[node] = value
			this.maxTree[node] = value
			return
		}
		const mid = (start + end) >> 1
		const left = node << 1
		const right = left | 1
		if (index <= mid) {
			this._update(left, start, mid, index, value)
		} else {
			this._update(right, mid + 1, end, index, value)
		}
		this.sumTree[node] = (this.sumTree[left] as number) + (this.sumTree[right] as number)
		this.minTree[node] = Math.min(this.minTree[left] as number, this.minTree[right] as number)
		this.maxTree[node] = Math.max(this.maxTree[left] as number, this.maxTree[right] as number)
	}

	querySum(lo: number, hi: number): number {
		if (lo > hi || lo >= this.n || hi < 0) return 0
		lo = Math.max(lo, 0)
		hi = Math.min(hi, this.n - 1)
		return this._querySum(1, 0, this.n - 1, lo, hi)
	}

	private _querySum(node: number, start: number, end: number, lo: number, hi: number): number {
		if (lo > end || hi < start) return 0
		if (lo <= start && end <= hi) return this.sumTree[node] as number
		const mid = (start + end) >> 1
		const left = node << 1
		const right = left | 1
		return this._querySum(left, start, mid, lo, hi) + this._querySum(right, mid + 1, end, lo, hi)
	}

	queryMin(lo: number, hi: number): number {
		if (lo > hi || lo >= this.n || hi < 0) return Infinity
		lo = Math.max(lo, 0)
		hi = Math.min(hi, this.n - 1)
		return this._queryMin(1, 0, this.n - 1, lo, hi)
	}

	private _queryMin(node: number, start: number, end: number, lo: number, hi: number): number {
		if (lo > end || hi < start) return Infinity
		if (lo <= start && end <= hi) return this.minTree[node] as number
		const mid = (start + end) >> 1
		const left = node << 1
		const right = left | 1
		return Math.min(
			this._queryMin(left, start, mid, lo, hi),
			this._queryMin(right, mid + 1, end, lo, hi),
		)
	}

	queryMax(lo: number, hi: number): number {
		if (lo > hi || lo >= this.n || hi < 0) return -Infinity
		lo = Math.max(lo, 0)
		hi = Math.min(hi, this.n - 1)
		return this._queryMax(1, 0, this.n - 1, lo, hi)
	}

	private _queryMax(node: number, start: number, end: number, lo: number, hi: number): number {
		if (lo > end || hi < start) return -Infinity
		if (lo <= start && end <= hi) return this.maxTree[node] as number
		const mid = (start + end) >> 1
		const left = node << 1
		const right = left | 1
		return Math.max(
			this._queryMax(left, start, mid, lo, hi),
			this._queryMax(right, mid + 1, end, lo, hi),
		)
	}

	get size(): number {
		return this.n
	}
}

export class ColumnSegmentTrees {
	private readonly trees = new Map<number, SegmentTree>()
	private readonly capacity: number

	constructor(rowCapacity: number) {
		this.capacity = rowCapacity
	}

	getOrCreate(col: number): SegmentTree {
		let tree = this.trees.get(col)
		if (!tree) {
			tree = new SegmentTree(this.capacity)
			this.trees.set(col, tree)
		}
		return tree
	}

	get(col: number): SegmentTree | undefined {
		return this.trees.get(col)
	}

	update(row: number, col: number, value: number): void {
		this.getOrCreate(col).update(row, value)
	}

	querySum(col: number, startRow: number, endRow: number): number {
		const tree = this.trees.get(col)
		return tree ? tree.querySum(startRow, endRow) : 0
	}

	queryMin(col: number, startRow: number, endRow: number): number {
		const tree = this.trees.get(col)
		return tree ? tree.queryMin(startRow, endRow) : Infinity
	}

	queryMax(col: number, startRow: number, endRow: number): number {
		const tree = this.trees.get(col)
		return tree ? tree.queryMax(startRow, endRow) : -Infinity
	}

	querySumRange(startRow: number, startCol: number, endRow: number, endCol: number): number {
		let sum = 0
		for (let c = startCol; c <= endCol; c++) {
			sum += this.querySum(c, startRow, endRow)
		}
		return sum
	}

	queryMinRange(startRow: number, startCol: number, endRow: number, endCol: number): number {
		let min = Infinity
		for (let c = startCol; c <= endCol; c++) {
			const v = this.queryMin(c, startRow, endRow)
			if (v < min) min = v
		}
		return min
	}

	queryMaxRange(startRow: number, startCol: number, endRow: number, endCol: number): number {
		let max = -Infinity
		for (let c = startCol; c <= endCol; c++) {
			const v = this.queryMax(c, startRow, endRow)
			if (v > max) max = v
		}
		return max
	}

	has(col: number): boolean {
		return this.trees.has(col)
	}

	clear(): void {
		this.trees.clear()
	}
}
