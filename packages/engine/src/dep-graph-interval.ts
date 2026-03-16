/**
 * IntervalIndex: range-to-formula dependency tracking for dirty propagation.
 *
 * Overlapping ranges: When a cell changes in A1:A50, both SUM(A1:A100) and
 * SUM(A1:A50) are correctly marked dirty — query(row,col) returns any formula
 * whose range contains that cell.
 *
 * aggregateRangeCache (formulas package): Caches exact range aggregates per
 * key `name:sheet:row:col:endRow:endCol`. Overlapping ranges (e.g. A1:A100
 * vs A1:A50) are separate entries; the cache does not derive subrange results
 * from a superset. Future optimization: when SUM(A1:A100) is cached, SUM(A1:A50)
 * could be computed from the cached iteration instead of from scratch.
 */
import type { CellKey } from './dep-graph.ts'

interface RangeEntry {
	startRow: number
	endRow: number
	startCol: number
	endCol: number
	formulaKey: CellKey
}

export class IntervalIndex {
	private entries: RangeEntry[] = []
	private formulaIndices = new Map<CellKey, Set<number>>()
	private subMax: Int32Array | null = null
	private dirty = true

	insert(
		startRow: number,
		endRow: number,
		startCol: number,
		endCol: number,
		formulaKey: CellKey,
	): void {
		const index = this.entries.length
		this.entries.push({ startRow, endRow, startCol, endCol, formulaKey })
		let indices = this.formulaIndices.get(formulaKey)
		if (!indices) {
			indices = new Set()
			this.formulaIndices.set(formulaKey, indices)
		}
		indices.add(index)
		this.dirty = true
	}

	query(row: number, col: number): CellKey[] {
		if (this.entries.length === 0) return []
		if (this.dirty) this.rebuild()
		const result: CellKey[] = []
		this.queryRange(0, this.entries.length - 1, row, col, result)
		return result
	}

	/**
	 * Batch query: find all formula keys whose range contains at least one of the
	 * given cells. Cells are grouped by column with sorted row arrays so each
	 * entry can be checked via binary search — O(M · C_avg · log N) total instead
	 * of O(N · (log M + k)) for N separate point queries.
	 */
	queryBatch(cellsByCol: ReadonlyMap<number, readonly number[]>): CellKey[] {
		if (this.entries.length === 0) return []
		if (this.dirty) this.rebuild()
		const result: CellKey[] = []
		const seen = new Set<CellKey>()
		for (let i = 0; i < this.entries.length; i++) {
			const e = this.entries[i] as RangeEntry
			if (seen.has(e.formulaKey)) continue
			let found = false
			for (let c = e.startCol; c <= e.endCol && !found; c++) {
				const sortedRows = cellsByCol.get(c)
				if (!sortedRows || sortedRows.length === 0) continue
				found = hasRowInRange(sortedRows, e.startRow, e.endRow)
			}
			if (found) {
				seen.add(e.formulaKey)
				result.push(e.formulaKey)
			}
		}
		return result
	}

	remove(formulaKey: CellKey): void {
		const indices = this.formulaIndices.get(formulaKey)
		if (!indices || indices.size === 0) return
		while (indices.size > 0) {
			const index = indices.values().next().value
			if (index === undefined) break
			this.removeAt(index)
		}
		this.formulaIndices.delete(formulaKey)
		this.dirty = true
	}

	get size(): number {
		return this.entries.length
	}

	private rebuild(): void {
		this.entries.sort((a, b) => a.startRow - b.startRow)
		const n = this.entries.length
		this.subMax = new Int32Array(n)
		this.formulaIndices.clear()
		for (let i = 0; i < n; i++) {
			const entry = this.entries[i] as RangeEntry
			let indices = this.formulaIndices.get(entry.formulaKey)
			if (!indices) {
				indices = new Set()
				this.formulaIndices.set(entry.formulaKey, indices)
			}
			indices.add(i)
		}
		this.computeSubMax(0, n - 1)
		this.dirty = false
	}

	private removeAt(index: number): void {
		const lastIndex = this.entries.length - 1
		if (index < 0 || index > lastIndex) return
		const removed = this.entries[index] as RangeEntry
		const removedIndices = this.formulaIndices.get(removed.formulaKey)
		removedIndices?.delete(index)
		if (index !== lastIndex) {
			const lastEntry = this.entries[lastIndex] as RangeEntry
			this.entries[index] = lastEntry
			const lastIndices = this.formulaIndices.get(lastEntry.formulaKey)
			lastIndices?.delete(lastIndex)
			lastIndices?.add(index)
		}
		this.entries.pop()
		if (removedIndices && removedIndices.size === 0) {
			this.formulaIndices.delete(removed.formulaKey)
		}
	}

	private computeSubMax(lo: number, hi: number): number {
		if (lo > hi) return -1
		const mid = (lo + hi) >>> 1
		let max = (this.entries[mid] as RangeEntry).endRow
		if (lo < mid) {
			const leftMax = this.computeSubMax(lo, mid - 1)
			if (leftMax > max) max = leftMax
		}
		if (mid < hi) {
			const rightMax = this.computeSubMax(mid + 1, hi)
			if (rightMax > max) max = rightMax
		}
		;(this.subMax as Int32Array)[mid] = max
		return max
	}

	private queryRange(lo: number, hi: number, row: number, col: number, result: CellKey[]): void {
		if (lo > hi) return
		const mid = (lo + hi) >>> 1
		if (((this.subMax as Int32Array)[mid] as number) < row) return
		if (lo < mid) this.queryRange(lo, mid - 1, row, col, result)
		const e = this.entries[mid] as RangeEntry
		if (e.startRow > row) return
		if (e.endRow >= row && col >= e.startCol && col <= e.endCol) {
			result.push(e.formulaKey)
		}
		if (mid < hi) this.queryRange(mid + 1, hi, row, col, result)
	}
}

function hasRowInRange(sortedRows: readonly number[], startRow: number, endRow: number): boolean {
	let lo = 0
	let hi = sortedRows.length - 1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		if ((sortedRows[mid] as number) < startRow) lo = mid + 1
		else hi = mid - 1
	}
	return lo < sortedRows.length && (sortedRows[lo] as number) <= endRow
}
