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
	private subMax: Int32Array | null = null
	private dirty = true

	insert(
		startRow: number,
		endRow: number,
		startCol: number,
		endCol: number,
		formulaKey: CellKey,
	): void {
		this.entries.push({ startRow, endRow, startCol, endCol, formulaKey })
		this.dirty = true
	}

	query(row: number, col: number): CellKey[] {
		if (this.entries.length === 0) return []
		if (this.dirty) this.rebuild()
		const result: CellKey[] = []
		this.queryRange(0, this.entries.length - 1, row, col, result)
		return result
	}

	remove(formulaKey: CellKey): void {
		let write = 0
		for (let i = 0; i < this.entries.length; i++) {
			const entry = this.entries[i] as RangeEntry
			if (entry.formulaKey !== formulaKey) {
				this.entries[write++] = entry
			}
		}
		if (write < this.entries.length) {
			this.entries.length = write
			this.dirty = true
		}
	}

	get size(): number {
		return this.entries.length
	}

	private rebuild(): void {
		this.entries.sort((a, b) => a.startRow - b.startRow)
		const n = this.entries.length
		this.subMax = new Int32Array(n)
		this.computeSubMax(0, n - 1)
		this.dirty = false
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
