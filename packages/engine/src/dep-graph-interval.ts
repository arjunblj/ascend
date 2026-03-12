import type { CellKey } from './dep-graph.ts'

export class IntervalIndex {
	private entries: Array<{
		startRow: number
		endRow: number
		startCol: number
		endCol: number
		formulaKey: CellKey
	}> = []
	private sorted = true

	insert(
		startRow: number,
		endRow: number,
		startCol: number,
		endCol: number,
		formulaKey: CellKey,
	): void {
		this.entries.push({ startRow, endRow, startCol, endCol, formulaKey })
		this.sorted = false
	}

	query(row: number, col: number): CellKey[] {
		if (this.entries.length === 0) return []
		this.ensureSorted()
		let lo = 0
		let hi = this.entries.length
		while (lo < hi) {
			const mid = (lo + hi) >>> 1
			const midEntry = this.entries[mid]
			if (midEntry && midEntry.startRow > row) hi = mid
			else lo = mid + 1
		}
		const result: CellKey[] = []
		for (let i = 0; i < hi; i++) {
			const e = this.entries[i]
			if (e && e.endRow >= row && col >= e.startCol && col <= e.endCol) {
				result.push(e.formulaKey)
			}
		}
		return result
	}

	remove(formulaKey: CellKey): void {
		let write = 0
		for (let i = 0; i < this.entries.length; i++) {
			const entry = this.entries[i]
			if (entry && entry.formulaKey !== formulaKey) {
				this.entries[write++] = entry
			}
		}
		this.entries.length = write
	}

	get size(): number {
		return this.entries.length
	}

	private ensureSorted(): void {
		if (this.sorted) return
		this.entries.sort((a, b) => a.startRow - b.startRow)
		this.sorted = true
	}
}
