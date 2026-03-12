/**
 * DEP-1 Incremental dependency graph: addFormula/removeFormula already support
 * incremental updates. A full DEP-1 optimization would add batch add/remove and
 * avoid clearing _cachedDirtyIndex on every change. Would require tracking
 * dirty-index invalidation per key rather than full rebuild.
 *
 * DEP-2 Range dependency indexing: rangeDependents and getCachedDirtyIndex already
 * provide O(log n) range lookups. Further optimization: spatial index (R-tree or
 * interval tree) for range containment when many overlapping ranges exist.
 *
 * DEP-5 Formula group execution: implemented in EVAL-5. Shared formula groups
 * are identified by getSharedFormulaGroups in analysis and batch-evaluated in
 * the calc eval loop, eliminating per-cell overhead for group members.
 */
export type CellKey = number

export interface RangeDependency {
	readonly sheetIndex: number
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

const COL_FACTOR = 16_384
const SHEET_FACTOR = 16_384 * 1_048_576

export function cellKey(sheetIndex: number, row: number, col: number): CellKey {
	return sheetIndex * SHEET_FACTOR + row * COL_FACTOR + col
}

export interface CellCoords {
	sheetIndex: number
	row: number
	col: number
}

export function parseCellKey(key: CellKey): readonly [number, number, number] {
	const col = key % COL_FACTOR
	const remainder = (key - col) / COL_FACTOR
	const row = remainder % 1_048_576
	const sheetIndex = (remainder - row) / 1_048_576
	return [sheetIndex, row, col] as const
}

export function parseCellKeyInto(key: CellKey, out: CellCoords): void {
	const col = key % COL_FACTOR
	const remainder = (key - col) / COL_FACTOR
	const row = remainder % 1_048_576
	const sheetIndex = (remainder - row) / 1_048_576
	out.sheetIndex = sheetIndex
	out.row = row
	out.col = col
}

interface FormulaEntry {
	readonly dependsOn: ReadonlySet<CellKey>
	readonly rangeDeps: readonly RangeDependency[]
	readonly volatile: boolean
}

export class DependencyGraph {
	private readonly formulas = new Map<CellKey, FormulaEntry>()
	private readonly dependents = new Map<CellKey, Set<CellKey>>()
	private readonly rangeDependents = new Map<
		number,
		Array<{ range: RangeDependency; formulaKey: CellKey }>
	>()
	private readonly _rangeDepsSortedBySheet = new Map<number, boolean>()
	private _cachedDirtyIndex: {
		set: ReadonlySet<CellKey>
		size: number
		index: Map<number, Map<number, Set<number>>>
	} | null = null
	private _cachedEvalOrder: CellKey[] | null = null

	addFormula(
		key: CellKey,
		dependsOn: CellKey[],
		isVolatile: boolean,
		rangeDeps: readonly RangeDependency[] = [],
	): void {
		this.removeFormula(key)
		const aboveKey = key - COL_FACTOR
		const aboveEntry = aboveKey >= 0 ? this.formulas.get(aboveKey) : undefined
		let depSet: ReadonlySet<CellKey>
		let sharedRangeDeps = rangeDeps
		if (aboveEntry && setsMatchArray(aboveEntry.dependsOn, dependsOn)) {
			depSet = aboveEntry.dependsOn
			if (rangeDepsEqual(aboveEntry.rangeDeps, rangeDeps)) {
				sharedRangeDeps = aboveEntry.rangeDeps
			}
		} else {
			depSet = new Set(dependsOn)
		}
		this._cachedEvalOrder = null
		this.formulas.set(key, { dependsOn: depSet, rangeDeps: sharedRangeDeps, volatile: isVolatile })
		for (const dep of depSet) {
			let set = this.dependents.get(dep)
			if (!set) {
				set = new Set()
				this.dependents.set(dep, set)
			}
			set.add(key)
		}
		for (const range of rangeDeps) {
			let arr = this.rangeDependents.get(range.sheetIndex)
			if (!arr) {
				arr = []
				this.rangeDependents.set(range.sheetIndex, arr)
			}
			arr.push({ range, formulaKey: key })
			this._rangeDepsSortedBySheet.set(range.sheetIndex, false)
		}
	}

	removeFormula(key: CellKey): void {
		const entry = this.formulas.get(key)
		if (!entry) return
		this._cachedEvalOrder = null
		for (const dep of entry.dependsOn) {
			const set = this.dependents.get(dep)
			if (set) {
				set.delete(key)
				if (set.size === 0) this.dependents.delete(dep)
			}
		}
		for (const range of entry.rangeDeps) {
			const arr = this.rangeDependents.get(range.sheetIndex)
			if (!arr) continue
			const filtered = arr.filter((candidate) => candidate.formulaKey !== key)
			if (filtered.length === 0) {
				this.rangeDependents.delete(range.sheetIndex)
				this._rangeDepsSortedBySheet.delete(range.sheetIndex)
			} else if (filtered.length !== arr.length) {
				arr.length = 0
				arr.push(...filtered)
				this._rangeDepsSortedBySheet.set(range.sheetIndex, false)
			}
		}
		this.formulas.delete(key)
	}

	/** Precedents as cells and ranges; ranges are not expanded to avoid O(n) for large refs like SUM(A1:A10000). */
	getPrecedents(key: CellKey): { cells: CellKey[]; ranges: RangeDependency[] } {
		const entry = this.formulas.get(key)
		if (!entry) return { cells: [], ranges: [] }
		return {
			cells: [...entry.dependsOn],
			ranges: [...entry.rangeDeps],
		}
	}

	getDependents(key: CellKey): CellKey[] {
		const direct = this.dependents.get(key)
		const result = direct ? new Set(direct) : new Set<CellKey>()
		const [sheetIndex, row, col] = parseCellKey(key)
		const entries = this.rangeDependents.get(sheetIndex)
		if (entries && entries.length > 0) {
			this.ensureSorted(sheetIndex, entries)
			let lo = 0
			let hi = entries.length
			while (lo < hi) {
				const mid = (lo + hi) >>> 1
				const midEntry = entries[mid]
				if (midEntry && midEntry.range.startRow > row) hi = mid
				else lo = mid + 1
			}
			for (let i = 0; i < hi; i++) {
				const entry = entries[i]
				if (entry && entry.range.endRow >= row && containsCell(entry.range, row, col)) {
					result.add(entry.formulaKey)
				}
			}
		}
		return [...result]
	}

	getVolatiles(): CellKey[] {
		const result: CellKey[] = []
		for (const [key, entry] of this.formulas) {
			if (entry.volatile) result.push(key)
		}
		return result
	}

	getDirtySet(changedKeys: CellKey[]): Set<CellKey> {
		const dirty = new Set<CellKey>()
		const queue = [...changedKeys]
		while (queue.length > 0) {
			const current = queue.pop()
			if (current === undefined || dirty.has(current)) continue
			dirty.add(current)
			for (const dep of this.getDependents(current)) {
				if (!dirty.has(dep)) queue.push(dep)
			}
		}
		return dirty
	}

	getEvalOrder(dirtySet: Set<CellKey>): CellKey[] {
		if (this._cachedEvalOrder === null) {
			const allFormulaKeys = new Set(this.formulas.keys())
			this._cachedEvalOrder = this._computeEvalOrder(allFormulaKeys)
		}
		const formulaOrder = this._cachedEvalOrder.filter((k) => dirtySet.has(k))
		const nonFormulaKeys = [...dirtySet].filter((k) => !this.formulas.has(k))
		return [...nonFormulaKeys, ...formulaOrder]
	}

	private _computeEvalOrder(dirtySet: Set<CellKey>): CellKey[] {
		const visited = new Set<CellKey>()
		const onStack = new Set<CellKey>()
		const order: CellKey[] = []
		const dirtyBySheetRow = this.getCachedDirtyIndex(dirtySet)

		const visit = (key: CellKey): void => {
			if (visited.has(key)) return
			if (onStack.has(key)) return
			onStack.add(key)
			visited.add(key)
			const entry = this.formulas.get(key)
			if (entry) {
				for (const dep of entry.dependsOn) {
					if (dirtySet.has(dep)) visit(dep)
				}
				for (const range of entry.rangeDeps) {
					const sheetRows = dirtyBySheetRow.get(range.sheetIndex)
					if (!sheetRows) continue
					for (const [row, cols] of sheetRows) {
						for (const col of cols) {
							if (containsCell(range, row, col)) {
								visit(cellKey(range.sheetIndex, row, col))
							}
						}
					}
				}
			}
			onStack.delete(key)
			order.push(key)
		}

		for (const key of dirtySet) {
			visit(key)
		}
		return order
	}

	detectCycles(): CellKey[][] {
		let index = 0
		const stack: CellKey[] = []
		const onStack = new Set<CellKey>()
		const indices = new Map<CellKey, number>()
		const lowlinks = new Map<CellKey, number>()
		const sccs: CellKey[][] = []

		const strongconnect = (v: CellKey): void => {
			indices.set(v, index)
			lowlinks.set(v, index)
			index++
			stack.push(v)
			onStack.add(v)

			for (const w of this.getDependents(v)) {
				if (!this.formulas.has(w)) continue
				if (!indices.has(w)) {
					strongconnect(w)
					lowlinks.set(v, Math.min(lowlinks.get(v) ?? 0, lowlinks.get(w) ?? 0))
				} else if (onStack.has(w)) {
					lowlinks.set(v, Math.min(lowlinks.get(v) ?? 0, indices.get(w) ?? 0))
				}
			}

			if (lowlinks.get(v) === indices.get(v)) {
				const scc: CellKey[] = []
				let w: CellKey
				do {
					w = stack.pop() as CellKey
					onStack.delete(w)
					scc.push(w)
				} while (w !== v)

				if (scc.length > 1) {
					sccs.push(scc)
				} else {
					const solo = scc[0] as CellKey
					const entry = this.formulas.get(solo)
					if (entry?.dependsOn.has(solo)) {
						sccs.push(scc)
					}
				}
			}
		}

		for (const key of this.formulas.keys()) {
			if (!indices.has(key)) strongconnect(key)
		}
		return sccs
	}

	getAllFormulaCells(): CellKey[] {
		return [...this.formulas.keys()]
	}

	hasFormula(key: CellKey): boolean {
		return this.formulas.has(key)
	}

	clear(): void {
		this.formulas.clear()
		this.dependents.clear()
		this.rangeDependents.clear()
		this._rangeDepsSortedBySheet.clear()
		this._cachedDirtyIndex = null
		this._cachedEvalOrder = null
	}

	private ensureSorted(
		sheetIndex: number,
		entries: Array<{ range: RangeDependency; formulaKey: CellKey }>,
	): void {
		if (this._rangeDepsSortedBySheet.get(sheetIndex)) return
		entries.sort((a, b) => a.range.startRow - b.range.startRow)
		this._rangeDepsSortedBySheet.set(sheetIndex, true)
	}

	private getCachedDirtyIndex(
		dirtySet: ReadonlySet<CellKey>,
	): Map<number, Map<number, Set<number>>> {
		if (
			this._cachedDirtyIndex &&
			this._cachedDirtyIndex.set === dirtySet &&
			this._cachedDirtyIndex.size === dirtySet.size
		) {
			return this._cachedDirtyIndex.index
		}
		const index = indexDirtyCellsBySheetRow(dirtySet)
		this._cachedDirtyIndex = { set: dirtySet, size: dirtySet.size, index }
		return index
	}
}

function containsCell(range: RangeDependency, row: number, col: number): boolean {
	return (
		row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
	)
}

function setsMatchArray(existing: ReadonlySet<CellKey>, incoming: CellKey[]): boolean {
	if (existing.size !== incoming.length) return false
	for (const dep of incoming) {
		if (!existing.has(dep)) return false
	}
	return true
}

function rangeDepsEqual(a: readonly RangeDependency[], b: readonly RangeDependency[]): boolean {
	if (a.length !== b.length) return false
	for (let i = 0; i < a.length; i++) {
		const ra = a[i]
		const rb = b[i]
		if (!ra || !rb) return false
		if (
			ra.sheetIndex !== rb.sheetIndex ||
			ra.startRow !== rb.startRow ||
			ra.startCol !== rb.startCol ||
			ra.endRow !== rb.endRow ||
			ra.endCol !== rb.endCol
		)
			return false
	}
	return true
}

function indexDirtyCellsBySheetRow(
	dirtySet: ReadonlySet<CellKey>,
): Map<number, Map<number, Set<number>>> {
	const indexed = new Map<number, Map<number, Set<number>>>()
	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	for (const key of dirtySet) {
		parseCellKeyInto(key, coords)
		const { sheetIndex, row, col } = coords
		let rows = indexed.get(sheetIndex)
		if (!rows) {
			rows = new Map()
			indexed.set(sheetIndex, rows)
		}
		let cols = rows.get(row)
		if (!cols) {
			cols = new Set()
			rows.set(row, cols)
		}
		cols.add(col)
	}
	return indexed
}
