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
	readonly dependsOn: readonly CellKey[]
	readonly rangeDeps: readonly RangeDependency[]
	readonly volatile: boolean
}

export { IntervalIndex } from './dep-graph-interval.ts'

import { IntervalIndex } from './dep-graph-interval.ts'

const _depCoords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
const TOPO_RANK_GAP = 1_024

export class DependencyGraph {
	private readonly formulas = new Map<CellKey, FormulaEntry>()
	private readonly rangeIndex = new Map<number, IntervalIndex>()
	private _cachedDirtyIndex: {
		set: ReadonlySet<CellKey>
		size: number
		hash: number
		index: Map<number, Map<number, number[]>>
	} | null = null
	private _cachedEvalOrder: CellKey[] | null = null
	private _rankByKey: Map<CellKey, number> | null = null
	private _compiledDirectIndex: {
		sourceIndex: Map<CellKey, number>
		starts: Int32Array
		counts: Int32Array
		dependents: readonly CellKey[]
		formulaKeys: readonly CellKey[]
		formulaKeySet: ReadonlySet<CellKey>
		volatileKeys: readonly CellKey[]
	} | null = null

	addFormula(
		key: CellKey,
		dependsOn: readonly CellKey[],
		isVolatile: boolean,
		rangeDeps: readonly RangeDependency[] = [],
	): void {
		this.removeFormula(key)
		const aboveKey = key - COL_FACTOR
		const aboveEntry = aboveKey >= 0 ? this.formulas.get(aboveKey) : undefined
		let deps: readonly CellKey[]
		let sharedRangeDeps = rangeDeps
		if (aboveEntry && arraysMatch(aboveEntry.dependsOn, dependsOn)) {
			deps = aboveEntry.dependsOn
			if (rangeDepsEqual(aboveEntry.rangeDeps, rangeDeps)) {
				sharedRangeDeps = aboveEntry.rangeDeps
			}
		} else {
			deps = dependsOn
		}
		this._cachedEvalOrder = null
		this._compiledDirectIndex = null
		this.formulas.set(key, { dependsOn: deps, rangeDeps: sharedRangeDeps, volatile: isVolatile })
		this._onFormulaAdded(key, deps)
		for (const range of rangeDeps) {
			let index = this.rangeIndex.get(range.sheetIndex)
			if (!index) {
				index = new IntervalIndex()
				this.rangeIndex.set(range.sheetIndex, index)
			}
			index.insert(range.startRow, range.endRow, range.startCol, range.endCol, key)
		}
	}

	removeFormula(key: CellKey): void {
		const entry = this.formulas.get(key)
		if (!entry) return
		this._cachedEvalOrder = null
		this._compiledDirectIndex = null
		this._rankByKey?.delete(key)
		const removedSheets = new Set<number>()
		for (const range of entry.rangeDeps) {
			if (removedSheets.has(range.sheetIndex)) continue
			removedSheets.add(range.sheetIndex)
			const index = this.rangeIndex.get(range.sheetIndex)
			if (!index) continue
			index.remove(key)
			if (index.size === 0) this.rangeIndex.delete(range.sheetIndex)
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
		const direct = this._readDirectDependents(key)
		parseCellKeyInto(key, _depCoords)
		const index = this.rangeIndex.get(_depCoords.sheetIndex)
		if (direct.length === 0 && !index) return []
		if (!index) return direct
		const rangeHits = index.query(_depCoords.row, _depCoords.col)
		if (direct.length === 0) return rangeHits
		if (rangeHits.length === 0) return [...direct]
		const directSet = new Set(direct)
		const result: CellKey[] = [...direct]
		for (let i = 0; i < rangeHits.length; i++) {
			const k = rangeHits[i] as CellKey
			if (!directSet.has(k)) result.push(k)
		}
		return result
	}

	getVolatiles(): CellKey[] {
		return [...this._ensureDirectIndex().volatileKeys]
	}

	getDirtySet(changedKeys: CellKey[]): Set<CellKey> {
		const dirty = new Set<CellKey>()
		for (const key of changedKeys) dirty.add(key)

		const directIndex = this._ensureDirectIndex()
		let frontier = changedKeys

		while (frontier.length > 0) {
			const nextFrontier: CellKey[] = []

			for (let f = 0; f < frontier.length; f++) {
				const key = frontier[f] as CellKey
				const slot = directIndex.sourceIndex.get(key)
				if (slot === undefined) continue
				const start = directIndex.starts[slot] ?? 0
				const count = directIndex.counts[slot] ?? 0
				for (let i = start; i < start + count; i++) {
					const dep = directIndex.dependents[i]
					if (dep !== undefined && !dirty.has(dep)) {
						dirty.add(dep)
						nextFrontier.push(dep)
					}
				}
			}

			if (this.rangeIndex.size > 0) {
				const bySheet = groupCellKeysBySheetCol(frontier)
				for (const [sheet, cellsByCol] of bySheet) {
					const index = this.rangeIndex.get(sheet)
					if (!index) continue

					let totalCells = 0
					for (const rows of cellsByCol.values()) totalCells += rows.length

					if (totalCells * 4 > index.size) {
						const hits = index.queryBatch(cellsByCol)
						for (let i = 0; i < hits.length; i++) {
							const hit = hits[i] as CellKey
							if (!dirty.has(hit)) {
								dirty.add(hit)
								nextFrontier.push(hit)
							}
						}
					} else {
						for (const [col, rows] of cellsByCol) {
							for (const row of rows) {
								const hits = index.query(row, col)
								for (let i = 0; i < hits.length; i++) {
									const hit = hits[i] as CellKey
									if (!dirty.has(hit)) {
										dirty.add(hit)
										nextFrontier.push(hit)
									}
								}
							}
						}
					}
				}
			}

			frontier = nextFrontier
		}

		return dirty
	}

	/**
	 * Returns cells in topological eval order. When the full order cache is stale,
	 * we normally rebuild it from all formulas. For dirty-only recalc, when
	 * dirtySet is smaller than all formulas, we skip the full rebuild and compute
	 * order only for the dirty subset (lazy order). This avoids O(all formulas)
	 * topological sort when only a small subset needs evaluation.
	 */
	getEvalOrder(dirtySet: Set<CellKey>): CellKey[] {
		const directIndex = this._ensureDirectIndex()
		const allFormulaKeys = directIndex.formulaKeys
		const formulaKeySet = directIndex.formulaKeySet
		let dirtyFormulaCount = 0
		const nonFormulaKeys: CellKey[] = []
		for (const k of dirtySet) {
			if (!formulaKeySet.has(k)) nonFormulaKeys.push(k)
			if (this.formulas.has(k)) dirtyFormulaCount++
		}

		const usePartialOrder =
			dirtySet.size < allFormulaKeys.length && dirtyFormulaCount < allFormulaKeys.length

		if (usePartialOrder) {
			const partialOrder = this._computeEvalOrder(dirtySet)
			const formulaOrder = partialOrder.filter((k) => this.formulas.has(k))
			return [...nonFormulaKeys, ...formulaOrder]
		}

		if (this._cachedEvalOrder === null) {
			const allSet = new Set(allFormulaKeys)
			this._cachedEvalOrder =
				dirtyFormulaCount === allFormulaKeys.length && !this._hasFormulaPrecedents(allSet)
					? [...allFormulaKeys]
					: this._computeEvalOrder(allSet)
			this._rebuildRanksFromOrder(this._cachedEvalOrder)
		}
		if (dirtyFormulaCount === allFormulaKeys.length) {
			return [...nonFormulaKeys, ...this._cachedEvalOrder]
		}
		const formulaQueue = new BinaryMinHeap(
			(a, b) => (this._rankByKey?.get(a) ?? 0) - (this._rankByKey?.get(b) ?? 0),
		)
		for (const key of dirtySet) {
			if (this.formulas.has(key)) formulaQueue.push(key)
		}
		const formulaOrder: CellKey[] = []
		while (formulaQueue.size > 0) {
			const key = formulaQueue.pop()
			if (key !== undefined) formulaOrder.push(key)
		}
		return [...nonFormulaKeys, ...formulaOrder]
	}

	private _computeEvalOrder(dirtySet: Set<CellKey>): CellKey[] {
		const visited = new Set<CellKey>()
		const onStack = new Set<CellKey>()
		const order: CellKey[] = []
		const dirtyFormulaBySheetCol = this.getCachedDirtyFormulaIndex(dirtySet)

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
					const sheetCols = dirtyFormulaBySheetCol.get(range.sheetIndex)
					if (!sheetCols) continue
					for (let col = range.startCol; col <= range.endCol; col++) {
						const rows = sheetCols.get(col)
						if (!rows) continue
						for (const row of rows) {
							if (row > range.endRow) break
							if (row >= range.startRow) visit(cellKey(range.sheetIndex, row, col))
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

	private _hasFormulaPrecedents(formulaSet: Set<CellKey>): boolean {
		const dirtyFormulaBySheetCol = this.getCachedDirtyFormulaIndex(formulaSet)
		for (const entry of this.formulas.values()) {
			for (let i = 0; i < entry.dependsOn.length; i++) {
				if (formulaSet.has(entry.dependsOn[i] as CellKey)) return true
			}
			for (const range of entry.rangeDeps) {
				const sheetCols = dirtyFormulaBySheetCol.get(range.sheetIndex)
				if (!sheetCols) continue
				for (let col = range.startCol; col <= range.endCol; col++) {
					const rows = sheetCols.get(col)
					if (rows && sortedRowsOverlapRange(rows, range.startRow, range.endRow)) return true
				}
			}
		}
		return false
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
					if (entry?.dependsOn.includes(solo)) {
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
		return [...this._ensureDirectIndex().formulaKeys]
	}

	/**
	 * Returns independent subgraphs (connected components) of the formula dependency graph.
	 * Cells in different subgraphs have no formula-to-formula dependency path between them
	 * and could theoretically be evaluated in parallel.
	 * Uses only direct cell dependencies (dependsOn); range deps are not expanded.
	 */
	getIndependentSubgraphs(): readonly (readonly CellKey[])[] {
		const formulaKeys = this._ensureDirectIndex().formulaKeys
		if (formulaKeys.length === 0) return []
		const keyToId = new Map<CellKey, number>()
		for (let i = 0; i < formulaKeys.length; i++) {
			keyToId.set(formulaKeys[i] as CellKey, i)
		}
		const parent = new Uint32Array(formulaKeys.length)
		for (let i = 0; i < parent.length; i++) parent[i] = i

		const find = (x: number): number => {
			if (parent[x] !== x) parent[x] = find(parent[x] as number)
			return parent[x] as number
		}
		const union = (a: number, b: number): void => {
			const ra = find(a)
			const rb = find(b)
			if (ra !== rb) parent[ra] = rb
		}

		for (const key of formulaKeys) {
			const entry = this.formulas.get(key)
			if (!entry) continue
			const idA = keyToId.get(key)
			if (idA === undefined) continue
			for (const dep of entry.dependsOn) {
				if (!this.formulas.has(dep)) continue
				const idB = keyToId.get(dep)
				if (idB !== undefined) union(idA, idB)
			}
		}

		const byRoot = new Map<number, CellKey[]>()
		for (let i = 0; i < formulaKeys.length; i++) {
			const root = find(i)
			const bucket = byRoot.get(root)
			const key = formulaKeys[i] as CellKey
			if (bucket) bucket.push(key)
			else byRoot.set(root, [key])
		}
		return [...byRoot.values()]
	}

	hasFormula(key: CellKey): boolean {
		return this.formulas.has(key)
	}

	clear(): void {
		this.formulas.clear()
		this.rangeIndex.clear()
		this._cachedDirtyIndex = null
		this._cachedEvalOrder = null
		this._rankByKey = null
		this._compiledDirectIndex = null
	}

	private _onFormulaAdded(key: CellKey, dependsOn: readonly CellKey[]): void {
		if (this._rankByKey === null) return
		if (!this._rankByKey.has(key)) {
			let maxRank = 0
			for (const rank of this._rankByKey.values()) {
				if (rank > maxRank) maxRank = rank
			}
			this._rankByKey.set(key, maxRank + TOPO_RANK_GAP)
		}
		for (let i = 0; i < dependsOn.length; i++) {
			const dep = dependsOn[i] as CellKey
			if (!this.formulas.has(dep)) continue
			const depRank = this._rankByKey.get(dep)
			const keyRank = this._rankByKey.get(key)
			if (depRank === undefined || keyRank === undefined || depRank < keyRank) continue
			if (!this._repairRanks(dep, key)) {
				this._cachedEvalOrder = null
				this._rankByKey = null
				return
			}
		}
	}

	private _repairRanks(source: CellKey, target: CellKey): boolean {
		const ranks = this._rankByKey
		if (!ranks) return false
		const upperRank = ranks.get(source)
		const lowerRank = ranks.get(target)
		if (upperRank === undefined || lowerRank === undefined || upperRank < lowerRank) return true
		const forward = new Set<CellKey>()
		const backward = new Set<CellKey>()
		const forwardStack = [target]
		while (forwardStack.length > 0) {
			const current = forwardStack.pop() as CellKey
			if (forward.has(current)) continue
			const rank = ranks.get(current)
			if (rank === undefined || rank > upperRank) continue
			forward.add(current)
			for (const dependent of this._scanFormulaDependents(current)) {
				forwardStack.push(dependent)
			}
		}
		if (forward.has(source)) return false
		const backwardStack = [source]
		while (backwardStack.length > 0) {
			const current = backwardStack.pop() as CellKey
			if (backward.has(current)) continue
			const rank = ranks.get(current)
			if (rank === undefined || rank < lowerRank) continue
			backward.add(current)
			const entry = this.formulas.get(current)
			if (!entry) continue
			for (let i = 0; i < entry.dependsOn.length; i++) {
				const precedent = entry.dependsOn[i] as CellKey
				if (this.formulas.has(precedent)) backwardStack.push(precedent)
			}
		}
		const affected = new Set<CellKey>([...backward, ...forward])
		if (affected.size === 0) return true
		const indegree = new Map<CellKey, number>()
		const dependents = new Map<CellKey, CellKey[]>()
		for (const key of affected) {
			indegree.set(key, 0)
			dependents.set(key, [])
		}
		for (const key of affected) {
			const entry = this.formulas.get(key)
			if (!entry) continue
			for (let i = 0; i < entry.dependsOn.length; i++) {
				const precedent = entry.dependsOn[i] as CellKey
				if (!affected.has(precedent)) continue
				indegree.set(key, (indegree.get(key) ?? 0) + 1)
				const bucket = dependents.get(precedent)
				if (bucket) bucket.push(key)
			}
		}
		const queue = [...affected]
			.filter((key) => (indegree.get(key) ?? 0) === 0)
			.sort((a, b) => (ranks.get(a) ?? 0) - (ranks.get(b) ?? 0))
		const ordered: CellKey[] = []
		while (queue.length > 0) {
			const key = queue.shift() as CellKey
			ordered.push(key)
			for (const dependent of dependents.get(key) ?? []) {
				const next = (indegree.get(dependent) ?? 0) - 1
				indegree.set(dependent, next)
				if (next === 0) {
					queue.push(dependent)
					queue.sort((a, b) => (ranks.get(a) ?? 0) - (ranks.get(b) ?? 0))
				}
			}
		}
		if (ordered.length !== affected.size) return false
		const minRank = Math.min(...ordered.map((key) => ranks.get(key) ?? 0))
		for (let i = 0; i < ordered.length; i++) {
			ranks.set(ordered[i] as CellKey, minRank + i * TOPO_RANK_GAP)
		}
		return true
	}

	private _scanFormulaDependents(precedent: CellKey): CellKey[] {
		const directIndex = this._compiledDirectIndex
		if (directIndex) {
			const slot = directIndex.sourceIndex.get(precedent)
			if (slot === undefined) return []
			const start = directIndex.starts[slot] ?? 0
			const count = directIndex.counts[slot] ?? 0
			return directIndex.dependents.slice(start, start + count)
		}
		const result: CellKey[] = []
		for (const [formulaKey, entry] of this.formulas) {
			if (entry.dependsOn.includes(precedent)) result.push(formulaKey)
		}
		return result
	}

	private _rebuildRanksFromOrder(order: readonly CellKey[]): void {
		const ranks = new Map<CellKey, number>()
		for (let i = 0; i < order.length; i++) {
			ranks.set(order[i] as CellKey, (i + 1) * TOPO_RANK_GAP)
		}
		this._rankByKey = ranks
	}

	private _readDirectDependents(key: CellKey): CellKey[] {
		const directIndex = this._ensureDirectIndex()
		const slot = directIndex.sourceIndex.get(key)
		if (slot === undefined) return []
		const start = directIndex.starts[slot] ?? 0
		const count = directIndex.counts[slot] ?? 0
		if (count === 0) return []
		return directIndex.dependents.slice(start, start + count)
	}

	private _ensureDirectIndex(): {
		sourceIndex: Map<CellKey, number>
		starts: Int32Array
		counts: Int32Array
		dependents: readonly CellKey[]
		formulaKeys: readonly CellKey[]
		formulaKeySet: ReadonlySet<CellKey>
		volatileKeys: readonly CellKey[]
	} {
		if (this._compiledDirectIndex) return this._compiledDirectIndex
		const formulaKeys = [...this.formulas.keys()]
		const formulaKeySet = new Set(formulaKeys)
		const volatileKeys: CellKey[] = []
		const directBuckets = new Map<CellKey, CellKey[]>()
		for (const [formulaKey, entry] of this.formulas) {
			if (entry.volatile) volatileKeys.push(formulaKey)
			for (let i = 0; i < entry.dependsOn.length; i++) {
				const dep = entry.dependsOn[i] as CellKey
				const bucket = directBuckets.get(dep)
				if (bucket) bucket.push(formulaKey)
				else directBuckets.set(dep, [formulaKey])
			}
		}
		const sourceIndex = new Map<CellKey, number>()
		const starts = new Int32Array(directBuckets.size)
		const counts = new Int32Array(directBuckets.size)
		const dependents: CellKey[] = []
		let offset = 0
		let slot = 0
		for (const [source, bucket] of directBuckets) {
			sourceIndex.set(source, slot)
			starts[slot] = offset
			counts[slot] = bucket.length
			for (let i = 0; i < bucket.length; i++) {
				dependents.push(bucket[i] as CellKey)
			}
			offset += bucket.length
			slot++
		}
		this._compiledDirectIndex = {
			sourceIndex,
			starts,
			counts,
			dependents,
			formulaKeys,
			formulaKeySet,
			volatileKeys,
		}
		return this._compiledDirectIndex
	}

	private getCachedDirtyFormulaIndex(
		dirtySet: ReadonlySet<CellKey>,
	): Map<number, Map<number, number[]>> {
		if (this._cachedDirtyIndex) {
			const c = this._cachedDirtyIndex
			if (c.set === dirtySet && c.size === dirtySet.size) return c.index
			if (c.size === dirtySet.size && c.hash === hashCellKeySet(dirtySet)) return c.index
		}
		const index = indexDirtyFormulaCellsBySheetCol(dirtySet, this.formulas)
		this._cachedDirtyIndex = {
			set: dirtySet,
			size: dirtySet.size,
			hash: hashCellKeySet(dirtySet),
			index,
		}
		return index
	}
}

function arraysMatch(existing: readonly CellKey[], incoming: readonly CellKey[]): boolean {
	if (existing.length !== incoming.length) return false
	for (let i = 0; i < existing.length; i++) {
		if ((existing[i] as CellKey) !== (incoming[i] as CellKey)) return false
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

function hashCellKeySet(set: ReadonlySet<CellKey>): number {
	let h = 0x811c9dc5
	for (const key of set) {
		const low = key >>> 0
		const high = Math.trunc(key / 0x1_0000_0000) >>> 0
		h = fnv1a32(h, low)
		h = fnv1a32(h, high)
	}
	return h | 0
}

function fnv1a32(hash: number, value: number): number {
	hash ^= value & 0xff
	hash = Math.imul(hash, 0x01000193)
	hash ^= (value >>> 8) & 0xff
	hash = Math.imul(hash, 0x01000193)
	hash ^= (value >>> 16) & 0xff
	hash = Math.imul(hash, 0x01000193)
	hash ^= (value >>> 24) & 0xff
	return Math.imul(hash, 0x01000193)
}

function groupCellKeysBySheetCol(keys: readonly CellKey[]): Map<number, Map<number, number[]>> {
	const result = new Map<number, Map<number, number[]>>()
	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	for (const key of keys) {
		parseCellKeyInto(key, coords)
		let cols = result.get(coords.sheetIndex)
		if (!cols) {
			cols = new Map()
			result.set(coords.sheetIndex, cols)
		}
		let rows = cols.get(coords.col)
		if (!rows) {
			rows = []
			cols.set(coords.col, rows)
		}
		rows.push(coords.row)
	}
	for (const cols of result.values()) {
		for (const rows of cols.values()) {
			rows.sort((a, b) => a - b)
		}
	}
	return result
}

function indexDirtyFormulaCellsBySheetCol(
	dirtySet: ReadonlySet<CellKey>,
	formulas: ReadonlyMap<CellKey, FormulaEntry>,
): Map<number, Map<number, number[]>> {
	const indexed = new Map<number, Map<number, number[]>>()
	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	for (const key of dirtySet) {
		if (!formulas.has(key)) continue
		parseCellKeyInto(key, coords)
		const { sheetIndex, row, col } = coords
		let cols = indexed.get(sheetIndex)
		if (!cols) {
			cols = new Map()
			indexed.set(sheetIndex, cols)
		}
		let rows = cols.get(col)
		if (!rows) {
			rows = []
			cols.set(col, rows)
		}
		rows.push(row)
	}
	for (const cols of indexed.values()) {
		for (const rows of cols.values()) {
			rows.sort((a, b) => a - b)
		}
	}
	return indexed
}

function sortedRowsOverlapRange(
	rows: readonly number[],
	startRow: number,
	endRow: number,
): boolean {
	let lo = 0
	let hi = rows.length - 1
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1
		if ((rows[mid] as number) < startRow) lo = mid + 1
		else hi = mid - 1
	}
	return lo < rows.length && (rows[lo] as number) <= endRow
}

class BinaryMinHeap {
	private readonly values: CellKey[] = []

	constructor(private readonly compare: (left: CellKey, right: CellKey) => number) {}

	get size(): number {
		return this.values.length
	}

	push(value: CellKey): void {
		this.values.push(value)
		this.bubbleUp(this.values.length - 1)
	}

	pop(): CellKey | undefined {
		if (this.values.length === 0) return undefined
		const first = this.values[0]
		const last = this.values.pop()
		if (last !== undefined && this.values.length > 0) {
			this.values[0] = last
			this.bubbleDown(0)
		}
		return first
	}

	private bubbleUp(index: number): void {
		let current = index
		while (current > 0) {
			const parent = (current - 1) >>> 1
			if (this.compare(this.values[current] as CellKey, this.values[parent] as CellKey) >= 0) break
			;[this.values[current], this.values[parent]] = [
				this.values[parent] as CellKey,
				this.values[current] as CellKey,
			]
			current = parent
		}
	}

	private bubbleDown(index: number): void {
		let current = index
		const length = this.values.length
		while (true) {
			const left = current * 2 + 1
			const right = left + 1
			let next = current
			if (
				left < length &&
				this.compare(this.values[left] as CellKey, this.values[next] as CellKey) < 0
			) {
				next = left
			}
			if (
				right < length &&
				this.compare(this.values[right] as CellKey, this.values[next] as CellKey) < 0
			) {
				next = right
			}
			if (next === current) break
			;[this.values[current], this.values[next]] = [
				this.values[next] as CellKey,
				this.values[current] as CellKey,
			]
			current = next
		}
	}
}
