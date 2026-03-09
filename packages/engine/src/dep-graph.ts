export type CellKey = string

export interface RangeDependency {
	readonly sheetIndex: number
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

export function cellKey(sheetIndex: number, row: number, col: number): CellKey {
	return `${sheetIndex}:${row}:${col}`
}

export function parseCellKey(key: CellKey): readonly [number, number, number] {
	const parts = key.split(':')
	return [Number(parts[0]), Number(parts[1]), Number(parts[2])] as const
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

	addFormula(
		key: CellKey,
		dependsOn: CellKey[],
		isVolatile: boolean,
		rangeDeps: readonly RangeDependency[] = [],
	): void {
		this.removeFormula(key)
		const depSet = new Set(dependsOn)
		this.formulas.set(key, { dependsOn: depSet, rangeDeps, volatile: isVolatile })
		for (const dep of depSet) {
			let set = this.dependents.get(dep)
			if (!set) {
				set = new Set()
				this.dependents.set(dep, set)
			}
			set.add(key)
		}
		for (const range of rangeDeps) {
			const entries = this.rangeDependents.get(range.sheetIndex)
			if (entries) entries.push({ range, formulaKey: key })
			else this.rangeDependents.set(range.sheetIndex, [{ range, formulaKey: key }])
		}
	}

	removeFormula(key: CellKey): void {
		const entry = this.formulas.get(key)
		if (!entry) return
		for (const dep of entry.dependsOn) {
			const set = this.dependents.get(dep)
			if (set) {
				set.delete(key)
				if (set.size === 0) this.dependents.delete(dep)
			}
		}
		for (const [sheetIndex, entries] of this.rangeDependents) {
			const filtered = entries.filter((entry) => entry.formulaKey !== key)
			if (filtered.length === 0) this.rangeDependents.delete(sheetIndex)
			else if (filtered.length !== entries.length) this.rangeDependents.set(sheetIndex, filtered)
		}
		this.formulas.delete(key)
	}

	getPrecedents(key: CellKey): CellKey[] {
		const entry = this.formulas.get(key)
		if (!entry) return []
		const precedents = [...entry.dependsOn]
		for (const range of entry.rangeDeps) {
			for (let row = range.startRow; row <= range.endRow; row++) {
				for (let col = range.startCol; col <= range.endCol; col++) {
					precedents.push(cellKey(range.sheetIndex, row, col))
				}
			}
		}
		return precedents
	}

	getDependents(key: CellKey): CellKey[] {
		const direct = this.dependents.get(key)
		const result = new Set(direct ? [...direct] : [])
		const [sheetIndex, row, col] = parseCellKey(key)
		for (const entry of this.rangeDependents.get(sheetIndex) ?? []) {
			if (containsCell(entry.range, row, col)) {
				result.add(entry.formulaKey)
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
		const visited = new Set<CellKey>()
		const onStack = new Set<CellKey>()
		const order: CellKey[] = []

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
					for (let row = range.startRow; row <= range.endRow; row++) {
						for (let col = range.startCol; col <= range.endCol; col++) {
							const dep = cellKey(range.sheetIndex, row, col)
							if (dirtySet.has(dep)) visit(dep)
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
	}
}

function containsCell(range: RangeDependency, row: number, col: number): boolean {
	return (
		row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
	)
}
