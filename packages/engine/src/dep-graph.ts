export type CellKey = string

export function cellKey(sheetIndex: number, row: number, col: number): CellKey {
	return `${sheetIndex}:${row}:${col}`
}

export function parseCellKey(key: CellKey): readonly [number, number, number] {
	const parts = key.split(':')
	return [Number(parts[0]), Number(parts[1]), Number(parts[2])] as const
}

interface FormulaEntry {
	readonly dependsOn: ReadonlySet<CellKey>
	readonly volatile: boolean
}

export class DependencyGraph {
	private readonly formulas = new Map<CellKey, FormulaEntry>()
	private readonly dependents = new Map<CellKey, Set<CellKey>>()

	addFormula(key: CellKey, dependsOn: CellKey[], isVolatile: boolean): void {
		this.removeFormula(key)
		const depSet = new Set(dependsOn)
		this.formulas.set(key, { dependsOn: depSet, volatile: isVolatile })
		for (const dep of depSet) {
			let set = this.dependents.get(dep)
			if (!set) {
				set = new Set()
				this.dependents.set(dep, set)
			}
			set.add(key)
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
		this.formulas.delete(key)
	}

	getPrecedents(key: CellKey): CellKey[] {
		const entry = this.formulas.get(key)
		return entry ? [...entry.dependsOn] : []
	}

	getDependents(key: CellKey): CellKey[] {
		const set = this.dependents.get(key)
		return set ? [...set] : []
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
			const deps = this.dependents.get(current)
			if (deps) {
				for (const dep of deps) {
					if (!dirty.has(dep)) queue.push(dep)
				}
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

			const deps = this.dependents.get(v)
			if (deps) {
				for (const w of deps) {
					if (!this.formulas.has(w)) continue
					if (!indices.has(w)) {
						strongconnect(w)
						lowlinks.set(v, Math.min(lowlinks.get(v) ?? 0, lowlinks.get(w) ?? 0))
					} else if (onStack.has(w)) {
						lowlinks.set(v, Math.min(lowlinks.get(v) ?? 0, indices.get(w) ?? 0))
					}
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
	}
}
