import type { SheetId } from './ids.ts'

export type DefinedNameScope =
	| { readonly kind: 'workbook' }
	| { readonly kind: 'sheet'; readonly sheetId: SheetId }

export interface DefinedName {
	readonly name: string
	readonly formula: string
	readonly scope: DefinedNameScope
	readonly hidden?: boolean
}

export interface DefinedNameOptions {
	readonly hidden?: boolean
}

const WORKBOOK_SCOPE: DefinedNameScope = { kind: 'workbook' }

export class DefinedNameCollection {
	private readonly items: DefinedName[] = []
	private readonly workbookIndex = new Map<string, DefinedName>()
	private readonly sheetIndex = new Map<SheetId, Map<string, DefinedName>>()

	get size(): number {
		return this.items.length
	}

	set(
		name: string,
		formula: string,
		scope: DefinedNameScope = WORKBOOK_SCOPE,
		options: DefinedNameOptions = {},
	): void {
		const index = this.findIndex(name, scope)
		const entry: DefinedName = {
			name,
			formula,
			scope,
			...(options.hidden !== undefined ? { hidden: options.hidden } : {}),
		}
		if (index >= 0) {
			const previous = this.items[index]
			if (previous) this.removeFromIndex(previous)
			this.items[index] = entry
			this.addToIndex(entry)
			return
		}
		this.items.push(entry)
		this.addToIndex(entry)
	}

	get(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): string | undefined {
		return this.getEntry(name, scope)?.formula
	}

	getEntry(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): DefinedName | undefined {
		return this.getFromIndex(name, scope)
	}

	resolve(
		name: string,
		currentSheetId?: SheetId,
		explicitSheetId?: SheetId,
	): DefinedName | undefined {
		if (explicitSheetId) {
			return this.sheetIndex.get(explicitSheetId)?.get(name.toLowerCase())
		}
		if (currentSheetId) {
			const local = this.sheetIndex.get(currentSheetId)?.get(name.toLowerCase())
			if (local) return local
		}
		return this.workbookIndex.get(name.toLowerCase())
	}

	has(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): boolean {
		return this.findIndex(name, scope) >= 0
	}

	delete(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): boolean {
		const index = this.findIndex(name, scope)
		if (index < 0) return false
		const [removed] = this.items.splice(index, 1)
		if (removed) this.removeFromIndex(removed)
		return true
	}

	list(): readonly DefinedName[] {
		return this.items
	}

	workbookKeys(): string[] {
		return this.items.filter((item) => item.scope.kind === 'workbook').map((item) => item.name)
	}

	clone(): DefinedNameCollection {
		const clone = new DefinedNameCollection()
		clone.copyFrom(this)
		return clone
	}

	copyFrom(other: DefinedNameCollection): void {
		this.items.length = 0
		this.workbookIndex.clear()
		this.sheetIndex.clear()
		this.items.push(...other.items)
		for (const [name, entry] of other.workbookIndex) {
			this.workbookIndex.set(name, entry)
		}
		for (const [sheetId, names] of other.sheetIndex) {
			this.sheetIndex.set(sheetId, new Map(names))
		}
	}

	private findIndex(name: string, scope: DefinedNameScope): number {
		const entry = this.getFromIndex(name, scope)
		if (!entry) return -1
		for (let i = 0; i < this.items.length; i++) {
			if (this.items[i] === entry) return i
		}
		return -1
	}

	private getFromIndex(name: string, scope: DefinedNameScope): DefinedName | undefined {
		const lower = name.toLowerCase()
		if (scope.kind === 'workbook') return this.workbookIndex.get(lower)
		return this.sheetIndex.get(scope.sheetId)?.get(lower)
	}

	private addToIndex(entry: DefinedName): void {
		const lower = entry.name.toLowerCase()
		if (entry.scope.kind === 'workbook') {
			this.workbookIndex.set(lower, entry)
			return
		}
		let names = this.sheetIndex.get(entry.scope.sheetId)
		if (!names) {
			names = new Map()
			this.sheetIndex.set(entry.scope.sheetId, names)
		}
		names.set(lower, entry)
	}

	private removeFromIndex(entry: DefinedName): void {
		const lower = entry.name.toLowerCase()
		if (entry.scope.kind === 'workbook') {
			this.workbookIndex.delete(lower)
			return
		}
		const names = this.sheetIndex.get(entry.scope.sheetId)
		if (!names) return
		names.delete(lower)
		if (names.size === 0) this.sheetIndex.delete(entry.scope.sheetId)
	}
}
