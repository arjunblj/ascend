import type { SheetId } from './ids.ts'

export type DefinedNameScope =
	| { readonly kind: 'workbook' }
	| { readonly kind: 'sheet'; readonly sheetId: SheetId }

export interface DefinedName {
	readonly name: string
	readonly formula: string
	readonly scope: DefinedNameScope
}

const WORKBOOK_SCOPE: DefinedNameScope = { kind: 'workbook' }

export class DefinedNameCollection {
	private readonly items: DefinedName[] = []

	get size(): number {
		return this.items.length
	}

	set(name: string, formula: string, scope: DefinedNameScope = WORKBOOK_SCOPE): void {
		const index = this.findIndex(name, scope)
		const entry: DefinedName = { name, formula, scope }
		if (index >= 0) {
			this.items[index] = entry
			return
		}
		this.items.push(entry)
	}

	get(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): string | undefined {
		return this.getEntry(name, scope)?.formula
	}

	getEntry(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): DefinedName | undefined {
		const lower = name.toLowerCase()
		return this.items.find(
			(item) => sameScope(item.scope, scope) && item.name.toLowerCase() === lower,
		)
	}

	resolve(
		name: string,
		currentSheetId?: SheetId,
		explicitSheetId?: SheetId,
	): DefinedName | undefined {
		const lower = name.toLowerCase()
		if (explicitSheetId) {
			return this.items.find(
				(item) =>
					item.scope.kind === 'sheet' &&
					item.scope.sheetId === explicitSheetId &&
					item.name.toLowerCase() === lower,
			)
		}
		if (currentSheetId) {
			const local = this.items.find(
				(item) =>
					item.scope.kind === 'sheet' &&
					item.scope.sheetId === currentSheetId &&
					item.name.toLowerCase() === lower,
			)
			if (local) return local
		}
		return this.items.find(
			(item) => item.scope.kind === 'workbook' && item.name.toLowerCase() === lower,
		)
	}

	has(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): boolean {
		return this.findIndex(name, scope) >= 0
	}

	delete(name: string, scope: DefinedNameScope = WORKBOOK_SCOPE): boolean {
		const index = this.findIndex(name, scope)
		if (index < 0) return false
		this.items.splice(index, 1)
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
		for (const item of this.items) {
			clone.set(item.name, item.formula, item.scope)
		}
		return clone
	}

	private findIndex(name: string, scope: DefinedNameScope): number {
		const lower = name.toLowerCase()
		return this.items.findIndex(
			(item) => sameScope(item.scope, scope) && item.name.toLowerCase() === lower,
		)
	}
}

function sameScope(left: DefinedNameScope, right: DefinedNameScope): boolean {
	if (left.kind !== right.kind) return false
	if (left.kind === 'workbook') return true
	return right.kind === 'sheet' && left.sheetId === right.sheetId
}
