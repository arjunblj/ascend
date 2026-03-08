import type { DefinedName, Workbook } from '@ascend/core'
import { toA1 } from '@ascend/core'
import type { CellValue } from '@ascend/schema'

export interface CellChange {
	readonly ref: string
	readonly before: CellValue
	readonly after: CellValue
	readonly formulaBefore: string | null
	readonly formulaAfter: string | null
}

export interface SheetDiff {
	readonly name: string
	readonly cellsAdded: string[]
	readonly cellsRemoved: string[]
	readonly cellsChanged: CellChange[]
}

export interface WorkbookDiff {
	readonly sheets: SheetDiff[]
	readonly namesAdded: string[]
	readonly namesRemoved: string[]
	readonly namesChanged: string[]
}

export function cellValuesEqual(a: CellValue, b: CellValue): boolean {
	if (a.kind !== b.kind) return false
	if (a.kind === 'empty') return true
	return JSON.stringify(a) === JSON.stringify(b)
}

export function diffWorkbooks(before: Workbook, after: Workbook): WorkbookDiff {
	const beforeSheets = new Map(before.sheets.map((s) => [s.name, s]))
	const afterSheets = new Map(after.sheets.map((s) => [s.name, s]))

	const sheets: SheetDiff[] = []

	for (const [name, beforeSheet] of beforeSheets) {
		const afterSheet = afterSheets.get(name)
		if (!afterSheet) {
			const cellsRemoved: string[] = []
			for (const [row, col] of beforeSheet.cells.iterate()) {
				cellsRemoved.push(toA1({ row, col }))
			}
			sheets.push({ name, cellsAdded: [], cellsRemoved, cellsChanged: [] })
			continue
		}

		const cellsAdded: string[] = []
		const cellsRemoved: string[] = []
		const cellsChanged: CellChange[] = []

		const beforeKeys = new Set<string>()
		for (const [row, col, cellBefore] of beforeSheet.cells.iterate()) {
			const ref = toA1({ row, col })
			beforeKeys.add(ref)
			const cellAfter = afterSheet.cells.get(row, col)
			if (!cellAfter) {
				cellsRemoved.push(ref)
			} else if (
				!cellValuesEqual(cellBefore.value, cellAfter.value) ||
				cellBefore.formula !== cellAfter.formula
			) {
				cellsChanged.push({
					ref,
					before: cellBefore.value,
					after: cellAfter.value,
					formulaBefore: cellBefore.formula,
					formulaAfter: cellAfter.formula,
				})
			}
		}

		for (const [row, col] of afterSheet.cells.iterate()) {
			const ref = toA1({ row, col })
			if (!beforeKeys.has(ref)) {
				cellsAdded.push(ref)
			}
		}

		if (cellsAdded.length > 0 || cellsRemoved.length > 0 || cellsChanged.length > 0) {
			sheets.push({ name, cellsAdded, cellsRemoved, cellsChanged })
		}
	}

	for (const [name, afterSheet] of afterSheets) {
		if (!beforeSheets.has(name)) {
			const cellsAdded: string[] = []
			for (const [row, col] of afterSheet.cells.iterate()) {
				cellsAdded.push(toA1({ row, col }))
			}
			sheets.push({ name, cellsAdded, cellsRemoved: [], cellsChanged: [] })
		}
	}

	const namesAdded: string[] = []
	const namesRemoved: string[] = []
	const namesChanged: string[] = []

	const beforeNames = new Map(
		before.definedNames.list().map((entry) => [definedNameKey(before, entry), entry.formula]),
	)
	const afterNames = new Map(
		after.definedNames.list().map((entry) => [definedNameKey(after, entry), entry.formula]),
	)

	for (const [name, ref] of beforeNames) {
		const afterRef = afterNames.get(name)
		if (afterRef === undefined) {
			namesRemoved.push(name)
		} else if (afterRef !== ref) {
			namesChanged.push(name)
		}
	}

	for (const name of afterNames.keys()) {
		if (!beforeNames.has(name)) {
			namesAdded.push(name)
		}
	}

	return { sheets, namesAdded, namesRemoved, namesChanged }
}

function definedNameKey(workbook: Workbook, entry: DefinedName): string {
	if (entry.scope.kind === 'workbook') return entry.name
	const scope = entry.scope
	const sheetName = workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name ?? 'Sheet'
	return `${sheetName}!${entry.name}`
}
