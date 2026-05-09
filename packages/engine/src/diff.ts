import type { DefinedName, Workbook } from '@ascend/core'
import { toA1 } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import { valuesEqual } from '@ascend/schema'

export const cellValuesEqual = valuesEqual

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

export interface SheetFeatureDiff {
	readonly name: string
	readonly mergesChanged: boolean
	readonly tablesChanged: boolean
	readonly dataValidationsChanged: boolean
	readonly conditionalFormatsChanged: boolean
	readonly sheetProtectionChanged: boolean
}

export interface WorkbookDiff {
	readonly sheets: SheetDiff[]
	readonly namesAdded: string[]
	readonly namesRemoved: string[]
	readonly namesChanged: string[]
	readonly workbookProtectionChanged: boolean
	readonly sheetFeatures: readonly SheetFeatureDiff[]
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

		for (const [row, col, cellBefore] of beforeSheet.cells.iterate()) {
			if (!afterSheet.cells.has(row, col)) {
				cellsRemoved.push(toA1({ row, col }))
			} else if (
				!valuesEqual(cellBefore.value, afterSheet.cells.readValue(row, col)) ||
				cellBefore.formula !== (afterSheet.cells.readFormula(row, col) ?? null)
			) {
				const ref = toA1({ row, col })
				cellsChanged.push({
					ref,
					before: cellBefore.value,
					after: afterSheet.cells.readValue(row, col),
					formulaBefore: cellBefore.formula,
					formulaAfter: afterSheet.cells.readFormula(row, col) ?? null,
				})
			}
		}

		for (const [row, col] of afterSheet.cells.iterate()) {
			if (!beforeSheet.cells.has(row, col)) {
				cellsAdded.push(toA1({ row, col }))
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

	const workbookProtectionChanged =
		JSON.stringify(before.workbookProtection) !== JSON.stringify(after.workbookProtection)

	const sheetNames = new Set([...beforeSheets.keys(), ...afterSheets.keys()])
	const sheetFeatures: SheetFeatureDiff[] = []
	for (const name of sheetNames) {
		const bs = beforeSheets.get(name)
		const as = afterSheets.get(name)
		if (!bs || !as) continue
		const featureDiff: SheetFeatureDiff = {
			name,
			mergesChanged: JSON.stringify(bs.merges) !== JSON.stringify(as.merges),
			tablesChanged: JSON.stringify(bs.tables) !== JSON.stringify(as.tables),
			dataValidationsChanged:
				JSON.stringify(bs.dataValidations) !== JSON.stringify(as.dataValidations),
			conditionalFormatsChanged:
				JSON.stringify(bs.conditionalFormats) !== JSON.stringify(as.conditionalFormats),
			sheetProtectionChanged: JSON.stringify(bs.protection) !== JSON.stringify(as.protection),
		}
		if (
			featureDiff.mergesChanged ||
			featureDiff.tablesChanged ||
			featureDiff.dataValidationsChanged ||
			featureDiff.conditionalFormatsChanged ||
			featureDiff.sheetProtectionChanged
		) {
			sheetFeatures.push(featureDiff)
		}
	}

	return {
		sheets,
		namesAdded,
		namesRemoved,
		namesChanged,
		workbookProtectionChanged,
		sheetFeatures,
	}
}

function definedNameKey(workbook: Workbook, entry: DefinedName): string {
	if (entry.scope.kind === 'workbook') return entry.name
	const scope = entry.scope
	const sheetName = workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name ?? 'Sheet'
	return `${sheetName}!${entry.name}`
}
