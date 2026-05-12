import type { Workbook } from '@ascend/core'
import { toA1 } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import type { CellChange, SheetDiff, SheetFeatureDiff, WorkbookDiff } from './diff.ts'
import { cellValuesEqual } from './diff.ts'

export interface SheetSnapshot {
	readonly name: string
	readonly cells: Record<string, { value: CellValue; formula: string | null }>
	readonly mergesJson: string
	readonly tablesJson: string
	readonly dataValidationsJson: string
	readonly x14DataValidationsJson: string
	readonly conditionalFormatsJson: string
	readonly x14ConditionalFormatsJson: string
	readonly protectionJson: string
}

export interface WorkbookSnapshot {
	readonly sheets: SheetSnapshot[]
	readonly names: Record<string, string>
	readonly workbookProtectionJson: string
	readonly timestamp: number
}

export function createSnapshot(workbook: Workbook): WorkbookSnapshot {
	const sheets: SheetSnapshot[] = workbook.sheets.map((sheet) => {
		const cells: Record<string, { value: CellValue; formula: string | null }> = {}
		for (const [row, col, c] of sheet.cells.iterate()) {
			cells[toA1({ row, col })] = { value: c.value, formula: c.formula }
		}
		return {
			name: sheet.name,
			cells,
			mergesJson: JSON.stringify(sheet.merges),
			tablesJson: JSON.stringify(sheet.tables),
			dataValidationsJson: JSON.stringify(sheet.dataValidations),
			x14DataValidationsJson: JSON.stringify(sheet.x14DataValidations),
			conditionalFormatsJson: JSON.stringify(sheet.conditionalFormats),
			x14ConditionalFormatsJson: JSON.stringify(sheet.x14ConditionalFormats),
			protectionJson: JSON.stringify(sheet.protection),
		}
	})

	const names: Record<string, string> = {}
	for (const definedName of workbook.definedNames.list()) {
		let key = definedName.name
		if (definedName.scope.kind === 'sheet') {
			const scope = definedName.scope
			const sheetName = workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name ?? 'Sheet'
			key = `${sheetName}!${definedName.name}`
		}
		names[key] = definedName.formula
	}

	return {
		sheets,
		names,
		workbookProtectionJson: JSON.stringify(workbook.workbookProtection),
		timestamp: Date.now(),
	}
}

export function compareSnapshots(a: WorkbookSnapshot, b: WorkbookSnapshot): WorkbookDiff {
	const sheetsA = new Map(a.sheets.map((s) => [s.name, s]))
	const sheetsB = new Map(b.sheets.map((s) => [s.name, s]))

	const sheets: SheetDiff[] = []

	for (const [name, sheetA] of sheetsA) {
		const sheetB = sheetsB.get(name)
		if (!sheetB) {
			sheets.push({
				name,
				cellsAdded: [],
				cellsRemoved: Object.keys(sheetA.cells),
				cellsChanged: [],
			})
			continue
		}

		const cellsAdded: string[] = []
		const cellsRemoved: string[] = []
		const cellsChanged: CellChange[] = []

		for (const ref of Object.keys(sheetA.cells)) {
			const cellA = sheetA.cells[ref]
			const cellB = sheetB.cells[ref]
			if (!cellA) continue
			if (!cellB) {
				cellsRemoved.push(ref)
			} else if (!cellValuesEqual(cellA.value, cellB.value) || cellA.formula !== cellB.formula) {
				cellsChanged.push({
					ref,
					before: cellA.value,
					after: cellB.value,
					formulaBefore: cellA.formula,
					formulaAfter: cellB.formula,
				})
			}
		}

		for (const ref of Object.keys(sheetB.cells)) {
			const cellA = sheetA.cells[ref]
			if (!cellA) {
				cellsAdded.push(ref)
			}
		}

		if (cellsAdded.length > 0 || cellsRemoved.length > 0 || cellsChanged.length > 0) {
			sheets.push({ name, cellsAdded, cellsRemoved, cellsChanged })
		}
	}

	for (const [name, sheetB] of sheetsB) {
		if (!sheetsA.has(name)) {
			sheets.push({
				name,
				cellsAdded: Object.keys(sheetB.cells),
				cellsRemoved: [],
				cellsChanged: [],
			})
		}
	}

	const namesAdded: string[] = []
	const namesRemoved: string[] = []
	const namesChanged: string[] = []

	for (const nameKey of Object.keys(a.names)) {
		const aRef = a.names[nameKey]
		const bRef = b.names[nameKey]
		if (!aRef) continue
		if (!bRef) {
			namesRemoved.push(nameKey)
		} else if (aRef !== bRef) {
			namesChanged.push(nameKey)
		}
	}

	for (const nameKey of Object.keys(b.names)) {
		const aRef = a.names[nameKey]
		if (!aRef) {
			namesAdded.push(nameKey)
		}
	}

	const workbookProtectionChanged = a.workbookProtectionJson !== b.workbookProtectionJson
	const sheetNames = new Set([...sheetsA.keys(), ...sheetsB.keys()])
	const sheetFeatures: SheetFeatureDiff[] = []
	for (const name of sheetNames) {
		const sheetA = sheetsA.get(name)
		const sheetB = sheetsB.get(name)
		if (!sheetA || !sheetB) continue
		const dataValidationsChanged = sheetA.dataValidationsJson !== sheetB.dataValidationsJson
		const x14DataValidationsChanged =
			sheetA.x14DataValidationsJson !== sheetB.x14DataValidationsJson
		const conditionalFormatsChanged =
			sheetA.conditionalFormatsJson !== sheetB.conditionalFormatsJson
		const x14ConditionalFormatsChanged =
			sheetA.x14ConditionalFormatsJson !== sheetB.x14ConditionalFormatsJson
		const featureDiff: SheetFeatureDiff = {
			name,
			mergesChanged: sheetA.mergesJson !== sheetB.mergesJson,
			tablesChanged: sheetA.tablesJson !== sheetB.tablesJson,
			dataValidationsChanged: dataValidationsChanged || x14DataValidationsChanged,
			x14DataValidationsChanged,
			conditionalFormatsChanged: conditionalFormatsChanged || x14ConditionalFormatsChanged,
			x14ConditionalFormatsChanged,
			sheetProtectionChanged: sheetA.protectionJson !== sheetB.protectionJson,
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
