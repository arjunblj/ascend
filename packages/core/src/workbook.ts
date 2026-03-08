import { type CalcSettings, DEFAULT_CALC_SETTINGS } from '@ascend/schema'
import { createWorkbookId, type SheetId, type WorkbookId } from './ids.ts'
import { createSheet, type Sheet } from './sheet.ts'
import { StyleRegistry } from './style-registry.ts'

export class Workbook {
	readonly id: WorkbookId
	readonly sheets: Sheet[] = []
	readonly definedNames = new Map<string, string>()
	readonly styles = new StyleRegistry()
	calcSettings: CalcSettings

	constructor(id?: WorkbookId) {
		this.id = id ?? createWorkbookId()
		this.calcSettings = DEFAULT_CALC_SETTINGS
	}

	addSheet(name: string, id?: SheetId): Sheet {
		const sheet = createSheet(name, id)
		this.sheets.push(sheet)
		return sheet
	}

	getSheet(nameOrId: string): Sheet | undefined {
		return this.sheets.find((s) => s.name === nameOrId || s.id === nameOrId)
	}

	removeSheet(nameOrId: string): boolean {
		const index = this.sheets.findIndex((s) => s.name === nameOrId || s.id === nameOrId)
		if (index === -1) return false
		this.sheets.splice(index, 1)
		return true
	}
}

export function createWorkbook(id?: WorkbookId): Workbook {
	return new Workbook(id)
}
