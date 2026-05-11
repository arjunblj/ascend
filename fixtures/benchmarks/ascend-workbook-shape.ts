import type { Workbook } from '../../packages/core/src/index.ts'

export type WorkbookSummarySheetEntry = {
	readonly name: string
	readonly sheet?: Workbook['sheets'][number]
}

export function workbookSheetEntriesForSummary(workbook: Workbook): WorkbookSummarySheetEntry[] {
	const worksheetsByName = new Map(workbook.sheets.map((sheet) => [sheet.name, sheet]))
	const preservedEntries = workbook.preservedXml?.sheetEntries
	if (preservedEntries?.length) {
		return preservedEntries.map((entry) => {
			if (entry.kind !== 'worksheet') return { name: entry.name }
			const sheet = worksheetsByName.get(entry.name)
			return sheet ? { name: entry.name, sheet } : { name: entry.name }
		})
	}
	return [
		...workbook.sheets.map((sheet) => ({ name: sheet.name, sheet })),
		...workbook.chartSheets.map((sheet) => ({ name: sheet.name })),
		...workbook.macroSheets.map((sheet) => ({ name: sheet.name })),
	]
}
