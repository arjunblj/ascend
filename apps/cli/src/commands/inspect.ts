import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { bullet, heading, table } from '../output/pretty.ts'

export async function inspectCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend inspect <file> [sheet]')
		return 1
	}

	const sheetArg = args[1] ?? flags.get('sheet')
	const wb = await AscendWorkbook.open(
		file,
		sheetArg ? { sheets: [sheetArg] } : { mode: 'metadata-only' },
	)
	const info = wb.inspect()

	if (sheetArg) {
		const sheet = info.sheets.find((s) => s.name === sheetArg)
		if (!sheet) {
			console.error(`Sheet "${sheetArg}" not found`)
			return 1
		}
		if (flags.has('json')) {
			console.log(jsonOut(sheet))
		} else {
			console.log(heading(`Sheet: ${sheet.name}`))
			console.log(bullet('Cell data loaded', sheet.cellDataLoaded ? 'yes' : 'no'))
			console.log(bullet('Rows', formatCount(sheet.rowCount)))
			console.log(bullet('Columns', formatCount(sheet.colCount)))
			console.log(bullet('Cells', formatCount(sheet.cellCount)))
			console.log(bullet('Tables', formatCount(sheet.tableCount)))
			console.log(bullet('Comments', formatCount(sheet.commentCount)))
			console.log(bullet('Conditional formats', formatCount(sheet.conditionalFormatCount)))
			console.log(bullet('Data validations', formatCount(sheet.dataValidationCount)))
			console.log(bullet('Images', formatCount(sheet.imageCount)))
			console.log(bullet('Frozen panes', formatLoadedBool(sheet.hasFrozenPanes)))
			console.log(bullet('Column widths', formatCount(sheet.colWidthCount)))
			console.log(bullet('Row heights', formatCount(sheet.rowHeightCount)))
			console.log(bullet('Hyperlinks', formatCount(sheet.hyperlinkCount)))
			console.log(bullet('Ignored errors', formatCount(sheet.ignoredErrorCount)))
			console.log(bullet('Auto filter', formatLoadedBool(sheet.hasAutoFilter)))
			console.log(bullet('Drawing refs', formatLoadedBool(sheet.hasDrawingRefs)))
			console.log(bullet('Page metadata', formatLoadedBool(sheet.hasPageMetadata)))
			console.log(bullet('Protection', formatLoadedBool(sheet.hasProtection)))
		}
		return 0
	}

	if (flags.has('json')) {
		console.log(jsonOut(info))
		return 0
	}

	console.log(heading(`Workbook: ${file}`))
	console.log(bullet('Format', info.sourceFormat))
	console.log(bullet('Source sheets', info.sheetCount))
	console.log(bullet('Loaded sheets', info.loadedSheetCount))
	console.log(bullet('Load mode', info.load.mode))
	console.log(bullet('Partial view', info.load.isPartial ? 'yes' : 'no'))
	console.log(bullet('Cell data loaded', info.load.cellsHydrated ? 'yes' : 'no'))
	console.log(bullet('Total cells', formatCount(info.cellCount)))
	console.log(bullet('Comments', formatCount(info.commentCount)))
	console.log(bullet('Conditional formats', formatCount(info.conditionalFormatCount)))
	console.log(bullet('Data validations', formatCount(info.dataValidationCount)))
	console.log(bullet('Images', formatCount(info.imageCount)))
	console.log(bullet('Workbook views', info.workbookViewCount))
	console.log(bullet('External references', info.externalReferenceCount))
	console.log(bullet('Workbook protection', info.hasWorkbookProtection ? 'yes' : 'no'))
	console.log(bullet('Cell styles', info.styleSummary.cellXfCount))
	console.log(bullet('Diff styles', info.styleSummary.dxfCount))
	console.log(bullet('Theme part', info.themeSummary.hasThemePart ? 'yes' : 'no'))
	console.log(bullet('Theme colors', info.themeSummary.colorCount))
	if (info.themeSummary.name) {
		console.log(bullet('Theme name', info.themeSummary.name))
	}
	if (info.themeSummary.majorFontLatin || info.themeSummary.minorFontLatin) {
		console.log(
			bullet(
				'Theme fonts',
				[
					info.themeSummary.majorFontLatin
						? `major=${info.themeSummary.majorFontLatin}`
						: undefined,
					info.themeSummary.minorFontLatin
						? `minor=${info.themeSummary.minorFontLatin}`
						: undefined,
				]
					.filter(Boolean)
					.join(', '),
			),
		)
	}
	console.log(bullet('Compatibility', info.compatibility.status))
	if (info.definedNames.length > 0) {
		console.log(bullet('Defined names', info.definedNames.join(', ')))
	}

	if (info.sheets.length > 0) {
		console.log('')
		console.log(
			table(
				['Sheet', 'Loaded', 'Rows', 'Cols', 'Cells', 'Tables', 'CF', 'DV', 'Images', 'Filter'],
				info.sheets.map((s) => [
					s.name,
					s.cellDataLoaded ? 'yes' : 'no',
					formatCount(s.rowCount),
					formatCount(s.colCount),
					formatCount(s.cellCount),
					formatCount(s.tableCount),
					formatCount(s.conditionalFormatCount),
					formatCount(s.dataValidationCount),
					formatCount(s.imageCount),
					formatLoadedBool(s.hasAutoFilter),
				]),
			),
		)
	}

	return 0
}

function formatCount(value: number | null): string {
	return value === null ? 'unknown (not hydrated)' : String(value)
}

function formatLoadedBool(value: boolean | null): string {
	if (value === null) return 'unknown (not hydrated)'
	return value ? 'yes' : 'no'
}
