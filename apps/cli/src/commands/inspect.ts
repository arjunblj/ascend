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
			console.log(bullet('Rows', sheet.rowCount))
			console.log(bullet('Columns', sheet.colCount))
			console.log(bullet('Cells', sheet.cellCount))
			console.log(bullet('Tables', sheet.tableCount))
			console.log(bullet('Frozen panes', sheet.hasFrozenPanes ? 'yes' : 'no'))
		}
		return 0
	}

	if (flags.has('json')) {
		console.log(jsonOut(info))
		return 0
	}

	console.log(heading(`Workbook: ${file}`))
	console.log(bullet('Format', info.sourceFormat))
	console.log(bullet('Sheets', info.sheetCount))
	console.log(bullet('Total cells', info.cellCount))
	if (info.definedNames.length > 0) {
		console.log(bullet('Defined names', info.definedNames.join(', ')))
	}

	if (info.sheets.length > 0) {
		console.log('')
		console.log(
			table(
				['Sheet', 'Rows', 'Cols', 'Cells', 'Tables'],
				info.sheets.map((s) => [
					s.name,
					String(s.rowCount),
					String(s.colCount),
					String(s.cellCount),
					String(s.tableCount),
				]),
			),
		)
	}

	return 0
}
