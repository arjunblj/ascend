import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { formatCellValue, table } from '../output/pretty.ts'

export async function readCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const range = args[1]
	if (!file || !range) {
		console.error('Usage: ascend read <file> <range> [--sheet <name>]')
		return 1
	}

	const requestedSheet = flags.get('sheet')
	const wb = await AscendWorkbook.open(
		file,
		requestedSheet ? { sheets: [requestedSheet] } : undefined,
	)
	const sheetName = requestedSheet ?? wb.sheets[0]
	if (!sheetName) {
		console.error('No sheets in workbook')
		return 1
	}

	const sheet = wb.sheet(sheetName)
	if (!sheet) {
		console.error(`Sheet "${sheetName}" not found`)
		return 1
	}

	const info = sheet.range(range)

	if (flags.has('json')) {
		const data = info.cells.map((c) => ({
			ref: c.ref,
			value: c.value,
			formula: c.formula,
		}))
		console.log(jsonOut(data))
		return 0
	}

	const grid: string[][] = []
	const cellMap = new Map(info.cells.map((cell) => [`${cell.row}:${cell.col}`, cell] as const))
	for (let r = 0; r < info.rowCount; r++) {
		const row: string[] = []
		for (let c = 0; c < info.colCount; c++) {
			const cell = cellMap.get(`${info.ref.start.row + r}:${info.ref.start.col + c}`)
			row.push(cell ? formatCellValue(cell.value) : '')
		}
		grid.push(row)
	}

	const headers = Array.from({ length: info.colCount }, (_, i) => `Col${i + 1}`)
	console.log(table(headers, grid))
	return 0
}
