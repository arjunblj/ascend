import { writeFile } from 'node:fs/promises'
import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'

export async function exportCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const output = args[1]
	if (!file || !output) {
		console.error('Usage: ascend export <file> <output> [--format csv|json]')
		return 1
	}

	const wb = await AscendWorkbook.open(file)
	const format = flags.get('format') ?? inferFormat(output)

	if (format === 'json') {
		const data = wb.toJSON()
		await writeFile(output, JSON.stringify(data, null, 2), 'utf-8')
	} else if (format === 'csv') {
		const sheetName = flags.get('sheet')
		const csv = wb.toCsv(sheetName ? { sheet: sheetName } : undefined)
		await writeFile(output, csv, 'utf-8')
	} else {
		await wb.save(output)
	}

	if (flags.has('json')) {
		console.log(jsonOut({ exported: output, format }))
	} else {
		console.log(`Exported ${file} → ${output} (${format})`)
	}
	return 0
}

function inferFormat(path: string): string {
	const ext = path.split('.').pop()?.toLowerCase() ?? ''
	if (ext === 'csv' || ext === 'tsv') return 'csv'
	if (ext === 'json') return 'json'
	return ext
}
