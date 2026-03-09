import { writeFile } from 'node:fs/promises'
import { jsonOut } from '../output/json.ts'
import { openWorkbookWithProgress, withProgress } from '../progress.ts'

export const usage = `Usage: ascend export <file> <output> [flags]

  Export workbook to another format.

Arguments:
  <file>              Path to the source workbook
  <output>            Path to the output file

Flags:
  --format csv|json   Output format (inferred from extension if omitted)
  --sheet <name>      Sheet to export (for CSV)
  --json              Output status as JSON
`

export async function exportCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const output = args[1]
	if (!file || !output) {
		console.error('Usage: ascend export <file> <output> [--format csv|json]')
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const format = flags.get('format') ?? inferFormat(output)

	if (format === 'json') {
		const { value: data } = await withProgress('Serializing workbook', () =>
			JSON.stringify(wb.toJSON(), null, 2),
		)
		await withProgress(`Writing ${output}`, () => writeFile(output, data, 'utf-8'))
	} else if (format === 'csv') {
		const sheetName = flags.get('sheet')
		const { value: csv } = await withProgress('Rendering CSV export', () =>
			wb.toCsv(sheetName ? { sheet: sheetName } : undefined),
		)
		await withProgress(`Writing ${output}`, () => writeFile(output, csv, 'utf-8'))
	} else {
		await withProgress(`Saving ${output}`, () => wb.save(output))
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
