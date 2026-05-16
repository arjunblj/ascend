import { writeFile } from 'node:fs/promises'
import { ascendError } from '@ascend/schema'
import { inferExportFormat, normalizeExportFormat } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { openWorkbookWithProgress, withProgress } from '../progress.ts'

export const usage = `Usage: ascend export <file> <output> [flags]

  Export workbook to another format.

Arguments:
  <file>              Path to the source workbook
  <output>            Path to the output file

Flags:
  --format csv|tsv|json|xlsx|xlsm   Output format (inferred from extension if omitted)
  --sheet <name>      Sheet to export (for CSV/TSV)
  --json              Output status as JSON
`

export async function exportCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const output = args[1]
	if (!file || !output) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required export input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'export',
					required: ['file', 'output'],
					missing: [...(!file ? ['file'] : []), ...(!output ? ['output'] : [])],
					workflow: ['reopen', 'verify', 'export'],
				},
				suggestedFix:
					'Run ascend export <file> <output> --format csv|tsv|json|xlsx|xlsm --json after verifying the workbook.',
			}),
			flags,
		)
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const format = normalizeExportFormat(flags.get('format') ?? '') ?? inferExportFormat(output)
	if (!format) {
		cliError(
			ascendError(
				'INVALID_ARGUMENT',
				'Invalid export format. Use one of: csv, tsv, json, xlsx, xlsm',
				{
					retryable: true,
					retryStrategy: 'modified',
					details: {
						command: 'export',
						flag: flags.has('format') ? 'format' : 'output',
						received: flags.get('format') ?? output,
						allowed: ['csv', 'tsv', 'json', 'xlsx', 'xlsm'],
						workflow: ['reopen', 'verify', 'export'],
					},
					suggestedFix:
						'Pass --format csv, tsv, json, xlsx, or xlsm, or choose an output path with one of those extensions.',
				},
			),
			flags,
		)
		return 1
	}

	if (format === 'json') {
		const { value: data } = await withProgress('Serializing workbook', () =>
			JSON.stringify(wb.toJSON(), null, 2),
		)
		await withProgress(`Writing ${output}`, () => writeFile(output, data, 'utf-8'))
	} else if (format === 'csv' || format === 'tsv') {
		const sheetName = flags.get('sheet')
		const { value: csv } = await withProgress('Rendering CSV export', () =>
			wb.toCsv({
				...(sheetName ? { sheet: sheetName } : {}),
				...(format === 'tsv' ? { dialect: { delimiter: '\t' } } : {}),
			}),
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
