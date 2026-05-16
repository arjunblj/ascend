import { ascendError } from '@ascend/schema'
import { cliError, jsonOut } from '../output/json.ts'
import { openWorkbookDocumentWithProgress } from '../progress.ts'

export const usage = `Usage: ascend list <file> [flags]

  List sheets and tables in a workbook.

Arguments:
  <file>    Path to the workbook file

Flags:
  --json    Output as JSON
`

export async function listCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required list input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'list',
					required: ['file'],
					missing: ['file'],
					workflow: ['inspect', 'read'],
				},
				suggestedFix: 'Run ascend list <file> --json to enumerate workbook sheets and tables.',
			}),
			flags,
		)
		return 1
	}

	const result = await openWorkbookDocumentWithProgress(file, { mode: 'metadata-only' })
	if (!result) return 1

	const info = result.document.inspect()
	const sheets = info.sheets.map((s) => ({
		name: s.name,
		rows: s.rowCount,
		cols: s.colCount,
		tableCount: s.tableCount,
	}))

	if (flags.has('json')) {
		console.log(jsonOut({ sheets }))
		return 0
	}

	console.log(`\n${file}`)
	console.log(`${'─'.repeat(file.length)}`)
	for (const s of sheets) {
		const size = s.rows != null && s.cols != null ? `  ${s.rows} rows × ${s.cols} cols` : ''
		const tables = s.tableCount ? `, ${s.tableCount} table(s)` : ''
		console.log(`  • ${s.name}${size}${tables}`)
	}
	return 0
}
