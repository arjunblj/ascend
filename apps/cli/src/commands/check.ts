import { jsonOut } from '../output/json.ts'
import { table } from '../output/pretty.ts'
import { openWorkbookSessionWithProgress } from '../progress.ts'

export const usage = `Usage: ascend check <file> [flags]

  Run structural checks on a workbook.

Arguments:
  <file>          Path to the workbook file

Flags:
  --json          Output as JSON
`

export async function checkCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend check <file>')
		return 1
	}

	const { session: wb } = await openWorkbookSessionWithProgress(file, { mode: 'formula' })
	const result = wb.check()

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else if (result.valid) {
		console.log(`${file}: all checks passed`)
	} else {
		console.log(`${file}: ${result.issues.length} issue(s) found\n`)
		console.log(
			table(
				['Severity', 'Message', 'Ref'],
				result.issues.map((i) => [i.severity, i.message, i.ref ?? '']),
			),
		)
	}

	return result.valid ? 0 : 2
}
