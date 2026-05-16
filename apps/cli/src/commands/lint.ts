import { ascendError } from '@ascend/schema'
import { cliError, jsonOut } from '../output/json.ts'
import { table } from '../output/pretty.ts'
import { openWorkbookDocumentWithProgress } from '../progress.ts'

export const usage = `Usage: ascend lint <file> [flags]

  Run formula lint checks on a workbook.

Arguments:
  <file>          Path to the workbook file

Flags:
  --json          Output as JSON
`

export async function lintCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required lint input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'lint',
					required: ['file'],
					missing: ['file'],
					workflow: ['reopen', 'verify'],
				},
				suggestedFix: 'Run ascend lint <file> --json after committing or reopening a workbook.',
			}),
			flags,
		)
		return 1
	}

	const { document: wb } = await openWorkbookDocumentWithProgress(file, { mode: 'formula' })
	const result = wb.lint()

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else if (result.clean) {
		console.log(`${file}: no lint warnings`)
	} else {
		console.log(`${file}: ${result.warnings.length} warning(s)\n`)
		console.log(
			table(
				['Rule', 'Message', 'Ref'],
				result.warnings.map((w) => [w.rule, w.message, w.ref ?? '']),
			),
		)
	}

	return result.clean ? 0 : 2
}
