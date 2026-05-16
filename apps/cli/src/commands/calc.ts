import { ascendError } from '@ascend/schema'
import { cliError, jsonErr, jsonOut } from '../output/json.ts'
import { openWorkbookWithProgress, withProgress } from '../progress.ts'

export const usage = `Usage: ascend calc <file> [flags]

  Recalculate all formulas in the workbook.

Arguments:
  <file>          Path to the workbook file

Flags:
  --json          Output as JSON
`

export async function calcCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required calc input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'calc',
					required: ['file'],
					missing: ['file'],
					workflow: ['reopen', 'verify'],
				},
				suggestedFix: 'Run ascend calc <file> --json after reopening a workbook with formulas.',
			}),
			flags,
		)
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const { value: result } = await withProgress('Recalculating formulas', () => wb.recalc())

	if (result.errors.length > 0) {
		if (flags.has('json')) {
			const first = result.errors[0]
			console.log(
				jsonErr(
					first
						? {
								...first.error,
								...(first.error.refs ? {} : { refs: [first.ref] }),
								details: { ...(first.error.details ?? {}), recalc: result },
							}
						: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed', {
								details: { recalc: result },
							}),
				),
			)
		} else {
			for (const e of result.errors) cliError(`${e.ref}: ${e.error.message}`, flags)
		}
		return 1
	}

	await withProgress(`Saving ${file}`, () => wb.save(file))

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		console.log(`Recalculated ${file} in ${result.duration}ms`)
		console.log(`  Changed: ${result.changed.length} cell(s)`)
		if (result.errors.length > 0) {
			console.log(`  Errors: ${result.errors.length}`)
		}
	}
	return 0
}
