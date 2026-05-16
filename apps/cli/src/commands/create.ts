import { ascendError } from '@ascend/schema'
import { AscendWorkbook } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { withProgress } from '../progress.ts'

export const usage = `Usage: ascend create <output.xlsx> [flags]

  Create a new empty .xlsx workbook.

Arguments:
  <output.xlsx>   Path to the output workbook file

Flags:
  --json          Output as JSON
`

export async function createCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required create input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'create',
					required: ['output'],
					missing: ['output'],
				},
				suggestedFix: 'Run ascend create <output.xlsx> --json to create a workbook.',
			}),
			flags,
		)
		return 1
	}

	const wb = AscendWorkbook.create()
	await withProgress(`Saving ${file}`, () => wb.save(file))

	if (flags.has('json')) {
		console.log(jsonOut({ created: file }))
	} else {
		console.log(`Created ${file}`)
	}
	return 0
}
