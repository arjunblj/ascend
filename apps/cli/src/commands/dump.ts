import { ascendError } from '@ascend/schema'
import { AscendWorkbook } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'

export const usage = `Usage: ascend dump <file> [flags]

  Dump supported workbook cells and formulas as a replayable operation batch.

Arguments:
  <file>              Path to the workbook file

Flags:
  --sheet <name>      Limit dump to one sheet
  --values-only       Omit formulas
  --formulas-only     Omit literal values
  --json              Output as JSON
`

export async function dumpCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required dump input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'dump',
					required: ['file'],
					missing: ['file'],
					workflow: ['inspect', 'dump', 'plan'],
				},
				suggestedFix:
					'Run ascend dump <file> --json to produce a replayable operation batch for supported cells and formulas.',
			}),
			flags,
		)
		return 1
	}
	if (flags.has('values-only') && flags.has('formulas-only')) {
		cliError('Use either --values-only or --formulas-only, not both.', flags)
		return 1
	}

	const wb = await AscendWorkbook.open(file)
	const result = wb.dumpBatch({
		...(flags.get('sheet') ? { sheets: [flags.get('sheet') as string] } : {}),
		...(flags.has('values-only') ? { includeFormulas: false } : {}),
		...(flags.has('formulas-only') ? { includeValues: false } : {}),
	})
	if (flags.has('json')) {
		console.log(jsonOut(result))
		return result.replayable ? 0 : 2
	}

	console.log(heading(`Dump: ${file}`))
	console.log(bullet('Sheets', result.sheetCount))
	console.log(bullet('Cells', result.cellCount))
	console.log(bullet('Formulas', result.formulaCount))
	console.log(bullet('Operations', result.ops.length))
	console.log(bullet('Replayable', result.replayable ? 'yes' : 'no'))
	if (result.unsupported.length > 0) {
		console.log(bullet('Unsupported cells', result.unsupported.length))
		for (const cell of result.unsupported.slice(0, 10)) {
			console.log(bullet(`${cell.sheet}!${cell.ref}`, `${cell.valueKind}: ${cell.reason}`))
		}
	}
	return result.replayable ? 0 : 2
}
