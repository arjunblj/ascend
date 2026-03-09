import { jsonOut } from '../output/json.ts'
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
		console.error('Usage: ascend calc <file>')
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const { value: result } = await withProgress('Recalculating formulas', () => wb.recalc())

	if (result.errors.length > 0) {
		for (const e of result.errors) console.error(`${e.ref}: ${e.error.message}`)
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
