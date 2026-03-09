import { jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import { openWorkbookWithProgress } from '../progress.ts'

export const usage = `Usage: ascend trace <file> <ref> [flags]

  Trace precedents and dependents for a cell.

Arguments:
  <file>          Path to the workbook file
  <ref>           Cell reference (e.g. Sheet1!A1)

Flags:
  --json          Output as JSON
`

export async function traceCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	if (!file || !cellRef) {
		console.error('Usage: ascend trace <file> <cell>')
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const result = wb.trace(cellRef)

	if (!result) {
		console.error(`Could not trace "${cellRef}"`)
		return 1
	}

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		console.log(heading(`Trace: ${result.ref}`))
		if (result.formula) {
			console.log(bullet('Formula', `=${result.formula}`))
		}
		console.log(
			bullet('Depends on', result.dependsOn.length > 0 ? result.dependsOn.join(', ') : '(none)'),
		)
		console.log(
			bullet('Feeds into', result.feedsInto.length > 0 ? result.feedsInto.join(', ') : '(none)'),
		)
	}
	return 0
}
