import { ascendError } from '@ascend/schema'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, formatCellValue, heading } from '../output/pretty.ts'
import { openWorkbookDocumentWithProgress } from '../progress.ts'

export const usage = `Usage: ascend trace <file> <ref> [flags]

  Trace precedents and dependents for a cell.

Arguments:
  <file>          Path to the workbook file
  <ref>           Cell reference (e.g. Sheet1!A1)

Flags:
  --json          Output as JSON
  --max-depth <N> Limit trace depth
`

export async function traceCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	if (!file || !cellRef) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required trace input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'trace',
					required: ['file', 'cell'],
					missing: [...(!file ? ['file'] : []), ...(!cellRef ? ['cell'] : [])],
					workflow: ['reopen', 'verify'],
				},
				suggestedFix: 'Run ascend trace <file> <sheet!cell> --json for formula dependency proof.',
			}),
			flags,
		)
		return 1
	}

	const { document: session } = await openWorkbookDocumentWithProgress(file, { mode: 'formula' })
	const maxDepth = parseOptionalInt(flags.get('max-depth'))
	if (flags.has('max-depth') && (maxDepth == null || maxDepth < 0)) {
		cliError('Invalid --max-depth. Use a non-negative integer.', flags)
		return 1
	}
	const validatedMaxDepth = maxDepth ?? undefined
	const result = session.trace(
		cellRef,
		validatedMaxDepth !== undefined ? { maxDepth: validatedMaxDepth } : undefined,
	)

	if (!result) {
		cliError(`Could not trace "${cellRef}"`, flags)
		return 1
	}

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		console.log(heading(`Trace: ${result.ref}`))
		if (result.formula) {
			console.log(bullet('Formula', `=${result.formula}`))
		}
		console.log(bullet('Value', formatCellValue(result.value)))
		console.log(heading('Precedents'))
		if (result.precedents.length === 0) {
			console.log('  (none)')
		} else {
			for (const node of result.precedents) {
				console.log(
					`  [${node.depth}] ${node.ref} value=${formatCellValue(node.value)}${node.formula ? ` formula=${node.formula}` : ''}`,
				)
			}
		}
		console.log(heading('Dependents'))
		if (result.dependents.length === 0) {
			console.log('  (none)')
		} else {
			for (const node of result.dependents) {
				console.log(
					`  [${node.depth}] ${node.ref} value=${formatCellValue(node.value)}${node.formula ? ` formula=${node.formula}` : ''}`,
				)
			}
		}
	}
	return 0
}

function parseOptionalInt(value: string | undefined): number | null | undefined {
	if (value === undefined || value === '') return undefined
	const parsed = Number.parseInt(value, 10)
	if (Number.isNaN(parsed)) return null
	return parsed
}
