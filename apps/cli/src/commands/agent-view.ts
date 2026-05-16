import { ascendError } from '@ascend/schema'
import { WorkbookDocument } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, formatCellValue, heading } from '../output/pretty.ts'
import { openWorkbookDocumentWithProgress } from '../progress.ts'

export const usage = `Usage: ascend agent-view <file> [flags]

  Get an AI-friendly summary of a sheet (column summaries, sample rows, formula patterns).

Arguments:
  <file>              Path to the workbook file

Flags:
  --sheet <name>      Sheet name (defaults to first sheet)
  --range <range>     Cell range (e.g. A1:Z100, defaults to used range)
  --tokens <count>    Approximate maximum tokens for JSON/pretty output
  --json              Output as JSON
`

export async function agentViewCommand(
	args: string[],
	flags: Map<string, string>,
): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required agent-view input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'agent-view',
					required: ['file'],
					missing: ['file'],
					workflow: ['inspect', 'read', 'plan'],
				},
				suggestedFix:
					'Run ascend agent-view <file> --sheet <name> --range <range> --json after a trust preflight.',
			}),
			flags,
		)
		return 1
	}

	const sheetName = flags.get('sheet')
	const range = flags.get('range')
	const maxApproxTokens = parsePositiveInt(flags.get('tokens'))
	if (flags.has('tokens') && maxApproxTokens == null) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Invalid --tokens. Use a positive integer.', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'agent-view',
					flag: 'tokens',
					received: flags.get('tokens'),
					expected: 'positive integer',
					workflow: ['inspect', 'agent-view', 'plan'],
				},
				suggestedFix:
					'Use --tokens 1 or another positive integer to cap agent-view output for context.',
			}),
			flags,
		)
		return 1
	}

	const { document: wb } = await openWorkbookDocumentWithProgress(file, {
		mode: 'formula',
		...(sheetName ? { sheets: [sheetName] } : {}),
	})

	const targetSheet = sheetName ?? wb.sheets[0]
	if (!targetSheet) {
		cliError('No sheets in workbook', flags)
		return 1
	}

	const sheet = wb.sheet(targetSheet)
	if (!sheet) {
		cliError(
			agentViewSheetNotFoundError(
				targetSheet,
				await loadAvailableSheetsForAgentView(file, wb.sheets),
			),
			flags,
		)
		return 1
	}

	const usedRange = sheet.usedRange()
	const effectiveRange = range ?? (usedRange ? rangeRefToString(usedRange) : 'A1:A1')

	const view = wb.agentView(targetSheet, effectiveRange, {
		...(maxApproxTokens != null ? { maxApproxTokens } : {}),
	})
	if (!view) {
		cliError(`Could not generate agent view for ${targetSheet}!${effectiveRange}`, flags)
		return 1
	}

	if (flags.has('json')) {
		console.log(jsonOut(view))
		return 0
	}

	console.log(heading(`Agent View: ${targetSheet}!${effectiveRange}`))
	console.log(bullet('Rows', view.rowCount))
	console.log(bullet('Columns', view.colCount))
	console.log(bullet('Non-empty cells', view.nonEmptyCount))
	console.log(bullet('Formula cells', view.formulaCount))
	if (view.budget) {
		console.log(bullet('Requested tokens', view.budget.requestedApproxTokens))
		console.log(bullet('Estimated tokens', view.budget.estimatedApproxTokens))
		console.log(bullet('Truncated', view.budget.truncated ? 'yes' : 'no'))
	}
	if (view.distinctFunctions.length > 0) {
		console.log(bullet('Functions', view.distinctFunctions.join(', ')))
	}

	if (view.formulaPatterns.length > 0) {
		console.log('')
		console.log(heading('Formula Patterns'))
		for (const p of view.formulaPatterns) {
			console.log(bullet(p.pattern, `${p.count} occurrence(s)`))
		}
	}

	if (view.columns.length > 0) {
		console.log('')
		console.log(heading('Column Summaries'))
		for (const col of view.columns) {
			const header = col.header != null ? formatSample(col.header) : '(none)'
			const samples =
				col.sampleValues.length > 0
					? col.sampleValues.map((s) => formatSample(s)).join(', ')
					: '(none)'
			console.log(
				bullet(
					`${col.ref} (${col.kind})`,
					`header=${header} nonEmpty=${col.nonEmptyCount} formulas=${col.formulaCount}`,
				),
			)
			if (samples !== '(none)') console.log(`    samples: ${samples}`)
		}
	}

	if (view.samples.length > 0) {
		console.log('')
		console.log(heading('Sample Rows'))
		for (const row of view.samples) {
			const cells = row.cells.map((c) => formatCellValue(c.value, { display: true })).join(' | ')
			console.log(bullet(`Row ${row.row + 1}`, cells))
		}
	}

	return 0
}

async function loadAvailableSheetsForAgentView(
	file: string,
	fallbackSheets: readonly string[],
): Promise<readonly string[]> {
	if (fallbackSheets.length > 0) return fallbackSheets
	try {
		const workbook = await WorkbookDocument.open(file, { mode: 'metadata-only' })
		return workbook.sheets
	} catch {
		return fallbackSheets
	}
}

function agentViewSheetNotFoundError(sheetName: string, availableSheets: readonly string[]) {
	return ascendError('SHEET_NOT_FOUND', `Sheet "${sheetName}" not found`, {
		retryable: true,
		retryStrategy: availableSheets.length > 0 ? 'modified' : 'none',
		details: {
			command: 'agent-view',
			sheet: sheetName,
			availableSheets,
			workflow: ['inspect', 'agent-view', 'plan'],
		},
		suggestedFix:
			availableSheets.length > 0
				? `Retry with --sheet set to one of: ${availableSheets.join(', ')}.`
				: 'Run ascend inspect <file> --json to confirm workbook sheets before retrying agent-view.',
	})
}

function rangeRefToString(ref: {
	start: { row: number; col: number }
	end: { row: number; col: number }
}): string {
	return `${colToA1(ref.start.col)}${ref.start.row + 1}:${colToA1(ref.end.col)}${ref.end.row + 1}`
}

function colToA1(col: number): string {
	let n = col
	let label = ''
	while (n >= 0) {
		label = String.fromCharCode(65 + (n % 26)) + label
		n = Math.floor(n / 26) - 1
	}
	return label
}

function parsePositiveInt(value: string | undefined): number | null | undefined {
	if (value === undefined || value === '') return undefined
	if (!/^\d+$/.test(value)) return null
	const parsed = Number.parseInt(value, 10)
	if (!Number.isSafeInteger(parsed) || parsed < 1) return null
	return parsed
}

function formatSample(v: unknown): string {
	if (v === null || v === undefined) return ''
	if (typeof v === 'number') return String(v)
	if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
	if (typeof v === 'string') return v
	return String(v)
}
