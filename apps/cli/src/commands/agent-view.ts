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
  --json              Output as JSON
`

export async function agentViewCommand(
	args: string[],
	flags: Map<string, string>,
): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError('Usage: ascend agent-view <file> [--sheet <name>] [--range <range>]', flags)
		return 1
	}

	const sheetName = flags.get('sheet')
	const range = flags.get('range')

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
		cliError(`Sheet "${targetSheet}" not found`, flags)
		return 1
	}

	const usedRange = sheet.usedRange()
	const effectiveRange = range ?? (usedRange ? rangeRefToString(usedRange) : 'A1:A1')

	const view = wb.agentView(targetSheet, effectiveRange)
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

function formatSample(v: unknown): string {
	if (v === null || v === undefined) return ''
	if (typeof v === 'number') return String(v)
	if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
	if (typeof v === 'string') return v
	return String(v)
}
