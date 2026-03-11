import { readFile } from 'node:fs/promises'
import { ascendError, type CellValue, type Operation } from '@ascend/schema'
import { jsonErr, jsonOut } from '../output/json.ts'
import { bullet, heading, table } from '../output/pretty.ts'
import { openWorkbookWithProgress } from '../progress.ts'
import { buildSetCellOps, parseWriteSelector, resolveSheetName } from './mutation-helpers.ts'

export const usage = `Usage: ascend preview <file> <selector> <values...> [flags]
       ascend preview <file> --ops <file.json>

  Preview workbook changes without saving them.

Arguments:
  <file>              Path to the workbook file
  <selector>          Cell reference (e.g. Sheet1!A1)
  <values...>         JSON-encoded values to preview

Flags:
  --sheet <name>      Sheet name when selector has no sheet qualifier
  --ops <file.json>   Preview operations from a JSON file instead
  --json              Output as JSON
`

export async function previewCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend preview <file> <selector> <json-values> [--sheet <name>]')
		console.error('       ascend preview <file> --ops <file.json>')
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const ops = await resolvePreviewOps(wb, args, flags)
	if (!ops) return 1

	const result = wb.preview(ops)
	if (flags.has('json')) {
		if (result.errors.length > 0) {
			const first = result.errors[0]
			console.log(
				jsonErr(
					first
						? {
								...first,
								details: { ...(first.details ?? {}), preview: result },
							}
						: ascendError('VALIDATION_ERROR', 'Preview failed', { details: { preview: result } }),
				),
			)
		} else {
			console.log(jsonOut(result))
		}
	} else {
		console.log(heading(`Preview: ${file}`))
		console.log(bullet('Sheets changed', result.sheetDiffs.length))
		console.log(bullet('Cell changes', result.cellChanges.length))
		console.log(bullet('Errors', result.errors.length))
		if (result.cellChanges.length > 0) {
			console.log('')
			console.log(
				table(
					['Ref', 'Before', 'After', 'Formula Before', 'Formula After'],
					result.cellChanges
						.slice(0, 12)
						.map((change) => [
							change.ref,
							formatValue(change.before),
							formatValue(change.after),
							change.formulaBefore ?? '',
							change.formulaAfter ?? '',
						]),
				),
			)
			if (result.cellChanges.length > 12) {
				console.log(`\nShowing first 12 of ${result.cellChanges.length} cell change(s).`)
			}
		}
		if (result.errors.length > 0) {
			console.log('')
			for (const error of result.errors) {
				console.log(bullet('Error', error.message))
			}
		}
	}

	return result.errors.length === 0 ? 0 : 1
}

async function resolvePreviewOps(
	wb: Awaited<ReturnType<typeof openWorkbookWithProgress>>['workbook'],
	args: string[],
	flags: Map<string, string>,
): Promise<readonly Operation[] | null> {
	const opsFile = flags.get('ops')
	if (opsFile) {
		const raw = await readFile(opsFile, 'utf-8')
		return JSON.parse(raw) as Operation[]
	}

	const selectorArg = args[1]
	const valuesStr = args[2]
	if (!selectorArg || !valuesStr) {
		console.error('Usage: ascend preview <file> <selector> <json-values> [--sheet <name>]')
		return null
	}

	const values: unknown = JSON.parse(valuesStr)
	const selector = parseWriteSelector(selectorArg, flags.get('sheet'))
	const sheetName = resolveSheetName(wb, selector.sheet)
	if (!sheetName) {
		console.error(
			wb.sheets.length === 0
				? 'No sheets in workbook'
				: 'Multiple sheets available; specify a sheet explicitly',
		)
		return null
	}
	return buildSetCellOps(sheetName, selector.ref, values)
}

function formatValue(value: CellValue): string {
	switch (value.kind) {
		case 'empty':
			return ''
		case 'number':
		case 'string':
		case 'boolean':
		case 'error':
			return String(value.value ?? '')
		case 'date':
			return String(value.serial ?? '')
		case 'richText':
			return '[richText]'
		case 'array':
			return '[array]'
		default:
			return ''
	}
}
