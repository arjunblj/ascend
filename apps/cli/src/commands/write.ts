import { readFile } from 'node:fs/promises'
import { ascendError, type Operation } from '@ascend/schema'
import { cliError, jsonErr, jsonOut } from '../output/json.ts'
import { openWorkbookWithProgress, withProgress } from '../progress.ts'
import { buildSetCellOps, parseWriteSelector, resolveSheetName } from './mutation-helpers.ts'

export const usage = `Usage: ascend write <file> <selector> <values...> [flags]
       ascend write <file> --ops <file.json>

  Write cell values or apply operations to a workbook.

Arguments:
  <file>              Path to the workbook file
  <selector>          Cell reference (e.g. Sheet1!A1)
  <values...>         JSON-encoded values to write

Flags:
  --sheet <name>      Sheet name when selector has no sheet qualifier
  --ops <file.json>   Apply operations from a JSON file instead
  --json              Output as JSON
`

export async function writeCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(missingWriteInputError(['file']), flags)
		return 1
	}

	const opsFile = flags.get('ops')
	const selectorArg = args[1]
	const valuesStr = args[2]
	if (!opsFile && (!selectorArg || !valuesStr)) {
		cliError(
			missingWriteInputError([
				...(!selectorArg ? ['selector'] : []),
				...(!valuesStr ? ['values'] : []),
			]),
			flags,
		)
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)

	if (opsFile) {
		const raw = await readFile(opsFile, 'utf-8')
		const ops: Operation[] = JSON.parse(raw)
		const result = wb.apply(ops)
		if (result.errors.length > 0) {
			if (flags.has('json')) {
				const first = result.errors[0]
				console.log(
					jsonErr(
						first ??
							ascendError('VALIDATION_ERROR', 'Failed to apply operations', {
								details: { apply: result },
							}),
					),
				)
			} else {
				for (const e of result.errors) cliError(e.message, flags)
			}
			return 1
		}
		if (result.recalcRequired) {
			const { value: recalc } = await withProgress('Recalculating formulas', () => wb.recalc())
			if (recalc.errors.length > 0) {
				if (flags.has('json')) {
					const first = recalc.errors[0]
					console.log(
						jsonErr(
							first
								? {
										...first.error,
										...(first.error.refs ? {} : { refs: [first.ref] }),
										details: { ...(first.error.details ?? {}), recalc },
									}
								: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed', {
										details: { recalc },
									}),
						),
					)
				} else {
					for (const error of recalc.errors) cliError(`${error.ref}: ${error.error.message}`, flags)
				}
				return 1
			}
		}
		await withProgress(`Saving ${file}`, () => wb.save(file))
		if (flags.has('json')) {
			console.log(jsonOut(result))
		} else {
			console.log(`Applied ${ops.length} operation(s) to ${file}`)
		}
		return 0
	}

	if (!selectorArg || !valuesStr) {
		cliError(
			missingWriteInputError([
				...(!selectorArg ? ['selector'] : []),
				...(!valuesStr ? ['values'] : []),
			]),
			flags,
		)
		return 1
	}

	const values: unknown = JSON.parse(valuesStr)
	const selector = parseWriteSelector(selectorArg, flags.get('sheet'))
	const sheetName = resolveSheetName(wb, selector.sheet)
	if (!sheetName) {
		cliError(
			wb.sheets.length === 0
				? 'No sheets in workbook'
				: 'Multiple sheets available; specify a sheet explicitly',
			flags,
		)
		return 1
	}

	const ops = buildSetCellOps(sheetName, selector.ref, values)
	const result = wb.apply(ops)
	if (result.errors.length > 0) {
		if (flags.has('json')) {
			const first = result.errors[0]
			console.log(
				jsonErr(
					first ??
						ascendError('VALIDATION_ERROR', 'Failed to apply operations', {
							details: { apply: result },
						}),
				),
			)
		} else {
			for (const e of result.errors) cliError(e.message, flags)
		}
		return 1
	}
	if (result.recalcRequired) {
		const { value: recalc } = await withProgress('Recalculating formulas', () => wb.recalc())
		if (recalc.errors.length > 0) {
			if (flags.has('json')) {
				const first = recalc.errors[0]
				console.log(
					jsonErr(
						first
							? {
									...first.error,
									...(first.error.refs ? {} : { refs: [first.ref] }),
									details: { ...(first.error.details ?? {}), recalc },
								}
							: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed', {
									details: { recalc },
								}),
					),
				)
			} else {
				for (const error of recalc.errors) cliError(`${error.ref}: ${error.error.message}`, flags)
			}
			return 1
		}
	}

	await withProgress(`Saving ${file}`, () => wb.save(file))
	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		const updateCount = ops[0]?.op === 'setCells' ? ops[0].updates.length : 0
		console.log(`Wrote ${updateCount} cell(s) to ${file}`)
	}
	return 0
}

function missingWriteInputError(missing: readonly string[]) {
	return ascendError('INVALID_ARGUMENT', 'Missing required write input', {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			command: 'write',
			required: ['file', 'selector or ops', 'values when selector is used'],
			missing,
			workflow: ['write', 'reopen', 'verify'],
		},
		suggestedFix:
			'Prefer ascend plan/commit for auditable edits. For direct writes, run ascend write <file> <selector> <json-values> --json or ascend write <file> --ops <file.json> --json.',
	})
}
