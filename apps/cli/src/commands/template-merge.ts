import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { ascendError } from '@ascend/schema'
import { AscendWorkbook } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'

export const usage = `Usage: ascend template-merge <file> --data <json-or-file> [flags]

  Compile {{key}} workbook template placeholders into replayable operations.

Arguments:
  <file>              Path to the workbook file

Flags:
  --data <json|file>  JSON object or path to a JSON file with scalar merge values
  --sheet <name>      Limit merge scan to one sheet
  --values-only       Omit formulas
  --formulas-only     Omit literal values
  --open <text>       Placeholder opening delimiter (default: {{)
  --close <text>      Placeholder closing delimiter (default: }})
  --json              Output as JSON
`

export async function templateMergeCommand(
	args: string[],
	flags: Map<string, string>,
): Promise<number> {
	const file = args[0]
	const dataArg = flags.get('data')
	if (!file || !dataArg) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required template-merge input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'template-merge',
					required: ['file', 'data'],
					missing: [...(!file ? ['file'] : []), ...(!dataArg ? ['data'] : [])],
					workflow: ['inspect', 'template-merge', 'plan'],
				},
				suggestedFix:
					'Run ascend template-merge <file> --data <json-or-file> --json, then inspect replayable ops before commit.',
			}),
			flags,
		)
		return 1
	}
	if (flags.has('values-only') && flags.has('formulas-only')) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Use either --values-only or --formulas-only, not both.', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'template-merge',
					conflictingFlags: ['values-only', 'formulas-only'],
					workflow: ['inspect', 'template-merge', 'plan'],
				},
				suggestedFix:
					'Remove one flag so template placeholders are scanned in values, formulas, or both.',
			}),
			flags,
		)
		return 1
	}

	const data = await readTemplateData(dataArg)
	if (!data || Array.isArray(data) || typeof data !== 'object') {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Template data must be a JSON object.', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'template-merge',
					flag: 'data',
					expected: 'JSON object',
					workflow: ['inspect', 'template-merge', 'plan'],
				},
				suggestedFix:
					'Pass --data with an object such as {"name":"Acme"} or a path to a JSON object file.',
			}),
			flags,
		)
		return 1
	}

	const wb = await AscendWorkbook.open(file)
	const result = wb.templateMerge(data as Record<string, unknown>, {
		...(flags.get('sheet') ? { sheets: [flags.get('sheet') as string] } : {}),
		...(flags.has('values-only') ? { includeFormulas: false } : {}),
		...(flags.has('formulas-only') ? { includeValues: false } : {}),
		...(flags.get('open') || flags.get('close')
			? {
					delimiters: {
						...(flags.get('open') ? { open: flags.get('open') as string } : {}),
						...(flags.get('close') ? { close: flags.get('close') as string } : {}),
					},
				}
			: {}),
	})
	if (flags.has('json')) {
		console.log(jsonOut(result))
		return result.replayable ? 0 : 2
	}

	console.log(heading(`Template merge: ${file}`))
	console.log(bullet('Sheets', result.sheetCount))
	console.log(bullet('Template cells', result.cellCount))
	console.log(bullet('Template formulas', result.formulaCount))
	console.log(bullet('Replacements', result.replacementCount))
	console.log(bullet('Operations', result.ops.length))
	console.log(bullet('Replayable', result.replayable ? 'yes' : 'no'))
	if (result.unresolved.length > 0) {
		console.log(bullet('Unresolved placeholders', result.unresolved.length))
		for (const placeholder of result.unresolved.slice(0, 10)) {
			console.log(bullet(`${placeholder.sheet}!${placeholder.ref}`, placeholder.placeholder))
		}
	}
	if (result.unsupported.length > 0) {
		console.log(bullet('Unsupported cells', result.unsupported.length))
		for (const cell of result.unsupported.slice(0, 10)) {
			console.log(bullet(`${cell.sheet}!${cell.ref}`, `${cell.valueKind}: ${cell.reason}`))
		}
	}
	return result.replayable ? 0 : 2
}

async function readTemplateData(dataArg: string): Promise<unknown> {
	try {
		return JSON.parse(dataArg)
	} catch {
		const raw = await readFile(existsSync(dataArg) ? dataArg : dataArg, 'utf-8')
		return JSON.parse(raw)
	}
}
