import { ascendError, levenshtein } from '@ascend/schema'
import { formulaAssist, type WorkbookLoadInfo } from '@ascend/sdk'
import { cliError, jsonErr, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import {
	openWorkbookDocumentWithProgress,
	openWorkbookWithProgress,
	withProgress,
} from '../progress.ts'

const FORMULA_SUBCOMMANDS = ['show', 'assist', 'set', 'fill'] as const

export const usage = `Usage: ascend formula <subcommand> [args] [flags]

  Inspect and edit cell formulas.

Subcommands:
  show <file> <ref>           Show formula at a cell reference
  assist <expr>               Return formula diagnostics, completions, signature help, and code actions
  set <file> <ref> <expr>     Set a formula on a cell
  fill <file> <range> <expr>  Fill a formula across a range

Flags:
  --cursor <n>                    Zero-based cursor offset for assist
  --prefix <text>                 Function completion prefix for assist
  --completion-limit <n>          Maximum completions for assist
  --function-name <name>          Function signature lookup for assist
  --reference <ref>               Reference text to insert for assist
  --replace-reference-at-cursor   Replace active reference when inserting
  --cycle-reference               Return Excel F4-style reference cycling
  --json                          Output as JSON
`

export async function formulaCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const action = args[0]
	switch (action) {
		case 'show':
			return showFormula(args.slice(1), flags)
		case 'assist':
			return assistFormula(args.slice(1), flags)
		case 'set':
			return setFormula(args.slice(1), flags)
		case 'fill':
			return fillFormula(args.slice(1), flags)
		default: {
			cliError(unknownFormulaSubcommandError(action), flags)
			return 1
		}
	}
}

function unknownFormulaSubcommandError(action: string | undefined) {
	const received = action ?? '(missing)'
	const suggestion = suggestClosest(action ?? '', FORMULA_SUBCOMMANDS)
	const message = suggestion
		? `Unknown formula subcommand: ${received}. Did you mean "${suggestion}"?`
		: `Unknown formula subcommand: ${received}`
	return ascendError('INVALID_ARGUMENT', message, {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			command: 'formula',
			subcommand: received,
			allowedSubcommands: FORMULA_SUBCOMMANDS,
			workflow: ['inspect', 'formula-assist', 'plan'],
			...(suggestion ? { suggestion } : {}),
		},
		suggestedFix: suggestion
			? `Did you mean "${suggestion}"? Retry with subcommand "${suggestion}".`
			: 'Use one of: show, assist, set, fill.',
	})
}

async function showFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	if (!file || !cellRef) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required formula show input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'formula show',
					required: ['file', 'cell'],
					missing: [...(!file ? ['file'] : []), ...(!cellRef ? ['cell'] : [])],
					workflow: ['inspect', 'read', 'plan'],
				},
				suggestedFix: 'Run ascend formula show <file> <sheet!cell> --json.',
			}),
			flags,
		)
		return 1
	}

	const { document: session } = await openWorkbookDocumentWithProgress(file, { mode: 'formula' })
	const info = session.formula(cellRef)
	if (!info) {
		cliError(formulaNotFoundError(cellRef, session.inspect().load), flags)
		return 1
	}

	if (flags.has('json')) {
		console.log(jsonOut(info))
		return 0
	}

	console.log(heading(`Formula: ${info.ref}`))
	console.log(bullet('Formula', `=${info.formula}`))
	console.log(bullet('Normalized', `=${info.normalizedFormula}`))
	console.log(bullet('Volatile', info.volatile ? 'yes' : 'no'))
	console.log(bullet('Functions', info.functions.length > 0 ? info.functions.join(', ') : '(none)'))
	console.log(
		bullet(
			'References',
			info.references.length > 0
				? info.references.map(formatReferenceSummary).join(', ')
				: info.refs.length > 0
					? info.refs.join(', ')
					: '(none)',
		),
	)
	if (info.parseError) {
		console.log(bullet('Parse error', info.parseError))
	}
	return 0
}

function formulaNotFoundError(cellRef: string, load: WorkbookLoadInfo) {
	return ascendError('VALIDATION_ERROR', 'Formula not found', {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			command: 'formula show',
			cell: cellRef,
			load,
			workflow: ['inspect', 'read', 'formula-assist'],
		},
		suggestedFix:
			'Run ascend inspect or ascend read to confirm the target sheet and cell before retrying formula show.',
	})
}

async function assistFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const expr = args[0]
	if (!expr) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required formula assist input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'formula assist',
					required: ['formula'],
					missing: ['formula'],
					workflow: ['inspect', 'formula-assist', 'plan'],
				},
				suggestedFix: "Run ascend formula assist '<formula>' --json before planning formula edits.",
			}),
			flags,
		)
		return 1
	}

	const cursor = optionalIntegerFlag(flags, 'cursor')
	const completionLimit = optionalIntegerFlag(flags, 'completion-limit')
	if (cursor === null || completionLimit === null) return 1

	const result = formulaAssist(expr, {
		...(cursor !== undefined ? { cursor } : {}),
		...(flags.has('prefix') ? { prefix: flags.get('prefix') ?? '' } : {}),
		...(completionLimit !== undefined ? { completionLimit } : {}),
		...(flags.has('function-name') ? { functionName: flags.get('function-name') ?? '' } : {}),
		...(flags.has('reference') ? { reference: flags.get('reference') ?? '' } : {}),
		replaceReferenceAtCursor: flags.has('replace-reference-at-cursor'),
		cycleReference: flags.has('cycle-reference'),
	})

	if (flags.has('json')) {
		console.log(jsonOut(result))
		return 0
	}

	console.log(heading('Formula Assist'))
	console.log(bullet('Formula', result.formula))
	console.log(bullet('Parse', result.diagnostics.parseOk ? 'ok' : 'error'))
	if (result.diagnostics.diagnostics.length > 0) {
		for (const diagnostic of result.diagnostics.diagnostics) {
			console.log(bullet('Diagnostic', `${diagnostic.code}: ${diagnostic.message}`))
		}
	}
	if (result.completions.length > 0) {
		console.log(
			bullet('Completions', result.completions.map((completion) => completion.name).join(', ')),
		)
	}
	if (result.signature) console.log(bullet('Signature', result.signature.label))
	if (result.signatureHelp) {
		console.log(
			bullet(
				'Active signature',
				`${result.signatureHelp.signature.label} arg ${result.signatureHelp.activeParameter + 1}`,
			),
		)
	}
	if (result.activeReference) {
		console.log(bullet('Active reference', result.activeReference.text))
	}
	if (result.insertion) console.log(bullet('Insertion', result.insertion.formula))
	if (result.cycle) console.log(bullet('Cycle', result.cycle.formula))
	return 0
}

async function setFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	const expr = args[2]
	if (!file || !cellRef || !expr) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required formula set input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'formula set',
					required: ['file', 'cell', 'formula'],
					missing: [
						...(!file ? ['file'] : []),
						...(!cellRef ? ['cell'] : []),
						...(!expr ? ['formula'] : []),
					],
					workflow: ['formula-assist', 'plan', 'commit', 'verify'],
				},
				suggestedFix:
					"Prefer ascend plan/commit for auditable formula edits. For direct edits, run ascend formula set <file> <sheet!cell> '<formula>' --json.",
			}),
			flags,
		)
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const result = wb.setFormula(cellRef, expr)
	if (result.errors.length > 0) {
		if (flags.has('json')) {
			const first = result.errors[0]
			console.log(
				jsonErr(
					first ??
						ascendError('VALIDATION_ERROR', 'Failed to set formula', {
							details: { apply: result },
						}),
				),
			)
		} else {
			for (const error of result.errors) cliError(error.message, flags)
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
		console.log(`Set formula at ${cellRef}`)
	}
	return 0
}

async function fillFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const rangeRef = args[1]
	const expr = args[2]
	if (!file || !rangeRef || !expr) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required formula fill input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'formula fill',
					required: ['file', 'range', 'formula'],
					missing: [
						...(!file ? ['file'] : []),
						...(!rangeRef ? ['range'] : []),
						...(!expr ? ['formula'] : []),
					],
					workflow: ['formula-assist', 'plan', 'commit', 'verify'],
				},
				suggestedFix:
					"Prefer ascend plan/commit for auditable formula edits. For direct edits, run ascend formula fill <file> <sheet!range> '<formula>' --json.",
			}),
			flags,
		)
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const result = wb.fillFormula(rangeRef, expr)
	if (result.errors.length > 0) {
		if (flags.has('json')) {
			const first = result.errors[0]
			console.log(
				jsonErr(
					first ??
						ascendError('VALIDATION_ERROR', 'Failed to fill formula', {
							details: { apply: result },
						}),
				),
			)
		} else {
			for (const error of result.errors) cliError(error.message, flags)
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
		console.log(`Filled formula across ${rangeRef}`)
	}
	return 0
}

function optionalIntegerFlag(flags: Map<string, string>, name: string): number | undefined | null {
	if (!flags.has(name)) return undefined
	const value = Number(flags.get(name))
	if (!Number.isInteger(value) || value < 0) {
		cliError(`Invalid --${name}: expected a non-negative integer`, flags)
		return null
	}
	return value
}

function suggestClosest(input: string, candidates: readonly string[]): string | undefined {
	let best: { candidate: string; distance: number } | undefined
	for (const candidate of candidates) {
		const distance = levenshtein(input, candidate)
		if (!best || distance < best.distance) best = { candidate, distance }
	}
	if (!best) return undefined
	return best.distance <= Math.max(2, Math.floor(best.candidate.length / 3))
		? best.candidate
		: undefined
}

function formatReferenceSummary(reference: {
	readonly kind: string
	readonly text: string
	readonly members?: readonly { readonly text: string }[]
}): string {
	if (reference.kind === 'union' || reference.kind === 'intersection') {
		const members = reference.members?.map((member) => member.text).join(', ') ?? reference.text
		return `${reference.kind}(${members})`
	}
	return reference.text
}
