import { ascendError, levenshtein } from '@ascend/schema'
import { cliError, jsonErr, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import {
	openWorkbookDocumentWithProgress,
	openWorkbookWithProgress,
	withProgress,
} from '../progress.ts'

export const usage = `Usage: ascend formula <subcommand> <file> <ref> [expr] [flags]

  Inspect and edit cell formulas.

Subcommands:
  show <file> <ref>           Show formula at a cell reference
  set <file> <ref> <expr>     Set a formula on a cell
  fill <file> <range> <expr>  Fill a formula across a range

Flags:
  --json          Output as JSON
`

export async function formulaCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const action = args[0]
	switch (action) {
		case 'show':
			return showFormula(args.slice(1), flags)
		case 'set':
			return setFormula(args.slice(1), flags)
		case 'fill':
			return fillFormula(args.slice(1), flags)
		default: {
			const msg = `Unknown formula subcommand: ${action ?? '(missing)'}`
			const suggestion = suggestClosest(action ?? '', ['show', 'set', 'fill'])
			const hint = suggestion ? `\nDid you mean "${suggestion}"?` : ''
			cliError(`${msg}${hint}\nUsage: ascend formula <show|set|fill> ...`, flags)
			return 1
		}
	}
}

async function showFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	if (!file || !cellRef) {
		cliError('Usage: ascend formula show <file> <sheet!cell>', flags)
		return 1
	}

	const { document: session } = await openWorkbookDocumentWithProgress(file, { mode: 'formula' })
	const info = session.formula(cellRef)
	if (!info) {
		cliError(`No formula found at "${cellRef}"`, flags)
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

async function setFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	const expr = args[2]
	if (!file || !cellRef || !expr) {
		cliError("Usage: ascend formula set <file> <sheet!cell> '<formula>'", flags)
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
		cliError("Usage: ascend formula fill <file> <sheet!range> '<formula>'", flags)
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
