import { jsonOut } from '../output/json.ts'
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
			console.error(`Unknown formula subcommand: ${action ?? '(missing)'}`)
			const suggestion = suggestClosest(action ?? '', ['show', 'set', 'fill'])
			if (suggestion) console.error(`Did you mean "${suggestion}"?`)
			console.error('Usage: ascend formula <show|set|fill> ...')
			return 1
		}
	}
}

async function showFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	if (!file || !cellRef) {
		console.error('Usage: ascend formula show <file> <sheet!cell>')
		return 1
	}

	const { document: session } = await openWorkbookDocumentWithProgress(file, { mode: 'formula' })
	const info = session.formula(cellRef)
	if (!info) {
		console.error(`No formula found at "${cellRef}"`)
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
		console.error("Usage: ascend formula set <file> <sheet!cell> '<formula>'")
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const result = wb.setFormula(cellRef, expr)
	if (result.errors.length > 0) {
		for (const error of result.errors) console.error(error.message)
		return 1
	}
	if (result.recalcRequired) {
		const { value: recalc } = await withProgress('Recalculating formulas', () => wb.recalc())
		if (recalc.errors.length > 0) {
			for (const error of recalc.errors) console.error(`${error.ref}: ${error.error.message}`)
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
		console.error("Usage: ascend formula fill <file> <sheet!range> '<formula>'")
		return 1
	}

	const { workbook: wb } = await openWorkbookWithProgress(file)
	const result = wb.fillFormula(rangeRef, expr)
	if (result.errors.length > 0) {
		for (const error of result.errors) console.error(error.message)
		return 1
	}
	if (result.recalcRequired) {
		const { value: recalc } = await withProgress('Recalculating formulas', () => wb.recalc())
		if (recalc.errors.length > 0) {
			for (const error of recalc.errors) console.error(`${error.ref}: ${error.error.message}`)
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

function levenshtein(a: string, b: string): number {
	if (a === b) return 0
	if (a.length === 0) return b.length
	if (b.length === 0) return a.length
	const prev = Array.from({ length: b.length + 1 }, (_, i) => i)
	const curr = new Array<number>(b.length + 1).fill(0)
	for (let i = 0; i < a.length; i++) {
		curr[0] = i + 1
		for (let j = 0; j < b.length; j++) {
			const left = curr[j] ?? 0
			const up = prev[j + 1] ?? 0
			const diag = prev[j] ?? 0
			const cost = (a[i] ?? '') === (b[j] ?? '') ? 0 : 1
			curr[j + 1] = Math.min(left + 1, up + 1, diag + cost)
		}
		for (let j = 0; j < prev.length; j++) prev[j] = curr[j] ?? 0
	}
	return prev[b.length] ?? Math.max(a.length, b.length)
}
