import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'

export async function formulaCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const action = args[0]
	switch (action) {
		case 'show':
			return showFormula(args.slice(1), flags)
		case 'set':
			return setFormula(args.slice(1), flags)
		case 'fill':
			return fillFormula(args.slice(1), flags)
		default:
			console.error('Usage: ascend formula <show|set|fill> ...')
			return 1
	}
}

async function showFormula(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const cellRef = args[1]
	if (!file || !cellRef) {
		console.error('Usage: ascend formula show <file> <sheet!cell>')
		return 1
	}

	const wb = await AscendWorkbook.open(file)
	const info = wb.formula(cellRef)
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
	console.log(bullet('Refs', info.refs.length > 0 ? info.refs.join(', ') : '(none)'))
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

	const wb = await AscendWorkbook.open(file)
	const result = wb.setFormula(cellRef, expr)
	if (result.errors.length > 0) {
		for (const error of result.errors) console.error(error.message)
		return 1
	}
	if (result.recalcRequired) wb.recalc()
	await wb.save(file)

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

	const wb = await AscendWorkbook.open(file)
	const result = wb.fillFormula(rangeRef, expr)
	if (result.errors.length > 0) {
		for (const error of result.errors) console.error(error.message)
		return 1
	}
	if (result.recalcRequired) wb.recalc()
	await wb.save(file)

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		console.log(`Filled formula across ${rangeRef}`)
	}
	return 0
}
