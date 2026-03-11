import type { CellValue } from '@ascend/schema'
import { booleanValue, errorValue, numberValue } from '@ascend/schema'
import { cellOf, type EvalArg, numArg, registerFunction } from './registry.ts'

function isblank(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind === 'empty')
}

function iserror(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind === 'error')
}

function iserr(args: EvalArg[]): CellValue {
	const v = cellOf(args[0])
	return booleanValue(v.kind === 'error' && v.value !== '#N/A')
}

function isna(args: EvalArg[]): CellValue {
	const v = cellOf(args[0])
	return booleanValue(v.kind === 'error' && v.value === '#N/A')
}

function isnumber(args: EvalArg[]): CellValue {
	const v = cellOf(args[0])
	return booleanValue(v.kind === 'number' || v.kind === 'date')
}

function istext(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind === 'string')
}

function islogical(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind === 'boolean')
}

function typeFn(args: EvalArg[]): CellValue {
	const arg = args[0]
	if (arg?.kind === 'range') return numberValue(64)
	const v = arg?.value
	if (!v) return numberValue(1)
	if (v.kind === 'array') return numberValue(64)
	switch (v.kind) {
		case 'number':
		case 'date':
		case 'empty':
			return numberValue(1)
		case 'string':
		case 'richText':
			return numberValue(2)
		case 'boolean':
			return numberValue(4)
		case 'error':
			return numberValue(16)
		default:
			return numberValue(1)
	}
}

function nFn(args: EvalArg[]): CellValue {
	const v = cellOf(args[0])
	if (v.kind === 'error') return v
	if (v.kind === 'number') return v
	if (v.kind === 'date') return numberValue(v.serial)
	if (v.kind === 'boolean') return numberValue(v.value ? 1 : 0)
	return numberValue(0)
}

function na(_args: EvalArg[]): CellValue {
	return errorValue('#N/A')
}

function iseven(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	return booleanValue(Math.trunc(n) % 2 === 0)
}

function isodd(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	return booleanValue(Math.trunc(n) % 2 !== 0)
}

function isnontext(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind !== 'string')
}

const ERROR_TYPE_MAP: Record<string, number> = {
	'#NULL!': 1,
	'#DIV/0!': 2,
	'#VALUE!': 3,
	'#REF!': 4,
	'#NAME?': 5,
	'#NUM!': 6,
	'#N/A': 7,
	'#GETTING_DATA': 8,
	'#SPILL!': 9,
	'#CALC!': 10,
}

function errorType(args: EvalArg[]): CellValue {
	const v = cellOf(args[0])
	if (v.kind !== 'error') return errorValue('#N/A')
	const code = ERROR_TYPE_MAP[v.value]
	return code !== undefined ? numberValue(code) : errorValue('#N/A')
}

// --- Registration ---

registerFunction({ name: 'ISBLANK', minArgs: 1, maxArgs: 1, evaluate: isblank })
registerFunction({
	name: 'ISERROR',
	minArgs: 1,
	maxArgs: 1,
	evaluate: iserror,
})
registerFunction({ name: 'ISERR', minArgs: 1, maxArgs: 1, evaluate: iserr })
registerFunction({ name: 'ISNA', minArgs: 1, maxArgs: 1, evaluate: isna })
registerFunction({
	name: 'ISNUMBER',
	minArgs: 1,
	maxArgs: 1,
	evaluate: isnumber,
})
registerFunction({ name: 'ISTEXT', minArgs: 1, maxArgs: 1, evaluate: istext })
registerFunction({
	name: 'ISLOGICAL',
	minArgs: 1,
	maxArgs: 1,
	evaluate: islogical,
})
registerFunction({
	name: 'ISREF',
	minArgs: 1,
	maxArgs: 1,
	evaluate: (args) => booleanValue(args[0]?.ref !== undefined),
})
registerFunction({ name: 'TYPE', minArgs: 1, maxArgs: 1, evaluate: typeFn })
registerFunction({ name: 'N', minArgs: 1, maxArgs: 1, evaluate: nFn })
registerFunction({ name: 'NA', minArgs: 0, maxArgs: 0, evaluate: na })
registerFunction({ name: 'ISEVEN', minArgs: 1, maxArgs: 1, evaluate: iseven })
registerFunction({ name: 'ISODD', minArgs: 1, maxArgs: 1, evaluate: isodd })
registerFunction({ name: 'ISNONTEXT', minArgs: 1, maxArgs: 1, evaluate: isnontext })
registerFunction({ name: 'ERROR.TYPE', minArgs: 1, maxArgs: 1, evaluate: errorType })
