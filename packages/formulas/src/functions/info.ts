import type { CellValue, ScalarCellValue } from '@ascend/schema'
import { arrayValue, booleanValue, errorValue, numberValue, topLeftScalar } from '@ascend/schema'
import type { FunctionDef } from './registry.ts'
import { cellOf, type EvalArg, getRange, numArg } from './registry.ts'

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
	return mapInfoPredicate(args[0], (v) => v.kind === 'number' || v.kind === 'date')
}

function mapInfoPredicate(
	arg: EvalArg | undefined,
	predicate: (value: CellValue) => boolean,
): CellValue {
	if (arg?.kind === 'range' || arg?.value.kind === 'array') {
		const rows = getRange(arg).map((row) =>
			row.map(
				(value): ScalarCellValue => topLeftScalar(booleanValue(predicate(topLeftScalar(value)))),
			),
		)
		return arrayValue(rows)
	}
	return booleanValue(predicate(cellOf(arg)))
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

function nScalar(v: CellValue): ScalarCellValue {
	if (v.kind === 'error') return v
	if (v.kind === 'number') return v
	if (v.kind === 'date') return topLeftScalar(numberValue(v.serial))
	if (v.kind === 'boolean') return topLeftScalar(numberValue(v.value ? 1 : 0))
	return topLeftScalar(numberValue(0))
}

function nFn(args: EvalArg[]): CellValue {
	const arg = args[0]
	if (arg?.kind === 'range' || arg?.value.kind === 'array') {
		return arrayValue(getRange(arg).map((row) => row.map((value) => nScalar(topLeftScalar(value)))))
	}
	return nScalar(cellOf(arg))
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

export const infoFunctions: FunctionDef[] = [
	{ name: 'ISBLANK', minArgs: 1, maxArgs: 1, evaluate: isblank },
	{ name: 'ISERROR', minArgs: 1, maxArgs: 1, evaluate: iserror },
	{ name: 'ISERR', minArgs: 1, maxArgs: 1, evaluate: iserr },
	{ name: 'ISNA', minArgs: 1, maxArgs: 1, evaluate: isna },
	{ name: 'ISNUMBER', minArgs: 1, maxArgs: 1, evaluate: isnumber },
	{ name: 'ISTEXT', minArgs: 1, maxArgs: 1, evaluate: istext },
	{ name: 'ISLOGICAL', minArgs: 1, maxArgs: 1, evaluate: islogical },
	{
		name: 'ISREF',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => booleanValue(args[0]?.ref !== undefined),
	},
	{ name: 'ISFORMULA', minArgs: 1, maxArgs: 1, evaluate: () => booleanValue(false) },
	{ name: 'TYPE', minArgs: 1, maxArgs: 1, evaluate: typeFn },
	{ name: 'N', minArgs: 1, maxArgs: 1, evaluate: nFn },
	{ name: 'NA', minArgs: 0, maxArgs: 0, evaluate: na },
	{ name: 'ISEVEN', minArgs: 1, maxArgs: 1, evaluate: iseven },
	{ name: 'ISODD', minArgs: 1, maxArgs: 1, evaluate: isodd },
	{ name: 'ISNONTEXT', minArgs: 1, maxArgs: 1, evaluate: isnontext },
	{ name: 'ERROR.TYPE', minArgs: 1, maxArgs: 1, evaluate: errorType },
]
