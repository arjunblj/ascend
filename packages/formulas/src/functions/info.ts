import type { CellValue } from '@ascend/schema'
import { booleanValue, errorValue, numberValue } from '@ascend/schema'
import { cellOf, type EvalArg, registerFunction } from './registry.ts'

function isblank(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind === 'empty')
}

function iserror(args: EvalArg[]): CellValue {
	return booleanValue(cellOf(args[0]).kind === 'error')
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

// --- Registration ---

registerFunction({ name: 'ISBLANK', minArgs: 1, maxArgs: 1, evaluate: isblank })
registerFunction({
	name: 'ISERROR',
	minArgs: 1,
	maxArgs: 1,
	evaluate: iserror,
})
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
