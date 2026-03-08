export { dateToSerial, serialToDate } from './date.ts'
export {
	cellOf,
	collectNumbers,
	compareValues,
	type EvalArg,
	type FnArg,
	type FunctionDef,
	type FunctionEvalContext,
	flattenArgs,
	functionRegistry,
	getRange,
	numArg,
	registerFunction,
	toNumber,
	valuesEqual,
	wildcardMatch,
} from './registry.ts'

import { logicalFunctions } from './logical.ts'
import { mathFunctions } from './math.ts'
import { registerFunction } from './registry.ts'
import { textFunctions } from './text.ts'

for (const fn of [...mathFunctions, ...textFunctions, ...logicalFunctions]) {
	registerFunction(fn)
}

import './date.ts'
import './info.ts'
import './lookup.ts'
import './stats.ts'
