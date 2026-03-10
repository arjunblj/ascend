export { dateToSerial, serialToDate } from './date.ts'
export {
	cellOf,
	collectNumbers,
	compareValues,
	type EvalArea,
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

import { dynamicFunctions } from './dynamic.ts'
import { financialFunctions } from './financial.ts'
import { logicalFunctions } from './logical.ts'
import { mathFunctions } from './math.ts'
import { registerFunction } from './registry.ts'
import { textFunctions } from './text.ts'

for (const fn of [
	...mathFunctions,
	...textFunctions,
	...logicalFunctions,
	...financialFunctions,
	...dynamicFunctions,
]) {
	registerFunction(fn)
}

import './date.ts'
import './info.ts'
import './lookup.ts'
import './stats.ts'
