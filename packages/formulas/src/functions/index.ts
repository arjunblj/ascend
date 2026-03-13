export { dateToSerial, serialToDate } from './date.ts'
export {
	type AggregateRangeCache,
	cellOf,
	collectNumbers,
	compareValues,
	type EvalArea,
	type EvalArg,
	type ExactLookupCache,
	type ExactLookupHit,
	type FnArg,
	type FunctionDef,
	type FunctionEvalContext,
	flattenArgs,
	functionRegistry,
	getRange,
	iterAreaRows,
	type LookupVectorCache,
	numArg,
	rangeShape,
	registerFunction,
	toNumber,
	valuesEqual,
	wildcardMatch,
} from './registry.ts'

import { databaseFunctions } from './database.ts'
import { dateFunctions } from './date.ts'
import { dynamicFunctions } from './dynamic.ts'
import { engineeringFunctions } from './engineering.ts'
import { financialFunctions } from './financial.ts'
import { infoFunctions } from './info.ts'
import { logicalFunctions } from './logical.ts'
import { lookupFunctions } from './lookup.ts'

export { clearCriteriaMatchCache } from './math/index.ts'

import { convertFunction } from './convert.ts'
import { mathFunctions } from './math/index.ts'
import { registerFunction } from './registry.ts'
import { statsFunctions } from './stats.ts'
import { textFunctions } from './text.ts'

for (const fn of [
	...mathFunctions,
	...textFunctions,
	...logicalFunctions,
	...financialFunctions,
	...dynamicFunctions,
	...dateFunctions,
	...databaseFunctions,
	...engineeringFunctions,
	...infoFunctions,
	...lookupFunctions,
	...statsFunctions,
	convertFunction,
]) {
	registerFunction(fn)
}
