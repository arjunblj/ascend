import type { FunctionDef } from '../index.ts'
import { aggregationFunctions } from './aggregation.ts'
import { basicFunctions } from './basic.ts'
import { combinatoricsFunctions } from './combinatorics.ts'
import { compositeFunctions } from './composite.ts'
import { conditionalFunctions } from './conditional.ts'
import { randomFunctions } from './random.ts'
import { roundingFunctions } from './rounding.ts'
import { trigFunctions } from './trig.ts'

export { clearCriteriaMatchCache } from './conditional.ts'
export { tanExcel } from './trig.ts'

export const mathFunctions: FunctionDef[] = [
	...aggregationFunctions,
	...roundingFunctions,
	...trigFunctions,
	...basicFunctions,
	...combinatoricsFunctions,
	...randomFunctions,
	...conditionalFunctions,
	...compositeFunctions,
]
