import { errorValue, numberValue } from '@ascend/schema'
import type { FunctionDef } from '../index.ts'
import { fn, numArg, seededRandom } from './helpers.ts'

export const randomFunctions: FunctionDef[] = [
	fn('RAND', 0, 0, (_args, ctx) => numberValue(seededRandom(ctx)), true),

	fn(
		'RANDBETWEEN',
		2,
		2,
		(args, ctx) => {
			const lo = numArg(args[0])
			if (typeof lo !== 'number') return lo
			const hi = numArg(args[1])
			if (typeof hi !== 'number') return hi
			const bottom = Math.ceil(lo)
			const top = Math.floor(hi)
			if (bottom > top) return errorValue('#NUM!')
			return numberValue(Math.floor(seededRandom(ctx) * (top - bottom + 1)) + bottom)
		},
		true,
	),
]
