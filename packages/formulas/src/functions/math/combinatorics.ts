import { errorValue, numberValue } from '@ascend/schema'
import type { FunctionDef } from '../index.ts'
import { fn, numArg } from './helpers.ts'

export const combinatoricsFunctions: FunctionDef[] = [
	fn('FACT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = Math.trunc(n)
		if (k < 0) return errorValue('#NUM!')
		let f = 1
		for (let i = 2; i <= k; i++) f *= i
		return numberValue(f)
	}),

	fn('FACTDOUBLE', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = Math.trunc(n)
		if (k < 0) return errorValue('#NUM!')
		let f = 1
		for (let i = k; i > 0; i -= 2) f *= i
		return numberValue(f)
	}),

	fn('COMBIN', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = numArg(args[1])
		if (typeof k !== 'number') return k
		const ni = Math.trunc(n)
		const ki = Math.trunc(k)
		if (ni < 0 || ki < 0 || ki > ni) return errorValue('#NUM!')
		if (ki === 0 || ki === ni) return numberValue(1)
		let c = 1
		for (let i = 0; i < ki; i++) c = (c * (ni - i)) / (i + 1)
		return numberValue(Math.round(c))
	}),

	fn('PERMUT', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = numArg(args[1])
		if (typeof k !== 'number') return k
		const ni = Math.trunc(n)
		const ki = Math.trunc(k)
		if (ni < 0 || ki < 0 || ki > ni) return errorValue('#NUM!')
		let p = 1
		for (let i = 0; i < ki; i++) p *= ni - i
		return numberValue(p)
	}),

	fn('COMBINA', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = numArg(args[1])
		if (typeof k !== 'number') return k
		const ni = Math.trunc(n)
		const ki = Math.trunc(k)
		if (ni < 0 || ki < 0) return errorValue('#NUM!')
		if (ni === 0 && ki === 0) return numberValue(1)
		if (ni === 0) return errorValue('#NUM!')
		const total = ni + ki - 1
		let c = 1
		for (let i = 0; i < ki; i++) c = (c * (total - i)) / (i + 1)
		return numberValue(Math.round(c))
	}),

	fn('PERMUTATIONA', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const k = numArg(args[1])
		if (typeof k !== 'number') return k
		const ni = Math.trunc(n)
		const ki = Math.trunc(k)
		if (ni < 0 || ki < 0) return errorValue('#NUM!')
		return numberValue(ni ** ki)
	}),
]
