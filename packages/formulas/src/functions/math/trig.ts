import { errorValue, numberValue } from '@ascend/schema'
import type { FunctionDef } from '../index.ts'
import { fn, numArg } from './helpers.ts'

export const trigFunctions: FunctionDef[] = [
	fn('SIN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.sin(n))
	}),

	fn('COS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.cos(n))
	}),

	fn('TAN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.tan(n))
	}),

	fn('ASIN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < -1 || n > 1) return errorValue('#NUM!')
		return numberValue(Math.asin(n))
	}),

	fn('ACOS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < -1 || n > 1) return errorValue('#NUM!')
		return numberValue(Math.acos(n))
	}),

	fn('ATAN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.atan(n))
	}),

	fn('ATAN2', 2, 2, (args) => {
		const x = numArg(args[0])
		if (typeof x !== 'number') return x
		const y = numArg(args[1])
		if (typeof y !== 'number') return y
		if (x === 0 && y === 0) return errorValue('#DIV/0!')
		return numberValue(Math.atan2(y, x))
	}),

	fn('DEGREES', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue((n * 180) / Math.PI)
	}),

	fn('RADIANS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue((n * Math.PI) / 180)
	}),

	fn('PI', 0, 0, () => numberValue(Math.PI)),

	fn('SINH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.sinh(n))
	}),

	fn('COSH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.cosh(n))
	}),

	fn('TANH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.tanh(n))
	}),

	fn('ASINH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.asinh(n))
	}),

	fn('ACOSH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < 1) return errorValue('#NUM!')
		return numberValue(Math.acosh(n))
	}),

	fn('ATANH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= -1 || n >= 1) return errorValue('#NUM!')
		return numberValue(Math.atanh(n))
	}),
]
