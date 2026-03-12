import { errorValue, numberValue } from '@ascend/schema'
import type { FunctionDef } from '../index.ts'
import { fn, numArg } from './helpers.ts'

export const roundingFunctions: FunctionDef[] = [
	fn('ROUND', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue((Math.sign(n) * Math.round(Math.abs(n) * factor)) / factor)
	}),

	fn('ROUNDUP', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue((Math.sign(n) * Math.ceil(Math.abs(n) * factor)) / factor)
	}),

	fn('ROUNDDOWN', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue((Math.sign(n) * Math.floor(Math.abs(n) * factor)) / factor)
	}),

	fn('INT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.floor(n))
	}),

	fn('TRUNC', 1, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = args.length >= 2 ? numArg(args[1]) : 0
		if (typeof d !== 'number') return d
		const factor = 10 ** Math.trunc(d)
		return numberValue(Math.trunc(n * factor) / factor)
	}),

	fn('MROUND', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const m = numArg(args[1])
		if (typeof m !== 'number') return m
		if (n > 0 && m < 0) return errorValue('#NUM!')
		if (n < 0 && m > 0) return errorValue('#NUM!')
		if (m === 0) return numberValue(0)
		return numberValue(Math.round(n / m) * m)
	}),

	fn('CEILING', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const sig = numArg(args[1])
		if (typeof sig !== 'number') return sig
		if (sig === 0) return numberValue(0)
		if (n > 0 && sig < 0) return errorValue('#NUM!')
		return numberValue(Math.ceil(n / sig) * sig)
	}),

	fn('FLOOR', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const sig = numArg(args[1])
		if (typeof sig !== 'number') return sig
		if (sig === 0) return errorValue('#DIV/0!')
		if (n > 0 && sig < 0) return errorValue('#NUM!')
		return numberValue(Math.floor(n / sig) * sig)
	}),

	fn('ODD', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const x = Math.ceil(Math.abs(n))
		const odd = x % 2 === 0 ? x + 1 : x
		return numberValue(n >= 0 ? odd : -odd)
	}),

	fn('EVEN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const x = Math.ceil(Math.abs(n))
		const even = x % 2 === 1 ? x + 1 : x
		return numberValue(n >= 0 ? even : -even)
	}),
]
