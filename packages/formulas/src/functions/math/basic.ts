import { errorValue, numberValue } from '@ascend/schema'
import type { FunctionDef } from '../index.ts'
import { fn, numArg } from './helpers.ts'

export const basicFunctions: FunctionDef[] = [
	fn('ABS', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.abs(n))
	}),

	fn('MOD', 2, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const d = numArg(args[1])
		if (typeof d !== 'number') return d
		if (d === 0) return errorValue('#DIV/0!')
		return numberValue(n - d * Math.floor(n / d))
	}),

	fn('POWER', 2, 2, (args) => {
		const base = numArg(args[0])
		if (typeof base !== 'number') return base
		const exp = numArg(args[1])
		if (typeof exp !== 'number') return exp
		const result = base ** exp
		if (!Number.isFinite(result)) return errorValue('#NUM!')
		return numberValue(result)
	}),

	fn('SQRT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < 0) return errorValue('#NUM!')
		return numberValue(Math.sqrt(n))
	}),

	fn('SIGN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.sign(n))
	}),

	fn('QUOTIENT', 2, 2, (args) => {
		const num = numArg(args[0])
		if (typeof num !== 'number') return num
		const den = numArg(args[1])
		if (typeof den !== 'number') return den
		if (den === 0) return errorValue('#DIV/0!')
		return numberValue(Math.trunc(num / den))
	}),

	fn('GCD', 1, 255, (args) => {
		const nums: number[] = []
		for (const arg of args) {
			const n = numArg(arg)
			if (typeof n !== 'number') return n
			nums.push(Math.trunc(Math.abs(n)))
		}
		if (nums.length === 0) return numberValue(0)
		let g = nums[0] ?? 0
		for (let i = 1; i < nums.length; i++) {
			let a = g
			let b = nums[i] ?? 0
			while (b !== 0) {
				const t = b
				b = a % b
				a = t
			}
			g = a
		}
		return numberValue(g)
	}),

	fn('LCM', 1, 255, (args) => {
		const nums: number[] = []
		for (const arg of args) {
			const n = numArg(arg)
			if (typeof n !== 'number') return n
			nums.push(Math.trunc(Math.abs(n)))
		}
		if (nums.length === 0) return numberValue(0)
		const gcd = (a: number, b: number): number => (b === 0 ? a : gcd(b, a % b))
		let l = nums[0] ?? 0
		for (let i = 1; i < nums.length; i++) {
			const n = nums[i] ?? 0
			if (l === 0 || n === 0) {
				l = 0
				break
			}
			l = Math.abs(l * n) / gcd(l, n)
		}
		return numberValue(l)
	}),

	fn('LOG', 1, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= 0) return errorValue('#NUM!')
		const base = args.length >= 2 ? numArg(args[1]) : 10
		if (typeof base !== 'number') return base
		if (base <= 0 || base === 1) return errorValue('#NUM!')
		return numberValue(Math.log(n) / Math.log(base))
	}),

	fn('LOG10', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= 0) return errorValue('#NUM!')
		return numberValue(Math.log10(n))
	}),

	fn('LN', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n <= 0) return errorValue('#NUM!')
		return numberValue(Math.log(n))
	}),

	fn('EXP', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.exp(n))
	}),
]
