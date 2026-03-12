import { EMPTY, errorValue, numberValue, stringValue, topLeftScalar } from '@ascend/schema'
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

	fn('SUMSQ', 1, 255, (args) => {
		let sum = 0
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (cell.kind === 'error') return cell
						if (cell.kind === 'number') sum += cell.value ** 2
						else if (cell.kind === 'date') sum += cell.serial ** 2
					}
				}
			} else {
				const n = numArg(arg)
				if (typeof n !== 'number') return n
				sum += n * n
			}
		}
		return numberValue(sum)
	}),

	fn('ROMAN', 1, 2, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const num = Math.trunc(n)
		if (num < 0 || num > 3999) return errorValue('#VALUE!')
		if (num === 0) return stringValue('')
		const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
		const syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
		let result = ''
		let remaining = num
		for (let i = 0; i < vals.length; i++) {
			while (remaining >= (vals[i] as number)) {
				result += syms[i]
				remaining -= vals[i] as number
			}
		}
		return stringValue(result)
	}),

	fn('ARABIC', 1, 1, (args) => {
		const v = topLeftScalar(args[0]?.value ?? EMPTY)
		if (v.kind === 'error') return v
		if (v.kind !== 'string') return errorValue('#VALUE!')
		let s = v.value.trim().toUpperCase()
		if (s === '' || s === '-') return numberValue(0)
		let sign = 1
		if (s.startsWith('-')) {
			sign = -1
			s = s.slice(1)
		}
		const romanVals: Record<string, number> = {
			I: 1,
			V: 5,
			X: 10,
			L: 50,
			C: 100,
			D: 500,
			M: 1000,
		}
		let result = 0
		let prev = 0
		for (let i = s.length - 1; i >= 0; i--) {
			const val = romanVals[s[i] as string]
			if (val === undefined) return errorValue('#VALUE!')
			if (val < prev) result -= val
			else result += val
			prev = val
		}
		return numberValue(sign * result)
	}),

	fn('BASE', 2, 3, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const radix = numArg(args[1])
		if (typeof radix !== 'number') return radix
		const r = Math.trunc(radix)
		if (r < 2 || r > 36) return errorValue('#NUM!')
		const num = Math.trunc(n)
		if (num < 0) return errorValue('#NUM!')
		let result = num.toString(r).toUpperCase()
		if (args.length >= 3) {
			const minLen = numArg(args[2])
			if (typeof minLen !== 'number') return minLen
			const ml = Math.trunc(minLen)
			if (ml < 0) return errorValue('#NUM!')
			result = result.padStart(ml, '0')
		}
		return stringValue(result)
	}),

	fn('DECIMAL', 2, 2, (args) => {
		const v = topLeftScalar(args[0]?.value ?? EMPTY)
		if (v.kind === 'error') return v
		const s = v.kind === 'string' ? v.value : v.kind === 'number' ? String(v.value) : ''
		const radix = numArg(args[1])
		if (typeof radix !== 'number') return radix
		const r = Math.trunc(radix)
		if (r < 2 || r > 36) return errorValue('#NUM!')
		const result = Number.parseInt(s, r)
		if (Number.isNaN(result)) return errorValue('#NUM!')
		return numberValue(result)
	}),
]
