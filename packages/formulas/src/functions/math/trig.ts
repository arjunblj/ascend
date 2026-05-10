import { errorValue, numberValue } from '@ascend/schema'
import type { FunctionDef } from '../index.ts'
import { fn, numArg } from './helpers.ts'

const PI_OVER_4_HI = 0.7853981633974483
const PI_OVER_4_LO = 3.061616997868383e-17
const TAN_REDUCTION_LIMIT = 1_048_576

function reducedSinCos(value: number): readonly [number, number] {
	const squared = value * value
	let sin = value
	let sinTerm = value
	let cos = 1
	let cosTerm = 1
	for (let index = 1; index < 24; index++) {
		const twice = 2 * index
		sinTerm *= -squared / (twice * (twice + 1))
		cosTerm *= -squared / ((twice - 1) * twice)
		sin += sinTerm
		cos += cosTerm
	}
	return [sin, cos]
}

export function tanExcel(n: number): number {
	if (!Number.isFinite(n) || Math.abs(n) > TAN_REDUCTION_LIMIT) return Math.tan(n)
	const quadrant = Math.round(n / PI_OVER_4_HI)
	const reduced = n - quadrant * PI_OVER_4_HI - quadrant * PI_OVER_4_LO
	const [sin, cos] = reducedSinCos(reduced)
	const tangent = sin / cos
	switch (((quadrant % 4) + 4) % 4) {
		case 0:
			return tangent
		case 1:
			return (1 + tangent) / (1 - tangent)
		case 2:
			return -1 / tangent
		default:
			return (tangent - 1) / (1 + tangent)
	}
}

function tanhExcel(n: number): number {
	if (!Number.isFinite(n)) return Math.tanh(n)
	const magnitude = Math.abs(n)
	if (magnitude === 0) return n
	if (magnitude < 1e-8) return n
	if (magnitude > 20) return Math.sign(n)
	const positive = Math.exp(magnitude)
	const negative = Math.exp(-magnitude)
	const value = (positive - negative) / (positive + negative)
	return n < 0 ? -value : value
}

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
		return numberValue(tanExcel(n))
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
		return numberValue(tanhExcel(n))
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

	fn('COT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) >= 2 ** 27) return errorValue('#NUM!')
		const s = Math.sin(n)
		if (s === 0) return errorValue('#DIV/0!')
		return numberValue(Math.cos(n) / s)
	}),

	fn('COTH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) >= 2 ** 27) return errorValue('#NUM!')
		const s = Math.sinh(n)
		if (s === 0) return errorValue('#DIV/0!')
		return numberValue(Math.cosh(n) / s)
	}),

	fn('CSC', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) >= 2 ** 27) return errorValue('#NUM!')
		const s = Math.sin(n)
		if (s === 0) return errorValue('#DIV/0!')
		return numberValue(1 / s)
	}),

	fn('CSCH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) >= 2 ** 27) return errorValue('#NUM!')
		const s = Math.sinh(n)
		if (s === 0) return errorValue('#DIV/0!')
		return numberValue(1 / s)
	}),

	fn('SEC', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) >= 2 ** 27) return errorValue('#NUM!')
		const c = Math.cos(n)
		if (c === 0) return errorValue('#DIV/0!')
		return numberValue(1 / c)
	}),

	fn('SECH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) >= 2 ** 27) return errorValue('#NUM!')
		return numberValue(1 / Math.cosh(n))
	}),

	fn('ACOT', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		return numberValue(Math.PI / 2 - Math.atan(n))
	}),

	fn('ACOTH', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (Math.abs(n) <= 1) return errorValue('#NUM!')
		return numberValue(0.5 * Math.log((n + 1) / (n - 1)))
	}),
]
