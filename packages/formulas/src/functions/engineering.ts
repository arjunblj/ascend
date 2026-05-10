import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue, stringValue } from '@ascend/schema'
import type { FunctionDef } from './registry.ts'
import { cellOf, type EvalArg, flattenArgs, numArg } from './registry.ts'

function strArg(arg: EvalArg | undefined): string | CellValue {
	const v = cellOf(arg)
	if (v.kind === 'error') return v
	if (v.kind === 'string') return v.value
	if (v.kind === 'number') return String(Math.trunc(v.value))
	if (v.kind === 'boolean') return v.value ? '1' : '0'
	if (v.kind === 'empty') return ''
	return errorValue('#VALUE!')
}

const BIN_RE = /^[01]{1,10}$/
function parseBin(s: string): number | null {
	const t = s.trim().toUpperCase()
	if (!BIN_RE.test(t)) return null
	if (t.length > 10) return null
	let n = 0
	for (let i = 0; i < t.length; i++) {
		n = (n << 1) | (t[i] === '1' ? 1 : 0)
	}
	if (t.length === 10 && t[0] === '1') {
		n = n - (1 << 10)
	}
	return n
}

function bin2dec(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = parseBin(s)
	return n === null ? errorValue('#NUM!') : numberValue(n)
}

const DEC2BIN_MIN = -512
const DEC2BIN_MAX = 511
function dec2bin(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < DEC2BIN_MIN || x > DEC2BIN_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + (1 << 10)) & 0x3ff : x
	let s = abs.toString(2)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

const HEX_RE = /^[0-9A-Fa-f]{1,10}$/
function parseHex(s: string): number | null {
	const t = s.trim().toUpperCase()
	if (!HEX_RE.test(t)) return null
	if (t.length > 10) return null
	let n = Number.parseInt(t, 16)
	if (t.length === 10 && (t[0] ?? '') >= '8') {
		n = n - 0x10000000000
	}
	return n
}

function hex2dec(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = parseHex(s)
	return n === null ? errorValue('#NUM!') : numberValue(n)
}

const DEC2HEX_MIN = -549_755_813_888
const DEC2HEX_MAX = 549_755_813_887
function dec2hex(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < DEC2HEX_MIN || x > DEC2HEX_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x10000000000) >>> 0 : x
	let s = abs.toString(16).toUpperCase()
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

const OCT_RE = /^[0-7]{1,10}$/
function parseOct(s: string): number | null {
	const t = s.trim()
	if (!OCT_RE.test(t)) return null
	if (t.length > 10) return null
	let n = Number.parseInt(t, 8)
	if (t.length === 10 && (t[0] ?? '') >= '4') {
		n = n - 0x40000000
	}
	return n
}

function oct2dec(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = parseOct(s)
	return n === null ? errorValue('#NUM!') : numberValue(n)
}

const DEC2OCT_MIN = -536_870_912
const DEC2OCT_MAX = 536_870_911
function dec2oct(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < DEC2OCT_MIN || x > DEC2OCT_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x40000000) >>> 0 : x
	let s = abs.toString(8)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function bin2hex(args: EvalArg[]): CellValue {
	const dec = bin2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const x = Math.trunc(dec.value)
	const abs = x < 0 ? (x + 0x10000000000) >>> 0 : x
	let s = abs.toString(16).toUpperCase()
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function bin2oct(args: EvalArg[]): CellValue {
	const dec = bin2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const x = Math.trunc(dec.value)
	const abs = x < 0 ? (x + 0x40000000) >>> 0 : x
	let s = abs.toString(8)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function hex2bin(args: EvalArg[]): CellValue {
	const dec = hex2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2BIN_MIN || x > DEC2BIN_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + (1 << 10)) & 0x3ff : x
	let s = abs.toString(2)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function hex2oct(args: EvalArg[]): CellValue {
	const dec = hex2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2OCT_MIN || x > DEC2OCT_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x40000000) >>> 0 : x
	let s = abs.toString(8)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function oct2bin(args: EvalArg[]): CellValue {
	const dec = oct2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2BIN_MIN || x > DEC2BIN_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + (1 << 10)) & 0x3ff : x
	let s = abs.toString(2)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function oct2hex(args: EvalArg[]): CellValue {
	const dec = oct2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2HEX_MIN || x > DEC2HEX_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x10000000000) >>> 0 : x
	let s = abs.toString(16).toUpperCase()
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function delta(args: EvalArg[]): CellValue {
	const n1 = numArg(args[0])
	if (typeof n1 !== 'number') return n1
	const n2 = args.length >= 2 ? numArg(args[1]) : 0
	if (typeof n2 !== 'number') return n2
	return numberValue(n1 === n2 ? 1 : 0)
}

function gestep(args: EvalArg[]): CellValue {
	const num = numArg(args[0])
	if (typeof num !== 'number') return num
	const step = args.length >= 2 ? numArg(args[1]) : 0
	if (typeof step !== 'number') return step
	return numberValue(num >= step ? 1 : 0)
}

const MAX_48BIT = 2 ** 48 - 1

function bitIntArg(arg: EvalArg | undefined): number | CellValue {
	const n = numArg(arg)
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < 0 || x > MAX_48BIT) return errorValue('#NUM!')
	return x
}

function bitand(args: EvalArg[]): CellValue {
	const a = bitIntArg(args[0])
	if (typeof a !== 'number') return a
	const b = bitIntArg(args[1])
	if (typeof b !== 'number') return b
	return numberValue(Number(BigInt(a) & BigInt(b)))
}

function bitor(args: EvalArg[]): CellValue {
	const a = bitIntArg(args[0])
	if (typeof a !== 'number') return a
	const b = bitIntArg(args[1])
	if (typeof b !== 'number') return b
	return numberValue(Number(BigInt(a) | BigInt(b)))
}

function bitxor(args: EvalArg[]): CellValue {
	const a = bitIntArg(args[0])
	if (typeof a !== 'number') return a
	const b = bitIntArg(args[1])
	if (typeof b !== 'number') return b
	return numberValue(Number(BigInt(a) ^ BigInt(b)))
}

function bitlshift(args: EvalArg[]): CellValue {
	const n = bitIntArg(args[0])
	if (typeof n !== 'number') return n
	const shift = numArg(args[1])
	if (typeof shift !== 'number') return shift
	const s = Math.trunc(shift)
	if (Math.abs(s) > 53) return errorValue('#NUM!')
	const result = s >= 0 ? Number(BigInt(n) << BigInt(s)) : Number(BigInt(n) >> BigInt(-s))
	if (result < 0 || result > MAX_48BIT) return errorValue('#NUM!')
	return numberValue(result)
}

function bitrshift(args: EvalArg[]): CellValue {
	const n = bitIntArg(args[0])
	if (typeof n !== 'number') return n
	const shift = numArg(args[1])
	if (typeof shift !== 'number') return shift
	const s = Math.trunc(shift)
	if (Math.abs(s) > 53) return errorValue('#NUM!')
	const result = s >= 0 ? Number(BigInt(n) >> BigInt(s)) : Number(BigInt(n) << BigInt(-s))
	if (result < 0 || result > MAX_48BIT) return errorValue('#NUM!')
	return numberValue(result)
}

const ERF_P = 0.3275911
const ERF_A1 = 0.254829592
const ERF_A2 = -0.284496736
const ERF_A3 = 1.421413741
const ERF_A4 = -1.453152027
const ERF_A5 = 1.061405429

function erfCore(x: number): number {
	const t = 1 / (1 + ERF_P * x)
	const poly = t * (ERF_A1 + t * (ERF_A2 + t * (ERF_A3 + t * (ERF_A4 + t * ERF_A5))))
	return 1 - poly * Math.exp(-x * x)
}

function erfFn(x: number): number {
	return x >= 0 ? erfCore(x) : -erfCore(-x)
}

function erfcFn(x: number): number {
	return 1 - erfFn(x)
}

function erf(args: EvalArg[]): CellValue {
	const lower = numArg(args[0])
	if (typeof lower !== 'number') return lower
	if (args.length >= 2 && args[1]) {
		const upper = numArg(args[1])
		if (typeof upper !== 'number') return upper
		return numberValue(erfFn(upper) - erfFn(lower))
	}
	return numberValue(erfFn(lower))
}

function erfPrecise(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	return numberValue(erfFn(x))
}

function erfc(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	return numberValue(erfcFn(x))
}

function erfcPrecise(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	return numberValue(erfcFn(x))
}

const TWO_OVER_PI_APPROX = 0.636619772

function horner(coefficients: readonly number[], value: number): number {
	let result = 0
	for (let index = coefficients.length - 1; index >= 0; index--) {
		result = value * result + (coefficients[index] ?? 0)
	}
	return result
}

function besselIter(x: number, n: number, f0: number, f1: number, sign: 1 | -1): number {
	if (n === 0) return f0
	if (n === 1) return f1
	const tdx = 2 / x
	let previous = f0
	let current = f1
	for (let order = 1; order < n; order++) {
		const next = current * order * tdx + sign * previous
		previous = current
		current = next
	}
	return current
}

function besselJ0(x: number): number {
	if (x === 0) return 1
	const ax = Math.abs(x)
	const y = ax * ax
	if (ax < 8) {
		return (
			horner(
				[
					57_568_490_574, -13_362_590_354, 651_619_640.7, -11_214_424.18, 77_392.33017,
					-184.9052456,
				],
				y,
			) / horner([57_568_490_411, 1_029_532_985, 9_494_680.718, 59_272.64853, 267.8532712, 1], y)
		)
	}
	const z = 8 / ax
	const yAsymptotic = z * z
	const xx = ax - 0.785398164
	const p = horner(
		[1, -0.001098628627, 0.00002734510407, -0.000002073370639, 0.0000002093887211],
		yAsymptotic,
	)
	const q = horner(
		[-0.01562499995, 0.0001430488765, -0.000006911147651, 0.0000007621095161, -0.0000000934935152],
		yAsymptotic,
	)
	return Math.sqrt(TWO_OVER_PI_APPROX / ax) * (Math.cos(xx) * p - z * Math.sin(xx) * q)
}

function besselJ1(x: number): number {
	const ax = Math.abs(x)
	const y = x * x
	let result: number
	if (ax < 8) {
		result =
			(x *
				horner(
					[
						72_362_614_232, -7_895_059_235, 242_396_853.1, -2_972_611.439, 15_704.4826,
						-30.16036606,
					],
					y,
				)) /
			horner([144_725_228_442, 2_300_535_178, 18_583_304.74, 99_447.43394, 376.9991397, 1], y)
	} else {
		const z = 8 / ax
		const yAsymptotic = z * z
		const xx = ax - 2.356194491
		const p = horner(
			[1, 0.00183105, -0.00003516396496, 0.000002457520174, -0.000000240337019],
			yAsymptotic,
		)
		const q = horner(
			[0.04687499995, -0.0002002690873, 0.000008449199096, -0.00000088228987, 0.000000105787412],
			yAsymptotic,
		)
		result = Math.sqrt(TWO_OVER_PI_APPROX / ax) * (Math.cos(xx) * p - z * Math.sin(xx) * q)
		if (x < 0) result = -result
	}
	return result
}

function besselJ(x: number, n: number): number {
	if (!Number.isFinite(x)) return Number.isNaN(x) ? Number.NaN : 0
	if (n === 0) return besselJ0(x)
	if (n === 1) return besselJ1(x)
	if (x < 0) return n % 2 === 0 ? besselJ(-x, n) : -besselJ(-x, n)
	if (x === 0) return 0
	if (x > n) return besselIter(x, n, besselJ0(x), besselJ1(x), -1)

	let result = 0
	let sum = 0
	let bjp = 0
	let bj = 1
	let jsum = false
	const tox = 2 / x
	const m = 2 * Math.floor((n + Math.floor(Math.sqrt(40 * n))) / 2)
	for (let order = m; order > 0; order--) {
		const bjm = order * tox * bj - bjp
		bjp = bj
		bj = bjm
		if (Math.abs(bj) > 1e10) {
			bj *= 1e-10
			bjp *= 1e-10
			result *= 1e-10
			sum *= 1e-10
		}
		if (jsum) sum += bj
		jsum = !jsum
		if (order === n) result = bjp
	}
	return result / (2 * sum - bj)
}

function besselI0(x: number): number {
	const ax = Math.abs(x)
	if (ax <= 3.75) {
		const y = (x / 3.75) ** 2
		return horner([1, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.0360768, 0.0045813], y)
	}
	return (
		(Math.exp(ax) / Math.sqrt(ax)) *
		horner(
			[
				0.39894228, 0.01328592, 0.00225319, -0.00157565, 0.00916281, -0.02057706, 0.02635537,
				-0.01647633, 0.00392377,
			],
			3.75 / ax,
		)
	)
}

function besselI1(x: number): number {
	const ax = Math.abs(x)
	let result: number
	if (ax < 3.75) {
		const y = (x / 3.75) ** 2
		result =
			x * horner([0.5, 0.87890594, 0.51498869, 0.15084934, 0.02658733, 0.00301532, 0.00032411], y)
	} else {
		result =
			(Math.exp(ax) / Math.sqrt(ax)) *
			horner(
				[
					0.39894228, -0.03988024, -0.00362018, 0.00163801, -0.01031555, 0.02282967, -0.02895312,
					0.01787654, -0.00420059,
				],
				3.75 / ax,
			)
		if (x < 0) result = -result
	}
	return result
}

function besselI(x: number, n: number): number {
	if (n === 0) return besselI0(x)
	if (n === 1) return besselI1(x)
	if (x === 0) return 0
	if (x === Infinity) return Infinity

	let result = 0
	let bip = 0
	let bi = 1
	const tox = 2 / Math.abs(x)
	const m = 2 * Math.round((n + Math.round(Math.sqrt(40 * n))) / 2)
	for (let order = m; order > 0; order--) {
		const bim = order * tox * bi + bip
		bip = bi
		bi = bim
		if (Math.abs(bi) > 1e10) {
			bi *= 1e-10
			bip *= 1e-10
			result *= 1e-10
		}
		if (order === n) result = bip
	}
	result *= besselI0(x) / bi
	return x < 0 && n % 2 === 1 ? -result : result
}

function besselY0(x: number): number {
	if (x < 8) {
		const y = x * x
		return (
			horner(
				[-2_957_821_389, 7_062_834_065, -512_359_803.6, 10_879_881.29, -86_327.92757, 228.4622733],
				y,
			) /
				horner([40_076_544_269, 745_249_964.8, 7_189_466.438, 47_447.2647, 226.1030244, 1], y) +
			TWO_OVER_PI_APPROX * besselJ0(x) * Math.log(x)
		)
	}
	const z = 8 / x
	const y = z * z
	const xx = x - 0.785398164
	const p = horner(
		[1, -0.001098628627, 0.00002734510407, -0.000002073370639, 0.0000002093887211],
		y,
	)
	const q = horner(
		[-0.01562499995, 0.0001430488765, -0.000006911147651, 0.0000007621095161, -0.0000000934945152],
		y,
	)
	return Math.sqrt(TWO_OVER_PI_APPROX / x) * (Math.sin(xx) * p + z * Math.cos(xx) * q)
}

function besselY1(x: number): number {
	if (x < 8) {
		const y = x * x
		return (
			(x *
				horner(
					[
						-4_900_604_943_000, 1_275_274_390_000, -51_534_381_390, 734_926_455.1, -4_237_922.726,
						8_511.937935,
					],
					y,
				)) /
				horner(
					[
						24_995_805_700_000, 424_441_966_400, 3_733_650_367, 22_459_040.02, 102_042.605,
						354.9632885, 1,
					],
					y,
				) +
			TWO_OVER_PI_APPROX * (besselJ1(x) * Math.log(x) - 1 / x)
		)
	}
	const z = 8 / x
	const y = z * z
	const xx = x - 2.356194491
	const p = horner([1, 0.00183105, -0.00003516396496, 0.000002457520174, -0.000000240337019], y)
	const q = horner(
		[0.04687499995, -0.0002002690873, 0.000008449199096, -0.00000088228987, 0.000000105787412],
		y,
	)
	return Math.sqrt(TWO_OVER_PI_APPROX / x) * (Math.sin(xx) * p + z * Math.cos(xx) * q)
}

function besselY(x: number, n: number): number {
	if (n === 0) return besselY0(x)
	if (n === 1) return besselY1(x)
	let y0 = besselY0(x)
	let y1 = besselY1(x)
	for (let i = 1; i < n; i++) {
		const y2 = ((2 * i) / x) * y1 - y0
		y0 = y1
		y1 = y2
	}
	return y1
}

function besselK0(x: number): number {
	if (x <= 2) {
		const y = (x * x) / 4
		return (
			-Math.log(x / 2) * besselI0(x) +
			horner([-0.57721566, 0.4227842, 0.23069756, 0.0348859, 0.00262698, 0.0001075, 0.0000074], y)
		)
	}
	return (
		(Math.exp(-x) / Math.sqrt(x)) *
		horner(
			[1.25331414, -0.07832358, 0.02189568, -0.01062446, 0.00587872, -0.0025154, 0.00053208],
			2 / x,
		)
	)
}

function besselK1(x: number): number {
	if (x <= 2) {
		const y = (x * x) / 4
		return (
			Math.log(x / 2) * besselI1(x) +
			(1 / x) *
				horner([1, 0.15443144, -0.67278579, -0.18156897, -0.01919402, -0.00110404, -0.00004686], y)
		)
	}
	return (
		(Math.exp(-x) / Math.sqrt(x)) *
		horner(
			[1.25331414, 0.23498619, -0.0365562, 0.01504268, -0.00780353, 0.00325614, -0.00068245],
			2 / x,
		)
	)
}

function besselK(x: number, n: number): number {
	if (n === 0) return besselK0(x)
	if (n === 1) return besselK1(x)
	let k0 = besselK0(x)
	let k1 = besselK1(x)
	for (let i = 1; i < n; i++) {
		const k2 = ((2 * i) / x) * k1 + k0
		k0 = k1
		k1 = k2
	}
	return k1
}

function besselOrderArg(arg: EvalArg | undefined): number | CellValue {
	const n = numArg(arg)
	if (typeof n !== 'number') return n
	const k = Math.trunc(n)
	if (k < 0) return errorValue('#NUM!')
	return k
}

function besselI_(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const n = besselOrderArg(args[1])
	if (typeof n !== 'number') return n
	if (Number.isNaN(x)) return errorValue('#NUM!')
	return numberValue(besselI(x, n))
}

function besselJ_(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const n = besselOrderArg(args[1])
	if (typeof n !== 'number') return n
	if (Number.isNaN(x)) return errorValue('#NUM!')
	return numberValue(besselJ(x, n))
}

function besselK_(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const n = besselOrderArg(args[1])
	if (typeof n !== 'number') return n
	if (Number.isNaN(x) || x <= 0) return errorValue('#NUM!')
	return numberValue(besselK(x, n))
}

function besselY_(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const n = besselOrderArg(args[1])
	if (typeof n !== 'number') return n
	if (Number.isNaN(x) || x <= 0) return errorValue('#NUM!')
	return numberValue(besselY(x, n))
}

interface Complex {
	re: number
	im: number
	suffix: string
}

function parseComplex(s: string): Complex | null {
	const trimmed = s.trim()
	if (trimmed === '') return null
	const lastChar = trimmed[trimmed.length - 1]
	if (lastChar !== 'i' && lastChar !== 'j') {
		const n = Number(trimmed)
		if (Number.isNaN(n)) return null
		return { re: n, im: 0, suffix: '' }
	}
	const suffix = lastChar
	const body = trimmed.slice(0, -1)
	if (body === '' || body === '+') return { re: 0, im: 1, suffix }
	if (body === '-') return { re: 0, im: -1, suffix }
	let splitIdx = -1
	for (let i = body.length - 1; i > 0; i--) {
		const ch = body[i]
		if ((ch === '+' || ch === '-') && body[i - 1] !== 'e' && body[i - 1] !== 'E') {
			splitIdx = i
			break
		}
	}
	if (splitIdx === -1) {
		const n = Number(body)
		if (Number.isNaN(n)) return null
		return { re: 0, im: n, suffix }
	}
	const realStr = body.slice(0, splitIdx)
	const imStr = body.slice(splitIdx)
	const re = Number(realStr)
	const im = imStr === '+' ? 1 : imStr === '-' ? -1 : Number(imStr)
	if (Number.isNaN(re) || Number.isNaN(im)) return null
	return { re, im, suffix }
}

function cleanFloat(n: number): number {
	if (Math.abs(n) < 5e-15) return 0
	const r = Math.round(n)
	if (r !== 0 && Math.abs(n - r) / Math.abs(r) < 1e-12) return r
	return n
}

function formatComplexNumber(n: number): string {
	if (Number.isInteger(n)) return String(n)
	return n
		.toPrecision(15)
		.replace(/(\.\d*?)0+(E|$)/, '$1$2')
		.replace(/\.E/, 'E')
		.replace(/e/g, 'E')
}

function formatComplex(re: number, im: number, sfx = 'i'): string {
	const suffix = sfx || 'i'
	const cRe = cleanFloat(re)
	const cIm = cleanFloat(im)
	if (cIm === 0 && cRe === 0) return '0'
	if (cIm === 0) return formatComplexNumber(cRe)
	let imStr: string
	if (cIm === 1) imStr = suffix
	else if (cIm === -1) imStr = `-${suffix}`
	else imStr = `${formatComplexNumber(Math.abs(cIm))}${suffix}`
	if (cRe === 0) return cIm < 0 && !imStr.startsWith('-') ? `-${imStr}` : imStr
	if (cIm > 0) return `${formatComplexNumber(cRe)}+${imStr}`
	return `${formatComplexNumber(cRe)}-${imStr}`
}

function complexArg(arg: EvalArg | undefined): Complex | CellValue {
	const v = cellOf(arg)
	if (v.kind === 'error') return v
	if (v.kind === 'number') return { re: v.value, im: 0, suffix: '' }
	if (v.kind === 'string') {
		const c = parseComplex(v.value)
		if (c === null) return errorValue('#VALUE!')
		return c
	}
	if (v.kind === 'empty') return { re: 0, im: 0, suffix: '' }
	return errorValue('#VALUE!')
}

function isComplex(v: Complex | CellValue): v is Complex {
	return 're' in v
}

function resolveSuffix(a: string, b: string): string | null {
	if (a && b && a !== b) return null
	return a || b || 'i'
}

function complexFn(args: EvalArg[]): CellValue {
	const re = numArg(args[0])
	if (typeof re !== 'number') return re
	const im = numArg(args[1])
	if (typeof im !== 'number') return im
	let suffix = 'i'
	if (args[2]) {
		const s = strArg(args[2])
		if (typeof s !== 'string') return s
		if (s !== 'i' && s !== 'j') return errorValue('#VALUE!')
		suffix = s
	}
	return stringValue(formatComplex(re, im, suffix))
}

function imreal(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	return numberValue(c.re)
}

function imaginary(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	return numberValue(c.im)
}

function imabs(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	return numberValue(Math.sqrt(c.re * c.re + c.im * c.im))
}

function imargument(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	if (c.re === 0 && c.im === 0) return errorValue('#DIV/0!')
	return numberValue(Math.atan2(c.im, c.re))
}

function imconjugate(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	return stringValue(formatComplex(c.re, -c.im, c.suffix || 'i'))
}

function imsum(args: EvalArg[]): CellValue {
	const cells = flattenArgs(args)
	let totalRe = 0
	let totalIm = 0
	let suffix = ''
	for (const cell of cells) {
		if (cell.kind === 'error') return cell
		let c: Complex | null
		if (cell.kind === 'number') {
			c = { re: cell.value, im: 0, suffix: '' }
		} else if (cell.kind === 'string') {
			c = parseComplex(cell.value)
		} else if (cell.kind === 'empty') {
			continue
		} else {
			return errorValue('#VALUE!')
		}
		if (c === null) return errorValue('#VALUE!')
		if (c.suffix && suffix && c.suffix !== suffix) return errorValue('#VALUE!')
		if (c.suffix) suffix = c.suffix
		totalRe += c.re
		totalIm += c.im
	}
	return stringValue(formatComplex(totalRe, totalIm, suffix || 'i'))
}

function imsub(args: EvalArg[]): CellValue {
	const a = complexArg(args[0])
	if (!isComplex(a)) return a
	const b = complexArg(args[1])
	if (!isComplex(b)) return b
	const sfx = resolveSuffix(a.suffix, b.suffix)
	if (sfx === null) return errorValue('#VALUE!')
	return stringValue(formatComplex(a.re - b.re, a.im - b.im, sfx))
}

function improduct(args: EvalArg[]): CellValue {
	const cells = flattenArgs(args)
	let re = 1
	let im = 0
	let suffix = ''
	for (const cell of cells) {
		if (cell.kind === 'error') return cell
		let c: Complex | null
		if (cell.kind === 'number') {
			c = { re: cell.value, im: 0, suffix: '' }
		} else if (cell.kind === 'string') {
			c = parseComplex(cell.value)
		} else if (cell.kind === 'empty') {
			continue
		} else {
			return errorValue('#VALUE!')
		}
		if (c === null) return errorValue('#VALUE!')
		if (c.suffix && suffix && c.suffix !== suffix) return errorValue('#VALUE!')
		if (c.suffix) suffix = c.suffix
		const newRe = re * c.re - im * c.im
		const newIm = re * c.im + im * c.re
		re = newRe
		im = newIm
	}
	return stringValue(formatComplex(re, im, suffix || 'i'))
}

function imdiv(args: EvalArg[]): CellValue {
	const a = complexArg(args[0])
	if (!isComplex(a)) return a
	const b = complexArg(args[1])
	if (!isComplex(b)) return b
	const sfx = resolveSuffix(a.suffix, b.suffix)
	if (sfx === null) return errorValue('#VALUE!')
	if (b.re === 0 && b.im === 0) return errorValue('#NUM!')
	if (Math.abs(b.re) >= Math.abs(b.im)) {
		const ratio = b.im / b.re
		const denom = b.re + b.im * ratio
		return stringValue(
			formatComplex((a.re + a.im * ratio) / denom, (a.im - a.re * ratio) / denom, sfx),
		)
	}
	const ratio = b.re / b.im
	const denom = b.im + b.re * ratio
	return stringValue(
		formatComplex((a.re * ratio + a.im) / denom, (a.im * ratio - a.re) / denom, sfx),
	)
}

function impower(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	const n = numArg(args[1])
	if (typeof n !== 'number') return n
	if (c.re === 0 && c.im === 0 && n < 0) return errorValue('#NUM!')
	if (c.re === 0 && c.im === 0) return stringValue(formatComplex(0, 0, c.suffix || 'i'))
	const sfx = c.suffix || 'i'
	if (Number.isInteger(n) && Math.abs(n) <= 100) {
		let absN = Math.abs(n)
		let rr = 1
		let ri = 0
		let br = c.re
		let bi = c.im
		while (absN > 0) {
			if (absN & 1) {
				const nr = rr * br - ri * bi
				const ni = rr * bi + ri * br
				rr = nr
				ri = ni
			}
			const nr = br * br - bi * bi
			const ni = 2 * br * bi
			br = nr
			bi = ni
			absN >>= 1
		}
		if (n < 0) {
			const d = rr * rr + ri * ri
			rr = rr / d
			ri = -ri / d
		}
		return stringValue(formatComplex(rr, ri, sfx))
	}
	const r = Math.sqrt(c.re * c.re + c.im * c.im)
	const theta = Math.atan2(c.im, c.re)
	const rn = r ** n
	return stringValue(formatComplex(rn * Math.cos(n * theta), rn * Math.sin(n * theta), sfx))
}

function imsqrt(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	const r = Math.sqrt(c.re * c.re + c.im * c.im)
	const theta = Math.atan2(c.im, c.re)
	const sqrtR = Math.sqrt(r)
	return stringValue(
		formatComplex(sqrtR * Math.cos(theta / 2), sqrtR * Math.sin(theta / 2), c.suffix || 'i'),
	)
}

function imexp(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	const ea = Math.exp(c.re)
	return stringValue(formatComplex(ea * Math.cos(c.im), ea * Math.sin(c.im), c.suffix || 'i'))
}

function imln(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	const r = Math.sqrt(c.re * c.re + c.im * c.im)
	if (r === 0) return errorValue('#NUM!')
	const theta = Math.atan2(c.im, c.re)
	return stringValue(formatComplex(Math.log(r), theta, c.suffix || 'i'))
}

function imsin(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	return stringValue(
		formatComplex(
			Math.sin(c.re) * Math.cosh(c.im),
			Math.cos(c.re) * Math.sinh(c.im),
			c.suffix || 'i',
		),
	)
}

function imcos(args: EvalArg[]): CellValue {
	const c = complexArg(args[0])
	if (!isComplex(c)) return c
	return stringValue(
		formatComplex(
			Math.cos(c.re) * Math.cosh(c.im),
			-Math.sin(c.re) * Math.sinh(c.im),
			c.suffix || 'i',
		),
	)
}

export const engineeringFunctions: FunctionDef[] = [
	{ name: 'BIN2DEC', minArgs: 1, maxArgs: 1, evaluate: bin2dec },
	{ name: 'DEC2BIN', minArgs: 1, maxArgs: 2, evaluate: dec2bin },
	{ name: 'HEX2DEC', minArgs: 1, maxArgs: 1, evaluate: hex2dec },
	{ name: 'DEC2HEX', minArgs: 1, maxArgs: 2, evaluate: dec2hex },
	{ name: 'OCT2DEC', minArgs: 1, maxArgs: 1, evaluate: oct2dec },
	{ name: 'DEC2OCT', minArgs: 1, maxArgs: 2, evaluate: dec2oct },
	{ name: 'BIN2HEX', minArgs: 1, maxArgs: 2, evaluate: bin2hex },
	{ name: 'BIN2OCT', minArgs: 1, maxArgs: 2, evaluate: bin2oct },
	{ name: 'HEX2BIN', minArgs: 1, maxArgs: 2, evaluate: hex2bin },
	{ name: 'HEX2OCT', minArgs: 1, maxArgs: 2, evaluate: hex2oct },
	{ name: 'OCT2BIN', minArgs: 1, maxArgs: 2, evaluate: oct2bin },
	{ name: 'OCT2HEX', minArgs: 1, maxArgs: 2, evaluate: oct2hex },
	{ name: 'DELTA', minArgs: 1, maxArgs: 2, evaluate: delta },
	{ name: 'GESTEP', minArgs: 1, maxArgs: 2, evaluate: gestep },
	{ name: 'BITAND', minArgs: 2, maxArgs: 2, evaluate: bitand },
	{ name: 'BITOR', minArgs: 2, maxArgs: 2, evaluate: bitor },
	{ name: 'BITXOR', minArgs: 2, maxArgs: 2, evaluate: bitxor },
	{ name: 'BITLSHIFT', minArgs: 2, maxArgs: 2, evaluate: bitlshift },
	{ name: 'BITRSHIFT', minArgs: 2, maxArgs: 2, evaluate: bitrshift },
	{ name: 'ERF', minArgs: 1, maxArgs: 2, evaluate: erf },
	{ name: 'ERF.PRECISE', minArgs: 1, maxArgs: 1, evaluate: erfPrecise },
	{ name: 'ERFC', minArgs: 1, maxArgs: 1, evaluate: erfc },
	{ name: 'ERFC.PRECISE', minArgs: 1, maxArgs: 1, evaluate: erfcPrecise },
	{ name: 'BESSELI', minArgs: 2, maxArgs: 2, evaluate: besselI_ },
	{ name: 'BESSELJ', minArgs: 2, maxArgs: 2, evaluate: besselJ_ },
	{ name: 'BESSELK', minArgs: 2, maxArgs: 2, evaluate: besselK_ },
	{ name: 'BESSELY', minArgs: 2, maxArgs: 2, evaluate: besselY_ },
	{ name: 'COMPLEX', minArgs: 2, maxArgs: 3, evaluate: complexFn },
	{ name: 'IMREAL', minArgs: 1, maxArgs: 1, evaluate: imreal },
	{ name: 'IMAGINARY', minArgs: 1, maxArgs: 1, evaluate: imaginary },
	{ name: 'IMABS', minArgs: 1, maxArgs: 1, evaluate: imabs },
	{ name: 'IMARGUMENT', minArgs: 1, maxArgs: 1, evaluate: imargument },
	{ name: 'IMCONJUGATE', minArgs: 1, maxArgs: 1, evaluate: imconjugate },
	{ name: 'IMSUM', minArgs: 1, maxArgs: 255, evaluate: imsum },
	{ name: 'IMSUB', minArgs: 2, maxArgs: 2, evaluate: imsub },
	{ name: 'IMPRODUCT', minArgs: 1, maxArgs: 255, evaluate: improduct },
	{ name: 'IMDIV', minArgs: 2, maxArgs: 2, evaluate: imdiv },
	{ name: 'IMPOWER', minArgs: 2, maxArgs: 2, evaluate: impower },
	{ name: 'IMSQRT', minArgs: 1, maxArgs: 1, evaluate: imsqrt },
	{ name: 'IMEXP', minArgs: 1, maxArgs: 1, evaluate: imexp },
	{ name: 'IMLN', minArgs: 1, maxArgs: 1, evaluate: imln },
	{ name: 'IMSIN', minArgs: 1, maxArgs: 1, evaluate: imsin },
	{ name: 'IMCOS', minArgs: 1, maxArgs: 1, evaluate: imcos },
	{
		name: 'IMLOG10',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const r = Math.sqrt(c.re * c.re + c.im * c.im)
			if (r === 0) return errorValue('#NUM!')
			const ln10 = Math.log(10)
			return stringValue(
				formatComplex(Math.log(r) / ln10, Math.atan2(c.im, c.re) / ln10, c.suffix || 'i'),
			)
		},
	},
	{
		name: 'IMLOG2',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const r = Math.sqrt(c.re * c.re + c.im * c.im)
			if (r === 0) return errorValue('#NUM!')
			const ln2 = Math.log(2)
			return stringValue(
				formatComplex(Math.log(r) / ln2, Math.atan2(c.im, c.re) / ln2, c.suffix || 'i'),
			)
		},
	},
	{
		name: 'IMTAN',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const d = Math.cos(2 * c.re) + Math.cosh(2 * c.im)
			if (d === 0) return errorValue('#NUM!')
			return stringValue(
				formatComplex(Math.sin(2 * c.re) / d, Math.sinh(2 * c.im) / d, c.suffix || 'i'),
			)
		},
	},
	{
		name: 'IMSINH',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			return stringValue(
				formatComplex(
					Math.sinh(c.re) * Math.cos(c.im),
					Math.cosh(c.re) * Math.sin(c.im),
					c.suffix || 'i',
				),
			)
		},
	},
	{
		name: 'IMCOSH',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			return stringValue(
				formatComplex(
					Math.cosh(c.re) * Math.cos(c.im),
					Math.sinh(c.re) * Math.sin(c.im),
					c.suffix || 'i',
				),
			)
		},
	},
	{
		name: 'IMSEC',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const cosRe = Math.cos(c.re) * Math.cosh(c.im)
			const cosIm = -Math.sin(c.re) * Math.sinh(c.im)
			const d = cosRe * cosRe + cosIm * cosIm
			if (d === 0) return errorValue('#NUM!')
			return stringValue(formatComplex(cosRe / d, -cosIm / d, c.suffix || 'i'))
		},
	},
	{
		name: 'IMCSC',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const sinRe = Math.sin(c.re) * Math.cosh(c.im)
			const sinIm = Math.cos(c.re) * Math.sinh(c.im)
			const d = sinRe * sinRe + sinIm * sinIm
			if (d === 0) return errorValue('#NUM!')
			return stringValue(formatComplex(sinRe / d, -sinIm / d, c.suffix || 'i'))
		},
	},
	{
		name: 'IMCOT',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const d = Math.cosh(2 * c.im) - Math.cos(2 * c.re)
			if (d === 0) return errorValue('#NUM!')
			return stringValue(
				formatComplex(Math.sin(2 * c.re) / d, -Math.sinh(2 * c.im) / d, c.suffix || 'i'),
			)
		},
	},
	{
		name: 'IMSECH',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const coshRe = Math.cosh(c.re) * Math.cos(c.im)
			const coshIm = Math.sinh(c.re) * Math.sin(c.im)
			const d = coshRe * coshRe + coshIm * coshIm
			if (d === 0) return errorValue('#NUM!')
			return stringValue(formatComplex(coshRe / d, -coshIm / d, c.suffix || 'i'))
		},
	},
	{
		name: 'IMCSCH',
		minArgs: 1,
		maxArgs: 1,
		evaluate: (args) => {
			const c = complexArg(args[0])
			if (!isComplex(c)) return c
			const sinhRe = Math.sinh(c.re) * Math.cos(c.im)
			const sinhIm = Math.cosh(c.re) * Math.sin(c.im)
			const d = sinhRe * sinhRe + sinhIm * sinhIm
			if (d === 0) return errorValue('#NUM!')
			return stringValue(formatComplex(sinhRe / d, -sinhIm / d, c.suffix || 'i'))
		},
	},
]
