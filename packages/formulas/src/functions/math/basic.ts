import type { CellValue } from '@ascend/schema'
import {
	arrayValue,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
	topLeftScalar,
} from '@ascend/schema'
import type { EvalArg, FunctionDef } from '../index.ts'
import { collectNumbers, getRange } from '../registry.ts'
import { fn, numArg } from './helpers.ts'

function matrixNumber(v: CellValue): number | null {
	if (v.kind === 'number') return v.value
	if (v.kind === 'date') return v.serial
	if (v.kind === 'boolean') return v.value ? 1 : 0
	return null
}

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
				const n = directMathNumber(arg)
				if (n === null) continue
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

	fn('SQRTPI', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		if (n < 0) return errorValue('#NUM!')
		return numberValue(Math.sqrt(n * Math.PI))
	}),

	fn('MULTINOMIAL', 1, 255, (args) => {
		const numsOrErr = collectNumbers(args)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		let total = 0
		let denom = 1
		for (const v of numsOrErr) {
			const k = Math.trunc(v)
			if (k < 0) return errorValue('#NUM!')
			total += k
			let f = 1
			for (let i = 2; i <= k; i++) f *= i
			denom *= f
		}
		let numer = 1
		for (let i = 2; i <= total; i++) numer *= i
		return numberValue(Math.round(numer / denom))
	}),

	fn('SERIESSUM', 4, 4, (args) => {
		const x = numArg(args[0])
		if (typeof x !== 'number') return x
		const n = numArg(args[1])
		if (typeof n !== 'number') return n
		const m = numArg(args[2])
		if (typeof m !== 'number') return m
		const coeffRange = getRange(args[3])
		const coeffs: number[] = []
		for (const row of coeffRange) {
			for (const cell of row) {
				if (cell.kind === 'error') return cell
				if (cell.kind === 'number') coeffs.push(cell.value)
				else if (cell.kind === 'date') coeffs.push(cell.serial)
				else coeffs.push(0)
			}
		}
		let sum = 0
		for (let i = 0; i < coeffs.length; i++) {
			sum += (coeffs[i] as number) * x ** (n + i * m)
		}
		return numberValue(sum)
	}),

	fn('SUMX2MY2', 2, 2, (args) => {
		const paired = collectPairedNumbers(args[0], args[1])
		if (!Array.isArray(paired)) return paired
		const [xs, ys] = paired
		let sum = 0
		for (let i = 0; i < xs.length; i++) {
			sum += (xs[i] as number) ** 2 - (ys[i] as number) ** 2
		}
		return numberValue(sum)
	}),

	fn('SUMX2PY2', 2, 2, (args) => {
		const paired = collectPairedNumbers(args[0], args[1])
		if (!Array.isArray(paired)) return paired
		const [xs, ys] = paired
		let sum = 0
		for (let i = 0; i < xs.length; i++) {
			sum += (xs[i] as number) ** 2 + (ys[i] as number) ** 2
		}
		return numberValue(sum)
	}),

	fn('SUMXMY2', 2, 2, (args) => {
		const paired = collectPairedNumbers(args[0], args[1])
		if (!Array.isArray(paired)) return paired
		const [xs, ys] = paired
		let sum = 0
		for (let i = 0; i < xs.length; i++) {
			sum += ((xs[i] as number) - (ys[i] as number)) ** 2
		}
		return numberValue(sum)
	}),

	fn('MMULT', 2, 2, (args) => {
		const a = getRange(args[0])
		const b = getRange(args[1])
		const aRows = a.length
		const aCols = a[0]?.length ?? 0
		const bRows = b.length
		const bCols = b[0]?.length ?? 0
		if (aCols !== bRows) return errorValue('#VALUE!')
		const rows: CellValue[][] = []
		for (let i = 0; i < aRows; i++) {
			const row: CellValue[] = []
			for (let j = 0; j < bCols; j++) {
				let sum = 0
				for (let k = 0; k < aCols; k++) {
					const av = matrixNumber(a[i]?.[k] ?? EMPTY)
					const bv = matrixNumber(b[k]?.[j] ?? EMPTY)
					if (av === null || bv === null) return errorValue('#VALUE!')
					sum += av * bv
				}
				row.push(numberValue(sum))
			}
			rows.push(row)
		}
		if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0] ?? EMPTY
		return arrayValue(rows.map((r) => r.map((c) => topLeftScalar(c))))
	}),

	fn('MDETERM', 1, 1, (args) => {
		const m = getRange(args[0])
		const n = m.length
		const cols = m[0]?.length ?? 0
		if (n !== cols || n === 0) return errorValue('#VALUE!')
		const mat: number[][] = []
		for (let i = 0; i < n; i++) {
			const row: number[] = []
			for (let j = 0; j < n; j++) {
				const v = matrixNumber(m[i]?.[j] ?? EMPTY)
				if (v === null) return errorValue('#VALUE!')
				row.push(v)
			}
			mat.push(row)
		}
		return numberValue(determinant(mat, n))
	}),

	fn('MINVERSE', 1, 1, (args) => {
		const m = getRange(args[0])
		const n = m.length
		const cols = m[0]?.length ?? 0
		if (n !== cols || n === 0) return errorValue('#VALUE!')
		const aug: number[][] = []
		for (let i = 0; i < n; i++) {
			const row: number[] = []
			for (let j = 0; j < n; j++) {
				const v = matrixNumber(m[i]?.[j] ?? EMPTY)
				if (v === null) return errorValue('#VALUE!')
				row.push(v)
			}
			for (let j = 0; j < n; j++) row.push(i === j ? 1 : 0)
			aug.push(row)
		}
		for (let col = 0; col < n; col++) {
			let maxRow = col
			let maxVal = Math.abs(aug[col]?.[col] ?? 0)
			for (let row = col + 1; row < n; row++) {
				const val = Math.abs(aug[row]?.[col] ?? 0)
				if (val > maxVal) {
					maxRow = row
					maxVal = val
				}
			}
			if (maxVal < 1e-15) return errorValue('#NUM!')
			if (maxRow !== col) {
				const tmp = aug[col]
				aug[col] = aug[maxRow] as number[]
				aug[maxRow] = tmp as number[]
			}
			const pivot = aug[col]?.[col] ?? 1
			const pivotRow = aug[col]
			if (!pivotRow) return errorValue('#NUM!')
			for (let j = 0; j < 2 * n; j++) pivotRow[j] = (pivotRow[j] ?? 0) / pivot
			for (let row = 0; row < n; row++) {
				if (row === col) continue
				const augRow = aug[row]
				if (!augRow) continue
				const factor = augRow[col] ?? 0
				for (let j = 0; j < 2 * n; j++) augRow[j] = (augRow[j] ?? 0) - factor * (pivotRow[j] ?? 0)
			}
		}
		const result: CellValue[][] = []
		for (let i = 0; i < n; i++) {
			const row: CellValue[] = []
			for (let j = 0; j < n; j++) row.push(numberValue(aug[i]?.[n + j] ?? 0))
			result.push(row)
		}
		if (n === 1) return result[0]?.[0] ?? EMPTY
		return arrayValue(result.map((r) => r.map((c) => topLeftScalar(c))))
	}),

	fn('MUNIT', 1, 1, (args) => {
		const d = numArg(args[0])
		if (typeof d !== 'number') return d
		const n = Math.trunc(d)
		if (n < 1) return errorValue('#VALUE!')
		const rows: CellValue[][] = []
		for (let i = 0; i < n; i++) {
			const row: CellValue[] = []
			for (let j = 0; j < n; j++) row.push(numberValue(i === j ? 1 : 0))
			rows.push(row)
		}
		if (n === 1) return numberValue(1)
		return arrayValue(rows.map((r) => r.map((c) => topLeftScalar(c))))
	}),
]

function collectPairedNumbers(
	arg1: EvalArg | undefined,
	arg2: EvalArg | undefined,
): [number[], number[]] | CellValue {
	const range1 = getRange(arg1)
	const range2 = getRange(arg2)
	const xs: number[] = []
	const ys: number[] = []
	const values1 = flattenRangeValues(range1)
	const values2 = flattenRangeValues(range2)
	if (values1.length !== values2.length) return errorValue('#N/A')
	for (let i = 0; i < values1.length; i++) {
		const v1 = values1[i] ?? EMPTY
		const v2 = values2[i] ?? EMPTY
		if (v1.kind === 'error') return v1
		if (v2.kind === 'error') return v2
		const n1 = v1.kind === 'number' ? v1.value : v1.kind === 'date' ? v1.serial : null
		const n2 = v2.kind === 'number' ? v2.value : v2.kind === 'date' ? v2.serial : null
		if (n1 !== null && n2 !== null) {
			xs.push(n1)
			ys.push(n2)
		}
	}
	if (xs.length === 0) return errorValue('#DIV/0!')
	return [xs, ys]
}

function directMathNumber(arg: EvalArg | undefined): number | null | CellValue {
	if (arg?.ref) {
		const value = arg.value ?? EMPTY
		if (value.kind === 'error') return value
		return value.kind === 'number' ? value.value : value.kind === 'date' ? value.serial : null
	}
	return numArg(arg)
}

function flattenRangeValues(range: readonly (readonly CellValue[])[]): CellValue[] {
	const values: CellValue[] = []
	for (const row of range) {
		for (const cell of row) values.push(cell)
	}
	return values
}

function determinant(m: number[][], n: number): number {
	if (n === 1) return m[0]?.[0] ?? 0
	if (n === 2) return (m[0]?.[0] ?? 0) * (m[1]?.[1] ?? 0) - (m[0]?.[1] ?? 0) * (m[1]?.[0] ?? 0)
	const mat = m.map((r) => [...r])
	let det = 1
	for (let col = 0; col < n; col++) {
		let maxRow = col
		let maxVal = Math.abs(mat[col]?.[col] ?? 0)
		for (let row = col + 1; row < n; row++) {
			const val = Math.abs(mat[row]?.[col] ?? 0)
			if (val > maxVal) {
				maxRow = row
				maxVal = val
			}
		}
		if (maxVal < 1e-15) return 0
		if (maxRow !== col) {
			const tmp = mat[col]
			mat[col] = mat[maxRow] as number[]
			mat[maxRow] = tmp as number[]
			det *= -1
		}
		det *= mat[col]?.[col] ?? 1
		const pivot = mat[col]?.[col] ?? 1
		for (let row = col + 1; row < n; row++) {
			const factor = (mat[row]?.[col] ?? 0) / pivot
			const matRow = mat[row]
			if (!matRow) continue
			for (let j = col + 1; j < n; j++) {
				matRow[j] = (matRow[j] ?? 0) - factor * (mat[col]?.[j] ?? 0)
			}
		}
	}
	return det
}
