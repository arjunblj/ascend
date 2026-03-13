import type { CellValue } from '@ascend/schema'
import { arrayValue, errorValue, numberValue, topLeftScalar } from '@ascend/schema'
import type { FunctionDef } from './registry.ts'
import { collectNumbers, type EvalArg, getRange, numArg } from './registry.ts'

function collectFrom(arg: EvalArg | undefined): number[] | CellValue {
	if (!arg) return []
	return collectNumbers([arg])
}

function quickselect(arr: number[], k: number, ascending: boolean): number {
	let lo = 0
	let hi = arr.length - 1
	while (lo < hi) {
		const pivot = arr[lo + ((hi - lo) >> 1)] as number
		let i = lo
		let j = hi
		while (i <= j) {
			if (ascending) {
				while ((arr[i] as number) < pivot) i++
				while ((arr[j] as number) > pivot) j--
			} else {
				while ((arr[i] as number) > pivot) i++
				while ((arr[j] as number) < pivot) j--
			}
			if (i <= j) {
				const tmp = arr[i] as number
				arr[i] = arr[j] as number
				arr[j] = tmp
				i++
				j--
			}
		}
		if (k <= j) hi = j
		else if (k >= i) lo = i
		else break
	}
	return arr[k] as number
}

function largeFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const ki = Math.floor(k)
	if (ki < 1 || ki > numsOrErr.length) return errorValue('#NUM!')
	return numberValue(quickselect(numsOrErr, ki - 1, false))
}

function smallFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const ki = Math.floor(k)
	if (ki < 1 || ki > numsOrErr.length) return errorValue('#NUM!')
	return numberValue(quickselect(numsOrErr, ki - 1, true))
}

function rankFn(args: EvalArg[]): CellValue {
	const num = numArg(args[0])
	if (typeof num !== 'number') return num
	const numsOrErr = collectFrom(args[1])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const order = args.length > 2 ? numArg(args[2]) : 0
	if (typeof order !== 'number') return order

	const ascending = order !== 0
	let rank = 1
	let found = false
	for (const v of numsOrErr) {
		if (v === num) found = true
		else if (ascending ? v < num : v > num) rank++
	}
	if (!found) return errorValue('#N/A')
	return numberValue(rank)
}

function percentileFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	if (k < 0 || k > 1 || numsOrErr.length === 0) return errorValue('#NUM!')

	numsOrErr.sort((a, b) => a - b)
	const n = numsOrErr.length
	const x = k * (n - 1)
	const i = Math.floor(x)
	const frac = x - i
	if (i + 1 >= n) return numberValue(numsOrErr[n - 1] ?? 0)
	return numberValue((numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)))
}

function medianFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#NUM!')

	numsOrErr.sort((a, b) => a - b)
	const mid = Math.floor(numsOrErr.length / 2)
	if (numsOrErr.length % 2 === 0) {
		return numberValue(((numsOrErr[mid - 1] ?? 0) + (numsOrErr[mid] ?? 0)) / 2)
	}
	return numberValue(numsOrErr[mid] ?? 0)
}

function stdevFn(args: EvalArg[], population = false): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const divisor = population ? numsOrErr.length : numsOrErr.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')

	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(Math.sqrt(sumSq / divisor))
}

function varFn(args: EvalArg[], population = false): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const divisor = population ? numsOrErr.length : numsOrErr.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')

	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(sumSq / divisor)
}

function percentileExcFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const n = numsOrErr.length
	if (n === 0 || k <= 0 || k >= 1) return errorValue('#NUM!')
	if (k < 1 / (n + 1) || k > n / (n + 1)) return errorValue('#NUM!')

	numsOrErr.sort((a, b) => a - b)
	const x = k * (n + 1) - 1
	const i = Math.floor(x)
	const frac = x - i
	if (i < 0) return numberValue(numsOrErr[0] ?? 0)
	if (i + 1 >= n) return numberValue(numsOrErr[n - 1] ?? 0)
	return numberValue((numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)))
}

function quartileFn(args: EvalArg[], exclusive = false): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const q = numArg(args[1])
	if (typeof q !== 'number') return q
	const qi = Math.trunc(q)
	if (exclusive) {
		if (qi < 1 || qi > 3 || numsOrErr.length === 0) return errorValue('#NUM!')
		numsOrErr.sort((a, b) => a - b)
		const n = numsOrErr.length
		const k = (qi / 4) * (n + 1) - 1
		const i = Math.floor(k)
		const frac = k - i
		if (i < 0) return numberValue(numsOrErr[0] as number)
		if (i + 1 >= n) return numberValue(numsOrErr[n - 1] as number)
		return numberValue(
			(numsOrErr[i] as number) + frac * ((numsOrErr[i + 1] as number) - (numsOrErr[i] as number)),
		)
	}
	if (qi < 0 || qi > 4 || numsOrErr.length === 0) return errorValue('#NUM!')
	numsOrErr.sort((a, b) => a - b)
	const n = numsOrErr.length
	if (qi === 0) return numberValue(numsOrErr[0] as number)
	if (qi === 4) return numberValue(numsOrErr[n - 1] as number)
	const k = (qi / 4) * (n - 1)
	const i = Math.floor(k)
	const frac = k - i
	if (i + 1 >= n) return numberValue(numsOrErr[n - 1] as number)
	return numberValue(
		(numsOrErr[i] as number) + frac * ((numsOrErr[i + 1] as number) - (numsOrErr[i] as number)),
	)
}

function modeFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#N/A')
	const freq = new Map<number, number>()
	let maxCount = 0
	let modeVal = 0
	for (const n of numsOrErr) {
		const c = (freq.get(n) ?? 0) + 1
		freq.set(n, c)
		if (c > maxCount) {
			maxCount = c
			modeVal = n
		}
	}
	if (maxCount < 2) return errorValue('#N/A')
	return numberValue(modeVal)
}

function collectNumbersA(args: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of args) {
		if (arg.kind === 'range' && arg.values) {
			for (const row of arg.values) {
				for (const cell of row) {
					if (cell.kind === 'error') return cell
					if (cell.kind === 'number') nums.push(cell.value)
					else if (cell.kind === 'date') nums.push(cell.serial)
					else if (cell.kind === 'boolean') nums.push(cell.value ? 1 : 0)
					else if (cell.kind === 'string') nums.push(0)
				}
			}
		} else {
			const n = numArg(arg)
			if (typeof n !== 'number') return n
			nums.push(n)
		}
	}
	return nums
}

function averageaFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbersA(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#DIV/0!')
	const sum = numsOrErr.reduce((a, b) => a + b, 0)
	return numberValue(sum / numsOrErr.length)
}

function maxaFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbersA(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return numberValue(0)
	let max = Number.NEGATIVE_INFINITY
	for (const v of numsOrErr) if (v > max) max = v
	return numberValue(max)
}

function minaFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbersA(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return numberValue(0)
	let min = Number.POSITIVE_INFINITY
	for (const v of numsOrErr) if (v < min) min = v
	return numberValue(min)
}

function rankAvgFn(args: EvalArg[]): CellValue {
	const num = numArg(args[0])
	if (typeof num !== 'number') return num
	const numsOrErr = collectFrom(args[1])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const order = args.length > 2 ? numArg(args[2]) : 0
	if (typeof order !== 'number') return order
	const ascending = order !== 0
	let rank = 1
	let found = false
	let tieCount = 0
	for (const v of numsOrErr) {
		if (v === num) {
			found = true
			tieCount++
		} else if (ascending ? v < num : v > num) rank++
	}
	if (!found) return errorValue('#N/A')
	return numberValue(rank + (tieCount - 1) / 2)
}

function geomeanFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#NUM!')
	let logSum = 0
	for (const v of numsOrErr) {
		if (v <= 0) return errorValue('#NUM!')
		logSum += Math.log(v)
	}
	return numberValue(Math.exp(logSum / numsOrErr.length))
}

function harmeanFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#NUM!')
	let recipSum = 0
	for (const v of numsOrErr) {
		if (v <= 0) return errorValue('#NUM!')
		recipSum += 1 / v
	}
	return numberValue(numsOrErr.length / recipSum)
}

function trimmeanFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const pct = numArg(args[1])
	if (typeof pct !== 'number') return pct
	if (pct < 0 || pct >= 1 || numsOrErr.length === 0) return errorValue('#NUM!')
	numsOrErr.sort((a, b) => a - b)
	const trimCount = Math.floor((numsOrErr.length * pct) / 2)
	const trimmed = numsOrErr.slice(trimCount, numsOrErr.length - trimCount)
	if (trimmed.length === 0) return errorValue('#NUM!')
	const sum = trimmed.reduce((a, b) => a + b, 0)
	return numberValue(sum / trimmed.length)
}

function percentrankIncFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const x = numArg(args[1])
	if (typeof x !== 'number') return x
	const sig = args.length > 2 ? numArg(args[2]) : 3
	if (typeof sig !== 'number') return sig
	const s = Math.floor(sig)
	if (s < 1 || numsOrErr.length === 0) return errorValue('#NUM!')
	numsOrErr.sort((a, b) => a - b)
	const n = numsOrErr.length
	if (x < (numsOrErr[0] as number) || x > (numsOrErr[n - 1] as number)) return errorValue('#N/A')
	if (n === 1) return numberValue(1)
	let i = 0
	while (i < n && (numsOrErr[i] as number) < x) i++
	let rank: number
	if (i < n && (numsOrErr[i] as number) === x) {
		rank = i / (n - 1)
	} else {
		const frac =
			(x - (numsOrErr[i - 1] as number)) / ((numsOrErr[i] as number) - (numsOrErr[i - 1] as number))
		rank = (i - 1 + frac) / (n - 1)
	}
	const factor = 10 ** s
	return numberValue(Math.floor(rank * factor) / factor)
}

function percentrankExcFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const x = numArg(args[1])
	if (typeof x !== 'number') return x
	const sig = args.length > 2 ? numArg(args[2]) : 3
	if (typeof sig !== 'number') return sig
	const s = Math.floor(sig)
	if (s < 1 || numsOrErr.length === 0) return errorValue('#NUM!')
	numsOrErr.sort((a, b) => a - b)
	const n = numsOrErr.length
	if (x < (numsOrErr[0] as number) || x > (numsOrErr[n - 1] as number)) return errorValue('#N/A')
	let i = 0
	while (i < n && (numsOrErr[i] as number) < x) i++
	let rank: number
	if (i < n && (numsOrErr[i] as number) === x) {
		rank = (i + 1) / (n + 1)
	} else {
		const frac =
			(x - (numsOrErr[i - 1] as number)) / ((numsOrErr[i] as number) - (numsOrErr[i - 1] as number))
		rank = (i + frac) / (n + 1)
	}
	const factor = 10 ** s
	return numberValue(Math.floor(rank * factor) / factor)
}

function collectPaired(
	arg1: EvalArg | undefined,
	arg2: EvalArg | undefined,
): [number[], number[]] | CellValue {
	const range1 = getRange(arg1)
	const range2 = getRange(arg2)
	const rows1 = range1.length
	const cols1 = range1[0]?.length ?? 0
	const rows2 = range2.length
	const cols2 = range2[0]?.length ?? 0
	if (rows1 !== rows2 || cols1 !== cols2) return errorValue('#N/A')
	const a: number[] = []
	const b: number[] = []
	for (let r = 0; r < rows1; r++) {
		for (let c = 0; c < cols1; c++) {
			const v1 = range1[r]?.[c]
			const v2 = range2[r]?.[c]
			if (!v1 || !v2) continue
			if (v1.kind === 'error') return v1
			if (v2.kind === 'error') return v2
			const n1 = v1.kind === 'number' ? v1.value : v1.kind === 'date' ? v1.serial : null
			const n2 = v2.kind === 'number' ? v2.value : v2.kind === 'date' ? v2.serial : null
			if (n1 !== null && n2 !== null) {
				a.push(n1)
				b.push(n2)
			}
		}
	}
	if (a.length === 0) return errorValue('#N/A')
	return [a, b]
}

function linregSums(ys: number[], xs: number[]) {
	const n = xs.length
	let sumX = 0
	let sumY = 0
	let sumXY = 0
	let sumX2 = 0
	let sumY2 = 0
	for (let i = 0; i < n; i++) {
		const x = xs[i] as number
		const y = ys[i] as number
		sumX += x
		sumY += y
		sumXY += x * y
		sumX2 += x * x
		sumY2 += y * y
	}
	const ssxx = n * sumX2 - sumX * sumX
	const ssyy = n * sumY2 - sumY * sumY
	const ssxy = n * sumXY - sumX * sumY
	return { n, sumX, sumY, ssxx, ssyy, ssxy }
}

function forecastLinearFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const paired = collectPaired(args[1], args[2])
	if (!Array.isArray(paired)) return paired
	const [ys, xs] = paired
	const { n, sumX, sumY, ssxx, ssxy } = linregSums(ys, xs)
	if (ssxx === 0) return errorValue('#DIV/0!')
	const slope = ssxy / ssxx
	const intercept = (sumY - slope * sumX) / n
	return numberValue(intercept + slope * x)
}

function slopeFn(args: EvalArg[]): CellValue {
	const paired = collectPaired(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const [ys, xs] = paired
	const { ssxx, ssxy } = linregSums(ys, xs)
	if (ssxx === 0) return errorValue('#DIV/0!')
	return numberValue(ssxy / ssxx)
}

function interceptFn(args: EvalArg[]): CellValue {
	const paired = collectPaired(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const [ys, xs] = paired
	const { n, sumX, sumY, ssxx, ssxy } = linregSums(ys, xs)
	if (ssxx === 0) return errorValue('#DIV/0!')
	const slope = ssxy / ssxx
	return numberValue((sumY - slope * sumX) / n)
}

function rsqFn(args: EvalArg[]): CellValue {
	const paired = collectPaired(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const [ys, xs] = paired
	const { ssxx, ssyy, ssxy } = linregSums(ys, xs)
	const denom = ssxx * ssyy
	if (denom === 0) return errorValue('#DIV/0!')
	return numberValue((ssxy * ssxy) / denom)
}

function correlFn(args: EvalArg[]): CellValue {
	const paired = collectPaired(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const [a, b] = paired
	const { ssxx, ssyy, ssxy } = linregSums(a, b)
	const denom = Math.sqrt(ssxx * ssyy)
	if (denom === 0) return errorValue('#DIV/0!')
	return numberValue(ssxy / denom)
}

function steyxFn(args: EvalArg[]): CellValue {
	const paired = collectPaired(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const [ys, xs] = paired
	const { n, sumX, sumY, ssxx, ssxy } = linregSums(ys, xs)
	if (n < 3) return errorValue('#DIV/0!')
	if (ssxx === 0) return errorValue('#DIV/0!')
	const slope = ssxy / ssxx
	const intercept = (sumY - slope * sumX) / n
	let sse = 0
	for (let i = 0; i < n; i++) {
		const residual = (ys[i] as number) - (intercept + slope * (xs[i] as number))
		sse += residual * residual
	}
	return numberValue(Math.sqrt(sse / (n - 2)))
}

function covarianceFn(args: EvalArg[], population: boolean): CellValue {
	const paired = collectPaired(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const [a, b] = paired
	const n = a.length
	const divisor = population ? n : n - 1
	if (divisor < 1) return errorValue('#DIV/0!')
	let sumA = 0
	let sumB = 0
	for (let i = 0; i < n; i++) {
		sumA += a[i] as number
		sumB += b[i] as number
	}
	const meanA = sumA / n
	const meanB = sumB / n
	let cov = 0
	for (let i = 0; i < n; i++) {
		cov += ((a[i] as number) - meanA) * ((b[i] as number) - meanB)
	}
	return numberValue(cov / divisor)
}

function avedevFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#NUM!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sum = numsOrErr.reduce((acc, v) => acc + Math.abs(v - mean), 0)
	return numberValue(sum / numsOrErr.length)
}

function devsqFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#NUM!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	return numberValue(numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0))
}

function kurtFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const n = numsOrErr.length
	if (n < 4) return errorValue('#DIV/0!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / n
	const s2 = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / (n - 1)
	if (s2 === 0) return errorValue('#DIV/0!')
	const s = Math.sqrt(s2)
	let m4 = 0
	for (const v of numsOrErr) m4 += ((v - mean) / s) ** 4
	return numberValue(
		(n * (n + 1) * m4) / ((n - 1) * (n - 2) * (n - 3)) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3)),
	)
}

function skewFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const n = numsOrErr.length
	if (n < 3) return errorValue('#DIV/0!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / n
	const s2 = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / (n - 1)
	if (s2 === 0) return errorValue('#DIV/0!')
	const s = Math.sqrt(s2)
	let m3 = 0
	for (const v of numsOrErr) m3 += ((v - mean) / s) ** 3
	return numberValue((n * m3) / ((n - 1) * (n - 2)))
}

function frequencyFn(args: EvalArg[]): CellValue {
	const dataOrErr = collectFrom(args[0])
	if (!Array.isArray(dataOrErr)) return dataOrErr
	const binsOrErr = collectFrom(args[1])
	if (!Array.isArray(binsOrErr)) return binsOrErr
	binsOrErr.sort((a, b) => a - b)
	const bLen = binsOrErr.length
	const counts = new Array<number>(bLen + 1).fill(0)
	for (const v of dataOrErr) {
		let idx = bLen
		for (let i = 0; i < bLen; i++) {
			if (v <= (binsOrErr[i] as number)) {
				idx = i
				break
			}
		}
		counts[idx] = (counts[idx] as number) + 1
	}
	return arrayValue(counts.map((c) => [topLeftScalar(numberValue(c))]))
}

function modeMultFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length === 0) return errorValue('#N/A')
	const freq = new Map<number, number>()
	let maxCount = 0
	for (const n of numsOrErr) {
		const c = (freq.get(n) ?? 0) + 1
		freq.set(n, c)
		if (c > maxCount) maxCount = c
	}
	if (maxCount < 2) return errorValue('#N/A')
	const modes: number[] = []
	for (const [val, count] of freq) {
		if (count === maxCount) modes.push(val)
	}
	modes.sort((a, b) => a - b)
	return arrayValue(modes.map((m) => [topLeftScalar(numberValue(m))]))
}

const SQRT_2PI = Math.sqrt(2 * Math.PI)

function normPdf(z: number): number {
	return Math.exp(-0.5 * z * z) / SQRT_2PI
}

function normCdf(z: number): number {
	if (z < 0) return 1 - normCdf(-z)
	const t = 1 / (1 + 0.2316419 * z)
	const poly =
		t * (0.31938153 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429))))
	return 1 - normPdf(z) * poly
}

function normSInvImpl(p: number): number {
	const y = p - 0.5
	if (Math.abs(y) < 0.42) {
		const r = y * y
		return (
			(y * (((-25.44106049637 * r + 41.39119773534) * r + -18.61500062529) * r + 2.50662823884)) /
			((((3.13082909833 * r + -21.06224101826) * r + 23.08336743743) * r + -8.4735109309) * r + 1)
		)
	}
	let r = y > 0 ? 1 - p : p
	r = Math.log(-Math.log(r))
	let z =
		0.3374754822726147 +
		r *
			(0.9761690190917186 +
				r *
					(0.1607979714918209 +
						r *
							(0.0276438810333863 +
								r *
									(0.0038405729373609 +
										r *
											(0.0003951896511919 +
												r *
													(0.0000321767881768 +
														r * (0.0000002888167364 + r * 0.0000003960315187)))))))
	if (y < 0) z = -z
	return z
}

const LANCZOS_C = [
	0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
	-176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
	1.5056327351493116e-7,
]

function gammalnImpl(x: number): number {
	if (x <= 0 && x === Math.floor(x)) return Number.POSITIVE_INFINITY
	if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - gammalnImpl(1 - x)
	const y = x - 1
	let a = LANCZOS_C[0] as number
	const t = y + 7.5
	for (let i = 1; i < LANCZOS_C.length; i++) a += (LANCZOS_C[i] as number) / (y + i)
	return 0.5 * Math.log(2 * Math.PI) + (y + 0.5) * Math.log(t) - t + Math.log(a)
}

function betaln(a: number, b: number): number {
	return gammalnImpl(a) + gammalnImpl(b) - gammalnImpl(a + b)
}

function betacf(x: number, a: number, b: number): number {
	const FPMIN = 1e-30
	const qab = a + b
	const qap = a + 1
	const qam = a - 1
	let c = 1
	let d = 1 - (qab * x) / qap
	if (Math.abs(d) < FPMIN) d = FPMIN
	d = 1 / d
	let h = d
	for (let m = 1; m <= 200; m++) {
		const m2 = 2 * m
		let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2))
		d = 1 + aa * d
		if (Math.abs(d) < FPMIN) d = FPMIN
		c = 1 + aa / c
		if (Math.abs(c) < FPMIN) c = FPMIN
		d = 1 / d
		h *= d * c
		aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2))
		d = 1 + aa * d
		if (Math.abs(d) < FPMIN) d = FPMIN
		c = 1 + aa / c
		if (Math.abs(c) < FPMIN) c = FPMIN
		d = 1 / d
		const del = d * c
		h *= del
		if (Math.abs(del - 1) < 3e-7) break
	}
	return h
}

function ibeta(x: number, a: number, b: number): number {
	if (x <= 0) return 0
	if (x >= 1) return 1
	const bt = Math.exp(
		gammalnImpl(a + b) - gammalnImpl(a) - gammalnImpl(b) + a * Math.log(x) + b * Math.log(1 - x),
	)
	if (x < (a + 1) / (a + b + 2)) return (bt * betacf(x, a, b)) / a
	return 1 - (bt * betacf(1 - x, b, a)) / b
}

function ibetainv(p: number, a: number, b: number): number {
	if (p <= 0) return 0
	if (p >= 1) return 1
	let x = 0.5
	for (let i = 0; i < 100; i++) {
		const fx = ibeta(x, a, b) - p
		const bt = Math.exp(
			gammalnImpl(a + b) -
				gammalnImpl(a) -
				gammalnImpl(b) +
				(a - 1) * Math.log(x) +
				(b - 1) * Math.log(1 - x),
		)
		if (bt === 0) break
		const dx = fx / bt
		x -= dx
		x = Math.max(1e-15, Math.min(1 - 1e-15, x))
		if (Math.abs(dx) < 1e-12) break
	}
	return x
}

function lowRegGamma(a: number, x: number): number {
	if (x <= 0) return 0
	if (a <= 0) return 1
	const aln = gammalnImpl(a)
	const maxIter = Math.ceil(Math.log(a >= 1 ? a : 1 / a) * 8.5 + a * 0.4 + 17)
	if (x < a + 1) {
		let ap = a
		let sum = 1 / a
		let del = sum
		for (let i = 1; i <= maxIter; i++) {
			ap += 1
			del *= x / ap
			sum += del
		}
		return sum * Math.exp(-x + a * Math.log(x) - aln)
	}
	let b2 = x + 1 - a
	let c = 1 / 1e-30
	let d = 1 / b2
	let h = d
	for (let i = 1; i <= maxIter; i++) {
		const an = -i * (i - a)
		b2 += 2
		d = an * d + b2
		c = b2 + an / c
		if (Math.abs(d) < 1e-30) d = 1e-30
		if (Math.abs(c) < 1e-30) c = 1e-30
		d = 1 / d
		h *= d * c
	}
	return 1 - h * Math.exp(-x + a * Math.log(x) - aln)
}

function combiln(n: number, k: number): number {
	return gammalnImpl(n + 1) - gammalnImpl(k + 1) - gammalnImpl(n - k + 1)
}

function normDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const mean = numArg(args[1])
	if (typeof mean !== 'number') return mean
	const stdev = numArg(args[2])
	if (typeof stdev !== 'number') return stdev
	const cumulative = numArg(args[3])
	if (typeof cumulative !== 'number') return cumulative
	if (stdev <= 0) return errorValue('#NUM!')
	const z = (x - mean) / stdev
	if (cumulative !== 0) return numberValue(normCdf(z))
	return numberValue(normPdf(z) / stdev)
}

function normInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const mean = numArg(args[1])
	if (typeof mean !== 'number') return mean
	const stdev = numArg(args[2])
	if (typeof stdev !== 'number') return stdev
	if (p <= 0 || p >= 1 || stdev <= 0) return errorValue('#NUM!')
	return numberValue(mean + stdev * normSInvImpl(p))
}

function normSDistFn(args: EvalArg[]): CellValue {
	const z = numArg(args[0])
	if (typeof z !== 'number') return z
	const cumulative = numArg(args[1])
	if (typeof cumulative !== 'number') return cumulative
	if (cumulative !== 0) return numberValue(normCdf(z))
	return numberValue(normPdf(z))
}

function normSInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	if (p <= 0 || p >= 1) return errorValue('#NUM!')
	return numberValue(normSInvImpl(p))
}

function gammaLnFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	if (x <= 0) return errorValue('#NUM!')
	return numberValue(gammalnImpl(x))
}

function gammaFnExcel(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	if (x <= 0 && x === Math.floor(x)) return errorValue('#NUM!')
	return numberValue(Math.exp(gammalnImpl(x)))
}

function tDistFn(args: EvalArg[]): CellValue {
	const t = numArg(args[0])
	if (typeof t !== 'number') return t
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	const cum = numArg(args[2])
	if (typeof cum !== 'number') return cum
	if (df < 1) return errorValue('#NUM!')
	if (cum !== 0) {
		const x = df / (df + t * t)
		const p = 0.5 * ibeta(x, df / 2, 0.5)
		return numberValue(t >= 0 ? 1 - p : p)
	}
	return numberValue(
		Math.exp(gammalnImpl((df + 1) / 2) - gammalnImpl(df / 2)) /
			(Math.sqrt(df * Math.PI) * (1 + (t * t) / df) ** ((df + 1) / 2)),
	)
}

function tDist2TFn(args: EvalArg[]): CellValue {
	const t = numArg(args[0])
	if (typeof t !== 'number') return t
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (t < 0 || df < 1) return errorValue('#NUM!')
	const x = df / (df + t * t)
	return numberValue(ibeta(x, df / 2, 0.5))
}

function tDistRTFn(args: EvalArg[]): CellValue {
	const t = numArg(args[0])
	if (typeof t !== 'number') return t
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (t < 0 || df < 1) return errorValue('#NUM!')
	const x = df / (df + t * t)
	return numberValue(0.5 * ibeta(x, df / 2, 0.5))
}

function tInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (p <= 0 || p >= 1 || df < 1) return errorValue('#NUM!')
	const x = ibetainv(2 * Math.min(p, 1 - p), df / 2, 0.5)
	const t = Math.sqrt((df * (1 - x)) / x)
	return numberValue(p < 0.5 ? -t : t)
}

function tInv2TFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (p <= 0 || p > 1 || df < 1) return errorValue('#NUM!')
	const x = ibetainv(p, df / 2, 0.5)
	return numberValue(Math.sqrt((df * (1 - x)) / x))
}

function fDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	if (x < 0 || d1 < 1 || d2 < 1) return errorValue('#NUM!')
	if (x === 0) return cum !== 0 ? numberValue(0) : errorValue('#NUM!')
	if (cum !== 0) {
		const w = (d1 * x) / (d1 * x + d2)
		return numberValue(ibeta(w, d1 / 2, d2 / 2))
	}
	return numberValue(
		Math.sqrt(((d1 * x) ** d1 * d2 ** d2) / (d1 * x + d2) ** (d1 + d2)) /
			(x * Math.exp(betaln(d1 / 2, d2 / 2))),
	)
}

function fDistRTFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	if (x < 0 || d1 < 1 || d2 < 1) return errorValue('#NUM!')
	if (x === 0) return numberValue(1)
	const w = (d1 * x) / (d1 * x + d2)
	return numberValue(1 - ibeta(w, d1 / 2, d2 / 2))
}

function fInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	if (p < 0 || p > 1 || d1 < 1 || d2 < 1) return errorValue('#NUM!')
	if (p === 0) return numberValue(0)
	if (p === 1) return errorValue('#NUM!')
	const x = ibetainv(p, d1 / 2, d2 / 2)
	return numberValue((x * d2) / (d1 * (1 - x)))
}

function fInvRTFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	if (p < 0 || p > 1 || d1 < 1 || d2 < 1) return errorValue('#NUM!')
	if (p === 1) return numberValue(0)
	if (p === 0) return errorValue('#NUM!')
	const x = ibetainv(1 - p, d1 / 2, d2 / 2)
	return numberValue((x * d2) / (d1 * (1 - x)))
}

function chisqDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	const cum = numArg(args[2])
	if (typeof cum !== 'number') return cum
	if (x < 0 || df < 1) return errorValue('#NUM!')
	if (cum !== 0) return numberValue(lowRegGamma(df / 2, x / 2))
	const hk = df / 2
	return numberValue(Math.exp((hk - 1) * Math.log(x) - x / 2 - hk * Math.log(2) - gammalnImpl(hk)))
}

function chisqDistRTFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (x < 0 || df < 1) return errorValue('#NUM!')
	return numberValue(1 - lowRegGamma(df / 2, x / 2))
}

function chisqInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (p < 0 || p > 1 || df < 1) return errorValue('#NUM!')
	if (p === 0) return numberValue(0)
	if (p === 1) return errorValue('#NUM!')
	let x = df
	for (let i = 0; i < 100; i++) {
		const fx = lowRegGamma(df / 2, x / 2) - p
		const pdf = Math.exp(
			(df / 2 - 1) * Math.log(x) - x / 2 - (df / 2) * Math.log(2) - gammalnImpl(df / 2),
		)
		if (pdf === 0) break
		const dx = fx / pdf
		x -= dx
		if (x <= 0) x = 1e-10
		if (Math.abs(dx) < 1e-12) break
	}
	return numberValue(x)
}

function chisqInvRTFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const df = numArg(args[1])
	if (typeof df !== 'number') return df
	if (p < 0 || p > 1 || df < 1) return errorValue('#NUM!')
	const result = chisqInvFn([{ value: numberValue(1 - p) }, { value: numberValue(df) }])
	return result
}

function betaDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const alpha = numArg(args[1])
	if (typeof alpha !== 'number') return alpha
	const beta = numArg(args[2])
	if (typeof beta !== 'number') return beta
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	const A = args.length > 4 ? numArg(args[4]) : 0
	if (typeof A !== 'number') return A
	const B = args.length > 5 ? numArg(args[5]) : 1
	if (typeof B !== 'number') return B
	if (alpha <= 0 || beta <= 0 || x < A || x > B) return errorValue('#NUM!')
	const z = (x - A) / (B - A)
	if (cum !== 0) return numberValue(ibeta(z, alpha, beta))
	return numberValue(
		(z ** (alpha - 1) * (1 - z) ** (beta - 1)) / ((B - A) * Math.exp(betaln(alpha, beta))),
	)
}

function betaInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const alpha = numArg(args[1])
	if (typeof alpha !== 'number') return alpha
	const beta = numArg(args[2])
	if (typeof beta !== 'number') return beta
	const A = args.length > 3 ? numArg(args[3]) : 0
	if (typeof A !== 'number') return A
	const B = args.length > 4 ? numArg(args[4]) : 1
	if (typeof B !== 'number') return B
	if (p < 0 || p > 1 || alpha <= 0 || beta <= 0) return errorValue('#NUM!')
	return numberValue(A + ibetainv(p, alpha, beta) * (B - A))
}

function binomDistFn(args: EvalArg[]): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const n = numArg(args[1])
	if (typeof n !== 'number') return n
	const p = numArg(args[2])
	if (typeof p !== 'number') return p
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	const k = Math.trunc(s)
	const ni = Math.trunc(n)
	if (k < 0 || k > ni || ni < 0 || p < 0 || p > 1) return errorValue('#NUM!')
	const pmf = (i: number) => {
		if (p === 0) return i === 0 ? 1 : 0
		if (p === 1) return i === ni ? 1 : 0
		return Math.exp(combiln(ni, i) + i * Math.log(p) + (ni - i) * Math.log(1 - p))
	}
	if (cum === 0) return numberValue(pmf(k))
	let sum = 0
	for (let i = 0; i <= k; i++) sum += pmf(i)
	return numberValue(Math.min(1, sum))
}

function binomInvFn(args: EvalArg[]): CellValue {
	const trials = numArg(args[0])
	if (typeof trials !== 'number') return trials
	const p = numArg(args[1])
	if (typeof p !== 'number') return p
	const alpha = numArg(args[2])
	if (typeof alpha !== 'number') return alpha
	const n = Math.trunc(trials)
	if (n < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) return errorValue('#NUM!')
	let sum = 0
	const pmf = (i: number) => {
		if (p === 0) return i === 0 ? 1 : 0
		if (p === 1) return i === n ? 1 : 0
		return Math.exp(combiln(n, i) + i * Math.log(p) + (n - i) * Math.log(1 - p))
	}
	for (let i = 0; i <= n; i++) {
		sum += pmf(i)
		if (sum >= alpha) return numberValue(i)
	}
	return numberValue(n)
}

function binomDistRangeFn(args: EvalArg[]): CellValue {
	const trials = numArg(args[0])
	if (typeof trials !== 'number') return trials
	const p = numArg(args[1])
	if (typeof p !== 'number') return p
	const s1 = numArg(args[2])
	if (typeof s1 !== 'number') return s1
	const s2 = args.length > 3 ? numArg(args[3]) : s1
	if (typeof s2 !== 'number') return s2
	const n = Math.trunc(trials)
	const lo = Math.trunc(s1)
	const hi = Math.trunc(s2)
	if (n < 0 || p < 0 || p > 1 || lo < 0 || hi > n || lo > hi) return errorValue('#NUM!')
	let sum = 0
	for (let i = lo; i <= hi; i++) {
		sum += Math.exp(combiln(n, i) + i * Math.log(p) + (n - i) * Math.log(1 - p))
	}
	return numberValue(sum)
}

function poissonDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const mean = numArg(args[1])
	if (typeof mean !== 'number') return mean
	const cum = numArg(args[2])
	if (typeof cum !== 'number') return cum
	const k = Math.trunc(x)
	if (k < 0 || mean < 0) return errorValue('#NUM!')
	const pmf = (i: number) =>
		mean === 0 ? (i === 0 ? 1 : 0) : Math.exp(-mean + i * Math.log(mean) - gammalnImpl(i + 1))
	if (cum === 0) return numberValue(pmf(k))
	let sum = 0
	for (let i = 0; i <= k; i++) sum += pmf(i)
	return numberValue(Math.min(1, sum))
}

function exponDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const lambda = numArg(args[1])
	if (typeof lambda !== 'number') return lambda
	const cum = numArg(args[2])
	if (typeof cum !== 'number') return cum
	if (x < 0 || lambda <= 0) return errorValue('#NUM!')
	if (cum !== 0) return numberValue(1 - Math.exp(-lambda * x))
	return numberValue(lambda * Math.exp(-lambda * x))
}

function weibullDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const alpha = numArg(args[1])
	if (typeof alpha !== 'number') return alpha
	const beta = numArg(args[2])
	if (typeof beta !== 'number') return beta
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	if (x < 0 || alpha <= 0 || beta <= 0) return errorValue('#NUM!')
	if (cum !== 0) return numberValue(1 - Math.exp(-((x / beta) ** alpha)))
	return numberValue((alpha / beta) * (x / beta) ** (alpha - 1) * Math.exp(-((x / beta) ** alpha)))
}

function gammaDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const alpha = numArg(args[1])
	if (typeof alpha !== 'number') return alpha
	const beta = numArg(args[2])
	if (typeof beta !== 'number') return beta
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	if (x < 0 || alpha <= 0 || beta <= 0) return errorValue('#NUM!')
	if (cum !== 0) return numberValue(lowRegGamma(alpha, x / beta))
	return numberValue(
		Math.exp((alpha - 1) * Math.log(x) - x / beta - alpha * Math.log(beta) - gammalnImpl(alpha)),
	)
}

function gammaInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const alpha = numArg(args[1])
	if (typeof alpha !== 'number') return alpha
	const beta = numArg(args[2])
	if (typeof beta !== 'number') return beta
	if (p < 0 || p > 1 || alpha <= 0 || beta <= 0) return errorValue('#NUM!')
	if (p === 0) return numberValue(0)
	if (p === 1) return errorValue('#NUM!')
	let x = alpha * beta
	for (let i = 0; i < 100; i++) {
		const fx = lowRegGamma(alpha, x / beta) - p
		const pdf = Math.exp(
			(alpha - 1) * Math.log(x) - x / beta - alpha * Math.log(beta) - gammalnImpl(alpha),
		)
		if (pdf === 0) break
		const dx = fx / pdf
		x -= dx
		if (x <= 0) x = 1e-10
		if (Math.abs(dx) < 1e-12) break
	}
	return numberValue(x)
}

function lognormDistFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const mu = numArg(args[1])
	if (typeof mu !== 'number') return mu
	const sigma = numArg(args[2])
	if (typeof sigma !== 'number') return sigma
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	if (x <= 0 || sigma <= 0) return errorValue('#NUM!')
	const z = (Math.log(x) - mu) / sigma
	if (cum !== 0) return numberValue(normCdf(z))
	return numberValue(normPdf(z) / (x * sigma))
}

function lognormInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const mu = numArg(args[1])
	if (typeof mu !== 'number') return mu
	const sigma = numArg(args[2])
	if (typeof sigma !== 'number') return sigma
	if (p <= 0 || p >= 1 || sigma <= 0) return errorValue('#NUM!')
	return numberValue(Math.exp(mu + sigma * normSInvImpl(p)))
}

function negbinomDistFn(args: EvalArg[]): CellValue {
	const f = numArg(args[0])
	if (typeof f !== 'number') return f
	const s = numArg(args[1])
	if (typeof s !== 'number') return s
	const p = numArg(args[2])
	if (typeof p !== 'number') return p
	const cum = numArg(args[3])
	if (typeof cum !== 'number') return cum
	const k = Math.trunc(f)
	const r = Math.trunc(s)
	if (k < 0 || r < 1 || p <= 0 || p >= 1) return errorValue('#NUM!')
	const pmf = (i: number) => Math.exp(combiln(i + r - 1, i) + r * Math.log(p) + i * Math.log(1 - p))
	if (cum === 0) return numberValue(pmf(k))
	let sum = 0
	for (let i = 0; i <= k; i++) sum += pmf(i)
	return numberValue(Math.min(1, sum))
}

function hypgeomDistFn(args: EvalArg[]): CellValue {
	const sampleS = numArg(args[0])
	if (typeof sampleS !== 'number') return sampleS
	const nSample = numArg(args[1])
	if (typeof nSample !== 'number') return nSample
	const popS = numArg(args[2])
	if (typeof popS !== 'number') return popS
	const nPop = numArg(args[3])
	if (typeof nPop !== 'number') return nPop
	const cum = numArg(args[4])
	if (typeof cum !== 'number') return cum
	const k = Math.trunc(sampleS)
	const n = Math.trunc(nSample)
	const K = Math.trunc(popS)
	const N = Math.trunc(nPop)
	if (k < 0 || n < 0 || K < 0 || N < 0 || n > N || K > N || k > n || k > K)
		return errorValue('#NUM!')
	const pmf = (i: number) => Math.exp(combiln(K, i) + combiln(N - K, n - i) - combiln(N, n))
	if (cum === 0) return numberValue(pmf(k))
	let sum = 0
	const lo = Math.max(0, n - (N - K))
	for (let i = lo; i <= k; i++) sum += pmf(i)
	return numberValue(Math.min(1, sum))
}

function fisherFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	if (x <= -1 || x >= 1) return errorValue('#NUM!')
	return numberValue(0.5 * Math.log((1 + x) / (1 - x)))
}

function fisherInvFn(args: EvalArg[]): CellValue {
	const y = numArg(args[0])
	if (typeof y !== 'number') return y
	const e2y = Math.exp(2 * y)
	return numberValue((e2y - 1) / (e2y + 1))
}

function standardizeFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const mean = numArg(args[1])
	if (typeof mean !== 'number') return mean
	const stdev = numArg(args[2])
	if (typeof stdev !== 'number') return stdev
	if (stdev <= 0) return errorValue('#NUM!')
	return numberValue((x - mean) / stdev)
}

function phiFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	return numberValue(normPdf(x))
}

function gaussFn(args: EvalArg[]): CellValue {
	const z = numArg(args[0])
	if (typeof z !== 'number') return z
	return numberValue(normCdf(z) - 0.5)
}

function confidenceNormFn(args: EvalArg[]): CellValue {
	const alpha = numArg(args[0])
	if (typeof alpha !== 'number') return alpha
	const stdev = numArg(args[1])
	if (typeof stdev !== 'number') return stdev
	const size = numArg(args[2])
	if (typeof size !== 'number') return size
	if (alpha <= 0 || alpha >= 1 || stdev <= 0 || size < 1) return errorValue('#NUM!')
	return numberValue((normSInvImpl(1 - alpha / 2) * stdev) / Math.sqrt(size))
}

function confidenceTFn(args: EvalArg[]): CellValue {
	const alpha = numArg(args[0])
	if (typeof alpha !== 'number') return alpha
	const stdev = numArg(args[1])
	if (typeof stdev !== 'number') return stdev
	const size = numArg(args[2])
	if (typeof size !== 'number') return size
	const n = Math.trunc(size)
	if (alpha <= 0 || alpha >= 1 || stdev <= 0 || n < 1) return errorValue('#NUM!')
	const df = n - 1
	if (df < 1) return errorValue('#NUM!')
	const x = ibetainv(alpha, df / 2, 0.5)
	const t = Math.sqrt((df * (1 - x)) / x)
	return numberValue((t * stdev) / Math.sqrt(n))
}

function skewPFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const n = numsOrErr.length
	if (n < 3) return errorValue('#DIV/0!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / n
	const s2 = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / n
	if (s2 === 0) return errorValue('#DIV/0!')
	const s = Math.sqrt(s2)
	let m3 = 0
	for (const v of numsOrErr) m3 += ((v - mean) / s) ** 3
	return numberValue(m3 / n)
}

function stdevAFn(args: EvalArg[], population: boolean): CellValue {
	const numsOrErr = collectNumbersA(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const divisor = population ? numsOrErr.length : numsOrErr.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(Math.sqrt(sumSq / divisor))
}

function varAFn(args: EvalArg[], population: boolean): CellValue {
	const numsOrErr = collectNumbersA(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const divisor = population ? numsOrErr.length : numsOrErr.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')
	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(sumSq / divisor)
}

function linestFn(args: EvalArg[]): CellValue {
	const yRange = getRange(args[0])
	const xRange = args.length > 1 ? getRange(args[1]) : yRange.map((_, i) => [numberValue(i + 1)])
	const ys: number[] = []
	const xs: number[] = []
	const yr = yRange.length
	const yc = yRange[0]?.length ?? 0
	const xr = xRange.length
	const xc = xRange[0]?.length ?? 0
	if (yr !== xr || yc !== xc) return errorValue('#REF!')
	for (let r = 0; r < yr; r++) {
		for (let c = 0; c < yc; c++) {
			const yv = yRange[r]?.[c]
			const xv = xRange[r]?.[c]
			if (!yv || !xv) continue
			const yn = yv.kind === 'number' ? yv.value : yv.kind === 'date' ? yv.serial : null
			const xn = xv.kind === 'number' ? xv.value : xv.kind === 'date' ? xv.serial : null
			if (yn !== null && xn !== null) {
				ys.push(yn)
				xs.push(xn)
			}
		}
	}
	if (ys.length < 2) return errorValue('#N/A')
	const { n, sumX, sumY, ssxx, ssxy } = linregSums(ys, xs)
	if (ssxx === 0) return errorValue('#DIV/0!')
	const slope = ssxy / ssxx
	const intercept = (sumY - slope * sumX) / n
	return arrayValue([[topLeftScalar(numberValue(slope)), topLeftScalar(numberValue(intercept))]])
}

export const statsFunctions: FunctionDef[] = [
	{ name: 'LARGE', minArgs: 2, maxArgs: 2, evaluate: largeFn },
	{ name: 'SMALL', minArgs: 2, maxArgs: 2, evaluate: smallFn },
	{ name: 'RANK', minArgs: 2, maxArgs: 3, evaluate: rankFn },
	{ name: 'PERCENTILE', minArgs: 2, maxArgs: 2, evaluate: percentileFn },
	{ name: 'MEDIAN', minArgs: 1, maxArgs: 255, evaluate: medianFn },
	{ name: 'STDEV', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevFn(args) },
	{ name: 'STDEV.S', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevFn(args) },
	{ name: 'STDEV.P', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevFn(args, true) },
	{ name: 'STDEVP', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevFn(args, true) },
	{ name: 'VAR', minArgs: 1, maxArgs: 255, evaluate: (args) => varFn(args) },
	{ name: 'VAR.S', minArgs: 1, maxArgs: 255, evaluate: (args) => varFn(args) },
	{ name: 'VAR.P', minArgs: 1, maxArgs: 255, evaluate: (args) => varFn(args, true) },
	{ name: 'VARP', minArgs: 1, maxArgs: 255, evaluate: (args) => varFn(args, true) },
	{ name: 'PERCENTILE.INC', minArgs: 2, maxArgs: 2, evaluate: percentileFn },
	{ name: 'PERCENTILE.EXC', minArgs: 2, maxArgs: 2, evaluate: percentileExcFn },
	{ name: 'QUARTILE', minArgs: 2, maxArgs: 2, evaluate: (args) => quartileFn(args) },
	{ name: 'QUARTILE.INC', minArgs: 2, maxArgs: 2, evaluate: (args) => quartileFn(args) },
	{ name: 'QUARTILE.EXC', minArgs: 2, maxArgs: 2, evaluate: (args) => quartileFn(args, true) },
	{ name: 'MODE', minArgs: 1, maxArgs: 255, evaluate: modeFn },
	{ name: 'MODE.SNGL', minArgs: 1, maxArgs: 255, evaluate: modeFn },
	{ name: 'AVERAGEA', minArgs: 1, maxArgs: 255, evaluate: averageaFn },
	{ name: 'MAXA', minArgs: 1, maxArgs: 255, evaluate: maxaFn },
	{ name: 'MINA', minArgs: 1, maxArgs: 255, evaluate: minaFn },
	{ name: 'RANK.EQ', minArgs: 2, maxArgs: 3, evaluate: rankFn },
	{ name: 'RANK.AVG', minArgs: 2, maxArgs: 3, evaluate: rankAvgFn },
	{ name: 'GEOMEAN', minArgs: 1, maxArgs: 255, evaluate: geomeanFn },
	{ name: 'HARMEAN', minArgs: 1, maxArgs: 255, evaluate: harmeanFn },
	{ name: 'TRIMMEAN', minArgs: 2, maxArgs: 2, evaluate: trimmeanFn },
	{ name: 'PERCENTRANK.INC', minArgs: 2, maxArgs: 3, evaluate: percentrankIncFn },
	{ name: 'PERCENTRANK.EXC', minArgs: 2, maxArgs: 3, evaluate: percentrankExcFn },
	{ name: 'FORECAST.LINEAR', minArgs: 3, maxArgs: 3, evaluate: forecastLinearFn },
	{ name: 'FORECAST', minArgs: 3, maxArgs: 3, evaluate: forecastLinearFn },
	{ name: 'SLOPE', minArgs: 2, maxArgs: 2, evaluate: slopeFn },
	{ name: 'INTERCEPT', minArgs: 2, maxArgs: 2, evaluate: interceptFn },
	{ name: 'RSQ', minArgs: 2, maxArgs: 2, evaluate: rsqFn },
	{ name: 'CORREL', minArgs: 2, maxArgs: 2, evaluate: correlFn },
	{ name: 'PEARSON', minArgs: 2, maxArgs: 2, evaluate: correlFn },
	{ name: 'STEYX', minArgs: 2, maxArgs: 2, evaluate: steyxFn },
	{ name: 'COVARIANCE.P', minArgs: 2, maxArgs: 2, evaluate: (args) => covarianceFn(args, true) },
	{ name: 'COVARIANCE.S', minArgs: 2, maxArgs: 2, evaluate: (args) => covarianceFn(args, false) },
	{ name: 'AVEDEV', minArgs: 1, maxArgs: 255, evaluate: avedevFn },
	{ name: 'DEVSQ', minArgs: 1, maxArgs: 255, evaluate: devsqFn },
	{ name: 'KURT', minArgs: 1, maxArgs: 255, evaluate: kurtFn },
	{ name: 'SKEW', minArgs: 1, maxArgs: 255, evaluate: skewFn },
	{ name: 'FREQUENCY', minArgs: 2, maxArgs: 2, evaluate: frequencyFn },
	{ name: 'MODE.MULT', minArgs: 1, maxArgs: 255, evaluate: modeMultFn },
	{ name: 'NORM.DIST', minArgs: 4, maxArgs: 4, evaluate: normDistFn },
	{ name: 'NORM.INV', minArgs: 3, maxArgs: 3, evaluate: normInvFn },
	{ name: 'NORM.S.DIST', minArgs: 2, maxArgs: 2, evaluate: normSDistFn },
	{ name: 'NORM.S.INV', minArgs: 1, maxArgs: 1, evaluate: normSInvFn },
	{ name: 'GAMMALN', minArgs: 1, maxArgs: 1, evaluate: gammaLnFn },
	{ name: 'GAMMALN.PRECISE', minArgs: 1, maxArgs: 1, evaluate: gammaLnFn },
	{ name: 'GAMMA', minArgs: 1, maxArgs: 1, evaluate: gammaFnExcel },
	{ name: 'T.DIST', minArgs: 3, maxArgs: 3, evaluate: tDistFn },
	{ name: 'T.DIST.2T', minArgs: 2, maxArgs: 2, evaluate: tDist2TFn },
	{ name: 'T.DIST.RT', minArgs: 2, maxArgs: 2, evaluate: tDistRTFn },
	{ name: 'T.INV', minArgs: 2, maxArgs: 2, evaluate: tInvFn },
	{ name: 'T.INV.2T', minArgs: 2, maxArgs: 2, evaluate: tInv2TFn },
	{ name: 'F.DIST', minArgs: 4, maxArgs: 4, evaluate: fDistFn },
	{ name: 'F.DIST.RT', minArgs: 3, maxArgs: 3, evaluate: fDistRTFn },
	{ name: 'F.INV', minArgs: 3, maxArgs: 3, evaluate: fInvFn },
	{ name: 'F.INV.RT', minArgs: 3, maxArgs: 3, evaluate: fInvRTFn },
	{ name: 'CHISQ.DIST', minArgs: 3, maxArgs: 3, evaluate: chisqDistFn },
	{ name: 'CHISQ.DIST.RT', minArgs: 2, maxArgs: 2, evaluate: chisqDistRTFn },
	{ name: 'CHISQ.INV', minArgs: 2, maxArgs: 2, evaluate: chisqInvFn },
	{ name: 'CHISQ.INV.RT', minArgs: 2, maxArgs: 2, evaluate: chisqInvRTFn },
	{ name: 'BETA.DIST', minArgs: 4, maxArgs: 6, evaluate: betaDistFn },
	{ name: 'BETA.INV', minArgs: 3, maxArgs: 5, evaluate: betaInvFn },
	{ name: 'BINOM.DIST', minArgs: 4, maxArgs: 4, evaluate: binomDistFn },
	{ name: 'BINOM.INV', minArgs: 3, maxArgs: 3, evaluate: binomInvFn },
	{ name: 'BINOM.DIST.RANGE', minArgs: 3, maxArgs: 4, evaluate: binomDistRangeFn },
	{ name: 'POISSON.DIST', minArgs: 3, maxArgs: 3, evaluate: poissonDistFn },
	{ name: 'EXPON.DIST', minArgs: 3, maxArgs: 3, evaluate: exponDistFn },
	{ name: 'WEIBULL.DIST', minArgs: 4, maxArgs: 4, evaluate: weibullDistFn },
	{ name: 'GAMMA.DIST', minArgs: 4, maxArgs: 4, evaluate: gammaDistFn },
	{ name: 'GAMMA.INV', minArgs: 3, maxArgs: 3, evaluate: gammaInvFn },
	{ name: 'LOGNORM.DIST', minArgs: 4, maxArgs: 4, evaluate: lognormDistFn },
	{ name: 'LOGNORM.INV', minArgs: 3, maxArgs: 3, evaluate: lognormInvFn },
	{ name: 'NEGBINOM.DIST', minArgs: 4, maxArgs: 4, evaluate: negbinomDistFn },
	{ name: 'HYPGEOM.DIST', minArgs: 5, maxArgs: 5, evaluate: hypgeomDistFn },
	{ name: 'FISHER', minArgs: 1, maxArgs: 1, evaluate: fisherFn },
	{ name: 'FISHERINV', minArgs: 1, maxArgs: 1, evaluate: fisherInvFn },
	{ name: 'STANDARDIZE', minArgs: 3, maxArgs: 3, evaluate: standardizeFn },
	{ name: 'PHI', minArgs: 1, maxArgs: 1, evaluate: phiFn },
	{ name: 'GAUSS', minArgs: 1, maxArgs: 1, evaluate: gaussFn },
	{ name: 'CONFIDENCE.NORM', minArgs: 3, maxArgs: 3, evaluate: confidenceNormFn },
	{ name: 'CONFIDENCE.T', minArgs: 3, maxArgs: 3, evaluate: confidenceTFn },
	{ name: 'SKEW.P', minArgs: 1, maxArgs: 255, evaluate: skewPFn },
	{ name: 'STDEVA', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevAFn(args, false) },
	{ name: 'STDEVPA', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevAFn(args, true) },
	{ name: 'VARA', minArgs: 1, maxArgs: 255, evaluate: (args) => varAFn(args, false) },
	{ name: 'VARPA', minArgs: 1, maxArgs: 255, evaluate: (args) => varAFn(args, true) },
	{ name: 'LINEST', minArgs: 1, maxArgs: 4, evaluate: linestFn },
]
