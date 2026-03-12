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
]
