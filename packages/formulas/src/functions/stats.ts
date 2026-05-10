import type { CellValue } from '@ascend/schema'
import { arrayValue, EMPTY, errorValue, numberValue, topLeftScalar } from '@ascend/schema'
import type { FunctionDef, FunctionEvalContext } from './registry.ts'
import { collectNumbers, type EvalArg, getRange, numArg } from './registry.ts'

function collectFrom(arg: EvalArg | undefined): number[] | CellValue {
	if (!arg) return []
	return collectNumbers([arg])
}

function optionalNumArg(args: EvalArg[], index: number, defaultValue: number): number | CellValue {
	const arg = args[index]
	if (!arg || topLeftScalar(arg.value).kind === 'empty') return defaultValue
	return numArg(arg)
}

function numericRangeCacheKey(arg: EvalArg | undefined): string | null {
	const ref = arg?.ref
	if (!ref || ref.kind !== 'range') return null
	return `NUMSORT:${ref.sheetIndex}:${ref.row}:${ref.col}:${ref.endRow ?? ref.row}:${ref.endCol ?? ref.col}`
}

function sortedNumericRange(
	arg: EvalArg | undefined,
	ctx: FunctionEvalContext | undefined,
): number[] | CellValue {
	const key = ctx?.numericVectorCache ? numericRangeCacheKey(arg) : null
	if (key && ctx?.numericVectorCache) {
		const cached = ctx.numericVectorCache.get(key)
		if (cached) return cached
	}
	const numsOrErr = collectFrom(arg)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	numsOrErr.sort((a, b) => a - b)
	if (key && ctx?.numericVectorCache) ctx.numericVectorCache.set(key, numsOrErr)
	return numsOrErr
}

function largeFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const numsOrErr = sortedNumericRange(args[0], ctx)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const ki = Math.floor(k)
	if (ki < 1 || ki > numsOrErr.length) return errorValue('#NUM!')
	return numberValue(numsOrErr[numsOrErr.length - ki] as number)
}

function smallFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const numsOrErr = sortedNumericRange(args[0], ctx)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const ki = Math.floor(k)
	if (ki < 1 || ki > numsOrErr.length) return errorValue('#NUM!')
	return numberValue(numsOrErr[ki - 1] as number)
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

function collectStatNumbers(args: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of args) {
		if (
			arg.kind === 'range' ||
			arg.value.kind === 'array' ||
			arg.areas?.length ||
			arg.forEachValue
		) {
			const collected = collectNumbers([arg])
			if (!Array.isArray(collected)) return collected
			nums.push(...collected)
			continue
		}
		const n = directStatNumber(arg)
		if (typeof n === 'number') nums.push(n)
		else if (n !== null) return n
	}
	return nums
}

function directStatNumber(arg: EvalArg): number | null | CellValue {
	if (arg.ref) {
		const value = arg.value ?? EMPTY
		if (value.kind === 'error') return value
		return value.kind === 'number' ? value.value : value.kind === 'date' ? value.serial : null
	}
	return numArg(arg)
}

function stdevFn(args: EvalArg[], population = false): CellValue {
	const numsOrErr = collectStatNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const variance = varianceValue(numsOrErr, population)
	return typeof variance === 'number' ? numberValue(Math.sqrt(variance)) : variance
}

function varFn(args: EvalArg[], population = false): CellValue {
	const numsOrErr = collectStatNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const variance = varianceValue(numsOrErr, population)
	return typeof variance === 'number' ? numberValue(variance) : variance
}

function varianceValue(nums: readonly number[], population: boolean): number | CellValue {
	const divisor = population ? nums.length : nums.length - 1
	if (divisor < 1) return errorValue('#DIV/0!')

	let mean = 0
	let m2 = 0
	let count = 0
	for (const value of nums) {
		count++
		const delta = value - mean
		mean += delta / count
		m2 += delta * (value - mean)
	}
	return m2 / divisor
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
	const numsOrErr = collectModeNumbers(args)
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
	const modeVal = numsOrErr.find((n) => freq.get(n) === maxCount)
	return numberValue(modeVal ?? 0)
}

function collectModeNumbers(args: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of args) {
		if (arg.kind === 'range' || arg.areas?.length || arg.forEachValue) {
			const collected = collectNumbers([arg])
			if (!Array.isArray(collected)) return collected
			nums.push(...collected)
			continue
		}
		if (arg.value.kind === 'array') {
			for (const row of arg.value.rows) {
				for (const cell of row) {
					const scalar = topLeftScalar(cell)
					if (scalar.kind === 'error') return scalar
					if (scalar.kind === 'number') nums.push(scalar.value)
					else if (scalar.kind === 'date') nums.push(scalar.serial)
				}
			}
			continue
		}
		const scalar = topLeftScalar(arg.value ?? EMPTY)
		if (scalar.kind === 'error') return scalar
		if (scalar.kind === 'number') nums.push(scalar.value)
		else if (scalar.kind === 'date') nums.push(scalar.serial)
		else if (scalar.kind === 'empty' && arg.ref?.kind === 'cell') return errorValue('#VALUE!')
		else if (scalar.kind !== 'empty') return errorValue('#VALUE!')
	}
	return nums
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
			const value = arg.value ?? EMPTY
			if (value.kind === 'error') return value
			const n = arg.ref ? toReferencedANumber(value) : numArg(arg)
			if (n === null) continue
			if (typeof n !== 'number') return n
			nums.push(n)
		}
	}
	return nums
}

function toReferencedANumber(value: CellValue): number | null {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'number':
			return scalar.value
		case 'date':
			return scalar.serial
		case 'boolean':
			return scalar.value ? 1 : 0
		case 'string':
			return 0
		default:
			return null
	}
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
	if (arg1?.kind !== 'range' && arg1?.value.kind === 'error') return arg1.value
	if (arg2?.kind !== 'range' && arg2?.value.kind === 'error') return arg2.value
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

interface CompensatedSum {
	sum: number
	correction: number
}

function addCompensated(total: CompensatedSum, value: number): void {
	const adjusted = value - total.correction
	const next = total.sum + adjusted
	total.correction = next - total.sum - adjusted
	total.sum = next
}

function linregSums(ys: number[], xs: number[]) {
	const n = xs.length
	const x = { sum: 0, correction: 0 }
	const y = { sum: 0, correction: 0 }
	const xy = { sum: 0, correction: 0 }
	const x2 = { sum: 0, correction: 0 }
	for (let i = 0; i < n; i++) {
		const xi = xs[i] as number
		const yi = ys[i] as number
		addCompensated(x, xi)
		addCompensated(y, yi)
		addCompensated(xy, xi * yi)
		addCompensated(x2, xi * xi)
	}
	const sumX = x.sum
	const sumY = y.sum
	const sumX2 = x2.sum
	const sumXY = xy.sum
	const meanX = sumX / n
	const meanY = sumY / n
	const centeredX2 = { sum: 0, correction: 0 }
	const centeredY2 = { sum: 0, correction: 0 }
	const centeredXY = { sum: 0, correction: 0 }
	for (let i = 0; i < n; i++) {
		const dx = (xs[i] as number) - meanX
		const dy = (ys[i] as number) - meanY
		addCompensated(centeredX2, dx * dx)
		addCompensated(centeredY2, dy * dy)
		addCompensated(centeredXY, dx * dy)
	}
	const ssxx = centeredX2.sum
	const ssyy = centeredY2.sum
	const ssxy = centeredXY.sum
	return { n, sumX, sumY, sumX2, sumXY, ssxx, ssyy, ssxy }
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
	const numsOrErr = collectModeNumbers(args)
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
const NORMAL_INV_LOW = 0.02425
const NORMAL_INV_HIGH = 1 - NORMAL_INV_LOW
const NORMAL_INV_A = [
	-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2,
	-3.066479806614716e1, 2.506628277459239,
]
const NORMAL_INV_B = [
	-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1,
	-1.328068155288572e1, 1,
]
const NORMAL_INV_C = [
	-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734,
	4.374664141464968, 2.938163982698783,
]
const NORMAL_INV_D = [
	7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416, 1,
]
const NORMAL_SQRT_HALF = Math.SQRT1_2
const NORMAL_MAX_LOG = 7.097827128933839e2
const NORMAL_ERFC_P = [
	2.461969814735305e-10, 5.641895648310687e-1, 7.463210564422699, 4.863719709856814e1,
	1.965208329560771e2, 5.264451949954774e2, 9.345285271719576e2, 1.027551886895157e3,
	5.575353353693993e2,
]
const NORMAL_ERFC_Q = [
	1.32281951154745e1, 8.670721408859897e1, 3.549377788878199e2, 9.757085017432055e2,
	1.8239091668790975e3, 2.24633760818711e3, 1.6566630919416134e3, 5.575353408177277e2,
]
const NORMAL_ERFC_R = [
	5.641895835477551e-1, 1.275366707599781, 5.019050422511805, 6.160210979930536, 7.409742699504489,
	2.9788666537210022,
]
const NORMAL_ERFC_S = [
	2.2605286322011726, 9.396035249380015, 1.2048953980809666e1, 1.708144507475659e1,
	9.608968090632859, 3.369076451000815,
]
const NORMAL_ERF_T = [
	9.604973739870516, 9.002601972038427e1, 2.232005345946843e3, 7.003325141128051e3,
	5.55923013010395e4,
]
const NORMAL_ERF_U = [
	3.356171416475031e1, 5.213579497801527e2, 4.594323829709801e3, 2.262900006138909e4,
	4.926739426086359e4,
]

function normPdf(z: number): number {
	return Math.exp(-0.5 * z * z) / SQRT_2PI
}

function evalPolynomial(coefficients: readonly number[], x: number): number {
	let result = 0
	for (const coefficient of coefficients) result = result * x + coefficient
	return result
}

function evalMonicPolynomial(coefficients: readonly number[], x: number): number {
	let result = x + (coefficients[0] ?? 0)
	for (let i = 1; i < coefficients.length; i++) result = result * x + (coefficients[i] ?? 0)
	return result
}

function erfApprox(x: number): number {
	if (Math.abs(x) > 1) return 1 - erfcApprox(x)
	const z = x * x
	return (x * evalPolynomial(NORMAL_ERF_T, z)) / evalMonicPolynomial(NORMAL_ERF_U, z)
}

function erfcApprox(x: number): number {
	const ax = Math.abs(x)
	if (ax < 1) return 1 - erfApprox(x)
	const z = -x * x
	if (z < -NORMAL_MAX_LOG) return x < 0 ? 2 : 0
	const exp = Math.exp(z)
	const numerator = ax < 8 ? evalPolynomial(NORMAL_ERFC_P, ax) : evalPolynomial(NORMAL_ERFC_R, ax)
	const denominator =
		ax < 8 ? evalMonicPolynomial(NORMAL_ERFC_Q, ax) : evalMonicPolynomial(NORMAL_ERFC_S, ax)
	const y = (exp * numerator) / denominator
	return x < 0 ? 2 - y : y === 0 ? 0 : y
}

function normCdf(z: number): number {
	if (z === 0) return 0.5
	const x = z * NORMAL_SQRT_HALF
	const ax = Math.abs(x)
	if (ax < 1) return 0.5 + 0.5 * erfApprox(x)
	const tail = 0.5 * erfcApprox(ax)
	return x > 0 ? 1 - tail : tail
}

function normSInvImpl(p: number): number {
	let z: number
	if (p < NORMAL_INV_LOW) {
		const q = Math.sqrt(-2 * Math.log(p))
		z = evalPolynomial(NORMAL_INV_C, q) / evalPolynomial(NORMAL_INV_D, q)
	} else if (p <= NORMAL_INV_HIGH) {
		const q = p - 0.5
		const r = q * q
		z = (evalPolynomial(NORMAL_INV_A, r) * q) / evalPolynomial(NORMAL_INV_B, r)
	} else {
		const q = Math.sqrt(-2 * Math.log(1 - p))
		z = -evalPolynomial(NORMAL_INV_C, q) / evalPolynomial(NORMAL_INV_D, q)
	}
	for (let i = 0; i < 2; i++) {
		const error = normCdf(z) - p
		z -= error / normPdf(z)
	}
	const rounded = Math.round(z)
	if (Math.abs(z - rounded) <= 1e-11 && Math.abs(normCdf(rounded) - p) <= 5e-15) return rounded
	return z
}

const LANCZOS_C = [
	0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
	-176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
	1.5056327351493116e-7,
]
const LOG_FACTORIALS = (() => {
	const values = [0]
	let sum = 0
	for (let n = 1; n <= 170; n++) {
		sum += Math.log(n)
		values[n] = sum
	}
	return values
})()

function gammalnImpl(x: number): number {
	if (x <= 0 && x === Math.floor(x)) return Number.POSITIVE_INFINITY
	if (Number.isInteger(x) && x > 0 && x <= LOG_FACTORIALS.length) {
		return LOG_FACTORIALS[x - 1] ?? 0
	}
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
	const EPS = 3e-14
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
		if (Math.abs(del - 1) < EPS) break
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

function ibetaComplement(x: number, a: number, b: number): number {
	if (x <= 0) return 1
	if (x >= 1) return 0
	const bt = Math.exp(
		gammalnImpl(a + b) - gammalnImpl(a) - gammalnImpl(b) + a * Math.log(x) + b * Math.log(1 - x),
	)
	if (x < (a + 1) / (a + b + 2)) return 1 - (bt * betacf(x, a, b)) / a
	return (bt * betacf(1 - x, b, a)) / b
}

function ibetainv(p: number, a: number, b: number): number {
	if (p <= 0) return 0
	if (p >= 1) return 1
	let lo = 0
	let hi = 1
	for (let i = 0; i < 120; i++) {
		const x = (lo + hi) / 2
		const fx = ibeta(x, a, b) - p
		if (fx < 0) lo = x
		else hi = x
		if (hi - lo < 1e-14) break
	}
	return (lo + hi) / 2
}

function gammaIterationLimit(a: number): number {
	return Math.ceil(Math.log(a >= 1 ? a : 1 / a) * 8.5 + a * 0.4 + 17)
}

function lowRegGammaSeries(a: number, x: number, aln: number): number {
	let ap = a
	let sum = 1 / a
	let del = sum
	for (let i = 1; i <= gammaIterationLimit(a); i++) {
		ap += 1
		del *= x / ap
		sum += del
	}
	return sum * Math.exp(-x + a * Math.log(x) - aln)
}

function upperRegGammaFraction(a: number, x: number, aln: number): number {
	const maxIter = gammaIterationLimit(a)
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
	return h * Math.exp(-x + a * Math.log(x) - aln)
}

function lowRegGamma(a: number, x: number): number {
	if (x <= 0) return 0
	if (a <= 0) return 1
	const aln = gammalnImpl(a)
	if (x < a + 1) return lowRegGammaSeries(a, x, aln)
	return 1 - upperRegGammaFraction(a, x, aln)
}

function upperRegGamma(a: number, x: number): number {
	if (x <= 0) return 1
	if (a <= 0) return 0
	const aln = gammalnImpl(a)
	if (x < a + 1) return 1 - lowRegGammaSeries(a, x, aln)
	return upperRegGammaFraction(a, x, aln)
}

function combiln(n: number, k: number): number {
	return gammalnImpl(n + 1) - gammalnImpl(k + 1) - gammalnImpl(n - k + 1)
}

function combinationNumber(n: number, k: number): number {
	if (k < 0 || k > n) return 0
	const kk = Math.min(k, n - k)
	let out = 1
	for (let i = 1; i <= kk; i++) out = (out * (n - kk + i)) / i
	return out
}

function sumKahan(values: Iterable<number>): number {
	let sum = 0
	let compensation = 0
	for (const value of values) {
		const y = value - compensation
		const t = sum + y
		compensation = t - sum - y
		sum = t
	}
	return sum
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
	const df1 = Math.trunc(d1)
	const df2 = Math.trunc(d2)
	if (x < 0 || df1 < 1 || df2 < 1) return errorValue('#NUM!')
	if (x === 0) return cum !== 0 ? numberValue(0) : errorValue('#NUM!')
	if (cum !== 0) {
		const w = (df1 * x) / (df1 * x + df2)
		return numberValue(ibeta(w, df1 / 2, df2 / 2))
	}
	return numberValue(
		Math.sqrt(((df1 * x) ** df1 * df2 ** df2) / (df1 * x + df2) ** (df1 + df2)) /
			(x * Math.exp(betaln(df1 / 2, df2 / 2))),
	)
}

function fDistRTFn(args: EvalArg[]): CellValue {
	const x = numArg(args[0])
	if (typeof x !== 'number') return x
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	const df1 = Math.trunc(d1)
	const df2 = Math.trunc(d2)
	if (x < 0 || df1 < 1 || df2 < 1) return errorValue('#NUM!')
	if (x === 0) return numberValue(1)
	const w = (df1 * x) / (df1 * x + df2)
	return numberValue(ibetaComplement(w, df1 / 2, df2 / 2))
}

function snapFInvRoundTrip(
	result: number,
	p: number,
	d1: number,
	d2: number,
	rightTail: boolean,
): number {
	const rounded = Math.round(result)
	if (Math.abs(result - rounded) > 1e-12) return result
	const w = (d1 * rounded) / (d1 * rounded + d2)
	const roundTrip = ibeta(w, d1 / 2, d2 / 2)
	const target = rightTail ? 1 - roundTrip : roundTrip
	return Math.abs(target - p) <= 5e-15 ? rounded : result
}

function fInvFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	const df1 = Math.trunc(d1)
	const df2 = Math.trunc(d2)
	if (p < 0 || p > 1 || df1 < 1 || df2 < 1) return errorValue('#NUM!')
	if (p === 0) return numberValue(0)
	if (p === 1) return errorValue('#NUM!')
	const x = ibetainv(p, df1 / 2, df2 / 2)
	const result = (x * df2) / (df1 * (1 - x))
	return numberValue(snapFInvRoundTrip(result, p, df1, df2, false))
}

function fInvRTFn(args: EvalArg[]): CellValue {
	const p = numArg(args[0])
	if (typeof p !== 'number') return p
	const d1 = numArg(args[1])
	if (typeof d1 !== 'number') return d1
	const d2 = numArg(args[2])
	if (typeof d2 !== 'number') return d2
	const df1 = Math.trunc(d1)
	const df2 = Math.trunc(d2)
	if (p < 0 || p > 1 || df1 < 1 || df2 < 1) return errorValue('#NUM!')
	if (p === 1) return numberValue(0)
	if (p === 0) return errorValue('#NUM!')
	const x = ibetainv(1 - p, df1 / 2, df2 / 2)
	const result = (x * df2) / (df1 * (1 - x))
	return numberValue(snapFInvRoundTrip(result, p, df1, df2, true))
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
	return numberValue(upperRegGamma(df / 2, x / 2))
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
	const A = optionalNumArg(args, 4, 0)
	if (typeof A !== 'number') return A
	const B = optionalNumArg(args, 5, 1)
	if (typeof B !== 'number') return B
	if (alpha <= 0 || beta <= 0 || B <= A || x < A || x > B) return errorValue('#NUM!')
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
	const A = optionalNumArg(args, 3, 0)
	if (typeof A !== 'number') return A
	const B = optionalNumArg(args, 4, 1)
	if (typeof B !== 'number') return B
	if (p < 0 || p > 1 || alpha <= 0 || beta <= 0 || B <= A) return errorValue('#NUM!')
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
		if (ni <= 170) return combinationNumber(ni, i) * p ** i * (1 - p) ** (ni - i)
		return Math.exp(combiln(ni, i) + i * Math.log(p) + (ni - i) * Math.log(1 - p))
	}
	if (cum === 0) return numberValue(pmf(k))
	const sum = sumKahan(
		(function* () {
			for (let i = 0; i <= k; i++) yield pmf(i)
		})(),
	)
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
	const sum = sumKahan(
		(function* () {
			for (let i = lo; i <= hi; i++) {
				if (n <= 170) yield combinationNumber(n, i) * p ** i * (1 - p) ** (n - i)
				else yield Math.exp(combiln(n, i) + i * Math.log(p) + (n - i) * Math.log(1 - p))
			}
		})(),
	)
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
	const pmf = (i: number) => {
		if (mean === 0) return i === 0 ? 1 : 0
		if (i <= 170) return (Math.exp(-mean) * mean ** i) / factorialNumber(i)
		return Math.exp(-mean + i * Math.log(mean) - gammalnImpl(i + 1))
	}
	if (cum === 0) return numberValue(pmf(k))
	let sum = 0
	for (let i = 0; i <= k; i++) sum += pmf(i)
	return numberValue(Math.min(1, sum))
}

function factorialNumber(n: number): number {
	let out = 1
	for (let i = 2; i <= n; i++) out *= i
	return out
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
	const useFiniteCombination = N <= 170 && cum === 0
	const pmf = (i: number) => {
		if (useFiniteCombination) {
			return (combinationNumber(K, i) * combinationNumber(N - K, n - i)) / combinationNumber(N, n)
		}
		return Math.exp(combiln(K, i) + combiln(N - K, n - i) - combiln(N, n))
	}
	if (cum === 0) return numberValue(pmf(k))
	const lo = Math.max(0, n - (N - K))
	const sum = sumKahan(
		(function* () {
			for (let i = lo; i <= k; i++) yield pmf(i)
		})(),
	)
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

function probNumber(value: CellValue): number | CellValue {
	const scalar = topLeftScalar(value)
	if (scalar.kind === 'error') return scalar
	if (scalar.kind === 'number') return scalar.value
	if (scalar.kind === 'date') return scalar.serial
	return errorValue('#VALUE!')
}

function probFn(args: EvalArg[]): CellValue {
	const xs = getRange(args[0])
	const probs = getRange(args[1])
	const rows = xs.length
	const cols = xs[0]?.length ?? 0
	if (rows !== probs.length || cols !== (probs[0]?.length ?? 0)) return errorValue('#N/A')

	const lower = numArg(args[2])
	if (typeof lower !== 'number') return lower
	const upper = args.length > 3 ? numArg(args[3]) : lower
	if (typeof upper !== 'number') return upper
	if (upper < lower) return errorValue('#NUM!')

	let probabilitySum = 0
	let result = 0
	for (let r = 0; r < rows; r++) {
		const xRow = xs[r]
		const pRow = probs[r]
		if ((xRow?.length ?? 0) !== cols || (pRow?.length ?? 0) !== cols) return errorValue('#N/A')
		for (let c = 0; c < cols; c++) {
			const x = probNumber(xRow?.[c] ?? errorValue('#VALUE!'))
			if (typeof x !== 'number') return x
			const probability = probNumber(pRow?.[c] ?? errorValue('#VALUE!'))
			if (typeof probability !== 'number') return probability
			if (probability <= 0 || probability > 1) return errorValue('#NUM!')
			probabilitySum += probability
			if (x >= lower && x <= upper) result += probability
		}
	}
	if (Math.abs(probabilitySum - 1) > 1e-7) return errorValue('#NUM!')
	return numberValue(result)
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

function extractNumbersFromRange(range: readonly (readonly CellValue[])[]): number[] | CellValue {
	const nums: number[] = []
	for (let r = 0; r < range.length; r++) {
		const row = range[r]
		const cols = row?.length ?? 0
		for (let c = 0; c < cols; c++) {
			const v = row?.[c]
			if (!v) continue
			if (v.kind === 'error') return v
			const n = v.kind === 'number' ? v.value : v.kind === 'date' ? v.serial : null
			if (n !== null) nums.push(n)
		}
	}
	if (nums.length === 0) return errorValue('#N/A')
	return nums
}

function trendFn(args: EvalArg[]): CellValue {
	const yRange = getRange(args[0])
	const ysOrErr = extractNumbersFromRange(yRange)
	if (!Array.isArray(ysOrErr)) return ysOrErr
	const ys = ysOrErr
	const n = ys.length

	let xs: number[]
	if (args.length > 1 && args[1]?.value !== undefined) {
		const xRange = getRange(args[1])
		const xsOrErr = extractNumbersFromRange(xRange)
		if (!Array.isArray(xsOrErr)) return xsOrErr
		if (xsOrErr.length !== n) return errorValue('#N/A')
		xs = xsOrErr
	} else {
		xs = Array.from({ length: n }, (_, i) => i + 1)
	}

	let useConst = true
	if (args.length > 3) {
		const v = numArg(args[3])
		if (typeof v !== 'number') return v
		useConst = v !== 0
	}

	const { n: nn, sumX, sumY, sumX2, sumXY, ssxx, ssxy } = linregSums(ys, xs)
	const slope = useConst
		? ssxx === 0
			? errorValue('#DIV/0!')
			: ssxy / ssxx
		: sumX2 === 0
			? errorValue('#DIV/0!')
			: sumXY / sumX2
	if (typeof slope !== 'number') return slope
	const intercept = useConst ? (sumY - slope * sumX) / nn : 0

	let newXs: number[]
	if (args.length > 2 && args[2]?.value !== undefined) {
		const newXRange = getRange(args[2])
		const newXsOrErr = extractNumbersFromRange(newXRange)
		if (!Array.isArray(newXsOrErr)) return newXsOrErr
		newXs = newXsOrErr
	} else {
		newXs = xs
	}

	const pred = newXs.map((x) => intercept + slope * x)
	return arrayValue(pred.map((v) => [topLeftScalar(numberValue(v))]))
}

function growthFn(args: EvalArg[]): CellValue {
	const yRange = getRange(args[0])
	const ysOrErr = extractNumbersFromRange(yRange)
	if (!Array.isArray(ysOrErr)) return ysOrErr
	const ys = ysOrErr
	for (const y of ys) {
		if (y <= 0) return errorValue('#NUM!')
	}
	const logYs = ys.map((y) => Math.log(y))
	const n = logYs.length

	let xs: number[]
	if (args.length > 1 && args[1]?.value !== undefined) {
		const xRange = getRange(args[1])
		const xsOrErr = extractNumbersFromRange(xRange)
		if (!Array.isArray(xsOrErr)) return xsOrErr
		if (xsOrErr.length !== n) return errorValue('#N/A')
		xs = xsOrErr
	} else {
		xs = Array.from({ length: n }, (_, i) => i + 1)
	}

	let useConst = true
	if (args.length > 3) {
		const v = numArg(args[3])
		if (typeof v !== 'number') return v
		useConst = v !== 0
	}

	const { n: nn, sumX, sumY, sumX2, sumXY, ssxx, ssxy } = linregSums(logYs, xs)
	const slope = useConst
		? ssxx === 0
			? errorValue('#DIV/0!')
			: ssxy / ssxx
		: sumX2 === 0
			? errorValue('#DIV/0!')
			: sumXY / sumX2
	if (typeof slope !== 'number') return slope
	const intercept = useConst ? (sumY - slope * sumX) / nn : 0
	const m = Math.exp(slope)
	const b = Math.exp(intercept)

	let newXs: number[]
	if (args.length > 2 && args[2]?.value !== undefined) {
		const newXRange = getRange(args[2])
		const newXsOrErr = extractNumbersFromRange(newXRange)
		if (!Array.isArray(newXsOrErr)) return newXsOrErr
		newXs = newXsOrErr
	} else {
		newXs = xs
	}

	const pred = newXs.map((x) => b * m ** x)
	return arrayValue(pred.map((v) => [topLeftScalar(numberValue(v))]))
}

function logestFn(args: EvalArg[]): CellValue {
	const yRange = getRange(args[0])
	const ysOrErr = extractNumbersFromRange(yRange)
	if (!Array.isArray(ysOrErr)) return ysOrErr
	const ys = ysOrErr
	for (const y of ys) {
		if (y <= 0) return errorValue('#NUM!')
	}
	const logYs = ys.map((y) => Math.log(y))
	const n = logYs.length

	let xs: number[]
	if (args.length > 1 && args[1]?.value !== undefined) {
		const xRange = getRange(args[1])
		const xsOrErr = extractNumbersFromRange(xRange)
		if (!Array.isArray(xsOrErr)) return xsOrErr
		if (xsOrErr.length !== n) return errorValue('#N/A')
		xs = xsOrErr
	} else {
		xs = Array.from({ length: n }, (_, i) => i + 1)
	}

	let useConst = true
	if (args.length > 2) {
		const v = numArg(args[2])
		if (typeof v !== 'number') return v
		useConst = v !== 0
	}

	const { n: nn, sumX, sumY, sumX2, sumXY, ssxx, ssxy } = linregSums(logYs, xs)
	const slope = useConst
		? ssxx === 0
			? errorValue('#DIV/0!')
			: ssxy / ssxx
		: sumX2 === 0
			? errorValue('#DIV/0!')
			: sumXY / sumX2
	if (typeof slope !== 'number') return slope
	const intercept = useConst ? (sumY - slope * sumX) / nn : 0
	const m = Math.exp(slope)
	const b = Math.exp(intercept)

	return arrayValue([[topLeftScalar(numberValue(m)), topLeftScalar(numberValue(b))]])
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

function sampleMeanAndVar(arr: number[]): { mean: number; var: number; n: number } {
	const n = arr.length
	const mean = arr.reduce((a, b) => a + b, 0) / n
	const sumSq = arr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	const v = n > 1 ? sumSq / (n - 1) : 0
	return { mean, var: v, n }
}

function tTestPValue(t: number, df: number, tails: number): number {
	if (df < 1) return NaN
	const x = df / (df + t * t)
	const p = ibeta(x, df / 2, 0.5)
	return tails === 1 ? 0.5 * p : p
}

function tTestFn(args: EvalArg[]): CellValue {
	const tails = numArg(args[2])
	if (typeof tails !== 'number') return tails
	const typeArg = numArg(args[3])
	if (typeof typeArg !== 'number') return typeArg
	const tailsInt = Math.trunc(tails)
	const typeInt = Math.trunc(typeArg)
	if (tailsInt !== 1 && tailsInt !== 2) return errorValue('#NUM!')
	if (typeInt < 1 || typeInt > 3) return errorValue('#NUM!')

	let t: number
	let df: number

	if (typeInt === 1) {
		const paired = collectPaired(args[0], args[1])
		if (!Array.isArray(paired)) return paired
		const [a1, a2] = paired
		const diffs: number[] = []
		for (let i = 0; i < a1.length; i++) diffs.push((a1[i] as number) - (a2[i] as number))
		const n = diffs.length
		if (n < 2) return errorValue('#DIV/0!')
		const mean = diffs.reduce((a, b) => a + b, 0) / n
		const sumSq = diffs.reduce((acc, v) => acc + (v - mean) ** 2, 0)
		const se = Math.sqrt(sumSq / (n * (n - 1)))
		if (se === 0) return errorValue('#DIV/0!')
		t = mean / se
		df = n - 1
	} else {
		const a1 = collectFrom(args[0])
		if (!Array.isArray(a1)) return a1
		const a2 = collectFrom(args[1])
		if (!Array.isArray(a2)) return a2
		if (a1.length < 2 || a2.length < 2) return errorValue('#DIV/0!')
		const s1 = sampleMeanAndVar(a1)
		const s2 = sampleMeanAndVar(a2)
		if (typeInt === 2) {
			const pooled = ((s1.n - 1) * s1.var + (s2.n - 1) * s2.var) / (s1.n + s2.n - 2)
			if (pooled === 0) return errorValue('#DIV/0!')
			const se = Math.sqrt(pooled * (1 / s1.n + 1 / s2.n))
			t = (s1.mean - s2.mean) / se
			df = s1.n + s2.n - 2
		} else {
			const se2 = s1.var / s1.n + s2.var / s2.n
			if (se2 === 0) return errorValue('#DIV/0!')
			const se = Math.sqrt(se2)
			t = (s1.mean - s2.mean) / se
			const v1 = s1.var / s1.n
			const v2 = s2.var / s2.n
			df = (v1 + v2) ** 2 / (v1 ** 2 / (s1.n - 1) + v2 ** 2 / (s2.n - 1))
		}
	}

	const p = tTestPValue(Math.abs(t), df, tailsInt)
	if (Number.isNaN(p)) return errorValue('#NUM!')
	return numberValue(p)
}

function fTestFn(args: EvalArg[]): CellValue {
	const a1 = collectFrom(args[0])
	if (!Array.isArray(a1)) return a1
	const a2 = collectFrom(args[1])
	if (!Array.isArray(a2)) return a2
	if (a1.length < 2 || a2.length < 2) return errorValue('#DIV/0!')
	const s1 = sampleMeanAndVar(a1)
	const s2 = sampleMeanAndVar(a2)
	if (s1.var === 0 || s2.var === 0) return errorValue('#DIV/0!')
	const df1 = s1.n - 1
	const df2 = s2.n - 1
	const F = s1.var >= s2.var ? s1.var / s2.var : s2.var / s1.var
	const d1 = s1.var >= s2.var ? df1 : df2
	const d2 = s1.var >= s2.var ? df2 : df1
	const w = (d1 * F) / (d1 * F + d2)
	const pOneTail = ibetaComplement(w, d1 / 2, d2 / 2)
	return numberValue(Math.min(1, 2 * pOneTail))
}

function chisqTestFn(args: EvalArg[]): CellValue {
	const actual = collectFrom(args[0])
	if (!Array.isArray(actual)) return actual
	const expected = collectFrom(args[1])
	if (!Array.isArray(expected)) return expected
	if (actual.length !== expected.length) return errorValue('#N/A')
	if (actual.length < 2) return errorValue('#NUM!')
	let chi2 = 0
	for (let i = 0; i < actual.length; i++) {
		const a = actual[i] as number
		const e = expected[i] as number
		if (e === 0) return errorValue('#DIV/0!')
		chi2 += (a - e) ** 2 / e
	}
	const df = actual.length - 1
	const p = upperRegGamma(df / 2, chi2 / 2)
	return numberValue(p)
}

function zTestFn(args: EvalArg[]): CellValue {
	const arr = collectFrom(args[0])
	if (!Array.isArray(arr)) return arr
	const x = numArg(args[1])
	if (typeof x !== 'number') return x
	if (arr.length === 0) return errorValue('#DIV/0!')
	const mean = arr.reduce((a, b) => a + b, 0) / arr.length
	const n = arr.length
	let se: number
	if (args.length > 2 && args[2]?.value !== undefined) {
		const sigma = numArg(args[2])
		if (typeof sigma !== 'number') return sigma
		if (sigma <= 0) return errorValue('#NUM!')
		se = sigma / Math.sqrt(n)
	} else {
		if (n < 2) return errorValue('#DIV/0!')
		const sumSq = arr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
		const stdev = Math.sqrt(sumSq / (n - 1))
		if (stdev === 0) return errorValue('#DIV/0!')
		se = stdev / Math.sqrt(n)
	}
	const z = (mean - x) / se
	return numberValue(1 - normCdf(z))
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
	{ name: 'PROB', minArgs: 3, maxArgs: 4, evaluate: probFn },
	{ name: 'SKEW.P', minArgs: 1, maxArgs: 255, evaluate: skewPFn },
	{ name: 'STDEVA', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevAFn(args, false) },
	{ name: 'STDEVPA', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevAFn(args, true) },
	{ name: 'VARA', minArgs: 1, maxArgs: 255, evaluate: (args) => varAFn(args, false) },
	{ name: 'VARPA', minArgs: 1, maxArgs: 255, evaluate: (args) => varAFn(args, true) },
	{ name: 'LINEST', minArgs: 1, maxArgs: 4, evaluate: linestFn },
	{ name: 'TREND', minArgs: 1, maxArgs: 4, evaluate: trendFn },
	{ name: 'GROWTH', minArgs: 1, maxArgs: 4, evaluate: growthFn },
	{ name: 'LOGEST', minArgs: 1, maxArgs: 4, evaluate: logestFn },
	{ name: 'T.TEST', minArgs: 4, maxArgs: 4, evaluate: tTestFn },
	{ name: 'TTEST', minArgs: 4, maxArgs: 4, evaluate: tTestFn },
	{ name: 'F.TEST', minArgs: 2, maxArgs: 2, evaluate: fTestFn },
	{ name: 'FTEST', minArgs: 2, maxArgs: 2, evaluate: fTestFn },
	{ name: 'CHISQ.TEST', minArgs: 2, maxArgs: 2, evaluate: chisqTestFn },
	{ name: 'CHITEST', minArgs: 2, maxArgs: 2, evaluate: chisqTestFn },
	{ name: 'Z.TEST', minArgs: 2, maxArgs: 3, evaluate: zTestFn },
	{ name: 'ZTEST', minArgs: 2, maxArgs: 3, evaluate: zTestFn },
]
