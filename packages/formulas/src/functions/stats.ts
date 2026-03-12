import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue } from '@ascend/schema'
import { collectNumbers, type EvalArg, numArg, registerFunction } from './registry.ts'

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

// --- Registration ---

registerFunction({ name: 'LARGE', minArgs: 2, maxArgs: 2, evaluate: largeFn })
registerFunction({ name: 'SMALL', minArgs: 2, maxArgs: 2, evaluate: smallFn })
registerFunction({ name: 'RANK', minArgs: 2, maxArgs: 3, evaluate: rankFn })
registerFunction({
	name: 'PERCENTILE',
	minArgs: 2,
	maxArgs: 2,
	evaluate: percentileFn,
})
registerFunction({
	name: 'MEDIAN',
	minArgs: 1,
	maxArgs: 255,
	evaluate: medianFn,
})
registerFunction({ name: 'STDEV', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevFn(args) })
registerFunction({ name: 'STDEV.S', minArgs: 1, maxArgs: 255, evaluate: (args) => stdevFn(args) })
registerFunction({
	name: 'STDEV.P',
	minArgs: 1,
	maxArgs: 255,
	evaluate: (args) => stdevFn(args, true),
})
registerFunction({
	name: 'STDEVP',
	minArgs: 1,
	maxArgs: 255,
	evaluate: (args) => stdevFn(args, true),
})
registerFunction({ name: 'VAR', minArgs: 1, maxArgs: 255, evaluate: (args) => varFn(args) })
registerFunction({ name: 'VAR.S', minArgs: 1, maxArgs: 255, evaluate: (args) => varFn(args) })
registerFunction({
	name: 'VAR.P',
	minArgs: 1,
	maxArgs: 255,
	evaluate: (args) => varFn(args, true),
})
registerFunction({
	name: 'VARP',
	minArgs: 1,
	maxArgs: 255,
	evaluate: (args) => varFn(args, true),
})
registerFunction({
	name: 'PERCENTILE.INC',
	minArgs: 2,
	maxArgs: 2,
	evaluate: percentileFn,
})
registerFunction({
	name: 'PERCENTILE.EXC',
	minArgs: 2,
	maxArgs: 2,
	evaluate: percentileExcFn,
})
registerFunction({
	name: 'QUARTILE',
	minArgs: 2,
	maxArgs: 2,
	evaluate: (args) => quartileFn(args),
})
registerFunction({
	name: 'QUARTILE.INC',
	minArgs: 2,
	maxArgs: 2,
	evaluate: (args) => quartileFn(args),
})
registerFunction({
	name: 'QUARTILE.EXC',
	minArgs: 2,
	maxArgs: 2,
	evaluate: (args) => quartileFn(args, true),
})
registerFunction({ name: 'MODE', minArgs: 1, maxArgs: 255, evaluate: modeFn })
registerFunction({ name: 'MODE.SNGL', minArgs: 1, maxArgs: 255, evaluate: modeFn })
registerFunction({ name: 'AVERAGEA', minArgs: 1, maxArgs: 255, evaluate: averageaFn })
registerFunction({ name: 'MAXA', minArgs: 1, maxArgs: 255, evaluate: maxaFn })
registerFunction({ name: 'MINA', minArgs: 1, maxArgs: 255, evaluate: minaFn })
registerFunction({ name: 'RANK.EQ', minArgs: 2, maxArgs: 3, evaluate: rankFn })
registerFunction({ name: 'RANK.AVG', minArgs: 2, maxArgs: 3, evaluate: rankAvgFn })
registerFunction({ name: 'GEOMEAN', minArgs: 1, maxArgs: 255, evaluate: geomeanFn })
registerFunction({ name: 'HARMEAN', minArgs: 1, maxArgs: 255, evaluate: harmeanFn })
registerFunction({ name: 'TRIMMEAN', minArgs: 2, maxArgs: 2, evaluate: trimmeanFn })
registerFunction({
	name: 'PERCENTRANK.INC',
	minArgs: 2,
	maxArgs: 3,
	evaluate: percentrankIncFn,
})
registerFunction({
	name: 'PERCENTRANK.EXC',
	minArgs: 2,
	maxArgs: 3,
	evaluate: percentrankExcFn,
})
