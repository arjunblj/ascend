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

	if (!numsOrErr.some((v) => v === num)) return errorValue('#N/A')
	const ascending = order !== 0
	let rank = 1
	for (const v of numsOrErr) {
		if (ascending ? v < num : v > num) rank++
	}
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
