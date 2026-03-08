import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue } from '@ascend/schema'
import { collectNumbers, type EvalArg, numArg, registerFunction } from './registry.ts'

function collectFrom(arg: EvalArg | undefined): number[] | CellValue {
	if (!arg) return []
	return collectNumbers([arg])
}

function largeFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const ki = Math.floor(k)
	if (ki < 1 || ki > numsOrErr.length) return errorValue('#NUM!')
	numsOrErr.sort((a, b) => b - a)
	return numberValue(numsOrErr[ki - 1] ?? 0)
}

function smallFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectFrom(args[0])
	if (!Array.isArray(numsOrErr)) return numsOrErr
	const k = numArg(args[1])
	if (typeof k !== 'number') return k
	const ki = Math.floor(k)
	if (ki < 1 || ki > numsOrErr.length) return errorValue('#NUM!')
	numsOrErr.sort((a, b) => a - b)
	return numberValue(numsOrErr[ki - 1] ?? 0)
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

function stdevFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length < 2) return errorValue('#DIV/0!')

	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(Math.sqrt(sumSq / (numsOrErr.length - 1)))
}

function varFn(args: EvalArg[]): CellValue {
	const numsOrErr = collectNumbers(args)
	if (!Array.isArray(numsOrErr)) return numsOrErr
	if (numsOrErr.length < 2) return errorValue('#DIV/0!')

	const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
	const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
	return numberValue(sumSq / (numsOrErr.length - 1))
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
registerFunction({
	name: 'STDEV',
	minArgs: 1,
	maxArgs: 255,
	evaluate: stdevFn,
})
registerFunction({ name: 'VAR', minArgs: 1, maxArgs: 255, evaluate: varFn })
