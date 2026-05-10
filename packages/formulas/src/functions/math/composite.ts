import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, isEmpty, isError, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef } from '../index.ts'
import { fn, numArg, numericVal, toNum } from './helpers.ts'

function aggregateCollectNumbers(args: EvalArg[], ignoreErrors: boolean): number[] | CellValue {
	const nums: number[] = []
	for (const arg of args) {
		if (arg.forEachValue) {
			let err: CellValue | undefined
			arg.forEachValue((cell) => {
				if (err && !ignoreErrors) return
				if (cell.kind === 'error') {
					if (!ignoreErrors) err = cell
					return
				}
				const n = numericVal(cell)
				if (n !== null) nums.push(n)
			})
			if (err) return err
		} else if (arg.kind === 'range' && arg.values) {
			for (const row of arg.values) {
				for (const cell of row) {
					if (cell.kind === 'error') {
						if (!ignoreErrors) return cell
						continue
					}
					const n = numericVal(cell)
					if (n !== null) nums.push(n)
				}
			}
		} else {
			const v = arg.value ?? EMPTY
			if (v.kind === 'error') {
				if (!ignoreErrors) return v
				continue
			}
			const n = toNum(v)
			if (typeof n === 'number') nums.push(n)
		}
	}
	return nums
}

function aggregateFn(args: EvalArg[]): CellValue {
	const fnNum = numArg(args[0])
	if (typeof fnNum !== 'number') return fnNum
	const code = Math.trunc(fnNum)
	const opt = numArg(args[1])
	if (typeof opt !== 'number') return opt
	const options = Math.trunc(opt)
	const ignoreErrors = options === 2 || options === 3 || options === 6 || options === 7

	if (code >= 1 && code <= 11) {
		const data = args.slice(2)
		const numsOrErr = aggregateCollectNumbers(data, ignoreErrors)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		switch (code) {
			case 1: {
				if (numsOrErr.length === 0) return errorValue('#DIV/0!')
				const sum = numsOrErr.reduce((a, b) => a + b, 0)
				return numberValue(sum / numsOrErr.length)
			}
			case 2: {
				return numberValue(numsOrErr.length)
			}
			case 3: {
				const dataAll = args.slice(2)
				let count = 0
				for (const arg of dataAll) {
					if (arg.forEachValue) {
						let err: CellValue | undefined
						arg.forEachValue((cell) => {
							if (cell.kind === 'error') {
								if (!ignoreErrors) err = cell
								return
							}
							if (!isEmpty(cell)) count++
						})
						if (err) return err
					} else if (arg.kind === 'range' && arg.values) {
						for (const row of arg.values) {
							for (const cell of row) {
								if (cell.kind === 'error') {
									if (!ignoreErrors) return cell
									continue
								}
								if (!isEmpty(cell)) count++
							}
						}
					} else if (!isEmpty(arg.value ?? EMPTY)) count++
				}
				return numberValue(count)
			}
			case 4:
				return numberValue(numsOrErr.length === 0 ? 0 : Math.max(...numsOrErr))
			case 5:
				return numberValue(numsOrErr.length === 0 ? 0 : Math.min(...numsOrErr))
			case 6:
				return numberValue(numsOrErr.length === 0 ? 0 : numsOrErr.reduce((a, b) => a * b, 1))
			case 7: {
				if (numsOrErr.length < 2) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
				return numberValue(Math.sqrt(sumSq / (numsOrErr.length - 1)))
			}
			case 8: {
				if (numsOrErr.length < 2) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				const sumSq = numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0)
				return numberValue(sumSq / (numsOrErr.length - 1))
			}
			case 9:
				return numberValue(numsOrErr.reduce((a, b) => a + b, 0))
			case 10: {
				if (numsOrErr.length < 2) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				return numberValue(
					numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / (numsOrErr.length - 1),
				)
			}
			case 11: {
				if (numsOrErr.length === 0) return errorValue('#DIV/0!')
				const mean = numsOrErr.reduce((a, b) => a + b, 0) / numsOrErr.length
				return numberValue(
					numsOrErr.reduce((acc, v) => acc + (v - mean) ** 2, 0) / numsOrErr.length,
				)
			}
			default:
				return errorValue('#VALUE!')
		}
	}

	if (code >= 12 && code <= 13) {
		const data = args.slice(2)
		const numsOrErr = aggregateCollectNumbers(data, ignoreErrors)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		if (numsOrErr.length === 0) return errorValue('#NUM!')
		if (code === 12) {
			numsOrErr.sort((a, b) => a - b)
			const mid = Math.floor(numsOrErr.length / 2)
			return numberValue(
				numsOrErr.length % 2 === 0
					? ((numsOrErr[mid - 1] ?? 0) + (numsOrErr[mid] ?? 0)) / 2
					: (numsOrErr[mid] ?? 0),
			)
		}
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
		return maxCount < 2 ? errorValue('#N/A') : numberValue(modeVal)
	}

	if (code >= 14 && code <= 19) {
		if (args.length < 4) return errorValue('#VALUE!')
		const kArg = numArg(args[2])
		if (typeof kArg !== 'number') return kArg
		const k = Math.trunc(kArg)
		const data = args.slice(3)
		const numsOrErr = aggregateCollectNumbers(data, ignoreErrors)
		if (!Array.isArray(numsOrErr)) return numsOrErr
		if (numsOrErr.length === 0) return errorValue('#NUM!')

		if (code === 14) {
			if (k < 1 || k > numsOrErr.length) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => b - a)
			return numberValue(numsOrErr[k - 1] ?? 0)
		}
		if (code === 15) {
			if (k < 1 || k > numsOrErr.length) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			return numberValue(numsOrErr[k - 1] ?? 0)
		}
		if (code === 16) {
			if (k < 0 || k > 1) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const n = numsOrErr.length
			const x = k * (n - 1)
			const i = Math.floor(x)
			const frac = x - i
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
		if (code === 17) {
			if (k < 0 || k > 4) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const n = numsOrErr.length
			if (k === 0) return numberValue(numsOrErr[0] ?? 0)
			if (k === 4) return numberValue(numsOrErr[n - 1] ?? 0)
			const q = (k / 4) * (n - 1)
			const i = Math.floor(q)
			const frac = q - i
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
		if (code === 18) {
			const n = numsOrErr.length
			if (k <= 0 || k >= 1) return errorValue('#NUM!')
			if (k < 1 / (n + 1) || k > n / (n + 1)) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const x = k * (n + 1) - 1
			const i = Math.floor(x)
			const frac = x - i
			if (i < 0) return numberValue(numsOrErr[0] ?? 0)
			if (i + 1 >= n) return numberValue(numsOrErr[n - 1] ?? 0)
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
		if (code === 19) {
			if (k < 1 || k > 3) return errorValue('#NUM!')
			numsOrErr.sort((a, b) => a - b)
			const n = numsOrErr.length
			const q = (k / 4) * (n + 1) - 1
			const i = Math.floor(q)
			const frac = q - i
			if (i < 0) return numberValue(numsOrErr[0] ?? 0)
			if (i + 1 >= n) return numberValue(numsOrErr[n - 1] ?? 0)
			return numberValue(
				(numsOrErr[i] ?? 0) + frac * ((numsOrErr[i + 1] ?? 0) - (numsOrErr[i] ?? 0)),
			)
		}
	}

	return errorValue('#VALUE!')
}

function subtotalNums(data: EvalArg[]): number[] | CellValue {
	const nums: number[] = []
	for (const arg of data) {
		if (arg.kind === 'range' && arg.values) {
			for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
				if (arg.rowHiddenAtOffset?.(rowIndex)) continue
				for (const cell of arg.values[rowIndex] ?? []) {
					if (isError(cell)) return cell
					const n = numericVal(cell)
					if (n !== null) nums.push(n)
				}
			}
		} else if (arg.forEachValue) {
			let err: CellValue | undefined
			arg.forEachValue((cell) => {
				if (err) return
				if (isError(cell)) {
					err = cell
					return
				}
				const n = numericVal(cell)
				if (n !== null) nums.push(n)
			})
			if (err) return err
		} else {
			const n = toNum(arg.value ?? EMPTY)
			if (typeof n !== 'number') return n
			nums.push(n)
		}
	}
	return nums
}

function subtotalDelegated(code: number, data: EvalArg[]): CellValue {
	return subtotalFn([{ value: numberValue(code) }, ...data])
}

function subtotalFn(args: EvalArg[]): CellValue {
	const fnNum = numArg(args[0])
	if (typeof fnNum !== 'number') return fnNum
	const code = Math.trunc(fnNum)
	const data = args.slice(1)
	switch (code) {
		case 1: {
			let sum = 0
			let count = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) {
								sum += n
								count++
							}
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					sum += n
					count++
				}
			}
			return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
		}
		case 2: {
			let count = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (cell.kind === 'number' || cell.kind === 'date') count++
						}
					}
				} else {
					const v = arg.value ?? EMPTY
					if (v.kind === 'number' || v.kind === 'date' || v.kind === 'boolean') count++
				}
			}
			return numberValue(count)
		}
		case 3: {
			let count = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (!isEmpty(cell)) count++
						}
					}
				} else if (!isEmpty(arg.value ?? EMPTY)) {
					count++
				}
			}
			return numberValue(count)
		}
		case 4: {
			let max = Number.NEGATIVE_INFINITY
			let found = false
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) {
								max = Math.max(max, n)
								found = true
							}
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					max = Math.max(max, n)
					found = true
				}
			}
			return numberValue(found ? max : 0)
		}
		case 5: {
			let min = Number.POSITIVE_INFINITY
			let found = false
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) {
								min = Math.min(min, n)
								found = true
							}
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					min = Math.min(min, n)
					found = true
				}
			}
			return numberValue(found ? min : 0)
		}
		case 6:
		case 106: {
			let product = 1
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) product *= n
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					product *= n
				}
			}
			return numberValue(product)
		}
		case 7:
		case 107: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			if (nums.length < 2) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
			return numberValue(Math.sqrt(sumSq / (nums.length - 1)))
		}
		case 8:
		case 108: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			if (nums.length < 2) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			const sumSq = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0)
			return numberValue(sumSq / (nums.length - 1))
		}
		case 9:
		case 109: {
			let sum = 0
			for (const arg of data) {
				if (arg.kind === 'range' && arg.values) {
					for (let rowIndex = 0; rowIndex < arg.values.length; rowIndex++) {
						if (arg.rowHiddenAtOffset?.(rowIndex)) continue
						for (const cell of arg.values[rowIndex] ?? []) {
							if (isError(cell)) return cell
							const n = numericVal(cell)
							if (n !== null) sum += n
						}
					}
				} else {
					const n = toNum(arg.value ?? EMPTY)
					if (typeof n !== 'number') return n
					sum += n
				}
			}
			return numberValue(sum)
		}
		case 101:
			return subtotalDelegated(1, data)
		case 102:
			return subtotalDelegated(2, data)
		case 103:
			return subtotalDelegated(3, data)
		case 104:
			return subtotalDelegated(4, data)
		case 105:
			return subtotalDelegated(5, data)
		case 10:
		case 110: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			const divisor = nums.length - 1
			if (divisor < 1) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			return numberValue(nums.reduce((acc, v) => acc + (v - mean) ** 2, 0) / divisor)
		}
		case 11:
		case 111: {
			const nums = subtotalNums(data)
			if (!Array.isArray(nums)) return nums
			if (nums.length === 0) return errorValue('#DIV/0!')
			const mean = nums.reduce((a, b) => a + b, 0) / nums.length
			return numberValue(nums.reduce((acc, v) => acc + (v - mean) ** 2, 0) / nums.length)
		}
		default:
			return errorValue('#VALUE!')
	}
}

export const compositeFunctions: FunctionDef[] = [
	fn('SUBTOTAL', 2, 255, subtotalFn),
	fn('AGGREGATE', 3, 253, aggregateFn),
]
