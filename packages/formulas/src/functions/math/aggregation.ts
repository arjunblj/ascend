import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, isEmpty, isError, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef, FunctionEvalContext } from '../index.ts'
import { fn, getRange, numericVal, sameShape, toNum } from './helpers.ts'

function singleRangeAggregateCacheKey(name: string, args: EvalArg[]): string | null {
	if (args.length !== 1) return null
	const ref = args[0]?.ref
	if (!ref || ref.kind !== 'range') return null
	return `${name}:${ref.sheetIndex}:${ref.row}:${ref.col}:${ref.endRow ?? ref.row}:${ref.endCol ?? ref.col}`
}

function withCachedSingleRangeAggregate(
	name: 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX',
	args: EvalArg[],
	ctx: FunctionEvalContext | undefined,
	compute: () => CellValue,
): CellValue {
	const key = ctx?.aggregateRangeCache ? singleRangeAggregateCacheKey(name, args) : null
	if (!key || !ctx?.aggregateRangeCache) return compute()
	const cached = ctx.aggregateRangeCache.get(key)
	if (cached) return cached
	const value = compute()
	ctx.aggregateRangeCache.set(key, value)
	return value
}

export const aggregationFunctions: FunctionDef[] = [
	fn('SUM', 1, 255, (args, ctx) =>
		withCachedSingleRangeAggregate('SUM', args, ctx, () => {
			let sum = 0
			for (const arg of args) {
				if (arg.forEachValue) {
					let err: CellValue | undefined
					arg.forEachValue((cell) => {
						if (err) return
						if (isError(cell)) {
							err = cell
							return
						}
						const n = numericVal(cell)
						if (n !== null) sum += n
					})
					if (err) return err
				} else if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
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
		}),
	),

	fn('SUMPRODUCT', 1, 255, (args) => {
		const ranges = args.map(getRange)
		const rows = ranges[0]?.length ?? 0
		const cols = ranges[0]?.[0]?.length ?? 0
		for (let i = 1; i < ranges.length; i++) {
			if (!sameShape(ranges[0] ?? [], ranges[i] ?? [])) return errorValue('#VALUE!')
		}
		const useFastPath = (() => {
			for (let c = 0; c < cols; c++) {
				for (const range of ranges) {
					const cell = range[0]?.[c] ?? EMPTY
					if (cell.kind === 'error' || cell.kind === 'string' || cell.kind === 'boolean')
						return false
				}
			}
			return true
		})()
		if (useFastPath) {
			let total = 0
			for (let r = 0; r < rows; r++) {
				for (let c = 0; c < cols; c++) {
					let product = 1
					for (const range of ranges) {
						const cell = range[r]?.[c] ?? EMPTY
						product *= cell.kind === 'number' ? cell.value : cell.kind === 'date' ? cell.serial : 0
					}
					total += product
				}
			}
			return numberValue(total)
		}
		let total = 0
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				let product = 1
				for (const range of ranges) {
					const cell = range[r]?.[c] ?? EMPTY
					if (isError(cell)) return cell
					const n = numericVal(cell)
					product *= n ?? (cell.kind === 'boolean' ? (cell.value ? 1 : 0) : 0)
				}
				total += product
			}
		}
		return numberValue(total)
	}),

	fn('AVERAGE', 1, 255, (args, ctx) =>
		withCachedSingleRangeAggregate('AVERAGE', args, ctx, () => {
			let sum = 0
			let count = 0
			for (const arg of args) {
				if (arg.forEachValue) {
					let err: CellValue | undefined
					arg.forEachValue((cell) => {
						if (err) return
						if (isError(cell)) {
							err = cell
							return
						}
						const n = numericVal(cell)
						if (n !== null) {
							sum += n
							count++
						}
					})
					if (err) return err
				} else if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
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
		}),
	),

	fn('COUNT', 1, 255, (args, ctx) =>
		withCachedSingleRangeAggregate('COUNT', args, ctx, () => {
			let count = 0
			for (const arg of args) {
				if (arg.forEachValue) {
					arg.forEachValue((cell) => {
						if (cell.kind === 'number' || cell.kind === 'date') count++
					})
				} else if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
							if (cell.kind === 'number' || cell.kind === 'date') count++
						}
					}
				} else {
					const v = arg.value ?? EMPTY
					if (v.kind === 'number' || v.kind === 'date' || v.kind === 'boolean') count++
				}
			}
			return numberValue(count)
		}),
	),

	fn('COUNTA', 1, 255, (args) => {
		let count = 0
		for (const arg of args) {
			if (arg.forEachValue) {
				arg.forEachValue((cell) => {
					if (!isEmpty(cell)) count++
				})
			} else if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (!isEmpty(cell)) count++
					}
				}
			} else {
				if (!isEmpty(arg.value ?? EMPTY)) count++
			}
		}
		return numberValue(count)
	}),

	fn('COUNTBLANK', 1, 1, (args) => {
		let count = 0
		const arg = args[0]
		const countBlank = (cell: CellValue) => {
			if (isEmpty(cell) || (cell.kind === 'string' && cell.value === '')) count++
		}
		if (arg?.forEachCellInRange) {
			arg.forEachCellInRange(countBlank)
		} else if (arg?.forEachValue) {
			arg.forEachValue(countBlank)
		} else {
			for (const row of getRange(arg)) {
				for (const cell of row) countBlank(cell)
			}
		}
		return numberValue(count)
	}),

	fn('MIN', 1, 255, (args, ctx) =>
		withCachedSingleRangeAggregate('MIN', args, ctx, () => {
			let min = Number.POSITIVE_INFINITY
			let found = false
			for (const arg of args) {
				if (arg.forEachValue) {
					let err: CellValue | undefined
					arg.forEachValue((cell) => {
						if (err) return
						if (isError(cell)) {
							err = cell
							return
						}
						const n = numericVal(cell)
						if (n !== null) {
							min = Math.min(min, n)
							found = true
						}
					})
					if (err) return err
				} else if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
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
		}),
	),

	fn('MAX', 1, 255, (args, ctx) =>
		withCachedSingleRangeAggregate('MAX', args, ctx, () => {
			let max = Number.NEGATIVE_INFINITY
			let found = false
			for (const arg of args) {
				if (arg.forEachValue) {
					let err: CellValue | undefined
					arg.forEachValue((cell) => {
						if (err) return
						if (isError(cell)) {
							err = cell
							return
						}
						const n = numericVal(cell)
						if (n !== null) {
							max = Math.max(max, n)
							found = true
						}
					})
					if (err) return err
				} else if (arg.kind === 'range' && arg.values) {
					for (const row of arg.values) {
						for (const cell of row) {
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
		}),
	),

	fn('PRODUCT', 1, 255, (args) => {
		let product = 1
		let found = false
		for (const arg of args) {
			if (arg.forEachValue) {
				let err: CellValue | undefined
				arg.forEachValue((cell) => {
					if (err) return
					if (isError(cell)) {
						err = cell
						return
					}
					const n = numericVal(cell)
					if (n !== null) {
						product *= n
						found = true
					}
				})
				if (err) return err
			} else if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						const n = numericVal(cell)
						if (n !== null) {
							product *= n
							found = true
						}
					}
				}
			} else {
				const n = toNum(arg.value ?? EMPTY)
				if (typeof n !== 'number') return n
				product *= n
				found = true
			}
		}
		return numberValue(found ? product : 0)
	}),
]
