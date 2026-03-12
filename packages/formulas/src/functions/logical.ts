import type { CellValue } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, isError, topLeftScalar } from '@ascend/schema'
import type { EvalArg, FunctionDef } from './index.ts'

function fn(
	name: string,
	minArgs: number,
	maxArgs: number,
	evaluate: (args: EvalArg[]) => CellValue,
): FunctionDef {
	return { name, minArgs, maxArgs, volatile: false, evaluate }
}

function toBool(v: CellValue): boolean | CellValue {
	v = topLeftScalar(v)
	switch (v.kind) {
		case 'empty':
			return false
		case 'number':
			return v.value !== 0
		case 'string': {
			const u = v.value.toUpperCase()
			if (u === 'TRUE') return true
			if (u === 'FALSE') return false
			return errorValue('#VALUE!')
		}
		case 'boolean':
			return v.value
		case 'error':
			return v
		case 'date':
			return v.serial !== 0
		case 'richText':
			return errorValue('#VALUE!')
	}
}

function valuesMatch(a: CellValue, b: CellValue): boolean {
	if (a.kind === 'number' && b.kind === 'number') return a.value === b.value
	if (a.kind === 'string' && b.kind === 'string')
		return a.value.toLowerCase() === b.value.toLowerCase()
	if (a.kind === 'boolean' && b.kind === 'boolean') return a.value === b.value
	if (a.kind === 'error' && b.kind === 'error') return a.value === b.value
	if (a.kind === 'empty' && b.kind === 'empty') return true
	return false
}

export const logicalFunctions: FunctionDef[] = [
	fn('IF', 2, 3, (args) => {
		const cond = toBool(args[0]?.value ?? EMPTY)
		if (typeof cond !== 'boolean') return cond
		if (cond) return args[1]?.value ?? EMPTY
		return args.length >= 3 ? (args[2]?.value ?? EMPTY) : booleanValue(false)
	}),

	fn('IFS', 2, 254, (args) => {
		for (let i = 0; i + 1 < args.length; i += 2) {
			const v = args[i]?.value ?? EMPTY
			if (isError(v)) return v
			const cond = toBool(v)
			if (typeof cond !== 'boolean') return cond
			if (cond) return args[i + 1]?.value ?? EMPTY
		}
		return errorValue('#N/A')
	}),

	fn('AND', 1, 255, (args) => {
		let found = false
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						if (cell.kind === 'boolean') {
							found = true
							if (!cell.value) return booleanValue(false)
						} else if (cell.kind === 'number') {
							found = true
							if (cell.value === 0) return booleanValue(false)
						}
					}
				}
			} else {
				const b = toBool(arg.value ?? EMPTY)
				if (typeof b !== 'boolean') return b
				found = true
				if (!b) return booleanValue(false)
			}
		}
		return found ? booleanValue(true) : errorValue('#VALUE!')
	}),

	fn('OR', 1, 255, (args) => {
		let found = false
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						if (cell.kind === 'boolean') {
							found = true
							if (cell.value) return booleanValue(true)
						} else if (cell.kind === 'number') {
							found = true
							if (cell.value !== 0) return booleanValue(true)
						}
					}
				}
			} else {
				const b = toBool(arg.value ?? EMPTY)
				if (typeof b !== 'boolean') return b
				found = true
				if (b) return booleanValue(true)
			}
		}
		return found ? booleanValue(false) : errorValue('#VALUE!')
	}),

	fn('NOT', 1, 1, (args) => {
		const b = toBool(args[0]?.value ?? EMPTY)
		if (typeof b !== 'boolean') return b
		return booleanValue(!b)
	}),

	fn('XOR', 1, 255, (args) => {
		let trueCount = 0
		let found = false
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						if (cell.kind === 'boolean') {
							found = true
							if (cell.value) trueCount++
						} else if (cell.kind === 'number') {
							found = true
							if (cell.value !== 0) trueCount++
						}
					}
				}
			} else {
				const b = toBool(arg.value ?? EMPTY)
				if (typeof b !== 'boolean') return b
				found = true
				if (b) trueCount++
			}
		}
		return found ? booleanValue(trueCount % 2 === 1) : errorValue('#VALUE!')
	}),

	fn('IFERROR', 2, 2, (args) => {
		const v = args[0]?.value ?? EMPTY
		return isError(v) ? (args[1]?.value ?? EMPTY) : v
	}),

	fn('IFNA', 2, 2, (args) => {
		const v = args[0]?.value ?? EMPTY
		return v.kind === 'error' && v.value === '#N/A' ? (args[1]?.value ?? EMPTY) : v
	}),

	fn('TRUE', 0, 0, () => booleanValue(true)),

	fn('FALSE', 0, 0, () => booleanValue(false)),

	fn('SWITCH', 3, 254, (args) => {
		const expr = args[0]?.value ?? EMPTY
		if (isError(expr)) return expr

		const remaining = args.length - 1
		const hasDefault = remaining % 2 === 1
		const pairEnd = hasDefault ? args.length - 1 : args.length

		for (let i = 1; i < pairEnd; i += 2) {
			const val = args[i]?.value ?? EMPTY
			if (isError(val)) return val
			if (valuesMatch(expr, val)) return args[i + 1]?.value ?? EMPTY
		}

		return hasDefault ? (args[args.length - 1]?.value ?? EMPTY) : errorValue('#N/A')
	}),
]
