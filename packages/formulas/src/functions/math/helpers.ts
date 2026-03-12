import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue, topLeftScalar } from '@ascend/schema'
import type { EvalArg, FunctionDef, FunctionEvalContext } from '../index.ts'
import { iterAreaRows } from '../registry.ts'

export function fn(
	name: string,
	minArgs: number,
	maxArgs: number,
	evaluate: (args: EvalArg[], ctx?: FunctionEvalContext) => CellValue,
	volatile = false,
): FunctionDef {
	return { name, minArgs, maxArgs, volatile, evaluate }
}

export function seededRandom(ctx?: FunctionEvalContext): number {
	const seed = ctx?.randomSeed ?? 42
	const row = ctx?.row ?? 0
	const col = ctx?.col ?? 0
	const sheet = ctx?.sheetIndex ?? 0
	let state =
		(seed ^ ((sheet + 1) * 0x9e3779b1) ^ ((row + 1) * 0x85ebca6b) ^ ((col + 1) * 0xc2b2ae35)) >>> 0
	state ^= state >>> 16
	state = Math.imul(state, 0x7feb352d) >>> 0
	state ^= state >>> 15
	state = Math.imul(state, 0x846ca68b) >>> 0
	state ^= state >>> 16
	return (state >>> 0) / 0x1_0000_0000
}

export function toNum(v: CellValue): number | CellValue {
	v = topLeftScalar(v)
	switch (v.kind) {
		case 'empty':
			return 0
		case 'number':
			return v.value
		case 'string': {
			if (v.value.trim() === '') return 0
			const n = Number(v.value)
			return Number.isNaN(n) ? errorValue('#VALUE!') : n
		}
		case 'boolean':
			return v.value ? 1 : 0
		case 'error':
			return v
		case 'date':
			return v.serial
		case 'richText':
			return errorValue('#VALUE!')
	}
}

export function numArg(arg: EvalArg | undefined): number | CellValue {
	return toNum(arg?.value ?? EMPTY)
}

export function getRange(arg: EvalArg | undefined): readonly (readonly CellValue[])[] {
	return iterAreaRows(arg)
}

export function numericVal(cell: CellValue): number | null {
	const v = topLeftScalar(cell)
	if (v.kind === 'number') return v.value
	if (v.kind === 'date') return v.serial
	return null
}

export function sameShape(
	left: readonly (readonly CellValue[])[],
	right: readonly (readonly CellValue[])[],
): boolean {
	if (left.length !== right.length) return false
	for (let row = 0; row < left.length; row++) {
		if ((left[row]?.length ?? 0) !== (right[row]?.length ?? 0)) return false
	}
	return true
}
