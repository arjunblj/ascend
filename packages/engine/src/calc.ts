import { indexToColumn, type RangeRef, type StyleId, type Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import { EMPTY, errorValue } from '@ascend/schema'
import { analyzeWorkbook } from './analysis.ts'
import type { CalcContext } from './calc-context.ts'
import { type DependencyGraph, parseCellKey } from './dep-graph.ts'
import { type EvalContext, evaluate } from './evaluator.ts'

export interface RecalcResult {
	readonly changed: string[]
	readonly errors: Array<{ ref: string; error: AscendError }>
	readonly duration: number
}

function valuesEqual(a: CellValue, b: CellValue): boolean {
	if (a.kind !== b.kind) return false
	switch (a.kind) {
		case 'empty':
			return true
		case 'number':
			return b.kind === 'number' && a.value === b.value
		case 'string':
			return b.kind === 'string' && a.value === b.value
		case 'boolean':
			return b.kind === 'boolean' && a.value === b.value
		case 'error':
			return b.kind === 'error' && a.value === b.value
		case 'date':
			return b.kind === 'date' && a.serial === b.serial
		default:
			return false
	}
}

function cellRefString(wb: Workbook, sheetIndex: number, row: number, col: number): string {
	const sheet = wb.sheets[sheetIndex]
	const name = sheet ? sheet.name : `Sheet${sheetIndex + 1}`
	return `${name}!${indexToColumn(col)}${row + 1}`
}

export function recalculate(
	workbook: Workbook,
	ctx: CalcContext,
	opts?: { dirtyOnly?: boolean; range?: RangeRef },
): RecalcResult {
	const start = performance.now()
	const asts = new Map<string, FormulaNode>()
	const changed: string[] = []
	const errors: Array<{ ref: string; error: AscendError }> = []

	const analysis = analyzeWorkbook(workbook, opts?.range ? { range: opts.range } : undefined)
	const graph = analysis.dependencyGraph

	for (const analyzed of analysis.formulas.values()) {
		if (analyzed.parseError || !analyzed.ast) {
			errors.push({
				ref: cellRefString(workbook, analyzed.sheetIndex, analyzed.row, analyzed.col),
				error: {
					code: 'FORMULA_PARSE_ERROR',
					message: `Failed to parse: ${analyzed.formula}`,
					retryable: false,
				},
			})
			continue
		}
		asts.set(analyzed.key, analyzed.ast)
	}

	let evalOrder: string[]

	if (opts?.dirtyOnly) {
		const volatiles = graph.getVolatiles()
		const dirty = graph.getDirtySet(volatiles)
		evalOrder = graph.getEvalOrder(dirty)
	} else {
		const allKeys = graph.getAllFormulaCells()
		const allSet = new Set(allKeys)
		evalOrder = graph.getEvalOrder(allSet)
	}

	const cycles = graph.detectCycles()
	const cycleKeys = new Set<string>()
	for (const scc of cycles) {
		for (const key of scc) cycleKeys.add(key)
	}

	if (ctx.iterativeCalc.enabled && cycleKeys.size > 0) {
		evalIterative(workbook, ctx, graph, asts, evalOrder, cycleKeys, changed, errors)
	} else {
		for (const key of evalOrder) {
			if (cycleKeys.has(key)) {
				const [si, row, col] = parseCellKey(key)
				const sheet = workbook.sheets[si]
				if (sheet) {
					const oldCell = sheet.cells.get(row, col)
					const newValue = errorValue('#REF!')
					if (!oldCell || !valuesEqual(oldCell.value, newValue)) {
						sheet.cells.set(row, col, {
							value: newValue,
							formula: oldCell?.formula ?? null,
							styleId: oldCell?.styleId ?? (0 as StyleId),
						})
						changed.push(cellRefString(workbook, si, row, col))
					}
					errors.push({
						ref: cellRefString(workbook, si, row, col),
						error: {
							code: 'CIRCULAR_REF',
							message: 'Circular reference detected',
							retryable: false,
						},
					})
				}
				continue
			}

			const ast = asts.get(key)
			if (!ast) continue

			const [si, row, col] = parseCellKey(key)
			const sheet = workbook.sheets[si]
			if (!sheet) continue

			const evalCtx: EvalContext = {
				workbook,
				calcContext: ctx,
				sheetIndex: si,
				row,
				col,
			}
			const newValue = evaluate(ast, evalCtx)
			const oldCell = sheet.cells.get(row, col)
			if (!oldCell || !valuesEqual(oldCell.value, newValue)) {
				sheet.cells.set(row, col, {
					value: newValue,
					formula: oldCell?.formula ?? null,
					styleId: oldCell?.styleId ?? (0 as StyleId),
				})
				changed.push(cellRefString(workbook, si, row, col))
			}
		}
	}

	return {
		changed,
		errors,
		duration: performance.now() - start,
	}
}

function evalIterative(
	workbook: Workbook,
	ctx: CalcContext,
	_graph: DependencyGraph,
	asts: Map<string, FormulaNode>,
	evalOrder: string[],
	cycleKeys: Set<string>,
	changed: string[],
	_errors: Array<{ ref: string; error: AscendError }>,
): void {
	const maxIter = ctx.iterativeCalc.maxIterations
	const maxChange = ctx.iterativeCalc.maxChange

	for (let iter = 0; iter < maxIter; iter++) {
		let maxDelta = 0
		for (const key of evalOrder) {
			if (!cycleKeys.has(key)) continue
			const ast = asts.get(key)
			if (!ast) continue

			const [si, row, col] = parseCellKey(key)
			const sheet = workbook.sheets[si]
			if (!sheet) continue

			const evalCtx: EvalContext = {
				workbook,
				calcContext: ctx,
				sheetIndex: si,
				row,
				col,
			}
			const newValue = evaluate(ast, evalCtx)
			const oldCell = sheet.cells.get(row, col)
			const oldValue = oldCell?.value ?? EMPTY

			if (oldValue.kind === 'number' && newValue.kind === 'number') {
				maxDelta = Math.max(maxDelta, Math.abs(newValue.value - oldValue.value))
			}

			sheet.cells.set(row, col, {
				value: newValue,
				formula: oldCell?.formula ?? null,
				styleId: oldCell?.styleId ?? (0 as StyleId),
			})
		}

		if (maxDelta <= maxChange) break
	}

	for (const key of cycleKeys) {
		const [si, row, col] = parseCellKey(key)
		changed.push(cellRefString(workbook, si, row, col))
	}

	for (const key of evalOrder) {
		if (cycleKeys.has(key)) continue
		const ast = asts.get(key)
		if (!ast) continue

		const [si, row, col] = parseCellKey(key)
		const sheet = workbook.sheets[si]
		if (!sheet) continue

		const evalCtx: EvalContext = {
			workbook,
			calcContext: ctx,
			sheetIndex: si,
			row,
			col,
		}
		const newValue = evaluate(ast, evalCtx)
		const oldCell = sheet.cells.get(row, col)
		if (!oldCell || !valuesEqual(oldCell.value, newValue)) {
			sheet.cells.set(row, col, {
				value: newValue,
				formula: oldCell?.formula ?? null,
				styleId: oldCell?.styleId ?? (0 as StyleId),
			})
			changed.push(cellRefString(workbook, si, row, col))
		}
	}
}
