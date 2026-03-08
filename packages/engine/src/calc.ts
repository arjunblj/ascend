import { indexToColumn, type RangeRef, type StyleId, type Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { extractRefs, functionRegistry, parseFormula } from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import { EMPTY, errorValue } from '@ascend/schema'
import type { CalcContext } from './calc-context.ts'
import { cellKey, DependencyGraph, parseCellKey } from './dep-graph.ts'
import { type EvalContext, evaluate } from './evaluator.ts'

export interface RecalcResult {
	readonly changed: string[]
	readonly errors: Array<{ ref: string; error: AscendError }>
	readonly duration: number
}

interface ParsedFormula {
	readonly ast: FormulaNode
	readonly volatile: boolean
	readonly deps: string[]
}

function parseAndAnalyze(formula: string, sheetIndex: number, wb: Workbook): ParsedFormula | null {
	const result = parseFormula(formula)
	if (!result.ok) return null

	const ast = result.value
	const refs = extractRefs(ast)
	const deps: string[] = []
	let volatile = false

	if (ast.type === 'function') {
		const fn = functionRegistry.get(ast.name.toUpperCase())
		if (fn?.volatile) volatile = true
	}
	checkVolatile(ast, (v) => {
		volatile = volatile || v
	})

	for (const ref of refs) {
		if (ref.kind === 'cell') {
			const si = resolveSheetIdx(wb, ref.sheet, sheetIndex)
			if (si >= 0) deps.push(cellKey(si, ref.ref.row, ref.ref.col))
		} else {
			const si = resolveSheetIdx(wb, ref.sheet, sheetIndex)
			if (si >= 0) {
				for (let r = ref.start.row; r <= ref.end.row; r++) {
					for (let c = ref.start.col; c <= ref.end.col; c++) {
						deps.push(cellKey(si, r, c))
					}
				}
			}
		}
	}

	return { ast, volatile, deps }
}

function checkVolatile(node: FormulaNode, cb: (v: boolean) => void): void {
	if (node.type === 'function') {
		const fn = functionRegistry.get(node.name.toUpperCase())
		if (fn?.volatile) cb(true)
		for (const arg of node.args) checkVolatile(arg, cb)
	} else if (node.type === 'binary') {
		checkVolatile(node.left, cb)
		checkVolatile(node.right, cb)
	} else if (node.type === 'unary') {
		checkVolatile(node.operand, cb)
	}
}

function resolveSheetIdx(wb: Workbook, sheetName: string | undefined, fallback: number): number {
	if (sheetName === undefined) return fallback
	return wb.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase())
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
	const graph = new DependencyGraph()
	const asts = new Map<string, FormulaNode>()
	const changed: string[] = []
	const errors: Array<{ ref: string; error: AscendError }> = []

	for (let si = 0; si < workbook.sheets.length; si++) {
		const sheet = workbook.sheets[si]
		if (!sheet) continue
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue

			if (opts?.range) {
				if (
					opts.range.sheet !== undefined &&
					opts.range.sheet.toLowerCase() !== sheet.name.toLowerCase()
				) {
					continue
				}
				if (
					row < opts.range.start.row ||
					row > opts.range.end.row ||
					col < opts.range.start.col ||
					col > opts.range.end.col
				) {
					continue
				}
			}

			const key = cellKey(si, row, col)
			const parsed = parseAndAnalyze(cell.formula, si, workbook)
			if (parsed) {
				graph.addFormula(key, parsed.deps, parsed.volatile)
				asts.set(key, parsed.ast)
			} else {
				errors.push({
					ref: cellRefString(workbook, si, row, col),
					error: {
						code: 'FORMULA_PARSE_ERROR',
						message: `Failed to parse: ${cell.formula}`,
						retryable: false,
					},
				})
			}
		}
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
