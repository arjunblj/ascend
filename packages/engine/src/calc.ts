import { indexToColumn, parseRange, type RangeRef, type StyleId, type Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import { EMPTY, errorValue, topLeftScalar } from '@ascend/schema'
import { analyzeWorkbook } from './analysis.ts'
import type { CalcContext } from './calc-context.ts'
import { type CompiledFormula, compileFormula, evaluateCompiled } from './compiled-eval.ts'
import { type CellKey, cellKey, type DependencyGraph, parseCellKey } from './dep-graph.ts'
import { type EvalContext, evaluate } from './evaluator.ts'

export interface RecalcResult {
	readonly changed: string[]
	readonly errors: Array<{ ref: string; error: AscendError }>
	readonly duration: number
}

function valuesEqual(a: CellValue, b: CellValue): boolean {
	a = topLeftScalar(a)
	b = topLeftScalar(b)
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

function toScalarMatrix(value: CellValue): readonly (readonly CellValue[])[] | null {
	if (value.kind !== 'array') return null
	return value.rows
}

function topLeftValue(value: CellValue): CellValue {
	return topLeftScalar(value)
}

const compiledCache = new WeakMap<FormulaNode, CompiledFormula | false>()

export function clearCompiledFormulaCache(): void {
	// WeakMap entries are reclaimed automatically; this is a no-op placeholder
	// kept for API symmetry with clearFormulaParseCache.
}

function evalFormula(_key: CellKey, ast: FormulaNode, ctx: EvalContext): CellValue {
	let compiled = compiledCache.get(ast)
	if (compiled === undefined) {
		const result = compileFormula(ast)
		compiled = result ?? false
		compiledCache.set(ast, compiled)
	}
	if (compiled !== false) {
		return evaluateCompiled(compiled, ctx)
	}
	return evaluate(ast, ctx)
}

function isSpillBinding(
	binding: unknown,
): binding is { kind: 'spill'; anchorRef: string; ref: string; isAnchor: boolean } {
	return (
		typeof binding === 'object' &&
		binding !== null &&
		(binding as { kind?: string }).kind === 'spill'
	)
}

function clearSpillFootprint(
	sheet: Workbook['sheets'][number],
	anchorRef: string,
	changed: string[],
): void {
	if (!sheet) return
	const removals: Array<{ row: number; col: number; ref: string }> = []
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (
			cell.formulaInfo &&
			isSpillBinding(cell.formulaInfo) &&
			cell.formulaInfo.anchorRef === anchorRef
		) {
			removals.push({ row, col, ref: toA1Ref(row, col) })
		}
	}
	for (const removal of removals) {
		sheet.cells.delete(removal.row, removal.col)
		changed.push(`${sheet.name}!${removal.ref}`)
	}
}

function isSpillBlocked(
	sheet: Workbook['sheets'][number],
	anchorRow: number,
	anchorCol: number,
	anchorRef: string,
	matrix: readonly (readonly CellValue[])[],
): boolean {
	if (!sheet) return true
	for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) {
		const sourceRow = matrix[rowOffset] ?? []
		for (let colOffset = 0; colOffset < sourceRow.length; colOffset++) {
			if (rowOffset === 0 && colOffset === 0) continue
			const targetRow = anchorRow + rowOffset
			const targetCol = anchorCol + colOffset
			const existing = sheet.cells.get(targetRow, targetCol)
			if (!existing) continue
			if (
				existing.formulaInfo &&
				isSpillBinding(existing.formulaInfo) &&
				existing.formulaInfo.anchorRef === anchorRef
			) {
				continue
			}
			return true
		}
	}
	return false
}

function applyArrayResult(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	oldCell:
		| { value: CellValue; formula: string | null; styleId: StyleId; formulaInfo?: unknown }
		| undefined,
	matrix: readonly (readonly CellValue[])[],
	changed: string[],
): CellValue {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return errorValue('#REF!')
	const anchorRef = `${sheet.name}!${toA1Ref(row, col)}`
	clearSpillFootprint(sheet, anchorRef, changed)
	if (matrix.length === 0 || (matrix[0]?.length ?? 0) === 0) return EMPTY
	if (isSpillBlocked(sheet, row, col, anchorRef, matrix)) {
		const spillError = errorValue('#SPILL!')
		sheet.cells.set(row, col, {
			value: spillError,
			formula: oldCell?.formula ?? null,
			styleId: oldCell?.styleId ?? (0 as StyleId),
		})
		changed.push(anchorRef)
		return spillError
	}

	const anchorValue = topLeftValue(matrix[0]?.[0] ?? EMPTY)
	const spillRef = `${toA1Ref(row, col)}:${toA1Ref(
		row + matrix.length - 1,
		col + Math.max(...matrix.map((matrixRow) => matrixRow.length), 1) - 1,
	)}`

	sheet.cells.set(row, col, {
		value: anchorValue,
		formula: oldCell?.formula ?? null,
		styleId: oldCell?.styleId ?? (0 as StyleId),
		formulaInfo: {
			kind: 'spill',
			anchorRef,
			ref: spillRef,
			isAnchor: true,
		},
	})
	changed.push(anchorRef)

	for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) {
		const sourceRow = matrix[rowOffset] ?? []
		for (let colOffset = 0; colOffset < sourceRow.length; colOffset++) {
			if (rowOffset === 0 && colOffset === 0) continue
			const targetRow = row + rowOffset
			const targetCol = col + colOffset
			sheet.cells.set(targetRow, targetCol, {
				value: topLeftValue(sourceRow[colOffset] ?? EMPTY),
				formula: null,
				styleId: oldCell?.styleId ?? (0 as StyleId),
				formulaInfo: {
					kind: 'spill',
					anchorRef,
					ref: spillRef,
					isAnchor: false,
				},
			})
			changed.push(`${sheet.name}!${toA1Ref(targetRow, targetCol)}`)
		}
	}
	return anchorValue
}

function toA1Ref(row: number, col: number): string {
	return `${indexToColumn(col)}${row + 1}`
}

export function recalculate(
	workbook: Workbook,
	ctx: CalcContext,
	opts?: { dirtyOnly?: boolean; range?: RangeRef; dirtyRefs?: readonly string[] },
): RecalcResult {
	const start = performance.now()
	const asts = new Map<CellKey, FormulaNode>()
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

	let evalOrder: CellKey[]

	if (opts?.dirtyOnly || (opts?.dirtyRefs?.length ?? 0) > 0) {
		const dirtySeeds = [...graph.getVolatiles(), ...resolveDirtyKeys(workbook, opts?.dirtyRefs)]
		const dirty = graph.getDirtySet(dirtySeeds)
		evalOrder = graph.getEvalOrder(dirty)
	} else {
		const allKeys = graph.getAllFormulaCells()
		const allSet = new Set(allKeys)
		evalOrder = graph.getEvalOrder(allSet)
	}

	const cycles = graph.detectCycles()
	const cycleKeys = new Set<CellKey>()
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
					if (
						oldCell?.formulaInfo &&
						isSpillBinding(oldCell.formulaInfo) &&
						oldCell.formulaInfo.isAnchor
					) {
						clearSpillFootprint(sheet, oldCell.formulaInfo.anchorRef, changed)
					}
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
			const newValue = evalFormula(key, ast, evalCtx)
			const oldCell = sheet.cells.get(row, col)
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				applyArrayResult(workbook, si, row, col, oldCell, spillMatrix, changed)
				continue
			}
			if (
				oldCell?.formulaInfo &&
				isSpillBinding(oldCell.formulaInfo) &&
				oldCell.formulaInfo.isAnchor
			) {
				clearSpillFootprint(sheet, oldCell.formulaInfo.anchorRef, changed)
			}
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

function resolveDirtyKeys(workbook: Workbook, refs: readonly string[] | undefined): CellKey[] {
	if (!refs || refs.length === 0) return []
	const keys: CellKey[] = []
	for (const ref of refs) {
		const bang = ref.indexOf('!')
		const sheetName =
			bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : workbook.sheets[0]?.name
		const localRef = bang >= 0 ? ref.slice(bang + 1) : ref
		if (!sheetName || !localRef) continue
		const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.name === sheetName)
		if (sheetIndex === -1) continue
		const range = parseRange(localRef)
		for (let row = range.start.row; row <= range.end.row; row++) {
			for (let col = range.start.col; col <= range.end.col; col++) {
				keys.push(cellKey(sheetIndex, row, col))
			}
		}
	}
	return keys
}

function evalIterative(
	workbook: Workbook,
	ctx: CalcContext,
	_graph: DependencyGraph,
	asts: Map<CellKey, FormulaNode>,
	evalOrder: CellKey[],
	cycleKeys: Set<CellKey>,
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
			const newValue = evalFormula(key, ast, evalCtx)
			const oldCell = sheet.cells.get(row, col)
			const oldValue = oldCell?.value ?? EMPTY
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				applyArrayResult(workbook, si, row, col, oldCell, spillMatrix, changed)
				continue
			}
			if (
				oldCell?.formulaInfo &&
				isSpillBinding(oldCell.formulaInfo) &&
				oldCell.formulaInfo.isAnchor
			) {
				clearSpillFootprint(sheet, oldCell.formulaInfo.anchorRef, changed)
			}

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
		const newValue = evalFormula(key, ast, evalCtx)
		const oldCell = sheet.cells.get(row, col)
		const spillMatrix = toScalarMatrix(newValue)
		if (spillMatrix) {
			applyArrayResult(workbook, si, row, col, oldCell, spillMatrix, changed)
			continue
		}
		if (
			oldCell?.formulaInfo &&
			isSpillBinding(oldCell.formulaInfo) &&
			oldCell.formulaInfo.isAnchor
		) {
			clearSpillFootprint(sheet, oldCell.formulaInfo.anchorRef, changed)
		}
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
