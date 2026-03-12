/**
 * EVAL-5 Formula group batch execution: shared formulas (formulaInfo.kind==='shared')
 * pointing to the same master are evaluated as a group when encountered in evalOrder.
 * getSharedFormulaGroups identifies groups; the eval loop batches them to eliminate
 * per-cell overhead (AST lookup, context setup, dirty checks) for subsequent members.
 *
 * EVAL-6 Common subexpression elimination (DEFERRED): formulas with repeated
 * subexpressions (e.g. A1:A1000 used twice) are evaluated independently each
 * time. CSE would: (1) identify shared sub-DAGs in the formula AST, (2) compute
 * once and cache by (formula, cell-context) key, (3) reuse on subsequent refs.
 * Requires eval context to support memoization and invalidation. Large structural
 * change; deferred for future work.
 */
import { indexToColumn, parseRange, type RangeRef, type StyleId, type Workbook } from '@ascend/core'
import type { ExactLookupCache, FormulaNode, LookupVectorCache } from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import { EMPTY, errorValue, topLeftScalar } from '@ascend/schema'
import { analyzeWorkbook, getSharedFormulaGroups } from './analysis.ts'
import type { CalcContext } from './calc-context.ts'
import { type CompiledFormula, compileFormula, evaluateCompiled } from './compiled-eval.ts'
import {
	type CellCoords,
	type CellKey,
	cellKey,
	type DependencyGraph,
	parseCellKeyInto,
} from './dep-graph.ts'
import { type EvalContext, evaluate, MutableEvalContext } from './evaluator.ts'

export interface RecalcResult {
	readonly changed: string[]
	readonly errors: Array<{ ref: string; error: AscendError }>
	readonly duration: number
}

interface SpillEntry {
	readonly row: number
	readonly col: number
	readonly ref: string
}

interface SpillIndexState {
	readonly bySheet: Map<number, Map<string, SpillEntry[]>>
	readonly initializedSheets: Set<number>
}

function valuesEqual(a: CellValue, b: CellValue): boolean {
	if (a === b) return true
	if (a.kind === 'array') a = topLeftScalar(a)
	if (b.kind === 'array') b = topLeftScalar(b)
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

function getSheetSpillIndex(
	spillIndex: SpillIndexState,
	sheetIndex: number,
	sheet: Workbook['sheets'][number],
): Map<string, SpillEntry[]> {
	const existing = spillIndex.bySheet.get(sheetIndex)
	if (spillIndex.initializedSheets.has(sheetIndex)) return existing ?? new Map()
	const entries = existing ?? new Map<string, SpillEntry[]>()
	if (sheet) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formulaInfo || !isSpillBinding(cell.formulaInfo)) continue
			const spillEntry = { row, col, ref: toA1Ref(row, col) }
			const bucket = entries.get(cell.formulaInfo.anchorRef)
			if (bucket) bucket.push(spillEntry)
			else entries.set(cell.formulaInfo.anchorRef, [spillEntry])
		}
	}
	spillIndex.bySheet.set(sheetIndex, entries)
	spillIndex.initializedSheets.add(sheetIndex)
	return entries
}

function recordSpillCell(
	spillIndex: SpillIndexState,
	sheetIndex: number,
	sheet: Workbook['sheets'][number],
	anchorRef: string,
	row: number,
	col: number,
): void {
	const entries = getSheetSpillIndex(spillIndex, sheetIndex, sheet)
	const spillEntry = { row, col, ref: toA1Ref(row, col) }
	const bucket = entries.get(anchorRef)
	if (bucket) bucket.push(spillEntry)
	else entries.set(anchorRef, [spillEntry])
}

function clearSpillFootprint(
	sheet: Workbook['sheets'][number],
	sheetIndex: number,
	anchorRef: string,
	changed: string[],
	spillIndex: SpillIndexState,
): boolean {
	if (!sheet) return false
	const entries = getSheetSpillIndex(spillIndex, sheetIndex, sheet)
	const removals = entries.get(anchorRef)
	if (!removals || removals.length === 0) return false
	entries.delete(anchorRef)
	for (const removal of removals) {
		sheet.cells.delete(removal.row, removal.col)
		changed.push(`${sheet.name}!${removal.ref}`)
	}
	return true
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
	spillIndex: SpillIndexState,
): CellValue {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return errorValue('#REF!')
	const anchorRef = `${sheet.name}!${toA1Ref(row, col)}`
	clearSpillFootprint(sheet, sheetIndex, anchorRef, changed, spillIndex)
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
	let maxCols = 1
	for (const matrixRow of matrix) {
		const len = matrixRow?.length ?? 0
		if (len > maxCols) maxCols = len
	}
	const spillRef = `${toA1Ref(row, col)}:${toA1Ref(row + matrix.length - 1, col + maxCols - 1)}`

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
	recordSpillCell(spillIndex, sheetIndex, sheet, anchorRef, row, col)
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
			recordSpillCell(spillIndex, sheetIndex, sheet, anchorRef, targetRow, targetCol)
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
	const changed: string[] = []
	const errors: Array<{ ref: string; error: AscendError }> = []
	const spillIndex: SpillIndexState = {
		bySheet: new Map(),
		initializedSheets: new Set(),
	}
	const exactLookupCache: ExactLookupCache = new Map()
	const lookupVectorCache: LookupVectorCache = new Map()

	const analysis = analyzeWorkbook(workbook, opts?.range ? { range: opts.range } : undefined)
	const graph = analysis.dependencyGraph
	const isDirtyRecalc = opts?.dirtyOnly || (opts?.dirtyRefs?.length ?? 0) > 0

	let evalOrder: CellKey[]
	let dirtySeeds: CellKey[] = []

	if (isDirtyRecalc) {
		dirtySeeds = [
			...graph.getVolatiles(),
			...resolveDirtyKeys(workbook, analysis.sheetNameIndex, opts?.dirtyRefs),
		]
		const dirty = graph.getDirtySet(dirtySeeds)
		evalOrder = graph.getEvalOrder(dirty)
	} else {
		const allKeys = graph.getAllFormulaCells()
		const allSet = new Set(allKeys)
		evalOrder = graph.getEvalOrder(allSet)
	}

	const cycleKeys = analysis.cycleKeys
	const volatileKeys = new Set(graph.getVolatiles())
	const mustEval: Set<CellKey> = new Set()
	if (isDirtyRecalc) {
		for (const seed of dirtySeeds) {
			mustEval.add(seed)
			for (const dep of graph.getDependents(seed)) {
				mustEval.add(dep)
			}
		}
	}
	const asts = new Map<CellKey, FormulaNode>()
	if (isDirtyRecalc) {
		for (const key of evalOrder) {
			const analyzed = analysis.formulas.get(key)
			if (!analyzed) continue
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
			asts.set(key, analyzed.ast)
		}
	} else {
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
	}

	if (ctx.iterativeCalc.enabled && cycleKeys.size > 0) {
		evalIterative(
			workbook,
			ctx,
			graph,
			asts,
			evalOrder,
			cycleKeys,
			changed,
			errors,
			spillIndex,
			exactLookupCache,
			lookupVectorCache,
		)
	} else {
		const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
		const mutableCtx = new MutableEvalContext()
		mutableCtx.workbook = workbook
		mutableCtx.calcContext = ctx
		mutableCtx.exactLookupCache = exactLookupCache
		mutableCtx.lookupVectorCache = lookupVectorCache

		const sharedGroups = getSharedFormulaGroups(workbook, analysis.formulas)
		const evalOrderIndex = new Map<CellKey, number>()
		let idx = 0
		for (const k of evalOrder) {
			evalOrderIndex.set(k, idx++)
		}
		for (const members of sharedGroups.values()) {
			members.sort((a, b) => (evalOrderIndex.get(a) ?? -1) - (evalOrderIndex.get(b) ?? -1))
		}
		const cellToGroup = new Map<CellKey, string>()
		for (const [gk, members] of sharedGroups) {
			if (members.length < 2) continue
			for (const member of members) cellToGroup.set(member, gk)
		}
		const processed = new Set<CellKey>()

		const evalCell = (key: CellKey) => {
			if (cycleKeys.has(key)) {
				parseCellKeyInto(key, coords)
				const { sheetIndex: si, row, col } = coords
				const sheet = workbook.sheets[si]
				if (sheet) {
					const oldCell = sheet.cells.get(row, col)
					const clearedSpill =
						oldCell?.formulaInfo &&
						isSpillBinding(oldCell.formulaInfo) &&
						oldCell.formulaInfo.isAnchor
							? clearSpillFootprint(sheet, si, oldCell.formulaInfo.anchorRef, changed, spillIndex)
							: false
					const newValue = errorValue('#REF!')
					if (!oldCell || clearedSpill || !valuesEqual(oldCell.value, newValue)) {
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
				if (isDirtyRecalc) {
					for (const dep of graph.getDependents(key)) {
						mustEval.add(dep)
					}
				}
				return
			}

			if (isDirtyRecalc && !mustEval.has(key) && !volatileKeys.has(key)) {
				return
			}

			const ast = asts.get(key)
			if (!ast) return

			parseCellKeyInto(key, coords)
			const { sheetIndex: si, row, col } = coords
			const sheet = workbook.sheets[si]
			if (!sheet) return

			mutableCtx.sheetIndex = si
			mutableCtx.row = row
			mutableCtx.col = col
			const newValue = evalFormula(key, ast, mutableCtx)
			const oldCell = sheet.cells.get(row, col)
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				applyArrayResult(workbook, si, row, col, oldCell, spillMatrix, changed, spillIndex)
				if (isDirtyRecalc) {
					for (const dep of graph.getDependents(key)) {
						mustEval.add(dep)
					}
				}
				return
			}
			const clearedSpill =
				oldCell?.formulaInfo && isSpillBinding(oldCell.formulaInfo) && oldCell.formulaInfo.isAnchor
					? clearSpillFootprint(sheet, si, oldCell.formulaInfo.anchorRef, changed, spillIndex)
					: false
			const valueChanged = !oldCell || clearedSpill || !valuesEqual(oldCell.value, newValue)
			if (valueChanged) {
				sheet.cells.set(row, col, {
					value: newValue,
					formula: oldCell?.formula ?? null,
					styleId: oldCell?.styleId ?? (0 as StyleId),
				})
				changed.push(cellRefString(workbook, si, row, col))
				if (isDirtyRecalc) {
					for (const dep of graph.getDependents(key)) {
						mustEval.add(dep)
					}
				}
			}
		}

		for (const key of evalOrder) {
			if (processed.has(key)) continue
			const groupKey = cellToGroup.get(key)
			if (groupKey !== undefined) {
				const groupMembers = sharedGroups.get(groupKey)
				if (!groupMembers) continue
				for (const memberKey of groupMembers) processed.add(memberKey)
				for (const memberKey of groupMembers) evalCell(memberKey)
				continue
			}
			evalCell(key)
		}
	}

	return {
		changed,
		errors,
		duration: performance.now() - start,
	}
}

function resolveDirtyKeys(
	workbook: Workbook,
	sheetNameIndex: ReadonlyMap<string, number>,
	refs: readonly string[] | undefined,
): CellKey[] {
	if (!refs || refs.length === 0) return []
	const keys: CellKey[] = []
	for (const ref of refs) {
		const bang = ref.indexOf('!')
		const sheetName =
			bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : workbook.sheets[0]?.name
		const localRef = bang >= 0 ? ref.slice(bang + 1) : ref
		if (!sheetName || !localRef) continue
		const sheetIndex = sheetNameIndex.get(sheetName.toLowerCase()) ?? -1
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
	cycleKeys: ReadonlySet<CellKey>,
	changed: string[],
	_errors: Array<{ ref: string; error: AscendError }>,
	spillIndex: SpillIndexState,
	exactLookupCache: ExactLookupCache,
	lookupVectorCache: LookupVectorCache,
): void {
	const maxIter = ctx.iterativeCalc.maxIterations
	const maxChange = ctx.iterativeCalc.maxChange

	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	const mutableCtx = new MutableEvalContext()
	mutableCtx.workbook = workbook
	mutableCtx.calcContext = ctx
	mutableCtx.exactLookupCache = exactLookupCache
	mutableCtx.lookupVectorCache = lookupVectorCache
	for (let iter = 0; iter < maxIter; iter++) {
		let maxDelta = 0
		for (const key of evalOrder) {
			if (!cycleKeys.has(key)) continue
			const ast = asts.get(key)
			if (!ast) continue

			parseCellKeyInto(key, coords)
			const { sheetIndex: si, row, col } = coords
			const sheet = workbook.sheets[si]
			if (!sheet) continue

			mutableCtx.sheetIndex = si
			mutableCtx.row = row
			mutableCtx.col = col
			const newValue = evalFormula(key, ast, mutableCtx)
			const oldCell = sheet.cells.get(row, col)
			const oldValue = oldCell?.value ?? EMPTY
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				applyArrayResult(workbook, si, row, col, oldCell, spillMatrix, changed, spillIndex)
				continue
			}
			const clearedSpill =
				oldCell?.formulaInfo && isSpillBinding(oldCell.formulaInfo) && oldCell.formulaInfo.isAnchor
					? clearSpillFootprint(sheet, si, oldCell.formulaInfo.anchorRef, changed, spillIndex)
					: false

			if (oldValue.kind === 'number' && newValue.kind === 'number') {
				maxDelta = Math.max(maxDelta, Math.abs(newValue.value - oldValue.value))
			}

			if (!oldCell || clearedSpill || !valuesEqual(oldCell.value, newValue)) {
				sheet.cells.set(row, col, {
					value: newValue,
					formula: oldCell?.formula ?? null,
					styleId: oldCell?.styleId ?? (0 as StyleId),
				})
			}
		}

		if (maxDelta <= maxChange) break
	}

	for (const key of cycleKeys) {
		parseCellKeyInto(key, coords)
		changed.push(cellRefString(workbook, coords.sheetIndex, coords.row, coords.col))
	}

	for (const key of evalOrder) {
		if (cycleKeys.has(key)) continue
		const ast = asts.get(key)
		if (!ast) continue

		parseCellKeyInto(key, coords)
		const { sheetIndex: si, row, col } = coords
		const sheet = workbook.sheets[si]
		if (!sheet) continue

		mutableCtx.sheetIndex = si
		mutableCtx.row = row
		mutableCtx.col = col
		const newValue = evalFormula(key, ast, mutableCtx)
		const oldCell = sheet.cells.get(row, col)
		const spillMatrix = toScalarMatrix(newValue)
		if (spillMatrix) {
			applyArrayResult(workbook, si, row, col, oldCell, spillMatrix, changed, spillIndex)
			continue
		}
		const clearedSpill =
			oldCell?.formulaInfo && isSpillBinding(oldCell.formulaInfo) && oldCell.formulaInfo.isAnchor
				? clearSpillFootprint(sheet, si, oldCell.formulaInfo.anchorRef, changed, spillIndex)
				: false
		if (!oldCell || clearedSpill || !valuesEqual(oldCell.value, newValue)) {
			sheet.cells.set(row, col, {
				value: newValue,
				formula: oldCell?.formula ?? null,
				styleId: oldCell?.styleId ?? (0 as StyleId),
			})
			changed.push(cellRefString(workbook, si, row, col))
		}
	}
}
