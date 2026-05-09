import type { StyleId } from '@ascend/core'
import {
	DEFAULT_STYLE_ID,
	indexToColumn,
	parseA1,
	parseRange,
	type RangeRef,
	type Workbook,
} from '@ascend/core'
import {
	type AggregateRangeCache,
	clearCriteriaMatchCache,
	type ExactLookupCache,
	type FormulaNode,
	type LookupVectorCache,
	type NumericVectorCache,
} from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import { EMPTY, errorValue, numberValue, topLeftScalar, valuesEqual } from '@ascend/schema'
import {
	type AnalyzedFormula,
	analyzeWorkbook,
	type GrowingAggregateAppendIndex,
	type GrowingRangeAggregateFunction,
	invalidateWorkbookAnalysis,
} from './analysis.ts'
import type { CalcContext } from './calc-context.ts'
import { type CodegenFn, codegenFormula, codegenSharedFormula } from './codegen.ts'
import { type CompiledFormula, compileFormula, evaluateCompiled } from './compiled-eval.ts'
import {
	type CellCoords,
	type CellKey,
	cellKey,
	type DependencyGraph,
	parseCellKeyInto,
} from './dep-graph.ts'
import {
	clearRangeValueCache,
	type EvalContext,
	evaluate,
	MutableEvalContext,
	setRangeValueCache,
} from './evaluator.ts'

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

interface GrowingRangeAggregateOptimization {
	readonly functionName: GrowingRangeAggregateFunction
	readonly previousKey: CellKey
	readonly previousSheetIndex: number
	readonly previousRow: number
	readonly previousCol: number
	readonly appendSheetIndex: number
	readonly appendStartRow: number
	readonly appendStartCol: number
	readonly appendEndRow: number
	readonly appendEndCol: number
}

interface RangeAggregateOptimization {
	readonly functionName: GrowingRangeAggregateFunction
	readonly sheetIndex: number
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

interface RangeAggregateState {
	readonly sum: number
	readonly count: number
	readonly min: number
	readonly max: number
	readonly error: CellValue | null
}

interface RecalcScratch {
	readonly spillBySheet: Map<number, Map<string, SpillEntry[]>>
	readonly spillInitializedSheets: Set<number>
	readonly exactLookupCache: ExactLookupCache
	readonly lookupVectorCache: LookupVectorCache
	readonly aggregateRangeCache: AggregateRangeCache
	readonly numericVectorCache: NumericVectorCache
	readonly rangeValueCache: Map<number, readonly (readonly CellValue[])[]>
	readonly growingAggregateStateCache: Map<CellKey, RangeAggregateState>
	readonly evalContext: MutableEvalContext
}

const recalcScratchByWorkbook = new WeakMap<Workbook, RecalcScratch>()

function getRecalcScratch(workbook: Workbook): RecalcScratch {
	let scratch = recalcScratchByWorkbook.get(workbook)
	if (!scratch) {
		scratch = {
			spillBySheet: new Map(),
			spillInitializedSheets: new Set(),
			exactLookupCache: new Map(),
			lookupVectorCache: new Map(),
			aggregateRangeCache: new Map(),
			numericVectorCache: new Map(),
			rangeValueCache: new Map(),
			growingAggregateStateCache: new Map(),
			evalContext: new MutableEvalContext(),
		}
		recalcScratchByWorkbook.set(workbook, scratch)
	}
	scratch.spillBySheet.clear()
	scratch.spillInitializedSheets.clear()
	scratch.aggregateRangeCache.clear()
	scratch.numericVectorCache.clear()
	scratch.rangeValueCache.clear()
	return scratch
}

function scanRangeAggregateState(
	workbook: Workbook,
	functionName: GrowingRangeAggregateFunction,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
	base?: RangeAggregateState,
): RangeAggregateState | null {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return null
	let sum = base?.sum ?? 0
	let count = base?.count ?? 0
	let min = base?.min ?? Number.POSITIVE_INFINITY
	let max = base?.max ?? Number.NEGATIVE_INFINITY
	let error = base?.error ?? null
	for (let row = startRow; row <= endRow; row++) {
		for (let col = startCol; col <= endCol; col++) {
			const directNumber = sheet.cells.readNumber(row, col)
			if (directNumber !== null) {
				switch (functionName) {
					case 'SUM':
						sum += directNumber
						break
					case 'COUNT':
						count++
						break
					case 'AVERAGE':
						sum += directNumber
						count++
						break
					case 'MIN':
						count++
						if (directNumber < min) min = directNumber
						break
					case 'MAX':
						count++
						if (directNumber > max) max = directNumber
						break
				}
				continue
			}
			const scalar = topLeftScalar(sheet.cells.readValue(row, col))
			if (scalar.kind === 'error' && functionName !== 'COUNT') {
				error = scalar
				return { sum, count, min, max, error }
			}
			const numeric =
				scalar.kind === 'number' ? scalar.value : scalar.kind === 'date' ? scalar.serial : null
			if (numeric === null) continue
			switch (functionName) {
				case 'SUM':
					sum += numeric
					break
				case 'COUNT':
					count++
					break
				case 'AVERAGE':
					sum += numeric
					count++
					break
				case 'MIN':
					count++
					if (numeric < min) min = numeric
					break
				case 'MAX':
					count++
					if (numeric > max) max = numeric
					break
			}
		}
	}
	return { sum, count, min, max, error }
}

function rangeAggregateStateToValue(
	functionName: GrowingRangeAggregateFunction,
	state: RangeAggregateState,
): CellValue {
	if (state.error) return state.error
	switch (functionName) {
		case 'SUM':
			return numberValue(state.sum)
		case 'COUNT':
			return numberValue(state.count)
		case 'AVERAGE':
			return state.count === 0 ? errorValue('#DIV/0!') : numberValue(state.sum / state.count)
		case 'MIN':
			return numberValue(state.count === 0 ? 0 : state.min)
		case 'MAX':
			return numberValue(state.count === 0 ? 0 : state.max)
	}
}

function needsRangeAggregateState(functionName: GrowingRangeAggregateFunction): boolean {
	return functionName === 'AVERAGE' || functionName === 'MIN' || functionName === 'MAX'
}

function tryEvaluateGrowingRangeScalarAggregate(
	workbook: Workbook,
	optimization: GrowingRangeAggregateOptimization,
	canUsePreviousValue: boolean,
): CellValue | null {
	if (
		optimization.functionName !== 'SUM' &&
		optimization.functionName !== 'COUNT' &&
		optimization.functionName !== 'MIN' &&
		optimization.functionName !== 'MAX'
	) {
		return null
	}
	if (!canUsePreviousValue) return null
	const previousNumber = workbook.sheets[optimization.previousSheetIndex]?.cells.readNumber(
		optimization.previousRow,
		optimization.previousCol,
	)
	let value: number
	if (previousNumber !== null && previousNumber !== undefined) {
		value = previousNumber
	} else {
		const previousValue = workbook.sheets[optimization.previousSheetIndex]?.cells.readValue(
			optimization.previousRow,
			optimization.previousCol,
		)
		if (!previousValue) return null
		const previousScalar = topLeftScalar(previousValue)
		if (previousScalar.kind === 'error') return previousScalar
		if (previousScalar.kind !== 'number') return null
		value = previousScalar.value
	}
	const appendSheet = workbook.sheets[optimization.appendSheetIndex]
	if (!appendSheet) return null
	for (let row = optimization.appendStartRow; row <= optimization.appendEndRow; row++) {
		for (let col = optimization.appendStartCol; col <= optimization.appendEndCol; col++) {
			const directNumber = appendSheet.cells.readNumber(row, col)
			if (directNumber !== null) {
				switch (optimization.functionName) {
					case 'SUM':
						value += directNumber
						break
					case 'COUNT':
						value += 1
						break
					case 'MIN':
						if (value === 0 && directNumber > 0) return null
						value = Math.min(value, directNumber)
						break
					case 'MAX':
						if (value === 0 && directNumber < 0) return null
						value = Math.max(value, directNumber)
						break
				}
				continue
			}
			const appendScalar = topLeftScalar(appendSheet.cells.readValue(row, col))
			if (appendScalar.kind === 'error') {
				if (optimization.functionName !== 'COUNT') return appendScalar
				continue
			}
			const appendNumeric =
				appendScalar.kind === 'number'
					? appendScalar.value
					: appendScalar.kind === 'date'
						? appendScalar.serial
						: null
			if (appendNumeric === null) continue
			switch (optimization.functionName) {
				case 'SUM':
					value += appendNumeric
					break
				case 'COUNT':
					value += 1
					break
				case 'MIN':
					if (value === 0 && appendNumeric > 0) return null
					value = Math.min(value, appendNumeric)
					break
				case 'MAX':
					if (value === 0 && appendNumeric < 0) return null
					value = Math.max(value, appendNumeric)
					break
			}
		}
	}
	return numberValue(value)
}

function tryEvaluateGrowingRangeAggregate(
	workbook: Workbook,
	key: CellKey,
	optimization: GrowingRangeAggregateOptimization,
	rangeAggregateStates: Map<CellKey, RangeAggregateState>,
): CellValue | null {
	const previousState = rangeAggregateStates.get(optimization.previousKey)
	if (!previousState) return null
	const state = scanRangeAggregateState(
		workbook,
		optimization.functionName,
		optimization.appendSheetIndex,
		optimization.appendStartRow,
		optimization.appendStartCol,
		optimization.appendEndRow,
		optimization.appendEndCol,
		previousState,
	)
	if (!state) return null
	rangeAggregateStates.set(key, state)
	return rangeAggregateStateToValue(optimization.functionName, state)
}

function seedPreviousRangeAggregateState(
	workbook: Workbook,
	optimization: GrowingRangeAggregateOptimization,
	rangeAggregates: ReadonlyMap<CellKey, RangeAggregateOptimization>,
	formulas: ReadonlyMap<
		CellKey,
		{ readonly rangeAggregate?: RangeAggregateOptimization | undefined }
	>,
	rangeAggregateStates: Map<CellKey, RangeAggregateState>,
): boolean {
	if (rangeAggregateStates.has(optimization.previousKey)) return true
	const previous =
		rangeAggregates.get(optimization.previousKey) ??
		formulas.get(optimization.previousKey)?.rangeAggregate
	if (!previous || previous.functionName !== optimization.functionName) return false
	const state = scanRangeAggregateState(
		workbook,
		previous.functionName,
		previous.sheetIndex,
		previous.startRow,
		previous.startCol,
		previous.endRow,
		previous.endCol,
	)
	if (!state) return false
	rangeAggregateStates.set(optimization.previousKey, state)
	return true
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

interface SharedFormulaPlan {
	readonly members: readonly CellKey[]
	readonly evaluator: CodegenFn | null
}

function parseSharedAnchorRef(
	masterRef: string | undefined,
	fallbackRow: number,
	fallbackCol: number,
): { row: number; col: number } {
	if (!masterRef) return { row: fallbackRow, col: fallbackCol }
	try {
		const parsed = parseA1(masterRef)
		return { row: parsed.row, col: parsed.col }
	} catch {
		return { row: fallbackRow, col: fallbackCol }
	}
}

const compiledCache = new WeakMap<FormulaNode, CompiledFormula | false>()

function shouldPreferCompiled(ast: FormulaNode): boolean {
	return ast.type === 'function'
		? ast.name.toUpperCase() === 'IF' ||
				ast.name.toUpperCase() === 'IFERROR' ||
				ast.name.toUpperCase() === 'IFNA'
		: false
}

const ARRAY_RETURNING_FUNCTIONS = new Set([
	'CHOOSECOLS',
	'CHOOSEROWS',
	'DROP',
	'EXPAND',
	'FILTER',
	'HSTACK',
	'MAKEARRAY',
	'MAP',
	'RANDARRAY',
	'SCAN',
	'SEQUENCE',
	'SORT',
	'SORTBY',
	'TAKE',
	'TEXTSPLIT',
	'TOCOL',
	'TOROW',
	'TRANSPOSE',
	'UNIQUE',
	'VSTACK',
	'WRAPCOLS',
	'WRAPROWS',
])

function needsInterpreterArraySemantics(ast: FormulaNode): boolean {
	return (
		ast.type === 'binary' &&
		ast.op !== ',' &&
		ast.op !== ' ' &&
		(nodeCanReturnArray(ast.left) || nodeCanReturnArray(ast.right))
	)
}

function nodeCanReturnArray(node: FormulaNode): boolean {
	switch (node.type) {
		case 'array':
		case 'rangeRef':
		case 'wholeColumnRange':
		case 'wholeRowRange':
		case 'structuredRef':
		case 'spillRef':
		case 'sheetSpanRef':
			return true
		case 'binary':
			return nodeCanReturnArray(node.left) || nodeCanReturnArray(node.right)
		case 'unary':
			return nodeCanReturnArray(node.operand)
		case 'function':
			return ARRAY_RETURNING_FUNCTIONS.has(node.name.toUpperCase())
		default:
			return false
	}
}

export function clearCompiledFormulaCache(): void {
	// WeakMap entries are reclaimed automatically; this is a no-op placeholder
	// kept for API symmetry with clearFormulaParseCache.
}

function evalFormula(
	_key: CellKey,
	formulaText: string,
	ast: FormulaNode,
	ctx: EvalContext,
): CellValue {
	if (needsInterpreterArraySemantics(ast)) return evaluate(ast, ctx)
	if (shouldPreferCompiled(ast)) {
		let compiled = compiledCache.get(ast)
		if (compiled === undefined) {
			const result = compileFormula(ast)
			compiled = result ?? false
			compiledCache.set(ast, compiled)
		}
		if (compiled !== false) {
			return evaluateCompiled(compiled, ctx)
		}
	}
	const generated = codegenFormula(formulaText, ast)
	if (generated) return generated(ctx)
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
			const existingFormulaInfo = sheet.cells.readFormulaInfo(targetRow, targetCol)
			if (existingFormulaInfo === undefined && !sheet.cells.has(targetRow, targetCol)) continue
			if (
				existingFormulaInfo &&
				isSpillBinding(existingFormulaInfo) &&
				existingFormulaInfo.anchorRef === anchorRef
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
	oldFormula: string | null,
	oldStyleId: StyleId,
	matrix: readonly (readonly CellValue[])[],
	changed: string[],
	spillIndex: SpillIndexState,
): CellValue {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return errorValue('#REF!')
	const anchorRef = `${sheet.name}!${toA1Ref(row, col)}`
	if (matrix.length === 0 || (matrix[0]?.length ?? 0) === 0) return EMPTY

	const anchorValue = topLeftValue(matrix[0]?.[0] ?? EMPTY)
	let maxCols = 1
	for (const matrixRow of matrix) {
		const len = matrixRow?.length ?? 0
		if (len > maxCols) maxCols = len
	}
	const spillRef = `${toA1Ref(row, col)}:${toA1Ref(row + matrix.length - 1, col + maxCols - 1)}`
	if (spillMatchesMatrix(sheet, sheetIndex, row, col, anchorRef, spillRef, matrix, spillIndex)) {
		return anchorValue
	}
	clearSpillFootprint(sheet, sheetIndex, anchorRef, changed, spillIndex)
	if (isSpillBlocked(sheet, row, col, anchorRef, matrix)) {
		const spillError = errorValue('#SPILL!')
		const currentValue = sheet.cells.readValue(row, col)
		const currentBinding = sheet.cells.readFormulaInfo(row, col)
		const alreadyBlocked = currentBinding === undefined && valuesEqual(currentValue, spillError)
		if (!alreadyBlocked) {
			sheet.cells.setResolved(row, col, spillError, oldFormula, oldStyleId)
			changed.push(anchorRef)
		}
		return spillError
	}

	sheet.cells.setResolved(row, col, anchorValue, oldFormula, oldStyleId, {
		kind: 'spill',
		anchorRef,
		ref: spillRef,
		isAnchor: true,
	})
	recordSpillCell(spillIndex, sheetIndex, sheet, anchorRef, row, col)
	changed.push(anchorRef)

	for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) {
		const sourceRow = matrix[rowOffset] ?? []
		for (let colOffset = 0; colOffset < sourceRow.length; colOffset++) {
			if (rowOffset === 0 && colOffset === 0) continue
			const targetRow = row + rowOffset
			const targetCol = col + colOffset
			sheet.cells.setResolved(
				targetRow,
				targetCol,
				topLeftValue(sourceRow[colOffset] ?? EMPTY),
				null,
				oldStyleId,
				{
					kind: 'spill',
					anchorRef,
					ref: spillRef,
					isAnchor: false,
				},
			)
			recordSpillCell(spillIndex, sheetIndex, sheet, anchorRef, targetRow, targetCol)
			changed.push(`${sheet.name}!${toA1Ref(targetRow, targetCol)}`)
		}
	}
	return anchorValue
}

function spillMatchesMatrix(
	sheet: Workbook['sheets'][number],
	sheetIndex: number,
	row: number,
	col: number,
	anchorRef: string,
	spillRef: string,
	matrix: readonly (readonly CellValue[])[],
	spillIndex: SpillIndexState,
): boolean {
	const existing = getSheetSpillIndex(spillIndex, sheetIndex, sheet).get(anchorRef)
	if (!existing || existing.length === 0) return false
	const expected = new Map<string, { value: CellValue; isAnchor: boolean }>()
	for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) {
		const sourceRow = matrix[rowOffset] ?? []
		for (let colOffset = 0; colOffset < sourceRow.length; colOffset++) {
			const targetRow = row + rowOffset
			const targetCol = col + colOffset
			expected.set(toA1Ref(targetRow, targetCol), {
				value: topLeftValue(sourceRow[colOffset] ?? EMPTY),
				isAnchor: rowOffset === 0 && colOffset === 0,
			})
		}
	}
	if (existing.length !== expected.size) return false
	for (const entry of existing) {
		const match = expected.get(entry.ref)
		if (!match) return false
		const value = sheet.cells.readValue(entry.row, entry.col)
		if (!valuesEqual(value, match.value)) return false
		const binding = sheet.cells.readFormulaInfo(entry.row, entry.col)
		if (!binding || !isSpillBinding(binding)) return false
		if (
			binding.anchorRef !== anchorRef ||
			binding.ref !== spillRef ||
			binding.isAnchor !== match.isAnchor
		) {
			return false
		}
	}
	return true
}

function toA1Ref(row: number, col: number): string {
	return `${indexToColumn(col)}${row + 1}`
}

function clearOrphanedSpills(
	workbook: Workbook,
	spillIndex: SpillIndexState,
	changed: string[],
): void {
	const sheetIndexByName = new Map<string, number>()
	for (let i = 0; i < workbook.sheets.length; i++) {
		const s = workbook.sheets[i]
		if (s) sheetIndexByName.set(s.name.toLowerCase(), i)
	}
	for (let si = 0; si < workbook.sheets.length; si++) {
		const sheet = workbook.sheets[si]
		if (!sheet) continue
		const entries = getSheetSpillIndex(spillIndex, si, sheet)
		for (const anchorRef of [...entries.keys()]) {
			const bang = anchorRef.lastIndexOf('!')
			if (bang < 0) continue
			const sheetName = anchorRef.slice(0, bang).replace(/^'|'$/g, '')
			const cellPart = anchorRef.slice(bang + 1)
			const anchorIdx = sheetIndexByName.get(sheetName.toLowerCase())
			const anchorSheet = anchorIdx !== undefined ? workbook.sheets[anchorIdx] : undefined
			if (!anchorSheet) continue
			let anchorRow: number
			let anchorCol: number
			try {
				const parsed = parseA1(cellPart)
				anchorRow = parsed.row
				anchorCol = parsed.col
			} catch {
				continue
			}
			if (anchorSheet.cells.readFormula(anchorRow, anchorCol) != null) continue
			const preserveAnchor = anchorSheet.cells.has(anchorRow, anchorCol)
			if (!preserveAnchor) {
				anchorSheet.cells.delete(anchorRow, anchorCol)
				changed.push(anchorRef)
			}
			const removals = entries.get(anchorRef)
			if (removals && removals.length > 0) {
				entries.delete(anchorRef)
				for (const removal of removals) {
					if (preserveAnchor && removal.row === anchorRow && removal.col === anchorCol) continue
					anchorSheet.cells.delete(removal.row, removal.col)
					changed.push(`${anchorSheet.name}!${removal.ref}`)
				}
				invalidateWorkbookAnalysis(workbook)
			}
		}
	}
}

/**
 * Recalculate all formula cells, or an incremental subset when `dirtyOnly` / `dirtyRefs` is set.
 * Incremental mode uses the dependency graph (`getDirtySet` + partial topological order) so small
 * edits avoid rebuilding global eval order over the entire workbook when possible.
 */
export function recalculate(
	workbook: Workbook,
	ctx: CalcContext,
	opts?: { dirtyOnly?: boolean; range?: RangeRef; dirtyRefs?: readonly string[] },
): RecalcResult {
	const start = performance.now()
	const changed: string[] = []
	const errors: Array<{ ref: string; error: AscendError }> = []
	const scratch = getRecalcScratch(workbook)
	const spillIndex: SpillIndexState = {
		bySheet: scratch.spillBySheet,
		initializedSheets: scratch.spillInitializedSheets,
	}
	const exactLookupCache = scratch.exactLookupCache
	const lookupVectorCache = scratch.lookupVectorCache
	const aggregateRangeCache = scratch.aggregateRangeCache
	const numericVectorCache = scratch.numericVectorCache
	const mutableCtx = scratch.evalContext
	mutableCtx.workbook = workbook
	mutableCtx.calcContext = ctx
	mutableCtx.exactLookupCache = exactLookupCache
	mutableCtx.lookupVectorCache = lookupVectorCache
	mutableCtx.aggregateRangeCache = aggregateRangeCache
	mutableCtx.numericVectorCache = numericVectorCache
	setRangeValueCache(scratch.rangeValueCache)
	const isDirtyRecalc = opts?.dirtyOnly || (opts?.dirtyRefs?.length ?? 0) > 0

	if (!isDirtyRecalc) clearOrphanedSpills(workbook, spillIndex, changed)

	const analysis = analyzeWorkbook(workbook, opts?.range ? { range: opts.range } : undefined)
	const graph = analysis.dependencyGraph
	// Parallelism note: graph.getIndependentSubgraphs() returns connected components; cells in
	// different subgraphs could be evaluated in parallel. The evaluator mutates workbook cell
	// values during evaluation, so true parallelism would require thread-safe writes or a
	// different architecture (e.g. immutable snapshots per subgraph).
	const volatileKeysList = graph.getVolatiles()
	const dirtyRefKeys = isDirtyRecalc
		? resolveDirtyKeys(workbook, analysis.sheetNameIndex, opts?.dirtyRefs)
		: []
	const dirtyRefsCanUnblockSpill =
		isDirtyRecalc && dirtyRefsMayUnblockSpill(workbook, analysis.sheetNameIndex, opts?.dirtyRefs)
	if (isDirtyRecalc) {
		invalidateLookupCachesForDirtyKeys(exactLookupCache, lookupVectorCache, dirtyRefKeys)
	} else {
		exactLookupCache.clear()
		lookupVectorCache.clear()
		scratch.growingAggregateStateCache.clear()
	}
	if (isDirtyRecalc && volatileKeysList.length === 0 && !dirtyRefsCanUnblockSpill) {
		const fast =
			tryFastDirtyGrowingAggregateRecalc(
				workbook,
				graph,
				analysis.formulas,
				analysis.growingAggregateAppendIndex,
				scratch.growingAggregateStateCache,
				dirtyRefKeys,
				start,
			) ??
			tryFastDirtyPrefixAggregateRecalc(workbook, graph, analysis.formulas, dirtyRefKeys, start)
		if (fast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return fast
		}
	}

	let evalOrder: CellKey[]
	let dirtySeeds: CellKey[] = []

	if (isDirtyRecalc) {
		dirtySeeds = [...volatileKeysList, ...dirtyRefKeys]
		if (dirtyRefsCanUnblockSpill) {
			dirtySeeds.push(
				...resolveBlockedSpillKeys(
					workbook,
					analysis.formulas,
					analysis.sheetNameIndex,
					opts?.dirtyRefs,
				),
			)
		}
		const dirty = graph.getDirtySet(dirtySeeds)
		evalOrder = graph.getEvalOrder(dirty)
	} else {
		const allKeys = graph.getAllFormulaCells()
		const allSet = new Set(allKeys)
		evalOrder = graph.getEvalOrder(allSet)
	}

	const cycleKeys = analysis.cycleKeys
	const volatileKeys = volatileKeysList.length > 0 ? new Set(volatileKeysList) : null
	const mustEval: Set<CellKey> = new Set()
	const rangeAggregates = new Map<CellKey, RangeAggregateOptimization>()
	const growingRangeAggregates = new Map<CellKey, GrowingRangeAggregateOptimization>()
	if (isDirtyRecalc) {
		for (const seed of dirtySeeds) {
			mustEval.add(seed)
			if (!analysis.formulas.has(seed)) {
				for (const dep of graph.getDependents(seed)) {
					mustEval.add(dep)
				}
			}
		}
	}
	const asts = new Map<CellKey, FormulaNode>()
	const formulaTexts = new Map<CellKey, string>()
	const evalOrderSet = new Set(evalOrder)
	const handleParseError = (analyzed: AnalyzedFormula) => {
		if (!analyzed.parseError) return
		errors.push({
			ref: cellRefString(workbook, analyzed.sheetIndex, analyzed.row, analyzed.col),
			error: {
				code: 'FORMULA_PARSE_ERROR',
				message: `Failed to parse: ${analyzed.formula}`,
				retryable: false,
			},
		})
		const sheet = workbook.sheets[analyzed.sheetIndex]
		if (!sheet) return
		const oldFormula = sheet.cells.readFormula(analyzed.row, analyzed.col) ?? null
		const oldStyleId = sheet.cells.readStyleId(analyzed.row, analyzed.col) ?? DEFAULT_STYLE_ID
		const parseErrorValue = errorValue('#VALUE!')
		const oldValue = sheet.cells.readValue(analyzed.row, analyzed.col)
		if (!valuesEqual(oldValue, parseErrorValue)) {
			sheet.cells.setResolved(analyzed.row, analyzed.col, parseErrorValue, oldFormula, oldStyleId)
			changed.push(cellRefString(workbook, analyzed.sheetIndex, analyzed.row, analyzed.col))
		}
	}
	if (isDirtyRecalc) {
		for (const key of evalOrder) {
			const analyzed = analysis.formulas.get(key)
			if (analyzed) handleParseError(analyzed)
		}
	} else {
		for (const analyzed of analysis.formulas.values()) handleParseError(analyzed)
	}

	if (isDirtyRecalc) {
		for (const key of evalOrder) {
			const analyzed = analysis.formulas.get(key)
			if (!analyzed) continue
			if (analyzed.parseError || !analyzed.ast) continue
			asts.set(key, analyzed.ast)
			formulaTexts.set(key, analyzed.formula)
			if (analyzed.rangeAggregate) {
				rangeAggregates.set(key, analyzed.rangeAggregate)
			}
			if (analyzed.growingRangeAggregate) {
				growingRangeAggregates.set(key, analyzed.growingRangeAggregate)
			}
		}
	} else {
		for (const analyzed of analysis.formulas.values()) {
			if (analyzed.parseError || !analyzed.ast) continue
			asts.set(analyzed.key, analyzed.ast)
			formulaTexts.set(analyzed.key, analyzed.formula)
			if (analyzed.rangeAggregate) {
				rangeAggregates.set(analyzed.key, analyzed.rangeAggregate)
			}
			if (analyzed.growingRangeAggregate) {
				growingRangeAggregates.set(analyzed.key, analyzed.growingRangeAggregate)
			}
		}
	}

	if (ctx.iterativeCalc.enabled && cycleKeys.size > 0) {
		evalIterative(
			workbook,
			ctx,
			graph,
			asts,
			formulaTexts,
			evalOrder,
			cycleKeys,
			changed,
			errors,
			spillIndex,
			exactLookupCache,
			lookupVectorCache,
			aggregateRangeCache,
			numericVectorCache,
			mutableCtx,
		)
	} else {
		const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }

		const sharedGroups = analysis.sharedFormulaGroups
		const cellToGroup = new Map<CellKey, SharedFormulaPlan>()
		let hasSharedFormulaGroups = false
		const hasGrowingRangeAggregates = growingRangeAggregates.size > 0
		const completedKeys = hasGrowingRangeAggregates ? new Set<CellKey>() : null
		const rangeAggregateStates =
			hasGrowingRangeAggregates &&
			[...growingRangeAggregates.values()].some((aggregate) =>
				needsRangeAggregateState(aggregate.functionName),
			)
				? scratch.growingAggregateStateCache
				: null
		if (sharedGroups.size > 0) {
			const evalOrderIndex = new Map<CellKey, number>()
			let idx = 0
			for (const k of evalOrder) {
				evalOrderIndex.set(k, idx++)
			}
			for (const [_gk, members] of sharedGroups) {
				if (members.length < 2) continue
				hasSharedFormulaGroups = true
				members.sort((a, b) => (evalOrderIndex.get(a) ?? -1) - (evalOrderIndex.get(b) ?? -1))
				let evaluator: CodegenFn | null = null
				for (const memberKey of members) {
					parseCellKeyInto(memberKey, coords)
					const memberSheet = workbook.sheets[coords.sheetIndex]
					const binding = memberSheet?.cells.readFormulaInfo(coords.row, coords.col) as
						| { kind?: string; isMaster?: boolean; masterRef?: string }
						| undefined
					if (!memberSheet || binding?.kind !== 'shared' || !binding.isMaster) continue
					const masterAst = asts.get(memberKey)
					const masterFormulaText = formulaTexts.get(memberKey)
					if (!masterAst || !masterFormulaText) break
					const anchor = parseSharedAnchorRef(binding.masterRef, coords.row, coords.col)
					evaluator = codegenSharedFormula(masterFormulaText, masterAst, anchor)
					break
				}
				const plan: SharedFormulaPlan = { members, evaluator }
				for (const member of members) cellToGroup.set(member, plan)
			}
		}

		const evalCell = (key: CellKey, groupEvaluator?: CodegenFn | null) => {
			if (cycleKeys.has(key)) {
				parseCellKeyInto(key, coords)
				const { sheetIndex: si, row, col } = coords
				const sheet = workbook.sheets[si]
				if (sheet) {
					const hadCell = sheet.cells.has(row, col)
					const oldValue = sheet.cells.readValue(row, col)
					const oldFormula = sheet.cells.readFormula(row, col) ?? null
					const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
					const oldFormulaInfo = sheet.cells.readFormulaInfo(row, col)
					const clearedSpill =
						oldFormulaInfo && isSpillBinding(oldFormulaInfo) && oldFormulaInfo.isAnchor
							? clearSpillFootprint(sheet, si, oldFormulaInfo.anchorRef, changed, spillIndex)
							: false
					const newValue = errorValue('#REF!')
					if (!hadCell || clearedSpill || !valuesEqual(oldValue, newValue)) {
						sheet.cells.setResolved(row, col, newValue, oldFormula, oldStyleId)
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

			if (isDirtyRecalc && !mustEval.has(key) && !volatileKeys?.has(key)) {
				return
			}

			parseCellKeyInto(key, coords)
			const { sheetIndex: si, row, col } = coords
			const sheet = workbook.sheets[si]
			if (!sheet) return

			const growingRangeAggregate = hasGrowingRangeAggregates
				? growingRangeAggregates.get(key)
				: undefined
			let newValue: CellValue | null = null
			if (growingRangeAggregate && completedKeys) {
				const canUsePreviousValue =
					completedKeys.has(growingRangeAggregate.previousKey) ||
					(isDirtyRecalc && !evalOrderSet.has(growingRangeAggregate.previousKey))
				newValue = tryEvaluateGrowingRangeScalarAggregate(
					workbook,
					growingRangeAggregate,
					canUsePreviousValue,
				)
				if (!newValue && rangeAggregateStates) {
					if (canUsePreviousValue && !completedKeys.has(growingRangeAggregate.previousKey)) {
						seedPreviousRangeAggregateState(
							workbook,
							growingRangeAggregate,
							rangeAggregates,
							analysis.formulas,
							rangeAggregateStates,
						)
					}
					newValue = tryEvaluateGrowingRangeAggregate(
						workbook,
						key,
						growingRangeAggregate,
						rangeAggregateStates,
					)
				}
			}
			if (!newValue) {
				const ast = asts.get(key)
				if (!ast) return
				const formulaText = formulaTexts.get(key)
				if (!formulaText) return
				mutableCtx.sheetIndex = si
				mutableCtx.row = row
				mutableCtx.col = col
				newValue = groupEvaluator
					? groupEvaluator(mutableCtx)
					: evalFormula(key, formulaText, ast, mutableCtx)
				const rangeAggregate = rangeAggregateStates ? rangeAggregates.get(key) : undefined
				if (rangeAggregate && needsRangeAggregateState(rangeAggregate.functionName)) {
					const state = scanRangeAggregateState(
						workbook,
						rangeAggregate.functionName,
						rangeAggregate.sheetIndex,
						rangeAggregate.startRow,
						rangeAggregate.startCol,
						rangeAggregate.endRow,
						rangeAggregate.endCol,
					)
					if (state) rangeAggregateStates?.set(key, state)
				}
			}
			const hadCell = sheet.cells.has(row, col)
			const oldValue = sheet.cells.readValue(row, col)
			const oldFormula = sheet.cells.readFormula(row, col) ?? null
			const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
			const oldFormulaInfo = sheet.cells.readFormulaInfo(row, col)
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				const changedBefore = changed.length
				applyArrayResult(
					workbook,
					si,
					row,
					col,
					oldFormula,
					oldStyleId,
					spillMatrix,
					changed,
					spillIndex,
				)
				if (isDirtyRecalc && changed.length > changedBefore) {
					for (const dep of graph.getDependents(key)) {
						mustEval.add(dep)
					}
				}
				return
			}
			const clearedSpill =
				oldFormulaInfo && isSpillBinding(oldFormulaInfo) && oldFormulaInfo.isAnchor
					? clearSpillFootprint(sheet, si, oldFormulaInfo.anchorRef, changed, spillIndex)
					: false
			// Volatile optimization: when value unchanged (e.g. TODAY() same date), we do not add
			// dependents to mustEval, so they are skipped in evalCell and not re-evaluated.
			const valueChanged = !hadCell || clearedSpill || !valuesEqual(oldValue, newValue)
			if (valueChanged) {
				if (newValue.kind === 'number') {
					sheet.cells.setNumberResolved(row, col, newValue.value, oldFormula, oldStyleId)
				} else {
					sheet.cells.setResolved(row, col, newValue, oldFormula, oldStyleId)
				}
				changed.push(cellRefString(workbook, si, row, col))
				if (isDirtyRecalc) {
					for (const dep of graph.getDependents(key)) {
						mustEval.add(dep)
					}
				}
			}
			completedKeys?.add(key)
		}

		if (!hasSharedFormulaGroups) {
			for (const key of evalOrder) {
				evalCell(key)
			}
		} else {
			const processed = new Set<CellKey>()
			for (const key of evalOrder) {
				if (processed.has(key)) continue
				const plan = cellToGroup.get(key)
				if (plan) {
					for (const memberKey of plan.members) processed.add(memberKey)
					for (const memberKey of plan.members) evalCell(memberKey, plan.evaluator)
					continue
				}
				evalCell(key)
			}
		}
	}

	clearRangeValueCache()
	clearCriteriaMatchCache()

	return {
		changed,
		errors,
		duration: performance.now() - start,
	}
}

function tryFastDirtyPrefixAggregateRecalc(
	workbook: Workbook,
	graph: DependencyGraph,
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
	dirtyRefKeys: readonly CellKey[],
	start: number,
): RecalcResult | null {
	if (dirtyRefKeys.length !== 1) return null
	const sourceKey = dirtyRefKeys[0] as CellKey
	const sourceCoords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	parseCellKeyInto(sourceKey, sourceCoords)

	const directDependents = graph.getDependents(sourceKey)
	if (directDependents.length === 0) return null
	const affected: AnalyzedFormula[] = []
	let groupKey: string | undefined
	for (const dependent of directDependents) {
		const formula = formulas.get(dependent)
		if (!formula) return null
		const aggregate = formula.rangeAggregate
		if (!aggregate || !canFastRecalculatePrefixAggregate(aggregate, sourceCoords)) return null
		const currentGroupKey = [
			aggregate.functionName,
			aggregate.sheetIndex,
			aggregate.startRow,
			aggregate.startCol,
			aggregate.endCol,
		].join(':')
		if (groupKey === undefined) groupKey = currentGroupKey
		else if (groupKey !== currentGroupKey) return null
		if (graph.getDependents(formula.key).length > 0) return null
		affected.push(formula)
	}

	affected.sort((a, b) => {
		const ar = a.rangeAggregate?.endRow ?? a.row
		const br = b.rangeAggregate?.endRow ?? b.row
		return ar - br || a.key - b.key
	})

	const firstAggregate = affected[0]?.rangeAggregate
	if (!firstAggregate) return null
	let state: RangeAggregateState | null = null
	let previousEndRow = firstAggregate.startRow - 1
	const changed: string[] = []
	for (const formula of affected) {
		const aggregate = formula.rangeAggregate
		if (!aggregate) return null
		const sheet = workbook.sheets[formula.sheetIndex]
		if (!sheet) return null
		const oldFormulaInfo = sheet.cells.readFormulaInfo(formula.row, formula.col)
		if (oldFormulaInfo !== undefined) return null
		const oldFormula = sheet.cells.readFormula(formula.row, formula.col) ?? null
		if (oldFormula === null) return null
		if (aggregate.endRow < previousEndRow) return null
		if (aggregate.endRow > previousEndRow) {
			state = scanRangeAggregateState(
				workbook,
				aggregate.functionName,
				aggregate.sheetIndex,
				previousEndRow + 1,
				aggregate.startCol,
				aggregate.endRow,
				aggregate.endCol,
				state ?? undefined,
			)
			if (!state) return null
			previousEndRow = aggregate.endRow
		}
		if (!state) return null
		const newValue = rangeAggregateStateToValue(aggregate.functionName, state)
		const oldValue = sheet.cells.readValue(formula.row, formula.col)
		if (valuesEqual(oldValue, newValue)) continue
		const oldStyleId = sheet.cells.readStyleId(formula.row, formula.col) ?? DEFAULT_STYLE_ID
		if (newValue.kind === 'number') {
			sheet.cells.setNumberResolved(
				formula.row,
				formula.col,
				newValue.value,
				oldFormula,
				oldStyleId,
			)
		} else {
			sheet.cells.setResolved(formula.row, formula.col, newValue, oldFormula, oldStyleId)
		}
		changed.push(cellRefString(workbook, formula.sheetIndex, formula.row, formula.col))
	}
	return { changed, errors: [], duration: performance.now() - start }
}

function canFastRecalculatePrefixAggregate(
	aggregate: RangeAggregateOptimization,
	source: CellCoords,
): boolean {
	if (
		aggregate.functionName !== 'SUM' &&
		aggregate.functionName !== 'COUNT' &&
		aggregate.functionName !== 'AVERAGE' &&
		aggregate.functionName !== 'MIN' &&
		aggregate.functionName !== 'MAX'
	) {
		return false
	}
	return (
		aggregate.sheetIndex === source.sheetIndex &&
		aggregate.startCol === aggregate.endCol &&
		aggregate.startCol === source.col &&
		source.row >= aggregate.startRow &&
		source.row <= aggregate.endRow
	)
}

function tryFastDirtyGrowingAggregateRecalc(
	workbook: Workbook,
	graph: DependencyGraph,
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
	growingAggregateAppendIndex: GrowingAggregateAppendIndex,
	growingAggregateStateCache: Map<CellKey, RangeAggregateState>,
	dirtyRefKeys: readonly CellKey[],
	start: number,
): RecalcResult | null {
	if (dirtyRefKeys.length !== 1) return null
	const sourceKey = dirtyRefKeys[0] as CellKey
	const dependents = growingAggregateAppendIndex.get(sourceKey)
	if (!dependents) return null
	if (dependents.length !== 1) return null
	const formulaKey = dependents[0] as CellKey
	if (graph.getDependents(formulaKey).length > 0) return null
	const analyzed = formulas.get(formulaKey)
	const growingRangeAggregate = analyzed?.growingRangeAggregate
	if (!analyzed || analyzed.parseError || !growingRangeAggregate) return null
	if (
		growingRangeAggregate.functionName !== 'SUM' &&
		growingRangeAggregate.functionName !== 'COUNT' &&
		growingRangeAggregate.functionName !== 'AVERAGE' &&
		growingRangeAggregate.functionName !== 'MIN' &&
		growingRangeAggregate.functionName !== 'MAX'
	) {
		return null
	}
	const sourceCoords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	parseCellKeyInto(sourceKey, sourceCoords)
	if (
		sourceCoords.sheetIndex !== growingRangeAggregate.appendSheetIndex ||
		sourceCoords.row < growingRangeAggregate.appendStartRow ||
		sourceCoords.row > growingRangeAggregate.appendEndRow ||
		sourceCoords.col < growingRangeAggregate.appendStartCol ||
		sourceCoords.col > growingRangeAggregate.appendEndCol
	) {
		return null
	}
	const sheet = workbook.sheets[analyzed.sheetIndex]
	if (!sheet) return null
	const oldFormulaInfo = sheet.cells.readFormulaInfo(analyzed.row, analyzed.col)
	if (oldFormulaInfo !== undefined) return null
	const oldFormula = sheet.cells.readFormula(analyzed.row, analyzed.col) ?? null
	if (oldFormula === null) return null
	let newValue = tryEvaluateGrowingRangeScalarAggregate(workbook, growingRangeAggregate, true)
	if (!newValue && needsRangeAggregateState(growingRangeAggregate.functionName)) {
		newValue = tryEvaluateGrowingRangeAggregate(
			workbook,
			formulaKey,
			growingRangeAggregate,
			growingAggregateStateCache,
		)
	}
	if (!newValue) return null
	const oldValue = sheet.cells.readValue(analyzed.row, analyzed.col)
	const changed: string[] = []
	if (!valuesEqual(oldValue, newValue)) {
		const oldStyleId = sheet.cells.readStyleId(analyzed.row, analyzed.col) ?? DEFAULT_STYLE_ID
		if (newValue.kind === 'number') {
			sheet.cells.setNumberResolved(
				analyzed.row,
				analyzed.col,
				newValue.value,
				oldFormula,
				oldStyleId,
			)
		} else {
			sheet.cells.setResolved(analyzed.row, analyzed.col, newValue, oldFormula, oldStyleId)
		}
		changed.push(cellRefString(workbook, analyzed.sheetIndex, analyzed.row, analyzed.col))
	}
	return { changed, errors: [], duration: performance.now() - start }
}

function invalidateLookupCachesForDirtyKeys(
	exactLookupCache: ExactLookupCache,
	lookupVectorCache: LookupVectorCache,
	dirtyKeys: readonly CellKey[],
): void {
	if (dirtyKeys.length === 0) return
	invalidateLookupCache(exactLookupCache, dirtyKeys)
	invalidateLookupCache(lookupVectorCache, dirtyKeys)
}

function invalidateLookupCache(cache: Map<string, unknown>, dirtyKeys: readonly CellKey[]): void {
	for (const key of [...cache.keys()]) {
		if (lookupCacheKeyOverlapsDirtyKeys(key, dirtyKeys)) cache.delete(key)
	}
}

function lookupCacheKeyOverlapsDirtyKeys(cacheKey: string, dirtyKeys: readonly CellKey[]): boolean {
	const parts = cacheKey.split(':')
	if (parts.length !== 5) return true
	const [axis, rawSheet, rawFixed, rawStart, rawEnd] = parts
	const sheetIndex = Number(rawSheet)
	const fixed = Number(rawFixed)
	const start = Number(rawStart)
	const end = Number(rawEnd)
	if (
		(axis !== 'column' && axis !== 'row') ||
		!Number.isInteger(sheetIndex) ||
		!Number.isInteger(fixed) ||
		!Number.isInteger(start) ||
		!Number.isInteger(end)
	) {
		return true
	}
	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	for (const key of dirtyKeys) {
		parseCellKeyInto(key, coords)
		if (coords.sheetIndex !== sheetIndex) continue
		if (axis === 'column') {
			if (coords.col === fixed && coords.row >= start && coords.row <= end) return true
		} else if (coords.row === fixed && coords.col >= start && coords.col <= end) {
			return true
		}
	}
	return false
}

function resolveBlockedSpillKeys(
	workbook: Workbook,
	formulas: ReadonlyMap<CellKey, { key: CellKey; sheetIndex: number; row: number; col: number }>,
	sheetNameIndex: ReadonlyMap<string, number>,
	refs: readonly string[] | undefined,
): CellKey[] {
	if (!refs || refs.length === 0) return []
	const dirtySheets = new Set<number>()
	for (const ref of refs) {
		const bang = ref.indexOf('!')
		const sheetName =
			bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : workbook.sheets[0]?.name
		if (!sheetName) continue
		const sheetIndex = sheetNameIndex.get(sheetName.toLowerCase())
		if (sheetIndex !== undefined) dirtySheets.add(sheetIndex)
	}
	if (dirtySheets.size === 0) return []

	const blocked: CellKey[] = []
	for (const analyzed of formulas.values()) {
		if (!dirtySheets.has(analyzed.sheetIndex)) continue
		const sheet = workbook.sheets[analyzed.sheetIndex]
		if (!sheet || sheet.cells.readKind(analyzed.row, analyzed.col) !== 'error') continue
		if (sheet.cells.readError(analyzed.row, analyzed.col) === '#SPILL!') blocked.push(analyzed.key)
	}
	return blocked
}

function dirtyRefsMayUnblockSpill(
	workbook: Workbook,
	sheetNameIndex: ReadonlyMap<string, number>,
	refs: readonly string[] | undefined,
): boolean {
	if (!refs || refs.length === 0) return false
	for (const ref of refs) {
		const bang = ref.indexOf('!')
		const sheetName =
			bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : workbook.sheets[0]?.name
		const localRef = bang >= 0 ? ref.slice(bang + 1) : ref
		if (!sheetName || !localRef) continue
		const sheetIndex = sheetNameIndex.get(sheetName.toLowerCase()) ?? -1
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		const range = parseRange(localRef)
		for (let row = range.start.row; row <= range.end.row; row++) {
			for (let col = range.start.col; col <= range.end.col; col++) {
				const value = sheet.cells.readValue(row, col)
				if (!value || value.kind === 'empty') return true
			}
		}
	}
	return false
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
	formulaTexts: Map<CellKey, string>,
	evalOrder: CellKey[],
	cycleKeys: ReadonlySet<CellKey>,
	changed: string[],
	_errors: Array<{ ref: string; error: AscendError }>,
	spillIndex: SpillIndexState,
	exactLookupCache: ExactLookupCache,
	lookupVectorCache: LookupVectorCache,
	aggregateRangeCache: AggregateRangeCache,
	numericVectorCache: NumericVectorCache,
	mutableCtx: MutableEvalContext,
): void {
	const maxIter = ctx.iterativeCalc.maxIterations
	const maxChange = ctx.iterativeCalc.maxChange

	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	mutableCtx.workbook = workbook
	mutableCtx.calcContext = ctx
	mutableCtx.exactLookupCache = exactLookupCache
	mutableCtx.lookupVectorCache = lookupVectorCache
	mutableCtx.aggregateRangeCache = aggregateRangeCache
	mutableCtx.numericVectorCache = numericVectorCache
	for (let iter = 0; iter < maxIter; iter++) {
		let maxDelta = 0
		for (const key of evalOrder) {
			if (!cycleKeys.has(key)) continue
			const ast = asts.get(key)
			if (!ast) continue
			const formulaText = formulaTexts.get(key)
			if (!formulaText) continue

			parseCellKeyInto(key, coords)
			const { sheetIndex: si, row, col } = coords
			const sheet = workbook.sheets[si]
			if (!sheet) continue

			mutableCtx.sheetIndex = si
			mutableCtx.row = row
			mutableCtx.col = col
			const newValue = evalFormula(key, formulaText, ast, mutableCtx)
			const hadCell = sheet.cells.has(row, col)
			const oldValue = sheet.cells.readValue(row, col)
			const oldFormula = sheet.cells.readFormula(row, col) ?? null
			const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
			const oldFormulaInfo = sheet.cells.readFormulaInfo(row, col)
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				const changedBefore = changed.length
				applyArrayResult(
					workbook,
					si,
					row,
					col,
					oldFormula,
					oldStyleId,
					spillMatrix,
					changed,
					spillIndex,
				)
				if (changed.length === changedBefore) continue
				continue
			}
			const clearedSpill =
				oldFormulaInfo && isSpillBinding(oldFormulaInfo) && oldFormulaInfo.isAnchor
					? clearSpillFootprint(sheet, si, oldFormulaInfo.anchorRef, changed, spillIndex)
					: false

			if (oldValue.kind === 'number' && newValue.kind === 'number') {
				maxDelta = Math.max(maxDelta, Math.abs(newValue.value - oldValue.value))
			}

			if (!hadCell || clearedSpill || !valuesEqual(oldValue, newValue)) {
				sheet.cells.setResolved(row, col, newValue, oldFormula, oldStyleId)
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
		const formulaText = formulaTexts.get(key)
		if (!formulaText) continue

		parseCellKeyInto(key, coords)
		const { sheetIndex: si, row, col } = coords
		const sheet = workbook.sheets[si]
		if (!sheet) continue

		mutableCtx.sheetIndex = si
		mutableCtx.row = row
		mutableCtx.col = col
		const newValue = evalFormula(key, formulaText, ast, mutableCtx)
		const hadCell = sheet.cells.has(row, col)
		const oldValue = sheet.cells.readValue(row, col)
		const oldFormula = sheet.cells.readFormula(row, col) ?? null
		const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
		const oldFormulaInfo = sheet.cells.readFormulaInfo(row, col)
		const spillMatrix = toScalarMatrix(newValue)
		if (spillMatrix) {
			applyArrayResult(
				workbook,
				si,
				row,
				col,
				oldFormula,
				oldStyleId,
				spillMatrix,
				changed,
				spillIndex,
			)
			continue
		}
		const clearedSpill =
			oldFormulaInfo && isSpillBinding(oldFormulaInfo) && oldFormulaInfo.isAnchor
				? clearSpillFootprint(sheet, si, oldFormulaInfo.anchorRef, changed, spillIndex)
				: false
		if (!hadCell || clearedSpill || !valuesEqual(oldValue, newValue)) {
			sheet.cells.setResolved(row, col, newValue, oldFormula, oldStyleId)
			changed.push(cellRefString(workbook, si, row, col))
		}
	}
}
