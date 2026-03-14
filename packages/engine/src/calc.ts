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
} from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import { EMPTY, errorValue, numberValue, topLeftScalar, valuesEqual } from '@ascend/schema'
import { analyzeWorkbook, getSharedFormulaGroups } from './analysis.ts'
import type { CalcContext } from './calc-context.ts'
import { type CodegenFn, codegenFormula, codegenSharedFormula } from './codegen.ts'
import { type CompiledFormula, compileFormula, evaluateCompiled } from './compiled-eval.ts'
import {
	type CellCoords,
	type CellKey,
	cellKey,
	type DependencyGraph,
	parseCellKey,
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
	readonly functionName: 'SUM'
	readonly previousKey: CellKey
	readonly appendSheetIndex: number
	readonly appendStartRow: number
	readonly appendStartCol: number
	readonly appendEndRow: number
	readonly appendEndCol: number
}

interface RecalcScratch {
	readonly spillBySheet: Map<number, Map<string, SpillEntry[]>>
	readonly spillInitializedSheets: Set<number>
	readonly exactLookupCache: ExactLookupCache
	readonly lookupVectorCache: LookupVectorCache
	readonly aggregateRangeCache: AggregateRangeCache
	readonly rangeValueCache: Map<string, readonly (readonly CellValue[])[]>
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
			rangeValueCache: new Map(),
		}
		recalcScratchByWorkbook.set(workbook, scratch)
	}
	scratch.spillBySheet.clear()
	scratch.spillInitializedSheets.clear()
	scratch.exactLookupCache.clear()
	scratch.lookupVectorCache.clear()
	scratch.aggregateRangeCache.clear()
	scratch.rangeValueCache.clear()
	return scratch
}

function tryEvaluateGrowingRangeAggregate(
	workbook: Workbook,
	optimization: GrowingRangeAggregateOptimization,
	completedKeys: ReadonlySet<CellKey>,
): CellValue | null {
	if (optimization.functionName !== 'SUM') return null
	if (!completedKeys.has(optimization.previousKey)) return null
	const [prevSheetIndex, prevRow, prevCol] = parseCellKey(optimization.previousKey)
	const previousValue = workbook.sheets[prevSheetIndex]?.cells.readValue(prevRow, prevCol)
	if (!previousValue) return null
	const previousScalar = topLeftScalar(previousValue)
	if (previousScalar.kind === 'error') return previousScalar
	if (previousScalar.kind !== 'number') return null
	const appendSheet = workbook.sheets[optimization.appendSheetIndex]
	if (!appendSheet) return null
	let sum = previousScalar.value
	for (let row = optimization.appendStartRow; row <= optimization.appendEndRow; row++) {
		for (let col = optimization.appendStartCol; col <= optimization.appendEndCol; col++) {
			const appendScalar = topLeftScalar(appendSheet.cells.readValue(row, col))
			if (appendScalar.kind === 'error') return appendScalar
			if (appendScalar.kind === 'number') sum += appendScalar.value
			else if (appendScalar.kind === 'date') sum += appendScalar.serial
		}
	}
	return numberValue(sum)
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
	setRangeValueCache(scratch.rangeValueCache)

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
		dirtySeeds.push(
			...resolveBlockedSpillKeys(
				workbook,
				analysis.formulas,
				analysis.sheetNameIndex,
				opts?.dirtyRefs,
			),
		)
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
	for (const analyzed of analysis.formulas.values()) {
		if (!analyzed.parseError) continue
		if (isDirtyRecalc && !evalOrderSet.has(analyzed.key)) continue
		errors.push({
			ref: cellRefString(workbook, analyzed.sheetIndex, analyzed.row, analyzed.col),
			error: {
				code: 'FORMULA_PARSE_ERROR',
				message: `Failed to parse: ${analyzed.formula}`,
				retryable: false,
			},
		})
		const sheet = workbook.sheets[analyzed.sheetIndex]
		if (sheet) {
			const oldFormula = sheet.cells.readFormula(analyzed.row, analyzed.col) ?? null
			const oldStyleId = sheet.cells.readStyleId(analyzed.row, analyzed.col) ?? DEFAULT_STYLE_ID
			const parseErrorValue = errorValue('#VALUE!')
			const oldValue = sheet.cells.readValue(analyzed.row, analyzed.col)
			if (!valuesEqual(oldValue, parseErrorValue)) {
				sheet.cells.setResolved(analyzed.row, analyzed.col, parseErrorValue, oldFormula, oldStyleId)
				changed.push(cellRefString(workbook, analyzed.sheetIndex, analyzed.row, analyzed.col))
			}
		}
	}

	if (isDirtyRecalc) {
		for (const key of evalOrder) {
			const analyzed = analysis.formulas.get(key)
			if (!analyzed) continue
			if (analyzed.parseError || !analyzed.ast) continue
			asts.set(key, analyzed.ast)
			formulaTexts.set(key, analyzed.formula)
			if (analyzed.growingRangeAggregate) {
				growingRangeAggregates.set(key, analyzed.growingRangeAggregate)
			}
		}
	} else {
		for (const analyzed of analysis.formulas.values()) {
			if (analyzed.parseError || !analyzed.ast) continue
			asts.set(analyzed.key, analyzed.ast)
			formulaTexts.set(analyzed.key, analyzed.formula)
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
		)
	} else {
		const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
		const mutableCtx = new MutableEvalContext()
		mutableCtx.workbook = workbook
		mutableCtx.calcContext = ctx
		mutableCtx.exactLookupCache = exactLookupCache
		mutableCtx.lookupVectorCache = lookupVectorCache
		mutableCtx.aggregateRangeCache = aggregateRangeCache

		const sharedGroups = getSharedFormulaGroups(workbook, analysis.formulas)
		const cellToGroup = new Map<CellKey, SharedFormulaPlan>()
		let hasSharedFormulaGroups = false
		const hasGrowingRangeAggregates = growingRangeAggregates.size > 0
		const completedKeys = hasGrowingRangeAggregates ? new Set<CellKey>() : null
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

			if (isDirtyRecalc && !mustEval.has(key) && !volatileKeys.has(key)) {
				return
			}

			const ast = asts.get(key)
			if (!ast) return
			const formulaText = formulaTexts.get(key)
			if (!formulaText) return

			parseCellKeyInto(key, coords)
			const { sheetIndex: si, row, col } = coords
			const sheet = workbook.sheets[si]
			if (!sheet) return

			mutableCtx.sheetIndex = si
			mutableCtx.row = row
			mutableCtx.col = col
			const growingRangeAggregate =
				!isDirtyRecalc && hasGrowingRangeAggregates ? growingRangeAggregates.get(key) : undefined
			const newValue =
				growingRangeAggregate && completedKeys
					? (tryEvaluateGrowingRangeAggregate(workbook, growingRangeAggregate, completedKeys) ??
						(groupEvaluator
							? groupEvaluator(mutableCtx)
							: evalFormula(key, formulaText, ast, mutableCtx)))
					: groupEvaluator
						? groupEvaluator(mutableCtx)
						: evalFormula(key, formulaText, ast, mutableCtx)
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
			const valueChanged = !hadCell || clearedSpill || !valuesEqual(oldValue, newValue)
			if (valueChanged) {
				sheet.cells.setResolved(row, col, newValue, oldFormula, oldStyleId)
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
): void {
	const maxIter = ctx.iterativeCalc.maxIterations
	const maxChange = ctx.iterativeCalc.maxChange

	const coords: CellCoords = { sheetIndex: 0, row: 0, col: 0 }
	const mutableCtx = new MutableEvalContext()
	mutableCtx.workbook = workbook
	mutableCtx.calcContext = ctx
	mutableCtx.exactLookupCache = exactLookupCache
	mutableCtx.lookupVectorCache = lookupVectorCache
	mutableCtx.aggregateRangeCache = aggregateRangeCache
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
