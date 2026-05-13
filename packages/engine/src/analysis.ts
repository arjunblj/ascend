import { parseA1Safe, type RangeRef, type Workbook } from '@ascend/core'
import type { FormulaCellRef, FormulaNode, FormulaRef } from '@ascend/formulas'
import {
	cachedParseFormula,
	extractRefs,
	functionRegistry,
	printFormula,
	rewriteRefs,
} from '@ascend/formulas'
import { ascendError } from '@ascend/schema'
import {
	type CellKey,
	cellKey,
	DependencyGraph,
	parseCellKey,
	type RangeDependency,
} from './dep-graph.ts'
import { resolveSheetIndexByMap } from './sheet-index.ts'
import {
	formulaAstReferencesSheet,
	rewriteFormulaAstForShift,
} from './structural/formula-rewrite.ts'
import { shiftIndex } from './structural/ref-shift.ts'
import { createStructuredRefResolver, type StructuredRefResolver } from './structured-refs.ts'

export interface AnalyzedFormula {
	readonly key: CellKey
	readonly sheetIndex: number
	readonly sheetName: string
	readonly row: number
	readonly col: number
	readonly formula: string
	readonly ast?: FormulaNode
	readonly refs: readonly FormulaRef[]
	readonly deps: readonly CellKey[]
	readonly rangeDeps: readonly RangeDependency[]
	readonly volatile: boolean
	readonly rangeAggregate?:
		| {
				readonly functionName: GrowingRangeAggregateFunction
				readonly sheetIndex: number
				readonly startRow: number
				readonly startCol: number
				readonly endRow: number
				readonly endCol: number
		  }
		| undefined
	readonly growingRangeAggregate?: GrowingRangeAggregate | undefined
	readonly parseError?: string
}

export interface IndexedFormula {
	readonly key: CellKey
	readonly sheetIndex: number
	readonly sheetName: string
	readonly row: number
	readonly col: number
	readonly formula: string
	readonly ast?: FormulaNode
	readonly refs: readonly FormulaRef[]
	readonly volatile: boolean
	readonly rangeAggregate?:
		| {
				readonly functionName: GrowingRangeAggregateFunction
				readonly sheetIndex: number
				readonly startRow: number
				readonly startCol: number
				readonly endRow: number
				readonly endCol: number
		  }
		| undefined
	readonly growingRangeAggregate?: GrowingRangeAggregate | undefined
	readonly parseError?: string
}

export type GrowingRangeAggregateFunction = 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'

export interface GrowingRangeAggregate {
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

export interface WorkbookFormulaAnalysis {
	readonly formulas: ReadonlyMap<CellKey, IndexedFormula>
	readonly sheetNameIndex: ReadonlyMap<string, number>
}

export interface WorkbookDependencyAnalysis {
	readonly dependencyGraph: DependencyGraph
	readonly resolvedFormulas: ReadonlyMap<CellKey, AnalyzedFormula>
	readonly cycles: readonly (readonly CellKey[])[]
	readonly cycleKeys: ReadonlySet<CellKey>
	readonly sheetNameIndex: ReadonlyMap<string, number>
}

export interface WorkbookAnalysis {
	readonly formulas: ReadonlyMap<CellKey, AnalyzedFormula>
	readonly dependencyGraph: DependencyGraph
	readonly cycles: readonly (readonly CellKey[])[]
	readonly cycleKeys: ReadonlySet<CellKey>
	readonly sheetNameIndex: ReadonlyMap<string, number>
	readonly sharedFormulaGroups: ReadonlyMap<string, CellKey[]>
	readonly growingAggregateAppendIndex: GrowingAggregateAppendIndex
}

export interface GrowingAggregateAppendIndex {
	readonly size: number
	get(key: CellKey): readonly CellKey[] | undefined
}

export interface AnalyzeWorkbookOptions {
	readonly range?: RangeRef
}

interface MutableWorkbookDependencyAnalysis {
	dependencyGraph: DependencyGraph
	resolvedFormulas: Map<CellKey, AnalyzedFormula>
	cycles: readonly (readonly CellKey[])[]
	cycleKeys: Set<CellKey>
}

interface MutableWorkbookAnalysis {
	formulas: Map<CellKey, AnalyzedFormula>
	dependencyGraph: DependencyGraph
	cycles: readonly (readonly CellKey[])[]
	cycleKeys: Set<CellKey>
	sharedFormulaGroups: Map<string, CellKey[]>
	growingAggregateAppendIndex: GrowingAggregateAppendIndex
}

/**
 * Analysis caching: formulas are parsed once per workbook and shared across read-oriented tools
 * (trace, check, lint). WeakMap keyed by Workbook identity; no generation counter needed because
 * the engine invalidates or patches the cache when formulas change via applyOperation.
 *
 * - workbookFormulaAnalysisCache: parsed ASTs, refs, volatile flags (analyzeWorkbookFormulas)
 * - workbookDependencyAnalysisCache: dependency graph, cycles (analyzeWorkbookDependencies)
 * - workbookAnalysisCache: full analysis including shared formula groups (analyzeWorkbook)
 *
 * Invalidation: invalidateWorkbookAnalysis(workbook) clears all caches (used when operations
 * like deleteSheet, addSheet, etc. affect formulas in non-incremental ways).
 * Incremental updates: patchWorkbookAnalysis(workbook, changedCells) and
 * shiftWorkbookAnalysisForAxis(...) update cached data for setFormula, copyRange, insertRows, etc.
 */
interface AnalysisCacheEntry {
	formulas?: WorkbookFormulaAnalysis
	dependencies?: WorkbookDependencyAnalysis
	full?: WorkbookAnalysis
}

const analysisCache = new WeakMap<Workbook, AnalysisCacheEntry>()

function getCacheEntry(workbook: Workbook): AnalysisCacheEntry {
	let entry = analysisCache.get(workbook)
	if (!entry) {
		entry = {}
		analysisCache.set(workbook, entry)
	}
	return entry
}

const workbookFormulaAnalysisCache = {
	get(wb: Workbook) {
		return getCacheEntry(wb).formulas
	},
	set(wb: Workbook, v: WorkbookFormulaAnalysis) {
		getCacheEntry(wb).formulas = v
	},
	delete(wb: Workbook) {
		const e = analysisCache.get(wb)
		if (e) delete e.formulas
	},
}
const workbookDependencyAnalysisCache = {
	get(wb: Workbook) {
		return getCacheEntry(wb).dependencies
	},
	set(wb: Workbook, v: WorkbookDependencyAnalysis) {
		getCacheEntry(wb).dependencies = v
	},
	delete(wb: Workbook) {
		const e = analysisCache.get(wb)
		if (e) delete e.dependencies
	},
}
const workbookAnalysisCache = {
	get(wb: Workbook) {
		return getCacheEntry(wb).full
	},
	set(wb: Workbook, v: WorkbookAnalysis) {
		getCacheEntry(wb).full = v
	},
	delete(wb: Workbook) {
		const e = analysisCache.get(wb)
		if (e) delete e.full
	},
}

export function createSheetNameIndex(workbook: Workbook): Map<string, number> {
	const index = new Map<string, number>()
	for (let i = 0; i < workbook.sheets.length; i++) {
		const sheet = workbook.sheets[i]
		if (!sheet) continue
		index.set(sheet.name.toLowerCase(), i)
	}
	return index
}

export function resolveSheetIndex(
	sheetNameIndex: ReadonlyMap<string, number>,
	sheetName: string | undefined,
	currentSheet: number,
): number {
	return resolveSheetIndexByMap(sheetNameIndex, sheetName, currentSheet)
}

function tryReuseFromCellAbove(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	cell: { formula: string | null; formulaInfo?: { kind?: string } | undefined },
	formulas: ReadonlyMap<CellKey, IndexedFormula>,
): FormulaNode | null {
	if (row === 0) return null
	const prevKey = cellKey(sheetIndex, row - 1, col)
	const prev = formulas.get(prevKey)
	if (!prev?.ast) return null
	const rewritten = rewriteRefs(prev.ast, (ref: FormulaCellRef) => ({
		...ref,
		row: ref.rowAbsolute ? ref.row : ref.row + 1,
		col: ref.colAbsolute ? ref.col : ref.col,
	}))
	const rewrittenText = printFormula(rewritten)
	const currentText = resolveCellFormulaText(workbook, sheetIndex, row, col, cell)
	if (currentText !== rewrittenText) return null
	return rewritten
}

function tryGetSingleRangeAggregate(
	sheetNameIndex: ReadonlyMap<string, number>,
	current: {
		sheetIndex: number
		row: number
		col: number
		ast: FormulaNode
	},
): IndexedFormula['rangeAggregate'] {
	if (current.ast.type !== 'function') return undefined
	const functionName = current.ast.name.toUpperCase()
	if (!isGrowingRangeAggregateFunction(functionName)) return undefined
	if (current.ast.args.length !== 1) return undefined
	const currentArg = current.ast.args[0]
	if (!currentArg || currentArg.type !== 'rangeRef') return undefined
	const currentRangeSheet =
		currentArg.sheet === undefined
			? current.sheetIndex
			: resolveSheetIndex(sheetNameIndex, currentArg.sheet, current.sheetIndex)
	if (currentRangeSheet < 0) return undefined
	return {
		functionName,
		sheetIndex: currentRangeSheet,
		startRow: currentArg.start.row,
		startCol: currentArg.start.col,
		endRow: currentArg.end.row,
		endCol: currentArg.end.col,
	}
}

function isGrowingRangeAggregateFunction(name: string): name is GrowingRangeAggregateFunction {
	return (
		name === 'SUM' || name === 'COUNT' || name === 'AVERAGE' || name === 'MIN' || name === 'MAX'
	)
}

function tryDetectGrowingRangeAggregate(
	currentRangeAggregate: IndexedFormula['rangeAggregate'],
	previousInColumn: IndexedFormula | undefined,
): IndexedFormula['growingRangeAggregate'] {
	if (!currentRangeAggregate) return undefined

	const previous = previousInColumn
	const previousRangeAggregate = previous?.rangeAggregate
	if (!previousRangeAggregate) return undefined
	if (previousRangeAggregate.functionName !== currentRangeAggregate.functionName) return undefined
	if (previousRangeAggregate.sheetIndex !== currentRangeAggregate.sheetIndex) return undefined
	if (
		currentRangeAggregate.startRow !== previousRangeAggregate.startRow ||
		currentRangeAggregate.startCol !== previousRangeAggregate.startCol ||
		currentRangeAggregate.endRow <= previousRangeAggregate.endRow ||
		currentRangeAggregate.endCol !== previousRangeAggregate.endCol ||
		previous.row >= currentRangeAggregate.endRow
	) {
		return undefined
	}

	return {
		functionName: currentRangeAggregate.functionName,
		previousKey: previous.key,
		previousSheetIndex: previous.sheetIndex,
		previousRow: previous.row,
		previousCol: previous.col,
		appendSheetIndex: currentRangeAggregate.sheetIndex,
		appendStartRow: previousRangeAggregate.endRow + 1,
		appendStartCol: previousRangeAggregate.startCol,
		appendEndRow: currentRangeAggregate.endRow,
		appendEndCol: currentRangeAggregate.endCol,
	}
}

export function analyzeWorkbookFormulas(
	workbook: Workbook,
	options: AnalyzeWorkbookOptions = {},
): WorkbookFormulaAnalysis {
	if (!options.range) {
		const cached = workbookFormulaAnalysisCache.get(workbook)
		if (cached) return cached
	}
	const sheetNameIndex = createSheetNameIndex(workbook)
	const formulas = new Map<CellKey, IndexedFormula>()
	const sharedMasterCache = new Map<string, FormulaNode>()
	const nameResolveCache = new Map<string, FormulaRef[]>()

	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		const previousFormulaByCol = new Map<number, IndexedFormula>()
		for (const [row, rowCells] of sheet.cells.iterateRows()) {
			for (const [col, cell] of rowCells) {
				if (!cellHasFormula(cell)) continue
				if (!inRange(sheet.name, row, col, options.range)) continue

				const key = cellKey(sheetIndex, row, col)
				const formulaText = resolveCellFormulaText(workbook, sheetIndex, row, col, cell)
				let ast: FormulaNode
				const reused = tryReuseFromCellAbove(workbook, sheetIndex, row, col, cell, formulas)
				if (reused) {
					ast = reused
				} else {
					const parsed = parseIndexedFormula(
						workbook,
						sheetIndex,
						row,
						col,
						cell,
						sharedMasterCache,
					)
					if (!parsed.ok) {
						formulas.set(key, {
							key,
							sheetIndex,
							sheetName: sheet.name,
							row,
							col,
							formula: formulaText ?? cell.formula ?? '',
							refs: [],
							volatile: false,
							parseError: parsed.error.message,
						})
						continue
					}
					ast = parsed.value
				}
				const refs = extractRefsWithNames(
					ast,
					workbook,
					sheetNameIndex,
					sheetIndex,
					[],
					nameResolveCache,
				)
				const volatile = hasVolatileFunction(ast)
				const rangeAggregate = tryGetSingleRangeAggregate(sheetNameIndex, {
					sheetIndex,
					row,
					col,
					ast,
				})
				const growingRangeAggregate = tryDetectGrowingRangeAggregate(
					rangeAggregate,
					previousFormulaByCol.get(col),
				)
				const indexed = {
					key,
					sheetIndex,
					sheetName: sheet.name,
					row,
					col,
					formula: formulaText ?? printFormula(ast),
					ast,
					refs,
					volatile,
					...(rangeAggregate ? { rangeAggregate } : {}),
					...(growingRangeAggregate ? { growingRangeAggregate } : {}),
				} satisfies IndexedFormula
				formulas.set(key, indexed)
				previousFormulaByCol.set(col, indexed)
			}
		}
	}

	const analysis = { formulas, sheetNameIndex }
	if (!options.range) workbookFormulaAnalysisCache.set(workbook, analysis)
	return analysis
}

export function cellHasFormula(
	cell:
		| {
				formula: string | null
				formulaInfo?:
					| {
							kind?: string
							sharedIndex?: string
							isMaster?: boolean
							masterRef?: string
					  }
					| undefined
		  }
		| null
		| undefined,
): boolean {
	if (!cell) return false
	return cell.formula !== null || cell.formulaInfo?.kind === 'shared'
}

export function resolveCellFormulaText(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	cell: {
		formula: string | null
		formulaInfo?:
			| {
					kind?: string
					sharedIndex?: string
					isMaster?: boolean
					masterRef?: string
			  }
			| undefined
	},
): string | null {
	if (cell.formula) return cell.formula
	const binding = cell.formulaInfo
	if (binding?.kind !== 'shared' || binding.isMaster) return null
	const masterAst = loadSharedMasterAst(workbook, sheetIndex, binding)
	if (!masterAst) return null
	return printFormula(rewriteSharedFormulaAst(masterAst, binding.masterRef, row, col))
}

function parseIndexedFormula(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	cell: {
		formula: string | null
		formulaInfo?:
			| {
					kind?: string
					sharedIndex?: string
					isMaster?: boolean
					masterRef?: string
			  }
			| undefined
	},
	sharedMasterCache: Map<string, FormulaNode>,
): ReturnType<typeof cachedParseFormula> {
	const binding = cell.formulaInfo
	if (binding?.kind === 'shared') {
		const cacheKey = `${sheetIndex}:${binding.sharedIndex ?? ''}`
		if (binding.isMaster && cell.formula) {
			const parsed = cachedParseFormula(cell.formula)
			if (parsed.ok) sharedMasterCache.set(cacheKey, parsed.value)
			return parsed
		}
		const masterAst =
			sharedMasterCache.get(cacheKey) ?? loadSharedMasterAst(workbook, sheetIndex, binding)
		if (masterAst) {
			sharedMasterCache.set(cacheKey, masterAst)
			return {
				ok: true as const,
				value: rewriteSharedFormulaAst(masterAst, binding.masterRef, row, col),
			}
		}
	}
	if (!cell.formula) {
		return {
			ok: false as const,
			error: ascendError('FORMULA_PARSE_ERROR', 'Formula text is missing for this cell.'),
		}
	}
	return cachedParseFormula(cell.formula)
}

function loadSharedMasterAst(
	workbook: Workbook,
	sheetIndex: number,
	binding: { masterRef?: string },
): FormulaNode | null {
	if (!binding.masterRef) return null
	const ref = parseA1Safe(binding.masterRef)
	if (!ref) return null
	const formula = workbook.sheets[sheetIndex]?.cells.readFormula(ref.row, ref.col)
	if (!formula) return null
	const parsed = cachedParseFormula(formula)
	return parsed.ok ? parsed.value : null
}

function rewriteSharedFormulaAst(
	masterAst: FormulaNode,
	masterRef: string | undefined,
	row: number,
	col: number,
): FormulaNode {
	const anchor = parseA1Safe(masterRef)
	if (!anchor) return masterAst
	const rowDelta = row - anchor.row
	const colDelta = col - anchor.col
	return rewriteRefs(masterAst, (ref: FormulaCellRef) => ({
		...ref,
		row: ref.rowAbsolute ? ref.row : ref.row + rowDelta,
		col: ref.colAbsolute ? ref.col : ref.col + colDelta,
	}))
}

export function analyzeWorkbook(
	workbook: Workbook,
	options: AnalyzeWorkbookOptions = {},
): WorkbookAnalysis {
	if (!options.range) {
		const cached = workbookAnalysisCache.get(workbook)
		if (cached) return cached
	}
	const indexed = analyzeWorkbookFormulas(workbook, options)
	const dependency = analyzeWorkbookDependenciesFrom(workbook, indexed)
	const analysis = {
		formulas: dependency.resolvedFormulas,
		dependencyGraph: dependency.dependencyGraph,
		cycles: dependency.cycles,
		cycleKeys: dependency.cycleKeys,
		sheetNameIndex: indexed.sheetNameIndex,
		sharedFormulaGroups: getSharedFormulaGroups(workbook, dependency.resolvedFormulas),
		growingAggregateAppendIndex: getGrowingAggregateAppendIndex(dependency.resolvedFormulas),
	}
	if (!options.range) workbookAnalysisCache.set(workbook, analysis)
	return analysis
}

export function analyzeWorkbookDependencies(
	workbook: Workbook,
	options: AnalyzeWorkbookOptions = {},
): WorkbookDependencyAnalysis {
	if (!options.range) {
		const cached = workbookDependencyAnalysisCache.get(workbook)
		if (cached) return cached
	}
	const indexed = analyzeWorkbookFormulas(workbook, options)
	return analyzeWorkbookDependenciesFrom(workbook, indexed, !options.range)
}

function analyzeWorkbookDependenciesFrom(
	workbook: Workbook,
	indexed: WorkbookFormulaAnalysis,
	cache = true,
): WorkbookDependencyAnalysis {
	const dependencyGraph = new DependencyGraph()
	const resolvedFormulas = new Map<CellKey, AnalyzedFormula>()
	const structuredRefResolver = createStructuredRefResolver(workbook)
	for (const formula of indexed.formulas.values()) {
		const resolved = resolveFormulaDependencies(
			workbook,
			indexed.sheetNameIndex,
			formula,
			structuredRefResolver,
		)
		resolvedFormulas.set(resolved.key, resolved)
		if (resolved.parseError) continue
		dependencyGraph.addFormula(resolved.key, resolved.deps, resolved.volatile, resolved.rangeDeps)
	}
	const cycles = dependencyGraph.detectCycles()
	const cycleKeys = new Set<CellKey>()
	for (const cycle of cycles) {
		for (const key of cycle) cycleKeys.add(key)
	}
	const analysis = {
		dependencyGraph,
		resolvedFormulas,
		cycles,
		cycleKeys,
		sheetNameIndex: indexed.sheetNameIndex,
	}
	if (cache) workbookDependencyAnalysisCache.set(workbook, analysis)
	return analysis
}

export function invalidateWorkbookAnalysis(workbook: Workbook): void {
	workbookFormulaAnalysisCache.delete(workbook)
	workbookDependencyAnalysisCache.delete(workbook)
	workbookAnalysisCache.delete(workbook)
}

export function patchWorkbookAnalysis(workbook: Workbook, changedCells: CellKey[]): void {
	const cachedFormulas = workbookFormulaAnalysisCache.get(workbook)
	if (!cachedFormulas) return

	const cachedDeps = workbookDependencyAnalysisCache.get(workbook)
	const cachedFull = workbookAnalysisCache.get(workbook)
	const sheetNameIndex = cachedFormulas.sheetNameIndex
	const formulas = cachedFormulas.formulas as Map<CellKey, IndexedFormula>
	const nameResolveCache = new Map<string, FormulaRef[]>()

	for (const key of changedCells) {
		const [sheetIndex, row, col] = parseCellKey(key)
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) {
			formulas.delete(key)
			cachedDeps?.dependencyGraph.removeFormula(key)
			cachedFull?.dependencyGraph.removeFormula(key)
			continue
		}
		const hasCell = sheet.cells.has(row, col)
		if (!hasCell) {
			formulas.delete(key)
			cachedDeps?.dependencyGraph.removeFormula(key)
			cachedFull?.dependencyGraph.removeFormula(key)
			if (cachedDeps) {
				;(cachedDeps.resolvedFormulas as Map<CellKey, AnalyzedFormula>).delete(key)
			}
			if (cachedFull) {
				;(cachedFull.formulas as Map<CellKey, AnalyzedFormula>).delete(key)
			}
			continue
		}
		const cell = {
			formula: sheet.cells.readFormula(row, col) ?? null,
			formulaInfo: sheet.cells.readFormulaInfo(row, col),
		}
		if (!cellHasFormula(cell)) {
			formulas.delete(key)
			cachedDeps?.dependencyGraph.removeFormula(key)
			cachedFull?.dependencyGraph.removeFormula(key)
			if (cachedDeps) {
				;(cachedDeps.resolvedFormulas as Map<CellKey, AnalyzedFormula>).delete(key)
			}
			if (cachedFull) {
				;(cachedFull.formulas as Map<CellKey, AnalyzedFormula>).delete(key)
			}
			continue
		}

		const formulaText = resolveCellFormulaText(workbook, sheetIndex, row, col, cell)
		const parsed = cachedParseFormula(formulaText ?? cell.formula ?? '')
		if (!parsed.ok) {
			const entry: IndexedFormula = {
				key,
				sheetIndex,
				sheetName: sheet.name,
				row,
				col,
				formula: formulaText ?? cell.formula ?? '',
				refs: [],
				volatile: false,
				parseError: parsed.error.message,
			}
			formulas.set(key, entry)
			cachedDeps?.dependencyGraph.removeFormula(key)
			cachedFull?.dependencyGraph.removeFormula(key)
			continue
		}

		const ast = parsed.value
		const refs = extractRefsWithNames(
			ast,
			workbook,
			sheetNameIndex,
			sheetIndex,
			[],
			nameResolveCache,
		)
		const volatile = hasVolatileFunction(ast)
		const indexed: IndexedFormula = {
			key,
			sheetIndex,
			sheetName: sheet.name,
			row,
			col,
			formula: formulaText ?? printFormula(ast),
			ast,
			refs,
			volatile,
		}
		formulas.set(key, indexed)

		const resolved = resolveFormulaDependencies(workbook, sheetNameIndex, indexed)
		if (cachedDeps) {
			;(cachedDeps.resolvedFormulas as Map<CellKey, AnalyzedFormula>).set(key, resolved)
			cachedDeps.dependencyGraph.removeFormula(key)
			if (!resolved.parseError) {
				cachedDeps.dependencyGraph.addFormula(
					key,
					resolved.deps,
					resolved.volatile,
					resolved.rangeDeps,
				)
			}
		}
		if (cachedFull) {
			;(cachedFull.formulas as Map<CellKey, AnalyzedFormula>).set(key, resolved)
			cachedFull.dependencyGraph.removeFormula(key)
			if (!resolved.parseError) {
				cachedFull.dependencyGraph.addFormula(
					key,
					resolved.deps,
					resolved.volatile,
					resolved.rangeDeps,
				)
			}
		}
	}
	if (cachedFull) {
		;(cachedFull as unknown as MutableWorkbookAnalysis).growingAggregateAppendIndex =
			getGrowingAggregateAppendIndex(cachedFull.formulas)
	}
}

export function shiftWorkbookAnalysisForAxis(
	workbook: Workbook,
	targetSheetName: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	const cachedFormulas = workbookFormulaAnalysisCache.get(workbook)
	if (!cachedFormulas) return

	const targetSheetIndex = cachedFormulas.sheetNameIndex.get(targetSheetName.toLowerCase())
	if (targetSheetIndex === undefined) {
		invalidateWorkbookAnalysis(workbook)
		return
	}

	const formulas = cachedFormulas.formulas as Map<CellKey, IndexedFormula>
	const nextFormulas = new Map<CellKey, IndexedFormula>()
	const nameResolveCache = new Map<string, FormulaRef[]>()

	for (const formula of formulas.values()) {
		const nextIndex =
			formula.sheetIndex === targetSheetIndex
				? axis === 'row'
					? shiftIndex(formula.row, at, delta)
					: shiftIndex(formula.col, at, delta)
				: axis === 'row'
					? formula.row
					: formula.col
		if (nextIndex === null) continue

		const nextRow =
			formula.sheetIndex === targetSheetIndex && axis === 'row' ? nextIndex : formula.row
		const nextCol =
			formula.sheetIndex === targetSheetIndex && axis === 'col' ? nextIndex : formula.col
		const nextKey = cellKey(formula.sheetIndex, nextRow, nextCol)

		let nextAst = formula.ast
		if (nextAst) {
			const needsRewrite =
				formula.sheetIndex === targetSheetIndex ||
				formulaAstReferencesSheet(nextAst, targetSheetName)
			if (needsRewrite) {
				nextAst = rewriteFormulaAstForShift(
					nextAst,
					targetSheetName,
					formula.sheetName,
					axis,
					at,
					delta,
				)
			}
		}

		let nextFormulaText = formula.formula
		let nextRefs = formula.refs
		let nextParseError = formula.parseError
		let nextVolatile = formula.volatile
		if (nextAst) {
			nextFormulaText = printFormula(nextAst)
			nextRefs = extractRefsWithNames(
				nextAst,
				workbook,
				cachedFormulas.sheetNameIndex,
				formula.sheetIndex,
				[],
				nameResolveCache,
			)
			nextVolatile = hasVolatileFunction(nextAst)
			nextParseError = undefined
		}

		nextFormulas.set(nextKey, {
			key: nextKey,
			sheetIndex: formula.sheetIndex,
			sheetName: formula.sheetName,
			row: nextRow,
			col: nextCol,
			formula: nextFormulaText,
			...(nextAst ? { ast: nextAst } : {}),
			refs: nextRefs,
			volatile: nextVolatile,
			...(nextParseError ? { parseError: nextParseError } : {}),
		})
	}

	formulas.clear()
	for (const [key, formula] of nextFormulas) formulas.set(key, formula)

	const nextDependency = analyzeWorkbookDependenciesFrom(
		workbook,
		{ formulas, sheetNameIndex: cachedFormulas.sheetNameIndex },
		false,
	)

	const cachedDeps = workbookDependencyAnalysisCache.get(workbook)
	if (cachedDeps) {
		const mutableDeps = cachedDeps as unknown as MutableWorkbookDependencyAnalysis
		mutableDeps.dependencyGraph = nextDependency.dependencyGraph
		mutableDeps.resolvedFormulas = nextDependency.resolvedFormulas as Map<CellKey, AnalyzedFormula>
		mutableDeps.cycles = nextDependency.cycles
		mutableDeps.cycleKeys = nextDependency.cycleKeys as Set<CellKey>
	}

	const cachedFull = workbookAnalysisCache.get(workbook)
	if (cachedFull) {
		const mutableFull = cachedFull as unknown as MutableWorkbookAnalysis
		mutableFull.formulas = nextDependency.resolvedFormulas as Map<CellKey, AnalyzedFormula>
		mutableFull.dependencyGraph = nextDependency.dependencyGraph
		mutableFull.cycles = nextDependency.cycles
		mutableFull.cycleKeys = nextDependency.cycleKeys as Set<CellKey>
		mutableFull.sharedFormulaGroups = getSharedFormulaGroups(workbook, mutableFull.formulas)
		mutableFull.growingAggregateAppendIndex = getGrowingAggregateAppendIndex(mutableFull.formulas)
	}
}

export function getGrowingAggregateAppendIndex(
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
): GrowingAggregateAppendIndex {
	const appendKeys: CellKey[] = []
	const formulaKeys: CellKey[] = []
	let sorted = true
	let previousAppendKey = Number.NEGATIVE_INFINITY
	for (const formula of formulas.values()) {
		const aggregate = formula.growingRangeAggregate
		if (!aggregate) continue
		if (
			aggregate.functionName !== 'SUM' &&
			aggregate.functionName !== 'COUNT' &&
			aggregate.functionName !== 'AVERAGE' &&
			aggregate.functionName !== 'MIN' &&
			aggregate.functionName !== 'MAX'
		) {
			continue
		}
		if (
			aggregate.appendStartRow !== aggregate.appendEndRow ||
			aggregate.appendStartCol !== aggregate.appendEndCol
		) {
			continue
		}
		const key = cellKey(
			aggregate.appendSheetIndex,
			aggregate.appendStartRow,
			aggregate.appendStartCol,
		)
		if (key < previousAppendKey) sorted = false
		previousAppendKey = key
		appendKeys.push(key)
		formulaKeys.push(formula.key)
	}
	if (!sorted) {
		const pairs = appendKeys.map((key, index) => [key, formulaKeys[index] as CellKey] as const)
		pairs.sort((a, b) => a[0] - b[0])
		for (let i = 0; i < pairs.length; i++) {
			const pair = pairs[i] as readonly [CellKey, CellKey]
			appendKeys[i] = pair[0]
			formulaKeys[i] = pair[1]
		}
	}
	return new SortedGrowingAggregateAppendIndex(appendKeys, formulaKeys)
}

class SortedGrowingAggregateAppendIndex implements GrowingAggregateAppendIndex {
	private readonly single: CellKey[] = []
	private readonly multi: CellKey[] = []

	constructor(
		private readonly appendKeys: readonly CellKey[],
		private readonly formulaKeys: readonly CellKey[],
	) {}

	get size(): number {
		return this.appendKeys.length
	}

	get(key: CellKey): readonly CellKey[] | undefined {
		let lo = 0
		let hi = this.appendKeys.length - 1
		while (lo <= hi) {
			const mid = (lo + hi) >>> 1
			const midKey = this.appendKeys[mid] as CellKey
			if (midKey < key) lo = mid + 1
			else if (midKey > key) hi = mid - 1
			else {
				let start = mid
				while (start > 0 && this.appendKeys[start - 1] === key) start--
				let end = mid + 1
				while (end < this.appendKeys.length && this.appendKeys[end] === key) end++
				if (end - start === 1) {
					this.single[0] = this.formulaKeys[start] as CellKey
					this.single.length = 1
					return this.single
				}
				this.multi.length = 0
				for (let i = start; i < end; i++) this.multi.push(this.formulaKeys[i] as CellKey)
				return this.multi
			}
		}
		return undefined
	}
}

export function getSharedFormulaGroups(
	workbook: Workbook,
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
): Map<string, CellKey[]> {
	const groups = new Map<string, CellKey[]>()
	for (const formula of formulas.values()) {
		const sheet = workbook.sheets[formula.sheetIndex]
		if (!sheet) continue
		const binding = sheet.cells.readFormulaInfo(formula.row, formula.col) as
			| { kind?: string; sharedIndex?: string }
			| undefined
		if (!binding) continue
		if (binding.kind !== 'shared' || binding.sharedIndex === undefined) continue
		const groupKey = `${formula.sheetIndex}:${binding.sharedIndex}`
		let group = groups.get(groupKey)
		if (!group) {
			group = []
			groups.set(groupKey, group)
		}
		group.push(formula.key)
	}
	return groups
}

export function resolveFormulaDependencies(
	workbook: Workbook,
	sheetNameIndex: ReadonlyMap<string, number>,
	formula: IndexedFormula,
	structuredRefResolver = createStructuredRefResolver(workbook),
): AnalyzedFormula {
	if (!formula.ast) {
		return {
			...formula,
			deps: [],
			rangeDeps: [],
		}
	}
	const deps: CellKey[] = []
	const rangeDeps: RangeDependency[] = []
	for (const ref of formula.refs) {
		if (ref.kind === 'sheetSpan') {
			const span = resolveSheetSpan(sheetNameIndex, ref.startSheet, ref.endSheet)
			if (!span) continue
			for (const refSheetIndex of span) {
				const target = ref.target
				if (target.kind === 'sheetSpan') continue
				if (target.kind === 'cell') {
					deps.push(cellKey(refSheetIndex, target.ref.row, target.ref.col))
				} else {
					const rangeDep = formulaRefToRangeDependency(workbook, refSheetIndex, target)
					if (rangeDep) rangeDeps.push(rangeDep)
				}
			}
			continue
		}
		const refSheetIndex = resolveSheetIndex(sheetNameIndex, ref.sheet, formula.sheetIndex)
		if (refSheetIndex < 0) continue
		if (ref.kind === 'cell') {
			deps.push(cellKey(refSheetIndex, ref.ref.row, ref.ref.col))
		} else {
			const rangeDep = formulaRefToRangeDependency(workbook, refSheetIndex, ref)
			if (rangeDep) rangeDeps.push(rangeDep)
		}
	}
	rangeDeps.push(
		...collectStructuredRefDependenciesWithNames(
			formula.ast,
			workbook,
			structuredRefResolver,
			formula.sheetIndex,
			formula.row,
			formula.col,
		),
	)
	if (
		!deps.includes(formula.key) &&
		formulaHasValueDependentSelfRange(
			{
				workbook,
				sheetNameIndex,
				structuredRefResolver,
				formulaSheetIndex: formula.sheetIndex,
				formulaRow: formula.row,
				formulaCol: formula.col,
				seenNames: [],
			},
			formula.ast,
			formula.sheetIndex,
		)
	) {
		deps.push(formula.key)
	}
	return {
		...formula,
		deps,
		rangeDeps,
		...(formula.growingRangeAggregate
			? { growingRangeAggregate: formula.growingRangeAggregate }
			: {}),
	}
}

const REFERENCE_SHAPE_FUNCTIONS = new Set(['AREAS', 'COLUMNS', 'FORMULATEXT', 'ISREF', 'ROWS'])

interface SelfCycleContext {
	readonly workbook: Workbook
	readonly sheetNameIndex: ReadonlyMap<string, number>
	readonly structuredRefResolver: StructuredRefResolver
	readonly formulaSheetIndex: number
	readonly formulaRow: number
	readonly formulaCol: number
	readonly seenNames: readonly string[]
}

interface SelfCycleRange {
	readonly sheetIndex: number
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

function formulaHasValueDependentSelfRange(
	ctx: SelfCycleContext,
	node: FormulaNode,
	currentSheetIndex: number,
	inValueContext = true,
): boolean {
	switch (node.type) {
		case 'rangeRef': {
			if (!inValueContext) return false
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			return rangeContainsFormulaCell(ctx, {
				sheetIndex,
				startRow: Math.min(node.start.row, node.end.row),
				startCol: Math.min(node.start.col, node.end.col),
				endRow: Math.max(node.start.row, node.end.row),
				endCol: Math.max(node.start.col, node.end.col),
			})
		}
		case 'wholeRowRange': {
			if (!inValueContext) return false
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			const sheet = ctx.workbook.sheets[sheetIndex]
			const usedRange = sheet?.cells.usedRange()
			if (!usedRange) return false
			return rangeContainsFormulaCell(ctx, {
				sheetIndex,
				startRow: Math.min(node.startRow, node.endRow),
				startCol: usedRange.start.col,
				endRow: Math.max(node.startRow, node.endRow),
				endCol: usedRange.end.col,
			})
		}
		case 'wholeColumnRange': {
			if (!inValueContext) return false
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			const sheet = ctx.workbook.sheets[sheetIndex]
			const usedRange = sheet?.cells.usedRange()
			if (!usedRange) return false
			return rangeContainsFormulaCell(ctx, {
				sheetIndex,
				startRow: usedRange.start.row,
				startCol: Math.min(node.startCol, node.endCol),
				endRow: usedRange.end.row,
				endCol: Math.max(node.startCol, node.endCol),
			})
		}
		case 'structuredRef': {
			if (!inValueContext) return false
			const range = resolveSelfCycleRange(ctx, node, currentSheetIndex)
			return range ? rangeContainsFormulaCell(ctx, range) : false
		}
		case 'name':
			if (!inValueContext) return false
			return definedNameHasValueDependentSelfRange(ctx, node.name, node.sheet, currentSheetIndex)
		case 'function': {
			if (node.name.toUpperCase() === 'INDEX') {
				const narrowed = indexFunctionHasValueDependentSelfRange(ctx, node, currentSheetIndex)
				if (narrowed !== null) return narrowed
			}
			const valueContext = !REFERENCE_SHAPE_FUNCTIONS.has(node.name.toUpperCase())
			return node.args.some((arg) =>
				formulaHasValueDependentSelfRange(ctx, arg, currentSheetIndex, valueContext),
			)
		}
		case 'binary':
			if (node.op === ' ') {
				if (!inValueContext) return false
				const left = resolveSelfCycleRange(ctx, node.left, currentSheetIndex)
				const right = resolveSelfCycleRange(ctx, node.right, currentSheetIndex)
				const intersection = left && right ? intersectSelfCycleRanges(left, right) : null
				return intersection ? rangeContainsFormulaCell(ctx, intersection) : false
			}
			return (
				formulaHasValueDependentSelfRange(ctx, node.left, currentSheetIndex, inValueContext) ||
				formulaHasValueDependentSelfRange(ctx, node.right, currentSheetIndex, inValueContext)
			)
		case 'dynamicRangeRef':
			return (
				formulaHasValueDependentSelfRange(ctx, node.start, currentSheetIndex, inValueContext) ||
				formulaHasValueDependentSelfRange(ctx, node.end, currentSheetIndex, inValueContext)
			)
		case 'unary':
			return formulaHasValueDependentSelfRange(ctx, node.operand, currentSheetIndex, inValueContext)
		case 'array':
			return node.rows.some((row) =>
				row.some((cell) =>
					formulaHasValueDependentSelfRange(ctx, cell, currentSheetIndex, inValueContext),
				),
			)
		case 'spillRef':
			return formulaHasValueDependentSelfRange(ctx, node.target, currentSheetIndex, inValueContext)
		case 'sheetSpanRef':
			return formulaHasValueDependentSelfRange(ctx, node.target, currentSheetIndex, inValueContext)
		default:
			return false
	}
}

function definedNameHasValueDependentSelfRange(
	ctx: SelfCycleContext,
	name: string,
	sheet: string | undefined,
	currentSheetIndex: number,
): boolean {
	const parsed = resolveDefinedNameAstForSelfCycle(ctx, name, sheet, currentSheetIndex)
	if (!parsed) return false
	return formulaHasValueDependentSelfRange(
		withSeenName(ctx, parsed.entryKey),
		parsed.ast,
		parsed.sheetIndex,
		true,
	)
}

function indexFunctionHasValueDependentSelfRange(
	ctx: SelfCycleContext,
	node: Extract<FormulaNode, { type: 'function' }>,
	currentSheetIndex: number,
): boolean | null {
	const range = narrowedIndexStructuredRefRange(
		node,
		ctx.structuredRefResolver,
		ctx.formulaSheetIndex,
		ctx.formulaRow,
		ctx.formulaCol,
	)
	if (!range) return null
	if (rangeContainsFormulaCell(ctx, range)) return true
	for (let i = 1; i < node.args.length; i++) {
		const arg = node.args[i]
		if (arg && formulaHasValueDependentSelfRange(ctx, arg, currentSheetIndex, true)) return true
	}
	return false
}

function resolveSelfCycleRange(
	ctx: SelfCycleContext,
	node: FormulaNode,
	currentSheetIndex: number,
): SelfCycleRange | null {
	switch (node.type) {
		case 'cellRef': {
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			return {
				sheetIndex,
				startRow: node.ref.row,
				startCol: node.ref.col,
				endRow: node.ref.row,
				endCol: node.ref.col,
			}
		}
		case 'rangeRef': {
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			return {
				sheetIndex,
				startRow: Math.min(node.start.row, node.end.row),
				startCol: Math.min(node.start.col, node.end.col),
				endRow: Math.max(node.start.row, node.end.row),
				endCol: Math.max(node.start.col, node.end.col),
			}
		}
		case 'wholeRowRange': {
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			const sheet = ctx.workbook.sheets[sheetIndex]
			const usedRange = sheet?.cells.usedRange()
			if (!usedRange) return null
			return {
				sheetIndex,
				startRow: Math.min(node.startRow, node.endRow),
				startCol: usedRange.start.col,
				endRow: Math.max(node.startRow, node.endRow),
				endCol: usedRange.end.col,
			}
		}
		case 'wholeColumnRange': {
			const sheetIndex = resolveSheetIndex(ctx.sheetNameIndex, node.sheet, currentSheetIndex)
			const sheet = ctx.workbook.sheets[sheetIndex]
			const usedRange = sheet?.cells.usedRange()
			if (!usedRange) return null
			return {
				sheetIndex,
				startRow: usedRange.start.row,
				startCol: Math.min(node.startCol, node.endCol),
				endRow: usedRange.end.row,
				endCol: Math.max(node.startCol, node.endCol),
			}
		}
		case 'structuredRef': {
			const resolved = ctx.structuredRefResolver.resolve(
				node,
				currentSheetIndex,
				ctx.formulaRow,
				ctx.formulaCol,
			)
			return resolved
				? {
						sheetIndex: resolved.sheetIndex,
						startRow: resolved.startRow,
						startCol: resolved.startCol,
						endRow: resolved.endRow,
						endCol: resolved.endCol,
					}
				: null
		}
		case 'name': {
			const parsed = resolveDefinedNameAstForSelfCycle(
				ctx,
				node.name,
				node.sheet,
				currentSheetIndex,
			)
			return parsed
				? resolveSelfCycleRange(withSeenName(ctx, parsed.entryKey), parsed.ast, parsed.sheetIndex)
				: null
		}
		case 'unary':
			return node.op === '@' ? resolveSelfCycleRange(ctx, node.operand, currentSheetIndex) : null
		case 'sheetSpanRef':
			return resolveSelfCycleRange(ctx, node.target, currentSheetIndex)
		default:
			return null
	}
}

function resolveDefinedNameAstForSelfCycle(
	ctx: SelfCycleContext,
	name: string,
	sheet: string | undefined,
	currentSheetIndex: number,
): { readonly ast: FormulaNode; readonly sheetIndex: number; readonly entryKey: string } | null {
	const currentSheet = ctx.workbook.sheets[currentSheetIndex]
	const explicitSheet = sheet ? ctx.workbook.getSheet(sheet) : undefined
	const entry = ctx.workbook.definedNames.resolve(name, currentSheet?.id, explicitSheet?.id)
	if (!entry) return null

	const entryKey =
		entry.scope.kind === 'workbook'
			? `workbook:${entry.name.toLowerCase()}`
			: `sheet:${entry.scope.sheetId}:${entry.name.toLowerCase()}`
	if (ctx.seenNames.includes(entryKey)) return null

	const parsed = cachedParseFormula(entry.formula)
	if (!parsed.ok) return null

	let sheetIndex = currentSheetIndex
	if (entry.scope.kind === 'sheet') {
		const scope = entry.scope
		const localSheetIndex = ctx.workbook.sheets.findIndex(
			(workbookSheet) => workbookSheet.id === scope.sheetId,
		)
		if (localSheetIndex >= 0) sheetIndex = localSheetIndex
	}

	return { ast: parsed.value, sheetIndex, entryKey }
}

function withSeenName(ctx: SelfCycleContext, entryKey: string): SelfCycleContext {
	return { ...ctx, seenNames: [...ctx.seenNames, entryKey] }
}

function intersectSelfCycleRanges(
	left: SelfCycleRange,
	right: SelfCycleRange,
): SelfCycleRange | null {
	if (left.sheetIndex !== right.sheetIndex) return null
	const startRow = Math.max(left.startRow, right.startRow)
	const startCol = Math.max(left.startCol, right.startCol)
	const endRow = Math.min(left.endRow, right.endRow)
	const endCol = Math.min(left.endCol, right.endCol)
	if (startRow > endRow || startCol > endCol) return null
	return {
		sheetIndex: left.sheetIndex,
		startRow,
		startCol,
		endRow,
		endCol,
	}
}

function rangeContainsFormulaCell(ctx: SelfCycleContext, range: SelfCycleRange): boolean {
	return (
		range.sheetIndex === ctx.formulaSheetIndex &&
		ctx.formulaRow >= range.startRow &&
		ctx.formulaRow <= range.endRow &&
		ctx.formulaCol >= range.startCol &&
		ctx.formulaCol <= range.endCol
	)
}

function narrowedIndexStructuredRefRange(
	node: Extract<FormulaNode, { type: 'function' }>,
	structuredRefResolver: StructuredRefResolver,
	sheetIndex: number,
	row: number,
	col: number,
): RangeDependency | null {
	if (node.name.toUpperCase() !== 'INDEX') return null
	const source = node.args[0]
	if (!source || source.type !== 'structuredRef') return null
	const resolved = structuredRefResolver.resolve(source, sheetIndex, row, col)
	if (!resolved) return null

	const height = resolved.endRow - resolved.startRow + 1
	const width = resolved.endCol - resolved.startCol + 1
	let startRow = resolved.startRow
	let endRow = resolved.endRow
	let startCol = resolved.startCol
	let endCol = resolved.endCol

	const rowNum = constantIndexNumber(node.args[1])
	const colNum = constantIndexNumber(node.args[2])
	if (colNum !== null) {
		if (colNum < 1 || colNum > width) return null
		startCol = resolved.startCol + colNum - 1
		endCol = startCol
	}
	if (rowNum !== null) {
		if (rowNum < 1 || rowNum > height) return null
		startRow = resolved.startRow + rowNum - 1
		endRow = startRow
	}
	if (rowNum === null && colNum === null) return null
	return { sheetIndex: resolved.sheetIndex, startRow, startCol, endRow, endCol }
}

function constantIndexNumber(node: FormulaNode | undefined): number | null {
	if (!node || node.type === 'missing') return null
	if (node.type === 'number') {
		const value = Math.trunc(node.value)
		return value === 0 ? null : value
	}
	if (
		node.type === 'unary' &&
		(node.op === '+' || node.op === '-') &&
		node.operand.type === 'number'
	) {
		const value = Math.trunc(node.operand.value)
		if (value === 0) return null
		return node.op === '-' ? -value : value
	}
	return null
}

function resolveSheetSpan(
	sheetNameIndex: ReadonlyMap<string, number>,
	startSheet: string,
	endSheet: string,
): number[] | null {
	const start = sheetNameIndex.get(startSheet.toLowerCase())
	const end = sheetNameIndex.get(endSheet.toLowerCase())
	if (start === undefined || end === undefined || start > end) return null
	return Array.from({ length: end - start + 1 }, (_, offset) => start + offset)
}

function inRange(
	sheetName: string,
	row: number,
	col: number,
	range: RangeRef | undefined,
): boolean {
	if (!range) return true
	if (range.sheet !== undefined && range.sheet.toLowerCase() !== sheetName.toLowerCase())
		return false
	return (
		row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col
	)
}

const volatileFunctionNames: ReadonlySet<string> = (() => {
	const names = new Set<string>()
	for (const [name, def] of functionRegistry.entries()) {
		if (def.volatile) names.add(name)
	}
	return names
})()

function hasVolatileFunction(node: FormulaNode): boolean {
	switch (node.type) {
		case 'function':
			if (volatileFunctionNames.has(node.name.toUpperCase())) return true
			return node.args.some((arg) => hasVolatileFunction(arg))
		case 'binary':
			return hasVolatileFunction(node.left) || hasVolatileFunction(node.right)
		case 'dynamicRangeRef':
			return hasVolatileFunction(node.start) || hasVolatileFunction(node.end)
		case 'unary':
			return hasVolatileFunction(node.operand)
		case 'array':
			return node.rows.some((row) => row.some((cell) => hasVolatileFunction(cell)))
		case 'spillRef':
			return hasVolatileFunction(node.target)
		case 'sheetSpanRef':
			return hasVolatileFunction(node.target)
		case 'wholeRowRange':
		case 'wholeColumnRange':
			return false
		default:
			return false
	}
}

function formulaRefToRangeDependency(
	workbook: Workbook,
	sheetIndex: number,
	ref: Exclude<FormulaRef, { kind: 'cell' } | { kind: 'sheetSpan' }>,
): RangeDependency | null {
	if (ref.kind === 'range') {
		return {
			sheetIndex,
			startRow: ref.start.row,
			startCol: ref.start.col,
			endRow: ref.end.row,
			endCol: ref.end.col,
		}
	}
	const targetSheet = workbook.sheets[sheetIndex]
	const usedRange = targetSheet?.cells.usedRange()
	if (!targetSheet || !usedRange) return null
	if (ref.kind === 'wholeRowRange') {
		return {
			sheetIndex,
			startRow: ref.startRow,
			startCol: usedRange.start.col,
			endRow: ref.endRow,
			endCol: usedRange.end.col,
		}
	}
	return {
		sheetIndex,
		startRow: usedRange.start.row,
		startCol: ref.startCol,
		endRow: usedRange.end.row,
		endCol: ref.endCol,
	}
}

function extractRefsWithNames(
	node: FormulaNode,
	workbook: Workbook,
	sheetNameIndex: ReadonlyMap<string, number>,
	sheetIndex: number,
	seenNames: readonly string[],
	nameResolveCache: Map<string, FormulaRef[]>,
	qualifyImplicitRefsToSheet?: string,
): FormulaRef[] {
	const refs = qualifyImplicitRefsToSheet
		? extractRefs(node).map((ref) => qualifyImplicitFormulaRef(ref, qualifyImplicitRefsToSheet))
		: extractRefs(node)
	const nameRefs = collectNameRefs(node)
	for (const nameRef of nameRefs) {
		const currentSheet = workbook.sheets[sheetIndex]
		const explicitSheet = nameRef.sheet ? workbook.getSheet(nameRef.sheet) : undefined
		const entry = workbook.definedNames.resolve(nameRef.name, currentSheet?.id, explicitSheet?.id)
		if (!entry) continue

		const entryKey =
			entry.scope.kind === 'workbook'
				? `workbook:${entry.name.toLowerCase()}`
				: `sheet:${entry.scope.sheetId}:${entry.name.toLowerCase()}`
		if (seenNames.includes(entryKey)) continue

		let formulaSheetIndex = sheetIndex
		if (entry.scope.kind === 'sheet') {
			const scope = entry.scope
			const localSheetIndex = workbook.sheets.findIndex(
				(workbookSheet) => workbookSheet.id === scope.sheetId,
			)
			if (localSheetIndex >= 0) formulaSheetIndex = localSheetIndex
		}

		const cacheKey = `${entryKey}:${formulaSheetIndex}`
		let resolved = nameResolveCache.get(cacheKey)
		if (resolved === undefined) {
			const parsed = cachedParseFormula(entry.formula)
			if (!parsed.ok) continue
			resolved = extractRefsWithNames(
				parsed.value,
				workbook,
				sheetNameIndex,
				formulaSheetIndex,
				[...seenNames, entryKey],
				nameResolveCache,
				workbook.sheets[formulaSheetIndex]?.name,
			)
			nameResolveCache.set(cacheKey, resolved)
		}
		refs.push(...resolved)
	}
	return refs
}

function qualifyImplicitFormulaRef(ref: FormulaRef, sheet: string): FormulaRef {
	if (ref.kind === 'sheetSpan') return ref
	if (ref.sheet !== undefined) return ref
	return { ...ref, sheet }
}

function collectNameRefs(node: FormulaNode): Array<{ name: string; sheet?: string }> {
	const result: Array<{ name: string; sheet?: string }> = []
	walkNameRefs(node, result, new Set<string>())
	return result
}

function collectStructuredRefDependencies(
	node: FormulaNode,
	structuredRefResolver: StructuredRefResolver,
	sheetIndex: number,
	row: number,
	col: number,
): RangeDependency[] {
	const result: RangeDependency[] = []
	walkStructuredRefDependencies(node, structuredRefResolver, sheetIndex, row, col, result)
	return result
}

function collectStructuredRefDependenciesWithNames(
	node: FormulaNode,
	workbook: Workbook,
	structuredRefResolver: StructuredRefResolver,
	sheetIndex: number,
	row: number,
	col: number,
	seenNames: readonly string[] = [],
	nameDependencyCache = new Map<string, RangeDependency[]>(),
): RangeDependency[] {
	const result = collectStructuredRefDependencies(node, structuredRefResolver, sheetIndex, row, col)
	const nameRefs = collectNameRefs(node)
	for (const nameRef of nameRefs) {
		const currentSheet = workbook.sheets[sheetIndex]
		const explicitSheet = nameRef.sheet ? workbook.getSheet(nameRef.sheet) : undefined
		const entry = workbook.definedNames.resolve(nameRef.name, currentSheet?.id, explicitSheet?.id)
		if (!entry) continue

		const entryKey =
			entry.scope.kind === 'workbook'
				? `workbook:${entry.name.toLowerCase()}`
				: `sheet:${entry.scope.sheetId}:${entry.name.toLowerCase()}`
		if (seenNames.includes(entryKey)) continue

		let formulaSheetIndex = sheetIndex
		if (entry.scope.kind === 'sheet') {
			const scope = entry.scope
			const localSheetIndex = workbook.sheets.findIndex(
				(workbookSheet) => workbookSheet.id === scope.sheetId,
			)
			if (localSheetIndex >= 0) formulaSheetIndex = localSheetIndex
		}

		const cacheKey = `${entryKey}:${formulaSheetIndex}:${row}:${col}`
		let resolved = nameDependencyCache.get(cacheKey)
		if (resolved === undefined) {
			const parsed = cachedParseFormula(entry.formula)
			if (!parsed.ok) continue
			resolved = collectStructuredRefDependenciesWithNames(
				parsed.value,
				workbook,
				structuredRefResolver,
				formulaSheetIndex,
				row,
				col,
				[...seenNames, entryKey],
				nameDependencyCache,
			)
			nameDependencyCache.set(cacheKey, resolved)
		}
		result.push(...resolved)
	}
	return result
}

function walkStructuredRefDependencies(
	node: FormulaNode,
	structuredRefResolver: StructuredRefResolver,
	sheetIndex: number,
	row: number,
	col: number,
	result: RangeDependency[],
): void {
	switch (node.type) {
		case 'structuredRef': {
			const resolved = structuredRefResolver.resolve(node, sheetIndex, row, col)
			if (resolved) result.push(resolved)
			break
		}
		case 'binary':
			walkStructuredRefDependencies(node.left, structuredRefResolver, sheetIndex, row, col, result)
			walkStructuredRefDependencies(node.right, structuredRefResolver, sheetIndex, row, col, result)
			break
		case 'dynamicRangeRef':
			walkStructuredRefDependencies(node.start, structuredRefResolver, sheetIndex, row, col, result)
			walkStructuredRefDependencies(node.end, structuredRefResolver, sheetIndex, row, col, result)
			break
		case 'unary':
			walkStructuredRefDependencies(
				node.operand,
				structuredRefResolver,
				sheetIndex,
				row,
				col,
				result,
			)
			break
		case 'function': {
			const narrowed = narrowedIndexStructuredRefRange(
				node,
				structuredRefResolver,
				sheetIndex,
				row,
				col,
			)
			if (narrowed) {
				result.push(narrowed)
				for (let i = 1; i < node.args.length; i++) {
					const arg = node.args[i]
					if (arg) {
						walkStructuredRefDependencies(arg, structuredRefResolver, sheetIndex, row, col, result)
					}
				}
				break
			}
			for (const arg of node.args) {
				walkStructuredRefDependencies(arg, structuredRefResolver, sheetIndex, row, col, result)
			}
			break
		}
		case 'array':
			for (const sourceRow of node.rows) {
				for (const cell of sourceRow) {
					walkStructuredRefDependencies(cell, structuredRefResolver, sheetIndex, row, col, result)
				}
			}
			break
		case 'spillRef':
			walkStructuredRefDependencies(
				node.target,
				structuredRefResolver,
				sheetIndex,
				row,
				col,
				result,
			)
			break
		case 'sheetSpanRef':
			walkStructuredRefDependencies(
				node.target,
				structuredRefResolver,
				sheetIndex,
				row,
				col,
				result,
			)
			break
		case 'wholeRowRange':
		case 'wholeColumnRange':
			break
		default:
			break
	}
}

function walkNameRefs(
	node: FormulaNode,
	result: Array<{ name: string; sheet?: string }>,
	shadowed: ReadonlySet<string>,
): void {
	switch (node.type) {
		case 'name':
			if (!node.sheet && shadowed.has(node.name.toLowerCase())) break
			result.push(
				node.sheet !== undefined ? { name: node.name, sheet: node.sheet } : { name: node.name },
			)
			break
		case 'binary':
			walkNameRefs(node.left, result, shadowed)
			walkNameRefs(node.right, result, shadowed)
			break
		case 'dynamicRangeRef':
			walkNameRefs(node.start, result, shadowed)
			walkNameRefs(node.end, result, shadowed)
			break
		case 'unary':
			walkNameRefs(node.operand, result, shadowed)
			break
		case 'function':
			if (node.name.toUpperCase() === 'LET') {
				const nextShadowed = new Set(shadowed)
				for (let i = 0; i < node.args.length - 1; i += 2) {
					const binder = node.args[i]
					if (binder?.type !== 'name' || binder.sheet) break
					const valueNode = node.args[i + 1]
					if (valueNode) walkNameRefs(valueNode, result, nextShadowed)
					nextShadowed.add(binder.name.toLowerCase())
				}
				const body = node.args[node.args.length - 1]
				if (body) walkNameRefs(body, result, nextShadowed)
				break
			}
			if (node.name.toUpperCase() === 'LAMBDA') {
				const nextShadowed = new Set(shadowed)
				for (let i = 0; i < node.args.length - 1; i++) {
					const binder = node.args[i]
					if (binder?.type !== 'name' || binder.sheet) break
					nextShadowed.add(binder.name.toLowerCase())
				}
				const body = node.args[node.args.length - 1]
				if (body) walkNameRefs(body, result, nextShadowed)
				break
			}
			for (const arg of node.args) walkNameRefs(arg, result, shadowed)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walkNameRefs(cell, result, shadowed)
			}
			break
		case 'spillRef':
			walkNameRefs(node.target, result, shadowed)
			break
		case 'sheetSpanRef':
			walkNameRefs(node.target, result, shadowed)
			break
		case 'wholeRowRange':
		case 'wholeColumnRange':
			break
		default:
			break
	}
}
