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
import { resolveStructuredRefRange } from './structured-refs.ts'

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
	readonly growingRangeAggregate?:
		| {
				readonly functionName: 'SUM'
				readonly previousKey: CellKey
				readonly appendSheetIndex: number
				readonly appendStartRow: number
				readonly appendStartCol: number
				readonly appendEndRow: number
				readonly appendEndCol: number
		  }
		| undefined
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
	readonly growingRangeAggregate?:
		| {
				readonly functionName: 'SUM'
				readonly previousKey: CellKey
				readonly appendSheetIndex: number
				readonly appendStartRow: number
				readonly appendStartCol: number
				readonly appendEndRow: number
				readonly appendEndCol: number
		  }
		| undefined
	readonly parseError?: string
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
}

const workbookFormulaAnalysisCache = new WeakMap<Workbook, WorkbookFormulaAnalysis>()
const workbookDependencyAnalysisCache = new WeakMap<Workbook, WorkbookDependencyAnalysis>()
const workbookAnalysisCache = new WeakMap<Workbook, WorkbookAnalysis>()
// Uses shared cachedParseFormula from @ascend/formulas

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

function tryDetectGrowingRangeAggregate(
	sheetNameIndex: ReadonlyMap<string, number>,
	current: {
		sheetIndex: number
		row: number
		col: number
		ast: FormulaNode
	},
	previousInColumn: IndexedFormula | undefined,
): IndexedFormula['growingRangeAggregate'] {
	if (current.ast.type !== 'function') return undefined
	if (current.ast.name.toUpperCase() !== 'SUM') return undefined
	if (current.ast.args.length !== 1) return undefined
	const currentArg = current.ast.args[0]
	if (!currentArg || currentArg.type !== 'rangeRef') return undefined
	const currentRangeSheet =
		currentArg.sheet === undefined
			? current.sheetIndex
			: resolveSheetIndex(sheetNameIndex, currentArg.sheet, current.sheetIndex)
	if (currentRangeSheet < 0) return undefined

	const previous = previousInColumn
	if (!previous?.ast || previous.ast.type !== 'function') return undefined
	if (previous.ast.name.toUpperCase() !== 'SUM' || previous.ast.args.length !== 1) return undefined
	const previousArg = previous.ast.args[0]
	if (!previousArg || previousArg.type !== 'rangeRef') return undefined
	const previousRangeSheet =
		previousArg.sheet === undefined
			? previous.sheetIndex
			: resolveSheetIndex(sheetNameIndex, previousArg.sheet, previous.sheetIndex)
	if (previousRangeSheet !== currentRangeSheet) return undefined
	if (
		currentArg.start.row !== previousArg.start.row ||
		currentArg.start.col !== previousArg.start.col ||
		currentArg.end.row <= previousArg.end.row ||
		currentArg.end.col !== previousArg.end.col ||
		current.row <= previous.row
	) {
		return undefined
	}

	return {
		functionName: 'SUM',
		previousKey: previous.key,
		appendSheetIndex: currentRangeSheet,
		appendStartRow: previousArg.end.row + 1,
		appendStartCol: previousArg.start.col,
		appendEndRow: currentArg.end.row,
		appendEndCol: currentArg.end.col,
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
				const growingRangeAggregate = tryDetectGrowingRangeAggregate(
					sheetNameIndex,
					{ sheetIndex, row, col, ast },
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
	for (const formula of indexed.formulas.values()) {
		const resolved = resolveFormulaDependencies(workbook, indexed.sheetNameIndex, formula)
		resolvedFormulas.set(resolved.key, resolved)
		if (resolved.parseError) continue
		dependencyGraph.addFormula(
			resolved.key,
			[...resolved.deps],
			resolved.volatile,
			resolved.rangeDeps,
		)
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
					[...resolved.deps],
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
					[...resolved.deps],
					resolved.volatile,
					resolved.rangeDeps,
				)
			}
		}
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
	for (const structuredRef of collectStructuredRefs(formula.ast)) {
		const resolved = resolveStructuredRefRange(
			workbook,
			structuredRef,
			formula.sheetIndex,
			formula.row,
			formula.col,
		)
		if (!resolved) continue
		rangeDeps.push({
			sheetIndex: resolved.sheetIndex,
			startRow: resolved.startRow,
			startCol: resolved.startCol,
			endRow: resolved.endRow,
			endCol: resolved.endCol,
		})
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
): FormulaRef[] {
	const refs = extractRefs(node)
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
			)
			nameResolveCache.set(cacheKey, resolved)
		}
		refs.push(...resolved)
	}
	return refs
}

function collectNameRefs(node: FormulaNode): Array<{ name: string; sheet?: string }> {
	const result: Array<{ name: string; sheet?: string }> = []
	walkNameRefs(node, result, new Set<string>())
	return result
}

function collectStructuredRefs(
	node: FormulaNode,
): Array<Extract<FormulaNode, { type: 'structuredRef' }>> {
	const result: Array<Extract<FormulaNode, { type: 'structuredRef' }>> = []
	walkStructuredRefs(node, result)
	return result
}

function walkStructuredRefs(
	node: FormulaNode,
	result: Array<Extract<FormulaNode, { type: 'structuredRef' }>>,
): void {
	switch (node.type) {
		case 'structuredRef':
			result.push(node)
			break
		case 'binary':
			walkStructuredRefs(node.left, result)
			walkStructuredRefs(node.right, result)
			break
		case 'unary':
			walkStructuredRefs(node.operand, result)
			break
		case 'function':
			for (const arg of node.args) walkStructuredRefs(arg, result)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walkStructuredRefs(cell, result)
			}
			break
		case 'spillRef':
			walkStructuredRefs(node.target, result)
			break
		case 'sheetSpanRef':
			walkStructuredRefs(node.target, result)
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
