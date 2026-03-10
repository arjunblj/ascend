export {
	type AnalyzedFormula,
	type AnalyzeWorkbookOptions,
	analyzeWorkbook,
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	cellHasFormula,
	createSheetNameIndex,
	type IndexedFormula,
	invalidateWorkbookAnalysis,
	resolveCellFormulaText,
	resolveFormulaDependencies,
	resolveSheetIndex,
	type WorkbookAnalysis,
	type WorkbookDependencyAnalysis,
	type WorkbookFormulaAnalysis,
} from './analysis.ts'
export { type RecalcResult, recalculate } from './calc.ts'
export { type CalcContext, defaultCalcContext } from './calc-context.ts'
export {
	type CellKey,
	cellKey,
	DependencyGraph,
	parseCellKey,
} from './dep-graph.ts'
export {
	type CellChange,
	cellValuesEqual,
	diffWorkbooks,
	type SheetDiff,
	type WorkbookDiff,
} from './diff.ts'
export { type EvalContext, evaluate } from './evaluator.ts'
export { applyOperation, applyOperations, type PatchResult } from './operations.ts'
export {
	compareSnapshots,
	createSnapshot,
	type SheetSnapshot,
	type WorkbookSnapshot,
} from './snapshot.ts'
