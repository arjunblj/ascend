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
	patchWorkbookAnalysis,
	resolveCellFormulaText,
	resolveFormulaDependencies,
	resolveSheetIndex,
	type WorkbookAnalysis,
	type WorkbookDependencyAnalysis,
	type WorkbookFormulaAnalysis,
} from './analysis.ts'
export { clearCompiledFormulaCache, type RecalcResult, recalculate } from './calc.ts'
export { type CalcContext, defaultCalcContext } from './calc-context.ts'
export { type CompiledFormula, compileFormula, evaluateCompiled } from './compiled-eval.ts'
export {
	type ConditionalFormatResult,
	evaluateConditionalFormats,
} from './conditional-format.ts'
export {
	type ValidationResult,
	validateCellValue,
} from './data-validation.ts'
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
export {
	clearFormulaParseCache,
	clearRangeValueCache,
	type EvalContext,
	evaluate,
	invalidateSheetIndexCache,
	setRangeValueCache,
} from './evaluator.ts'
export {
	type ApplyOperationsErrors,
	type ApplyOperationsOptions,
	applyOperation,
	applyOperations,
	applyWithTransaction,
	type PatchResult,
} from './operations.ts'
export {
	compareSnapshots,
	createSnapshot,
	type SheetSnapshot,
	type WorkbookSnapshot,
} from './snapshot.ts'
