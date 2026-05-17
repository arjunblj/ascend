import type { CellFormulaBinding, StyleId } from '@ascend/core'
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
	cachedParseFormula,
	clearCriteriaMatchCache,
	type ExactLookupCache,
	type FormulaCellRef,
	type FormulaNode,
	type LookupVectorCache,
	type NumericVectorCache,
} from '@ascend/formulas'
import type { AscendError, CellValue } from '@ascend/schema'
import {
	dateValue,
	EMPTY,
	errorValue,
	numberValue,
	topLeftScalar,
	valuesEqual,
} from '@ascend/schema'
import {
	type AnalyzedFormula,
	analyzeWorkbook,
	type GrowingAggregateAppendIndex,
	type GrowingRangeAggregateFunction,
	invalidateWorkbookAnalysis,
} from './analysis.ts'
import type { CalcContext } from './calc-context.ts'
import {
	type CodegenFn,
	clearCodegenCache,
	codegenFormula,
	codegenSharedFormula,
} from './codegen.ts'
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
	evaluateLegacyTopLevelFormula,
	MutableEvalContext,
	setRangeValueCache,
} from './evaluator.ts'
import { createStructuredRefResolver } from './structured-refs.ts'

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

type BlockedSpillReason = 'occupied-cell' | 'sheet-edge'

interface BlockedSpillCheck {
	readonly reason: BlockedSpillReason
	readonly blockingRefs: readonly string[]
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

interface PrefixAggregateBlock {
	readonly key: CellKey
	readonly formula: AnalyzedFormula
	readonly aggregate: RangeAggregateOptimization
}

interface TextPrefixAggregateBlock {
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly formula: string
	readonly aggregate: RangeAggregateOptimization
}

interface ScalarIfAggregateFallbackFormula {
	readonly conditionRef: { row: number; col: number }
	readonly comparison: '>' | '>=' | '<' | '<=' | '=' | '<>'
	readonly threshold: number
	readonly trueRef: { row: number; col: number }
	readonly fallbackAggregate: RangeAggregateOptimization
}

const EXCEL_MAX_ROWS = 1_048_576
const EXCEL_MAX_COLS = 16_384

interface ScalarIfAggregateFallbackBlock {
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly formula: string
	readonly parsed: ScalarIfAggregateFallbackFormula
}

interface TailPrefixAggregateResult {
	readonly value: CellValue
	readonly state?: RangeAggregateState
}

interface IndexMatchReturnPattern {
	readonly sheetIndex: number
	readonly returnStartRow: number
	readonly returnCol: number
	readonly returnEndRow: number
	readonly lookupCellRow: number
	readonly lookupCellCol: number
	readonly lookupStartRow: number
	readonly lookupCol: number
	readonly lookupEndRow: number
}

interface IndexMatchReturnFormula {
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly formula: string
	readonly pattern: IndexMatchReturnPattern
}

type TextPrefixSourceSnapshot = number | null | 'error'

interface RecalcScratch {
	readonly spillBySheet: Map<number, Map<string, SpillEntry[]>>
	readonly spillInitializedSheets: Set<number>
	readonly exactLookupCache: ExactLookupCache
	readonly lookupVectorCache: LookupVectorCache
	readonly aggregateRangeCache: AggregateRangeCache
	readonly numericVectorCache: NumericVectorCache
	readonly rangeValueCache: Map<number, readonly (readonly CellValue[])[]>
	readonly growingAggregateStateCache: Map<CellKey, RangeAggregateState>
	readonly textPrefixTailIndex: Map<CellKey, TextPrefixAggregateBlock>
	readonly textPrefixFormulaByGroupEnd: Map<string, TextPrefixAggregateBlock>
	readonly textPrefixGroups: Map<string, readonly TextPrefixAggregateBlock[]>
	readonly textPrefixAggregateStates: Map<string, RangeAggregateState>
	readonly textPrefixSourceSnapshots: Map<CellKey, TextPrefixSourceSnapshot>
	readonly indexMatchReturnBySource: Map<CellKey, readonly IndexMatchReturnFormula[]>
	indexMatchReturnFormulaCount: number
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
			textPrefixTailIndex: new Map(),
			textPrefixFormulaByGroupEnd: new Map(),
			textPrefixGroups: new Map(),
			textPrefixAggregateStates: new Map(),
			textPrefixSourceSnapshots: new Map(),
			indexMatchReturnBySource: new Map(),
			indexMatchReturnFormulaCount: 0,
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
	sheet.cells.forEachValueInRange(startRow, startCol, endRow, endCol, (cellValue) => {
		if (error) return
		const rangeValue = rangeAggregateNumericValue(cellValue)
		if (typeof rangeValue === 'number') {
			switch (functionName) {
				case 'SUM':
					sum += rangeValue
					break
				case 'COUNT':
					count++
					break
				case 'AVERAGE':
					sum += rangeValue
					count++
					break
				case 'MIN':
					count++
					if (rangeValue < min) min = rangeValue
					break
				case 'MAX':
					count++
					if (rangeValue > max) max = rangeValue
					break
			}
			return
		}
		if (rangeValue?.kind === 'error' && functionName !== 'COUNT') error = rangeValue
	})
	return { sum, count, min, max, error }
}

function rangeAggregateNumericValue(value: CellValue): number | CellValue | null {
	const scalar = topLeftScalar(value)
	if (scalar.kind === 'number') return scalar.value
	if (scalar.kind === 'date') return scalar.serial
	if (scalar.kind === 'error') return scalar
	return null
}

function readRangeAggregateNumericCell(
	sheet: Workbook['sheets'][number],
	row: number,
	col: number,
): number | CellValue | null {
	const kind = sheet.cells.readKind(row, col)
	if (kind === undefined || kind === 'empty') return null
	if (kind === 'number' || kind === 'date') return sheet.cells.readNumber(row, col) ?? 0
	if (kind === 'error') return sheet.cells.readValue(row, col)
	const scalar = topLeftScalar(sheet.cells.readValue(row, col))
	if (scalar.kind === 'number') return scalar.value
	if (scalar.kind === 'date') return scalar.serial
	if (scalar.kind === 'error') return scalar
	return null
}

function normalizeFormulaResultForExistingCell(
	oldValue: CellValue,
	newValue: CellValue,
): CellValue {
	if (oldValue.kind === 'date' && newValue.kind === 'number') {
		return dateValue(newValue.value)
	}
	return newValue
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
			const rangeValue = readRangeAggregateNumericCell(appendSheet, row, col)
			if (typeof rangeValue === 'number') {
				switch (optimization.functionName) {
					case 'SUM':
						value += rangeValue
						break
					case 'COUNT':
						value += 1
						break
					case 'MIN':
						if (value === 0 && rangeValue > 0) return null
						value = Math.min(value, rangeValue)
						break
					case 'MAX':
						if (value === 0 && rangeValue < 0) return null
						value = Math.max(value, rangeValue)
						break
				}
				continue
			}
			if (rangeValue?.kind === 'error') {
				if (optimization.functionName !== 'COUNT') return rangeValue
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

interface SharedRelativeBinaryGroup {
	readonly sheetIndex: number
	readonly anchorRow: number
	readonly anchorCol: number
	readonly op: '+' | '-' | '*' | '/'
	readonly left: FormulaCellRef
	readonly right: FormulaCellRef
	readonly members: SharedRelativeBinaryMember[]
}

interface SharedRelativeBinaryMember {
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly formula: string | null
	readonly formulaInfo: CellFormulaBinding
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

function tryFastFullSharedRelativeBinaryRecalc(
	workbook: Workbook,
	changed: string[],
	start: number,
): RecalcResult | null {
	const groups = new Map<string, SharedRelativeBinaryGroup>()
	let formulaCount = 0
	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, entries] of sheet.cells.iterateRows()) {
			for (const [col, cell] of entries) {
				if (cell.formula === null && cell.formulaInfo === undefined) continue
				const binding = cell.formulaInfo
				if (binding?.kind !== 'shared' || binding.sharedIndex === undefined) return null
				const anchor = parseSharedAnchorRef(binding.masterRef, row, col)
				const masterFormula =
					binding.isMaster && cell.formula
						? cell.formula
						: (workbook.sheets[sheetIndex]?.cells.readFormula(anchor.row, anchor.col) ?? null)
				if (!masterFormula) return null
				const template = parseSharedRelativeBinaryTemplate(masterFormula)
				if (!template) return null
				const groupKey = `${sheetIndex}:${binding.sharedIndex}`
				let group = groups.get(groupKey)
				if (!group) {
					group = {
						sheetIndex,
						anchorRow: anchor.row,
						anchorCol: anchor.col,
						...template,
						members: [],
					}
					groups.set(groupKey, group)
				} else if (
					group.sheetIndex !== sheetIndex ||
					group.anchorRow !== anchor.row ||
					group.anchorCol !== anchor.col ||
					group.op !== template.op ||
					!formulaRefsEqual(group.left, template.left) ||
					!formulaRefsEqual(group.right, template.right)
				) {
					return null
				}
				group.members.push({ sheetIndex, row, col, formula: cell.formula, formulaInfo: binding })
				formulaCount++
			}
		}
	}
	if (formulaCount === 0) return null
	for (const group of groups.values()) {
		for (const member of group.members) {
			if (!sharedRelativeBinarySourcesAreNumeric(workbook, group, member)) return null
		}
	}
	for (const group of groups.values()) {
		for (const member of group.members) {
			const sheet = workbook.sheets[member.sheetIndex]
			if (!sheet) return null
			const newValue = evaluateSharedRelativeBinary(workbook, group, member)
			if (!newValue) return null
			const oldValue = sheet.cells.readValue(member.row, member.col)
			if (valuesEqual(oldValue, newValue)) continue
			const oldStyleId = sheet.cells.readStyleId(member.row, member.col) ?? DEFAULT_STYLE_ID
			if (newValue.kind === 'number') {
				sheet.cells.setNumberResolved(
					member.row,
					member.col,
					newValue.value,
					member.formula,
					oldStyleId,
					member.formulaInfo,
				)
			} else {
				sheet.cells.setResolved(
					member.row,
					member.col,
					newValue,
					member.formula,
					oldStyleId,
					member.formulaInfo,
				)
			}
			changed.push(cellRefString(workbook, member.sheetIndex, member.row, member.col))
		}
	}
	return { changed, errors: [], duration: performance.now() - start }
}

function parseSharedRelativeBinaryTemplate(formula: string): {
	readonly op: '+' | '-' | '*' | '/'
	readonly left: FormulaCellRef
	readonly right: FormulaCellRef
} | null {
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return null
	const ast = parsed.value
	if (ast.type !== 'binary') return null
	if (ast.op !== '+' && ast.op !== '-' && ast.op !== '*' && ast.op !== '/') return null
	if (ast.left.type !== 'cellRef' || ast.right.type !== 'cellRef') return null
	if (ast.left.sheet !== undefined || ast.right.sheet !== undefined) return null
	return { op: ast.op, left: ast.left.ref, right: ast.right.ref }
}

function formulaRefsEqual(left: FormulaCellRef, right: FormulaCellRef): boolean {
	return (
		left.row === right.row &&
		left.col === right.col &&
		left.rowAbsolute === right.rowAbsolute &&
		left.colAbsolute === right.colAbsolute
	)
}

function sharedRelativeBinarySourcesAreNumeric(
	workbook: Workbook,
	group: SharedRelativeBinaryGroup,
	member: SharedRelativeBinaryMember,
): boolean {
	const left = readSharedRelativeBinarySource(workbook, group, member, group.left)
	if (left?.kind !== 'number') return false
	const right = readSharedRelativeBinarySource(workbook, group, member, group.right)
	return right?.kind === 'number'
}

function evaluateSharedRelativeBinary(
	workbook: Workbook,
	group: SharedRelativeBinaryGroup,
	member: SharedRelativeBinaryMember,
): CellValue | null {
	const left = readSharedRelativeBinarySource(workbook, group, member, group.left)
	const right = readSharedRelativeBinarySource(workbook, group, member, group.right)
	if (left?.kind !== 'number' || right?.kind !== 'number') return null
	switch (group.op) {
		case '+':
			return numberValue(left.value + right.value)
		case '-':
			return numberValue(left.value - right.value)
		case '*':
			return numberValue(left.value * right.value)
		case '/':
			return right.value === 0 ? errorValue('#DIV/0!') : numberValue(left.value / right.value)
	}
}

function readSharedRelativeBinarySource(
	workbook: Workbook,
	group: SharedRelativeBinaryGroup,
	member: SharedRelativeBinaryMember,
	ref: FormulaCellRef,
): CellValue | null {
	const rowOffset = member.row - group.anchorRow
	const colOffset = member.col - group.anchorCol
	const row = ref.rowAbsolute ? ref.row : ref.row + rowOffset
	const col = ref.colAbsolute ? ref.col : ref.col + colOffset
	const sheet = workbook.sheets[group.sheetIndex]
	if (!sheet) return null
	if (row < 0 || col < 0) return null
	const sourceFormula = sheet.cells.readFormula(row, col)
	if (sourceFormula !== null && sourceFormula !== undefined) return null
	if (sheet.cells.readFormulaInfo(row, col) !== undefined) return null
	return sheet.cells.readValue(row, col)
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
	'FREQUENCY',
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

const LEGACY_SINGLE_CELL_ARRAY_FUNCTIONS = new Set([
	'FREQUENCY',
	'GROWTH',
	'INDEX',
	'LINEST',
	'LOGEST',
	'MINVERSE',
	'MMULT',
	'TREND',
])

const REFERENCE_RETURNING_FUNCTIONS = new Set(['INDEX', 'INDIRECT', 'OFFSET'])

const ARRAY_MAPPING_FUNCTIONS = new Set([
	'ABS',
	'ACOS',
	'ACOSH',
	'ACOT',
	'ACOTH',
	'ASIN',
	'ASINH',
	'ATAN',
	'ATAN2',
	'ATANH',
	'ARABIC',
	'BASE',
	'BESSELI',
	'BESSELJ',
	'BESSELK',
	'BESSELY',
	'BETADIST',
	'BETA.DIST',
	'BETAINV',
	'BETA.INV',
	'BIN2DEC',
	'BIN2HEX',
	'BIN2OCT',
	'BINOMDIST',
	'BINOM.DIST',
	'BINOM.DIST.RANGE',
	'BINOM.INV',
	'BITAND',
	'BITLSHIFT',
	'BITOR',
	'BITRSHIFT',
	'BITXOR',
	'CEILING',
	'CEILING.MATH',
	'CEILING.PRECISE',
	'CHAR',
	'CHIDIST',
	'CHIINV',
	'CHISQ.DIST',
	'CHISQ.DIST.RT',
	'CHISQ.INV',
	'CHISQ.INV.RT',
	'CLEAN',
	'CODE',
	'COLUMN',
	'COMBIN',
	'COMBINA',
	'COMPLEX',
	'CONFIDENCE',
	'CONFIDENCE.NORM',
	'CONFIDENCE.T',
	'CONVERT',
	'COS',
	'COSH',
	'COT',
	'COTH',
	'CSC',
	'CSCH',
	'CUMIPMT',
	'CUMPRINC',
	'CRITBINOM',
	'DATE',
	'DATEVALUE',
	'DATEDIF',
	'DAY',
	'DAYS',
	'DAYS360',
	'DB',
	'DECIMAL',
	'DEC2BIN',
	'DEC2HEX',
	'DEC2OCT',
	'DEGREES',
	'DELTA',
	'DDB',
	'DOLLAR',
	'DOLLARDE',
	'DOLLARFR',
	'EDATE',
	'EFFECT',
	'EOMONTH',
	'ERROR.TYPE',
	'ERF',
	'ERF.PRECISE',
	'ERFC',
	'ERFC.PRECISE',
	'EVEN',
	'EXACT',
	'EXP',
	'EXPONDIST',
	'EXPON.DIST',
	'FACT',
	'FACTDOUBLE',
	'FDIST',
	'F.DIST',
	'F.DIST.RT',
	'FIND',
	'FINV',
	'F.INV',
	'F.INV.RT',
	'FISHER',
	'FISHERINV',
	'FIXED',
	'FLOOR',
	'FLOOR.MATH',
	'FLOOR.PRECISE',
	'FV',
	'GAMMADIST',
	'GAMMA',
	'GAMMA.DIST',
	'GAMMAINV',
	'GAMMA.INV',
	'GAMMALN',
	'GAMMALN.PRECISE',
	'GAUSS',
	'GESTEP',
	'HEX2BIN',
	'HEX2DEC',
	'HEX2OCT',
	'HOUR',
	'HYPGEOM.DIST',
	'HYPGEOMDIST',
	'IF',
	'IFERROR',
	'IFNA',
	'IFS',
	'IMABS',
	'IMAGINARY',
	'IMARGUMENT',
	'IMCONJUGATE',
	'IMCOS',
	'IMCOSH',
	'IMCOT',
	'IMCSC',
	'IMCSCH',
	'IMDIV',
	'IMEXP',
	'IMLN',
	'IMLOG10',
	'IMLOG2',
	'IMPOWER',
	'IMREAL',
	'IMSEC',
	'IMSECH',
	'IMSIN',
	'IMSINH',
	'IMSQRT',
	'IMSUB',
	'IMTAN',
	'INT',
	'IPMT',
	'ISBLANK',
	'ISERR',
	'ISERROR',
	'ISEVEN',
	'ISFORMULA',
	'ISLOGICAL',
	'ISNA',
	'ISNONTEXT',
	'ISNUMBER',
	'ISODD',
	'ISPMT',
	'ISOWEEKNUM',
	'ISTEXT',
	'LEFT',
	'LEN',
	'LN',
	'LOG',
	'LOG10',
	'LOGINV',
	'LOGNORMDIST',
	'LOGNORM.DIST',
	'LOGNORM.INV',
	'LOWER',
	'MID',
	'MINUTE',
	'MOD',
	'MONTH',
	'MROUND',
	'N',
	'NETWORKDAYS',
	'NETWORKDAYS.INTL',
	'NEGBINOMDIST',
	'NEGBINOM.DIST',
	'NOMINAL',
	'NORMDIST',
	'NORM.DIST',
	'NORMINV',
	'NORM.INV',
	'NORMSDIST',
	'NORM.S.DIST',
	'NORMSINV',
	'NORM.S.INV',
	'NOT',
	'NPER',
	'ODD',
	'OCT2BIN',
	'OCT2DEC',
	'OCT2HEX',
	'PHI',
	'PMT',
	'PDURATION',
	'POISSON',
	'POISSON.DIST',
	'POWER',
	'PPMT',
	'PROPER',
	'PV',
	'QUOTIENT',
	'RADIANS',
	'RATE',
	'REPLACE',
	'REPT',
	'RIGHT',
	'ROMAN',
	'ROUND',
	'ROUNDDOWN',
	'ROUNDUP',
	'RRI',
	'ROW',
	'SEARCH',
	'SEC',
	'SECH',
	'SECOND',
	'SIGN',
	'SIN',
	'SINH',
	'SLN',
	'SQRT',
	'SQRTPI',
	'STANDARDIZE',
	'SUBSTITUTE',
	'SYD',
	'SWITCH',
	'TAN',
	'TANH',
	'TDIST',
	'T.DIST',
	'T.DIST.2T',
	'T.DIST.RT',
	'TEXT',
	'TEXTAFTER',
	'TEXTBEFORE',
	'TIME',
	'TIMEVALUE',
	'TINV',
	'T.INV',
	'T.INV.2T',
	'TRIM',
	'TRUNC',
	'UNICHAR',
	'UNICODE',
	'UPPER',
	'VALUE',
	'VDB',
	'WEIBULL',
	'WEIBULL.DIST',
	'WEEKDAY',
	'WEEKNUM',
	'WORKDAY',
	'WORKDAY.INTL',
	'YEAR',
	'YEARFRAC',
])

const REFERENCE_SENSITIVE_SHARED_FUNCTIONS = new Set(['INDIRECT', 'OFFSET'])

function formulaContainsFunction(node: FormulaNode, names: ReadonlySet<string>): boolean {
	switch (node.type) {
		case 'function':
			return (
				names.has(node.name.toUpperCase()) ||
				node.args.some((arg) => formulaContainsFunction(arg, names))
			)
		case 'binary':
			return formulaContainsFunction(node.left, names) || formulaContainsFunction(node.right, names)
		case 'unary':
			return formulaContainsFunction(node.operand, names)
		case 'array':
			return node.rows.some((row) => row.some((cell) => formulaContainsFunction(cell, names)))
		default:
			return false
	}
}

function needsInterpreterArraySemantics(ast: FormulaNode): boolean {
	return expressionNeedsInterpreterArraySemantics(ast)
}

function expressionNeedsInterpreterArraySemantics(node: FormulaNode): boolean {
	switch (node.type) {
		case 'binary':
			return (
				(node.op !== ',' &&
					node.op !== ' ' &&
					(nodeCanReturnArray(node.left) || nodeCanReturnArray(node.right))) ||
				expressionNeedsInterpreterArraySemantics(node.left) ||
				expressionNeedsInterpreterArraySemantics(node.right)
			)
		case 'unary':
			return (
				nodeCanReturnArray(node.operand) || expressionNeedsInterpreterArraySemantics(node.operand)
			)
		case 'function':
			return functionCanReturnArray(node) || node.args.some(nodeCanReturnComputedArray)
		case 'array':
			return true
		default:
			return false
	}
}

function nodeCanReturnComputedArray(node: FormulaNode): boolean {
	switch (node.type) {
		case 'array':
			return true
		case 'binary':
			return (
				(node.op !== ',' &&
					node.op !== ' ' &&
					(nodeCanReturnArray(node.left) || nodeCanReturnArray(node.right))) ||
				nodeCanReturnComputedArray(node.left) ||
				nodeCanReturnComputedArray(node.right)
			)
		case 'unary':
			return nodeCanReturnArray(node.operand) || nodeCanReturnComputedArray(node.operand)
		case 'function':
			return functionCanReturnArray(node) || node.args.some(nodeCanReturnComputedArray)
		default:
			return false
	}
}

function nodeCanReturnArray(node: FormulaNode): boolean {
	switch (node.type) {
		case 'array':
		case 'rangeRef':
		case 'dynamicRangeRef':
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
			return functionCanReturnArray(node)
		default:
			return false
	}
}

function functionCanReturnArray(node: Extract<FormulaNode, { type: 'function' }>): boolean {
	const name = node.name.toUpperCase()
	return (
		ARRAY_RETURNING_FUNCTIONS.has(name) ||
		(ARRAY_MAPPING_FUNCTIONS.has(name) && node.args.some(nodeCanReturnArray))
	)
}

function hasExternalWorkbookReference(node: FormulaNode): boolean {
	switch (node.type) {
		case 'cellRef':
		case 'rangeRef':
		case 'wholeColumnRange':
		case 'wholeRowRange':
		case 'name':
			return isExternalSheetToken(node.sheet)
		case 'sheetSpanRef':
			return isExternalSheetToken(node.startSheet) || isExternalSheetToken(node.endSheet)
		case 'spillRef':
			return hasExternalWorkbookReference(node.target)
		case 'dynamicRangeRef':
			return hasExternalWorkbookReference(node.start) || hasExternalWorkbookReference(node.end)
		case 'binary':
			return hasExternalWorkbookReference(node.left) || hasExternalWorkbookReference(node.right)
		case 'unary':
			return hasExternalWorkbookReference(node.operand)
		case 'function':
			return node.args.some(hasExternalWorkbookReference)
		case 'array':
			return node.rows.some((row) => row.some(hasExternalWorkbookReference))
		default:
			return false
	}
}

function isExternalSheetToken(sheet: string | undefined): boolean {
	const open = sheet?.indexOf('[') ?? -1
	const close = open >= 0 && sheet ? sheet.indexOf(']', open + 1) : -1
	if (!sheet || open < 0 || close <= open) return false
	const workbook = `${sheet.slice(0, open)}${sheet.slice(open + 1, close)}`
	const sheetName = sheet.slice(close + 1)
	return workbook.length > 0 && sheetName.length > 0
}

export function clearCompiledFormulaCache(): void {
	clearCodegenCache()
}

function evalFormula(
	_key: CellKey,
	formulaText: string,
	ast: FormulaNode,
	ctx: EvalContext,
): CellValue {
	if (hasExternalWorkbookReference(ast))
		return normalizeTopLevelFormulaValue(ast, evaluate(ast, ctx), ctx)
	if (!usesArrayFormulaSemantics(ctx)) {
		const legacyValue = evaluateLegacyTopLevelFormula(ast, ctx)
		if (legacyValue) return normalizeTopLevelFormulaValue(ast, legacyValue, ctx)
	}
	if (needsInterpreterArraySemantics(ast))
		return normalizeTopLevelFormulaValue(ast, evaluate(ast, withArrayFormulaSemantics(ctx)), ctx)
	if (containsEmptyStringLiteralNumericBinary(ast)) {
		return normalizeTopLevelFormulaValue(ast, evaluate(ast, ctx), ctx)
	}
	if (shouldPreferCompiled(ast)) {
		let compiled = compiledCache.get(ast)
		if (compiled === undefined) {
			const result = compileFormula(ast)
			compiled = result ?? false
			compiledCache.set(ast, compiled)
		}
		if (compiled !== false) {
			return normalizeTopLevelFormulaValue(ast, evaluateCompiled(compiled, ctx), ctx)
		}
	}
	const generated = codegenFormula(formulaText, ast)
	if (generated) return normalizeTopLevelFormulaValue(ast, generated(ctx), ctx)
	let compiled = compiledCache.get(ast)
	if (compiled === undefined) {
		const result = compileFormula(ast)
		compiled = result ?? false
		compiledCache.set(ast, compiled)
	}
	if (compiled !== false) {
		return normalizeTopLevelFormulaValue(ast, evaluateCompiled(compiled, ctx), ctx)
	}
	return normalizeTopLevelFormulaValue(ast, evaluate(ast, ctx), ctx)
}

function withArrayFormulaSemantics(ctx: EvalContext): EvalContext {
	return ctx.arrayFormulaSemantics === true ? ctx : { ...ctx, arrayFormulaSemantics: true }
}

function containsEmptyStringLiteralNumericBinary(node: FormulaNode): boolean {
	let found = false
	function scan(n: FormulaNode): void {
		if (found) return
		if (n.type === 'binary') {
			if (
				isNumericBinaryOp(n.op) &&
				(isEmptyStringLiteral(n.left) || isEmptyStringLiteral(n.right))
			) {
				found = true
				return
			}
			scan(n.left)
			scan(n.right)
			return
		}
		if (n.type === 'unary') scan(n.operand)
		else if (n.type === 'function') {
			for (const arg of n.args) scan(arg)
		}
	}
	scan(node)
	return found
}

function isNumericBinaryOp(op: string): boolean {
	return op === '+' || op === '-' || op === '*' || op === '/' || op === '^'
}

function isEmptyStringLiteral(node: FormulaNode): boolean {
	return node.type === 'string' && node.value === ''
}

function normalizeTopLevelFormulaValue(
	ast: FormulaNode,
	value: CellValue,
	ctx: EvalContext,
): CellValue {
	if (value.kind === 'empty' && ast.type === 'cellRef') return numberValue(0)
	if (
		value.kind === 'empty' &&
		ast.type === 'function' &&
		REFERENCE_RETURNING_FUNCTIONS.has(ast.name.toUpperCase())
	) {
		return numberValue(0)
	}
	if (
		value.kind === 'array' &&
		ast.type === 'function' &&
		LEGACY_SINGLE_CELL_ARRAY_FUNCTIONS.has(ast.name.toUpperCase()) &&
		ctx.workbook.sourceArchiveBytes !== null &&
		!usesArrayFormulaSemantics(ctx)
	) {
		if (ast.name.toUpperCase() === 'INDEX') return errorValue('#VALUE!')
		return topLeftValue(value)
	}
	if (
		value.kind === 'array' &&
		ast.type === 'function' &&
		(ast.name.toUpperCase() === 'ROW' || ast.name.toUpperCase() === 'COLUMN') &&
		!usesArrayFormulaSemantics(ctx)
	) {
		return topLeftValue(value)
	}
	return value
}

function usesArrayFormulaSemantics(ctx: EvalContext): boolean {
	const binding = ctx.workbook.sheets[ctx.sheetIndex]?.cells.get(ctx.row, ctx.col)?.formulaInfo
	return (
		binding?.kind === 'array' ||
		binding?.kind === 'dynamicArray' ||
		binding?.kind === 'spill' ||
		binding?.kind === 'blockedSpill'
	)
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

function persistentScalarFormulaBinding(
	binding: CellFormulaBinding | undefined,
): CellFormulaBinding | undefined {
	return binding?.kind === 'spill' || binding?.kind === 'blockedSpill' ? undefined : binding
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

function findBlockedSpill(
	sheet: Workbook['sheets'][number],
	anchorRow: number,
	anchorCol: number,
	anchorRef: string,
	matrix: readonly (readonly CellValue[])[],
): BlockedSpillCheck | null {
	if (!sheet) return { reason: 'sheet-edge', blockingRefs: [] }
	const blocked: string[] = []
	let blockedBySheetEdge = false
	for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) {
		const sourceRow = matrix[rowOffset] ?? []
		for (let colOffset = 0; colOffset < sourceRow.length; colOffset++) {
			const targetRow = anchorRow + rowOffset
			const targetCol = anchorCol + colOffset
			if (!isWithinExcelGrid(targetRow, targetCol)) {
				blockedBySheetEdge = true
				continue
			}
			if (cellHasSpillLayoutBlocker(sheet, targetRow, targetCol)) {
				blocked.push(toA1Ref(targetRow, targetCol))
				continue
			}
			if (rowOffset === 0 && colOffset === 0) continue
			const existingFormulaInfo = sheet.cells.readFormulaInfo(targetRow, targetCol)
			if (existingFormulaInfo === undefined && !sheet.cells.has(targetRow, targetCol)) continue
			if (
				existingFormulaInfo &&
				isSpillBinding(existingFormulaInfo) &&
				existingFormulaInfo.anchorRef === anchorRef
			) {
				continue
			}
			blocked.push(toA1Ref(targetRow, targetCol))
		}
	}
	if (blockedBySheetEdge) return { reason: 'sheet-edge', blockingRefs: [] }
	return blocked.length > 0 ? { reason: 'occupied-cell', blockingRefs: blocked } : null
}

function isWithinExcelGrid(row: number, col: number): boolean {
	return row >= 0 && row < EXCEL_MAX_ROWS && col >= 0 && col < EXCEL_MAX_COLS
}

function cellHasSpillLayoutBlocker(
	sheet: Workbook['sheets'][number],
	row: number,
	col: number,
): boolean {
	return (
		sheet.merges.some((merge) => rangeContainsCell(merge, row, col)) ||
		sheet.tables.some((table) => rangeContainsCell(table.ref, row, col))
	)
}

function rangeContainsCell(range: RangeRef, row: number, col: number): boolean {
	return (
		row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col
	)
}

function stringArraysEqual(a: readonly string[], b: readonly string[]): boolean {
	if (a.length !== b.length) return false
	for (let i = 0; i < a.length; i++) {
		if (a[i] !== b[i]) return false
	}
	return true
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
	const blockedSpill = findBlockedSpill(sheet, row, col, anchorRef, matrix)
	if (blockedSpill) {
		const spillError = errorValue('#SPILL!')
		const currentValue = sheet.cells.readValue(row, col)
		const currentBinding = sheet.cells.readFormulaInfo(row, col)
		const blockedInfo = {
			kind: 'blockedSpill' as const,
			anchorRef,
			ref: spillRef,
			...(blockedSpill.reason === 'sheet-edge' ? { reason: blockedSpill.reason } : {}),
			blockingRefs: blockedSpill.blockingRefs,
		}
		const alreadyBlocked =
			currentBinding?.kind === 'blockedSpill' &&
			currentBinding.ref === spillRef &&
			(currentBinding.reason ?? 'occupied-cell') === blockedSpill.reason &&
			stringArraysEqual(currentBinding.blockingRefs, blockedSpill.blockingRefs) &&
			valuesEqual(currentValue, spillError)
		if (!alreadyBlocked) {
			sheet.cells.setResolved(row, col, spillError, oldFormula, oldStyleId, blockedInfo)
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

function applyLegacyArrayResult(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	oldFormula: string | null,
	oldStyleId: StyleId,
	formulaInfo: CellFormulaBinding & { kind: 'array'; ref: string },
	matrix: readonly (readonly CellValue[])[],
	changed: string[],
): CellValue {
	const range = parseRange(formulaInfo.ref)
	const targetSheetIndex =
		range.sheet === undefined
			? sheetIndex
			: workbook.sheets.findIndex(
					(sheet) => sheet?.name.toLowerCase() === range.sheet?.toLowerCase(),
				)
	if (targetSheetIndex < 0) return errorValue('#REF!')
	const sheet = workbook.sheets[targetSheetIndex]
	if (!sheet) return errorValue('#REF!')
	const anchorValue = topLeftValue(matrix[0]?.[0] ?? EMPTY)
	const binding = { kind: 'array' as const, ref: formulaInfo.ref }

	for (let targetRow = range.start.row; targetRow <= range.end.row; targetRow++) {
		for (let targetCol = range.start.col; targetCol <= range.end.col; targetCol++) {
			const rowOffset = targetRow - range.start.row
			const colOffset = targetCol - range.start.col
			const nextValue = topLeftValue(matrix[rowOffset]?.[colOffset] ?? EMPTY)
			const targetFormula = targetRow === row && targetCol === col ? oldFormula : null
			const targetStyleId = sheet.cells.readStyleId(targetRow, targetCol) ?? oldStyleId
			const currentValue = sheet.cells.readValue(targetRow, targetCol)
			const currentFormula = sheet.cells.readFormula(targetRow, targetCol)
			const currentInfo = sheet.cells.readFormulaInfo(targetRow, targetCol)
			const isSameArrayInfo =
				currentInfo?.kind === 'array' && (currentInfo.ref ?? formulaInfo.ref) === formulaInfo.ref
			if (
				valuesEqual(currentValue, nextValue) &&
				currentFormula === targetFormula &&
				isSameArrayInfo
			) {
				continue
			}
			if (nextValue.kind === 'number') {
				sheet.cells.setNumberResolved(
					targetRow,
					targetCol,
					nextValue.value,
					targetFormula,
					targetStyleId,
					binding,
				)
			} else {
				sheet.cells.setResolved(
					targetRow,
					targetCol,
					nextValue,
					targetFormula,
					targetStyleId,
					binding,
				)
			}
			changed.push(cellRefString(workbook, targetSheetIndex, targetRow, targetCol))
		}
	}
	return anchorValue
}

function applyArrayOrLegacyResult(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	oldFormula: string | null,
	oldStyleId: StyleId,
	oldFormulaInfo: CellFormulaBinding | undefined,
	matrix: readonly (readonly CellValue[])[],
	changed: string[],
	spillIndex: SpillIndexState,
): CellValue {
	if (oldFormulaInfo?.kind === 'array' && oldFormulaInfo.ref) {
		return applyLegacyArrayResult(
			workbook,
			sheetIndex,
			row,
			col,
			oldFormula,
			oldStyleId,
			oldFormulaInfo as CellFormulaBinding & { kind: 'array'; ref: string },
			matrix,
			changed,
		)
	}
	return applyArrayResult(
		workbook,
		sheetIndex,
		row,
		col,
		oldFormula,
		oldStyleId,
		matrix,
		changed,
		spillIndex,
	)
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
	if (!workbookHasFormulaInfo(workbook) && spillIndex.bySheet.size === 0) return
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

function workbookHasFormulaInfo(workbook: Workbook): boolean {
	for (const sheet of workbook.sheets) {
		if (sheet?.cells.formulaInfoCellCount() > 0) return true
	}
	return false
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
	mutableCtx.structuredRefResolver = createStructuredRefResolver(workbook)
	setRangeValueCache(scratch.rangeValueCache)
	const isDirtyRecalc = opts?.dirtyOnly || (opts?.dirtyRefs?.length ?? 0) > 0

	clearOrphanedSpills(workbook, spillIndex, changed)
	if (!isDirtyRecalc && !opts?.range) {
		exactLookupCache.clear()
		lookupVectorCache.clear()
		scratch.growingAggregateStateCache.clear()
		scratch.indexMatchReturnBySource.clear()
		scratch.indexMatchReturnFormulaCount = 0
		const linearChainFast = tryFastFullPreviousRowAddendRecalc(workbook, changed, start)
		if (linearChainFast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return linearChainFast
		}
		const scalarIfFast = tryFastFullScalarIfAggregateFallbackTextRecalc(workbook, changed, start)
		if (scalarIfFast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return scalarIfFast
		}
		const fast = tryFastFullPrefixAggregateTextRecalc(workbook, changed, start, scratch)
		if (fast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return fast
		}
		const sharedFast = tryFastFullSharedRelativeBinaryRecalc(workbook, changed, start)
		if (sharedFast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return sharedFast
		}
	}
	if (isDirtyRecalc && !opts?.range) {
		const fast = tryFastDirtyPrefixAggregateTextRecalc(
			workbook,
			opts?.dirtyRefs,
			changed,
			start,
			scratch,
		)
		if (fast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return fast
		}
		const lookupFast = tryFastDirtyIndexMatchReturnTextRecalc(
			workbook,
			opts?.dirtyRefs,
			changed,
			start,
			scratch,
		)
		if (lookupFast) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return lookupFast
		}
	}

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
			tryFastDirtySingleFormulaRecalc(
				workbook,
				graph,
				analysis.formulas,
				dirtyRefKeys,
				mutableCtx,
				start,
			) ??
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
		const blockCompletedKeys =
			!isDirtyRecalc && cycleKeys.size === 0 && rangeAggregates.size > 0
				? tryFastFullPrefixAggregateBlocks(
						workbook,
						analysis.formulas,
						evalOrderSet,
						rangeAggregateStates,
						changed,
					)
				: null
		if (blockCompletedKeys && completedKeys) {
			for (const key of blockCompletedKeys) completedKeys.add(key)
		}
		if (blockCompletedKeys?.size === evalOrder.length) {
			clearRangeValueCache()
			clearCriteriaMatchCache()
			return {
				changed,
				errors,
				duration: performance.now() - start,
			}
		}
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
					evaluator = formulaContainsFunction(masterAst, REFERENCE_SENSITIVE_SHARED_FUNCTIONS)
						? null
						: codegenSharedFormula(masterFormulaText, masterAst, anchor)
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
					if (hadCell && !clearedSpill && oldValue.kind !== 'empty') {
						return
					}
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
			const liveFormula = sheet.cells.readFormula(row, col)
			const liveFormulaInfo = sheet.cells.readFormulaInfo(row, col)
			if (liveFormula === null && liveFormulaInfo?.kind !== 'shared') return
			const hadCell = sheet.cells.has(row, col)
			const oldValue = sheet.cells.readValue(row, col)
			const oldFormula = liveFormula ?? null
			const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
			const oldFormulaInfo = liveFormulaInfo

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
				if (
					hasExternalWorkbookReference(ast) &&
					!ctx.externalReferences &&
					hadCell &&
					oldValue.kind !== 'empty'
				) {
					newValue = oldValue
				} else {
					mutableCtx.sheetIndex = si
					mutableCtx.row = row
					mutableCtx.col = col
					newValue = groupEvaluator
						? normalizeTopLevelFormulaValue(ast, groupEvaluator(mutableCtx), mutableCtx)
						: evalFormula(key, formulaText, ast, mutableCtx)
				}
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
			newValue = normalizeFormulaResultForExistingCell(oldValue, newValue)
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				const changedBefore = changed.length
				applyArrayOrLegacyResult(
					workbook,
					si,
					row,
					col,
					oldFormula,
					oldStyleId,
					oldFormulaInfo,
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
				const formulaInfo = persistentScalarFormulaBinding(oldFormulaInfo)
				if (newValue.kind === 'number') {
					sheet.cells.setNumberResolved(
						row,
						col,
						newValue.value,
						oldFormula,
						oldStyleId,
						formulaInfo,
					)
				} else {
					sheet.cells.setResolved(row, col, newValue, oldFormula, oldStyleId, formulaInfo)
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
				if (blockCompletedKeys?.has(key)) continue
				evalCell(key)
			}
		} else {
			const processed = new Set<CellKey>()
			for (const key of evalOrder) {
				if (processed.has(key)) continue
				if (blockCompletedKeys?.has(key)) continue
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
	if (!isDirtyRecalc && !opts?.range) {
		cacheIndexMatchReturnFormulas(workbook, scratch, analysis.formulas)
	}

	return {
		changed,
		errors,
		duration: performance.now() - start,
	}
}

function tryFastFullPreviousRowAddendRecalc(
	workbook: Workbook,
	changed: string[],
	start: number,
): RecalcResult | null {
	const formulas: Array<{
		readonly sheetIndex: number
		readonly row: number
		readonly col: number
		readonly formula: string
		readonly addend: number
	}> = []
	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, entries] of sheet.cells.iterateRows()) {
			for (const [col, cell] of entries) {
				if (cell.formula === null && cell.formulaInfo === undefined) continue
				if (cell.formula === null || cell.formulaInfo !== undefined) return null
				const addend = parsePreviousRowAddendFormula(cell.formula, row, col)
				if (addend === undefined) return null
				formulas.push({ sheetIndex, row, col, formula: cell.formula, addend })
			}
		}
	}
	if (formulas.length === 0) return null
	const plannedNumbers = new Map<CellKey, number>()
	const plans: Array<{
		readonly sheetIndex: number
		readonly row: number
		readonly col: number
		readonly formula: string
		readonly nextNumber: number
	}> = []
	for (const formula of formulas) {
		const sheet = workbook.sheets[formula.sheetIndex]
		if (!sheet) return null
		const sourceKey = cellKey(formula.sheetIndex, formula.row - 1, formula.col)
		const plannedPrevious = plannedNumbers.get(sourceKey)
		let previousNumber = plannedPrevious
		if (previousNumber === undefined) {
			const previous = sheet.cells.readValue(formula.row - 1, formula.col)
			if (previous?.kind !== 'number') return null
			previousNumber = previous.value
		}
		const nextNumber = previousNumber + formula.addend
		plannedNumbers.set(cellKey(formula.sheetIndex, formula.row, formula.col), nextNumber)
		plans.push({
			sheetIndex: formula.sheetIndex,
			row: formula.row,
			col: formula.col,
			formula: formula.formula,
			nextNumber,
		})
	}
	for (const plan of plans) {
		const sheet = workbook.sheets[plan.sheetIndex]
		if (!sheet) return null
		const oldValue = sheet.cells.readValue(plan.row, plan.col)
		const newValue = numberValue(plan.nextNumber)
		if (valuesEqual(oldValue, newValue)) continue
		const oldStyleId = sheet.cells.readStyleId(plan.row, plan.col) ?? DEFAULT_STYLE_ID
		sheet.cells.setNumberResolved(plan.row, plan.col, plan.nextNumber, plan.formula, oldStyleId)
		changed.push(cellRefString(workbook, plan.sheetIndex, plan.row, plan.col))
	}
	return { changed, errors: [], duration: performance.now() - start }
}

function parsePreviousRowAddendFormula(
	formula: string,
	row: number,
	col: number,
): number | undefined {
	if (row <= 0) return undefined
	const text = formula.charCodeAt(0) === 61 ? formula.slice(1) : formula
	const ref = `${indexToColumn(col)}${row}`
	if (!text.startsWith(ref)) return undefined
	const op = text.charCodeAt(ref.length)
	if (op !== 43 && op !== 45) return undefined
	const addend = parseFiniteNumberLiteral(text, ref.length + 1)
	if (!Number.isFinite(addend)) return undefined
	return op === 45 ? -addend : addend
}

function parseFiniteNumberLiteral(text: string, offset: number): number {
	let index = offset
	let sawDigit = false
	let sawDot = false
	for (; index < text.length; index++) {
		const char = text.charCodeAt(index)
		if (char >= 48 && char <= 57) {
			sawDigit = true
			continue
		}
		if (char === 46 && !sawDot) {
			sawDot = true
			continue
		}
		break
	}
	if (!sawDigit) return Number.NaN
	if (index < text.length) {
		const exponent = text.charCodeAt(index)
		if (exponent !== 69 && exponent !== 101) return Number.NaN
		index++
		const sign = text.charCodeAt(index)
		if (sign === 43 || sign === 45) index++
		let sawExponentDigit = false
		for (; index < text.length; index++) {
			const char = text.charCodeAt(index)
			if (char < 48 || char > 57) return Number.NaN
			sawExponentDigit = true
		}
		if (!sawExponentDigit) return Number.NaN
	}
	return Number(text.slice(offset))
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

function tryFastFullScalarIfAggregateFallbackTextRecalc(
	workbook: Workbook,
	changed: string[],
	start: number,
): RecalcResult | null {
	const blocks: ScalarIfAggregateFallbackBlock[] = []
	const formulaKeys = new Set<CellKey>()
	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, entries] of sheet.cells.iterateRows()) {
			for (const [col, cell] of entries) {
				if (cell.formula === null) {
					if (cell.formulaInfo !== undefined) return null
					continue
				}
				if (cell.formulaInfo !== undefined) return null
				const parsed = parseScalarIfAggregateFallbackFormula(cell.formula, sheetIndex)
				if (!parsed) return null
				formulaKeys.add(cellKey(sheetIndex, row, col))
				blocks.push({ sheetIndex, row, col, formula: cell.formula, parsed })
			}
		}
	}
	if (blocks.length === 0) return null
	const aggregateCache = new Map<string, CellValue>()
	const aggregateReadyCache = new Map<string, boolean>()
	for (const block of blocks) {
		if (
			!scalarIfAggregateFallbackSourcesAreReady(workbook, block, formulaKeys, aggregateReadyCache)
		) {
			return null
		}
	}
	for (const block of blocks) {
		const sheet = workbook.sheets[block.sheetIndex]
		if (!sheet) return null
		const newValue = evaluateScalarIfAggregateFallback(workbook, block, aggregateCache)
		if (!newValue) return null
		const oldValue = sheet.cells.readValue(block.row, block.col)
		if (valuesEqual(oldValue, newValue)) continue
		const oldStyleId = sheet.cells.readStyleId(block.row, block.col) ?? DEFAULT_STYLE_ID
		if (newValue.kind === 'number') {
			sheet.cells.setNumberResolved(block.row, block.col, newValue.value, block.formula, oldStyleId)
		} else {
			sheet.cells.setResolved(block.row, block.col, newValue, block.formula, oldStyleId)
		}
		changed.push(cellRefString(workbook, block.sheetIndex, block.row, block.col))
	}
	return { changed, errors: [], duration: performance.now() - start }
}

function parseScalarIfAggregateFallbackFormula(
	formula: string,
	sheetIndex: number,
): ScalarIfAggregateFallbackFormula | null {
	const text = formula.trim().replace(/^=/, '')
	const match =
		/^IF\s*\(\s*([$]?[A-Z]+[$]?\d+)\s*(>=|<=|<>|>|<|=)\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\s*,\s*([$]?[A-Z]+[$]?\d+)\s*,\s*([A-Z]+)\s*\(\s*([$]?[A-Z]+[$]?\d+)\s*:\s*([$]?[A-Z]+[$]?\d+)\s*\)\s*\)$/i.exec(
			text,
		)
	if (!match) return null
	const [, conditionText, comparison, thresholdText, trueText, functionText, startText, endText] =
		match
	if (
		!conditionText ||
		!comparison ||
		!thresholdText ||
		!trueText ||
		!functionText ||
		!startText ||
		!endText
	) {
		return null
	}
	const conditionRef = parseSimpleCellRef(conditionText, 0, conditionText.length)
	const trueRef = parseSimpleCellRef(trueText, 0, trueText.length)
	const startRef = parseSimpleCellRef(startText, 0, startText.length)
	const endRef = parseSimpleCellRef(endText, 0, endText.length)
	if (!conditionRef || !trueRef || !startRef || !endRef) return null
	if (endRef.row < startRef.row || endRef.col < startRef.col) return null
	const functionName = parseGrowingAggregateFunction(functionText)
	if (functionName !== 'SUM') return null
	const threshold = Number(thresholdText)
	if (!Number.isFinite(threshold)) return null
	return {
		conditionRef,
		comparison: comparison as ScalarIfAggregateFallbackFormula['comparison'],
		threshold,
		trueRef,
		fallbackAggregate: {
			functionName,
			sheetIndex,
			startRow: startRef.row,
			startCol: startRef.col,
			endRow: endRef.row,
			endCol: endRef.col,
		},
	}
}

function scalarIfAggregateFallbackSourcesAreReady(
	workbook: Workbook,
	block: ScalarIfAggregateFallbackBlock,
	formulaKeys: ReadonlySet<CellKey>,
	aggregateReadyCache: Map<string, boolean>,
): boolean {
	const sheet = workbook.sheets[block.sheetIndex]
	if (!sheet) return false
	const conditionRef = block.parsed.conditionRef
	const trueRef = block.parsed.trueRef
	if (formulaKeys.has(cellKey(block.sheetIndex, conditionRef.row, conditionRef.col))) return false
	if (formulaKeys.has(cellKey(block.sheetIndex, trueRef.row, trueRef.col))) return false
	const conditionValue = readRangeAggregateNumericCell(sheet, conditionRef.row, conditionRef.col)
	if (typeof conditionValue !== 'number') return false
	const aggregate = block.parsed.fallbackAggregate
	const aggregateKey = scalarIfAggregateFallbackKey(aggregate)
	const cached = aggregateReadyCache.get(aggregateKey)
	if (cached !== undefined) return cached
	const aggregateSheet = workbook.sheets[aggregate.sheetIndex]
	if (!aggregateSheet) return false
	for (let row = aggregate.startRow; row <= aggregate.endRow; row++) {
		for (let col = aggregate.startCol; col <= aggregate.endCol; col++) {
			if (formulaKeys.has(cellKey(aggregate.sheetIndex, row, col))) {
				aggregateReadyCache.set(aggregateKey, false)
				return false
			}
		}
	}
	aggregateReadyCache.set(aggregateKey, true)
	return true
}

function evaluateScalarIfAggregateFallback(
	workbook: Workbook,
	block: ScalarIfAggregateFallbackBlock,
	aggregateCache: Map<string, CellValue>,
): CellValue | null {
	const sheet = workbook.sheets[block.sheetIndex]
	if (!sheet) return null
	const conditionValue = readRangeAggregateNumericCell(
		sheet,
		block.parsed.conditionRef.row,
		block.parsed.conditionRef.col,
	)
	if (typeof conditionValue !== 'number') return null
	if (compareScalarIfCondition(conditionValue, block.parsed.comparison, block.parsed.threshold)) {
		return readScalarFormulaReferenceValue(
			sheet,
			block.parsed.trueRef.row,
			block.parsed.trueRef.col,
		)
	}
	const aggregate = block.parsed.fallbackAggregate
	const aggregateKey = scalarIfAggregateFallbackKey(aggregate)
	const cached = aggregateCache.get(aggregateKey)
	if (cached) return cached
	const state = scanRangeAggregateState(
		workbook,
		aggregate.functionName,
		aggregate.sheetIndex,
		aggregate.startRow,
		aggregate.startCol,
		aggregate.endRow,
		aggregate.endCol,
	)
	if (!state) return null
	const value = rangeAggregateStateToValue(aggregate.functionName, state)
	aggregateCache.set(aggregateKey, value)
	return value
}

function readScalarFormulaReferenceValue(
	sheet: Workbook['sheets'][number],
	row: number,
	col: number,
): CellValue {
	const value = topLeftScalar(sheet.cells.readValue(row, col))
	return value.kind === 'empty' ? numberValue(0) : value
}

function scalarIfAggregateFallbackKey(aggregate: RangeAggregateOptimization): string {
	return [
		aggregate.functionName,
		aggregate.sheetIndex,
		aggregate.startRow,
		aggregate.startCol,
		aggregate.endRow,
		aggregate.endCol,
	].join(':')
}

function compareScalarIfCondition(
	left: number,
	comparison: ScalarIfAggregateFallbackFormula['comparison'],
	right: number,
): boolean {
	switch (comparison) {
		case '>':
			return left > right
		case '>=':
			return left >= right
		case '<':
			return left < right
		case '<=':
			return left <= right
		case '=':
			return left === right
		case '<>':
			return left !== right
	}
}

function tryFastFullPrefixAggregateTextRecalc(
	workbook: Workbook,
	changed: string[],
	start: number,
	scratch: RecalcScratch,
): RecalcResult | null {
	scratch.textPrefixTailIndex.clear()
	scratch.textPrefixFormulaByGroupEnd.clear()
	scratch.textPrefixGroups.clear()
	scratch.textPrefixAggregateStates.clear()
	scratch.textPrefixSourceSnapshots.clear()
	const groups = new Map<string, TextPrefixAggregateBlock[]>()
	const formulaCells: TextPrefixAggregateBlock[] = []
	const tailIndex = new Map<CellKey, TextPrefixAggregateBlock>()
	const formulaByGroupEnd = new Map<string, TextPrefixAggregateBlock>()
	const aggregateStates = new Map<string, RangeAggregateState>()
	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, entries] of sheet.cells.iterateRows()) {
			for (const [col, cell] of entries) {
				if (cell.formula === null) {
					if (cell.formulaInfo !== undefined) return null
					continue
				}
				if (cell.formulaInfo !== undefined) return null
				const aggregate = parseSimplePrefixAggregateFormula(cell.formula, sheetIndex)
				if (!aggregate) return null
				const groupKey = [
					aggregate.functionName,
					aggregate.sheetIndex,
					aggregate.startRow,
					aggregate.startCol,
					aggregate.endCol,
				].join(':')
				const block = { sheetIndex, row, col, formula: cell.formula, aggregate }
				formulaCells.push(block)
				const groupEndKey = `${prefixAggregateTextGroupKey(aggregate)}:${aggregate.endRow}`
				formulaByGroupEnd.set(groupEndKey, block)
				tailIndex.set(cellKey(aggregate.sheetIndex, aggregate.endRow, aggregate.startCol), block)
				let group = groups.get(groupKey)
				if (!group) {
					group = []
					groups.set(groupKey, group)
				}
				group.push(block)
			}
		}
	}
	if (formulaCells.length === 0 || groups.size === 0) return null

	for (const group of groups.values()) {
		group.sort((a, b) => a.aggregate.endRow - b.aggregate.endRow || a.row - b.row || a.col - b.col)
		const first = group[0]
		const last = group[group.length - 1]
		if (!first || !last) return null
		if (
			textPrefixSourceContainsFormula(
				formulaCells,
				first.aggregate.sheetIndex,
				first.aggregate.startRow,
				first.aggregate.startCol,
				last.aggregate.endRow,
				first.aggregate.endCol,
			)
		) {
			return null
		}

		let state: RangeAggregateState | null = null
		let previousEndRow = first.aggregate.startRow - 1
		for (const block of group) {
			const aggregate = block.aggregate
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
			const sheet = workbook.sheets[block.sheetIndex]
			if (!sheet) return null
			if (needsRangeAggregateState(aggregate.functionName)) {
				aggregateStates.set(`${prefixAggregateTextGroupKey(aggregate)}:${aggregate.endRow}`, state)
			}
			const newValue = rangeAggregateStateToValue(aggregate.functionName, state)
			const oldValue = sheet.cells.readValue(block.row, block.col)
			if (valuesEqual(oldValue, newValue)) continue
			const oldStyleId = sheet.cells.readStyleId(block.row, block.col) ?? DEFAULT_STYLE_ID
			if (newValue.kind === 'number') {
				sheet.cells.setNumberResolved(
					block.row,
					block.col,
					newValue.value,
					block.formula,
					oldStyleId,
				)
			} else {
				sheet.cells.setResolved(block.row, block.col, newValue, block.formula, oldStyleId)
			}
			changed.push(cellRefString(workbook, block.sheetIndex, block.row, block.col))
		}
	}
	for (const [key, block] of tailIndex) scratch.textPrefixTailIndex.set(key, block)
	for (const [key, block] of formulaByGroupEnd) scratch.textPrefixFormulaByGroupEnd.set(key, block)
	for (const [key, group] of groups) scratch.textPrefixGroups.set(key, group)
	for (const [key, state] of aggregateStates) scratch.textPrefixAggregateStates.set(key, state)
	cacheTextPrefixSourceSnapshots(workbook, groups, scratch.textPrefixSourceSnapshots)
	return { changed, errors: [], duration: performance.now() - start }
}

function tryFastDirtyPrefixAggregateTextRecalc(
	workbook: Workbook,
	dirtyRefs: readonly string[] | undefined,
	changed: string[],
	start: number,
	scratch: RecalcScratch,
): RecalcResult | null {
	const source = resolveSingleDirtyCell(workbook, dirtyRefs)
	if (!source) return null
	const indexed = scratch.textPrefixTailIndex.get(
		cellKey(source.sheetIndex, source.row, source.col),
	)
	if (indexed && indexed.aggregate.endRow > indexed.aggregate.startRow) {
		const sheet = workbook.sheets[indexed.sheetIndex]
		const currentFormula = sheet?.cells.readFormula(indexed.row, indexed.col)
		if (
			sheet &&
			currentFormula === indexed.formula &&
			sheet.cells.readFormulaInfo(indexed.row, indexed.col) === undefined
		) {
			const tail = tryFastDirtyTailPrefixValue(
				workbook,
				source,
				indexed,
				scratch.textPrefixFormulaByGroupEnd,
				scratch.textPrefixAggregateStates,
			)
			if (tail && writeTextPrefixAggregateValue(workbook, indexed, tail.value, changed)) {
				cacheTailPrefixAggregateState(indexed, tail, scratch.textPrefixAggregateStates)
				return { changed, errors: [], duration: performance.now() - start }
			}
		}
	}
	const cached = tryFastDirtyCachedPrefixAggregateTextRecalc(
		workbook,
		source,
		changed,
		start,
		scratch,
	)
	if (cached) return cached
	const groups = new Map<string, TextPrefixAggregateBlock[]>()
	const formulaCells: TextPrefixAggregateBlock[] = []
	const formulaByGroupEnd = new Map<string, TextPrefixAggregateBlock>()
	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, entries] of sheet.cells.iterateRows()) {
			for (const [col, cell] of entries) {
				if (cell.formula === null) {
					if (cell.formulaInfo !== undefined) return null
					continue
				}
				if (cell.formulaInfo !== undefined) return null
				const aggregate = parseSimplePrefixAggregateFormula(cell.formula, sheetIndex)
				if (!aggregate) return null
				const groupKey = prefixAggregateTextGroupKey(aggregate)
				const block = { sheetIndex, row, col, formula: cell.formula, aggregate }
				formulaCells.push(block)
				formulaByGroupEnd.set(`${groupKey}:${aggregate.endRow}`, block)
				if (sourceOverlapsAggregate(source, aggregate)) {
					let group = groups.get(groupKey)
					if (!group) {
						group = []
						groups.set(groupKey, group)
					}
					group.push(block)
				}
			}
		}
	}
	if (formulaCells.length === 0 || groups.size === 0) return null
	for (const group of groups.values()) {
		group.sort((a, b) => a.aggregate.endRow - b.aggregate.endRow || a.row - b.row || a.col - b.col)
		const first = group[0]
		const last = group[group.length - 1]
		if (!first || !last) return null
		if (
			textPrefixSourceContainsFormula(
				formulaCells,
				first.aggregate.sheetIndex,
				first.aggregate.startRow,
				first.aggregate.startCol,
				last.aggregate.endRow,
				first.aggregate.endCol,
			)
		) {
			return null
		}
		let state: RangeAggregateState | null = null
		let previousEndRow = first.aggregate.startRow - 1
		for (const block of group) {
			const aggregate = block.aggregate
			const tail = tryFastDirtyTailPrefixValue(workbook, source, block, formulaByGroupEnd)
			let newValue = tail?.value ?? null
			if (newValue?.kind === 'number') {
				if (aggregate.functionName === 'SUM') {
					state = {
						sum: newValue.value,
						count: 0,
						min: Number.POSITIVE_INFINITY,
						max: Number.NEGATIVE_INFINITY,
						error: null,
					}
					previousEndRow = aggregate.endRow
				} else if (aggregate.functionName === 'COUNT') {
					state = {
						sum: 0,
						count: newValue.value,
						min: Number.POSITIVE_INFINITY,
						max: Number.NEGATIVE_INFINITY,
						error: null,
					}
					previousEndRow = aggregate.endRow
				}
			} else if (tail?.state) {
				state = tail.state
				previousEndRow = aggregate.endRow
			}
			if (!newValue) {
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
				newValue = rangeAggregateStateToValue(aggregate.functionName, state)
			}
			if (!writeTextPrefixAggregateValue(workbook, block, newValue, changed)) return null
			cacheTextPrefixAggregateState(block, state, scratch.textPrefixAggregateStates)
		}
	}
	return { changed, errors: [], duration: performance.now() - start }
}

function tryFastDirtyCachedPrefixAggregateTextRecalc(
	workbook: Workbook,
	source: CellCoords,
	changed: string[],
	start: number,
	scratch: RecalcScratch,
): RecalcResult | null {
	if (scratch.textPrefixGroups.size === 0) return null
	const sourceSheet = workbook.sheets[source.sheetIndex]
	if (!sourceSheet || sourceSheet.cells.readFormula(source.row, source.col) !== null) return null
	const affectedGroups: (readonly TextPrefixAggregateBlock[])[] = []
	for (const group of scratch.textPrefixGroups.values()) {
		const first = group[0]
		const last = group[group.length - 1]
		if (!first || !last) continue
		if (!sourceOverlapsAggregate(source, last.aggregate)) continue
		if (
			source.sheetIndex !== first.aggregate.sheetIndex ||
			source.col < first.aggregate.startCol ||
			source.col > first.aggregate.endCol
		) {
			continue
		}
		affectedGroups.push(group)
	}
	if (affectedGroups.length === 0) return null

	for (const group of affectedGroups) {
		if (!textPrefixGroupStillValid(workbook, group)) return null
		const deltaValue = tryTextPrefixSumDeltaValue(
			workbook,
			source,
			group,
			scratch.textPrefixSourceSnapshots,
		)
		if (deltaValue !== null) {
			for (const block of group) {
				if (!sourceOverlapsAggregate(source, block.aggregate)) continue
				const newValue = deltaValue(workbook, block)
				if (!newValue || !writeTextPrefixAggregateValue(workbook, block, newValue, changed)) {
					return null
				}
			}
			continue
		}

		let state: RangeAggregateState | null = null
		let previousEndRow = (group[0]?.aggregate.startRow ?? 0) - 1
		for (const block of group) {
			if (!sourceOverlapsAggregate(source, block.aggregate)) continue
			const aggregate = block.aggregate
			const tail = tryFastDirtyTailPrefixValue(
				workbook,
				source,
				block,
				scratch.textPrefixFormulaByGroupEnd,
				scratch.textPrefixAggregateStates,
			)
			let newValue = tail?.value ?? null
			if (newValue?.kind === 'number') {
				if (aggregate.functionName === 'SUM') {
					state = {
						sum: newValue.value,
						count: 0,
						min: Number.POSITIVE_INFINITY,
						max: Number.NEGATIVE_INFINITY,
						error: null,
					}
					previousEndRow = aggregate.endRow
				} else if (aggregate.functionName === 'COUNT') {
					state = {
						sum: 0,
						count: newValue.value,
						min: Number.POSITIVE_INFINITY,
						max: Number.NEGATIVE_INFINITY,
						error: null,
					}
					previousEndRow = aggregate.endRow
				}
			} else if (tail?.state) {
				state = tail.state
				previousEndRow = aggregate.endRow
			}
			if (!newValue) {
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
				newValue = rangeAggregateStateToValue(aggregate.functionName, state)
			}
			if (!writeTextPrefixAggregateValue(workbook, block, newValue, changed)) return null
			cacheTextPrefixAggregateState(block, state, scratch.textPrefixAggregateStates)
		}
	}
	scratch.textPrefixSourceSnapshots.set(
		cellKey(source.sheetIndex, source.row, source.col),
		readTextPrefixSourceSnapshot(workbook, source.sheetIndex, source.row, source.col),
	)
	return { changed, errors: [], duration: performance.now() - start }
}

function textPrefixGroupStillValid(
	workbook: Workbook,
	group: readonly TextPrefixAggregateBlock[],
): boolean {
	for (const block of group) {
		const sheet = workbook.sheets[block.sheetIndex]
		if (!sheet) return false
		if (sheet.cells.readFormula(block.row, block.col) !== block.formula) return false
		if (sheet.cells.readFormulaInfo(block.row, block.col) !== undefined) return false
	}
	return true
}

function tryTextPrefixSumDeltaValue(
	workbook: Workbook,
	source: CellCoords,
	group: readonly TextPrefixAggregateBlock[],
	sourceSnapshots: Map<CellKey, TextPrefixSourceSnapshot>,
): ((workbook: Workbook, block: TextPrefixAggregateBlock) => CellValue | null) | null {
	const first = group[0]
	if (!first || first.aggregate.functionName !== 'SUM') return null
	const sourceKey = cellKey(source.sheetIndex, source.row, source.col)
	const previous = sourceSnapshots.get(sourceKey)
	if (previous === undefined || previous === 'error') return null
	const current = readTextPrefixSourceSnapshot(workbook, source.sheetIndex, source.row, source.col)
	if (current === 'error') return null
	const delta = (current ?? 0) - (previous ?? 0)
	sourceSnapshots.set(sourceKey, current)
	return (wb, block) => {
		const sheet = wb.sheets[block.sheetIndex]
		if (!sheet) return null
		const oldNumber = sheet.cells.readNumber(block.row, block.col)
		if (oldNumber === null || oldNumber === undefined) return null
		return numberValue(oldNumber + delta)
	}
}

function cacheTextPrefixSourceSnapshots(
	workbook: Workbook,
	groups: ReadonlyMap<string, readonly TextPrefixAggregateBlock[]>,
	sourceSnapshots: Map<CellKey, TextPrefixSourceSnapshot>,
): void {
	for (const group of groups.values()) {
		const first = group[0]
		const last = group[group.length - 1]
		if (!first || !last || first.aggregate.functionName !== 'SUM') continue
		const aggregate = first.aggregate
		const sheet = workbook.sheets[aggregate.sheetIndex]
		if (!sheet) continue
		for (let row = aggregate.startRow; row <= last.aggregate.endRow; row++) {
			for (let col = aggregate.startCol; col <= aggregate.endCol; col++) {
				sourceSnapshots.set(
					cellKey(aggregate.sheetIndex, row, col),
					readTextPrefixSourceSnapshot(workbook, aggregate.sheetIndex, row, col),
				)
			}
		}
	}
}

function readTextPrefixSourceSnapshot(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): TextPrefixSourceSnapshot {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return null
	const value = readRangeAggregateNumericCell(sheet, row, col)
	if (typeof value === 'number') return value
	if (value?.kind === 'error') return 'error'
	return null
}

function parseSimplePrefixAggregateFormula(
	formula: string,
	sheetIndex: number,
): RangeAggregateOptimization | null {
	const text = formula.trim().replace(/^=/, '')
	const open = text.indexOf('(')
	if (open <= 0 || text.charCodeAt(text.length - 1) !== 41) return null
	const functionName = parseGrowingAggregateFunction(text.slice(0, open))
	if (!functionName) return null
	const colon = text.indexOf(':', open + 1)
	if (colon < 0) return null
	const startCell = parseSimpleCellRef(text, open + 1, colon)
	const endCell = parseSimpleCellRef(text, colon + 1, text.length - 1)
	if (!startCell || !endCell) return null
	if (startCell.col !== endCell.col || endCell.row < startCell.row) return null
	return {
		functionName,
		sheetIndex,
		startRow: startCell.row,
		startCol: startCell.col,
		endRow: endCell.row,
		endCol: endCell.col,
	}
}

function parseGrowingAggregateFunction(name: string): GrowingRangeAggregateFunction | null {
	switch (name.trim().toUpperCase()) {
		case 'SUM':
			return 'SUM'
		case 'COUNT':
			return 'COUNT'
		case 'AVERAGE':
			return 'AVERAGE'
		case 'MIN':
			return 'MIN'
		case 'MAX':
			return 'MAX'
		default:
			return null
	}
}

function parseSimpleCellRef(
	text: string,
	rawStart: number,
	rawEnd: number,
): { row: number; col: number } | null {
	let i = rawStart
	let end = rawEnd
	while (i < end && text.charCodeAt(i) <= 32) i++
	while (end > i && text.charCodeAt(end - 1) <= 32) end--
	if (text.charCodeAt(i) === 36) i++
	let col = 0
	let sawCol = false
	while (i < end) {
		const code = text.charCodeAt(i)
		const upper = code >= 97 && code <= 122 ? code - 32 : code
		if (upper < 65 || upper > 90) break
		col = col * 26 + (upper - 64)
		sawCol = true
		i++
	}
	if (!sawCol) return null
	if (text.charCodeAt(i) === 36) i++
	let row = 0
	let sawRow = false
	while (i < end) {
		const code = text.charCodeAt(i)
		if (code < 48 || code > 57) return null
		row = row * 10 + (code - 48)
		sawRow = true
		i++
	}
	if (!sawRow || row < 1 || col < 1 || col > 16_384) return null
	return { row: row - 1, col: col - 1 }
}

function parseIndexMatchReturnAst(
	ast: FormulaNode,
	sheetIndex: number,
): IndexMatchReturnPattern | null {
	if (ast.type !== 'function' || ast.name.toUpperCase() !== 'INDEX' || ast.args.length !== 2) {
		return null
	}
	const returnRange = ast.args[0]
	const rowArg = ast.args[1]
	if (!returnRange || returnRange.type !== 'rangeRef' || returnRange.sheet !== undefined) {
		return null
	}
	if (
		returnRange.start.col !== returnRange.end.col ||
		returnRange.end.row < returnRange.start.row
	) {
		return null
	}
	if (!rowArg || rowArg.type !== 'function' || rowArg.name.toUpperCase() !== 'MATCH') return null
	if (rowArg.args.length !== 3) return null
	const lookupCell = rowArg.args[0]
	const lookupRange = rowArg.args[1]
	const matchType = rowArg.args[2]
	if (!lookupCell || lookupCell.type !== 'cellRef' || lookupCell.sheet !== undefined) return null
	if (!lookupRange || lookupRange.type !== 'rangeRef' || lookupRange.sheet !== undefined)
		return null
	if (matchType?.type !== 'number' || matchType.value !== 0) return null
	if (
		lookupRange.start.col !== lookupRange.end.col ||
		lookupRange.end.row < lookupRange.start.row
	) {
		return null
	}
	if (returnRange.end.row - returnRange.start.row !== lookupRange.end.row - lookupRange.start.row) {
		return null
	}
	return {
		sheetIndex,
		returnStartRow: returnRange.start.row,
		returnCol: returnRange.start.col,
		returnEndRow: returnRange.end.row,
		lookupCellRow: lookupCell.ref.row,
		lookupCellCol: lookupCell.ref.col,
		lookupStartRow: lookupRange.start.row,
		lookupCol: lookupRange.start.col,
		lookupEndRow: lookupRange.end.row,
	}
}

function prefixAggregateTextGroupKey(aggregate: RangeAggregateOptimization): string {
	return [
		aggregate.functionName,
		aggregate.sheetIndex,
		aggregate.startRow,
		aggregate.startCol,
		aggregate.endCol,
	].join(':')
}

function sourceOverlapsAggregate(
	source: CellCoords,
	aggregate: RangeAggregateOptimization,
): boolean {
	return (
		source.sheetIndex === aggregate.sheetIndex &&
		source.row >= aggregate.startRow &&
		source.row <= aggregate.endRow &&
		source.col >= aggregate.startCol &&
		source.col <= aggregate.endCol
	)
}

function resolveSingleDirtyCell(
	workbook: Workbook,
	dirtyRefs: readonly string[] | undefined,
): CellCoords | null {
	if (!dirtyRefs || dirtyRefs.length !== 1) return null
	const ref = dirtyRefs[0]
	if (!ref) return null
	const bang = ref.lastIndexOf('!')
	const sheetName = bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : workbook.sheets[0]?.name
	const localRef = bang >= 0 ? ref.slice(bang + 1) : ref
	if (!sheetName || !localRef) return null
	const sheetIndex = workbook.sheets.findIndex(
		(sheet) => sheet?.name.toLowerCase() === sheetName.toLowerCase(),
	)
	if (sheetIndex < 0) return null
	const range = parseRange(localRef)
	if (range.start.row !== range.end.row || range.start.col !== range.end.col) return null
	return { sheetIndex, row: range.start.row, col: range.start.col }
}

function tryFastDirtyTailPrefixValue(
	workbook: Workbook,
	source: CellCoords,
	block: TextPrefixAggregateBlock,
	formulaByGroupEnd: ReadonlyMap<string, TextPrefixAggregateBlock>,
	stateByGroupEnd?: ReadonlyMap<string, RangeAggregateState>,
): TailPrefixAggregateResult | null {
	const aggregate = block.aggregate
	if (aggregate.endRow !== source.row || aggregate.endRow <= aggregate.startRow) return null
	const groupKey = prefixAggregateTextGroupKey(aggregate)
	const previousEndKey = `${groupKey}:${aggregate.endRow - 1}`
	const previous = formulaByGroupEnd.get(previousEndKey)
	if (!previous) return null
	if (needsRangeAggregateState(aggregate.functionName)) {
		const previousState = stateByGroupEnd?.get(previousEndKey)
		const sourceSheet = workbook.sheets[source.sheetIndex]
		if (!previousState || !sourceSheet) return null
		const state = appendCellToRangeAggregateState(
			aggregate.functionName,
			previousState,
			readRangeAggregateNumericCell(sourceSheet, source.row, source.col),
		)
		return { value: rangeAggregateStateToValue(aggregate.functionName, state), state }
	}
	const optimization = {
		functionName: aggregate.functionName,
		previousKey: cellKey(previous.sheetIndex, previous.row, previous.col),
		previousSheetIndex: previous.sheetIndex,
		previousRow: previous.row,
		previousCol: previous.col,
		appendSheetIndex: source.sheetIndex,
		appendStartRow: source.row,
		appendStartCol: source.col,
		appendEndRow: source.row,
		appendEndCol: source.col,
	}
	const scalarValue = tryEvaluateGrowingRangeScalarAggregate(workbook, optimization, true)
	if (scalarValue) return { value: scalarValue }
	return null
}

function cacheTailPrefixAggregateState(
	block: TextPrefixAggregateBlock,
	tail: TailPrefixAggregateResult,
	stateByGroupEnd: Map<string, RangeAggregateState>,
): void {
	if (tail.state) {
		stateByGroupEnd.set(
			`${prefixAggregateTextGroupKey(block.aggregate)}:${block.aggregate.endRow}`,
			tail.state,
		)
	}
}

function cacheTextPrefixAggregateState(
	block: TextPrefixAggregateBlock,
	state: RangeAggregateState | null,
	stateByGroupEnd: Map<string, RangeAggregateState>,
): void {
	if (state && needsRangeAggregateState(block.aggregate.functionName)) {
		stateByGroupEnd.set(
			`${prefixAggregateTextGroupKey(block.aggregate)}:${block.aggregate.endRow}`,
			state,
		)
	}
}

function appendCellToRangeAggregateState(
	functionName: GrowingRangeAggregateFunction,
	previous: RangeAggregateState,
	rangeValue: number | CellValue | null,
): RangeAggregateState {
	if (previous.error) return previous
	if (typeof rangeValue !== 'number') {
		return rangeValue?.kind === 'error' ? { ...previous, error: rangeValue } : previous
	}
	switch (functionName) {
		case 'AVERAGE':
			return {
				...previous,
				sum: previous.sum + rangeValue,
				count: previous.count + 1,
			}
		case 'MIN':
			return {
				...previous,
				count: previous.count + 1,
				min: Math.min(previous.min, rangeValue),
			}
		case 'MAX':
			return {
				...previous,
				count: previous.count + 1,
				max: Math.max(previous.max, rangeValue),
			}
		default:
			return previous
	}
}

function tryFastDirtyIndexMatchReturnTextRecalc(
	workbook: Workbook,
	dirtyRefs: readonly string[] | undefined,
	changed: string[],
	start: number,
	scratch: RecalcScratch,
): RecalcResult | null {
	const source = resolveSingleDirtyCell(workbook, dirtyRefs)
	if (!source) return null
	const sourceSheet = workbook.sheets[source.sheetIndex]
	if (!sourceSheet || sourceSheet.cells.readFormula(source.row, source.col) !== null) return null
	const sourceKey = cellKey(source.sheetIndex, source.row, source.col)
	const cached = scratch.indexMatchReturnBySource.get(sourceKey)
	if (cached && cached.length > 0) {
		if (countPlainFormulas(workbook) !== scratch.indexMatchReturnFormulaCount) return null
		for (const formula of cached) {
			if (sourceOverlapsIndexMatchLookup(source, formula.pattern)) {
				clearIndexMatchReturnCache(scratch)
				return null
			}
			const sheet = workbook.sheets[formula.sheetIndex]
			if (!sheet) return null
			if (sheet.cells.readFormula(formula.row, formula.col) !== formula.formula) return null
			if (sheet.cells.readFormulaInfo(formula.row, formula.col) !== undefined) return null
			if (
				!writeFormulaValue(
					workbook,
					formula.sheetIndex,
					formula.row,
					formula.col,
					formula.formula,
					sourceSheet.cells.readValue(source.row, source.col),
					changed,
				)
			) {
				return null
			}
		}
		return { changed, errors: [], duration: performance.now() - start }
	}
	if (scratch.indexMatchReturnFormulaCount > 0) clearIndexMatchReturnCache(scratch)
	return null
}

function firstExactLookupOffset(
	sheet: NonNullable<Workbook['sheets'][number]>,
	startRow: number,
	col: number,
	endRow: number,
	lookupValue: CellValue,
): number {
	for (let row = startRow; row <= endRow; row++) {
		if (valuesEqual(sheet.cells.readValue(row, col), lookupValue)) return row - startRow
	}
	return -1
}

function cacheIndexMatchReturnFormulas(
	workbook: Workbook,
	scratch: RecalcScratch,
	formulas?: ReadonlyMap<CellKey, AnalyzedFormula>,
): void {
	clearIndexMatchReturnCache(scratch)
	const firstOffsetIndexes = new Map<string, ReadonlyMap<string, number>>()
	if (formulas) {
		const formulaCount = countPlainFormulas(workbook)
		if (formulaCount < 1 || formulaCount !== formulas.size) return
		scratch.indexMatchReturnFormulaCount = formulaCount
		for (const formula of formulas.values()) {
			if (formula.parseError || !formula.ast) {
				scratch.indexMatchReturnBySource.clear()
				scratch.indexMatchReturnFormulaCount = 0
				return
			}
			const sheet = workbook.sheets[formula.sheetIndex]
			if (!sheet || sheet.cells.readFormulaInfo(formula.row, formula.col) !== undefined) {
				scratch.indexMatchReturnBySource.clear()
				scratch.indexMatchReturnFormulaCount = 0
				return
			}
			const pattern = parseIndexMatchReturnAst(formula.ast, formula.sheetIndex)
			if (!pattern) {
				scratch.indexMatchReturnBySource.clear()
				scratch.indexMatchReturnFormulaCount = 0
				return
			}
			cacheIndexMatchReturnFormula(workbook, scratch, firstOffsetIndexes, {
				sheetIndex: formula.sheetIndex,
				row: formula.row,
				col: formula.col,
				formula: formula.formula,
				pattern,
			})
		}
		return
	}
}

function sourceOverlapsIndexMatchLookup(
	source: CellCoords,
	pattern: IndexMatchReturnPattern,
): boolean {
	return (
		source.sheetIndex === pattern.sheetIndex &&
		source.col === pattern.lookupCol &&
		source.row >= pattern.lookupStartRow &&
		source.row <= pattern.lookupEndRow
	)
}

function clearIndexMatchReturnCache(scratch: RecalcScratch): void {
	scratch.indexMatchReturnBySource.clear()
	scratch.indexMatchReturnFormulaCount = 0
}

function cacheIndexMatchReturnFormula(
	workbook: Workbook,
	scratch: RecalcScratch,
	firstOffsetIndexes: Map<string, ReadonlyMap<string, number>>,
	formula: IndexMatchReturnFormula,
): void {
	const sheet = workbook.sheets[formula.pattern.sheetIndex]
	if (!sheet) return
	const pattern = formula.pattern
	const lookupValue = sheet.cells.readValue(pattern.lookupCellRow, pattern.lookupCellCol)
	const offset =
		firstExactLookupOffsetFromIndex(sheet, pattern, lookupValue, firstOffsetIndexes) ??
		firstExactLookupOffset(
			sheet,
			pattern.lookupStartRow,
			pattern.lookupCol,
			pattern.lookupEndRow,
			lookupValue,
		)
	if (offset < 0) return
	const sourceRow = pattern.returnStartRow + offset
	if (sourceRow > pattern.returnEndRow) return
	const sourceKey = cellKey(pattern.sheetIndex, sourceRow, pattern.returnCol)
	const existing = scratch.indexMatchReturnBySource.get(sourceKey)
	if (existing) {
		scratch.indexMatchReturnBySource.set(sourceKey, [...existing, formula])
	} else {
		scratch.indexMatchReturnBySource.set(sourceKey, [formula])
	}
}

function firstExactLookupOffsetFromIndex(
	sheet: NonNullable<Workbook['sheets'][number]>,
	pattern: IndexMatchReturnPattern,
	lookupValue: CellValue,
	firstOffsetIndexes: Map<string, ReadonlyMap<string, number>>,
): number | null {
	const valueKey = exactCellValueKey(lookupValue)
	if (valueKey === null) return null
	const rangeKey = [
		pattern.sheetIndex,
		pattern.lookupStartRow,
		pattern.lookupCol,
		pattern.lookupEndRow,
	].join(':')
	let index = firstOffsetIndexes.get(rangeKey)
	if (!index) {
		const next = new Map<string, number>()
		for (let row = pattern.lookupStartRow; row <= pattern.lookupEndRow; row++) {
			const key = exactCellValueKey(sheet.cells.readValue(row, pattern.lookupCol))
			if (key !== null && !next.has(key)) next.set(key, row - pattern.lookupStartRow)
		}
		index = next
		firstOffsetIndexes.set(rangeKey, index)
	}
	return index.get(valueKey) ?? -1
}

function exactCellValueKey(value: CellValue): string | null {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'empty':
			return 'empty:'
		case 'number':
			return `number:${scalar.value}`
		case 'date':
			return `date:${scalar.serial}`
		case 'string':
			return `string:${scalar.value}`
		case 'boolean':
			return `boolean:${scalar.value ? 1 : 0}`
		case 'error':
			return `error:${scalar.value}`
		default:
			return null
	}
}

function countPlainFormulas(workbook: Workbook): number {
	let count = 0
	for (const sheet of workbook.sheets) {
		if (!sheet) continue
		if (sheet.cells.formulaInfoCellCount() > 0) return -1
		count += sheet.cells.formulaCellCount()
	}
	return count
}

function writeFormulaValue(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	formula: string,
	newValue: CellValue,
	changed: string[],
): boolean {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return false
	const oldValue = sheet.cells.readValue(row, col)
	if (valuesEqual(oldValue, newValue)) return true
	const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
	if (newValue.kind === 'number') {
		sheet.cells.setNumberResolved(row, col, newValue.value, formula, oldStyleId)
	} else {
		sheet.cells.setResolved(row, col, newValue, formula, oldStyleId)
	}
	changed.push(cellRefString(workbook, sheetIndex, row, col))
	return true
}

function writeTextPrefixAggregateValue(
	workbook: Workbook,
	block: TextPrefixAggregateBlock,
	newValue: CellValue,
	changed: string[],
): boolean {
	const sheet = workbook.sheets[block.sheetIndex]
	if (!sheet) return false
	const oldValue = sheet.cells.readValue(block.row, block.col)
	if (valuesEqual(oldValue, newValue)) return true
	const oldStyleId = sheet.cells.readStyleId(block.row, block.col) ?? DEFAULT_STYLE_ID
	if (newValue.kind === 'number') {
		sheet.cells.setNumberResolved(block.row, block.col, newValue.value, block.formula, oldStyleId)
	} else {
		sheet.cells.setResolved(block.row, block.col, newValue, block.formula, oldStyleId)
	}
	changed.push(cellRefString(workbook, block.sheetIndex, block.row, block.col))
	return true
}

function textPrefixSourceContainsFormula(
	formulas: readonly TextPrefixAggregateBlock[],
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): boolean {
	for (const formula of formulas) {
		if (formula.sheetIndex !== sheetIndex) continue
		if (
			formula.row >= startRow &&
			formula.row <= endRow &&
			formula.col >= startCol &&
			formula.col <= endCol
		) {
			return true
		}
	}
	return false
}

function tryFastFullPrefixAggregateBlocks(
	workbook: Workbook,
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
	evalOrderSet: ReadonlySet<CellKey>,
	rangeAggregateStates: Map<CellKey, RangeAggregateState> | null,
	changed: string[],
): Set<CellKey> | null {
	const groups = new Map<string, PrefixAggregateBlock[]>()
	for (const [key, formula] of formulas) {
		if (!evalOrderSet.has(key) || formula.parseError) continue
		const aggregate = formula.rangeAggregate
		if (!aggregate || !canFastFullPrefixAggregate(aggregate)) continue
		const groupKey = [
			aggregate.functionName,
			aggregate.sheetIndex,
			aggregate.startRow,
			aggregate.startCol,
			aggregate.endCol,
		].join(':')
		let group = groups.get(groupKey)
		if (!group) {
			group = []
			groups.set(groupKey, group)
		}
		group.push({ key, formula, aggregate })
	}
	if (groups.size === 0) return null

	const completed = new Set<CellKey>()
	const hasFormulaInfo = workbookHasFormulaInfo(workbook)
	for (const group of groups.values()) {
		if (group.length < 4) continue
		group.sort((a, b) => a.aggregate.endRow - b.aggregate.endRow || a.key - b.key)
		const first = group[0]
		const last = group[group.length - 1]
		if (!first || !last) continue
		const sourceEndRow = last.aggregate.endRow
		if (
			rangeContainsFormula(
				formulas,
				first.aggregate.sheetIndex,
				first.aggregate.startRow,
				first.aggregate.startCol,
				sourceEndRow,
				first.aggregate.endCol,
			)
		) {
			continue
		}

		let state: RangeAggregateState | null = null
		let previousEndRow = first.aggregate.startRow - 1
		for (const block of group) {
			const { formula, aggregate } = block
			if (aggregate.endRow < previousEndRow) continue
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
				if (!state) break
				previousEndRow = aggregate.endRow
			}
			if (!state) continue
			const sheet = workbook.sheets[formula.sheetIndex]
			if (!sheet) continue
			if (hasFormulaInfo && sheet.cells.readFormulaInfo(formula.row, formula.col) !== undefined) {
				continue
			}
			const oldFormula = sheet.cells.readFormula(formula.row, formula.col) ?? null
			if (oldFormula === null) continue

			const newValue = rangeAggregateStateToValue(aggregate.functionName, state)
			const oldValue = sheet.cells.readValue(formula.row, formula.col)
			if (!valuesEqual(oldValue, newValue)) {
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
			if (rangeAggregateStates && needsRangeAggregateState(aggregate.functionName)) {
				rangeAggregateStates.set(block.key, state)
			}
			completed.add(block.key)
		}
	}
	return completed.size > 0 ? completed : null
}

function canFastFullPrefixAggregate(aggregate: RangeAggregateOptimization): boolean {
	return (
		aggregate.startCol === aggregate.endCol &&
		aggregate.endRow >= aggregate.startRow &&
		(aggregate.functionName === 'SUM' ||
			aggregate.functionName === 'COUNT' ||
			aggregate.functionName === 'AVERAGE' ||
			aggregate.functionName === 'MIN' ||
			aggregate.functionName === 'MAX')
	)
}

function rangeContainsFormula(
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): boolean {
	for (const formula of formulas.values()) {
		if (formula.sheetIndex !== sheetIndex) continue
		if (
			formula.row >= startRow &&
			formula.row <= endRow &&
			formula.col >= startCol &&
			formula.col <= endCol
		) {
			return true
		}
	}
	return false
}

function tryFastDirtySingleFormulaRecalc(
	workbook: Workbook,
	graph: DependencyGraph,
	formulas: ReadonlyMap<CellKey, AnalyzedFormula>,
	dirtyRefKeys: readonly CellKey[],
	mutableCtx: MutableEvalContext,
	start: number,
): RecalcResult | null {
	if (dirtyRefKeys.length !== 1) return null
	const sourceKey = dirtyRefKeys[0] as CellKey
	const dependents = graph.getDependents(sourceKey)
	if (dependents.length !== 1) return null
	const formulaKey = dependents[0] as CellKey
	if (graph.getDependents(formulaKey).length > 0) return null
	const analyzed = formulas.get(formulaKey)
	if (!analyzed || analyzed.parseError || !analyzed.ast) return null
	if (analyzed.deps.includes(formulaKey)) return null
	if (hasExternalWorkbookReference(analyzed.ast)) return null
	const sheet = workbook.sheets[analyzed.sheetIndex]
	if (!sheet) return null
	const oldFormulaInfo = sheet.cells.readFormulaInfo(analyzed.row, analyzed.col)
	if (oldFormulaInfo !== undefined) return null
	const oldFormula = sheet.cells.readFormula(analyzed.row, analyzed.col) ?? null
	if (oldFormula === null) return null
	const oldValue = sheet.cells.readValue(analyzed.row, analyzed.col)
	const oldStyleId = sheet.cells.readStyleId(analyzed.row, analyzed.col) ?? DEFAULT_STYLE_ID
	mutableCtx.sheetIndex = analyzed.sheetIndex
	mutableCtx.row = analyzed.row
	mutableCtx.col = analyzed.col
	const newValue = evalFormula(formulaKey, analyzed.formula, analyzed.ast, mutableCtx)
	if (toScalarMatrix(newValue)) return null
	const changed: string[] = []
	if (!valuesEqual(oldValue, newValue)) {
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
		const bang = ref.lastIndexOf('!')
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
		const bang = ref.lastIndexOf('!')
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
		const bang = ref.lastIndexOf('!')
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
			const newValue = normalizeFormulaResultForExistingCell(
				sheet.cells.readValue(row, col),
				evalFormula(key, formulaText, ast, mutableCtx),
			)
			const hadCell = sheet.cells.has(row, col)
			const oldValue = sheet.cells.readValue(row, col)
			const oldFormula = sheet.cells.readFormula(row, col) ?? null
			const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
			const oldFormulaInfo = sheet.cells.readFormulaInfo(row, col)
			const spillMatrix = toScalarMatrix(newValue)
			if (spillMatrix) {
				const changedBefore = changed.length
				applyArrayOrLegacyResult(
					workbook,
					si,
					row,
					col,
					oldFormula,
					oldStyleId,
					oldFormulaInfo,
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
		const newValue = normalizeFormulaResultForExistingCell(
			sheet.cells.readValue(row, col),
			evalFormula(key, formulaText, ast, mutableCtx),
		)
		const hadCell = sheet.cells.has(row, col)
		const oldValue = sheet.cells.readValue(row, col)
		const oldFormula = sheet.cells.readFormula(row, col) ?? null
		const oldStyleId = sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID
		const oldFormulaInfo = sheet.cells.readFormulaInfo(row, col)
		const spillMatrix = toScalarMatrix(newValue)
		if (spillMatrix) {
			applyArrayOrLegacyResult(
				workbook,
				si,
				row,
				col,
				oldFormula,
				oldStyleId,
				oldFormulaInfo,
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
