import {
	type CellStyle,
	type Color,
	type CustomFilter,
	DEFAULT_STYLE_ID,
	type FilterColumn,
	type FilterDateGroupItem,
	indexToColumn,
	parseRange,
	pivotDataFieldCaptionsMatch,
	type Sheet,
	type Workbook,
} from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import {
	type AggregateRangeCache,
	dateToSerial,
	type EvalArea,
	type EvalArg,
	type ExactLookupCache,
	type FunctionDef,
	type FunctionEvalContext,
	functionRegistry,
	getRange,
	type LookupVectorCache,
	type NumericVectorCache,
	serialToDate,
	cachedParseFormula as sharedCachedParseFormula,
	toNumber,
} from '@ascend/formulas'
import type { CellValue, ScalarCellValue } from '@ascend/schema'
import {
	arrayValue,
	assertUnreachable,
	booleanValue,
	coerceCellValueToString,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
	topLeftScalar,
} from '@ascend/schema'
import type { CalcContext } from './calc-context.ts'
import { aggregateNumericRange } from './compiled-eval.ts'
import { evaluateConditionalFormats } from './conditional-format.ts'
import { computeIconFilterRows } from './icon-filter.ts'
import { resolveSheetIndexInWorkbook as resolveSheetIndex } from './sheet-index.ts'
import { createStructuredRefResolver, type StructuredRefResolver } from './structured-refs.ts'

export interface EvalContext {
	readonly workbook: Workbook
	readonly calcContext: CalcContext
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly definedNameStack?: readonly string[]
	readonly letBindings?: ReadonlyMap<string, LetBinding>
	readonly exactLookupCache?: ExactLookupCache
	readonly lookupVectorCache?: LookupVectorCache
	readonly aggregateRangeCache?: AggregateRangeCache
	readonly numericVectorCache?: NumericVectorCache
	readonly structuredRefResolver?: StructuredRefResolver
}

export class MutableEvalContext implements EvalContext {
	workbook!: Workbook
	calcContext!: CalcContext
	sheetIndex = 0
	row = 0
	col = 0
	exactLookupCache?: ExactLookupCache
	lookupVectorCache?: LookupVectorCache
	aggregateRangeCache?: AggregateRangeCache
	numericVectorCache?: NumericVectorCache
	structuredRefResolver?: StructuredRefResolver
}

class FunctionEvalCtx implements FunctionEvalContext {
	now!: Date
	today!: Date
	randomSeed!: number
	locale!: string
	dateSystem!: '1900' | '1904'
	sheetIndex?: number
	row?: number
	col?: number
	exactLookupCache: ExactLookupCache | undefined
	lookupVectorCache: LookupVectorCache | undefined
	aggregateRangeCache: AggregateRangeCache | undefined
	numericVectorCache: NumericVectorCache | undefined

	update(ctx: EvalContext): this {
		const cc = ctx.calcContext
		this.now = cc.now
		this.today = cc.today
		this.randomSeed = cc.randomSeed
		this.locale = cc.locale
		this.dateSystem = cc.dateSystem
		this.sheetIndex = ctx.sheetIndex
		this.row = ctx.row
		this.col = ctx.col
		this.exactLookupCache = ctx.exactLookupCache
		this.lookupVectorCache = ctx.lookupVectorCache
		this.aggregateRangeCache = ctx.aggregateRangeCache
		this.numericVectorCache = ctx.numericVectorCache
		return this
	}
}

const sharedFnCtx = new FunctionEvalCtx()

const cachedParseFormula = sharedCachedParseFormula

export { clearGlobalParseCache as clearFormulaParseCache } from '@ascend/formulas'
export { invalidateSheetIndexCache } from './sheet-index.ts'

const SCALAR_IMPLICIT_INTERSECTION_FUNCTIONS = new Set([
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
	'CEILING',
	'CHAR',
	'CLEAN',
	'CODE',
	'COMBIN',
	'COMBINA',
	'COS',
	'COSH',
	'COT',
	'COTH',
	'CSC',
	'CSCH',
	'DEGREES',
	'DOLLAR',
	'EVEN',
	'EXACT',
	'EXP',
	'FACT',
	'FACTDOUBLE',
	'FIND',
	'FIXED',
	'FLOOR',
	'FLOOR.MATH',
	'FLOOR.PRECISE',
	'INT',
	'LEFT',
	'LEN',
	'LN',
	'LOG',
	'LOG10',
	'LOWER',
	'MID',
	'MOD',
	'ODD',
	'PROPER',
	'POWER',
	'RADIANS',
	'REPLACE',
	'REPT',
	'RIGHT',
	'ROUND',
	'ROUNDDOWN',
	'ROUNDUP',
	'SEARCH',
	'SEC',
	'SECH',
	'SIGN',
	'SIN',
	'SINH',
	'SQRT',
	'SUBSTITUTE',
	'T',
	'TAN',
	'TANH',
	'TEXT',
	'TEXTAFTER',
	'TEXTBEFORE',
	'TEXTSPLIT',
	'TRIM',
	'TRUNC',
	'UNICHAR',
	'UNICODE',
	'UPPER',
	'VALUE',
])
SCALAR_IMPLICIT_INTERSECTION_FUNCTIONS.delete('T')

const ARRAY_CONTEXT_MAPPABLE_FUNCTIONS = new Set(['SQRT'])

const LEGACY_TOP_LEVEL_SCALAR_FUNCTIONS = new Set([
	...SCALAR_IMPLICIT_INTERSECTION_FUNCTIONS,
	'AND',
	'CONCATENATE',
	'ISBLANK',
	'ISERR',
	'ISERROR',
	'ISEVEN',
	'ISLOGICAL',
	'ISNA',
	'ISNONTEXT',
	'ISNUMBER',
	'ISODD',
	'ISREF',
	'ISTEXT',
	'OR',
])

const EXCEL_MAX_ROWS = 1_048_576
const EXCEL_MAX_COLS = 16_384

interface LambdaInfo {
	readonly params: readonly string[]
	readonly body: FormulaNode
	readonly ctx: EvalContext
}

type LetBinding = CellValue | LambdaInfo

function isLambdaBinding(binding: LetBinding): binding is LambdaInfo {
	return typeof binding === 'object' && binding !== null && 'params' in binding && 'body' in binding
}

type RangeValueCache = Map<number, readonly (readonly CellValue[])[]>
let activeRangeValueCache: RangeValueCache | null = null

const RANGE_KEY_SHIFT_SI = 40
const RANGE_KEY_SHIFT_R1 = 30
const RANGE_KEY_SHIFT_C1 = 20
const RANGE_KEY_SHIFT_R2 = 10

function rangeCacheKey(si: number, r1: number, c1: number, r2: number, c2: number): number {
	return (
		si * 2 ** RANGE_KEY_SHIFT_SI +
		r1 * 2 ** RANGE_KEY_SHIFT_R1 +
		c1 * 2 ** RANGE_KEY_SHIFT_C1 +
		r2 * 2 ** RANGE_KEY_SHIFT_R2 +
		c2
	)
}

export function setRangeValueCache(cache: RangeValueCache): void {
	activeRangeValueCache = cache
}

export function clearRangeValueCache(): void {
	activeRangeValueCache = null
}

function getCellValue(wb: Workbook, sheetIndex: number, row: number, col: number): CellValue {
	if (!isCellInBounds(row, col)) return errorValue('#REF!')
	const sheet = wb.sheets[sheetIndex]
	if (!sheet) return errorValue('#REF!')
	return sheet.cells.readValue(row, col)
}

function externalReferenceTarget(
	sheet: string | undefined,
): { workbook: string; sheet: string } | null {
	if (!sheet?.startsWith('[')) return null
	const close = sheet.indexOf(']')
	if (close <= 1) return null
	const workbook = sheet.slice(1, close)
	const sheetName = sheet.slice(close + 1)
	if (sheetName.length === 0) return null
	return { workbook, sheet: sheetName }
}

function resolveExternalCell(
	ctx: EvalContext,
	sheet: string | undefined,
	row: number,
	col: number,
): CellValue | null {
	const target = externalReferenceTarget(sheet)
	if (!target) return null
	return (
		ctx.calcContext.externalReferences?.resolveCell?.({
			workbook: target.workbook,
			sheet: target.sheet,
			row,
			col,
		}) ?? errorValue('#REF!')
	)
}

function materializeExternalRange(
	ctx: EvalContext,
	workbook: string,
	sheet: string,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): readonly (readonly CellValue[])[] | null {
	const resolver = ctx.calcContext.externalReferences
	if (resolver?.resolveRange) {
		const resolved = resolver.resolveRange({
			workbook,
			sheet,
			row: startRow,
			col: startCol,
			endRow,
			endCol,
		})
		if (resolved) return normalizeExternalRangeValues(resolved, startRow, startCol, endRow, endCol)
	}
	if (!resolver?.resolveCell) return null
	const rows: CellValue[][] = []
	for (let row = startRow; row <= endRow; row++) {
		const values: CellValue[] = []
		for (let col = startCol; col <= endCol; col++) {
			values.push(resolver.resolveCell({ workbook, sheet, row, col }) ?? errorValue('#REF!'))
		}
		rows.push(values)
	}
	return rows
}

function normalizeExternalRangeValues(
	values: readonly (readonly CellValue[])[],
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): readonly (readonly CellValue[])[] {
	const rowCount = Math.max(0, endRow - startRow + 1)
	const colCount = Math.max(0, endCol - startCol + 1)
	const normalized: CellValue[][] = []
	for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
		const sourceRow = values[rowOffset]
		const row: CellValue[] = []
		for (let colOffset = 0; colOffset < colCount; colOffset++) {
			row.push(sourceRow?.[colOffset] ?? EMPTY)
		}
		normalized.push(row)
	}
	return normalized
}

function resolveExternalRange(
	ctx: EvalContext,
	sheet: string | undefined,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): EvalArg | null {
	const target = externalReferenceTarget(sheet)
	if (!target) return null
	if (!isRangeInBounds(startRow, startCol, endRow, endCol)) return { value: errorValue('#REF!') }
	const values = materializeExternalRange(
		ctx,
		target.workbook,
		target.sheet,
		startRow,
		startCol,
		endRow,
		endCol,
	)
	if (!values) return { value: errorValue('#REF!') }
	return makeExternalRangeArg(ctx.sheetIndex, startRow, startCol, endRow, endCol, values)
}

function isCellInBounds(row: number, col: number): boolean {
	return row >= 0 && row < EXCEL_MAX_ROWS && col >= 0 && col < EXCEL_MAX_COLS
}

function isRangeInBounds(
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): boolean {
	return (
		isCellInBounds(startRow, startCol) &&
		isCellInBounds(endRow, endCol) &&
		startRow <= endRow &&
		startCol <= endCol
	)
}

function getRangeValues(
	wb: Workbook,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): readonly (readonly CellValue[])[] {
	const sheet = wb.sheets[sheetIndex]
	if (!sheet) return [[errorValue('#REF!')]]
	const key = rangeCacheKey(sheetIndex, startRow, startCol, endRow, endCol)
	if (activeRangeValueCache) {
		const cached = activeRangeValueCache.get(key)
		if (cached) return cached
	}
	const rows: CellValue[][] = []
	for (let r = startRow; r <= endRow; r++) {
		const row: CellValue[] = []
		for (let c = startCol; c <= endCol; c++) {
			row.push(sheet.cells.readValue(r, c))
		}
		rows.push(row)
	}
	activeRangeValueCache?.set(key, rows)
	return rows
}

function coerceToNumber(v: CellValue): number | null {
	return toNumber(v)
}

function coerceToString(v: CellValue): string {
	return coerceCellValueToString(v)
}

function coerceToBoolean(v: CellValue): boolean | CellValue {
	v = topLeftScalar(v)
	switch (v.kind) {
		case 'empty':
			return false
		case 'number':
			return v.value !== 0
		case 'string': {
			const upper = v.value.toUpperCase()
			if (upper === 'TRUE') return true
			if (upper === 'FALSE') return false
			return errorValue('#VALUE!')
		}
		case 'boolean':
			return v.value
		case 'error':
			return v
		case 'date':
			return v.serial !== 0
		case 'richText':
			return errorValue('#VALUE!')
	}
}

function isReferenceBinaryOp(op: string): op is ',' | ' ' {
	return op === ',' || op === ' '
}

function isNumericBinaryOp(op: string): boolean {
	return op === '+' || op === '-' || op === '*' || op === '/' || op === '^'
}

function isEmptyStringLiteral(node: FormulaNode): boolean {
	return node.type === 'string' && node.value === ''
}

function isTextLikeValue(value: CellValue): boolean {
	const scalar = topLeftScalar(value)
	return scalar.kind === 'string' || scalar.kind === 'richText'
}

function evalBinary(op: string, left: CellValue, right: CellValue): CellValue {
	if (left.kind === 'array' || right.kind === 'array') {
		return evalArrayBinary(op, left, right)
	}
	return evalScalarBinary(op, left, right)
}

function evalArrayBinary(op: string, left: CellValue, right: CellValue): CellValue {
	const leftRows = left.kind === 'array' ? left.rows.length : 1
	const rightRows = right.kind === 'array' ? right.rows.length : 1
	const leftCols = left.kind === 'array' ? maxRowLength(left.rows) : 1
	const rightCols = right.kind === 'array' ? maxRowLength(right.rows) : 1
	const rows = broadcastLength(leftRows, rightRows)
	const cols = broadcastLength(leftCols, rightCols)
	if (rows === null || cols === null) return errorValue('#VALUE!')

	const result: ScalarCellValue[][] = []
	for (let row = 0; row < rows; row++) {
		const resultRow: ScalarCellValue[] = []
		for (let col = 0; col < cols; col++) {
			resultRow.push(
				topLeftScalar(
					evalScalarBinary(op, arrayCellAt(left, row, col), arrayCellAt(right, row, col)),
				),
			)
		}
		result.push(resultRow)
	}
	return arrayValue(result)
}

function maxRowLength(rows: readonly (readonly ScalarCellValue[])[]): number {
	let max = 0
	for (const row of rows) max = Math.max(max, row.length)
	return max
}

function broadcastLength(left: number, right: number): number | null {
	if (left === right) return left
	if (left === 1) return right
	if (right === 1) return left
	return null
}

function arrayCellAt(value: CellValue, row: number, col: number): CellValue {
	if (value.kind !== 'array') return value
	const sourceRow = value.rows[value.rows.length === 1 ? 0 : row]
	return sourceRow?.[sourceRow.length === 1 ? 0 : col] ?? EMPTY
}

function evalScalarBinary(op: string, left: CellValue, right: CellValue): CellValue {
	left = topLeftScalar(left)
	right = topLeftScalar(right)
	if (left.kind === 'error') return left
	if (right.kind === 'error') return right

	if (op === '&') {
		return stringValue(coerceToString(left) + coerceToString(right))
	}

	if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
		return evalComparison(op, left, right)
	}

	const ln = coerceToNumber(left)
	const rn = coerceToNumber(right)
	if (ln === null || rn === null) return errorValue('#VALUE!')

	switch (op) {
		case '+':
			return numberValue(ln + rn)
		case '-':
			return numberValue(ln - rn)
		case '*':
			return numberValue(ln * rn)
		case '/':
			return rn === 0 ? errorValue('#DIV/0!') : numberValue(ln / rn)
		case '^':
			return powerValue(ln, rn)
		default:
			return errorValue('#VALUE!')
	}
}

function powerValue(base: number, exponent: number): CellValue {
	const result = base ** exponent
	if (Number.isNaN(result)) return errorValue('#NUM!')
	if (!Number.isFinite(result)) {
		return base === 0 && exponent < 0 ? errorValue('#DIV/0!') : errorValue('#NUM!')
	}
	return numberValue(result)
}

function evalComparison(op: string, left: CellValue, right: CellValue): CellValue {
	const emptyComparison = compareEmptyOperand(op, left, right)
	if (emptyComparison !== undefined) return booleanValue(emptyComparison)

	const leftRank = comparisonRank(left)
	const rightRank = comparisonRank(right)

	if (leftRank !== rightRank) {
		return booleanValue(comparePrimitive(op, leftRank, rightRank))
	}

	if (leftRank === 0) {
		const ln = coerceToNumber(left)
		const rn = coerceToNumber(right)
		if (ln === null || rn === null) return errorValue('#VALUE!')
		return booleanValue(comparePrimitive(op, ln, rn))
	}

	if (leftRank === 1) {
		const ls = coerceToString(left).toLowerCase()
		const rs = coerceToString(right).toLowerCase()
		return booleanValue(comparePrimitive(op, ls, rs))
	}

	const lb = left.kind === 'boolean' && left.value ? 1 : 0
	const rb = right.kind === 'boolean' && right.value ? 1 : 0
	return booleanValue(comparePrimitive(op, lb, rb))
}

function comparisonRank(value: CellValue): 0 | 1 | 2 {
	const scalar = topLeftScalar(value)
	if (scalar.kind === 'boolean') return 2
	if (scalar.kind === 'string' || scalar.kind === 'richText') return 1
	return 0
}

function compareEmptyOperand(op: string, left: CellValue, right: CellValue): boolean | undefined {
	left = topLeftScalar(left)
	right = topLeftScalar(right)
	if (left.kind !== 'empty' && right.kind !== 'empty') return undefined
	if (left.kind === 'error' || right.kind === 'error') return undefined
	if (
		left.kind === 'string' ||
		right.kind === 'string' ||
		left.kind === 'richText' ||
		right.kind === 'richText'
	) {
		return comparePrimitive(
			op,
			coerceToString(left).toLowerCase(),
			coerceToString(right).toLowerCase(),
		)
	}
	if (left.kind === 'boolean' || right.kind === 'boolean') {
		const l = left.kind === 'boolean' && left.value ? 1 : 0
		const r = right.kind === 'boolean' && right.value ? 1 : 0
		return comparePrimitive(op, l, r)
	}
	const ln = coerceToNumber(left)
	const rn = coerceToNumber(right)
	if (ln === null || rn === null) return undefined
	return comparePrimitive(op, ln, rn)
}

function comparePrimitive<T extends number | string>(op: string, a: T, b: T): boolean {
	switch (op) {
		case '=':
			return a === b
		case '<>':
			return a !== b
		case '<':
			return a < b
		case '>':
			return a > b
		case '<=':
			return a <= b
		case '>=':
			return a >= b
		default:
			return false
	}
}

export function evaluate(node: FormulaNode, ctx: EvalContext): CellValue {
	switch (node.type) {
		case 'number':
			return numberValue(node.value)
		case 'string':
			return stringValue(node.value)
		case 'boolean':
			return booleanValue(node.value)
		case 'error':
			return errorValue(node.value)
		case 'missing':
			return EMPTY

		case 'cellRef': {
			const external = resolveExternalCell(ctx, node.sheet, node.ref.row, node.ref.col)
			if (external) return blankAsTopLevelReferenceValue(external)
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return errorValue('#REF!')
			return blankAsTopLevelReferenceValue(
				getCellValue(ctx.workbook, si, node.ref.row, node.ref.col),
			)
		}

		case 'rangeRef': {
			return evaluateReferenceNode(node, ctx)
		}

		case 'dynamicRangeRef': {
			return evaluateReferenceNode(node, ctx)
		}

		case 'wholeRowRange': {
			return evaluateReferenceNode(node, ctx)
		}

		case 'wholeColumnRange': {
			return evaluateReferenceNode(node, ctx)
		}

		case 'name': {
			return evaluateDefinedName(node.name, node.sheet, ctx)
		}

		case 'binary': {
			if (isReferenceBinaryOp(node.op)) return evaluateReferenceNode(node, ctx)
			let left = evaluateBinaryOperand(node.left, ctx)
			let right = evaluateBinaryOperand(node.right, ctx)
			if (isNumericBinaryOp(node.op)) {
				if (isEmptyStringLiteral(node.left) && !isTextLikeValue(right)) left = numberValue(0)
				if (isEmptyStringLiteral(node.right) && !isTextLikeValue(left)) right = numberValue(0)
			}
			return evalBinary(node.op, left, right)
		}

		case 'unary': {
			if (node.op === '@') {
				const refOperand = resolveReferenceNode(node.operand, ctx)
				if (refOperand) return implicitIntersect(refOperand, ctx)
				return topLeftScalar(evaluate(node.operand, ctx))
			}
			const operand = evaluate(node.operand, ctx)
			if (operand.kind === 'error') return operand
			if (operand.kind === 'array') return evalArrayUnary(node.op, operand)
			const n = coerceToNumber(operand)
			if (n === null) return errorValue('#VALUE!')
			switch (node.op) {
				case '+':
					return numberValue(n)
				case '-':
					return numberValue(-n)
				case '%':
					return numberValue(n / 100)
			}
			return EMPTY
		}

		case 'function':
			return evalFunction(node.name, node.args, ctx)

		case 'array': {
			const rows: ScalarCellValue[][] = []
			for (const sourceRow of node.rows) {
				const targetRow: ScalarCellValue[] = []
				for (const sourceNode of sourceRow) {
					targetRow.push(topLeftScalar(evaluate(sourceNode, ctx)))
				}
				rows.push(targetRow)
			}
			return arrayValue(rows)
		}

		case 'structuredRef': {
			const ref = resolveReferenceNode(node, ctx)
			if (ref) return implicitIntersect(ref, ctx)
			return evaluateReferenceNode(node, ctx)
		}
		case 'spillRef': {
			return evaluateReferenceNode(node, ctx)
		}
		case 'sheetSpanRef': {
			return evaluateReferenceNode(node, ctx)
		}
		default:
			return assertUnreachable(node)
	}
}

export function evaluateLegacyTopLevelFormula(
	node: FormulaNode,
	ctx: EvalContext,
): CellValue | null {
	switch (node.type) {
		case 'unary': {
			if (node.op === '@') return null
			const ref = resolveReferenceNode(node.operand, ctx)
			if (!ref) return null
			const operand = implicitIntersect(ref, ctx)
			if (operand.kind === 'error') return operand
			if (node.op === '+') {
				if (operand.kind === 'string' || operand.kind === 'richText') return operand
			}
			const n = coerceToNumber(operand)
			if (n === null) return errorValue('#VALUE!')
			switch (node.op) {
				case '+':
					return numberValue(n)
				case '-':
					return numberValue(-n)
				case '%':
					return numberValue(n / 100)
				default:
					return null
			}
		}
		case 'binary': {
			if (isReferenceBinaryOp(node.op)) return null
			const leftOperand = evalLegacyScalarOperand(node.left, ctx)
			const rightOperand = evalLegacyScalarOperand(node.right, ctx)
			if (!leftOperand.handled && !rightOperand.handled) return null
			const left = leftOperand.handled ? leftOperand.value : evaluate(node.left, ctx)
			const right = rightOperand.handled ? rightOperand.value : evaluate(node.right, ctx)
			return evalBinary(node.op, left, right)
		}
		case 'function':
			return evalLegacyTopLevelFunction(node.name, node.args, ctx)
		default:
			return null
	}
}

function evalLegacyScalarOperand(
	node: FormulaNode,
	ctx: EvalContext,
): { value: CellValue; handled: true } | { handled: false } {
	const ref = resolveReferenceNode(node, ctx)
	if (ref) return { value: implicitIntersect(ref, ctx), handled: true }
	if (node.type === 'unary') {
		const value = evaluateLegacyTopLevelFormula(node, ctx)
		if (value) return { value, handled: true }
	}
	if (node.type === 'binary' && !isReferenceBinaryOp(node.op)) {
		const left = evalLegacyScalarOperand(node.left, ctx)
		const right = evalLegacyScalarOperand(node.right, ctx)
		if (!left.handled && !right.handled) return { handled: false }
		const leftValue = left.handled ? left.value : evaluate(node.left, ctx)
		const rightValue = right.handled ? right.value : evaluate(node.right, ctx)
		return { value: evalBinary(node.op, leftValue, rightValue), handled: true }
	}
	return { handled: false }
}

function evalLegacyTopLevelFunction(
	name: string,
	argNodes: readonly FormulaNode[],
	ctx: EvalContext,
): CellValue | null {
	const upperName = name.toUpperCase()
	if (!LEGACY_TOP_LEVEL_SCALAR_FUNCTIONS.has(upperName)) return null
	const def = functionRegistry.get(upperName)
	if (!def) return null
	if (argNodes.length < def.minArgs || argNodes.length > def.maxArgs) return errorValue('#VALUE!')
	const args: EvalArg[] = new Array(argNodes.length)
	for (let i = 0; i < argNodes.length; i++) {
		const node = argNodes[i] as FormulaNode
		const ref = resolveReferenceNode(node, ctx)
		args[i] =
			ref && upperName !== 'ISREF' && upperName !== 'AND' && upperName !== 'OR'
				? { value: implicitIntersect(ref, ctx) }
				: resolveArg(node, ctx)
	}
	return def.evaluate(args, sharedFnCtx.update(ctx))
}

function evalArrayUnary(op: string, operand: Extract<CellValue, { kind: 'array' }>): CellValue {
	const result: ScalarCellValue[][] = []
	for (const sourceRow of operand.rows) {
		const resultRow: ScalarCellValue[] = []
		for (const cell of sourceRow) {
			if (cell.kind === 'error') {
				resultRow.push(cell)
				continue
			}
			const n = coerceToNumber(cell)
			if (n === null) {
				resultRow.push(topLeftScalar(errorValue('#VALUE!')))
				continue
			}
			switch (op) {
				case '+':
					resultRow.push(topLeftScalar(numberValue(n)))
					break
				case '-':
					resultRow.push(topLeftScalar(numberValue(-n)))
					break
				case '%':
					resultRow.push(topLeftScalar(numberValue(n / 100)))
					break
				default:
					resultRow.push(topLeftScalar(EMPTY))
			}
		}
		result.push(resultRow)
	}
	return arrayValue(result)
}

function blankAsTopLevelReferenceValue(value: CellValue): CellValue {
	return value.kind === 'empty' ? numberValue(0) : value
}

function evaluateBinaryOperand(node: FormulaNode, ctx: EvalContext): CellValue {
	if (node.type === 'wholeColumnRange' || node.type === 'wholeRowRange') {
		const ref = resolveReferenceNode(node, ctx)
		if (ref) return implicitIntersect(ref, ctx)
	}
	if (node.type === 'name' && !node.sheet && ctx.letBindings?.has(node.name.toLowerCase())) {
		return evaluate(node, ctx)
	}
	const ref = resolveReferenceNode(node, ctx)
	if (ref) return referenceArgToValue(ref)
	return evaluate(node, ctx)
}

function evalFunction(name: string, argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const upperName = name.toUpperCase()
	if (upperName === 'ROW' && argNodes.length === 0) {
		return numberValue(ctx.row + 1)
	}
	if (upperName === 'COLUMN' && argNodes.length === 0) {
		return numberValue(ctx.col + 1)
	}
	if (upperName === 'INDIRECT' || upperName === 'OFFSET') {
		return implicitIntersect(resolveReferenceFunction(upperName, argNodes, ctx), ctx)
	}
	if (upperName === 'FORMULATEXT') {
		return evalFormulaText(argNodes, ctx)
	}
	if (upperName === 'GETPIVOTDATA') {
		return evalGetPivotData(argNodes, ctx)
	}
	if (upperName === 'LET') {
		return evalLet(argNodes, ctx)
	}
	if (upperName === 'LAMBDA') {
		return errorValue('#CALC!')
	}
	if (upperName === 'MAP') {
		return evalMap(argNodes, ctx)
	}
	if (upperName === 'REDUCE') {
		return evalReduce(argNodes, ctx)
	}
	if (upperName === 'SCAN') {
		return evalScan(argNodes, ctx)
	}
	if (upperName === 'BYROW') {
		return evalByRow(argNodes, ctx)
	}
	if (upperName === 'BYCOL') {
		return evalByCol(argNodes, ctx)
	}
	if (upperName === 'MAKEARRAY') {
		return evalMakeArray(argNodes, ctx)
	}
	if (upperName === 'ISFORMULA') {
		return evalIsFormula(argNodes, ctx)
	}
	if (upperName === 'CELL') {
		return evalCellInfo(argNodes, ctx)
	}
	if (upperName === 'SHEETS') {
		return evalSheets(argNodes, ctx)
	}
	if (upperName === 'SHEET') {
		return evalSheet(argNodes, ctx)
	}
	if (upperName === 'IF') {
		return evalIf(argNodes, ctx)
	}
	if (upperName === 'IFERROR') {
		return evalIfError(argNodes, ctx)
	}
	if (upperName === 'IFNA') {
		return evalIfNa(argNodes, ctx)
	}
	if (upperName === 'CHOOSE') {
		return evalChoose(argNodes, ctx)
	}
	if (upperName === 'SWITCH') {
		return evalSwitch(argNodes, ctx)
	}
	if (upperName === 'IFS') {
		return evalIfs(argNodes, ctx)
	}
	if (upperName === '__CALL__') {
		return evalCall(argNodes, ctx)
	}

	if (
		argNodes.length === 1 &&
		argNodes[0]?.type === 'rangeRef' &&
		(upperName === 'SUM' ||
			upperName === 'COUNT' ||
			upperName === 'AVERAGE' ||
			upperName === 'MIN' ||
			upperName === 'MAX')
	) {
		const rangeNode = argNodes[0]
		const si = resolveSheetIndex(ctx.workbook, rangeNode.sheet, ctx.sheetIndex)
		if (si >= 0) {
			const { sum, count, min, max, error } = aggregateNumericRange(
				ctx.workbook.sheets[si],
				rangeNode.start.row,
				rangeNode.start.col,
				rangeNode.end.row,
				rangeNode.end.col,
			)
			if (error) return error
			switch (upperName) {
				case 'SUM':
					return numberValue(sum)
				case 'COUNT':
					return numberValue(count)
				case 'AVERAGE':
					return count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
				case 'MIN':
					return numberValue(count === 0 ? 0 : min)
				case 'MAX':
					return numberValue(count === 0 ? 0 : max)
			}
		}
	}

	const def = functionRegistry.get(upperName)
	if (!def) {
		const lambda = extractLambdaFromName(name, ctx)
		if (lambda) {
			const evaluatedArgs = argNodes.map((node) => evaluate(node, ctx))
			return invokeLambda(lambda, evaluatedArgs)
		}
		return errorValue('#NAME?')
	}
	if (argNodes.length < def.minArgs || argNodes.length > def.maxArgs) {
		return errorValue('#VALUE!')
	}
	if (ARRAY_CONTEXT_MAPPABLE_FUNCTIONS.has(upperName) && usesFormulaArraySemantics(ctx)) {
		const mappedArgs = argNodes.map((node) => resolveArg(node, ctx))
		const mapped = evalMappedScalarFunction(def, mappedArgs, ctx)
		if (mapped) return mapped
	}

	// EvalArg pooling: Not implemented. resolveArg returns many shapes (simple { value },
	// range refs with ref/areas/forEachValue, multi-area with getters). Only the scalar path
	// is poolable; range/ref paths create complex objects. V8 allocates small objects quickly;
	// calc benchmarks show no allocation bottleneck. Pooling would require lifecycle management
	// (reset between cell evals) and conditional fill logic, adding complexity without proven gain.
	const args: EvalArg[] = new Array(argNodes.length)
	for (let i = 0; i < argNodes.length; i++) {
		args[i] = resolveFunctionArg(argNodes[i] as FormulaNode, ctx, upperName)
	}
	return def.evaluate(args, sharedFnCtx.update(ctx))
}

function usesFormulaArraySemantics(ctx: EvalContext): boolean {
	const binding = ctx.workbook.sheets[ctx.sheetIndex]?.cells.get(ctx.row, ctx.col)?.formulaInfo
	return (
		binding?.kind === 'array' ||
		binding?.kind === 'dynamicArray' ||
		binding?.kind === 'spill' ||
		binding?.kind === 'blockedSpill'
	)
}

function evalMappedScalarFunction(
	def: FunctionDef,
	args: readonly EvalArg[],
	ctx: EvalContext,
): CellValue | null {
	let rows = 1
	let cols = 1
	const ranges = args.map((arg) => {
		const range = getRange(arg)
		const rangeRows = range.length
		let rangeCols = 0
		for (const row of range) rangeCols = Math.max(rangeCols, row.length)
		rows = Math.max(rows, rangeRows)
		cols = Math.max(cols, rangeCols)
		return { arg, range, rows: rangeRows, cols: rangeCols }
	})
	if (rows === 1 && cols === 1) return null
	for (const range of ranges) {
		if ((range.rows !== 1 && range.rows !== rows) || (range.cols !== 1 && range.cols !== cols)) {
			return errorValue('#VALUE!')
		}
	}
	const mappedRows: ScalarCellValue[][] = []
	for (let row = 0; row < rows; row++) {
		const mappedRow: ScalarCellValue[] = []
		for (let col = 0; col < cols; col++) {
			const cellArgs = ranges.map(({ range, rows: rangeRows, cols: rangeCols }) => {
				const sourceRow = rangeRows === 1 ? 0 : row
				const sourceCol = rangeCols === 1 ? 0 : col
				return { value: topLeftScalar(range[sourceRow]?.[sourceCol] ?? EMPTY) }
			})
			mappedRow.push(topLeftScalar(def.evaluate(cellArgs, sharedFnCtx.update(ctx))))
		}
		mappedRows.push(mappedRow)
	}
	return arrayValue(mappedRows)
}

function resolveFunctionArg(node: FormulaNode, ctx: EvalContext, functionName: string): EvalArg {
	if (SCALAR_IMPLICIT_INTERSECTION_FUNCTIONS.has(functionName)) {
		if (node.type === 'name' && !node.sheet && ctx.letBindings?.has(node.name.toLowerCase())) {
			return resolveArg(node, ctx)
		}
		const ref = resolveReferenceNode(node, ctx)
		if (ref) return { value: implicitIntersect(ref, ctx) }
	}
	return resolveArg(node, ctx)
}

function evalLazyArg(node: FormulaNode | undefined, ctx: EvalContext): CellValue {
	return referenceArgToValue(resolveArg(node ?? { type: 'missing' }, ctx))
}

function evalIf(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const conditionValue = evalLazyArg(argNodes[0], ctx)
	if (conditionValue.kind === 'array') {
		const trueValue = evalLazyArg(argNodes[1], ctx)
		const falseValue = argNodes.length >= 3 ? evalLazyArg(argNodes[2], ctx) : booleanValue(false)
		return evalIfArray(conditionValue, trueValue, falseValue)
	}
	const condition = coerceToBoolean(conditionValue)
	if (typeof condition !== 'boolean') return condition
	if (condition) return evalLazyArg(argNodes[1], ctx)
	return argNodes.length >= 3 ? evalLazyArg(argNodes[2], ctx) : booleanValue(false)
}

function evalIfArray(
	conditionValue: CellValue,
	trueValue: CellValue,
	falseValue: CellValue,
): CellValue {
	const condRows = conditionValue.kind === 'array' ? conditionValue.rows.length : 1
	const trueRows = trueValue.kind === 'array' ? trueValue.rows.length : 1
	const falseRows = falseValue.kind === 'array' ? falseValue.rows.length : 1
	const condCols = conditionValue.kind === 'array' ? maxRowLength(conditionValue.rows) : 1
	const trueCols = trueValue.kind === 'array' ? maxRowLength(trueValue.rows) : 1
	const falseCols = falseValue.kind === 'array' ? maxRowLength(falseValue.rows) : 1
	const conditionTrueRows = broadcastLength(condRows, trueRows)
	const rows = conditionTrueRows === null ? null : broadcastLength(conditionTrueRows, falseRows)
	const conditionTrueCols = broadcastLength(condCols, trueCols)
	const cols = conditionTrueCols === null ? null : broadcastLength(conditionTrueCols, falseCols)
	if (rows === null || cols === null) return errorValue('#VALUE!')

	const result: ScalarCellValue[][] = []
	for (let row = 0; row < rows; row++) {
		const resultRow: ScalarCellValue[] = []
		for (let col = 0; col < cols; col++) {
			const condition = coerceToBoolean(arrayCellAt(conditionValue, row, col))
			if (typeof condition !== 'boolean') {
				resultRow.push(topLeftScalar(condition))
			} else {
				resultRow.push(topLeftScalar(arrayCellAt(condition ? trueValue : falseValue, row, col)))
			}
		}
		result.push(resultRow)
	}
	return arrayValue(result)
}

function evalIfError(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const value = evalLazyArg(argNodes[0], ctx)
	return value.kind === 'error' ? evalLazyArg(argNodes[1], ctx) : value
}

function evalIfNa(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const value = evalLazyArg(argNodes[0], ctx)
	return value.kind === 'error' && value.value === '#N/A' ? evalLazyArg(argNodes[1], ctx) : value
}

function evalChoose(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length < 2) return errorValue('#VALUE!')
	const idxVal = evaluate(argNodes[0] ?? { type: 'missing' }, ctx)
	if (idxVal.kind === 'error') return idxVal
	const n = coerceToNumber(idxVal)
	if (n === null) return errorValue('#VALUE!')
	const idx = Math.floor(n)
	if (idx < 1 || idx >= argNodes.length) return errorValue('#VALUE!')
	return evalLazyArg(argNodes[idx], ctx)
}

function evalSwitch(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length < 3) return errorValue('#VALUE!')
	const expr = evaluate(argNodes[0] ?? { type: 'missing' }, ctx)
	if (expr.kind === 'error') return expr
	const remaining = argNodes.length - 1
	const hasDefault = remaining % 2 === 1
	const pairEnd = hasDefault ? argNodes.length - 1 : argNodes.length
	for (let i = 1; i < pairEnd; i += 2) {
		const val = evaluate(argNodes[i] ?? { type: 'missing' }, ctx)
		if (val.kind === 'error') return val
		if (switchValuesMatch(expr, val)) return evalLazyArg(argNodes[i + 1], ctx)
	}
	return hasDefault ? evalLazyArg(argNodes[argNodes.length - 1], ctx) : errorValue('#N/A')
}

function switchValuesMatch(a: CellValue, b: CellValue): boolean {
	if (a.kind === 'number' && b.kind === 'number') return a.value === b.value
	if (a.kind === 'string' && b.kind === 'string')
		return a.value.toLowerCase() === b.value.toLowerCase()
	if (a.kind === 'boolean' && b.kind === 'boolean') return a.value === b.value
	if (a.kind === 'error' && b.kind === 'error') return a.value === b.value
	if (a.kind === 'empty' && b.kind === 'empty') return true
	return false
}

function evalIfs(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length < 2 || argNodes.length % 2 !== 0) return errorValue('#VALUE!')
	for (let i = 0; i + 1 < argNodes.length; i += 2) {
		const v = evaluate(argNodes[i] ?? { type: 'missing' }, ctx)
		if (v.kind === 'error') return v
		const cond = coerceToBoolean(v)
		if (typeof cond !== 'boolean') return cond
		if (cond) return evalLazyArg(argNodes[i + 1], ctx)
	}
	return errorValue('#N/A')
}

function evalFormulaText(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const arg = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	if (!arg.ref) {
		if (arg.value.kind === 'error') return arg.value
		return errorValue('#N/A')
	}
	const sheet = ctx.workbook.sheets[arg.ref.sheetIndex]
	if (!sheet) return errorValue('#N/A')
	const formula = sheet.cells.readFormula(arg.ref.row, arg.ref.col)
	if (!formula) return errorValue('#N/A')
	return stringValue(`=${formula}`)
}

function evalGetPivotData(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length < 2 || argNodes.length % 2 !== 0) return errorValue('#REF!')
	const dataFieldValue = evaluate(argNodes[0] ?? { type: 'missing' }, ctx)
	if (dataFieldValue.kind === 'error') return dataFieldValue
	const dataField = coerceToString(dataFieldValue)
	if (normalizePivotText(dataField).length === 0) return errorValue('#REF!')

	const anchor = resolveReferenceNode(argNodes[1] ?? { type: 'missing' }, ctx)
	if (!anchor?.ref || anchor.ref.kind !== 'cell') {
		return anchor?.value.kind === 'error' ? anchor.value : errorValue('#REF!')
	}
	const anchorSheet = ctx.workbook.sheets[anchor.ref.sheetIndex]
	if (!anchorSheet) return errorValue('#REF!')

	const filters: { field: string; item: string }[] = []
	for (let i = 2; i + 1 < argNodes.length; i += 2) {
		const fieldValue = evaluate(argNodes[i] ?? { type: 'missing' }, ctx)
		if (fieldValue.kind === 'error') return fieldValue
		const itemValue = evaluate(argNodes[i + 1] ?? { type: 'missing' }, ctx)
		if (itemValue.kind === 'error') return itemValue
		filters.push({
			field: normalizePivotText(coerceToString(fieldValue)),
			item: normalizePivotText(coerceToString(itemValue)),
		})
	}

	for (const pivot of ctx.workbook.pivotTables) {
		if (pivot.sheetName !== anchorSheet.name || !pivot.locationRef) continue
		const bounds = parsePivotLocation(pivot.locationRef)
		if (!bounds) continue
		if (
			anchor.ref.row < bounds.startRow ||
			anchor.ref.row > bounds.endRow ||
			anchor.ref.col < bounds.startCol ||
			anchor.ref.col > bounds.endCol
		) {
			continue
		}
		const value = lookupVisiblePivotValue(ctx, anchor.ref.sheetIndex, bounds, dataField, filters)
		if (value) return value
	}
	return errorValue('#REF!')
}

function parsePivotLocation(locationRef: string): {
	startRow: number
	startCol: number
	endRow: number
	endCol: number
} | null {
	try {
		const range = parseRange(locationRef)
		return {
			startRow: range.start.row,
			startCol: range.start.col,
			endRow: range.end.row,
			endCol: range.end.col,
		}
	} catch {
		return null
	}
}

function lookupVisiblePivotValue(
	ctx: EvalContext,
	sheetIndex: number,
	bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
	dataField: string,
	filters: readonly { field: string; item: string }[],
): CellValue | null {
	const header = findPivotDataHeader(ctx, sheetIndex, bounds, dataField)
	if (!header) return null
	const fieldColumns = filters.map((filter) => ({
		filter,
		col: findPivotFieldColumn(ctx, sheetIndex, bounds, header.row, header.col, filter.field),
	}))
	if (fieldColumns.some((entry) => entry.col === null)) return null

	for (let row = header.row + 1; row <= bounds.endRow; row++) {
		let matches = true
		if (fieldColumns.length === 0) {
			matches =
				normalizePivotText(
					coerceToString(getCellValue(ctx.workbook, sheetIndex, row, bounds.startCol)),
				) === 'grand total'
		} else {
			for (const entry of fieldColumns) {
				const col = entry.col
				if (col === null) return null
				const value = normalizePivotText(
					coerceToString(getCellValue(ctx.workbook, sheetIndex, row, col)),
				)
				if (value !== entry.filter.item) {
					matches = false
					break
				}
			}
		}
		if (matches) {
			if (fieldColumns.length === 0 && header.col === bounds.startCol) {
				const rowTotal = rightmostVisiblePivotRowValue(ctx, sheetIndex, bounds, row)
				if (rowTotal) return rowTotal
			}
			return getCellValue(ctx.workbook, sheetIndex, row, header.col)
		}
	}
	return null
}

function rightmostVisiblePivotRowValue(
	ctx: EvalContext,
	sheetIndex: number,
	bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
	row: number,
): CellValue | null {
	for (let col = bounds.endCol; col > bounds.startCol; col--) {
		const value = getCellValue(ctx.workbook, sheetIndex, row, col)
		if (value.kind !== 'empty') return value
	}
	return null
}

function findPivotDataHeader(
	ctx: EvalContext,
	sheetIndex: number,
	bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
	dataField: string,
): { row: number; col: number } | null {
	const maxHeaderRows = Math.min(bounds.endRow, bounds.startRow + 8)
	for (let row = bounds.startRow; row <= maxHeaderRows; row++) {
		for (let col = bounds.startCol; col <= bounds.endCol; col++) {
			const value = normalizePivotText(
				coerceToString(getCellValue(ctx.workbook, sheetIndex, row, col)),
			)
			if (pivotDataFieldCaptionsMatch(dataField, value)) return { row, col }
		}
	}
	return null
}

function findPivotFieldColumn(
	ctx: EvalContext,
	sheetIndex: number,
	bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
	headerRow: number,
	dataCol: number,
	fieldName: string,
): number | null {
	for (let col = bounds.startCol; col < dataCol; col++) {
		const header = normalizePivotText(
			coerceToString(getCellValue(ctx.workbook, sheetIndex, headerRow, col)),
		)
		if (header === fieldName) return col
	}
	if (dataCol > bounds.startCol) return bounds.startCol
	return null
}

function normalizePivotText(value: string): string {
	return value.trim().replace(/\s+/g, ' ').toLowerCase()
}

function resolveArg(node: FormulaNode, ctx: EvalContext): EvalArg {
	if (node.type === 'function') {
		const upperName = node.name.toUpperCase()
		if (upperName === 'INDIRECT' || upperName === 'OFFSET') {
			return resolveReferenceFunction(upperName, node.args, ctx)
		}
		if (upperName === 'INDEX') {
			const ref = resolveReferenceNode(node, ctx)
			if (ref) return ref
		}
		if (upperName === 'LET') {
			const value = evalLet(node.args, ctx)
			return value.kind === 'array' ? { value, kind: 'range', values: value.rows } : { value }
		}
	}

	if (node.type === 'name') {
		if (!node.sheet && ctx.letBindings?.has(node.name.toLowerCase())) {
			const bound = ctx.letBindings.get(node.name.toLowerCase()) as LetBinding
			if (isLambdaBinding(bound)) return { value: errorValue('#CALC!') }
			if (bound.kind === 'array') return { value: bound, kind: 'range', values: bound.rows }
			return { value: bound }
		}
		const resolved = resolveDefinedName(node.name, node.sheet, ctx)
		if (!resolved) return { value: errorValue('#NAME?') }
		return resolveArg(resolved.ast, resolved.ctx)
	}

	if (node.type === 'spillRef') {
		return resolveSpillReference(node.target, ctx) ?? { value: errorValue('#REF!') }
	}

	const refResult = resolveReferenceNode(node, ctx)
	if (refResult) return refResult

	const value = evaluate(node, ctx)
	if (value.kind === 'array') {
		return {
			value,
			kind: 'range',
			values: value.rows,
		}
	}
	return { value }
}

function resolveReferenceFunction(
	name: 'INDIRECT' | 'OFFSET' | 'INDEX',
	argNodes: readonly FormulaNode[],
	ctx: EvalContext,
): EvalArg {
	switch (name) {
		case 'INDIRECT':
			return resolveIndirectReference(argNodes, ctx)
		case 'OFFSET':
			return resolveOffsetReference(argNodes, ctx)
		case 'INDEX':
			return resolveIndexReference(argNodes, ctx)
	}
}

function resolveIndirectReference(argNodes: readonly FormulaNode[], ctx: EvalContext): EvalArg {
	const textValue = evaluate(argNodes[0] ?? { type: 'missing' }, ctx)
	if (textValue.kind === 'error') return { value: textValue }
	const refText = coerceToString(textValue).trim()
	if (refText.length === 0) return { value: errorValue('#REF!') }

	const useA1Value =
		argNodes.length > 1 ? evaluate(argNodes[1] as FormulaNode, ctx) : booleanValue(true)
	if (useA1Value.kind === 'error') return { value: useA1Value }
	const useA1 =
		useA1Value.kind === 'boolean'
			? useA1Value.value
			: useA1Value.kind === 'number'
				? useA1Value.value !== 0
				: true

	if (useA1) {
		const parsed = cachedParseFormula(refText)
		if (!parsed.ok) return { value: errorValue('#REF!') }
		const resolved = resolveReferenceNode(parsed.value, ctx)
		return resolved ?? { value: errorValue('#REF!') }
	}

	return resolveR1C1Reference(refText, ctx) ?? { value: errorValue('#REF!') }
}

function resolveOffsetReference(argNodes: readonly FormulaNode[], ctx: EvalContext): EvalArg {
	const base = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	if (base.value.kind === 'error') return base
	if (!base.ref) return { value: errorValue('#VALUE!') }

	const rowOffset = offsetScalarNumberArg(argNodes[1], ctx)
	if (typeof rowOffset !== 'number') return { value: rowOffset }
	const colOffset = offsetScalarNumberArg(argNodes[2], ctx)
	if (typeof colOffset !== 'number') return { value: colOffset }

	const baseHeight =
		base.ref.kind === 'range' ? (base.ref.endRow ?? base.ref.row) - base.ref.row + 1 : 1
	const baseWidth =
		base.ref.kind === 'range' ? (base.ref.endCol ?? base.ref.col) - base.ref.col + 1 : 1

	const height =
		argNodes.length > 3 && argNodes[3]?.type !== 'missing'
			? offsetScalarNumberArg(argNodes[3], ctx)
			: baseHeight
	if (typeof height !== 'number') return { value: height }
	const width =
		argNodes.length > 4 && argNodes[4]?.type !== 'missing'
			? offsetScalarNumberArg(argNodes[4], ctx)
			: baseWidth
	if (typeof width !== 'number') return { value: width }

	const targetHeight = Math.trunc(height)
	const targetWidth = Math.trunc(width)
	if (targetHeight <= 0 || targetWidth <= 0) return { value: errorValue('#REF!') }

	const startRow = base.ref.row + Math.trunc(rowOffset)
	const startCol = base.ref.col + Math.trunc(colOffset)
	const endRow = startRow + targetHeight - 1
	const endCol = startCol + targetWidth - 1
	if (startRow < 0 || startCol < 0 || endRow < 0 || endCol < 0) {
		return { value: errorValue('#REF!') }
	}

	return makeRangeArg(
		ctx.workbook,
		base.ref.sheetIndex,
		startRow,
		startCol,
		endRow,
		endCol,
		ctx.calcContext.dateSystem,
		ctx.calcContext.today,
	)
}

function offsetNumberArg(node: FormulaNode | undefined, ctx: EvalContext): number | CellValue {
	const value = evaluate(node ?? { type: 'missing' }, ctx)
	if (value.kind === 'error') return value
	const number = coerceToNumber(value)
	return number === null ? errorValue('#VALUE!') : number
}

function offsetScalarNumberArg(
	node: FormulaNode | undefined,
	ctx: EvalContext,
): number | CellValue {
	const scalarNode = node ?? { type: 'missing' }
	const ref = resolveReferenceNode(scalarNode, ctx)
	const value = ref ? implicitIntersect(ref, ctx) : evaluate(scalarNode, ctx)
	if (value.kind === 'error') return value
	const number = coerceToNumber(value)
	return number === null ? errorValue('#VALUE!') : number
}

function resolveIndexReference(argNodes: readonly FormulaNode[], ctx: EvalContext): EvalArg {
	const source = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	if (source.value.kind === 'error' && !source.ref && !source.areas?.length) return source
	const areas = areasOf(source)
	if (!areas || areas.length !== 1) return { value: errorValue('#VALUE!') }
	const area = areas[0]
	if (!area) return { value: errorValue('#VALUE!') }
	const bounds = toAreaBounds(area.ref)
	const height = bounds.endRow - bounds.startRow + 1
	const width = bounds.endCol - bounds.startCol + 1

	const rowNum = offsetNumberArg(argNodes[1], ctx)
	if (typeof rowNum !== 'number') return { value: rowNum }
	const row = Math.floor(rowNum)

	if (argNodes.length > 2) {
		const colNum = offsetNumberArg(argNodes[2], ctx)
		if (typeof colNum !== 'number') return { value: colNum }
		const col = Math.floor(colNum)
		if (row === 0 && col === 0) return { value: errorValue('#VALUE!') }
		if (row === 0) {
			if (col < 1 || col > width) return { value: errorValue('#REF!') }
			const targetCol = bounds.startCol + col - 1
			return makeRangeArg(
				ctx.workbook,
				bounds.sheetIndex,
				bounds.startRow,
				targetCol,
				bounds.endRow,
				targetCol,
				ctx.calcContext.dateSystem,
				ctx.calcContext.today,
			)
		}
		if (col === 0) {
			if (row < 1 || row > height) return { value: errorValue('#REF!') }
			const targetRow = bounds.startRow + row - 1
			return makeRangeArg(
				ctx.workbook,
				bounds.sheetIndex,
				targetRow,
				bounds.startCol,
				targetRow,
				bounds.endCol,
				ctx.calcContext.dateSystem,
				ctx.calcContext.today,
			)
		}
		if (row < 1 || row > height || col < 1 || col > width) {
			return { value: errorValue('#REF!') }
		}
		return makeRangeArg(
			ctx.workbook,
			bounds.sheetIndex,
			bounds.startRow + row - 1,
			bounds.startCol + col - 1,
			bounds.startRow + row - 1,
			bounds.startCol + col - 1,
			ctx.calcContext.dateSystem,
			ctx.calcContext.today,
		)
	}

	if (height === 1) {
		if (row < 1 || row > width) return { value: errorValue('#REF!') }
		return makeRangeArg(
			ctx.workbook,
			bounds.sheetIndex,
			bounds.startRow,
			bounds.startCol + row - 1,
			bounds.startRow,
			bounds.startCol + row - 1,
			ctx.calcContext.dateSystem,
			ctx.calcContext.today,
		)
	}
	if (row < 1 || row > height) return { value: errorValue('#REF!') }
	return makeRangeArg(
		ctx.workbook,
		bounds.sheetIndex,
		bounds.startRow + row - 1,
		bounds.startCol,
		bounds.startRow + row - 1,
		bounds.startCol,
		ctx.calcContext.dateSystem,
		ctx.calcContext.today,
	)
}

function resolveReferenceNode(node: FormulaNode, ctx: EvalContext): EvalArg | null {
	switch (node.type) {
		case 'cellRef': {
			const external = resolveExternalCell(ctx, node.sheet, node.ref.row, node.ref.col)
			if (external) {
				return {
					value: external,
					ref: {
						kind: 'cell',
						sheetIndex: ctx.sheetIndex,
						row: node.ref.row,
						col: node.ref.col,
					},
					valueAtOffset: (rowOffset, colOffset) =>
						resolveExternalCell(
							ctx,
							node.sheet,
							node.ref.row + rowOffset,
							node.ref.col + colOffset,
						) ?? errorValue('#REF!'),
				}
			}
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return {
				value: getCellValue(ctx.workbook, si, node.ref.row, node.ref.col),
				ref: {
					kind: 'cell',
					sheetIndex: si,
					row: node.ref.row,
					col: node.ref.col,
				},
				valueAtOffset: (rowOffset, colOffset) =>
					getCellValue(ctx.workbook, si, node.ref.row + rowOffset, node.ref.col + colOffset),
				formulaAtOffset: (rowOffset, colOffset) =>
					ctx.workbook.sheets[si]?.cells.readFormula(
						node.ref.row + rowOffset,
						node.ref.col + colOffset,
					),
			}
		}
		case 'rangeRef': {
			const external = resolveExternalRange(
				ctx,
				node.sheet,
				node.start.row,
				node.start.col,
				node.end.row,
				node.end.col,
			)
			if (external) return external
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeRangeArg(
				ctx.workbook,
				si,
				node.start.row,
				node.start.col,
				node.end.row,
				node.end.col,
				ctx.calcContext.dateSystem,
				ctx.calcContext.today,
			)
		}
		case 'dynamicRangeRef':
			return resolveDynamicRangeReference(node.start, node.end, ctx)
		case 'wholeRowRange': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeWholeRowArg(
				ctx.workbook,
				si,
				node.startRow,
				node.endRow,
				ctx.calcContext.dateSystem,
				ctx.calcContext.today,
			)
		}
		case 'wholeColumnRange': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeWholeColumnArg(
				ctx.workbook,
				si,
				node.startCol,
				node.endCol,
				ctx.calcContext.dateSystem,
				ctx.calcContext.today,
			)
		}
		case 'name': {
			const resolved = resolveDefinedName(node.name, node.sheet, ctx)
			if (!resolved) return { value: errorValue('#NAME?') }
			return (
				resolveReferenceNode(resolved.ast, resolved.ctx) ?? resolveArg(resolved.ast, resolved.ctx)
			)
		}
		case 'structuredRef': {
			const resolved = (
				ctx.structuredRefResolver ?? createStructuredRefResolver(ctx.workbook)
			).resolve(node, ctx.sheetIndex, ctx.row, ctx.col)
			if (!resolved) return { value: errorValue('#REF!') }
			return makeRangeArg(
				ctx.workbook,
				resolved.sheetIndex,
				resolved.startRow,
				resolved.startCol,
				resolved.endRow,
				resolved.endCol,
				ctx.calcContext.dateSystem,
				ctx.calcContext.today,
			)
		}
		case 'spillRef':
			return resolveSpillReference(node.target, ctx)
		case 'sheetSpanRef': {
			const sheetIndices = resolveSheetSpanIndices(ctx.workbook, node.startSheet, node.endSheet)
			if (!sheetIndices) return { value: errorValue('#REF!') }
			const areas: EvalArea[] = []
			for (const sheetIndex of sheetIndices) {
				const targetRef = resolveReferenceNode(
					applySheetToReferenceNode(node.target, sheetIndex, ctx),
					ctx,
				)
				if (!targetRef) return { value: errorValue('#REF!') }
				const targetAreas = areasOf(targetRef)
				if (!targetAreas) return { value: errorValue('#REF!') }
				areas.push(...targetAreas)
			}
			return areas.length > 0 ? makeMultiAreaArg(areas) : { value: errorValue('#REF!') }
		}
		case 'binary': {
			if (!isReferenceBinaryOp(node.op)) return null
			const left = resolveReferenceNode(node.left, ctx)
			const right = resolveReferenceNode(node.right, ctx)
			if (!left || !right) return null
			const leftAreas = areasOf(left)
			const rightAreas = areasOf(right)
			if (!leftAreas || !rightAreas) return { value: errorValue('#VALUE!') }
			if (node.op === ',') {
				return makeMultiAreaArg([...leftAreas, ...rightAreas])
			}
			const intersections: EvalArea[] = []
			for (const leftArea of leftAreas) {
				for (const rightArea of rightAreas) {
					const leftRef = toAreaBounds(leftArea.ref)
					const rightRef = toAreaBounds(rightArea.ref)
					if (leftRef.sheetIndex !== rightRef.sheetIndex) continue
					const startRow = Math.max(leftRef.startRow, rightRef.startRow)
					const startCol = Math.max(leftRef.startCol, rightRef.startCol)
					const endRow = Math.min(leftRef.endRow, rightRef.endRow)
					const endCol = Math.min(leftRef.endCol, rightRef.endCol)
					if (startRow > endRow || startCol > endCol) continue
					intersections.push(
						makeRangeArea(ctx.workbook, leftRef.sheetIndex, startRow, startCol, endRow, endCol, {
							dateSystem: ctx.calcContext.dateSystem,
							today: ctx.calcContext.today,
						}),
					)
				}
			}
			return intersections.length > 0
				? makeMultiAreaArg(intersections)
				: { value: errorValue('#NULL!') }
		}
		case 'function': {
			const upperName = node.name.toUpperCase()
			if (upperName === 'INDIRECT' || upperName === 'OFFSET' || upperName === 'INDEX') {
				return resolveReferenceFunction(upperName, node.args, ctx)
			}
			return null
		}
		default:
			return null
	}
}

function resolveDynamicRangeReference(
	startNode: FormulaNode,
	endNode: FormulaNode,
	ctx: EvalContext,
): EvalArg {
	const start = resolveReferenceNode(startNode, ctx)
	if (!start) return { value: errorValue('#VALUE!') }
	if (start.value.kind === 'error') return start
	const end = resolveReferenceNode(endNode, ctx)
	if (!end) return { value: errorValue('#VALUE!') }
	if (end.value.kind === 'error') return end
	const startAreas = areasOf(start)
	const endAreas = areasOf(end)
	if (!startAreas || !endAreas || startAreas.length !== 1 || endAreas.length !== 1) {
		return { value: errorValue('#VALUE!') }
	}
	const startArea = startAreas[0]
	const endArea = endAreas[0]
	if (!startArea || !endArea) return { value: errorValue('#VALUE!') }
	const startBounds = toAreaBounds(startArea.ref)
	const endBounds = toAreaBounds(endArea.ref)
	if (startBounds.sheetIndex !== endBounds.sheetIndex) return { value: errorValue('#VALUE!') }
	return makeRangeArg(
		ctx.workbook,
		startBounds.sheetIndex,
		startBounds.startRow,
		startBounds.startCol,
		endBounds.endRow,
		endBounds.endCol,
		ctx.calcContext.dateSystem,
		ctx.calcContext.today,
	)
}

function resolveSpillReference(target: FormulaNode, ctx: EvalContext): EvalArg | null {
	const targetRef = resolveReferenceNode(target, ctx) ?? resolveArg(target, ctx)
	if (!targetRef.ref) return null
	const sheet = ctx.workbook.sheets[targetRef.ref.sheetIndex]
	if (!sheet) return null
	const binding = sheet.cells.readFormulaInfo(targetRef.ref.row, targetRef.ref.col)
	if (!binding || !isSpillFormulaBinding(binding)) return { value: errorValue('#REF!') }
	const parsed = cachedParseFormula(binding.ref)
	if (!parsed.ok || parsed.value.type !== 'rangeRef') return null
	return makeRangeArg(
		ctx.workbook,
		targetRef.ref.sheetIndex,
		parsed.value.start.row,
		parsed.value.start.col,
		parsed.value.end.row,
		parsed.value.end.col,
		ctx.calcContext.dateSystem,
		ctx.calcContext.today,
	)
}

function isSpillFormulaBinding(
	binding: unknown,
): binding is { kind: 'spill'; anchorRef: string; ref: string; isAnchor: boolean } {
	return (
		typeof binding === 'object' &&
		binding !== null &&
		(binding as { kind?: string }).kind === 'spill'
	)
}

function makeRangeArg(
	workbook: Workbook,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
	dateSystem: '1900' | '1904' = '1900',
	today?: Date,
): EvalArg {
	if (!isRangeInBounds(startRow, startCol, endRow, endCol)) {
		return { value: errorValue('#REF!') }
	}
	return makeMultiAreaArg([
		makeRangeArea(workbook, sheetIndex, startRow, startCol, endRow, endCol, {
			dateSystem,
			...(today ? { today } : {}),
		}),
	])
}

function makeExternalRangeArg(
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
	values: readonly (readonly CellValue[])[],
): EvalArg {
	const topLeft = values[0]?.[0] ?? EMPTY
	return makeMultiAreaArg([
		{
			ref: {
				kind: 'range',
				sheetIndex,
				row: startRow,
				col: startCol,
				endRow,
				endCol,
			},
			values,
			topLeft,
			valueAtOffset: (rowOffset, colOffset) => values[rowOffset]?.[colOffset] ?? EMPTY,
			forEachValue: (fn) => {
				for (const row of values) {
					for (const value of row) fn(value)
				}
			},
			forEachCellInRange: (fn) => {
				for (const row of values) {
					for (const value of row) fn(value)
				}
			},
		},
	])
}

function evaluateReferenceNode(node: FormulaNode, ctx: EvalContext): CellValue {
	const resolved = resolveReferenceNode(node, ctx)
	if (!resolved) return errorValue('#VALUE!')
	return referenceArgToValue(resolved)
}

function referenceArgToValue(arg: EvalArg): CellValue {
	if (arg.value.kind === 'error') return arg.value
	const areas = areasOf(arg)
	if (areas?.length) {
		if (areas.length !== 1) return errorValue('#VALUE!')
		const area = areas[0]
		if (!area) return errorValue('#VALUE!')
		return areaToValue(area)
	}
	if (arg.kind === 'range' && arg.values) return matrixToValue(arg.values)
	return arg.value
}

function implicitIntersect(arg: EvalArg, ctx: EvalContext): CellValue {
	const areas = areasOf(arg)
	if (arg.value.kind === 'error' && !areas?.length) return arg.value
	if (arg.value.kind === 'array') return topLeftScalar(arg.value)
	if (!areas?.length) return topLeftScalar(arg.value)
	if (areas.length !== 1) return errorValue('#VALUE!')
	const area = areas[0]
	if (!area) return errorValue('#VALUE!')
	const bounds = toAreaBounds(area.ref)
	const height = bounds.endRow - bounds.startRow + 1
	const width = bounds.endCol - bounds.startCol + 1
	if (height === 1 && width === 1) return area.values[0]?.[0] ?? EMPTY
	if (height === 1) {
		if (ctx.col < bounds.startCol || ctx.col > bounds.endCol) return errorValue('#VALUE!')
		return getCellValue(ctx.workbook, bounds.sheetIndex, bounds.startRow, ctx.col)
	}
	if (width === 1) {
		if (ctx.row < bounds.startRow || ctx.row > bounds.endRow) return errorValue('#VALUE!')
		return getCellValue(ctx.workbook, bounds.sheetIndex, ctx.row, bounds.startCol)
	}
	if (
		ctx.row >= bounds.startRow &&
		ctx.row <= bounds.endRow &&
		ctx.col >= bounds.startCol &&
		ctx.col <= bounds.endCol
	) {
		return getCellValue(ctx.workbook, bounds.sheetIndex, ctx.row, ctx.col)
	}
	return errorValue('#VALUE!')
}

function matrixToValue(values: readonly (readonly CellValue[])[]): CellValue {
	const rows = values.map((row) => row.map((value) => topLeftScalar(value)))
	if (rows.length === 0) return EMPTY
	if (rows.length === 1 && (rows[0]?.length ?? 0) === 1) return rows[0]?.[0] ?? EMPTY
	return arrayValue(rows)
}

function areaToValue(area: EvalArea): CellValue {
	return matrixToValue(area.values)
}

function areasOf(arg: EvalArg): readonly EvalArea[] | null {
	if (arg.areas?.length) return arg.areas
	if (!arg.ref) return null
	if (arg.ref.kind === 'cell') {
		return [
			{
				ref: {
					kind: 'range',
					sheetIndex: arg.ref.sheetIndex,
					row: arg.ref.row,
					col: arg.ref.col,
					endRow: arg.ref.row,
					endCol: arg.ref.col,
				},
				topLeft: arg.value,
				values: [[arg.value]],
				forEachValue: (fn) => fn(arg.value),
			},
		]
	}
	return [
		{
			ref: arg.ref,
			topLeft: arg.value,
			values: arg.values ?? [[arg.value]],
			...(arg.formulaAtOffset ? { formulaAtOffset: arg.formulaAtOffset } : {}),
			...(arg.rowHiddenAtOffset ? { rowHiddenAtOffset: arg.rowHiddenAtOffset } : {}),
			...(arg.rowFilteredAtOffset ? { rowFilteredAtOffset: arg.rowFilteredAtOffset } : {}),
			...(arg.forEachValue ? { forEachValue: arg.forEachValue } : {}),
		},
	]
}

interface ActiveFilterRange {
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly dateSystem: '1900' | '1904'
	readonly today: Date
	readonly columns: readonly PreparedFilterColumn[]
	readonly savedHiddenRows?: ReadonlySet<number>
}

type FilterRangeBounds = Pick<ActiveFilterRange, 'startRow' | 'startCol' | 'endRow'>

interface PreparedFilterColumn {
	readonly source: FilterColumn
	readonly acceptedValues?: ReadonlySet<string>
	readonly colorMatchedRows?: ReadonlySet<number>
	readonly dynamicAverage?: number
	readonly iconMatchedRows?: ReadonlySet<number>
	readonly top10Threshold?: number
}

function activeFilterRanges(
	workbook: Workbook,
	sheetIndex: number,
	dateSystem: '1900' | '1904',
	today: Date,
): readonly ActiveFilterRange[] {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return []
	const ranges: ActiveFilterRange[] = []
	const needsConditionalFormats =
		sheet.autoFilter?.columns.some((column) => column.kind === 'colorFilter') === true ||
		sheet.tables.some(
			(table) => table.autoFilter?.columns.some((column) => column.kind === 'colorFilter') === true,
		)
	const conditionalFormats =
		needsConditionalFormats && sheet.conditionalFormats.length > 0
			? evaluateConditionalFormats(sheet, workbook)
			: undefined
	if (sheet.autoFilter && sheet.autoFilter.columns.length > 0) {
		const range = parseFilterRange(sheet.autoFilter.ref)
		if (range)
			ranges.push({
				...range,
				dateSystem,
				today,
				columns: prepareFilterColumns(
					workbook,
					sheet,
					range,
					sheet.autoFilter.columns,
					conditionalFormats,
				),
				...savedHiddenRowsForFilterRange(sheet, range),
			})
	}
	for (const table of sheet.tables) {
		if (!table.autoFilter || table.autoFilter.columns.length === 0) continue
		const range = parseFilterRange(table.autoFilter.ref)
		if (range)
			ranges.push({
				...range,
				dateSystem,
				today,
				columns: prepareFilterColumns(
					workbook,
					sheet,
					range,
					table.autoFilter.columns,
					conditionalFormats,
				),
				...savedHiddenRowsForFilterRange(sheet, range),
			})
	}
	return ranges
}

function savedHiddenRowsForFilterRange(
	sheet: Sheet,
	range: FilterRangeBounds,
): { readonly savedHiddenRows?: ReadonlySet<number> } {
	const rows = new Set<number>()
	for (let row = range.startRow + 1; row <= range.endRow; row++) {
		if (sheet.rowDefs.get(row)?.hidden === true) rows.add(row)
	}
	return rows.size > 0 ? { savedHiddenRows: rows } : {}
}

function prepareFilterColumns(
	workbook: Workbook,
	sheet: Sheet,
	range: FilterRangeBounds,
	columns: readonly FilterColumn[],
	conditionalFormats?: ReturnType<typeof evaluateConditionalFormats>,
): readonly PreparedFilterColumn[] {
	return columns.map((source) => {
		const prepared: {
			source: FilterColumn
			acceptedValues?: ReadonlySet<string>
			colorMatchedRows?: ReadonlySet<number>
			dynamicAverage?: number
			iconMatchedRows?: ReadonlySet<number>
			top10Threshold?: number
		} = { source }
		if (source.kind === 'filters') {
			prepared.acceptedValues = new Set((source.values ?? []).map((value) => value.toLowerCase()))
		}
		if (source.kind === 'colorFilter') {
			const rows = computeColorFilterRows(workbook, sheet, range, source, conditionalFormats)
			if (rows) prepared.colorMatchedRows = rows
		}
		if (source.kind === 'iconFilter') {
			const rows = computeIconFilterRows(sheet, range, source)
			if (rows) prepared.iconMatchedRows = rows
		}
		if (source.kind === 'top10') {
			const threshold = computeTop10FilterThreshold(sheet, range, source)
			if (threshold !== undefined) prepared.top10Threshold = threshold
		}
		if (source.kind === 'dynamicFilter') {
			const average = computeDynamicFilterAverage(sheet, range, source)
			if (average !== undefined) prepared.dynamicAverage = average
		}
		return prepared
	})
}

function parseFilterRange(ref: string): FilterRangeBounds | null {
	try {
		const range = parseRange(ref)
		return { startRow: range.start.row, startCol: range.start.col, endRow: range.end.row }
	} catch {
		return null
	}
}

function rowFailsFilterCriteria(sheet: Sheet, row: number, range: ActiveFilterRange): boolean {
	for (const column of range.columns) {
		const cellValue = sheet.cells.readValue(row, range.startCol + column.source.colId)
		const matches = cellMatchesFilterColumn(cellValue, row, column, range.dateSystem, range.today)
		if (matches === false) return true
	}
	return false
}

function cellMatchesFilterColumn(
	value: CellValue,
	row: number,
	column: PreparedFilterColumn,
	dateSystem: '1900' | '1904',
	today: Date,
): boolean | null {
	const source = column.source
	if (source.kind === 'colorFilter') {
		return column.colorMatchedRows ? column.colorMatchedRows.has(row) : null
	}
	if (source.kind === 'iconFilter') {
		return column.iconMatchedRows ? column.iconMatchedRows.has(row) : null
	}
	if (source.kind === 'customFilters') return cellMatchesCustomFilters(value, source)
	if (source.kind === 'dynamicFilter') {
		return cellMatchesDynamicFilter(value, source, column.dynamicAverage, dateSystem, today)
	}
	if (source.kind === 'top10') return cellMatchesTop10Filter(value, source, column.top10Threshold)
	if (source.kind !== 'filters') return null
	const acceptsBlank = source.blank === true
	const dateGroupItems = source.dateGroupItems ?? []
	if ((column.acceptedValues?.size ?? 0) === 0 && !acceptsBlank && dateGroupItems.length === 0) {
		return null
	}
	if (isBlankFilterValue(value)) return acceptsBlank
	const text = coerceCellValueToString(value).toLowerCase()
	return (
		column.acceptedValues?.has(text) === true ||
		cellMatchesDateGroupItems(value, dateGroupItems, dateSystem)
	)
}

function computeTop10FilterThreshold(
	sheet: Sheet,
	range: FilterRangeBounds,
	column: FilterColumn,
): number | undefined {
	if (column.filterVal !== undefined && Number.isFinite(column.filterVal)) {
		return column.filterVal
	}
	const values: number[] = []
	const col = range.startCol + column.colId
	for (let row = range.startRow + 1; row <= range.endRow; row++) {
		const value = filterTop10ComparableNumber(sheet.cells.readValue(row, col))
		if (value !== null) values.push(value)
	}
	if (values.length === 0) return undefined
	const rawCount = column.val ?? 10
	const count =
		column.percent === true ? Math.ceil((values.length * rawCount) / 100) : Math.floor(rawCount)
	if (!Number.isFinite(count) || count <= 0) return undefined
	const selectedCount = Math.min(values.length, Math.max(1, count))
	values.sort((a, b) => a - b)
	return column.top === false ? values[selectedCount - 1] : values[values.length - selectedCount]
}

function cellMatchesTop10Filter(
	value: CellValue,
	column: FilterColumn,
	threshold: number | undefined,
): boolean | null {
	if (threshold === undefined) return null
	const comparable = filterTop10ComparableNumber(value)
	if (comparable === null) return false
	return column.top === false ? comparable <= threshold : comparable >= threshold
}

function computeColorFilterRows(
	workbook: Workbook,
	sheet: Sheet,
	range: FilterRangeBounds,
	column: FilterColumn,
	conditionalFormats: ReturnType<typeof evaluateConditionalFormats> | undefined,
): ReadonlySet<number> | undefined {
	if (column.dxfId === undefined) return undefined
	const target = workbook.differentialStyles[column.dxfId]
	if (!target) return undefined
	const cellColor = column.cellColor !== false
	if (!filterColorStyleHasTarget(target, cellColor)) return undefined
	const col = range.startCol + column.colId
	const rows = new Set<number>()
	for (let row = range.startRow + 1; row <= range.endRow; row++) {
		const directStyle = workbook.styles.get(sheet.cells.readStyleId(row, col) ?? DEFAULT_STYLE_ID)
		if (filterColorStyleMatches(directStyle, target, cellColor)) {
			rows.add(row)
			continue
		}
		const a1 = `${indexToColumn(col)}${row + 1}`
		for (const result of conditionalFormats?.get(a1) ?? []) {
			if (filterColorStyleMatches(result.format, target, cellColor)) {
				rows.add(row)
				break
			}
		}
	}
	return rows
}

function filterColorStyleHasTarget(style: CellStyle, cellColor: boolean): boolean {
	return cellColor ? fillColors(style.fill).length > 0 : style.font?.color !== undefined
}

function filterColorStyleMatches(
	candidate: CellStyle | undefined,
	target: CellStyle,
	cellColor: boolean,
): boolean {
	if (cellColor) return fillColorsOverlap(candidate?.fill, target.fill)
	return colorsEqual(candidate?.font?.color, target.font?.color)
}

function fillColorsOverlap(
	candidate: CellStyle['fill'] | undefined,
	target: CellStyle['fill'] | undefined,
): boolean {
	const targetColors = fillColors(target)
	if (targetColors.length === 0) return false
	const candidateColors = fillColors(candidate)
	return targetColors.some((targetColor) =>
		candidateColors.some((candidateColor) => colorsEqual(candidateColor, targetColor)),
	)
}

function fillColors(fill: CellStyle['fill'] | undefined): readonly Color[] {
	const colors: Color[] = []
	if (fill?.fgColor) colors.push(fill.fgColor)
	if (fill?.bgColor) colors.push(fill.bgColor)
	return colors
}

function colorsEqual(left: Color | undefined, right: Color | undefined): boolean {
	return left !== undefined && right !== undefined && colorKey(left) === colorKey(right)
}

function colorKey(color: Color): string {
	switch (color.kind) {
		case 'rgb':
			return `rgb:${normalizeRgb(color.rgb)}`
		case 'theme':
			return `theme:${color.theme}:${color.tint ?? 0}`
		case 'indexed':
			return `indexed:${color.index}`
		case 'auto':
			return 'auto'
	}
}

function normalizeRgb(value: string): string {
	const normalized = value.trim().toUpperCase()
	return normalized.length === 6 ? `FF${normalized}` : normalized
}

function computeDynamicFilterAverage(
	sheet: Sheet,
	range: FilterRangeBounds,
	column: FilterColumn,
): number | undefined {
	const type = column.dynamicFilterType
	if (type !== 'aboveAverage' && type !== 'belowAverage') return undefined
	if (column.dynamicFilterVal !== undefined && Number.isFinite(column.dynamicFilterVal)) {
		return column.dynamicFilterVal
	}
	let sum = 0
	let count = 0
	const col = range.startCol + column.colId
	for (let row = range.startRow + 1; row <= range.endRow; row++) {
		const value = filterDynamicComparableNumber(sheet.cells.readValue(row, col))
		if (value === null) continue
		sum += value
		count += 1
	}
	return count > 0 ? sum / count : undefined
}

function cellMatchesDynamicFilter(
	value: CellValue,
	column: FilterColumn,
	average: number | undefined,
	dateSystem: '1900' | '1904',
	today: Date,
): boolean | null {
	const type = column.dynamicFilterType
	if (!type) return null
	const comparable = filterDynamicComparableNumber(value)
	if (type === 'aboveAverage' || type === 'belowAverage') {
		if (average === undefined || comparable === null) return average === undefined ? null : false
		return type === 'aboveAverage' ? comparable > average : comparable < average
	}
	const range = dynamicFilterRange(column, dateSystem, today)
	if (range) {
		if (comparable === null) return false
		return comparable >= range.min && comparable < range.max
	}
	const parts = value.kind === 'date' ? serialToDateTimeParts(value.serial, dateSystem) : null
	if (!parts) return null
	const month = dynamicFilterMonth(type)
	if (month !== null) return parts.month === month
	const quarter = dynamicFilterQuarter(type)
	if (quarter !== null) return Math.ceil(parts.month / 3) === quarter
	return null
}

function dynamicFilterRange(
	column: FilterColumn,
	dateSystem: '1900' | '1904',
	today: Date,
): { readonly min: number; readonly max: number } | null {
	const min = dynamicFilterBound(column.dynamicFilterVal, column.dynamicFilterValIso, dateSystem)
	const max = dynamicFilterBound(
		column.dynamicFilterMaxVal,
		column.dynamicFilterMaxValIso,
		dateSystem,
	)
	if (min !== undefined && max !== undefined && max > min) return { min, max }
	return dynamicFilterRelativeRange(column.dynamicFilterType ?? '', today, dateSystem)
}

function dynamicFilterBound(
	numeric: number | undefined,
	iso: string | undefined,
	dateSystem: '1900' | '1904',
): number | undefined {
	if (numeric !== undefined && Number.isFinite(numeric)) return numeric
	return iso ? parseIsoDateTimeSerial(iso, dateSystem) : undefined
}

function dynamicFilterRelativeRange(
	type: string,
	today: Date,
	dateSystem: '1900' | '1904',
): { readonly min: number; readonly max: number } | null {
	const day = startOfLocalDay(today)
	switch (type) {
		case 'yesterday':
			return serialDateRange(addDays(day, -1), day, dateSystem)
		case 'today':
			return serialDateRange(day, addDays(day, 1), dateSystem)
		case 'tomorrow':
			return serialDateRange(addDays(day, 1), addDays(day, 2), dateSystem)
		case 'lastWeek': {
			const start = addDays(startOfWeek(day), -7)
			return serialDateRange(start, addDays(start, 7), dateSystem)
		}
		case 'thisWeek': {
			const start = startOfWeek(day)
			return serialDateRange(start, addDays(start, 7), dateSystem)
		}
		case 'nextWeek': {
			const start = addDays(startOfWeek(day), 7)
			return serialDateRange(start, addDays(start, 7), dateSystem)
		}
		case 'lastMonth': {
			const start = addMonths(startOfMonth(day), -1)
			return serialDateRange(start, addMonths(start, 1), dateSystem)
		}
		case 'thisMonth': {
			const start = startOfMonth(day)
			return serialDateRange(start, addMonths(start, 1), dateSystem)
		}
		case 'nextMonth': {
			const start = addMonths(startOfMonth(day), 1)
			return serialDateRange(start, addMonths(start, 1), dateSystem)
		}
		case 'lastQuarter': {
			const start = addMonths(startOfQuarter(day), -3)
			return serialDateRange(start, addMonths(start, 3), dateSystem)
		}
		case 'thisQuarter': {
			const start = startOfQuarter(day)
			return serialDateRange(start, addMonths(start, 3), dateSystem)
		}
		case 'nextQuarter': {
			const start = addMonths(startOfQuarter(day), 3)
			return serialDateRange(start, addMonths(start, 3), dateSystem)
		}
		case 'lastYear':
			return serialDateRange(
				new Date(day.getFullYear() - 1, 0, 1),
				new Date(day.getFullYear(), 0, 1),
				dateSystem,
			)
		case 'thisYear':
			return serialDateRange(
				new Date(day.getFullYear(), 0, 1),
				new Date(day.getFullYear() + 1, 0, 1),
				dateSystem,
			)
		case 'nextYear':
			return serialDateRange(
				new Date(day.getFullYear() + 1, 0, 1),
				new Date(day.getFullYear() + 2, 0, 1),
				dateSystem,
			)
		case 'yearToDate':
			return serialDateRange(new Date(day.getFullYear(), 0, 1), addDays(day, 1), dateSystem)
		default:
			return null
	}
}

function startOfLocalDay(date: Date): Date {
	return new Date(date.getFullYear(), date.getMonth(), date.getDate())
}

function startOfWeek(date: Date): Date {
	return addDays(date, -date.getDay())
}

function startOfMonth(date: Date): Date {
	return new Date(date.getFullYear(), date.getMonth(), 1)
}

function startOfQuarter(date: Date): Date {
	return new Date(date.getFullYear(), Math.floor(date.getMonth() / 3) * 3, 1)
}

function addDays(date: Date, days: number): Date {
	return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days)
}

function addMonths(date: Date, months: number): Date {
	return new Date(date.getFullYear(), date.getMonth() + months, 1)
}

function serialDateRange(
	start: Date,
	end: Date,
	dateSystem: '1900' | '1904',
): { readonly min: number; readonly max: number } {
	return {
		min: dateToSerial(start.getFullYear(), start.getMonth() + 1, start.getDate(), dateSystem),
		max: dateToSerial(end.getFullYear(), end.getMonth() + 1, end.getDate(), dateSystem),
	}
}

function parseIsoDateTimeSerial(value: string, dateSystem: '1900' | '1904'): number | undefined {
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2})(?::(\d{2})(?::(\d{2}(?:\.\d+)?))?)?)?/.exec(
		value.trim(),
	)
	if (!match) return undefined
	const year = Number(match[1])
	const month = Number(match[2])
	const day = Number(match[3])
	const hour = match[4] ? Number(match[4]) : 0
	const minute = match[5] ? Number(match[5]) : 0
	const second = match[6] ? Number(match[6]) : 0
	if (
		!Number.isInteger(year) ||
		!Number.isInteger(month) ||
		!Number.isInteger(day) ||
		month < 1 ||
		month > 12 ||
		day < 1 ||
		day > 31 ||
		hour < 0 ||
		hour >= 24 ||
		minute < 0 ||
		minute >= 60 ||
		second < 0 ||
		second >= 60
	) {
		return undefined
	}
	return (
		dateToSerial(year, month, day, dateSystem) +
		(hour * 3600 + minute * 60 + second) / SECONDS_PER_DAY
	)
}

function dynamicFilterMonth(type: string): number | null {
	const match = /^M([1-9]|1[0-2])$/.exec(type)
	return match ? Number(match[1]) : null
}

function dynamicFilterQuarter(type: string): number | null {
	const match = /^Q([1-4])$/.exec(type)
	return match ? Number(match[1]) : null
}

function cellMatchesDateGroupItems(
	value: CellValue,
	items: readonly FilterDateGroupItem[],
	dateSystem: '1900' | '1904',
): boolean {
	if (items.length === 0 || value.kind !== 'date') return false
	const dateParts = serialToDateTimeParts(value.serial, dateSystem)
	if (!dateParts) return false
	for (const item of items) {
		if (dateGroupItemMatches(dateParts, item)) return true
	}
	return false
}

const SECONDS_PER_DAY = 86_400

function serialToDateTimeParts(
	serial: number,
	dateSystem: '1900' | '1904',
): {
	readonly year: number
	readonly month: number
	readonly day: number
	readonly hour: number
	readonly minute: number
	readonly second: number
} | null {
	const wholeDays = Math.floor(serial)
	const dateParts = serialToDate(wholeDays, dateSystem)
	if (!dateParts) return null
	const fractionalDay = Math.max(0, serial - wholeDays)
	const wholeSeconds = Math.min(SECONDS_PER_DAY - 1, Math.round(fractionalDay * SECONDS_PER_DAY))
	return {
		...dateParts,
		hour: Math.floor(wholeSeconds / 3600),
		minute: Math.floor((wholeSeconds % 3600) / 60),
		second: wholeSeconds % 60,
	}
}

function dateGroupItemMatches(
	dateParts: {
		readonly year: number
		readonly month: number
		readonly day: number
		readonly hour: number
		readonly minute: number
		readonly second: number
	},
	item: FilterDateGroupItem,
): boolean {
	return (
		(item.year === undefined || item.year === dateParts.year) &&
		(item.month === undefined || item.month === dateParts.month) &&
		(item.day === undefined || item.day === dateParts.day) &&
		(item.hour === undefined || item.hour === dateParts.hour) &&
		(item.minute === undefined || item.minute === dateParts.minute) &&
		(item.second === undefined || item.second === dateParts.second)
	)
}

function cellMatchesCustomFilters(value: CellValue, column: FilterColumn): boolean | null {
	const filters = column.customFilters ?? []
	if (filters.length === 0) return null
	return column.and === true
		? filters.every((filter) => cellMatchesCustomFilter(value, filter))
		: filters.some((filter) => cellMatchesCustomFilter(value, filter))
}

function cellMatchesCustomFilter(value: CellValue, filter: CustomFilter): boolean {
	const operator = filter.operator ?? 'equal'
	if (operator === 'equal' || operator === 'notEqual') {
		const matches = hasWildcardFilterPattern(filter.val)
			? wildcardFilterMatches(value, filter.val)
			: compareFilterValues(value, filter.val) === 0
		return operator === 'notEqual' ? !matches : matches
	}
	const comparison = compareFilterValues(value, filter.val)
	if (comparison === null) return false
	switch (operator) {
		case 'greaterThan':
			return comparison > 0
		case 'greaterThanOrEqual':
			return comparison >= 0
		case 'lessThan':
			return comparison < 0
		case 'lessThanOrEqual':
			return comparison <= 0
		default:
			return false
	}
}

function compareFilterValues(value: CellValue, criterion: string): number | null {
	const valueNumber = filterComparableNumber(value)
	const criterionNumber = filterCriterionNumber(criterion)
	if (valueNumber !== null && criterionNumber !== null)
		return Math.sign(valueNumber - criterionNumber)
	const left = coerceCellValueToString(value).toLowerCase()
	const right = criterion.toLowerCase()
	if (left === right) return 0
	return left < right ? -1 : 1
}

function filterComparableNumber(value: CellValue): number | null {
	if (value.kind === 'number') return value.value
	if (value.kind === 'date') return value.serial
	if (value.kind === 'boolean') return value.value ? 1 : 0
	if (value.kind !== 'string') return null
	return filterCriterionNumber(value.value)
}

function filterTop10ComparableNumber(value: CellValue): number | null {
	if (value.kind === 'number') return Number.isFinite(value.value) ? value.value : null
	if (value.kind === 'date') return Number.isFinite(value.serial) ? value.serial : null
	return null
}

function filterDynamicComparableNumber(value: CellValue): number | null {
	return filterTop10ComparableNumber(value)
}

function filterCriterionNumber(value: string): number | null {
	const trimmed = value.trim()
	if (trimmed === '') return null
	const parsed = Number(trimmed)
	return Number.isFinite(parsed) ? parsed : null
}

function hasWildcardFilterPattern(pattern: string): boolean {
	return pattern.includes('*') || pattern.includes('?')
}

function wildcardFilterMatches(value: CellValue, pattern: string): boolean {
	return wildcardFilterRegex(pattern).test(coerceCellValueToString(value))
}

function wildcardFilterRegex(pattern: string): RegExp {
	let source = '^'
	for (let index = 0; index < pattern.length; index++) {
		const char = pattern[index] ?? ''
		if (char === '*') {
			source += '.*'
		} else if (char === '?') {
			source += '.'
		} else if (char === '~' && index + 1 < pattern.length) {
			index++
			source += escapeRegexLiteral(pattern[index] ?? '')
		} else {
			source += escapeRegexLiteral(char)
		}
	}
	return new RegExp(`${source}$`, 'i')
}

function escapeRegexLiteral(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function isBlankFilterValue(value: CellValue): boolean {
	return value.kind === 'empty' || (value.kind === 'string' && value.value === '')
}

function makeRangeArea(
	workbook: Workbook,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
	options: {
		dateSystem?: '1900' | '1904'
		today?: Date
		materializedStartRow?: number
		materializedStartCol?: number
		materializedEndRow?: number
		materializedEndCol?: number
	} = {},
): EvalArea {
	const sheet = workbook.sheets[sheetIndex]
	let cachedValues: readonly (readonly CellValue[])[] | undefined
	const dateSystem = options.dateSystem ?? '1900'
	const today = options.today ?? startOfLocalDay(new Date())
	const materializedStartRow = options.materializedStartRow ?? startRow
	const materializedStartCol = options.materializedStartCol ?? startCol
	const materializedEndRow = options.materializedEndRow ?? endRow
	const materializedEndCol = options.materializedEndCol ?? endCol
	const filteredRanges = sheet ? activeFilterRanges(workbook, sheetIndex, dateSystem, today) : []
	return {
		ref: {
			kind: 'range',
			sheetIndex,
			row: startRow,
			col: startCol,
			endRow,
			endCol,
		},
		get topLeft() {
			return sheet?.cells.readValue(materializedStartRow, materializedStartCol) ?? EMPTY
		},
		valueAtOffset: (rowOffset: number, colOffset: number) =>
			getCellValue(workbook, sheetIndex, startRow + rowOffset, startCol + colOffset),
		formulaAtOffset: (rowOffset: number, colOffset: number) =>
			sheet?.cells.readFormula(materializedStartRow + rowOffset, materializedStartCol + colOffset),
		rowHiddenAtOffset: (rowOffset: number) =>
			sheet?.rowDefs.get(materializedStartRow + rowOffset)?.hidden === true,
		rowFilteredAtOffset: (rowOffset: number) => {
			const row = materializedStartRow + rowOffset
			return filteredRanges.some(
				(range) =>
					row > range.startRow &&
					row <= range.endRow &&
					(range.savedHiddenRows
						? range.savedHiddenRows.has(row)
						: sheet?.rowDefs.get(row)?.hidden === true ||
							(sheet ? rowFailsFilterCriteria(sheet, row, range) : false)),
			)
		},
		get values() {
			if (!cachedValues) {
				cachedValues = getRangeValues(
					workbook,
					sheetIndex,
					materializedStartRow,
					materializedStartCol,
					materializedEndRow,
					materializedEndCol,
				)
			}
			return cachedValues
		},
		...(sheet
			? {
					forEachValue: (fn: (value: CellValue) => void) => {
						sheet.cells.forEachValueInRange(
							materializedStartRow,
							materializedStartCol,
							materializedEndRow,
							materializedEndCol,
							(value) => fn(value),
						)
					},
					forEachCellInRange: (fn: (value: CellValue) => void) => {
						for (let r = materializedStartRow; r <= materializedEndRow; r++) {
							for (let c = materializedStartCol; c <= materializedEndCol; c++) {
								fn(sheet.cells.readValue(r, c))
							}
						}
					},
				}
			: {}),
	}
}

function makeMultiAreaArg(areas: readonly EvalArea[]): EvalArg {
	if (areas.length === 0) return { value: errorValue('#NULL!') }
	const firstArea = areas[0]
	if (!firstArea) return { value: errorValue('#NULL!') }
	return {
		get value() {
			return firstArea.topLeft ?? firstArea.values[0]?.[0] ?? EMPTY
		},
		kind: 'range',
		...(areas.length === 1
			? {
					get values() {
						return firstArea.values
					},
					ref: firstArea.ref,
					shapeRows: (firstArea.ref.endRow ?? firstArea.ref.row) - firstArea.ref.row + 1,
					shapeCols: (firstArea.ref.endCol ?? firstArea.ref.col) - firstArea.ref.col + 1,
					...(firstArea.valueAtOffset ? { valueAtOffset: firstArea.valueAtOffset } : {}),
					...(firstArea.formulaAtOffset ? { formulaAtOffset: firstArea.formulaAtOffset } : {}),
					...(firstArea.rowHiddenAtOffset
						? { rowHiddenAtOffset: firstArea.rowHiddenAtOffset }
						: {}),
					...(firstArea.rowFilteredAtOffset
						? { rowFilteredAtOffset: firstArea.rowFilteredAtOffset }
						: {}),
				}
			: {}),
		areas,
		forEachValue: (fn) => {
			for (const area of areas) {
				if (area.forEachValue) area.forEachValue(fn)
				else {
					for (const row of area.values) {
						for (const value of row) fn(value)
					}
				}
			}
		},
		forEachCellInRange: (fn) => {
			for (const area of areas) {
				if (area.forEachCellInRange) area.forEachCellInRange(fn)
				else if (area.forEachValue) area.forEachValue(fn)
				else {
					for (const row of area.values) {
						for (const value of row) fn(value)
					}
				}
			}
		},
	}
}

function toAreaBounds(ref: EvalArea['ref']): {
	sheetIndex: number
	startRow: number
	startCol: number
	endRow: number
	endCol: number
} {
	return {
		sheetIndex: ref.sheetIndex,
		startRow: ref.row,
		startCol: ref.col,
		endRow: ref.endRow ?? ref.row,
		endCol: ref.endCol ?? ref.col,
	}
}

function resolveSheetSpanIndices(
	workbook: Workbook,
	startSheet: string,
	endSheet: string,
): number[] | null {
	const startLower = startSheet.toLowerCase()
	const endLower = endSheet.toLowerCase()
	let start = -1
	let end = -1
	for (let i = 0; i < workbook.sheets.length; i++) {
		const sheet = workbook.sheets[i]
		if (!sheet) continue
		const nameLower = sheet.name.toLowerCase()
		if (nameLower === startLower) start = i
		if (nameLower === endLower) end = i
		if (start !== -1 && end !== -1) break
	}
	if (start === -1 || end === -1 || start > end) return null
	const result: number[] = new Array(end - start + 1)
	for (let i = start; i <= end; i++) result[i - start] = i
	return result
}

function applySheetToReferenceNode(
	node: FormulaNode,
	sheetIndex: number,
	ctx: EvalContext,
): FormulaNode {
	const sheet = ctx.workbook.sheets[sheetIndex]
	if (!sheet) return node
	switch (node.type) {
		case 'cellRef':
			return { type: 'cellRef', ref: node.ref, sheet: sheet.name }
		case 'rangeRef':
			return { type: 'rangeRef', start: node.start, end: node.end, sheet: sheet.name }
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: applySheetToReferenceNode(node.start, sheetIndex, ctx),
				end: applySheetToReferenceNode(node.end, sheetIndex, ctx),
			}
		case 'wholeRowRange':
			return {
				type: 'wholeRowRange',
				startRow: node.startRow,
				endRow: node.endRow,
				sheet: sheet.name,
			}
		case 'wholeColumnRange':
			return {
				type: 'wholeColumnRange',
				startCol: node.startCol,
				endCol: node.endCol,
				...(node.startColAbsolute ? { startColAbsolute: true } : {}),
				...(node.endColAbsolute ? { endColAbsolute: true } : {}),
				sheet: sheet.name,
			}
		case 'name':
			return { type: 'name', name: node.name, sheet: sheet.name }
		default:
			return node
	}
}

function makeWholeRowArg(
	workbook: Workbook,
	sheetIndex: number,
	startRow: number,
	endRow: number,
	dateSystem: '1900' | '1904' = '1900',
	today?: Date,
): EvalArg {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return { value: errorValue('#REF!') }
	const used = sheet.cells.usedRange()
	const materializedStartCol = used?.start.col ?? 0
	const materializedEndCol = used?.end.col ?? materializedStartCol
	return makeMultiAreaArg([
		makeRangeArea(workbook, sheetIndex, startRow, 0, endRow, EXCEL_MAX_COLS - 1, {
			dateSystem,
			...(today ? { today } : {}),
			materializedStartRow: startRow,
			materializedStartCol,
			materializedEndRow: endRow,
			materializedEndCol,
		}),
	])
}

function makeWholeColumnArg(
	workbook: Workbook,
	sheetIndex: number,
	startCol: number,
	endCol: number,
	dateSystem: '1900' | '1904' = '1900',
	today?: Date,
): EvalArg {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return { value: errorValue('#REF!') }
	const used = sheet.cells.usedRange()
	const materializedStartRow = used?.start.row ?? 0
	const materializedEndRow = used?.end.row ?? materializedStartRow
	return makeMultiAreaArg([
		makeRangeArea(workbook, sheetIndex, 0, startCol, EXCEL_MAX_ROWS - 1, endCol, {
			dateSystem,
			...(today ? { today } : {}),
			materializedStartRow,
			materializedStartCol: startCol,
			materializedEndRow,
			materializedEndCol: endCol,
		}),
	])
}

function extractLambda(node: FormulaNode, ctx: EvalContext): LambdaInfo | null {
	if (node.type === 'function' && node.name.toUpperCase() === 'LAMBDA') {
		if (node.args.length < 2) return null
		const params: string[] = []
		for (let i = 0; i < node.args.length - 1; i++) {
			const p = node.args[i]
			if (!p || p.type !== 'name') return null
			params.push(p.name.toLowerCase())
		}
		return { params, body: node.args[node.args.length - 1] as FormulaNode, ctx }
	}
	if (node.type === 'name') {
		if (!node.sheet && ctx.letBindings?.has(node.name.toLowerCase())) {
			const bound = ctx.letBindings.get(node.name.toLowerCase()) as LetBinding
			return isLambdaBinding(bound) ? bound : null
		}
		const resolved = resolveDefinedName(node.name, node.sheet, ctx)
		if (!resolved) return null
		return extractLambda(resolved.ast, resolved.ctx)
	}
	return null
}

function extractLambdaFromName(name: string, ctx: EvalContext): LambdaInfo | null {
	const letBound = ctx.letBindings?.get(name.toLowerCase())
	if (letBound && isLambdaBinding(letBound)) return letBound
	const resolved = resolveDefinedName(name, undefined, ctx)
	if (!resolved) return null
	return extractLambda(resolved.ast, resolved.ctx)
}

function invokeLambda(lambda: LambdaInfo, args: readonly CellValue[]): CellValue {
	if (lambda.params.length !== args.length) return errorValue('#VALUE!')
	const bindings = new Map<string, LetBinding>(lambda.ctx.letBindings)
	for (let i = 0; i < lambda.params.length; i++) {
		bindings.set(lambda.params[i] as string, args[i] as CellValue)
	}
	return evaluate(lambda.body, { ...lambda.ctx, letBindings: bindings })
}

function evalCall(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const [callee, ...argExprs] = argNodes
	if (!callee) return errorValue('#VALUE!')
	const lambda = extractLambda(callee, ctx)
	if (!lambda) return errorValue('#VALUE!')
	return invokeLambda(
		lambda,
		argExprs.map((arg) => evaluate(arg, ctx)),
	)
}

function evalMap(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length < 2) return errorValue('#VALUE!')
	const lambdaNode = argNodes[argNodes.length - 1]
	if (!lambdaNode) return errorValue('#VALUE!')
	const lambda = extractLambda(lambdaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	const arrayNodes = argNodes.slice(0, -1)
	if (lambda.params.length !== arrayNodes.length) return errorValue('#VALUE!')
	const ranges = arrayNodes.map((node) => getRange(resolveArg(node, ctx)))
	const shape = rangeShapeOf(ranges[0])
	if (!shape) return errorValue('#VALUE!')
	for (let i = 1; i < ranges.length; i++) {
		const candidateShape = rangeShapeOf(ranges[i])
		if (
			!candidateShape ||
			candidateShape.rows !== shape.rows ||
			candidateShape.cols !== shape.cols
		) {
			return errorValue('#VALUE!')
		}
	}
	const rows: ScalarCellValue[][] = []
	for (let rowIndex = 0; rowIndex < shape.rows; rowIndex++) {
		const resultRow: ScalarCellValue[] = []
		for (let colIndex = 0; colIndex < shape.cols; colIndex++) {
			const args = ranges.map((range) => range[rowIndex]?.[colIndex] ?? EMPTY)
			const result = invokeLambda(lambda, args)
			if (result.kind === 'error') return result
			resultRow.push(topLeftScalar(result))
		}
		rows.push(resultRow)
	}
	if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0] ?? EMPTY
	return arrayValue(rows)
}

function rangeShapeOf(
	range: readonly (readonly CellValue[])[] | undefined,
): { rows: number; cols: number } | null {
	const rows = range?.length ?? 0
	const cols = range?.[0]?.length ?? 0
	if (rows === 0 || cols === 0) return null
	for (const row of range ?? []) {
		if (row.length !== cols) return null
	}
	return { rows, cols }
}

function evalReduce(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length !== 3) return errorValue('#VALUE!')
	let accumulator = evaluate(argNodes[0] as FormulaNode, ctx)
	const arrayArg = resolveArg(argNodes[1] as FormulaNode, ctx)
	const lambda = extractLambda(argNodes[2] as FormulaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	if (lambda.params.length !== 2) return errorValue('#VALUE!')
	const range = getRange(arrayArg)
	for (const row of range) {
		for (const cell of row) {
			accumulator = invokeLambda(lambda, [accumulator, cell])
			if (accumulator.kind === 'error') return accumulator
		}
	}
	return accumulator
}

function evalScan(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length !== 3) return errorValue('#VALUE!')
	let accumulator = evaluate(argNodes[0] as FormulaNode, ctx)
	const arrayArg = resolveArg(argNodes[1] as FormulaNode, ctx)
	const lambda = extractLambda(argNodes[2] as FormulaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	if (lambda.params.length !== 2) return errorValue('#VALUE!')
	const range = getRange(arrayArg)
	const rows: ScalarCellValue[][] = []
	for (const row of range) {
		const resultRow: ScalarCellValue[] = []
		for (const cell of row) {
			accumulator = invokeLambda(lambda, [accumulator, cell])
			if (accumulator.kind === 'error') return accumulator
			resultRow.push(topLeftScalar(accumulator))
		}
		rows.push(resultRow)
	}
	if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0] ?? EMPTY
	return arrayValue(rows)
}

function evalByRow(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length !== 2) return errorValue('#VALUE!')
	const arrayArg = resolveArg(argNodes[0] as FormulaNode, ctx)
	const lambda = extractLambda(argNodes[1] as FormulaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	if (lambda.params.length !== 1) return errorValue('#VALUE!')
	const range = getRange(arrayArg)
	const rows: ScalarCellValue[][] = []
	for (const row of range) {
		const rowArray = arrayValue([row.map((c) => topLeftScalar(c))])
		const result = invokeLambda(lambda, [rowArray])
		if (result.kind === 'error') return result
		if (result.kind === 'array') return errorValue('#CALC!')
		rows.push([topLeftScalar(result)])
	}
	if (rows.length === 1) return rows[0]?.[0] ?? EMPTY
	return arrayValue(rows)
}

function evalByCol(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length !== 2) return errorValue('#VALUE!')
	const arrayArg = resolveArg(argNodes[0] as FormulaNode, ctx)
	const lambda = extractLambda(argNodes[1] as FormulaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	if (lambda.params.length !== 1) return errorValue('#VALUE!')
	const range = getRange(arrayArg)
	const colCount = range.reduce((max, row) => Math.max(max, row.length), 0)
	const results: ScalarCellValue[] = []
	for (let col = 0; col < colCount; col++) {
		const colData: ScalarCellValue[][] = range.map((row) => [topLeftScalar(row[col] ?? EMPTY)])
		const colArray = arrayValue(colData)
		const result = invokeLambda(lambda, [colArray])
		if (result.kind === 'error') return result
		if (result.kind === 'array') return errorValue('#CALC!')
		results.push(topLeftScalar(result))
	}
	if (results.length === 1) return results[0] ?? EMPTY
	return arrayValue([results])
}

function evalMakeArray(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length !== 3) return errorValue('#VALUE!')
	const rowCountVal = evaluate(argNodes[0] as FormulaNode, ctx)
	if (rowCountVal.kind === 'error') return rowCountVal
	const rowCount = coerceToNumber(rowCountVal)
	if (rowCount === null || rowCount < 1) return errorValue('#VALUE!')
	const colCountVal = evaluate(argNodes[1] as FormulaNode, ctx)
	if (colCountVal.kind === 'error') return colCountVal
	const colCount = coerceToNumber(colCountVal)
	if (colCount === null || colCount < 1) return errorValue('#VALUE!')
	const lambda = extractLambda(argNodes[2] as FormulaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	if (lambda.params.length !== 2) return errorValue('#VALUE!')
	const r = Math.trunc(rowCount)
	const c = Math.trunc(colCount)
	const rows: ScalarCellValue[][] = []
	for (let row = 0; row < r; row++) {
		const resultRow: ScalarCellValue[] = []
		for (let col = 0; col < c; col++) {
			const result = invokeLambda(lambda, [numberValue(row + 1), numberValue(col + 1)])
			if (result.kind === 'error') return result
			resultRow.push(topLeftScalar(result))
		}
		rows.push(resultRow)
	}
	if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0] ?? EMPTY
	return arrayValue(rows)
}

function evalIsFormula(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const arg = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	if (!arg.ref) {
		if (arg.value.kind === 'error') return arg.value
		return booleanValue(false)
	}
	return booleanValue(arg.formulaAtOffset?.(0, 0) != null)
}

function evalSheet(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length === 0) return numberValue(ctx.sheetIndex + 1)
	const arg = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	const sheetIndex = firstSheetIndexOf(arg)
	if (sheetIndex !== null) return numberValue(sheetIndex + 1)
	const name = coerceToString(arg.value)
	for (let i = 0; i < ctx.workbook.sheets.length; i++) {
		if (ctx.workbook.sheets[i]?.name.toLowerCase() === name.toLowerCase()) return numberValue(i + 1)
	}
	return errorValue('#N/A')
}

function evalSheets(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length === 0) return numberValue(ctx.workbook.sheets.length)
	const arg = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	const sheetCount = countSheetsOf(arg)
	if (sheetCount !== null) return numberValue(sheetCount)
	const name = coerceToString(arg.value)
	for (const sheet of ctx.workbook.sheets) {
		if (sheet?.name.toLowerCase() === name.toLowerCase()) return numberValue(1)
	}
	return errorValue('#N/A')
}

function evalCellInfo(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const infoTypeValue = evaluate(argNodes[0] ?? { type: 'missing' }, ctx)
	if (infoTypeValue.kind === 'error') return infoTypeValue
	const infoType = coerceToString(infoTypeValue).toLowerCase()
	const arg = resolveArg(argNodes[1] ?? { type: 'missing' }, ctx)
	const targetRef = arg.ref ?? {
		kind: 'cell' as const,
		sheetIndex: ctx.sheetIndex,
		row: ctx.row,
		col: ctx.col,
	}
	const sheet = ctx.workbook.sheets[targetRef.sheetIndex]
	const cellStyle = sheet
		? ctx.workbook.styles.get(
				sheet.cells.readStyleId(targetRef.row, targetRef.col) ?? DEFAULT_STYLE_ID,
			)
		: undefined

	switch (infoType) {
		case 'address':
			return stringValue(`$${indexToColumn(targetRef.col)}$${targetRef.row + 1}`)
		case 'col':
			return numberValue(targetRef.col + 1)
		case 'row':
			return numberValue(targetRef.row + 1)
		case 'contents':
			return topLeftScalar(arg.value)
		case 'type': {
			const value = topLeftScalar(arg.value)
			if (value.kind === 'empty') return stringValue('b')
			if (value.kind === 'string' || value.kind === 'richText') return stringValue('l')
			return stringValue('v')
		}
		case 'width':
			return numberValue(sheet?.colWidths.get(targetRef.col) ?? 8)
		case 'prefix': {
			const value = topLeftScalar(arg.value)
			if (value.kind !== 'string' && value.kind !== 'richText') return stringValue('')
			const h = cellStyle?.alignment?.horizontal
			if (h === 'left') return stringValue("'")
			if (h === 'center') return stringValue('^')
			if (h === 'right') return stringValue('"')
			return stringValue('')
		}
		case 'format': {
			const fmt = cellStyle?.numberFormat ?? 'G'
			if (fmt === 'General' || fmt === 'G') return stringValue('G')
			return stringValue(fmt)
		}
		case 'color':
			return numberValue(cellStyle?.numberFormat?.includes('[Red') ? 1 : 0)
		case 'protect':
			return numberValue(cellStyle?.protection?.locked === false ? 0 : 1)
		case 'parentheses':
			return numberValue(cellStyle?.numberFormat?.includes('(') ? 1 : 0)
		case 'filename': {
			const sheetObj = ctx.workbook.sheets[targetRef.sheetIndex]
			const sheetName = sheetObj?.name ?? `Sheet${targetRef.sheetIndex + 1}`
			return stringValue(`[Workbook]${sheetName}`)
		}
		case 'sheet':
			return stringValue(
				ctx.workbook.sheets[targetRef.sheetIndex]?.name ?? `Sheet${targetRef.sheetIndex + 1}`,
			)
		default:
			return errorValue('#VALUE!')
	}
}

function firstSheetIndexOf(arg: EvalArg): number | null {
	if (arg.areas?.length) return arg.areas[0]?.ref.sheetIndex ?? null
	if (arg.ref) return arg.ref.sheetIndex
	return null
}

function countSheetsOf(arg: EvalArg): number | null {
	if (arg.areas?.length) {
		const sheets = new Set<number>()
		for (const area of arg.areas) sheets.add(area.ref.sheetIndex)
		return sheets.size
	}
	if (arg.ref) return 1
	return null
}

function evalLet(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length < 3 || argNodes.length % 2 === 0) return errorValue('#VALUE!')
	const bindings = new Map<string, LetBinding>(ctx.letBindings)
	for (let i = 0; i < argNodes.length - 1; i += 2) {
		const nameNode = argNodes[i] as FormulaNode
		const valueNode = argNodes[i + 1] as FormulaNode
		if (nameNode.type !== 'name') return errorValue('#VALUE!')
		const boundCtx: EvalContext = { ...ctx, letBindings: bindings }
		const lambda = extractLambda(valueNode, boundCtx)
		bindings.set(nameNode.name.toLowerCase(), lambda ?? evaluate(valueNode, boundCtx))
	}
	const bodyNode = argNodes[argNodes.length - 1] as FormulaNode
	return evaluate(bodyNode, { ...ctx, letBindings: bindings })
}

function resolveR1C1Reference(refText: string, ctx: EvalContext): EvalArg | null {
	const sheetSplit = refText.lastIndexOf('!')
	const sheetName =
		sheetSplit === -1 ? undefined : refText.slice(0, sheetSplit).replace(/^'|'$/g, '')
	const body = sheetSplit === -1 ? refText : refText.slice(sheetSplit + 1)
	const [startToken, endToken = body] = body.split(':')
	if (!startToken) return null
	const start = parseR1C1CellRef(startToken, ctx.row, ctx.col)
	const end = parseR1C1CellRef(endToken, ctx.row, ctx.col)
	if (!start || !end) return null
	const sheetIndex = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
	if (sheetIndex < 0) return { value: errorValue('#REF!') }
	if (start.row < 0 || start.col < 0 || end.row < 0 || end.col < 0) {
		return { value: errorValue('#REF!') }
	}

	return makeRangeArg(
		ctx.workbook,
		sheetIndex,
		start.row,
		start.col,
		end.row,
		end.col,
		ctx.calcContext.dateSystem,
		ctx.calcContext.today,
	)
}

function parseR1C1CellRef(
	token: string,
	currentRow: number,
	currentCol: number,
): { row: number; col: number } | null {
	const match = /^(R(?:\[-?\d+\]|\d+)?)(C(?:\[-?\d+\]|\d+)?)$/i.exec(token.trim())
	if (!match) return null
	const rowToken = match[1]
	const colToken = match[2]
	if (!rowToken || !colToken) return null
	return {
		row: parseR1C1Coordinate(rowToken, currentRow),
		col: parseR1C1Coordinate(colToken, currentCol),
	}
}

function parseR1C1Coordinate(token: string, currentIndex: number): number {
	const value = token.slice(1)
	if (value.startsWith('[') && value.endsWith(']')) {
		return currentIndex + Number.parseInt(value.slice(1, -1) || '0', 10)
	}
	if (value === '') return currentIndex
	return Number.parseInt(value, 10) - 1
}

function evaluateDefinedName(name: string, sheet: string | undefined, ctx: EvalContext): CellValue {
	if (!sheet && ctx.letBindings?.has(name.toLowerCase())) {
		const bound = ctx.letBindings.get(name.toLowerCase()) as LetBinding
		return isLambdaBinding(bound) ? errorValue('#CALC!') : bound
	}
	const resolved = resolveDefinedName(name, sheet, ctx)
	if (!resolved) return errorValue('#NAME?')
	return evaluate(resolved.ast, resolved.ctx)
}

function resolveDefinedName(
	name: string,
	sheet: string | undefined,
	ctx: EvalContext,
): { ast: FormulaNode; ctx: EvalContext } | null {
	const explicitSheet = sheet ? ctx.workbook.getSheet(sheet) : undefined
	const currentSheet = ctx.workbook.sheets[ctx.sheetIndex]
	const entry = ctx.workbook.definedNames.resolve(name, currentSheet?.id, explicitSheet?.id)
	if (!entry) return null

	const entryKey =
		entry.scope.kind === 'workbook'
			? `workbook:${entry.name.toLowerCase()}`
			: `sheet:${entry.scope.sheetId}:${entry.name.toLowerCase()}`
	if (ctx.definedNameStack?.includes(entryKey)) return null

	const parsed = cachedParseFormula(entry.formula)
	if (!parsed.ok) return null

	let sheetIndex = ctx.sheetIndex
	if (entry.scope.kind === 'sheet') {
		const scope = entry.scope
		const localSheetIndex = ctx.workbook.sheets.findIndex(
			(workbookSheet) => workbookSheet.id === scope.sheetId,
		)
		if (localSheetIndex >= 0) sheetIndex = localSheetIndex
	}

	return {
		ast: parsed.value,
		ctx: {
			...ctx,
			sheetIndex,
			definedNameStack: [...(ctx.definedNameStack ?? []), entryKey],
		},
	}
}
