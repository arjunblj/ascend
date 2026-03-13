import { indexToColumn, type Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import {
	type EvalArea,
	type EvalArg,
	type ExactLookupCache,
	type FunctionEvalContext,
	functionRegistry,
	getRange,
	type LookupVectorCache,
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
import { resolveStructuredRefRange } from './structured-refs.ts'

export interface EvalContext {
	readonly workbook: Workbook
	readonly calcContext: CalcContext
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly definedNameStack?: readonly string[]
	readonly letBindings?: ReadonlyMap<string, CellValue>
	readonly exactLookupCache?: ExactLookupCache
	readonly lookupVectorCache?: LookupVectorCache
}

export class MutableEvalContext implements EvalContext {
	workbook!: Workbook
	calcContext!: CalcContext
	sheetIndex = 0
	row = 0
	col = 0
	exactLookupCache?: ExactLookupCache
	lookupVectorCache?: LookupVectorCache
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
		return this
	}
}

const sharedFnCtx = new FunctionEvalCtx()

const cachedParseFormula = sharedCachedParseFormula

export { clearGlobalParseCache as clearFormulaParseCache } from '@ascend/formulas'

const EXCEL_MAX_ROWS = 1_048_576
const EXCEL_MAX_COLS = 16_384

type RangeValueCache = Map<string, readonly (readonly CellValue[])[]>
let activeRangeValueCache: RangeValueCache | null = null

export function setRangeValueCache(cache: RangeValueCache): void {
	activeRangeValueCache = cache
}

export function clearRangeValueCache(): void {
	activeRangeValueCache = null
}

const sheetIndexCache = new WeakMap<Workbook, Map<string, number>>()

function resolveSheetIndex(
	wb: Workbook,
	sheetName: string | undefined,
	currentSheet: number,
): number {
	if (sheetName === undefined) return currentSheet
	let cache = sheetIndexCache.get(wb)
	if (!cache) {
		cache = new Map()
		for (let i = 0; i < wb.sheets.length; i++) {
			const s = wb.sheets[i]
			if (s) cache.set(s.name.toLowerCase(), i)
		}
		sheetIndexCache.set(wb, cache)
	}
	return cache.get(sheetName.toLowerCase()) ?? -1
}

export function invalidateSheetIndexCache(wb: Workbook): void {
	sheetIndexCache.delete(wb)
}

function getCellValue(wb: Workbook, sheetIndex: number, row: number, col: number): CellValue {
	const sheet = wb.sheets[sheetIndex]
	if (!sheet) return errorValue('#REF!')
	return sheet.cells.readValue(row, col)
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
	const cacheKey = `${sheetIndex}:${startRow}:${startCol}:${endRow}:${endCol}`
	if (activeRangeValueCache) {
		const cached = activeRangeValueCache.get(cacheKey)
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
	activeRangeValueCache?.set(cacheKey, rows)
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

function evalBinary(op: string, left: CellValue, right: CellValue): CellValue {
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
			return numberValue(ln ** rn)
		default:
			return errorValue('#VALUE!')
	}
}

function evalComparison(op: string, left: CellValue, right: CellValue): CellValue {
	const ln = coerceToNumber(left)
	const rn = coerceToNumber(right)

	if (ln !== null && rn !== null) {
		return booleanValue(comparePrimitive(op, ln, rn))
	}

	if (left.kind === 'string' || right.kind === 'string') {
		const ls = coerceToString(left).toLowerCase()
		const rs = coerceToString(right).toLowerCase()
		return booleanValue(comparePrimitive(op, ls, rs))
	}

	if (left.kind === 'boolean' && right.kind === 'boolean') {
		const lb = left.value ? 1 : 0
		const rb = right.value ? 1 : 0
		return booleanValue(comparePrimitive(op, lb, rb))
	}

	return booleanValue(comparePrimitive(op, coerceToString(left), coerceToString(right)))
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
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return errorValue('#REF!')
			return getCellValue(ctx.workbook, si, node.ref.row, node.ref.col)
		}

		case 'rangeRef': {
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
			const left = evaluate(node.left, ctx)
			const right = evaluate(node.right, ctx)
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

function evalFunction(name: string, argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const upperName = name.toUpperCase()
	if (upperName === 'ROW' && argNodes.length === 0) {
		return numberValue(ctx.row + 1)
	}
	if (upperName === 'COLUMN' && argNodes.length === 0) {
		return numberValue(ctx.col + 1)
	}
	if (upperName === 'INDIRECT' || upperName === 'OFFSET') {
		return resolveReferenceFunction(upperName, argNodes, ctx).value
	}
	if (upperName === 'FORMULATEXT') {
		return evalFormulaText(argNodes, ctx)
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

	const args: EvalArg[] = new Array(argNodes.length)
	for (let i = 0; i < argNodes.length; i++) {
		args[i] = resolveArg(argNodes[i] as FormulaNode, ctx)
	}
	return def.evaluate(args, sharedFnCtx.update(ctx))
}

function evalLazyArg(node: FormulaNode | undefined, ctx: EvalContext): CellValue {
	return referenceArgToValue(resolveArg(node ?? { type: 'missing' }, ctx))
}

function evalIf(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const condition = coerceToBoolean(resolveArg(argNodes[0] ?? { type: 'missing' }, ctx).value)
	if (typeof condition !== 'boolean') return condition
	if (condition) return evalLazyArg(argNodes[1], ctx)
	return argNodes.length >= 3 ? evalLazyArg(argNodes[2], ctx) : booleanValue(false)
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

function resolveArg(node: FormulaNode, ctx: EvalContext): EvalArg {
	if (node.type === 'function') {
		const upperName = node.name.toUpperCase()
		if (upperName === 'INDIRECT' || upperName === 'OFFSET') {
			return resolveReferenceFunction(upperName, node.args, ctx)
		}
		if (upperName === 'LET') {
			const value = evalLet(node.args, ctx)
			return value.kind === 'array' ? { value, kind: 'range', values: value.rows } : { value }
		}
	}

	if (node.type === 'name') {
		if (!node.sheet && ctx.letBindings?.has(node.name.toLowerCase())) {
			const bound = ctx.letBindings.get(node.name.toLowerCase()) as CellValue
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
	name: 'INDIRECT' | 'OFFSET',
	argNodes: readonly FormulaNode[],
	ctx: EvalContext,
): EvalArg {
	return name === 'INDIRECT'
		? resolveIndirectReference(argNodes, ctx)
		: resolveOffsetReference(argNodes, ctx)
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

	const rowOffset = coerceToNumber(evaluate(argNodes[1] ?? { type: 'missing' }, ctx))
	if (rowOffset === null) return { value: errorValue('#VALUE!') }
	const colOffset = coerceToNumber(evaluate(argNodes[2] ?? { type: 'missing' }, ctx))
	if (colOffset === null) return { value: errorValue('#VALUE!') }

	const baseHeight =
		base.ref.kind === 'range' ? (base.ref.endRow ?? base.ref.row) - base.ref.row + 1 : 1
	const baseWidth =
		base.ref.kind === 'range' ? (base.ref.endCol ?? base.ref.col) - base.ref.col + 1 : 1

	const height =
		argNodes.length > 3 ? coerceToNumber(evaluate(argNodes[3] as FormulaNode, ctx)) : baseHeight
	if (height === null) return { value: errorValue('#VALUE!') }
	const width =
		argNodes.length > 4 ? coerceToNumber(evaluate(argNodes[4] as FormulaNode, ctx)) : baseWidth
	if (width === null) return { value: errorValue('#VALUE!') }

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

	return makeRangeArg(ctx.workbook, base.ref.sheetIndex, startRow, startCol, endRow, endCol)
}

function resolveReferenceNode(node: FormulaNode, ctx: EvalContext): EvalArg | null {
	switch (node.type) {
		case 'cellRef': {
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
			}
		}
		case 'rangeRef': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeRangeArg(
				ctx.workbook,
				si,
				node.start.row,
				node.start.col,
				node.end.row,
				node.end.col,
			)
		}
		case 'wholeRowRange': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeWholeRowArg(ctx.workbook, si, node.startRow, node.endRow)
		}
		case 'wholeColumnRange': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeWholeColumnArg(ctx.workbook, si, node.startCol, node.endCol)
		}
		case 'name': {
			const resolved = resolveDefinedName(node.name, node.sheet, ctx)
			if (!resolved) return { value: errorValue('#NAME?') }
			return (
				resolveReferenceNode(resolved.ast, resolved.ctx) ?? resolveArg(resolved.ast, resolved.ctx)
			)
		}
		case 'structuredRef': {
			const resolved = resolveStructuredRefRange(
				ctx.workbook,
				node,
				ctx.sheetIndex,
				ctx.row,
				ctx.col,
			)
			if (!resolved) return { value: errorValue('#REF!') }
			return makeRangeArg(
				ctx.workbook,
				resolved.sheetIndex,
				resolved.startRow,
				resolved.startCol,
				resolved.endRow,
				resolved.endCol,
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
						makeRangeArea(ctx.workbook, leftRef.sheetIndex, startRow, startCol, endRow, endCol),
					)
				}
			}
			return intersections.length > 0
				? makeMultiAreaArg(intersections)
				: { value: errorValue('#NULL!') }
		}
		default:
			return null
	}
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
): EvalArg {
	return makeMultiAreaArg([makeRangeArea(workbook, sheetIndex, startRow, startCol, endRow, endCol)])
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
	if (arg.value.kind === 'error') return arg.value
	if (arg.value.kind === 'array') return topLeftScalar(arg.value)
	const areas = areasOf(arg)
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
			...(arg.forEachValue ? { forEachValue: arg.forEachValue } : {}),
		},
	]
}

function makeRangeArea(
	workbook: Workbook,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
	options: {
		materializedStartRow?: number
		materializedStartCol?: number
		materializedEndRow?: number
		materializedEndCol?: number
	} = {},
): EvalArea {
	const sheet = workbook.sheets[sheetIndex]
	let cachedValues: readonly (readonly CellValue[])[] | undefined
	const materializedStartRow = options.materializedStartRow ?? startRow
	const materializedStartCol = options.materializedStartCol ?? startCol
	const materializedEndRow = options.materializedEndRow ?? endRow
	const materializedEndCol = options.materializedEndCol ?? endCol
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
): EvalArg {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return { value: errorValue('#REF!') }
	const used = sheet.cells.usedRange()
	const materializedStartCol = used?.start.col ?? 0
	const materializedEndCol = used?.end.col ?? materializedStartCol
	return makeMultiAreaArg([
		makeRangeArea(workbook, sheetIndex, startRow, 0, endRow, EXCEL_MAX_COLS - 1, {
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
): EvalArg {
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return { value: errorValue('#REF!') }
	const used = sheet.cells.usedRange()
	const materializedStartRow = used?.start.row ?? 0
	const materializedEndRow = used?.end.row ?? materializedStartRow
	return makeMultiAreaArg([
		makeRangeArea(workbook, sheetIndex, 0, startCol, EXCEL_MAX_ROWS - 1, endCol, {
			materializedStartRow,
			materializedStartCol: startCol,
			materializedEndRow,
			materializedEndCol: endCol,
		}),
	])
}

interface LambdaInfo {
	readonly params: readonly string[]
	readonly body: FormulaNode
	readonly ctx: EvalContext
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
		if (!node.sheet && ctx.letBindings?.has(node.name.toLowerCase())) return null
		const resolved = resolveDefinedName(node.name, node.sheet, ctx)
		if (!resolved) return null
		return extractLambda(resolved.ast, resolved.ctx)
	}
	return null
}

function extractLambdaFromName(name: string, ctx: EvalContext): LambdaInfo | null {
	const resolved = resolveDefinedName(name, undefined, ctx)
	if (!resolved) return null
	return extractLambda(resolved.ast, resolved.ctx)
}

function invokeLambda(lambda: LambdaInfo, args: readonly CellValue[]): CellValue {
	if (lambda.params.length !== args.length) return errorValue('#VALUE!')
	const bindings = new Map<string, CellValue>(lambda.ctx.letBindings)
	for (let i = 0; i < lambda.params.length; i++) {
		bindings.set(lambda.params[i] as string, args[i] as CellValue)
	}
	return evaluate(lambda.body, { ...lambda.ctx, letBindings: bindings })
}

function evalMap(argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	if (argNodes.length !== 2) return errorValue('#VALUE!')
	const arrayArg = resolveArg(argNodes[0] as FormulaNode, ctx)
	const lambda = extractLambda(argNodes[1] as FormulaNode, ctx)
	if (!lambda) return errorValue('#VALUE!')
	if (lambda.params.length !== 1) return errorValue('#VALUE!')
	const range = getRange(arrayArg)
	const rows: ScalarCellValue[][] = []
	for (const row of range) {
		const resultRow: ScalarCellValue[] = []
		for (const cell of row) {
			const result = invokeLambda(lambda, [cell])
			if (result.kind === 'error') return result
			resultRow.push(topLeftScalar(result))
		}
		rows.push(resultRow)
	}
	if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0] ?? EMPTY
	return arrayValue(rows)
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
	const sheet = ctx.workbook.sheets[arg.ref.sheetIndex]
	if (!sheet) return booleanValue(false)
	return booleanValue(sheet.cells.readFormula(arg.ref.row, arg.ref.col) != null)
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
	const bindings = new Map<string, CellValue>(ctx.letBindings)
	for (let i = 0; i < argNodes.length - 1; i += 2) {
		const nameNode = argNodes[i] as FormulaNode
		const valueNode = argNodes[i + 1] as FormulaNode
		if (nameNode.type !== 'name') return errorValue('#VALUE!')
		const boundCtx: EvalContext = { ...ctx, letBindings: bindings }
		bindings.set(nameNode.name.toLowerCase(), evaluate(valueNode, boundCtx))
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

	return makeRangeArg(ctx.workbook, sheetIndex, start.row, start.col, end.row, end.col)
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
		return ctx.letBindings.get(name.toLowerCase()) as CellValue
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
