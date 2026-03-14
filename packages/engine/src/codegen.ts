import type { FormulaNode } from '@ascend/formulas'
import { toNumber } from '@ascend/formulas'
import type { CellValue } from '@ascend/schema'
import {
	booleanValue,
	coerceCellValueToString,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
	topLeftScalar,
} from '@ascend/schema'
import type { EvalContext } from './evaluator.ts'
import { evaluate as treeEvaluate } from './evaluator.ts'
import { resolveSheetIndexInWorkbook as resolveSheetIndex } from './sheet-index.ts'

export type CodegenFn = (ctx: EvalContext) => CellValue

const CODEGEN_CACHE_MAX = 4096
const codegenCache = new Map<string, CodegenFn | null>()
const sharedCodegenCache = new Map<string, CodegenFn | null>()

const CODEGEN_FUNCTIONS = new Set(['IF', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT'])
const RANGE_AGG_FUNCTIONS = new Set(['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX'])
const PARTIAL_CODEGEN_FUNCTIONS = new Set(['VLOOKUP', 'MATCH', 'INDEX'])

interface SharedAnchor {
	readonly row: number
	readonly col: number
}

function getCachedCodegen(
	cache: Map<string, CodegenFn | null>,
	key: string,
): CodegenFn | null | undefined {
	const cached = cache.get(key)
	if (cached === undefined) return undefined
	cache.delete(key)
	cache.set(key, cached)
	return cached
}

function setCachedCodegen(
	cache: Map<string, CodegenFn | null>,
	key: string,
	value: CodegenFn | null,
): void {
	if (cache.has(key)) cache.delete(key)
	else if (cache.size >= CODEGEN_CACHE_MAX) {
		const oldest = cache.keys().next().value
		if (oldest !== undefined) cache.delete(oldest)
	}
	cache.set(key, value)
}

function canCodegen(node: FormulaNode): boolean {
	switch (node.type) {
		case 'number':
		case 'string':
		case 'boolean':
		case 'error':
		case 'missing':
			return true
		case 'cellRef':
			return true
		case 'binary':
			if (node.op === ',' || node.op === ' ') return false
			return canCodegen(node.left) && canCodegen(node.right)
		case 'unary':
			if (node.op === '@') return false
			return canCodegen(node.operand)
		case 'function': {
			const upper = node.name.toUpperCase()
			if (CODEGEN_FUNCTIONS.has(upper)) return node.args.every(canCodegen)
			if (
				RANGE_AGG_FUNCTIONS.has(upper) &&
				node.args.length === 1 &&
				(node.args[0] as FormulaNode).type === 'rangeRef'
			) {
				return true
			}
			if (upper === 'ROUND' && node.args.length >= 1 && node.args.length <= 2) {
				return node.args.every(canCodegen)
			}
			return PARTIAL_CODEGEN_FUNCTIONS.has(upper)
		}
		default:
			return false
	}
}

interface CodegenState {
	lines: string[]
	varCounter: number
	closureVars: Map<string, string>
	treeNodes: Map<string, FormulaNode>
	sharedAnchor?: SharedAnchor
}

function freshVar(state: CodegenState, prefix = 'v'): string {
	return `${prefix}${state.varCounter++}`
}

function emitReadCell(
	state: CodegenState,
	ref: {
		row: number
		col: number
		rowAbsolute: boolean
		colAbsolute: boolean
	},
	sheet?: string,
): string {
	const result = freshVar(state)
	if (state.sharedAnchor) {
		state.lines.push(
			`var ${result} = _readSharedCell(ctx, _sheet, ${sheet === undefined ? 'null' : JSON.stringify(sheet)}, ${ref.row}, ${ref.col}, ${ref.rowAbsolute}, ${ref.colAbsolute}, ${state.sharedAnchor.row}, ${state.sharedAnchor.col});`,
		)
		return result
	}
	if (sheet !== undefined) {
		const sheetVar = freshVar(state, 'si')
		state.lines.push(
			`var ${sheetVar} = _resolveSheet(ctx.workbook, ${JSON.stringify(sheet)}, ctx.sheetIndex);`,
		)
		state.lines.push(`if (${sheetVar} < 0) return _errorValue('#REF!');`)
		state.lines.push(
			`var ${result} = ctx.workbook.sheets[${sheetVar}]?.cells.readValue(${ref.row}, ${ref.col}) ?? _EMPTY;`,
		)
	} else {
		state.lines.push(`var ${result} = _sheet.cells.readValue(${ref.row}, ${ref.col}) ?? _EMPTY;`)
	}
	return result
}

function emitCoerceNumber(state: CodegenState, valueVar: string): string {
	const result = freshVar(state, 'n')
	state.lines.push(
		`var ${result} = _toNumber(${valueVar}); if (${result} === null) return _errorValue('#VALUE!');`,
	)
	return result
}

function emitNode(state: CodegenState, node: FormulaNode): string {
	switch (node.type) {
		case 'number': {
			const v = freshVar(state)
			state.lines.push(`var ${v} = _numberValue(${node.value});`)
			return v
		}
		case 'string': {
			const v = freshVar(state)
			state.lines.push(`var ${v} = _stringValue(${JSON.stringify(node.value)});`)
			return v
		}
		case 'boolean': {
			const v = freshVar(state)
			state.lines.push(`var ${v} = _booleanValue(${node.value});`)
			return v
		}
		case 'error': {
			const v = freshVar(state)
			state.lines.push(`var ${v} = _errorValue(${JSON.stringify(node.value)});`)
			return v
		}
		case 'missing': {
			const v = freshVar(state)
			state.lines.push(`var ${v} = _EMPTY;`)
			return v
		}
		case 'cellRef': {
			return emitReadCell(state, node.ref, node.sheet)
		}
		case 'binary': {
			return emitBinary(state, node)
		}
		case 'unary': {
			return emitUnary(state, node)
		}
		case 'function': {
			return emitFunction(state, node)
		}
		default: {
			return emitTreeFallback(state, node)
		}
	}
}

function emitBinary(state: CodegenState, node: FormulaNode & { type: 'binary' }): string {
	const op = node.op
	if (op === '&') {
		const lv = emitNode(state, node.left)
		const rv = emitNode(state, node.right)
		const ls = freshVar(state, 's')
		const rs = freshVar(state, 's')
		const result = freshVar(state)
		state.lines.push(`var ${ls} = _topLeft(${lv}); var ${rs} = _topLeft(${rv});`)
		state.lines.push(
			`if (${ls}.kind === 'error') return ${ls}; if (${rs}.kind === 'error') return ${rs};`,
		)
		state.lines.push(`var ${result} = _stringValue(_coerceStr(${ls}) + _coerceStr(${rs}));`)
		return result
	}

	if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
		const lv = emitNode(state, node.left)
		const rv = emitNode(state, node.right)
		const ls = freshVar(state, 'cl')
		const rs = freshVar(state, 'cr')
		const result = freshVar(state)
		state.lines.push(`var ${ls} = _topLeft(${lv}); var ${rs} = _topLeft(${rv});`)
		state.lines.push(
			`if (${ls}.kind === 'error') return ${ls}; if (${rs}.kind === 'error') return ${rs};`,
		)
		state.lines.push(
			`var ${result} = _booleanValue(_evalCmp(${JSON.stringify(op)}, ${ls}, ${rs}));`,
		)
		return result
	}

	const lv = emitNode(state, node.left)
	const rv = emitNode(state, node.right)
	const ls = freshVar(state, 'sl')
	const rs = freshVar(state, 'sr')
	state.lines.push(`var ${ls} = _topLeft(${lv}); var ${rs} = _topLeft(${rv});`)
	state.lines.push(
		`if (${ls}.kind === 'error') return ${ls}; if (${rs}.kind === 'error') return ${rs};`,
	)
	const ln = emitCoerceNumber(state, ls)
	const rn = emitCoerceNumber(state, rs)
	const result = freshVar(state)

	switch (op) {
		case '+':
			state.lines.push(`var ${result} = _numberValue(${ln} + ${rn});`)
			break
		case '-':
			state.lines.push(`var ${result} = _numberValue(${ln} - ${rn});`)
			break
		case '*':
			state.lines.push(`var ${result} = _numberValue(${ln} * ${rn});`)
			break
		case '/':
			state.lines.push(
				`if (${rn} === 0) return _errorValue('#DIV/0!'); var ${result} = _numberValue(${ln} / ${rn});`,
			)
			break
		case '^':
			state.lines.push(`var ${result} = _numberValue(${ln} ** ${rn});`)
			break
		default:
			state.lines.push(`var ${result} = _EMPTY;`)
	}
	return result
}

function emitUnary(state: CodegenState, node: FormulaNode & { type: 'unary' }): string {
	const operandVar = emitNode(state, node.operand)
	const sv = freshVar(state, 'su')
	state.lines.push(`var ${sv} = _topLeft(${operandVar});`)
	state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
	const n = emitCoerceNumber(state, sv)
	const result = freshVar(state)
	switch (node.op) {
		case '+':
			state.lines.push(`var ${result} = _numberValue(${n});`)
			break
		case '-':
			state.lines.push(`var ${result} = _numberValue(-${n});`)
			break
		case '%':
			state.lines.push(`var ${result} = _numberValue(${n} / 100);`)
			break
		default:
			state.lines.push(`var ${result} = _EMPTY;`)
	}
	return result
}

function emitFunction(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const upper = node.name.toUpperCase()

	if (upper === 'IF' && node.args.length >= 2 && node.args.length <= 3) {
		return emitIf(state, node.args)
	}
	if (upper === 'IFERROR' && node.args.length === 2) {
		return emitIfError(state, node.args)
	}
	if (upper === 'IFNA' && node.args.length === 2) {
		return emitIfNa(state, node.args)
	}
	if (upper === 'AND' || upper === 'OR') return emitAndOr(state, node.args, upper === 'AND')
	if (upper === 'NOT' && node.args.length === 1) return emitNot(state, node.args)
	if (upper === 'ROUND' && node.args.length >= 1 && node.args.length <= 2) {
		return emitRound(state, node.args)
	}
	if (
		RANGE_AGG_FUNCTIONS.has(upper) &&
		node.args.length === 1 &&
		(node.args[0] as FormulaNode).type === 'rangeRef'
	) {
		return emitRangeAggregate(state, node.args[0] as FormulaNode & { type: 'rangeRef' }, upper)
	}
	if (upper === 'INDEX') return emitIndex(state, node)
	if (upper === 'MATCH') return emitMatch(state, node)
	if (upper === 'VLOOKUP') return emitVlookup(state, node)

	return emitTreeFallback(state, node)
}

function emitRangeAggregate(
	state: CodegenState,
	node: FormulaNode & { type: 'rangeRef' },
	func: string,
): string {
	const result = freshVar(state)
	const sheetExpr = node.sheet === undefined ? 'null' : JSON.stringify(node.sheet)
	const anchorRow = state.sharedAnchor?.row ?? -1
	const anchorCol = state.sharedAnchor?.col ?? -1
	state.lines.push(
		`var ${result} = _rangeAgg(ctx, _sheet, ${sheetExpr}, ${node.start.row}, ${node.start.col}, ${node.start.rowAbsolute}, ${node.start.colAbsolute}, ${node.end.row}, ${node.end.col}, ${node.end.rowAbsolute}, ${node.end.colAbsolute}, ${anchorRow}, ${anchorCol}, ${JSON.stringify(func)});`,
	)
	return result
}

function emitAndOr(state: CodegenState, args: readonly FormulaNode[], isAnd: boolean): string {
	const result = freshVar(state)
	state.lines.push(`var ${result};`)
	for (let i = 0; i < args.length; i++) {
		const argVar = emitNode(state, args[i] as FormulaNode)
		const boolVar = freshVar(state, 'ab')
		state.lines.push(`var ${boolVar} = _coerceBool(${argVar});`)
		state.lines.push(`if (typeof ${boolVar} !== 'boolean') return ${boolVar};`)
		if (isAnd) {
			state.lines.push(`if (!${boolVar}) { ${result} = _booleanValue(false); }`)
		} else {
			state.lines.push(`if (${boolVar}) { ${result} = _booleanValue(true); }`)
		}
	}
	state.lines.push(`if (${result} === undefined) ${result} = _booleanValue(${isAnd});`)
	return result
}

function emitNot(state: CodegenState, args: readonly FormulaNode[]): string {
	const argVar = emitNode(state, args[0] as FormulaNode)
	const boolVar = freshVar(state, 'nb')
	state.lines.push(`var ${boolVar} = _coerceBool(${argVar});`)
	state.lines.push(`if (typeof ${boolVar} !== 'boolean') return ${boolVar};`)
	const result = freshVar(state)
	state.lines.push(`var ${result} = _booleanValue(!${boolVar});`)
	return result
}

function emitRound(state: CodegenState, args: readonly FormulaNode[]): string {
	const valueVar = emitNode(state, args[0] as FormulaNode)
	const sv = freshVar(state, 'rv')
	state.lines.push(`var ${sv} = _topLeft(${valueVar});`)
	state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
	const nv = emitCoerceNumber(state, sv)
	let digitsExpr = '0'
	if (args.length >= 2) {
		const digitsVar = emitNode(state, args[1] as FormulaNode)
		const ds = freshVar(state, 'ds')
		state.lines.push(`var ${ds} = _topLeft(${digitsVar});`)
		state.lines.push(`if (${ds}.kind === 'error') return ${ds};`)
		digitsExpr = emitCoerceNumber(state, ds)
	}
	const result = freshVar(state)
	state.lines.push(
		`var _m${state.varCounter} = Math.pow(10, Math.floor(${digitsExpr})); var ${result} = _numberValue(Math.round(${nv} * _m${state.varCounter}) / _m${state.varCounter});`,
	)
	state.varCounter++
	return result
}

function emitIndex(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const args = node.args
	if (args.length < 2 || args.length > 3) return emitTreeFallback(state, node)
	const rangeArg = args[0]
	if (!rangeArg || rangeArg.type !== 'rangeRef') return emitTreeFallback(state, node)

	const rowVar = emitNode(state, args[1] as FormulaNode)
	let colExpr = 'null'
	if (args.length > 2) colExpr = emitNode(state, args[2] as FormulaNode)

	const nodeKey = `_tree${state.treeNodes.size}`
	state.treeNodes.set(nodeKey, node)

	const result = freshVar(state)
	const sheetExpr = rangeArg.sheet === undefined ? 'null' : JSON.stringify(rangeArg.sheet)
	const anchorRow = state.sharedAnchor?.row ?? -1
	const anchorCol = state.sharedAnchor?.col ?? -1
	state.lines.push(
		`var ${result} = _indexRange(ctx, _sheet, ${sheetExpr}, ${rangeArg.start.row}, ${rangeArg.start.col}, ${rangeArg.start.rowAbsolute}, ${rangeArg.start.colAbsolute}, ${rangeArg.end.row}, ${rangeArg.end.col}, ${rangeArg.end.rowAbsolute}, ${rangeArg.end.colAbsolute}, ${anchorRow}, ${anchorCol}, ${rowVar}, ${colExpr}, ${nodeKey});`,
	)
	return result
}

function emitMatch(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const args = node.args
	if (args.length < 2 || args.length > 3) return emitTreeFallback(state, node)
	const rangeArg = args[1]
	if (!rangeArg || rangeArg.type !== 'rangeRef') return emitTreeFallback(state, node)

	if (args.length >= 3) {
		const typeArg = args[2]
		if (!typeArg || typeArg.type !== 'number' || typeArg.value !== 0)
			return emitTreeFallback(state, node)
	} else {
		return emitTreeFallback(state, node)
	}

	const lookupVar = emitNode(state, args[0] as FormulaNode)
	const nodeKey = `_tree${state.treeNodes.size}`
	state.treeNodes.set(nodeKey, node)

	const result = freshVar(state)
	const sheetExpr = rangeArg.sheet === undefined ? 'null' : JSON.stringify(rangeArg.sheet)
	const anchorRow = state.sharedAnchor?.row ?? -1
	const anchorCol = state.sharedAnchor?.col ?? -1
	state.lines.push(
		`var ${result} = _matchExact(ctx, _sheet, ${sheetExpr}, ${rangeArg.start.row}, ${rangeArg.start.col}, ${rangeArg.start.rowAbsolute}, ${rangeArg.start.colAbsolute}, ${rangeArg.end.row}, ${rangeArg.end.col}, ${rangeArg.end.rowAbsolute}, ${rangeArg.end.colAbsolute}, ${anchorRow}, ${anchorCol}, ${lookupVar}, ${nodeKey});`,
	)
	return result
}

function emitVlookup(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const args = node.args
	if (args.length < 3 || args.length > 4) return emitTreeFallback(state, node)
	const tableArg = args[1]
	if (!tableArg || tableArg.type !== 'rangeRef') return emitTreeFallback(state, node)

	if (args.length >= 4) {
		const approxArg = args[3]
		if (!approxArg) return emitTreeFallback(state, node)
		const isExact =
			(approxArg.type === 'boolean' && approxArg.value === false) ||
			(approxArg.type === 'number' && approxArg.value === 0)
		if (!isExact) return emitTreeFallback(state, node)
	} else {
		return emitTreeFallback(state, node)
	}

	const lookupVar = emitNode(state, args[0] as FormulaNode)
	const colIdxVar = emitNode(state, args[2] as FormulaNode)

	const nodeKey = `_tree${state.treeNodes.size}`
	state.treeNodes.set(nodeKey, node)

	const result = freshVar(state)
	const sheetExpr = tableArg.sheet === undefined ? 'null' : JSON.stringify(tableArg.sheet)
	const anchorRow = state.sharedAnchor?.row ?? -1
	const anchorCol = state.sharedAnchor?.col ?? -1
	state.lines.push(
		`var ${result} = _vlookupExact(ctx, _sheet, ${sheetExpr}, ${tableArg.start.row}, ${tableArg.start.col}, ${tableArg.start.rowAbsolute}, ${tableArg.start.colAbsolute}, ${tableArg.end.row}, ${tableArg.end.col}, ${tableArg.end.rowAbsolute}, ${tableArg.end.colAbsolute}, ${anchorRow}, ${anchorCol}, ${lookupVar}, ${colIdxVar}, ${nodeKey});`,
	)
	return result
}

function emitIf(state: CodegenState, args: readonly FormulaNode[]): string {
	const condVar = emitNode(state, args[0] as FormulaNode)
	const boolVar = freshVar(state, 'b')
	const result = freshVar(state)
	state.lines.push(`var ${boolVar} = _coerceBool(${condVar});`)
	state.lines.push(`if (typeof ${boolVar} !== 'boolean') return ${boolVar};`)
	state.lines.push(`var ${result};`)
	state.lines.push(`if (${boolVar}) {`)
	const trueVar = emitNode(state, args[1] as FormulaNode)
	state.lines.push(`${result} = ${trueVar};`)
	state.lines.push('} else {')
	if (args.length >= 3) {
		const falseVar = emitNode(state, args[2] as FormulaNode)
		state.lines.push(`${result} = ${falseVar};`)
	} else {
		state.lines.push(`${result} = _booleanValue(false);`)
	}
	state.lines.push('}')
	return result
}

function emitIfError(state: CodegenState, args: readonly FormulaNode[]): string {
	const innerState: CodegenState = {
		lines: [],
		varCounter: state.varCounter,
		closureVars: state.closureVars,
		treeNodes: state.treeNodes,
	}
	const innerVar = emitNode(innerState, args[0] as FormulaNode)
	innerState.lines.push(`return ${innerVar};`)
	state.varCounter = innerState.varCounter

	const valueVar = freshVar(state, 'iev')
	state.lines.push(`var ${valueVar} = (function() { ${innerState.lines.join('\n')} }).call(this);`)
	const result = freshVar(state)
	const scalar = freshVar(state, 'ie')
	state.lines.push(`var ${scalar} = _topLeft(${valueVar});`)
	state.lines.push(`var ${result};`)
	state.lines.push(`if (${scalar}.kind === 'error') {`)
	const fallbackVar = emitNode(state, args[1] as FormulaNode)
	state.lines.push(`${result} = ${fallbackVar};`)
	state.lines.push('} else {')
	state.lines.push(`${result} = ${valueVar};`)
	state.lines.push('}')
	return result
}

function emitIfNa(state: CodegenState, args: readonly FormulaNode[]): string {
	const innerState: CodegenState = {
		lines: [],
		varCounter: state.varCounter,
		closureVars: state.closureVars,
		treeNodes: state.treeNodes,
	}
	const innerVar = emitNode(innerState, args[0] as FormulaNode)
	innerState.lines.push(`return ${innerVar};`)
	state.varCounter = innerState.varCounter

	const valueVar = freshVar(state, 'inv')
	state.lines.push(`var ${valueVar} = (function() { ${innerState.lines.join('\n')} }).call(this);`)
	const result = freshVar(state)
	const scalar = freshVar(state, 'in')
	state.lines.push(`var ${scalar} = _topLeft(${valueVar});`)
	state.lines.push(`var ${result};`)
	state.lines.push(`if (${scalar}.kind === 'error' && ${scalar}.value === '#N/A') {`)
	const fallbackVar = emitNode(state, args[1] as FormulaNode)
	state.lines.push(`${result} = ${fallbackVar};`)
	state.lines.push('} else {')
	state.lines.push(`${result} = ${valueVar};`)
	state.lines.push('}')
	return result
}

function emitTreeFallback(state: CodegenState, node: FormulaNode): string {
	const key = `_tree${state.treeNodes.size}`
	state.treeNodes.set(key, node)
	const result = freshVar(state)
	state.lines.push(`var ${result} = _treeEval(${key}, ctx);`)
	return result
}

function evalCmp(op: string, left: CellValue, right: CellValue): boolean {
	const ln = toNumber(left)
	const rn = toNumber(right)
	if (ln !== null && rn !== null) return cmpPrimitive(op, ln, rn)
	if (left.kind === 'string' || right.kind === 'string') {
		return cmpPrimitive(
			op,
			coerceCellValueToString(left).toLowerCase(),
			coerceCellValueToString(right).toLowerCase(),
		)
	}
	if (left.kind === 'boolean' && right.kind === 'boolean') {
		return cmpPrimitive(op, left.value ? 1 : 0, right.value ? 1 : 0)
	}
	return cmpPrimitive(op, coerceCellValueToString(left), coerceCellValueToString(right))
}

function cmpPrimitive<T extends number | string>(op: string, a: T, b: T): boolean {
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

function coerceToBoolForCodegen(v: CellValue): boolean | CellValue {
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

function readSharedCell(
	ctx: EvalContext,
	currentSheet: import('@ascend/core').Workbook['sheets'][number] | undefined,
	sheetName: string | null,
	row: number,
	col: number,
	rowAbsolute: boolean,
	colAbsolute: boolean,
	anchorRow: number,
	anchorCol: number,
): CellValue {
	const targetRow = rowAbsolute ? row : ctx.row + (row - anchorRow)
	const targetCol = colAbsolute ? col : ctx.col + (col - anchorCol)
	if (targetRow < 0 || targetCol < 0) return errorValue('#REF!')
	if (sheetName === null) {
		if (!currentSheet) return errorValue('#REF!')
		return currentSheet.cells.readValue(targetRow, targetCol)
	}
	const sheetIndex = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
	if (sheetIndex < 0) return errorValue('#REF!')
	const targetSheet = ctx.workbook.sheets[sheetIndex]
	if (!targetSheet) return errorValue('#REF!')
	return targetSheet.cells.readValue(targetRow, targetCol)
}

function resolveRelativeIndex(
	value: number,
	absolute: boolean,
	current: number,
	anchor: number,
): number {
	return absolute || anchor < 0 ? value : current + (value - anchor)
}

function rangeAggregate(
	ctx: EvalContext,
	currentSheet: import('@ascend/core').Workbook['sheets'][number] | undefined,
	sheetName: string | null,
	startRow: number,
	startCol: number,
	startRowAbsolute: boolean,
	startColAbsolute: boolean,
	endRow: number,
	endCol: number,
	endRowAbsolute: boolean,
	endColAbsolute: boolean,
	anchorRow: number,
	anchorCol: number,
	func: string,
): CellValue {
	const resolvedStartRow = resolveRelativeIndex(startRow, startRowAbsolute, ctx.row, anchorRow)
	const resolvedStartCol = resolveRelativeIndex(startCol, startColAbsolute, ctx.col, anchorCol)
	const resolvedEndRow = resolveRelativeIndex(endRow, endRowAbsolute, ctx.row, anchorRow)
	const resolvedEndCol = resolveRelativeIndex(endCol, endColAbsolute, ctx.col, anchorCol)
	if (resolvedStartRow < 0 || resolvedStartCol < 0 || resolvedEndRow < 0 || resolvedEndCol < 0) {
		return errorValue('#REF!')
	}
	let targetSheet = currentSheet
	if (sheetName !== null) {
		const sheetIndex = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
		if (sheetIndex < 0) return errorValue('#REF!')
		targetSheet = ctx.workbook.sheets[sheetIndex]
	}
	if (!targetSheet) return errorValue('#REF!')
	const fromRow = Math.min(resolvedStartRow, resolvedEndRow)
	const toRow = Math.max(resolvedStartRow, resolvedEndRow)
	const fromCol = Math.min(resolvedStartCol, resolvedEndCol)
	const toCol = Math.max(resolvedStartCol, resolvedEndCol)

	let acc = 0
	let count = 0
	let countA = 0
	let minVal = Number.POSITIVE_INFINITY
	let maxVal = Number.NEGATIVE_INFINITY

	for (let row = fromRow; row <= toRow; row++) {
		for (let col = fromCol; col <= toCol; col++) {
			const kind = targetSheet.cells.readKind(row, col)
			if (kind === undefined || kind === 'empty') continue
			if (kind === 'error') {
				return errorValue(targetSheet.cells.readError(row, col) ?? '#VALUE!')
			}
			countA++
			if (kind !== 'number' && kind !== 'date') continue
			const n = targetSheet.cells.readNumber(row, col) ?? 0
			acc += n
			count++
			if (n < minVal) minVal = n
			if (n > maxVal) maxVal = n
		}
	}

	switch (func) {
		case 'SUM':
			return numberValue(acc)
		case 'AVERAGE':
			return count === 0 ? errorValue('#DIV/0!') : numberValue(acc / count)
		case 'COUNT':
			return numberValue(count)
		case 'COUNTA':
			return numberValue(countA)
		case 'MIN':
			return count === 0 ? numberValue(0) : numberValue(minVal)
		case 'MAX':
			return count === 0 ? numberValue(0) : numberValue(maxVal)
		default:
			return numberValue(acc)
	}
}

function lookupValueKey(value: CellValue): string | null {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'number':
			return `n:${String(scalar.value === 0 ? 0 : scalar.value)}`
		case 'date':
			return `d:${scalar.serial}`
		case 'string':
			return `s:${scalar.value.toLowerCase()}`
		case 'boolean':
			return scalar.value ? 'b:1' : 'b:0'
		case 'empty':
			return 'e'
		default:
			return null
	}
}

function getLookupIndex(
	ctx: EvalContext,
	sheetIndex: number,
	col: number,
	startRow: number,
	endRow: number,
): ReadonlyMap<string, { first: number; last: number }> {
	const cacheKey = `column:${sheetIndex}:${col}:${startRow}:${endRow}`
	if (ctx.exactLookupCache) {
		const cached = ctx.exactLookupCache.get(cacheKey)
		if (cached) return cached
	}
	const sheet = ctx.workbook.sheets[sheetIndex]
	const index = new Map<string, { first: number; last: number }>()
	if (sheet) {
		for (let i = startRow; i <= endRow; i++) {
			const key = lookupValueKey(sheet.cells.readValue(i, col))
			if (key === null) continue
			const offset = i - startRow
			const existing = index.get(key)
			if (existing) index.set(key, { first: existing.first, last: offset })
			else index.set(key, { first: offset, last: offset })
		}
	}
	ctx.exactLookupCache?.set(cacheKey, index)
	return index
}

function indexedLookup(
	lookup: CellValue,
	index: ReadonlyMap<string, { first: number; last: number }>,
): number {
	const key = lookupValueKey(lookup)
	if (key === null) return -1
	const hit = index.get(key)
	return hit ? hit.first : -1
}

function hasWildcardChars(text: string): boolean {
	for (let i = 0; i < text.length; i++) {
		const ch = text[i]
		if ((ch === '*' || ch === '?') && text[i - 1] !== '~') return true
	}
	return false
}

function indexRange(
	ctx: EvalContext,
	currentSheet: import('@ascend/core').Workbook['sheets'][number] | undefined,
	sheetName: string | null,
	startRow: number,
	startCol: number,
	startRowAbsolute: boolean,
	startColAbsolute: boolean,
	endRow: number,
	endCol: number,
	endRowAbsolute: boolean,
	endColAbsolute: boolean,
	anchorRow: number,
	anchorCol: number,
	rowVal: CellValue,
	colVal: CellValue | null,
	fallbackNode: FormulaNode,
): CellValue {
	const sr = resolveRelativeIndex(startRow, startRowAbsolute, ctx.row, anchorRow)
	const sc = resolveRelativeIndex(startCol, startColAbsolute, ctx.col, anchorCol)
	const er = resolveRelativeIndex(endRow, endRowAbsolute, ctx.row, anchorRow)
	const ec = resolveRelativeIndex(endCol, endColAbsolute, ctx.col, anchorCol)
	if (sr < 0 || sc < 0 || er < 0 || ec < 0) return errorValue('#REF!')

	let sheet = currentSheet
	if (sheetName !== null) {
		const si = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
		if (si < 0) return errorValue('#REF!')
		sheet = ctx.workbook.sheets[si]
	}
	if (!sheet) return errorValue('#REF!')

	const rowScalar = topLeftScalar(rowVal)
	if (rowScalar.kind === 'error') return rowScalar
	const rn = toNumber(rowScalar)
	if (rn === null) return errorValue('#VALUE!')
	const row = Math.floor(rn)
	const rows = er - sr + 1
	const cols = ec - sc + 1

	if (colVal !== null) {
		const colScalar = topLeftScalar(colVal)
		if (colScalar.kind === 'error') return colScalar
		const cn = toNumber(colScalar)
		if (cn === null) return errorValue('#VALUE!')
		const col = Math.floor(cn)
		if (row === 0 || col === 0) return treeEvaluate(fallbackNode, ctx)
		if (row < 1 || row > rows || col < 1 || col > cols) return errorValue('#REF!')
		return sheet.cells.readValue(sr + row - 1, sc + col - 1)
	}

	if (cols === 1) {
		if (row < 1 || row > rows) return errorValue('#REF!')
		return sheet.cells.readValue(sr + row - 1, sc)
	}
	if (rows === 1) {
		if (row < 1 || row > cols) return errorValue('#REF!')
		return sheet.cells.readValue(sr, sc + row - 1)
	}
	if (row < 1 || row > rows) return errorValue('#REF!')
	return sheet.cells.readValue(sr + row - 1, sc)
}

function matchExact(
	ctx: EvalContext,
	_currentSheet: import('@ascend/core').Workbook['sheets'][number] | undefined,
	sheetName: string | null,
	startRow: number,
	startCol: number,
	startRowAbsolute: boolean,
	startColAbsolute: boolean,
	endRow: number,
	endCol: number,
	endRowAbsolute: boolean,
	endColAbsolute: boolean,
	anchorRow: number,
	anchorCol: number,
	lookupVal: CellValue,
	fallbackNode: FormulaNode,
): CellValue {
	if (lookupVal.kind === 'array') return treeEvaluate(fallbackNode, ctx)
	const lookup = topLeftScalar(lookupVal)
	if (lookup.kind === 'error') return lookup
	if (lookup.kind === 'string' && hasWildcardChars(lookup.value))
		return treeEvaluate(fallbackNode, ctx)

	const sr = resolveRelativeIndex(startRow, startRowAbsolute, ctx.row, anchorRow)
	const sc = resolveRelativeIndex(startCol, startColAbsolute, ctx.col, anchorCol)
	const er = resolveRelativeIndex(endRow, endRowAbsolute, ctx.row, anchorRow)
	const ec = resolveRelativeIndex(endCol, endColAbsolute, ctx.col, anchorCol)
	if (sr < 0 || sc < 0 || er < 0 || ec < 0) return errorValue('#REF!')

	let targetSheetIndex = ctx.sheetIndex
	if (sheetName !== null) {
		targetSheetIndex = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
		if (targetSheetIndex < 0) return errorValue('#REF!')
	}

	const rows = er - sr + 1
	const cols = ec - sc + 1
	const isRow = rows === 1 && cols > 1

	if (isRow) {
		const rowCacheKey = `row:${targetSheetIndex}:${sr}:${sc}:${ec}`
		let rowIndex: ReadonlyMap<string, { first: number; last: number }> | undefined
		if (ctx.exactLookupCache) {
			rowIndex = ctx.exactLookupCache.get(rowCacheKey)
		}
		if (!rowIndex) {
			const sheet = ctx.workbook.sheets[targetSheetIndex]
			const built = new Map<string, { first: number; last: number }>()
			if (sheet) {
				for (let i = sc; i <= ec; i++) {
					const key = lookupValueKey(sheet.cells.readValue(sr, i))
					if (key === null) continue
					const offset = i - sc
					const existing = built.get(key)
					if (existing) built.set(key, { first: existing.first, last: offset })
					else built.set(key, { first: offset, last: offset })
				}
			}
			rowIndex = built
			ctx.exactLookupCache?.set(rowCacheKey, built)
		}
		const idx = indexedLookup(lookup, rowIndex)
		return idx < 0 ? errorValue('#N/A') : numberValue(idx + 1)
	}

	const index = getLookupIndex(ctx, targetSheetIndex, sc, sr, er)
	const idx = indexedLookup(lookup, index)
	return idx < 0 ? errorValue('#N/A') : numberValue(idx + 1)
}

function vlookupExact(
	ctx: EvalContext,
	currentSheet: import('@ascend/core').Workbook['sheets'][number] | undefined,
	sheetName: string | null,
	startRow: number,
	startCol: number,
	startRowAbsolute: boolean,
	startColAbsolute: boolean,
	endRow: number,
	endCol: number,
	endRowAbsolute: boolean,
	endColAbsolute: boolean,
	anchorRow: number,
	anchorCol: number,
	lookupVal: CellValue,
	colIdxVal: CellValue,
	fallbackNode: FormulaNode,
): CellValue {
	if (lookupVal.kind === 'array') return treeEvaluate(fallbackNode, ctx)

	const sr = resolveRelativeIndex(startRow, startRowAbsolute, ctx.row, anchorRow)
	const sc = resolveRelativeIndex(startCol, startColAbsolute, ctx.col, anchorCol)
	const er = resolveRelativeIndex(endRow, endRowAbsolute, ctx.row, anchorRow)
	const ec = resolveRelativeIndex(endCol, endColAbsolute, ctx.col, anchorCol)
	if (sr < 0 || sc < 0 || er < 0 || ec < 0) return errorValue('#REF!')

	let sheet = currentSheet
	let targetSheetIndex = ctx.sheetIndex
	if (sheetName !== null) {
		targetSheetIndex = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
		if (targetSheetIndex < 0) return errorValue('#REF!')
		sheet = ctx.workbook.sheets[targetSheetIndex]
	}
	if (!sheet) return errorValue('#REF!')

	const colScalar = topLeftScalar(colIdxVal)
	if (colScalar.kind === 'error') return colScalar
	const cn = toNumber(colScalar)
	if (cn === null) return errorValue('#VALUE!')
	const colInt = Math.floor(cn)
	const cols = ec - sc + 1
	if (colInt < 1 || colInt > cols) return errorValue('#REF!')

	const lookup = topLeftScalar(lookupVal)
	if (lookup.kind === 'error') return lookup

	const index = getLookupIndex(ctx, targetSheetIndex, sc, sr, er)
	const idx = indexedLookup(lookup, index)
	return idx < 0 ? errorValue('#N/A') : sheet.cells.readValue(sr + idx, sc + colInt - 1)
}

function buildCodegenFn(node: FormulaNode, sharedAnchor?: SharedAnchor): CodegenFn | null {
	if (!canCodegen(node)) return null

	const state: CodegenState = {
		lines: [],
		varCounter: 0,
		closureVars: new Map(),
		treeNodes: new Map(),
	}
	if (sharedAnchor) state.sharedAnchor = sharedAnchor

	const resultVar = emitNode(state, node)
	state.lines.push(`return ${resultVar};`)

	const closureNames: string[] = [
		'_numberValue',
		'_stringValue',
		'_booleanValue',
		'_errorValue',
		'_EMPTY',
		'_toNumber',
		'_topLeft',
		'_coerceStr',
		'_evalCmp',
		'_coerceBool',
		'_resolveSheet',
		'_readSharedCell',
		'_rangeAgg',
		'_treeEval',
		'_indexRange',
		'_matchExact',
		'_vlookupExact',
	]
	const closureValues: unknown[] = [
		numberValue,
		stringValue,
		booleanValue,
		errorValue,
		EMPTY,
		toNumber,
		topLeftScalar,
		coerceCellValueToString,
		evalCmp,
		coerceToBoolForCodegen,
		resolveSheetIndex,
		readSharedCell,
		rangeAggregate,
		treeEvaluate,
		indexRange,
		matchExact,
		vlookupExact,
	]

	for (const [key, node] of state.treeNodes) {
		closureNames.push(key)
		closureValues.push(node)
	}

	const body = `var _sheet = ctx.workbook.sheets[ctx.sheetIndex];\n${state.lines.join('\n')}`

	const factory = new Function(
		...closureNames,
		`return function codegenFormula(ctx) {\n${body}\n};`,
	)
	return factory(...closureValues) as CodegenFn
}

export function codegenFormula(formulaText: string, ast: FormulaNode): CodegenFn | null {
	const cached = getCachedCodegen(codegenCache, formulaText)
	if (cached !== undefined) return cached

	const fn = buildCodegenFn(ast)
	setCachedCodegen(codegenCache, formulaText, fn)
	return fn
}

export function codegenSharedFormula(
	formulaText: string,
	ast: FormulaNode,
	anchor: SharedAnchor,
): CodegenFn | null {
	const cacheKey = `${formulaText}@${anchor.row}:${anchor.col}`
	const cached = getCachedCodegen(sharedCodegenCache, cacheKey)
	if (cached !== undefined) return cached

	const fn = buildCodegenFn(ast, anchor)
	setCachedCodegen(sharedCodegenCache, cacheKey, fn)
	return fn
}

export function clearCodegenCache(): void {
	codegenCache.clear()
	sharedCodegenCache.clear()
}
