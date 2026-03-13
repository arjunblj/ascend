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

export type CodegenFn = (ctx: EvalContext) => CellValue

const codegenCache = new Map<string, CodegenFn | null>()

const CODEGEN_FUNCTIONS = new Set(['IF', 'IFERROR', 'IFNA'])

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
			if (!CODEGEN_FUNCTIONS.has(upper)) return false
			return node.args.every(canCodegen)
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
}

function freshVar(state: CodegenState, prefix = 'v'): string {
	return `${prefix}${state.varCounter++}`
}

function emitReadCell(state: CodegenState, row: number, col: number, sheet?: string): string {
	const result = freshVar(state)
	if (sheet !== undefined) {
		const sheetVar = freshVar(state, 'si')
		state.lines.push(
			`var ${sheetVar} = _resolveSheet(ctx.workbook, ${JSON.stringify(sheet)}, ctx.sheetIndex);`,
		)
		state.lines.push(`if (${sheetVar} < 0) return _errorValue('#REF!');`)
		state.lines.push(
			`var ${result} = ctx.workbook.sheets[${sheetVar}]?.cells.readValue(${row}, ${col}) ?? _EMPTY;`,
		)
	} else {
		state.lines.push(`var ${result} = _sheet.cells.readValue(${row}, ${col}) ?? _EMPTY;`)
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
			return emitReadCell(state, node.ref.row, node.ref.col, node.sheet)
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

	return emitTreeFallback(state, node)
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

const sheetIndexCache = new WeakMap<import('@ascend/core').Workbook, Map<string, number>>()

function resolveSheetIndex(
	wb: import('@ascend/core').Workbook,
	sheetName: string,
	_currentSheet: number,
): number {
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

function buildCodegenFn(node: FormulaNode): CodegenFn | null {
	if (!canCodegen(node)) return null

	const state: CodegenState = {
		lines: [],
		varCounter: 0,
		closureVars: new Map(),
		treeNodes: new Map(),
	}

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
		'_treeEval',
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
		treeEvaluate,
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
	const cached = codegenCache.get(formulaText)
	if (cached !== undefined) return cached

	const fn = buildCodegenFn(ast)
	codegenCache.set(formulaText, fn)
	return fn
}

export function clearCodegenCache(): void {
	codegenCache.clear()
}
