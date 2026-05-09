import type { FormulaNode } from '@ascend/formulas'
import { dateToSerial, formatNumber, serialToDate, toNumber, wildcardMatch } from '@ascend/formulas'
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
import { aggregateNumericRange } from './compiled-eval.ts'
import type { EvalContext } from './evaluator.ts'
import { evaluate as treeEvaluate } from './evaluator.ts'
import { resolveSheetIndexInWorkbook as resolveSheetIndex } from './sheet-index.ts'

export type CodegenFn = (ctx: EvalContext) => CellValue

const CODEGEN_CACHE_MAX = 4096

class GenerationalCache<V> {
	private young = new Map<string, V>()
	private old = new Map<string, V>()
	private readonly maxSize: number

	constructor(maxSize: number) {
		this.maxSize = maxSize
	}

	get(key: string): V | undefined {
		const fromYoung = this.young.get(key)
		if (fromYoung !== undefined) return fromYoung
		const fromOld = this.old.get(key)
		if (fromOld !== undefined) {
			this.young.set(key, fromOld)
			return fromOld
		}
		return undefined
	}

	set(key: string, value: V): void {
		this.young.set(key, value)
		if (this.young.size >= this.maxSize) {
			this.old = this.young
			this.young = new Map()
		}
	}

	clear(): void {
		this.young.clear()
		this.old.clear()
	}
}

const codegenCache = new GenerationalCache<CodegenFn | null>(CODEGEN_CACHE_MAX)
const sharedCodegenCache = new GenerationalCache<CodegenFn | null>(CODEGEN_CACHE_MAX)

const CODEGEN_FUNCTIONS = new Set(['IF', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT'])
const RANGE_AGG_FUNCTIONS = new Set(['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX'])
const PARTIAL_CODEGEN_FUNCTIONS = new Set(['VLOOKUP', 'MATCH', 'INDEX'])
const DATE_EXTRACT_FUNCTIONS = new Set(['YEAR', 'MONTH', 'DAY'])

interface SharedAnchor {
	readonly row: number
	readonly col: number
}

const SIMPLE_MATH_FUNCTIONS = new Set([
	'ABS',
	'SIGN',
	'INT',
	'MOD',
	'POWER',
	'SQRT',
	'LN',
	'LOG10',
	'EXP',
	'ROUNDUP',
	'ROUNDDOWN',
	'TRUNC',
	'CEILING',
	'FLOOR',
	'PI',
	'DEGREES',
	'RADIANS',
	'SIN',
	'COS',
	'TAN',
	'ASIN',
	'ACOS',
	'ATAN',
])

const SIMPLE_TEXT_FUNCTIONS = new Set(['LEN', 'UPPER', 'LOWER', 'TRIM', 'PROPER'])

const TEXT_WITH_ARGS_FUNCTIONS = new Set(['LEFT', 'RIGHT', 'MID', 'CONCATENATE', 'CONCAT', 'REPT'])

function isConstantNode(node: FormulaNode): boolean {
	switch (node.type) {
		case 'number':
		case 'string':
		case 'boolean':
		case 'error':
			return true
		case 'binary':
			if (node.op === ',' || node.op === ' ') return false
			return isConstantNode(node.left) && isConstantNode(node.right)
		case 'unary':
			if (node.op === '@') return false
			return isConstantNode(node.operand)
		case 'function': {
			const upper = node.name.toUpperCase()
			if (upper === 'PI') return true
			if (SIMPLE_MATH_FUNCTIONS.has(upper) || upper === 'ROUND') {
				return node.args.every(isConstantNode)
			}
			return false
		}
		default:
			return false
	}
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
			if (SIMPLE_MATH_FUNCTIONS.has(upper)) {
				const expected1 = new Set([
					'ABS',
					'SIGN',
					'INT',
					'SQRT',
					'LN',
					'LOG10',
					'EXP',
					'DEGREES',
					'RADIANS',
					'SIN',
					'COS',
					'TAN',
					'ASIN',
					'ACOS',
					'ATAN',
				])
				const expected0 = new Set(['PI'])
				const expected2 = new Set(['MOD', 'POWER', 'ROUNDUP', 'ROUNDDOWN', 'CEILING', 'FLOOR'])
				const expected12 = new Set(['TRUNC'])
				if (expected0.has(upper) && node.args.length === 0) return true
				if (expected1.has(upper) && node.args.length === 1)
					return canCodegen(node.args[0] as FormulaNode)
				if (expected2.has(upper) && node.args.length === 2) return node.args.every(canCodegen)
				if (expected12.has(upper) && node.args.length >= 1 && node.args.length <= 2)
					return node.args.every(canCodegen)
				return false
			}
			if (SIMPLE_TEXT_FUNCTIONS.has(upper) && node.args.length === 1) {
				return canCodegen(node.args[0] as FormulaNode)
			}
			if (TEXT_WITH_ARGS_FUNCTIONS.has(upper)) {
				return node.args.every(canCodegen)
			}
			if (DATE_EXTRACT_FUNCTIONS.has(upper) && node.args.length === 1) {
				return canCodegen(node.args[0] as FormulaNode)
			}
			if ((upper === 'TODAY' || upper === 'NOW') && node.args.length === 0) return true
			if (
				(upper === 'SUMIF' || upper === 'COUNTIF') &&
				node.args.length === 2 &&
				(node.args[0] as FormulaNode).type === 'rangeRef'
			) {
				return canCodegen(node.args[1] as FormulaNode)
			}
			if (upper === 'TEXT' && node.args.length === 2) return node.args.every(canCodegen)
			if (upper === 'TEXTJOIN' && node.args.length >= 3) return node.args.every(canCodegen)
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
	nodeHashes: Map<string, string>
}

function hashNode(node: FormulaNode): string | null {
	switch (node.type) {
		case 'cellRef':
			return `cell:${node.ref.row}:${node.ref.col}:${node.ref.rowAbsolute}:${node.ref.colAbsolute}:${node.sheet ?? ''}`
		case 'function': {
			const upper = node.name.toUpperCase()
			if (
				RANGE_AGG_FUNCTIONS.has(upper) &&
				node.args.length === 1 &&
				(node.args[0] as FormulaNode).type === 'rangeRef'
			) {
				const r = node.args[0] as FormulaNode & { type: 'rangeRef' }
				return `range:${upper}:${r.start.row}:${r.start.col}:${r.start.rowAbsolute}:${r.start.colAbsolute}:${r.end.row}:${r.end.col}:${r.end.rowAbsolute}:${r.end.colAbsolute}:${r.sheet ?? ''}`
			}
			return null
		}
		default:
			return null
	}
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

function tryFoldConstant(node: FormulaNode): number | null {
	switch (node.type) {
		case 'number':
			return node.value
		case 'boolean':
			return node.value ? 1 : 0
		case 'binary': {
			const l = tryFoldConstant(node.left)
			const r = tryFoldConstant(node.right)
			if (l === null || r === null) return null
			switch (node.op) {
				case '+':
					return l + r
				case '-':
					return l - r
				case '*':
					return l * r
				case '/':
					return r === 0 ? null : l / r
				case '^':
					return l ** r
				default:
					return null
			}
		}
		case 'unary': {
			const v = tryFoldConstant(node.operand)
			if (v === null) return null
			switch (node.op) {
				case '+':
					return v
				case '-':
					return -v
				case '%':
					return v / 100
				default:
					return null
			}
		}
		case 'function': {
			const upper = node.name.toUpperCase()
			if (upper === 'PI' && node.args.length === 0) return Math.PI
			if (node.args.length === 1) {
				const a = tryFoldConstant(node.args[0] as FormulaNode)
				if (a === null) return null
				switch (upper) {
					case 'ABS':
						return Math.abs(a)
					case 'SIGN':
						return Math.sign(a)
					case 'INT':
						return Math.floor(a)
					case 'SQRT':
						return a < 0 ? null : Math.sqrt(a)
					case 'LN':
						return a <= 0 ? null : Math.log(a)
					case 'LOG10':
						return a <= 0 ? null : Math.log10(a)
					case 'EXP':
						return Math.exp(a)
					case 'DEGREES':
						return a * (180 / Math.PI)
					case 'RADIANS':
						return a * (Math.PI / 180)
					case 'SIN':
						return Math.sin(a)
					case 'COS':
						return Math.cos(a)
					case 'TAN':
						return Math.tan(a)
					case 'ASIN':
						return a < -1 || a > 1 ? null : Math.asin(a)
					case 'ACOS':
						return a < -1 || a > 1 ? null : Math.acos(a)
					case 'ATAN':
						return Math.atan(a)
					default:
						return null
				}
			}
			if (node.args.length === 2) {
				const a = tryFoldConstant(node.args[0] as FormulaNode)
				const b = tryFoldConstant(node.args[1] as FormulaNode)
				if (a === null || b === null) return null
				switch (upper) {
					case 'MOD':
						return b === 0 ? null : a - b * Math.floor(a / b)
					case 'POWER':
						return a < 0 && b !== Math.floor(b) ? null : a ** b
					case 'ROUND': {
						const m = 10 ** Math.floor(b)
						return Math.round(a * m) / m
					}
					default:
						return null
				}
			}
			return null
		}
		default:
			return null
	}
}

function emitNode(state: CodegenState, node: FormulaNode): string {
	if (
		isConstantNode(node) &&
		node.type !== 'number' &&
		node.type !== 'string' &&
		node.type !== 'boolean' &&
		node.type !== 'error'
	) {
		const folded = tryFoldConstant(node)
		if (folded !== null && Number.isFinite(folded)) {
			const v = freshVar(state)
			state.lines.push(`var ${v} = _numberValue(${folded});`)
			return v
		}
	}

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
			const h = hashNode(node)
			if (h) {
				const cached = state.nodeHashes.get(h)
				if (cached !== undefined) return cached
			}
			const result = emitReadCell(state, node.ref, node.sheet)
			if (h) state.nodeHashes.set(h, result)
			return result
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
		const h = hashNode(node)
		if (h) {
			const cached = state.nodeHashes.get(h)
			if (cached !== undefined) return cached
		}
		const result = emitRangeAggregate(
			state,
			node.args[0] as FormulaNode & { type: 'rangeRef' },
			upper,
		)
		if (h) state.nodeHashes.set(h, result)
		return result
	}
	if (SIMPLE_MATH_FUNCTIONS.has(upper)) return emitSimpleMath(state, node)
	if (SIMPLE_TEXT_FUNCTIONS.has(upper)) return emitSimpleText(state, node)
	if (TEXT_WITH_ARGS_FUNCTIONS.has(upper)) return emitTextWithArgs(state, node)
	if (DATE_EXTRACT_FUNCTIONS.has(upper)) return emitDateExtract(state, node, upper)
	if (upper === 'TODAY') return emitToday(state)
	if (upper === 'NOW') return emitNow(state)
	if (upper === 'SUMIF') return emitSumif(state, node)
	if (upper === 'COUNTIF') return emitCountif(state, node)
	if (upper === 'TEXT') return emitTextFormat(state, node)
	if (upper === 'TEXTJOIN') return emitTextJoin(state, node)
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

function emitSimpleMath(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const upper = node.name.toUpperCase()
	if (upper === 'PI' && node.args.length === 0) {
		const result = freshVar(state)
		state.lines.push(`var ${result} = _numberValue(${Math.PI});`)
		return result
	}

	if (node.args.length === 1) {
		const argVar = emitNode(state, node.args[0] as FormulaNode)
		const sv = freshVar(state, 'sm')
		state.lines.push(`var ${sv} = _topLeft(${argVar});`)
		state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
		const nv = emitCoerceNumber(state, sv)
		const result = freshVar(state)
		switch (upper) {
			case 'ABS':
				state.lines.push(`var ${result} = _numberValue(Math.abs(${nv}));`)
				break
			case 'SIGN':
				state.lines.push(`var ${result} = _numberValue(Math.sign(${nv}));`)
				break
			case 'INT':
				state.lines.push(`var ${result} = _numberValue(Math.floor(${nv}));`)
				break
			case 'SQRT':
				state.lines.push(
					`if (${nv} < 0) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.sqrt(${nv}));`,
				)
				break
			case 'LN':
				state.lines.push(
					`if (${nv} <= 0) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.log(${nv}));`,
				)
				break
			case 'LOG10':
				state.lines.push(
					`if (${nv} <= 0) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.log10(${nv}));`,
				)
				break
			case 'EXP':
				state.lines.push(`var ${result} = _numberValue(Math.exp(${nv}));`)
				break
			case 'DEGREES':
				state.lines.push(`var ${result} = _numberValue(${nv} * ${180 / Math.PI});`)
				break
			case 'RADIANS':
				state.lines.push(`var ${result} = _numberValue(${nv} * ${Math.PI / 180});`)
				break
			case 'SIN':
				state.lines.push(`var ${result} = _numberValue(Math.sin(${nv}));`)
				break
			case 'COS':
				state.lines.push(`var ${result} = _numberValue(Math.cos(${nv}));`)
				break
			case 'TAN':
				state.lines.push(`var ${result} = _numberValue(Math.tan(${nv}));`)
				break
			case 'ASIN':
				state.lines.push(
					`if (${nv} < -1 || ${nv} > 1) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.asin(${nv}));`,
				)
				break
			case 'ACOS':
				state.lines.push(
					`if (${nv} < -1 || ${nv} > 1) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.acos(${nv}));`,
				)
				break
			case 'ATAN':
				state.lines.push(`var ${result} = _numberValue(Math.atan(${nv}));`)
				break
			case 'TRUNC':
				state.lines.push(`var ${result} = _numberValue(Math.trunc(${nv}));`)
				break
			default:
				state.lines.push(`var ${result} = _numberValue(${nv});`)
		}
		return result
	}

	if (node.args.length === 2) {
		const arg1Var = emitNode(state, node.args[0] as FormulaNode)
		const arg2Var = emitNode(state, node.args[1] as FormulaNode)
		const s1 = freshVar(state, 'sa')
		const s2 = freshVar(state, 'sb')
		state.lines.push(`var ${s1} = _topLeft(${arg1Var}); var ${s2} = _topLeft(${arg2Var});`)
		state.lines.push(
			`if (${s1}.kind === 'error') return ${s1}; if (${s2}.kind === 'error') return ${s2};`,
		)
		const n1 = emitCoerceNumber(state, s1)
		const n2 = emitCoerceNumber(state, s2)
		const result = freshVar(state)
		switch (upper) {
			case 'MOD':
				state.lines.push(
					`if (${n2} === 0) return _errorValue('#DIV/0!'); var ${result} = _numberValue(${n1} - ${n2} * Math.floor(${n1} / ${n2}));`,
				)
				break
			case 'POWER': {
				state.lines.push(
					`if (${n1} < 0 && ${n2} !== Math.floor(${n2})) return _errorValue('#NUM!'); var ${result} = _numberValue(${n1} ** ${n2});`,
				)
				break
			}
			case 'ROUNDUP': {
				state.lines.push(
					`var _m${state.varCounter} = Math.pow(10, Math.floor(${n2})); var ${result} = _numberValue(${n1} >= 0 ? Math.ceil(${n1} * _m${state.varCounter}) / _m${state.varCounter} : Math.floor(${n1} * _m${state.varCounter}) / _m${state.varCounter});`,
				)
				state.varCounter++
				break
			}
			case 'ROUNDDOWN': {
				state.lines.push(
					`var _m${state.varCounter} = Math.pow(10, Math.floor(${n2})); var ${result} = _numberValue(Math.trunc(${n1} * _m${state.varCounter}) / _m${state.varCounter});`,
				)
				state.varCounter++
				break
			}
			case 'TRUNC': {
				state.lines.push(
					`var _m${state.varCounter} = Math.pow(10, Math.floor(${n2})); var ${result} = _numberValue(Math.trunc(${n1} * _m${state.varCounter}) / _m${state.varCounter});`,
				)
				state.varCounter++
				break
			}
			case 'CEILING': {
				state.lines.push(
					`if (${n2} === 0) return _numberValue(0); if (${n1} > 0 && ${n2} < 0) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.ceil(${n1} / ${n2}) * ${n2});`,
				)
				break
			}
			case 'FLOOR': {
				state.lines.push(
					`if (${n2} === 0) return _errorValue('#DIV/0!'); if (${n1} > 0 && ${n2} < 0) return _errorValue('#NUM!'); var ${result} = _numberValue(Math.floor(${n1} / ${n2}) * ${n2});`,
				)
				break
			}
			default:
				state.lines.push(`var ${result} = _numberValue(${n1});`)
		}
		return result
	}

	return emitTreeFallback(state, node)
}

function emitSimpleText(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const upper = node.name.toUpperCase()
	const argVar = emitNode(state, node.args[0] as FormulaNode)
	const sv = freshVar(state, 'st')
	state.lines.push(`var ${sv} = _topLeft(${argVar});`)
	state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
	const str = freshVar(state, 'ts')
	state.lines.push(`var ${str} = _coerceStr(${sv});`)
	const result = freshVar(state)
	switch (upper) {
		case 'LEN':
			state.lines.push(`var ${result} = _numberValue(${str}.length);`)
			break
		case 'UPPER':
			state.lines.push(`var ${result} = _stringValue(${str}.toUpperCase());`)
			break
		case 'LOWER':
			state.lines.push(`var ${result} = _stringValue(${str}.toLowerCase());`)
			break
		case 'TRIM':
			state.lines.push(`var ${result} = _stringValue(${str}.replace(/^ +| +$|( ) +/g, '$1'));`)
			break
		case 'PROPER':
			state.lines.push(
				`var ${result} = _stringValue(${str}.toLowerCase().replace(/(?:^|[^a-zA-Z\u00C0-\u024F])([a-zA-Z\u00C0-\u024F])/g, function(m,c){return m.slice(0,-1)+c.toUpperCase();}));`,
			)
			break
		default:
			state.lines.push(`var ${result} = _stringValue(${str});`)
	}
	return result
}

function emitTextWithArgs(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const upper = node.name.toUpperCase()

	if ((upper === 'CONCATENATE' || upper === 'CONCAT') && node.args.length >= 1) {
		const parts: string[] = []
		for (const arg of node.args) {
			const argVar = emitNode(state, arg as FormulaNode)
			const sv = freshVar(state, 'cc')
			state.lines.push(`var ${sv} = _topLeft(${argVar});`)
			state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
			parts.push(`_coerceStr(${sv})`)
		}
		const result = freshVar(state)
		state.lines.push(`var ${result} = _stringValue(${parts.join(' + ')});`)
		return result
	}

	if (upper === 'LEFT' && node.args.length >= 1 && node.args.length <= 2) {
		const argVar = emitNode(state, node.args[0] as FormulaNode)
		const sv = freshVar(state, 'lt')
		state.lines.push(`var ${sv} = _topLeft(${argVar});`)
		state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
		const str = freshVar(state, 'ls')
		state.lines.push(`var ${str} = _coerceStr(${sv});`)
		let countExpr = '1'
		if (node.args.length === 2) {
			const countVar = emitNode(state, node.args[1] as FormulaNode)
			const cs = freshVar(state, 'lc')
			state.lines.push(`var ${cs} = _topLeft(${countVar});`)
			state.lines.push(`if (${cs}.kind === 'error') return ${cs};`)
			countExpr = emitCoerceNumber(state, cs)
		}
		const result = freshVar(state)
		state.lines.push(
			`if (${countExpr} < 0) return _errorValue('#VALUE!'); var ${result} = _stringValue(${str}.substring(0, ${countExpr}));`,
		)
		return result
	}

	if (upper === 'RIGHT' && node.args.length >= 1 && node.args.length <= 2) {
		const argVar = emitNode(state, node.args[0] as FormulaNode)
		const sv = freshVar(state, 'rt')
		state.lines.push(`var ${sv} = _topLeft(${argVar});`)
		state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
		const str = freshVar(state, 'rs')
		state.lines.push(`var ${str} = _coerceStr(${sv});`)
		let countExpr = '1'
		if (node.args.length === 2) {
			const countVar = emitNode(state, node.args[1] as FormulaNode)
			const cs = freshVar(state, 'rc')
			state.lines.push(`var ${cs} = _topLeft(${countVar});`)
			state.lines.push(`if (${cs}.kind === 'error') return ${cs};`)
			countExpr = emitCoerceNumber(state, cs)
		}
		const result = freshVar(state)
		state.lines.push(
			`if (${countExpr} < 0) return _errorValue('#VALUE!'); var ${result} = _stringValue(${countExpr} === 0 ? '' : ${str}.slice(-Math.trunc(${countExpr})));`,
		)
		return result
	}

	if (upper === 'MID' && node.args.length === 3) {
		const argVar = emitNode(state, node.args[0] as FormulaNode)
		const startVar = emitNode(state, node.args[1] as FormulaNode)
		const lenVar = emitNode(state, node.args[2] as FormulaNode)
		const sv = freshVar(state, 'mt')
		state.lines.push(`var ${sv} = _topLeft(${argVar});`)
		state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
		const str = freshVar(state, 'ms')
		state.lines.push(`var ${str} = _coerceStr(${sv});`)
		const ss = freshVar(state, 'ms2')
		state.lines.push(`var ${ss} = _topLeft(${startVar});`)
		state.lines.push(`if (${ss}.kind === 'error') return ${ss};`)
		const startN = emitCoerceNumber(state, ss)
		const ls = freshVar(state, 'ml')
		state.lines.push(`var ${ls} = _topLeft(${lenVar});`)
		state.lines.push(`if (${ls}.kind === 'error') return ${ls};`)
		const lenN = emitCoerceNumber(state, ls)
		const result = freshVar(state)
		state.lines.push(
			`if (${startN} < 1 || ${lenN} < 0) return _errorValue('#VALUE!'); var ${result} = _stringValue(${str}.substring(${startN} - 1, ${startN} - 1 + ${lenN}));`,
		)
		return result
	}

	if (upper === 'REPT' && node.args.length === 2) {
		const argVar = emitNode(state, node.args[0] as FormulaNode)
		const countVar = emitNode(state, node.args[1] as FormulaNode)
		const sv = freshVar(state, 'rpt')
		state.lines.push(`var ${sv} = _topLeft(${argVar});`)
		state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
		const str = freshVar(state, 'rps')
		state.lines.push(`var ${str} = _coerceStr(${sv});`)
		const cs = freshVar(state, 'rpc')
		state.lines.push(`var ${cs} = _topLeft(${countVar});`)
		state.lines.push(`if (${cs}.kind === 'error') return ${cs};`)
		const cn = emitCoerceNumber(state, cs)
		const result = freshVar(state)
		state.lines.push(
			`if (${cn} < 0) return _errorValue('#VALUE!'); var ${result} = _stringValue(${str}.repeat(Math.floor(${cn})));`,
		)
		return result
	}

	return emitTreeFallback(state, node)
}

function emitDateExtract(
	state: CodegenState,
	node: FormulaNode & { type: 'function' },
	field: string,
): string {
	const argVar = emitNode(state, node.args[0] as FormulaNode)
	const sv = freshVar(state, 'de')
	state.lines.push(`var ${sv} = _topLeft(${argVar});`)
	state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
	const nv = emitCoerceNumber(state, sv)
	const parts = freshVar(state, 'dp')
	state.lines.push(`var ${parts} = _serialToDate(Math.floor(${nv}), ctx.calcContext.dateSystem);`)
	state.lines.push(`if (!${parts}) return _errorValue('#NUM!');`)
	const result = freshVar(state)
	const prop = field === 'YEAR' ? 'year' : field === 'MONTH' ? 'month' : 'day'
	state.lines.push(`var ${result} = _numberValue(${parts}.${prop});`)
	return result
}

function emitToday(state: CodegenState): string {
	const d = freshVar(state, 'td')
	state.lines.push(`var ${d} = ctx.calcContext.today;`)
	const result = freshVar(state)
	state.lines.push(
		`var ${result} = _numberValue(_dateToSerial(${d}.getFullYear(), ${d}.getMonth() + 1, ${d}.getDate(), ctx.calcContext.dateSystem));`,
	)
	return result
}

function emitNow(state: CodegenState): string {
	const d = freshVar(state, 'nd')
	state.lines.push(`var ${d} = ctx.calcContext.now;`)
	const serial = freshVar(state, 'ns')
	state.lines.push(
		`var ${serial} = _dateToSerial(${d}.getFullYear(), ${d}.getMonth() + 1, ${d}.getDate(), ctx.calcContext.dateSystem);`,
	)
	const frac = freshVar(state, 'nf')
	state.lines.push(
		`var ${frac} = (${d}.getHours() * 3600 + ${d}.getMinutes() * 60 + ${d}.getSeconds()) / 86400;`,
	)
	const result = freshVar(state)
	state.lines.push(`var ${result} = _numberValue(${serial} + ${frac});`)
	return result
}

function emitSumif(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const rangeArg = node.args[0] as FormulaNode & { type: 'rangeRef' }
	const criteriaVar = emitNode(state, node.args[1] as FormulaNode)
	const result = freshVar(state)
	const sheetExpr = rangeArg.sheet === undefined ? 'null' : JSON.stringify(rangeArg.sheet)
	const anchorRow = state.sharedAnchor?.row ?? -1
	const anchorCol = state.sharedAnchor?.col ?? -1
	state.lines.push(
		`var ${result} = _sumifRange(ctx, _sheet, ${sheetExpr}, ${rangeArg.start.row}, ${rangeArg.start.col}, ${rangeArg.start.rowAbsolute}, ${rangeArg.start.colAbsolute}, ${rangeArg.end.row}, ${rangeArg.end.col}, ${rangeArg.end.rowAbsolute}, ${rangeArg.end.colAbsolute}, ${anchorRow}, ${anchorCol}, ${criteriaVar});`,
	)
	return result
}

function emitCountif(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const rangeArg = node.args[0] as FormulaNode & { type: 'rangeRef' }
	const criteriaVar = emitNode(state, node.args[1] as FormulaNode)
	const result = freshVar(state)
	const sheetExpr = rangeArg.sheet === undefined ? 'null' : JSON.stringify(rangeArg.sheet)
	const anchorRow = state.sharedAnchor?.row ?? -1
	const anchorCol = state.sharedAnchor?.col ?? -1
	state.lines.push(
		`var ${result} = _countifRange(ctx, _sheet, ${sheetExpr}, ${rangeArg.start.row}, ${rangeArg.start.col}, ${rangeArg.start.rowAbsolute}, ${rangeArg.start.colAbsolute}, ${rangeArg.end.row}, ${rangeArg.end.col}, ${rangeArg.end.rowAbsolute}, ${rangeArg.end.colAbsolute}, ${anchorRow}, ${anchorCol}, ${criteriaVar});`,
	)
	return result
}

function emitTextFormat(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const valVar = emitNode(state, node.args[0] as FormulaNode)
	const fmtVar = emitNode(state, node.args[1] as FormulaNode)
	const vs = freshVar(state, 'tv')
	state.lines.push(`var ${vs} = _topLeft(${valVar});`)
	state.lines.push(`if (${vs}.kind === 'error') return ${vs};`)
	const nv = emitCoerceNumber(state, vs)
	const fs = freshVar(state, 'tf')
	state.lines.push(`var ${fs} = _topLeft(${fmtVar});`)
	state.lines.push(`if (${fs}.kind === 'error') return ${fs};`)
	const fmtStr = freshVar(state, 'tfs')
	state.lines.push(`var ${fmtStr} = _coerceStr(${fs});`)
	const result = freshVar(state)
	state.lines.push(`var ${result} = _stringValue(_formatNumber(${nv}, ${fmtStr}));`)
	return result
}

function emitTextJoin(state: CodegenState, node: FormulaNode & { type: 'function' }): string {
	const delimVar = emitNode(state, node.args[0] as FormulaNode)
	const ds = freshVar(state, 'tjd')
	state.lines.push(`var ${ds} = _topLeft(${delimVar});`)
	state.lines.push(`if (${ds}.kind === 'error') return ${ds};`)
	const delimStr = freshVar(state, 'tds')
	state.lines.push(`var ${delimStr} = _coerceStr(${ds});`)

	const ieVar = emitNode(state, node.args[1] as FormulaNode)
	const ies = freshVar(state, 'tji')
	state.lines.push(`var ${ies} = _topLeft(${ieVar});`)
	state.lines.push(`if (${ies}.kind === 'error') return ${ies};`)
	const ieBool = freshVar(state, 'tjib')
	state.lines.push(
		`var ${ieBool} = ${ies}.kind === 'boolean' ? ${ies}.value : ${ies}.kind !== 'empty';`,
	)

	const partsVar = freshVar(state, 'tjp')
	state.lines.push(`var ${partsVar} = [];`)

	for (let i = 2; i < node.args.length; i++) {
		const argVar = emitNode(state, node.args[i] as FormulaNode)
		const sv = freshVar(state, 'tjs')
		state.lines.push(`var ${sv} = _topLeft(${argVar});`)
		state.lines.push(`if (${sv}.kind === 'error') return ${sv};`)
		state.lines.push(
			`if (!${ieBool} || (${sv}.kind !== 'empty' && !(${sv}.kind === 'string' && ${sv}.value === ''))) ${partsVar}.push(_coerceStr(${sv}));`,
		)
	}

	const result = freshVar(state)
	state.lines.push(`var ${result} = _stringValue(${partsVar}.join(${delimStr}));`)
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
		nodeHashes: new Map(),
		...(state.sharedAnchor !== undefined && { sharedAnchor: state.sharedAnchor }),
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
		nodeHashes: new Map(),
		...(state.sharedAnchor !== undefined && { sharedAnchor: state.sharedAnchor }),
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

	if (func === 'COUNTA') {
		let countA = 0
		for (let row = fromRow; row <= toRow; row++) {
			for (let col = fromCol; col <= toCol; col++) {
				const kind = targetSheet.cells.readKind(row, col)
				if (kind === undefined || kind === 'empty') continue
				if (kind === 'error') {
					return errorValue(targetSheet.cells.readError(row, col) ?? '#VALUE!')
				}
				countA++
			}
		}
		return numberValue(countA)
	}

	if (
		func === 'SUM' ||
		func === 'AVERAGE' ||
		func === 'COUNT' ||
		func === 'MIN' ||
		func === 'MAX'
	) {
		const r = aggregateNumericRange(targetSheet, fromRow, fromCol, toRow, toCol)
		if (r.error) return r.error
		switch (func) {
			case 'SUM':
				return numberValue(r.sum)
			case 'AVERAGE':
				return r.count === 0 ? errorValue('#DIV/0!') : numberValue(r.sum / r.count)
			case 'COUNT':
				return numberValue(r.count)
			case 'MIN':
				return r.count === 0 ? numberValue(0) : numberValue(r.min)
			case 'MAX':
				return r.count === 0 ? numberValue(0) : numberValue(r.max)
			default:
				return numberValue(r.sum)
		}
	}

	let acc = 0
	for (let row = fromRow; row <= toRow; row++) {
		for (let col = fromCol; col <= toCol; col++) {
			const kind = targetSheet.cells.readKind(row, col)
			if (kind === undefined || kind === 'empty') continue
			if (kind === 'error') {
				return errorValue(targetSheet.cells.readError(row, col) ?? '#VALUE!')
			}
			if (kind !== 'number' && kind !== 'date') continue
			acc += targetSheet.cells.readNumber(row, col) ?? 0
		}
	}
	return numberValue(acc)
}

function lookupValueKey(value: CellValue): string | null {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'number':
			return `v:${String(scalar.value === 0 ? 0 : scalar.value)}`
		case 'date':
			return `v:${String(scalar.serial === 0 ? 0 : scalar.serial)}`
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

function hasWildcardCriteria(text: string): boolean {
	for (let i = 0; i < text.length; i++) {
		if (text[i] === '~') {
			i++
			continue
		}
		if (text[i] === '*' || text[i] === '?') return true
	}
	return false
}

function buildCriteriaMatcher(criteria: CellValue): (v: CellValue) => boolean {
	if (criteria.kind === 'number') {
		const t = criteria.value
		return (v) => {
			const n = toNumber(v)
			return n !== null && n === t
		}
	}
	if (criteria.kind === 'boolean') {
		const t = criteria.value
		return (v) => v.kind === 'boolean' && v.value === t
	}
	if (criteria.kind !== 'string') return () => false

	const s = criteria.value
	let op = ''
	let rest = s
	for (const prefix of ['>=', '<=', '<>', '>', '<', '='] as const) {
		if (s.startsWith(prefix)) {
			op = prefix
			rest = s.slice(prefix.length)
			break
		}
	}

	const numRest = Number(rest)
	const isNumeric = rest.trim() !== '' && !Number.isNaN(numRest)

	const hasWild = hasWildcardCriteria(rest)

	if (!op) {
		if (s === '') return (v) => v.kind === 'empty' || (v.kind === 'string' && v.value === '')
		if (hasWild) return (v) => v.kind === 'string' && wildcardMatch(s, v.value)
		const lower = s.toLowerCase()
		return (v) => {
			if (v.kind === 'string') return v.value.toLowerCase() === lower
			if (isNumeric && v.kind === 'number') return v.value === numRest
			if (v.kind === 'boolean') return v.value === (lower === 'true')
			return false
		}
	}

	if (isNumeric) {
		return (v) => {
			const n = toNumber(v)
			if (n === null) return false
			switch (op) {
				case '>=':
					return n >= numRest
				case '<=':
					return n <= numRest
				case '>':
					return n > numRest
				case '<':
					return n < numRest
				case '<>':
					return n !== numRest
				case '=':
					return n === numRest
				default:
					return false
			}
		}
	}

	return (v) => {
		if (op === '=' && rest === '')
			return v.kind === 'empty' || (v.kind === 'string' && v.value === '')
		if (op === '<>' && rest === '')
			return v.kind !== 'empty' && !(v.kind === 'string' && v.value === '')
		if (v.kind !== 'string') return false
		if (hasWild) {
			const m = wildcardMatch(rest, v.value)
			return op === '=' ? m : !m
		}
		const lower = rest.toLowerCase()
		if (op === '=') return v.value.toLowerCase() === lower
		if (op === '<>') return v.value.toLowerCase() !== lower
		return false
	}
}

function resolveRangeSheet(
	ctx: EvalContext,
	currentSheet: import('@ascend/core').Workbook['sheets'][number] | undefined,
	sheetName: string | null,
): import('@ascend/core').Workbook['sheets'][number] | undefined {
	if (sheetName === null) return currentSheet
	const si = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
	if (si < 0) return undefined
	return ctx.workbook.sheets[si]
}

function sumifRange(
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
	criteriaVal: CellValue,
): CellValue {
	const sr = resolveRelativeIndex(startRow, startRowAbsolute, ctx.row, anchorRow)
	const sc = resolveRelativeIndex(startCol, startColAbsolute, ctx.col, anchorCol)
	const er = resolveRelativeIndex(endRow, endRowAbsolute, ctx.row, anchorRow)
	const ec = resolveRelativeIndex(endCol, endColAbsolute, ctx.col, anchorCol)
	if (sr < 0 || sc < 0 || er < 0 || ec < 0) return errorValue('#REF!')

	const sheet = resolveRangeSheet(ctx, currentSheet, sheetName)
	if (!sheet) return errorValue('#REF!')

	const fromRow = Math.min(sr, er)
	const toRow = Math.max(sr, er)
	const fromCol = Math.min(sc, ec)
	const toCol = Math.max(sc, ec)

	const match = buildCriteriaMatcher(topLeftScalar(criteriaVal))
	let sum = 0
	for (let row = fromRow; row <= toRow; row++) {
		for (let col = fromCol; col <= toCol; col++) {
			const cell = sheet.cells.readValue(row, col)
			if (match(cell)) {
				const n = toNumber(cell)
				if (n !== null) sum += n
			}
		}
	}
	return numberValue(sum)
}

function countifRange(
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
	criteriaVal: CellValue,
): CellValue {
	const sr = resolveRelativeIndex(startRow, startRowAbsolute, ctx.row, anchorRow)
	const sc = resolveRelativeIndex(startCol, startColAbsolute, ctx.col, anchorCol)
	const er = resolveRelativeIndex(endRow, endRowAbsolute, ctx.row, anchorRow)
	const ec = resolveRelativeIndex(endCol, endColAbsolute, ctx.col, anchorCol)
	if (sr < 0 || sc < 0 || er < 0 || ec < 0) return errorValue('#REF!')

	const sheet = resolveRangeSheet(ctx, currentSheet, sheetName)
	if (!sheet) return errorValue('#REF!')

	const fromRow = Math.min(sr, er)
	const toRow = Math.max(sr, er)
	const fromCol = Math.min(sc, ec)
	const toCol = Math.max(sc, ec)

	const match = buildCriteriaMatcher(topLeftScalar(criteriaVal))
	let count = 0
	for (let row = fromRow; row <= toRow; row++) {
		for (let col = fromCol; col <= toCol; col++) {
			if (match(sheet.cells.readValue(row, col))) count++
		}
	}
	return numberValue(count)
}

function buildCodegenFn(node: FormulaNode, sharedAnchor?: SharedAnchor): CodegenFn | null {
	if (!canCodegen(node)) return null

	const state: CodegenState = {
		lines: [],
		varCounter: 0,
		closureVars: new Map(),
		treeNodes: new Map(),
		nodeHashes: new Map(),
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
		'_serialToDate',
		'_dateToSerial',
		'_formatNumber',
		'_sumifRange',
		'_countifRange',
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
		serialToDate,
		dateToSerial,
		formatNumber,
		sumifRange,
		countifRange,
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

export function codegenSharedFormula(
	formulaText: string,
	ast: FormulaNode,
	anchor: SharedAnchor,
): CodegenFn | null {
	const cacheKey = `${formulaText}@${anchor.row}:${anchor.col}`
	const cached = sharedCodegenCache.get(cacheKey)
	if (cached !== undefined) return cached

	const fn = buildCodegenFn(ast, anchor)
	sharedCodegenCache.set(cacheKey, fn)
	return fn
}

export function clearCodegenCache(): void {
	codegenCache.clear()
	sharedCodegenCache.clear()
}
