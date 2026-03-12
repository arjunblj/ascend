import type { FormulaNode } from '@ascend/formulas'
import { toNumber } from '@ascend/formulas'
import type { CellValue, ExcelError } from '@ascend/schema'
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

const Op = {
	NUM: 1,
	STR: 2,
	BOOL: 3,
	ERR: 4,
	EMPTY_VAL: 5,
	CELL: 10,
	CELL_SHEET: 11,
	ADD: 20,
	SUB: 21,
	MUL: 22,
	DIV: 23,
	POW: 24,
	NEG: 25,
	PCT: 26,
	CONCAT: 27,
	EQ: 30,
	NE: 31,
	LT: 32,
	GT: 33,
	LE: 34,
	GE: 35,
	IF: 40,
	JMP: 41,
	IFERROR_JMP: 42,
	IFNA_JMP: 43,
	CELL_ADD: 50,
	CELL_SUB: 51,
	CELL_MUL: 52,
	TREE: 99,
} as const

export interface CompiledFormula {
	readonly ops: number[]
	readonly constants: readonly (string | number | boolean | CellValue | FormulaNode)[]
	readonly numericOnly: boolean
}

const coerceStr = coerceCellValueToString

function comparePrimitive<T extends number | string>(op: number, a: T, b: T): boolean {
	switch (op) {
		case Op.EQ:
			return a === b
		case Op.NE:
			return a !== b
		case Op.LT:
			return a < b
		case Op.GT:
			return a > b
		case Op.LE:
			return a <= b
		case Op.GE:
			return a >= b
		default:
			return false
	}
}

function evalCmp(op: number, left: CellValue, right: CellValue): boolean {
	const ln = toNumber(left)
	const rn = toNumber(right)

	if (ln !== null && rn !== null) {
		return comparePrimitive(op, ln, rn)
	}

	if (left.kind === 'string' || right.kind === 'string') {
		return comparePrimitive(op, coerceStr(left).toLowerCase(), coerceStr(right).toLowerCase())
	}

	if (left.kind === 'boolean' && right.kind === 'boolean') {
		return comparePrimitive(op, left.value ? 1 : 0, right.value ? 1 : 0)
	}

	return comparePrimitive(op, coerceStr(left), coerceStr(right))
}

function coerceToBoolForIf(v: CellValue): boolean | CellValue {
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

const COMPILABLE_FUNCTIONS = new Set(['IF', 'IFERROR', 'IFNA'])

function shouldCompile(node: FormulaNode): boolean {
	let compilableNodes = 0
	function scan(n: FormulaNode): boolean {
		switch (n.type) {
			case 'number':
			case 'string':
			case 'boolean':
			case 'error':
			case 'missing':
				compilableNodes++
				return true
			case 'cellRef':
				compilableNodes++
				return true
			case 'binary':
				if (n.op === ',' || n.op === ' ') return true
				compilableNodes++
				scan(n.left)
				scan(n.right)
				return true
			case 'unary':
				if (n.op === '@') return true
				compilableNodes++
				scan(n.operand)
				return true
			case 'function': {
				const upper = n.name.toUpperCase()
				if (COMPILABLE_FUNCTIONS.has(upper)) {
					compilableNodes++
					for (const arg of n.args) scan(arg)
				}
				return true
			}
			default:
				return true
		}
	}
	scan(node)
	return compilableNodes >= 3
}

function isNumericFormula(ops: readonly number[]): boolean {
	let ip = 0
	while (ip < ops.length) {
		const op = ops[ip++] as number
		switch (op) {
			case Op.NUM:
				ip++
				break
			case Op.EMPTY_VAL:
			case Op.ADD:
			case Op.SUB:
			case Op.MUL:
			case Op.DIV:
			case Op.POW:
			case Op.NEG:
			case Op.PCT:
			case Op.EQ:
			case Op.NE:
			case Op.LT:
			case Op.GT:
			case Op.LE:
			case Op.GE:
				break
			case Op.CELL:
			case Op.CELL_ADD:
			case Op.CELL_SUB:
			case Op.CELL_MUL:
				ip += 2
				break
			case Op.CELL_SHEET:
				ip += 3
				break
			default:
				return false
		}
	}
	return true
}

export function compileFormula(node: FormulaNode): CompiledFormula | null {
	if (!shouldCompile(node)) return null

	const ops: number[] = []
	const constants: (string | number | boolean | CellValue | FormulaNode)[] = []

	function addConst(val: string | number | boolean | CellValue | FormulaNode): number {
		constants.push(val)
		return constants.length - 1
	}

	function tryFoldConstant(n: FormulaNode): number | null {
		if (n.type === 'number') return n.value
		if (n.type === 'boolean') return n.value ? 1 : 0
		if (n.type === 'unary' && n.op === '-') {
			const inner = tryFoldConstant(n.operand)
			if (inner !== null) return -inner
		}
		if (n.type === 'unary' && n.op === '%') {
			const inner = tryFoldConstant(n.operand)
			if (inner !== null) return inner / 100
		}
		if (n.type === 'binary') {
			const left = tryFoldConstant(n.left)
			const right = tryFoldConstant(n.right)
			if (left !== null && right !== null) {
				switch (n.op) {
					case '+':
						return left + right
					case '-':
						return left - right
					case '*':
						return left * right
					case '/':
						return right !== 0 ? left / right : null
					case '^':
						return left ** right
				}
			}
		}
		return null
	}

	function emit(n: FormulaNode): void {
		switch (n.type) {
			case 'number':
				ops.push(Op.NUM, addConst(n.value))
				break
			case 'string':
				ops.push(Op.STR, addConst(n.value))
				break
			case 'boolean':
				ops.push(Op.BOOL, addConst(n.value))
				break
			case 'error':
				ops.push(Op.ERR, addConst(n.value))
				break
			case 'missing':
				ops.push(Op.EMPTY_VAL)
				break
			case 'cellRef':
				if (n.sheet === undefined) {
					ops.push(Op.CELL, n.ref.row, n.ref.col)
				} else {
					ops.push(Op.CELL_SHEET, addConst(n.sheet), n.ref.row, n.ref.col)
				}
				break
			case 'binary': {
				if (n.op === ',' || n.op === ' ') {
					ops.push(Op.TREE, addConst(n))
					break
				}
				const folded = tryFoldConstant(n)
				if (folded !== null) {
					ops.push(Op.NUM, addConst(folded))
					break
				}
				const superOp =
					n.op === '+' ? Op.CELL_ADD : n.op === '-' ? Op.CELL_SUB : n.op === '*' ? Op.CELL_MUL : 0
				if (superOp) {
					if (n.right.type === 'cellRef' && n.right.sheet === undefined) {
						emit(n.left)
						ops.push(superOp, n.right.ref.row, n.right.ref.col)
						break
					}
					if (n.left.type === 'cellRef' && n.left.sheet === undefined && n.op !== '-') {
						emit(n.right)
						ops.push(superOp, n.left.ref.row, n.left.ref.col)
						break
					}
				}
				emit(n.left)
				emit(n.right)
				switch (n.op) {
					case '+':
						ops.push(Op.ADD)
						break
					case '-':
						ops.push(Op.SUB)
						break
					case '*':
						ops.push(Op.MUL)
						break
					case '/':
						ops.push(Op.DIV)
						break
					case '^':
						ops.push(Op.POW)
						break
					case '&':
						ops.push(Op.CONCAT)
						break
					case '=':
						ops.push(Op.EQ)
						break
					case '<>':
						ops.push(Op.NE)
						break
					case '<':
						ops.push(Op.LT)
						break
					case '>':
						ops.push(Op.GT)
						break
					case '<=':
						ops.push(Op.LE)
						break
					case '>=':
						ops.push(Op.GE)
						break
				}
				break
			}
			case 'unary': {
				if (n.op === '@') {
					ops.push(Op.TREE, addConst(n))
					break
				}
				const foldedUnary = tryFoldConstant(n)
				if (foldedUnary !== null) {
					ops.push(Op.NUM, addConst(foldedUnary))
					break
				}
				emit(n.operand)
				switch (n.op) {
					case '-':
						ops.push(Op.NEG)
						break
					case '+':
						break
					case '%':
						ops.push(Op.PCT)
						break
				}
				break
			}
			case 'function':
				emitFunction(n)
				break
			default:
				ops.push(Op.TREE, addConst(n))
		}
	}

	function emitFunction(n: FormulaNode & { type: 'function' }): void {
		const upper = n.name.toUpperCase()
		if (upper === 'IF' && n.args.length >= 2 && n.args.length <= 3) {
			emit(n.args[0] as FormulaNode)
			const ifPos = ops.length
			ops.push(Op.IF, 0, 0)
			emit(n.args[1] as FormulaNode)
			const jmpPos = ops.length
			ops.push(Op.JMP, 0)
			ops[ifPos + 1] = ops.length
			if (n.args.length >= 3) {
				emit(n.args[2] as FormulaNode)
			} else {
				ops.push(Op.BOOL, addConst(false))
			}
			ops[ifPos + 2] = ops.length
			ops[jmpPos + 1] = ops.length
			return
		}
		if (upper === 'IFERROR' && n.args.length === 2) {
			emit(n.args[0] as FormulaNode)
			const jmpPos = ops.length
			ops.push(Op.IFERROR_JMP, 0)
			emit(n.args[1] as FormulaNode)
			ops[jmpPos + 1] = ops.length
			return
		}
		if (upper === 'IFNA' && n.args.length === 2) {
			emit(n.args[0] as FormulaNode)
			const jmpPos = ops.length
			ops.push(Op.IFNA_JMP, 0)
			emit(n.args[1] as FormulaNode)
			ops[jmpPos + 1] = ops.length
			return
		}
		ops.push(Op.TREE, addConst(n))
	}

	emit(node)
	return { ops, constants, numericOnly: isNumericFormula(ops) }
}

const NUMERIC_STACK_SIZE = 64
const numericStack = new Float64Array(NUMERIC_STACK_SIZE)

function evaluateCompiledNumeric(compiled: CompiledFormula, ctx: EvalContext): CellValue {
	const { ops, constants } = compiled
	const sheet = ctx.workbook.sheets[ctx.sheetIndex]
	let ip = 0
	let sp = 0
	const len = ops.length

	while (ip < len) {
		const op = ops[ip] as number
		ip++
		switch (op) {
			case Op.NUM:
				numericStack[sp++] = constants[ops[ip++] as number] as number
				break
			case Op.EMPTY_VAL:
				numericStack[sp++] = 0
				break
			case Op.CELL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				numericStack[sp++] = toNumber(sheet?.cells.readValue(row, col) ?? EMPTY) ?? NaN
				break
			}
			case Op.CELL_SHEET: {
				const nameIdx = ops[ip] as number
				const row = ops[ip + 1] as number
				const col = ops[ip + 2] as number
				ip += 3
				const si = resolveSheetIdx(ctx.workbook, constants[nameIdx] as string, ctx.sheetIndex)
				if (si < 0) {
					numericStack[sp++] = NaN
				} else {
					numericStack[sp++] =
						toNumber(ctx.workbook.sheets[si]?.cells.readValue(row, col) ?? EMPTY) ?? NaN
				}
				break
			}
			case Op.ADD: {
				const b = numericStack[--sp] as number
				numericStack[sp - 1] = (numericStack[sp - 1] as number) + b
				break
			}
			case Op.SUB: {
				const b = numericStack[--sp] as number
				numericStack[sp - 1] = (numericStack[sp - 1] as number) - b
				break
			}
			case Op.MUL: {
				const b = numericStack[--sp] as number
				numericStack[sp - 1] = (numericStack[sp - 1] as number) * b
				break
			}
			case Op.DIV: {
				const b = numericStack[--sp] as number
				numericStack[sp - 1] = b === 0 ? NaN : (numericStack[sp - 1] as number) / b
				break
			}
			case Op.POW: {
				const b = numericStack[--sp] as number
				numericStack[sp - 1] = (numericStack[sp - 1] as number) ** b
				break
			}
			case Op.NEG:
				numericStack[sp - 1] = -(numericStack[sp - 1] as number)
				break
			case Op.PCT:
				numericStack[sp - 1] = (numericStack[sp - 1] as number) / 100
				break
			case Op.CELL_ADD:
			case Op.CELL_SUB:
			case Op.CELL_MUL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				const cn = toNumber(sheet?.cells.readValue(row, col) ?? EMPTY) ?? NaN
				const left = numericStack[sp - 1] as number
				numericStack[sp - 1] =
					op === Op.CELL_ADD ? left + cn : op === Op.CELL_SUB ? left - cn : left * cn
				break
			}
			case Op.EQ:
			case Op.NE:
			case Op.LT:
			case Op.GT:
			case Op.LE:
			case Op.GE: {
				const rb = numericStack[--sp] as number
				const la = numericStack[sp - 1] as number
				let cmp: boolean
				switch (op) {
					case Op.EQ:
						cmp = la === rb
						break
					case Op.NE:
						cmp = la !== rb
						break
					case Op.LT:
						cmp = la < rb
						break
					case Op.GT:
						cmp = la > rb
						break
					case Op.LE:
						cmp = la <= rb
						break
					default:
						cmp = la >= rb
						break
				}
				numericStack[sp - 1] = cmp ? 1 : 0
				break
			}
		}
	}

	const result = numericStack[0] as number
	if (!Number.isFinite(result)) return errorValue('#VALUE!')
	return numberValue(result)
}

const STACK_SIZE = 64
const sharedStack: CellValue[] = new Array(STACK_SIZE)
let stackDepth = 0

export function evaluateCompiled(compiled: CompiledFormula, ctx: EvalContext): CellValue {
	if (compiled.numericOnly) {
		return evaluateCompiledNumeric(compiled, ctx)
	}

	const { ops, constants } = compiled
	const baseDepth = stackDepth
	const sheet = ctx.workbook.sheets[ctx.sheetIndex]
	let ip = 0
	const len = ops.length

	while (ip < len) {
		const op = ops[ip] as number
		ip++
		switch (op) {
			case Op.NUM: {
				const idx = ops[ip] as number
				ip++
				sharedStack[stackDepth++] = numberValue(constants[idx] as number)
				break
			}
			case Op.STR: {
				const idx = ops[ip] as number
				ip++
				sharedStack[stackDepth++] = stringValue(constants[idx] as string)
				break
			}
			case Op.BOOL: {
				const idx = ops[ip] as number
				ip++
				sharedStack[stackDepth++] = booleanValue(constants[idx] as boolean)
				break
			}
			case Op.ERR: {
				const idx = ops[ip] as number
				ip++
				sharedStack[stackDepth++] = errorValue(constants[idx] as ExcelError)
				break
			}
			case Op.EMPTY_VAL:
				sharedStack[stackDepth++] = EMPTY
				break
			case Op.CELL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				sharedStack[stackDepth++] = sheet?.cells.readValue(row, col) ?? EMPTY
				break
			}
			case Op.CELL_SHEET: {
				const nameIdx = ops[ip] as number
				const row = ops[ip + 1] as number
				const col = ops[ip + 2] as number
				ip += 3
				const sheetName = constants[nameIdx] as string
				const si = resolveSheetIdx(ctx.workbook, sheetName, ctx.sheetIndex)
				if (si < 0) {
					sharedStack[stackDepth++] = errorValue('#REF!')
				} else {
					const target = ctx.workbook.sheets[si]
					sharedStack[stackDepth++] = target?.cells.readValue(row, col) ?? EMPTY
				}
				break
			}
			case Op.ADD:
			case Op.SUB:
			case Op.MUL:
			case Op.DIV:
			case Op.POW: {
				const right = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				const left = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				if (left.kind === 'error') {
					sharedStack[stackDepth++] = left
					break
				}
				if (right.kind === 'error') {
					sharedStack[stackDepth++] = right
					break
				}
				const ln = toNumber(left)
				const rn = toNumber(right)
				if (ln === null || rn === null) {
					sharedStack[stackDepth++] = errorValue('#VALUE!')
					break
				}
				switch (op) {
					case Op.ADD:
						sharedStack[stackDepth++] = numberValue(ln + rn)
						break
					case Op.SUB:
						sharedStack[stackDepth++] = numberValue(ln - rn)
						break
					case Op.MUL:
						sharedStack[stackDepth++] = numberValue(ln * rn)
						break
					case Op.DIV:
						sharedStack[stackDepth++] = rn === 0 ? errorValue('#DIV/0!') : numberValue(ln / rn)
						break
					case Op.POW:
						sharedStack[stackDepth++] = numberValue(ln ** rn)
						break
				}
				break
			}
			case Op.CELL_ADD:
			case Op.CELL_SUB:
			case Op.CELL_MUL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				const cellVal = sheet?.cells.readValue(row, col) ?? EMPTY
				const left = sharedStack[stackDepth - 1] ?? EMPTY
				const ln = toNumber(left)
				const cn = toNumber(cellVal)
				if (ln !== null && cn !== null) {
					sharedStack[stackDepth - 1] = numberValue(
						op === Op.CELL_ADD ? ln + cn : op === Op.CELL_SUB ? ln - cn : ln * cn,
					)
				} else {
					const sv = topLeftScalar(left)
					const cv = topLeftScalar(cellVal)
					if (sv.kind === 'error') sharedStack[stackDepth - 1] = sv
					else if (cv.kind === 'error') sharedStack[stackDepth - 1] = cv
					else sharedStack[stackDepth - 1] = errorValue('#VALUE!')
				}
				break
			}
			case Op.NEG: {
				const v = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				if (v.kind === 'error') {
					sharedStack[stackDepth++] = v
					break
				}
				const n = toNumber(v)
				sharedStack[stackDepth++] = n === null ? errorValue('#VALUE!') : numberValue(-n)
				break
			}
			case Op.PCT: {
				const v = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				if (v.kind === 'error') {
					sharedStack[stackDepth++] = v
					break
				}
				const n = toNumber(v)
				sharedStack[stackDepth++] = n === null ? errorValue('#VALUE!') : numberValue(n / 100)
				break
			}
			case Op.CONCAT: {
				const right = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				const left = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				if (left.kind === 'error') {
					sharedStack[stackDepth++] = left
					break
				}
				if (right.kind === 'error') {
					sharedStack[stackDepth++] = right
					break
				}
				sharedStack[stackDepth++] = stringValue(coerceStr(left) + coerceStr(right))
				break
			}
			case Op.EQ:
			case Op.NE:
			case Op.LT:
			case Op.GT:
			case Op.LE:
			case Op.GE: {
				const right = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				const left = topLeftScalar(sharedStack[--stackDepth] ?? EMPTY)
				if (left.kind === 'error') {
					sharedStack[stackDepth++] = left
					break
				}
				if (right.kind === 'error') {
					sharedStack[stackDepth++] = right
					break
				}
				sharedStack[stackDepth++] = booleanValue(evalCmp(op, left, right))
				break
			}
			case Op.IF: {
				const falseTarget = ops[ip] as number
				const endTarget = ops[ip + 1] as number
				ip += 2
				const cond = coerceToBoolForIf(sharedStack[--stackDepth] ?? EMPTY)
				if (typeof cond !== 'boolean') {
					sharedStack[stackDepth++] = cond
					ip = endTarget
				} else if (!cond) {
					ip = falseTarget
				}
				break
			}
			case Op.JMP: {
				ip = ops[ip] as number
				break
			}
			case Op.IFERROR_JMP: {
				const endTarget = ops[ip] as number
				ip++
				const v = topLeftScalar(sharedStack[stackDepth - 1] ?? EMPTY)
				if (v.kind !== 'error') {
					ip = endTarget
				} else {
					stackDepth--
				}
				break
			}
			case Op.IFNA_JMP: {
				const endTarget = ops[ip] as number
				ip++
				const v = topLeftScalar(sharedStack[stackDepth - 1] ?? EMPTY)
				if (!(v.kind === 'error' && v.value === '#N/A')) {
					ip = endTarget
				} else {
					stackDepth--
				}
				break
			}
			case Op.TREE: {
				const idx = ops[ip] as number
				ip++
				const node = constants[idx] as FormulaNode
				sharedStack[stackDepth++] = treeEvaluate(node, ctx)
				break
			}
		}
	}
	const result = sharedStack[baseDepth]
	stackDepth = baseDepth
	return result ?? EMPTY
}

const sheetIdxCache = new WeakMap<import('@ascend/core').Workbook, Map<string, number>>()

function resolveSheetIdx(
	wb: import('@ascend/core').Workbook,
	sheetName: string,
	_currentSheet: number,
): number {
	let cache = sheetIdxCache.get(wb)
	if (!cache) {
		cache = new Map()
		for (let i = 0; i < wb.sheets.length; i++) {
			const s = wb.sheets[i]
			if (s) cache.set(s.name.toLowerCase(), i)
		}
		sheetIdxCache.set(wb, cache)
	}
	return cache.get(sheetName.toLowerCase()) ?? -1
}
