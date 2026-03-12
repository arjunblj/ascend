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
	TREE: 99,
} as const

export interface CompiledFormula {
	readonly ops: number[]
	readonly constants: readonly (string | number | boolean | CellValue | FormulaNode)[]
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
				if (n.sheet === undefined) compilableNodes++
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
			default:
				return true
		}
	}
	scan(node)
	return compilableNodes >= 5
}

export function compileFormula(node: FormulaNode): CompiledFormula | null {
	if (!shouldCompile(node)) return null

	const ops: number[] = []
	const constants: (string | number | boolean | CellValue | FormulaNode)[] = []

	function addConst(val: string | number | boolean | CellValue | FormulaNode): number {
		constants.push(val)
		return constants.length - 1
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
					ops.push(Op.TREE, addConst(n))
				}
				break
			case 'binary':
				if (n.op === ',' || n.op === ' ') {
					ops.push(Op.TREE, addConst(n))
					break
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
			case 'unary':
				if (n.op === '@') {
					ops.push(Op.TREE, addConst(n))
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
			default:
				ops.push(Op.TREE, addConst(n))
		}
	}

	emit(node)
	return { ops, constants }
}

export function evaluateCompiled(compiled: CompiledFormula, ctx: EvalContext): CellValue {
	const { ops, constants } = compiled
	const stack: CellValue[] = []
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
				stack.push(numberValue(constants[idx] as number))
				break
			}
			case Op.STR: {
				const idx = ops[ip] as number
				ip++
				stack.push(stringValue(constants[idx] as string))
				break
			}
			case Op.BOOL: {
				const idx = ops[ip] as number
				ip++
				stack.push(booleanValue(constants[idx] as boolean))
				break
			}
			case Op.ERR: {
				const idx = ops[ip] as number
				ip++
				stack.push(errorValue(constants[idx] as ExcelError))
				break
			}
			case Op.EMPTY_VAL:
				stack.push(EMPTY)
				break
			case Op.CELL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				stack.push(sheet?.cells.getValue(row, col) ?? EMPTY)
				break
			}
			case Op.ADD:
			case Op.SUB:
			case Op.MUL:
			case Op.DIV:
			case Op.POW: {
				const right = topLeftScalar(stack.pop() ?? EMPTY)
				const left = topLeftScalar(stack.pop() ?? EMPTY)
				if (left.kind === 'error') {
					stack.push(left)
					break
				}
				if (right.kind === 'error') {
					stack.push(right)
					break
				}
				const ln = toNumber(left)
				const rn = toNumber(right)
				if (ln === null || rn === null) {
					stack.push(errorValue('#VALUE!'))
					break
				}
				switch (op) {
					case Op.ADD:
						stack.push(numberValue(ln + rn))
						break
					case Op.SUB:
						stack.push(numberValue(ln - rn))
						break
					case Op.MUL:
						stack.push(numberValue(ln * rn))
						break
					case Op.DIV:
						stack.push(rn === 0 ? errorValue('#DIV/0!') : numberValue(ln / rn))
						break
					case Op.POW:
						stack.push(numberValue(ln ** rn))
						break
				}
				break
			}
			case Op.NEG: {
				const v = topLeftScalar(stack.pop() ?? EMPTY)
				if (v.kind === 'error') {
					stack.push(v)
					break
				}
				const n = toNumber(v)
				stack.push(n === null ? errorValue('#VALUE!') : numberValue(-n))
				break
			}
			case Op.PCT: {
				const v = topLeftScalar(stack.pop() ?? EMPTY)
				if (v.kind === 'error') {
					stack.push(v)
					break
				}
				const n = toNumber(v)
				stack.push(n === null ? errorValue('#VALUE!') : numberValue(n / 100))
				break
			}
			case Op.CONCAT: {
				const right = topLeftScalar(stack.pop() ?? EMPTY)
				const left = topLeftScalar(stack.pop() ?? EMPTY)
				if (left.kind === 'error') {
					stack.push(left)
					break
				}
				if (right.kind === 'error') {
					stack.push(right)
					break
				}
				stack.push(stringValue(coerceStr(left) + coerceStr(right)))
				break
			}
			case Op.EQ:
			case Op.NE:
			case Op.LT:
			case Op.GT:
			case Op.LE:
			case Op.GE: {
				const right = topLeftScalar(stack.pop() ?? EMPTY)
				const left = topLeftScalar(stack.pop() ?? EMPTY)
				if (left.kind === 'error') {
					stack.push(left)
					break
				}
				if (right.kind === 'error') {
					stack.push(right)
					break
				}
				stack.push(booleanValue(evalCmp(op, left, right)))
				break
			}
			case Op.TREE: {
				const idx = ops[ip] as number
				ip++
				const node = constants[idx] as FormulaNode
				stack.push(treeEvaluate(node, ctx))
				break
			}
		}
	}
	return stack[0] ?? EMPTY
}
