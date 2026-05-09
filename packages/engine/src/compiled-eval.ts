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
import { resolveSheetIndexInWorkbook } from './sheet-index.ts'
import { getWasmRangeOps } from './wasm-range.ts'

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
	CELL_DIV: 53,
	CELL_POW: 54,
	ROUND: 60,
	ROUNDUP: 61,
	ROUNDDOWN: 62,
	INT_OP: 63,
	TRUNC: 64,
	ABS: 65,
	NOT: 70,
	AND_JF: 71,
	OR_JT: 72,
	DUP: 73,
	SUM_RANGE: 80,
	COUNT_RANGE: 81,
	AVERAGE_RANGE: 82,
	MIN_RANGE: 83,
	MAX_RANGE: 84,
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

const COMPILABLE_FUNCTIONS = new Set([
	'IF',
	'IFERROR',
	'IFNA',
	'ROUND',
	'ROUNDUP',
	'ROUNDDOWN',
	'INT',
	'TRUNC',
	'ABS',
	'NOT',
	'AND',
	'OR',
	'SUM',
	'COUNT',
	'AVERAGE',
	'MIN',
	'MAX',
])

function shouldCompile(node: FormulaNode): boolean {
	if (node.type === 'function' && node.args.length === 1 && node.args[0]?.type === 'rangeRef') {
		const upper = node.name.toUpperCase()
		if (
			upper === 'SUM' ||
			upper === 'COUNT' ||
			upper === 'AVERAGE' ||
			upper === 'MIN' ||
			upper === 'MAX'
		) {
			return true
		}
	}
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
			case 'rangeRef':
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

function canEvaluateNumericCondition(node: FormulaNode): boolean {
	switch (node.type) {
		case 'number':
		case 'boolean':
		case 'missing':
		case 'cellRef':
			return true
		case 'unary':
			if (node.op === '@') return false
			if (node.op === '-' || node.op === '%') return canEvaluateNumericValue(node.operand, false)
			return canEvaluateNumericCondition(node.operand)
		case 'binary':
			if (node.op === ',' || node.op === ' ' || node.op === '&') return false
			if (
				node.op === '=' ||
				node.op === '<>' ||
				node.op === '<' ||
				node.op === '>' ||
				node.op === '<=' ||
				node.op === '>='
			) {
				return (
					canEvaluateNumericValue(node.left, false) && canEvaluateNumericValue(node.right, false)
				)
			}
			return canEvaluateNumericValue(node.left, false) && canEvaluateNumericValue(node.right, false)
		case 'function': {
			const upper = node.name.toUpperCase()
			if (upper === 'AND' || upper === 'OR') return node.args.every(canEvaluateNumericCondition)
			if (upper === 'NOT' && node.args.length === 1) {
				return canEvaluateNumericCondition(node.args[0] as FormulaNode)
			}
			return canEvaluateNumericValue(node, false)
		}
		default:
			return false
	}
}

function isSameCellRef(
	left: FormulaNode,
	right: FormulaNode,
): left is FormulaNode & { type: 'cellRef' } {
	return (
		left.type === 'cellRef' &&
		right.type === 'cellRef' &&
		left.sheet === right.sheet &&
		left.ref.row === right.ref.row &&
		left.ref.col === right.ref.col &&
		left.ref.rowAbsolute === right.ref.rowAbsolute &&
		left.ref.colAbsolute === right.ref.colAbsolute
	)
}

function canEvaluateNumericValue(node: FormulaNode, atRoot = true): boolean {
	switch (node.type) {
		case 'number':
		case 'missing':
		case 'cellRef':
			return true
		case 'boolean':
		case 'string':
		case 'error':
			return false
		case 'binary':
			if (node.op === ',' || node.op === ' ') return false
			if (node.op === '&') return false
			if (
				node.op === '=' ||
				node.op === '<>' ||
				node.op === '<' ||
				node.op === '>' ||
				node.op === '<=' ||
				node.op === '>='
			) {
				if (atRoot) return false
				return (
					canEvaluateNumericValue(node.left, false) && canEvaluateNumericValue(node.right, false)
				)
			}
			return canEvaluateNumericValue(node.left, false) && canEvaluateNumericValue(node.right, false)
		case 'unary':
			if (node.op === '@') return false
			return canEvaluateNumericValue(node.operand, false)
		case 'function': {
			const upper = node.name.toUpperCase()
			if (upper === 'IF' && node.args.length === 3) {
				return (
					canEvaluateNumericCondition(node.args[0] as FormulaNode) &&
					canEvaluateNumericValue(node.args[1] as FormulaNode, false) &&
					canEvaluateNumericValue(node.args[2] as FormulaNode, false)
				)
			}
			if (
				(upper === 'ROUND' || upper === 'ROUNDUP' || upper === 'ROUNDDOWN' || upper === 'TRUNC') &&
				node.args.length === 2
			) {
				return (
					canEvaluateNumericValue(node.args[0] as FormulaNode, false) &&
					canEvaluateNumericValue(node.args[1] as FormulaNode, false)
				)
			}
			if ((upper === 'INT' || upper === 'ABS') && node.args.length === 1) {
				return canEvaluateNumericValue(node.args[0] as FormulaNode, false)
			}
			if (
				(upper === 'SUM' ||
					upper === 'COUNT' ||
					upper === 'AVERAGE' ||
					upper === 'MIN' ||
					upper === 'MAX') &&
				node.args.length === 1 &&
				(node.args[0] as FormulaNode).type === 'rangeRef'
			) {
				return true
			}
			return false
		}
		default:
			return false
	}
}

export function compileFormula(node: FormulaNode): CompiledFormula | null {
	if (!shouldCompile(node)) return null

	const ops: number[] = []
	const constants: (string | number | boolean | CellValue | FormulaNode)[] = []
	const andOrJumps: number[] = []

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
				if (n.op === '^' && n.right.type === 'number' && n.right.value === 2) {
					emit(n.left)
					ops.push(Op.DUP, Op.MUL)
					break
				}
				if (n.op === '/' && n.right.type === 'number' && n.right.value === 2) {
					emit(n.left)
					ops.push(Op.NUM, addConst(0.5), Op.MUL)
					break
				}
				if (isSameCellRef(n.left, n.right)) {
					emit(n.left)
					ops.push(Op.DUP)
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
				const superOp =
					n.op === '+'
						? Op.CELL_ADD
						: n.op === '-'
							? Op.CELL_SUB
							: n.op === '*'
								? Op.CELL_MUL
								: n.op === '/'
									? Op.CELL_DIV
									: n.op === '^'
										? Op.CELL_POW
										: 0
				if (superOp) {
					if (n.right.type === 'cellRef' && n.right.sheet === undefined) {
						emit(n.left)
						ops.push(superOp, n.right.ref.row, n.right.ref.col)
						break
					}
					if (
						n.left.type === 'cellRef' &&
						n.left.sheet === undefined &&
						n.op !== '-' &&
						n.op !== '/' &&
						n.op !== '^'
					) {
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
			const condition = tryFoldConstant(n.args[0] as FormulaNode)
			if (condition !== null) {
				if (condition !== 0) {
					emit(n.args[1] as FormulaNode)
				} else if (n.args.length >= 3) {
					emit(n.args[2] as FormulaNode)
				} else {
					ops.push(Op.BOOL, addConst(false))
				}
				return
			}
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
		if (upper === 'ROUND' && n.args.length === 2) {
			emit(n.args[0] as FormulaNode)
			emit(n.args[1] as FormulaNode)
			ops.push(Op.ROUND)
			return
		}
		if (upper === 'ROUNDUP' && n.args.length === 2) {
			emit(n.args[0] as FormulaNode)
			emit(n.args[1] as FormulaNode)
			ops.push(Op.ROUNDUP)
			return
		}
		if (upper === 'ROUNDDOWN' && n.args.length === 2) {
			emit(n.args[0] as FormulaNode)
			emit(n.args[1] as FormulaNode)
			ops.push(Op.ROUNDDOWN)
			return
		}
		if (upper === 'TRUNC' && n.args.length >= 1 && n.args.length <= 2) {
			emit(n.args[0] as FormulaNode)
			if (n.args.length === 2) {
				emit(n.args[1] as FormulaNode)
			} else {
				ops.push(Op.NUM, addConst(0))
			}
			ops.push(Op.TRUNC)
			return
		}
		if (upper === 'INT' && n.args.length === 1) {
			emit(n.args[0] as FormulaNode)
			ops.push(Op.INT_OP)
			return
		}
		if (upper === 'ABS' && n.args.length === 1) {
			emit(n.args[0] as FormulaNode)
			ops.push(Op.ABS)
			return
		}
		if (upper === 'NOT' && n.args.length === 1) {
			emit(n.args[0] as FormulaNode)
			ops.push(Op.NOT)
			return
		}
		if (upper === 'AND' && n.args.length >= 1) {
			for (let i = 0; i < n.args.length; i++) {
				emit(n.args[i] as FormulaNode)
				if (i < n.args.length - 1) {
					const jfPos = ops.length
					ops.push(Op.AND_JF, 0)
					andOrJumps.push(jfPos)
				}
			}
			ops.push(Op.NOT, Op.NOT)
			const endTarget = ops.length
			for (const pos of andOrJumps) ops[pos + 1] = endTarget
			andOrJumps.length = 0
			return
		}
		if (upper === 'OR' && n.args.length >= 1) {
			for (let i = 0; i < n.args.length; i++) {
				emit(n.args[i] as FormulaNode)
				if (i < n.args.length - 1) {
					const jtPos = ops.length
					ops.push(Op.OR_JT, 0)
					andOrJumps.push(jtPos)
				}
			}
			ops.push(Op.NOT, Op.NOT)
			const endTarget = ops.length
			for (const pos of andOrJumps) ops[pos + 1] = endTarget
			andOrJumps.length = 0
			return
		}
		if (
			(upper === 'SUM' ||
				upper === 'COUNT' ||
				upper === 'AVERAGE' ||
				upper === 'MIN' ||
				upper === 'MAX') &&
			n.args.length === 1 &&
			(n.args[0] as FormulaNode).type === 'rangeRef'
		) {
			const arg = n.args[0] as FormulaNode & { type: 'rangeRef' }
			const rangeOp =
				upper === 'SUM'
					? Op.SUM_RANGE
					: upper === 'COUNT'
						? Op.COUNT_RANGE
						: upper === 'AVERAGE'
							? Op.AVERAGE_RANGE
							: upper === 'MIN'
								? Op.MIN_RANGE
								: Op.MAX_RANGE
			const constIdx = arg.sheet !== undefined ? addConst(arg.sheet) : -1
			ops.push(rangeOp, constIdx, arg.start.row, arg.start.col, arg.end.row, arg.end.col)
			return
		}
		ops.push(Op.TREE, addConst(n))
	}

	emit(node)
	return { ops, constants, numericOnly: canEvaluateNumericValue(node) }
}

// Module-level scratch buffers — reused across evaluations to avoid allocation.
// NOT safe for concurrent use. The engine is single-threaded by design; if
// multi-threaded evaluation is ever introduced, these must become thread-local.
let numericStackSize = 64
let numericStack = new Float64Array(numericStackSize)
let rangeScratchSize = 256
let rangeScratch = new Float64Array(rangeScratchSize)
const WASM_RANGE_THRESHOLD = 128
const EXCEL_MAX_ROWS = 1_048_576
const EXCEL_MAX_COLS = 16_384

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

function readCellNumeric(
	sheet: import('@ascend/core').Sheet | undefined,
	row: number,
	col: number,
): number | CellValue {
	if (!isCellInBounds(row, col)) return errorValue('#REF!')
	const cells = sheet?.cells
	if (!cells) return 0
	const kind = cells.readKind(row, col)
	if (kind === undefined || kind === 'empty') return 0
	if (kind === 'error') return errorValue(cells.readError(row, col) ?? '#VALUE!')
	const directNumber = cells.readNumber(row, col)
	if (directNumber !== null) return directNumber
	if (kind === 'string') {
		const raw = cells.readString(row, col) ?? ''
		const trimmed = raw.trim()
		if (trimmed === '') return 0
		const parsed = Number(trimmed)
		return Number.isNaN(parsed) ? errorValue('#VALUE!') : parsed
	}
	const cv = cells.readValue(row, col)
	if (cv.kind === 'error') return cv
	const n = toNumber(cv)
	return n ?? errorValue('#VALUE!')
}

export function aggregateNumericRange(
	sheet: import('@ascend/core').Sheet | undefined,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): { sum: number; count: number; min: number; max: number; error: CellValue | null } {
	if (!isRangeInBounds(startRow, startCol, endRow, endCol)) {
		return { sum: 0, count: 0, min: Infinity, max: -Infinity, error: errorValue('#REF!') }
	}
	const cells = sheet?.cells
	if (!cells) return { sum: 0, count: 0, min: Infinity, max: -Infinity, error: null }
	const maxCells = Math.max(0, endRow - startRow + 1) * Math.max(0, endCol - startCol + 1)
	const wasmEarly = maxCells >= WASM_RANGE_THRESHOLD ? getWasmRangeOps() : null
	const wasmPad = wasmEarly?.numericScratch(maxCells) ?? null
	let count = 0
	for (let row = startRow; row <= endRow; row++) {
		for (let col = startCol; col <= endCol; col++) {
			const kind = cells.readKind(row, col)
			if (kind === undefined || kind === 'empty') continue
			if (kind === 'error') {
				return {
					sum: 0,
					count,
					min: Infinity,
					max: -Infinity,
					error: errorValue(cells.readError(row, col) ?? '#VALUE!'),
				}
			}
			if (kind !== 'number' && kind !== 'date') continue
			const n = cells.readNumber(row, col) ?? 0
			if (wasmPad) {
				wasmPad[count++] = n
			} else {
				if (count === rangeScratchSize) ensureRangeScratch(count + 1)
				rangeScratch[count++] = n
			}
		}
	}
	if (count === 0) {
		return { sum: 0, count: 0, min: Infinity, max: -Infinity, error: null }
	}
	const wasm = count >= WASM_RANGE_THRESHOLD ? getWasmRangeOps() : null
	if (wasm && wasmPad && count <= wasmPad.length) {
		return {
			sum: wasm.sum(count),
			count,
			min: wasm.min(count),
			max: wasm.max(count),
			error: null,
		}
	}
	if (wasm) {
		wasm.load(rangeScratch, count)
		return {
			sum: wasm.sum(count),
			count,
			min: wasm.min(count),
			max: wasm.max(count),
			error: null,
		}
	}
	const buf = wasmPad && count <= wasmPad.length ? wasmPad : rangeScratch
	let sum = 0
	let min = Infinity
	let max = -Infinity
	for (let i = 0; i < count; i++) {
		const value = buf[i] ?? 0
		sum += value
		if (value < min) min = value
		if (value > max) max = value
	}
	return { sum, count, min, max, error: null }
}

function evaluateCompiledNumeric(compiled: CompiledFormula, ctx: EvalContext): CellValue {
	const { ops, constants } = compiled
	const sheet = ctx.workbook.sheets[ctx.sheetIndex]
	let ip = 0
	let sp = 0
	const len = ops.length
	ensureNumericStack(len)

	while (ip < len) {
		const op = ops[ip] as number
		ip++
		switch (op) {
			case Op.NUM:
				numericStack[sp++] = constants[ops[ip++] as number] as number
				break
			case Op.BOOL:
				numericStack[sp++] = (constants[ops[ip++] as number] as boolean) ? 1 : 0
				break
			case Op.EMPTY_VAL:
				numericStack[sp++] = 0
				break
			case Op.CELL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				const n = readCellNumeric(sheet, row, col)
				if (typeof n !== 'number') return n
				numericStack[sp++] = n
				break
			}
			case Op.CELL_SHEET: {
				const nameIdx = ops[ip] as number
				const row = ops[ip + 1] as number
				const col = ops[ip + 2] as number
				ip += 3
				const si = resolveSheetIndexInWorkbook(
					ctx.workbook,
					constants[nameIdx] as string,
					ctx.sheetIndex,
				)
				if (si < 0) return errorValue('#REF!')
				const n = readCellNumeric(ctx.workbook.sheets[si], row, col)
				if (typeof n !== 'number') return n
				numericStack[sp++] = n
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
				if (b === 0) return errorValue('#DIV/0!')
				numericStack[sp - 1] = (numericStack[sp - 1] as number) / b
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
			case Op.DUP:
				numericStack[sp] = numericStack[sp - 1] as number
				sp++
				break
			case Op.EQ:
			case Op.NE:
			case Op.LT:
			case Op.GT:
			case Op.LE:
			case Op.GE: {
				const b = numericStack[--sp] as number
				const a = numericStack[sp - 1] as number
				numericStack[sp - 1] = comparePrimitive(op, a, b) ? 1 : 0
				break
			}
			case Op.CELL_ADD:
			case Op.CELL_SUB:
			case Op.CELL_MUL:
			case Op.CELL_DIV:
			case Op.CELL_POW: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				const cn = readCellNumeric(sheet, row, col)
				if (typeof cn !== 'number') return cn
				const left = numericStack[sp - 1] as number
				if (op === Op.CELL_ADD) numericStack[sp - 1] = left + cn
				else if (op === Op.CELL_SUB) numericStack[sp - 1] = left - cn
				else if (op === Op.CELL_MUL) numericStack[sp - 1] = left * cn
				else if (op === Op.CELL_DIV) {
					if (cn === 0) return errorValue('#DIV/0!')
					numericStack[sp - 1] = left / cn
				} else numericStack[sp - 1] = left ** cn
				break
			}
			case Op.ROUND: {
				const digits = numericStack[--sp] as number
				const val = numericStack[sp - 1] as number
				const factor = 10 ** Math.trunc(digits)
				numericStack[sp - 1] = (Math.sign(val) * Math.round(Math.abs(val) * factor)) / factor
				break
			}
			case Op.ROUNDUP: {
				const digits = numericStack[--sp] as number
				const val = numericStack[sp - 1] as number
				const factor = 10 ** Math.trunc(digits)
				const scaled = val * factor
				numericStack[sp - 1] = (scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled)) / factor
				break
			}
			case Op.ROUNDDOWN: {
				const digits = numericStack[--sp] as number
				const val = numericStack[sp - 1] as number
				const factor = 10 ** Math.trunc(digits)
				numericStack[sp - 1] = Math.trunc(val * factor) / factor
				break
			}
			case Op.INT_OP:
				numericStack[sp - 1] = Math.floor(numericStack[sp - 1] as number)
				break
			case Op.TRUNC: {
				const digits = numericStack[--sp] as number
				const val = numericStack[sp - 1] as number
				const factor = 10 ** Math.trunc(digits)
				numericStack[sp - 1] = Math.trunc(val * factor) / factor
				break
			}
			case Op.ABS:
				numericStack[sp - 1] = Math.abs(numericStack[sp - 1] as number)
				break
			case Op.NOT:
				numericStack[sp - 1] = (numericStack[sp - 1] as number) === 0 ? 1 : 0
				break
			case Op.AND_JF: {
				const endTarget = ops[ip] as number
				ip++
				if ((numericStack[sp - 1] as number) === 0) {
					numericStack[sp - 1] = 0
					ip = endTarget
				} else {
					sp--
				}
				break
			}
			case Op.OR_JT: {
				const endTarget = ops[ip] as number
				ip++
				if ((numericStack[sp - 1] as number) !== 0) {
					numericStack[sp - 1] = 1
					ip = endTarget
				} else {
					sp--
				}
				break
			}
			case Op.IF: {
				const falseTarget = ops[ip] as number
				ip += 2
				if ((numericStack[--sp] as number) === 0) ip = falseTarget
				break
			}
			case Op.JMP:
				ip = ops[ip] as number
				break
			case Op.SUM_RANGE:
			case Op.COUNT_RANGE:
			case Op.AVERAGE_RANGE:
			case Op.MIN_RANGE:
			case Op.MAX_RANGE: {
				const ci = ops[ip] as number
				const sr = ops[ip + 1] as number
				const sc = ops[ip + 2] as number
				const er = ops[ip + 3] as number
				const ec = ops[ip + 4] as number
				ip += 5
				let target = sheet
				if (ci !== -1) {
					const si = resolveSheetIndexInWorkbook(
						ctx.workbook,
						constants[ci] as string,
						ctx.sheetIndex,
					)
					if (si < 0) return errorValue('#REF!')
					target = ctx.workbook.sheets[si]
				}
				const { sum, count, min, max, error } = aggregateNumericRange(target, sr, sc, er, ec)
				if (error) return error
				if (op === Op.SUM_RANGE) numericStack[sp++] = sum
				else if (op === Op.COUNT_RANGE) numericStack[sp++] = count
				else if (op === Op.AVERAGE_RANGE) {
					if (count === 0) return errorValue('#DIV/0!')
					numericStack[sp++] = sum / count
				} else if (op === Op.MIN_RANGE) numericStack[sp++] = count === 0 ? 0 : min
				else numericStack[sp++] = count === 0 ? 0 : max
				break
			}
		}
	}

	return numberValue(numericStack[0] as number)
}

const MIXED_STACK_INITIAL_SIZE = 64

function ensureNumericStack(needed: number): void {
	if (needed <= numericStackSize) return
	const newSize = Math.max(numericStackSize * 2, needed)
	const newStack = new Float64Array(newSize)
	newStack.set(numericStack)
	numericStack = newStack
	numericStackSize = newSize
}

function ensureRangeScratch(needed: number): void {
	if (needed <= rangeScratchSize) return
	const newSize = Math.max(rangeScratchSize * 2, needed)
	const next = new Float64Array(newSize)
	next.set(rangeScratch)
	rangeScratch = next
	rangeScratchSize = newSize
}

export function evaluateCompiled(compiled: CompiledFormula, ctx: EvalContext): CellValue {
	if (compiled.numericOnly) {
		return evaluateCompiledNumeric(compiled, ctx)
	}

	const { ops, constants } = compiled
	let stack: CellValue[] = new Array(MIXED_STACK_INITIAL_SIZE)
	let stackDepth = 0
	const sheet = ctx.workbook.sheets[ctx.sheetIndex]
	let ip = 0
	const len = ops.length

	function ensureStack(needed: number): void {
		if (needed <= stack.length) return
		const newSize = Math.max(stack.length * 2, needed)
		const newStack: CellValue[] = new Array(newSize)
		for (let i = 0; i < stackDepth; i++) newStack[i] = stack[i] as CellValue
		stack = newStack
	}
	ensureStack(stackDepth + len)

	while (ip < len) {
		const op = ops[ip] as number
		ip++
		switch (op) {
			case Op.NUM: {
				const idx = ops[ip] as number
				ip++
				stack[stackDepth++] = numberValue(constants[idx] as number)
				break
			}
			case Op.STR: {
				const idx = ops[ip] as number
				ip++
				stack[stackDepth++] = stringValue(constants[idx] as string)
				break
			}
			case Op.BOOL: {
				const idx = ops[ip] as number
				ip++
				stack[stackDepth++] = booleanValue(constants[idx] as boolean)
				break
			}
			case Op.ERR: {
				const idx = ops[ip] as number
				ip++
				stack[stackDepth++] = errorValue(constants[idx] as ExcelError)
				break
			}
			case Op.EMPTY_VAL:
				stack[stackDepth++] = EMPTY
				break
			case Op.DUP:
				stack[stackDepth] = stack[stackDepth - 1] ?? EMPTY
				stackDepth++
				break
			case Op.CELL: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				stack[stackDepth++] = isCellInBounds(row, col)
					? (sheet?.cells.readValue(row, col) ?? EMPTY)
					: errorValue('#REF!')
				break
			}
			case Op.CELL_SHEET: {
				const nameIdx = ops[ip] as number
				const row = ops[ip + 1] as number
				const col = ops[ip + 2] as number
				ip += 3
				const sheetName = constants[nameIdx] as string
				const si = resolveSheetIndexInWorkbook(ctx.workbook, sheetName, ctx.sheetIndex)
				if (si < 0) {
					stack[stackDepth++] = errorValue('#REF!')
				} else {
					const target = ctx.workbook.sheets[si]
					stack[stackDepth++] = isCellInBounds(row, col)
						? (target?.cells.readValue(row, col) ?? EMPTY)
						: errorValue('#REF!')
				}
				break
			}
			case Op.ADD:
			case Op.SUB:
			case Op.MUL:
			case Op.DIV:
			case Op.POW: {
				const right = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				const left = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (left.kind === 'error') {
					stack[stackDepth++] = left
					break
				}
				if (right.kind === 'error') {
					stack[stackDepth++] = right
					break
				}
				const ln = toNumber(left)
				const rn = toNumber(right)
				if (ln === null || rn === null) {
					stack[stackDepth++] = errorValue('#VALUE!')
					break
				}
				switch (op) {
					case Op.ADD:
						stack[stackDepth++] = numberValue(ln + rn)
						break
					case Op.SUB:
						stack[stackDepth++] = numberValue(ln - rn)
						break
					case Op.MUL:
						stack[stackDepth++] = numberValue(ln * rn)
						break
					case Op.DIV:
						stack[stackDepth++] = rn === 0 ? errorValue('#DIV/0!') : numberValue(ln / rn)
						break
					case Op.POW:
						stack[stackDepth++] = numberValue(ln ** rn)
						break
				}
				break
			}
			case Op.CELL_ADD:
			case Op.CELL_SUB:
			case Op.CELL_MUL:
			case Op.CELL_DIV:
			case Op.CELL_POW: {
				const row = ops[ip] as number
				const col = ops[ip + 1] as number
				ip += 2
				const cellVal = isCellInBounds(row, col)
					? (sheet?.cells.readValue(row, col) ?? EMPTY)
					: errorValue('#REF!')
				const left = stack[stackDepth - 1] ?? EMPTY
				const ln = toNumber(left)
				const cn = toNumber(cellVal)
				if (ln !== null && cn !== null) {
					if (op === Op.CELL_DIV && cn === 0) {
						stack[stackDepth - 1] = errorValue('#DIV/0!')
					} else if (op === Op.CELL_ADD) stack[stackDepth - 1] = numberValue(ln + cn)
					else if (op === Op.CELL_SUB) stack[stackDepth - 1] = numberValue(ln - cn)
					else if (op === Op.CELL_MUL) stack[stackDepth - 1] = numberValue(ln * cn)
					else if (op === Op.CELL_DIV) stack[stackDepth - 1] = numberValue(ln / cn)
					else stack[stackDepth - 1] = numberValue(ln ** cn)
				} else {
					const sv = topLeftScalar(left)
					const cv = topLeftScalar(cellVal)
					if (sv.kind === 'error') stack[stackDepth - 1] = sv
					else if (cv.kind === 'error') stack[stackDepth - 1] = cv
					else stack[stackDepth - 1] = errorValue('#VALUE!')
				}
				break
			}
			case Op.NEG: {
				const v = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (v.kind === 'error') {
					stack[stackDepth++] = v
					break
				}
				const n = toNumber(v)
				stack[stackDepth++] = n === null ? errorValue('#VALUE!') : numberValue(-n)
				break
			}
			case Op.PCT: {
				const v = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (v.kind === 'error') {
					stack[stackDepth++] = v
					break
				}
				const n = toNumber(v)
				stack[stackDepth++] = n === null ? errorValue('#VALUE!') : numberValue(n / 100)
				break
			}
			case Op.CONCAT: {
				const right = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				const left = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (left.kind === 'error') {
					stack[stackDepth++] = left
					break
				}
				if (right.kind === 'error') {
					stack[stackDepth++] = right
					break
				}
				stack[stackDepth++] = stringValue(coerceStr(left) + coerceStr(right))
				break
			}
			case Op.EQ:
			case Op.NE:
			case Op.LT:
			case Op.GT:
			case Op.LE:
			case Op.GE: {
				const right = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				const left = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (left.kind === 'error') {
					stack[stackDepth++] = left
					break
				}
				if (right.kind === 'error') {
					stack[stackDepth++] = right
					break
				}
				stack[stackDepth++] = booleanValue(evalCmp(op, left, right))
				break
			}
			case Op.IF: {
				const falseTarget = ops[ip] as number
				const endTarget = ops[ip + 1] as number
				ip += 2
				const cond = coerceToBoolForIf(stack[--stackDepth] ?? EMPTY)
				if (typeof cond !== 'boolean') {
					stack[stackDepth++] = cond
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
				const v = topLeftScalar(stack[stackDepth - 1] ?? EMPTY)
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
				const v = topLeftScalar(stack[stackDepth - 1] ?? EMPTY)
				if (!(v.kind === 'error' && v.value === '#N/A')) {
					ip = endTarget
				} else {
					stackDepth--
				}
				break
			}
			case Op.ROUND:
			case Op.ROUNDUP:
			case Op.ROUNDDOWN:
			case Op.TRUNC: {
				const digitsVal = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				const numVal = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (numVal.kind === 'error') {
					stack[stackDepth++] = numVal
					break
				}
				if (digitsVal.kind === 'error') {
					stack[stackDepth++] = digitsVal
					break
				}
				const num = toNumber(numVal)
				const digits = toNumber(digitsVal)
				if (num === null || digits === null) {
					stack[stackDepth++] = errorValue('#VALUE!')
					break
				}
				const factor = 10 ** Math.trunc(digits)
				if (op === Op.ROUND) {
					stack[stackDepth++] = numberValue(
						(Math.sign(num) * Math.round(Math.abs(num) * factor)) / factor,
					)
				} else if (op === Op.ROUNDUP) {
					const scaled = num * factor
					stack[stackDepth++] = numberValue(
						(scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled)) / factor,
					)
				} else if (op === Op.ROUNDDOWN || op === Op.TRUNC)
					stack[stackDepth++] = numberValue(Math.trunc(num * factor) / factor)
				break
			}
			case Op.INT_OP: {
				const v = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (v.kind === 'error') {
					stack[stackDepth++] = v
					break
				}
				const n = toNumber(v)
				stack[stackDepth++] = n === null ? errorValue('#VALUE!') : numberValue(Math.floor(n))
				break
			}
			case Op.ABS: {
				const v = topLeftScalar(stack[--stackDepth] ?? EMPTY)
				if (v.kind === 'error') {
					stack[stackDepth++] = v
					break
				}
				const n = toNumber(v)
				stack[stackDepth++] = n === null ? errorValue('#VALUE!') : numberValue(Math.abs(n))
				break
			}
			case Op.NOT: {
				const v = stack[--stackDepth] ?? EMPTY
				const b = coerceToBoolForIf(v)
				if (typeof b !== 'boolean') {
					stack[stackDepth++] = b
					break
				}
				stack[stackDepth++] = booleanValue(!b)
				break
			}
			case Op.AND_JF: {
				const endTarget = ops[ip] as number
				ip++
				const v = stack[stackDepth - 1] ?? EMPTY
				const b = coerceToBoolForIf(v)
				if (typeof b !== 'boolean') {
					ip = endTarget
				} else if (!b) {
					stack[stackDepth - 1] = booleanValue(false)
					ip = endTarget
				} else {
					stackDepth--
				}
				break
			}
			case Op.OR_JT: {
				const endTarget = ops[ip] as number
				ip++
				const v = stack[stackDepth - 1] ?? EMPTY
				const b = coerceToBoolForIf(v)
				if (typeof b !== 'boolean') {
					ip = endTarget
				} else if (b) {
					stack[stackDepth - 1] = booleanValue(true)
					ip = endTarget
				} else {
					stackDepth--
				}
				break
			}
			case Op.SUM_RANGE:
			case Op.COUNT_RANGE:
			case Op.AVERAGE_RANGE:
			case Op.MIN_RANGE:
			case Op.MAX_RANGE: {
				const ci = ops[ip] as number
				const sr = ops[ip + 1] as number
				const sc = ops[ip + 2] as number
				const er = ops[ip + 3] as number
				const ec = ops[ip + 4] as number
				ip += 5
				let target = sheet
				if (ci !== -1) {
					const si = resolveSheetIndexInWorkbook(
						ctx.workbook,
						constants[ci] as string,
						ctx.sheetIndex,
					)
					if (si < 0) {
						stack[stackDepth++] = errorValue('#REF!')
						break
					}
					target = ctx.workbook.sheets[si]
				}
				const {
					sum,
					count,
					min,
					max,
					error: rangeErr,
				} = aggregateNumericRange(target, sr, sc, er, ec)
				if (rangeErr) {
					stack[stackDepth++] = rangeErr
				} else if (op === Op.SUM_RANGE) stack[stackDepth++] = numberValue(sum)
				else if (op === Op.COUNT_RANGE) stack[stackDepth++] = numberValue(count)
				else if (op === Op.AVERAGE_RANGE)
					stack[stackDepth++] = count === 0 ? errorValue('#DIV/0!') : numberValue(sum / count)
				else if (op === Op.MIN_RANGE) stack[stackDepth++] = numberValue(count === 0 ? 0 : min)
				else stack[stackDepth++] = numberValue(count === 0 ? 0 : max)
				break
			}
			case Op.TREE: {
				const idx = ops[ip] as number
				ip++
				const node = constants[idx] as FormulaNode
				stack[stackDepth++] = treeEvaluate(node, ctx)
				break
			}
		}
	}
	return stack[0] ?? EMPTY
}
