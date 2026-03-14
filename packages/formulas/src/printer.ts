import { indexToColumn } from '@ascend/core'
import type { BinaryNode, BinaryOp, FormulaCellRef, FormulaNode, StructuredRefNode } from './ast.ts'

const PRECEDENCE: Record<BinaryOp, number> = {
	',': 0,
	' ': 0,
	'=': 1,
	'<>': 1,
	'<': 1,
	'>': 1,
	'<=': 1,
	'>=': 1,
	'&': 2,
	'+': 3,
	'-': 3,
	'*': 4,
	'/': 4,
	'^': 5,
}

const CELL_REF_PATTERN = /^\$?[A-Za-z]{1,3}\$?\d+$/

function needsQuoting(sheet: string): boolean {
	if (!/^[A-Za-z_]\w*$/.test(sheet)) return true
	if (CELL_REF_PATTERN.test(sheet)) return true
	const upper = sheet.toUpperCase()
	return upper === 'TRUE' || upper === 'FALSE'
}

function formatSheet(sheet: string): string {
	if (sheet.startsWith('[') && !sheet.includes(' ')) return `${sheet}!`
	if (needsQuoting(sheet)) {
		return `'${sheet.replace(/'/g, "''")}'!`
	}
	return `${sheet}!`
}

function formatSheetSpan(startSheet: string, endSheet: string): string {
	if (needsQuoting(startSheet) || needsQuoting(endSheet)) {
		return `'${startSheet.replace(/'/g, "''")}:${endSheet.replace(/'/g, "''")}'!`
	}
	return `${startSheet}:${endSheet}!`
}

function formatCellRef(ref: FormulaCellRef): string {
	const col = ref.colAbsolute ? `$${indexToColumn(ref.col)}` : indexToColumn(ref.col)
	const row = ref.rowAbsolute ? `$${ref.row + 1}` : String(ref.row + 1)
	return col + row
}

function applyRefOffset(ref: FormulaCellRef, rowDelta: number, colDelta: number): FormulaCellRef {
	return {
		...ref,
		row: ref.rowAbsolute ? ref.row : ref.row + rowDelta,
		col: ref.colAbsolute ? ref.col : ref.col + colDelta,
	}
}

function binaryNeedsParens(
	child: FormulaNode,
	parentOp: BinaryOp,
	side: 'left' | 'right',
): boolean {
	if (child.type !== 'binary') return false
	const childPrec = PRECEDENCE[child.op]
	const parentPrec = PRECEDENCE[parentOp]
	if (childPrec < parentPrec) return true
	if (childPrec === parentPrec && side === 'right') return true
	return false
}

function printBinary(node: BinaryNode, ctx: PrintContext): string {
	const leftStr = binaryNeedsParens(node.left, node.op, 'left')
		? `(${printNode(node.left, ctx)})`
		: printNode(node.left, ctx)
	const rightStr = binaryNeedsParens(node.right, node.op, 'right')
		? `(${printNode(node.right)})`
		: printNode(node.right, ctx)
	return `${leftStr}${node.op}${rightStr}`
}

function printStructuredRef(node: StructuredRefNode): string {
	let inner: string
	if (node.specifiers.length > 0 && node.column !== undefined) {
		const specPart = node.specifiers.join(',')
		inner = `${specPart}${node.column}`
	} else if (node.specifiers.length > 0) {
		inner = node.specifiers.join(',')
	} else {
		inner = node.column ?? ''
	}
	return `${node.table}[${inner}]`
}

type PrintContext = { rowDelta: number; colDelta: number } | null

function printNode(node: FormulaNode, ctx: PrintContext = null): string {
	switch (node.type) {
		case 'number':
			return String(node.value)
		case 'string':
			return `"${node.value.replace(/"/g, '""')}"`
		case 'boolean':
			return node.value ? 'TRUE' : 'FALSE'
		case 'error':
			return node.value
		case 'cellRef': {
			const ref = ctx ? applyRefOffset(node.ref, ctx.rowDelta, ctx.colDelta) : node.ref
			return (node.sheet !== undefined ? formatSheet(node.sheet) : '') + formatCellRef(ref)
		}
		case 'rangeRef': {
			const start = ctx ? applyRefOffset(node.start, ctx.rowDelta, ctx.colDelta) : node.start
			const end = ctx ? applyRefOffset(node.end, ctx.rowDelta, ctx.colDelta) : node.end
			return (
				(node.sheet !== undefined ? formatSheet(node.sheet) : '') +
				formatCellRef(start) +
				':' +
				formatCellRef(end)
			)
		}
		case 'wholeRowRange':
			return `${node.sheet !== undefined ? formatSheet(node.sheet) : ''}${node.startRow + 1}:${node.endRow + 1}`
		case 'wholeColumnRange':
			return `${node.sheet !== undefined ? formatSheet(node.sheet) : ''}${indexToColumn(node.startCol)}:${indexToColumn(node.endCol)}`
		case 'name':
			return (node.sheet !== undefined ? formatSheet(node.sheet) : '') + node.name
		case 'function':
			if (node.name === '__CALL__') {
				const [callee, ...args] = node.args
				if (!callee) return '__CALL__()'
				const calleeStr =
					callee.type === 'binary' ? `(${printNode(callee, ctx)})` : printNode(callee, ctx)
				return `${calleeStr}(${args.map((a) => printFunctionArg(a, ctx)).join(',')})`
			}
			return `${node.name}(${node.args.map((a) => printFunctionArg(a, ctx)).join(',')})`
		case 'binary':
			return printBinary(node, ctx)
		case 'unary': {
			if (node.op === '%') {
				const inner =
					node.operand.type === 'binary' ||
					(node.operand.type === 'unary' && node.operand.op !== '%')
						? `(${printNode(node.operand, ctx)})`
						: printNode(node.operand, ctx)
				return `${inner}%`
			}
			const inner =
				node.operand.type === 'binary'
					? `(${printNode(node.operand, ctx)})`
					: printNode(node.operand, ctx)
			return `${node.op}${inner}`
		}
		case 'spillRef': {
			const inner =
				node.target.type === 'binary'
					? `(${printNode(node.target, ctx)})`
					: printNode(node.target, ctx)
			return `${inner}#`
		}
		case 'sheetSpanRef':
			return `${formatSheetSpan(node.startSheet, node.endSheet)}${printNode(node.target, ctx)}`
		case 'array':
			return `{${node.rows.map((row) => row.map((c) => printNode(c, ctx)).join(',')).join(';')}}`
		case 'structuredRef':
			return printStructuredRef(node)
		case 'missing':
			return ''
	}
}

export function printFormula(node: FormulaNode): string {
	return printNode(node, null)
}

export function printFormulaWithOffset(
	node: FormulaNode,
	rowDelta: number,
	colDelta: number,
): string {
	return printNode(node, { rowDelta, colDelta })
}

function printFunctionArg(node: FormulaNode, ctx: PrintContext): string {
	if (node.type === 'binary' && node.op === ',') {
		return `(${printNode(node, ctx)})`
	}
	return printNode(node, ctx)
}
