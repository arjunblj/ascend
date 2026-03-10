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
	if (needsQuoting(sheet)) {
		return `'${sheet.replace(/'/g, "''")}'!`
	}
	return `${sheet}!`
}

function formatCellRef(ref: FormulaCellRef): string {
	const col = ref.colAbsolute ? `$${indexToColumn(ref.col)}` : indexToColumn(ref.col)
	const row = ref.rowAbsolute ? `$${ref.row + 1}` : String(ref.row + 1)
	return col + row
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

function printBinary(node: BinaryNode): string {
	const leftStr = binaryNeedsParens(node.left, node.op, 'left')
		? `(${printNode(node.left)})`
		: printNode(node.left)
	const rightStr = binaryNeedsParens(node.right, node.op, 'right')
		? `(${printNode(node.right)})`
		: printNode(node.right)
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

function printNode(node: FormulaNode): string {
	switch (node.type) {
		case 'number':
			return String(node.value)
		case 'string':
			return `"${node.value.replace(/"/g, '""')}"`
		case 'boolean':
			return node.value ? 'TRUE' : 'FALSE'
		case 'error':
			return node.value
		case 'cellRef':
			return (node.sheet !== undefined ? formatSheet(node.sheet) : '') + formatCellRef(node.ref)
		case 'rangeRef':
			return (
				(node.sheet !== undefined ? formatSheet(node.sheet) : '') +
				formatCellRef(node.start) +
				':' +
				formatCellRef(node.end)
			)
		case 'wholeRowRange':
			return `${node.sheet !== undefined ? formatSheet(node.sheet) : ''}${node.startRow + 1}:${node.endRow + 1}`
		case 'wholeColumnRange':
			return `${node.sheet !== undefined ? formatSheet(node.sheet) : ''}${indexToColumn(node.startCol)}:${indexToColumn(node.endCol)}`
		case 'name':
			return (node.sheet !== undefined ? formatSheet(node.sheet) : '') + node.name
		case 'function':
			return `${node.name}(${node.args.map(printNode).join(',')})`
		case 'binary':
			return printBinary(node)
		case 'unary': {
			if (node.op === '%') {
				const inner =
					node.operand.type === 'binary' ||
					(node.operand.type === 'unary' && node.operand.op !== '%')
						? `(${printNode(node.operand)})`
						: printNode(node.operand)
				return `${inner}%`
			}
			const inner =
				node.operand.type === 'binary' ? `(${printNode(node.operand)})` : printNode(node.operand)
			return `${node.op}${inner}`
		}
		case 'spillRef': {
			const inner =
				node.target.type === 'binary' ? `(${printNode(node.target)})` : printNode(node.target)
			return `${inner}#`
		}
		case 'array':
			return `{${node.rows.map((row) => row.map(printNode).join(',')).join(';')}}`
		case 'structuredRef':
			return printStructuredRef(node)
		case 'missing':
			return ''
	}
}

export function printFormula(node: FormulaNode): string {
	return printNode(node)
}
