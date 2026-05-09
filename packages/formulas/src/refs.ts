import type { FormulaCellRef, FormulaNode } from './ast.ts'

export type FormulaRef =
	| {
			readonly kind: 'cell'
			readonly ref: FormulaCellRef
			readonly sheet?: string
	  }
	| {
			readonly kind: 'range'
			readonly start: FormulaCellRef
			readonly end: FormulaCellRef
			readonly sheet?: string
	  }
	| {
			readonly kind: 'wholeRowRange'
			readonly startRow: number
			readonly endRow: number
			readonly sheet?: string
	  }
	| {
			readonly kind: 'wholeColumnRange'
			readonly startCol: number
			readonly endCol: number
			readonly sheet?: string
	  }
	| {
			readonly kind: 'sheetSpan'
			readonly startSheet: string
			readonly endSheet: string
			readonly target: FormulaRef
	  }

function walk(node: FormulaNode, out: FormulaRef[]): void {
	switch (node.type) {
		case 'cellRef':
			out.push(
				node.sheet !== undefined
					? { kind: 'cell', ref: node.ref, sheet: node.sheet }
					: { kind: 'cell', ref: node.ref },
			)
			break
		case 'rangeRef':
			out.push(
				node.sheet !== undefined
					? { kind: 'range', start: node.start, end: node.end, sheet: node.sheet }
					: { kind: 'range', start: node.start, end: node.end },
			)
			break
		case 'dynamicRangeRef':
			walk(node.start, out)
			walk(node.end, out)
			break
		case 'wholeRowRange':
			out.push(
				node.sheet !== undefined
					? {
							kind: 'wholeRowRange',
							startRow: node.startRow,
							endRow: node.endRow,
							sheet: node.sheet,
						}
					: { kind: 'wholeRowRange', startRow: node.startRow, endRow: node.endRow },
			)
			break
		case 'wholeColumnRange':
			out.push(
				node.sheet !== undefined
					? {
							kind: 'wholeColumnRange',
							startCol: node.startCol,
							endCol: node.endCol,
							sheet: node.sheet,
						}
					: { kind: 'wholeColumnRange', startCol: node.startCol, endCol: node.endCol },
			)
			break
		case 'sheetSpanRef': {
			const nested: FormulaRef[] = []
			walk(node.target, nested)
			const target = nested[0]
			if (target) {
				out.push({
					kind: 'sheetSpan',
					startSheet: node.startSheet,
					endSheet: node.endSheet,
					target,
				})
			}
			break
		}
		case 'binary':
			walk(node.left, out)
			walk(node.right, out)
			break
		case 'unary':
			walk(node.operand, out)
			break
		case 'function':
			for (const arg of node.args) walk(arg, out)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walk(cell, out)
			}
			break
		case 'spillRef':
			walk(node.target, out)
			break
		default:
			break
	}
}

export function extractRefs(node: FormulaNode): FormulaRef[] {
	const refs: FormulaRef[] = []
	walk(node, refs)
	return refs
}

export function rewriteRefs(
	node: FormulaNode,
	transform: (ref: FormulaCellRef) => FormulaCellRef,
): FormulaNode {
	switch (node.type) {
		case 'cellRef': {
			const ref = transform(node.ref)
			if (node.sheet !== undefined) {
				return { type: 'cellRef', ref, sheet: node.sheet }
			}
			return { type: 'cellRef', ref }
		}
		case 'rangeRef': {
			const start = transform(node.start)
			const end = transform(node.end)
			if (node.sheet !== undefined) {
				return { type: 'rangeRef', start, end, sheet: node.sheet }
			}
			return { type: 'rangeRef', start, end }
		}
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: rewriteRefs(node.start, transform),
				end: rewriteRefs(node.end, transform),
			}
		case 'wholeColumnRange': {
			const start = transform({
				row: 0,
				col: node.startCol,
				rowAbsolute: true,
				colAbsolute: node.startColAbsolute ?? false,
			})
			const end = transform({
				row: 0,
				col: node.endCol,
				rowAbsolute: true,
				colAbsolute: node.endColAbsolute ?? false,
			})
			return {
				type: 'wholeColumnRange',
				startCol: start.col,
				endCol: end.col,
				...(node.startColAbsolute ? { startColAbsolute: true } : {}),
				...(node.endColAbsolute ? { endColAbsolute: true } : {}),
				...(node.sheet !== undefined ? { sheet: node.sheet } : {}),
			}
		}
		case 'wholeRowRange':
			return {
				type: 'wholeRowRange',
				startRow: transform({
					row: node.startRow,
					col: 0,
					rowAbsolute: false,
					colAbsolute: true,
				}).row,
				endRow: transform({
					row: node.endRow,
					col: 0,
					rowAbsolute: false,
					colAbsolute: true,
				}).row,
				...(node.sheet !== undefined ? { sheet: node.sheet } : {}),
			}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteRefs(node.left, transform),
				right: rewriteRefs(node.right, transform),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteRefs(node.operand, transform),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((a) => rewriteRefs(a, transform)),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((r) => r.map((c) => rewriteRefs(c, transform))),
			}
		case 'spillRef':
			return {
				type: 'spillRef',
				target: rewriteRefs(node.target, transform),
			}
		case 'sheetSpanRef':
			return {
				type: 'sheetSpanRef',
				startSheet: node.startSheet,
				endSheet: node.endSheet,
				target: rewriteRefs(node.target, transform),
			}
		default:
			return node
	}
}
