import type { Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { cachedParseFormula } from '@ascend/formulas'
import type { CellValue } from '@ascend/schema'
import { errorValue } from '@ascend/schema'
import { defaultCalcContext } from './calc-context.ts'
import { codegenSharedFormula } from './codegen.ts'
import type { EvalContext } from './evaluator.ts'
import { evaluate } from './evaluator.ts'

function shiftFormulaNode(node: FormulaNode, rowDelta: number, colDelta: number): FormulaNode {
	switch (node.type) {
		case 'cellRef':
			return {
				...node,
				ref: {
					...node.ref,
					row: node.ref.rowAbsolute ? node.ref.row : node.ref.row + rowDelta,
					col: node.ref.colAbsolute ? node.ref.col : node.ref.col + colDelta,
				},
			}
		case 'rangeRef':
			return {
				...node,
				start: {
					...node.start,
					row: node.start.rowAbsolute ? node.start.row : node.start.row + rowDelta,
					col: node.start.colAbsolute ? node.start.col : node.start.col + colDelta,
				},
				end: {
					...node.end,
					row: node.end.rowAbsolute ? node.end.row : node.end.row + rowDelta,
					col: node.end.colAbsolute ? node.end.col : node.end.col + colDelta,
				},
			}
		case 'dynamicRangeRef':
			return {
				...node,
				start: shiftFormulaNode(node.start, rowDelta, colDelta),
				end: shiftFormulaNode(node.end, rowDelta, colDelta),
			}
		case 'binary':
			return {
				...node,
				left: shiftFormulaNode(node.left, rowDelta, colDelta),
				right: shiftFormulaNode(node.right, rowDelta, colDelta),
			}
		case 'unary':
			return { ...node, operand: shiftFormulaNode(node.operand, rowDelta, colDelta) }
		case 'function':
			return { ...node, args: node.args.map((arg) => shiftFormulaNode(arg, rowDelta, colDelta)) }
		default:
			return node
	}
}

export function evaluateRelativeFormulaText(
	formulaText: string,
	workbook: Workbook,
	sheetIndex: number,
	anchorRow: number,
	anchorCol: number,
	row: number,
	col: number,
): CellValue {
	const normalized = formulaText.startsWith('=') ? formulaText.slice(1) : formulaText
	const parsed = cachedParseFormula(normalized)
	if (!parsed.ok) return errorValue('#VALUE!')
	const ast = parsed.value
	const shared = codegenSharedFormula(normalized, ast, { row: anchorRow, col: anchorCol })
	const ctx: EvalContext = {
		workbook,
		calcContext: defaultCalcContext({
			dateSystem: workbook.calcSettings.dateSystem,
			iterativeCalc: workbook.calcSettings.iterativeCalc,
		}),
		sheetIndex,
		row,
		col,
	}
	if (shared) return shared(ctx)
	const shifted = shiftFormulaNode(ast, row - anchorRow, col - anchorCol)
	return evaluate(shifted, ctx)
}
