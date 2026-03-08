import type { RangeRef, Workbook } from '@ascend/core'
import type { FormulaNode, FormulaRef } from '@ascend/formulas'
import { extractRefs, functionRegistry, parseFormula } from '@ascend/formulas'
import { type CellKey, cellKey, DependencyGraph } from './dep-graph.ts'

export interface AnalyzedFormula {
	readonly key: CellKey
	readonly sheetIndex: number
	readonly sheetName: string
	readonly row: number
	readonly col: number
	readonly formula: string
	readonly ast?: FormulaNode
	readonly refs: readonly FormulaRef[]
	readonly deps: readonly CellKey[]
	readonly volatile: boolean
	readonly parseError?: string
}

export interface WorkbookAnalysis {
	readonly formulas: ReadonlyMap<CellKey, AnalyzedFormula>
	readonly dependencyGraph: DependencyGraph
	readonly sheetNameIndex: ReadonlyMap<string, number>
}

export interface AnalyzeWorkbookOptions {
	readonly range?: RangeRef
}

export function createSheetNameIndex(workbook: Workbook): Map<string, number> {
	const index = new Map<string, number>()
	for (let i = 0; i < workbook.sheets.length; i++) {
		const sheet = workbook.sheets[i]
		if (!sheet) continue
		index.set(sheet.name.toLowerCase(), i)
	}
	return index
}

export function resolveSheetIndex(
	sheetNameIndex: ReadonlyMap<string, number>,
	sheetName: string | undefined,
	currentSheet: number,
): number {
	if (sheetName === undefined) return currentSheet
	return sheetNameIndex.get(sheetName.toLowerCase()) ?? -1
}

export function analyzeWorkbook(
	workbook: Workbook,
	options: AnalyzeWorkbookOptions = {},
): WorkbookAnalysis {
	const sheetNameIndex = createSheetNameIndex(workbook)
	const formulas = new Map<CellKey, AnalyzedFormula>()
	const dependencyGraph = new DependencyGraph()

	for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
		const sheet = workbook.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			if (!inRange(sheet.name, row, col, options.range)) continue

			const key = cellKey(sheetIndex, row, col)
			const parsed = parseFormula(cell.formula)
			if (!parsed.ok) {
				formulas.set(key, {
					key,
					sheetIndex,
					sheetName: sheet.name,
					row,
					col,
					formula: cell.formula,
					refs: [],
					deps: [],
					volatile: false,
					parseError: parsed.error.message,
				})
				continue
			}

			const ast = parsed.value
			const refs = extractRefs(ast)
			const deps: CellKey[] = []
			for (const ref of refs) {
				const refSheetIndex = resolveSheetIndex(sheetNameIndex, ref.sheet, sheetIndex)
				if (refSheetIndex < 0) continue
				if (ref.kind === 'cell') {
					deps.push(cellKey(refSheetIndex, ref.ref.row, ref.ref.col))
				} else {
					for (let r = ref.start.row; r <= ref.end.row; r++) {
						for (let c = ref.start.col; c <= ref.end.col; c++) {
							deps.push(cellKey(refSheetIndex, r, c))
						}
					}
				}
			}

			const volatile = hasVolatileFunction(ast)
			dependencyGraph.addFormula(key, deps, volatile)
			formulas.set(key, {
				key,
				sheetIndex,
				sheetName: sheet.name,
				row,
				col,
				formula: cell.formula,
				ast,
				refs,
				deps,
				volatile,
			})
		}
	}

	return { formulas, dependencyGraph, sheetNameIndex }
}

function inRange(
	sheetName: string,
	row: number,
	col: number,
	range: RangeRef | undefined,
): boolean {
	if (!range) return true
	if (range.sheet !== undefined && range.sheet.toLowerCase() !== sheetName.toLowerCase())
		return false
	return (
		row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col
	)
}

function hasVolatileFunction(node: FormulaNode): boolean {
	switch (node.type) {
		case 'function':
			if (functionRegistry.get(node.name.toUpperCase())?.volatile) return true
			return node.args.some((arg) => hasVolatileFunction(arg))
		case 'binary':
			return hasVolatileFunction(node.left) || hasVolatileFunction(node.right)
		case 'unary':
			return hasVolatileFunction(node.operand)
		case 'array':
			return node.rows.some((row) => row.some((cell) => hasVolatileFunction(cell)))
		default:
			return false
	}
}
