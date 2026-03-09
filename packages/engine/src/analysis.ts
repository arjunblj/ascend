import type { RangeRef, Workbook } from '@ascend/core'
import type { FormulaNode, FormulaRef } from '@ascend/formulas'
import { extractRefs, functionRegistry, parseFormula } from '@ascend/formulas'
import { type CellKey, cellKey, DependencyGraph, type RangeDependency } from './dep-graph.ts'
import { resolveStructuredRefRange } from './structured-refs.ts'

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
	readonly rangeDeps: readonly RangeDependency[]
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
					rangeDeps: [],
					volatile: false,
					parseError: parsed.error.message,
				})
				continue
			}

			const ast = parsed.value
			const refs = extractRefsWithNames(ast, workbook, sheetNameIndex, sheetIndex, [])
			const deps: CellKey[] = []
			const rangeDeps: RangeDependency[] = []
			for (const ref of refs) {
				const refSheetIndex = resolveSheetIndex(sheetNameIndex, ref.sheet, sheetIndex)
				if (refSheetIndex < 0) continue
				if (ref.kind === 'cell') {
					deps.push(cellKey(refSheetIndex, ref.ref.row, ref.ref.col))
				} else {
					rangeDeps.push({
						sheetIndex: refSheetIndex,
						startRow: ref.start.row,
						startCol: ref.start.col,
						endRow: ref.end.row,
						endCol: ref.end.col,
					})
				}
			}
			for (const structuredRef of collectStructuredRefs(ast)) {
				const resolved = resolveStructuredRefRange(workbook, structuredRef, sheetIndex, row, col)
				if (!resolved) continue
				rangeDeps.push({
					sheetIndex: resolved.sheetIndex,
					startRow: resolved.startRow,
					startCol: resolved.startCol,
					endRow: resolved.endRow,
					endCol: resolved.endCol,
				})
			}

			const volatile = hasVolatileFunction(ast)
			dependencyGraph.addFormula(key, deps, volatile, rangeDeps)
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
				rangeDeps,
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

function extractRefsWithNames(
	node: FormulaNode,
	workbook: Workbook,
	sheetNameIndex: ReadonlyMap<string, number>,
	sheetIndex: number,
	seenNames: readonly string[],
): FormulaRef[] {
	const refs = extractRefs(node)
	const nameRefs = collectNameRefs(node)
	for (const nameRef of nameRefs) {
		const currentSheet = workbook.sheets[sheetIndex]
		const explicitSheet = nameRef.sheet ? workbook.getSheet(nameRef.sheet) : undefined
		const entry = workbook.definedNames.resolve(nameRef.name, currentSheet?.id, explicitSheet?.id)
		if (!entry) continue

		const entryKey =
			entry.scope.kind === 'workbook'
				? `workbook:${entry.name.toLowerCase()}`
				: `sheet:${entry.scope.sheetId}:${entry.name.toLowerCase()}`
		if (seenNames.includes(entryKey)) continue

		const parsed = parseFormula(entry.formula)
		if (!parsed.ok) continue

		let formulaSheetIndex = sheetIndex
		if (entry.scope.kind === 'sheet') {
			const scope = entry.scope
			const localSheetIndex = workbook.sheets.findIndex(
				(workbookSheet) => workbookSheet.id === scope.sheetId,
			)
			if (localSheetIndex >= 0) formulaSheetIndex = localSheetIndex
		}

		refs.push(
			...extractRefsWithNames(parsed.value, workbook, sheetNameIndex, formulaSheetIndex, [
				...seenNames,
				entryKey,
			]),
		)
	}
	return refs
}

function collectNameRefs(node: FormulaNode): Array<{ name: string; sheet?: string }> {
	const result: Array<{ name: string; sheet?: string }> = []
	walkNameRefs(node, result)
	return result
}

function collectStructuredRefs(
	node: FormulaNode,
): Array<Extract<FormulaNode, { type: 'structuredRef' }>> {
	const result: Array<Extract<FormulaNode, { type: 'structuredRef' }>> = []
	walkStructuredRefs(node, result)
	return result
}

function walkStructuredRefs(
	node: FormulaNode,
	result: Array<Extract<FormulaNode, { type: 'structuredRef' }>>,
): void {
	switch (node.type) {
		case 'structuredRef':
			result.push(node)
			break
		case 'binary':
			walkStructuredRefs(node.left, result)
			walkStructuredRefs(node.right, result)
			break
		case 'unary':
			walkStructuredRefs(node.operand, result)
			break
		case 'function':
			for (const arg of node.args) walkStructuredRefs(arg, result)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walkStructuredRefs(cell, result)
			}
			break
		default:
			break
	}
}

function walkNameRefs(node: FormulaNode, result: Array<{ name: string; sheet?: string }>): void {
	switch (node.type) {
		case 'name':
			result.push(
				node.sheet !== undefined ? { name: node.name, sheet: node.sheet } : { name: node.name },
			)
			break
		case 'binary':
			walkNameRefs(node.left, result)
			walkNameRefs(node.right, result)
			break
		case 'unary':
			walkNameRefs(node.operand, result)
			break
		case 'function':
			for (const arg of node.args) walkNameRefs(arg, result)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walkNameRefs(cell, result)
			}
			break
		default:
			break
	}
}
