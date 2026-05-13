import { type Cell, type CellFormulaBinding, type Sheet, toA1, type Workbook } from '@ascend/core'
import { type FormulaNode, parseFormula } from '@ascend/formulas'

const ARRAY_MAPPABLE_FUNCTIONS = new Set(['SQRT'])

export function inferLegacyArrayFormulaBlocks(
	workbook: Workbook,
	candidateSheets: readonly Sheet[] = workbook.sheets,
): readonly string[] {
	const inferredSheets = new Set<string>()
	for (const sheet of candidateSheets) {
		if (inferSheetLegacyArrayFormulaBlocks(workbook, sheet)) inferredSheets.add(sheet.name)
	}
	return [...inferredSheets]
}

function inferSheetLegacyArrayFormulaBlocks(workbook: Workbook, sheet: Sheet): boolean {
	let inferred = false
	for (const [row, rowCells] of sheet.cells.iterateRows()) {
		for (const [col, cell] of rowCells) {
			if (cell.formula === null || cell.formulaInfo !== undefined) continue
			const parsed = parseFormula(cell.formula)
			if (!parsed.ok || !canProduceArray(parsed.value, workbook, sheet, new Set())) continue
			const block = findRelayBlock(sheet, row, col)
			if (!block) continue
			applyLegacyArrayBinding(sheet, row, col, block.endRow, block.endCol)
			inferred = true
		}
	}
	return inferred
}

function canProduceArray(
	node: FormulaNode,
	workbook: Workbook,
	sheet: Sheet,
	seenNames: Set<string>,
): boolean {
	switch (node.type) {
		case 'array':
		case 'rangeRef':
		case 'dynamicRangeRef':
		case 'wholeColumnRange':
		case 'wholeRowRange':
		case 'structuredRef':
		case 'spillRef':
		case 'sheetSpanRef':
			return true
		case 'name':
			return definedNameCanProduceArray(node, workbook, sheet, seenNames)
		case 'unary':
			return canProduceArray(node.operand, workbook, sheet, seenNames)
		case 'binary':
			return (
				node.op !== ',' &&
				node.op !== ' ' &&
				(canProduceArray(node.left, workbook, sheet, seenNames) ||
					canProduceArray(node.right, workbook, sheet, seenNames))
			)
		case 'function':
			return (
				ARRAY_MAPPABLE_FUNCTIONS.has(node.name.toUpperCase()) &&
				node.args.some((arg) => canProduceArray(arg, workbook, sheet, seenNames))
			)
		default:
			return false
	}
}

function definedNameCanProduceArray(
	node: Extract<FormulaNode, { type: 'name' }>,
	workbook: Workbook,
	sheet: Sheet,
	seenNames: Set<string>,
): boolean {
	const explicitSheet = node.sheet
		? workbook.sheets.find(
				(candidate) => candidate.name.toLowerCase() === node.sheet?.toLowerCase(),
			)
		: undefined
	const name = workbook.definedNames.resolve(node.name, sheet.id, explicitSheet?.id)
	if (!name) return false
	const key = `${name.scope.kind}:${name.scope.kind === 'sheet' ? name.scope.sheetId : ''}:${name.name.toLowerCase()}`
	if (seenNames.has(key)) return false
	seenNames.add(key)
	const parsed = parseFormula(name.formula)
	if (!parsed.ok) return false
	return canProduceArray(parsed.value, workbook, sheet, seenNames)
}

function findRelayBlock(
	sheet: Sheet,
	anchorRow: number,
	anchorCol: number,
): { readonly endRow: number; readonly endCol: number } | null {
	let endCol = anchorCol
	while (isRelayCell(sheet.cells.get(anchorRow, endCol + 1), sheet, anchorRow, anchorCol)) {
		endCol++
	}

	let endRow = anchorRow
	while (isRelayCell(sheet.cells.get(endRow + 1, anchorCol), sheet, anchorRow, anchorCol)) {
		endRow++
	}

	if (endRow === anchorRow && endCol === anchorCol) return null
	for (let row = anchorRow; row <= endRow; row++) {
		for (let col = anchorCol; col <= endCol; col++) {
			if (row === anchorRow && col === anchorCol) continue
			if (!isRelayCell(sheet.cells.get(row, col), sheet, anchorRow, anchorCol)) return null
		}
	}
	return { endRow, endCol }
}

function isRelayCell(
	cell: Cell | undefined,
	sheet: Sheet,
	anchorRow: number,
	anchorCol: number,
): boolean {
	if (!cell?.formula || cell.formulaInfo !== undefined) return false
	const parsed = parseFormula(cell.formula)
	if (!parsed.ok || parsed.value.type !== 'cellRef') return false
	const ref = parsed.value
	if (ref.sheet !== undefined && ref.sheet.toLowerCase() !== sheet.name.toLowerCase()) return false
	return ref.ref.row === anchorRow && ref.ref.col === anchorCol
}

function applyLegacyArrayBinding(
	sheet: Sheet,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): void {
	const ref = `${toA1({ row: startRow, col: startCol })}:${toA1({ row: endRow, col: endCol })}`
	const binding: CellFormulaBinding = { kind: 'array', ref }
	for (let row = startRow; row <= endRow; row++) {
		for (let col = startCol; col <= endCol; col++) {
			const cell = sheet.cells.get(row, col)
			if (!cell) continue
			const formula = row === startRow && col === startCol ? cell.formula : null
			sheet.cells.setResolved(row, col, cell.value, formula, cell.styleId, binding)
		}
	}
}
