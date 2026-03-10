import { parseA1, toA1, type Workbook } from '@ascend/core'
import { analyzeWorkbook, cellKey, parseCellKey, type WorkbookAnalysis } from '@ascend/engine'
import type { FormulaRef } from '@ascend/formulas'
import { ascendError, type CellValue, EMPTY, err, ok, type Result } from '@ascend/schema'

export interface TraceResult {
	readonly cell: string
	readonly sheet: string
	readonly formula: string | null
	readonly value: CellValue
	readonly precedents: readonly TraceNode[]
	readonly dependents: readonly TraceNode[]
}

export interface TraceNode {
	readonly ref: string
	readonly sheet: string
	readonly formula: string | null
	readonly value: CellValue
	readonly depth: number
}

function resolveCell(
	wb: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): { formula: string | null; value: CellValue } {
	const sheet = wb.sheets[sheetIndex]
	if (!sheet) return { formula: null, value: EMPTY }
	const cell = sheet.cells.get(row, col)
	if (!cell) return { formula: null, value: EMPTY }
	return { formula: cell.formula, value: cell.value }
}

export function trace(
	workbook: Workbook,
	sheetName: string,
	ref: string,
	opts?: { maxDepth?: number },
	analysis?: WorkbookAnalysis,
): Result<TraceResult> {
	const sheetIndex = workbook.sheets.findIndex(
		(s) => s.name.toLowerCase() === sheetName.toLowerCase(),
	)
	if (sheetIndex === -1) {
		return err(ascendError('SHEET_NOT_FOUND', `Sheet "${sheetName}" not found`))
	}

	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return err(ascendError('SHEET_NOT_FOUND', `Sheet "${sheetName}" not found`))
	const maxDepth = opts?.maxDepth ?? 10

	let cellRef: { row: number; col: number }
	try {
		cellRef = parseA1(ref)
	} catch {
		return err(ascendError('INVALID_REF', `Invalid cell reference: ${ref}`))
	}

	const cell = sheet.cells.get(cellRef.row, cellRef.col)
	const formula = cell?.formula ?? null
	const value = cell?.value ?? EMPTY

	const compiled = analysis ?? analyzeWorkbook(workbook)
	const graph = compiled.dependencyGraph
	const targetKey = cellKey(sheetIndex, cellRef.row, cellRef.col)
	const targetFormula = compiled.formulas.get(targetKey)

	const precedents: TraceNode[] = []
	const visitedPre = new Set<string>()
	const prQueue: { key: string; depth: number }[] = [{ key: targetKey, depth: 0 }]
	visitedPre.add(targetKey)
	let preIndex = 0

	while (preIndex < prQueue.length) {
		const item = prQueue[preIndex++]
		if (!item) break
		if (item.depth >= maxDepth) continue

		if (item.key === targetKey && targetFormula) {
			for (const refNode of targetFormula.refs) {
				const traceNode = formulaRefToTraceNode(workbook, sheetIndex, refNode, item.depth + 1)
				if (!traceNode) continue
				const visitKey = `${traceNode.sheet}!${traceNode.ref}`
				if (visitedPre.has(visitKey)) continue
				visitedPre.add(visitKey)
				precedents.push(traceNode)
				if (refNode.kind === 'cell') {
					const refSheetIndex = resolveFormulaRefSheetIndex(workbook, sheetIndex, refNode)
					if (refSheetIndex >= 0) {
						prQueue.push({
							key: cellKey(refSheetIndex, refNode.ref.row, refNode.ref.col),
							depth: item.depth + 1,
						})
					}
				}
			}
			continue
		}

		for (const pred of graph.getPrecedents(item.key)) {
			if (visitedPre.has(pred)) continue
			visitedPre.add(pred)
			const [si, r, c] = parseCellKey(pred)
			const s = workbook.sheets[si]
			if (!s) continue
			const resolved = resolveCell(workbook, si, r, c)
			precedents.push({
				ref: toA1({ row: r, col: c }),
				sheet: s.name,
				formula: resolved.formula,
				value: resolved.value,
				depth: item.depth + 1,
			})
			prQueue.push({ key: pred, depth: item.depth + 1 })
		}
	}

	const dependents: TraceNode[] = []
	const visitedDep = new Set<string>()
	const depQueue: { key: string; depth: number }[] = [{ key: targetKey, depth: 0 }]
	visitedDep.add(targetKey)
	let depIndex = 0

	while (depIndex < depQueue.length) {
		const item = depQueue[depIndex++]
		if (!item) break
		if (item.depth >= maxDepth) continue

		const deps = graph.getDependents(item.key)
		for (const dep of deps) {
			if (visitedDep.has(dep)) continue
			visitedDep.add(dep)
			const [si, r, c] = parseCellKey(dep)
			const s = workbook.sheets[si]
			if (!s) continue
			const resolved = resolveCell(workbook, si, r, c)
			dependents.push({
				ref: toA1({ row: r, col: c }),
				sheet: s.name,
				formula: resolved.formula,
				value: resolved.value,
				depth: item.depth + 1,
			})
			depQueue.push({ key: dep, depth: item.depth + 1 })
		}
	}

	return ok({
		cell: ref,
		sheet: sheetName,
		formula,
		value,
		precedents,
		dependents,
	})
}

function formulaRefToTraceNode(
	workbook: Workbook,
	currentSheetIndex: number,
	ref: FormulaRef,
	depth: number,
): TraceNode | null {
	const sheetIndex = resolveFormulaRefSheetIndex(workbook, currentSheetIndex, ref)
	if (sheetIndex < 0) return null
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return null
	if (ref.kind === 'cell') {
		const resolved = resolveCell(workbook, sheetIndex, ref.ref.row, ref.ref.col)
		return {
			ref: toA1({ row: ref.ref.row, col: ref.ref.col }),
			sheet: sheet.name,
			formula: resolved.formula,
			value: resolved.value,
			depth,
		}
	}
	if (ref.kind === 'range') {
		const resolved = resolveCell(workbook, sheetIndex, ref.start.row, ref.start.col)
		return {
			ref: `${toA1({ row: ref.start.row, col: ref.start.col })}:${toA1({ row: ref.end.row, col: ref.end.col })}`,
			sheet: sheet.name,
			formula: resolved.formula,
			value: resolved.value,
			depth,
		}
	}
	if (ref.kind === 'wholeRowRange') {
		const used = sheet.cells.usedRange()
		const firstCol = used?.start.col ?? 0
		const resolved = resolveCell(workbook, sheetIndex, ref.startRow, firstCol)
		return {
			ref: `${ref.startRow + 1}:${ref.endRow + 1}`,
			sheet: sheet.name,
			formula: resolved.formula,
			value: resolved.value,
			depth,
		}
	}
	const used = sheet.cells.usedRange()
	const firstRow = used?.start.row ?? 0
	const resolved = resolveCell(workbook, sheetIndex, firstRow, ref.startCol)
	return {
		ref: `${toA1({ row: firstRow, col: ref.startCol }).replace(/\d+$/, '')}:${toA1({ row: firstRow, col: ref.endCol }).replace(/\d+$/, '')}`,
		sheet: sheet.name,
		formula: resolved.formula,
		value: resolved.value,
		depth,
	}
}

function resolveFormulaRefSheetIndex(
	workbook: Workbook,
	currentSheetIndex: number,
	ref: FormulaRef,
): number {
	if (!ref.sheet) return currentSheetIndex
	return workbook.sheets.findIndex((sheet) => sheet.name.toLowerCase() === ref.sheet?.toLowerCase())
}
