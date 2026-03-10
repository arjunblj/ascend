import { parseA1, toA1, type Workbook } from '@ascend/core'
import {
	type AnalyzedFormula,
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	cellKey,
	parseCellKey,
	resolveFormulaDependencies,
	type WorkbookDependencyAnalysis,
	type WorkbookFormulaAnalysis,
} from '@ascend/engine'
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
	analysis?: {
		readonly formulas?: WorkbookFormulaAnalysis
		readonly dependencies?: WorkbookDependencyAnalysis
	},
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

	const formulas = analysis?.formulas ?? analyzeWorkbookFormulas(workbook)
	const dependencies = analysis?.dependencies ?? analyzeWorkbookDependencies(workbook)
	const graph = dependencies.dependencyGraph
	const targetKey = cellKey(sheetIndex, cellRef.row, cellRef.col)
	const targetFormula = targetKey ? formulas.formulas.get(targetKey) : undefined

	const precedents: TraceNode[] = []
	const visitedPre = new Set<string>()
	const prQueue: { key: string; depth: number }[] = [{ key: targetKey, depth: 0 }]
	visitedPre.add(targetKey)
	let preIndex = 0

	while (preIndex < prQueue.length) {
		const item = prQueue[preIndex++]
		if (!item) break
		if (item.depth >= maxDepth) continue

		const currentIndexed = item.key === targetKey ? targetFormula : formulas.formulas.get(item.key)
		if (currentIndexed) {
			const currentFormula = resolveFormulaDependencies(
				workbook,
				formulas.sheetNameIndex,
				currentIndexed,
			)
			addSymbolicPrecedents(
				workbook,
				currentFormula,
				item.depth + 1,
				visitedPre,
				precedents,
				prQueue,
			)
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

function addSymbolicPrecedents(
	workbook: Workbook,
	formula: AnalyzedFormula,
	depth: number,
	visitedPre: Set<string>,
	precedents: TraceNode[],
	prQueue: Array<{ key: string; depth: number }>,
): void {
	const dependencyKeys = new Set<string>()
	for (const refNode of formula.refs) {
		const traceNode = formulaRefToTraceNode(workbook, formula.sheetIndex, refNode, depth)
		if (!traceNode) continue
		const visitKey = `${traceNode.sheet}!${traceNode.ref}`
		if (visitedPre.has(visitKey)) continue
		visitedPre.add(visitKey)
		precedents.push(traceNode)
		const dependencyKey = formulaRefToDependencyKey(workbook, formula.sheetIndex, refNode)
		if (dependencyKey) dependencyKeys.add(dependencyKey)
		if (refNode.kind === 'cell') {
			const refSheetIndex = resolveFormulaRefSheetIndex(workbook, formula.sheetIndex, refNode)
			if (refSheetIndex >= 0) {
				prQueue.push({
					key: cellKey(refSheetIndex, refNode.ref.row, refNode.ref.col),
					depth,
				})
			}
		}
	}
	for (const dep of formula.deps) {
		if (dependencyKeys.has(dep) || visitedPre.has(dep)) continue
		dependencyKeys.add(dep)
		visitedPre.add(dep)
		const [si, r, c] = parseCellKey(dep)
		const s = workbook.sheets[si]
		if (!s) continue
		const resolved = resolveCell(workbook, si, r, c)
		precedents.push({
			ref: toA1({ row: r, col: c }),
			sheet: s.name,
			formula: resolved.formula,
			value: resolved.value,
			depth,
		})
		prQueue.push({ key: dep, depth })
	}
	for (const rangeDep of formula.rangeDeps) {
		const depKey = rangeDependencyKey(rangeDep)
		if (dependencyKeys.has(depKey)) continue
		dependencyKeys.add(depKey)
		const traceNode = rangeDependencyToTraceNode(workbook, rangeDep, depth)
		if (!traceNode) continue
		const visitKey = `${traceNode.sheet}!${traceNode.ref}`
		if (visitedPre.has(visitKey)) continue
		visitedPre.add(visitKey)
		precedents.push(traceNode)
	}
}

function formulaRefToTraceNode(
	workbook: Workbook,
	currentSheetIndex: number,
	ref: FormulaRef,
	depth: number,
): TraceNode | null {
	if (ref.kind === 'sheetSpan') {
		const sheetIndices = resolveSheetSpanIndices(workbook, ref.startSheet, ref.endSheet)
		if (!sheetIndices) return null
		const firstSheetIndex = sheetIndices[0]
		if (firstSheetIndex === undefined) return null
		const firstSheet = workbook.sheets[firstSheetIndex]
		if (!firstSheet) return null
		const target = ref.target
		if (target.kind === 'sheetSpan') return null
		const resolved =
			target.kind === 'cell'
				? resolveCell(workbook, firstSheetIndex, target.ref.row, target.ref.col)
				: target.kind === 'range'
					? resolveCell(workbook, firstSheetIndex, target.start.row, target.start.col)
					: target.kind === 'wholeRowRange'
						? resolveCell(
								workbook,
								firstSheetIndex,
								target.startRow,
								firstSheet.cells.usedRange()?.start.col ?? 0,
							)
						: resolveCell(
								workbook,
								firstSheetIndex,
								firstSheet.cells.usedRange()?.start.row ?? 0,
								target.startCol,
							)
		return {
			ref: formulaRefTargetText(target),
			sheet: `${ref.startSheet}:${ref.endSheet}`,
			formula: resolved.formula,
			value: resolved.value,
			depth,
		}
	}
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
	if (ref.kind === 'sheetSpan') return -1
	if (!ref.sheet) return currentSheetIndex
	return workbook.sheets.findIndex((sheet) => sheet.name.toLowerCase() === ref.sheet?.toLowerCase())
}

function formulaRefToDependencyKey(
	workbook: Workbook,
	currentSheetIndex: number,
	ref: FormulaRef,
): string | null {
	if (ref.kind === 'sheetSpan') {
		if (ref.target.kind === 'sheetSpan') return null
		return `sheetSpan:${ref.startSheet}:${ref.endSheet}:${formulaRefTargetText(ref.target)}`
	}
	const sheetIndex = resolveFormulaRefSheetIndex(workbook, currentSheetIndex, ref)
	if (sheetIndex < 0) return null
	if (ref.kind === 'cell') return cellKey(sheetIndex, ref.ref.row, ref.ref.col)
	const rangeDep = formulaRefToRangeDependency(workbook, sheetIndex, ref)
	return rangeDep ? rangeDependencyKey(rangeDep) : null
}

function formulaRefToRangeDependency(
	workbook: Workbook,
	sheetIndex: number,
	ref: Exclude<FormulaRef, { kind: 'cell' } | { kind: 'sheetSpan' }>,
): {
	sheetIndex: number
	startRow: number
	startCol: number
	endRow: number
	endCol: number
} | null {
	if (ref.kind === 'range') {
		return {
			sheetIndex,
			startRow: ref.start.row,
			startCol: ref.start.col,
			endRow: ref.end.row,
			endCol: ref.end.col,
		}
	}
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return null
	const used = sheet.cells.usedRange()
	if (!used) return null
	if (ref.kind === 'wholeRowRange') {
		return {
			sheetIndex,
			startRow: ref.startRow,
			startCol: used.start.col,
			endRow: ref.endRow,
			endCol: used.end.col,
		}
	}
	return {
		sheetIndex,
		startRow: used.start.row,
		startCol: ref.startCol,
		endRow: used.end.row,
		endCol: ref.endCol,
	}
}

function rangeDependencyToTraceNode(
	workbook: Workbook,
	rangeDep: {
		sheetIndex: number
		startRow: number
		startCol: number
		endRow: number
		endCol: number
	},
	depth: number,
): TraceNode | null {
	const sheet = workbook.sheets[rangeDep.sheetIndex]
	if (!sheet) return null
	const resolved = resolveCell(workbook, rangeDep.sheetIndex, rangeDep.startRow, rangeDep.startCol)
	return {
		ref: `${toA1({ row: rangeDep.startRow, col: rangeDep.startCol })}:${toA1({ row: rangeDep.endRow, col: rangeDep.endCol })}`,
		sheet: sheet.name,
		formula: resolved.formula,
		value: resolved.value,
		depth,
	}
}

function rangeDependencyKey(rangeDep: {
	sheetIndex: number
	startRow: number
	startCol: number
	endRow: number
	endCol: number
}): string {
	return `range:${rangeDep.sheetIndex}:${rangeDep.startRow}:${rangeDep.startCol}:${rangeDep.endRow}:${rangeDep.endCol}`
}

function resolveSheetSpanIndices(
	workbook: Workbook,
	startSheet: string,
	endSheet: string,
): number[] | null {
	const start = workbook.sheets.findIndex(
		(sheet) => sheet.name.toLowerCase() === startSheet.toLowerCase(),
	)
	const end = workbook.sheets.findIndex(
		(sheet) => sheet.name.toLowerCase() === endSheet.toLowerCase(),
	)
	if (start === -1 || end === -1 || start > end) return null
	return Array.from({ length: end - start + 1 }, (_, offset) => start + offset)
}

function formulaRefTargetText(ref: Exclude<FormulaRef, { kind: 'sheetSpan' }>): string {
	if (ref.kind === 'cell') return toA1({ row: ref.ref.row, col: ref.ref.col })
	if (ref.kind === 'range') {
		return `${toA1({ row: ref.start.row, col: ref.start.col })}:${toA1({ row: ref.end.row, col: ref.end.col })}`
	}
	if (ref.kind === 'wholeRowRange') return `${ref.startRow + 1}:${ref.endRow + 1}`
	return `${toA1({ row: 0, col: ref.startCol }).replace(/\d+$/, '')}:${toA1({ row: 0, col: ref.endCol }).replace(/\d+$/, '')}`
}
