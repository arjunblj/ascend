import { parseA1, toA1, type Workbook } from '@ascend/core'
import { analyzeWorkbook, cellKey, parseCellKey } from '@ascend/engine'
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

	const graph = analyzeWorkbook(workbook).dependencyGraph
	const targetKey = cellKey(sheetIndex, cellRef.row, cellRef.col)

	const precedents: TraceNode[] = []
	const visitedPre = new Set<string>()
	const prQueue: { key: string; depth: number }[] = [{ key: targetKey, depth: 0 }]
	visitedPre.add(targetKey)

	while (prQueue.length > 0) {
		const item = prQueue.shift()
		if (!item) break
		if (item.depth >= maxDepth) continue

		const preds = graph.getPrecedents(item.key)
		for (const pred of preds) {
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

	while (depQueue.length > 0) {
		const item = depQueue.shift()
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
