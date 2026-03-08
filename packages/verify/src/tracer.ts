import { parseA1, toA1, type Workbook } from '@ascend/core'
import { cellKey, DependencyGraph, parseCellKey } from '@ascend/engine'
import { extractRefs, parseFormula } from '@ascend/formulas'
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

function buildGraph(wb: Workbook): DependencyGraph {
	const graph = new DependencyGraph()

	for (let si = 0; si < wb.sheets.length; si++) {
		const sheet = wb.sheets[si]
		if (!sheet) continue
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			const parsed = parseFormula(cell.formula)
			if (!parsed.ok) continue

			const refs = extractRefs(parsed.value)
			const deps: string[] = []
			let isVolatile = false

			if (parsed.value.type === 'function') {
				const name = parsed.value.name.toUpperCase()
				if (name === 'NOW' || name === 'TODAY' || name === 'RAND' || name === 'RANDBETWEEN') {
					isVolatile = true
				}
			}

			for (const ref of refs) {
				const targetSheet = ref.sheet
					? wb.sheets.findIndex((s) => s.name.toLowerCase() === ref.sheet?.toLowerCase())
					: si
				if (targetSheet === -1) continue

				if (ref.kind === 'cell') {
					deps.push(cellKey(targetSheet, ref.ref.row, ref.ref.col))
				} else {
					for (let r = ref.start.row; r <= ref.end.row; r++) {
						for (let c = ref.start.col; c <= ref.end.col; c++) {
							deps.push(cellKey(targetSheet, r, c))
						}
					}
				}
			}

			graph.addFormula(cellKey(si, row, col), deps, isVolatile)
		}
	}

	return graph
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

	const graph = buildGraph(workbook)
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
