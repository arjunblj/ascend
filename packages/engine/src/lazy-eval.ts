import type { Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { cachedParseFormula } from '@ascend/formulas'
import type { CellValue } from '@ascend/schema'
import { EMPTY, errorValue } from '@ascend/schema'
import type { CalcContext } from './calc-context.ts'
import { type CompiledFormula, compileFormula, evaluateCompiled } from './compiled-eval.ts'
import { type CellKey, type DependencyGraph, parseCellKeyInto } from './dep-graph.ts'
import { evaluate, MutableEvalContext } from './evaluator.ts'

enum CellState {
	Clean = 0,
	Dirty = 1,
	InProgress = 2,
}

interface FormulaRecord {
	ast: FormulaNode
	compiled: CompiledFormula | false | undefined
	value: CellValue
	state: CellState
}

export class LazyEvalContext {
	private readonly workbook: Workbook
	private readonly graph: DependencyGraph
	private readonly formulas = new Map<CellKey, FormulaRecord>()
	private readonly evalCtx: MutableEvalContext

	constructor(workbook: Workbook, calcContext: CalcContext, graph: DependencyGraph) {
		this.workbook = workbook
		this.graph = graph
		this.evalCtx = new MutableEvalContext()
		this.evalCtx.workbook = workbook
		this.evalCtx.calcContext = calcContext
	}

	register(key: CellKey, formulaText: string): void {
		const parsed = cachedParseFormula(formulaText)
		if (!parsed.ok) return
		this.formulas.set(key, {
			ast: parsed.value,
			compiled: undefined,
			value: EMPTY,
			state: CellState.Dirty,
		})
	}

	registerAst(key: CellKey, ast: FormulaNode): void {
		this.formulas.set(key, {
			ast,
			compiled: undefined,
			value: EMPTY,
			state: CellState.Dirty,
		})
	}

	markDirty(key: CellKey): void {
		const visited = new Set<CellKey>()
		const queue = [key]
		while (queue.length > 0) {
			const current = queue.pop() as CellKey
			if (visited.has(current)) continue
			visited.add(current)
			const record = this.formulas.get(current)
			if (record && record.state !== CellState.Dirty) {
				record.state = CellState.Dirty
			}
			for (const dep of this.graph.getDependents(current)) {
				if (!visited.has(dep)) queue.push(dep)
			}
		}
	}

	getValue(key: CellKey): CellValue {
		const record = this.formulas.get(key)
		if (!record) {
			const coords = { sheetIndex: 0, row: 0, col: 0 }
			parseCellKeyInto(key, coords)
			const sheet = this.workbook.sheets[coords.sheetIndex]
			return sheet?.cells.readValue(coords.row, coords.col) ?? EMPTY
		}
		if (record.state === CellState.Clean) return record.value
		if (record.state === CellState.InProgress) return errorValue('#REF!')
		return this.evaluateCell(key, record)
	}

	recalcVisible(keys: CellKey[]): void {
		for (const key of keys) {
			this.getValue(key)
		}
	}

	isDirty(key: CellKey): boolean {
		const record = this.formulas.get(key)
		return record !== undefined && record.state === CellState.Dirty
	}

	getCachedValue(key: CellKey): CellValue | undefined {
		const record = this.formulas.get(key)
		return record?.value
	}

	private evaluateCell(key: CellKey, record: FormulaRecord): CellValue {
		record.state = CellState.InProgress

		const coords = { sheetIndex: 0, row: 0, col: 0 }
		parseCellKeyInto(key, coords)
		const sheet = this.workbook.sheets[coords.sheetIndex]
		const cell = sheet?.cells.get(coords.row, coords.col)

		if (sheet && cell) {
			sheet.cells.setResolved(
				coords.row,
				coords.col,
				errorValue('#REF!'),
				cell.formula,
				cell.styleId,
			)
		}

		this.ensurePrecedentsClean(key)

		this.evalCtx.sheetIndex = coords.sheetIndex
		this.evalCtx.row = coords.row
		this.evalCtx.col = coords.col

		let value: CellValue
		if (record.compiled === undefined) {
			const result = compileFormula(record.ast)
			record.compiled = result ?? false
		}
		if (record.compiled !== false) {
			value = evaluateCompiled(record.compiled, this.evalCtx)
		} else {
			value = evaluate(record.ast, this.evalCtx)
		}

		record.value = value
		record.state = CellState.Clean

		if (sheet && cell) {
			sheet.cells.setResolved(coords.row, coords.col, value, cell.formula, cell.styleId)
		}

		return value
	}

	private ensurePrecedentsClean(key: CellKey): void {
		const prec = this.graph.getPrecedents(key)
		for (const dep of prec.cells) {
			const depRecord = this.formulas.get(dep)
			if (depRecord && depRecord.state === CellState.Dirty) {
				this.evaluateCell(dep, depRecord)
			}
		}
	}
}
