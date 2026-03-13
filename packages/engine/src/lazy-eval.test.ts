import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { EMPTY, errorValue, numberValue } from '@ascend/schema'
import { defaultCalcContext } from './calc-context.ts'
import { cellKey, DependencyGraph } from './dep-graph.ts'
import { LazyEvalContext } from './lazy-eval.ts'

const sid = 0 as StyleId

function setup() {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	const graph = new DependencyGraph()
	const ctx = new LazyEvalContext(wb, defaultCalcContext(), graph)
	return { wb, sheet, graph, ctx }
}

describe('LazyEvalContext', () => {
	test('getValue returns plain cell value for non-formula cells', () => {
		const { sheet, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: sid })
		expect(ctx.getValue(cellKey(0, 0, 0))).toEqual(numberValue(42))
	})

	test('getValue evaluates dirty formula', () => {
		const { sheet, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+B1', styleId: sid })

		const key = cellKey(0, 1, 0)
		ctx.register(key, 'A1+B1')
		const result = ctx.getValue(key)
		expect(result).toEqual(numberValue(15))
	})

	test('clean cell is not re-evaluated', () => {
		const { sheet, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1*2', styleId: sid })

		const key = cellKey(0, 1, 0)
		ctx.register(key, 'A1*2')

		const first = ctx.getValue(key)
		expect(first).toEqual(numberValue(20))
		expect(ctx.isDirty(key)).toBe(false)

		const second = ctx.getValue(key)
		expect(second).toEqual(numberValue(20))
	})

	test('markDirty propagates to dependents', () => {
		const { sheet, graph, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1*2', styleId: sid })

		const a1Key = cellKey(0, 0, 0)
		const a2Key = cellKey(0, 1, 0)
		graph.addFormula(a2Key, [a1Key], false)
		ctx.register(a2Key, 'A1*2')

		ctx.getValue(a2Key)
		expect(ctx.isDirty(a2Key)).toBe(false)

		ctx.markDirty(a1Key)
		expect(ctx.isDirty(a2Key)).toBe(true)
	})

	test('re-evaluation after markDirty picks up new values', () => {
		const { sheet, graph, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1*2', styleId: sid })

		const a1Key = cellKey(0, 0, 0)
		const a2Key = cellKey(0, 1, 0)
		graph.addFormula(a2Key, [a1Key], false)
		ctx.register(a2Key, 'A1*2')

		expect(ctx.getValue(a2Key)).toEqual(numberValue(20))

		sheet.cells.set(0, 0, { value: numberValue(25), formula: null, styleId: sid })
		ctx.markDirty(a1Key)
		expect(ctx.getValue(a2Key)).toEqual(numberValue(50))
	})

	test('recalcVisible evaluates only specified cells', () => {
		const { sheet, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: sid })
		sheet.cells.set(0, 1, { value: numberValue(4), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(1, 1, { value: EMPTY, formula: 'B1+1', styleId: sid })

		const a2Key = cellKey(0, 1, 0)
		const b2Key = cellKey(0, 1, 1)
		ctx.register(a2Key, 'A1+1')
		ctx.register(b2Key, 'B1+1')

		ctx.recalcVisible([a2Key])

		expect(ctx.isDirty(a2Key)).toBe(false)
		expect(ctx.isDirty(b2Key)).toBe(true)
		expect(ctx.getCachedValue(a2Key)).toEqual(numberValue(4))
	})

	test('circular reference returns #REF! instead of stack overflow', () => {
		const { sheet, graph, ctx } = setup()
		sheet.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: sid })
		sheet.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: sid })

		const a1Key = cellKey(0, 0, 0)
		const b1Key = cellKey(0, 0, 1)
		graph.addFormula(a1Key, [b1Key], false)
		graph.addFormula(b1Key, [a1Key], false)
		ctx.register(a1Key, 'B1')
		ctx.register(b1Key, 'A1')

		const result = ctx.getValue(a1Key)
		expect(result).toEqual(errorValue('#REF!'))
	})

	test('transitive dependency chain evaluates correctly', () => {
		const { sheet, graph, ctx } = setup()
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: sid })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+1', styleId: sid })
		sheet.cells.set(2, 0, { value: EMPTY, formula: 'A2+1', styleId: sid })

		const a1Key = cellKey(0, 0, 0)
		const a2Key = cellKey(0, 1, 0)
		const a3Key = cellKey(0, 2, 0)
		graph.addFormula(a2Key, [a1Key], false)
		graph.addFormula(a3Key, [a2Key], false)
		ctx.register(a2Key, 'A1+1')
		ctx.register(a3Key, 'A2+1')

		expect(ctx.getValue(a3Key)).toEqual(numberValue(7))

		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: sid })
		ctx.markDirty(a1Key)
		expect(ctx.getValue(a3Key)).toEqual(numberValue(12))
	})
})
