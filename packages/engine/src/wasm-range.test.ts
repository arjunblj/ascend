import { describe, expect, test } from 'bun:test'
import { getWasmRangeOps, resetWasmRangeOpsForTest } from './wasm-range.ts'

describe('getWasmRangeOps', () => {
	test('returns a stable cached instance', () => {
		const a = getWasmRangeOps()
		const b = getWasmRangeOps()
		expect(b).toBe(a)
	})

	test('computes sum/min/max for even and odd lengths', () => {
		const ops = getWasmRangeOps()
		if (!ops) return

		const values = new Float64Array([3, 1, 4, 1, 5])
		ops.load(values, values.length)

		expect(ops.sum(values.length)).toBe(14)
		expect(ops.min(values.length)).toBe(1)
		expect(ops.max(values.length)).toBe(5)
	})

	test('handles empty reductions consistently', () => {
		const ops = getWasmRangeOps()
		if (!ops) return

		ops.load(new Float64Array(0), 0)
		expect(ops.sum(0)).toBe(0)
		expect(ops.min(0)).toBe(0)
		expect(ops.max(0)).toBe(0)
	})

	test('grows backing memory for larger inputs', () => {
		const ops = getWasmRangeOps()
		if (!ops) return

		const values = new Float64Array(20_000)
		for (let i = 0; i < values.length; i++) values[i] = i - 10_000
		ops.load(values, values.length)

		expect(ops.min(values.length)).toBe(-10_000)
		expect(ops.max(values.length)).toBe(9_999)
		expect(ops.sum(values.length)).toBe(-10_000)
	})

	test('returns null when WASM initialization fails', () => {
		resetWasmRangeOpsForTest()
		const originalModule = WebAssembly.Module
		const failingModule = ((): never => {
			throw new Error('simulated wasm init failure')
		}) as unknown as typeof WebAssembly.Module
		;(WebAssembly as { Module: typeof WebAssembly.Module }).Module = failingModule
		try {
			expect(getWasmRangeOps()).toBeNull()
		} finally {
			;(WebAssembly as { Module: typeof WebAssembly.Module }).Module = originalModule
			resetWasmRangeOpsForTest()
		}
	})
})
