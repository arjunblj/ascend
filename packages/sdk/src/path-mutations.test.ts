import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('path-addressed mutations', () => {
	test('compiles cell, formula, range, sheet, name, and table paths into replayable ops', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])

		const result = wb.compilePathMutations([
			{ path: 'sheets.Sheet1.cells.C1.value', value: 10 },
			{ path: '/sheets/Sheet1/cells/D1/formula', value: 'C1*2' },
			{ path: ['sheets', 'Sheet1', 'ranges', 'E1:F2', 'clear'], value: 'values' },
			{ path: '/names/ReportDate/ref', value: "'Sheet1'!$C$1" },
			{ path: '/tables/Sales/rows/append', value: [['Bob', 42]] },
			{ path: '/sheets/Sheet1/name', value: 'Summary' },
		])

		expect(result.replayable).toBe(true)
		expect(result.issues).toEqual([])
		expect(result.ops).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C1', value: 10 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'D1', formula: 'C1*2' },
			{ op: 'clearRange', sheet: 'Sheet1', range: 'E1:F2', what: 'values' },
			{ op: 'setDefinedName', name: 'ReportDate', ref: "'Sheet1'!$C$1" },
			{ op: 'appendRows', table: 'Sales', rows: [['Bob', 42]] },
			{ op: 'renameSheet', sheet: 'Sheet1', newName: 'Summary' },
		])

		const applied = wb.batch(result.ops)
		expect(applied.errors).toEqual([])
		expect(wb.sheets).toContain('Summary')
		expect(wb.table('Sales')?.rowCount).toBe(2)
	})

	test('supports JSON Pointer and escaped-dot paths for names that contain punctuation', () => {
		const wb = AscendWorkbook.create()
		wb.renameSheet('Sheet1', 'Q1.Forecast')

		const result = wb.compilePathMutations([
			{ path: '/sheets/Q1.Forecast/cells/A1/value', value: 'pointer' },
			{ path: 'sheets.Q1\\.Forecast.cells.A2.value', value: 'escaped' },
		])

		expect(result.replayable).toBe(true)
		expect(result.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Q1.Forecast',
				updates: [
					{ ref: 'A1', value: 'pointer' },
					{ ref: 'A2', value: 'escaped' },
				],
			},
		])
	})

	test('reports invalid paths and values without emitting unsafe operations', () => {
		const wb = AscendWorkbook.create()

		const result = wb.compilePathMutations([
			{ path: '/sheets/Missing/cells/A1/value', value: 1 },
			{ path: '/sheets/Sheet1/cells/NotARef/value', value: 1 },
			{ path: '/sheets/Sheet1/cells/A1/value', value: { nested: true } },
			{ path: '/tables/Missing/rows/append', value: [[1]] },
			{ path: '/workbook/properties/title', value: 'Unsupported' },
		])

		expect(result.replayable).toBe(false)
		expect(result.ops).toEqual([])
		expect(result.issues.map((entry) => entry.code)).toEqual([
			'sheet_not_found',
			'invalid_ref',
			'invalid_value',
			'table_not_found',
			'unsupported_path',
		])
	})
})
