import { describe, expect, test } from 'bun:test'
import { AscendWorkbook, SUPPORTED_PATH_MUTATION_SHAPES } from './index.ts'

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

		expect(result.issues).toEqual([])
		expect(result.replayable).toBe(true)
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

	test('keeps JSON Pointer and escaped-dot paths equivalent for agent-addressed names', () => {
		const sheetName = "Q1.Forecast's Café Δ"
		const tableName = 'Sales.Δ'
		const columnName = 'Gross.Profit'
		const definedName = "Revenue/North ~ Café's Δ"
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: sheetName },
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: columnName },
					{ ref: 'A2', value: 'North' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: sheetName, ref: 'A1:B2', name: tableName, hasHeaders: true },
		])

		const result = wb.compilePathMutations([
			{ path: `/sheets/${pointerSegment(sheetName)}/cells/A1/value`, value: 'pointer' },
			{ path: `sheets.${dotSegment(sheetName)}.cells.A2.value`, value: 'dot' },
			{
				path: `/tables/${pointerSegment(tableName)}/columns/${pointerSegment(columnName)}/formula`,
				value: 'SUM([Gross.Profit])',
			},
			{
				path: `tables.${dotSegment(tableName)}.columns.${dotSegment(columnName)}.totalsRowLabel`,
				value: 'Total',
			},
			{ path: `/names/${pointerSegment(definedName)}/ref`, value: `'${sheetName}'!$B$2` },
			{ path: `names.${dotSegment(definedName)}.ref`, value: `'${sheetName}'!$B$2` },
		])

		expect(result.issues).toEqual([])
		expect(result.replayable).toBe(true)
		expect(result.ops).toEqual([
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'pointer' },
					{ ref: 'A2', value: 'dot' },
				],
			},
			{
				op: 'setTableColumn',
				table: tableName,
				column: columnName,
				formula: 'SUM([Gross.Profit])',
			},
			{ op: 'setTableColumn', table: tableName, column: columnName, totalsRowLabel: 'Total' },
			{ op: 'setDefinedName', name: definedName, ref: `'${sheetName}'!$B$2` },
			{ op: 'setDefinedName', name: definedName, ref: `'${sheetName}'!$B$2` },
		])
	})

	test('compiles sheet metadata and table metadata paths into canonical ops', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Revenue' },
					{ ref: 'A2', value: 'Acme' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])

		const result = wb.compilePathMutations([
			{ path: '/sheets/Sheet1/cells/G2/comment', value: { text: 'review', author: 'agent' } },
			{
				path: '/sheets/Sheet1/cells/G2/hyperlink',
				value: { url: 'https://example.com', display: 'Open' },
			},
			{ path: '/sheets/Sheet1/ranges/C2:C10/numberFormat', value: '0.00' },
			{ path: '/sheets/Sheet1/ranges/C2:C10/style', value: { font: { bold: true } } },
			{
				path: '/sheets/Sheet1/ranges/C2:C10/validation',
				value: { type: 'whole', operator: 'greaterThan', formula1: '0' },
			},
			{
				path: '/sheets/Sheet1/ranges/C2:C10/conditionalFormat',
				value: {
					rule: { type: 'cellIs', operator: 'greaterThan', formula: '10' },
					mode: 'append',
					reassignPriorities: true,
				},
			},
			{ path: '/sheets/Sheet1/ranges/H1:I1/merge', value: true },
			{
				path: '/sheets/Sheet1/autofilter',
				value: { range: 'A1:B10', column: 1, values: ['Open'] },
			},
			{ path: '/tables/Sales/columns/Revenue/formula', value: 'SUM([Revenue])' },
			{
				path: '/tables/Sales/style',
				value: { styleName: 'TableStyleMedium2', showRowStripes: true },
			},
			{ path: '/tables/Sales/name', value: 'SalesData' },
		])

		expect(result.replayable).toBe(true)
		expect(result.issues).toEqual([])
		expect(result.ops).toEqual([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'G2', text: 'review', author: 'agent' },
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'G2',
				url: 'https://example.com',
				display: 'Open',
			},
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C2:C10', format: '0.00' },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'C2:C10', style: { font: { bold: true } } },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'C2:C10',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'C2:C10',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '10' },
				mode: 'append',
				reassignPriorities: true,
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'H1:I1' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:B10', column: 1, values: ['Open'] },
			{ op: 'setTableColumn', table: 'Sales', column: 'Revenue', formula: 'SUM([Revenue])' },
			{
				op: 'setTableStyle',
				table: 'Sales',
				styleName: 'TableStyleMedium2',
				showRowStripes: true,
			},
			{ op: 'renameTable', table: 'Sales', newName: 'SalesData' },
		])

		const applied = wb.batch(result.ops)
		expect(applied.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.comment('G2')?.text).toBe('review')
		expect(wb.sheet('Sheet1')?.hyperlink('G2')?.target).toBe('https://example.com')
		expect(wb.table('SalesData')?.styleInfo?.name).toBe('TableStyleMedium2')
	})

	test('reports invalid paths and values without emitting unsafe operations', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])

		const result = wb.compilePathMutations([
			{ path: '/sheets/Missing/cells/A1/value', value: 1 },
			{ path: '/sheets/Sheet1/cells/NotARef/value', value: 1 },
			{ path: '/sheets/Sheet1/cells/A1/value', value: { nested: true } },
			{ path: '/tables/Missing/rows/append', value: [[1]] },
			{ path: '/tables/Sales/columns/Missing/formula', value: '1+1' },
			{ path: '/sheets/Sheet1/cells/A1/hyperlink', value: { display: 'missing target' } },
			{ path: '/workbook/properties/title', value: 'Unsupported' },
		])

		expect(result.replayable).toBe(false)
		expect(result.ops).toEqual([])
		expect(result.issues.map((entry) => entry.code)).toEqual([
			'sheet_not_found',
			'invalid_ref',
			'invalid_value',
			'table_not_found',
			'invalid_path',
			'invalid_value',
			'unsupported_path',
		])
		expect(result.issues.at(-1)?.details?.supportedShapes).toEqual(SUPPORTED_PATH_MUTATION_SHAPES)
	})

	test('rejects malformed path syntax instead of repairing it silently', () => {
		const wb = AscendWorkbook.create()

		const result = wb.compilePathMutations([
			{ path: '/sheets//cells/A1/value', value: 1 },
			{ path: '/sheets/Sheet1/cells/A1/value/', value: 1 },
			{ path: 'sheets..Sheet1.cells.A1.value', value: 1 },
			{ path: 'sheets.Sheet1\\cells.A1.value', value: 1 },
			{ path: '/sheets/Sheet1~2/cells/A1/value', value: 1 },
			{ path: '/sheets/%E0%A4%A/cells/A1/value', value: 1 },
			{ path: ['sheets', '', 'cells', 'A1', 'value'], value: 1 },
		])

		expect(result.replayable).toBe(false)
		expect(result.ops).toEqual([])
		expect(result.issues.map((entry) => entry.code)).toEqual([
			'invalid_path',
			'invalid_path',
			'invalid_path',
			'invalid_path',
			'invalid_path',
			'invalid_path',
			'invalid_path',
		])
		expect(result.issues.map((entry) => entry.message)).toEqual([
			'Path segment 1 must not be empty.',
			'Path segment 5 must not be empty.',
			'Path segment 1 must not be empty.',
			'Invalid escaped character "\\c" in dot path.',
			'Invalid JSON Pointer escape in path segment "Sheet1~2".',
			'Invalid percent encoding in path segment "%E0%A4%A".',
			'Path segment 1 must not be empty.',
		])
	})
})

function pointerSegment(value: string): string {
	return encodeURIComponent(value.replace(/~/g, '~0').replace(/\//g, '~1'))
}

function dotSegment(value: string): string {
	return value.replace(/\\/g, '\\\\').replace(/\./g, '\\.')
}
