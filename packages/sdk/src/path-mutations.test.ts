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
		const columnName = "Gross.Profit / Δ~'s"
		const definedName = "Revenue/North ~ Café's Δ"
		const scopedName = 'Local.Rate_Δ'
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
			{ path: ['sheets', sheetName, 'cells', 'A3', 'value'], value: 'array' },
			{
				path: `/tables/${pointerSegment(tableName)}/columns/${pointerSegment(columnName)}/formula`,
				value: "SUM([Gross.Profit / Δ~'s])",
			},
			{
				path: `tables.${dotSegment(tableName)}.columns.${dotSegment(columnName)}.totalsRowLabel`,
				value: 'Total',
			},
			{ path: `/names/${pointerSegment(definedName)}/ref`, value: `'${sheetName}'!$B$2` },
			{ path: `names.${dotSegment(definedName)}.ref`, value: `'${sheetName}'!$B$2` },
			{
				path: `/sheets/${pointerSegment(sheetName)}/names/${pointerSegment(scopedName)}/ref`,
				value: `'${sheetName}'!$B$2`,
			},
			{
				path: `sheets.${dotSegment(sheetName)}.names.${dotSegment(scopedName)}.ref`,
				value: `'${sheetName}'!$B$2`,
			},
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
					{ ref: 'A3', value: 'array' },
				],
			},
			{
				op: 'setTableColumn',
				table: tableName,
				column: columnName,
				formula: "SUM([Gross.Profit / Δ~'s])",
			},
			{ op: 'setTableColumn', table: tableName, column: columnName, totalsRowLabel: 'Total' },
			{ op: 'setDefinedName', name: definedName, ref: `'${sheetName}'!$B$2` },
			{ op: 'setDefinedName', name: definedName, ref: `'${sheetName}'!$B$2` },
			{ op: 'setDefinedName', name: scopedName, scope: sheetName, ref: `'${sheetName}'!$B$2` },
			{ op: 'setDefinedName', name: scopedName, scope: sheetName, ref: `'${sheetName}'!$B$2` },
		])
	})

	test('defers sheet and table renames so replayable batches still apply', () => {
		const sheetName = 'Q1.Forecast'
		const tableName = 'Sales.Δ'
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: sheetName },
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Revenue' },
					{ ref: 'A2', value: 'North' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: sheetName, ref: 'A1:B2', name: tableName, hasHeaders: true },
		])

		const result = wb.compilePathMutations([
			{ path: `/sheets/${pointerSegment(sheetName)}/name`, value: 'Summary' },
			{ path: `/sheets/${pointerSegment(sheetName)}/cells/C1/value`, value: 'still old address' },
			{ path: `/tables/${pointerSegment(tableName)}/name`, value: 'SalesData' },
			{
				path: `/tables/${pointerSegment(tableName)}/columns/Revenue/formula`,
				value: 'SUM([Revenue])',
			},
		])

		expect(result.replayable).toBe(true)
		expect(result.issues).toEqual([])
		expect(result.ops).toEqual([
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [{ ref: 'C1', value: 'still old address' }],
			},
			{
				op: 'setTableColumn',
				table: tableName,
				column: 'Revenue',
				formula: 'SUM([Revenue])',
			},
			{ op: 'renameSheet', sheet: sheetName, newName: 'Summary' },
			{ op: 'renameTable', table: tableName, newName: 'SalesData' },
		])

		const applied = wb.batch(result.ops)
		expect(applied.errors).toEqual([])
		expect(wb.sheets).toContain('Summary')
		expect(wb.sheet('Summary')?.cell('C1')?.value).toEqual({
			kind: 'string',
			value: 'still old address',
		})
		expect(wb.table('SalesData')?.columns).toContain('Revenue')
	})

	test('canonicalizes case-insensitive table column path selectors', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Revenue' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])

		const result = wb.compilePathMutations([
			{ path: '/tables/Sales/columns/revenue/formula', value: 'SUM([Revenue])' },
			{ path: '/tables/Sales/columns/region/name', value: 'Market' },
		])

		expect(result.issues).toEqual([])
		expect(result.replayable).toBe(true)
		expect(result.ops).toEqual([
			{ op: 'setTableColumn', table: 'Sales', column: 'Revenue', formula: 'SUM([Revenue])' },
			{ op: 'setTableColumn', table: 'Sales', column: 'Region', newName: 'Market' },
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
			{ path: '/sheets/Sheet1/name', value: 'Bad/Name' },
			{ path: '/tables/Missing/rows/append', value: [[1]] },
			{ path: '/tables/Sales/columns/Missing/formula', value: '1+1' },
			{ path: '/tables/Sales/name', value: 'Bad Name' },
			{ path: '/sheets/Sheet1/cells/A1/hyperlink', value: { display: 'missing target' } },
			{ path: '/workbook/properties/title', value: 'Unsupported' },
		])

		expect(result.replayable).toBe(false)
		expect(result.ops).toEqual([])
		expect(result.issues.map((entry) => entry.code)).toEqual([
			'sheet_not_found',
			'invalid_ref',
			'invalid_value',
			'invalid_value',
			'table_not_found',
			'invalid_path',
			'invalid_value',
			'invalid_value',
			'unsupported_path',
		])
		expect(result.issues[3]?.message).toContain('invalid characters')
		expect(result.issues[6]?.message).toContain('invalid characters')
		expect(result.issues.at(-1)?.details?.supportedShapes).toEqual(SUPPORTED_PATH_MUTATION_SHAPES)
	})

	test('does not emit partial ops when any mutation makes the batch non-replayable', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])

		const result = wb.compilePathMutations([
			{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' },
			{ path: '/sheets/Missing/cells/A1/value', value: 1 },
		])

		expect(result.replayable).toBe(false)
		expect(result.ops).toEqual([])
		expect(result.issues).toEqual([
			expect.objectContaining({
				code: 'sheet_not_found',
				path: '/sheets/Missing/cells/A1/value',
			}),
		])
		const applied = wb.batch(result.ops)
		expect(applied.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'old' })
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

	test('reports runtime-invalid mutation path shapes without throwing', () => {
		const wb = AscendWorkbook.create()

		const result = wb.compilePathMutations([
			{ path: 123, value: 1 },
			{ path: ['sheets', 1, 'cells', 'A1', 'value'], value: 1 },
		] as never)

		expect(result.replayable).toBe(false)
		expect(result.ops).toEqual([])
		expect(result.issues).toEqual([
			expect.objectContaining({
				code: 'invalid_path',
				message: 'Path must be a string or string array.',
			}),
			expect.objectContaining({
				code: 'invalid_path',
				message: 'Path must be a string or string array.',
			}),
		])
	})
})

function pointerSegment(value: string): string {
	return encodeURIComponent(value.replace(/~/g, '~0').replace(/\//g, '~1'))
}

function dotSegment(value: string): string {
	return value.replace(/\\/g, '\\\\').replace(/\./g, '\\.')
}
