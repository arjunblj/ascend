import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { numberValue, stringValue } from '@ascend/schema'
import { AscendWorkbook } from './index.ts'

const MS_EXCEL_PIVOT_FIXTURE = new URL(
	'../../../research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx',
	import.meta.url,
)

describe('GETPIVOTDATA metadata queries', () => {
	test('resolves matching pivot metadata and reports output limitations', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				pivotTables: Array<Record<string, unknown>>
			}
		}
		internal.wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Summary',
			name: 'PivotTable1',
			cacheId: 34,
			locationRef: 'A3:D20',
			fields: [
				{ index: 0, name: 'Region' },
				{ index: 1, name: 'Product' },
			],
			rowFields: [{ index: 0, name: 'Region' }],
			columnFields: [{ index: 1, name: 'Product' }],
			pageFields: [],
			dataFields: [{ fieldIndex: 2, name: 'Sum of Sales', subtotal: 'sum' }],
		})

		const result = wb.getPivotData({
			pivotTable: 'PivotTable1',
			dataField: 'Sum of Sales',
			filters: [
				{ field: 'Region', item: 'West' },
				{ field: 'Missing', item: 'Nope' },
			],
		})

		expect(result.canResolveOutput).toBe(false)
		expect(result.matches).toHaveLength(1)
		expect(result.matches[0]?.pivotTable.name).toBe('PivotTable1')
		expect(result.matches[0]?.dataField.name).toBe('Sum of Sales')
		expect(result.matches[0]?.matchedFilters).toEqual([{ field: 'Region', item: 'West' }])
		expect(result.matches[0]?.unmatchedFilters).toEqual([{ field: 'Missing', item: 'Nope' }])
		expect(result.warnings.join('\n')).toContain('no matching saved pivot output')

		wb.setFormula('Sheet1!A1', '=GETPIVOTDATA("Sum of Sales",Summary!A3)')
		expect(wb.inspect().capabilityWarnings).toContainEqual(
			expect.objectContaining({
				capabilityId: 'analytics.getpivotdata',
				evidence: ['GETPIVOTDATA formulas=1'],
			}),
		)
	})

	test('returns a machine-readable miss when no pivot data field matches', () => {
		const wb = AscendWorkbook.create()

		const result = wb.getPivotData({ dataField: 'Sum of Sales' })

		expect(result.matches).toEqual([])
		expect(result.warnings.join('\n')).toContain('No matching pivot table/data field')
	})

	test('resolves saved visible pivot output values from SDK queries', () => {
		const wb = workbookWithSavedPivotOutput()

		const filtered = wb.getPivotData({
			pivotTable: 'PivotTable1',
			dataField: 'Sum of Sales',
			filters: [{ field: 'Region', item: 'West' }],
		})

		expect(filtered.canResolveOutput).toBe(true)
		expect(filtered.matches[0]?.output).toEqual({
			sheetName: 'Sheet1',
			ref: 'B2',
			value: numberValue(100),
		})
		expect(filtered.warnings.join('\n')).toContain('not recalculated')

		const grandTotal = wb.getPivotData({
			pivotTable: 'PivotTable1',
			dataField: 'Sum of Sales',
		})

		expect(grandTotal.canResolveOutput).toBe(true)
		expect(grandTotal.matches[0]?.output).toEqual({
			sheetName: 'Sheet1',
			ref: 'B4',
			value: numberValue(150),
		})

		const rootName = wb.getPivotData({
			pivotTable: 'PivotTable1',
			dataField: 'Sales',
			filters: [{ field: 'Region', item: 'West' }],
		})

		expect(rootName.canResolveOutput).toBe(true)
		expect(rootName.matches[0]?.dataField.name).toBe('Sum of Sales')
		expect(rootName.matches[0]?.output).toEqual({
			sheetName: 'Sheet1',
			ref: 'B2',
			value: numberValue(100),
		})
	})

	test('resolves real Excel saved pivot grand totals', async () => {
		const wb = await AscendWorkbook.open(readFileSync(MS_EXCEL_PIVOT_FIXTURE), {
			pivotCacheRecordMaterializeLimit: 'all',
		})

		const result = wb.getPivotData({
			pivotTable: 'PivotTable1',
			dataField: 'Count of outcome',
		})

		expect(result.canResolveOutput).toBe(true)
		expect(result.matches[0]?.output).toEqual({
			sheetName: 'Pivot 1',
			ref: 'F14',
			value: numberValue(4114),
		})
	})
})

function workbookWithSavedPivotOutput(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const internal = wb as unknown as {
		wb: {
			pivotTables: Array<Record<string, unknown>>
			sheets: Array<{
				cells: {
					set(row: number, col: number, cell: { value: unknown; formula: null; styleId: 0 }): void
				}
			}>
		}
	}
	const cells = internal.wb.sheets[0]?.cells
	if (!cells) throw new Error('Expected default sheet')
	cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: 0 })
	cells.set(0, 1, { value: stringValue('Sum of Sales'), formula: null, styleId: 0 })
	cells.set(1, 0, { value: stringValue('West'), formula: null, styleId: 0 })
	cells.set(1, 1, { value: numberValue(100), formula: null, styleId: 0 })
	cells.set(2, 0, { value: stringValue('East'), formula: null, styleId: 0 })
	cells.set(2, 1, { value: numberValue(50), formula: null, styleId: 0 })
	cells.set(3, 0, { value: stringValue('Grand Total'), formula: null, styleId: 0 })
	cells.set(3, 1, { value: numberValue(150), formula: null, styleId: 0 })
	internal.wb.pivotTables.push({
		partPath: 'xl/pivotTables/pivotTable1.xml',
		sheetName: 'Sheet1',
		name: 'PivotTable1',
		cacheId: 1,
		locationRef: 'A1:B4',
		fields: [{ index: 0, name: 'Region' }],
		rowFields: [{ index: 0 }],
		columnFields: [],
		pageFields: [],
		dataFields: [{ fieldIndex: 1, name: 'Sum of Sales', subtotal: 'sum' }],
	})
	return wb
}
