import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

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
		expect(result.warnings.join('\n')).toContain('not recalculated headlessly')
	})

	test('returns a machine-readable miss when no pivot data field matches', () => {
		const wb = AscendWorkbook.create()

		const result = wb.getPivotData({ dataField: 'Sum of Sales' })

		expect(result.matches).toEqual([])
		expect(result.warnings.join('\n')).toContain('No matching pivot table/data field')
	})
})
