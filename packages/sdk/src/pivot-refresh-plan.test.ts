import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('pivot refresh plans', () => {
	test('inspect exposes stale pivot output warnings and refresh operations', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				pivotTables: Array<Record<string, unknown>>
				pivotCaches: Array<Record<string, unknown>>
			}
		}
		internal.wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Summary',
			name: 'PivotTable1',
			cacheId: 34,
			locationRef: 'A3:D20',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D100',
			invalid: true,
			saveData: false,
			fields: [],
		})

		const plan = wb.inspect().pivotRefreshPlans[0]

		expect(plan).toMatchObject({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D100',
			outputState: 'stale',
			canRefreshHeadlessly: false,
			requiresExternalRefresh: true,
		})
		expect(plan?.pivotTables).toEqual([
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Summary',
				name: 'PivotTable1',
				locationRef: 'A3:D20',
			},
		])
		expect(plan?.warnings.join('\n')).toContain('does not recalculate pivot output cells')
		expect(plan?.recommendedOps).toEqual([
			{
				op: 'setPivotCache',
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 34,
				refreshOnLoad: true,
				invalid: true,
				saveData: false,
			},
		])
		expect(wb.pivotRefreshPlans()).toEqual(wb.inspect().pivotRefreshPlans)
	})

	test('cached pivot caches do not require refresh actions', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				pivotCaches: Array<Record<string, unknown>>
			}
		}
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 7,
			recordCount: 10,
			records: {
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 10,
				parsedCount: 10,
				preview: [],
				materializedCount: 10,
				materializedComplete: true,
				valueKindCounts: [
					{ kind: 'sharedItem', count: 12 },
					{ kind: 'number', count: 8 },
				],
			},
			fields: [],
		})

		expect(wb.pivotRefreshPlans()[0]).toMatchObject({
			outputState: 'cached',
			cacheRecords: {
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 10,
				parsedCount: 10,
				materializedCount: 10,
				materializedComplete: true,
				valueKindCounts: [
					{ kind: 'sharedItem', count: 12 },
					{ kind: 'number', count: 8 },
				],
			},
			requiresExternalRefresh: false,
			warnings: [],
			recommendedOps: [],
		})
	})

	test('cached pivot records still require refresh when output cells are not saved', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				pivotCaches: Array<Record<string, unknown>>
				pivotTables: Array<Record<string, unknown>>
			}
		}
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 8,
			recordCount: 2,
			sourceSheet: 'Raw',
			sourceRef: 'A1:B3',
			fields: [
				{ index: 0, name: 'Region' },
				{ index: 1, name: 'Sales' },
			],
		})
		internal.wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'MissingOutputPivot',
			cacheId: 8,
			locationRef: 'A1:B3',
			fields: [
				{ index: 0, name: 'Region', axis: 'axisRow' },
				{ index: 1, name: 'Sales', dataField: true },
			],
			rowFields: [{ index: 0 }],
			columnFields: [],
			pageFields: [],
			dataFields: [{ fieldIndex: 1, name: 'Sum of Sales', subtotal: 'sum' }],
		})

		const plan = wb.pivotRefreshPlans()[0]

		expect(plan).toMatchObject({
			outputState: 'not-saved',
			requiresExternalRefresh: true,
			recommendedOps: [
				{
					op: 'setPivotCache',
					partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
					cacheId: 8,
					refreshOnLoad: true,
					invalid: true,
					saveData: false,
				},
			],
		})
		expect(plan?.warnings.join('\n')).toContain('PivotTable output range(s) have no saved cells')
		expect(wb.refreshMetadata().notSavedCount).toBe(1)
	})
})
