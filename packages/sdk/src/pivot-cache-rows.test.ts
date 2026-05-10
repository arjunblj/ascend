import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { AscendWorkbook, WorkbookSession } from './index.ts'

describe('pivot cache materialized rows', () => {
	test('decodes saved cache rows by field name for agent audit workflows', () => {
		const wb = workbookWithPivotCacheRows()

		expect(wb.pivotCacheRows()).toEqual([
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 7,
				rowIndex: 0,
				values: [
					{
						fieldIndex: 0,
						fieldName: 'Region',
						rawKind: 'sharedItem',
						kind: 'string',
						value: 'West',
						sharedItemIndex: 0,
						sharedItemKind: 'string',
					},
					{
						fieldIndex: 1,
						fieldName: 'Sales',
						rawKind: 'number',
						kind: 'number',
						value: '125',
					},
				],
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 7,
				rowIndex: 1,
				values: [
					{
						fieldIndex: 0,
						fieldName: 'Region',
						rawKind: 'sharedItem',
						kind: 'string',
						value: 'East',
						sharedItemIndex: 1,
						sharedItemKind: 'string',
					},
					{
						fieldIndex: 1,
						fieldName: 'Sales',
						rawKind: 'number',
						kind: 'number',
						value: '250',
					},
				],
			},
		])
		expect(wb.pivotCacheRows({ cacheId: 7, limit: 1 })).toEqual([wb.pivotCacheRows()[0]])
		expect(wb.pivotCacheRows({ partPath: 'missing' })).toEqual([])
	})

	test('sessions expose decoded cache rows through the read-only document surface', async () => {
		const bytes = readFileSync(
			new URL(
				'../../../fixtures/xlsx/libreoffice/pivot-table/test_diff_aggregation.xlsx',
				import.meta.url,
			),
		)
		const session = await WorkbookSession.open(bytes)

		expect(session.pivotCacheRows({ cacheId: 4, limit: 2 })).toMatchObject({
			length: 2,
		})
		expect(session.pivotCacheRows({ cacheId: 4 })[0]?.values).toEqual([
			{
				fieldIndex: 0,
				fieldName: 'Year',
				rawKind: 'sharedItem',
				kind: 'number',
				value: '2010',
				sharedItemIndex: 0,
				sharedItemKind: 'number',
			},
			{ fieldIndex: 1, fieldName: 'Spend', rawKind: 'number', kind: 'number', value: '25' },
		])

		session.close()
	})
})

function workbookWithPivotCacheRows(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const internal = wb as unknown as {
		wb: {
			pivotCaches: Array<Record<string, unknown>>
		}
	}
	internal.wb.pivotCaches.push({
		partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
		cacheId: 7,
		fields: [
			{
				index: 0,
				name: 'Region',
				sharedItems: [
					{ index: 0, kind: 'string', value: 'West' },
					{ index: 1, kind: 'string', value: 'East' },
				],
			},
			{ index: 1, name: 'Sales' },
		],
		records: {
			partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			parsedCount: 2,
			materializedCount: 2,
			materializedComplete: true,
			preview: [],
			materializedRecords: [
				{
					index: 0,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 0 },
						{ index: 1, kind: 'number', value: '125' },
					],
				},
				{
					index: 1,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 1 },
						{ index: 1, kind: 'number', value: '250' },
					],
				},
			],
			valueKindCounts: [
				{ kind: 'sharedItem', count: 2 },
				{ kind: 'number', count: 2 },
			],
		},
	})
	return wb
}
