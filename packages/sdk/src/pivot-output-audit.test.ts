import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { numberValue, stringValue } from '@ascend/schema'
import { AscendWorkbook, WorkbookDocument, WorkbookSession } from './index.ts'

describe('pivot output audits', () => {
	test('audits LibreOffice calculated pivot output against materialized cache records', async () => {
		const wb = await AscendWorkbook.open(loadLibreOfficeFixture('pivot-table/tdf126858-1.xlsx'))

		expect(wb.pivotOutputAudits()).toEqual([
			{
				pivotTable: 'СводнаяТаблица3',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Лист2',
				cacheId: 32,
				status: 'passed',
				checkedValueCount: 4,
				mismatches: [],
				warnings: [],
			},
		])
	})

	test('audits SUM and COUNT pivot outputs with calculated cache fields', async () => {
		const wb = await AscendWorkbook.open(
			loadLibreOfficeFixture('pivot-table/test_diff_aggregation.xlsx'),
		)

		expect(wb.pivotOutputAudits()).toEqual([
			{
				pivotTable: 'DataPilot1',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Pivot Table_Sheet1_1',
				cacheId: 4,
				status: 'passed',
				checkedValueCount: 6,
				mismatches: [],
				warnings: [],
			},
			{
				pivotTable: 'DataPilot2',
				partPath: 'xl/pivotTables/pivotTable2.xml',
				sheetName: 'Pivot Table_Sheet1_2',
				cacheId: 4,
				status: 'passed',
				checkedValueCount: 6,
				mismatches: [],
				warnings: [],
			},
		])
	})

	test('exposes audits through read-only document and session APIs', async () => {
		const bytes = loadLibreOfficeFixture('pivot-table/test_diff_aggregation.xlsx')
		const document = await WorkbookDocument.open(bytes)
		const session = await WorkbookSession.open(bytes)

		expect(document.pivotOutputAudits().map((audit) => audit.status)).toEqual(['passed', 'passed'])
		expect(session.pivotOutputAudits().map((audit) => audit.checkedValueCount)).toEqual([6, 6])

		session.close()
	})

	test('reports unsupported instead of overclaiming when cache records are absent', () => {
		const wb = workbookWithoutMaterializedPivot()

		expect(wb.pivotOutputAudits()).toEqual([
			expect.objectContaining({
				pivotTable: 'PivotTable1',
				status: 'unsupported',
				checkedValueCount: 0,
				mismatches: [],
				warnings: ['Pivot cache records are not fully materialized.'],
			}),
		])
	})

	test('audits simple saved pivot outputs with page filters and hidden grand totals', () => {
		const wb = workbookWithFilteredPivot()

		expect(wb.pivotOutputAudits()).toEqual([
			{
				pivotTable: 'FilteredPivot',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Sheet1',
				cacheId: 7,
				status: 'passed',
				checkedValueCount: 2,
				mismatches: [],
				warnings: [],
			},
		])
	})

	test('audits saved pivot outputs with data fields on rows and columns on axis', () => {
		const wb = workbookWithDataFieldsOnRowsPivot()

		expect(wb.pivotOutputAudits()).toEqual([
			{
				pivotTable: 'DataFieldsOnRowsPivot',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Sheet1',
				cacheId: 8,
				status: 'passed',
				checkedValueCount: 6,
				mismatches: [],
				warnings: [],
			},
		])
	})

	test('audits empty saved pivot outputs without overclaiming value checks', () => {
		const wb = workbookWithEmptyPivotOutput()

		expect(wb.pivotOutputAudits()).toEqual([
			{
				pivotTable: 'EmptyPivot',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Sheet1',
				cacheId: 9,
				status: 'passed',
				checkedValueCount: 1,
				mismatches: [],
				warnings: [],
			},
		])
	})
})

function loadLibreOfficeFixture(file: string): Uint8Array {
	return readFileSync(new URL(`../../../fixtures/xlsx/libreoffice/${file}`, import.meta.url))
}

function workbookWithoutMaterializedPivot(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const internal = wb as unknown as {
		wb: {
			pivotCaches: Array<Record<string, unknown>>
			pivotTables: Array<Record<string, unknown>>
		}
	}
	internal.wb.pivotCaches.push({
		partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
		cacheId: 1,
		fields: [
			{ index: 0, name: 'Region' },
			{ index: 1, name: 'Sales' },
		],
	})
	internal.wb.pivotTables.push({
		partPath: 'xl/pivotTables/pivotTable1.xml',
		sheetName: 'Sheet1',
		name: 'PivotTable1',
		cacheId: 1,
		locationRef: 'A1:B3',
		fields: [],
		rowFields: [{ index: 0 }],
		columnFields: [],
		pageFields: [],
		dataFields: [{ fieldIndex: 1, name: 'Sum of Sales' }],
	})
	return wb
}

function workbookWithFilteredPivot(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const internal = wb as unknown as {
		wb: {
			pivotCaches: Array<Record<string, unknown>>
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
			{
				index: 2,
				name: 'Channel',
				sharedItems: [
					{ index: 0, kind: 'string', value: 'Retail' },
					{ index: 1, kind: 'string', value: 'Online' },
				],
			},
		],
		records: {
			partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			parsedCount: 3,
			materializedCount: 3,
			materializedComplete: true,
			preview: [],
			valueKindCounts: [],
			materializedRecords: [
				{
					index: 0,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 0 },
						{ index: 1, kind: 'number', value: '999' },
						{ index: 2, kind: 'sharedItem', sharedItemIndex: 0 },
					],
				},
				{
					index: 1,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 0 },
						{ index: 1, kind: 'number', value: '100' },
						{ index: 2, kind: 'sharedItem', sharedItemIndex: 1 },
					],
				},
				{
					index: 2,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 1 },
						{ index: 1, kind: 'number', value: '50' },
						{ index: 2, kind: 'sharedItem', sharedItemIndex: 1 },
					],
				},
			],
		},
	})
	internal.wb.pivotTables.push({
		partPath: 'xl/pivotTables/pivotTable1.xml',
		sheetName: 'Sheet1',
		name: 'FilteredPivot',
		cacheId: 7,
		locationRef: 'A1:B3',
		options: { rowGrandTotals: false },
		fields: [
			{ index: 0, axis: 'axisRow' },
			{ index: 1, dataField: true },
			{ index: 2, axis: 'axisPage', items: [{ index: 0, cacheIndex: 1 }] },
		],
		rowFields: [{ index: 0 }],
		columnFields: [],
		pageFields: [{ index: 2, item: 0 }],
		dataFields: [{ fieldIndex: 1, name: 'Sum of Sales' }],
	})
	return wb
}

function workbookWithDataFieldsOnRowsPivot(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const internal = wb as unknown as {
		wb: {
			pivotCaches: Array<Record<string, unknown>>
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
	cells.set(0, 1, { value: stringValue('Month'), formula: null, styleId: 0 })
	cells.set(1, 0, { value: stringValue('Values'), formula: null, styleId: 0 })
	cells.set(1, 1, { value: stringValue('Jan'), formula: null, styleId: 0 })
	cells.set(1, 2, { value: stringValue('Feb'), formula: null, styleId: 0 })
	cells.set(1, 3, { value: stringValue('Grand Total'), formula: null, styleId: 0 })
	cells.set(2, 0, { value: stringValue('Sum of Sales'), formula: null, styleId: 0 })
	cells.set(2, 1, { value: numberValue(10), formula: null, styleId: 0 })
	cells.set(2, 2, { value: numberValue(30), formula: null, styleId: 0 })
	cells.set(2, 3, { value: numberValue(40), formula: null, styleId: 0 })
	cells.set(3, 0, { value: stringValue('Count of Sales'), formula: null, styleId: 0 })
	cells.set(3, 1, { value: numberValue(1), formula: null, styleId: 0 })
	cells.set(3, 2, { value: numberValue(2), formula: null, styleId: 0 })
	cells.set(3, 3, { value: numberValue(3), formula: null, styleId: 0 })
	internal.wb.pivotCaches.push({
		partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
		cacheId: 8,
		fields: [
			{
				index: 0,
				name: 'Month',
				sharedItems: [
					{ index: 0, kind: 'string', value: 'Jan' },
					{ index: 1, kind: 'string', value: 'Feb' },
				],
			},
			{ index: 1, name: 'Sales' },
		],
		records: {
			partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			parsedCount: 3,
			materializedCount: 3,
			materializedComplete: true,
			preview: [],
			valueKindCounts: [],
			materializedRecords: [
				{
					index: 0,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 0 },
						{ index: 1, kind: 'number', value: '10' },
					],
				},
				{
					index: 1,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 1 },
						{ index: 1, kind: 'number', value: '20' },
					],
				},
				{
					index: 2,
					values: [
						{ index: 0, kind: 'sharedItem', sharedItemIndex: 1 },
						{ index: 1, kind: 'number', value: '10' },
					],
				},
			],
		},
	})
	internal.wb.pivotTables.push({
		partPath: 'xl/pivotTables/pivotTable1.xml',
		sheetName: 'Sheet1',
		name: 'DataFieldsOnRowsPivot',
		cacheId: 8,
		locationRef: 'A1:D4',
		location: { ref: 'A1:D4', firstDataRow: 2, firstDataCol: 1 },
		options: { dataOnRows: true },
		fields: [
			{
				index: 0,
				axis: 'axisCol',
				items: [
					{ index: 0, cacheIndex: 0 },
					{ index: 1, cacheIndex: 1 },
				],
			},
			{ index: 1, dataField: true },
		],
		rowFields: [{ index: -2 }],
		columnFields: [{ index: 0 }],
		columnItems: [
			{ index: 0, fieldItems: [{ index: 0, item: 0 }] },
			{ index: 1, fieldItems: [{ index: 0, item: 1 }] },
			{ index: 2, itemType: 'grand', fieldItems: [{ index: 0 }] },
		],
		pageFields: [],
		dataFields: [
			{ fieldIndex: 1, name: 'Sum of Sales' },
			{ fieldIndex: 1, name: 'Count of Sales', subtotal: 'count' },
		],
	})
	return wb
}

function workbookWithEmptyPivotOutput(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const internal = wb as unknown as {
		wb: {
			pivotCaches: Array<Record<string, unknown>>
			pivotTables: Array<Record<string, unknown>>
		}
	}
	internal.wb.pivotCaches.push({
		partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
		cacheId: 9,
		fields: [{ index: 0, name: 'Segment' }],
		records: {
			partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			parsedCount: 0,
			materializedCount: 0,
			materializedComplete: true,
			preview: [],
			valueKindCounts: [],
			materializedRecords: [],
		},
	})
	internal.wb.pivotTables.push({
		partPath: 'xl/pivotTables/pivotTable1.xml',
		sheetName: 'Sheet1',
		name: 'EmptyPivot',
		cacheId: 9,
		locationRef: 'C3',
		location: { ref: 'C3', firstDataRow: 0, firstDataCol: 0 },
		fields: [{ index: 0, axis: 'axisPage' }],
		rowFields: [],
		columnFields: [],
		pageFields: [{ index: 0 }],
		dataFields: [],
	})
	return wb
}
