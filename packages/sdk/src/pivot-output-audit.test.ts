import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
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
