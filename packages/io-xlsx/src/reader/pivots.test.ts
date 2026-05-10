import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { Workbook } from '@ascend/core'
import { readXlsx } from './index.ts'
import { parsePivotTableXml } from './pivots.ts'

function expectWorkbook(bytes: Uint8Array): Workbook {
	const result = readXlsx(bytes)
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
	return result.value.workbook
}

describe('pivot inventory', () => {
	test('parses page-field selections and pivot-field item flags', () => {
		const parsed = parsePivotTableXml(
			`<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="FilteredPivot" cacheId="4">
  <pivotFields count="1">
    <pivotField axis="axisPage" numFmtId="14" multipleItemSelectionAllowed="1" showAll="0">
      <items count="3">
        <item h="1" x="2"/>
        <item s="1" sd="0" f="1" m="1" c="1" d="1" e="0" n="Manual" x="5"/>
        <item t="default"/>
      </items>
    </pivotField>
  </pivotFields>
  <pageFields count="1"><pageField fld="0" item="5" hier="-1" name="[Date]" cap="Date"/></pageFields>
</pivotTableDefinition>`,
			'xl/pivotTables/pivotTable1.xml',
			'PivotSheet',
		)

		expect(parsed).toEqual({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'PivotSheet',
			name: 'FilteredPivot',
			cacheId: 4,
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					numFmtId: 14,
					multipleItemSelectionAllowed: true,
					showAll: false,
					items: [
						{ index: 0, cacheIndex: 2, hidden: true },
						{
							index: 1,
							cacheIndex: 5,
							caption: 'Manual',
							manualFilter: true,
							showDetails: false,
							calculated: true,
							missing: true,
							childItems: true,
							expanded: true,
							drillAcrossAttributes: false,
						},
						{ index: 2, itemType: 'default' },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0, item: 5, hierarchy: -1, name: '[Date]', caption: 'Date' }],
			dataFields: [],
		})
	})

	test('captures real Calamine page-filter hidden items', () => {
		const wb = expectWorkbook(
			readFileSync(new URL('../../../../fixtures/xlsx/calamine/pivots.xlsx', import.meta.url)),
		)
		const pivot = wb.pivotTables.find((entry) => entry.name === 'PivotTable1')
		expect(pivot?.pageFields).toEqual([{ index: 5, hierarchy: -1 }])

		const dateField = pivot?.fields[5]
		expect(dateField).toMatchObject({
			index: 5,
			axis: 'axisPage',
			numFmtId: 14,
			multipleItemSelectionAllowed: true,
			showAll: false,
		})
		expect(dateField?.items?.slice(0, 3)).toEqual([
			{ index: 0, cacheIndex: 7 },
			{ index: 1, cacheIndex: 5 },
			{ index: 2, cacheIndex: 1, hidden: true },
		])
		expect(dateField?.items?.at(-1)).toEqual({ index: 9, itemType: 'default' })

		const cache = wb.pivotCaches.find((entry) => entry.cacheId === 65)
		expect(cache?.fields[5]?.sharedItems?.[7]).toEqual({
			index: 7,
			kind: 'date',
			value: '1999-01-01T00:00:00',
		})
		expect(cache?.fields[10]).toMatchObject({
			index: 10,
			name: 'Value x 2',
			databaseField: false,
			formula: "'Value / Size'* 2",
		})
	})

	test('captures real ClosedXML page-field selected item ids and prefixed pivot caches', () => {
		const wb = expectWorkbook(
			readFileSync(
				new URL(
					'../../../../fixtures/xlsx/closedxml/PivotTables_PivotTables.xlsx',
					import.meta.url,
				),
			),
		)
		expect(wb.pivotCaches).toHaveLength(1)
		expect(wb.pivotCaches[0]?.fields.slice(0, 3).map((field) => field.name)).toEqual([
			'Name',
			'Code',
			'NumberOfOrders',
		])
		expect(wb.pivotCaches[0]?.fields[3]?.sharedItems?.[13]).toEqual({
			index: 13,
			kind: 'number',
			value: '5.19',
		})
		expect(wb.pivotCaches[0]?.fields[5]?.sharedItems?.[13]).toEqual({
			index: 13,
			kind: 'date',
			value: '2017-05-03T00:00:00',
		})

		const pivot = wb.pivotTables.find((entry) => entry.name === 'pvtFilter')
		expect(pivot?.pageFields).toEqual([
			{ index: 0 },
			{ index: 3, item: 13 },
			{ index: 5, item: 13 },
		])

		const nameField = pivot?.fields[0]
		expect(nameField).toMatchObject({
			index: 0,
			axis: 'axisPage',
			name: 'Name',
			multipleItemSelectionAllowed: true,
			showAll: false,
		})
		expect(nameField?.items).toEqual([
			{ index: 0, cacheIndex: 0, hidden: true },
			{ index: 1, cacheIndex: 1 },
			{ index: 2, cacheIndex: 2, hidden: true },
			{ index: 3, cacheIndex: 3, hidden: true },
			{ index: 4, cacheIndex: 4 },
			{ index: 5, itemType: 'default' },
		])
	})
})
