import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { Workbook } from '@ascend/core'
import { readXlsx } from './index.ts'
import {
	parsePivotCacheDefinitionXml,
	parsePivotCacheRecordsXml,
	parsePivotTableXml,
} from './pivots.ts'

function expectWorkbook(bytes: Uint8Array): Workbook {
	const result = readXlsx(bytes)
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
	return result.value.workbook
}

describe('pivot inventory', () => {
	test('parses bounded pivot cache record previews and kind counts', () => {
		const parsed = parsePivotCacheRecordsXml(
			`<?xml version="1.0"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <r><n v="1"/><x v="0"/><m/></r>
  <r><s v="West"/><b v="1"/><e v="#DIV/0!"/></r>
  <r><d v="2024-01-01T00:00:00"/><future v="raw"/></r>
</pivotCacheRecords>`,
			'xl/pivotCache/pivotCacheRecords1.xml',
			2,
		)
		expect(parsed).toEqual({
			partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			declaredCount: 3,
			parsedCount: 3,
			preview: [
				{
					index: 0,
					values: [
						{ index: 0, kind: 'number', value: '1' },
						{ index: 1, kind: 'sharedItem', sharedItemIndex: 0 },
						{ index: 2, kind: 'missing' },
					],
				},
				{
					index: 1,
					values: [
						{ index: 0, kind: 'string', value: 'West' },
						{ index: 1, kind: 'boolean', value: '1' },
						{ index: 2, kind: 'error', value: '#DIV/0!' },
					],
				},
			],
			valueKindCounts: [
				{ kind: 'number', count: 1 },
				{ kind: 'sharedItem', count: 1 },
				{ kind: 'missing', count: 1 },
				{ kind: 'string', count: 1 },
				{ kind: 'boolean', count: 1 },
				{ kind: 'error', count: 1 },
				{ kind: 'date', count: 1 },
				{ kind: 'unknown', count: 1 },
			],
		})
	})

	test('parses pivot cache range grouping metadata', () => {
		const parsed = parsePivotCacheDefinitionXml(
			`<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" upgradeOnRefresh="1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:D20" sheet="Data" name="SalesRange"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="OrderDate">
      <fieldGroup par="1" base="0">
        <rangePr groupBy="months" startDate="2024-01-01T00:00:00" endDate="2024-12-31T00:00:00" autoStart="0" autoEnd="1"/>
      </fieldGroup>
    </cacheField>
    <cacheField name="Amount">
      <fieldGroup base="1">
        <rangePr startNum="100" endNum="500" groupInterval="100"/>
      </fieldGroup>
    </cacheField>
  </cacheFields>
  <extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}"><x14:pivotCacheDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" pivotCacheId="123"/></ext></extLst>
</pivotCacheDefinition>`,
			'xl/pivotCache/pivotCacheDefinition1.xml',
			1,
			'rId1',
			[],
		)
		expect(parsed).toMatchObject({
			upgradeOnRefresh: true,
			extensionCacheId: 123,
			sourceType: 'worksheet',
			sourceSheet: 'Data',
			sourceRef: 'A1:D20',
			sourceName: 'SalesRange',
		})
		expect(parsed?.fields[0]?.fieldGroup).toEqual({
			parent: 1,
			base: 0,
			range: {
				groupBy: 'months',
				startDate: '2024-01-01T00:00:00',
				endDate: '2024-12-31T00:00:00',
				autoStart: false,
				autoEnd: true,
			},
		})
		expect(parsed?.fields[1]?.fieldGroup).toEqual({
			base: 1,
			range: {
				startNumber: 100,
				endNumber: 500,
				groupInterval: 100,
			},
		})
	})

	test('parses page-field selections and pivot-field item flags', () => {
		const parsed = parsePivotTableXml(
			`<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="FilteredPivot" cacheId="4" dataPosition="0" chartFormat="1">
  <pivotFields count="1">
    <pivotField axis="axisPage" numFmtId="14" multipleItemSelectionAllowed="1" showAll="0">
      <items count="3">
        <item h="1" x="2"/>
        <item s="1" sd="0" f="1" m="1" c="1" d="1" e="0" n="Manual" x="5"/>
        <item t="default"/>
      </items>
      <extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}"><x14:pivotField xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" fillDownLabels="1"/></ext></extLst>
    </pivotField>
  </pivotFields>
  <pageFields count="1"><pageField fld="0" item="5" hier="-1" name="[Date]" cap="Date"/></pageFields>
  <formats count="1">
    <format dxfId="7" action="format">
      <pivotArea type="button" axis="axisRow" field="0" fieldPosition="2" dataOnly="0" labelOnly="1" grandRow="0" grandCol="1" cacheIndex="0" outline="0" collapsedLevelsAreSubtotals="1">
        <references count="1"><reference field="0" count="1" selected="0"><x v="5"/></reference></references>
      </pivotArea>
    </format>
  </formats>
  <chartFormats count="1">
    <chartFormat chart="2" format="3" series="0">
      <pivotArea type="data" outline="0" fieldPosition="0">
        <references count="1"><reference field="5" count="1" selected="0"><x v="2"/></reference></references>
      </pivotArea>
    </chartFormat>
  </chartFormats>
  <extLst>
    <ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"><x14:pivotTableDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" fillDownLabelsDefault="1" hideValuesRow="1"/></ext>
    <ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}"><xpdl:pivotTableDefinition16 xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout" EnabledSubtotalsDefault="0" SubtotalsOnTopDefault="1"/></ext>
  </extLst>
</pivotTableDefinition>`,
			'xl/pivotTables/pivotTable1.xml',
			'PivotSheet',
		)

		expect(parsed).toEqual({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'PivotSheet',
			name: 'FilteredPivot',
			cacheId: 4,
			options: {
				dataPosition: 0,
				chartFormat: 1,
				hideValuesRow: true,
				fillDownLabelsDefault: true,
				enabledSubtotalsDefault: false,
				subtotalsOnTopDefault: true,
			},
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					numFmtId: 14,
					multipleItemSelectionAllowed: true,
					showAll: false,
					fillDownLabels: true,
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
			formats: [
				{
					index: 0,
					dxfId: 7,
					action: 'format',
					area: {
						type: 'button',
						axis: 'axisRow',
						field: 0,
						fieldPosition: 2,
						dataOnly: false,
						labelOnly: true,
						grandRow: false,
						grandCol: true,
						cacheIndex: false,
						outline: false,
						collapsedLevelsAreSubtotals: true,
						references: [
							{
								index: 0,
								field: 0,
								itemCount: 1,
								selected: false,
								items: [{ index: 0, item: 5 }],
							},
						],
					},
				},
			],
			chartFormats: [
				{
					index: 0,
					chart: 2,
					formatId: 3,
					series: false,
					area: {
						type: 'data',
						outline: false,
						fieldPosition: 0,
						references: [
							{
								index: 0,
								field: 5,
								itemCount: 1,
								selected: false,
								items: [{ index: 0, item: 2 }],
							},
						],
					},
				},
			],
		})
	})

	test('captures real Calamine page-filter hidden items', () => {
		const wb = expectWorkbook(
			readFileSync(new URL('../../../../fixtures/xlsx/calamine/pivots.xlsx', import.meta.url)),
		)
		const pivot = wb.pivotTables.find((entry) => entry.name === 'PivotTable1')
		expect(pivot?.location).toMatchObject({
			ref: 'A3:E5',
			firstHeaderRow: 0,
			firstDataRow: 1,
			firstDataCol: 1,
			rowPageCount: 1,
			colPageCount: 1,
		})
		expect(pivot?.options).toMatchObject({
			applyNumberFormats: false,
			applyWidthHeightFormats: true,
			useAutoFormatting: true,
			itemPrintTitles: true,
			multipleFieldFilters: false,
			hideValuesRow: true,
			outline: true,
			outlineData: true,
			createdVersion: 8,
			updatedVersion: 8,
			minRefreshableVersion: 3,
			indent: 0,
			dataCaption: 'Values',
		})
		expect(pivot?.style).toMatchObject({
			name: 'PivotStyleLight16',
			showRowHeaders: true,
			showColHeaders: true,
			showLastColumn: true,
		})
		expect(pivot?.pageFields).toEqual([{ index: 5, hierarchy: -1 }])
		expect(pivot?.dataFields[1]).toMatchObject({
			fieldIndex: 3,
			name: 'Average of Value',
			subtotal: 'average',
			baseField: 0,
			baseItem: 0,
		})
		expect(pivot?.rowItems).toEqual([
			{ index: 0, fieldItems: [{ index: 0, item: 1 }] },
			{ index: 1, itemType: 'grand', fieldItems: [{ index: 0 }] },
		])
		expect(pivot?.columnItems).toEqual([
			{ index: 0, fieldItems: [{ index: 0 }] },
			{ index: 1, dataFieldIndex: 1, fieldItems: [{ index: 0, item: 1 }] },
			{ index: 2, dataFieldIndex: 2, fieldItems: [{ index: 0, item: 2 }] },
			{ index: 3, dataFieldIndex: 3, fieldItems: [{ index: 0, item: 3 }] },
		])

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

		const percentPivot = wb.pivotTables[0]
		expect(percentPivot?.dataFields[0]).toMatchObject({
			fieldIndex: 2,
			name: 'NumberOfOrdersPercentageOfBearclaw',
			showDataAs: 'percent',
			baseField: 0,
			baseItem: 2,
			numFmtId: 9,
		})
		expect(percentPivot?.style).toMatchObject({
			name: 'PivotStyleLight16',
			showRowHeaders: true,
			showColHeaders: true,
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
