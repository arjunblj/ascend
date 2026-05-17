import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('timeline inventory', () => {
	test('discovers timeline and timeline cache parts', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<x15:timelineCacheDefinition xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <x15:data><x15:tabular pivotCacheId="3"/></x15:data>
  <x15:pivotTables><x15:pivotTable name="PivotTable1"/></x15:pivotTables>
  <x15:state filterId="7" filterPivotName="PivotTable1" filterType="dateRange" filterTabId="2" lastRefreshVersion="6" minimalRefreshVersion="4" pivotCacheId="3" singleRangeFilterState="1">
    <x15:selection startDate="2024-01-01T00:00:00" endDate="2024-03-31T00:00:00"/>
    <x15:bounds startDate="2023-01-01T00:00:00" endDate="2024-12-31T00:00:00"/>
  </x15:state>
</x15:timelineCacheDefinition>`,
			'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<x15:timelines xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <x15:timeline name="Order_Date" cache="Timeline_Order_Date" caption="Order Date"/>
</x15:timelines>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.timelineCaches).toEqual([
			{
				partPath: 'xl/timelineCaches/timelineCache1.xml',
				name: 'Timeline_Order_Date',
				sourceName: 'Order Date',
				pivotCacheId: 3,
				pivotTableNames: ['PivotTable1'],
				state: {
					filterId: 7,
					filterPivotName: 'PivotTable1',
					filterType: 'dateRange',
					filterTabId: 2,
					lastRefreshVersion: 6,
					minimalRefreshVersion: 4,
					pivotCacheId: 3,
					singleRangeFilterState: true,
					selection: {
						startDate: '2024-01-01T00:00:00',
						endDate: '2024-03-31T00:00:00',
					},
					bounds: {
						startDate: '2023-01-01T00:00:00',
						endDate: '2024-12-31T00:00:00',
					},
				},
			},
		])
		expect(result.value.workbook.timelines).toEqual([
			{
				partPath: 'xl/timelines/timeline1.xml',
				name: 'Order_Date',
				cacheName: 'Timeline_Order_Date',
				caption: 'Order Date',
			},
		])
		expect(
			result.value.report.features.find((entry) => entry.feature === 'preservedTimeline'),
		).toBeDefined()
	})
})
