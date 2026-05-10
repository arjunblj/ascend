import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../io-xlsx/test/helpers.ts'
import { AscendWorkbook } from './index.ts'

describe('timeline SDK inventory and edits', () => {
	test('inspect exposes timeline state and setTimelineRange preserves cache XML', async () => {
		const wb = await AscendWorkbook.open(timelineWorkbook())
		const before = wb.inspect()

		expect(before.timelineCaches[0]?.state).toMatchObject({
			filterType: 'dateRange',
			singleRangeFilterState: true,
			selection: {
				startDate: '2024-01-01T00:00:00',
				endDate: '2024-03-31T00:00:00',
			},
			bounds: {
				startDate: '2023-01-01T00:00:00',
				endDate: '2024-12-31T00:00:00',
			},
		})

		const applied = wb.apply([
			{
				op: 'setTimelineRange',
				timelineCache: 'Timeline_Order_Date',
				startDate: '2024-04-01T00:00:00',
				endDate: '2024-06-30T00:00:00',
			},
		])
		expect(applied.errors).toHaveLength(0)
		expect(applied.warnings[0]?.message).toContain('Timeline range changed')

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.inspect().timelineCaches[0]?.state?.selection).toEqual({
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-06-30T00:00:00',
		})
	})
})

function timelineWorkbook(): Uint8Array {
	return makeXlsx({
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
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdTimelineCache" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="timelineCaches/timelineCache1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables>
  <state filterType="dateRange" singleRangeFilterState="1">
    <selection startDate="2024-01-01T00:00:00" endDate="2024-03-31T00:00:00"/>
    <bounds startDate="2023-01-01T00:00:00" endDate="2024-12-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`,
		'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="Order_Date" cache="Timeline_Order_Date" caption="Order Date"/>
</timelines>`,
	})
}
