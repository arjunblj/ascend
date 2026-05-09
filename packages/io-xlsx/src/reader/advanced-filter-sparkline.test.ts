import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('advanced filter and sparkline inventory', () => {
	test('parses custom sheet view filters and sparkline extension groups', () => {
		const result = readXlsx(advancedFilterSparklineWorkbook())
		expectOk(result)
		const sheet = result.value.workbook.sheets[0]

		expect(sheet?.advancedFilters).toEqual([
			{
				viewName: 'WestOnly',
				guid: '{11111111-1111-1111-1111-111111111111}',
				ref: 'A1:C20',
				filterColumnCount: 1,
				sortConditionCount: 1,
				autoFilter: {
					ref: 'A1:C20',
					columns: [{ colId: 0, kind: 'filters', values: ['West'] }],
					sortState: {
						ref: 'A2:C20',
						conditions: [{ ref: 'C2:C20', descending: true }],
					},
				},
			},
		])
		expect(sheet?.sparklineGroups).toEqual([
			{
				groupIndex: 0,
				type: 'line',
				displayEmptyCellsAs: 'gap',
				markers: true,
				highPoint: true,
				displayXAxis: true,
				colorSeries: 'FF4472C4',
				range: 'Data!B2:B4',
				locationRange: 'D2:D4',
				count: 1,
			},
		])
		expect(sheet?.preservedExtLst).toContain('sparklineGroups')
	})
})

export function advancedFilterSparklineWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData/>
  <customSheetViews>
    <customSheetView name="WestOnly" guid="{11111111-1111-1111-1111-111111111111}">
      <autoFilter ref="A1:C20">
        <filterColumn colId="0"><filters><filter val="West"/></filters></filterColumn>
        <sortState ref="A2:C20"><sortCondition ref="C2:C20" descending="1"/></sortState>
      </autoFilter>
    </customSheetView>
  </customSheetViews>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
      <x14:sparklineGroups>
        <x14:sparklineGroup type="line" displayEmptyCellsAs="gap" markers="1" high="1" displayXAxis="1">
          <x14:colorSeries rgb="FF4472C4"/>
          <x14:sparklines>
            <x14:sparkline><xm:f>Data!B2:B4</xm:f><xm:sqref>D2:D4</xm:sqref></x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>`,
	})
}
