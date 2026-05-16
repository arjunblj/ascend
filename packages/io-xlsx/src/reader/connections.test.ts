import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('connection part inventory', () => {
	test('discovers query tables, workbook connections, and Power Query mashup parts', () => {
		const result = readXlsx(connectionWorkbook())
		expectOk(result)

		expect(result.value.workbook.connectionParts).toEqual([
			{
				kind: 'connection',
				partPath: 'xl/connections.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections',
				relationshipCount: 0,
				name: 'SalesConnection',
				connectionId: 1,
				connectionType: 6,
				description: 'CSV import',
				deleted: false,
				backgroundRefresh: true,
				keepAlive: true,
				refreshInterval: 15,
				refreshOnLoad: true,
				saveData: false,
				savePassword: false,
				refreshedVersion: 8,
				sourceFile: 'C:\\data\\sales.csv',
				command: 'SELECT * FROM [Sales]',
				hasConnectionString: true,
			},
			{
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
				sheetName: 'Data',
				relationshipCount: 0,
				name: 'SalesQuery',
				connectionId: 1,
				refreshOnLoad: true,
				saveData: false,
			},
			{
				kind: 'powerQueryMashup',
				partPath: 'xl/customData/item1.data',
				contentType: 'application/vnd.ms-excel.customData',
				relType: 'http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup',
				relationshipCount: 0,
			},
		])
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedConnection'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/connections.xml'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedPowerQuery'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/customData/item1.data'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedQueryTable'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/queryTables/queryTable1.xml'],
		})
	})
})

function connectionWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
  <Override PartName="/xl/customData/item1.data" ContentType="application/vnd.ms-excel.customData"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdConn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
  <Relationship Id="rIdMashup" Type="http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup" Target="customData/item1.data"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdQuery" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/>
</Relationships>`,
		'xl/connections.xml': `<?xml version="1.0"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <connection id="1" name="SalesConnection" description="CSV import" type="6" deleted="0" background="1" keepAlive="1" interval="15" refreshOnLoad="1" saveData="0" savePassword="0" refreshedVersion="8">
    <dbPr connection="Provider=Microsoft.ACE.OLEDB.12.0;" command="SELECT * FROM [Sales]"/>
    <textPr sourceFile="C:\\data\\sales.csv"/>
  </connection>
</connections>`,
		'xl/queryTables/queryTable1.xml': `<?xml version="1.0"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="SalesQuery" connectionId="1" refreshOnLoad="1" removeDataOnSave="1"/>`,
		'xl/customData/item1.data': 'mashup-bytes',
	})
}
