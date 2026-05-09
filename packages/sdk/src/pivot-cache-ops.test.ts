import { describe, expect, test } from 'bun:test'
import { extractZip } from '../../io-xlsx/src/reader/zip.ts'
import { createZip, encode } from '../../io-xlsx/src/writer/zip.ts'
import { AscendWorkbook } from './index.ts'

describe('pivot cache operations', () => {
	test('setPivotCache writes source and refresh metadata back to XLSX', async () => {
		const wb = await AscendWorkbook.open(pivotWorkbook())
		const applied = wb.apply([
			{
				op: 'setPivotCache',
				pivotTable: 'PivotTable1',
				sourceSheet: 'RawData',
				sourceRef: 'A1:E200',
				refreshOnLoad: true,
				invalid: true,
				saveData: false,
			},
		])

		expect(applied.errors).toEqual([])
		expect(applied.sheetsModified).toEqual(['PivotSheet'])
		const out = wb.toBytes()
		const zip = extractZip(out)
		const pivotXml = zip?.readText('xl/pivotCache/pivotCacheDefinition1.xml') ?? ''
		expect(pivotXml).toContain('refreshOnLoad="1"')
		expect(pivotXml).toContain('invalid="1"')
		expect(pivotXml).toContain('saveData="0"')
		expect(pivotXml).toContain('<worksheetSource ref="A1:E200" sheet="RawData"/>')

		const reopened = await AscendWorkbook.open(out)
		expect(reopened.pivotCaches()[0]).toMatchObject({
			sourceSheet: 'RawData',
			sourceRef: 'A1:E200',
			refreshOnLoad: true,
			invalid: true,
			saveData: false,
		})
	})
})

function pivotWorkbook(): Uint8Array {
	const entries = new Map<string, Uint8Array>()
	for (const [path, content] of Object.entries({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches><pivotCache cacheId="34" r:id="rIdPivotCache"/></pivotCaches>
  <sheets><sheet name="PivotSheet" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
		'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <location ref="A1"/>
</pivotTableDefinition>`,
		'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="25">
  <cacheSource type="worksheet"><worksheetSource ref="A1:D100" sheet="Raw"/></cacheSource>
</pivotCacheDefinition>`,
	})) {
		entries.set(path, encode(content))
	}
	return createZip(entries)
}
