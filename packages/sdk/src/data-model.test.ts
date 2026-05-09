import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../io-xlsx/test/helpers.ts'
import { AscendWorkbook } from './index.ts'

describe('data model SDK inventory', () => {
	test('inspect exposes data model parts and SDK accessor returns copies', async () => {
		const wb = await AscendWorkbook.open(dataModelWorkbook())
		const info = wb.inspect()

		expect(info.dataModelPartCount).toBe(3)
		expect(wb.dataModelParts()).toEqual(info.dataModelParts)
		expect(info.dataModelParts.map((part) => part.kind)).toEqual([
			'modelData',
			'modelTable',
			'modelRelationship',
		])
		expect(info.dataModelParts[0]).toMatchObject({
			partPath: 'xl/model/item.data',
			relType: 'http://schemas.microsoft.com/office/2011/relationships/model',
			relationshipCount: 1,
		})
	})
})

function dataModelWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="data" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/model/item.data" ContentType="application/vnd.ms-excel.model+data"/>
  <Override PartName="/xl/model/tables/table1.xml" ContentType="application/vnd.ms-excel.modelTable+xml"/>
  <Override PartName="/xl/model/relationships/relationship1.xml" ContentType="application/vnd.ms-excel.modelRelationship+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdModel" Type="http://schemas.microsoft.com/office/2011/relationships/model" Target="model/item.data"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/model/item.data': 'model-bytes',
		'xl/model/_rels/item.data.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable" Type="http://schemas.microsoft.com/office/2011/relationships/modelTable" Target="tables/table1.xml"/>
</Relationships>`,
		'xl/model/tables/table1.xml': `<?xml version="1.0"?><modelTable name="Sales"/>`,
		'xl/model/relationships/relationship1.xml': `<?xml version="1.0"?><modelRelationship fromTable="Sales" toTable="Calendar"/>`,
	})
}
