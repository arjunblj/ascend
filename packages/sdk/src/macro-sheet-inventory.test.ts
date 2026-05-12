import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../io-xlsx/test/helpers.ts'
import { AscendWorkbook } from './index.ts'

describe('macro sheet SDK inventory', () => {
	test('inspect exposes Excel 4 macro sheets separately from worksheets', async () => {
		const wb = await AscendWorkbook.open(macroSheetWorkbook())
		const info = wb.inspect()

		expect(info.sheetCount).toBe(1)
		expect(info.macroSheetCount).toBe(1)
		expect(info.macroSheets).toEqual([
			{
				name: 'Macro1',
				sheetId: '2',
				relId: 'rIdMacro',
				partPath: 'xl/macrosheets/sheet1.xml',
				state: 'hidden',
				relationshipCount: 0,
				dimensionRef: 'A1',
				cellCount: 1,
				formulaCount: 1,
			},
		])
		expect(info.activeContent).toContainEqual({
			kind: 'macroSheet',
			partPath: 'xl/macrosheets/sheet1.xml',
			contentType: 'application/vnd.ms-excel.macrosheet+xml',
			anchor: 'sheet',
			sheetName: 'Macro1',
			relType: 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
			sourceRelationshipId: 'rIdMacro',
			relationshipCount: 0,
			opaque: true,
			executionPolicy: 'blocked',
		})
		expect(
			info.compatibility.features.find((feature) => feature.feature === 'preservedMacroSheet'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/macrosheets/sheet1.xml'],
		})
	})
})

function macroSheetWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMacro" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rIdData"/>
    <sheet name="Macro1" sheetId="2" r:id="rIdMacro" state="hidden"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/macrosheets/sheet1.xml': `<?xml version="1.0"?>
<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><f>RUN("Task")</f><v>0</v></c></row></sheetData>
</xm:macrosheet>`,
	})
}
