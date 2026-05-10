import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../io-xlsx/test/helpers.ts'
import { AscendWorkbook } from './index.ts'

describe('active content SDK inventory', () => {
	test('inspect exposes ActiveX and form-control metadata for agents', async () => {
		const wb = await AscendWorkbook.open(controlWorkbook())
		const info = wb.inspect()

		expect(info.activeContentCount).toBe(2)
		expect(info.activeContent).toContainEqual({
			kind: 'activeX',
			partPath: 'xl/activeX/activeX1.xml',
			contentType: 'application/vnd.ms-office.activeX+xml',
			anchor: 'sheet',
			sheetName: 'Data',
			relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControl',
			sourceRelationshipId: 'rIdActiveX',
			relationshipCount: 1,
			activeX: {
				classId: '{8BD21D40-EC42-11CE-9E0D-00AA006002F3}',
				persistence: 'persistStreamInit',
				relationshipId: 'rId1',
				binaryRelationshipId: 'rId1',
				binaryTarget: 'activeX1.bin',
			},
		})
		expect(info.activeContent).toContainEqual({
			kind: 'formControl',
			partPath: 'xl/ctrlProps/ctrlProp1.xml',
			contentType: 'application/vnd.ms-excel.controlproperties+xml',
			anchor: 'sheet',
			sheetName: 'Data',
			relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
			sourceRelationshipId: 'rIdCtrl',
			relationshipCount: 0,
			formControl: {
				macro: 'Module1.Run',
				linkedCell: '$A$1',
				listFillRange: '$A$2:$A$4',
			},
		})
	})
})

function controlWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdActiveX" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rIdCtrl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
		'xl/activeX/activeX1.xml': `<?xml version="1.0"?><ax:ocx ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
		'xl/activeX/_rels/activeX1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
</Relationships>`,
		'xl/ctrlProps/ctrlProp1.xml': `<?xml version="1.0"?><formControlPr macro="Module1.Run" fmlaLink="$A$1" fmlaRange="$A$2:$A$4"/>`,
	})
}
