import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('active content inventory', () => {
	test('discovers macros, ActiveX controls, and form control property parts', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
			'xl/vbaProject.bin': 'macro-bytes',
			'xl/activeX/activeX1.xml': `<?xml version="1.0"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"/>`,
			'xl/ctrlProps/ctrlProp1.xml': `<?xml version="1.0"?><formControlPr macro="Module1.Run"/>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.activeContent).toEqual([
			{
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				contentType: 'application/vnd.ms-office.vbaProject',
				anchor: 'workbook',
				relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
				relationshipCount: 0,
				byteSize: 11,
				opaque: true,
				executionPolicy: 'blocked',
			},
			{
				kind: 'activeX',
				partPath: 'xl/activeX/activeX1.xml',
				contentType: 'application/vnd.ms-office.activeX+xml',
				anchor: 'sheet',
				sheetName: 'Data',
				relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControl',
				relationshipCount: 0,
			},
			{
				kind: 'formControl',
				partPath: 'xl/ctrlProps/ctrlProp1.xml',
				contentType: 'application/vnd.ms-excel.controlproperties+xml',
				anchor: 'sheet',
				sheetName: 'Data',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
				relationshipCount: 0,
			},
		])
		expect(result.value.report.status).toBe('has-unsupported')
		expect(
			result.value.report.features.find((feature) => feature.feature === 'vbaProject'),
		).toMatchObject({
			tier: 'unsupported',
			locations: ['xl/vbaProject.bin'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'activeX'),
		).toMatchObject({
			tier: 'unsupported',
			locations: ['xl/activeX/activeX1.xml'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedMacro'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/vbaProject.bin'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedActiveX'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/activeX/activeX1.xml'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedControl'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/ctrlProps/ctrlProp1.xml'],
		})
	})
})
