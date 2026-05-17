import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('theme inventory', () => {
	test('parses theme palette slots and font names while preserving the source part', () => {
		const result = readXlsx(themeWorkbook())
		expectOk(result)

		expect(result.value.workbook.themeMetadata).toEqual({
			name: 'Office Twist',
			colorSchemeName: 'Office',
			colorCount: 4,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		})
		expect(result.value.workbook.themeColors).toEqual([
			{ slot: 'dk1', systemColor: 'windowText', lastColor: '000000' },
			{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
			{ slot: 'accent1', rgb: '4F81BD' },
			{ slot: 'hlink', rgb: '0000FF' },
		])
		expect(result.value.workbook.preservedTheme).toEqual({
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
		})
	})
})

function themeWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<d:theme xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Twist">
  <d:themeElements>
    <d:clrScheme name="Office">
      <d:dk1><d:sysClr val="windowText" lastClr="000000"/></d:dk1>
      <d:lt1><d:sysClr val="window" lastClr="FFFFFF"/></d:lt1>
      <d:accent1><d:srgbClr val="4F81BD"/></d:accent1>
      <d:hlink><d:srgbClr val="0000FF"/></d:hlink>
    </d:clrScheme>
    <d:fontScheme name="Office">
      <d:majorFont><d:latin typeface="Aptos Display"/></d:majorFont>
      <d:minorFont><d:latin typeface="Aptos"/></d:minorFont>
    </d:fontScheme>
  </d:themeElements>
</d:theme>`,
	})
}
