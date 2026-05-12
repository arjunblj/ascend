import { describe, expect, test } from 'bun:test'
import { unzipSync } from 'fflate'
import { makeXlsx } from '../../test/helpers.ts'
import { writeXlsx } from '../writer/index.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('macro sheet inventory', () => {
	test('discovers Excel 4 macro sheets without modeling them as worksheet grids', () => {
		const result = readXlsx(macroSheetWorkbook())
		expectOk(result)

		expect(result.value.workbook.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
		expect(result.value.workbook.macroSheets).toEqual([
			{
				name: 'Macro1',
				sheetId: '2',
				relId: 'rIdMacro',
				partPath: 'xl/macrosheets/sheet1.xml',
				state: 'veryHidden',
				relationshipCount: 1,
				dimensionRef: 'A1:B2',
				cellCount: 2,
				formulaCount: 1,
			},
		])
		expect(result.value.workbook.activeContent).toContainEqual({
			kind: 'macroSheet',
			partPath: 'xl/macrosheets/sheet1.xml',
			contentType: 'application/vnd.ms-excel.macrosheet+xml',
			anchor: 'sheet',
			sheetName: 'Macro1',
			relType: 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
			sourceRelationshipId: 'rIdMacro',
			relationshipCount: 1,
			opaque: true,
			executionPolicy: 'blocked',
		})
		expect(result.value.report.features).toContainEqual({
			feature: 'preservedMacroSheet',
			tier: 'preserved',
			count: 1,
			locations: ['xl/macrosheets/sheet1.xml'],
			note: 'Excel 4 macro sheets are inventoried and preserved exactly where possible; macro formulas are not executed or semantically edited.',
		})

		const written = writeXlsx(result.value.workbook, result.value.capsules, {
			workbookMetaDirty: true,
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		expect(zip['xl/macrosheets/sheet1.xml']).toBeDefined()
		expect(zip['xl/macrosheets/_rels/sheet1.xml.rels']).toBeDefined()
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		expect(workbookXml).toContain('name="Macro1"')
		expect(workbookXml).toContain('sheetId="2"')
		expect(workbookXml).toContain('state="veryHidden"')
		expect(workbookRels).toContain(
			'Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"',
		)
		expect(workbookRels).toContain('Target="macrosheets/sheet1.xml"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.macroSheets[0]).toMatchObject({
			name: 'Macro1',
			sheetId: '2',
			partPath: 'xl/macrosheets/sheet1.xml',
			state: 'veryHidden',
			dimensionRef: 'A1:B2',
			cellCount: 2,
			formulaCount: 1,
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
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
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
    <sheet name="Macro1" sheetId="2" r:id="rIdMacro" state="veryHidden"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/macrosheets/sheet1.xml': `<?xml version="1.0"?>
<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><f>ACTIVATE()</f><v>0</v></c></row>
    <row r="2"><c r="B2"><v>7</v></c></row>
  </sheetData>
  <drawing r:id="rIdDrawing"/>
</xm:macrosheet>`,
		'xl/macrosheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
		'xl/drawings/drawing1.xml': `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`,
	})
}
