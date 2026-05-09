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

describe('chartsheet inventory', () => {
	test('discovers chartsheets without modeling them as worksheet grids', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdChartSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rIdData"/>
    <sheet name="Sales Chart" sheetId="2" r:id="rIdChartSheet" state="hidden"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/chartsheets/sheet1.xml': `<?xml version="1.0"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="rIdChart"/></chartsheet>`,
			'xl/chartsheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
			'xl/charts/chart1.xml': `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart/></c:plotArea></c:chart>
</c:chartSpace>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
		expect(result.value.workbook.chartSheets).toEqual([
			{
				name: 'Sales Chart',
				sheetId: '2',
				relId: 'rIdChartSheet',
				partPath: 'xl/chartsheets/sheet1.xml',
				state: 'hidden',
				chartPartPaths: ['xl/charts/chart1.xml'],
			},
		])
		expect(result.value.workbook.chartParts[0]).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sales Chart',
			title: 'Sales',
		})
		expect(result.value.report.features).toContainEqual({
			feature: 'chartSheet',
			tier: 'unsupported',
			count: 1,
			locations: ['xl/chartsheets/sheet1.xml'],
			note: 'Chartsheets are inventoried but not editable as worksheet grids; writes require explicit loss approval.',
		})

		const written = writeXlsx(result.value.workbook, result.value.capsules, {
			workbookMetaDirty: true,
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		expect(zip['xl/chartsheets/sheet1.xml']).toBeDefined()
		expect(zip['xl/chartsheets/_rels/sheet1.xml.rels']).toBeDefined()
		expect(zip['xl/charts/chart1.xml']).toBeDefined()
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		expect(workbookXml).toContain('name="Sales Chart"')
		expect(workbookXml).toContain('sheetId="2"')
		expect(workbookXml).toContain('state="hidden"')
		expect(workbookRels).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"',
		)
		expect(workbookRels).toContain('Target="chartsheets/sheet1.xml"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.chartSheets[0]).toMatchObject({
			name: 'Sales Chart',
			sheetId: '2',
			partPath: 'xl/chartsheets/sheet1.xml',
			state: 'hidden',
			chartPartPaths: ['xl/charts/chart1.xml'],
		})
	})
})
