import { describe, expect, test } from 'bun:test'
import { type StyleId, Workbook } from '@ascend/core'
import { numberValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	if (!result.ok) throw new Error(result.error.message)
	expect(result.ok).toBe(true)
}

describe('visual conditional formatting', () => {
	test('round-trips color scales, data bars, and icon sets', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: S0 })
		}
		sheet.conditionalFormats.push(
			{
				sqref: 'A1:A5',
				rules: [
					{
						type: 'colorScale',
						priority: 1,
						formulas: [],
						colorScale: {
							cfvo: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
							colors: [{ rgb: 'FFF8696B' }, { rgb: 'FFFFEB84' }, { rgb: 'FF63BE7B' }],
						},
					},
				],
			},
			{
				sqref: 'B1:B5',
				rules: [
					{
						type: 'dataBar',
						priority: 2,
						formulas: [],
						dataBar: {
							cfvo: [{ type: 'min' }, { type: 'max' }],
							color: { rgb: 'FF638EC6' },
							showValue: false,
							minLength: 10,
							maxLength: 90,
						},
					},
				],
			},
			{
				sqref: 'C1:C5',
				rules: [
					{
						type: 'iconSet',
						priority: 3,
						formulas: [],
						iconSet: {
							iconSet: '3TrafficLights1',
							cfvo: [
								{ type: 'percent', value: '0' },
								{ type: 'percent', value: '33', gte: false },
								{ type: 'percent', value: '67' },
							],
							showValue: true,
							percent: true,
							reverse: false,
						},
					},
				],
			},
		)

		const written = writeXlsx(wb)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)

		expect(reopened.value.workbook.sheets[0]?.conditionalFormats).toEqual(sheet.conditionalFormats)
	})

	test('keeps x14 visual threshold formulas nested on dirty round-trip', () => {
		const source = x14DataBarFormulaThresholdWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)

		expect(opened.value.workbook.sheets[0]?.x14ConditionalFormats).toEqual([
			{
				index: 0,
				sqref: 'A1:A5',
				formulas: [],
				type: 'dataBar',
				priority: 4,
				dataBar: { cfvo: [{ type: 'formula', value: 'A1' }] },
			},
		])

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const sheetXml = decodeXml(unzipSync(written.value)['xl/worksheets/sheet1.xml'])

		expect(sheetXml).toContain('<x14:cfvo type="formula"><xm:f>A1</xm:f></x14:cfvo>')
		expect(sheetXml).not.toMatch(/<x14:cfRule\b[^>]*>\s*<xm:f>A1<\/xm:f>\s*<x14:dataBar>/)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.x14ConditionalFormats[0]).toMatchObject({
			formulas: [],
			dataBar: { cfvo: [{ type: 'formula', value: 'A1' }] },
		})
	})
})

function decodeXml(bytes: Uint8Array | undefined): string {
	expect(bytes).toBeDefined()
	return new TextDecoder().decode(bytes)
}

function x14DataBarFormulaThresholdWorkbook(): Uint8Array {
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
      <x14:conditionalFormattings>
        <x14:conditionalFormatting>
          <x14:cfRule type="dataBar" priority="4">
            <x14:dataBar><x14:cfvo type="formula"><xm:f>A1</xm:f></x14:cfvo></x14:dataBar>
          </x14:cfRule>
          <xm:sqref>A1:A5</xm:sqref>
        </x14:conditionalFormatting>
      </x14:conditionalFormattings>
    </ext>
  </extLst>
</worksheet>`,
	})
}
