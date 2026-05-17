import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook, indexToColumn, Workbook } from '@ascend/core'
import { booleanValue, dateValue, EMPTY, numberValue, stringValue } from '@ascend/schema'
import { defaultCalcContext, recalculate } from '../../../engine/src/index.ts'
import { makeEmbeddedChartXlsx, makeXlsx } from '../../test/helpers.ts'
import { writeXlsx } from '../writer/index.ts'
import { readXlsx } from './index.ts'
import {
	emptySharedStrings,
	parseSharedStrings,
	parseSharedStringsBytes,
	parseSharedStringsChunks,
} from './shared-strings.ts'
import type { StreamedSheetRow } from './sheet.ts'
import {
	parseSheetFormulaOnlyBytes,
	parseSheetFullScalarBytes,
	parseSheetValuesOnlyBytes,
	streamSheetRowsByteChunks,
	streamSheetRowsTextChunks,
} from './sheet.ts'
import { readXlsxRowsStream } from './stream.ts'

const S0 = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function expectErr<T, E>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is {
	ok: false
	error: E
} {
	expect(result.ok).toBe(false)
	if (result.ok) throw new Error('Expected readXlsx to fail')
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`

const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`

const SHARED_STRINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
</sst>`

const SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>42</v></c>
      <c r="C1" t="b"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>1</v></c>
      <c r="B2"><f>B1*2</f><v>84</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:B1"/>
  </mergeCells>
</worksheet>`

function minimalXlsx(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': CONTENT_TYPES,
		'_rels/.rels': ROOT_RELS,
		'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
		'xl/workbook.xml': WORKBOOK_XML,
		'xl/sharedStrings.xml': SHARED_STRINGS,
		'xl/worksheets/sheet1.xml': SHEET_XML,
	})
}

describe('readXlsx', () => {
	it('parses a minimal XLSX with correct sheet and cell data', () => {
		const result = readXlsx(minimalXlsx())
		expectOk(result)

		const { workbook, report } = result.value

		expect(workbook.sheets).toHaveLength(1)
		const sheet = workbook.sheets[0]
		expect(sheet).toBeDefined()
		if (!sheet) return
		expect(sheet.name).toBe('Data')

		const a1 = sheet.cells.get(0, 0)
		expect(a1).toBeDefined()
		expect(a1?.value).toEqual({ kind: 'string', value: 'Hello' })

		const b1 = sheet.cells.get(0, 1)
		expect(b1).toBeDefined()
		expect(b1?.value).toEqual({ kind: 'number', value: 42 })

		const c1 = sheet.cells.get(0, 2)
		expect(c1).toBeDefined()
		expect(c1?.value).toEqual({ kind: 'boolean', value: true })

		const a2 = sheet.cells.get(1, 0)
		expect(a2).toBeDefined()
		expect(a2?.value).toEqual({ kind: 'string', value: 'World' })

		const b2 = sheet.cells.get(1, 1)
		expect(b2).toBeDefined()
		expect(b2?.value).toEqual({ kind: 'number', value: 84 })
		expect(b2?.formula).toBe('B1*2')

		expect(report.status).toBe('clean')
	})

	it('maps default workbook cell styles to the default style id', () => {
		const source = createWorkbook()
		const sourceSheet = source.addSheet('Sheet1')
		sourceSheet.cells.setPlainNumberSpan(0, 0, [1, 2, 3])

		const written = writeXlsx(source)
		expectOk(written)
		const result = readXlsx(written.value, { mode: 'full', richMetadata: true })
		expectOk(result)

		const workbook = result.value.workbook
		const sheet = workbook.getSheet('Sheet1')
		expect(sheet?.cells.get(0, 0)?.styleId).toBe(S0)
		expect(workbook.styles.size).toBe(1)
		expect(workbook.preservedStyles?.xfByStyleId[S0]).toBe(0)
		expect(sheet?.cells.storageStats().styleArrayBufferBytes).toBe(0)
	})

	it('parses compact inline strings without explicit cell refs', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c t="inlineStr"><is><t>alpha</t></is></c><c><v>42</v></c></row>
    <row r="2"><c t="inlineStr"><is><t>A&amp;B</t></is></c><c><v>84</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual(stringValue('alpha'))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(numberValue(42))
		expect(sheet?.cells.get(1, 0)?.value).toEqual(stringValue('A&B'))
		expect(sheet?.cells.get(1, 1)?.value).toEqual(numberValue(84))
	})

	it('parses XML-legal single-quoted worksheet attributes', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref='A1:F2'/>
  <sheetData>
    <row r='1' ht='22' customHeight='1' hidden='1' collapsed='1' outlineLevel='2'>
      <c r='A1' t='s'><v>0</v></c>
      <c r='B1'><v>42</v></c>
      <c r='C1' t='b'><v>1</v></c>
      <c r='D1'><f t='shared' si='0' ref='D1:D2'>B1*2</f><v>84</v></c>
      <c r='E1' t='inlineStr'><is><r><rPr><b/><rFont val='Aptos'/><sz val='11'/><color rgb='FF00AA00'/></rPr><t>Rich</t></r></is></c>
    </row>
    <row r='2'>
      <c r='B2'><v>84</v></c>
      <c r='D2'><f t='shared' si='0'/><v>168</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		const sheet = result.value.workbook.sheets[0]

		expect(sheet?.preservedDimensionRef).toBe('A1:F2')
		expect(sheet?.cells.get(0, 0)?.value).toEqual(stringValue('Hello'))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(numberValue(42))
		expect(sheet?.cells.get(0, 2)?.value).toEqual(booleanValue(true))
		expect(sheet?.cells.get(0, 3)?.formula).toBe('B1*2')
		expect(sheet?.cells.get(0, 3)?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			ref: 'D1:D2',
		})
		expect(sheet?.cells.get(1, 3)?.formula).toBeNull()
		expect(sheet?.cells.get(1, 3)?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'D1',
		})
		expect(sheet?.cells.get(0, 4)?.value).toEqual({
			kind: 'richText',
			runs: [
				{
					text: 'Rich',
					bold: true,
					fontName: 'Aptos',
					fontSize: 11,
					color: 'FF00AA00',
				},
			],
		})
		expect(sheet?.rowHeights.get(0)).toBe(22)
		expect(sheet?.rowDefs.get(0)).toEqual({
			customHeight: true,
			hidden: true,
			collapsed: true,
			outlineLevel: 2,
		})
	})

	it('parses XML-legal single-quoted package and workbook attributes', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
  <Default Extension='xml' ContentType='application/xml'/>
  <Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>
  <Override PartName='/xl/worksheets/sheet1.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>
  <Override PartName='/xl/sharedStrings.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'/>
</Types>`,
			'_rels/.rels': `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rIdOffice' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rIdSheet' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet1.xml'/>
  <Relationship Id='rIdStrings' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings' Target='sharedStrings.xml'/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
  <workbookPr date1904='1' codeName='Book&amp;One'/>
  <bookViews><workbookView activeTab='0'/></bookViews>
  <sheets>
    <sheet name='Data &amp; Model' sheetId='1' r:id='rIdSheet'/>
  </sheets>
  <definedNames><definedName name='Total' localSheetId='0'>A1</definedName></definedNames>
  <calcPr calcMode='manual'/>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.name).toBe('Data & Model')
		expect(result.value.workbook.calcSettings.dateSystem).toBe('1904')
		expect(result.value.workbook.calcSettings.calcMode).toBe('manual')
		expect(result.value.workbook.workbookProperties.codeName).toBe('Book&One')
		expect(result.value.workbook.definedNames.list()).toContainEqual({
			name: 'Total',
			formula: 'A1',
			scope: { kind: 'sheet', sheetId: result.value.workbook.sheets[0]?.id },
		})
		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual(stringValue('Hello'))
	})

	it('normalizes backslash ZIP entry paths from non-standard producers', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels\\.rels': ROOT_RELS,
			'xl\\_rels\\workbook.xml.rels': WORKBOOK_RELS,
			'xl\\workbook.xml': WORKBOOK_XML,
			'xl\\sharedStrings.xml': SHARED_STRINGS,
			'xl\\worksheets\\sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.name).toBe('Data')
		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Hello',
		})
	})

	it('resolves URI-encoded internal relationship targets to package parts', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/work book.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet 1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/shared Strings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/work%20book.xml"/>
</Relationships>`,
			'xl/_rels/work book.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets\\sheet%201.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="shared%20Strings.xml"/>
</Relationships>`,
			'xl/work book.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/shared Strings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet 1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Hello',
		})
	})

	it('parses prefixed workbook namespace elements', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0" encoding="utf-8"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr date1904="1" codeName="ThisWorkbook"/>
  <x:bookViews><x:workbookView activeTab="0"/></x:bookViews>
  <x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets>
  <x:definedNames><x:definedName name="Total">Data!$B$1</x:definedName></x:definedNames>
  <x:calcPr calcMode="manual"/>
</x:workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
		expect(result.value.workbook.calcSettings.dateSystem).toBe('1904')
		expect(result.value.workbook.calcSettings.calcMode).toBe('manual')
		expect(result.value.workbook.workbookProperties.codeName).toBe('ThisWorkbook')
		expect(result.value.workbook.workbookViews).toEqual([{ activeTab: 0 }])
		expect(result.value.workbook.definedNames.get('Total')).toBe('Data!$B$1')
	})

	it('parses XML-legal strict workbook namespace declarations', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0" encoding="utf-8"?>
<x:workbook xmlns:x='http://purl.oclc.org/ooxml/spreadsheetml/main'
  xmlns:rel='http://purl.oclc.org/ooxml/officeDocument/relationships'>
  <x:workbookPr date1904='1' codeName='StrictWorkbook'/>
  <x:sheets><x:sheet name='Data' sheetId='1' rel:id='rId1'/></x:sheets>
  <x:definedNames><x:definedName name='StrictName'>Data!$A$1</x:definedName></x:definedNames>
  <x:calcPr calcMode='manual'/>
</x:workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
		expect(result.value.workbook.calcSettings.dateSystem).toBe('1904')
		expect(result.value.workbook.calcSettings.calcMode).toBe('manual')
		expect(result.value.workbook.workbookProperties.codeName).toBe('StrictWorkbook')
		expect(result.value.workbook.definedNames.get('StrictName')).toBe('Data!$A$1')
	})

	it('parses sheet relationship ids with non-r namespace prefixes', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:AlternateContent><mc:Fallback/></mc:AlternateContent>
  <sheets>
    <sheet name="Data" sheetId="1"
      xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      relationships:id="rId1"/>
  </sheets>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.name).toBe('Data')
		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Hello',
		})
	})

	it('parses prefixed shared-string namespace elements', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="utf-8"?>
<x:sst xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' count="2" uniqueCount="2">
  <x:si><x:t>A</x:t></x:si>
  <x:si><x:r><x:rPr><x:b/></x:rPr><x:t>B</x:t></x:r></x:si>
</x:sst>`,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'A',
		})
		expect(result.value.workbook.sheets[0]?.cells.get(1, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'B', bold: true }],
		})
	})

	it('parses XML-legal prefixed worksheet namespace declarations', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>
  <x:sheetData>
    <x:row r='1'><x:c r='A1' t='s'><x:v>0</x:v></x:c></x:row>
  </x:sheetData>
  <x:mergeCells count='1'><x:mergeCell ref='A1:B1'/></x:mergeCells>
</x:worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(sheet?.merges).toEqual([{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }])
	})

	it('values mode reads plain string cells without full cell XML parsing', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="str"><v>plain &amp; fast</v></c>
      <c r="B1" t="str"><f>TEXT(1,"0")</f><v>cached text</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'plain & fast' })
		expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'cached text' })
		expect(sheet?.cells.get(0, 1)?.formula).toBeNull()
	})

	it('advances omitted cell references across skipped empty cells', async () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c/>
      <c t="str"><v>after empty</v></c>
      <c s="1"/>
      <c><v>42</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		for (const mode of ['formula', 'values'] as const) {
			const result = readXlsx(bytes, { mode })
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet?.cells.get(0, 0)).toBeUndefined()
			expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'after empty' })
			expect(sheet?.cells.get(0, 2)).toBeUndefined()
			expect(sheet?.cells.get(0, 3)?.value).toEqual({ kind: 'number', value: 42 })
		}

		const streamResult = await readXlsxRowsStream(bytes, { sheet: 'Data', mode: 'values' })
		expectOk(streamResult)
		const rows: StreamedSheetRow[] = []
		for await (const row of streamResult.value) rows.push(row)
		expect(rows).toEqual([
			{
				row: 0,
				cells: [
					[1, { value: { kind: 'string', value: 'after empty' }, formula: null, styleId: S0 }],
					[3, { value: { kind: 'number', value: 42 }, formula: null, styleId: S0 }],
				],
			},
		])
	})

	it('does not materialize typed blank placeholders without cached values', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="str"/>
      <c r="B1" t="str"></c>
      <c r="C1" t="n"/>
      <c r="D1" t="s"></c>
	      <c r="E1" t="str"><f>A2</f></c>
	      <c r="F1" t="n"><f>A3</f></c>
	      <c r="G1" t="str"><v></v></c>
	      <c r="H1" t="s"><v/></c>
	      <c r="I1" t="str"><v /></c>
	    </row>
	  </sheetData>
	</worksheet>`,
		})

		for (const mode of ['formula', 'values'] as const) {
			const result = readXlsx(bytes, { mode })
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet?.cells.get(0, 0)).toBeUndefined()
			expect(sheet?.cells.get(0, 1)).toBeUndefined()
			expect(sheet?.cells.get(0, 2)).toBeUndefined()
			expect(sheet?.cells.get(0, 3)).toBeUndefined()
			expect(sheet?.cells.get(0, 4)?.value).toEqual(EMPTY)
			expect(sheet?.cells.get(0, 5)?.value).toEqual(EMPTY)
			expect(sheet?.cells.get(0, 4)?.formula).toBe(mode === 'formula' ? 'A2' : null)
			expect(sheet?.cells.get(0, 5)?.formula).toBe(mode === 'formula' ? 'A3' : null)
			expect(sheet?.cells.get(0, 6)?.value).toEqual({ kind: 'string', value: '' })
			expect(sheet?.cells.get(0, 7)?.value).toEqual({ kind: 'string', value: '' })
			expect(sheet?.cells.get(0, 8)?.value).toEqual({ kind: 'string', value: '' })
		}
	})

	it('does not materialize blank inline-string placeholders', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"/>
      <c r="B1" t="inlineStr"></c>
      <c r="C1" t="inlineStr"><is><t></t></is></c>
      <c r="D1" t="inlineStr"><is><t>value</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		for (const mode of ['formula', 'values'] as const) {
			const result = readXlsx(bytes, { mode })
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet?.cells.get(0, 0)).toBeUndefined()
			expect(sheet?.cells.get(0, 1)).toBeUndefined()
			expect(sheet?.cells.get(0, 2)?.value).toEqual({ kind: 'string', value: '' })
			expect(sheet?.cells.get(0, 3)?.value).toEqual({ kind: 'string', value: 'value' })
		}
	})

	it('streams worksheet rows through the async reader path', async () => {
		const result = await readXlsxRowsStream(minimalXlsx(), { sheet: 'Data', mode: 'formula' })
		expectOk(result)
		const rows: StreamedSheetRow[] = []
		for await (const row of result.value) rows.push(row)
		expect(rows).toEqual([
			{
				row: 0,
				cells: [
					[0, { value: { kind: 'string', value: 'Hello' }, formula: null, styleId: S0 }],
					[1, { value: { kind: 'number', value: 42 }, formula: null, styleId: S0 }],
					[2, { value: { kind: 'boolean', value: true }, formula: null, styleId: S0 }],
				],
			},
			{
				row: 1,
				cells: [
					[0, { value: { kind: 'string', value: 'World' }, formula: null, styleId: S0 }],
					[1, { value: { kind: 'number', value: 84 }, formula: 'B1*2', styleId: S0 }],
				],
			},
		])
	})

	it('streams a bounded values row window without requiring callers to drain the sheet', async () => {
		const result = await readXlsxRowsStream(minimalXlsx(), {
			sheet: 'Data',
			mode: 'values',
			maxRows: 1,
		})
		expectOk(result)
		const rows: StreamedSheetRow[] = []
		for await (const row of result.value) rows.push(row)
		expect(rows).toEqual([
			{
				row: 0,
				cells: [
					[0, { value: { kind: 'string', value: 'Hello' }, formula: null, styleId: S0 }],
					[1, { value: { kind: 'number', value: 42 }, formula: null, styleId: S0 }],
					[2, { value: { kind: 'boolean', value: true }, formula: null, styleId: S0 }],
				],
			},
		])
	})

	it('streams worksheet rows from an async iterable byte source', async () => {
		const bytes = minimalXlsx()
		async function* chunks() {
			yield bytes.subarray(0, Math.floor(bytes.length / 2))
			yield bytes.subarray(Math.floor(bytes.length / 2))
		}

		const result = await readXlsxRowsStream(chunks(), { sheet: 'Data', mode: 'values' })
		expectOk(result)
		const rows: StreamedSheetRow[] = []
		for await (const row of result.value) rows.push(row)
		expect(rows[0]?.cells[0]?.[1]?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(rows[1]?.cells[0]?.[1]?.value).toEqual({ kind: 'string', value: 'World' })
	})

	it('streams an empty worksheet without yielding rows', async () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData />
</worksheet>`,
		})

		const result = await readXlsxRowsStream(bytes, { sheet: 'Data', mode: 'values' })
		expectOk(result)
		const rows: StreamedSheetRow[] = []
		for await (const row of result.value) rows.push(row)
		expect(rows).toEqual([])
	})

	it('streams worksheet text chunks split inside rows, cells, and values', () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c><v>2</v></c></row><row><c t="str"><v>plain &amp; text</v></c></row></sheetData></worksheet>'
		const rows = [
			...streamSheetRowsTextChunks(
				'Data',
				[xml.slice(0, 29), xml.slice(29, 53), xml.slice(53, 91), xml.slice(91)],
				{
					sharedStrings: emptySharedStrings(),
					styleIds: [S0],
					isDateFormat: [false],
					valuesOnly: true,
				},
			),
		]
		expect(rows).toEqual([
			{
				row: 0,
				cells: [
					[0, { value: { kind: 'number', value: 1 }, formula: null, styleId: S0 }],
					[1, { value: { kind: 'number', value: 2 }, formula: null, styleId: S0 }],
				],
			},
			{
				row: 1,
				cells: [
					[0, { value: { kind: 'string', value: 'plain & text' }, formula: null, styleId: S0 }],
				],
			},
		])
	})

	it('streams worksheet byte chunks split inside rows, cells, and values', () => {
		const xml =
			'<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c><v>2</v></c><c r="C1" s="1"><v>45293</v></c><c r="D1" s="1"><f>C1+1</f><v>45294</v></c></row><row><c t="str"><v>plain &amp; text</v></c></row><row r="3"><c r="A3"><v>3'
		const bytes = new TextEncoder().encode(xml)
		const rows = [
			...streamSheetRowsByteChunks(
				'Data',
				[bytes.subarray(0, 29), bytes.subarray(29, 53), bytes.subarray(53, 91), bytes.subarray(91)],
				{
					sharedStrings: emptySharedStrings(),
					styleIds: [S0],
					isDateFormat: [false, true],
					hasDateStyles: true,
					valuesOnly: true,
					maxRows: 2,
				},
			),
		]
		expect(rows).toEqual([
			{
				row: 0,
				cells: [
					[0, { value: { kind: 'number', value: 1 }, formula: null, styleId: S0 }],
					[1, { value: { kind: 'number', value: 2 }, formula: null, styleId: S0 }],
					[2, { value: { kind: 'date', serial: 45293 }, formula: null, styleId: S0 }],
					[3, { value: { kind: 'date', serial: 45294 }, formula: null, styleId: S0 }],
				],
			},
			{
				row: 1,
				cells: [
					[0, { value: { kind: 'string', value: 'plain & text' }, formula: null, styleId: S0 }],
				],
			},
		])
	})

	it('parses merge cells', () => {
		const result = readXlsx(minimalXlsx())
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		if (!sheet) return
		expect(sheet.merges).toHaveLength(1)
		expect(sheet.merges[0]).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 1 },
		})
	})

	it('returns error for invalid ZIP data', () => {
		const garbage = new Uint8Array([1, 2, 3, 4, 5])
		const result = readXlsx(garbage)
		expectErr(result)
		expect(result.error.code).toBe('CORRUPT_FILE')
	})

	it('returns error for ZIP missing required parts', () => {
		const empty = makeXlsx({ 'dummy.txt': 'nothing' })
		const result = readXlsx(empty)
		expectErr(result)
		expect(result.error.code).toBe('CORRUPT_FILE')
	})

	it('keeps parsed dense sheets in dense chunks', () => {
		const wb = createWorkbook()
		wb.addSheet('Dense')
		const sheet = wb.sheets[0]
		if (!sheet) throw new Error('no sheet')
		for (let r = 0; r < 64; r++) {
			for (let c = 0; c < 64; c++) {
				sheet.cells.set(r, c, {
					value: numberValue(r * 64 + c),
					formula: null,
					styleId: S0,
				})
			}
		}
		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(written.error.message)
		const result = readXlsx(written.value)
		expectOk(result)
		const readSheet = result.value.workbook.sheets[0]
		expect(readSheet?.cells.getChunkKindAt(0, 0)).toBe('dense')
	})

	it('keeps parsed narrow dense sheets in dense chunks', () => {
		const rows: string[] = []
		for (let r = 1; r <= 160; r++) {
			const cells: string[] = []
			for (let c = 0; c < 5; c++) {
				const ref = `${indexToColumn(c)}${r}`
				cells.push(`<c r="${ref}"><v>${r * 5 + c}</v></c>`)
			}
			rows.push(`<row r="${r}">${cells.join('')}</row>`)
		}
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:E160"/>
  <sheetData>${rows.join('')}</sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		const readSheet = result.value.workbook.sheets[0]
		expect(readSheet?.cells.get(159, 4)?.value).toEqual({ kind: 'number', value: 804 })
		expect(readSheet?.cells.getChunkKindAt(0, 0)).toBe('dense')
	})

	it('does not densify sparse sheets from a stale large dimension', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:XFD1048576"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="1048576"><c r="XFD1048576"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		const readSheet = result.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(1_048_575, 16_383)?.value).toEqual({
			kind: 'number',
			value: 2,
		})
		expect(readSheet?.cells.getChunkKindAt(0, 0)).toBe('sparse')
		expect(readSheet?.cells.getChunkKindAt(1_048_575, 16_383)).toBe('sparse')
	})

	it('handles workbook with no shared strings', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>99</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const cell = result.value.workbook.sheets[0]?.cells.get(0, 0)
		expect(cell?.value).toEqual({ kind: 'number', value: 99 })
	})

	it('infers omitted row and cell references during sheet parsing', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row>
      <c t="s"><v>0</v></c>
      <c><v>7</v></c>
    </row>
    <row>
      <c><f>A1+B1</f><v>14</v></c>
    </row>
    <row r="5">
      <c r="C5"><v>3</v></c>
      <c><v>4</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 7 })
		expect(sheet?.cells.get(1, 0)?.formula).toBe('A1+B1')
		expect(sheet?.cells.get(1, 0)?.value).toEqual({ kind: 'number', value: 14 })
		expect(sheet?.cells.get(4, 2)?.value).toEqual({ kind: 'number', value: 3 })
		expect(sheet?.cells.get(4, 3)?.value).toEqual({ kind: 'number', value: 4 })
	})

	it('captures shared strings as a preserved workbook part', () => {
		const result = readXlsx(minimalXlsx())
		expectOk(result)

		expect(result.value.workbook.preservedSharedStrings).toEqual({
			path: 'xl/sharedStrings.xml',
		})
	})

	it('parses rich-text shared strings without losing formatting runs', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Hello</t></r>
    <r><t>World</t></r>
  </si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'Hello', bold: true }, { text: 'World' }],
		})
	})

	it('parses escaped and empty shared strings without DOM fallback', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t xml:space="preserve">  &lt;tag&gt; &amp; &quot;q&quot;  </t></si>
  <si><t/></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: '  <tag> & "q"  ',
		})
		expect(result.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'string',
			value: '',
		})
	})

	it('eager shared strings preserve plain text, entities, and xml:space fallbacks', () => {
		const sharedStrings =
			parseSharedStrings(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Alpha</t></si>
  <si><t>Beta &amp; Gamma</t></si>
  <si><t xml:space="preserve">  Delta  </t></si>
</sst>`)

		expect(sharedStrings.get(0)).toEqual({ kind: 'string', value: 'Alpha' })
		expect(sharedStrings.get(1)).toEqual({ kind: 'string', value: 'Beta & Gamma' })
		expect(sharedStrings.get(2)).toEqual({ kind: 'string', value: '  Delta  ' })
	})

	it('eager byte shared strings fast-path plain text and fallback for rich text', () => {
		const plainBytes =
			new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Alpha</t></si><si><t>Beta &amp; Gamma</t></si></sst>`)
		const plain = parseSharedStringsBytes(plainBytes)
		expect(plain.getString?.(0)).toBe('Alpha')
		expect(plain.get(1)).toEqual({ kind: 'string', value: 'Beta & Gamma' })

		const richBytes =
			new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><r><rPr><b/></rPr><t>Bold</t></r><r><t>Text</t></r></si><si><t xml:space="preserve">  Delta  </t></si></sst>`)
		const rich = parseSharedStringsBytes(richBytes)
		expect(rich.get(0)).toEqual({
			kind: 'richText',
			runs: [{ text: 'Bold', bold: true }, { text: 'Text' }],
		})
		expect(rich.getString?.(1)).toBe('  Delta  ')
	})

	it('lazy byte shared strings preserve simple, escaped, rich, and xml:space entries', () => {
		const bytes = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Alpha</t></si>
  <si><t>Beta &amp; Gamma</t></si>
  <si><r><rPr><b/></rPr><t>Bold</t></r><r><t>Text</t></r></si>
  <si><t xml:space="preserve">  Delta  </t></si>
</sst>`)
		const sharedStrings = parseSharedStringsBytes(bytes, { lazy: true })

		expect(sharedStrings.getString?.(0)).toBe('Alpha')
		expect(sharedStrings.getString?.(1)).toBe('Beta & Gamma')
		expect(sharedStrings.getString?.(2)).toBeUndefined()
		expect(sharedStrings.get(2)).toEqual({
			kind: 'richText',
			runs: [{ text: 'Bold', bold: true }, { text: 'Text' }],
		})
		expect(sharedStrings.getString?.(3)).toBe('  Delta  ')
		expect(sharedStrings.get(1)).toEqual({ kind: 'string', value: 'Beta & Gamma' })
	})

	it('keeps lazy plain shared strings bounded when phonetic metadata follows text', () => {
		const sharedStrings = parseSharedStrings(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>First</t><phoneticPr fontId="2" type="noConversion"/></si>
  <si><t>Second</t></si>
</sst>`,
			{ lazy: true },
		)

		expect(sharedStrings.getString?.(0)).toBe('First')
		expect(sharedStrings.getString?.(1)).toBe('Second')
		expect(sharedStrings.get(0)).toEqual({ kind: 'string', value: 'First' })
	})

	it('streams shared string chunks on demand with fallback for unsupported markup', () => {
		const chunks = [
			'<?xml version="1.0"?><sst>',
			'<si><t>Alpha</t></si><si><t>Be',
			'ta</t></si><si><r><t>Rich</t></r></si></sst>',
		]
		let fallbackCalls = 0
		const sharedStrings = parseSharedStringsChunks(chunks, {
			fallback: () => {
				fallbackCalls += 1
				return emptySharedStrings()
			},
		})
		expect(sharedStrings.getString?.(1)).toBe('Beta')
		expect(sharedStrings.get(2)).toEqual({ kind: 'string', value: 'Rich' })
		expect(fallbackCalls).toBe(0)

		const prefixed = parseSharedStringsChunks(['<x:sst><x:si><x:t>Fallback</x:t></x:si></x:sst>'], {
			fallback: () => {
				fallbackCalls += 1
				return parseSharedStrings(
					'<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:si><x:t>Fallback</x:t></x:si></x:sst>',
				)
			},
		})
		expect(prefixed.getString?.(0)).toBe('Fallback')
		expect(fallbackCalls).toBe(1)
	})

	it('parses theme and indexed colors in shared string rich text runs', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/><color theme="1" tint="0.25"/></rPr><t>theme</t></r>
    <r><rPr><i/><color indexed="64"/></rPr><t>indexed</t></r>
  </si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const cell = result.value.workbook.sheets[0]?.cells.get(0, 0)
		expect(cell?.value).toEqual({
			kind: 'richText',
			runs: [
				{ text: 'theme', bold: true, color: { kind: 'theme', theme: 1, tint: 0.25 } },
				{ text: 'indexed', italic: true, color: { kind: 'indexed', index: 64 } },
			],
		})
	})

	it('preserves xml:space text in DOM fallback rich-text parsing', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><r><rPr><b/></rPr><t xml:space="preserve">  hi  </t></r></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: '  hi  ', bold: true }],
		})
	})

	it('handles error cell type', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const cell = result.value.workbook.sheets[0]?.cells.get(0, 0)
		expect(cell?.value).toEqual({ kind: 'error', value: '#DIV/0!' })
	})

	it('reads defined names from workbook', () => {
		const wbXml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="Total">Data!$A$1</definedName>
  </definedNames>
</workbook>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': wbXml,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.definedNames.get('Total')).toBe('Data!$A$1')
	})

	it('preserves duplicate defined names in the same scope', () => {
		const wbXml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$2</definedName>
    <definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="0">Data!$A$1:$B$3</definedName>
  </definedNames>
</workbook>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': wbXml,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.getSheet('Data')
		expect(sheet).toBeDefined()
		if (!sheet) return

		const entries = result.value.workbook.definedNames.list()
		expect(entries).toHaveLength(2)
		expect(entries.map((entry) => entry.formula)).toEqual(['Data!$A$1:$B$2', 'Data!$A$1:$B$3'])
		expect(
			result.value.workbook.definedNames.resolve('_xlnm._FilterDatabase', sheet.id)?.formula,
		).toBe('Data!$A$1:$B$2')
	})

	it('decodes escaped workbook defined-name formulas', () => {
		const wbXml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="Escaped">Data!$A$1&amp;&quot;x&quot;</definedName>
  </definedNames>
</workbook>`
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': wbXml,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.definedNames.get('Escaped')).toBe('Data!$A$1&"x"')
	})

	it('reads sheet-scoped defined names from localSheetId', () => {
		const wbXml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Summary" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="Budget" localSheetId="1">Summary!$A$1</definedName>
  </definedNames>
</workbook>`
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
			'xl/workbook.xml': wbXml,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const summary = result.value.workbook.getSheet('Summary')
		expect(summary).toBeDefined()
		if (!summary) return

		const resolved = result.value.workbook.definedNames.resolve('Budget', summary.id)
		expect(resolved?.scope.kind).toBe('sheet')
		expect(resolved?.formula).toBe('Summary!$A$1')
	})

	it('reads hidden and extended defined-name metadata', () => {
		const wbXml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1" comment="Filter menu" description="Visible &amp; hidden rows" function="1" vbProcedure="1" xlm="1" functionGroupId="7" shortcutKey="K">Data!$A$1:$B$10</definedName>
  </definedNames>
</workbook>`
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': wbXml,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.getSheet('Data')
		expect(sheet).toBeDefined()
		if (!sheet) return

		const entry = result.value.workbook.definedNames.resolve('_xlnm._FilterDatabase', sheet.id)
		expect(entry?.hidden).toBe(true)
		expect(entry?.formula).toBe('Data!$A$1:$B$10')
		expect(entry?.extraAttributes).toEqual([
			{ name: 'comment', value: 'Filter menu' },
			{ name: 'description', value: 'Visible & hidden rows' },
			{ name: 'function', value: '1' },
			{ name: 'vbProcedure', value: '1' },
			{ name: 'xlm', value: '1' },
			{ name: 'functionGroupId', value: '7' },
			{ name: 'shortcutKey', value: 'K' },
		])
	})

	it('reads array formula text from structured formula nodes', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>3</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const cell = result.value.workbook.sheets[0]?.cells.get(0, 0)
		expect(cell?.formula).toBe('SUM(B1:B2)')
		expect(cell?.formulaInfo).toEqual({ kind: 'array', ref: 'A1:A2' })
		expect(
			result.value.report.features.find((feature) => feature.feature === 'arrayFormula'),
		).toMatchObject({
			feature: 'arrayFormula',
			tier: 'normalized',
		})
	})

	it('infers imported legacy array blocks from cached relay formulas', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="Global4">{1,2,3;4,5,6}</definedName>
  </definedNames>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="4">
      <c r="A4"><f>Global4</f><v>1</v></c>
      <c r="B4"><f>A4</f><v>2</v></c>
      <c r="C4"><f>A4</f><v>3</v></c>
    </row>
    <row r="5">
      <c r="A5"><f>A4</f><v>4</v></c>
      <c r="B5"><f>A4</f><v>5</v></c>
      <c r="C5"><f>A4</f><v>6</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		if (!sheet) return

		const binding = { kind: 'array' as const, ref: 'A4:C5' }
		expect(sheet.cells.get(3, 0)?.formula).toBe('Global4')
		expect(sheet.cells.get(3, 0)?.formulaInfo).toEqual(binding)
		expect(sheet.cells.get(3, 1)?.formula).toBeNull()
		expect(sheet.cells.get(4, 2)?.formula).toBeNull()
		expect(sheet.cells.get(3, 1)?.formulaInfo).toEqual(binding)
		expect(sheet.cells.get(4, 2)?.formulaInfo).toEqual(binding)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'arrayFormula'),
		).toMatchObject({
			feature: 'arrayFormula',
			tier: 'normalized',
			count: 1,
			locations: ['Data'],
		})

		recalculate(result.value.workbook, defaultCalcContext())
		expect(sheet.cells.get(3, 0)?.value).toEqual(numberValue(1))
		expect(sheet.cells.get(3, 1)?.value).toEqual(numberValue(2))
		expect(sheet.cells.get(3, 2)?.value).toEqual(numberValue(3))
		expect(sheet.cells.get(4, 0)?.value).toEqual(numberValue(4))
		expect(sheet.cells.get(4, 1)?.value).toEqual(numberValue(5))
		expect(sheet.cells.get(4, 2)?.value).toEqual(numberValue(6))
	})

	it('reads data-table formula anchors as symbolic formula bindings', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>0.5</v></c></row>
    <row r="2"><c r="B2"><f>SUM(A1)</f><v>0.5</v></c></row>
    <row r="3"><c r="C3"><f t="dataTable" ref="C3:C5" dt2D="0" dtr="1" r1="A1"/><v>10</v></c></row>
    <row r="4"><c r="C4"><v>20</v></c></row>
    <row r="5"><c r="C5"><v>30</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		const dataTableCell = sheet?.cells.get(2, 2)
		expect(sheet?.cells.formulaCellCount()).toBe(2)
		expect(sheet?.cells.formulaInfoCellCount()).toBe(1)
		expect(dataTableCell?.formula).toBeNull()
		expect(dataTableCell?.formulaInfo).toEqual({
			kind: 'dataTable',
			ref: 'C3:C5',
			dt2D: false,
			dtr: true,
			r1: 'A1',
		})
		expect(dataTableCell?.value).toEqual({ kind: 'number', value: 10 })
	})

	it('flags shared formulas in the compatibility report', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="shared" si="0">B1*2</f><v>84</v></c>
      <c r="A2"><f t="shared" si="0"/><v>168</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(
			result.value.report.features.find((feature) => feature.feature === 'sharedFormula'),
		).toMatchObject({
			feature: 'sharedFormula',
			tier: 'normalized',
		})
		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'A1',
		})
		expect(result.value.workbook.sheets[0]?.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A1',
		})
		expect(result.value.workbook.sheets[0]?.cells.get(1, 0)?.formula).toBeNull()
	})

	it('reads shared formulas with master formula text, member shared info, and correct values after recalc', async () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B1"><v>42</v></c>
      <c r="A1"><f t="shared" si="0">B1*2</f><v>84</v></c>
    </row>
    <row r="2">
      <c r="B2"><v>84</v></c>
      <c r="A2"><f t="shared" si="0"/><v>168</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet).toBeDefined()

		const masterCell = sheet?.cells.get(0, 0)
		expect(masterCell?.formula).toBe('B1*2')
		expect(masterCell?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'A1',
		})

		const memberCell = sheet?.cells.get(1, 0)
		expect(memberCell?.formula).toBeNull()
		expect(memberCell?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A1',
		})

		recalculate(result.value.workbook, defaultCalcContext())

		expect(sheet?.cells.get(0, 0)?.value).toMatchObject({ kind: 'number', value: 84 })
		expect(sheet?.cells.get(1, 0)?.value).toMatchObject({ kind: 'number', value: 168 })
	})

	it('imports dynamic-array metadata and normalizes storage formulas', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/metadata.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c>
      <c r="B1"><f>SUM(_xlfn.ANCHORARRAY(A1))</f><v>6</v></c>
      <c r="C1"><f>_xlfn.SINGLE(A1)</f><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(sheet?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(sheet?.cells.get(0, 1)?.formula).toBe('SUM(A1#)')
		expect(sheet?.cells.get(0, 2)?.formula).toBe('@A1')
		expect(
			result.value.report.features.find((feature) => feature.feature === 'dynamicArray'),
		).toMatchObject({
			feature: 'dynamicArray',
			tier: 'normalized',
		})
	})

	it('parses worksheet tables from sheet relationships', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Name</v></c><c r="B1" t="str"><v>Score</v></c></row>
    <row r="2"><c r="A2" t="str"><v>Mina</v></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="str"><v>Noah</v></c><c r="B3"><v>12</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
			'xl/tables/table1.xml': `<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" id="1" xr:uid="{TABLE-UID}" name="Table1" displayName="Scores" ref="A1:B3" headerRowCount="1" totalsRowCount="0" dataCellStyle="TableData" headerRowDxfId="5" headerRowCellStyle="TableHeader" dataDxfId="6" tableBorderDxfId="9">
  <autoFilter ref="A1:B3" xr:uid="{FILTER-UID}"><sortState ref="A2:B3"><sortCondition ref="B2:B3"/></sortState></autoFilter>
  <sortState ref="A2:B3" xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"><sortCondition ref="A2:A3"/></sortState>
  <tableColumns count="2">
    <tableColumn id="1" xr3:uid="{COLUMN-1-UID}" name="Name " totalsRowLabel="Total" dataDxfId="7"><calculatedColumnFormula array="1">[@Score]*2</calculatedColumnFormula></tableColumn>
    <tableColumn id="2" xr3:uid="{COLUMN-2-UID}" name="Score" totalsRowFunction="sum" totalsRowDxfId="8"><xmlColumnPr mapId="2" xpath="/ns1:Scores/ns1:Score" xmlDataType="double"/></tableColumn>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
  <extLst><ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:table altText="Scores" altTextSummary="Student score table"/></ext></extLst>
</table>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.tables).toHaveLength(1)
		expect(sheet?.tables[0]?.name).toBe('Scores')
		expect(sheet?.tables[0]).toMatchObject({
			partPath: 'xl/tables/table1.xml',
			ooxmlId: 1,
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
			contentTypeSource: 'override',
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			sourceRelationshipPart: 'xl/worksheets/_rels/sheet1.xml.rels',
			sourceRelationshipId: 'rId1',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			sourceRelationshipRawTarget: '../tables/table1.xml',
			sourceRelationshipResolvedTarget: 'xl/tables/table1.xml',
		})
		expect(sheet?.tables[0]?.nameAttribute).toBe('Table1')
		expect(sheet?.tables[0]?.hasHeaders).toBe(true)
		expect(sheet?.tables[0]?.columns).toEqual([
			{
				id: 1,
				uid: '{COLUMN-1-UID}',
				name: 'Name ',
				formula: '[@Score]*2',
				formulaIsArray: true,
				totalsRowLabel: 'Total',
				dataDxfId: 7,
			},
			{
				id: 2,
				uid: '{COLUMN-2-UID}',
				name: 'Score',
				xmlColumnPr: {
					mapId: 2,
					xpath: '/ns1:Scores/ns1:Score',
					xmlDataType: 'double',
				},
				totalsRowFunction: 'sum',
				totalsRowDxfId: 8,
			},
		])
		expect(sheet?.tables[0]?.autoFilter).toEqual({
			ref: 'A1:B3',
			uid: '{FILTER-UID}',
			columns: [],
			sortState: {
				ref: 'A2:B3',
				conditions: [{ ref: 'B2:B3' }],
			},
		})
		expect(sheet?.tables[0]?.sortState).toEqual({
			ref: 'A2:B3',
			preservedAttributes: {
				'xmlns:xlrd2': 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2',
			},
			conditions: [{ ref: 'A2:A3' }],
		})
		expect(sheet?.tables[0]?.uid).toBe('{TABLE-UID}')
		expect(sheet?.tables[0]?.headerRowDxfId).toBe(5)
		expect(sheet?.tables[0]?.headerRowCellStyle).toBe('TableHeader')
		expect(sheet?.tables[0]?.altText).toBe('Scores')
		expect(sheet?.tables[0]?.altTextSummary).toBe('Student score table')
		expect(sheet?.tables[0]?.dataCellStyle).toBe('TableData')
		expect(sheet?.tables[0]?.dataDxfId).toBe(6)
		expect(sheet?.tables[0]?.tableBorderDxfId).toBe(9)
		expect(sheet?.tables[0]?.tableStyleInfo).toEqual({
			name: 'TableStyleMedium2',
			showFirstColumn: false,
			showLastColumn: false,
			showRowStripes: true,
			showColumnStripes: false,
		})
		expect(result.value.report.features.some((feature) => feature.feature === 'table')).toBe(true)
	})

	it('parses tables when sheet relationships use strict OOXML namespace (purl.oclc.org)', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>10</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
			'xl/tables/table1.xml': `<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="StrictTable" displayName="StrictTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0">
  <autoFilter ref="A1:B2"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Col1"/>
    <tableColumn id="2" name="Col2"/>
  </tableColumns>
</table>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.tables).toHaveLength(1)
		expect(sheet?.tables[0]?.name).toBe('StrictTable')
		expect(sheet?.tables[0]).toMatchObject({
			partPath: 'xl/tables/table1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
			contentTypeSource: 'override',
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			sourceRelationshipPart: 'xl/worksheets/_rels/sheet1.xml.rels',
			sourceRelationshipId: 'rId1',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			sourceRelationshipRawType: 'http://purl.oclc.org/ooxml/officeDocument/relationships/table',
			sourceRelationshipRawTarget: '../tables/table1.xml',
			sourceRelationshipResolvedTarget: 'xl/tables/table1.xml',
		})
		expect(sheet?.tables[0]?.ref).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } })
		expect(sheet?.tables[0]?.columns).toHaveLength(2)
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 42 })
	})

	it('parses worksheet autoFilter criteria and sort state', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:B10">
    <filterColumn colId="0">
      <filters blank="1">
        <filter val="Open"/>
        <filter val="Closed"/>
      </filters>
    </filterColumn>
    <sortState ref="A2:B10" caseSensitive="1">
      <sortCondition ref="B2:B10" descending="1"/>
    </sortState>
  </autoFilter>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		expect(result.value.workbook.sheets[0]?.autoFilter).toEqual({
			ref: 'A1:B10',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					blank: true,
					values: ['Open', 'Closed'],
				},
			],
			sortState: {
				ref: 'A2:B10',
				caseSensitive: true,
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		})
	})

	it('parses advanced worksheet autoFilter criteria', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:D10">
    <filterColumn colId="0"><dynamicFilter type="thisMonth" valIso="2026-03-01T00:00:00" maxValIso="2026-04-01T00:00:00"/></filterColumn>
    <filterColumn colId="1"><top10 top="0" percent="1" val="40" filterVal="25"/></filterColumn>
    <filterColumn colId="2"><colorFilter dxfId="3" cellColor="0"/></filterColumn>
    <filterColumn colId="3"><iconFilter iconSet="3TrafficLights1" iconId="1"/></filterColumn>
  </autoFilter>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		expect(result.value.workbook.sheets[0]?.autoFilter?.columns).toEqual([
			{
				colId: 0,
				kind: 'dynamicFilter',
				dynamicFilterType: 'thisMonth',
				dynamicFilterValIso: '2026-03-01T00:00:00',
				dynamicFilterMaxValIso: '2026-04-01T00:00:00',
			},
			{
				colId: 1,
				kind: 'top10',
				top: false,
				percent: true,
				val: 40,
				filterVal: 25,
			},
			{
				colId: 2,
				kind: 'colorFilter',
				dxfId: 3,
				cellColor: false,
			},
			{
				colId: 3,
				kind: 'iconFilter',
				iconSet: '3TrafficLights1',
				iconId: 1,
			},
		])
	})

	it('parses top-level worksheet sort state', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <sortState ref="D1:I2707">
    <sortCondition ref="D1"/>
  </sortState>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		expect(result.value.workbook.sheets[0]?.sortState).toEqual({
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1' }],
		})
	})

	it('reports workbook freshness signals from calc settings and calc chain', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcCompleted="0" forceFullCalc="1"/>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
			'xl/calcChain.xml': `<?xml version="1.0"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(
			result.value.report.features.some((feature) => feature.feature === 'formulaFreshness'),
		).toBe(true)
		const freshness = result.value.report.features.find(
			(feature) => feature.feature === 'formulaFreshness',
		)
		expect(freshness?.note).toContain('calculation not completed')
		expect(freshness?.note).toContain('forced full recalculation')
		expect(result.value.report.features.some((feature) => feature.feature === 'calcChain')).toBe(
			true,
		)
	})

	it('discovers pivot cache and pivot table inventory', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
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
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable11" cacheId="34"
  dataOnRows="1" compact="0" rowGrandTotals="0" dataCaption="Values" updatedVersion="7">
  <location ref="O17" firstHeaderRow="0" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"/>
  <pivotFields count="4">
    <pivotField axis="axisRow" showAll="0" sortType="ascending" compact="0" dragToPage="0"/>
    <pivotField axis="axisCol" includeNewItemsInFilter="1" itemPageCount="10"/>
    <pivotField axis="axisPage" hidden="1"/>
    <pivotField dataField="1" defaultSubtotal="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="1"/></colFields>
  <rowItems count="2"><i><x v="0"/></i><i t="grand"><x/></i></rowItems>
  <colItems count="2"><i><x v="1"/></i><i r="1" i="0"><x v="2"/></i></colItems>
  <pageFields count="1"><pageField fld="2" name="Region filter"/></pageFields>
  <dataFields count="1"><dataField fld="3" name="Sum of Sales" subtotal="sum" numFmtId="4" showDataAs="percent" baseField="0" baseItem="2"/></dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="1" showLastColumn="0"/>
</pivotTableDefinition>`,
			'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rIdRecords" recordCount="25" refreshOnLoad="1" enableRefresh="1" saveData="0"
  invalid="1" optimizeMemory="1" refreshedVersion="7" minRefreshableVersion="3"
  createdVersion="6" refreshedBy="Ascend" refreshedDate="45123.5">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D100" sheet="raw data"/>
  </cacheSource>
  <cacheFields count="4">
    <cacheField name="Region" databaseField="1">
      <sharedItems count="2"><s v="West"/><s v="East"/></sharedItems>
    </cacheField>
    <cacheField name="Quarter"/>
    <cacheField name="Channel"/>
    <cacheField name="Sales" numFmtId="4">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="5" maxValue="100"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
			'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
			'xl/pivotCache/pivotCacheRecords1.xml':
				'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.pivotTables).toEqual([
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'PivotSheet',
				name: 'PivotTable11',
				cacheId: 34,
				locationRef: 'O17',
				location: {
					ref: 'O17',
					firstHeaderRow: 0,
					firstDataRow: 1,
					firstDataCol: 1,
					rowPageCount: 1,
					colPageCount: 1,
				},
				options: {
					dataOnRows: true,
					compact: false,
					rowGrandTotals: false,
					updatedVersion: 7,
					dataCaption: 'Values',
				},
				style: {
					name: 'PivotStyleLight16',
					showRowHeaders: true,
					showColHeaders: true,
					showRowStripes: false,
					showColStripes: true,
					showLastColumn: false,
				},
				fields: [
					{
						index: 0,
						axis: 'axisRow',
						showAll: false,
						compact: false,
						dragToPage: false,
						sortType: 'ascending',
					},
					{ index: 1, axis: 'axisCol', includeNewItemsInFilter: true, itemPageCount: 10 },
					{ index: 2, axis: 'axisPage', hidden: true },
					{ index: 3, dataField: true, defaultSubtotal: false },
				],
				rowFields: [{ index: 0 }],
				columnFields: [{ index: 1 }],
				pageFields: [{ index: 2, name: 'Region filter' }],
				dataFields: [
					{
						fieldIndex: 3,
						name: 'Sum of Sales',
						subtotal: 'sum',
						numFmtId: 4,
						showDataAs: 'percent',
						baseField: 0,
						baseItem: 2,
					},
				],
				rowItems: [
					{ index: 0, fieldItems: [{ index: 0, item: 0 }] },
					{ index: 1, itemType: 'grand', fieldItems: [{ index: 0 }] },
				],
				columnItems: [
					{ index: 0, fieldItems: [{ index: 0, item: 1 }] },
					{
						index: 1,
						repeatedItemCount: 1,
						dataFieldIndex: 0,
						fieldItems: [{ index: 0, item: 2 }],
					},
				],
			},
		])
		expect(result.value.workbook.pivotCaches).toEqual([
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 34,
				relId: 'rIdPivotCache',
				recordCount: 25,
				refreshedVersion: 7,
				minRefreshableVersion: 3,
				createdVersion: 6,
				refreshedBy: 'Ascend',
				refreshedDate: 45123.5,
				refreshOnLoad: true,
				enableRefresh: true,
				invalid: true,
				saveData: false,
				optimizeMemory: true,
				sourceType: 'worksheet',
				sourceSheet: 'raw data',
				sourceRef: 'A1:D100',
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				records: {
					partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
					parsedCount: 0,
					preview: [],
					valueKindCounts: [],
				},
				fields: [
					{
						index: 0,
						name: 'Region',
						databaseField: true,
						sharedItemsInfo: { count: 2 },
						sharedItems: [
							{ index: 0, kind: 'string', value: 'West' },
							{ index: 1, kind: 'string', value: 'East' },
						],
					},
					{ index: 1, name: 'Quarter' },
					{ index: 2, name: 'Channel' },
					{
						index: 3,
						name: 'Sales',
						numFmtId: 4,
						sharedItemsInfo: {
							containsSemiMixedTypes: false,
							containsString: false,
							containsNumber: true,
							containsInteger: true,
							minValue: 5,
							maxValue: 100,
						},
					},
				],
			},
		])
	})

	it('discovers embedded chart ownership, type, title, and series source ranges', () => {
		const bytes = makeEmbeddedChartXlsx({ sheetName: 'Data' })

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.chartParts).toEqual([
			{
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Data',
				chartType: 'barChart',
				title: 'Revenue & margin',
				series: [
					{
						nameRef: 'Data!$B$1',
						categoryRef: 'Data!$A$2:$A$4',
						valueRef: 'Data!$B$2:$B$4',
					},
				],
			},
		])
	})

	it('discovers slicer and slicer cache inventory', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/slicerCaches/slicerCache1.xml': `<?xml version="1.0"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Product_Category" sourceName="Product_Category">
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="1332190931"><items count="3"><i x="0" s="1"/><i x="1"/><i x="2" nd="1"/></items></tabular></data>
</slicerCacheDefinition>`,
			'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Product_Category" cache="Slicer_Product_Category" caption="Product Category"/>
</slicers>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.slicerCaches).toEqual([
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				name: 'Slicer_Product_Category',
				sourceName: 'Product_Category',
				pivotCacheId: 1332190931,
				pivotTableNames: ['PivotTable1'],
				items: [{ index: 0, selected: true }, { index: 1 }, { index: 2, noData: true }],
			},
		])
		expect(result.value.workbook.slicers).toEqual([
			{
				partPath: 'xl/slicers/slicer1.xml',
				name: 'Product_Category',
				cacheName: 'Slicer_Product_Category',
				caption: 'Product Category',
			},
		])
	})

	it('parses workbook views, workbook properties, and external references', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1" filterPrivacy="1" codeName="Model"/>
  <bookViews>
    <workbookView activeTab="1" firstSheet="2" visibility="visible" tabRatio="600" minimized="1" showSheetTabs="0" windowWidth="16800" windowHeight="9000"/>
  </bookViews>
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
  <calcPr calcMode="manual" fullCalcOnLoad="1"/>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdExt"/></externalLink>`,
			'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMetadata" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
  <Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../sources/source.xlsx" TargetMode="External"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.workbookProperties).toEqual({
			date1904: true,
			filterPrivacy: true,
			codeName: 'Model',
		})
		expect(result.value.workbook.workbookViews).toEqual([
			{
				activeTab: 1,
				firstSheet: 2,
				visibility: 'visible',
				tabRatio: 600,
				extraAttributes: [
					{ name: 'minimized', value: '1' },
					{ name: 'showSheetTabs', value: '0' },
					{ name: 'windowWidth', value: '16800' },
					{ name: 'windowHeight', value: '9000' },
				],
			},
		])
		expect(result.value.workbook.externalReferences).toEqual(['xl/externalLinks/externalLink1.xml'])
		expect(result.value.workbook.externalReferenceDetails).toEqual([
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				relId: 'rId2',
				sourcePartPath: 'xl/workbook.xml',
				sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
				sourceRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
				sourceRelationshipRawType:
					'http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink',
				sourceRelationshipRawTarget: 'externalLinks/externalLink1.xml',
				sourceRelationshipResolvedTarget: 'xl/externalLinks/externalLink1.xml',
				externalBookRelId: 'rIdExt',
				externalLinkKind: 'externalBook',
				externalLinkRelId: 'rIdExt',
				linkRelId: 'rIdExt',
				linkRelationshipPart: 'xl/externalLinks/_rels/externalLink1.xml.rels',
				linkRelationshipKind: 'externalLinkPath',
				linkBindingStatus: 'externalBookRelId',
				linkRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
				linkRelationshipRawTarget: '../sources/source.xlsx',
				target: '../sources/source.xlsx',
				targetMode: 'External',
			},
		])
	})

	it('uses externalBook r:id to resolve the external workbook path relationship', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdChosen"/></externalLink>`,
			'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOther" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../sources/wrong.xlsx" TargetMode="External"/>
  <Relationship Id="rIdChosen" Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup" Target="personal.xls" TargetMode="External"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.externalReferenceDetails).toEqual([
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				relId: 'rId2',
				sourcePartPath: 'xl/workbook.xml',
				sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
				sourceRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
				sourceRelationshipRawTarget: 'externalLinks/externalLink1.xml',
				sourceRelationshipResolvedTarget: 'xl/externalLinks/externalLink1.xml',
				externalBookRelId: 'rIdChosen',
				externalLinkKind: 'externalBook',
				externalLinkRelId: 'rIdChosen',
				linkRelId: 'rIdChosen',
				linkRelationshipPart: 'xl/externalLinks/_rels/externalLink1.xml.rels',
				linkRelationshipKind: 'xlStartup',
				linkBindingStatus: 'externalBookRelId',
				linkRelationshipType:
					'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup',
				linkRelationshipRawTarget: 'personal.xls',
				target: 'personal.xls',
				targetMode: 'External',
			},
		])
	})

	it('marks external link path fallback when externalBook r:id does not bind', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdMissing"/></externalLink>`,
			'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPath" Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary" Target="library.xlsx" TargetMode="External"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.externalReferenceDetails).toEqual([
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				relId: 'rId2',
				sourcePartPath: 'xl/workbook.xml',
				sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
				sourceRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
				sourceRelationshipRawTarget: 'externalLinks/externalLink1.xml',
				sourceRelationshipResolvedTarget: 'xl/externalLinks/externalLink1.xml',
				externalBookRelId: 'rIdMissing',
				externalLinkKind: 'externalBook',
				externalLinkRelId: 'rIdMissing',
				linkRelId: 'rIdPath',
				linkRelationshipPart: 'xl/externalLinks/_rels/externalLink1.xml.rels',
				linkRelationshipKind: 'xlLibrary',
				linkBindingStatus: 'fallbackPathRelationship',
				linkRelationshipType:
					'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary',
				linkRelationshipRawTarget: 'library.xlsx',
				target: 'library.xlsx',
				targetMode: 'External',
			},
		])
	})

	it('does not bind externalBook r:id to a same-id non-path relationship', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdChosen"/></externalLink>`,
			'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChosen" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
  <Relationship Id="rIdPath" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../sources/actual.xlsx" TargetMode="External"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.externalReferenceDetails).toEqual([
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				relId: 'rId2',
				sourcePartPath: 'xl/workbook.xml',
				sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
				sourceRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
				sourceRelationshipRawTarget: 'externalLinks/externalLink1.xml',
				sourceRelationshipResolvedTarget: 'xl/externalLinks/externalLink1.xml',
				externalBookRelId: 'rIdChosen',
				externalLinkKind: 'externalBook',
				externalLinkRelId: 'rIdChosen',
				linkRelId: 'rIdPath',
				linkRelationshipPart: 'xl/externalLinks/_rels/externalLink1.xml.rels',
				linkRelationshipKind: 'externalLinkPath',
				linkBindingStatus: 'fallbackPathRelationship',
				linkRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
				linkRelationshipRawTarget: '../sources/actual.xlsx',
				target: '../sources/actual.xlsx',
				targetMode: 'External',
			},
		])
	})

	it('parses workbook and sheet protection metadata', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookProtection lockStructure="1" workbookPassword="ABCD" workbookAlgorithmName="SHA-512" workbookSpinCount="100000"/>
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <sheetProtection sheet="1" objects="1" scenarios="1" password="1234" sort="0" autoFilter="0" selectUnlockedCells="1"/>
  <protectedRanges>
    <protectedRange name="Editable" sqref="C:C" password="ABCD"/>
  </protectedRanges>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.workbookProtection).toEqual({
			lockStructure: true,
			workbookPassword: 'ABCD',
			workbookAlgorithmName: 'SHA-512',
			workbookSpinCount: 100000,
		})
		expect(result.value.workbook.sheets[0]?.protection).toEqual({
			sheet: true,
			objects: true,
			scenarios: true,
			password: '1234',
			sort: false,
			autoFilter: false,
			selectUnlockedCells: true,
		})
		expect(result.value.workbook.sheets[0]?.protectedRanges).toEqual([
			{
				name: 'Editable',
				sqref: 'C:C',
				password: 'ABCD',
			},
		])
	})

	it('parses and preserves workbook theme metadata', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Twist">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`,
		})

		const result = readXlsx(bytes, { mode: 'full' })
		expectOk(result)

		expect(result.value.workbook.themeMetadata).toEqual({
			name: 'Office Twist',
			colorSchemeName: 'Office',
			colorCount: 12,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		})
		expect(result.value.workbook.preservedTheme).toEqual({
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
		})
	})

	it('parses worksheet layout metadata and hyperlinks', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetPr codeName="SheetCode1" filterMode="1" enableFormatConditionsCalculation="0">
    <pageSetUpPr fitToPage="1" autoPageBreaks="0"/>
  </sheetPr>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="2" ySplit="1" state="frozen" topLeftCell="C2" activePane="bottomRight"/>
      <selection pane="topRight" activeCell="C1" sqref="C1"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
      <selection pane="bottomRight" activeCell="C2" activeCellId="0" sqref="C2:D4"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="12" defaultRowHeight="14.5" zeroHeight="1" x14ac:dyDescent="0.25"/>
  <cols>
    <col min="1" max="2" width="18.5" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" spans="1:2" ht="24" customHeight="0" s="2" customFormat="1" thickTop="1" thickBot="1" x14ac:dyDescent="0.3">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:B10"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup orientation="landscape" fitToWidth="1" fitToHeight="2" firstPageNumber="2" copies="3" horizontalDpi="600" verticalDpi="600" pageOrder="overThenDown" cellComments="asDisplayed" errors="dash" blackAndWhite="1" draft="1" useFirstPageNumber="1" usePrinterDefaults="0" r:id="rIdPrinter"/>
  <printOptions gridLines="1" gridLinesSet="1" headings="1"/>
  <headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="1"><oddHeader>&amp;LTest</oddHeader><oddFooter>&amp;R1</oddFooter></headerFooter>
  <phoneticPr fontId="1" type="noConversion" alignment="center"/>
  <rowBreaks count="1" manualBreakCount="1"><brk id="5" min="0" max="16383" man="1"/></rowBreaks>
  <colBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="1048575" man="1"/></colBreaks>
  <ignoredErrors><ignoredError sqref="A1:B2" numberStoredAsText="1"/></ignoredErrors>
  <hyperlinks><hyperlink ref="A1" r:id="rIdHyper" display="Docs" tooltip="Open docs"/></hyperlinks>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHyper" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.codeName).toBe('SheetCode1')
		expect(sheet?.filterMode).toBe(true)
		expect(sheet?.enableFormatConditionsCalculation).toBe(false)
		expect(sheet?.pageSetupPr).toEqual({ fitToPage: true, autoPageBreaks: false })
		expect(sheet?.frozenRows).toBe(1)
		expect(sheet?.frozenCols).toBe(2)
		expect(sheet?.preservedPaneAttributes).toEqual({
			xSplit: '2',
			ySplit: '1',
			state: 'frozen',
			topLeftCell: 'C2',
			activePane: 'bottomRight',
		})
		expect(sheet?.preservedSheetViewSelections).toEqual([
			{ pane: 'topRight', activeCell: 'C1', sqref: 'C1' },
			{ pane: 'bottomLeft', activeCell: 'A2', sqref: 'A2' },
			{ pane: 'bottomRight', activeCell: 'C2', activeCellId: '0', sqref: 'C2:D4' },
		])
		expect(sheet?.sheetFormatPr).toEqual({
			baseColWidth: 12,
			defaultRowHeight: 14.5,
			zeroHeight: true,
			dyDescent: 0.25,
		})
		expect(sheet?.colWidths.get(0)).toBe(18.5)
		expect(sheet?.colWidths.get(1)).toBe(18.5)
		expect(sheet?.rowHeights.get(0)).toBe(24)
		expect(sheet?.rowDefs.get(0)).toEqual({
			spans: '1:2',
			style: 2,
			customFormat: true,
			customHeight: false,
			thickTop: true,
			thickBot: true,
			dyDescent: 0.3,
		})
		expect(sheet?.autoFilter).toEqual({
			ref: 'A1:B10',
			columns: [],
		})
		expect(sheet?.pageMargins).toEqual({
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		})
		expect(sheet?.pageSetup).toEqual({
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 2,
			firstPageNumber: 2,
			copies: 3,
			horizontalDpi: 600,
			verticalDpi: 600,
			pageOrder: 'overThenDown',
			cellComments: 'asDisplayed',
			errors: 'dash',
			blackAndWhite: true,
			draft: true,
			useFirstPageNumber: true,
			usePrinterDefaults: false,
			printerSettingsRelId: 'rIdPrinter',
		})
		expect(sheet?.printOptions).toEqual({
			gridLines: true,
			gridLinesSet: true,
			headings: true,
			horizontalCentered: undefined,
			verticalCentered: undefined,
		})
		expect(sheet?.headerFooter).toEqual({
			differentOddEven: true,
			differentFirst: true,
			scaleWithDoc: false,
			alignWithMargins: true,
			oddHeader: '&LTest',
			oddFooter: '&R1',
			evenHeader: undefined,
			evenFooter: undefined,
			firstHeader: undefined,
			firstFooter: undefined,
		})
		expect(sheet?.phoneticPr).toEqual({
			fontId: 1,
			type: 'noConversion',
			alignment: 'center',
		})
		expect(sheet?.rowBreaks).toEqual([{ id: 5, min: 0, max: 16383, man: true }])
		expect(sheet?.colBreaks).toEqual([{ id: 2, min: 0, max: 1048575, man: true }])
		expect(sheet?.ignoredErrors).toEqual([{ sqref: 'A1:B2', numberStoredAsText: true }])
		expect(sheet?.hyperlinks.get('A1')).toEqual({
			target: 'https://example.com/docs',
			location: undefined,
			display: 'Docs',
			tooltip: 'Open docs',
		})
	})

	it('parses full sheetView attributes (zoomScale, zoomScaleNormal, zoomScaleSheetLayoutView, showGridLines, showFormulas, rightToLeft, tabSelected, view, topLeftCell)', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0" zoomScale="75" zoomScaleNormal="100" zoomScaleSheetLayoutView="214" showGridLines="0" showFormulas="1" rightToLeft="1" tabSelected="1" view="pageBreakPreview" topLeftCell="E1">
      <pane ySplit="2" xSplit="1" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.frozenRows).toBe(2)
		expect(sheet?.frozenCols).toBe(1)
		expect(sheet?.sheetView).toEqual({
			zoomScale: 75,
			zoomScaleNormal: 100,
			zoomScaleSheetLayoutView: 214,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageBreakPreview',
			topLeftCell: 'E1',
		})
	})

	it('parses drawing and legacyDrawing references on worksheets', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDraw"/>
  <legacyDrawing r:id="rIdLegacy"/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rIdLegacy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)
		expect(result.value.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: true,
			hasLegacyDrawing: true,
		})
	})

	it('discovers image relationships hanging off worksheet drawings', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDraw"/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
			'xl/drawings/drawing1.xml': `<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row><xdr:colOff>10</xdr:colOff><xdr:rowOff>20</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row><xdr:colOff>30</xdr:colOff><xdr:rowOff>40</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="Image 1" descr="Hero"/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rIdImg"/></xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
			'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
			'xl/media/image1.png': 'fakepng',
		})

		const result = readXlsx(bytes)
		expectOk(result)
		expect(result.value.workbook.sheets[0]?.imageRefs).toEqual([
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				relId: 'rIdImg',
				targetPath: 'xl/media/image1.png',
				name: 'Image 1',
				description: 'Hero',
				anchor: {
					kind: 'twoCell',
					editAs: 'oneCell',
					from: { col: 1, row: 2, colOff: 10, rowOff: 20 },
					to: { col: 3, row: 4, colOff: 30, rowOff: 40 },
				},
			},
		])
	})

	it('parses classic comments from worksheet relationships', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <legacyDrawing r:id="rIdLegacy"/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rIdLegacy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`,
			'xl/comments1.xml': `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Ada</author></authors>
  <commentList>
    <comment ref="B2" authorId="0"><text><t>Hello</t></text></comment>
  </commentList>
</comments>`,
			'xl/drawings/vmlDrawing1.vml': '<xml xmlns:v="urn:schemas-microsoft-com:vml"/>',
		})

		const result = readXlsx(bytes)
		expectOk(result)
		expect(result.value.workbook.sheets[0]?.comments.get('B2')).toEqual({
			text: 'Hello',
			author: 'Ada',
		})
	})

	it('values mode can opt into rich sheet metadata without hydrating formulas', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames><definedName name="FeatureRange">Data!$A$1:$B$2</definedName></definedNames>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>7</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1"><cfRule type="cellIs" operator="greaterThan" priority="1"><formula>0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="B2"><formula1>"A,B"</formula1></dataValidation></dataValidations>
  <hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink" display="Ascend"/></hyperlinks>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rIdHyperlink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/ascend" TargetMode="External"/>
</Relationships>`,
			'xl/comments1.xml': `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Ada</author></authors>
  <commentList><comment ref="B2" authorId="0"><text><t>Review</t></text></comment></commentList>
</comments>`,
		})

		const valuesOnly = readXlsx(bytes, { mode: 'values' })
		expectOk(valuesOnly)
		expect(valuesOnly.value.loadInfo.richSheetMetadataHydrated).toBe(false)
		expect(valuesOnly.value.workbook.sheets[0]?.comments.size).toBe(0)
		expect(valuesOnly.value.workbook.sheets[0]?.hyperlinks.size).toBe(0)
		expect(valuesOnly.value.workbook.sheets[0]?.dataValidations).toHaveLength(0)
		expect(valuesOnly.value.workbook.sheets[0]?.conditionalFormats).toHaveLength(0)
		expect(valuesOnly.value.workbook.definedNames.get('FeatureRange')).toBe('Data!$A$1:$B$2')

		const valuesWithMetadata = readXlsx(bytes, { mode: 'values', richMetadata: true })
		expectOk(valuesWithMetadata)
		const sheet = valuesWithMetadata.value.workbook.sheets[0]
		expect(valuesWithMetadata.value.loadInfo.richSheetMetadataHydrated).toBe(true)
		expect(valuesWithMetadata.value.loadInfo.isPartial).toBe(true)
		expect(sheet?.cells.get(0, 0)?.formula).toBeNull()
		expect(sheet?.comments.get('B2')).toEqual({ text: 'Review', author: 'Ada' })
		expect(sheet?.hyperlinks.get('A1')).toEqual({
			target: 'https://example.com/ascend',
			display: 'Ascend',
		})
		expect(sheet?.dataValidations).toEqual([{ sqref: 'B2', type: 'list', formula1: '"A,B"' }])
		expect(sheet?.conditionalFormats).toEqual([
			{
				sqref: 'A1',
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						priority: 1,
						formulas: ['0'],
					},
				],
			},
		])
	})

	it('captures style metadata richness for read-time inspection', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>
  <fonts count="2"><font/><font><b/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="1" borderId="0"/></cellXfs>
  <dxfs count="2"><dxf><font><b/></font></dxf><dxf><fill><patternFill patternType="solid"/></fill></dxf></dxfs>
  <tableStyles count="1" defaultTableStyle="TableStyleMedium2"><tableStyle name="TableStyleMedium2"/></tableStyles>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.styleMetadata).toEqual({
			numFmtCount: 1,
			fontCount: 2,
			fillCount: 2,
			borderCount: 1,
			cellXfCount: 2,
			dxfCount: 2,
			tableStyleCount: 1,
		})
	})

	it('parses gradient fills into the style model', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <gradientFill type="linear" degree="45">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color theme="1" tint="0.25"/></stop>
      </gradientFill>
    </fill>
  </fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0"/>
  </cellXfs>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const styleId = result.value.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0
		expect(result.value.workbook.styles.get(styleId)?.fill).toEqual({
			gradient: {
				type: 'linear',
				degree: 45,
				stops: [
					{ position: 0, color: { kind: 'rgb', rgb: 'FFFF0000' } },
					{ position: 1, color: { kind: 'theme', theme: 1, tint: 0.25 } },
				],
			},
		})
	})

	it('parses conditional formatting and data validations', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font/><font><b/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <dxfs count="1"><dxf><font><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></dxf></dxfs>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>4</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A10" pivot="1">
    <cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1" stopIfTrue="1">
      <formula>3</formula>
    </cfRule>
    <cfRule type="aboveAverage" priority="2" aboveAverage="0" stdDev="2"/>
    <cfRule type="containsText" priority="3" operator="containsText" text="Grain"/>
  </conditionalFormatting>
  <dataValidations count="1" disablePrompts="1" xWindow="220" yWindow="120">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" sqref="B2:B4" xr:uid="{CAFD7DE3-F94F-4BD6-B5E6-7E794FD6EC31}">
      <formula1>"Q1,Q2,Q3"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A10',
				pivot: true,
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						dxfId: 0,
						priority: 1,
						stopIfTrue: true,
						formulas: ['3'],
						style: {
							font: { bold: true },
							fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
						},
					},
					{
						type: 'aboveAverage',
						priority: 2,
						formulas: [],
						aboveAverage: false,
						stdDev: 2,
					},
					{
						type: 'containsText',
						operator: 'containsText',
						priority: 3,
						formulas: [],
						text: 'Grain',
					},
				],
			},
		])
		expect(sheet?.dataValidations).toEqual([
			{
				sqref: 'B2:B4',
				uid: '{CAFD7DE3-F94F-4BD6-B5E6-7E794FD6EC31}',
				type: 'list',
				allowBlank: true,
				showInputMessage: true,
				formula1: '"Q1,Q2,Q3"',
			},
		])
		expect(sheet?.dataValidationSettings).toEqual({
			disablePrompts: true,
			xWindow: 220,
			yWindow: 120,
		})
		expect(result.value.workbook.differentialStyles).toEqual([
			{
				font: { bold: true },
				fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
			},
		])
	})

	it('parses extension-list data validations', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
  xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <extLst>
    <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
      <x14:dataValidations count="1">
        <x14:dataValidation type="list" showInputMessage="1" showErrorMessage="1" xr:uid="{DV-UID}" customFlag="1">
          <x14:formula1><xm:f>Lookup!$E$2:$E$123</xm:f></x14:formula1>
          <x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>
          <xm:sqref>E8:E11</xm:sqref>
        </x14:dataValidation>
      </x14:dataValidations>
    </ext>
  </extLst>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.dataValidations).toEqual([
			{
				sqref: 'E8:E11',
				source: 'x14',
				type: 'list',
				showInputMessage: true,
				showErrorMessage: true,
				formula1: 'Lookup!$E$2:$E$123',
				uid: '{DV-UID}',
			},
		])
		expect(sheet?.x14DataValidations).toEqual([
			{
				index: 0,
				sqref: 'E8:E11',
				type: 'list',
				showInputMessage: true,
				showErrorMessage: true,
				formula1: 'Lookup!$E$2:$E$123',
				preservedAttributes: {
					'xr:uid': '{DV-UID}',
					customFlag: '1',
				},
				preservedChildXml: ['<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>'],
			},
		])
	})

	it('parses regular conditional formatting extension payloads', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <sheetData/>
  <conditionalFormatting sqref="A1:A5">
    <cfRule type="dataBar" priority="1" activePresent="1" xr:uid="{REGULAR-CF-UID}">
      <dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>
      <extLst><ext uri="{regular-cf-extension}"><x14ac:metadata flag="1"/></ext></extLst>
    </cfRule>
  </conditionalFormatting>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.conditionalFormats[0]?.rules[0]).toMatchObject({
			type: 'dataBar',
			priority: 1,
			preservedRuleAttributes: {
				activePresent: '1',
				'xr:uid': '{REGULAR-CF-UID}',
			},
			preservedRuleChildXml: [
				'<extLst><ext uri="{regular-cf-extension}"><x14ac:metadata flag="1"/></ext></extLst>',
			],
			dataBar: {
				cfvo: [{ type: 'min' }, { type: 'max' }],
				color: { rgb: 'FF638EC6' },
			},
		})
	})

	it('preserves document-order indexes for mixed x14 data-validation elements', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
      <x14:dataValidations count="2">
        <x14:dataValidation type="whole" sqref="A1"/>
        <x14:dataValidation type="list">
          <x14:formula1><xm:f>Lookup!$A$1:$A$2</xm:f></x14:formula1>
          <xm:sqref>B1</xm:sqref>
        </x14:dataValidation>
      </x14:dataValidations>
    </ext>
  </extLst>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.x14DataValidations).toEqual([
			{ index: 0, sqref: 'A1', type: 'whole' },
			{ index: 1, sqref: 'B1', type: 'list', formula1: 'Lookup!$A$1:$A$2' },
		])
	})

	it('parses extension-list conditional formatting rule payloads', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
  xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <sheetData/>
  <extLst>
    <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
      <x14:conditionalFormattings>
        <x14:conditionalFormatting>
          <x14:cfRule type="dataBar" priority="4" activePresent="1" xr:uid="{CF-UID}">
            <x14:dataBar minLength="3" maxLength="97" border="1" showValue="0" gradient="0" direction="rightToLeft" axisPosition="middle" negativeBarColorSameAsPositive="0" negativeBarBorderColorSameAsPositive="1"><x14:cfvo type="formula"><xm:f>A1</xm:f></x14:cfvo><x14:fillColor rgb="FF638EC6"/><x14:borderColor theme="4" tint="-0.25"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor auto="1"/></x14:dataBar>
            <x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>
          </x14:cfRule>
          <xm:sqref>A1:A5</xm:sqref>
        </x14:conditionalFormatting>
      </x14:conditionalFormattings>
    </ext>
  </extLst>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.x14ConditionalFormats).toEqual([
			{
				index: 0,
				sqref: 'A1:A5',
				formulas: [],
				type: 'dataBar',
				priority: 4,
				preservedRuleAttributes: {
					activePresent: '1',
					'xr:uid': '{CF-UID}',
				},
				preservedRuleChildXml: [
					'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
				],
				dataBar: {
					cfvo: [{ type: 'formula', value: 'A1' }],
					minLength: 3,
					maxLength: 97,
					border: true,
					showValue: false,
					gradient: false,
					direction: 'rightToLeft',
					axisPosition: 'middle',
					negativeBarColorSameAsPositive: false,
					negativeBarBorderColorSameAsPositive: true,
					fillColor: { rgb: 'FF638EC6' },
					borderColor: { theme: 4, tint: -0.25 },
					negativeFillColor: { rgb: 'FFFF0000' },
					axisColor: { auto: true },
				},
			},
		])
	})

	it('collapses equivalent normal and extension-list data validations', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData/>
  <dataValidations count="1">
    <dataValidation type="list" showInputMessage="1" showErrorMessage="1" sqref="E8:E11">
      <formula1>Lookup!$E$2:$E$123</formula1>
    </dataValidation>
  </dataValidations>
  <extLst>
    <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
      <x14:dataValidations count="1">
        <x14:dataValidation type="list" showInputMessage="1" showErrorMessage="1">
          <x14:formula1><xm:f>Lookup!$E$2:$E$123</xm:f></x14:formula1>
          <xm:sqref>E8:E11</xm:sqref>
        </x14:dataValidation>
      </x14:dataValidations>
    </ext>
  </extLst>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.dataValidations).toHaveLength(1)
	})

	it('supports metadata-only reads without parsing sheet cells', () => {
		const result = readXlsx(minimalXlsx(), { mode: 'metadata-only' })
		expectOk(result)

		expect(result.value.workbook.sheets).toHaveLength(1)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.name).toBe('Data')
		expect(sheet?.cells.cellCount()).toBe(0)
		expect(result.value.report.features.some((feature) => feature.feature === 'partialLoad')).toBe(
			true,
		)
	})

	it('supports values mode with hydrated cells but without formula fidelity', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>A2+1</f><v>7</v></c>
      <c r="A2"><v>6</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 7 })
		expect(sheet?.cells.get(0, 0)?.formula).toBeNull()
		expect(result.value.loadInfo.mode).toBe('values')
		expect(result.value.loadInfo.cellsHydrated).toBe(true)
		expect(result.value.loadInfo.isPartial).toBe(true)
		expect(result.value.report.features.some((feature) => feature.feature === 'partialLoad')).toBe(
			true,
		)
	})

	it('values mode returns cell values but never formulas', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>100</v></c>
      <c r="B1"><f>SUM(A:A)</f></c>
      <c r="C1"><f>A1*2</f><v>200</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 100 })
		expect(sheet?.cells.get(0, 0)?.formula).toBeNull()

		expect(sheet?.cells.get(0, 1)?.value).toEqual(EMPTY)
		expect(sheet?.cells.get(0, 1)?.formula).toBeNull()

		expect(sheet?.cells.get(0, 2)?.value).toEqual({ kind: 'number', value: 200 })
		expect(sheet?.cells.get(0, 2)?.formula).toBeNull()
	})

	it('values mode reads sequential A1-style refs beyond Z without shifting columns', () => {
		const cells = Array.from(
			{ length: 703 },
			(_, col) => `<c r="${indexToColumn(col)}1"><v>${col}</v></c>`,
		).join('')
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1">${cells}</row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 0 })
		expect(sheet?.cells.get(0, 25)?.value).toEqual({ kind: 'number', value: 25 })
		expect(sheet?.cells.get(0, 26)?.value).toEqual({ kind: 'number', value: 26 })
		expect(sheet?.cells.get(0, 701)?.value).toEqual({ kind: 'number', value: 701 })
		expect(sheet?.cells.get(0, 702)?.value).toEqual({ kind: 'number', value: 702 })
	})

	it('full mode preserves simple numeric cell references and style indexes', () => {
		const cells = [
			'<c r="A1"><v>7</v></c>',
			'<c r="AA1" s="1"><v>0.25</v></c>',
			'<c r="AAA1"><v>-3.5</v></c>',
		].join('')
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1">${cells}</row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const workbook = result.value.workbook
		const sheet = workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual(numberValue(7))
		const styled = sheet?.cells.get(0, 26)
		expect(styled?.value).toEqual(numberValue(0.25))
		expect(workbook.styles.get(styled?.styleId ?? S0)?.numberFormat).toBe('0.0%')
		expect(sheet?.cells.get(0, 702)?.value).toEqual(numberValue(-3.5))
	})

	it('full mode preserves simple shared-string cell references', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="AA1" t="s"><v>1</v></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual(stringValue('Hello'))
		expect(sheet?.cells.get(0, 26)?.value).toEqual(stringValue('World'))
	})

	it('full mode preserves simple inline-string cell references', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Hello &amp; fast</t></is></c><c r="AA1" t="inlineStr"><is><t>World</t></is></c><c r="AB1"><v>7</v></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual(stringValue('Hello & fast'))
		expect(sheet?.cells.get(0, 26)?.value).toEqual(stringValue('World'))
		expect(sheet?.cells.get(0, 27)?.value).toEqual(numberValue(7))
	})

	it('full mode infers omitted shared-string cell references in dense rows', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c><v>7</v></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual(stringValue('Hello'))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(stringValue('World'))
		expect(sheet?.cells.get(0, 2)?.value).toEqual(numberValue(7))
	})

	it('full mode preserves simple boolean cell references', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="b"><v>1</v></c><c r="AA1" t="b"><v>0</v></c><c r="AB1" t="b"><f>A1</f><v>1</v></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual(booleanValue(true))
		expect(sheet?.cells.get(0, 26)?.value).toEqual(booleanValue(false))
		expect(sheet?.cells.get(0, 27)?.value).toEqual(booleanValue(true))
		expect(sheet?.cells.get(0, 27)?.formula).toBe('A1')
	})

	it('full mode preserves simple formula cells without cached values', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>A1*2</f></c><c r="AA1"><f>A1&lt;10</f></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 1)?.value).toEqual(EMPTY)
		expect(sheet?.cells.get(0, 1)?.formula).toBe('A1*2')
		expect(sheet?.cells.get(0, 26)?.value).toEqual(EMPTY)
		expect(sheet?.cells.get(0, 26)?.formula).toBe('A1<10')
	})

	it('values mode reads simple inline strings without full cell XML parsing', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>hello &amp; goodbye</t></is></c>
      <c r="B1" t="inlineStr"><f>A1</f><is><t>cached</t></is></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'hello & goodbye',
		})
		expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'cached' })
		expect(sheet?.cells.get(0, 1)?.formula).toBeNull()
	})

	it('values-only byte parser preserves supported direct cell semantics', () => {
		const sharedStrings = parseSharedStrings(
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Shared &amp; decoded</t></si>
</sst>`,
		)
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:L2"/>
  <sheetData>
    <row r="1" ht="27" customHeight="1" hidden="1" collapsed="1" outlineLevel="2">
      <c r="A1"><v>42.5</v></c>
      <c r="B1" t="s"><v>0</v></c>
      <c r="C1" t="inlineStr"><is><t>inline &amp; entity</t></is></c>
      <c r="D1" t="b"><v>0</v></c>
      <c r="E1" t="e"><v>#N/A</v></c>
      <c r="F1" t="str"><v>plain &lt;text&gt;</v></c>
      <c r="G1"><f>SUM(A1:A1)</f><v>43</v></c>
      <c r="H1"><f>SUM(A1:A1)</f></c>
      <c r="I1" s="1"><v>45292</v></c>
      <c s="1" r="J1"><v>45293</v></c>
      <c r="K1" s="1"><f>I1+2</f><v>45294</v></c>
      <c r="L1" t="n" s="1"><v>45295</v></c>
    </row>
    <row r="2"><c r="A2"><v>99</v></c></row>
  </sheetData>
</worksheet>`
		const bytes = new TextEncoder().encode(xml)

		const sheet = parseSheetValuesOnlyBytes('Sheet1', bytes, {
			sharedStrings,
			styleIds: [S0],
			isDateFormat: [false, true],
			hasDateStyles: true,
			valuesOnly: true,
		})

		expect(sheet).not.toBeNull()
		expect(sheet?.cells.get(0, 0)?.value).toEqual(numberValue(42.5))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(stringValue('Shared & decoded'))
		expect(sheet?.cells.get(0, 2)?.value).toEqual(stringValue('inline & entity'))
		expect(sheet?.cells.get(0, 3)?.value).toEqual(booleanValue(false))
		expect(sheet?.cells.get(0, 4)?.value).toEqual({ kind: 'error', value: '#N/A' })
		expect(sheet?.cells.get(0, 5)?.value).toEqual(stringValue('plain <text>'))
		expect(sheet?.cells.get(0, 6)?.value).toEqual(numberValue(43))
		expect(sheet?.cells.get(0, 6)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 7)?.value).toEqual(EMPTY)
		expect(sheet?.cells.get(0, 7)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 8)?.value).toEqual(dateValue(45292))
		expect(sheet?.cells.get(0, 9)?.value).toEqual(dateValue(45293))
		expect(sheet?.cells.get(0, 10)?.value).toEqual(dateValue(45294))
		expect(sheet?.cells.get(0, 10)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 11)?.value).toEqual(dateValue(45295))
		expect(sheet?.cells.get(1, 0)?.value).toEqual(numberValue(99))
		expect(sheet?.rowHeights.get(0)).toBe(27)
		expect(sheet?.rowDefs.get(0)).toEqual({
			customHeight: true,
			hidden: true,
			collapsed: true,
			outlineLevel: 2,
		})

		const cappedSheet = parseSheetValuesOnlyBytes('Sheet1', bytes, {
			sharedStrings,
			styleIds: [S0],
			isDateFormat: [false, true],
			hasDateStyles: true,
			valuesOnly: true,
			maxRows: 1,
		})
		expect(cappedSheet?.cells.get(0, 8)?.value).toEqual(dateValue(45292))
		expect(cappedSheet?.cells.get(1, 0)).toBeUndefined()
	})

	it('values-only byte parser accepts XML single-quoted worksheet attributes', () => {
		const sharedStrings = parseSharedStrings(
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Shared &amp; decoded</t></si>
</sst>`,
		)
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref='A1:I2'/>
  <sheetData>
    <row r='1' ht='27' customHeight='1' hidden='1' collapsed='1' outlineLevel='2'>
      <c r='A1'><v>42.5</v></c>
      <c r='B1' t='s'><v>0</v></c>
      <c r='C1' t='inlineStr'><is><t>inline &amp; entity</t></is></c>
      <c r='D1' t='b'><v>0</v></c>
      <c r='E1' t='e'><v>#N/A</v></c>
      <c r='F1' t='str'><v>plain &lt;text&gt;</v></c>
      <c r='G1'><f>SUM(A1:A1)</f><v>43</v></c>
      <c r='H1'><f>SUM(A1:A1)</f></c>
      <c r='I1' s='1'><v>45292</v></c>
    </row>
    <row r='2'><c r='A2'><v>99</v></c></row>
  </sheetData>
</worksheet>`

		const sheet = parseSheetValuesOnlyBytes('Sheet1', new TextEncoder().encode(xml), {
			sharedStrings,
			styleIds: [S0],
			isDateFormat: [false, true],
			hasDateStyles: true,
			valuesOnly: true,
		})

		expect(sheet).not.toBeNull()
		expect(sheet?.preservedDimensionRef).toBe('A1:I2')
		expect(sheet?.cells.get(0, 0)?.value).toEqual(numberValue(42.5))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(stringValue('Shared & decoded'))
		expect(sheet?.cells.get(0, 2)?.value).toEqual(stringValue('inline & entity'))
		expect(sheet?.cells.get(0, 3)?.value).toEqual(booleanValue(false))
		expect(sheet?.cells.get(0, 4)?.value).toEqual({ kind: 'error', value: '#N/A' })
		expect(sheet?.cells.get(0, 5)?.value).toEqual(stringValue('plain <text>'))
		expect(sheet?.cells.get(0, 6)?.value).toEqual(numberValue(43))
		expect(sheet?.cells.get(0, 6)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 7)?.value).toEqual(EMPTY)
		expect(sheet?.cells.get(0, 7)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 8)?.value).toEqual(dateValue(45292))
		expect(sheet?.cells.get(1, 0)?.value).toEqual(numberValue(99))
		expect(sheet?.rowHeights.get(0)).toBe(27)
		expect(sheet?.rowDefs.get(0)).toEqual({
			customHeight: true,
			hidden: true,
			collapsed: true,
			outlineLevel: 2,
		})
	})

	it('values-only byte parser rejects prefixed sheetData children for XML fallback', () => {
		const xml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c></x:row>
  </sheetData>
</worksheet>`

		const sheet = parseSheetValuesOnlyBytes('Sheet1', new TextEncoder().encode(xml), {
			sharedStrings: emptySharedStrings(),
			styleIds: [S0],
			isDateFormat: [false],
			hasDateStyles: false,
			valuesOnly: true,
		})

		expect(sheet).toBeNull()
	})

	it('values-only byte parser hydrates canonical typed rows without generic cell parsing', () => {
		const sharedStrings = parseSharedStrings(
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Shared &amp; typed</t></si>
</sst>`,
		)
		const sequentialWideCells = Array.from({ length: 28 }, (_, col) => {
			const ref = `${indexToColumn(col)}2`
			return `<c r="${ref}"><v>${col + 1}</v></c>`
		}).join('')
		const xml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:AB3"/>
  <sheetData>
    <row r="1"><c r="A1"><v>7</v></c><c r="B1" t="inlineStr"><is><t>fast &amp; safe</t></is></c><c r="C1" t="s"><v>0</v></c><c r="D1" t="b"><v>1</v></c><c r="AA1"><v>-3.5</v></c></row>
    <row r="2">${sequentialWideCells}</row>
    <row r="3"><c><v>-12</v></c><c><v>1.25E3</v></c><c><v>12345678901234567</v></c></row>
  </sheetData>
</worksheet>`
		const sheet = parseSheetValuesOnlyBytes('Sheet1', new TextEncoder().encode(xml), {
			sharedStrings,
			styleIds: [S0],
			isDateFormat: [false],
			hasDateStyles: false,
			valuesOnly: true,
		})

		expect(sheet?.cells.get(0, 0)?.value).toEqual(numberValue(7))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(stringValue('fast & safe'))
		expect(sheet?.cells.get(0, 2)?.value).toEqual(stringValue('Shared & typed'))
		expect(sheet?.cells.get(0, 3)?.value).toEqual(booleanValue(true))
		expect(sheet?.cells.get(0, 26)?.value).toEqual(numberValue(-3.5))
		expect(sheet?.cells.get(1, 25)?.value).toEqual(numberValue(26))
		expect(sheet?.cells.get(1, 26)?.value).toEqual(numberValue(27))
		expect(sheet?.cells.get(1, 27)?.value).toEqual(numberValue(28))
		expect(sheet?.cells.get(2, 0)?.value).toEqual(numberValue(-12))
		expect(sheet?.cells.get(2, 1)?.value).toEqual(numberValue(1250))
		expect(sheet?.cells.get(2, 2)?.value).toEqual(numberValue(Number('12345678901234567')))
	})

	it('formula-only byte parser hydrates no-formula sheets and falls back for formulas', () => {
		const sharedStrings = parseSharedStrings(
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Shared formula-free</t></si>
</sst>`,
		)
		const ctx = {
			sharedStrings,
			styleIds: [S0],
			isDateFormat: [false],
			hasDateStyles: false,
			formulaOnly: true,
		}
		const noFormulaXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B1"/>
  <sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData>
</worksheet>`

		const sheet = parseSheetFormulaOnlyBytes('Sheet1', new TextEncoder().encode(noFormulaXml), ctx)

		expect(sheet).not.toBeNull()
		expect(sheet?.preservedDimensionRef).toBe('A1:B1')
		expect(sheet?.cells.get(0, 0)?.value).toEqual(numberValue(7))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(stringValue('Shared formula-free'))
		expect(sheet?.cells.get(0, 1)?.formula).toBeNull()

		const scanOnlySheet = parseSheetFormulaOnlyBytes(
			'Sheet1',
			new TextEncoder().encode(noFormulaXml),
			{ ...ctx, formulaModeHydrateValues: false },
		)
		expect(scanOnlySheet).not.toBeNull()
		expect(scanOnlySheet?.preservedDimensionRef).toBe('A1:B1')
		expect(scanOnlySheet?.cells.get(0, 0)).toBeUndefined()

		const formulaXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData>
</worksheet>`
		expect(
			parseSheetFormulaOnlyBytes('Sheet1', new TextEncoder().encode(formulaXml), ctx),
		).toBeNull()
	})

	it('full scalar byte parser preserves simple rich-read cell semantics', () => {
		const sharedStrings = parseSharedStrings(
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Shared &amp; full</t></si>
</sst>`,
		)
		const xml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:E1"/>
  <sheetData>
    <row r="1" ht="18" customHeight="1"><c r="A1"><v>7</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>inline &amp; full</t></is></c><c r="D1" t="b"><v>1</v></c><c r="E1" s="1"><v>45292</v></c></row>
  </sheetData>
</worksheet>`
		const sheet = parseSheetFullScalarBytes('Sheet1', new TextEncoder().encode(xml), {
			sharedStrings,
			styleIds: [S0, 1 as StyleId],
			isDateFormat: [false, true],
			hasDateStyles: true,
			valuesOnly: false,
			richMetadata: true,
		})

		expect(sheet).not.toBeNull()
		expect(sheet?.preservedDimensionRef).toBe('A1:E1')
		expect(sheet?.cells.get(0, 0)?.value).toEqual(numberValue(7))
		expect(sheet?.cells.get(0, 1)?.value).toEqual(stringValue('Shared & full'))
		expect(sheet?.cells.get(0, 1)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 2)?.value).toEqual(stringValue('inline & full'))
		expect(sheet?.cells.get(0, 3)?.value).toEqual(booleanValue(true))
		expect(sheet?.cells.get(0, 4)?.value).toEqual(dateValue(45292))
		expect(sheet?.cells.get(0, 4)?.styleId).toBe(1)
		expect(sheet?.rowHeights.get(0)).toBe(18)
		expect(sheet?.rowDefs.get(0)).toEqual({ customHeight: true })

		const formulaModeSheet = parseSheetFullScalarBytes('Sheet1', new TextEncoder().encode(xml), {
			sharedStrings,
			styleIds: [S0, 1 as StyleId],
			isDateFormat: [false, true],
			hasDateStyles: true,
			formulaOnly: true,
			richMetadata: true,
		})
		expect(formulaModeSheet?.cells.get(0, 1)?.value).toEqual(stringValue('Shared & full'))
		expect(formulaModeSheet?.cells.get(0, 4)?.styleId).toBe(1)
	})

	it('full scalar byte parser rejects formulas and hydrates rich sheet metadata', () => {
		const ctx = {
			sharedStrings: emptySharedStrings(),
			styleIds: [S0],
			isDateFormat: [false],
			hasDateStyles: false,
			valuesOnly: false,
			richMetadata: true,
		}
		const formulaSheet = parseSheetFullScalarBytes(
			'Sheet1',
			new TextEncoder().encode(
				`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`,
			),
			ctx,
		)
		const metadataSheet = parseSheetFullScalarBytes(
			'Sheet1',
			new TextEncoder().encode(
				`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews/><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData><autoFilter ref="A1:A1"/></worksheet>`,
			),
			ctx,
		)

		expect(formulaSheet).toBeNull()
		expect(metadataSheet?.cells.get(0, 0)?.value).toEqual(numberValue(2))
		expect(metadataSheet?.autoFilter?.ref).toBe('A1:A1')
	})

	it('values mode preserves mixed direct values, dates, row metadata, maxRows, and sheet metadata', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Shared &amp; decoded</t></si>
</sst>`,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen"/></sheetView></sheetViews>
  <sheetData>
    <row r="1" ht="27" customHeight="1" hidden="1" collapsed="1" outlineLevel="2">
      <c r="A1"><v>42.5</v></c>
      <c r="B1" t="s"><v>0</v></c>
      <c r="C1" t="inlineStr"><is><t>inline &amp; entity</t></is></c>
      <c r="D1" t="b"><v>0</v></c>
      <c r="E1" t="e"><v>#N/A</v></c>
      <c r="F1" t="str"><v>plain &lt;text&gt;</v></c>
      <c r="G1"><f>SUM(A1:A1)</f><v>43</v></c>
      <c r="H1"><f>SUM(A1:A1)</f></c>
      <c r="I1" s="1"><v>45292</v></c>
    </row>
    <row r="2"><c r="A2"><v>99</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <autoFilter ref="A1:I2"/>
  <pageMargins left="0.7" right="0.8" top="0.9" bottom="1"/>
  <pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0"/>
  <rowBreaks count="1" manualBreakCount="1"><brk id="1" min="0" max="16383" man="1"/></rowBreaks>
  <ignoredErrors><ignoredError sqref="F1" numberStoredAsText="1"/></ignoredErrors>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)
		expect(result.value.loadInfo.richSheetMetadataHydrated).toBe(false)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 42.5 })
		expect(sheet?.cells.get(0, 1)?.value).toEqual({
			kind: 'string',
			value: 'Shared & decoded',
		})
		expect(sheet?.cells.get(0, 2)?.value).toEqual({
			kind: 'string',
			value: 'inline & entity',
		})
		expect(sheet?.cells.get(0, 3)?.value).toEqual({ kind: 'boolean', value: false })
		expect(sheet?.cells.get(0, 4)?.value).toEqual({ kind: 'error', value: '#N/A' })
		expect(sheet?.cells.get(0, 5)?.value).toEqual({ kind: 'string', value: 'plain <text>' })
		expect(sheet?.cells.get(0, 6)?.value).toEqual({ kind: 'number', value: 43 })
		expect(sheet?.cells.get(0, 6)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 7)?.value).toEqual(EMPTY)
		expect(sheet?.cells.get(0, 7)?.formula).toBeNull()
		expect(sheet?.cells.get(0, 8)?.value).toEqual({ kind: 'date', serial: 45292 })
		expect(sheet?.cells.get(1, 0)?.value).toEqual({ kind: 'number', value: 99 })

		expect(sheet?.rowHeights.get(0)).toBe(27)
		expect(sheet?.rowDefs.get(0)).toEqual({
			customHeight: true,
			hidden: true,
			collapsed: true,
			outlineLevel: 2,
		})
		expect(sheet?.frozenRows).toBe(1)
		expect(sheet?.merges).toEqual([{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }])
		expect(sheet?.autoFilter?.ref).toBe('A1:I2')
		expect(sheet?.pageMargins).toEqual({ left: 0.7, right: 0.8, top: 0.9, bottom: 1 })
		expect(sheet?.pageSetup).toEqual({ orientation: 'portrait', fitToWidth: 1, fitToHeight: 0 })
		expect(sheet?.rowBreaks).toEqual([{ id: 1, min: 0, max: 16383, man: true }])
		expect(sheet?.ignoredErrors).toEqual([{ sqref: 'F1', numberStoredAsText: true }])

		const capped = readXlsx(bytes, { mode: 'values', maxRows: 1 })
		expectOk(capped)
		expect(capped.value.loadInfo.isPartial).toBe(true)
		const cappedSheet = capped.value.workbook.sheets[0]
		expect(cappedSheet?.cells.get(0, 8)?.value).toEqual({ kind: 'date', serial: 45292 })
		expect(cappedSheet?.cells.get(1, 0)).toBeUndefined()
		expect(cappedSheet?.rowHeights.get(0)).toBe(27)
		expect(cappedSheet?.autoFilter?.ref).toBe('A1:I2')
	})

	it('values mode reads dense workbooks successfully across repeated runs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		if (!sheet) throw new Error('Missing sheet')
		for (let r = 0; r < 500; r++) {
			for (let c = 0; c < 20; c++) {
				sheet.cells.set(r, c, {
					value: numberValue(r * 20 + c),
					formula: null,
					styleId: S0,
				})
			}
		}
		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(written.error.message)
		const bytes = new Uint8Array(written.value)

		const iterations = 8
		for (let i = 0; i < iterations; i++) {
			const fullResult = readXlsx(bytes)
			if (!fullResult.ok) throw new Error(fullResult.error.message)
			const valuesResult = readXlsx(bytes, { mode: 'values' })
			if (!valuesResult.ok) throw new Error(valuesResult.error.message)
			const fullSheet = fullResult.value.workbook.sheets[0]
			const valuesSheet = valuesResult.value.workbook.sheets[0]
			expect(valuesSheet?.cells.get(499, 19)?.value).toEqual(fullSheet?.cells.get(499, 19)?.value)
		}
	})

	it('preserves worksheet layout metadata in values mode', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane xSplit="12540" ySplit="3840" topLeftCell="K23" activePane="bottomRight"/><selection pane="bottomRight" activeCell="K23" sqref="K23"/></sheetView></sheetViews>
  <cols><col min="1" max="1" width="24" customWidth="1"/></cols>
  <sheetData>
    <row r="1" spans="1:1" ht="12.8" customHeight="0" s="2" customFormat="1" thickBot="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
  <mergeCells><mergeCell ref="A1:B2"/></mergeCells>
  <autoFilter ref="A1:B5"/>
  <pageMargins left="0.7" right="0.8" top="0.9" bottom="1"/>
  <printOptions gridLines="1"/>
  <headerFooter><oddHeader>&amp;LTest</oddHeader></headerFooter>
  <drawing r:id="rIdDrawing"/>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.frozenRows).toBe(0)
		expect(sheet?.frozenCols).toBe(0)
		expect(sheet?.preservedPaneAttributes).toEqual({
			xSplit: '12540',
			ySplit: '3840',
			topLeftCell: 'K23',
			activePane: 'bottomRight',
		})
		expect(sheet?.preservedSheetViewSelections).toEqual([
			{ pane: 'bottomRight', activeCell: 'K23', sqref: 'K23' },
		])
		expect(sheet?.colWidths.get(0)).toBe(24)
		expect(sheet?.rowHeights.get(0)).toBe(12.8)
		expect(sheet?.rowDefs.get(0)).toEqual({
			spans: '1:1',
			style: 2,
			customFormat: true,
			customHeight: false,
			thickBot: true,
		})
		expect(sheet?.merges).toEqual([{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }])
		expect(sheet?.autoFilter?.ref).toBe('A1:B5')
		expect(sheet?.pageMargins).toEqual({ left: 0.7, right: 0.8, top: 0.9, bottom: 1 })
		expect(sheet?.printOptions).toEqual({ gridLines: true })
		expect(sheet?.headerFooter).toEqual({ oddHeader: '&LTest' })
		expect(sheet?.drawingRefs).toEqual({ hasDrawing: true, hasLegacyDrawing: false })
	})

	it('preserves date decoding in values mode without full style hydration', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1"><v>45292</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'date',
			serial: 45292,
		})
	})

	it('detects custom date formats with conditions and locale markers', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/styles.xml': `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="165" formatCode="[&lt;=60]d-mmm;d-mmm"/>
    <numFmt numFmtId="166" formatCode="[$-409]d-mmm-yyyy"/>
    <numFmt numFmtId="167" formatCode="[Red]0.00"/>
    <numFmt numFmtId="168" formatCode="[&lt;=60]0;0"/>
  </numFmts>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="167" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="168" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>60</v></c>
      <c r="B1" s="2"><v>45292</v></c>
      <c r="C1" s="3"><v>12.5</v></c>
      <c r="D1" s="4"><v>60</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		for (const mode of ['full', 'values'] as const) {
			const result = mode === 'values' ? readXlsx(bytes, { mode: 'values' }) : readXlsx(bytes)
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'date', serial: 60 })
			expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'date', serial: 45292 })
			expect(sheet?.cells.get(0, 2)?.value).toEqual({ kind: 'number', value: 12.5 })
			expect(sheet?.cells.get(0, 3)?.value).toEqual({ kind: 'number', value: 60 })
		}
	})

	it('supports selective sheet parsing', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Archive" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData>
</worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>20</v></c></row></sheetData>
</worksheet>`,
		})

		const result = readXlsx(bytes, { sheets: ['Archive'] })
		expectOk(result)

		expect(result.value.workbook.sheets).toHaveLength(1)
		expect(result.value.workbook.sheets[0]?.name).toBe('Archive')
		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 20,
		})
		expect(result.value.report.features.some((feature) => feature.feature === 'partialLoad')).toBe(
			true,
		)
	})

	it('handles 100K+ unique shared strings with lazy parsing', () => {
		const count = 100_001
		const sstEntries = Array.from({ length: count }, (_, i) => `<si><t>str_${i}</t></si>`)
		const sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">
${sstEntries.join('\n')}
</sst>`

		const rowEntries: string[] = []
		for (let r = 1; r <= 1000; r++) {
			rowEntries.push(`<row r="${r}"><c r="A${r}" t="s"><v>${r - 1}</v></c></row>`)
		}
		rowEntries.push(`<row r="1001"><c r="A1001" t="s"><v>100000</v></c></row>`)
		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
${rowEntries.join('\n')}
  </sheetData>
</worksheet>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': sstXml,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const result = readXlsx(bytes, { mode: 'values' })
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		if (!sheet) return

		expect(sheet.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'str_0' })
		expect(sheet.cells.get(999, 0)?.value).toEqual({ kind: 'string', value: 'str_999' })
		expect(sheet.cells.get(1000, 0)?.value).toEqual({ kind: 'string', value: 'str_100000' })

		for (let i = 0; i <= 10; i++) {
			const cell = sheet.cells.get(i, 0)
			expect(cell?.value).toEqual({ kind: 'string', value: `str_${i}` })
		}
	})

	it('stops parsing after maxRows when maxRows option is set', () => {
		const rowEntries = Array.from(
			{ length: 20 },
			(_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c></row>`,
		)
		const sheetXml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
${rowEntries.join('\n')}
  </sheetData>
</worksheet>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const result = readXlsx(bytes, { maxRows: 10 })
		expectOk(result)
		expect(result.value.loadInfo.isPartial).toBe(true)
		expect(result.value.loadInfo.maxRows).toBe(10)
		expect(result.value.loadInfo.partialReasons).toContain(
			'only the first 10 row(s) are hydrated per loaded sheet',
		)
		expect(result.value.workbook.sourceArchiveBytes).toBeNull()

		const sheet = result.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		if (!sheet) return

		expect(sheet.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(sheet.cells.get(9, 0)?.value).toEqual({ kind: 'number', value: 10 })
		expect(sheet.cells.get(10, 0)).toBeUndefined()
		expect(sheet.cells.get(19, 0)).toBeUndefined()
	})

	it('preserves row metadata in streamed capped values reads', () => {
		const rowEntries = [
			'<row r="1" ht="21" customHeight="1" hidden="1" outlineLevel="2"><c r="A1"><v>1</v></c></row>',
			'<row r="2" collapsed="1"><c r="A2"><v>2</v></c></row>',
			...Array.from(
				{ length: 4000 },
				(_, i) => `<row r="${i + 3}"><c r="A${i + 3}"><v>${i + 3}</v></c></row>`,
			),
		]
		const sheetXml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
${rowEntries.join('\n')}
  </sheetData>
</worksheet>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const result = readXlsx(bytes, { mode: 'values', maxRows: 2 })
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(sheet?.cells.get(1, 0)?.value).toEqual({ kind: 'number', value: 2 })
		expect(sheet?.cells.get(2, 0)).toBeUndefined()
		expect(sheet?.rowHeights.get(0)).toBe(21)
		expect(sheet?.rowDefs.get(0)).toEqual({
			customHeight: true,
			hidden: true,
			outlineLevel: 2,
		})
		expect(sheet?.rowDefs.get(1)).toEqual({ collapsed: true })
	})

	it('reports preserved non-semantic parts explicitly', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/custom/custom1.xml" ContentType="application/custom+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
  <Override PartName="/xl/theme/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
  <Override PartName="/xl/xmlMaps.xml" ContentType="application/xml"/>
  <Override PartName="/xl/customProperty1.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty"/>
  <Override PartName="/xl/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
  <Override PartName="/xl/revisions/revisionHeaders.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml"/>
  <Override PartName="/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rIdXmlMaps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps" Target="xmlMaps.xml"/>
  <Relationship Id="rIdCustomProperty" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="customProperty1.bin"/>
  <Relationship Id="rIdDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
  <Relationship Id="rIdRevisionHeaders" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders" Target="revisions/revisionHeaders.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': SHEET_XML,
			'xl/custom/custom1.xml': '<custom>preserve me</custom>',
			'xl/calcChain.xml': '<calcChain/>',
			'xl/comments1.xml': '<comments/>',
			'xl/externalLinks/externalLink1.xml': '<externalLink/>',
			'xl/theme/theme.xml': '<a:theme/>',
			'xl/charts/style1.xml': '<cs:chartStyle/>',
			'xl/charts/colors1.xml': '<cs:colors/>',
			'xl/richData/richValueTypes.xml': '<rvTypes/>',
			'xl/xmlMaps.xml': '<MapInfo/>',
			'xl/customProperty1.bin': 'custom-property-bytes',
			'xl/diagrams/data1.xml': '<dgm:dataModel/>',
			'xl/revisions/revisionHeaders.xml': '<headers/>',
			'pivotCache/pivotCacheDefinition1.xml': '<pivotCacheDefinition/>',
			'docProps/core.xml': '<cp:coreProperties/>',
			'docProps/app.xml': '<Properties/>',
			'xl/.DS_Store': 'junk',
			'__MACOSX/._custom1.xml': 'junk',
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.report.status).toBe('has-preserved')
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedOther'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/custom/custom1.xml'],
			}),
		)
		expect(result.value.report.features.some((feature) => feature.feature === 'calcChain')).toBe(
			true,
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedCalcChain'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/calcChain.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedComments'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/comments1.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedExternalLink'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/externalLinks/externalLink1.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedTheme'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/theme/theme.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedChartStyle'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/charts/style1.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedChartColor'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/charts/colors1.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedMetadata'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 2,
				locations: ['xl/richData/richValueTypes.xml', 'xl/customProperty1.bin'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedCustomXml'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/xmlMaps.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedDrawing'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/diagrams/data1.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedRevision'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/revisions/revisionHeaders.xml'],
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedPivot'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['pivotCache/pivotCacheDefinition1.xml'],
			}),
		)
		expect(
			result.value.report.features.find(
				(feature) => feature.feature === 'preservedDocumentProperties',
			),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 2,
				locations: ['docProps/core.xml', 'docProps/app.xml'],
			}),
		)
	})

	it('reports vendor security and orphan worksheet sidecars without preservedOther ambiguity', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="xen" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rDellEncryptedDoc" Type="http://schemas.dell.com/ddp/2016/relationships/xenFile" Target="ddp/ddpfile.xen"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': SHEET_XML,
			'xl/worksheets/sheet1_formatted.xml': '<worksheet/>',
			'ddp/ddpfile.xen': 'opaque-vendor-security',
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.report.features).toContainEqual(
			expect.objectContaining({
				feature: 'preservedVendorSecurity',
				tier: 'preserved',
				count: 1,
				locations: ['ddp/ddpfile.xen'],
				note: expect.stringContaining('Vendor security policy'),
			}),
		)
		expect(result.value.report.features).toContainEqual(
			expect.objectContaining({
				feature: 'preservedWorksheetSidecar',
				tier: 'preserved',
				count: 1,
				locations: ['xl/worksheets/sheet1_formatted.xml'],
				note: expect.stringContaining('without workbook relationships'),
			}),
		)
		expect(
			result.value.report.features.some((feature) => feature.feature === 'preservedOther'),
		).toBe(false)
	})

	it('identifies and preserves threaded comments as preservedThreadedComments', () => {
		const threadedCommentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Comment text</text>
  </threadedComment>
</ThreadedComments>`

		const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`

		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
			'xl/worksheets/_rels/sheet1.xml.rels': sheetRels,
			'xl/threadedComments/threadedComment1.xml': threadedCommentXml,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(
			result.value.report.features.find((f) => f.feature === 'preservedThreadedComments'),
		).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/threadedComments/threadedComment1.xml'],
			}),
		)

		const tcCapsule = result.value.capsules.find(
			(c) => c.partPath === 'xl/threadedComments/threadedComment1.xml',
		)
		expect(tcCapsule).toBeDefined()
		expect(tcCapsule?.anchor).toEqual(expect.objectContaining({ kind: 'sheet', sheetName: 'Data' }))
	})
})

describe('stub cells', () => {
	it('skips cells with type attributes but no cached value', () => {
		const stubSheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"/>
      <c r="B1" t="n"/>
      <c r="C1" t="b"/>
      <c r="D1" t="e"/>
      <c r="E1" t="str"/>
    </row>
  </sheetData>
</worksheet>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': stubSheetXml,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		expect(sheet?.cells.cellCount()).toBe(0)
	})
})

describe('round-trip fidelity', () => {
	function roundTrip(wb: Workbook) {
		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(`write failed: ${written.error.message}`)
		const read = readXlsx(written.value)
		if (!read.ok) throw new Error(`read failed: ${read.error.message}`)
		return { bytes: written.value, result: read.value }
	}

	it('preserves cell values (number, string, boolean, date, empty)', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Values')
		sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Hello'), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: booleanValue(true), formula: null, styleId: S0 })
		const dateFmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
		sheet.cells.set(0, 3, {
			value: dateValue(45292),
			formula: null,
			styleId: dateFmtId,
		})
		sheet.cells.set(0, 4, { value: EMPTY, formula: null, styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s).toBeDefined()
		expect(s?.name).toBe('Values')
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 42 })
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'boolean', value: true })
		expect(s?.cells.get(0, 3)?.value).toEqual({ kind: 'date', serial: 45292 })
		const emptyCell = s?.cells.get(0, 4)
		expect(emptyCell === undefined || emptyCell?.value?.kind === 'empty').toBe(true)
	})

	it('preserves formula text strings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: numberValue(30), formula: 'A1+B1', styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(55), formula: 'SUM(A1:A10)', styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 2)?.formula).toBe('A1+B1')
		expect(s?.cells.get(1, 0)?.formula).toBe('SUM(A1:A10)')
	})

	it('preserves styles (number formats, bold font)', () => {
		const wb = new Workbook()
		const boldId = wb.styles.register({ font: { bold: true } })
		const pctId = wb.styles.register({ numberFormat: '0.0%' })
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })
		sheet.cells.set(0, 1, { value: numberValue(0.75), formula: null, styleId: pctId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell0 = s?.cells.get(0, 0)
		const cell1 = s?.cells.get(0, 1)
		expect(result.workbook.styles.get(cell0?.styleId ?? S0)?.font?.bold).toBe(true)
		expect(result.workbook.styles.get(cell1?.styleId ?? S0)?.numberFormat).toBe('0.0%')
	})

	it('preserves multi-sheet workbook (names and data)', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('Input')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		s1.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: S0 })
		const s2 = wb.addSheet('Summary')
		s2.cells.set(0, 0, { value: stringValue('Total'), formula: null, styleId: S0 })
		s2.cells.set(0, 1, { value: numberValue(3), formula: 'Input!A1+Input!B1', styleId: S0 })
		const s3 = wb.addSheet('Archive')
		s3.cells.set(0, 0, { value: stringValue('Archived'), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)
		expect(result.workbook.sheets).toHaveLength(3)
		expect(result.workbook.sheets[0]?.name).toBe('Input')
		expect(result.workbook.sheets[1]?.name).toBe('Summary')
		expect(result.workbook.sheets[2]?.name).toBe('Archive')
		expect(result.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 1,
		})
		expect(result.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'number',
			value: 2,
		})
		expect(result.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Total',
		})
		expect(result.workbook.sheets[1]?.cells.get(0, 1)?.formula).toBe('Input!A1+Input!B1')
		expect(result.workbook.sheets[2]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Archived',
		})
	})
})
