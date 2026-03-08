import { describe, expect, it } from 'bun:test'
import { strToU8, zipSync } from 'fflate'
import { readXlsx } from './index.ts'

function makeXlsx(parts: Record<string, string>): Uint8Array {
	const entries: Record<string, Uint8Array> = {}
	for (const [path, content] of Object.entries(parts)) {
		entries[path] = strToU8(content)
	}
	return zipSync(entries)
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

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

	it('parses merge cells', () => {
		const result = readXlsx(minimalXlsx())
		expect(result.ok).toBe(true)
		if (!result.ok) return

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
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect(result.error.code).toBe('CORRUPT_FILE')
	})

	it('returns error for ZIP missing required parts', () => {
		const empty = makeXlsx({ 'dummy.txt': 'nothing' })
		const result = readXlsx(empty)
		expect(result.ok).toBe(false)
		if (result.ok) return
		expect(result.error.code).toBe('CORRUPT_FILE')
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const cell = result.value.workbook.sheets[0]?.cells.get(0, 0)
		expect(cell?.value).toEqual({ kind: 'number', value: 99 })
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

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
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.definedNames.get('Total')).toBe('Data!$A$1')
	})
})
