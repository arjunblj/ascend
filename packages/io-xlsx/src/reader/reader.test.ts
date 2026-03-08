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
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const summary = result.value.workbook.getSheet('Summary')
		expect(summary).toBeDefined()
		if (!summary) return

		const resolved = result.value.workbook.definedNames.resolve('Budget', summary.id)
		expect(resolved?.scope.kind).toBe('sheet')
		expect(resolved?.formula).toBe('Summary!$A$1')
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const cell = result.value.workbook.sheets[0]?.cells.get(0, 0)
		expect(cell?.formula).toBe('SUM(B1:B2)')
		expect(result.value.report.features.some((feature) => feature.feature === 'arrayFormula')).toBe(
			true,
		)
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(
			result.value.report.features.some((feature) => feature.feature === 'sharedFormula'),
		).toBe(true)
		expect(result.value.workbook.sheets[0]?.cells.get(1, 0)?.formula).toBeNull()
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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Scores" ref="A1:B3" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Score"/>
  </tableColumns>
</table>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.tables).toHaveLength(1)
		expect(sheet?.tables[0]?.name).toBe('Scores')
		expect(sheet?.tables[0]?.hasHeaders).toBe(true)
		expect(sheet?.tables[0]?.columns.map((column) => column.name)).toEqual(['Name', 'Score'])
		expect(result.value.report.features.some((feature) => feature.feature === 'table')).toBe(true)
	})

	it('supports metadata-only reads without parsing sheet cells', () => {
		const result = readXlsx(minimalXlsx(), { mode: 'metadata-only' })
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.sheets).toHaveLength(1)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.name).toBe('Data')
		expect(sheet?.cells.cellCount()).toBe(0)
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.sheets).toHaveLength(1)
		expect(result.value.workbook.sheets[0]?.name).toBe('Archive')
		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 20,
		})
	})
})
