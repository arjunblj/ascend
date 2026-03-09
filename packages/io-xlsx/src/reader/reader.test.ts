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
		expect(result.value.workbook.sheets[0]?.cells.get(1, 0)?.formula).toBe('B2*2')
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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Scores" ref="A1:B3" headerRowCount="1" totalsRowCount="0" headerRowDxfId="5" dataDxfId="6">
  <autoFilter ref="A1:B3"><sortState ref="A2:B3"><sortCondition ref="B2:B3"/></sortState></autoFilter>
  <tableColumns count="2">
    <tableColumn id="1" name="Name" totalsRowLabel="Total" dataDxfId="7"/>
    <tableColumn id="2" name="Score" totalsRowFunction="sum" totalsRowDxfId="8"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.tables).toHaveLength(1)
		expect(sheet?.tables[0]?.name).toBe('Scores')
		expect(sheet?.tables[0]?.hasHeaders).toBe(true)
		expect(sheet?.tables[0]?.columns).toEqual([
			{ id: 1, name: 'Name', totalsRowLabel: 'Total', dataDxfId: 7 },
			{ id: 2, name: 'Score', totalsRowFunction: 'sum', totalsRowDxfId: 8 },
		])
		expect(sheet?.tables[0]?.autoFilter).toEqual({
			ref: 'A1:B3',
			columns: [],
			sortState: {
				ref: 'A2:B3',
				conditions: [{ ref: 'B2:B3' }],
			},
		})
		expect(sheet?.tables[0]?.headerRowDxfId).toBe(5)
		expect(sheet?.tables[0]?.dataDxfId).toBe(6)
		expect(sheet?.tables[0]?.tableStyleInfo).toEqual({
			name: 'TableStyleMedium2',
			showFirstColumn: false,
			showLastColumn: false,
			showRowStripes: true,
			showColumnStripes: false,
		})
		expect(result.value.report.features.some((feature) => feature.feature === 'table')).toBe(true)
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
		expect(result.ok).toBe(true)
		if (!result.ok) return
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
  <calcPr calcMode="manual" fullCalcOnLoad="1"/>
</workbook>`,
			'xl/sharedStrings.xml': SHARED_STRINGS,
			'xl/worksheets/sheet1.xml': SHEET_XML,
			'xl/calcChain.xml': `<?xml version="1.0"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(
			result.value.report.features.some((feature) => feature.feature === 'formulaFreshness'),
		).toBe(true)
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
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable11" cacheId="34">
  <location ref="O17"/>
</pivotTableDefinition>`,
			'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rIdRecords" recordCount="25">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D100" sheet="raw data"/>
  </cacheSource>
</pivotCacheDefinition>`,
			'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
			'xl/pivotCache/pivotCacheRecords1.xml':
				'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.pivotTables).toEqual([
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'PivotSheet',
				name: 'PivotTable11',
				cacheId: 34,
				locationRef: 'O17',
			},
		])
		expect(result.value.workbook.pivotCaches).toEqual([
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 34,
				relId: 'rIdPivotCache',
				recordCount: 25,
				sourceSheet: 'raw data',
				sourceRef: 'A1:D100',
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
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
  <data><tabular pivotCacheId="1332190931"/></data>
</slicerCacheDefinition>`,
			'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Product_Category" cache="Slicer_Product_Category" caption="Product Category"/>
</slicers>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.slicerCaches).toEqual([
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				name: 'Slicer_Product_Category',
				sourceName: 'Product_Category',
				pivotCacheId: 1332190931,
				pivotTableNames: ['PivotTable1'],
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1" filterPrivacy="1" codeName="Model"/>
  <bookViews>
    <workbookView activeTab="1" firstSheet="2" visibility="visible" tabRatio="600"/>
  </bookViews>
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
  <calcPr calcMode="manual" fullCalcOnLoad="1"/>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.workbookProperties).toEqual({
			date1904: true,
			filterPrivacy: true,
			codeName: 'Model',
		})
		expect(result.value.workbook.workbookViews).toEqual([
			{ activeTab: 1, firstSheet: 2, visibility: 'visible', tabRatio: 600 },
		])
		expect(result.value.workbook.externalReferences).toEqual(['xl/externalLinks/externalLink1.xml'])
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
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

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

		const result = readXlsx(bytes, { mode: 'metadata-only' })
		expect(result.ok).toBe(true)
		if (!result.ok) return

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
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="2" ySplit="1" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="2" width="18.5" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="24" customHeight="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:B10"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup orientation="landscape" fitToWidth="1" fitToHeight="2"/>
  <printOptions gridLines="1" headings="1"/>
  <headerFooter><oddHeader>&amp;LTest</oddHeader><oddFooter>&amp;R1</oddFooter></headerFooter>
  <ignoredErrors><ignoredError sqref="A1:B2" numberStoredAsText="1"/></ignoredErrors>
  <hyperlinks><hyperlink ref="A1" r:id="rIdHyper" display="Docs" tooltip="Open docs"/></hyperlinks>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHyper" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs"/>
</Relationships>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.frozenRows).toBe(1)
		expect(sheet?.frozenCols).toBe(2)
		expect(sheet?.colWidths.get(0)).toBe(18.5)
		expect(sheet?.colWidths.get(1)).toBe(18.5)
		expect(sheet?.rowHeights.get(0)).toBe(24)
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
		})
		expect(sheet?.printOptions).toEqual({
			gridLines: true,
			headings: true,
			horizontalCentered: undefined,
			verticalCentered: undefined,
		})
		expect(sheet?.headerFooter).toEqual({
			oddHeader: '&LTest',
			oddFooter: '&R1',
			evenHeader: undefined,
			evenFooter: undefined,
			firstHeader: undefined,
			firstFooter: undefined,
		})
		expect(sheet?.ignoredErrors).toEqual([{ sqref: 'A1:B2', numberStoredAsText: true }])
		expect(sheet?.hyperlinks.get('A1')).toEqual({
			target: 'https://example.com/docs',
			location: undefined,
			display: 'Docs',
			tooltip: 'Open docs',
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
		expect(result.ok).toBe(true)
		if (!result.ok) return
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
		expect(result.ok).toBe(true)
		if (!result.ok) return
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
		expect(result.ok).toBe(true)
		if (!result.ok) return
		expect(result.value.workbook.sheets[0]?.comments.get('B2')).toEqual({
			text: 'Hello',
			author: 'Ada',
		})
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
		expect(result.ok).toBe(true)
		if (!result.ok) return

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
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1" stopIfTrue="1">
      <formula>3</formula>
    </cfRule>
  </conditionalFormatting>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" sqref="B2:B4">
      <formula1>"Q1,Q2,Q3"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A10',
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
				],
			},
		])
		expect(sheet?.dataValidations).toEqual([
			{
				sqref: 'B2:B4',
				type: 'list',
				allowBlank: true,
				showInputMessage: true,
				formula1: '"Q1,Q2,Q3"',
			},
		])
		expect(result.value.workbook.differentialStyles).toEqual([
			{
				font: { bold: true },
				fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
			},
		])
	})

	it('supports metadata-only reads without parsing sheet cells', () => {
		const result = readXlsx(minimalXlsx(), { mode: 'metadata-only' })
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.sheets).toHaveLength(1)
		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.name).toBe('Data')
		expect(sheet?.cells.cellCount()).toBe(0)
		expect(result.value.report.features.some((feature) => feature.feature === 'partialLoad')).toBe(
			true,
		)
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
		expect(result.value.report.features.some((feature) => feature.feature === 'partialLoad')).toBe(
			true,
		)
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
</Types>`,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': WORKBOOK_XML,
			'xl/worksheets/sheet1.xml': SHEET_XML,
			'xl/custom/custom1.xml': '<custom>preserve me</custom>',
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

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
	})
})
