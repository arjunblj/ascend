import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { Workbook } from '@ascend/core'
import { stringValue } from '@ascend/schema'
import { strToU8, zipSync } from 'fflate'
import { writeXlsx } from '../writer/index.ts'
import { readXlsx } from './index.ts'

const S0 = 0 as StyleId

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

describe('Unicode cell values', () => {
	it('parses CJK, emoji, RTL, accented, and mixed Unicode in shared strings', () => {
		const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>你好世界</t></si>
  <si><t>🎉🚀</t></si>
  <si><t>مرحبا</t></si>
  <si><t>café, naïve, résumé</t></si>
  <si><t>Hello 世界 🌍</t></si>
</sst>`

		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
      <c r="D1" t="s"><v>3</v></c>
      <c r="E1" t="s"><v>4</v></c>
    </row>
  </sheetData>
</worksheet>`

		const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': workbookXml,
			'xl/sharedStrings.xml': sharedStrings,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: '你好世界' })
		expect(sheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: '🎉🚀' })
		expect(sheet?.cells.get(0, 2)?.value).toEqual({ kind: 'string', value: 'مرحبا' })
		expect(sheet?.cells.get(0, 3)?.value).toEqual({
			kind: 'string',
			value: 'café, naïve, résumé',
		})
		expect(sheet?.cells.get(0, 4)?.value).toEqual({ kind: 'string', value: 'Hello 世界 🌍' })
	})

	it('round-trips Unicode cell values through write → read', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: stringValue('你好世界'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('🎉🚀'), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: stringValue('مرحبا'), formula: null, styleId: S0 })
		sheet.cells.set(0, 3, { value: stringValue('café, naïve, résumé'), formula: null, styleId: S0 })
		sheet.cells.set(0, 4, { value: stringValue('Hello 世界 🌍'), formula: null, styleId: S0 })

		const written = writeXlsx(wb)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const read = readXlsx(written.value)
		expect(read.ok).toBe(true)
		if (!read.ok) return

		const s = read.value.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: '你好世界' })
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: '🎉🚀' })
		expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'string', value: 'مرحبا' })
		expect(s?.cells.get(0, 3)?.value).toEqual({
			kind: 'string',
			value: 'café, naïve, résumé',
		})
		expect(s?.cells.get(0, 4)?.value).toEqual({ kind: 'string', value: 'Hello 世界 🌍' })
	})
})

describe('Unicode sheet names', () => {
	it('parses workbook with Unicode sheet names and survives round-trip', () => {
		const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

		const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`

		const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="データ" sheetId="1" r:id="rId1"/>
    <sheet name="Résumé" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`

		const bytes = makeXlsx({
			'[Content_Types].xml': contentTypes,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': workbookRels,
			'xl/workbook.xml': workbookXml,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.sheets[0]?.name).toBe('データ')
		expect(result.value.workbook.sheets[1]?.name).toBe('Résumé')

		const written = writeXlsx(result.value.workbook, result.value.capsules)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const reread = readXlsx(written.value)
		expect(reread.ok).toBe(true)
		if (!reread.ok) return

		expect(reread.value.workbook.sheets[0]?.name).toBe('データ')
		expect(reread.value.workbook.sheets[1]?.name).toBe('Résumé')
	})
})

describe('Unicode in formulas', () => {
	it('preserves formula referencing Unicode sheet name through round-trip', () => {
		const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

		const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`

		const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="データ" sheetId="1" r:id="rId1"/>
    <sheet name="Résumé" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`

		const bytes = makeXlsx({
			'[Content_Types].xml': contentTypes,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': workbookRels,
			'xl/workbook.xml': workbookXml,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>'データ'!A1</f><v>42</v></c></row></sheetData></worksheet>`,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const resumeSheet = result.value.workbook.getSheet('Résumé')
		expect(resumeSheet).toBeDefined()
		const cell = resumeSheet?.cells.get(0, 0)
		expect(cell?.formula).toBe("'データ'!A1")

		const written = writeXlsx(result.value.workbook, result.value.capsules)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const reread = readXlsx(written.value)
		expect(reread.ok).toBe(true)
		if (!reread.ok) return

		const rereadResume = reread.value.workbook.getSheet('Résumé')
		expect(rereadResume?.cells.get(0, 0)?.formula).toBe("'データ'!A1")
	})
})

describe('Rich text round-trip', () => {
	it('preserves rich text (bold/italic runs) through read → write → read', () => {
		const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><rPr><i/></rPr><t>Italic</t></r>
    <r><t>Normal</t></r>
  </si>
</sst>`

		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`

		const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

		const bytes = makeXlsx({
			'[Content_Types].xml': CONTENT_TYPES,
			'_rels/.rels': ROOT_RELS,
			'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
			'xl/workbook.xml': workbookXml,
			'xl/sharedStrings.xml': sharedStrings,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const result = readXlsx(bytes)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'Bold', bold: true }, { text: 'Italic', italic: true }, { text: 'Normal' }],
		})

		const written = writeXlsx(result.value.workbook, result.value.capsules)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const reread = readXlsx(written.value)
		expect(reread.ok).toBe(true)
		if (!reread.ok) return

		expect(reread.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'Bold', bold: true }, { text: 'Italic', italic: true }, { text: 'Normal' }],
		})
	})
})
