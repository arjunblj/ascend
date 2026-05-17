import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { writeXlsx } from '../writer/index.ts'
import { parseExternalBookRelationshipId, parseExternalLinkInfo } from './external-links.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('external link metadata', () => {
	test('parses the externalBook relationship id across namespace prefixes', () => {
		expect(
			parseExternalBookRelationshipId(
				'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId2"/></externalLink>',
			),
		).toBe('rId2')
		expect(
			parseExternalBookRelationshipId(
				'<x:externalLink xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"><x:externalBook rel:id="rIdStrict"/></x:externalLink>',
			),
		).toBe('rIdStrict')
		expect(parseExternalBookRelationshipId('<externalLink/>')).toBeUndefined()
	})

	test('parses externalBook sheet names and external defined names', () => {
		expect(
			parseExternalLinkInfo(
				`<x:externalLink xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:externalBook r:id="rIdBook">
    <x:sheetNames>
      <x:sheetName val="Summary &amp; Ops"/>
      <x:sheetName val="Data"/>
    </x:sheetNames>
    <x:definedNames>
      <x:definedName name="ExtTotal" refersTo="'Summary &amp; Ops'!$B$2" sheetId="0"/>
      <x:definedName name="GlobalFlag" refersTo="TRUE"/>
    </x:definedNames>
  </x:externalBook>
</x:externalLink>`,
			),
		).toEqual({
			kind: 'externalBook',
			relationshipId: 'rIdBook',
			externalBookSheetNames: ['Summary & Ops', 'Data'],
			externalBookDefinedNames: [
				{ name: 'ExtTotal', refersTo: "'Summary & Ops'!$B$2", sheetId: 0 },
				{ name: 'GlobalFlag', refersTo: 'TRUE' },
			],
		})
	})

	test('ignores externalBook id attributes outside the relationships namespace', () => {
		expect(
			parseExternalBookRelationshipId(
				'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:foo="urn:not-relationships"><externalBook foo:id="rIdBad"/></externalLink>',
			),
		).toBeUndefined()
		expect(
			parseExternalBookRelationshipId(
				'<externalLink xmlns:foo="urn:not-relationships"><externalBook foo:id="rIdBad" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdGood"/></externalLink>',
			),
		).toBe('rIdGood')
		expect(parseExternalBookRelationshipId('<externalBook rel:id="rIdUnbound"/>')).toBeUndefined()
	})

	test('decodes XML entities in relationship ids', () => {
		expect(
			parseExternalBookRelationshipId(
				'<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId&amp;2"/>',
			),
		).toBe('rId&2')
	})

	test('parses XML-legal single-quoted externalBook relationship ids', () => {
		expect(
			parseExternalBookRelationshipId(
				"<externalBook xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' r:id='rId&amp;Single'/>",
			),
		).toBe('rId&Single')
	})

	test('parses DDE and OLE external link source metadata', () => {
		expect(
			parseExternalLinkInfo(
				`<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <ddeLink ddeService="Excel" ddeTopic="Sheet1!R1C1">
    <ddeItems count="2">
      <ddeItem name="R1C1" advise="1" preferPic="0"/>
      <ddeItem name="R2C1" ole="1"/>
    </ddeItems>
  </ddeLink>
</externalLink>`,
			),
		).toEqual({
			kind: 'ddeLink',
			ddeService: 'Excel',
			ddeTopic: 'Sheet1!R1C1',
			ddeItems: [
				{ name: 'R1C1', advise: true, preferPicture: false },
				{ name: 'R2C1', ole: true },
			],
		})
		expect(
			parseExternalLinkInfo(
				'<externalLink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><oleLink r:id="rIdOle"/></externalLink>',
			),
		).toEqual({
			kind: 'oleLink',
			relationshipId: 'rIdOle',
		})
		expect(parseExternalBookRelationshipId('<oleLink id="rIdOle"/>')).toBeUndefined()
	})

	test('inventories OLE external link relationship binding across save and reopen', () => {
		const source = readXlsx(externalLinkWorkbook())
		expectOk(source)
		expect(source.value.workbook.externalReferenceDetails).toEqual([
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink1.xml',
				externalLinkKind: 'oleLink',
				externalLinkRelId: 'rIdOle',
				linkRelId: 'rIdOle',
				linkBindingStatus: 'externalLinkRelId',
				target: '../linked/source.xlsx',
				targetMode: 'External',
			}),
		])

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.externalReferenceDetails[0]).toMatchObject({
			externalLinkKind: 'oleLink',
			externalLinkRelId: 'rIdOle',
			linkBindingStatus: 'externalLinkRelId',
			target: '../linked/source.xlsx',
		})
	})

	test('inventories DDE external link item metadata across save and reopen', () => {
		const source = readXlsx(ddeLinkWorkbook())
		expectOk(source)
		expect(source.value.workbook.externalReferenceDetails[0]).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			externalLinkKind: 'ddeLink',
			externalLinkDdeService: 'Excel',
			externalLinkDdeTopic: 'Sheet1!R1C1',
			externalLinkDdeItems: [
				{ name: 'R1C1', advise: true, preferPicture: false },
				{ name: 'R2C1', ole: true },
			],
			linkBindingStatus: 'missingPathRelationship',
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.externalReferenceDetails[0]).toMatchObject({
			externalLinkKind: 'ddeLink',
			externalLinkDdeItems: [
				{ name: 'R1C1', advise: true, preferPicture: false },
				{ name: 'R2C1', ole: true },
			],
		})
	})

	test('inventories externalBook sheet and defined-name metadata across save and reopen', () => {
		const source = readXlsx(externalBookWorkbook())
		expectOk(source)
		expect(source.value.workbook.externalReferenceDetails[0]).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			externalLinkKind: 'externalBook',
			externalBookRelId: 'rIdBook',
			externalBookSheetNames: ['Summary', 'Data'],
			externalBookDefinedNames: [
				{ name: 'ExtTotal', refersTo: 'Summary!$B$2', sheetId: 0 },
				{ name: 'GlobalFlag', refersTo: 'TRUE' },
			],
			linkBindingStatus: 'externalBookRelId',
			target: '../linked/source.xlsx',
			targetMode: 'External',
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.externalReferenceDetails[0]).toMatchObject({
			externalLinkKind: 'externalBook',
			externalBookSheetNames: ['Summary', 'Data'],
			externalBookDefinedNames: [
				{ name: 'ExtTotal', refersTo: 'Summary!$B$2', sheetId: 0 },
				{ name: 'GlobalFlag', refersTo: 'TRUE' },
			],
			target: '../linked/source.xlsx',
		})
	})
})

function ddeLinkWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <ddeLink ddeService="Excel" ddeTopic="Sheet1!R1C1">
    <ddeItems count="2">
      <ddeItem name="R1C1" advise="1" preferPic="0"/>
      <ddeItem name="R2C1" ole="1"/>
    </ddeItems>
  </ddeLink>
</externalLink>`,
	})
}

function externalLinkWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <oleLink r:id="rIdOle"/>
</externalLink>`,
		'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../linked/source.xlsx" TargetMode="External"/>
</Relationships>`,
	})
}

function externalBookWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rIdBook">
    <sheetNames>
      <sheetName val="Summary"/>
      <sheetName val="Data"/>
    </sheetNames>
    <definedNames>
      <definedName name="ExtTotal" refersTo="Summary!$B$2" sheetId="0"/>
      <definedName name="GlobalFlag" refersTo="TRUE"/>
    </definedNames>
  </externalBook>
</externalLink>`,
		'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../linked/source.xlsx" TargetMode="External"/>
</Relationships>`,
	})
}
