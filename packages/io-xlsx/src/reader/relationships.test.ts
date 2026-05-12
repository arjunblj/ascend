import { describe, expect, test } from 'bun:test'
import { isExternalLinkPathRelationshipType, parseRelationships } from './relationships.ts'

describe('relationships', () => {
	test('parses XML-legal single-quoted relationship attributes', () => {
		const rels = parseRelationships(`<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rIdOffice' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target='https://example.com/?q=Tom&amp;Jane' TargetMode='External'/>
</Relationships>`)

		expect(rels).toEqual([
			{
				id: 'rIdOffice',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
				target: 'xl/workbook.xml',
			},
			{
				id: 'rIdExternal',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
				target: 'https://example.com/?q=Tom&Jane',
				targetMode: 'External',
			},
		])
	})

	test('recognizes Excel external workbook path relationship types', () => {
		expect(
			isExternalLinkPathRelationshipType(
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
			),
		).toBe(true)
		expect(
			isExternalLinkPathRelationshipType(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup',
			),
		).toBe(true)
		expect(
			isExternalLinkPathRelationshipType(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup',
			),
		).toBe(true)
		expect(
			isExternalLinkPathRelationshipType(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary',
			),
		).toBe(true)
		expect(
			isExternalLinkPathRelationshipType(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing',
			),
		).toBe(true)
		expect(
			isExternalLinkPathRelationshipType(
				'http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath',
			),
		).toBe(true)
		expect(
			isExternalLinkPathRelationshipType(
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
			),
		).toBe(false)
	})
})
