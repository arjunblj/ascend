import { describe, expect, test } from 'bun:test'
import { isExternalLinkPathRelationshipType } from './relationships.ts'

describe('relationships', () => {
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
