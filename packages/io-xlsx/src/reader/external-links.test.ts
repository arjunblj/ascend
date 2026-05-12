import { describe, expect, test } from 'bun:test'
import { parseExternalBookRelationshipId } from './external-links.ts'

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
})
