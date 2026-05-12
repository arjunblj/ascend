import { describe, expect, test } from 'bun:test'
import {
	externalLinkPathRelationshipKind,
	isExternalLinkPathRelationshipType,
	parseRelationships,
	resolvePath,
} from './relationships.ts'

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

	test('classifies Excel external workbook path relationship base semantics', () => {
		expect(
			externalLinkPathRelationshipKind(
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
			),
		).toBe('externalLinkPath')
		expect(
			externalLinkPathRelationshipKind(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup',
			),
		).toBe('xlStartup')
		expect(
			externalLinkPathRelationshipKind(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup',
			),
		).toBe('xlAlternateStartup')
		expect(
			externalLinkPathRelationshipKind(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary',
			),
		).toBe('xlLibrary')
		expect(
			externalLinkPathRelationshipKind(
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing',
			),
		).toBe('xlPathMissing')
		expect(
			externalLinkPathRelationshipKind(
				'http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath',
			),
		).toBe('externalLinkPath')
		expect(
			externalLinkPathRelationshipKind(
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
			),
		).toBe('unknown')
		expect(externalLinkPathRelationshipKind(undefined)).toBeUndefined()
	})

	test('keeps raw strict relationship types while exposing normalized lookup types', () => {
		const rels = parseRelationships(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMetadata" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`)

		expect(rels[0]).toEqual({
			id: 'rIdSheet',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
			rawType: 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet',
			target: 'worksheets/sheet1.xml',
		})
		expect(rels[1]).toEqual({
			id: 'rIdMetadata',
			type: 'http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata',
			target: 'metadata.xml',
		})
	})

	test('resolves internal relationship URI targets to package part paths', () => {
		expect(resolvePath('', 'xl/work%20book.xml')).toBe('xl/work book.xml')
		expect(resolvePath('xl/work book.xml', 'worksheets/sheet%201.xml')).toBe(
			'xl/worksheets/sheet 1.xml',
		)
		expect(resolvePath('xl/worksheets/sheet 1.xml', '..\\drawings\\drawing%201.xml')).toBe(
			'xl/drawings/drawing 1.xml',
		)
		expect(resolvePath('xl/workbook.xml', 'worksheets/sheet%2F1.xml')).toBe(
			'xl/worksheets/sheet%2F1.xml',
		)
	})
})
