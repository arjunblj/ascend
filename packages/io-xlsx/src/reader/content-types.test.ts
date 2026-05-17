import { describe, expect, test } from 'bun:test'
import { parseContentTypes } from './content-types.ts'

describe('content types', () => {
	test('parses XML-legal single-quoted Default and Override attributes', () => {
		const contentTypes = parseContentTypes(`<?xml version='1.0' encoding='UTF-8'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='xml' ContentType='application/xml'/>
  <Default Extension="bin" ContentType='application/vnd.ms-office.vbaProject'/>
  <Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>
  <Override PartName="/customXml/item1.xml" ContentType='application/xml'/>
</Types>`)

		expect(contentTypes.defaults.get('xml')).toBe('application/xml')
		expect(contentTypes.defaults.get('bin')).toBe('application/vnd.ms-office.vbaProject')
		expect(contentTypes.overrides.get('xl/workbook.xml')).toBe(
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
		)
		expect(contentTypes.overrides.get('customXml/item1.xml')).toBe('application/xml')
		expect(contentTypes.defaultEntries).toEqual([
			{ extension: 'xml', contentType: 'application/xml' },
			{ extension: 'bin', contentType: 'application/vnd.ms-office.vbaProject' },
		])
		expect(contentTypes.overrideEntries).toEqual([
			{
				partPath: 'xl/workbook.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
			},
			{ partPath: 'customXml/item1.xml', contentType: 'application/xml' },
		])
	})

	test('keeps extension attributes on content type entries', () => {
		const contentTypes = parseContentTypes(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
  xmlns:ct2="urn:ascend:content-type-entry">
  <Default Extension="xml" ContentType="application/xml" ct2:role="generic"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" ct2:role="main" ct2:checksum="Tom&amp;Jane"/>
</Types>`)

		expect(contentTypes.defaultEntries).toEqual([
			{
				extension: 'xml',
				contentType: 'application/xml',
				extraAttributes: [{ name: 'ct2:role', value: 'generic' }],
			},
		])
		expect(contentTypes.overrideEntries).toEqual([
			{
				partPath: 'xl/workbook.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
				extraAttributes: [
					{ name: 'ct2:role', value: 'main' },
					{ name: 'ct2:checksum', value: 'Tom&Jane' },
				],
			},
		])
	})

	test('parses explicitly closed empty Default and Override elements', () => {
		const contentTypes = parseContentTypes(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"></Default>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml">
  </Override>
</Types>`)

		expect(contentTypes.defaultEntries).toEqual([
			{ extension: 'xml', contentType: 'application/xml' },
		])
		expect(contentTypes.overrideEntries).toEqual([
			{
				partPath: 'xl/workbook.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
			},
		])
	})

	test('parses namespace-prefixed content type entries', () => {
		const contentTypes = parseContentTypes(`<?xml version="1.0" encoding="UTF-8"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml">
  </ct:Override>
</ct:Types>`)

		expect(contentTypes.defaultEntries).toEqual([
			{ extension: 'xml', contentType: 'application/xml' },
		])
		expect(contentTypes.overrideEntries).toEqual([
			{
				partPath: 'xl/workbook.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
			},
		])
	})
})
