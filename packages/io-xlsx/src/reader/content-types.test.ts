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
	})
})
