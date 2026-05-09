import { describe, expect, test } from 'bun:test'
import { escapeXml, parseXml } from './xml.ts'

describe('escapeXml', () => {
	test('returns safe strings unchanged and escapes XML-sensitive characters', () => {
		const safe = 'text-00001234'
		expect(escapeXml(safe)).toBe(safe)
		expect(escapeXml('A&B<"C">')).toBe('A&amp;B&lt;&quot;C&quot;&gt;')
	})
})

describe('parseXml', () => {
	test('normalizes mc:AlternateContent to fallback content', () => {
		const parsed = parseXml(`<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:AlternateContent>
    <mc:Choice Requires="x14">
      <value>choice</value>
    </mc:Choice>
    <mc:Fallback>
      <value>fallback</value>
    </mc:Fallback>
  </mc:AlternateContent>
</root>`)

		expect(parsed.root).toBeDefined()
		expect((parsed.root as { value?: string }).value).toBe('fallback')
	})
})
