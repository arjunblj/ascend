import { describe, expect, test } from 'bun:test'
import { parseXml } from './xml.ts'

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
