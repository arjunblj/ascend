import { describe, expect, test } from 'bun:test'
import { parseSharedStringsBytes } from './shared-strings.ts'

describe('shared string table parsing', () => {
	test('byte parser falls back for prefixed SpreadsheetML shared strings', () => {
		const xml = `<?xml version="1.0"?>
<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <x:si><x:t>Alpha</x:t></x:si>
  <x:si><x:r><x:t>Beta</x:t></x:r></x:si>
</x:sst>`

		const sharedStrings = parseSharedStringsBytes(new TextEncoder().encode(xml))

		expect(sharedStrings.count).toBe(2)
		expect(sharedStrings.getString?.(0)).toBe('Alpha')
		expect(sharedStrings.get(1)).toEqual({ kind: 'string', value: 'Beta' })
	})
})
