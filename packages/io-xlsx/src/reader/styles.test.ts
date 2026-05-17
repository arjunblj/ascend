import { describe, expect, test } from 'bun:test'
import { parseStyles, parseStylesLite } from './styles.ts'

describe('style inventory', () => {
	test('parses full and lite styles from prefixed SpreadsheetML style sheets', () => {
		const stylesXml = `<?xml version="1.0"?>
<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="m/d/yyyy"/></x:numFmts>
  <x:fonts count="2"><x:font/><x:font><x:b/></x:font></x:fonts>
  <x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FFC6EFCE"/></x:patternFill></x:fill></x:fills>
  <x:borders count="1"><x:border/></x:borders>
  <x:cellXfs count="2">
    <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <x:xf numFmtId="164" fontId="1" fillId="1" borderId="0"/>
  </x:cellXfs>
  <x:dxfs count="1"><x:dxf><x:font><x:i/></x:font></x:dxf></x:dxfs>
  <x:tableStyles count="1"><x:tableStyle name="TableStyleMedium2"/></x:tableStyles>
</x:styleSheet>`

		const full = parseStyles(stylesXml)
		expect(full.metadata).toMatchObject({
			numFmtCount: 1,
			fontCount: 2,
			fillCount: 2,
			borderCount: 1,
			cellXfCount: 2,
			dxfCount: 1,
			tableStyleCount: 1,
		})
		expect(full.cellStyles[1]).toMatchObject({
			font: { bold: true },
			fill: { pattern: 'solid', fgColor: { rgb: 'FFC6EFCE' } },
			numberFormat: 'm/d/yyyy',
		})
		expect(full.isDateFormat[1]).toBe(true)
		expect(full.differentialStyles[0]).toMatchObject({ font: { italic: true } })

		const lite = parseStylesLite(stylesXml)
		expect(lite.metadata).toMatchObject({
			numFmtCount: 1,
			fontCount: 2,
			fillCount: 2,
			borderCount: 1,
			cellXfCount: 2,
			dxfCount: 1,
			tableStyleCount: 1,
		})
		expect(lite.isDateFormat[1]).toBe(true)
	})
})
