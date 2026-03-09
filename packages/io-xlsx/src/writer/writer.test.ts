import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, Workbook } from '@ascend/core'
import { booleanValue, numberValue, stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { fingerprintXlsx } from '../../test/fidelity-harness.ts'
import type { PreservationCapsule } from '../preserve.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function roundTrip(wb: Workbook, capsules?: PreservationCapsule[]) {
	const written = writeXlsx(wb, capsules)
	if (!written.ok) throw new Error(`write failed: ${written.error.message}`)
	const read = readXlsx(written.value)
	if (!read.ok) throw new Error(`read failed: ${read.error.message}`)
	return { bytes: written.value, result: read.value }
}

describe('writeXlsx', () => {
	it('round-trips cell values correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: booleanValue(true), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(3.14), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: booleanValue(false), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s).toBeDefined()
		expect(s?.name).toBe('Test')
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 42 })
		expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'boolean', value: true })
		expect(s?.cells.get(1, 0)?.value).toEqual({ kind: 'number', value: 3.14 })
		expect(s?.cells.get(1, 1)?.value).toEqual({ kind: 'boolean', value: false })
	})

	it('round-trips multiple sheets', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('First')
		s1.cells.set(0, 0, { value: stringValue('A'), formula: null, styleId: S0 })
		const s2 = wb.addSheet('Second')
		s2.cells.set(0, 0, { value: numberValue(99), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)

		expect(result.workbook.sheets).toHaveLength(2)
		expect(result.workbook.sheets[0]?.name).toBe('First')
		expect(result.workbook.sheets[1]?.name).toBe('Second')
		expect(result.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'A',
		})
		expect(result.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
	})

	it('preserves formula text on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: 'A1*2', styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(30), formula: 'SUM(A1,B1)', styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 1)?.formula).toBe('A1*2')
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 20 })
		expect(s?.cells.get(1, 0)?.formula).toBe('SUM(A1,B1)')
	})

	it('preserves bold style on round-trip', () => {
		const wb = new Workbook()
		const boldId = wb.styles.register({ font: { bold: true } })
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell = s?.cells.get(0, 0)
		expect(cell).toBeDefined()
		const style = result.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.font?.bold).toBe(true)
	})

	it('preserves number format on round-trip', () => {
		const wb = new Workbook()
		const pctId = wb.styles.register({ numberFormat: '0.00%' })
		const sheet = wb.addSheet('Fmt')
		sheet.cells.set(0, 0, { value: numberValue(0.75), formula: null, styleId: pctId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell = s?.cells.get(0, 0)
		expect(cell).toBeDefined()
		const style = result.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.numberFormat).toBe('0.00%')
	})

	it('preserves merges on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Merges')
		sheet.cells.set(0, 0, { value: stringValue('Merged'), formula: null, styleId: S0 })
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } })
		sheet.merges.push({ start: { row: 3, col: 0 }, end: { row: 3, col: 1 } })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.merges).toHaveLength(2)
		expect(s?.merges[0]).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } })
		expect(s?.merges[1]).toEqual({ start: { row: 3, col: 0 }, end: { row: 3, col: 1 } })
	})

	it('preserves worksheet layout metadata on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Layout')
		sheet.cells.set(0, 0, { value: stringValue('Header'), formula: null, styleId: S0 })
		sheet.frozenRows = 1
		sheet.frozenCols = 2
		sheet.colWidths.set(0, 18.5)
		sheet.colWidths.set(1, 18.5)
		sheet.rowHeights.set(0, 24)
		sheet.autoFilter = 'A1:B10'
		sheet.pageMargins = {
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		}
		sheet.pageSetup = {
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 2,
		}
		sheet.printOptions = {
			gridLines: true,
			headings: true,
		}
		sheet.headerFooter = {
			oddHeader: '&LTest',
			oddFooter: '&R1',
		}

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.frozenRows).toBe(1)
		expect(s?.frozenCols).toBe(2)
		expect(s?.colWidths.get(0)).toBe(18.5)
		expect(s?.colWidths.get(1)).toBe(18.5)
		expect(s?.rowHeights.get(0)).toBe(24)
		expect(s?.autoFilter).toBe('A1:B10')
		expect(s?.pageMargins).toEqual({
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		})
		expect(s?.pageSetup).toEqual({
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 2,
		})
		expect(s?.printOptions).toEqual({
			gridLines: true,
			headings: true,
		})
		expect(s?.headerFooter).toEqual({
			oddHeader: '&LTest',
			oddFooter: '&R1',
		})
	})

	it('preserves hyperlinks on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Links')
		sheet.cells.set(0, 0, { value: stringValue('Docs'), formula: null, styleId: S0 })
		sheet.hyperlinks.set('A1', {
			target: 'https://example.com/docs',
			display: 'Docs',
			tooltip: 'Open docs',
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.hyperlinks.get('A1')).toEqual({
			target: 'https://example.com/docs',
			display: 'Docs',
			tooltip: 'Open docs',
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			hyperlinks: 1,
			hyperlink: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.normalized).toContain(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
		)
		expect(fingerprint.sheetRels[0]?.xml.normalized).toContain('TargetMode="External"')
	})

	it('preserves conditional formatting and data validations on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Rules')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: S0 })
		wb.differentialStyles.push({
			font: { bold: true },
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
		})
		sheet.conditionalFormats.push({
			sqref: 'A1:A10',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					dxfId: 0,
					priority: 1,
					stopIfTrue: true,
					formulas: ['3'],
					style: wb.differentialStyles[0],
				},
			],
		})
		sheet.dataValidations.push({
			sqref: 'B2:B4',
			type: 'list',
			allowBlank: true,
			showInputMessage: true,
			formula1: '"Q1,Q2,Q3"',
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A10',
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						dxfId: 0,
						priority: 1,
						stopIfTrue: true,
						formulas: ['3'],
						style: {
							font: { bold: true },
							fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
						},
					},
				],
			},
		])
		expect(result.workbook.sheets[0]?.dataValidations).toEqual([
			{
				sqref: 'B2:B4',
				type: 'list',
				allowBlank: true,
				showInputMessage: true,
				formula1: '"Q1,Q2,Q3"',
			},
		])
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.styles?.tagCounts).toMatchObject({
			dxfs: 1,
			dxf: 1,
		})
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			conditionalFormatting: 1,
			cfRule: 1,
			dataValidations: 1,
			dataValidation: 1,
			formula: 1,
			formula1: 1,
		})
	})

	it('preserves defined names on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.definedNames.set('Total', 'Data!$A$1')

		const { result } = roundTrip(wb)
		expect(result.workbook.definedNames.get('Total')).toBe('Data!$A$1')
	})

	it('preserves sheet-scoped defined names on round-trip', () => {
		const wb = new Workbook()
		const data = wb.addSheet('Data')
		const summary = wb.addSheet('Summary')
		data.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		summary.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })
		wb.definedNames.set('Budget', 'Summary!$A$1', { kind: 'sheet', sheetId: summary.id })

		const { result } = roundTrip(wb)
		const resolved = result.workbook.definedNames.resolve(
			'Budget',
			result.workbook.getSheet('Summary')?.id,
		)
		expect(resolved?.scope.kind).toBe('sheet')
		expect(resolved?.formula).toBe('Summary!$A$1')
	})

	it('preserves workbook views and external reference wiring on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.workbookViews.push({ activeTab: 1, firstSheet: 0, visibility: 'visible', tabRatio: 600 })
		wb.workbookProperties = { codeName: 'Model', filterPrivacy: true }
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
				relationships: [],
				content: new TextEncoder().encode(
					'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
				),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			},
		]

		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.workbookViews).toEqual([
			{ activeTab: 1, firstSheet: 0, visibility: 'visible', tabRatio: 600 },
		])
		expect(result.workbook.externalReferences).toEqual(['xl/externalLinks/externalLink1.xml'])

		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.workbook?.tagCounts).toMatchObject({
			bookViews: 1,
			workbookView: 1,
			externalReferences: 1,
			externalReference: 1,
			calcPr: 1,
		})
	})

	it('preserves workbook theme parts on round-trip', () => {
		const wb = new Workbook()
		const themedStyle = wb.styles.register({
			font: { color: { kind: 'theme', theme: 4, tint: -0.25 } },
			fill: { pattern: 'solid', fgColor: { kind: 'theme', theme: 5 } },
		})
		const sheet = wb.addSheet('Theme')
		sheet.cells.set(0, 0, { value: stringValue('Brand'), formula: null, styleId: themedStyle })
		wb.preservedTheme = {
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
  <a:themeElements>
    <a:clrScheme name="Brand">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="123456"/></a:dk2>
      <a:lt2><a:srgbClr val="F2F2F2"/></a:lt2>
      <a:accent1><a:srgbClr val="004488"/></a:accent1>
      <a:accent2><a:srgbClr val="D64545"/></a:accent2>
      <a:accent3><a:srgbClr val="4CAF50"/></a:accent3>
      <a:accent4><a:srgbClr val="7E57C2"/></a:accent4>
      <a:accent5><a:srgbClr val="00ACC1"/></a:accent5>
      <a:accent6><a:srgbClr val="FB8C00"/></a:accent6>
      <a:hlink><a:srgbClr val="1A73E8"/></a:hlink>
      <a:folHlink><a:srgbClr val="7B1FA2"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Brand Fonts">
      <a:majorFont><a:latin typeface="Inter Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Inter"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Brand Formats"/>
  </a:themeElements>
</a:theme>`,
		}
		wb.themeMetadata = {
			name: 'Custom Theme',
			colorSchemeName: 'Brand',
			colorCount: 12,
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		}

		const { result, bytes } = roundTrip(wb)
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.partPaths).toContain('xl/theme/theme1.xml')
		expect(result.workbook.preservedTheme).toEqual({
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
			xml: expect.stringContaining('Custom Theme'),
		})
		expect(result.workbook.themeMetadata).toEqual({
			name: 'Custom Theme',
			colorSchemeName: 'Brand',
			colorCount: 12,
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		})
		const style = result.workbook.styles.get(
			result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
		)
		expect(style?.font?.color).toEqual({ kind: 'theme', theme: 4, tint: -0.25 })
		expect(style?.fill?.fgColor).toEqual({ kind: 'theme', theme: 5 })
		expect(fingerprint.workbookRels?.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 4,
		})
		expect(fingerprint.contentTypes?.normalized).toContain('/xl/theme/theme1.xml')
	})

	it('preserves table-part sheet wiring when table capsules are present', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Balance')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/tables/table1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
				relationships: [],
				content: new TextEncoder().encode(`<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="BalanceTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
</table>`),
				anchor: { kind: 'sheet', sheetName: 'Balance' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			},
		]

		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.sheets[0]?.tables).toHaveLength(1)
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			tableParts: 1,
			tablePart: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 1,
		})
	})

	it('emits table parts for semantic tables without capsules', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Inventory')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Qty'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Bolts'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'InventoryTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Qty' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.tables).toHaveLength(1)
		expect(result.workbook.sheets[0]?.tables[0]?.name).toBe('InventoryTable')

		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.partPaths).toContain('xl/tables/table1.xml')
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			tableParts: 1,
			tablePart: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 1,
		})

		const entries = unzipSync(bytes)
		const tableEntry = entries['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('InventoryTable')
		expect(tableXml).toContain('<tableColumns count="2">')
	})

	it('preserves capsule parts through write-read cycle', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })

		const capsuleContent = new TextEncoder().encode('<chart>test chart data</chart>')
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/charts/chart1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
				relationships: [],
				content: capsuleContent,
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
			},
		]

		const { bytes, result } = roundTrip(wb, capsules)

		const entries = unzipSync(bytes)
		expect(entries['xl/charts/chart1.xml']).toBeDefined()

		const decoded = new TextDecoder().decode(entries['xl/charts/chart1.xml'])
		expect(decoded).toBe('<chart>test chart data</chart>')

		const readCapsules = result.capsules
		const chart = readCapsules.find((c) => c.partPath === 'xl/charts/chart1.xml')
		expect(chart).toBeDefined()
		expect(chart?.contentType).toContain('chart')
		expect(new TextDecoder().decode(chart?.content)).toBe('<chart>test chart data</chart>')
	})

	it('produces a valid ZIP file', () => {
		const wb = new Workbook()
		wb.addSheet('Empty')

		const written = writeXlsx(wb)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const entries = unzipSync(written.value)
		expect(entries['[Content_Types].xml']).toBeDefined()
		expect(entries['_rels/.rels']).toBeDefined()
		expect(entries['xl/workbook.xml']).toBeDefined()
		expect(entries['xl/_rels/workbook.xml.rels']).toBeDefined()
		expect(entries['xl/styles.xml']).toBeDefined()
		expect(entries['xl/worksheets/sheet1.xml']).toBeDefined()
	})

	it('emits a stable structure fingerprint for synthetic workbooks', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Report')
		const percentId = wb.styles.register({
			font: { bold: true },
			numberFormat: '0.0%',
		})
		sheet.cells.set(0, 0, { value: stringValue('Revenue'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(0.25), formula: null, styleId: percentId })
		sheet.cells.set(1, 1, { value: numberValue(0.5), formula: 'B1*2', styleId: percentId })
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } })
		wb.definedNames.set('RevenuePct', 'Report!$B$1')
		wb.calcSettings = {
			...wb.calcSettings,
			calcMode: 'manual',
			fullCalcOnLoad: true,
		}

		const written = writeXlsx(wb)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const fingerprint = fingerprintXlsx(written.value)
		expect(fingerprint.partPaths).toEqual([
			'[Content_Types].xml',
			'_rels/.rels',
			'docProps/app.xml',
			'docProps/core.xml',
			'xl/_rels/workbook.xml.rels',
			'xl/sharedStrings.xml',
			'xl/styles.xml',
			'xl/workbook.xml',
			'xl/worksheets/sheet1.xml',
		])
		expect(fingerprint.workbook?.tagCounts).toMatchObject({
			workbook: 1,
			sheets: 1,
			sheet: 1,
			definedNames: 1,
			definedName: 1,
			calcPr: 1,
		})
		expect(fingerprint.workbookRels?.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 3,
		})
		expect(fingerprint.styles?.tagCounts).toMatchObject({
			styleSheet: 1,
			fonts: 1,
			fills: 1,
			borders: 1,
			numFmts: 1,
			numFmt: 1,
			cellXfs: 1,
			xf: 2,
		})
		expect(fingerprint.sheets).toHaveLength(1)
		expect(fingerprint.sheets[0]).toEqual(
			expect.objectContaining({
				path: 'xl/worksheets/sheet1.xml',
				xml: expect.objectContaining({
					tagCounts: expect.objectContaining({
						worksheet: 1,
						sheetData: 1,
						row: 2,
						c: 3,
						f: 1,
						mergeCells: 1,
						mergeCell: 1,
					}),
				}),
			}),
		)
	})
})
