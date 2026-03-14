import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, Workbook } from '@ascend/core'
import { booleanValue, numberValue, stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { applyOperations } from '../../../engine/src/index.ts'
import { fingerprintXlsx } from '../../test/fidelity-harness.ts'
import { makeXlsx } from '../../test/helpers.ts'
import type { PreservationCapsule } from '../preserve.ts'
import { readXlsx } from '../reader/index.ts'
import { planWriteXlsx, writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

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

	it('round-trips array formula bindings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, {
			value: numberValue(3),
			formula: 'SUM(B1:B2)',
			styleId: S0,
			formulaInfo: { kind: 'array', ref: 'A1:A2' },
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(xml).toContain('<f t="array" ref="A1:A2">SUM(B1:B2)</f>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'array',
			ref: 'A1:A2',
		})
	})

	it('writes dynamic-array metadata and storage formula syntax', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Dynamic')
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(3)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(0, 1, {
			value: numberValue(6),
			formula: 'SUM(A1#)',
			styleId: S0,
		})
		sheet.cells.set(0, 2, {
			value: numberValue(1),
			formula: '@A1',
			styleId: S0,
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		const metadataXml = new TextDecoder().decode(zip['xl/metadata.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(contentTypes).toContain(
			'/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"',
		)
		expect(workbookRels).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata"',
		)
		expect(metadataXml).toContain('name="XLDAPR"')
		expect(sheetXml).toContain('cm="1"')
		expect(sheetXml).toContain('_xlfn.SEQUENCE(3)')
		expect(sheetXml).toContain('SUM(_xlfn.ANCHORARRAY(A1))')
		expect(sheetXml).toContain('_xlfn.SINGLE(A1)')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 1)?.formula).toBe('SUM(A1#)')
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 2)?.formula).toBe('@A1')
	})

	it('preserves sharedStrings.xml when string indices are unchanged', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2">
  <si><t>World</t></si>
  <si><t>Hello</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
      <c r="B1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const sourceZip = unzipSync(sourceBytes)
		const originalSharedStrings = new TextDecoder().decode(
			sourceZip['xl/sharedStrings.xml'] ?? new Uint8Array(),
		)

		const sheet = source.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		sheet?.cells.set(0, 1, { value: numberValue(99), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
			sharedStringsDirty: false,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sharedStrings = new TextDecoder().decode(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
		expect(sharedStrings).toBe(originalSharedStrings)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Hello',
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(1, 0)?.value).toEqual({
			kind: 'string',
			value: 'World',
		})
	})

	it('classifies generated and preserved parts in the write plan', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
			'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
			'xl/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme"/>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)

		const plan = planWriteXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(plan)

		const workbookPart = plan.value.descriptors.find((entry) => entry.path === 'xl/workbook.xml')
		const stylesPart = plan.value.descriptors.find((entry) => entry.path === 'xl/styles.xml')
		const themePart = plan.value.descriptors.find((entry) => entry.path === 'xl/theme/theme1.xml')
		const sheetPart = plan.value.descriptors.find(
			(entry) => entry.path === 'xl/worksheets/sheet1.xml',
		)
		expect(workbookPart?.origin).toBe('preserved-source')
		expect(stylesPart?.origin).toBe('preserved-source')
		expect(themePart?.origin).toBe('preserved-source')
		expect(sheetPart?.origin).toBe('generated')
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

	it('1904 date system round-trips correctly', () => {
		const wb = new Workbook()
		wb.calcSettings = { ...wb.calcSettings, dateSystem: '1904' }
		const dateFmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
		const sheet = wb.addSheet('Dates')
		sheet.cells.set(0, 0, {
			value: { kind: 'date', serial: 0 },
			formula: null,
			styleId: dateFmtId,
		})
		sheet.cells.set(1, 0, {
			value: { kind: 'date', serial: 1 },
			formula: null,
			styleId: dateFmtId,
		})
		sheet.cells.set(2, 0, {
			value: { kind: 'date', serial: 100 },
			formula: null,
			styleId: dateFmtId,
		})

		const { bytes, result } = roundTrip(wb)
		expect(result.workbook.calcSettings?.dateSystem).toBe('1904')
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'date', serial: 0 })
		expect(s?.cells.get(1, 0)?.value).toEqual({ kind: 'date', serial: 1 })
		expect(s?.cells.get(2, 0)?.value).toEqual({ kind: 'date', serial: 100 })

		const zip = unzipSync(bytes)
		const wbXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(wbXml).toContain('date1904="1"')
	})

	describe('number format edge cases', () => {
		it('round-trips custom number format #,##0.00', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '#,##0.00' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(1234.5), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('#,##0.00')
		})

		it('round-trips percentage format 0.00%', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '0.00%' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(0.125), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('0.00%')
		})

		it('round-trips currency format $#,##0', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '$#,##0' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(999), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('$#,##0')
		})

		it('round-trips accounting format _($* #,##0.00_)', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '_($* #,##0.00_)' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(-1234.56), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('_($* #,##0.00_)')
		})

		it('round-trips date format yyyy-mm-dd', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, {
				value: { kind: 'date', serial: 45292 },
				formula: null,
				styleId: fmtId,
			})

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('yyyy-mm-dd')
		})

		it('round-trips date format dd/mm/yyyy', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: 'dd/mm/yyyy' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, {
				value: { kind: 'date', serial: 44927 },
				formula: null,
				styleId: fmtId,
			})

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('dd/mm/yyyy')
		})

		it('round-trips negative format with color [Red]', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '#,##0.00;[Red]-#,##0.00' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(-100), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('#,##0.00;[Red]-#,##0.00')
		})

		it('round-trips text format @', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '@' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: stringValue('ID-001'), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('@')
		})
	})

	it('round-trips gradient fills', () => {
		const wb = new Workbook()
		const gradientId = wb.styles.register({
			fill: {
				gradient: {
					type: 'linear',
					degree: 45,
					stops: [
						{ position: 0, color: { kind: 'rgb', rgb: 'FFFF0000' } },
						{ position: 1, color: { kind: 'theme', theme: 1, tint: 0.25 } },
					],
				},
			},
		})
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: gradientId })

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'] ?? new Uint8Array())
		expect(stylesXml).toContain('<gradientFill type="linear" degree="45">')
		expect(stylesXml).toContain('<stop position="0">')
		expect(stylesXml).toContain('<color rgb="FFFF0000"/>')
		expect(stylesXml).toContain('<color theme="1" tint="0.25"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const style = reopened.value.workbook.styles.get(
			reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
		)
		expect(style?.fill).toEqual({
			gradient: {
				type: 'linear',
				degree: 45,
				stops: [
					{ position: 0, color: { kind: 'rgb', rgb: 'FFFF0000' } },
					{ position: 1, color: { kind: 'theme', theme: 1, tint: 0.25 } },
				],
			},
		})
	})

	it('appends number-format styles onto preserved styles.xml without rebuilding it', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
    <tableStyle name="TableStyleMedium2"/>
  </tableStyles>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>0.25</v></c></row></sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const applied = applyOperations(source.value.workbook, [
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'A1:A1', format: '0.0%' },
		])
		expectOk(applied)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'] ?? new Uint8Array())
		expect(stylesXml).toContain('<tableStyles count="1" defaultTableStyle="TableStyleMedium2"')
		expect(stylesXml).toContain('<cellStyleXfs count="1"><xf/></cellStyleXfs>')
		expect(stylesXml).toContain('formatCode="0.0%"')
		expect(stylesXml).toContain('<cellXfs count="2">')
		expect(stylesXml).toContain('applyNumberFormat="1"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const cell = reopened.value.workbook.sheets[0]?.cells.get(0, 0)
		const style = reopened.value.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.numberFormat).toBe('0.0%')
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
		sheet.autoFilter = {
			ref: 'A1:B10',
			columns: [],
		}
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
		expect(s?.autoFilter).toEqual({
			ref: 'A1:B10',
			columns: [],
		})
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

	it('preserves macro-enabled workbook content type when workbook capsules are present', () => {
		const wb = new Workbook()
		wb.preservedXml = {
			workbookXml:
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
			workbookRelsXml:
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
			contentType: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
		}
		wb.addSheet('Sheet1')
		const macroCapsule: PreservationCapsule = {
			partPath: 'xl/vbaProject.bin',
			contentType: 'application/vnd.ms-office.vbaProject',
			relationships: [],
			content: new Uint8Array([0xde, 0xad, 0xbe, 0xef]),
			anchor: { kind: 'workbook' },
			relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
		}

		const written = writeXlsx(wb, [macroCapsule])
		if (!written.ok) throw new Error(`write failed: ${written.error.message}`)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)

		expect(contentTypes).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml')
		expect(contentTypes).toContain('application/vnd.ms-office.vbaProject')
		expect(workbookRels).toContain('relationships/vbaProject')
		expect(zip['xl/vbaProject.bin']).toBeDefined()
	})

	it('preserves worksheet autoFilter criteria and sort state on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Filter')
		sheet.autoFilter = {
			ref: 'A1:B10',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					blank: true,
					values: ['Open', 'Closed'],
				},
			],
			sortState: {
				ref: 'A2:B10',
				caseSensitive: true,
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		}

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.autoFilter).toEqual({
			ref: 'A1:B10',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					blank: true,
					values: ['Open', 'Closed'],
				},
			],
			sortState: {
				ref: 'A2:B10',
				caseSensitive: true,
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			autoFilter: 1,
			filterColumn: 1,
			filters: 1,
			filter: 2,
			sortState: 1,
			sortCondition: 1,
		})
	})

	it('preserves workbook and sheet protection on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Protected')
		sheet.cells.set(0, 0, { value: stringValue('Locked'), formula: null, styleId: S0 })
		wb.workbookProtection = {
			lockStructure: true,
			workbookPassword: 'ABCD',
			workbookAlgorithmName: 'SHA-512',
			workbookSpinCount: 100000,
		}
		sheet.protection = {
			sheet: true,
			objects: true,
			scenarios: true,
			password: '1234',
			sort: false,
			autoFilter: false,
			selectUnlockedCells: true,
		}

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.workbookProtection).toEqual({
			lockStructure: true,
			workbookPassword: 'ABCD',
			workbookAlgorithmName: 'SHA-512',
			workbookSpinCount: 100000,
		})
		expect(result.workbook.sheets[0]?.protection).toEqual({
			sheet: true,
			objects: true,
			scenarios: true,
			password: '1234',
			sort: false,
			autoFilter: false,
			selectUnlockedCells: true,
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.workbook?.tagCounts).toMatchObject({
			workbookProtection: 1,
		})
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			sheetProtection: 1,
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

	it('preserves drawing and legacyDrawing references when sheet capsules exist', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Visuals')
		sheet.cells.set(0, 0, { value: stringValue('Chart host'), formula: null, styleId: S0 })
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: true }

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/drawings/drawing1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
				relationships: [],
				content: new TextEncoder().encode(
					'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
				),
				anchor: { kind: 'sheet', sheetName: 'Visuals' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
			{
				partPath: 'xl/drawings/vmlDrawing1.vml',
				contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
				relationships: [],
				content: new TextEncoder().encode('<xml xmlns:v="urn:schemas-microsoft-com:vml"/>'),
				anchor: { kind: 'sheet', sheetName: 'Visuals' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
			},
		]

		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: true,
			hasLegacyDrawing: true,
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			drawing: 1,
			legacyDrawing: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 2,
		})
	})

	it('writes classic comments with generated comments and VML parts', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Notes')
		sheet.cells.set(1, 1, { value: stringValue('Cell'), formula: null, styleId: S0 })
		sheet.comments.set('B2', { text: 'Hello', author: 'Ada' })

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.comments.get('B2')).toEqual({
			text: 'Hello',
			author: 'Ada',
		})
		expect(result.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: false,
			hasLegacyDrawing: true,
		})

		const entries = unzipSync(bytes)
		expect(entries['xl/comments1.xml']).toBeDefined()
		expect(entries['xl/drawings/vmlDrawing1.vml']).toBeDefined()
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			legacyDrawing: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 2,
		})
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

	it('preserves rich text inline strings when useSharedStrings is false (xlsx-2)', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, {
			value: {
				kind: 'richText',
				runs: [{ text: 'bold text', bold: true }, { text: ' normal' }],
			},
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(wb, undefined, { useSharedStrings: false })
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<r><rPr><b/></rPr><t>bold text</t></r>')
		expect(sheetXml).toContain('<r><t> normal</t></r>')
		expect(sheetXml).toContain('t="inlineStr"')
		expect(sheetXml).toContain('<is>')

		const read = readXlsx(written.value)
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'bold text', bold: true }, { text: ' normal' }],
		})
	})

	it('creates dxfId for CF rules with style but no dxfId (xlsx-4)', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Rules')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.conditionalFormats.push({
			sqref: 'A1:A5',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					priority: 1,
					formulas: ['3'],
					style: {
						font: { bold: true },
						fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
					},
				},
			],
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('dxfId="0"')
		expect(stylesXml).toContain('<dxfs count="1">')
		expect(stylesXml).toContain('<dxf>')
		expect(stylesXml).toContain('<b/>')
		expect(stylesXml).toContain('FFC6EFCE')

		const { result } = roundTrip(wb)
		const rule = result.workbook.sheets[0]?.conditionalFormats[0]?.rules[0]
		expect(rule?.dxfId).toBeDefined()
		expect(rule?.dxfId).toBeGreaterThanOrEqual(0)
		expect(rule?.style).toEqual({
			font: { bold: true },
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
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

	it('preserves richer table metadata on round-trip', () => {
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
			columns: [
				{ id: 1, name: 'Name', totalsRowLabel: 'Total', dataDxfId: 7 },
				{ id: 2, name: 'Qty', totalsRowFunction: 'sum', totalsRowDxfId: 8 },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:B2',
				columns: [],
				sortState: {
					ref: 'A2:B2',
					conditions: [{ ref: 'B2:B2' }],
				},
			},
			headerRowDxfId: 5,
			dataDxfId: 6,
			tableStyleInfo: {
				name: 'TableStyleMedium2',
				showFirstColumn: false,
				showLastColumn: false,
				showRowStripes: true,
				showColumnStripes: false,
			},
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.tables[0]).toEqual(
			expect.objectContaining({
				name: 'InventoryTable',
				headerRowDxfId: 5,
				dataDxfId: 6,
				tableStyleInfo: {
					name: 'TableStyleMedium2',
					showFirstColumn: false,
					showLastColumn: false,
					showRowStripes: true,
					showColumnStripes: false,
				},
				columns: [
					{ id: 1, name: 'Name', totalsRowLabel: 'Total', dataDxfId: 7 },
					{ id: 2, name: 'Qty', totalsRowFunction: 'sum', totalsRowDxfId: 8 },
				],
			}),
		)
		const entries = unzipSync(bytes)
		const tableEntry = entries['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('headerRowDxfId="5"')
		expect(tableXml).toContain('dataDxfId="6"')
		expect(tableXml).toContain('totalsRowFunction="sum"')
		expect(tableXml).toContain('tableStyleInfo')
	})

	it('round-trips table created via createTable op through XLSX', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: stringValue('Product'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Widget'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(100), formula: null, styleId: S0 })
		sheet.cells.set(2, 0, { value: stringValue('Gadget'), formula: null, styleId: S0 })
		sheet.cells.set(2, 1, { value: numberValue(200), formula: null, styleId: S0 })

		const applyResult = applyOperations(wb, [
			{
				op: 'createTable',
				sheet: 'Data',
				ref: 'A1:B3',
				name: 'SalesTable',
				hasHeaders: true,
			},
		])
		expectOk(applyResult)

		const { result } = roundTrip(wb)
		const roundTrippedSheet = result.workbook.sheets[0]
		expect(roundTrippedSheet?.tables).toHaveLength(1)
		const table = roundTrippedSheet?.tables[0]
		expect(table?.name).toBe('SalesTable')
		expect(table?.hasHeaders).toBe(true)
		expect(table?.ref).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } })
		expect(table?.columns).toHaveLength(2)
		expect(table?.columns[0]?.name).toBe('Product')
		expect(table?.columns[1]?.name).toBe('Amount')
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
		expect(chart?.content).toBeUndefined()
	})

	it('preserves threaded comments XML through read-write round-trip', () => {
		const threadedCommentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Comment text</text>
  </threadedComment>
</ThreadedComments>`

		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/threadedComments/threadedComment1.xml',
				contentType: 'application/vnd.ms-excel.threadedcomments+xml',
				relationships: [],
				content: new TextEncoder().encode(threadedCommentXml),
				anchor: { kind: 'sheet', sheetName: 'Sheet1' },
				relType: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
			},
		]

		const { bytes, result } = roundTrip(wb, capsules)

		const entries = unzipSync(bytes)
		expect(entries['xl/threadedComments/threadedComment1.xml']).toBeDefined()
		const decoded = new TextDecoder().decode(entries['xl/threadedComments/threadedComment1.xml'])
		expect(decoded).toContain('<ThreadedComments')
		expect(decoded).toContain('threadedComment ref="A1"')
		expect(decoded).toContain('<text>Comment text</text>')

		const tcCapsule = result.capsules.find(
			(c) => c.partPath === 'xl/threadedComments/threadedComment1.xml',
		)
		expect(tcCapsule).toBeDefined()
		expect(result.report.features.find((f) => f.feature === 'preservedThreadedComments')).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/threadedComments/threadedComment1.xml'],
			}),
		)
	})

	it('produces a valid ZIP file', () => {
		const wb = new Workbook()
		wb.addSheet('Empty')

		const written = writeXlsx(wb)
		expectOk(written)

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
		expectOk(written)

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

	it('drops calcChain and marks workbook stale after formula-affecting edits', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="0" calcCompleted="1" calcOnSave="0" calcId="191029"/>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>A1*2</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
			'xl/calcChain.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const applied = applyOperations(source.value.workbook, [
			{ op: 'setCells', sheet: 'Calc', updates: [{ ref: 'A1', value: 3 }] },
		])
		expectOk(applied)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Calc'],
			calcStateDirty: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		expect(zip['xl/calcChain.xml']).toBeUndefined()
		expect(workbookRels).not.toContain('relationships/calcChain')
		expect(contentTypes).not.toContain('calcChain+xml')
		expect(workbookXml).toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="0"')
		expect(workbookXml).toContain('forceFullCalc="1"')
	})

	it('preserves calcPr fidelity on clean round-trip', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="1" calcCompleted="0" calcOnSave="0" forceFullCalc="1" calcId="191029"/>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(workbookXml).toContain('calcMode="manual"')
		expect(workbookXml).toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="0"')
		expect(workbookXml).toContain('calcOnSave="0"')
		expect(workbookXml).toContain('forceFullCalc="1"')
		expect(workbookXml).toContain('calcId="191029"')
	})

	it('preserves sheetView attributes (zoomScale, showGridLines, showFormulas, rightToLeft, tabSelected, view) on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('View')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.sheetView = {
			zoomScale: 125,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageLayout',
		}

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.sheetView).toEqual({
			zoomScale: 125,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageLayout',
		})
	})

	it('preserves sheetPr tabColor and sheetFormatPr on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Colored')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.tabColor = { rgb: 'FF0000FF', theme: 4, tint: -0.25 }
		sheet.sheetFormatPr = { defaultRowHeight: 14.5, defaultColWidth: 10.0 }

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.tabColor).toEqual({ rgb: 'FF0000FF', theme: 4, tint: -0.25 })
		expect(s.sheetFormatPr).toEqual({ defaultRowHeight: 14.5, defaultColWidth: 10 })
	})

	it('preserves row and column outline metadata on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Outline')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.outlinePr = { summaryBelow: false, summaryRight: false, showOutlineSymbols: true }
		sheet.sheetFormatPr = { outlineLevelRow: 1, outlineLevelCol: 2 }
		sheet.rowDefs.set(1, { hidden: true, outlineLevel: 1 })
		sheet.rowDefs.set(2, { collapsed: true })
		sheet.colDefs.push({ min: 0, max: 0, hidden: true, outlineLevel: 2 })
		sheet.colDefs.push({ min: 1, max: 1, collapsed: true })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.outlinePr).toEqual({
			summaryBelow: false,
			summaryRight: false,
			showOutlineSymbols: true,
		})
		expect(s.sheetFormatPr).toEqual({ outlineLevelRow: 1, outlineLevelCol: 2 })
		expect(s.rowDefs.get(1)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(s.rowDefs.get(2)).toEqual({ collapsed: true })
		expect(s.colDefs).toContainEqual({ min: 0, max: 0, hidden: true, outlineLevel: 2 })
		expect(s.colDefs).toContainEqual({ min: 1, max: 1, collapsed: true })
	})

	it('preserves extLst blocks through round-trip when present', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('ExtLst')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.preservedExtLst = '<extLst><ext uri="{test}"><x14:sparklines/></ext></extLst>'

		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(written.error.message)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<extLst>')
		expect(sheetXml).toContain('x14:sparklines')

		const read = readXlsx(written.value)
		if (!read.ok) throw new Error(read.error.message)
		expect(read.value.workbook.sheets[0]?.preservedExtLst).toContain('<extLst>')
	})

	it('preserves sparkline extLst XML through read-write round-trip', () => {
		const sparklineExtLst = `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
    <x14:sparklineGroup>
      <x14:f>Sheet1!A1:A5</x14:f>
      <x14:sparklines>
        <x14:sparkline f="Sheet1!B1:B5"/>
      </x14:sparklines>
    </x14:sparklineGroup>
  </x14:sparklineGroups>
</ext></extLst>`

		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
${sparklineExtLst}
</worksheet>`

		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const read = readXlsx(bytes)
		if (!read.ok) throw new Error(read.error.message)
		const preserved = read.value.workbook.sheets[0]?.preservedExtLst
		expect(preserved).toBeDefined()
		expect(preserved).toContain('x14:sparklineGroups')
		expect(preserved).toContain('x14:sparkline')

		const written = writeXlsx(read.value.workbook, read.value.capsules)
		if (!written.ok) throw new Error(written.error.message)
		const zip = unzipSync(written.value)
		const writtenSheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(writtenSheetXml).toContain('<extLst>')
		expect(writtenSheetXml).toContain('x14:sparklineGroups')
		expect(writtenSheetXml).toContain('Sheet1!A1:A5')
	})

	it('emits dimension element for non-empty sheets', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Dim')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.cells.set(9, 4, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(written.error.message)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<dimension ref="A1:E10"/>')
	})

	it('preserves full ignoredError attributes on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Errors')
		sheet.cells.set(0, 0, { value: stringValue('123'), formula: null, styleId: S0 })
		sheet.ignoredErrors.push({
			sqref: 'A1:B2',
			numberStoredAsText: true,
			formula: true,
			evalError: true,
		})

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.ignoredErrors).toEqual([
			{
				sqref: 'A1:B2',
				numberStoredAsText: true,
				formula: true,
				evalError: true,
			},
		])
	})

	describe('round-trip fidelity', () => {
		it('round-trips mixed data types (numbers, strings, booleans, dates, errors)', () => {
			const wb = new Workbook()
			const dateFmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
			const sheet = wb.addSheet('Mixed')
			sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: S0 })
			sheet.cells.set(0, 1, { value: stringValue('text'), formula: null, styleId: S0 })
			sheet.cells.set(0, 2, { value: booleanValue(true), formula: null, styleId: S0 })
			sheet.cells.set(0, 3, { value: booleanValue(false), formula: null, styleId: S0 })
			sheet.cells.set(0, 4, {
				value: { kind: 'date', serial: 45292 },
				formula: null,
				styleId: dateFmtId,
			})
			sheet.cells.set(0, 5, {
				value: { kind: 'error', value: '#DIV/0!' },
				formula: null,
				styleId: S0,
			})

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 42 })
			expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'text' })
			expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'boolean', value: true })
			expect(s?.cells.get(0, 3)?.value).toEqual({ kind: 'boolean', value: false })
			expect(s?.cells.get(0, 4)?.value).toEqual({ kind: 'date', serial: 45292 })
			expect(s?.cells.get(0, 5)?.value).toEqual({ kind: 'error', value: '#DIV/0!' })
		})

		it('round-trips formulas with various complexity', () => {
			const wb = new Workbook()
			const sheet = wb.addSheet('Formulas')
			sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
			sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: S0 })
			sheet.cells.set(1, 0, { value: numberValue(30), formula: 'A1+B1', styleId: S0 })
			sheet.cells.set(1, 1, { value: numberValue(600), formula: 'A1*B1*2', styleId: S0 })
			sheet.cells.set(2, 0, { value: numberValue(3), formula: 'SUM(A1:B2)', styleId: S0 })
			sheet.cells.set(2, 1, {
				value: numberValue(2),
				formula: 'IF(A1>B1,"big","small")',
				styleId: S0,
			})

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			expect(s?.cells.get(1, 0)?.formula).toBe('A1+B1')
			expect(s?.cells.get(1, 1)?.formula).toBe('A1*B1*2')
			expect(s?.cells.get(2, 0)?.formula).toBe('SUM(A1:B2)')
			expect(s?.cells.get(2, 1)?.formula).toBe('IF(A1>B1,"big","small")')
		})

		it('round-trips styled cells (bold, italic, colors, number formats)', () => {
			const wb = new Workbook()
			const boldId = wb.styles.register({ font: { bold: true } })
			const italicId = wb.styles.register({ font: { italic: true } })
			const colorId = wb.styles.register({
				font: { color: { kind: 'rgb', rgb: 'FF0000FF' } },
			})
			const pctId = wb.styles.register({ numberFormat: '0.00%' })
			const sheet = wb.addSheet('Styled')
			sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })
			sheet.cells.set(0, 1, { value: stringValue('Italic'), formula: null, styleId: italicId })
			sheet.cells.set(0, 2, { value: stringValue('Blue'), formula: null, styleId: colorId })
			sheet.cells.set(0, 3, { value: numberValue(0.75), formula: null, styleId: pctId })

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			const styles = result.workbook.styles
			expect(styles.get(s?.cells.get(0, 0)?.styleId ?? S0)?.font?.bold).toBe(true)
			expect(styles.get(s?.cells.get(0, 1)?.styleId ?? S0)?.font?.italic).toBe(true)
			expect(styles.get(s?.cells.get(0, 2)?.styleId ?? S0)?.font?.color).toEqual({
				kind: 'rgb',
				rgb: 'FF0000FF',
			})
			expect(styles.get(s?.cells.get(0, 3)?.styleId ?? S0)?.numberFormat).toBe('0.00%')
		})

		it('round-trips merged cells', () => {
			const wb = new Workbook()
			const sheet = wb.addSheet('Merged')
			sheet.cells.set(0, 0, { value: stringValue('Title'), formula: null, styleId: S0 })
			sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } })
			sheet.merges.push({ start: { row: 2, col: 0 }, end: { row: 4, col: 1 } })

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			expect(s?.merges).toHaveLength(2)
			expect(s?.merges[0]).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } })
			expect(s?.merges[1]).toEqual({ start: { row: 2, col: 0 }, end: { row: 4, col: 1 } })
		})

		it('round-trips multiple sheets with cross-sheet references', () => {
			const wb = new Workbook()
			const data = wb.addSheet('Data')
			const summary = wb.addSheet('Summary')
			data.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: S0 })
			data.cells.set(1, 0, { value: numberValue(200), formula: null, styleId: S0 })
			summary.cells.set(0, 0, {
				value: numberValue(300),
				formula: 'SUM(Data!A1:A2)',
				styleId: S0,
			})

			const { result } = roundTrip(wb)
			expect(result.workbook.sheets).toHaveLength(2)
			expect(result.workbook.sheets[0]?.name).toBe('Data')
			expect(result.workbook.sheets[1]?.name).toBe('Summary')
			expect(result.workbook.sheets[1]?.cells.get(0, 0)?.formula).toBe('SUM(Data!A1:A2)')
			expect(result.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
				kind: 'number',
				value: 300,
			})
		})
	})
})
