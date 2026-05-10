import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, Workbook } from '@ascend/core'
import { booleanValue, errorValue, numberValue, stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { applyOperations } from '../../../engine/src/index.ts'
import { fingerprintXlsx } from '../../test/fidelity-harness.ts'
import { makeXlsx } from '../../test/helpers.ts'
import type { PreservationCapsule } from '../preserve.ts'
import { readXlsx } from '../reader/index.ts'
import { writeDenseRowsXlsx, writeDenseRowsXlsxStreaming } from './dense-rows.ts'
import { planWriteXlsx, writeXlsx, writeXlsxStreaming } from './index.ts'
import { updatePivotTableDefinitionXml } from './pivot-table.ts'

const S0 = 0 as StyleId

function expectOk<T, E>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(formatTestError(result.error))
}

function formatTestError(error: unknown): string {
	if (error && typeof error === 'object') {
		const candidate = error as { message?: unknown; errors?: Array<{ message?: unknown }> }
		if (typeof candidate.message === 'string') return candidate.message
		if (Array.isArray(candidate.errors)) {
			const messages = candidate.errors
				.map((entry) => (typeof entry?.message === 'string' ? entry.message : null))
				.filter((entry): entry is string => entry !== null)
			if (messages.length > 0) return messages.join('; ')
		}
	}
	return 'Unknown test error'
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

	it('writes edited tabular slicer cache item state back into preserved package XML', () => {
		const slicerCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_State" sourceName="State">
  <pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data>
</slicerCacheDefinition>`
		const wb = new Workbook()
		wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 5,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			sourceName: 'State',
			pivotCacheId: 5,
			pivotTableNames: ['PivotTable1'],
			items: [{ index: 1 }, { index: 0, selected: true }],
		})
		const applied = applyOperations(wb, [
			{
				op: 'setSlicerCacheItem',
				slicerCache: 'Slicer_State',
				item: 1,
				selected: true,
				noData: true,
			},
			{
				op: 'setSlicerCacheItem',
				slicerCache: 'Slicer_State',
				item: 0,
				selected: null,
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				contentType: 'application/vnd.ms-excel.slicerCache+xml',
				relationships: [],
				content: new TextEncoder().encode(slicerCacheXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(
			zip['xl/slicerCaches/slicerCache1.xml'] ?? new Uint8Array(),
		)

		expect(xml).toContain('<i x="1" s="1" nd="1"/>')
		expect(xml).toContain('<i x="0"/>')
		expect(xml).toContain('<items count="2">')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.slicerCaches[0]?.items).toEqual([
			{ index: 1, selected: true, noData: true },
			{ index: 0 },
		])
	})

	it('writes edited pivot field item state back into preserved pivot XML', () => {
		const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <pivotFields count="1"><pivotField axis="axisPage" showAll="0"><items count="3"><item h="1" x="0"/><item x="1" sd="0"></item><item t="default"/></items></pivotField></pivotFields>
  <pageFields count="1"><pageField fld="0"></pageField></pageFields>
</pivotTableDefinition>`
		const pivotCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">
  <cacheFields count="1"><cacheField name="Region"><sharedItems count="2"><s v="West"/><s v="East"/></sharedItems></cacheField></cacheFields>
</pivotCacheDefinition>`
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					showAll: false,
					items: [
						{ index: 0, cacheIndex: 0, hidden: true },
						{ index: 1, cacheIndex: 1, showDetails: false },
						{ index: 2, itemType: 'default' },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0 }],
			dataFields: [],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			fields: [
				{
					index: 0,
					name: 'Region',
					sharedItems: [
						{ index: 0, kind: 'string', value: 'West' },
						{ index: 1, kind: 'string', value: 'East' },
					],
				},
			],
		})
		const applied = applyOperations(wb, [
			{
				op: 'setPivotFieldItem',
				pivotTable: 'PivotTable1',
				fieldIndex: 0,
				itemIndex: 0,
				hidden: null,
				manualFilter: true,
				selectedPageItem: 1,
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotTableXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: 'Sheet1' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotCacheXml),
				anchor: { kind: 'workbook' },
				relType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const pivotXml = new TextDecoder().decode(
			zip['xl/pivotTables/pivotTable1.xml'] ?? new Uint8Array(),
		)
		const cacheXml = new TextDecoder().decode(
			zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array(),
		)

		expect(pivotXml).toContain('<item x="0" s="1"/>')
		expect(pivotXml).toContain('<item x="1" sd="0"></item>')
		expect(pivotXml).toContain('<pageField fld="0" item="1"></pageField>')
		expect(cacheXml).toContain('<pivotCacheDefinition')
		expect(cacheXml).toContain('refreshOnLoad="1"')
		expect(cacheXml).toContain('invalid="1"')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.pivotTables[0]?.fields[0]?.items).toEqual([
			{ index: 0, cacheIndex: 0, manualFilter: true },
			{ index: 1, cacheIndex: 1, showDetails: false },
			{ index: 2, itemType: 'default' },
		])
		expect(reopened.value.workbook.pivotTables[0]?.pageFields).toEqual([{ index: 0, item: 1 }])
	})

	it('scopes pivot item rewrites to direct main pivot fields', () => {
		const pivotTableXml = `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="PivotTable1" cacheId="34">
  <pivotFields count="3">
    <pivotField><items count="1"><item x="0"/></items><extLst><ext><x14:pivotField><x14:items><x14:item x="99" h="1"/></x14:items></x14:pivotField></ext></extLst></pivotField>
    <pivotField><items count="1"><item x="1"/></items></pivotField>
    <pivotField><items count="1"><item x="2"/></items></pivotField>
  </pivotFields>
</pivotTableDefinition>`
		const updated = updatePivotTableDefinitionXml(pivotTableXml, {
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{ index: 0 },
				{ index: 1 },
				{ index: 2, items: [{ index: 0, cacheIndex: 2, manualFilter: true }] },
			],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})

		expect(updated).toContain('<x14:item x="99" h="1"/>')
		expect(updated).toContain('<pivotField><items count="1"><item x="1"/></items></pivotField>')
		expect(updated).toContain(
			'<pivotField><items count="1"><item x="2" s="1"/></items></pivotField>',
		)
	})

	it('preserves unchanged pivot and slicer capsules byte-for-byte', () => {
		const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <pivotFields count="1"><pivotField axis="axisPage" showAll="false"><items count="2"><item h="true" x="0"></item><item t="default"/></items></pivotField></pivotFields>
  <pageFields count="1"><pageField fld="0"></pageField></pageFields>
</pivotTableDefinition>`
		const pivotCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="true" invalid="false" saveData="true">
  <cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1:B2"/></cacheSource>
</pivotCacheDefinition>`
		const slicerCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_State" sourceName="State">
  <pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="34"><items count="2"><i s="true" x="0"></i><i x="1" nd="false"/></items></tabular></data>
</slicerCacheDefinition>`
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					showAll: false,
					items: [
						{ index: 0, cacheIndex: 0, hidden: true },
						{ index: 1, itemType: 'default' },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0 }],
			dataFields: [],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			refreshOnLoad: true,
			invalid: false,
			saveData: true,
			sourceSheet: 'Data',
			sourceRef: 'A1:B2',
			fields: [],
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			sourceName: 'State',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
			items: [
				{ index: 0, selected: true },
				{ index: 1, noData: false },
			],
		})
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotTableXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: 'Sheet1' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotCacheXml),
				anchor: { kind: 'workbook' },
				relType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			},
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				contentType: 'application/vnd.ms-excel.slicerCache+xml',
				relationships: [],
				content: new TextEncoder().encode(slicerCacheXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const decodePart = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(decodePart('xl/pivotTables/pivotTable1.xml')).toBe(pivotTableXml)
		expect(decodePart('xl/pivotCache/pivotCacheDefinition1.xml')).toBe(pivotCacheXml)
		expect(decodePart('xl/slicerCaches/slicerCache1.xml')).toBe(slicerCacheXml)
	})

	it('writes edited drawing text back into preserved drawing XML', () => {
		const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="10" name="Callout" descr="Revenue note"/></xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>Revenue </a:t></a:r><a:r><a:t>up</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: false }
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 10,
			name: 'Callout',
			description: 'Revenue note',
			text: 'Revenue up',
			anchor: {
				kind: 'twoCell',
				editAs: 'oneCell',
				from: { row: 2, col: 1 },
				to: { row: 4, col: 3 },
			},
		})
		const applied = applyOperations(wb, [
			{
				op: 'setDrawingText',
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 10,
				text: 'Revenue flat',
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/drawings/drawing1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
				relationships: [],
				content: new TextEncoder().encode(drawingXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: 'Sheet1' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/drawings/drawing1.xml'] ?? new Uint8Array())

		expect(xml).toContain('<xdr:twoCellAnchor editAs="oneCell">')
		expect(xml).toContain('<xdr:cNvPr id="10" name="Callout" descr="Revenue note"/>')
		expect(xml).toContain('<a:rPr b="1"/>')
		expect(xml).toContain('<a:t>Revenue flat</a:t>')
		expect(xml).toContain('<a:t></a:t>')
		expect(xml).not.toContain('Revenue </a:t>')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.drawingObjectRefs[0]).toMatchObject({
			kind: 'textBox',
			id: 10,
			name: 'Callout',
			text: 'Revenue flat',
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

	it('preserves original stored formula text when unrelated edits dirty the sheet', () => {
		const storedFormula = `_xlfn.LET(
  _xlpm.value, A1,
  SUM(--(_xlfn.VSTACK(_xlpm.value, 2)))
)`
		const source = makeXlsx({
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
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>${storedFormula}</f><v>3</v></c>
      <c r="C1"><v>4</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})
		const read = readXlsx(source)
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 1)?.formula).not.toContain('_xlfn.')
		expect(sheet?.storedFormulaText.get('0:1')).toBe(storedFormula)
		sheet?.cells.set(0, 2, { value: numberValue(5), formula: null, styleId: S0 })

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(xml).toContain(`<f>${storedFormula}</f>`)
		expect(xml).toContain('<c r="C1"><v>5</v></c>')
	})

	it('round-trips shared formula bindings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Shared')
		sheet.cells.set(0, 1, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(84), formula: null, styleId: S0 })
		sheet.cells.set(0, 0, {
			value: numberValue(84),
			formula: 'B1*2',
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'A1',
				ref: 'A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(168),
			formula: null,
			styleId: S0,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(xml).toContain('t="shared" si="0"')
		expect(xml).toContain('ref="A1:A2"')
		expect(xml).toContain('>B1*2</f>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const s = reopened.value.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(s?.cells.get(0, 0)?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'A1',
		})
		expect(s?.cells.get(1, 0)?.formula).toBeNull()
		expect(s?.cells.get(1, 0)?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A1',
		})
	})

	it('splits oversized shared formulas into explicit formulas per cell', () => {
		const wb = new Workbook()
		const other = wb.addSheet('Other')
		other.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: S0 })
		other.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: S0 })
		const sheet = wb.addSheet('SharedLong')
		const longFormula = Array.from({ length: 1000 }, () => 'Other!B1').join('+')
		sheet.cells.set(0, 0, {
			value: numberValue(1800),
			formula: longFormula,
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'A1',
				ref: 'A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(2700),
			formula: null,
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: false,
				masterRef: 'A1',
			},
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
		expect(xml).not.toContain('t="shared"')
		expect(xml).toContain('<f>Other!B1+Other!B1')
		expect(xml).toContain('<f>Other!B2+Other!B2')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[1]?.cells.get(0, 0)?.formula).toBe(longFormula)
		expect(
			reopened.value.workbook.sheets[1]?.cells.get(1, 0)?.formula?.startsWith('Other!B2'),
		).toBe(true)
	})

	it('round-trips regular formulas that could be shared, preserving formula text', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: S0 })
		sheet.cells.set(0, 0, { value: numberValue(20), formula: 'B1*2', styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(40), formula: 'B2*2', styleId: S0 })
		sheet.cells.set(2, 0, { value: numberValue(60), formula: 'B3*2', styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(s?.cells.get(1, 0)?.formula).toBe('B2*2')
		expect(s?.cells.get(2, 0)?.formula).toBe('B3*2')
		expect(s?.cells.get(0, 0)?.value).toMatchObject({ kind: 'number', value: 20 })
		expect(s?.cells.get(1, 0)?.value).toMatchObject({ kind: 'number', value: 40 })
		expect(s?.cells.get(2, 0)?.value).toMatchObject({ kind: 'number', value: 60 })
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

	it('preserves formula spacing when storage rewriting is a no-op', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, {
			value: errorValue('#REF!'),
			formula: 'CHOOSE(#REF!, 1)',
			styleId: S0,
		})
		sheet.cells.set(0, 1, {
			value: numberValue(1),
			formula: 'COUNT(2, "A",  "",#REF!, #DIV/0!)',
			styleId: S0,
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<f>CHOOSE(#REF!, 1)</f>')
		expect(sheetXml).toContain('<f>COUNT(2, &quot;A&quot;,  &quot;&quot;,#REF!, #DIV/0!)</f>')
	})

	it('writes all dynamic-array metadata records referenced by cells', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Dynamic')
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(2)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(0, 2, {
			value: numberValue(10),
			formula: 'SEQUENCE(2,1,10)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 2, collapsed: true },
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const metadataXml = new TextDecoder().decode(zip['xl/metadata.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(metadataXml).toContain('<futureMetadata name="XLDAPR" count="2">')
		expect(metadataXml).toContain('<cellMetadata count="2">')
		expect(metadataXml).toContain('v="0"')
		expect(metadataXml).toContain('v="1"')
		expect(metadataXml).toContain('fCollapsed="1"')
		expect(sheetXml).toContain('<c r="A1" cm="1"')
		expect(sheetXml).toContain('<c r="C1" cm="2"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 2)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 2,
			collapsed: true,
		})
	})

	it('regenerates preserved dynamic-array metadata when edits add records', () => {
		const source = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Dynamic" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/metadata.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" cm="1"><f>_xlfn.SEQUENCE(2)</f><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const read = readXlsx(source)
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		sheet?.cells.set(0, 2, {
			value: numberValue(10),
			formula: 'SEQUENCE(2,1,10)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 2, collapsed: true },
		})

		const written = writeXlsx(read.value.workbook, undefined, { dirtySheetNames: ['Dynamic'] })
		expectOk(written)
		const zip = unzipSync(written.value)
		const metadataXml = new TextDecoder().decode(zip['xl/metadata.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(metadataXml).toContain('<futureMetadata name="XLDAPR" count="2">')
		expect(metadataXml).toContain('<cellMetadata count="2">')
		expect(metadataXml).toContain('fCollapsed="1"')
		expect(sheetXml).toContain('<c r="C1" cm="2"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 2)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 2,
			collapsed: true,
		})
	})

	it('round-trips _xlfn. prefix for future functions', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Xlfn')
		const formulas = [
			'IFS(A1>0,"pos",A1<0,"neg",TRUE,"zero")',
			'TEXTJOIN(",",TRUE,A1:A3)',
			'SWITCH(A1,1,"one",2,"two","other")',
			'MAXIFS(A1:A3,B1:B3,">0")',
			'CONCAT(A1,B1)',
			'IFNA(A1,0)',
			'XOR(TRUE,FALSE)',
		]
		for (const [i, formula] of formulas.entries()) {
			sheet.cells.set(i, 0, {
				value: stringValue(''),
				formula,
				styleId: S0,
			})
		}

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(sheetXml).toContain('_xlfn.IFS(')
		expect(sheetXml).toContain('_xlfn.TEXTJOIN(')
		expect(sheetXml).toContain('_xlfn.SWITCH(')
		expect(sheetXml).toContain('_xlfn.MAXIFS(')
		expect(sheetXml).toContain('_xlfn.CONCAT(')
		expect(sheetXml).toContain('_xlfn.IFNA(')
		expect(sheetXml).toContain('_xlfn.XOR(')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const cells = reopened.value.workbook.sheets[0]?.cells
		for (let i = 0; i < formulas.length; i++) {
			expect(cells?.get(i, 0)?.formula).toBe(formulas[i])
		}
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
		expect(sharedStrings).toContain('<t>World</t>')
		expect(sharedStrings).toContain('<t>Hello</t>')
		expect(sharedStrings).not.toBe('')
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

	it('keeps shared string positions when preserved entries contain duplicates', () => {
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
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
  <si><t>Beta</t></si>
  <si><t>Gamma</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>2</v></c>
      <c r="B1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(99),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
			sharedStringsDirty: false,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sharedStrings = new TextDecoder().decode(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sharedStrings.match(/<si>/g)).toHaveLength(4)
		expect(sheetXml).toContain('<c r="A2" t="s"><v>3</v></c>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Beta',
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(1, 0)?.value).toEqual({
			kind: 'string',
			value: 'Gamma',
		})
	})

	it('preserves package docProps parts and custom property relationships', () => {
		const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Original Author</dc:creator>
  <cp:lastModifiedBy>Reviewer</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>
</cp:coreProperties>`
		const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Excel</Application>
  <Company>Acme Analytics</Company>
</Properties>`
		const customXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Desk">
    <vt:lpwstr>Research</vt:lpwstr>
  </property>
</Properties>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`,
			'docProps/core.xml': coreXml,
			'docProps/app.xml': appXml,
			'docProps/custom.xml': customXml,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)
		expect(source.value.capsules.map((capsule) => capsule.partPath).sort()).toContain(
			'docProps/custom.xml',
		)

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const decode = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(decode('docProps/core.xml')).toBe(coreXml)
		expect(decode('docProps/app.xml')).toBe(appXml)
		expect(decode('docProps/custom.xml')).toBe(customXml)
		expect(decode('_rels/.rels')).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"',
		)
		expect(decode('_rels/.rels')).toContain('Target="docProps/custom.xml"')
		expect(decode('xl/_rels/workbook.xml.rels')).not.toContain('docProps/custom.xml')
	})

	it('preserves nonstandard root core-properties topology', () => {
		const corePath = 'package/services/metadata/core-properties/source.psmdcp'
		const coreXml = `<?xml version="1.0" encoding="utf-8"?>
<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
  <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">ClosedXML</dc:creator>
</coreProperties>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Default Extension="psmdcp" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="/docProps/app.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="/package/services/metadata/core-properties/source.psmdcp"/>
</Relationships>`,
			'docProps/app.xml': '<Properties/>',
			[corePath]: coreXml,
			'xl/workbook.xml': `<?xml version="1.0" encoding="utf-8"?>
<x:workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets>
</x:workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="utf-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		sheet?.cells.set(0, 0, { value: numberValue(424242), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const decode = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(zip['docProps/core.xml']).toBeUndefined()
		expect(decode(corePath)).toBe(coreXml)
		expect(decode('[Content_Types].xml')).toContain(
			'<Default Extension="psmdcp" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
		)
		expect(decode('[Content_Types].xml')).not.toContain('PartName="/xl/workbook.xml"')
		expect(decode('[Content_Types].xml')).not.toContain('PartName="/docProps/core.xml"')
		expect(decode('_rels/.rels')).toContain(
			'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="/package/services/metadata/core-properties/source.psmdcp"',
		)
		expect(decode('xl/_rels/workbook.xml.rels')).not.toContain('core-properties')
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
		expect(workbookPart?.origin).toBe('generated')
		expect(stylesPart?.origin).toBe('generated')
		expect(themePart?.origin).toBe('preserved-source')
		expect(sheetPart?.origin).toBe('generated')
	})

	it('dirty-part patching: only regenerates modified sheet when dirtySheetNames provided', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const s3 = wb.addSheet('Sheet3')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })
		s3.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: S0 })

		const initialBytes = writeXlsx(wb)
		expectOk(initialBytes)
		const read = readXlsx(initialBytes.value)
		expectOk(read)

		read.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(99),
			formula: null,
			styleId: S0,
		})

		const dirtyWrite = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(dirtyWrite)

		const reopened = readXlsx(dirtyWrite.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets).toHaveLength(3)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 1,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
		expect(reopened.value.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 2,
		})
		expect(reopened.value.workbook.sheets[2]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 3,
		})

		const fullWrite = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Sheet1', 'Sheet2', 'Sheet3'],
		})
		expectOk(fullWrite)
		const fullReopened = readXlsx(fullWrite.value)
		expectOk(fullReopened)
		expect(fullReopened.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
	})

	it('dirty-part patching preserves shared string indexes for untouched sheets', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.cells.set(0, 0, { value: stringValue('alpha'), formula: null, styleId: S0 })
		s2.cells.set(0, 0, { value: stringValue('beta'), formula: null, styleId: S0 })
		const initial = writeXlsx(wb)
		expectOk(initial)
		const reopened = readXlsx(initial.value)
		expectOk(reopened)
		reopened.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: stringValue('gamma'),
			formula: null,
			styleId: S0,
		})
		const patched = writeXlsx(reopened.value.workbook, reopened.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(patched)
		const roundTripped = readXlsx(patched.value)
		expectOk(roundTripped)
		expect(roundTripped.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual(
			stringValue('alpha'),
		)
		expect(roundTripped.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual(
			stringValue('gamma'),
		)
		expect(roundTripped.value.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual(
			stringValue('beta'),
		)
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
		expect(stylesXml).toContain('formatCode="0.0%"')
		expect(stylesXml).toContain('applyNumberFormat="1"')
		expect(stylesXml).toContain('<cellStyleXfs')
		expect(stylesXml).toContain('<cellStyles')
		expect(stylesXml).toContain('defaultTableStyle="TableStyleMedium2"')
		expect(stylesXml).toContain('defaultPivotStyle="PivotStyleLight16"')
		expect(stylesXml).toContain('<tableStyle name="TableStyleMedium2"/>')

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
		sheet.rowBreaks = [{ id: 5, min: 0, max: 16383, man: true }]
		sheet.colBreaks = [{ id: 2, min: 0, max: 1048575, man: true }]

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
		expect(s?.rowBreaks).toEqual([{ id: 5, min: 0, max: 16383, man: true }])
		expect(s?.colBreaks).toEqual([{ id: 2, min: 0, max: 1048575, man: true }])
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

	it('preserves default-covered capsule content types and workbook-relative targets', () => {
		const wb = new Workbook()
		wb.preservedXml = {
			contentTypeDefaults: [{ extension: 'png', contentType: 'image/png' }],
		}
		wb.addSheet('Sheet1')
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'customXml/item1.xml',
				contentType: 'application/xml',
				contentTypeSource: 'default',
				relationships: [],
				content: new TextEncoder().encode('<root/>'),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
			},
			{
				partPath: 'xl/media/image1.png',
				contentType: 'image/png',
				contentTypeSource: 'default',
				relationships: [],
				content: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
				anchor: { kind: 'workbook' },
			},
		]

		const written = writeXlsx(wb, capsules)
		if (!written.ok) throw new Error(`write failed: ${written.error.message}`)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)

		expect(contentTypes).toContain('<Default Extension="png" ContentType="image/png"/>')
		expect(contentTypes).not.toContain('PartName="/customXml/item1.xml"')
		expect(contentTypes).not.toContain('PartName="/xl/media/image1.png"')
		expect(workbookRels).toContain('Target="../customXml/item1.xml"')
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
				{
					colId: 1,
					kind: 'dynamicFilter',
					dynamicFilterType: 'thisMonth',
					dynamicFilterValIso: '2026-03-01T00:00:00',
					dynamicFilterMaxValIso: '2026-04-01T00:00:00',
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
				{
					colId: 1,
					kind: 'dynamicFilter',
					dynamicFilterType: 'thisMonth',
					dynamicFilterValIso: '2026-03-01T00:00:00',
					dynamicFilterMaxValIso: '2026-04-01T00:00:00',
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
			filterColumn: 2,
			filters: 1,
			filter: 2,
			dynamicFilter: 1,
			sortState: 1,
			sortCondition: 1,
		})
	})

	it('preserves top-level worksheet sort state on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Sorted')
		sheet.cells.set(0, 3, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.sortState = {
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1' }],
		}

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.sortState).toEqual({
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1' }],
		})
		const sheetXml = new TextDecoder().decode(unzipSync(bytes)['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<sortState ref="D1:I2707"><sortCondition ref="D1"/></sortState>')
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
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
			{
				partPath: 'xl/drawings/vmlDrawing1.vml',
				contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
				relationships: [],
				content: new TextEncoder().encode('<xml xmlns:v="urn:schemas-microsoft-com:vml"/>'),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
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

	it('generates drawing XML and media parts for programmatic image refs', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Images')
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
			anchor: {
				kind: 'oneCell',
				from: { row: 1, col: 1 },
				cx: 320000,
				cy: 240000,
			},
			name: 'Logo',
			description: 'Brand logo',
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		expect(entries['xl/drawings/drawing1.xml']).toBeDefined()
		expect(entries['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined()
		expect(entries['xl/media/image1.png']).toBeDefined()

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.drawingRefs.hasDrawing).toBe(true)
		expect(reopened.value.workbook.sheets[0]?.imageRefs).toHaveLength(1)
		expect(reopened.value.workbook.sheets[0]?.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			targetPath: 'xl/media/image1.png',
			name: 'Logo',
			description: 'Brand logo',
			anchor: {
				kind: 'oneCell',
			},
		})
	})

	it('overrides replaced image media while preserving existing drawing capsules', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Images')
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
			anchor: {
				kind: 'oneCell',
				from: { row: 1, col: 1 },
				cx: 320000,
				cy: 240000,
			},
			name: 'Logo',
		})
		const source = writeXlsx(wb)
		expectOk(source)
		const read = readXlsx(source.value)
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		const image = readSheet?.imageRefs[0]
		expect(image).toBeDefined()
		if (!readSheet || !image) return
		readSheet.imageRefs[0] = {
			...image,
			contentType: 'image/png',
			content: new Uint8Array([4, 5, 6]),
		}

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Images'],
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		expect(Array.from(entries['xl/media/image1.png'] ?? [])).toEqual([4, 5, 6])

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			targetPath: 'xl/media/image1.png',
			name: 'Logo',
			anchor: { kind: 'oneCell' },
		})
	})

	it('keeps sheet-owned drawing capsules attached after renaming a sheet', () => {
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
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
			{
				partPath: 'xl/drawings/vmlDrawing1.vml',
				contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
				relationships: [],
				content: new TextEncoder().encode('<xml xmlns:v="urn:schemas-microsoft-com:vml"/>'),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
			},
		]

		sheet.name = 'Dashboard'
		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.sheets[0]?.name).toBe('Dashboard')
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

	it('writes inline strings with useInlineStrings and round-trips correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		for (let r = 0; r < 50; r++) {
			for (let c = 0; c < 10; c++) {
				sheet.cells.set(r, c, {
					value: stringValue(`unique-${r}-${c}`),
					formula: null,
					styleId: S0,
				})
			}
		}

		const written = writeXlsx(wb, undefined, { useInlineStrings: true })
		expectOk(written)

		const zip = unzipSync(written.value)
		expect(zip['xl/sharedStrings.xml']).toBeUndefined()

		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('t="inlineStr"')
		expect(sheetXml).toContain('<is><t>unique-0-0</t></is>')
		expect(sheetXml).toContain('<is><t>unique-25-5</t></is>')

		const read = readXlsx(written.value)
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'unique-0-0',
		})
		expect(readSheet?.cells.get(25, 5)?.value).toEqual({
			kind: 'string',
			value: 'unique-25-5',
		})
	})

	it('writes plain string cells with usePlainStrings and round-trips correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, {
			value: stringValue('plain & fast'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(wb, undefined, { usePlainStrings: true })
		expectOk(written)

		const zip = unzipSync(written.value)
		expect(zip['xl/sharedStrings.xml']).toBeUndefined()

		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('t="str"')
		expect(sheetXml).toContain('<v>plain &amp; fast</v>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'plain & fast',
		})
	})

	it('omits dense cell refs only when positions can be inferred', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('b'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: stringValue('sparse'), formula: null, styleId: S0 })

		const written = writeXlsx(wb, undefined, {
			useSharedStrings: false,
			usePlainStrings: true,
			omitDenseCellRefs: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row r="1"><c t="str"><v>a</v></c><c t="str"><v>b</v></c></row>')
		expect(sheetXml).toContain('<row r="2"><c r="B2" t="str"><v>sparse</v></c></row>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'a' })
		expect(readSheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'b' })
		expect(readSheet?.cells.get(1, 1)?.value).toEqual({ kind: 'string', value: 'sparse' })
	})

	it('writes dense rows without materializing a workbook', () => {
		const written = writeDenseRowsXlsx({
			rows: 2,
			cols: 3,
			omitCellRefs: true,
			allCellsPresent: true,
			valueAt: (row, col) => (row === 1 && col === 1 ? null : `r${row}c${col}`),
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row r="1"><c t="str"><v>r0c0</v></c>')
		expect(sheetXml).toContain(
			'<row r="2"><c r="A2" t="str"><v>r1c0</v></c><c r="C2" t="str"><v>r1c2</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'r0c0' })
		expect(readSheet?.cells.get(1, 1)).toBeUndefined()
		expect(readSheet?.cells.get(1, 2)?.value).toEqual({ kind: 'string', value: 'r1c2' })
	})

	it('reuses constant dense row bodies when explicitly requested', () => {
		let calls = 0
		const written = writeDenseRowsXlsx({
			rows: 4,
			cols: 3,
			omitCellRefs: true,
			constantRows: true,
			valueAt: (_row, col) => {
				calls++
				return col < 2 ? 'repeat' : col
			},
		})
		expectOk(written)
		expect(calls).toBe(3)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'repeat' })
		expect(readSheet?.cells.get(3, 0)?.value).toEqual({ kind: 'string', value: 'repeat' })
		expect(readSheet?.cells.get(3, 2)?.value).toEqual({ kind: 'number', value: 2 })
	})

	it('can skip string escaping for generated XML-safe dense values', () => {
		const written = writeDenseRowsXlsx({
			rows: 1,
			cols: 2,
			omitCellRefs: true,
			stringsAreXmlSafe: true,
			valueAt: (_row, col) => `safe-${col}`,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<c t="str"><v>safe-0</v></c>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'string',
			value: 'safe-1',
		})
	})

	it('uses dense value type hints while preserving fallback semantics', () => {
		const written = writeDenseRowsXlsx({
			rows: 1,
			cols: 3,
			omitCellRefs: true,
			valueType: 'number',
			valueAt: (_row, col) => (col === 2 ? 'fallback' : col + 1),
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain(
			'<row r="1"><c><v>1</v></c><c><v>2</v></c><c t="str"><v>fallback</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(0, 2)?.value).toEqual({ kind: 'string', value: 'fallback' })
	})

	it('uses dense column value type hints while preserving fallback semantics', () => {
		const written = writeDenseRowsXlsx({
			rows: 2,
			cols: 4,
			omitCellRefs: true,
			allCellsPresent: true,
			stringsAreXmlSafe: true,
			valueTypes: ['string', 'number', 'string', 'number'],
			valueAt: (row, col) => {
				if (row === 1 && col === 3) return 'fallback'
				return col % 2 === 0 ? `text-${row}-${col}` : row * 10 + col
			},
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain(
			'<row r="1"><c t="str"><v>text-0-0</v></c><c><v>1</v></c><c t="str"><v>text-0-2</v></c><c><v>3</v></c></row>',
		)
		expect(sheetXml).toContain(
			'<row r="2"><c t="str"><v>text-1-0</v></c><c><v>11</v></c><c t="str"><v>text-1-2</v></c><c t="str"><v>fallback</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'text-0-0' })
		expect(readSheet?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(1, 3)?.value).toEqual({ kind: 'string', value: 'fallback' })
	})

	it('streams dense rows without materializing a workbook', async () => {
		const written = await writeDenseRowsXlsxStreaming({
			rows: 2,
			cols: 2,
			omitCellRefs: true,
			valueAt: (row, col) => row * 2 + col,
		})
		expectOk(written)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 0 })
		expect(readSheet?.cells.get(1, 1)?.value).toEqual({ kind: 'number', value: 3 })
	})

	it('supports compact dense writer compression without changing values', async () => {
		const written = await writeDenseRowsXlsxStreaming({
			rows: 200,
			cols: 20,
			omitCellRefs: true,
			allCellsPresent: true,
			stringsAreXmlSafe: true,
			valueType: 'string',
			compressionProfile: 'compact',
			valueAt: (row, col) => `text-${row}-${col}`,
		})
		expectOk(written)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'text-0-0' })
		expect(readSheet?.cells.get(199, 19)?.value).toEqual({
			kind: 'string',
			value: 'text-199-19',
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

	it('preserves hidden defined names on round-trip', () => {
		const wb = new Workbook()
		const data = wb.addSheet('Data')
		data.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.definedNames.set(
			'_xlnm._FilterDatabase',
			'Data!$A$1:$B$10',
			{ kind: 'sheet', sheetId: data.id },
			{ hidden: true },
		)

		const { bytes } = roundTrip(wb)
		const workbookXml = new TextDecoder().decode(unzipSync(bytes)['xl/workbook.xml'])
		expect(workbookXml).toContain(
			'<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$10</definedName>',
		)

		const result = readXlsx(bytes)
		expectOk(result)
		const resolved = result.value.workbook.definedNames.resolve(
			'_xlnm._FilterDatabase',
			result.value.workbook.getSheet('Data')?.id,
		)
		expect(resolved?.hidden).toBe(true)
		expect(resolved?.formula).toBe('Data!$A$1:$B$10')
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

	it('rewrites external link relationship targets while preserving link parts', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			linkRelId: 'rIdExt',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
				relationships: [
					{
						id: 'rIdExt',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
						target: '../sources/source.xlsx',
						targetMode: 'External',
					},
				],
				content: new TextEncoder().encode(
					'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
				),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			},
		]

		const applied = applyOperations(wb, [
			{
				op: 'rewriteExternalLink',
				partPath: 'xl/externalLinks/externalLink1.xml',
				linkRelId: 'rIdExt',
				newTarget: '../sources/reforecast & final.xlsx',
			},
		])
		expectOk(applied)
		const written = writeXlsx(wb, capsules, { workbookMetaDirty: true })
		expectOk(written)

		const parts = unzipSync(written.value)
		const rels = new TextDecoder().decode(
			parts['xl/externalLinks/_rels/externalLink1.xml.rels'] ?? new Uint8Array(),
		)
		expect(rels).toContain('Target="../sources/reforecast &amp; final.xlsx"')
		expect(rels).toContain('TargetMode="External"')

		const read = readXlsx(written.value)
		expectOk(read)
		expect(read.value.workbook.externalReferenceDetails[0]).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			linkRelId: 'rIdExt',
			target: '../sources/reforecast & final.xlsx',
			targetMode: 'External',
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

	it('generates workbook theme part from theme metadata when no preserved theme exists', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('ThemeMeta')
		sheet.cells.set(0, 0, { value: stringValue('Brand'), formula: null, styleId: S0 })
		wb.themeMetadata = {
			name: 'Generated Theme',
			colorSchemeName: 'Generated Colors',
			colorCount: 12,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		}

		const { result, bytes } = roundTrip(wb)
		const zip = unzipSync(bytes)
		const themeXml = new TextDecoder().decode(zip['xl/theme/theme1.xml'] ?? new Uint8Array())
		expect(themeXml).toContain('name="Generated Theme"')
		expect(themeXml).toContain('name="Generated Colors"')
		expect(themeXml).toContain('typeface="Aptos Display"')
		expect(themeXml).toContain('typeface="Aptos"')
		expect(result.workbook.preservedTheme).toEqual({
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
		})
		expect(result.workbook.themeMetadata).toEqual({
			name: 'Generated Theme',
			colorSchemeName: 'Generated Colors',
			colorCount: 12,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		})
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
			columns: [
				{ id: 1, name: 'Name' },
				{ id: 2, name: 'Value' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		const capsuleContent = new TextEncoder().encode(`<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="BalanceTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
</table>`)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/tables/table1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
				relationships: [],
				content: capsuleContent,
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
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
		expect(unzipSync(bytes)['xl/tables/table1.xml']).toEqual(capsuleContent)
	})

	it('regenerates table parts when capsule-backed table metadata changes', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Balance')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(2, 0, { value: stringValue('Debt'), formula: null, styleId: S0 })
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Name' },
				{ id: 2, name: 'Value' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		const staleCapsuleContent = new TextEncoder().encode(`<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="BalanceTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
</table>`)
		const { bytes } = roundTrip(wb, [
			{
				partPath: 'xl/tables/table1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
				relationships: [],
				content: staleCapsuleContent,
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			},
		])

		const tableEntry = unzipSync(bytes)['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		expect(tableEntry).not.toEqual(staleCapsuleContent)
		expect(new TextDecoder().decode(tableEntry)).toContain('ref="A1:B3"')
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

	it('updates chart series refs inside preserved chart capsules', () => {
		const wb = new Workbook()
		wb.addSheet('Sheet1')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			chartType: 'barChart',
			series: [
				{
					nameRef: 'Data!$C$1',
					categoryRef: 'Data!$A$2:$A$10',
					valueRef: 'Data!$C$2:$C$10',
				},
			],
		})
		const chartXml = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart><c:ser>
    <c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>
    <c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f></c:strRef></c:cat>
    <c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:val>
  </c:ser></c:barChart></c:plotArea></c:chart>
</c:chartSpace>`
		const written = writeXlsx(wb, [
			{
				partPath: 'xl/charts/chart1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
				relationships: [],
				content: new TextEncoder().encode(chartXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
			},
		])
		expectOk(written)

		const entries = unzipSync(written.value)
		const updated = entries['xl/charts/chart1.xml']
		expect(updated).toBeDefined()
		if (!updated) return
		const xml = new TextDecoder().decode(updated)
		expect(xml).toContain('<c:f>Data!$C$1</c:f>')
		expect(xml).toContain('<c:f>Data!$A$2:$A$10</c:f>')
		expect(xml).toContain('<c:f>Data!$C$2:$C$10</c:f>')
		expect(xml).not.toContain('<c:f>Data!$B$2:$B$4</c:f>')
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
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
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

	it('streaming write produces valid XLSX that round-trips (50K rows)', async () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		const rows = 50_000
		const cols = 5
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				sheet.cells.set(r, c, {
					value: numberValue(r * cols + c + 1),
					formula: null,
					styleId: S0,
				})
			}
		}

		const written = await writeXlsxStreaming(wb, undefined, { streaming: true })
		expectOk(written)

		const entries = unzipSync(written.value)
		expect(entries['xl/worksheets/sheet1.xml']).toBeDefined()
		const sheetXml = new TextDecoder().decode(entries['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<sheetData>')
		expect(sheetXml).toContain(`<row r="${rows}"`)

		const read = readXlsx(written.value)
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(rows - 1, cols - 1)?.value).toEqual({
			kind: 'number',
			value: rows * cols,
		})
	}, 90_000)

	it('streaming write plan emits worksheet XML through streamingBuild', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		for (let r = 0; r < 10_000; r++) {
			sheet.cells.set(r, 0, {
				value: numberValue(r + 1),
				formula: null,
				styleId: S0,
			})
		}

		const plan = planWriteXlsx(wb, undefined, { streaming: true })
		expectOk(plan)
		const sheetPart = plan.value.descriptors.find(
			(entry) => entry.path === 'xl/worksheets/sheet1.xml',
		)
		expect(sheetPart?.streamingBuild).toBeDefined()
		expect(plan.value.parts.has('xl/worksheets/sheet1.xml')).toBe(false)
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

	it('preserves calcChain for value edits that do not change formula topology', () => {
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
			calcChainDirty: false,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		expect(zip['xl/calcChain.xml']).toBeDefined()
		expect(workbookRels).toContain('relationships/calcChain')
		expect(contentTypes).toContain('calcChain+xml')
		expect(workbookXml).toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="0"')
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

	it('preserves sheetView attributes (zoomScale, zoomScaleNormal, zoomScaleSheetLayoutView, showGridLines, showFormulas, rightToLeft, tabSelected, view, topLeftCell) on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('View')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.sheetView = {
			zoomScale: 125,
			zoomScaleNormal: 100,
			zoomScaleSheetLayoutView: 214,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageLayout',
			topLeftCell: 'E1',
		}

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.sheetView).toEqual({
			zoomScale: 125,
			zoomScaleNormal: 100,
			zoomScaleSheetLayoutView: 214,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageLayout',
			topLeftCell: 'E1',
		})
	})

	it('preserves plain sheetView presence on regenerated sheets', () => {
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
  <sheets><sheet name="View" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['View'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<sheetView workbookViewId="0">')
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

	it('preserves worksheet-level extLst instead of nested rule extensions', () => {
		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData>
  <x:conditionalFormatting sqref="A1">
    <x:cfRule type="dataBar" priority="1"><x:extLst><x:ext uri="{nested}"/></x:extLst></x:cfRule>
  </x:conditionalFormatting>
  <x:extLst><x:ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{worksheet}"><x14:conditionalFormattings/></x:ext></x:extLst>
</x:worksheet>`
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
		expectOk(read)
		const preserved = read.value.workbook.sheets[0]?.preservedExtLst
		expect(preserved).toContain('uri="{worksheet}"')
		expect(preserved).toContain('x14:conditionalFormattings')
		expect(preserved).not.toContain('uri="{nested}"')

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const writtenSheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(writtenSheetXml).toContain('<extLst>')
		expect(writtenSheetXml).toContain('x14:conditionalFormattings')
		expect(writtenSheetXml).not.toContain('uri="{nested}"')
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
