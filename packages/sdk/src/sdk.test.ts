import { describe, expect, test } from 'bun:test'
import { Buffer } from 'node:buffer'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { extractZip } from '../../io-xlsx/src/reader/zip.ts'
import { createZip, encode } from '../../io-xlsx/src/writer/zip.ts'
import {
	Ascend,
	AscendWorkbook,
	listCapabilities,
	ops,
	parseOperations,
	WorkbookDocument,
	WorkbookSession,
} from './index.ts'

describe('AscendWorkbook', () => {
	test('Ascend entry point exposes create, open, fromCsv', async () => {
		const wb = Ascend.create()
		expect(wb.sheets).toEqual(['Sheet1'])
		expect(wb.sheet('Sheet1')?.cell('A1')).toBeUndefined()

		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'via Ascend' }] }])
		const bytes = wb.toBytes()
		const reopened = await Ascend.open(bytes)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'via Ascend',
		})

		const csv = Ascend.fromCsv('a,b\n1,2')
		expect(csv.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'a' })
		expect(csv.sheet('Sheet1')?.cell('B2')?.value).toEqual({ kind: 'number', value: 2 })
	})

	test('opens encrypted XLSX fixtures with a password through the SDK', async () => {
		const encrypted = readFileSync('fixtures/xlsx/calamine/pass_protected.xlsx')
		const wb = await Ascend.open(new Uint8Array(encrypted), { password: '123' })
		expect(wb.sheets).toContain('Sheet1')
		expect(() => extractZip(wb.toBytes())).not.toThrow()
		await expect(Ascend.open(new Uint8Array(encrypted))).rejects.toThrow('requires a password')
	})

	test('create returns an empty workbook with one sheet', () => {
		const wb = AscendWorkbook.create()
		expect(wb.sheets).toEqual(['Sheet1'])
		const info = wb.inspect()
		expect(info.sheetCount).toBe(1)
		expect(info.loadedSheetCount).toBe(1)
		expect(info.cellCount).toBe(0)
		expect(info.sourceFormat).toBe('ascend')
		expect(info.load.isPartial).toBe(false)
		expect(info.capabilityWarnings).toEqual([])
	})

	test('inspect exposes package document properties in metadata-only mode', async () => {
		const bytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`,
			'docProps/core.xml': `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Forecast Pack</dc:title>
  <dc:creator>Finance Ops</dc:creator>
  <dcterms:created>2025-03-04T05:06:07Z</dcterms:created>
</cp:coreProperties>`,
			'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Excel</Application>
  <Company>Ascend Fixtures</Company>
</Properties>`,
			'docProps/custom.xml': `<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Reviewed">
    <vt:bool>true</vt:bool>
  </property>
</Properties>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
		})

		const wb = await AscendWorkbook.open(bytes, { mode: 'metadata-only' })

		expect(wb.inspect().documentProperties).toEqual({
			core: {
				title: 'Forecast Pack',
				creator: 'Finance Ops',
				created: '2025-03-04T05:06:07Z',
			},
			app: {
				Application: 'Microsoft Excel',
				Company: 'Ascend Fixtures',
			},
			custom: [
				{
					name: 'Reviewed',
					value: true,
					type: 'bool',
					pid: 2,
					fmtid: '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
				},
			],
		})
	})

	test('setDocumentProperties saves edited core app and custom docProps', async () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{
				op: 'setDocumentProperties',
				properties: {
					core: { title: 'Forecast Pack', creator: 'Finance Ops' },
					app: {
						Application: 'Ascend',
						Company: 'Ascend Fixtures',
						HeadingPairs: ['Worksheets', 1],
						TitlesOfParts: ['Sheet1'],
					},
					custom: [{ name: 'Reviewed', value: true, type: 'bool' }],
				},
			},
		])
		expect(result.errors).toEqual([])

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.inspect().documentProperties).toMatchObject({
			core: { title: 'Forecast Pack', creator: 'Finance Ops' },
			app: {
				Application: 'Ascend',
				Company: 'Ascend Fixtures',
				HeadingPairs: ['Worksheets', 1],
				TitlesOfParts: ['Sheet1'],
			},
			custom: [{ name: 'Reviewed', value: true, type: 'bool' }],
		})
	})

	test('cellStyle exposes a cloned exact cell style', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
			{
				op: 'setStyle',
				sheet: 'Sheet1',
				range: 'A1',
				style: { numberFormat: '0.00', font: { bold: true } },
			},
		])

		const style = wb.cellStyle('Sheet1!A1')
		expect(style?.numberFormat).toBe('0.00')
		expect(style?.font?.bold).toBe(true)
		expect(wb.cellStyle('Sheet1!B1')).toBeUndefined()

		if (style?.font) {
			;(style.font as { bold?: boolean }).bold = false
		}
		expect(wb.cellStyle('Sheet1!A1')?.font?.bold).toBe(true)
	})

	test('inspect returns correct sheet info', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'hello' }] }])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<Record<string, unknown>>
				workbookProtection: unknown
				pivotTables: Array<Record<string, unknown>>
				pivotCaches: Array<Record<string, unknown>>
				slicerCaches: Array<Record<string, unknown>>
				slicers: Array<Record<string, unknown>>
			}
		}
		const backingSheet = internal.wb.sheets[0] as
			| {
					comments: Map<string, { text: string }>
					dataValidations: Array<Record<string, unknown>>
					conditionalFormats: Array<Record<string, unknown>>
					x14ConditionalFormats: Array<Record<string, unknown>>
					x14DataValidations: Array<Record<string, unknown>>
					imageRefs: Array<Record<string, unknown>>
			  }
			| undefined
		backingSheet?.comments.set('A1', { text: 'note' })
		backingSheet?.dataValidations.push({ sqref: 'B1', type: 'list', formula1: '"A,B"' })
		backingSheet?.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'cellIs', formulas: ['1'] }],
		})
		backingSheet?.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C1:C2',
			formulas: ['C1>0'],
			type: 'dataBar',
		})
		backingSheet?.x14DataValidations.push({
			index: 0,
			sqref: 'D1:D2',
			type: 'list',
			formula1: 'Lookup!$A$1:$A$2',
		})
		backingSheet?.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: 'xl/media/image1.png',
		})
		internal.wb.workbookProtection = { lockStructure: true }
		internal.wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			locationRef: 'A1',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			fields: [],
		})
		internal.wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Product',
			pivotTableNames: ['PivotTable1'],
			items: [
				{ index: 0, selected: true },
				{ index: 1, noData: true },
			],
		})
		internal.wb.slicers.push({
			partPath: 'xl/slicers/slicer1.xml',
			name: 'Product',
			cacheName: 'Slicer_Product',
		})
		const info = wb.inspect()
		expect(info.cellCount).toBe(1)
		expect(info.commentCount).toBe(1)
		expect(info.conditionalFormatCount).toBe(1)
		expect(info.dataValidationCount).toBe(1)
		expect(info.x14ConditionalFormatCount).toBe(1)
		expect(info.x14DataValidationCount).toBe(1)
		expect(info.imageCount).toBe(1)
		expect(info.pivotTableCount).toBe(1)
		expect(info.pivotCacheCount).toBe(1)
		expect(info.capabilityWarnings).toContainEqual(
			expect.objectContaining({
				capabilityId: 'analytics.pivots',
				status: 'inspectable',
				priority: 'P0',
				evidence: ['pivotTables=1'],
			}),
		)
		expect(info.slicerCount).toBe(1)
		expect(info.slicerCacheCount).toBe(1)
		expect(info.hasWorkbookProtection).toBe(true)
		expect(info.pivotTables[0]?.name).toBe('PivotTable1')
		expect(info.pivotCaches[0]?.sourceSheet).toBe('Raw')
		expect(info.slicerCaches[0]?.name).toBe('Slicer_Product')
		expect(info.slicerCaches[0]?.items).toEqual([
			{ index: 0, selected: true },
			{ index: 1, noData: true },
		])
		expect(info.slicers[0]?.name).toBe('Product')
		expect(info.sheets[0]?.cellCount).toBe(1)
		expect(info.sheets[0]?.commentCount).toBe(1)
		expect(info.sheets[0]?.conditionalFormatCount).toBe(1)
		expect(info.sheets[0]?.dataValidationCount).toBe(1)
		expect(info.sheets[0]?.x14ConditionalFormatCount).toBe(1)
		expect(info.sheets[0]?.x14DataValidationCount).toBe(1)
		expect(info.sheets[0]?.imageCount).toBe(1)
		expect(info.sheets[0]?.hasProtection).toBe(false)
		expect(info.sheets[0]?.name).toBe('Sheet1')
		expect(info.sheets[0]?.cellDataLoaded).toBe(true)
	})

	test('WorkbookReadView and WorkbookDocument expose pivot and slicer query surfaces', async () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				pivotTables: Array<Record<string, unknown>>
				pivotCaches: Array<Record<string, unknown>>
				slicerCaches: Array<Record<string, unknown>>
				slicers: Array<Record<string, unknown>>
			}
		}
		internal.wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 7,
			locationRef: 'A1',
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 7,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			fields: [],
		})
		internal.wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Product',
			pivotTableNames: ['PivotTable1'],
			items: [
				{ index: 0, selected: true },
				{ index: 1, noData: true },
			],
		})
		internal.wb.slicers.push({
			partPath: 'xl/slicers/slicer1.xml',
			name: 'Product',
			cacheName: 'Slicer_Product',
		})

		expect(wb.pivotTables()).toHaveLength(1)
		expect(wb.pivotTables('Sheet1')[0]?.name).toBe('PivotTable1')
		expect(wb.pivotCaches()[0]?.sourceSheet).toBe('Raw')
		expect(wb.slicerCaches()[0]?.pivotTableNames).toEqual(['PivotTable1'])
		expect(wb.slicerCaches()[0]?.items).toEqual([
			{ index: 0, selected: true },
			{ index: 1, noData: true },
		])
		expect(wb.inspect().slicerCaches[0]?.items).toEqual([
			{ index: 0, selected: true },
			{ index: 1, noData: true },
		])
		expect(wb.slicers()[0]?.cacheName).toBe('Slicer_Product')

		const corpusPath = join(
			import.meta.dir,
			'../../../research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx',
		)
		if (existsSync(corpusPath)) {
			const session = await WorkbookDocument.open(corpusPath)
			expect(session.pivotTables().length).toBeGreaterThan(0)
			expect(session.pivotCaches().length).toBeGreaterThan(0)
			WorkbookDocument.clearCache()
		}
	})

	test('inspectSheet returns parsed worksheet structures', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'hello' }] },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://example.com' },
		])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<Record<string, unknown>>
			}
		}
		const backingSheet = internal.wb.sheets[0] as
			| {
					comments: Map<string, { text: string; author?: string }>
					ignoredErrors: Array<Record<string, unknown>>
					conditionalFormats: Array<Record<string, unknown>>
					dataValidations: Array<Record<string, unknown>>
					x14ConditionalFormats: Array<Record<string, unknown>>
					x14DataValidations: Array<Record<string, unknown>>
					drawingRefs: { hasDrawing: boolean; hasLegacyDrawing: boolean }
					pageMargins: Record<string, unknown> | null
			  }
			| undefined
		backingSheet?.comments.set('A1', { text: 'note', author: 'me' })
		backingSheet?.ignoredErrors.push({ sqref: 'A1', formula: true })
		backingSheet?.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'cellIs', formulas: ['1'] }],
		})
		backingSheet?.dataValidations.push({ sqref: 'A1', type: 'list', formula1: '"A,B"' })
		backingSheet?.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B1:B2',
			formulas: ['B1>0'],
			type: 'iconSet',
			iconSet: { cfvo: [{ type: 'num', value: '0' }] },
		})
		backingSheet?.x14DataValidations.push({
			index: 0,
			sqref: 'C1:C2',
			type: 'list',
			allowBlank: true,
			formula1: 'Lookup!$A$1:$A$2',
		})
		if (backingSheet) {
			backingSheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: true }
			backingSheet.pageMargins = { left: 0.5 }
		}

		const detail = wb.inspectSheet('Sheet1')
		expect(detail).toBeDefined()
		expect(detail?.usedRange).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 0, col: 0 },
		})
		expect(detail?.comments?.[0]?.ref).toBe('A1')
		expect(detail?.hyperlinks?.[0]?.target).toBe('https://example.com')
		expect(detail?.ignoredErrors).toHaveLength(1)
		expect(detail?.conditionalFormats).toHaveLength(1)
		expect(detail?.dataValidations).toHaveLength(1)
		expect(detail?.x14ConditionalFormats?.[0]).toMatchObject({
			sqref: 'B1:B2',
			formulas: ['B1>0'],
			iconSet: { cfvo: [{ type: 'num', value: '0' }] },
		})
		expect(detail?.x14DataValidations?.[0]).toMatchObject({
			sqref: 'C1:C2',
			type: 'list',
			allowBlank: true,
			formula1: 'Lookup!$A$1:$A$2',
		})
		expect(detail?.drawingRefs?.hasLegacyDrawing).toBe(true)
		expect(detail?.pageMargins?.left).toBe(0.5)
	})

	const x14ConditionalFixture = new URL(
		'../../../fixtures/xlsx/poi/NewStyleConditionalFormattings.xlsx',
		import.meta.url,
	)
	const x14ValidationFixture = new URL(
		'../../../fixtures/xlsx/closedxml/Misc_DataValidation.xlsx',
		import.meta.url,
	)

	test.skipIf(!existsSync(x14ConditionalFixture) || !existsSync(x14ValidationFixture))(
		'inspect exposes x14 extension inventories from real corpus workbooks',
		async () => {
			const conditionalWorkbook = await AscendWorkbook.open(readFileSync(x14ConditionalFixture))
			const conditionalInfo = conditionalWorkbook.inspect()
			const conditionalSheetName = conditionalInfo.sheets.find(
				(sheet) => sheet.x14ConditionalFormatCount === 3,
			)?.name
			expect(conditionalInfo.x14ConditionalFormatCount).toBe(3)
			expect(conditionalSheetName).toBeDefined()
			expect(
				conditionalWorkbook.inspectSheet(conditionalSheetName ?? '')?.x14ConditionalFormats?.[0],
			).toMatchObject({
				sqref: 'E2:E17',
				type: 'dataBar',
				dataBar: {
					cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
					borderColor: { rgb: 'FF63C384' },
				},
			})

			const validationWorkbook = await AscendWorkbook.open(readFileSync(x14ValidationFixture))
			expect(validationWorkbook.inspect().x14DataValidationCount).toBe(2)
			expect(
				validationWorkbook.inspectSheet('Data Validation - Copy')?.x14DataValidations?.[0],
			).toMatchObject({
				sqref: 'A5:A5',
				type: 'list',
				allowBlank: true,
				errorStyle: 'stop',
				formula1: "'Data Validation'!$C$1:$C$2",
			})
		},
	)

	test('find returns cells matching search criteria', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Total' },
					{ ref: 'A2', value: 'Total Revenue' },
					{ ref: 'A3', value: 'Grand Total' },
					{ ref: 'B1', value: 'other' },
					{ ref: 'C1', value: 42 },
					{ ref: 'C2', value: 42 },
					{ ref: 'C3', value: 100 },
				],
			},
		])

		const containsTotal = wb.find('Sheet1', { value: 'Total', match: 'contains' })
		expect(containsTotal.map((r) => r.ref)).toEqual(['A1', 'A2', 'A3'])
		expect(containsTotal[0]?.value).toEqual({ kind: 'string', value: 'Total' })
		expect(containsTotal[1]?.value).toEqual({ kind: 'string', value: 'Total Revenue' })
		expect(containsTotal[2]?.value).toEqual({ kind: 'string', value: 'Grand Total' })

		const exact42 = wb.find('Sheet1', { value: 42 })
		expect(exact42.map((r) => r.ref)).toEqual(['C1', 'C2'])
		expect(exact42[0]?.value).toEqual({ kind: 'number', value: 42 })
		expect(exact42[1]?.value).toEqual({ kind: 'number', value: 42 })
	})

	test('find uses case-insensitive string matching', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'HELLO' },
					{ ref: 'A2', value: 'hello' },
					{ ref: 'A3', value: 'HeLLo World' },
				],
			},
		])

		const exact = wb.find('Sheet1', { value: 'hello', match: 'exact' })
		expect(exact.map((r) => r.ref)).toEqual(['A1', 'A2'])

		const contains = wb.find('Sheet1', { value: 'hello', match: 'contains' })
		expect(contains.map((r) => r.ref)).toEqual(['A1', 'A2', 'A3'])
	})

	test('find returns empty array for nonexistent sheet', () => {
		const wb = AscendWorkbook.create()
		expect(wb.find('Missing', { value: 'x' })).toEqual([])
	})

	test('sheet handle reads cells', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 42 },
					{ ref: 'B2', value: 'text' },
				],
			},
		])
		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()

		const a1 = handle?.cell('A1')
		expect(a1).toBeDefined()
		expect(a1?.value).toEqual({ kind: 'number', value: 42 })

		const b2 = handle?.cell('B2')
		expect(b2).toBeDefined()
		expect(b2?.value).toEqual({ kind: 'string', value: 'text' })

		expect(handle?.cell('C3')).toBeUndefined()
	})

	test('builder fluent API batches set, formula, and commit', () => {
		const wb = AscendWorkbook.create()
		const result = wb.builder().set('Sheet1!A1', 42).formula('Sheet1!B1', '=A1*2').commit()
		expect(result.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 42 })
		expect(wb.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1*2')
		wb.recalc()
		expect(wb.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 84 })
	})

	test('typed cell selectors work for set/get and sheet cell access', () => {
		const wb = AscendWorkbook.create()
		expect(wb.set({ sheet: 'Sheet1', cell: { row: 0, col: 0 } }, 42).errors).toEqual([])
		expect(wb.get({ sheet: 'Sheet1', cell: { row: 0, col: 0 } })).toEqual({
			kind: 'number',
			value: 42,
		})
		expect(wb.sheet('Sheet1')?.cell({ cell: { row: 0, col: 0 } })?.value).toEqual({
			kind: 'number',
			value: 42,
		})
	})

	test('typed range selectors work for fillFormula', () => {
		const wb = AscendWorkbook.create()
		wb.set('Sheet1!A1', 2)
		wb.set('Sheet1!A2', 3)
		const result = wb.fillFormula(
			{
				sheet: 'Sheet1',
				range: {
					start: { row: 0, col: 1 },
					end: { row: 1, col: 1 },
				},
			},
			'=A1*10',
		)
		expect(result.errors).toEqual([])
		wb.recalc()
		expect(wb.get('Sheet1!B1')).toEqual({ kind: 'number', value: 20 })
		expect(wb.get('Sheet1!B2')).toEqual({ kind: 'number', value: 30 })
	})

	test('builder coalesces adjacent set() calls per sheet into a single setCells op', () => {
		const wb = AscendWorkbook.create()
		const builder = wb
			.builder()
			.set('Sheet1!A1', 1)
			.set('Sheet1!A2', 2)
			.set('Sheet2!A1', 3)
			.set('Sheet2!A2', 4) as unknown as {
			ops: Array<{ op: string; sheet?: string; updates?: unknown[] }>
		}

		expect(builder.ops).toHaveLength(2)
		expect(builder.ops[0]).toMatchObject({
			op: 'setCells',
			sheet: 'Sheet1',
		})
		expect(builder.ops[1]).toMatchObject({
			op: 'setCells',
			sheet: 'Sheet2',
		})
		expect(builder.ops[0]?.updates).toHaveLength(2)
		expect(builder.ops[1]?.updates).toHaveLength(2)
	})

	test('set auto-coerces primitive values through setCells', () => {
		const wb = AscendWorkbook.create()
		expect(wb.set('Sheet1!A1', 42).errors).toEqual([])
		expect(wb.set('Sheet1!A2', 'hello').errors).toEqual([])
		expect(wb.set('Sheet1!A3', true).errors).toEqual([])

		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 42 })
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'string', value: 'hello' })
		expect(wb.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'boolean', value: true })
	})

	test('eval evaluates formula against workbook state without writing to cell', () => {
		const wb = AscendWorkbook.create()
		wb.set('Sheet1!A1', 1)
		wb.set('Sheet1!A2', 2)
		wb.set('Sheet1!A3', 3)
		const result = wb.eval('=SUM(Sheet1!A1:A3)')
		expect(result).toEqual({ kind: 'number', value: 6 })
	})

	test('eval can resolve external workbook references with a caller hook', () => {
		const wb = AscendWorkbook.create()
		const result = wb.eval('=SUM([Budget.xlsx]Inputs!A1:A3)', {
			externalReferences: {
				resolveRange: ({ workbook, sheet, row, col, endRow, endCol }) => {
					expect({ workbook, sheet, row, col, endRow, endCol }).toEqual({
						workbook: 'Budget.xlsx',
						sheet: 'Inputs',
						row: 0,
						col: 0,
						endRow: 2,
						endCol: 0,
					})
					return [
						[{ kind: 'number', value: 1 }],
						[{ kind: 'number', value: 2 }],
						[{ kind: 'number', value: 3 }],
					]
				},
			},
		})
		expect(result).toEqual({ kind: 'number', value: 6 })
	})

	test('agentView summarizes a range for LLM-friendly reads', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
					{ ref: 'A3', value: 'Bob' },
				],
			},
		])
		wb.setFormula('Sheet1!B3', '=B2*2')
		wb.recalc()
		const view = wb.agentView('Sheet1', 'A1:B3')
		expect(view).toBeDefined()
		expect(view?.sheet).toBe('Sheet1')
		expect(view?.nonEmptyCount).toBe(6)
		expect(view?.formulaCount).toBe(1)
		expect(view?.distinctFunctions).toEqual([])
		expect(view?.columns[0]?.header).toBe('Name')
		expect(view?.columns[1]?.kind).toBe('mixed')
		expect(view?.samples[0]?.cells[0]?.value).toEqual({ kind: 'string', value: 'Name' })
	})

	test('sheet handle returns range info', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
					{ ref: 'B1', value: 3 },
				],
			},
		])
		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		const range = handle?.range('A1:B2')
		expect(range).toBeDefined()
		if (!range) return
		expect(range.rowCount).toBe(2)
		expect(range.colCount).toBe(2)
		expect(range.cells.length).toBe(3)
	})

	test('sheet handle streams range rows', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
					{ ref: 'A2', value: 3 },
				],
			},
		])
		const rows = [...(wb.sheet('Sheet1')?.streamRange('A1:B2') ?? [])]
		expect(rows).toHaveLength(2)
		expect(rows[0]?.map((cell) => cell.ref)).toEqual(['A1', 'B1'])
		expect(rows[1]?.map((cell) => cell.ref)).toEqual(['A2'])
	})

	test('sheet handle can read a windowed slice of a range', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
					{ ref: 'A3', value: 3 },
				],
			},
		])
		const window = wb.sheet('Sheet1')?.readWindow('A1:A3', { rowOffset: 1, rowLimit: 1 })
		expect(window?.cells.map((cell) => cell.ref)).toEqual(['A2'])
		expect(window?.hasMore).toBe(true)
		expect(window?.nextRowOffset).toBe(2)
	})

	test('sheet handle compact window omits refs when requested', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
		])
		const window = wb.sheet('Sheet1')?.readWindowCompact('A1:A2', {
			rowLimit: 2,
			includeRefs: false,
		})
		expect(window?.cells.map((cell) => cell.ref)).toEqual([undefined, undefined])
		expect(window?.cells.map((cell) => cell.formulaBinding)).toEqual([null, null])
	})

	test('sheet handle usedRange', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 1 },
					{ ref: 'D4', value: 2 },
				],
			},
		])
		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		const used = handle?.usedRange()
		expect(used).toBeDefined()
		expect(used?.start).toEqual({ row: 1, col: 1 })
		expect(used?.end).toEqual({ row: 3, col: 3 })
	})

	test('sheet returns undefined for nonexistent sheet', () => {
		const wb = AscendWorkbook.create()
		expect(wb.sheet('Nope')).toBeUndefined()
	})

	test('workbook readRange and streamRange helpers delegate to sheet handles', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'x' },
					{ ref: 'A2', value: 'y' },
				],
			},
		])

		const range = wb.readRange('Sheet1', 'A1:A2')
		expect(range?.cells.length).toBe(2)

		const rows = [...wb.streamRange('Sheet1', 'A1:A2')]
		expect(rows).toHaveLength(2)
		expect(rows[0]?.[0]?.ref).toBe('A1')
		expect(rows[1]?.[0]?.ref).toBe('A2')
	})

	test('workbook readWindow and streamWindows helpers paginate rows', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'x' },
					{ ref: 'A2', value: 'y' },
					{ ref: 'A3', value: 'z' },
				],
			},
		])

		const window = wb.readWindow('Sheet1', 'A1:A3', { rowLimit: 2 })
		expect(window?.cells.map((cell) => cell.ref)).toEqual(['A1', 'A2'])
		expect(window?.hasMore).toBe(true)

		const windows = [...wb.streamWindows('Sheet1', 'A1:A3', { rowLimit: 2 })]
		expect(windows).toHaveLength(2)
		expect(windows[0]?.cells.map((cell) => cell.ref)).toEqual(['A1', 'A2'])
		expect(windows[1]?.cells.map((cell) => cell.ref)).toEqual(['A3'])
	})

	test('workbook compact window helpers preserve row and column coordinates', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 'x' },
					{ ref: 'B3', value: 'y' },
				],
			},
		])
		const window = wb.readWindowCompact('Sheet1', 'B2:B3', { includeRefs: false })
		expect(window?.cells).toEqual([
			{
				row: 1,
				col: 1,
				ref: undefined,
				value: { kind: 'string', value: 'x' },
				formula: null,
				formulaBinding: null,
			},
			{
				row: 2,
				col: 1,
				ref: undefined,
				value: { kind: 'string', value: 'y' },
				formula: null,
				formulaBinding: null,
			},
		])
	})

	test('sheet handle can expose range rows and object rows for extraction workflows', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'name' },
					{ ref: 'B1', value: 'score' },
					{ ref: 'A2', value: 'Ada' },
					{ ref: 'B2', value: 42 },
				],
			},
		])
		const rows = wb.sheet('Sheet1')?.readRows('A1:B2')
		expect(rows?.rows).toEqual([
			[
				{ kind: 'string', value: 'name' },
				{ kind: 'string', value: 'score' },
			],
			[
				{ kind: 'string', value: 'Ada' },
				{ kind: 'number', value: 42 },
			],
		])
		const objects = wb.sheet('Sheet1')?.readObjects('A1:B2')
		expect(objects?.headers).toEqual(['name', 'score'])
		expect(objects?.rows).toEqual([
			{ name: { kind: 'string', value: 'Ada' }, score: { kind: 'number', value: 42 } },
		])
	})

	test('apply operations modifies workbook', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 100 }] },
		])
		expect(result.errors).toHaveLength(0)
		expect(result.affectedCells).toContain('A1')
		expect(result.sheetsModified).toContain('Sheet1')

		const cell = wb.sheet('Sheet1')?.cell('A1')
		expect(cell?.value).toEqual({ kind: 'number', value: 100 })
	})

	test('setCells returns warnings when data validation fails (write still succeeds)', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1',
				rule: { type: 'whole', formula1: '1', formula2: '10', operator: 'between' },
			},
		])
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] },
		])
		expect(result.errors).toHaveLength(0)
		expect(result.warnings).toHaveLength(1)
		expect(result.warnings?.[0]?.code).toBe('VALIDATION_ERROR')
		expect(result.warnings?.[0]?.refs).toContain('A1')
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 99 })
	})

	test('setCells with valid value has no warnings', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1',
				rule: { type: 'whole', formula1: '1', formula2: '10', operator: 'between' },
			},
		])
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		expect(result.errors).toHaveLength(0)
		expect(result.warnings).toBeUndefined()
	})

	test('apply returns error for invalid sheet', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
		])
		expect(result.errors).toHaveLength(1)
		expect(result.errors[0]?.code).toBe('SHEET_NOT_FOUND')
	})

	test('failed apply does not partially mutate workbook state', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'keep' }] },
			{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 'break' }] },
		])
		expect(result.errors).toHaveLength(1)
		expect(wb.sheet('Sheet1')?.cell('A1')).toBeUndefined()
	})

	test('apply with transaction:true rolls back on failure', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'first' }] },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 'second' }] },
				{ op: 'deleteSheet', sheet: 'NonExistent' },
			],
			{ transaction: true },
		)
		expect(result.errors).toHaveLength(1)
		expect(wb.sheet('Sheet1')?.cell('A1')).toBeUndefined()
		expect(wb.sheet('Sheet1')?.cell('A2')).toBeUndefined()
	})

	test('apply with collectAllErrors returns all validation errors', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply(
			[
				{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
				{ op: 'addSheet', name: 'Sheet1' },
				{ op: 'deleteSheet', sheet: 'NonExistent' },
			],
			{ collectAllErrors: true },
		)
		expect(result.errors).toHaveLength(3)
		expect(result.errors[0]?.code).toBe('SHEET_NOT_FOUND')
		expect(result.errors[0]?.message).toContain('Missing')
		expect(result.errors[1]?.code).toBe('NAME_CONFLICT')
		expect(result.errors[1]?.message).toContain('Sheet1')
		expect(result.errors[2]?.code).toBe('SHEET_NOT_FOUND')
		expect(result.errors[2]?.message).toContain('NonExistent')
	})

	test('recalc updates formula values', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'A2', value: 20 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A3', formula: 'A1+A2' },
		])
		const result = wb.recalc()
		expect(result.changed.length).toBeGreaterThanOrEqual(1)

		const a3 = wb.sheet('Sheet1')?.cell('A3')
		expect(a3).toBeDefined()
		expect(a3?.value).toEqual({ kind: 'number', value: 30 })
	})

	test('recalc persists clean calc freshness when formula values are unchanged', async () => {
		const sourceBytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr fullCalcOnLoad="1" calcCompleted="0" calcOnSave="0" forceFullCalc="1" calcId="191029"/>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>
</worksheet>`,
		})
		const wb = await AscendWorkbook.open(sourceBytes)

		const result = wb.recalc()
		expect(result.errors).toHaveLength(0)
		expect(result.changed).toHaveLength(0)
		const out = wb.toBytes()

		expect(out).not.toEqual(sourceBytes)
		const archive = extractZip(out)
		const workbookXml = archive.readText('xl/workbook.xml') ?? ''
		expect(workbookXml).not.toContain('fullCalcOnLoad="1"')
		expect(workbookXml).not.toContain('calcCompleted="0"')
		expect(workbookXml).not.toContain('calcOnSave="0"')
		expect(workbookXml).not.toContain('forceFullCalc="1"')
		expect(workbookXml).toContain('calcCompleted="1"')
		expect(workbookXml).toContain('calcOnSave="1"')
		expect(workbookXml).toContain('forceFullCalc="0"')
	})

	test('recalc can resolve external workbook references with a caller hook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=[Budget.xlsx]Inputs!B2' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: '=SUM([Budget.xlsx]Inputs!A1:A3)' },
		])

		const result = wb.recalc({
			externalReferences: {
				resolveCell: ({ workbook, sheet, row, col }) => {
					if (workbook !== 'Budget.xlsx' || sheet !== 'Inputs') return undefined
					if (row === 1 && col === 1) return { kind: 'number', value: 42 }
					if (col === 0) return { kind: 'number', value: row + 1 }
					return undefined
				},
			},
		})

		expect(result.errors).toEqual([])
		expect(wb.sheet('Sheet1').cell('A1')?.value).toEqual({ kind: 'number', value: 42 })
		expect(wb.sheet('Sheet1').cell('A2')?.value).toEqual({ kind: 'number', value: 6 })
	})

	test('applyAndRecalc applies operations and returns recalculation details', () => {
		const wb = AscendWorkbook.create()
		const result = wb.applyAndRecalc([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'A2', value: 20 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A3', formula: 'A1+A2' },
		])

		expect(result.apply.errors).toHaveLength(0)
		expect(result.apply.recalcRequired).toBe(true)
		expect(result.recalc?.errors).toHaveLength(0)
		expect(result.recalc?.changed.length).toBeGreaterThanOrEqual(1)
		expect(wb.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'number', value: 30 })
	})

	test('applyAndRecalc skips recalculation when apply fails', () => {
		const wb = AscendWorkbook.create()
		let recalcCount = 0
		const originalRecalc = wb.recalc.bind(wb)
		wb.recalc = (...args: Parameters<typeof wb.recalc>) => {
			recalcCount++
			return originalRecalc(...args)
		}
		const result = wb.applyAndRecalc([
			{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
		])

		expect(result.apply.errors).toHaveLength(1)
		expect(result.recalc).toBeNull()
		expect(recalcCount).toBe(0)
	})

	test('batch(ops) applies multiple operations without recalc, caller recalcs once', () => {
		const wb = AscendWorkbook.create()
		const batchResult = wb.batch([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 10 }] },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 20 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A3', formula: 'A1+A2' },
		])
		expect(batchResult.errors).toHaveLength(0)

		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 10 })
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 20 })
		expect(wb.sheet('Sheet1')?.cell('A3')?.formula).toBe('A1+A2')

		const recalcResult = wb.recalc()
		expect(recalcResult.errors).toHaveLength(0)
		expect(wb.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'number', value: 30 })
	})

	test('dumpBatch emits deterministic replay ops and reports unsupported values', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'B1', value: 'label' },
					{ ref: 'A2', value: true },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: 'A1*2' },
			{ op: 'setRichText', sheet: 'Sheet1', ref: 'C1', runs: [{ text: 'rich', bold: true }] },
		])

		const dump = wb.dumpBatch()
		expect(dump.replayable).toBe(false)
		expect(dump.cellCount).toBe(4)
		expect(dump.formulaCount).toBe(1)
		expect(dump.unsupported).toEqual([
			expect.objectContaining({
				sheet: 'Sheet1',
				ref: 'C1',
				valueKind: 'richText',
			}),
		])
		expect(dump.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'B1', value: 'label' },
					{ ref: 'A2', value: true },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: 'A1*2' },
		])

		const replayed = AscendWorkbook.create()
		expect(replayed.batch(dump.ops).errors).toEqual([])
		expect(replayed.recalc().errors).toEqual([])
		expect(replayed.get('Sheet1!B2')).toEqual({ kind: 'number', value: 20 })
	})

	test('batch(fn) defers recalc until all mutations complete, triggers single recalc', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A4', formula: 'A1+A2+A3' }])

		let recalcCount = 0
		const originalRecalc = wb.recalc.bind(wb)
		wb.recalc = (...args: Parameters<typeof wb.recalc>) => {
			recalcCount++
			return originalRecalc(...args)
		}

		wb.batch(() => {
			wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
			wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 2 }] }])
			wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A3', value: 3 }] }])
		})

		expect(recalcCount).toBe(1)
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 1 })
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 2 })
		expect(wb.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'number', value: 3 })
		expect(wb.sheet('Sheet1')?.cell('A4')?.value).toEqual({ kind: 'number', value: 6 })
	})

	test('save to bytes and re-open roundtrips', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'roundtrip' },
					{ ref: 'B1', value: 42 },
				],
			},
		])

		const bytes = wb.toBytes()
		expect(bytes).toBeInstanceOf(Uint8Array)
		expect(bytes.length).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(bytes)
		const cell = reopened.sheet('Sheet1')?.cell('A1')
		expect(cell).toBeDefined()
		expect(cell?.value).toEqual({ kind: 'string', value: 'roundtrip' })

		const b1 = reopened.sheet('Sheet1')?.cell('B1')
		expect(b1).toBeDefined()
		expect(b1?.value).toEqual({ kind: 'number', value: 42 })
	})

	test('setRichText saves to XLSX bytes and reopens with formatting runs', async () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{
				op: 'setRichText',
				sheet: 'Sheet1',
				ref: 'A1',
				runs: [
					{ text: 'Bold', bold: true },
					{ text: ' Italic', italic: true },
					{ text: ' Blue', color: 'FF0000FF' },
				],
			},
		])
		expect(result.errors).toHaveLength(0)

		const bytes = wb.toBytes()
		const archive = extractZip(bytes)
		const sheetXml = archive.readText('xl/worksheets/sheet1.xml') ?? ''
		const sharedStringsXml = archive.readText('xl/sharedStrings.xml') ?? ''
		expect(sheetXml).toContain('t="s"')
		expect(sharedStringsXml).toContain('<b/>')
		expect(sharedStringsXml).toContain('<i/>')
		expect(sharedStringsXml).toContain('rgb="FF0000FF"')

		const reopened = await AscendWorkbook.open(bytes)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'richText',
			runs: [
				{ text: 'Bold', bold: true },
				{ text: ' Italic', italic: true },
				{ text: ' Blue', color: 'FF0000FF' },
			],
		})
	})

	test('setHyperlink saves external and internal links to XLSX bytes', async () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply([
			{ op: 'addSheet', name: 'Sheet2' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'External' }] },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 'Internal' }] },
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'A1',
				url: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'B1',
				location: 'Sheet2!A1',
				display: 'Sheet2',
				tooltip: 'Jump inside workbook',
			},
		])
		expect(result.errors).toHaveLength(0)

		const bytes = wb.toBytes()
		const archive = extractZip(bytes)
		const sheetXml = archive.readText('xl/worksheets/sheet1.xml') ?? ''
		const sheetRelsXml = archive.readText('xl/worksheets/_rels/sheet1.xml.rels') ?? ''
		expect(sheetXml).toContain('ref="A1"')
		expect(sheetXml).toContain('display="Report"')
		expect(sheetXml).toContain('tooltip="Open report"')
		expect(sheetXml).toContain('ref="B1"')
		expect(sheetXml).toContain('location="Sheet2!A1"')
		expect(sheetXml).toContain('tooltip="Jump inside workbook"')
		expect(sheetRelsXml).toContain('Target="https://example.com/report"')
		expect(sheetRelsXml).toContain('TargetMode="External"')
		expect(sheetRelsXml).not.toContain('Target="Sheet2!A1"')

		const reopened = await AscendWorkbook.open(bytes)
		expect(reopened.sheet('Sheet1')?.hyperlink('A1')).toEqual({
			target: 'https://example.com/report',
			display: 'Report',
			tooltip: 'Open report',
		})
		expect(reopened.sheet('Sheet1')?.hyperlink('B1')).toEqual({
			location: 'Sheet2!A1',
			display: 'Sheet2',
			tooltip: 'Jump inside workbook',
		})
		expect(reopened.sheet('Sheet1')?.getHyperlinks()).toContainEqual({
			ref: 'B1',
			location: 'Sheet2!A1',
			display: 'Sheet2',
			tooltip: 'Jump inside workbook',
		})
	})

	test('dirty serialization rebases source bytes and clears dirty flags', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'hello' }] }])

		const bytes = wb.toBytes()
		const internal = wb as unknown as {
			originalBytes: Uint8Array | null
			dirty: boolean
			dirtySheets: Set<string>
			workbookMetaDirty: boolean
			sharedStringsDirty: boolean
			wb: { sourceArchiveBytes: Uint8Array | null }
		}

		expect(internal.dirty).toBe(false)
		expect(internal.originalBytes).toBe(bytes)
		expect(internal.wb.sourceArchiveBytes).toBe(bytes)
		expect(internal.dirtySheets.size).toBe(0)
		expect(internal.workbookMetaDirty).toBe(false)
		expect(internal.sharedStringsDirty).toBe(false)
		expect(wb.toBytes()).toBe(bytes)
	})

	test('value edits preserve imported calcChain while marking formula caches stale', async () => {
		const sourceBytes = makeSyntheticXlsx({
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
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="0" calcCompleted="1" calcOnSave="0" calcId="191029"/>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>
</worksheet>`,
			'xl/calcChain.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>`,
		})
		const wb = await AscendWorkbook.open(sourceBytes)
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Calc',
				updates: [
					{ ref: 'A1', value: 3 },
					{ ref: 'C3', value: '__ascend_feature_contract__' },
				],
			},
		])

		const archive = extractZip(wb.toBytes())
		expect(archive.has('xl/calcChain.xml')).toBe(true)
		expect(archive.readText('xl/_rels/workbook.xml.rels')).toContain('relationships/calcChain')
		const workbookXml = archive.readText('xl/workbook.xml') ?? ''
		expect(workbookXml).toContain('calcCompleted="0"')
		expect(workbookXml).toContain('forceFullCalc="1"')
	})

	test('workbook metadata edits regenerate workbook XML without forcing stale calc state', async () => {
		const sourceBytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr codeName="OldBook"/>
  <bookViews><workbookView activeTab="0" firstSheet="0" tabRatio="600"/></bookViews>
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="auto" fullCalcOnLoad="0" calcCompleted="1" calcOnSave="1" forceFullCalc="0" calcId="1"/>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})
		const wb = await AscendWorkbook.open(sourceBytes)
		const apply = wb.apply([
			{ op: 'setWorkbookProperties', properties: { codeName: 'NewBook' } },
			{ op: 'setWorkbookView', view: { activeTab: 0, firstSheet: 0, tabRatio: 720 } },
			{
				op: 'setCalcSettings',
				settings: {
					calcMode: 'manual',
					fullCalcOnLoad: false,
					calcCompleted: true,
					calcOnSave: false,
					forceFullCalc: false,
					calcId: 42,
				},
			},
		])
		expect(apply.errors).toHaveLength(0)

		const saved = wb.toBytes()
		const workbookXml = extractZip(saved).readText('xl/workbook.xml') ?? ''
		expect(workbookXml).toContain('codeName="NewBook"')
		expect(workbookXml).toContain('tabRatio="720"')
		expect(workbookXml).toContain('calcMode="manual"')
		expect(workbookXml).not.toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="1"')
		expect(workbookXml).toContain('calcOnSave="0"')
		expect(workbookXml).toContain('forceFullCalc="0"')
		expect(workbookXml).toContain('calcId="42"')

		const reopened = await AscendWorkbook.open(saved)
		expect(reopened.workbookViews()[0]?.tabRatio).toBe(720)
		expect(
			(reopened.toJSON() as { calcSettings: { calcMode: string; calcId?: number } }).calcSettings,
		).toMatchObject({ calcMode: 'manual', calcId: 42 })
	})

	test('value edits drop calcChain when replacing an existing formula cell', async () => {
		const sourceBytes = makeSyntheticXlsx({
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
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="0" calcCompleted="1" calcOnSave="0" calcId="191029"/>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>
</worksheet>`,
			'xl/calcChain.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>`,
		})
		const wb = await AscendWorkbook.open(sourceBytes)
		wb.apply([{ op: 'setCells', sheet: 'Calc', updates: [{ ref: 'B1', value: 123 }] }])

		const saved = wb.toBytes()
		const archive = extractZip(saved)
		expect(archive.has('xl/calcChain.xml')).toBe(false)
		expect(archive.readText('xl/_rels/workbook.xml.rels')).not.toContain('relationships/calcChain')
		expect(archive.readText('[Content_Types].xml')).not.toContain('/xl/calcChain.xml')

		const reopened = await AscendWorkbook.open(saved)
		expect(reopened.sheet('Calc')?.cell('B1')?.formula).toBeNull()
		expect(reopened.sheet('Calc')?.cell('B1')?.value).toEqual({ kind: 'number', value: 123 })
	})

	test('reusing an existing shared string avoids dirtying the shared string table', async () => {
		const bytes = makeSyntheticXlsx({
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
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>World</t></si>
  <si><t>Hello</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
      <c r="B1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const wb = await AscendWorkbook.open(bytes)
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 'Hello' }] }])

		const internal = wb as unknown as { sharedStringsDirty: boolean }
		expect(internal.sharedStringsDirty).toBe(false)

		const saved = wb.toBytes()
		const archive = extractZip(saved)
		const sharedStringsXml = archive.readText('xl/sharedStrings.xml')
		expect(sharedStringsXml).toContain('<t>World</t>')
		expect(sharedStringsXml).toContain('<t>Hello</t>')
		const reopened = await AscendWorkbook.open(saved)
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'string', value: 'Hello' })
	})

	test('untouched imported xlsx returns original bytes exactly', async () => {
		const bytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" si="0">B1*2</f><v>84</v></c></row>
    <row r="2"><c r="A2"><f t="shared" si="0"/><v>168</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const reopened = await AscendWorkbook.open(bytes)
		expect(reopened.toBytes()).toEqual(bytes)
	})

	test('modified workbook preserves untouched shared-formula sheets when safe', async () => {
		const bytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Input" sheetId="1" r:id="rId1"/>
    <sheet name="Calc" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" si="0">B1*2</f><v>84</v></c></row>
    <row r="2"><c r="A2"><f t="shared" si="0"/><v>168</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const reopened = await AscendWorkbook.open(bytes)
		reopened.apply([{ op: 'setCells', sheet: 'Input', updates: [{ ref: 'A1', value: 2 }] }])
		const out = reopened.toBytes()
		const archive = extractZip(out)
		const calcSheetXml = archive.readText('xl/worksheets/sheet2.xml')
		expect(calcSheetXml).toBeDefined()
		expect(calcSheetXml?.match(/<f\b[^>]*\bt="shared"/g)?.length).toBe(2)
	})

	test('string edits still preserve untouched shared-formula sheets', async () => {
		const bytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Input" sheetId="1" r:id="rId1"/>
    <sheet name="Calc" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>old</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
</worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" si="0">B1*2</f><v>84</v></c></row>
    <row r="2"><c r="A2"><f t="shared" si="0"/><v>168</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const reopened = await AscendWorkbook.open(bytes)
		reopened.apply([{ op: 'setCells', sheet: 'Input', updates: [{ ref: 'A1', value: 'new' }] }])
		const out = reopened.toBytes()
		const archive = extractZip(out)
		const calcSheetXml = archive.readText('xl/worksheets/sheet2.xml')
		expect(calcSheetXml).toBeDefined()
		expect(calcSheetXml?.match(/<f\b[^>]*\bt="shared"/g)?.length).toBe(2)
	})

	test('non-formula edits preserve shared formulas on the same sheet', async () => {
		const bytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="shared" si="0" ref="A1:A2">B1*2</f><v>84</v></c>
      <c r="B1"><v>42</v></c>
    </row>
    <row r="2"><c r="A2"><f t="shared" si="0"/><v>168</v></c></row>
    <row r="3"><c r="C3"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
		})

		const reopened = await AscendWorkbook.open(bytes)
		reopened.apply([{ op: 'setCells', sheet: 'Calc', updates: [{ ref: 'C3', value: 2 }] }])
		const out = reopened.toBytes()
		const archive = extractZip(out)
		const calcSheetXml = archive.readText('xl/worksheets/sheet1.xml')
		expect(calcSheetXml).toBeDefined()
		expect(calcSheetXml?.match(/<f\b[^>]*\bt="shared"/g)?.length).toBe(2)
	})

	test('metadata-only open preserves workbook structure without parsing cells', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'summary' }] },
			{ op: 'addSheet', name: 'Archive' },
		])
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { mode: 'metadata-only' })
		expect(reopened.sheets).toEqual(['Sheet1', 'Archive'])
		expect(reopened.inspect().sheetCount).toBe(2)
		expect(reopened.inspect().loadedSheetCount).toBe(2)
		expect(reopened.inspect().cellCount).toBeNull()
		expect(reopened.inspect().load.mode).toBe('metadata-only')
		expect(reopened.inspect().load.isPartial).toBe(true)
		expect(reopened.inspect().sheets[0]?.cellDataLoaded).toBe(false)
		expect(reopened.sheet('Sheet1')?.cell('A1')).toBeUndefined()
		const metadataInternal = reopened as unknown as {
			originalBytes: Uint8Array | null
			caps: readonly unknown[]
			wb: { sourceArchiveBytes: Uint8Array | null }
		}
		expect(metadataInternal.originalBytes).toBeNull()
		expect(metadataInternal.caps).toHaveLength(0)
		expect(metadataInternal.wb.sourceArchiveBytes).toBeNull()
		expect(() => reopened.toBytes()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('values mode opens hydrated cells as a read-only partial view', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*3' },
		])
		wb.recalc()
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { mode: 'values' })
		expect(reopened.inspect().load.mode).toBe('values')
		expect(reopened.inspect().load.isPartial).toBe(true)
		expect(reopened.inspect().load.cellsHydrated).toBe(true)
		expect(reopened.inspect().load.richSheetMetadataHydrated).toBe(false)
		expect(reopened.inspect().commentCount).toBeNull()
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 6 })
		expect(reopened.sheet('Sheet1')?.cell('A2')?.formula).toBeNull()
		const valuesInternal = reopened as unknown as {
			originalBytes: Uint8Array | null
			caps: readonly unknown[]
			wb: { sourceArchiveBytes: Uint8Array | null }
		}
		expect(valuesInternal.originalBytes).toBeNull()
		expect(valuesInternal.caps).toHaveLength(0)
		expect(valuesInternal.wb.sourceArchiveBytes).toBeNull()
		expect(() => reopened.toBytes()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('maxRows opens a row-capped partial workbook view through the SDK', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'loaded' },
					{ ref: 'A2', value: 'not loaded' },
				],
			},
		])
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { mode: 'values', maxRows: 1 })
		expect(reopened.inspect().load.mode).toBe('values')
		expect(reopened.inspect().load.isPartial).toBe(true)
		expect(reopened.inspect().load.maxRows).toBe(1)
		expect(reopened.inspect().load.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'loaded',
		})
		expect(reopened.sheet('Sheet1')?.cell('A2')).toBeUndefined()
		const check = reopened.check()
		expect(check.valid).toBe(false)
		expect(check.issues[0]?.message).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
		const lint = reopened.lint()
		expect(lint.clean).toBe(false)
		expect(lint.warnings[0]?.rule).toBe('partial-dependency-analysis')
		expect(lint.warnings[0]?.message).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
		expect(() => reopened.toBytes()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('values mode can expose rich sheet metadata as an opt-in partial view', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] }])
		const internal = wb as unknown as {
			wb: {
				definedNames: { set(name: string, formula: string): void }
				sheets: Array<{
					comments: Map<string, { text: string; author?: string }>
					hyperlinks: Map<string, { target?: string; display?: string }>
					dataValidations: Array<Record<string, unknown>>
					conditionalFormats: Array<Record<string, unknown>>
				}>
			}
		}
		const sheet = internal.wb.sheets[0]
		sheet?.comments.set('B2', { text: 'Review', author: 'Ada' })
		sheet?.hyperlinks.set('A1', { target: 'https://example.com/ascend', display: 'Ascend' })
		sheet?.dataValidations.push({ sqref: 'B2', type: 'list', formula1: '"A,B"' })
		sheet?.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'cellIs', operator: 'greaterThan', formulas: ['0'] }],
		})
		internal.wb.definedNames.set('FeatureRange', 'Sheet1!$A$1:$B$2')
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { mode: 'values', richMetadata: true })
		expect(reopened.inspect().load.mode).toBe('values')
		expect(reopened.inspect().load.isPartial).toBe(true)
		expect(reopened.inspect().load.richSheetMetadataHydrated).toBe(true)
		expect(reopened.inspect().commentCount).toBe(1)
		expect(reopened.inspect().dataValidationCount).toBe(1)
		expect(reopened.inspect().conditionalFormatCount).toBe(1)
		const detail = reopened.inspectSheet('Sheet1')
		expect(detail?.comments?.[0]).toMatchObject({
			ref: 'B2',
			author: 'Ada',
			text: 'Review',
			legacyDrawing: {
				row: 1,
				column: 1,
				visible: false,
			},
		})
		expect(detail?.hyperlinks?.[0]?.target).toBe('https://example.com/ascend')
		expect(detail?.dataValidations).toHaveLength(1)
		expect(detail?.conditionalFormats).toHaveLength(1)
		expect(reopened.definedNames()[0]?.name).toBe('FeatureRange')
		expect(() => reopened.toBytes()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('formula mode preserves formulas in a read-only partial view', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*3' },
		])
		wb.recalc()
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { mode: 'formula' })
		expect(reopened.inspect().load.mode).toBe('formula')
		expect(reopened.inspect().load.isPartial).toBe(true)
		expect(reopened.inspect().load.cellsHydrated).toBe(true)
		expect(reopened.inspect().load.richSheetMetadataHydrated).toBe(false)
		expect(reopened.inspect().commentCount).toBeNull()
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 6 })
		expect(reopened.sheet('Sheet1')?.cell('A2')?.formula).toBe('A1*3')
		expect(reopened.formula('Sheet1!A2')?.normalizedFormula).toBe('A1*3')
		const formulaInternal = reopened as unknown as {
			originalBytes: Uint8Array | null
			caps: readonly unknown[]
			wb: { sourceArchiveBytes: Uint8Array | null }
		}
		expect(formulaInternal.originalBytes).toBeNull()
		expect(formulaInternal.caps).toHaveLength(0)
		expect(formulaInternal.wb.sourceArchiveBytes).toBeNull()
		expect(() => reopened.toBytes()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('selective open parses only requested sheets', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'main' }] },
			{ op: 'addSheet', name: 'Archive' },
			{ op: 'setCells', sheet: 'Archive', updates: [{ ref: 'A1', value: 'extra' }] },
		])
		const bytes = wb.toBytes()

		const reopened = await AscendWorkbook.open(bytes, { sheets: ['Archive'] })
		expect(reopened.sheets).toEqual(['Archive'])
		expect(reopened.inspect().sheetCount).toBe(2)
		expect(reopened.inspect().loadedSheetCount).toBe(1)
		expect(reopened.inspect().load.mode).toBe('selective')
		expect(reopened.inspect().load.isPartial).toBe(true)
		expect(reopened.sheet('Archive')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'extra' })
		const selectiveInternal = reopened as unknown as {
			originalBytes: Uint8Array | null
			caps: readonly unknown[]
			wb: { sourceArchiveBytes: Uint8Array | null }
		}
		expect(selectiveInternal.originalBytes).toBeNull()
		expect(selectiveInternal.caps).toHaveLength(0)
		expect(selectiveInternal.wb.sourceArchiveBytes).toBeNull()
		expect(() => reopened.toCsv()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('selective open preserves only loaded sheet-scoped defined names', async () => {
		const bytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <definedNames>
    <definedName name="GlobalRange">Archive!A1</definedName>
    <definedName name="MainLocal" localSheetId="0">Main!A1</definedName>
    <definedName name="ArchiveLocal" localSheetId="1">Archive!A1</definedName>
  </definedNames>
  <sheets>
    <sheet name="Main" sheetId="1" r:id="rId1"/>
    <sheet name="Archive" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>
</worksheet>`,
		})

		const reopened = await AscendWorkbook.open(bytes, { sheets: ['Archive'], mode: 'formula' })
		expect(reopened.names).toContain('GlobalRange')
		expect(reopened.definedName('ArchiveLocal', 'Archive')?.formula).toBe('Archive!A1')
		expect(reopened.definedName('MainLocal', 'Archive')).toBeUndefined()
		expect(reopened.definedNames().map((entry) => entry.name)).not.toContain('MainLocal')
	})

	test('partial dependency verification fails honestly', async () => {
		const bytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Main" sheetId="1" r:id="rId1"/>
    <sheet name="Archive" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>
</worksheet>`,
			'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><f>Main!A1*2</f><v>4</v></c></row></sheetData>
</worksheet>`,
		})

		const reopened = await AscendWorkbook.open(bytes, { sheets: ['Archive'], mode: 'formula' })
		const check = reopened.check()
		expect(check.valid).toBe(false)
		expect(check.issues[0]?.rule).toBe('partial-dependency-analysis')
		expect(check.issues[0]?.suggestedFix).toContain('all referenced sheets')
		expect(check.issues[0]?.message).toContain(
			'Cannot verify workbook dependencies from this partial view',
		)
		const lint = reopened.lint()
		expect(lint.clean).toBe(false)
		expect(lint.warnings[0]?.rule).toBe('partial-dependency-analysis')
		expect(lint.warnings[0]?.message).toContain('not all sheets are loaded')
		const traceIssue = reopened.traceIssue('Archive!A1')
		expect(traceIssue?.rule).toBe('partial-dependency-analysis')
		expect(traceIssue?.message).toContain('not all sheets are loaded')
		expect(reopened.trace('Archive!A1')).toBeUndefined()
	})

	test('partial workbook views reject preview, apply, and recalc operations', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'locked' }] }])
		const bytes = wb.toBytes()
		const reopened = await AscendWorkbook.open(bytes, { mode: 'values', maxRows: 1 })

		const preview = reopened.preview([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'next' }] },
		])
		expect(preview.errors[0]?.message).toContain('Cannot modify a partial workbook view')
		expect(preview.errors[0]?.details).toMatchObject({
			partialWorkbookView: true,
			mode: 'values',
			maxRows: 1,
			requiredLoad: { mode: 'full', allSheets: true, maxRows: null },
		})
		expect(preview.errors[0]?.details?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)

		const apply = reopened.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'next' }] },
		])
		expect(apply.errors[0]?.message).toContain('Cannot modify a partial workbook view')
		expect(apply.errors[0]?.details?.partialWorkbookView).toBe(true)
		expect(apply.errors[0]?.details?.loadedSheets).toEqual(['Sheet1'])

		const recalc = reopened.recalc()
		expect(recalc.errors[0]?.error.message).toContain('Cannot modify a partial workbook view')
		expect(recalc.errors[0]?.error.details?.partialWorkbookView).toBe(true)
		expect(recalc.errors[0]?.error.details?.maxRows).toBe(1)
	})

	test('CSV import creates workbook', () => {
		const csv = 'Name,Age\nAlice,30\nBob,25'
		const wb = AscendWorkbook.fromCsv(csv)
		expect(wb.sheets).toEqual(['Sheet1'])

		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		expect(handle?.cell('A1')?.value).toEqual({ kind: 'string', value: 'Name' })
		expect(handle?.cell('B2')?.value).toEqual({ kind: 'number', value: 30 })
	})

	test('sheet handles stay current across workbook replacement', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'before' }] }])
		const handle = wb.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 handle')

		expect(handle.cell('A1')?.value).toEqual({ kind: 'string', value: 'before' })

		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'after' }] }])

		expect(handle.cell('A1')?.value).toEqual({ kind: 'string', value: 'after' })
	})

	test('CSV export works', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'x' },
					{ ref: 'B1', value: 1 },
					{ ref: 'A2', value: 'y' },
					{ ref: 'B2', value: 2 },
				],
			},
		])
		const csv = wb.toCsv()
		expect(csv).toContain('x')
		expect(csv).toContain('1')
		expect(csv).toContain('y')
		expect(csv).toContain('2')
	})

	test('preview shows changes without modifying workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'before' }] }])

		const preview = wb.preview([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'after' }] },
		])
		expect(preview.errors).toHaveLength(0)
		expect(preview.cellChanges.length).toBeGreaterThanOrEqual(1)
		expect(preview.changedCells).toHaveLength(1)
		expect(preview.changedCells[0]).toEqual({
			ref: 'Sheet1!A1',
			oldValue: { kind: 'string', value: 'before' },
			newValue: { kind: 'string', value: 'after' },
		})
		expect(preview.wouldSucceed).toBe(true)
		expect(preview.writePlan?.totalParts).toBeGreaterThan(0)

		const cell = wb.sheet('Sheet1')?.cell('A1')
		expect(cell?.value).toEqual({ kind: 'string', value: 'before' })
	})

	test('preview surfaces recalc errors from derived workbook state', () => {
		const wb = AscendWorkbook.create()
		const preview = wb.preview([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }])
		expect(preview.errors.some((error) => error.code === 'CIRCULAR_REF')).toBe(true)
		expect(preview.wouldSucceed).toBe(false)
	})

	test('preview shows recalc scope for formula dependents', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 10 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1*2' },
		])

		const preview = wb.preview([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 20 }] },
		])
		expect(preview.errors).toHaveLength(0)
		expect(preview.wouldSucceed).toBe(true)
		expect(preview.recalcScope).toBeGreaterThanOrEqual(1)
		expect(preview.changedCells.some((c) => c.ref === 'Sheet1!B1')).toBe(true)
		expect(preview.changedCells.find((c) => c.ref === 'Sheet1!B1')?.newValue).toEqual({
			kind: 'number',
			value: 40,
		})
	})

	test('preview captures validation warnings', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A5',
				rule: { type: 'whole', formula1: '1', formula2: '10', operator: 'between' },
			},
		])

		const preview = wb.preview([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] },
		])
		expect(preview.warnings.length).toBeGreaterThanOrEqual(1)
		expect(preview.warnings.some((w) => w.code === 'VALIDATION_ERROR')).toBe(true)
		expect(preview.wouldSucceed).toBe(true)
	})

	test('preview formatting changes do not grow the source style registry', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 0.25 }] }])
		const internal = wb as unknown as { wb: { styles: { size: number } } }
		const beforeSize = internal.wb.styles.size

		const preview = wb.preview([
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'A1:A1', format: '0.0%' },
		])
		expect(preview.errors).toHaveLength(0)
		expect(internal.wb.styles.size).toBe(beforeSize)
	})

	test('setNumberFormat preserves imported style metadata through SDK writes', async () => {
		const sourceBytes = makeSyntheticXlsx({
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

		const wb = await AscendWorkbook.open(sourceBytes)
		wb.apply([{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'A1:A1', format: '0.0%' }])

		const archive = extractZip(wb.toBytes())
		const stylesXml = archive.readText('xl/styles.xml') ?? ''
		expect(stylesXml).toContain('formatCode="0.0%"')
		expect(stylesXml).toContain('<cellStyleXfs')
		expect(stylesXml).toContain('<cellStyles')
		expect(stylesXml).toContain('defaultTableStyle="TableStyleMedium2"')
		expect(stylesXml).toContain('defaultPivotStyle="PivotStyleLight16"')
		expect(stylesXml).toContain('<tableStyle name="TableStyleMedium2"/>')
	})

	test('preview and writePlanSummary do not retain a cached source archive', async () => {
		const sourceBytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})
		const wb = await AscendWorkbook.open(sourceBytes)
		const internal = wb as unknown as { sourceArchive?: unknown }

		expect(internal.sourceArchive).toBeUndefined()
		wb.preview([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'preview' }] }])
		expect(internal.sourceArchive).toBeUndefined()
		wb.writePlanSummary()
		expect(internal.sourceArchive).toBeUndefined()
		wb.toBytes()
		expect(internal.sourceArchive).toBeUndefined()
	})

	test('check returns clean for valid workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		const result = wb.check()
		expect(result.valid).toBe(true)
		expect(result.issues).toHaveLength(0)
	})

	test('check surfaces source package graph integrity issues', async () => {
		const bytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdXmlMaps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps" Target="xmlMaps.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
		})

		const workbook = await AscendWorkbook.open(bytes)
		const document = await WorkbookDocument.open(bytes)

		for (const result of [workbook.check(), document.check()]) {
			const issue = result.issues.find((entry) => entry.rule === 'package-graph-integrity')
			expect(result.valid).toBe(false)
			expect(issue).toMatchObject({
				severity: 'error',
				refs: ['xl/_rels/workbook.xml.rels#rIdXmlMaps'],
				details: {
					code: 'package_relationship_target',
					sourcePartPath: 'xl/workbook.xml',
					relationshipId: 'rIdXmlMaps',
					expected: 'xmlMaps.xml',
					actual: 'xl/xmlMaps.xml',
				},
			})
			expect(issue?.suggestedFix).toContain('restore the referenced package part')
		}
	})

	test('check preserves structured verification metadata for agent repair', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'renameSheet', sheet: 'Sheet1', newName: 'SummaryData' }])
		wb.apply([{ op: 'setFormula', sheet: 'SummaryData', ref: 'A1', formula: '=Summary!B1' }])

		const result = wb.check()
		expect(result.valid).toBe(false)
		const issue = result.issues.find((entry) => entry.rule === 'broken-refs')
		expect(issue).toBeDefined()
		expect(issue?.severity).toBe('error')
		expect(issue?.ref).toBe('SummaryData!A1')
		expect(issue?.refs).toEqual(['SummaryData!A1'])
		expect(issue?.message).toContain('Summary')
		expect(issue?.suggestedFix).toContain('SummaryData')
	})

	test('check exposes blocked spill causes for agent repair', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SEQUENCE(3)' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 'blocker' }] },
		])
		wb.recalc()

		const result = wb.check()
		const issue = result.issues.find((entry) => entry.rule === 'spill-diagnostics')
		expect(result.valid).toBe(false)
		expect(issue?.ref).toBe('Sheet1!A1')
		expect(issue?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(issue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
		expect(wb.sheet('Sheet1').getFormulaBinding('A1')).toEqual({
			kind: 'blocked-spill',
			formula: 'SEQUENCE(3)',
			range: 'A1:A3',
			blockingRefs: ['A2'],
		})
	})

	test('lint detects formula parse errors', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=INVALID((' }])
		const result = wb.lint()
		expect(result.clean).toBe(false)
		expect(result.warnings.length).toBeGreaterThanOrEqual(1)
	})

	test('diff detects changes between workbooks', () => {
		const wb1 = AscendWorkbook.create()
		wb1.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const wb2 = AscendWorkbook.create()
		wb2.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] }])

		const diff = wb1.diff(wb2)
		expect(diff.sheets.length).toBeGreaterThanOrEqual(1)
	})

	test('snapshot captures workbook state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] }])
		const snap = wb.snapshot()
		expect(snap.sheets).toHaveLength(1)
		expect(snap.sheets[0]?.name).toBe('Sheet1')
		expect(snap.sheets[0]?.cells.A1?.value).toEqual({ kind: 'number', value: 99 })
		expect(snap.timestamp).toBeGreaterThan(0)
	})

	test('toJSON returns serializable structure', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'json' }] }])
		const json = wb.toJSON()
		const parsed = JSON.parse(JSON.stringify(json))
		expect(parsed.sheets).toBeDefined()
		expect(parsed.sheets[0].name).toBe('Sheet1')
	})

	test('writePlanSummary classifies generated and preserved write parts', async () => {
		const sourceBytes = makeSyntheticXlsx({
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
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
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
		})
		const wb = await AscendWorkbook.open(sourceBytes)
		const summary = wb.writePlanSummary()
		expect(summary.totalParts).toBeGreaterThan(0)
		expect(summary.byOrigin['preserved-source']).toBeGreaterThan(0)
		expect(summary.byOrigin.generated).toBeGreaterThan(0)
		expect(summary.byOwnerKind.workbook).toBeGreaterThan(0)
	})

	test('rawPackagePart returns bounded source text with package metadata', async () => {
		const binaryBytes = new Uint8Array([0, 1, 2, 3, 4, 255])
		const sourceBytes = makeSyntheticXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/media/image1.png': binaryBytes,
		})
		const wb = await AscendWorkbook.open(sourceBytes)

		const raw = wb.rawPackagePart({ partPath: '/xl/workbook.xml', maxBytes: 80 })

		expect(raw.found).toBe(true)
		expect(raw.partPath).toBe('xl/workbook.xml')
		expect(raw.validPath).toBe(true)
		expect(raw.normalizedFromRoot).toBe(true)
		expect(raw.origin).toBe('source')
		expect(raw.load?.mode).toBe('full')
		expect(raw.load?.isPartial).toBe(false)
		expect(raw.semantics).toBe('raw-package-bytes')
		expect(raw.featureFamily).toBe('workbook')
		expect(raw.ownerScope).toBe('workbook')
		expect(raw.byteLength).toBeGreaterThan(80)
		expect(raw.sha256).toMatch(/^[a-f0-9]{64}$/)
		expect(raw.truncated).toBe(true)
		expect(raw.previewByteLength).toBe(80)
		expect(raw.binaryLike).toBe(false)
		expect(raw.text).toContain('<?xml')
		expect(raw.text?.length).toBeLessThan(raw.byteLength ?? 0)

		const binary = wb.rawPackagePart({
			partPath: 'xl/media/image1.png',
			encoding: 'base64',
			maxBytes: 3,
		})
		expect(binary.found).toBe(true)
		expect(binary.encoding).toBe('base64')
		expect(binary.byteLength).toBe(binaryBytes.byteLength)
		expect(binary.previewByteLength).toBe(3)
		expect(binary.truncated).toBe(true)
		expect(binary.binaryLike).toBe(true)
		expect(binary.base64).toBe(Buffer.from(binaryBytes.subarray(0, 3)).toString('base64'))
		expect(binary.sha256).toBe(createHash('sha256').update(binaryBytes).digest('hex'))
		expect(binary.text).toBeUndefined()

		const metadataOnly = wb.rawPackagePart({
			partPath: 'xl/media/image1.png',
			encoding: 'none',
			maxBytes: 3,
		})
		expect(metadataOnly.encoding).toBe('none')
		expect(metadataOnly.previewByteLength).toBe(0)
		expect(metadataOnly.truncated).toBe(false)
		expect(metadataOnly.text).toBeUndefined()
		expect(metadataOnly.base64).toBeUndefined()
		expect(metadataOnly.sha256).toBe(binary.sha256)

		const missing = wb.rawPackagePart({ partPath: 'xl/missing.xml' })
		expect(missing.found).toBe(false)
		expect(missing.validPath).toBe(true)
		expect(missing.partPath).toBe('xl/missing.xml')

		const invalid = wb.rawPackagePart({ partPath: 'xl//workbook.xml' })
		expect(invalid.found).toBe(false)
		expect(invalid.validPath).toBe(false)
		expect(invalid.invalidReason).toContain('duplicate slashes')
		expect(wb.rawPackagePart({ partPath: 'xl/' }).invalidReason).toContain('empty segments')
		expect(wb.rawPackagePart({ partPath: 'xl\\workbook.xml' }).invalidReason).toContain(
			'forward slashes',
		)

		const wrongCase = wb.rawPackagePart({ partPath: 'XL/WORKBOOK.XML' })
		expect(wrongCase.found).toBe(false)
		expect(wrongCase.caseInsensitiveFallback).toBe(false)

		const fallback = wb.rawPackagePart({ partPath: 'XL/WORKBOOK.XML', caseInsensitive: true })
		expect(fallback.found).toBe(true)
		expect(fallback.partPath).toBe('xl/workbook.xml')
		expect(fallback.caseInsensitiveRequested).toBe(true)
		expect(fallback.caseInsensitiveFallback).toBe(true)

		const doc = await WorkbookDocument.open(sourceBytes, { mode: 'metadata-only' })
		const sessionRaw = await doc.rawPackagePart({ partPath: 'xl/workbook.xml', maxBytes: 16 })
		expect(sessionRaw.origin).toBe('source')
		expect(sessionRaw.load?.mode).toBe('metadata-only')
		expect(sessionRaw.load?.isPartial).toBe(true)
		expect(sessionRaw.semantics).toBe('raw-package-bytes')

		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'dirty' }] }])
		const dirtyRaw = wb.rawPackagePart({ partPath: 'xl/worksheets/sheet1.xml', maxBytes: 128 })
		expect(dirtyRaw.found).toBe(true)
		expect(dirtyRaw.origin).toBe('serialized-current')
		expect(dirtyRaw.load?.mode).toBe('full')
	})

	test('report returns compatibility info', () => {
		const wb = AscendWorkbook.create()
		expect(wb.report.status).toBe('clean')
		expect(wb.report.sourceFormat).toBe('ascend')
	})

	test('names returns defined name keys', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'MyRange', ref: 'Sheet1!A1:B10' }])
		expect(wb.names).toContain('MyRange')
	})

	test('definedName returns queryable name metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'MyRange', ref: 'Sheet1!A1:B10' }])
		expect(wb.definedName('MyRange')).toEqual({
			name: 'MyRange',
			formula: 'Sheet1!A1:B10',
			normalizedFormula: 'Sheet1!A1:B10',
			scope: 'workbook',
			references: [
				{ kind: 'range', text: 'Sheet1!A1:B10', scope: { kind: 'sheet', sheet: 'Sheet1' } },
			],
			refs: ['Sheet1!A1:B10'],
			functions: [],
			volatile: false,
		})
	})

	test('definedNames exposes parsed name inventory and workbook metadata inventories', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A3' }])
		const internal = wb as unknown as {
			wb: {
				workbookViews: Array<Record<string, unknown>>
				externalReferences: string[]
			}
		}
		internal.wb.workbookViews.push({ activeTab: 0, visibility: 'visible' })
		internal.wb.externalReferences.push('xl/externalLinks/externalLink1.xml')

		expect(wb.definedNames()).toEqual([
			{
				name: 'Budget',
				formula: 'Sheet1!A1:A3',
				normalizedFormula: 'Sheet1!A1:A3',
				scope: 'workbook',
				references: [
					{ kind: 'range', text: 'Sheet1!A1:A3', scope: { kind: 'sheet', sheet: 'Sheet1' } },
				],
				refs: ['Sheet1!A1:A3'],
				functions: [],
				volatile: false,
			},
		])
		expect(wb.workbookViews()).toEqual([{ activeTab: 0, visibility: 'visible' }])
		expect(wb.externalReferences()).toEqual(['xl/externalLinks/externalLink1.xml'])
	})

	test('formula returns parsed formula metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SUM(B1:B2)' }])
		const info = wb.formula('Sheet1!A1')
		expect(info).toBeDefined()
		expect(info?.normalizedFormula).toBe('SUM(B1:B2)')
		expect(info?.functions).toEqual(['SUM'])
		expect(info?.references).toEqual([{ kind: 'range', text: 'B1:B2', scope: { kind: 'local' } }])
		expect(info?.refs).toEqual(['B1:B2'])
	})

	test('formula metadata exposes whole-column references', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=SUM(A:A)' }])
		const info = wb.formula('Sheet1!B1')
		expect(info?.references).toEqual([
			{ kind: 'wholeColumn', text: 'A:A', scope: { kind: 'local' } },
		])
		expect(info?.refs).toContain('A:A')
	})

	test('formula metadata exposes structured reference column ranges', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A1',
				formula: '=SUM(Table1[[Revenue]:[Quantity]])',
			},
		])
		const info = wb.formula('Sheet1!A1')
		expect(info?.references).toEqual([
			{
				kind: 'structured',
				text: 'Table1[[Revenue]:[Quantity]]',
				scope: { kind: 'local' },
				table: 'Table1',
				specifiers: [],
				column: 'Revenue',
				endColumn: 'Quantity',
			},
		])
	})

	test('inspectSheet exposes table package relationship provenance', async () => {
		const wb = await AscendWorkbook.open(
			readFileSync(join(import.meta.dir, '../../../fixtures/xlsx/poi/StructuredReferences.xlsx')),
		)
		const detail = wb.inspectSheet('Table')
		expect(detail?.tables?.[0]).toMatchObject({
			name: '\\_Prime.1',
			partPath: 'xl/tables/table1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
			contentTypeSource: 'override',
			sourcePartPath: 'xl/worksheets/sheet2.xml',
			sourceRelationshipPart: 'xl/worksheets/_rels/sheet2.xml.rels',
			sourceRelationshipId: 'rId1',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			sourceRelationshipRawTarget: '../tables/table1.xml',
			sourceRelationshipResolvedTarget: 'xl/tables/table1.xml',
		})
	})

	test('inspectSheet exposes table queryTable package relationship provenance', async () => {
		const wb = await AscendWorkbook.open(
			makeSyntheticXlsx({
				'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
</Types>`,
				'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
				'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
				'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
				'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Name</v></c><c r="B1" t="str"><v>Value</v></c></row>
    <row r="2"><c r="A2" t="str"><v>A</v></c><c r="B2"><v>1</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>
</worksheet>`,
				'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
				'xl/tables/table1.xml': `<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="QueryTable" displayName="QueryTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0" tableType="queryTable">
  <autoFilter ref="A1:B2"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Name" queryTableFieldId="1"/>
    <tableColumn id="2" name="Value" queryTableFieldId="2"/>
  </tableColumns>
</table>`,
				'xl/tables/_rels/table1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdQuery99" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable" Target="../queryTables/queryTable1.xml"/>
</Relationships>`,
				'xl/queryTables/queryTable1.xml': `<?xml version="1.0"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="QueryTable" connectionId="1"/>`,
			}),
		)
		const queryTable = {
			relationshipId: 'rIdQuery99',
			partPath: 'xl/queryTables/queryTable1.xml',
			relationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			relationshipRawType: 'http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable',
			target: '../queryTables/queryTable1.xml',
		}

		expect(wb.inspectSheet('Data')?.tables?.[0]?.queryTable).toEqual(queryTable)
		expect(wb.table('QueryTable')?.queryTable).toEqual(queryTable)
	})

	test('formula metadata exposes workbook-qualified external references symbolically', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=[Book.xlsx]Sheet1!A1' }])
		const info = wb.formula('Sheet1!A1')
		expect(info?.references).toEqual([
			{
				kind: 'cell',
				text: '[Book.xlsx]Sheet1!A1',
				scope: { kind: 'external', workbook: 'Book.xlsx', sheet: 'Sheet1' },
			},
		])
	})

	test('formula metadata exposes 3D sheet-span references symbolically', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Sheet2' },
			{ op: 'addSheet', name: 'Sheet3' },
		])
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SUM(Sheet1:Sheet3!A1)' }])
		const info = wb.formula('Sheet1!A1')
		expect(info?.references).toEqual([
			{
				kind: 'cell',
				text: 'Sheet1:Sheet3!A1',
				scope: { kind: 'sheetSpan', startSheet: 'Sheet1', endSheet: 'Sheet3' },
			},
		])
	})

	test('formula and cell inspection expose formula binding metadata', async () => {
		const bytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>3</v></c>
      <c r="B1"><f t="shared" si="0">A1*2</f><v>6</v></c>
    </row>
    <row r="2">
      <c r="B2"><f t="shared" si="0"/><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})
		const wb = await AscendWorkbook.open(bytes)
		expect(wb.sheet('Calc')?.cell('A1')?.formulaBinding).toEqual({ kind: 'array', ref: 'A1:A2' })
		expect(wb.sheet('Calc')?.cell('B1')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'B1',
		})
		expect(wb.sheet('Calc')?.cell('B2')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'B1',
		})
		expect(wb.formula('Calc!A1')?.binding).toEqual({ kind: 'array', ref: 'A1:A2' })
		expect(wb.formula('Calc!B1')?.binding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'B1',
		})
		expect(wb.sheet('Calc')?.cell('B2')?.formula).toBe('A2*2')
		expect(wb.formula('Calc!B2')?.normalizedFormula).toBe('A2*2')
		const compact = wb.readWindowCompact('Calc', 'A1:B2', { includeRefs: true })
		expect(compact?.cells.map((cell) => [cell.ref, cell.formula, cell.formulaBinding])).toEqual([
			['A1', 'SUM(B1:B2)', { kind: 'array', ref: 'A1:A2' }],
			['B1', 'A1*2', { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' }],
			['B2', 'A2*2', { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' }],
		])
		const compactNoRefs = wb.readWindowCompact('Calc', 'A1:B2', { includeRefs: false })
		expect(compactNoRefs?.cells.map((cell) => cell.ref)).toEqual([undefined, undefined, undefined])
		expect(compactNoRefs?.cells.map((cell) => [cell.row, cell.col, cell.formula])).toEqual([
			[0, 0, 'SUM(B1:B2)'],
			[0, 1, 'A1*2'],
			[1, 1, 'A2*2'],
		])
	})

	test('getFormulaBinding returns summary for formula cells, null for non-formula', async () => {
		const bytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>A1*2</f><v>2</v></c>
      <c r="C1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>3</v></c>
      <c r="D1"><f t="shared" si="0">A1+1</f><v>2</v></c>
    </row>
    <row r="2">
      <c r="D2"><f t="shared" si="0"/><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})
		const wb = await AscendWorkbook.open(bytes)
		const sheet = wb.sheet('Calc')
		expect(sheet).toBeDefined()
		expect(sheet?.getFormulaBinding('A1')).toBeNull()
		expect(sheet?.getFormulaBinding('B1')).toEqual({ kind: 'normal', formula: 'A1*2' })
		expect(sheet?.getFormulaBinding('C1')).toEqual({
			kind: 'array',
			formula: 'SUM(B1:B2)',
			range: 'A1:A2',
		})
		expect(sheet?.getFormulaBinding('D1')).toMatchObject({
			kind: 'shared-anchor',
			formula: 'A1+1',
			sharedIndex: '0',
		})
		expect(sheet?.getFormulaBinding('D2')).toMatchObject({
			kind: 'shared-member',
			sharedIndex: '0',
			masterRef: 'D1',
		})
	})

	test('getFormulaCells returns all formula cells, optionally filtered by kind', async () => {
		const bytes = makeSyntheticXlsx({
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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>A1*2</f><v>2</v></c>
      <c r="C1"><f t="shared" si="0">A1+1</f><v>2</v></c>
    </row>
    <row r="2">
      <c r="C2"><f t="shared" si="0"/><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})
		const wb = await AscendWorkbook.open(bytes)
		const sheet = wb.sheet('Calc')
		expect(sheet).toBeDefined()
		const all = sheet?.getFormulaCells() ?? []
		expect(all.length).toBe(3)
		expect(all.map((e) => e.ref).sort()).toEqual(['B1', 'C1', 'C2'])
		const normal = sheet?.getFormulaCells({ kind: 'normal' }) ?? []
		expect(normal.length).toBe(1)
		expect(normal[0]).toEqual({ ref: 'B1', binding: { kind: 'normal', formula: 'A1*2' } })
		const anchors = sheet?.getFormulaCells({ kind: 'shared-anchor' }) ?? []
		expect(anchors.length).toBe(1)
		expect(anchors[0]?.ref).toBe('C1')
		expect(anchors[0]?.binding.kind).toBe('shared-anchor')
		const members = sheet?.getFormulaCells({ kind: 'shared-member' }) ?? []
		expect(members.length).toBe(1)
		expect(members[0]?.ref).toBe('C2')
		expect(members[0]?.binding.kind).toBe('shared-member')
	})

	test('setFormula and fillFormula helpers apply formula operations', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'A2', value: 3 },
				],
			},
		])
		wb.setFormula('Sheet1!B1', '=A1*2')
		wb.fillFormula('Sheet1!B1:B2', '=A1*2')
		wb.recalc()
		expect(wb.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1*2')
		expect(wb.sheet('Sheet1')?.cell('B2')?.formula).toBe('A2*2')
		expect(wb.sheet('Sheet1')?.cell('B2')?.value).toEqual({ kind: 'number', value: 6 })
	})

	test('add and delete sheets', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Extra' }])
		expect(wb.sheets).toEqual(['Sheet1', 'Extra'])

		wb.apply([{ op: 'deleteSheet', sheet: 'Extra' }])
		expect(wb.sheets).toEqual(['Sheet1'])
	})

	test('convenience methods addSheet, deleteSheet, insertRows, deleteRows', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'a' },
					{ ref: 'A2', value: 'b' },
					{ ref: 'A3', value: 'c' },
				],
			},
		])

		const addResult = wb.addSheet('Extra')
		expect(addResult.errors).toHaveLength(0)
		expect(wb.sheets).toEqual(['Sheet1', 'Extra'])

		const deleteResult = wb.deleteSheet('Extra')
		expect(deleteResult.errors).toHaveLength(0)
		expect(wb.sheets).toEqual(['Sheet1'])

		const insertResult = wb.insertRows('Sheet1', 1, 2)
		expect(insertResult.errors).toHaveLength(0)
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'a' })
		expect(wb.sheet('Sheet1')?.cell('A2')).toBeUndefined()
		expect(wb.sheet('Sheet1')?.cell('A3')).toBeUndefined()
		expect(wb.sheet('Sheet1')?.cell('A4')?.value).toEqual({ kind: 'string', value: 'b' })
		expect(wb.sheet('Sheet1')?.cell('A5')?.value).toEqual({ kind: 'string', value: 'c' })

		const deleteRowsResult = wb.deleteRows('Sheet1', 1, 2)
		expect(deleteRowsResult.errors).toHaveLength(0)
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'a' })
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'string', value: 'b' })
		expect(wb.sheet('Sheet1')?.cell('A3')?.value).toEqual({ kind: 'string', value: 'c' })
	})

	test('SheetHandle exposes comments, hyperlinks, merges, frozenPanes, protection, autoFilter', () => {
		const wb = AscendWorkbook.create()
		const handle = wb.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 to exist')

		expect(handle.comments().size).toBe(0)
		expect(handle.hyperlinks().size).toBe(0)
		expect(handle.merges).toEqual([])
		expect(handle.frozenRows).toBe(0)
		expect(handle.frozenCols).toBe(0)
		expect(handle.protection).toBeNull()
		expect(handle.autoFilter).toBeNull()
		expect(handle.conditionalFormats).toEqual([])
		expect(handle.dataValidations).toEqual([])
		expect(handle.imageRefs).toEqual([])
		expect(handle.state).toBe('visible')
		expect(handle.tabColor).toBeNull()
		expect(handle.sheetFormatPr).toBeNull()
	})

	test('SheetHandle comment and hyperlink single-cell accessors', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'test' }] },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://example.com' },
		])
		const handle = wb.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 to exist')
		expect(handle.hyperlink('A1')).toBeDefined()
		expect(handle.hyperlink('A1')?.target).toBe('https://example.com')
		expect(handle.hyperlink('B1')).toBeUndefined()
		expect(handle.comment('A1')).toBeUndefined()
	})

	test('getComments returns correct data when comments exist', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'cell' }] }])
		const internal = wb as unknown as {
			wb: { sheets: Array<{ comments: Map<string, { text: string; author?: string }> }> }
		}
		internal.wb.sheets[0]?.comments.set('A1', { text: 'Review this', author: 'Alice' })
		internal.wb.sheets[0]?.comments.set('B2', { text: 'Note' })

		const handle = wb.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 to exist')
		const comments = handle.getComments()
		expect(comments).toHaveLength(2)
		expect(comments).toContainEqual({ ref: 'A1', author: 'Alice', text: 'Review this' })
		expect(comments).toContainEqual({ ref: 'B2', text: 'Note' })
		expect(handle.getComments({ range: 'A1:A1' })).toEqual([
			{ ref: 'A1', author: 'Alice', text: 'Review this' },
		])
		expect(handle.getComments({ range: 'C1:D5' })).toEqual([])
	})

	test('getConditionalFormats returns rule summaries', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				sheets: Array<{
					conditionalFormats: Array<{
						sqref: string
						rules: Array<{ type: string; priority?: number }>
					}>
				}>
			}
		}
		internal.wb.sheets[0]?.conditionalFormats.push({
			sqref: 'A1:B10',
			rules: [
				{ type: 'cellIs', priority: 1 },
				{ type: 'colorScale', priority: 2 },
			],
		})

		const handle = wb.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 to exist')
		const cfs = handle.getConditionalFormats()
		expect(cfs).toHaveLength(2)
		expect(cfs[0]).toEqual({ type: 'cellIs', priority: 1, range: 'A1:B10' })
		expect(cfs[1]).toEqual({ type: 'colorScale', priority: 2, range: 'A1:B10' })
	})

	test('getHyperlinks returns link data', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'link' }] },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://example.com' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 'other' }] },
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'B2',
				url: 'mailto:test@example.com',
				display: 'Email',
			},
		])

		const handle = wb.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 to exist')
		const links = handle.getHyperlinks()
		expect(links).toHaveLength(2)
		expect(links).toContainEqual({ ref: 'A1', target: 'https://example.com' })
		expect(links).toContainEqual({ ref: 'B2', target: 'mailto:test@example.com', display: 'Email' })
		expect(handle.getHyperlinks({ range: 'B2:B2' })).toEqual([
			{ ref: 'B2', target: 'mailto:test@example.com', display: 'Email' },
		])
		expect(handle.getHyperlinks({ range: 'C1:D5' })).toEqual([])
	})

	test('TableHandle exposes ref, styleInfo, autoFilter, sortState, header/totals rows, and columnDefs', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 90 },
					{ ref: 'A3', value: 'Total' },
					{ ref: 'B3', value: 90 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'MyTable', hasHeaders: true },
		])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<{ tables: Array<Record<string, unknown>> }>
			}
		}
		const tableModel = internal.wb.sheets[0]?.tables[0] as
			| {
					sortState?: { ref: string; conditions: readonly { ref: string }[] }
					ref: { start: { row: number; col: number }; end: { row: number; col: number } }
					hasTotals: boolean
			  }
			| undefined
		if (tableModel) {
			tableModel.sortState = { ref: 'A1:B2', conditions: [{ ref: 'B2:B2' }] }
			tableModel.hasTotals = true
			tableModel.ref = { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }
		}

		const table = wb.table('MyTable')
		if (!table) throw new Error('Expected MyTable to exist')
		expect(table.name).toBe('MyTable')
		expect(table.ref).toBeDefined()
		expect(table.columns).toEqual(['Name', 'Score'])
		expect(table.hasHeaders).toBe(true)
		expect(table.hasTotals).toBe(true)
		expect(table.columnDefs).toHaveLength(2)
		expect(table.columnDefs[0]?.name).toBe('Name')
		expect(table.rowCount).toBe(1)
		expect(table.sortState?.ref).toBe('A1:B2')
		expect(table.headerRow()?.[0]).toEqual({ kind: 'string', value: 'Name' })
		expect(table.totalsRow()?.[1]).toEqual({ kind: 'number', value: 90 })
		expect(table.readRows({ offset: 0, limit: 1 })).toEqual({
			rowOffset: 0,
			rowLimit: 1,
			returnedRows: 1,
			totalRows: 1,
			hasMore: false,
			rows: [
				{
					index: 0,
					sheetRow: 1,
					values: {
						Name: { kind: 'string', value: 'Alice' },
						Score: { kind: 'number', value: 90 },
					},
				},
			],
		})
	})

	test('TableHandle columns getter returns correct column headers', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Product' },
					{ ref: 'B1', value: 'Revenue' },
					{ ref: 'C1', value: 'Quantity' },
					{ ref: 'A2', value: 'Widget' },
					{ ref: 'B2', value: 100 },
					{ ref: 'C2', value: 5 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C2', name: 'SalesTable', hasHeaders: true },
		])
		const table = wb.table('SalesTable')
		if (!table) throw new Error('Expected SalesTable to exist')
		expect(table.columns).toEqual(['Product', 'Revenue', 'Quantity'])
	})

	test('TableHandle exposes table metadata (style, totals) via getters', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Amount' },
					{ ref: 'A2', value: 'Item1' },
					{ ref: 'B2', value: 10 },
					{ ref: 'A3', value: 'Total' },
					{ ref: 'B3', value: 10 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'DataTable', hasHeaders: true },
		])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<{
					tables: Array<{
						ref: { start: { row: number; col: number }; end: { row: number; col: number } }
						hasTotals?: boolean
						tableStyleInfo?: { name?: string; showRowStripes?: boolean }
						autoFilter?: { ref: string; columns: unknown[] }
					}>
				}>
			}
		}
		const tableModel = internal.wb.sheets[0]?.tables[0]
		if (tableModel) {
			tableModel.hasTotals = true
			tableModel.ref = { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }
			tableModel.tableStyleInfo = { name: 'TableStyleMedium2', showRowStripes: true }
			tableModel.autoFilter = { ref: 'A1:B3', columns: [] }
		}

		const table = wb.table('DataTable')
		if (!table) throw new Error('Expected DataTable to exist')
		expect(table.styleInfo).toEqual({ name: 'TableStyleMedium2', showRowStripes: true })
		expect(table.totalsRow()).toEqual([
			{ kind: 'string', value: 'Total' },
			{ kind: 'number', value: 10 },
		])
		expect(table.hasFilters()).toBe(true)
		expect(table.sortState).toBeNull()
	})

	test('table handles stay current across workbook replacement', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 90 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'MyTable', hasHeaders: true },
		])
		const table = wb.table('MyTable')
		if (!table) throw new Error('Expected MyTable to exist')

		expect(table.readRows({ limit: 1 }).rows[0]?.values.Score).toEqual({
			kind: 'number',
			value: 90,
		})

		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 95 }] }])

		expect(table.readRows({ limit: 1 }).rows[0]?.values.Score).toEqual({
			kind: 'number',
			value: 95,
		})
	})

	test('inspectSheet exposes structured table metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 90 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'MyTable', hasHeaders: true },
		])
		const detail = wb.inspectSheet('Sheet1')
		expect(detail?.tables).toHaveLength(1)
		expect(detail?.tables?.[0]?.name).toBe('MyTable')
		expect(detail?.tables?.[0]?.rowCount).toBe(1)
		expect(detail?.tables?.[0]?.columnDefs.map((column) => column.name)).toEqual(['Name', 'Score'])
		expect(detail?.tables?.[0]?.headerRow?.[0]).toEqual({ kind: 'string', value: 'Name' })
	})

	test('trace exposes node values and depth', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'B1', value: 3 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=A1+B1' },
		])
		wb.recalc()

		const trace = wb.trace('Sheet1!C1')
		expect(trace).toBeDefined()
		expect(trace?.value).toEqual({ kind: 'number', value: 5 })
		expect(trace?.precedents).toEqual([
			{ ref: 'Sheet1!A1', formula: null, value: { kind: 'number', value: 2 }, depth: 1 },
			{ ref: 'Sheet1!B1', formula: null, value: { kind: 'number', value: 3 }, depth: 1 },
		])
	})

	test('trace keeps nested range precedents symbolic', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'A2', value: 3 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=SUM(A1:A2)' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=B1+1' },
		])
		wb.recalc()

		const trace = wb.trace('Sheet1!C1')
		expect(trace?.precedents).toEqual([
			{ ref: 'Sheet1!B1', formula: 'SUM(A1:A2)', value: { kind: 'number', value: 5 }, depth: 1 },
			{ ref: 'Sheet1!A1:A2', formula: null, value: { kind: 'number', value: 2 }, depth: 2 },
		])
	})

	test('WorkbookDocument reuses cached documents for unchanged files and invalidates on change', async () => {
		WorkbookDocument.clearCache()
		const path = join(
			tmpdir(),
			`ascend-session-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		try {
			const wb = AscendWorkbook.create()
			wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'v1' }] }])
			await wb.save(path)

			const first = await WorkbookDocument.open(path, { mode: 'values' })
			const second = await WorkbookDocument.open(path, { mode: 'values' })
			expect(first).toBe(second)
			expect(first.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'v1' })

			await Bun.sleep(20)
			const updated = AscendWorkbook.create()
			updated.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'v2' }] }])
			await updated.save(path)

			const third = await WorkbookDocument.open(path, { mode: 'values' })
			expect(third).not.toBe(first)
			expect(third.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'v2' })
		} finally {
			WorkbookDocument.clearCache()
			await unlink(path).catch(() => {})
		}
	})

	test('WorkbookDocument reuses cached documents for identical byte sources', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'bytes' }] }])
		const bytes = wb.toBytes()
		const first = await WorkbookDocument.open(bytes, { mode: 'values' })
		const second = await WorkbookDocument.open(bytes, { mode: 'values' })
		expect(first).toBe(second)
		expect(first.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'bytes' })
		WorkbookDocument.clearCache()
	})

	test('WorkbookDocument derives upgraded load snapshots immutably', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: '=A1*3' },
		])
		const bytes = wb.toBytes()
		const document = await WorkbookDocument.open(bytes, { mode: 'values' })
		expect(document.loadOptions.mode).toBe('values')
		const upgraded = await document.withLoad({ mode: 'formula' })
		expect(document.loadOptions.mode).toBe('values')
		expect(upgraded.loadOptions.mode).toBe('formula')
		expect(upgraded.formula('Sheet1!A2')?.normalizedFormula).toBe('A1*3')
		const reopened = await WorkbookDocument.open(bytes, { mode: 'formula' })
		expect(reopened).toBe(upgraded)
		WorkbookDocument.clearCache()
	})

	test('WorkbookDocument sheet handles stay snapshot-bound across withLoad', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'loaded' }] }])
		const bytes = wb.toBytes()
		const document = await WorkbookDocument.open(bytes, { mode: 'metadata-only' })
		const handle = document.sheet('Sheet1')
		if (!handle) throw new Error('Expected Sheet1 handle')

		expect(handle.cell('A1')).toBeUndefined()

		const hydrated = await document.withLoad({ mode: 'values' })

		expect(handle.cell('A1')).toBeUndefined()
		expect(hydrated.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'loaded' })
		WorkbookDocument.clearCache()
	})

	test('WorkbookDocument can derive broader sheet snapshots on demand', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Data' }])
		wb.apply([{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'loaded' }] }])
		const bytes = wb.toBytes()
		const document = await WorkbookDocument.open(bytes, { mode: 'values', sheets: ['Sheet1'] })
		expect(document.sheets).toEqual(['Sheet1'])
		const expanded = await document.withSheet('Data', { mode: 'values' })
		expect(document.sheets).toEqual(['Sheet1'])
		expect(expanded.sheets).toEqual(['Sheet1', 'Data'])
		expect(expanded.sheet('Data')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'loaded' })
		WorkbookDocument.clearCache()
	})
})

describe('WorkbookSession', () => {
	test('session open + inspect + read reuses same workbook', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'x' },
					{ ref: 'B1', value: 'y' },
				],
			},
		])
		const bytes = wb.toBytes()

		const session = await WorkbookSession.open(bytes, { mode: 'values' })
		const info1 = session.inspect()
		const read1 = session.read('Sheet1!A1:B1')
		const info2 = session.inspect()

		expect(info1).toBe(info2)
		expect(read1?.cells.length).toBe(2)
		expect(read1?.cells[0]?.value).toEqual({ kind: 'string', value: 'x' })
		session.close()
		WorkbookDocument.clearCache()
	})

	test('session is not stale immediately after open', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		const bytes = wb.toBytes()

		const session = await WorkbookSession.open(bytes)
		expect(session.isStale()).toBe(false)
		session.close()
		WorkbookDocument.clearCache()
	})

	test('multiple reads through session return consistent results', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 42 },
					{ ref: 'A2', value: 99 },
				],
			},
		])
		const bytes = wb.toBytes()

		const session = await WorkbookSession.open(bytes)
		const read1 = session.read('Sheet1!A1:A2')
		const read2 = session.read('Sheet1!A1:A2')

		expect(read1?.cells.length).toBe(2)
		expect(read2?.cells.length).toBe(2)
		expect(read1?.cells[0]?.value).toEqual(read2?.cells[0]?.value)
		expect(read1?.cells[1]?.value).toEqual(read2?.cells[1]?.value)
		expect(read1?.cells[0]?.value).toEqual({ kind: 'number', value: 42 })
		session.close()
		WorkbookDocument.clearCache()
	})

	test('session trace reuses analysis cache', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
				],
			},
		])
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=A1+B1' }])
		const bytes = wb.toBytes()

		const session = await WorkbookSession.open(bytes, { mode: 'formula' })
		const trace1 = session.trace('Sheet1!C1')
		const trace2 = session.trace('Sheet1!C1')

		expect(trace1?.dependsOn).toContain('Sheet1!A1')
		expect(trace1?.dependsOn).toContain('Sheet1!B1')
		expect(trace2?.dependsOn).toEqual(trace1?.dependsOn)
		session.close()
		WorkbookDocument.clearCache()
	})

	test('session read with sheet-qualified range', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Data' }])
		wb.apply([{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'from Data' }] }])
		const bytes = wb.toBytes()

		const session = await WorkbookSession.open(bytes)
		const read = session.read('Data!A1:A1')
		expect(read?.cells.length).toBe(1)
		expect(read?.cells[0]?.value).toEqual({ kind: 'string', value: 'from Data' })
		session.close()
		WorkbookDocument.clearCache()
	})

	test('session workbook() returns underlying document', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'via workbook' }] }])
		const bytes = wb.toBytes()

		const session = await WorkbookSession.open(bytes)
		const doc = session.workbook()
		expect(doc.inspect().sheetCount).toBe(1)
		expect(doc.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'via workbook',
		})
		session.close()
		WorkbookDocument.clearCache()
	})

	test('session close prevents further operations', async () => {
		WorkbookDocument.clearCache()
		const wb = AscendWorkbook.create()
		const bytes = wb.toBytes()
		const session = await WorkbookSession.open(bytes)
		session.close()
		expect(() => session.inspect()).toThrow('WorkbookSession is closed')
		expect(() => session.workbook()).toThrow('WorkbookSession is closed')
		WorkbookDocument.clearCache()
	})
})

describe('ops.listOperations', () => {
	test('returns all operation schemas with op, description, and requiredFields', () => {
		const schemas = ops.listOperations()
		expect(schemas.length).toBeGreaterThan(10)
		const setCells = schemas.find((s) => s.op === 'setCells')
		expect(setCells).toBeDefined()
		expect(setCells?.description).toBe('Set cell values')
		expect(setCells?.requiredFields).toEqual(['sheet', 'updates'])
		const addSheet = schemas.find((s) => s.op === 'addSheet')
		expect(addSheet?.optionalFields).toContain('position')
	})

	test('operation schemas include examples and shared parser rejects unknown ops', () => {
		const schema = ops.getOperationsSchema().find((entry) => entry.op === 'setCells')
		expect(schema?.examples[0]?.op).toBe('setCells')
		const parsed = parseOperations([{ op: 'missingOperation' }])
		expect(parsed.ok).toBe(false)
		if (!parsed.ok) expect(parsed.issues[0]).toContain('not supported')
	})
})

describe('capabilities registry', () => {
	test('classifies the major Excel feature families', () => {
		const capabilities = listCapabilities()
		for (const family of [
			'workbook/package',
			'sheets/ranges',
			'tables/data',
			'formula engine',
			'visuals',
			'analytics',
			'active content',
			'connections',
			'agent UX',
		]) {
			expect(capabilities.some((capability) => capability.family === family)).toBe(true)
		}
		expect(capabilities.find((capability) => capability.id === 'analytics.pivots')?.priority).toBe(
			'P0',
		)
		expect(capabilities.find((capability) => capability.id === 'visuals.charts')?.status).toBe(
			'editable',
		)
		const iterativeCalc = capabilities.find(
			(capability) => capability.id === 'formulas.iterative-calc',
		)
		expect(iterativeCalc?.status).toBe('editable')
		expect(iterativeCalc?.tests).toContain('packages/engine/src/calc.test.ts')
		const spillDiagnostics = capabilities.find(
			(capability) => capability.id === 'formulas.spill-diagnostics',
		)
		expect(spillDiagnostics?.status).toBe('editable')
		expect(spillDiagnostics?.tests).toContain('packages/verify/src/verify.test.ts')
		const externalRefs = capabilities.find(
			(capability) => capability.id === 'formulas.external-refs',
		)
		expect(externalRefs?.status).toBe('editable')
		expect(externalRefs?.tests).toContain('packages/sdk/src/external-reference-usages.test.ts')
		const refreshMetadata = capabilities.find(
			(capability) => capability.id === 'connections.refresh-metadata',
		)
		expect(refreshMetadata?.status).toBe('editable')
		expect(refreshMetadata?.tests).toContain('packages/sdk/src/connection-inventory.test.ts')
		const queryTables = capabilities.find(
			(capability) => capability.id === 'connections.query-tables',
		)
		expect(queryTables?.status).toBe('editable')
		expect(queryTables?.tests).toContain('packages/sdk/src/connection-inventory.test.ts')
		const textBoxes = capabilities.find((capability) => capability.id === 'visuals.text-boxes')
		expect(textBoxes?.status).toBe('editable')
		expect(textBoxes?.tests).toContain('packages/sdk/src/visual-inventory.test.ts')
		const sparklines = capabilities.find((capability) => capability.id === 'visuals.sparklines')
		expect(sparklines?.status).toBe('editable')
		expect(sparklines?.tests).toContain('packages/sdk/src/advanced-filter-sparkline.test.ts')
		const advancedFilters = capabilities.find(
			(capability) => capability.id === 'tables.advanced-filters',
		)
		expect(advancedFilters?.status).toBe('editable')
		expect(advancedFilters?.tests).toContain('packages/sdk/src/advanced-filter-sparkline.test.ts')
		const themes = capabilities.find((capability) => capability.id === 'workbook.themes')
		expect(themes?.status).toBe('editable')
		expect(themes?.tests).toContain('packages/sdk/src/theme-inventory.test.ts')
		const workbookProperties = capabilities.find(
			(capability) => capability.id === 'workbook.properties',
		)
		expect(workbookProperties?.tests).toContain('packages/sdk/src/sdk.test.ts')
		expect(workbookProperties?.gapReason).toContain('setDocumentProperties')
	})
})

function makeSyntheticXlsx(parts: Record<string, string | Uint8Array>): Uint8Array {
	const entries = new Map<string, Uint8Array>()
	for (const [path, content] of Object.entries(parts)) {
		entries.set(path, typeof content === 'string' ? encode(content) : content)
	}
	return createZip(entries)
}
