import { describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { extractZip } from '../../io-xlsx/src/reader/zip.ts'
import { createZip, encode } from '../../io-xlsx/src/writer/zip.ts'
import { Ascend, AscendWorkbook, WorkbookSession } from './index.ts'

describe('AscendWorkbook', () => {
	test('Ascend is an alias for AscendWorkbook', () => {
		expect(Ascend).toBe(AscendWorkbook)
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
					imageRefs: Array<Record<string, unknown>>
			  }
			| undefined
		backingSheet?.comments.set('A1', { text: 'note' })
		backingSheet?.dataValidations.push({ sqref: 'B1', type: 'list', formula1: '"A,B"' })
		backingSheet?.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'cellIs', formulas: ['1'] }],
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
		})
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
		})
		internal.wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Product',
			pivotTableNames: ['PivotTable1'],
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
		expect(info.imageCount).toBe(1)
		expect(info.pivotTableCount).toBe(1)
		expect(info.pivotCacheCount).toBe(1)
		expect(info.slicerCount).toBe(1)
		expect(info.slicerCacheCount).toBe(1)
		expect(info.hasWorkbookProtection).toBe(true)
		expect(info.pivotTables[0]?.name).toBe('PivotTable1')
		expect(info.pivotCaches[0]?.sourceSheet).toBe('Raw')
		expect(info.slicerCaches[0]?.name).toBe('Slicer_Product')
		expect(info.slicers[0]?.name).toBe('Product')
		expect(info.sheets[0]?.cellCount).toBe(1)
		expect(info.sheets[0]?.commentCount).toBe(1)
		expect(info.sheets[0]?.conditionalFormatCount).toBe(1)
		expect(info.sheets[0]?.dataValidationCount).toBe(1)
		expect(info.sheets[0]?.imageCount).toBe(1)
		expect(info.sheets[0]?.hasProtection).toBe(false)
		expect(info.sheets[0]?.name).toBe('Sheet1')
		expect(info.sheets[0]?.cellDataLoaded).toBe(true)
	})

	test('WorkbookReadView and WorkbookSession expose pivot and slicer query surfaces', async () => {
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
		})
		internal.wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 7,
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
		})
		internal.wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Product',
			pivotTableNames: ['PivotTable1'],
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
		expect(wb.slicers()[0]?.cacheName).toBe('Slicer_Product')

		const corpusPath = join(
			import.meta.dir,
			'../../../research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx',
		)
		const session = await WorkbookSession.open(corpusPath)
		expect(session.pivotTables().length).toBeGreaterThan(0)
		expect(session.pivotCaches().length).toBeGreaterThan(0)
		WorkbookSession.clearCache()
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
		expect(detail?.drawingRefs?.hasLegacyDrawing).toBe(true)
		expect(detail?.pageMargins?.left).toBe(0.5)
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
		expect(archive.readText('xl/sharedStrings.xml')).toBe(
			extractZip(bytes).readText('xl/sharedStrings.xml'),
		)
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

	test('CSV import creates workbook', () => {
		const csv = 'Name,Age\nAlice,30\nBob,25'
		const wb = AscendWorkbook.fromCsv(csv)
		expect(wb.sheets).toEqual(['Sheet1'])

		const handle = wb.sheet('Sheet1')
		expect(handle).toBeDefined()
		expect(handle?.cell('A1')?.value).toEqual({ kind: 'string', value: 'Name' })
		expect(handle?.cell('B2')?.value).toEqual({ kind: 'number', value: 30 })
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
		expect(preview.writePlan?.totalParts).toBeGreaterThan(0)

		const cell = wb.sheet('Sheet1')?.cell('A1')
		expect(cell?.value).toEqual({ kind: 'string', value: 'before' })
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
		})
		expect(wb.sheet('Calc')?.cell('B2')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
		})
		expect(wb.formula('Calc!A1')?.binding).toEqual({ kind: 'array', ref: 'A1:A2' })
		expect(wb.formula('Calc!B1')?.binding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
		})
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

	test('WorkbookSession reuses cached sessions for unchanged files and invalidates on change', async () => {
		WorkbookSession.clearCache()
		const path = join(
			tmpdir(),
			`ascend-session-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		try {
			const wb = AscendWorkbook.create()
			wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'v1' }] }])
			await wb.save(path)

			const first = await WorkbookSession.open(path, { mode: 'values' })
			const second = await WorkbookSession.open(path, { mode: 'values' })
			expect(first).toBe(second)
			expect(first.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'v1' })

			await Bun.sleep(20)
			const updated = AscendWorkbook.create()
			updated.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'v2' }] }])
			await updated.save(path)

			const third = await WorkbookSession.open(path, { mode: 'values' })
			expect(third).not.toBe(first)
			expect(third.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'v2' })
		} finally {
			WorkbookSession.clearCache()
			await unlink(path).catch(() => {})
		}
	})

	test('WorkbookSession reuses cached sessions for identical byte sources', async () => {
		WorkbookSession.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'bytes' }] }])
		const bytes = wb.toBytes()
		const first = await WorkbookSession.open(bytes, { mode: 'values' })
		const second = await WorkbookSession.open(bytes, { mode: 'values' })
		expect(first).toBe(second)
		expect(first.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'bytes' })
		WorkbookSession.clearCache()
	})

	test('WorkbookSession can upgrade load options in place', async () => {
		WorkbookSession.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: '=A1*3' },
		])
		const bytes = wb.toBytes()
		const session = await WorkbookSession.open(bytes, { mode: 'values' })
		expect(session.openOptions.mode).toBe('values')
		await session.upgrade({ mode: 'formula' })
		expect(session.openOptions.mode).toBe('formula')
		expect(session.formula('Sheet1!A2')?.normalizedFormula).toBe('A1*3')
		WorkbookSession.clearCache()
	})

	test('WorkbookSession can hydrate additional sheets on demand', async () => {
		WorkbookSession.clearCache()
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Data' }])
		wb.apply([{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'loaded' }] }])
		const bytes = wb.toBytes()
		const session = await WorkbookSession.open(bytes, { mode: 'values', sheets: ['Sheet1'] })
		expect(session.sheets).toEqual(['Sheet1'])
		await session.hydrateSheet('Data', { mode: 'values' })
		expect(session.sheets).toEqual(['Sheet1', 'Data'])
		expect(session.sheet('Data')?.cell('A1')?.value).toEqual({ kind: 'string', value: 'loaded' })
		WorkbookSession.clearCache()
	})
})

function makeSyntheticXlsx(parts: Record<string, string>): Uint8Array {
	const entries = new Map<string, Uint8Array>()
	for (const [path, content] of Object.entries(parts)) {
		entries.set(path, encode(content))
	}
	return createZip(entries)
}
