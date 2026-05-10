import { describe, expect, test } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import { runFormulaCorpusCorrectness } from '../benchmarks/formula-corpus-correctness.ts'
import {
	normalizeManifest,
	selectManifestEntries,
	validateManifestProvenance,
} from '../corpus/manifest.ts'
import { loadManifest } from './calamine/manifest.ts'

const calamineDir = fileURLToPath(new URL('./calamine/', import.meta.url))
const calamineManifest = fileURLToPath(new URL('./calamine/manifest.ts', import.meta.url))

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./calamine/${name}`, import.meta.url))
}

function fixtureReadOptions(name: string): { readonly password?: string } {
	return name === 'pass_protected.xlsx' ? { password: '123' } : {}
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('Calamine XLSX/XLSM fixture corpus', () => {
	test('manifest has pinned provenance for the vendored MIT fixture subset', async () => {
		expect(existsSync(new URL('./calamine/LICENSE', import.meta.url))).toBe(true)
		const entries = normalizeManifest(await loadManifest())
		expect(entries.length).toBe(53)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['calamine'] })).toHaveLength(entries.length)
		expect(selectManifestEntries(entries, { tags: ['formula-fidelity'] }).length).toBeGreaterThan(0)
		expect(selectManifestEntries(entries, { tags: ['issue-regression'] }).length).toBeGreaterThan(0)
		expect(selectManifestEntries(entries, { tags: ['table'] }).length).toBeGreaterThan(0)
		expect(selectManifestEntries(entries, { tags: ['macro'] }).length).toBe(2)
		expect(selectManifestEntries(entries, { tags: ['unsupported'] })).toHaveLength(0)
		expect(
			selectManifestEntries(entries, { tags: ['encrypted'] }).map((entry) => entry.file),
		).toEqual(['pass_protected.xlsx'])
	})

	test('opens and round-trips every Calamine OOXML fixture without crashing', async () => {
		const entries = await loadManifest()
		for (const entry of entries) {
			const initial = readXlsx(loadFixture(entry.file), fixtureReadOptions(entry.file))
			expectOk(initial)
			const written = writeXlsx(initial.value.workbook, initial.value.capsules)
			expectOk(written)
			const reopened = readXlsx(written.value)
			expectOk(reopened)
			expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
		}
	})

	test('reports password failures before opening encrypted OOXML fixtures', () => {
		const missing = readXlsx(loadFixture('pass_protected.xlsx'))
		expect(missing.ok).toBe(false)
		if (!missing.ok) expect(missing.error.code).toBe('PROTECTION_ERROR')
		const wrong = readXlsx(loadFixture('pass_protected.xlsx'), { password: 'wrong' })
		expect(wrong.ok).toBe(false)
		if (!wrong.ok) expect(wrong.error.message).toContain('Invalid XLSX password')
	})

	test('captures expected feature families from Calamine regression fixtures', async () => {
		const entries = normalizeManifest(await loadManifest())
		expect(entries.find((entry) => entry.file === 'date_1904.xlsx')?.featureTags).toContain('date')
		expect(entries.find((entry) => entry.file === 'pivots.xlsx')?.features.pivot_tables).toBe(true)
		expect(entries.find((entry) => entry.file === 'picture.xlsx')?.features.images_or_media).toBe(
			true,
		)
		expect(entries.find((entry) => entry.file === 'vba.xlsm')?.features.macros).toBe(true)
		expect(entries.find((entry) => entry.file === 'table-multiple.xlsx')?.features.tables).toBe(
			true,
		)
	})

	test('recovers sheet, shared-string, style, and chart links from issue252.xlsx', () => {
		const result = readXlsx(loadFixture('issue252.xlsx'))
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.name).toBe('Sheet1')
		expect(sheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'data' })
		expect(sheet?.cells.get(4, 4)?.formula).toBe('SUM(B1:D5)')
		expect(result.value.workbook.chartParts).toEqual([
			expect.objectContaining({
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				chartType: 'barChart',
				series: [expect.objectContaining({ valueRef: 'Sheet1!$B$8' })],
			}),
		])
	})

	test('inventories real Calamine image anchors from picture.xlsx', async () => {
		const wb = await AscendWorkbook.open(loadFixture('picture.xlsx'))
		const visuals = wb.visualInventory()
		expect(visuals.sheetImageCount).toBe(2)
		expect(
			visuals.sheets.map((sheet) => [
				sheet.sheet,
				sheet.imageRefs?.map((image) => ({
					drawingPartPath: image.drawingPartPath,
					relId: image.relId,
					targetPath: image.targetPath,
					anchorKind: image.anchor?.kind,
				})),
			]),
		).toEqual([
			[
				'Sheet1',
				[
					{
						drawingPartPath: 'xl/drawings/drawing1.xml',
						relId: 'rId1',
						targetPath: 'xl/media/image1.jpg',
						anchorKind: 'twoCell',
					},
				],
			],
			[
				'Sheet2',
				[
					{
						drawingPartPath: 'xl/drawings/drawing2.xml',
						relId: 'rId1',
						targetPath: 'xl/media/image2.png',
						anchorKind: 'twoCell',
					},
				],
			],
		])
	})

	test('resolves absolute table relationship targets from Calamine fixtures', () => {
		const result = readXlsx(loadFixture('table_with_absolute_paths.xlsx'))
		expectOk(result)

		const sheet = result.value.workbook.sheets.find((entry) => entry.name === 'table')
		expect(sheet?.tables).toHaveLength(1)
		expect(sheet?.tables[0]).toMatchObject({
			name: 'Table1',
			ref: { start: { row: 0, col: 0 }, end: { row: 4, col: 1 } },
			tableStyleInfo: { name: 'TableStyleLight1' },
			columns: [
				expect.objectContaining({ id: 1, name: 'Item' }),
				expect.objectContaining({ id: 2, name: 'Name' }),
			],
		})
	})

	test('captures Calamine insert-row tables and calculated columns', () => {
		const result = readXlsx(loadFixture('table_with_insertrow_attribute.xlsx'))
		expectOk(result)

		const sheet = result.value.workbook.sheets.find((entry) => entry.name === 'Sheet1')
		const tables = new Map((sheet?.tables ?? []).map((table) => [table.name, table]))
		expect([...tables.keys()].sort()).toEqual(['Table1', 'Table2', 'Table3', 'Table4'])
		expect(tables.get('Table1')?.columns[1]).toMatchObject({
			name: 'Price',
			formula: 'RAND()*100',
		})
		expect(tables.get('Table2')?.columns[1]).toMatchObject({
			name: 'Price',
			formula: 'RAND()*100',
		})
		expect(tables.get('Table3')).toMatchObject({
			name: 'Table3',
			insertRow: true,
			ref: { start: { row: 7, col: 12 }, end: { row: 8, col: 12 } },
		})
		expect(tables.get('Table4')).toMatchObject({
			name: 'Table4',
			insertRow: true,
			ref: { start: { row: 7, col: 14 }, end: { row: 8, col: 14 } },
		})
	})

	test('surfaces real pivot cache shared item bounds and grouped fields', () => {
		const result = readXlsx(loadFixture('pivots.xlsx'))
		expectOk(result)
		const cache = result.value.workbook.pivotCaches.find((entry) => entry.cacheId === 65)
		expect(cache).toBeDefined()
		expect(cache).toMatchObject({
			sourceType: 'worksheet',
			sourceSheet: 'DataSheet',
			sourceRef: 'A1:J11',
			records: {
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 10,
				parsedCount: 10,
			},
		})
		expect(cache?.records?.preview[0]?.values).toEqual([
			{ index: 0, kind: 'number', value: '1' },
			{ index: 1, kind: 'sharedItem', sharedItemIndex: 0 },
			{ index: 2, kind: 'sharedItem', sharedItemIndex: 0 },
			{ index: 3, kind: 'sharedItem', sharedItemIndex: 0 },
			{ index: 4, kind: 'sharedItem', sharedItemIndex: 0 },
			{ index: 5, kind: 'sharedItem', sharedItemIndex: 0 },
			{ index: 6, kind: 'number', value: '5.6179775280898872' },
			{ index: 7, kind: 'boolean', value: '1' },
			{ index: 8, kind: 'missing' },
			{ index: 9, kind: 'missing' },
		])
		expect(cache?.records?.valueKindCounts).toContainEqual({ kind: 'error', count: 2 })
		const fields = cache?.fields ?? []

		expect(fields.find((field) => field.name === 'Id')?.sharedItemsInfo).toMatchObject({
			containsNumber: true,
			containsInteger: true,
			containsString: false,
			minValue: 1,
			maxValue: 10,
		})
		expect(fields.find((field) => field.name === 'Date')?.sharedItemsInfo).toMatchObject({
			containsDate: true,
			containsNonDate: false,
			minDate: '1999-01-01T00:00:00',
			maxDate: '2024-12-02T00:00:00',
			count: 9,
		})
		expect(fields.find((field) => field.name === 'Value')?.fieldGroup).toEqual({ parent: 11 })
		expect(fields.find((field) => field.name === 'Value2')?.fieldGroup).toMatchObject({
			base: 3,
			discreteItems: [
				{ index: 0, value: 1 },
				{ index: 1, value: 1 },
				{ index: 2, value: 1 },
				{ index: 3, value: 0 },
			],
			groupItems: [
				{ index: 0, kind: 'number', value: '5' },
				{ index: 1, kind: 'string', value: 'Group1' },
			],
		})
	})

	test('cached formulas in the Calamine subset have no semantic mismatches', async () => {
		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: calamineDir,
			manifest: calamineManifest,
			tags: ['formula-fidelity'],
			tiers: [],
			maxReportedMismatches: 20,
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
			maxUnacceptedMismatches: 0,
			maxSemanticMismatches: 0,
			maxErrors: 0,
			minComparedFormulas: 74,
			minSemanticPerfectWorkbooks: 15,
		})
		expect(payload.summary).toMatchObject({
			workbookCount: 15,
			formulaCount: 74,
			comparedCount: 74,
			volatileOracleSkipCount: 8,
			mismatchCount: 0,
			acceptedMismatchCount: 0,
			unacceptedMismatchCount: 0,
			semanticMismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 15,
			semanticPerfectWorkbookCount: 15,
		})
	})
})
