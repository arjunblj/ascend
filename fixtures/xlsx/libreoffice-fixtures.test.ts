import { describe, expect, test } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import { runFormulaCorpusCorrectness } from '../benchmarks/formula-corpus-correctness.ts'
import {
	normalizeManifest,
	selectManifestEntries,
	validateManifestProvenance,
} from '../corpus/manifest.ts'
import { loadManifest } from './libreoffice/manifest.ts'

const libreOfficeDir = fileURLToPath(new URL('./libreoffice/', import.meta.url))
const libreOfficeManifest = fileURLToPath(new URL('./libreoffice/manifest.ts', import.meta.url))

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./libreoffice/${name}`, import.meta.url))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('LibreOffice XLSX fixture corpus', () => {
	test('manifest has pinned provenance for the vendored Calc QA fixture subset', async () => {
		expect(existsSync(new URL('./libreoffice/LICENSE', import.meta.url))).toBe(true)
		const entries = normalizeManifest(await loadManifest())
		expect(entries.length).toBe(36)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['libreoffice'] })).toHaveLength(entries.length)
		expect(selectManifestEntries(entries, { tags: ['formula-fidelity'] })).toHaveLength(15)
		expect(selectManifestEntries(entries, { tags: ['pivot-table'] })).toHaveLength(5)
		expect(selectManifestEntries(entries, { tags: ['table'] })).toHaveLength(6)
		expect(selectManifestEntries(entries, { tags: ['conditional-formatting'] })).toHaveLength(4)
		expect(selectManifestEntries(entries, { tags: ['date'] })).toHaveLength(2)
		expect(
			selectManifestEntries(entries, { tags: ['active-content'] }).map((entry) => entry.file),
		).toEqual(['activex_checkbox.xlsx'])
		expect(
			selectManifestEntries(entries, { tags: ['strict-ooxml'] }).map((entry) => entry.file),
		).toEqual(['universal-content-strict.xlsx'])
	})

	test('opens and round-trips every LibreOffice Calc QA fixture without crashing', async () => {
		const entries = await loadManifest()
		for (const entry of entries) {
			const initial = readXlsx(loadFixture(entry.file))
			expectOk(initial)
			expect(
				initial.value.workbook.sheets.length + initial.value.workbook.chartSheets.length,
			).toBeGreaterThan(0)
			const written = writeXlsx(initial.value.workbook, initial.value.capsules)
			expectOk(written)
			const reopened = readXlsx(written.value)
			expectOk(reopened)
			expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
		}
	})

	test('captures expected feature families from LibreOffice regression fixtures', async () => {
		const entries = normalizeManifest(await loadManifest())
		expect(
			entries.find((entry) => entry.file === 'PivotTable_CachedDefinitionAndDataInSync.xlsx')
				?.features.pivot_tables,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'activex_checkbox.xlsx')?.features.active_content,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'MissingPathExternal.xlsx')?.features.external_links,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'textLengthDataValidity.xlsx')?.features
				.data_validations,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'universal-content-strict.xlsx')?.features
				.strict_ooxml,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'colorscale.xlsx')?.features.conditional_formatting,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'functions-excel-2010.xlsx')?.counts.formulas,
		).toBe(517)
		expect(
			entries.find((entry) => entry.file === 'tdf167689_tableType.xlsx')?.features.tables,
		).toBe(true)
	})

	test('captures LibreOffice sparkline styling and multi-sparkline groups', () => {
		const initial = readXlsx(loadFixture('Sparklines.xlsx'))
		expectOk(initial)

		const sheet1 = initial.value.workbook.sheets.find((sheet) => sheet.name === 'Sheet1')
		expect(sheet1?.sparklineGroups).toHaveLength(2)
		expect(sheet1?.sparklineGroups[0]).toMatchObject({
			groupIndex: 0,
			count: 1,
			lineWeight: 1,
			displayEmptyCellsAs: 'gap',
			markers: true,
			highPoint: true,
			lowPoint: true,
			firstPoint: true,
			lastPoint: true,
			negative: true,
			displayXAxis: true,
			colorSeries: 'FF376092',
			colorNegative: 'FF00B050',
			colorHigh: 'FF92D050',
			range: 'Sheet1!B1:M1',
			locationRange: 'A2',
		})
		expect(sheet1?.sparklineGroups[1]).toMatchObject({
			type: 'column',
			highPoint: true,
			lowPoint: true,
			firstPoint: true,
			lastPoint: true,
			negative: true,
			range: 'Sheet1!B1:M1',
			locationRange: 'A3',
		})

		const sheet3 = initial.value.workbook.sheets.find((sheet) => sheet.name === 'Sheet3')
		expect(sheet3?.sparklineGroups.map((group) => [group.type, group.count])).toEqual([
			['stacked', 10],
			['column', 10],
			[undefined, 10],
		])
		expect(sheet3?.sparklineGroups[0]?.sparklines?.[0]).toEqual({
			range: 'Sheet3!A1:J1',
			locationRange: 'N1',
		})
		expect(sheet3?.sparklineGroups[0]?.sparklines?.[9]).toEqual({
			range: 'Sheet3!A10:J10',
			locationRange: 'N10',
		})
		expect(sheet3?.sparklineGroups[1]?.sparklines?.[9]).toEqual({
			range: 'Sheet3!A10:J10',
			locationRange: 'M10',
		})
		expect(sheet3?.sparklineGroups[2]?.sparklines?.[9]).toEqual({
			range: 'Sheet3!A10:J10',
			locationRange: 'L10',
		})
	})

	test('inventories real LibreOffice textbox drawing text and relationship ids', () => {
		const initial = readXlsx(loadFixture('textbox-hyperlink.xlsx'))
		expectOk(initial)

		const sheet = initial.value.workbook.sheets.find((entry) => entry.name === 'Sheet1')
		expect(sheet?.drawingObjectRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 2,
			name: 'TextBox 1',
			text: 'text',
			relIds: ['rId1'],
			relationshipRefs: [
				{
					id: 'rId1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
					target: 'https://www.google.com/',
					targetMode: 'External',
				},
			],
			anchor: {
				kind: 'twoCell',
				from: { col: 2, row: 3, colOff: 133350, rowOff: 152400 },
				to: { col: 10, row: 9, colOff: 28575, rowOff: 85725 },
			},
		})
		expect(sheet?.drawingObjectRefs).toHaveLength(1)
	})

	test('links LibreOffice ActiveX checkbox controls to worksheet and drawing metadata', () => {
		const initial = readXlsx(loadFixture('activex_checkbox.xlsx'))
		expectOk(initial)

		const activeX = initial.value.workbook.activeContent.find(
			(content) => content.kind === 'activeX' && content.partPath.endsWith('.xml'),
		)
		expect(activeX).toMatchObject({
			sheetName: 'Sheet1',
			sourceRelationshipId: 'rId3',
			worksheetControl: {
				shapeId: 1025,
				name: 'CheckBox1343',
				relationshipId: 'rId3',
				controlPrRelationshipId: 'rId4',
				controlPrTarget: 'xl/media/image1.emf',
				anchor: {
					kind: 'twoCell',
					from: { col: 1, row: 3, colOff: 438150, rowOff: 38100 },
					to: { col: 4, row: 6, colOff: 161925, rowOff: 114300 },
				},
				vmlMapOcx: true,
				vmlImageTarget: 'xl/media/image1.emf',
			},
		})

		const sheet = initial.value.workbook.sheets.find((entry) => entry.name === 'Sheet1')
		expect(sheet?.drawingObjectRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'shape',
			id: 1025,
			name: 'CheckBox1343',
			anchor: activeX?.worksheetControl?.anchor,
		})
	})

	test('surfaces and round-trips LibreOffice query table relationships', () => {
		const initial = readXlsx(loadFixture('TableEmptyHeaders.xlsx'))
		expectOk(initial)

		const sheet = initial.value.workbook.sheets.find((entry) => entry.name === 'BTC')
		const table = sheet?.tables.find((entry) => entry.name === 'Bitcoin')
		expect(table).toMatchObject({
			name: 'Bitcoin',
			tableType: 'queryTable',
			ref: { start: { row: 0, col: 0 }, end: { row: 15, col: 1 } },
			autoFilter: { ref: 'A1:B16' },
			tableStyleInfo: { name: 'TableStyleMedium7' },
			queryTable: {
				relationshipId: 'rId1',
				partPath: 'xl/queryTables/queryTable1.xml',
				relationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
				target: '../queryTables/queryTable1.xml',
			},
		})
		expect(table?.columns).toEqual([
			expect.objectContaining({
				id: 3,
				uniqueName: '3',
				name: 'Column1',
				queryTableFieldId: 1,
				dataDxfId: 1,
			}),
			expect.objectContaining({
				id: 2,
				uniqueName: '2',
				name: 'Column2',
				queryTableFieldId: 2,
				dataDxfId: 0,
			}),
		])

		const written = writeXlsx(initial.value.workbook, initial.value.capsules, {
			dirtySheetNames: ['BTC'],
		})
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedTable = reopened.value.workbook.sheets
			.find((entry) => entry.name === 'BTC')
			?.tables.find((entry) => entry.name === 'Bitcoin')
		expect(reopenedTable?.queryTable).toMatchObject({
			relationshipId: 'rId1',
			partPath: 'xl/queryTables/queryTable1.xml',
		})
		expect(reopenedTable?.columns[0]).toMatchObject({
			uniqueName: '3',
			queryTableFieldId: 1,
		})
	})

	test('captures LibreOffice table total-row style metadata', () => {
		const initial = readXlsx(loadFixture('totalsRowFunction.xlsx'))
		expectOk(initial)

		const table = initial.value.workbook.sheets[0]?.tables.find(
			(entry) => entry.name === 'PresentPlanner',
		)
		expect(table).toMatchObject({
			hasTotals: true,
			ref: { start: { row: 1, col: 1 }, end: { row: 6, col: 6 } },
			tableStyleInfo: { name: 'Present planner table' },
		})
		expect(table?.columns.map((column) => column.dataCellStyle)).toEqual([
			'Date',
			'Normal',
			'Normal',
			'Normal',
			'Amount',
			'Notes',
		])
		expect(table?.columns[4]).toMatchObject({
			name: 'HOW MUCH',
			totalsRowFunction: 'sum',
			dataDxfId: 1,
			totalsRowDxfId: 2,
		})
	})

	test('preserves LibreOffice pivot caches that intentionally omit cache records', () => {
		const initial = readXlsx(
			loadFixture(
				'PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithoutCacheData.xlsx',
			),
		)
		expectOk(initial)

		expect(initial.value.workbook.pivotCaches).toHaveLength(1)
		expect(initial.value.workbook.pivotCaches[0]?.recordsPartPath).toBeUndefined()
		expect(initial.value.workbook.pivotTables).toHaveLength(1)
	})

	test('surfaces LibreOffice pivot cache source and refresh upgrade metadata', () => {
		const initial = readXlsx(loadFixture('pivottable_date_field_filter.xlsx'))
		expectOk(initial)

		expect(initial.value.workbook.pivotCaches[0]).toMatchObject({
			cacheId: 1,
			upgradeOnRefresh: true,
			sourceType: 'worksheet',
			sourceSheet: 'Sheet1',
			sourceRef: 'C1:H4',
		})
	})

	test('surfaces LibreOffice pivot layout style and data-field metadata', () => {
		const initial = readXlsx(loadFixture('PivotTable_CachedDefinitionAndDataInSync.xlsx'))
		expectOk(initial)

		const pivot = initial.value.workbook.pivotTables[0]
		expect(pivot?.location).toMatchObject({
			ref: 'A3:B6',
			firstHeaderRow: 1,
			firstDataRow: 1,
			firstDataCol: 1,
		})
		expect(pivot?.style).toMatchObject({
			name: 'PivotStyleLight15',
			showRowHeaders: true,
			showColHeaders: true,
			showRowStripes: true,
			showColStripes: false,
			showLastColumn: true,
		})
		expect(pivot?.dataFields).toEqual([
			{ fieldIndex: 0, name: 'Sum of A', baseField: 0, baseItem: 0 },
		])
		expect(pivot?.rowItems).toEqual([
			{ index: 0, fieldItems: [{ index: 0 }] },
			{ index: 1, fieldItems: [{ index: 0, item: 1 }] },
			{ index: 2, itemType: 'grand', fieldItems: [{ index: 0 }] },
		])
	})

	test('cached formulas in the LibreOffice subset recalculate without mismatches', async () => {
		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: libreOfficeDir,
			manifest: libreOfficeManifest,
			tags: ['formula-fidelity'],
			tiers: [],
			maxReportedMismatches: 20,
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
			maxMismatches: 0,
			maxUnacceptedMismatches: 0,
			maxSemanticMismatches: 0,
			maxErrors: 0,
			minComparedFormulas: 286,
			minSemanticPerfectWorkbooks: 15,
		})
		expect(payload.summary).toMatchObject({
			workbookCount: 15,
			formulaCount: 286,
			comparedCount: 286,
			mismatchCount: 22,
			acceptedMismatchCount: 22,
			unacceptedMismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 14,
			semanticPerfectWorkbookCount: 15,
		})
	})
})
