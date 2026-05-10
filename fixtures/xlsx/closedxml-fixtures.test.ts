import { describe, expect, test } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import type { Workbook } from '@ascend/core'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import { runFormulaCorpusCorrectness } from '../benchmarks/formula-corpus-correctness.ts'
import {
	normalizeManifest,
	selectManifestEntries,
	validateManifestProvenance,
} from '../corpus/manifest.ts'
import { loadManifest } from './closedxml/manifest.ts'

const closedXmlDir = fileURLToPath(new URL('./closedxml/', import.meta.url))
const closedXmlManifest = fileURLToPath(new URL('./closedxml/manifest.ts', import.meta.url))

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./closedxml/${name}`, import.meta.url))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function expectNumberCell(
	workbook: Workbook,
	sheetName: string,
	row: number,
	col: number,
	expected: number,
): void {
	const sheet = workbook.sheets.find((entry) => entry.name === sheetName)
	expect(sheet).toBeDefined()
	const value = sheet?.cells.get(row, col)?.value
	expect(value?.kind).toBe('number')
	if (value?.kind === 'number') expect(value.value).toBeCloseTo(expected, 10)
}

function expectStringCell(
	workbook: Workbook,
	sheetName: string,
	row: number,
	col: number,
	expected: string,
): void {
	const sheet = workbook.sheets.find((entry) => entry.name === sheetName)
	expect(sheet).toBeDefined()
	expect(sheet?.cells.get(row, col)?.value).toEqual({ kind: 'string', value: expected })
}

describe('ClosedXML XLSX fixture corpus', () => {
	test('manifest has pinned provenance for the vendored MIT fixture subset', async () => {
		expect(existsSync(new URL('./closedxml/LICENSE', import.meta.url))).toBe(true)
		const entries = normalizeManifest(await loadManifest())
		expect(entries.length).toBe(26)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['closedxml'] })).toHaveLength(entries.length)
		expect(
			selectManifestEntries(entries, { tags: ['formula-fidelity'] }).map((entry) => entry.file),
		).toEqual([
			'Misc_FormulasWithEvaluation.xlsx',
			'Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
			'Other_Formulas_BooleanFormulaValues.xlsx',
			'Other_Formulas_DataTableFormula-Excel-Input.xlsx',
		])
		expect(selectManifestEntries(entries, { tags: ['pivot-table'] }).length).toBeGreaterThan(0)
		expect(selectManifestEntries(entries, { tags: ['conditional-formatting'] }).length).toBe(2)
	})

	test('opens and round-trips every ClosedXML fixture without crashing', async () => {
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

	test('captures expected feature families from ClosedXML regression fixtures', async () => {
		const entries = normalizeManifest(await loadManifest())
		expect(
			entries.find((entry) => entry.file === 'Other_Charts_PreserveCharts_inputfile.xlsx')?.features
				.charts,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'Other_ExternalLinks_WorkbookWithExternalLink.xlsx')
				?.features.external_links,
		).toBe(true)
		expect(entries.find((entry) => entry.file === 'Tables_UsingTables.xlsx')?.features.tables).toBe(
			true,
		)
		expect(
			entries.find((entry) => entry.file === 'Comments_AddingComments.xlsx')?.features.comments,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'Ranges_DefinedNames.xlsx')?.features.defined_names,
		).toBe(true)
	})

	test('captures ClosedXML comment VML visibility and position metadata', () => {
		const result = readXlsx(loadFixture('Comments_AddingComments.xlsx'))
		expectOk(result)
		const visibility = result.value.workbook.sheets.find((sheet) => sheet.name === 'Visibility')
		const position = result.value.workbook.sheets.find((sheet) => sheet.name === 'Position')
		expect(visibility?.comments.get('A1')?.legacyDrawing).toMatchObject({
			shapeId: '_x0000_s14',
			anchor: [1, 15, 0, 0, 3, 33, 3, 14],
			row: 0,
			column: 0,
			visible: false,
		})
		expect(visibility?.comments.get('A2')?.legacyDrawing).toMatchObject({
			shapeId: '_x0000_s15',
			visible: true,
		})
		expect(position?.comments.get('A1')?.legacyDrawing).toMatchObject({
			shapeId: '_x0000_s18',
			anchor: [2, 38, 4, 8, 4, 32, 8, 7],
			row: 0,
			column: 0,
			visible: true,
		})
	})

	test('resolves ClosedXML external-link formula usages to package targets', async () => {
		const wb = await AscendWorkbook.open(
			loadFixture('Other_ExternalLinks_WorkbookWithExternalLink.xlsx'),
		)
		expect(wb.externalReferenceUsages()).toEqual([
			{
				workbook: '1',
				sheet: 'Sheet1',
				sourceKind: 'cellFormula',
				sourceRef: 'Sheet1!B2',
				formula: '[1]Sheet1!$A$1',
				references: ['[1]Sheet1!$A$1'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					relId: 'rId2',
					linkRelId: 'rId1',
					target: 'book1.xlsx',
					targetMode: 'External',
				},
			},
		])
	})

	test('inventories real ClosedXML image anchors and media relationships', async () => {
		const wb = await AscendWorkbook.open(loadFixture('ImageHandling_ImageAnchors.xlsx'))
		const visuals = wb.visualInventory()
		expect(visuals.sheetImageCount).toBe(7)

		const sheets = new Map(
			visuals.sheets.map((sheet) => [sheet.sheet, sheet.imageRefs ?? []] as const),
		)
		expect([...sheets].map(([sheet, refs]) => [sheet, refs.length])).toEqual([
			['Images1', 2],
			['Images2', 1],
			['Images3', 3],
			['Images4', 1],
		])
		expect(
			visuals.sheets.flatMap((sheet) => sheet.imageRefs ?? []).map((image) => image.anchor?.kind),
		).toEqual(['oneCell', 'absolute', 'twoCell', 'oneCell', 'twoCell', 'twoCell', 'absolute'])
		expect(sheets.get('Images1')?.map((image) => [image.relId, image.targetPath])).toEqual([
			['rId10', 'xl/media/image2.png'],
			['rId9', 'xl/media/image.png'],
		])
		expect(sheets.get('Images3')?.map((image) => [image.relId, image.targetPath])).toEqual([
			['rId16', 'xl/media/image3.jpg'],
			['rId14', 'xl/media/image.jpg'],
			['rId15', 'xl/media/image2.jpg'],
		])
	})

	test('cached formulas in the ClosedXML subset recalculate without mismatches', async () => {
		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: closedXmlDir,
			manifest: closedXmlManifest,
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
			minComparedFormulas: 23,
			minSemanticPerfectWorkbooks: 4,
		})
		expect(payload.summary).toMatchObject({
			workbookCount: 4,
			formulaCount: 23,
			comparedCount: 23,
			mismatchCount: 0,
			acceptedMismatchCount: 0,
			unacceptedMismatchCount: 0,
			semanticMismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 4,
			semanticPerfectWorkbookCount: 4,
		})
	})

	test('formula workbooks without cached values recalculate deterministically', () => {
		const formulas = readXlsx(loadFixture('Misc_Formulas.xlsx'))
		expectOk(formulas)
		recalculate(formulas.value.workbook, defaultCalcContext())
		expectNumberCell(formulas.value.workbook, 'Formulas', 1, 2, 3)
		expectStringCell(formulas.value.workbook, 'Formulas', 1, 6, 'Yes')
		expectStringCell(formulas.value.workbook, 'Formulas', 3, 2, 'TestAR3C2')
		expectNumberCell(formulas.value.workbook, 'Formulas', 5, 1, 2)
		expectNumberCell(formulas.value.workbook, 'Formulas', 5, 2, 1)
		expect(formulas.value.workbook.sheets[0]?.cells.get(5, 2)?.formulaInfo).toEqual({
			kind: 'array',
			ref: 'C6:D6',
		})

		const shifting = readXlsx(loadFixture('Misc_ShiftingFormulas.xlsx'))
		expectOk(shifting)
		recalculate(shifting.value.workbook, defaultCalcContext())
		expectNumberCell(shifting.value.workbook, 'Shifting Formulas', 2, 4, 6)
		expectNumberCell(shifting.value.workbook, 'Shifting Formulas', 2, 6, 3.5)
		expectNumberCell(shifting.value.workbook, 'Shifting Formulas', 4, 2, 11)
		expectNumberCell(shifting.value.workbook, 'WS2', 0, 0, 5)
		expectNumberCell(shifting.value.workbook, 'WS2', 8, 0, 3.5)

		const arrayFormula = readXlsx(loadFixture('Other_Formulas_ArrayFormula.xlsx'))
		expectOk(arrayFormula)
		recalculate(arrayFormula.value.workbook, defaultCalcContext())
		expectNumberCell(arrayFormula.value.workbook, 'Sheet1', 0, 0, 3)
	})
})
