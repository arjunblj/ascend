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
		expect(
			entries.find((entry) => entry.file === 'Misc_SheetProtection.xlsx')?.features
				.sheet_protection,
		).toBe(true)
	})

	test('captures full ClosedXML sparkline group metadata and all member ranges', () => {
		const initial = readXlsx(loadFixture('Sparklines_SampleSparklines.xlsx'))
		expectOk(initial)

		const sheets = new Map(initial.value.workbook.sheets.map((sheet) => [sheet.name, sheet]))
		expect([...sheets.values()].map((sheet) => [sheet.name, sheet.sparklineGroups.length])).toEqual(
			[
				['Linear', 6],
				['Column', 6],
				['Stacked', 6],
			],
		)
		expect(sheets.get('Linear')?.sparklineGroups.map((group) => group.type)).toEqual([
			'line',
			'line',
			'line',
			'line',
			'line',
			'line',
		])
		expect(sheets.get('Column')?.sparklineGroups.map((group) => group.type)).toEqual([
			'column',
			'column',
			'column',
			'column',
			'column',
			'column',
		])
		expect(sheets.get('Stacked')?.sparklineGroups.map((group) => group.type)).toEqual([
			'stacked',
			'stacked',
			'stacked',
			'stacked',
			'stacked',
			'stacked',
		])

		const linear = sheets.get('Linear')?.sparklineGroups ?? []
		expect(linear[0]).toMatchObject({
			groupIndex: 0,
			count: 3,
			lineWeight: 0.75,
			markers: true,
			highPoint: true,
			lowPoint: true,
			firstPoint: true,
			lastPoint: true,
			negative: true,
			displayHidden: false,
			rightToLeft: false,
			minAxisType: 'group',
			maxAxisType: 'group',
			colorSeries: 'FF5F5F5F',
			colorNegative: 'FFFFB620',
			colorAxis: 'FF000000',
			colorMarkers: 'FFD70077',
			colorFirst: 'FF5687C2',
			colorLast: 'FF359CEB',
			colorHigh: 'FF56BE79',
			colorLow: 'FFFF5055',
		})
		expect(linear[0]?.sparklines).toEqual([
			{ range: 'Linear!C2:P2', locationRange: 'B2' },
			{ range: 'Linear!C3:P3', locationRange: 'B3' },
			{ range: 'Linear!C4:P4', locationRange: 'B4' },
		])
		expect(linear[2]).toMatchObject({
			manualMax: 100,
			manualMin: -80,
			minAxisType: 'custom',
			maxAxisType: 'custom',
			negative: true,
			colorSeries: 'FFC6EFCE',
		})
		expect(linear[2]?.sparklines?.[2]).toEqual({
			range: 'Linear!C10:P10',
			locationRange: 'B10',
		})
		expect(linear[3]).toMatchObject({
			dateAxis: true,
			dateAxisRange: 'Linear!C1:P1',
			range: 'Linear!C11:P11',
			locationRange: 'B11',
		})
		expect(linear[4]).toMatchObject({
			lineWeight: 2,
			displayXAxis: true,
			rightToLeft: true,
			colorSeries: 'FF00B050',
			colorAxis: 'FFFF0000',
		})
		expect(linear[5]).toMatchObject({ firstPoint: true, lastPoint: true })
		expect(linear[5]?.sparklines?.[2]).toEqual({
			range: 'Linear!C19:E19',
			locationRange: 'B19',
		})

		const written = writeXlsx(initial.value.workbook, initial.value.capsules, {
			dirtySheetNames: ['Linear'],
		})
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedLinear = reopened.value.workbook.sheets.find((sheet) => sheet.name === 'Linear')
		expect(reopenedLinear?.sparklineGroups[0]?.sparklines?.[2]).toEqual({
			range: 'Linear!C4:P4',
			locationRange: 'B4',
		})
		expect(reopenedLinear?.sparklineGroups[3]).toMatchObject({
			dateAxisRange: 'Linear!C1:P1',
			range: 'Linear!C11:P11',
		})
	})

	test('captures ClosedXML x14 data-bar semantics across all six regions', () => {
		const initial = readXlsx(loadFixture('ConditionalFormatting_CFDataBars.xlsx'))
		expectOk(initial)
		const sheet = initial.value.workbook.sheets.find((entry) => entry.name === 'Sheet1')
		expect(sheet).toBeDefined()

		expect(
			sheet?.conditionalFormats.map((format) => {
				const rule = format.rules[0]
				return {
					sqref: format.sqref,
					type: rule?.type,
					priority: rule?.priority,
					cfvo: rule?.dataBar?.cfvo,
					color: rule?.dataBar?.color,
					showValue: rule?.dataBar?.showValue,
				}
			}),
		).toEqual([
			{
				sqref: 'A2:A6',
				type: 'dataBar',
				priority: 1,
				cfvo: [{ type: 'min' }, { type: 'max' }],
				color: { rgb: 'FFFFBF00' },
				showValue: true,
			},
			{
				sqref: 'B2:B6',
				type: 'dataBar',
				priority: 2,
				cfvo: [
					{ type: 'min', value: '0' },
					{ type: 'max', value: '0' },
				],
				color: { rgb: 'FF21ABCD' },
				showValue: true,
			},
			{
				sqref: 'C2:C6',
				type: 'dataBar',
				priority: 3,
				cfvo: [
					{ type: 'num', value: '0' },
					{ type: 'num', value: '10' },
				],
				color: { rgb: 'FF536872' },
				showValue: true,
			},
			{
				sqref: 'D2:D6',
				type: 'dataBar',
				priority: 4,
				cfvo: [
					{ type: 'percent', value: '50' },
					{ type: 'percent', value: '100' },
				],
				color: { rgb: 'FFC19A6B' },
				showValue: true,
			},
			{
				sqref: 'E2:E6',
				type: 'dataBar',
				priority: 5,
				cfvo: [
					{ type: 'formula', value: '-SUM($A$2:$E$2)' },
					{ type: 'formula', value: 'SUM($A$6:$E$6)' },
				],
				color: { rgb: 'FFC2B280' },
				showValue: true,
			},
			{
				sqref: 'F2:F6',
				type: 'dataBar',
				priority: 6,
				cfvo: [
					{ type: 'percentile', value: '30' },
					{ type: 'percentile', value: '70' },
				],
				color: { rgb: 'FFB53389' },
				showValue: true,
			},
		])

		const expectedX14DataBars = [
			{
				index: 0,
				sqref: 'A2:A6',
				formulas: [],
				type: 'dataBar',
				id: '{38d9e5d3-73a7-4e18-b3b9-aa1b26e4cc06}',
				dataBar: {
					cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
					minLength: 0,
					maxLength: 100,
					negativeFillColor: { rgb: 'FFFFBF00' },
					axisColor: { rgb: 'FF000000' },
				},
			},
			{
				index: 1,
				sqref: 'B2:B6',
				formulas: ['0', '0'],
				type: 'dataBar',
				id: '{0671ebad-2a52-43f6-a3a9-585018af6f01}',
				dataBar: {
					cfvo: [
						{ type: 'num', value: '0' },
						{ type: 'num', value: '0' },
					],
					minLength: 0,
					maxLength: 100,
					negativeFillColor: { rgb: 'FF21ABCD' },
					axisColor: { rgb: 'FF000000' },
				},
			},
			{
				index: 2,
				sqref: 'C2:C6',
				formulas: ['0', '10'],
				type: 'dataBar',
				id: '{0b10551c-b4e0-4975-ac10-c045e09df4bf}',
				dataBar: {
					cfvo: [
						{ type: 'num', value: '0' },
						{ type: 'num', value: '10' },
					],
					minLength: 0,
					maxLength: 100,
					negativeFillColor: { rgb: 'FF536872' },
					axisColor: { rgb: 'FF000000' },
				},
			},
			{
				index: 3,
				sqref: 'D2:D6',
				formulas: ['50', '100'],
				type: 'dataBar',
				id: '{49ad88bb-d5a9-4684-afcc-8033d8d051eb}',
				dataBar: {
					cfvo: [
						{ type: 'num', value: '50' },
						{ type: 'num', value: '100' },
					],
					minLength: 0,
					maxLength: 100,
					negativeFillColor: { rgb: 'FFC19A6B' },
					axisColor: { rgb: 'FF000000' },
				},
			},
			{
				index: 4,
				sqref: 'E2:E6',
				formulas: ['-SUM($A$2:$E$2)', 'SUM($A$6:$E$6)'],
				type: 'dataBar',
				id: '{93055e58-8649-4a74-9a8e-1e4036062ece}',
				dataBar: {
					cfvo: [
						{ type: 'num', value: '-SUM($A$2:$E$2)' },
						{ type: 'num', value: 'SUM($A$6:$E$6)' },
					],
					minLength: 0,
					maxLength: 100,
					negativeFillColor: { rgb: 'FFC2B280' },
					axisColor: { rgb: 'FF000000' },
				},
			},
			{
				index: 5,
				sqref: 'F2:F6',
				formulas: ['30', '70'],
				type: 'dataBar',
				id: '{01b3b3bb-243e-4add-ae5b-fd208d5e5960}',
				dataBar: {
					cfvo: [
						{ type: 'num', value: '30' },
						{ type: 'num', value: '70' },
					],
					minLength: 0,
					maxLength: 100,
					negativeFillColor: { rgb: 'FFB53389' },
					axisColor: { rgb: 'FF000000' },
				},
			},
		]
		expect(sheet?.x14ConditionalFormats).toEqual(expectedX14DataBars)

		const written = writeXlsx(initial.value.workbook, initial.value.capsules)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.sheets.find((entry) => entry.name === 'Sheet1')
		expect(reopenedSheet?.x14ConditionalFormats).toEqual(expectedX14DataBars)
	})

	test('captures ClosedXML x14 data validation sheets, formulas, and flags', () => {
		const initial = readXlsx(loadFixture('Misc_DataValidation.xlsx'))
		expectOk(initial)
		const expectedX14Validation = {
			index: 0,
			sqref: 'A5:A5',
			type: 'list',
			operator: 'between',
			allowBlank: true,
			showInputMessage: true,
			showErrorMessage: true,
			showDropDown: false,
			errorStyle: 'stop',
			formula1: "'Data Validation'!$C$1:$C$2",
		}
		const validationBySheet = new Map(
			initial.value.workbook.sheets.map((sheet) => [sheet.name, sheet.x14DataValidations]),
		)
		expect([...validationBySheet].map(([name, validations]) => [name, validations.length])).toEqual(
			[
				['Data Validation', 0],
				['Validate Ranges', 0],
				['Data Validation - Copy', 1],
				['Validate Ranges - Copy', 0],
				['Copy From Range 1', 1],
				['Copy From Range 2', 0],
			],
		)
		expect(validationBySheet.get('Data Validation - Copy')).toEqual([expectedX14Validation])
		expect(validationBySheet.get('Copy From Range 1')).toEqual([expectedX14Validation])

		for (const sheetName of ['Data Validation - Copy', 'Copy From Range 1']) {
			const sheet = initial.value.workbook.sheets.find((entry) => entry.name === sheetName)
			expect(sheet?.dataValidations.find((validation) => validation.source === 'x14')).toEqual({
				sqref: expectedX14Validation.sqref,
				type: expectedX14Validation.type,
				operator: expectedX14Validation.operator,
				errorStyle: expectedX14Validation.errorStyle,
				allowBlank: expectedX14Validation.allowBlank,
				showInputMessage: expectedX14Validation.showInputMessage,
				showErrorMessage: expectedX14Validation.showErrorMessage,
				showDropDown: expectedX14Validation.showDropDown,
				formula1: expectedX14Validation.formula1,
				source: 'x14',
			})
		}

		const written = writeXlsx(initial.value.workbook, initial.value.capsules)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedValidationBySheet = new Map(
			reopened.value.workbook.sheets.map((sheet) => [sheet.name, sheet.x14DataValidations]),
		)
		expect(reopenedValidationBySheet.get('Data Validation - Copy')).toEqual([expectedX14Validation])
		expect(reopenedValidationBySheet.get('Copy From Range 1')).toEqual([expectedX14Validation])
	})

	test('captures ClosedXML sheet protection flags and password hash', () => {
		const result = readXlsx(loadFixture('Misc_SheetProtection.xlsx'))
		expectOk(result)

		expect(result.value.workbook.sheets.map((sheet) => [sheet.name, sheet.protection])).toEqual([
			[
				'Protected No-Password',
				{
					sheet: true,
					objects: true,
					scenarios: false,
					formatCells: false,
					insertColumns: false,
					deleteColumns: false,
					deleteRows: false,
				},
			],
			[
				'Protected Password = 123',
				{
					sheet: true,
					objects: true,
					insertColumns: false,
					insertRows: false,
					password: 'CF7A',
				},
			],
		])
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
