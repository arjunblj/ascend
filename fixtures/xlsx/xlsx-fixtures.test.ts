import { describe, expect, it } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { fileURLToPath } from 'node:url'
import { unzipSync } from 'fflate'
import type { Sheet, Workbook } from '../../packages/core/src/index.ts'
import { parseRange, toA1 } from '../../packages/core/src/index.ts'
import {
	applyOperation,
	defaultCalcContext,
	evaluateConditionalFormats,
	recalculate,
	validateCellValue,
} from '../../packages/engine/src/index.ts'
import { shiftSheetCellMetadata } from '../../packages/engine/src/structural/sheet-topology.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import {
	fingerprintXlsx,
	fingerprintXlsxPart,
} from '../../packages/io-xlsx/test/fidelity-harness.ts'
import { EMPTY } from '../../packages/schema/src/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import { runFormulaCorpusCorrectness } from '../benchmarks/formula-corpus-correctness.ts'

const poiDir = fileURLToPath(new URL('./poi/', import.meta.url))
const poiManifest = fileURLToPath(new URL('./poi/manifest.ts', import.meta.url))
const poiFixtures = readdirSync(poiDir)
	.filter((name) => name.endsWith('.xlsx'))
	.sort((a, b) => a.localeCompare(b))

interface NoOpFidelityFixture {
	readonly name: string
	readonly preservedParts?: readonly string[]
	readonly worksheetTags: Readonly<Record<string, number>>
}

const noOpFidelityFixtures: readonly NoOpFidelityFixture[] = [
	{
		name: 'WithChart.xlsx',
		preservedParts: ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml'],
		worksheetTags: { drawing: 1 },
	},
	{
		name: 'WithDrawing.xlsx',
		preservedParts: ['xl/drawings/drawing1.xml'],
		worksheetTags: { drawing: 1 },
	},
	{
		name: 'SimpleWithComments.xlsx',
		preservedParts: ['xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml'],
		worksheetTags: { legacyDrawing: 1 },
	},
	{
		name: 'DataValidationEvaluations.xlsx',
		worksheetTags: { dataValidations: 1, dataValidation: 17 },
	},
	{
		name: 'NewStyleConditionalFormattings.xlsx',
		worksheetTags: {
			conditionalFormatting: 18,
			cfRule: 19,
			extLst: 2,
			'x14:conditionalFormatting': 3,
			'x14:cfRule': 3,
		},
	},
	{
		name: 'StructuredReferences.xlsx',
		preservedParts: ['xl/tables/table1.xml'],
		worksheetTags: { tableParts: 1, tablePart: 1 },
	},
]

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./poi/${name}`, import.meta.url))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function noOpRoundTripFixture(name: string): { original: Uint8Array; written: Uint8Array } {
	const original = loadFixture(name)
	const initial = readXlsx(original)
	expectOk(initial)
	const written = writeXlsx(initial.value.workbook, initial.value.capsules)
	expectOk(written)
	const reopened = readXlsx(written.value)
	expectOk(reopened)
	expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
	return { original, written: written.value }
}

function readZipText(bytes: Uint8Array, path: string): string {
	const entry = unzipSync(bytes)[path]
	expect(entry).toBeDefined()
	return new TextDecoder().decode(entry)
}

function expectPreservedXmlParts(
	original: Uint8Array,
	written: Uint8Array,
	paths: readonly string[],
): void {
	for (const path of paths) {
		const before = fingerprintXlsxPart(original, path)
		const after = fingerprintXlsxPart(written, path)
		expect(before).toBeDefined()
		expect(after).toBeDefined()
		expect(after?.xml.normalized).toBe(before?.xml.normalized)
		expect(after?.xml.tagCounts).toEqual(before?.xml.tagCounts)
	}
}

function expectWorksheetTagCountsPreserved(
	original: Uint8Array,
	written: Uint8Array,
	expectedCounts: Readonly<Record<string, number>>,
): void {
	const before = fingerprintXlsx(original)
	const after = fingerprintXlsx(written)
	expect(after.sheets.map((sheet) => sheet.path)).toEqual(before.sheets.map((sheet) => sheet.path))

	for (const [tag, expectedCount] of Object.entries(expectedCounts)) {
		const beforeCount = countWorksheetTag(before, tag)
		expect(beforeCount).toBe(expectedCount)
		expect(countWorksheetTag(after, tag)).toBe(beforeCount)
	}
}

function countWorksheetTag(fingerprint: ReturnType<typeof fingerprintXlsx>, tag: string): number {
	return fingerprint.sheets.reduce((sum, sheet) => sum + (sheet.xml.tagCounts[tag] ?? 0), 0)
}

function matchingConditionalFormatRefs(
	workbook: Workbook,
	sheet: Sheet,
	type: string,
	priority: number,
): string[] {
	const matches = evaluateConditionalFormats(sheet, workbook)
	return [...matches.entries()]
		.filter(([, rules]) => rules.some((rule) => rule.type === type && rule.priority === priority))
		.map(([ref]) => ref)
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

if (poiFixtures.length > 0) {
	describe('POI XLSX fixtures', () => {
		it('opens all POI .xlsx files without crashing', () => {
			const results: { name: string; ok: boolean; error?: string }[] = []
			for (const name of poiFixtures) {
				const result = readXlsx(loadFixture(name), { mode: 'values' })
				results.push({
					name,
					ok: result.ok,
					error: result.ok ? undefined : result.error.message,
				})
			}
			const passed = results.filter((r) => r.ok).length
			const total = results.length
			const pct = total > 0 ? (passed / total) * 100 : 0
			for (const r of results) {
				if (r.ok) {
					console.log(`  ✓ ${r.name}`)
				} else {
					console.log(`  ✗ ${r.name}: ${r.error}`)
				}
			}
			console.log(`  ${passed}/${total} opened (${pct.toFixed(1)}%)`)
			expect(total).toBeGreaterThan(0)
			expect(pct).toBe(100)
		})

		for (const fixture of poiFixtures) {
			it(`reads ${fixture}`, () => {
				const result = readXlsx(loadFixture(fixture))
				expectOk(result)
				expect(result.value.workbook.sheets.length).toBeGreaterThan(0)
			})

			it(`round-trips ${fixture}`, () => {
				const initial = readXlsx(loadFixture(fixture))
				expectOk(initial)
				const written = writeXlsx(initial.value.workbook, initial.value.capsules)
				expectOk(written)
				const reopened = readXlsx(written.value)
				expectOk(reopened)
				expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
			})
		}

		it('evaluates conditional formatting rules from ConditionalFormattingSamples.xlsx', () => {
			const result = readXlsx(loadFixture('ConditionalFormattingSamples.xlsx'))
			expectOk(result)
			const workbook = result.value.workbook

			const products = workbook.sheets.find((sheet) => sheet.name === 'Products1')
			expect(products).toBeDefined()
			if (!products) return
			expect(matchingConditionalFormatRefs(workbook, products, 'containsText', 4)).toEqual([
				'B9',
				'B13',
				'B14',
				'B15',
				'B16',
			])
			expect(matchingConditionalFormatRefs(workbook, products, 'cellIs', 7)).toEqual([
				'D7',
				'D8',
				'D11',
				'D12',
				'D14',
				'D15',
				'D17',
				'D21',
				'D23',
			])

			const grades = workbook.sheets.find((sheet) => sheet.name === 'Grades')
			expect(grades).toBeDefined()
			if (!grades) return
			expect(matchingConditionalFormatRefs(workbook, grades, 'top10', 1)).toEqual(['F3', 'F10'])

			const customers = workbook.sheets.find((sheet) => sheet.name === 'Customers1')
			expect(customers).toBeDefined()
			if (!customers) return
			expect(matchingConditionalFormatRefs(workbook, customers, 'duplicateValues', 5)).toEqual([
				'A7',
				'A10',
				'A11',
				'A12',
				'A15',
				'A17',
				'A19',
				'A20',
				'A21',
			])

			const bandedRows = workbook.sheets.find((sheet) => sheet.name === 'Banded rows')
			expect(bandedRows).toBeDefined()
			if (!bandedRows) return
			expect(matchingConditionalFormatRefs(workbook, bandedRows, 'expression', 1)).toHaveLength(110)
			expect(
				matchingConditionalFormatRefs(workbook, bandedRows, 'expression', 1).slice(0, 10),
			).toEqual(['A5', 'A7', 'A9', 'A11', 'A13', 'A15', 'A17', 'A19', 'A21', 'A23'])
		})

		it('preserves ConditionalFormattingSamples package defaults and customXml rel targets', () => {
			const { written } = noOpRoundTripFixture('ConditionalFormattingSamples.xlsx')
			const contentTypes = readZipText(written, '[Content_Types].xml')
			const workbookRels = readZipText(written, 'xl/_rels/workbook.xml.rels')

			expect(contentTypes).toContain('<Default Extension="png" ContentType="image/png"/>')
			expect(contentTypes).toContain(
				'<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>',
			)
			expect(contentTypes).not.toContain('PartName="/xl/media/image1.png"')
			expect(contentTypes).not.toContain('PartName="/xl/printerSettings/printerSettings1.bin"')
			expect(contentTypes).not.toContain('PartName="/customXml/item1.xml"')
			expect(workbookRels).toContain('Target="../customXml/item1.xml"')
			expect(workbookRels).not.toContain('Target="customXml/item1.xml"')
		})

		it('preserves x14/extLst conditional formatting payloads on round-trip', () => {
			const initial = readXlsx(loadFixture('NewStyleConditionalFormattings.xlsx'))
			expectOk(initial)
			const sheet = initial.value.workbook.sheets[0]
			expect(sheet?.preservedExtLst?.includes('<extLst')).toBe(true)
			expect(sheet?.x14ConditionalFormats.map((format) => format.sqref)).toEqual([
				'E2:E17',
				'Q2:Q17',
				'U2:U17',
			])
			expect(sheet?.x14ConditionalFormats[0]).toMatchObject({
				index: 0,
				type: 'dataBar',
				id: '{9B4F274F-F774-40EE-9C50-A8B810847010}',
				dataBar: {
					minLength: 0,
					maxLength: 100,
					border: true,
					negativeBarBorderColorSameAsPositive: false,
					cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
					borderColor: { rgb: 'FF63C384' },
					negativeFillColor: { rgb: 'FFFF0000' },
					negativeBorderColor: { rgb: 'FFFF0000' },
					axisColor: { rgb: 'FF000000' },
				},
			})
			expect(sheet?.x14ConditionalFormats[1]).toMatchObject({
				index: 1,
				type: 'iconSet',
				priority: 6,
				id: '{23404A22-2791-4FFE-B3BF-645D3DA783FE}',
				formulas: ['0', '33', '67'],
				iconSet: {
					iconSet: '3Stars',
					cfvo: [
						{ type: 'percent', value: '0' },
						{ type: 'percent', value: '33' },
						{ type: 'percent', value: '67' },
					],
				},
			})
			expect(sheet?.x14ConditionalFormats[2]).toMatchObject({
				index: 2,
				type: 'iconSet',
				priority: 1,
				id: '{F6A45A70-DEA9-40B0-83E9-E53A5F349BE7}',
				iconSet: {
					custom: true,
					cfvo: [
						{ type: 'percent', value: '0' },
						{ type: 'num', value: '0' },
						{ type: 'num', value: '30' },
					],
					icons: [
						{ iconSet: '3Signs', iconId: 0 },
						{ iconSet: '3Flags', iconId: 1 },
						{ iconSet: '3Symbols2', iconId: 2 },
					],
				},
			})

			const written = writeXlsx(initial.value.workbook, initial.value.capsules)
			expectOk(written)
			const reopened = readXlsx(written.value)
			expectOk(reopened)
			expect(reopened.value.workbook.sheets[0]?.preservedExtLst?.includes('<extLst')).toBe(true)
		})

		it('rewrites real x14/extLst conditional formatting ranges after metadata shifts', () => {
			const initial = readXlsx(loadFixture('NewStyleConditionalFormattings.xlsx'))
			expectOk(initial)
			const workbook = initial.value.workbook
			const sheet = workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			shiftSheetCellMetadata(sheet, 'row', 1, 1)
			expect(sheet?.x14ConditionalFormats.map((format) => format.sqref)).toEqual([
				'E3:E18',
				'Q3:Q18',
				'U3:U18',
			])

			const written = writeXlsx(workbook, initial.value.capsules, { dirtySheetNames: ['CF'] })
			expectOk(written)
			const worksheetXml = readZipText(written.value, 'xl/worksheets/sheet1.xml')
			expect(worksheetXml).toContain('<xm:sqref>E3:E18</xm:sqref>')
			expect(worksheetXml).toContain('<xm:sqref>Q3:Q18</xm:sqref>')
			expect(worksheetXml).toContain('<xm:sqref>U3:U18</xm:sqref>')
			expect(worksheetXml).not.toContain('<xm:sqref>E2:E17</xm:sqref>')
			expect(worksheetXml).not.toContain('<xm:sqref>Q2:Q17</xm:sqref>')
			expect(worksheetXml).not.toContain('<xm:sqref>U2:U17</xm:sqref>')
		})

		it('removes real x14/extLst conditional formatting entries deleted by metadata shifts', () => {
			const initial = readXlsx(loadFixture('NewStyleConditionalFormattings.xlsx'))
			expectOk(initial)
			const workbook = initial.value.workbook
			const sheet = workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			shiftSheetCellMetadata(sheet, 'row', 1, -16)
			expect(sheet.x14ConditionalFormats.every((format) => format.deleted)).toBe(true)

			const written = writeXlsx(workbook, initial.value.capsules, { dirtySheetNames: ['CF'] })
			expectOk(written)
			const worksheetXml = readZipText(written.value, 'xl/worksheets/sheet1.xml')
			expect(worksheetXml).not.toContain('<xm:sqref>E2:E17</xm:sqref>')
			expect(worksheetXml).not.toContain('<xm:sqref>Q2:Q17</xm:sqref>')
			expect(worksheetXml).not.toContain('<xm:sqref>U2:U17</xm:sqref>')
			expect(worksheetXml).not.toContain('<x14:conditionalFormatting ')
			expect(worksheetXml).not.toContain('<x14:cfRule')
		})

		it('captures data validation rules from DataValidationEvaluations.xlsx', () => {
			const result = readXlsx(loadFixture('DataValidationEvaluations.xlsx'))
			expectOk(result)
			const count = result.value.workbook.sheets.reduce(
				(sum, sheet) => sum + (sheet?.dataValidations.length ?? 0),
				0,
			)
			expect(count).toBeGreaterThan(0)
		})

		it('evaluates POI data validation rules against expected Excel outcomes', () => {
			const result = readXlsx(loadFixture('DataValidationEvaluations.xlsx'))
			expectOk(result)
			const workbook = result.value.workbook
			const sheet = workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return

			for (let row = 2; row <= 34; row++) {
				const value = sheet.cells.get(row, 1)?.value ?? EMPTY
				const expected = sheet.cells.get(row, 2)?.value
				expect(expected?.kind).toBe('boolean')
				if (expected?.kind !== 'boolean') continue
				expect(validateCellValue(sheet, row, 1, value, workbook).valid).toBe(expected.value)
			}
		})

		it('captures structured references tables from StructuredReferences.xlsx', () => {
			const result = readXlsx(loadFixture('StructuredReferences.xlsx'))
			expectOk(result)
			const count = result.value.workbook.sheets.reduce(
				(sum, sheet) => sum + (sheet?.tables.length ?? 0),
				0,
			)
			expect(count).toBeGreaterThan(0)
		})

		it('captures comments from SimpleWithComments.xlsx', () => {
			const result = readXlsx(loadFixture('SimpleWithComments.xlsx'))
			expectOk(result)
			const count = result.value.workbook.sheets.reduce(
				(sum, sheet) => sum + (sheet?.comments.size ?? 0),
				0,
			)
			expect(count).toBeGreaterThan(0)
		})

		for (const fixture of noOpFidelityFixtures) {
			it(`preserves no-op package fidelity for ${fixture.name}`, () => {
				const { original, written } = noOpRoundTripFixture(fixture.name)
				if (fixture.preservedParts) {
					expectPreservedXmlParts(original, written, fixture.preservedParts)
				}
				expectWorksheetTagCountsPreserved(original, written, fixture.worksheetTags)
			})
		}

		it('loads multiple sheets from 55906-MultiSheetRefs.xlsx', () => {
			const result = readXlsx(loadFixture('55906-MultiSheetRefs.xlsx'))
			expectOk(result)
			expect(result.value.workbook.sheets.length).toBeGreaterThan(1)
		})

		it('preserves calcChain package wiring on no-op round-trip', () => {
			const { original, written } = noOpRoundTripFixture('55906-MultiSheetRefs.xlsx')
			const beforeChain = fingerprintXlsxPart(original, 'xl/calcChain.xml')
			const afterChain = fingerprintXlsxPart(written, 'xl/calcChain.xml')
			const afterContentTypes = fingerprintXlsxPart(written, '[Content_Types].xml')
			const afterWorkbookRels = fingerprintXlsxPart(written, 'xl/_rels/workbook.xml.rels')

			expect(beforeChain).toBeDefined()
			expect(afterChain?.xml.normalized).toBe(beforeChain?.xml.normalized)
			expect(afterContentTypes?.xml.normalized).toContain('calcChain+xml')
			expect(afterWorkbookRels?.xml.normalized).toContain('relationships/calcChain')
		})

		it('captures workbook or sheet protection from POI protection fixtures', () => {
			const workbookProtected = readXlsx(
				loadFixture('workbookProtection_workbook_structure_protected.xlsx'),
			)
			expectOk(workbookProtected)
			expect(workbookProtected.value.workbook.workbookProtection).toEqual({
				lockStructure: true,
			})

			const sheetProtected = readXlsx(loadFixture('sheetProtection_allLocked.xlsx'))
			expectOk(sheetProtected)
			expect(sheetProtected.value.workbook.sheets[0]?.name).toBe('Foglio1')
			expect(sheetProtected.value.workbook.sheets[0]?.protection).toEqual({
				sheet: true,
				objects: true,
				scenarios: true,
				selectLockedCells: true,
				selectUnlockedCells: true,
			})
			expect(
				sheetProtected.value.workbook.sheets.slice(1).every((sheet) => sheet.protection === null),
			).toBe(true)

			const notProtected = readXlsx(loadFixture('sheetProtection_not_protected.xlsx'))
			expectOk(notProtected)
			expect(notProtected.value.workbook.workbookProtection).toBeNull()
			expect(notProtected.value.workbook.sheets.every((sheet) => sheet.protection === null)).toBe(
				true,
			)
		})

		it('keeps fixture list stable and non-empty', () => {
			expect(poiFixtures.length).toBeGreaterThan(20)
			expect(poiFixtures.map((name) => basename(name))).toContain('shared_formulas.xlsx')
			expect(poiFixtures.map((name) => basename(name))).toContain('SimpleStrict.xlsx')
		})

		it('parses number format codes from NumberFormatTests.xlsx', () => {
			const result = readXlsx(loadFixture('NumberFormatTests.xlsx'))
			expectOk(result)
			const wb = result.value.workbook
			let hasCustomFormat = false
			for (let i = 0; i < wb.styles.size; i++) {
				const style = wb.styles.get(i as never)
				if (style?.numberFormat && style.numberFormat !== 'General') {
					hasCustomFormat = true
					break
				}
			}
			expect(hasCustomFormat).toBe(true)
		})

		it('parses fonts and fills from styles.xlsx', () => {
			const result = readXlsx(loadFixture('styles.xlsx'))
			expectOk(result)
			const wb = result.value.workbook
			let hasFontStyle = false
			let hasFillStyle = false
			for (let i = 0; i < wb.styles.size; i++) {
				const style = wb.styles.get(i as never)
				if (style?.font?.bold || style?.font?.italic || style?.font?.name) hasFontStyle = true
				if (style?.fill?.pattern && style.fill.pattern !== 'none') hasFillStyle = true
			}
			expect(hasFontStyle).toBe(true)
			expect(hasFillStyle).toBe(true)
		})

		it('parses CF rule types from WithConditionalFormatting.xlsx', () => {
			const result = readXlsx(loadFixture('WithConditionalFormatting.xlsx'))
			expectOk(result)
			const rules = result.value.workbook.sheets.flatMap((s) =>
				s.conditionalFormats.flatMap((cf) => cf.rules),
			)
			expect(rules.length).toBeGreaterThan(0)
			const types = new Set(rules.map((r) => r.type))
			expect(types.has('cellIs')).toBe(true)
		})

		it('parses theme metadata from Themes.xlsx', () => {
			const result = readXlsx(loadFixture('Themes.xlsx'))
			expectOk(result)
			expect(result.value.workbook.themeMetadata).toBeDefined()
		})

		it('parses comment text, author, and VML layout from SimpleWithComments.xlsx', () => {
			const result = readXlsx(loadFixture('SimpleWithComments.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			expect([...sheet.comments.keys()]).toEqual(['B1', 'B2', 'B3'])
			for (const ref of ['B1', 'B2', 'B3']) {
				expect(sheet.comments.get(ref)?.author).toBe('Yegor Kozlov')
				expect(sheet.comments.get(ref)?.text.length).toBeGreaterThan(0)
			}
			expect(sheet.comments.get('B1')?.legacyDrawing).toMatchObject({
				shapeId: '_x0000_s1025',
				anchor: [2, 15, 0, 2, 4, 15, 4, 8],
				row: 0,
				column: 1,
				visible: false,
				moveWithCells: true,
				sizeWithCells: true,
				autoFill: false,
			})
			expect(sheet.comments.get('B2')?.legacyDrawing).toMatchObject({
				row: 1,
				column: 1,
				visible: false,
			})
			expect(sheet.comments.get('B3')?.legacyDrawing).toMatchObject({
				shapeId: '_x0000_s1027',
				anchor: [2, 15, 1, 7, 4, 15, 5, 13],
				row: 2,
				column: 1,
				visible: true,
			})
		})

		it('recalculates shared formulas from shared_formulas.xlsx', () => {
			const result = readXlsx(loadFixture('shared_formulas.xlsx'))
			expectOk(result)
			const wb = result.value.workbook
			recalculate(wb, defaultCalcContext())
			const sheet = wb.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			let formulaCount = 0
			for (const [, , cell] of sheet.cells.iterate()) {
				if (cell.formula) formulaCount++
			}
			expect(formulaCount).toBeGreaterThan(0)
		})

		it('reads shared_formulas.xlsx with correct master/member structure', () => {
			const result = readXlsx(loadFixture('shared_formulas.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			const hasSharedMaster = [...sheet.cells.iterate()].some(
				([, , cell]) => cell.formulaInfo?.kind === 'shared' && cell.formulaInfo?.isMaster,
			)
			const hasSharedMember = [...sheet.cells.iterate()].some(
				([, , cell]) => cell.formulaInfo?.kind === 'shared' && !cell.formulaInfo?.isMaster,
			)
			expect(hasSharedMaster).toBe(true)
			expect(hasSharedMember).toBe(true)
			const masterCell = [...sheet.cells.iterate()].find(
				([, , cell]) => cell.formulaInfo?.kind === 'shared' && cell.formulaInfo?.isMaster,
			)
			expect(masterCell).toBeDefined()
			expect(masterCell?.[2]?.formula).toBeDefined()
			expect(masterCell?.[2]?.formula?.length ?? 0).toBeGreaterThan(0)
		})

		it('copies real POI shared formula members as effective formulas', () => {
			const result = readXlsx(loadFixture('shared_formulas.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			const member = [...sheet.cells.iterate()].find(
				([, , cell]) => cell.formulaInfo?.kind === 'shared' && !cell.formulaInfo?.isMaster,
			)
			expect(member).toBeDefined()
			if (!member) return
			const [row, col, cell] = member
			const sourceRef = toA1({ row, col })
			const targetRef = toA1({ row, col: col + 10 })

			const edit = applyOperation(result.value.workbook, {
				op: 'copyRange',
				sheet: sheet.name,
				source: sourceRef,
				target: targetRef,
			})
			expectOk(edit)
			const target = sheet.cells.get(row, col + 10)
			expect(cell.formula).toBeNull()
			expect(target?.formula).toBeTruthy()
			expect(target?.formulaInfo).toBeUndefined()
		})

		it('rejects structural edits that would drop real POI shared formula bindings', () => {
			const result = readXlsx(loadFixture('shared_formulas.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			const member = [...sheet.cells.iterate()].find(
				([, , cell]) => cell.formulaInfo?.kind === 'shared' && !cell.formulaInfo?.isMaster,
			)
			expect(member).toBeDefined()
			if (!member) return

			const edit = applyOperation(result.value.workbook, {
				op: 'insertRows',
				sheet: sheet.name,
				at: 100,
				count: 1,
			})
			expect(edit.ok).toBe(false)
			expect(member[2].formulaInfo).toEqual(sheet.cells.get(member[0], member[1])?.formulaInfo)
		})

		it('reads TestShiftRowSharedFormula.xlsx with shared formulas and correct values', () => {
			const result = readXlsx(loadFixture('TestShiftRowSharedFormula.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			expect(result.value.report.features.some((f) => f.feature === 'sharedFormula')).toBe(true)
			recalculate(result.value.workbook, defaultCalcContext())
			const d5 = sheet.cells.get(4, 3)
			const e5 = sheet.cells.get(4, 4)
			expect(d5?.value).toMatchObject({ kind: 'number', value: 15 })
			expect(e5?.value).toMatchObject({ kind: 'number', value: 18 })
		})

		it('reads and recalculates legacy array formulas from MatrixFormulaEvalTestData.xlsx', () => {
			const result = readXlsx(loadFixture('MatrixFormulaEvalTestData.xlsx'))
			expectOk(result)
			expect(result.value.report.features.some((f) => f.feature === 'arrayFormula')).toBe(true)
			const hasArrayMaster = result.value.workbook.sheets.some((sheet) =>
				[...sheet.cells.iterate()].some(([, , cell]) => cell.formulaInfo?.kind === 'array'),
			)
			expect(hasArrayMaster).toBe(true)
			const recalc = recalculate(result.value.workbook, defaultCalcContext())
			expect(recalc.errors).toEqual([])
		})

		it('rejects partial edits inside real POI legacy array formulas', () => {
			const result = readXlsx(loadFixture('MatrixFormulaEvalTestData.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets.find((candidate) =>
				[...candidate.cells.iterate()].some(([, , cell]) => cell.formulaInfo?.kind === 'array'),
			)
			expect(sheet).toBeDefined()
			if (!sheet) return
			const arrayCell = [...sheet.cells.iterate()].find(
				([, , cell]) => cell.formulaInfo?.kind === 'array' && cell.formulaInfo.ref,
			)
			expect(arrayCell).toBeDefined()
			const arrayRef =
				arrayCell?.[2]?.formulaInfo?.kind === 'array' ? arrayCell[2].formulaInfo.ref : undefined
			expect(arrayRef).toBeDefined()
			if (!arrayRef) return
			const range = parseRange(arrayRef)
			const targetRef = toA1(range.end)
			const before = sheet.cells.get(range.end.row, range.end.col)

			const edit = applyOperation(result.value.workbook, {
				op: 'setCells',
				sheet: sheet.name,
				updates: [{ ref: targetRef, value: 999 }],
			})
			expect(edit.ok).toBe(false)
			if (!edit.ok) expect(edit.error.message).toContain('legacy array formula')
			expect(sheet.cells.get(range.end.row, range.end.col)).toEqual(before)
		})

		it('reads hidden sheets from TwoSheetsOneHidden.xlsx', () => {
			const result = readXlsx(loadFixture('TwoSheetsOneHidden.xlsx'))
			expectOk(result)
			expect(result.value.workbook.sheets.length).toBe(2)
			expect(result.value.workbook.sheets.some((s) => s.state === 'hidden')).toBe(true)
		})

		it('reads row/column grouping from GroupTest.xlsx', () => {
			const result = readXlsx(loadFixture('GroupTest.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			const hasGroupedRows = [...sheet.rowDefs.values()].some(
				(def) => def.outlineLevel !== undefined && def.outlineLevel > 0,
			)
			const hasGroupedCols = sheet.colDefs.some(
				(def) => def.outlineLevel !== undefined && def.outlineLevel > 0,
			)
			expect(hasGroupedRows || hasGroupedCols).toBe(true)
		})

		it('reads 48495.xlsx without errors', () => {
			const result = readXlsx(loadFixture('48495.xlsx'))
			expectOk(result)
			expect(result.value.workbook.sheets.length).toBeGreaterThan(0)
		})

		it('reads theme-colored fonts from 50784-font_theme_colours.xlsx', () => {
			const result = readXlsx(loadFixture('50784-font_theme_colours.xlsx'))
			expectOk(result)
			const wb = result.value.workbook
			let hasThemeColor = false
			for (let i = 0; i < wb.styles.size; i++) {
				const style = wb.styles.get(i as never)
				if (style?.font?.color?.kind === 'theme') {
					hasThemeColor = true
					break
				}
			}
			expect(hasThemeColor).toBe(true)
		})

		it('reads indexed colors from 50786-indexed_colours.xlsx', () => {
			const result = readXlsx(loadFixture('50786-indexed_colours.xlsx'))
			expectOk(result)
			const wb = result.value.workbook
			let hasIndexedColor = false
			for (let i = 0; i < wb.styles.size; i++) {
				const style = wb.styles.get(i as never)
				const font = style?.font
				const fill = style?.fill
				if (font?.color?.kind === 'indexed' || fill?.fgColor?.kind === 'indexed') {
					hasIndexedColor = true
					break
				}
			}
			expect(hasIndexedColor).toBe(true)
		})

		it('reads formulas from FormulaEvalTestData_Copy.xlsx', () => {
			const result = readXlsx(loadFixture('FormulaEvalTestData_Copy.xlsx'))
			expectOk(result)
			const wb = result.value.workbook
			let formulaCount = 0
			for (const sheet of wb.sheets) {
				for (const [, , cell] of sheet.cells.iterate()) {
					if (cell.formula) formulaCount++
				}
			}
			expect(formulaCount).toBeGreaterThan(10)
		})

		it('keeps POI cached formula corpus semantically perfect', async () => {
			const payload = await runFormulaCorpusCorrectness({
				corpusRoot: poiDir,
				manifest: poiManifest,
				tags: ['formula-fidelity'],
				tiers: [],
				maxReportedMismatches: 20,
				sampleSeed: 1,
				oracle: 'cached-values',
				json: true,
				maxUnacceptedMismatches: 0,
				maxSemanticMismatches: 0,
				maxErrors: 0,
				minComparedFormulas: 1499,
				minSemanticPerfectWorkbooks: 22,
			})
			expect(payload.summary).toMatchObject({
				workbookCount: 22,
				formulaCount: 1499,
				comparedCount: 1499,
				noCachedFormulaCount: 0,
				volatileOracleSkipCount: 0,
				mismatchCount: 1,
				acceptedMismatchCount: 1,
				unacceptedMismatchCount: 0,
				semanticMismatchCount: 0,
				numericDriftMismatchCount: 1,
				errorCount: 0,
				perfectWorkbookCount: 21,
				semanticPerfectWorkbookCount: 22,
			})
		})

		it('reads defined names from named_ranges_2011.xlsx', () => {
			const result = readXlsx(loadFixture('named_ranges_2011.xlsx'))
			expectOk(result)
			expect(result.value.workbook.definedNames.size).toBeGreaterThan(0)
		})

		it('reads chart preservation capsules from WithChart.xlsx', () => {
			const result = readXlsx(loadFixture('WithChart.xlsx'))
			expectOk(result)
			const chartCapsules = result.value.capsules.filter(
				(c) => c.contentType?.includes('chart') || c.relType?.includes('chart'),
			)
			expect(chartCapsules.length).toBeGreaterThan(0)
		})

		it('inventories VML button drawings from BrNotClosed.xlsx', () => {
			const result = readXlsx(loadFixture('BrNotClosed.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]

			expect(sheet?.drawingObjectRefs).toEqual([
				expect.objectContaining({
					drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
					source: 'vml',
					kind: 'textBox',
					id: 1025,
					name: '_x0000_s1025',
					text: 'Multi Line Text',
					vmlShapeId: '_x0000_s1025',
					vmlObjectType: 'Button',
					anchor: {
						kind: 'twoCell',
						from: { col: 2, colOff: 14, row: 3, rowOff: 6 },
						to: { col: 4, colOff: 3, row: 8, rowOff: 0 },
					},
				}),
			])
		})

		it('preserves drawing-backed images when renaming sheets in WithDrawing.xlsx', async () => {
			const wb = await AscendWorkbook.open(loadFixture('WithDrawing.xlsx'))
			const originalSheetName = wb.sheets[0]
			expect(originalSheetName).toBeDefined()
			if (!originalSheetName) return

			const before = wb.inspect().sheets[0]
			const beforeVisuals = wb.visualInventory()
			const beforeSheetVisuals = beforeVisuals.sheets.find(
				(sheet) => sheet.sheet === originalSheetName,
			)
			expect(before?.hasDrawingRefs).toBe(true)
			expect(before?.imageCount).toBe(5)
			expect(beforeSheetVisuals?.imageRefs?.map((image) => image.targetPath)).toEqual([
				'xl/media/image1.jpeg',
				'xl/media/image2.emf',
				'xl/media/image3.png',
				'xl/media/image4.emf',
				'xl/media/image5.wmf',
			])
			expect(beforeSheetVisuals?.imageRefs?.map((image) => image.anchor?.kind)).toEqual([
				'twoCell',
				'twoCell',
				'twoCell',
				'twoCell',
				'twoCell',
			])
			expect(beforeSheetVisuals?.drawingObjectRefs).toEqual([
				expect.objectContaining({
					drawingPartPath: 'xl/drawings/drawing1.xml',
					kind: 'textBox',
					id: 9,
					name: 'TextBox 8',
					anchor: expect.objectContaining({ kind: 'oneCell' }),
				}),
			])

			const rename = wb.renameSheet(originalSheetName, `${originalSheetName} Renamed`)
			expect(rename.errors).toHaveLength(0)

			const reopened = await AscendWorkbook.open(wb.toBytes())
			const after = reopened.inspect().sheets[0]
			const afterVisuals = reopened.visualInventory().sheets[0]
			expect(after?.name).toBe(`${originalSheetName} Renamed`)
			expect(after?.hasDrawingRefs).toBe(true)
			expect(after?.imageCount).toBe(before?.imageCount)
			expect(afterVisuals?.drawingObjectCount).toBe(1)
		})
	})
} else {
	describe('POI XLSX fixtures', () => {
		it('skips when optional POI binaries are not downloaded', () => {
			expect(poiFixtures).toHaveLength(0)
		})
	})
}
