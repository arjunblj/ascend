import { describe, expect, it } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'

const poiDir = fileURLToPath(new URL('./poi/', import.meta.url))
const poiFixtures = readdirSync(poiDir)
	.filter((name) => name.endsWith('.xlsx'))
	.sort((a, b) => a.localeCompare(b))

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./poi/${name}`, import.meta.url))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

if (poiFixtures.length > 0) {
	describe('POI XLSX fixtures', () => {
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

		it('captures conditional formatting from ConditionalFormattingSamples.xlsx', () => {
			const result = readXlsx(loadFixture('ConditionalFormattingSamples.xlsx'))
			expectOk(result)
			const count = result.value.workbook.sheets.reduce(
				(sum, sheet) => sum + (sheet?.conditionalFormats.length ?? 0),
				0,
			)
			expect(count).toBeGreaterThan(0)
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

		it('loads multiple sheets from 55906-MultiSheetRefs.xlsx', () => {
			const result = readXlsx(loadFixture('55906-MultiSheetRefs.xlsx'))
			expectOk(result)
			expect(result.value.workbook.sheets.length).toBeGreaterThan(1)
		})

		it('captures workbook or sheet protection from POI protection fixtures', () => {
			const workbookProtected = readXlsx(
				loadFixture('workbookProtection_workbook_structure_protected.xlsx'),
			)
			expectOk(workbookProtected)
			expect(workbookProtected.value.workbook.workbookProtection).toBeDefined()

			const sheetProtected = readXlsx(loadFixture('sheetProtection_allLocked.xlsx'))
			expectOk(sheetProtected)
			expect(
				sheetProtected.value.workbook.sheets.some((sheet) => sheet?.protection !== undefined),
			).toBe(true)
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

		it('parses comment text and author from SimpleWithComments.xlsx', () => {
			const result = readXlsx(loadFixture('SimpleWithComments.xlsx'))
			expectOk(result)
			const sheet = result.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			const entries = [...sheet.comments.entries()]
			expect(entries.length).toBeGreaterThan(0)
			const [, comment] = entries[0] ?? ['', { text: '' }]
			expect(comment.text.length).toBeGreaterThan(0)
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
	})
} else {
	describe('POI XLSX fixtures', () => {
		it('skips when optional POI binaries are not downloaded', () => {
			expect(poiFixtures).toHaveLength(0)
		})
	})
}
