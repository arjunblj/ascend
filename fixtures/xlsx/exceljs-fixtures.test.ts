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
import { loadManifest } from './exceljs/manifest.ts'

const excelJsDir = fileURLToPath(new URL('./exceljs/', import.meta.url))
const excelJsManifest = fileURLToPath(new URL('./exceljs/manifest.ts', import.meta.url))

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./exceljs/${name}`, import.meta.url))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('ExcelJS XLSX fixture corpus', () => {
	test('manifest has pinned provenance for the vendored MIT fixture subset', async () => {
		expect(existsSync(new URL('./exceljs/LICENSE', import.meta.url))).toBe(true)
		const entries = normalizeManifest(await loadManifest())
		expect(entries.length).toBe(20)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['exceljs'] })).toHaveLength(entries.length)
		expect(selectManifestEntries(entries, { tags: ['formula-fidelity'] }).length).toBeGreaterThan(0)
		expect(selectManifestEntries(entries, { tags: ['issue-regression'] }).length).toBeGreaterThan(0)
		expect(selectManifestEntries(entries, { tags: ['chart'] }).map((entry) => entry.file)).toEqual([
			'chart-sheet.xlsx',
		])
	})

	test('opens and round-trips every ExcelJS fixture without crashing', async () => {
		const entries = await loadManifest()
		for (const entry of entries) {
			const initial = readXlsx(loadFixture(entry.file))
			expectOk(initial)
			expect(initial.value.workbook.sheets.length).toBeGreaterThan(0)
			const written = writeXlsx(initial.value.workbook, initial.value.capsules)
			expectOk(written)
			const reopened = readXlsx(written.value)
			expectOk(reopened)
			expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
		}
	})

	test('captures the expected feature families from ExcelJS issue fixtures', async () => {
		const entries = normalizeManifest(await loadManifest())
		expect(
			entries.find((entry) => entry.file === 'bogus-defined-name.xlsx')?.features.defined_names,
		).toBe(true)
		expect(entries.find((entry) => entry.file === 'test-issue-1669.xlsx')?.features.tables).toBe(
			true,
		)
		expect(
			entries.find((entry) => entry.file === 'test-issue-1842.xlsx')?.features.data_validations,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === 'test-issue-623.xlsx')?.features.merged_cells,
		).toBe(true)
	})

	test('links drawing-mediated chartsheet charts from chart-sheet.xlsx', () => {
		const result = readXlsx(loadFixture('chart-sheet.xlsx'))
		expectOk(result)

		expect(result.value.workbook.chartSheets).toEqual([
			{
				name: 'Chart1',
				sheetId: '9',
				relId: 'rId1',
				partPath: 'xl/chartsheets/sheet1.xml',
				state: 'visible',
				chartPartPaths: ['xl/charts/chart1.xml'],
			},
		])
		expect(result.value.workbook.chartParts[0]).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Chart1',
			chartType: 'barChart',
			title: 'Wildlife Population',
		})
	})

	test('cached formulas in the ExcelJS subset recalculate without mismatches', async () => {
		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: excelJsDir,
			manifest: excelJsManifest,
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
			minComparedFormulas: 9,
			minSemanticPerfectWorkbooks: 4,
		})
		expect(payload.summary).toMatchObject({
			workbookCount: 4,
			formulaCount: 9,
			comparedCount: 9,
			mismatchCount: 0,
			acceptedMismatchCount: 0,
			unacceptedMismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 4,
			semanticPerfectWorkbookCount: 4,
		})
	})
})
