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
		expect(entries.length).toBe(22)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['libreoffice'] })).toHaveLength(entries.length)
		expect(selectManifestEntries(entries, { tags: ['formula-fidelity'] })).toHaveLength(8)
		expect(selectManifestEntries(entries, { tags: ['pivot-table'] })).toHaveLength(2)
		expect(selectManifestEntries(entries, { tags: ['table'] })).toHaveLength(5)
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
			anchor: {
				kind: 'twoCell',
				from: { col: 2, row: 3, colOff: 133350, rowOff: 152400 },
				to: { col: 10, row: 9, colOff: 28575, rowOff: 85725 },
			},
		})
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
			minComparedFormulas: 107,
			minSemanticPerfectWorkbooks: 8,
		})
		expect(payload.summary).toMatchObject({
			workbookCount: 8,
			formulaCount: 107,
			comparedCount: 107,
			mismatchCount: 0,
			acceptedMismatchCount: 0,
			unacceptedMismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 8,
			semanticPerfectWorkbookCount: 8,
		})
	})
})
