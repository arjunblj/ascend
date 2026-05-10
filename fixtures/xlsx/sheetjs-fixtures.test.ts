import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import { runFormulaCorpusCorrectness } from '../benchmarks/formula-corpus-correctness.ts'
import {
	normalizeManifest,
	selectManifestEntries,
	validateManifestProvenance,
} from '../corpus/manifest.ts'
import { loadManifest } from './sheetjs/manifest.ts'

const sheetJsDir = fileURLToPath(new URL('./sheetjs/', import.meta.url))
const sheetJsManifest = fileURLToPath(new URL('./sheetjs/manifest.ts', import.meta.url))

function loadFixture(name: string): Uint8Array {
	return readFileSync(resolve(sheetJsDir, name))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('SheetJS XLSX fixture corpus', () => {
	test('manifest has pinned provenance for the vendored Apache-2.0 fixture subset', async () => {
		const entries = normalizeManifest(await loadManifest())
		if (entries.length === 0) return
		expect(entries.map((entry) => entry.file).sort()).toEqual([
			'../poi/AutoFilter.xlsx',
			'../poi/formula_stress_test.xlsx',
			'../poi/merge_cells.xlsx',
			'../poi/named_ranges_2011.xlsx',
		])
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['sheetjs'] })).toHaveLength(entries.length)
		expect(
			selectManifestEntries(entries, { tags: ['formula-fidelity'] }).map((entry) => entry.file),
		).toEqual(['../poi/formula_stress_test.xlsx', '../poi/named_ranges_2011.xlsx'])
		expect(
			selectManifestEntries(entries, { tags: ['merged-cells'] }).map((entry) => entry.file),
		).toEqual(['../poi/merge_cells.xlsx'])
		expect(
			selectManifestEntries(entries, { tags: ['defined-names'] }).map((entry) => entry.file),
		).toEqual(['../poi/AutoFilter.xlsx', '../poi/named_ranges_2011.xlsx'])
	})

	test('opens and round-trips every SheetJS fixture without crashing', async () => {
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

	test('captures expected feature families from SheetJS regression fixtures', async () => {
		const entries = normalizeManifest(await loadManifest())
		if (entries.length === 0) return
		expect(
			entries.find((entry) => entry.file === '../poi/AutoFilter.xlsx')?.features.defined_names,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === '../poi/merge_cells.xlsx')?.features.merged_cells,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === '../poi/named_ranges_2011.xlsx')?.features
				.defined_names,
		).toBe(true)
		expect(
			entries.find((entry) => entry.file === '../poi/formula_stress_test.xlsx')?.features
				.calc_chain,
		).toBe(true)
	})

	test('cached formulas in the SheetJS subset have no semantic mismatches', async () => {
		if ((await loadManifest()).length === 0) return
		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: sheetJsDir,
			manifest: sheetJsManifest,
			tags: ['formula-fidelity'],
			tiers: [],
			maxReportedMismatches: 50,
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
			maxUnacceptedMismatches: 0,
			maxSemanticMismatches: 0,
			maxErrors: 0,
			minComparedFormulas: 445,
			minSemanticPerfectWorkbooks: 2,
		})
		expect(payload.summary).toMatchObject({
			workbookCount: 2,
			formulaCount: 445,
			comparedCount: 445,
			volatileOracleSkipCount: 3,
			mismatchCount: 31,
			acceptedMismatchCount: 31,
			unacceptedMismatchCount: 0,
			semanticMismatchCount: 0,
			numericDriftMismatchCount: 31,
			errorCount: 0,
			perfectWorkbookCount: 1,
			semanticPerfectWorkbookCount: 2,
		})
	})
})
