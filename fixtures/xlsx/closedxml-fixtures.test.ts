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
})
