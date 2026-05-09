import { describe, expect, test } from 'bun:test'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createWorkbook, type StyleId } from '../../packages/core/src/index.ts'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { numberValue } from '../../packages/schema/src/index.ts'
import type { CorpusManifestEntry } from '../corpus/manifest.ts'
import {
	formulaCorpusCorrectnessAssertionFailures,
	runFormulaCorpusCorrectness,
} from './formula-corpus-correctness.ts'

const SID = 0 as StyleId
const runnerPath = fileURLToPath(new URL('./formula-corpus-correctness.ts', import.meta.url))

async function writeFormulaWorkbook(input: {
	readonly formulaValue: number
	readonly filename: string
	readonly root: string
}): Promise<void> {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')
	sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
	sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: SID })
	sheet.cells.set(2, 0, {
		value: numberValue(input.formulaValue),
		formula: 'SUM(A1:A2)',
		styleId: SID,
	})
	sheet.cells.set(3, 0, {
		value: numberValue(input.formulaValue * 2),
		formula: 'A3*2',
		styleId: SID,
	})
	const written = writeXlsx(workbook)
	if (!written.ok) throw new Error(written.error.message)
	await writeFile(join(input.root, input.filename), written.value)
}

function manifestEntry(file: string): CorpusManifestEntry {
	return {
		file,
		size_bytes: 1,
		features: {
			macros: false,
			charts: false,
			pivot_tables: false,
			tables: false,
			drawings: false,
			comments: false,
			threaded_comments: false,
			conditional_formatting: false,
			data_validations: false,
			merged_cells: false,
			hyperlinks: false,
			defined_names: false,
			external_links: false,
			connections: false,
			slicers: false,
			images_or_media: false,
			custom_xml: false,
			calc_chain: true,
		},
		counts: {
			worksheets: 1,
			formulas: 2,
		},
		source: 'generated-test-fixture',
		license: 'MIT',
		sha256: 'a'.repeat(64),
		redistributionAllowed: true,
		citation: 'Generated test fixture.',
		benchmarkTier: 'smoke',
		featureTags: ['formula-fidelity'],
	}
}

async function writeManifest(
	root: string,
	entries: readonly CorpusManifestEntry[],
): Promise<string> {
	const manifest = join(root, 'manifest.json')
	await writeFile(manifest, `${JSON.stringify(entries, null, 2)}\n`, 'utf-8')
	return manifest
}

describe('formula corpus correctness runner', () => {
	test('compares cached formula values against recalculated workbook outputs', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		await writeFormulaWorkbook({ root, filename: 'correct.xlsx', formulaValue: 3 })
		const manifest = await writeManifest(root, [manifestEntry('correct.xlsx')])

		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: root,
			manifest,
			tags: [],
			tiers: [],
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
		})

		expect(payload.summary).toMatchObject({
			workbookCount: 1,
			formulaCount: 2,
			comparedCount: 2,
			mismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 1,
		})
		expect(payload.results[0]?.beforeHash).toBe(payload.results[0]?.afterHash)
	})

	test('reports cached-value mismatches with formula references', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		await writeFormulaWorkbook({ root, filename: 'stale.xlsx', formulaValue: 99 })
		const manifest = await writeManifest(root, [manifestEntry('stale.xlsx')])

		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: root,
			manifest,
			tags: ['formula-fidelity'],
			tiers: ['smoke'],
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
		})

		expect(payload.summary.mismatchCount).toBeGreaterThan(0)
		expect(payload.results[0]?.mismatches[0]).toMatchObject({
			ref: 'Sheet1!A3',
			formula: 'SUM(A1:A2)',
			cached: 'n:99',
			calculated: 'n:3',
		})
	})

	test('assertion gates require enough real formula evidence', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		await writeFormulaWorkbook({ root, filename: 'correct.xlsx', formulaValue: 3 })
		const manifest = await writeManifest(root, [manifestEntry('correct.xlsx')])

		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: root,
			manifest,
			tags: [],
			tiers: [],
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
		})

		expect(
			formulaCorpusCorrectnessAssertionFailures(payload, {
				maxErrors: 0,
				maxMismatches: 0,
				minComparedFormulas: 2,
				minPerfectWorkbooks: 1,
			}),
		).toEqual([])
		expect(
			formulaCorpusCorrectnessAssertionFailures(payload, {
				minComparedFormulas: 3,
				minPerfectWorkbooks: 2,
			}),
		).toEqual(['compared formulas 2 below 3', 'perfect workbooks 1 below 2'])
	})

	test('CLI assertion gates accept strict zero mismatch and error thresholds', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		await writeFormulaWorkbook({ root, filename: 'correct.xlsx', formulaValue: 3 })
		const manifest = await writeManifest(root, [manifestEntry('correct.xlsx')])

		const proc = Bun.spawnSync({
			cmd: [
				Bun.argv[0],
				runnerPath,
				'--corpus-root',
				root,
				'--manifest',
				manifest,
				'--max-mismatches',
				'0',
				'--max-errors',
				'0',
				'--min-compared-formulas',
				'2',
				'--min-perfect-workbooks',
				'1',
				'--json',
			],
			stdout: 'pipe',
			stderr: 'pipe',
		})

		expect(proc.exitCode, new TextDecoder().decode(proc.stderr)).toBe(0)
		const payload = JSON.parse(new TextDecoder().decode(proc.stdout))
		expect(payload.summary).toMatchObject({
			comparedCount: 2,
			mismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 1,
		})
	})
})
