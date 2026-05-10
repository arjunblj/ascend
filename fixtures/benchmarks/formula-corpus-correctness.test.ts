import { describe, expect, test } from 'bun:test'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createWorkbook, type StyleId } from '../../packages/core/src/index.ts'
import { dateToSerial } from '../../packages/formulas/src/functions/date.ts'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { dateValue, EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'
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

async function writeTypeScriptManifest(
	root: string,
	entries: readonly CorpusManifestEntry[],
): Promise<string> {
	const manifest = join(root, 'manifest.ts')
	await writeFile(manifest, `export default ${JSON.stringify(entries, null, 2)}\n`, 'utf-8')
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

	test('loads TypeScript corpus manifests for vendored OSS fixture sets', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		await writeFormulaWorkbook({ root, filename: 'correct.xlsx', formulaValue: 3 })
		const manifest = await writeTypeScriptManifest(root, [manifestEntry('correct.xlsx')])

		const payload = await runFormulaCorpusCorrectness({
			corpusRoot: root,
			manifest,
			tags: ['formula-fidelity'],
			tiers: ['smoke'],
			sampleSeed: 1,
			oracle: 'cached-values',
			json: true,
			minComparedFormulas: 2,
			maxMismatches: 0,
			maxErrors: 0,
		})

		expect(payload.summary).toMatchObject({
			workbookCount: 1,
			comparedCount: 2,
			mismatchCount: 0,
			errorCount: 0,
		})
		expect(payload.manifest.endsWith('manifest.ts')).toBe(true)
	})

	test('uses imported workbook calc settings when comparing cached formulas', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		workbook.calcSettings = { ...workbook.calcSettings, dateSystem: '1904' }
		const sheet = workbook.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: dateValue(17816.607951388887),
			formula: null,
			styleId: SID,
		})
		sheet.cells.set(0, 1, {
			value: stringValue('11-10-52'),
			formula: 'TEXT(A1,"d-m-y")',
			styleId: SID,
		})
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'date1904.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('date1904.xlsx')])

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
			comparedCount: 1,
			mismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 1,
		})
	})

	test('compares control-character strings using OOXML escape form', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		sheet.cells.set(0, 0, {
			value: stringValue('_x0007_7_x0007_'),
			formula: 'CHAR(7)&"7"&CHAR(7)',
			styleId: SID,
		})
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'control-string.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('control-string.xlsx')])

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
			comparedCount: 1,
			mismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 1,
		})
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

	test('skips formula cells without cached oracle values', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
		sheet.cells.set(1, 0, { value: EMPTY, formula: 'A1+2', styleId: SID })
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'no-cache.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('no-cache.xlsx')])

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
			formulaCount: 1,
			comparedCount: 0,
			noCachedFormulaCount: 1,
			mismatchCount: 0,
			errorCount: 0,
			perfectWorkbookCount: 1,
		})
	})

	test('classifies semantic, numeric drift, and volatile cached-value oracle skips', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(3.0000000001), formula: '1+2', styleId: SID })
		sheet.cells.set(1, 0, { value: numberValue(99), formula: '1+2', styleId: SID })
		sheet.cells.set(2, 0, { value: numberValue(0.1), formula: 'RAND()', styleId: SID })
		sheet.cells.set(3, 0, {
			value: stringValue('3.32192809488737-8.28846662730513E-15j'),
			formula: 'COMPLEX(3.32192809488737,-8.11707666704866E-15,"j")',
			styleId: SID,
		})
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'classified.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('classified.xlsx')])

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
			volatileOracleSkipCount: 1,
			mismatchCount: 3,
			acceptedMismatchCount: 2,
			unacceptedMismatchCount: 1,
			semanticMismatchCount: 1,
			numericDriftMismatchCount: 2,
			semanticPerfectWorkbookCount: 0,
		})
		expect(payload.results[0]?.mismatches.map((mismatch) => mismatch.classification)).toEqual([
			'numeric-drift',
			'semantic',
			'numeric-drift',
		])
		expect(payload.results[0]?.volatileOracleSkips.map((skip) => skip.ref)).toEqual(['Sheet1!A3'])
	})

	test('classifies formulas downstream of volatile precedents as oracle skips', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(0.25), formula: 'RAND()', styleId: SID })
		sheet.cells.set(1, 0, { value: numberValue(1.25), formula: 'A1+1', styleId: SID })
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'volatile-downstream.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('volatile-downstream.xlsx')])

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
			volatileOracleSkipCount: 2,
			mismatchCount: 0,
			unacceptedMismatchCount: 0,
			semanticMismatchCount: 0,
			perfectWorkbookCount: 1,
		})
		expect(payload.results[0]?.volatileOracleSkips.map((skip) => skip.ref).sort()).toEqual([
			'Sheet1!A1',
			'Sheet1!A2',
		])
	})

	test('classifies formulas downstream of numeric-drift precedents as numeric drift', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Finance')
		const values = [-10000, 3000, 4200, 6800]
		const dates = [
			dateToSerial(2013, 1, 1, '1900'),
			dateToSerial(2013, 2, 1, '1900'),
			dateToSerial(2013, 3, 1, '1900'),
			dateToSerial(2013, 4, 1, '1900'),
		]
		for (let index = 0; index < values.length; index++) {
			sheet.cells.set(9, 4 + index, {
				value: numberValue(values[index] ?? 0),
				formula: null,
				styleId: SID,
			})
			sheet.cells.set(10, 4 + index, {
				value: dateValue(dates[index] ?? 0),
				formula: null,
				styleId: SID,
			})
		}
		sheet.cells.set(10, 1, {
			value: numberValue(5.315237760543824),
			formula: 'XIRR(E10:H10,E11:H11)',
			styleId: SID,
		})
		sheet.cells.set(11, 1, {
			value: numberValue(-0.000006211546860868111),
			formula: 'XNPV(B11,E10:H10,E11:H11)',
			styleId: SID,
		})
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'downstream-drift.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('downstream-drift.xlsx')])

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
			mismatchCount: 2,
			acceptedMismatchCount: 2,
			unacceptedMismatchCount: 0,
			semanticMismatchCount: 0,
			numericDriftMismatchCount: 2,
		})
		expect(payload.results[0]?.mismatches[1]).toMatchObject({
			ref: 'Finance!B12',
			classification: 'numeric-drift',
			reason: 'downstream of numeric-drift precedent Finance!B11',
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
				maxUnacceptedMismatches: 0,
				maxSemanticMismatches: 0,
				minComparedFormulas: 2,
				minPerfectWorkbooks: 1,
				minSemanticPerfectWorkbooks: 1,
			}),
		).toEqual([])
		expect(
			formulaCorpusCorrectnessAssertionFailures(payload, {
				minComparedFormulas: 3,
				minPerfectWorkbooks: 2,
				minSemanticPerfectWorkbooks: 2,
			}),
		).toEqual([
			'compared formulas 2 below 3',
			'perfect workbooks 1 below 2',
			'semantic-perfect workbooks 1 below 2',
		])
	})

	test('assertion gates distinguish accepted numeric drift from unaccepted mismatches', async () => {
		const root = await mkdtemp(join(tmpdir(), 'ascend-formula-corpus-'))
		await mkdir(root, { recursive: true })
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(3.0000000001), formula: '1+2', styleId: SID })
		const written = writeXlsx(workbook)
		if (!written.ok) throw new Error(written.error.message)
		await writeFile(join(root, 'drift.xlsx'), written.value)
		const manifest = await writeManifest(root, [manifestEntry('drift.xlsx')])

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
			mismatchCount: 1,
			acceptedMismatchCount: 1,
			unacceptedMismatchCount: 0,
			semanticPerfectWorkbookCount: 1,
		})
		expect(
			formulaCorpusCorrectnessAssertionFailures(payload, {
				maxUnacceptedMismatches: 0,
				maxSemanticMismatches: 0,
				minSemanticPerfectWorkbooks: 1,
			}),
		).toEqual([])
		expect(
			formulaCorpusCorrectnessAssertionFailures(payload, {
				maxMismatches: 0,
			}),
		).toEqual(['mismatches 1 exceeded 0'])
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
				'--max-unaccepted-mismatches',
				'0',
				'--min-compared-formulas',
				'2',
				'--min-perfect-workbooks',
				'1',
				'--min-semantic-perfect-workbooks',
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
