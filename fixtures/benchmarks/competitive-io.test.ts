import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createWorkbook, type StyleId } from '../../packages/core/src/index.ts'
import { readXlsx, writeDenseRowsXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import {
	buildGeneratedWriteDataSet,
	buildNonRankingResult,
	buildWorkloadDataSet,
	buildWorkloadValues,
	competitorMatches,
	type DenseDataSet,
	denseWorkbookAssertions,
	denseWriteAssertions,
	evaluateAssertions,
	expectedDenseValuesHash,
	expectedOrderedWorkloadValuesHash,
	expectedSparseWideValuesHash,
	expectedStringHeavyValuesHash,
	expectedWorkloadValuesHash,
	hashOrderedLines,
	libraryAllowed,
	normalizeExternalSamples,
	parseLibraryAllowlist,
	unsupportedExternalWriteWorkloadReason,
	type WorkloadName,
	writeOperationProfile,
} from './competitive-io.ts'
import {
	extractWorkbookFeatureSummary,
	normalizeExternalRunnerSpecs,
} from './competitive-real-workbook.ts'

const SID = 0 as StyleId

type CompetitiveIoJsonPayload = {
	readonly cases: readonly {
		readonly dimensions: { readonly library: string; readonly correctnessStatus?: string }
		readonly assertions?: Record<string, unknown>
	}[]
	readonly metadata: {
		readonly skipped: readonly { readonly library: string; readonly reason?: string }[]
		readonly externalWriteRunnersByWorkload?: Record<string, readonly { readonly name: string }[]>
	}
}

async function runCompetitiveIoJson(args: readonly string[]): Promise<CompetitiveIoJsonPayload> {
	const proc = Bun.spawn(['bun', 'run', 'fixtures/benchmarks/competitive-io.ts', ...args], {
		stdout: 'pipe',
		stderr: 'pipe',
	})
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	expect(stderr).toBe('')
	expect(exitCode).toBe(0)
	return JSON.parse(stdout) as CompetitiveIoJsonPayload
}

describe('competitive IO helpers', () => {
	test('semantic write hash validation catches corrupted XLSX output', () => {
		const input = denseInput(2, 2)
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Data')
		for (let row = 0; row < input.rows; row++) {
			for (let col = 0; col < input.cols; col++) {
				sheet.cells.set(row, col, {
					value: { kind: 'number', value: row * input.cols + col },
					formula: null,
					styleId: SID,
				})
			}
		}
		sheet.cells.set(0, 0, {
			value: { kind: 'number', value: 999 },
			formula: null,
			styleId: SID,
		})

		const written = writeXlsx(workbook)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const assertions = denseWriteAssertions(written.value, input)
		const evaluated = evaluateAssertions('write', input, assertions)

		expect(assertions.reopenOk).toBe(true)
		expect(assertions.cellCountMatches).toBe(true)
		expect(assertions.semanticCellValuesHashMatches).toBe(false)
		expect(evaluated.status).toBe('semantic-mismatch')
	})

	test('string-heavy workload write assertions use string semantic hashes', () => {
		const input = workloadInput('string-heavy', 2, 5)
		const written = writeXlsx(workbookFromInput(input))
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const assertions = denseWriteAssertions(written.value, input)
		const evaluated = evaluateAssertions('write', input, assertions)

		expect(input.semanticCellValuesHash).toBe(expectedStringHeavyValuesHash(2, 5))
		expect(input.semanticCellValuesHash).not.toBe(expectedDenseValuesHash(2, 5))
		expect(assertions.reopenOk).toBe(true)
		expect(assertions.semanticCellValuesHashMatches).toBe(true)
		expect(evaluated.status).toBe('pass')
	})

	test('string-heavy workload assertions catch corrupted string cells', () => {
		const input = workloadInput('string-heavy', 2, 5)
		const workbook = workbookFromInput(input)
		const sheet = workbook.getSheet('Data')
		expect(sheet).toBeTruthy()
		sheet?.cells.set(0, 0, {
			value: { kind: 'string', value: 'sku-corrupt' },
			formula: null,
			styleId: SID,
		})

		const written = writeXlsx(workbook)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const assertions = denseWriteAssertions(written.value, input)
		const evaluated = evaluateAssertions('write', input, assertions)

		expect(assertions.cellCountMatches).toBe(true)
		expect(assertions.semanticCellValuesHashMatches).toBe(false)
		expect(evaluated.status).toBe('semantic-mismatch')
	})

	test('sparse-wide workload counts only populated cells', () => {
		const input = workloadInput('sparse-wide', 3, 101)
		const written = writeXlsx(workbookFromInput(input))
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const assertions = denseWriteAssertions(written.value, input)
		const evaluated = evaluateAssertions('write', input, assertions)

		expect(input.cells).toBe(9)
		expect(input.semanticCellValuesHash).toBe(expectedSparseWideValuesHash(3, 101))
		expect(assertions.reopenOk).toBe(true)
		expect(assertions.cellCount).toBe(9)
		expect(assertions.semanticCellValuesHashMatches).toBe(true)
		expect(evaluated.status).toBe('pass')
	})

	test('sparse-wide workload hash is distinct from dense and string workloads', () => {
		const sparseHash = expectedSparseWideValuesHash(20, 32)
		expect(sparseHash).not.toBe(expectedDenseValuesHash(20, 32))
		expect(sparseHash).not.toBe(expectedStringHeavyValuesHash(20, 32))
	})

	test('ClosedXML mixed workload encodes 10 text columns and 5 number columns', () => {
		const values = buildWorkloadValues('mixed-closedxml-10text-5number', 2, 15)
		expect(values[0]?.slice(0, 10)).toEqual(Array.from({ length: 10 }, () => 'Hello world'))
		expect(values[0]?.slice(10, 15)).toEqual([0, 1, 2, 3, 4])
		expect(values[1]?.slice(0, 10)).toEqual(Array.from({ length: 10 }, () => 'Hello world'))
		expect(values[1]?.slice(10, 15)).toEqual([0, 1, 2, 3, 4])
		expect(expectedWorkloadValuesHash('mixed-closedxml-10text-5number', 2, 15)).not.toBe(
			expectedWorkloadValuesHash('mixed-50pct-text', 2, 15),
		)
	})

	test('sparse-wide generated data set tracks logical width separately from populated cells', async () => {
		const input = await buildWorkloadDataSet('sparse-wide', 20, 32)
		expect(input.cells).toBeLessThan(input.rows * input.cols)
		expect(input.semanticCellValuesHash).toBe(expectedSparseWideValuesHash(20, 32))
	})

	test('generated-write data set skips source XLSX while preserving ordered correctness target', () => {
		const input = buildGeneratedWriteDataSet('sparse-wide', 3, 101)
		const expectedOrderedHash = expectedOrderedWorkloadValuesHash('sparse-wide', 3, 101)

		expect(input.sourceMode).toBe('generated-write')
		expect(input.xlsxBytes.byteLength).toBe(0)
		expect(input.values).toEqual([])
		expect(input.cells).toBe(9)
		expect(input.orderedSemanticCellValuesHash).toBe(expectedOrderedHash)

		const evaluated = evaluateAssertions('write', input, {
			sheetCount: 1,
			cellCount: input.cells,
			orderedSemanticCellValuesHash: expectedOrderedHash,
			reopenOk: true,
		})
		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.orderedSemanticCellValuesHashMatches).toBe(true)
	})

	test('generated-write assertions use fast worksheet hash path for dense XML output', () => {
		const input = buildGeneratedWriteDataSet('plain-text', 3, 4)
		const written = writeDenseRowsXlsx({
			rows: input.rows,
			cols: input.cols,
			omitCellRefs: true,
			stringsAreXmlSafe: true,
			valueType: 'string',
			allCellsPresent: true,
			valueAt: (row, col) => `text-${String(row * input.cols + col).padStart(8, '0')}`,
		})
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const assertions = denseWriteAssertions(written.value, input)
		const evaluated = evaluateAssertions('write', input, assertions)

		expect(assertions.fastGeneratedWriteValidation).toBe(true)
		expect(assertions.semanticCellValuesHash).toBe('__not-computed-for-fast-generated-write__')
		expect(assertions.orderedSemanticCellValuesHash).toBe(input.orderedSemanticCellValuesHash)
		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.sortedSemanticCellValuesHashMatches).toBe(false)
		expect(evaluated.assertions.orderedSemanticCellValuesHashMatches).toBe(true)
	})

	test('advanced generated workloads have distinct semantic hashes', () => {
		const hashes = new Set(
			(['styles-heavy', 'formula-heavy', 'table-heavy', 'feature-rich'] as const).map((workload) =>
				expectedWorkloadValuesHash(workload, 6, 5),
			),
		)
		expect(hashes.size).toBe(4)
		expect(hashes.has(expectedDenseValuesHash(6, 5))).toBe(false)
	})

	test('advanced write workloads report feature-specific operation profiles', () => {
		expect(writeOperationProfile('styles-heavy')).toBe('write-styles')
		expect(writeOperationProfile('formula-heavy')).toBe('write-formulas')
		expect(writeOperationProfile('table-heavy')).toBe('write-tables')
		expect(writeOperationProfile('feature-rich')).toBe('write-rich-metadata')
		expect(writeOperationProfile('dense-values')).toBe('write-values')
	})

	test('formula-heavy workload round-trips cached formula values', async () => {
		for (const readSource of ['ascend-writer', 'raw-ooxml'] as const) {
			const input = await buildWorkloadDataSet('formula-heavy', 4, 5, readSource)
			const read = readXlsx(input.xlsxBytes, { mode: 'values' })
			expect(read.ok).toBe(true)
			if (!read.ok) continue
			const evaluated = evaluateAssertions(
				'read',
				input,
				denseWorkbookAssertions(read.value.workbook, input),
			)
			expect(evaluated.status).toBe('pass')
		}
	})

	test('formula-heavy write assertions require actual worksheet formulas', async () => {
		const input = await buildWorkloadDataSet('formula-heavy', 4, 5, 'ascend-writer')
		const formulaAssertions = denseWriteAssertions(input.xlsxBytes, input)
		const valuesOnlyInput = workloadInput('formula-heavy', 4, 5)
		const valuesOnlyWritten = writeXlsx(workbookFromInput(valuesOnlyInput))
		expect(valuesOnlyWritten.ok).toBe(true)
		if (!valuesOnlyWritten.ok) return

		const valuesOnlyAssertions = denseWriteAssertions(valuesOnlyWritten.value, valuesOnlyInput)

		expect(formulaAssertions.formulaCount).toBe(12)
		expect(evaluateAssertions('write', input, formulaAssertions).status).toBe('pass')
		expect(valuesOnlyAssertions.formulaCount).toBe(0)
		expect(evaluateAssertions('write', valuesOnlyInput, valuesOnlyAssertions).status).toBe(
			'semantic-mismatch',
		)
		expect(
			evaluateAssertions('write', valuesOnlyInput, valuesOnlyAssertions).assertions
				.formulaCountMatches,
		).toBe(false)
	})

	test('table-heavy raw OOXML source carries table parts while preserving values', async () => {
		const input = await buildWorkloadDataSet('table-heavy', 5, 4, 'raw-ooxml')
		const read = readXlsx(input.xlsxBytes, { mode: 'values' })
		expect(read.ok).toBe(true)
		if (!read.ok) return
		const evaluated = evaluateAssertions(
			'read',
			input,
			denseWorkbookAssertions(read.value.workbook, input),
		)
		expect(evaluated.status).toBe('pass')
		expect(extractWorkbookFeatureSummary(input.xlsxBytes).tablePartCount).toBe(1)

		const fullRead = readXlsx(input.xlsxBytes)
		expect(fullRead.ok).toBe(true)
		if (!fullRead.ok) return
		expect(denseWorkbookAssertions(fullRead.value.workbook, input).readTableCount).toBe(1)
	})

	test('table-heavy write assertions require an actual table part', async () => {
		const input = await buildWorkloadDataSet('table-heavy', 5, 4, 'ascend-writer')
		const tableAssertions = denseWriteAssertions(input.xlsxBytes, input)
		const noTableInput = workloadInput('table-heavy', 5, 4)
		const noTableWritten = writeXlsx(workbookFromInput(noTableInput))
		expect(noTableWritten.ok).toBe(true)
		if (!noTableWritten.ok) return

		const noTableAssertions = denseWriteAssertions(noTableWritten.value, noTableInput)

		expect(tableAssertions.tablePartCount).toBe(1)
		expect(evaluateAssertions('write', input, tableAssertions).status).toBe('pass')
		expect(noTableAssertions.tablePartCount).toBe(0)
		expect(evaluateAssertions('write', noTableInput, noTableAssertions).status).toBe(
			'semantic-mismatch',
		)
		expect(
			evaluateAssertions('write', noTableInput, noTableAssertions).assertions.tablePartMatches,
		).toBe(false)
	})

	test('feature-rich write assertions require workbook feature inventory', async () => {
		const input = await buildWorkloadDataSet('feature-rich', 5, 4, 'ascend-writer')
		const rawInput = await buildWorkloadDataSet('feature-rich', 5, 4, 'raw-ooxml')
		const featureAssertions = denseWriteAssertions(input.xlsxBytes, input)
		const valuesOnlyInput = workloadInput('feature-rich', 5, 4)
		const valuesOnlyWritten = writeXlsx(workbookFromInput(valuesOnlyInput))
		expect(valuesOnlyWritten.ok).toBe(true)
		if (!valuesOnlyWritten.ok) return

		const valuesOnlyAssertions = denseWriteAssertions(valuesOnlyWritten.value, valuesOnlyInput)
		const featureSummary = extractWorkbookFeatureSummary(input.xlsxBytes)
		const featureRead = readXlsx(input.xlsxBytes)
		expect(featureRead.ok).toBe(true)
		if (!featureRead.ok) return
		const valuesRead = readXlsx(input.xlsxBytes, { mode: 'values' })
		expect(valuesRead.ok).toBe(true)
		if (!valuesRead.ok) return
		const valuesMetadataRead = readXlsx(input.xlsxBytes, {
			mode: 'values',
			richMetadata: true,
		})
		expect(valuesMetadataRead.ok).toBe(true)
		if (!valuesMetadataRead.ok) return

		expect(featureSummary.commentPartCount).toBe(1)
		expect(featureSummary.vmlDrawingPartCount).toBe(1)
		expect(featureSummary.worksheetHyperlinkCount).toBe(1)
		expect(featureSummary.worksheetDataValidationCount).toBe(1)
		expect(featureSummary.worksheetConditionalFormattingCount).toBe(1)
		expect(featureSummary.definedNameCount).toBe(1)
		expect(denseWorkbookAssertions(featureRead.value.workbook, input)).toMatchObject({
			readCommentCount: 1,
			readHyperlinkCount: 1,
			readDataValidationCount: 1,
			readConditionalFormatCount: 1,
			readDefinedNameCount: 1,
		})
		expect(denseWorkbookAssertions(valuesRead.value.workbook, input)).toMatchObject({
			readCommentCount: 0,
			readHyperlinkCount: 0,
			readDataValidationCount: 0,
			readConditionalFormatCount: 0,
			readDefinedNameCount: 1,
		})
		const valuesMetadataAssertions = denseWorkbookAssertions(
			valuesMetadataRead.value.workbook,
			input,
		)
		expect(valuesMetadataAssertions).toMatchObject({
			readCommentCount: 1,
			readHyperlinkCount: 1,
			readDataValidationCount: 1,
			readConditionalFormatCount: 1,
			readDefinedNameCount: 1,
		})
		expect(
			evaluateAssertions('read', input, denseWorkbookAssertions(valuesRead.value.workbook, input))
				.status,
		).toBe('semantic-mismatch')
		expect(evaluateAssertions('read', input, valuesMetadataAssertions).status).toBe('pass')
		expect(extractWorkbookFeatureSummary(rawInput.xlsxBytes).featureInventoryHash).toBe(
			featureSummary.featureInventoryHash,
		)
		expect(evaluateAssertions('write', input, featureAssertions).status).toBe('pass')
		expect(evaluateAssertions('write', valuesOnlyInput, valuesOnlyAssertions).status).toBe(
			'semantic-mismatch',
		)
		expect(
			evaluateAssertions('write', valuesOnlyInput, valuesOnlyAssertions).assertions
				.featureRichMatches,
		).toBe(false)
	})

	test('selected-sheet workload hydrates only the requested sheet', async () => {
		const input = await buildWorkloadDataSet('selected-sheet', 5, 4, 'raw-ooxml')
		const read = readXlsx(input.xlsxBytes, { mode: 'values', sheets: ['Data'] })
		expect(read.ok).toBe(true)
		if (!read.ok) return
		const evaluated = evaluateAssertions('read', input, {
			...denseWorkbookAssertions(read.value.workbook, input),
			selectedSheetRead: true,
			sourceSheetCount: read.value.loadInfo.sourceSheetNames.length,
			loadedSheetCount: read.value.loadInfo.loadedSheetNames.length,
			loadedSheetNames: read.value.loadInfo.loadedSheetNames.join(','),
			hasAllSheets: read.value.loadInfo.hasAllSheets,
			cellsHydrated: read.value.loadInfo.cellsHydrated,
		})
		expect(read.value.loadInfo.sourceSheetNames).toEqual(['Data', 'Summary', 'Archive'])
		expect(read.value.loadInfo.loadedSheetNames).toEqual(['Data'])
		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.selectedSheetMatches).toBe(true)
	})

	test('SheetJS selected-sheet lane parses only the requested worksheet', async () => {
		const proc = Bun.spawn(
			[
				'bun',
				'run',
				'fixtures/benchmarks/competitive-io.ts',
				'--workload',
				'selected-sheet',
				'--category',
				'read',
				'--libraries',
				'sheetjs',
				'--rows',
				'5',
				'--cols',
				'4',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--json',
			],
			{ stdout: 'pipe', stderr: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(stderr).toBe('')
		expect(exitCode).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly cases: readonly [
				{
					readonly dimensions: { readonly correctnessStatus: string }
					readonly assertions: Record<string, unknown>
				},
			]
			readonly metadata: { readonly skipped: readonly { readonly library: string }[] }
		}
		const result = payload.cases[0]
		expect(result.dimensions.correctnessStatus).toBe('pass')
		expect(result.assertions.selectedSheetRead).toBe(true)
		expect(result.assertions.sourceSheetCount).toBe(3)
		expect(result.assertions.loadedSheetCount).toBe(1)
		expect(result.assertions.loadedSheetNames).toBe('Data')
		expect(result.assertions.hasAllSheets).toBe(false)
		expect(payload.metadata.skipped.some((entry) => entry.library === 'sheetjs')).toBe(false)
	})

	test('metadata-only workload checks workbook metadata without hydrating cells', async () => {
		const input = await buildWorkloadDataSet('metadata-only', 5, 4, 'raw-ooxml')
		const read = readXlsx(input.xlsxBytes, { mode: 'metadata-only' })
		expect(read.ok).toBe(true)
		if (!read.ok) return
		const evaluated = evaluateAssertions('read', input, {
			metadataOnlyRead: true,
			sourceSheetCount: read.value.loadInfo.sourceSheetNames.length,
			loadedSheetCount: read.value.loadInfo.loadedSheetNames.length,
			loadedSheetNames: read.value.loadInfo.loadedSheetNames.join(','),
			hasAllSheets: read.value.loadInfo.hasAllSheets,
			cellsHydrated: read.value.loadInfo.cellsHydrated,
			cellCount: read.value.workbook.sheets.reduce(
				(count, sheet) => count + sheet.cells.cellCount(),
				0,
			),
		})
		expect(read.value.loadInfo.cellsHydrated).toBe(false)
		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.cellsNotHydrated).toBe(true)
	})

	test('external SheetJS runner supports metadata-only reads without hydrated cells', async () => {
		const input = await buildWorkloadDataSet('metadata-only', 5, 4, 'raw-ooxml')
		const proc = Bun.spawn(
			[
				'bun',
				'fixtures/benchmarks/runners/js_read_runner.ts',
				'--library',
				'sheetjs',
				'--metadata-only',
				'--operation',
				'read',
				'--file',
				input.xlsxPath,
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--json',
			],
			{ stdout: 'pipe', stderr: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(stderr).toBe('')
		expect(exitCode).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly assertions: Record<string, unknown>
		}
		expect(payload.assertions.metadataOnlyRead).toBe(true)
		expect(payload.assertions.sourceSheetCount).toBe(3)
		expect(payload.assertions.loadedSheetCount).toBe(3)
		expect(payload.assertions.hasAllSheets).toBe(true)
		expect(payload.assertions.cellsHydrated).toBe(false)
		expect(payload.assertions.runnerLoadMode).toBe('metadata-only')

		const evaluated = evaluateAssertions('read', input, payload.assertions)
		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.cellsNotHydrated).toBe(true)
	})

	test('raw OOXML generated data set is independent and semantically equivalent', async () => {
		const input = await buildWorkloadDataSet('sparse-wide', 4, 16, 'raw-ooxml')
		const read = readXlsx(input.xlsxBytes, { mode: 'values' })
		expect(read.ok).toBe(true)
		if (!read.ok) return
		const assertions = evaluateAssertions(
			'read',
			input,
			// Use the same assertion path as Ascend read competitors.
			denseWorkbookAssertions(read.value.workbook, input),
		)
		expect(input.readSource).toBe('raw-ooxml')
		expect(input.semanticCellValuesHash).toBe(expectedSparseWideValuesHash(4, 16))
		expect(assertions.status).toBe('pass')
	})

	test('sparse-wide workload assertions catch corrupted far-right edge cells', () => {
		const input = workloadInput('sparse-wide', 3, 101)
		const workbook = workbookFromInput(input)
		const sheet = workbook.getSheet('Data')
		expect(sheet).toBeTruthy()
		sheet?.cells.set(1, 100, {
			value: { kind: 'string', value: 'edge-corrupt' },
			formula: null,
			styleId: SID,
		})

		const written = writeXlsx(workbook)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const assertions = denseWriteAssertions(written.value, input)
		const evaluated = evaluateAssertions('write', input, assertions)

		expect(assertions.cellCountMatches).toBe(true)
		expect(assertions.semanticCellValuesHashMatches).toBe(false)
		expect(evaluated.status).toBe('semantic-mismatch')
	})

	test('external batched sample validation rejects malformed samples', () => {
		expect(() =>
			normalizeExternalSamples({ samples: [{ durationMs: 1 }, { durationMs: 2 }] }, 3, 'runner'),
		).toThrow('runner reported 2 samples but repeat requested 3')
		expect(() => normalizeExternalSamples({ samples: [{ durationMs: '1' }] }, 1, 'runner')).toThrow(
			'runner sample 0 must provide a positive durationMs',
		)
		expect(() => normalizeExternalSamples({ samples: [null] }, 1, 'runner')).toThrow(
			'runner sample 0 must be an object',
		)
	})

	test('external batched samples preserve optional memory metrics', () => {
		expect(
			normalizeExternalSamples(
				{
					samples: [
						{
							durationMs: 1,
							peakRssBytes: 1024,
							rssAfterBytes: 2048,
							rssDeltaBytes: 512,
							retainedRssDeltaBytes: 256,
							heapDeltaBytes: 128,
						},
					],
				},
				1,
				'runner',
			),
		).toEqual([
			{
				durationMs: 1,
				peakRssBytes: 1024,
				rssAfterBytes: 2048,
				rssDeltaBytes: 512,
				retainedRssDeltaBytes: 256,
				heapDeltaBytes: 128,
			},
		])
		expect(() =>
			normalizeExternalSamples({ samples: [{ durationMs: 1, peakRssBytes: -1 }] }, 1, 'runner'),
		).toThrow('External runner sample field "peakRssBytes" must be a non-negative finite number')
	})

	test('external generated read assertions are evaluated against workload semantic hashes', () => {
		const input = workloadInput('sparse-wide', 3, 101)
		const evaluated = evaluateAssertions('read', input, {
			sheetCount: 1,
			cellCount: input.cells,
			semanticCellValuesHash: input.semanticCellValuesHash,
		})

		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.sheetCountMatches).toBe(true)
		expect(evaluated.assertions.cellCountMatches).toBe(true)
		expect(evaluated.assertions.semanticCellValuesHashMatches).toBe(true)
	})

	test('external generated read assertions accept ordered semantic hashes for streaming lanes', () => {
		const input = workloadInput('mixed-50pct-text', 4, 3)
		const evaluated = evaluateAssertions('read', input, {
			sheetCount: 1,
			cellCount: input.cells,
			orderedSemanticCellValuesHash: input.orderedSemanticCellValuesHash,
		})

		expect(evaluated.status).toBe('pass')
		expect(evaluated.assertions.sortedSemanticCellValuesHashMatches).toBe(false)
		expect(evaluated.assertions.orderedSemanticCellValuesHashMatches).toBe(true)
		expect(evaluated.assertions.semanticCellValuesHashMatches).toBe(true)
	})

	test('advanced external write workloads require declared runner capabilities', () => {
		expect(
			unsupportedExternalWriteWorkloadReason(
				{ name: 'values-only', capabilities: {} },
				'formula-heavy',
			),
		).toContain('writeFormulas=true')
		expect(
			unsupportedExternalWriteWorkloadReason(
				{ name: 'formula-writer', capabilities: { writeFormulas: true } },
				'formula-heavy',
			),
		).toBeUndefined()
		expect(
			unsupportedExternalWriteWorkloadReason(
				{ name: 'formula-writer', capabilities: { writeFormulas: true } },
				'table-heavy',
			),
		).toContain('writeTables=true')
		expect(
			unsupportedExternalWriteWorkloadReason(
				{ name: 'rich-writer', capabilities: { writeRichMetadata: true } },
				'feature-rich',
			),
		).toBeUndefined()
		expect(
			unsupportedExternalWriteWorkloadReason(
				{ name: 'values-only', capabilities: {} },
				'dense-values',
			),
		).toBeUndefined()
	})

	test('external write manifest skips unsupported advanced workloads before running', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'ascend-writer-capabilities-'))
		try {
			const manifestPath = join(dir, 'writers.json')
			writeFileSync(
				manifestPath,
				JSON.stringify([
					{
						name: 'unsupported-writer',
						command: [
							'bun',
							'fixtures/benchmarks/runners/js_writer_runner.ts',
							'--library',
							'sheetjs',
						],
						categories: ['write'],
						capabilities: { internalTiming: true, valueOnlyRead: true },
					},
					{
						name: 'formula-writer',
						command: ['bun', 'fixtures/benchmarks/runners/ascend_writer_runner.ts'],
						categories: ['write'],
						capabilities: {
							internalTiming: true,
							valueOnlyRead: true,
							finalValidation: true,
							writeFormulas: true,
						},
					},
				]),
			)
			const proc = Bun.spawn(
				[
					'bun',
					'run',
					'fixtures/benchmarks/competitive-io.ts',
					'--category',
					'write',
					'--workload',
					'formula-heavy',
					'--rows',
					'5',
					'--cols',
					'4',
					'--repeat',
					'1',
					'--warmup',
					'0',
					'--libraries',
					'unsupported-writer,formula-writer',
					'--write-runner-manifest',
					manifestPath,
					'--validation-mode',
					'final',
					'--json',
				],
				{ stdout: 'pipe', stderr: 'pipe' },
			)
			const [stdout, stderr, exitCode] = await Promise.all([
				new Response(proc.stdout).text(),
				new Response(proc.stderr).text(),
				proc.exited,
			])
			expect(stderr).toBe('')
			expect(exitCode).toBe(0)
			const payload = JSON.parse(stdout) as {
				readonly cases: readonly { readonly dimensions: { readonly library: string } }[]
				readonly metadata: {
					readonly skipped: readonly { readonly library: string; readonly reason: string }[]
					readonly externalWriteRunnersByWorkload: Record<
						string,
						readonly { readonly name: string }[]
					>
				}
			}
			expect(payload.cases.map((entry) => entry.dimensions.library)).toContain('formula-writer')
			expect(payload.cases.map((entry) => entry.dimensions.library)).not.toContain(
				'unsupported-writer',
			)
			expect(payload.metadata.skipped).toContainEqual({
				library: 'unsupported-writer',
				reason:
					'unsupported write workload "formula-heavy": runner does not declare capabilities.writeFormulas=true',
			})
			expect(
				payload.metadata.externalWriteRunnersByWorkload['formula-heavy']?.map(
					(entry) => entry.name,
				),
			).toEqual(['formula-writer'])
		} finally {
			rmSync(dir, { recursive: true, force: true })
		}
	}, 15_000)

	test('advanced workload gating also suppresses built-in fallback writers', async () => {
		const payload = await runCompetitiveIoJson([
			'--category',
			'write',
			'--workload',
			'formula-heavy',
			'--rows',
			'20',
			'--cols',
			'6',
			'--repeat',
			'1',
			'--warmup',
			'0',
			'--libraries',
			'sheetjs,openpyxl,pyexcelerate,pyexcelerate-range,pyexcelerate-cell',
			'--write-runner-manifest',
			'fixtures/benchmarks/runners/sota-writers.manifest.json',
			'--validation-mode',
			'final',
			'--json',
		])
		expect(payload.cases).toHaveLength(0)
		for (const library of [
			'sheetjs',
			'openpyxl',
			'pyexcelerate',
			'pyexcelerate-range',
			'pyexcelerate-cell',
		]) {
			expect(
				payload.metadata.skipped.some(
					(entry) =>
						entry.library === library && entry.reason?.includes('capabilities.writeFormulas=true'),
				),
			).toBe(true)
		}
	})

	test('capable in-process rich metadata writers still run when manifest adapters skip', async () => {
		const payload = await runCompetitiveIoJson([
			'--category',
			'write',
			'--workload',
			'feature-rich',
			'--rows',
			'20',
			'--cols',
			'6',
			'--repeat',
			'1',
			'--warmup',
			'0',
			'--libraries',
			'exceljs',
			'--write-runner-manifest',
			'fixtures/benchmarks/runners/sota-writers.manifest.json',
			'--validation-mode',
			'final',
			'--json',
		])
		expect(payload.cases.map((entry) => entry.dimensions.library)).toEqual(['exceljs'])
		expect(payload.cases[0]?.dimensions.correctnessStatus).toBe('pass')
		expect(payload.cases[0]?.assertions?.featureRichMatches).toBe(true)
	})

	test('external competitor selector includes non-JS SOTA runners', () => {
		expect(competitorMatches('sheetjs', 'external')).toBe(false)
		expect(competitorMatches('fastexcel', 'external')).toBe(true)
		expect(competitorMatches('fastexcel-java', 'external')).toBe(true)
		expect(competitorMatches('fastxlsx', 'external')).toBe(true)
		expect(competitorMatches('pyopenxlsx', 'external')).toBe(true)
		expect(competitorMatches('pyopenxlsx-bulk', 'external')).toBe(true)
		expect(competitorMatches('pyfastexcel', 'external')).toBe(true)
		expect(competitorMatches('pyexcelerate-range', 'external')).toBe(true)
		expect(competitorMatches('pyexcelerate-cell', 'external')).toBe(true)
		expect(competitorMatches('rust-calamine', 'external')).toBe(true)
		expect(competitorMatches('polars-calamine', 'external')).toBe(true)
		expect(competitorMatches('polars-xlsx2csv', 'external')).toBe(true)
		expect(competitorMatches('polars-openpyxl', 'external')).toBe(true)
		expect(competitorMatches('apache-poi', 'external')).toBe(true)
		expect(competitorMatches('excelize', 'external')).toBe(true)
		expect(competitorMatches('npoi', 'external')).toBe(true)
		expect(competitorMatches('closedxml', 'python')).toBe(true)
		expect(competitorMatches('ascend-external-values', 'external')).toBe(true)
		expect(competitorMatches('ascend-readxlsx-raw-values-bytes', 'external')).toBe(true)
		expect(competitorMatches('ascend-readxlsx-raw-values-operation-bytes', 'external')).toBe(true)
		expect(competitorMatches('ascend-readxlsx-raw-values-operation-path', 'external')).toBe(true)
		expect(competitorMatches('ascend-readxlsx-row-stream-bytes', 'external')).toBe(true)
		expect(competitorMatches('ascend-external-writer', 'external')).toBe(true)
		expect(competitorMatches('rust-xlsxwriter', 'external')).toBe(true)
	})

	test('library allowlist applies exact runner names after competitor selection', () => {
		const allowlist = parseLibraryAllowlist('ascend-external-writer, excelize ')
		expect(libraryAllowed('ascend-external-writer', allowlist)).toBe(true)
		expect(libraryAllowed('excelize', allowlist)).toBe(true)
		expect(libraryAllowed('xlsxwriter', allowlist)).toBe(false)
		expect(libraryAllowed('xlsxwriter', parseLibraryAllowlist(undefined))).toBe(true)
		expect(parseLibraryAllowlist(' , ')).toBeUndefined()
	})

	test('library allowlist expands the natural Ascend alias across read and write runners', () => {
		const allowlist = parseLibraryAllowlist('ascend')
		expect(libraryAllowed('ascend-external-writer', allowlist)).toBe(true)
		expect(libraryAllowed('ascend-readxlsx-row-stream-bytes', allowlist)).toBe(true)
		expect(libraryAllowed('ascend', allowlist)).toBe(true)
		expect(libraryAllowed('rust-xlsxwriter', allowlist)).toBe(false)
	})

	test('SOTA writer manifest includes Python writer baselines required by scoreboard coverage', () => {
		const manifest = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/sota-writers.manifest.json', 'utf-8'),
		) as unknown
		const specs = normalizeExternalRunnerSpecs(manifest)
		const names = new Set(specs.map((spec) => spec.name))
		for (const name of [
			'ascend-external-writer',
			'xlsxwriter',
			'xlsxwriter-constant-memory',
			'pyexcelerate',
			'pyexcelerate-range',
			'pyexcelerate-cell',
			'openpyxl',
			'openpyxl-write-only',
			'rust-xlsxwriter',
			'fastxlsx',
			'pyopenxlsx',
			'pyopenxlsx-bulk',
			'pyfastexcel',
		]) {
			expect(names.has(name)).toBe(true)
		}
		expect(
			specs.find((spec) => spec.name === 'ascend-external-writer')?.capabilities,
		).toMatchObject({
			writeFormulas: true,
			writeTables: true,
			writeRichMetadata: true,
		})
		expect(specs.find((spec) => spec.name === 'sheetjs')?.capabilities?.writeFormulas).toBe(
			undefined,
		)
		expect(specs.find((spec) => spec.name === 'xlsxwriter')?.capabilities).toMatchObject({
			writeFormulas: true,
			writeTables: true,
			writeRichMetadata: true,
		})
		expect(specs.find((spec) => spec.name === 'openpyxl')?.capabilities).toMatchObject({
			writeTables: true,
			writeRichMetadata: true,
		})
		expect(specs.find((spec) => spec.name === 'rust-xlsxwriter')?.capabilities).toMatchObject({
			writeFormulas: true,
			writeTables: true,
		})
	})

	test('SOTA reader manifest includes Ascend materialized read lane for pyopenxlsx comparisons', () => {
		const manifest = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/ascend-python-readers.manifest.json', 'utf-8'),
		) as unknown
		const specs = normalizeExternalRunnerSpecs(manifest)
		const materialized = specs.find(
			(spec) => spec.name === 'ascend-readxlsx-cell-materialization-bytes',
		)
		expect(materialized?.timingModel).toBe('external-internal-cell-materialization-timing')
		expect(materialized?.capabilities?.valueOnlyRead).toBe(true)
		expect(competitorMatches('ascend-readxlsx-cell-materialization-bytes', 'external')).toBe(true)
		const openOnly = specs.find((spec) => spec.name === 'ascend-readxlsx-raw-values-operation-path')
		expect(openOnly?.command).toContain(
			'fixtures/benchmarks/runners/ascend_readxlsx_open_runner.ts',
		)
		const bytesOnly = specs.find(
			(spec) => spec.name === 'ascend-readxlsx-raw-values-operation-bytes',
		)
		expect(bytesOnly?.command).toContain(
			'fixtures/benchmarks/runners/ascend_readxlsx_open_runner.ts',
		)
		const stream = specs.find((spec) => spec.name === 'ascend-readxlsx-row-stream-bytes')
		expect(stream?.timingModel).toBe('external-internal-row-stream-timing')
		expect(stream?.command).toContain(
			'fixtures/benchmarks/runners/ascend_readxlsx_stream_runner.ts',
		)
		const fastExcelJava = specs.find((spec) => spec.name === 'fastexcel-java')
		expect(fastExcelJava?.timingModel).toBe('external-internal-streaming-values-timing')
		expect(fastExcelJava?.capabilities?.valueOnlyRead).toBe(true)
		expect(fastExcelJava?.capabilities?.finalValidation).toBe(true)
	})

	test('external generated runner failures become non-ranking benchmark rows', () => {
		const input = workloadInput('sparse-wide', 3, 101)
		const result = buildNonRankingResult({
			name: 'broken-reader:xlsx-read-sparse-wide',
			library: 'broken-reader',
			category: 'read',
			executionScope: 'external-process',
			runnerProvenance: {
				adapterVersion: '1',
				libraryVersion: '2.3.4',
				runtime: 'python3',
			},
			timingModel: 'external-internal-file-path-materialization-timing',
			validationModel: 'external-post-operation-assertions',
			memoryModel: 'peak-rss-reported',
			dataSet: input,
			repeat: 3,
			status: 'error',
			reason: 'runner exploded',
		})

		expect(result.dimensions.correctnessStatus).toBe('error')
		expect(result.dimensions.rankingEligible).toBe(false)
		expect(result.dimensions.errorReason).toBe('runner exploded')
		expect(result.dimensions.runnerAdapterVersion).toBe('1')
		expect(result.dimensions.runnerManifestLibraryVersion).toBe('2.3.4')
		expect(result.dimensions.runnerRuntime).toBe('python3')
		expect(result.dimensions.timingLane).toBe(
			'external-internal-file-path-materialization-timing:sparse-wide',
		)
		expect(result.dimensions.timingModel).toBe('external-internal-file-path-materialization-timing')
		expect(result.dimensions.validationModel).toBe('external-post-operation-assertions')
		expect(result.dimensions.memoryModel).toBe('peak-rss-reported')
		expect(result.assertions?.errorReason).toBe('runner exploded')
	})
})

function denseInput(rows: number, cols: number): DenseDataSet {
	return workloadInput('dense-values', rows, cols)
}

function workloadInput(workloadName: WorkloadName, rows: number, cols: number): DenseDataSet {
	const values = buildWorkloadValues(workloadName, rows, cols)
	const semanticCellValuesHash =
		workloadName === 'string-heavy'
			? expectedStringHeavyValuesHash(rows, cols)
			: workloadName === 'sparse-wide'
				? expectedSparseWideValuesHash(rows, cols)
				: workloadName === 'dense-values'
					? expectedDenseValuesHash(rows, cols)
					: expectedWorkloadValuesHash(workloadName, rows, cols)
	return {
		workloadName,
		readSource: 'ascend-writer',
		rows,
		cols,
		cells: values.reduce((count, row) => count + row.filter((value) => value !== null).length, 0),
		values,
		semanticCellValuesHash,
		orderedSemanticCellValuesHash: hashOrderedLines(
			values.flatMap((row, rowIndex) =>
				row.flatMap((value, colIndex) =>
					value === null
						? []
						: [
								`Data!${String.fromCharCode(65 + colIndex)}${rowIndex + 1}\t${scalarPayload(value)}`,
							],
				),
			),
		),
		xlsxPath: '',
		xlsxBytes: new Uint8Array(),
	}
}

function workbookFromInput(input: DenseDataSet) {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Data')
	for (let row = 0; row < input.rows; row++) {
		const sourceRow = input.values[row]
		if (!sourceRow) continue
		for (let col = 0; col < input.cols; col++) {
			const value = sourceRow[col]
			if (value === undefined || value === null) continue
			sheet.cells.set(row, col, {
				value:
					typeof value === 'number'
						? { kind: 'number', value }
						: typeof value === 'boolean'
							? { kind: 'boolean', value }
							: { kind: 'string', value },
				formula: null,
				styleId: SID,
			})
		}
	}
	return workbook
}

function scalarPayload(value: unknown): string | null {
	if (typeof value === 'number') return `n:${value}`
	if (typeof value === 'string') return `s:${value}`
	if (typeof value === 'boolean') return `b:${value ? 'true' : 'false'}`
	return null
}
