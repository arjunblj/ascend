import { describe, expect, it, setDefaultTimeout } from 'bun:test'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { readXlsx } from '@ascend/io-xlsx'
import { AscendWorkbook } from '@ascend/sdk'
import { summarizeOoxmlPackage } from './package-summary.ts'

setDefaultTimeout(30_000)

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

function loadCorpusFile(filename: string): Uint8Array | null {
	const path = resolve(CORPUS_DIR, filename)
	if (!existsSync(path)) return null
	return new Uint8Array(readFileSync(path))
}

function requireBytes(bytes: Uint8Array | null): Uint8Array {
	if (!bytes) throw new Error('Corpus file not available')
	return bytes
}

interface CorpusEntry {
	file: string
	expectedSheets: number
	hasOpaqueFeatures: boolean
	expectedTables: number
	expectedCharts: number
	expectedPivotTables: number
	expectedDrawings: number
}

const CORPUS: CorpusEntry[] = [
	{
		file: 'bevreport-demo.xlsm',
		expectedSheets: 13,
		hasOpaqueFeatures: true,
		expectedTables: 10,
		expectedCharts: 11,
		expectedPivotTables: 0,
		expectedDrawings: 39,
	},
	{
		file: 'conditional-formatting.xlsx',
		expectedSheets: 4,
		hasOpaqueFeatures: false,
		expectedTables: 0,
		expectedCharts: 0,
		expectedPivotTables: 0,
		expectedDrawings: 0,
	},
	{
		file: 'excel-dashboard-v2.xlsx',
		expectedSheets: 3,
		hasOpaqueFeatures: true,
		expectedTables: 0,
		expectedCharts: 12,
		expectedPivotTables: 5,
		expectedDrawings: 3,
	},
	{
		file: 'large-macro-example.xlsm',
		expectedSheets: 4,
		hasOpaqueFeatures: true,
		expectedTables: 0,
		expectedCharts: 0,
		expectedPivotTables: 0,
		expectedDrawings: 1,
	},
	{
		file: 'ms-excel-formulas-and-pivot-tables.xlsx',
		expectedSheets: 7,
		hasOpaqueFeatures: true,
		expectedTables: 0,
		expectedCharts: 20,
		expectedPivotTables: 5,
		expectedDrawings: 10,
	},
]

for (const entry of CORPUS) {
	describe(`corpus: ${entry.file}`, () => {
		const bytes = loadCorpusFile(entry.file)

		it.skipIf(!bytes)('opens successfully with readXlsx', () => {
			const result = readXlsx(requireBytes(bytes))
			expect(result.ok).toBe(true)
		})

		it.skipIf(!bytes)(`has ${entry.expectedSheets} sheets`, () => {
			const result = readXlsx(requireBytes(bytes))
			if (!result.ok) throw new Error(result.error.message)
			expect(result.value.workbook.sheets).toHaveLength(entry.expectedSheets)
		})

		it.skipIf(!bytes)('has expected compatibility status', () => {
			const result = readXlsx(requireBytes(bytes))
			if (!result.ok) throw new Error(result.error.message)
			if (entry.hasOpaqueFeatures) {
				expect(result.value.report.status).toBe('has-preserved')
			} else {
				expect(['clean', 'has-preserved']).toContain(result.value.report.status)
			}
		})

		it.skipIf(!bytes || !entry.hasOpaqueFeatures)(
			'has capsules for preserved opaque features',
			() => {
				const result = readXlsx(requireBytes(bytes))
				if (!result.ok) throw new Error(result.error.message)
				expect(result.value.capsules.length).toBeGreaterThan(0)
			},
		)

		it.skipIf(!bytes)('no-op save produces byte-identical output', async () => {
			const sourceBytes = requireBytes(bytes)
			const wb = await AscendWorkbook.open(sourceBytes)
			const saved = wb.toBytes()
			expect(sha256(saved)).toBe(sha256(sourceBytes))
		})

		it.skipIf(!bytes)('reopen after save succeeds', async () => {
			const wb = await AscendWorkbook.open(requireBytes(bytes))
			const saved = wb.toBytes()
			const reopened = await AscendWorkbook.open(saved)
			expect(reopened.sheets.length).toBe(entry.expectedSheets)
		})

		if (entry.expectedTables > 0) {
			it.skipIf(!bytes)(`has ${entry.expectedTables} tables`, () => {
				const result = readXlsx(requireBytes(bytes))
				if (!result.ok) throw new Error(result.error.message)
				const totalTables = result.value.workbook.sheets.reduce(
					(sum, s) => sum + s.tables.length,
					0,
				)
				expect(totalTables).toBe(entry.expectedTables)
			})
		}

		if (entry.expectedPivotTables > 0) {
			it.skipIf(!bytes)(`has ${entry.expectedPivotTables} pivot tables`, () => {
				const result = readXlsx(requireBytes(bytes))
				if (!result.ok) throw new Error(result.error.message)
				expect(result.value.workbook.pivotTables).toHaveLength(entry.expectedPivotTables)
			})
		}

		it.skipIf(!bytes)(`has ${entry.expectedCharts} chart package parts`, () => {
			const summary = summarizeOoxmlPackage(requireBytes(bytes))
			expect(summary.families.charts).toBe(entry.expectedCharts)
		})

		it.skipIf(!bytes)(`has ${entry.expectedDrawings} drawing package parts`, () => {
			const summary = summarizeOoxmlPackage(requireBytes(bytes))
			expect(summary.families.drawings).toBe(entry.expectedDrawings)
		})

		if (entry.expectedPivotTables > 0) {
			it.skipIf(!bytes)(`has ${entry.expectedPivotTables} pivot table package parts`, () => {
				const summary = summarizeOoxmlPackage(requireBytes(bytes))
				expect(summary.families.pivotTables).toBe(entry.expectedPivotTables)
			})
		}

		describe.skipIf(!bytes)('SDK integration', () => {
			it('AscendWorkbook.open works', async () => {
				const wb = await AscendWorkbook.open(requireBytes(bytes))
				expect(wb.sheets.length).toBe(entry.expectedSheets)
			})

			it('inspect returns expected counts', async () => {
				const wb = await AscendWorkbook.open(requireBytes(bytes))
				const info = wb.inspect()
				expect(info.sheetCount).toBe(entry.expectedSheets)
				expect(info.loadedSheetCount).toBe(entry.expectedSheets)
				expect(info.pivotTableCount).toBe(entry.expectedPivotTables)
				expect(info.sourceFormat).toBe('xlsx')
				expect(info.load.mode).toBe('full')
				expect(info.load.isPartial).toBe(false)
			})

			it('SDK toBytes roundtrip succeeds', async () => {
				const wb = await AscendWorkbook.open(requireBytes(bytes))
				const saved = wb.toBytes()
				expect(saved.length).toBeGreaterThan(0)
				const reopened = await AscendWorkbook.open(saved)
				expect(reopened.sheets.length).toBe(entry.expectedSheets)
			})
		})
	})
}
