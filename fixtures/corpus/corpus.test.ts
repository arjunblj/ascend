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

describe('corpus: pivot formatting metadata', () => {
	const formulasAndPivots = loadCorpusFile('ms-excel-formulas-and-pivot-tables.xlsx')
	const dashboard = loadCorpusFile('excel-dashboard-v2.xlsx')

	it.skipIf(!formulasAndPivots)(
		'exposes pivot format areas from ms-excel-formulas-and-pivot-tables.xlsx',
		() => {
			const result = readXlsx(requireBytes(formulasAndPivots))
			if (!result.ok) throw new Error(result.error.message)
			const pivot = result.value.workbook.pivotTables.find((entry) => entry.name === 'PivotTable8')
			expect(pivot?.options).toMatchObject({ dataPosition: 0, chartFormat: 1 })
			expect(pivot?.formats).toEqual([
				{
					index: 0,
					dxfId: 0,
					area: {
						outline: false,
						collapsedLevelsAreSubtotals: true,
						fieldPosition: 0,
						references: [
							{
								index: 0,
								field: 4294967294,
								itemCount: 1,
								selected: false,
								items: [{ index: 0, item: 1 }],
							},
							{ index: 1, field: 5, itemCount: 0, selected: false, items: [] },
						],
					},
				},
			])
		},
	)

	it.skipIf(!dashboard)('exposes pivot formats without explicit references', () => {
		const result = readXlsx(requireBytes(dashboard))
		if (!result.ok) throw new Error(result.error.message)
		expect(result.value.workbook.pivotCaches[0]).toMatchObject({
			cacheId: 34,
			extensionCacheId: 1332190931,
			sourceType: 'worksheet',
			sourceSheet: 'raw data',
			sourceRef: 'A1:AB1048576',
		})
		const pivot = result.value.workbook.pivotTables.find((entry) => entry.name === 'PivotTable13')
		expect(pivot?.options).toMatchObject({
			fillDownLabelsDefault: true,
			enabledSubtotalsDefault: false,
			subtotalsOnTopDefault: false,
		})
		expect(pivot?.fields[0]?.fillDownLabels).toBe(true)
		expect(pivot?.formats?.[0]).toEqual({
			index: 0,
			dxfId: 62,
			area: {
				outline: false,
				collapsedLevelsAreSubtotals: true,
				fieldPosition: 0,
			},
		})
	})

	it.skipIf(!formulasAndPivots)('exposes PivotChart format bindings', () => {
		const result = readXlsx(requireBytes(formulasAndPivots))
		if (!result.ok) throw new Error(result.error.message)
		const pivot = result.value.workbook.pivotTables.find((entry) => entry.name === 'PivotTable9')
		const expectedChartFormat = (index: number, fieldItem: number) => ({
			index,
			chart: 2,
			formatId: index,
			series: true,
			area: {
				type: 'data',
				outline: false,
				fieldPosition: 0,
				references: [
					{
						index: 0,
						field: 4294967294,
						itemCount: 1,
						selected: false,
						items: [{ index: 0, item: 0 }],
					},
					{
						index: 1,
						field: 5,
						itemCount: 1,
						selected: false,
						items: [{ index: 0, item: fieldItem }],
					},
				],
			},
		})
		expect(pivot?.chartFormats).toEqual([
			expectedChartFormat(0, 0),
			expectedChartFormat(1, 1),
			expectedChartFormat(2, 2),
		])
	})

	it.skipIf(!formulasAndPivots)('exposes pivot cache range grouping bounds', () => {
		const result = readXlsx(requireBytes(formulasAndPivots))
		if (!result.ok) throw new Error(result.error.message)
		const cache = result.value.workbook.pivotCaches.find((entry) => entry.cacheId === 0)
		const fields = cache?.fields ?? []
		expect(
			fields.find((field) => field.name === 'Date Created Conversion')?.fieldGroup,
		).toMatchObject({
			parent: 21,
			base: 18,
			range: {
				groupBy: 'months',
				startDate: '1970-01-01T00:00:00',
				endDate: '2017-03-15T15:30:07',
			},
		})
		expect(fields.find((field) => field.name === 'Quarters')?.fieldGroup).toMatchObject({
			base: 18,
			range: {
				groupBy: 'quarters',
				startDate: '1970-01-01T00:00:00',
				endDate: '2017-03-15T15:30:07',
			},
		})
		expect(fields.find((field) => field.name === 'Years')?.fieldGroup).toMatchObject({
			base: 18,
			range: {
				groupBy: 'years',
				startDate: '1970-01-01T00:00:00',
				endDate: '2017-03-15T15:30:07',
			},
		})
	})

	it.skipIf(!formulasAndPivots)(
		'summarizes real pivot cache records and sentinel shared items for cache auditing',
		() => {
			const result = readXlsx(requireBytes(formulasAndPivots))
			if (!result.ok) throw new Error(result.error.message)
			const cache = result.value.workbook.pivotCaches.find((entry) => entry.cacheId === 0)
			const cache2 = result.value.workbook.pivotCaches.find((entry) => entry.cacheId === 2)
			expect(cache?.records).toMatchObject({
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 4115,
				parsedCount: 4115,
			})
			expect(cache?.records?.preview[0]?.values.slice(0, 6)).toEqual([
				{ index: 0, kind: 'missing' },
				{ index: 1, kind: 'string', value: 'Formulas' },
				{ index: 2, kind: 'missing' },
				{ index: 3, kind: 'missing' },
				{ index: 4, kind: 'missing' },
				{ index: 5, kind: 'sharedItem', sharedItemIndex: 0 },
			])
			expect(cache?.records?.valueKindCounts).toContainEqual({ kind: 'error', count: 413 })
			expect(cache?.fields.find((field) => field.name === 'outcome')?.sharedItems?.[0]).toEqual({
				index: 0,
				kind: 'missing',
			})
			expect(
				cache?.fields.find((field) => field.name === 'Parent Category')?.sharedItems?.[0],
			).toEqual({
				index: 0,
				kind: 'error',
				value: '#VALUE!',
			})
			expect(
				cache2?.fields.find((field) => field.name === 'Sub-Category')?.sharedItems?.[0],
			).toEqual({
				index: 0,
				kind: 'error',
				value: '#VALUE!',
			})
		},
	)
})
