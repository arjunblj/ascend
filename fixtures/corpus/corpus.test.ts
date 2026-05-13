import { describe, expect, it, setDefaultTimeout } from 'bun:test'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { readXlsx } from '@ascend/io-xlsx'
import { AscendWorkbook } from '@ascend/sdk'
import { summarizeOoxmlPackage } from './package-summary.ts'

setDefaultTimeout(30_000)

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')
const SDK_INTEGRATION_TIMEOUT_MS = 120_000
const LARGE_SDK_INTEGRATION_FILES = new Set(['excel-dashboard-v2.xlsx'])
const CHART_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const corpusFileCache = new Map<string, Uint8Array | null>()
const readResultCache = new WeakMap<Uint8Array, ReturnType<typeof readXlsx>>()
const packageSummaryCache = new WeakMap<Uint8Array, ReturnType<typeof summarizeOoxmlPackage>>()
const sdkWorkbookCache = new WeakMap<Uint8Array, Promise<AscendWorkbook>>()
const savedSdkBytesCache = new WeakMap<Uint8Array, Promise<Uint8Array>>()
const reopenedSavedSdkWorkbookCache = new WeakMap<Uint8Array, Promise<AscendWorkbook>>()
const fullPivotCacheSdkWorkbookCache = new WeakMap<Uint8Array, Promise<AscendWorkbook>>()

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

function loadCorpusFile(filename: string): Uint8Array | null {
	if (corpusFileCache.has(filename)) return corpusFileCache.get(filename) ?? null
	const path = resolve(CORPUS_DIR, filename)
	const bytes = existsSync(path) ? new Uint8Array(readFileSync(path)) : null
	corpusFileCache.set(filename, bytes)
	return bytes
}

function requireBytes(bytes: Uint8Array | null): Uint8Array {
	if (!bytes) throw new Error('Corpus file not available')
	return bytes
}

function readCorpusFile(bytes: Uint8Array | null): ReturnType<typeof readXlsx> {
	const sourceBytes = requireBytes(bytes)
	const cached = readResultCache.get(sourceBytes)
	if (cached) return cached
	const result = readXlsx(sourceBytes)
	readResultCache.set(sourceBytes, result)
	return result
}

function readCorpusWorkbook(bytes: Uint8Array | null) {
	const result = readCorpusFile(bytes)
	if (!result.ok) throw new Error(result.error.message)
	return result.value.workbook
}

function summarizeCorpusPackage(
	bytes: Uint8Array | null,
): ReturnType<typeof summarizeOoxmlPackage> {
	const sourceBytes = requireBytes(bytes)
	const cached = packageSummaryCache.get(sourceBytes)
	if (cached) return cached
	const summary = summarizeOoxmlPackage(sourceBytes)
	packageSummaryCache.set(sourceBytes, summary)
	return summary
}

function openReadOnlySdkWorkbook(bytes: Uint8Array | null): Promise<AscendWorkbook> {
	const sourceBytes = requireBytes(bytes)
	const cached = sdkWorkbookCache.get(sourceBytes)
	if (cached) return cached
	const workbook = AscendWorkbook.open(sourceBytes)
	sdkWorkbookCache.set(sourceBytes, workbook)
	return workbook
}

function saveSdkWorkbookBytes(bytes: Uint8Array | null): Promise<Uint8Array> {
	const sourceBytes = requireBytes(bytes)
	const cached = savedSdkBytesCache.get(sourceBytes)
	if (cached) return cached
	const saved = openReadOnlySdkWorkbook(sourceBytes).then((workbook) => workbook.toBytes())
	savedSdkBytesCache.set(sourceBytes, saved)
	return saved
}

function reopenSavedSdkWorkbook(bytes: Uint8Array | null): Promise<AscendWorkbook> {
	const sourceBytes = requireBytes(bytes)
	const cached = reopenedSavedSdkWorkbookCache.get(sourceBytes)
	if (cached) return cached
	const reopened = saveSdkWorkbookBytes(sourceBytes).then((saved) => AscendWorkbook.open(saved))
	reopenedSavedSdkWorkbookCache.set(sourceBytes, reopened)
	return reopened
}

function openFullPivotCacheSdkWorkbook(bytes: Uint8Array | null): Promise<AscendWorkbook> {
	const sourceBytes = requireBytes(bytes)
	const cached = fullPivotCacheSdkWorkbookCache.get(sourceBytes)
	if (cached) return cached
	const workbook = AscendWorkbook.open(sourceBytes, {
		pivotCacheRecordMaterializeLimit: 'all',
	})
	fullPivotCacheSdkWorkbookCache.set(sourceBytes, workbook)
	return workbook
}

function expectChartRelationshipLinks(
	workbook: ReturnType<typeof readCorpusWorkbook>,
	sheetName: string,
	expectedLinks: readonly {
		readonly drawingPartPath: string
		readonly name: string
		readonly relId: string
		readonly target: string
	}[],
): void {
	const sheet = workbook.sheets.find((entry) => entry.name === sheetName)
	const chartPartPaths = new Set(workbook.chartParts.map((chart) => chart.partPath))
	const links =
		sheet?.drawingObjectRefs
			.filter((object) => object.kind === 'graphicFrame')
			.flatMap((object) =>
				(object.relationshipRefs ?? [])
					.filter((relationship) => relationship.type === CHART_RELATIONSHIP_TYPE)
					.map((relationship) => ({
						drawingPartPath: object.drawingPartPath,
						name: object.name,
						relId: relationship.id,
						target: relationship.target,
					})),
			) ?? []

	expect(links).toEqual(expectedLinks)
	for (const link of links) {
		expect(chartPartPaths.has(link.target)).toBe(true)
	}
}

function chartPart(workbook: ReturnType<typeof readCorpusWorkbook>, partPath: string) {
	const chart = workbook.chartParts.find((entry) => entry.partPath === partPath)
	if (!chart) throw new Error(`Missing chart part ${partPath}`)
	return chart
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
		const sdkIntegrationTimeout = LARGE_SDK_INTEGRATION_FILES.has(entry.file)
			? SDK_INTEGRATION_TIMEOUT_MS
			: 30_000

		it.skipIf(!bytes)('opens successfully with readXlsx', () => {
			const result = readCorpusFile(bytes)
			expect(result.ok).toBe(true)
		})

		it.skipIf(!bytes)(`has ${entry.expectedSheets} sheets`, () => {
			const result = readCorpusFile(bytes)
			if (!result.ok) throw new Error(result.error.message)
			expect(result.value.workbook.sheets).toHaveLength(entry.expectedSheets)
		})

		it.skipIf(!bytes)('has expected compatibility status', () => {
			const result = readCorpusFile(bytes)
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
				const result = readCorpusFile(bytes)
				if (!result.ok) throw new Error(result.error.message)
				expect(result.value.capsules.length).toBeGreaterThan(0)
			},
		)

		it.skipIf(!bytes)(
			'no-op save produces byte-identical output',
			async () => {
				const sourceBytes = requireBytes(bytes)
				const saved = await saveSdkWorkbookBytes(sourceBytes)
				expect(sha256(saved)).toBe(sha256(sourceBytes))
			},
			sdkIntegrationTimeout,
		)

		it.skipIf(!bytes)('reopen after save succeeds', async () => {
			const reopened = await reopenSavedSdkWorkbook(bytes)
			expect(reopened.sheets.length).toBe(entry.expectedSheets)
		})

		if (entry.expectedTables > 0) {
			it.skipIf(!bytes)(`has ${entry.expectedTables} tables`, () => {
				const result = readCorpusFile(bytes)
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
				const result = readCorpusFile(bytes)
				if (!result.ok) throw new Error(result.error.message)
				expect(result.value.workbook.pivotTables).toHaveLength(entry.expectedPivotTables)
			})
		}

		it.skipIf(!bytes)(`has ${entry.expectedCharts} chart package parts`, () => {
			const summary = summarizeCorpusPackage(bytes)
			expect(summary.families.charts).toBe(entry.expectedCharts)
		})

		it.skipIf(!bytes)(`has ${entry.expectedDrawings} drawing package parts`, () => {
			const summary = summarizeCorpusPackage(bytes)
			expect(summary.families.drawings).toBe(entry.expectedDrawings)
		})

		if (entry.expectedPivotTables > 0) {
			it.skipIf(!bytes)(`has ${entry.expectedPivotTables} pivot table package parts`, () => {
				const summary = summarizeCorpusPackage(bytes)
				expect(summary.families.pivotTables).toBe(entry.expectedPivotTables)
			})
		}

		describe.skipIf(!bytes)('SDK integration', () => {
			const testOptions = LARGE_SDK_INTEGRATION_FILES.has(entry.file)
				? { timeout: SDK_INTEGRATION_TIMEOUT_MS }
				: undefined

			it(
				'AscendWorkbook.open works',
				async () => {
					const wb = await openReadOnlySdkWorkbook(bytes)
					expect(wb.sheets.length).toBe(entry.expectedSheets)
				},
				testOptions,
			)

			it(
				'inspect returns expected counts',
				async () => {
					const wb = await openReadOnlySdkWorkbook(bytes)
					const info = wb.inspect()
					expect(info.sheetCount).toBe(entry.expectedSheets)
					expect(info.loadedSheetCount).toBe(entry.expectedSheets)
					expect(info.pivotTableCount).toBe(entry.expectedPivotTables)
					expect(info.sourceFormat).toBe('xlsx')
					expect(info.load.mode).toBe('full')
					expect(info.load.isPartial).toBe(false)
				},
				testOptions,
			)

			it(
				'SDK toBytes roundtrip succeeds',
				async () => {
					const saved = await saveSdkWorkbookBytes(bytes)
					expect(saved.length).toBeGreaterThan(0)
					const reopened = await reopenSavedSdkWorkbook(bytes)
					expect(reopened.sheets.length).toBe(entry.expectedSheets)
				},
				testOptions,
			)
		})
	})
}

describe('corpus: semantic dashboard chart and drawing inventory', () => {
	const dashboard = loadCorpusFile('excel-dashboard-v2.xlsx')
	const bevReport = loadCorpusFile('bevreport-demo.xlsm')
	const formulasAndPivots = loadCorpusFile('ms-excel-formulas-and-pivot-tables.xlsx')

	it.skipIf(!dashboard)(
		'exposes dashboard chart types, titles, series ranges, and drawing chart links',
		() => {
			const workbook = readCorpusWorkbook(dashboard)
			const chart1 = chartPart(workbook, 'xl/charts/chart1.xml')
			const chart2 = chartPart(workbook, 'xl/charts/chart2.xml')
			const chart3 = chartPart(workbook, 'xl/charts/chart3.xml')
			expect(workbook.chartParts).toHaveLength(3)
			expect(chart1).toMatchObject({
				sheetName: 'Overall Performance Summary',
				chartType: 'barChart',
			})
			expect(chart1.series[0]).toMatchObject({
				nameText: 'Revenue',
				nameRef: "'Overall Performance Summary'!$B$31:$C$31",
				valueRef: "'Overall Performance Summary'!$D$31:$O$31",
			})
			expect(chart1.series[1]).toMatchObject({
				nameText: 'Registration Fee',
				nameRef: "'Overall Performance Summary'!$B$32:$C$32",
				valueRef: "'Overall Performance Summary'!$D$32:$O$32",
			})
			expect(chart2).toMatchObject({
				sheetName: 'Overall Performance Summary',
				chartType: 'barChart',
				title: 'FY21 Total Revenue',
			})
			expect(chart2.series[0]).toMatchObject({
				nameText: 'Total Revenue',
				categoryRef: "'Overall Performance Summary'!$P$38",
				valueRef: "'Overall Performance Summary'!$P$38",
			})
			expect(chart2.series[1]).toMatchObject({
				nameText: 'Target',
				categoryRef: "'Overall Performance Summary'!$P$39",
				valueRef: "'Overall Performance Summary'!$P$39",
			})
			expect(chart3).toMatchObject({
				sheetName: 'Overall Performance Summary',
				chartType: 'scatterChart',
				title: 'Revenue',
			})
			expect(chart3.series[0]).toMatchObject({
				nameText: 'Sum of Marketing Spend',
				nameRef: "'Pivot Tables'!$AU$4",
				categoryRef: "'Pivot Tables'!$AT$5:$AT$16498",
				valueRef: "'Pivot Tables'!$AU$5:$AU$16498",
			})
			expectChartRelationshipLinks(workbook, 'Overall Performance Summary', [
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 6',
					relId: 'rId2',
					target: 'xl/charts/chart1.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 16',
					relId: 'rId3',
					target: 'xl/charts/chart2.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 17',
					relId: 'rId4',
					target: 'xl/charts/chart3.xml',
				},
			])
		},
	)

	it.skipIf(!dashboard)('exposes dashboard pivot sources used by chart ranges', () => {
		const workbook = readCorpusWorkbook(dashboard)
		expect(workbook.pivotCaches[0]).toMatchObject({
			cacheId: 34,
			extensionCacheId: 1332190931,
			sourceType: 'worksheet',
			sourceSheet: 'raw data',
			sourceRef: 'A1:AB1048576',
			recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
		})
		expect(workbook.pivotTables.find((pivot) => pivot.name === 'PivotTable14')).toMatchObject({
			sheetName: 'Pivot Tables',
			cacheId: 34,
			locationRef: 'AP4:AR16498',
			dataFields: [
				{ fieldIndex: 3, name: 'Sum of Revenue' },
				{ fieldIndex: 26, name: 'Sum of Marketing Spend' },
			],
			rowFields: [{ index: 27 }],
			columnFields: [{ index: -2 }],
		})
		expect(
			workbook.chartParts.find((chart) => chart.partPath === 'xl/charts/chart3.xml'),
		).toMatchObject({ sheetName: 'Overall Performance Summary' })
		expect(chartPart(workbook, 'xl/charts/chart3.xml').series[0]).toMatchObject({
			nameText: 'Sum of Marketing Spend',
			categoryRef: "'Pivot Tables'!$AT$5:$AT$16498",
			valueRef: "'Pivot Tables'!$AU$5:$AU$16498",
		})
	})

	it.skipIf(!dashboard)(
		'audits saved dashboard pivot outputs from full cache records',
		async () => {
			const workbook = await openFullPivotCacheSdkWorkbook(dashboard)

			expect(workbook.pivotOutputAudits()).toEqual([
				expect.objectContaining({
					pivotTable: 'PivotTable11',
					status: 'passed',
					checkedValueCount: 1,
				}),
				expect.objectContaining({
					pivotTable: 'PivotTable12',
					status: 'passed',
					checkedValueCount: 56,
				}),
				expect.objectContaining({
					pivotTable: 'PivotTable13',
					status: 'passed',
					checkedValueCount: 28,
				}),
				expect.objectContaining({
					pivotTable: 'PivotTable14',
					status: 'passed',
					checkedValueCount: 32988,
				}),
				expect.objectContaining({
					pivotTable: 'PivotTable1',
					status: 'passed',
					checkedValueCount: 91,
				}),
			])
		},
		{ timeout: 120_000 },
	)

	it.skipIf(!bevReport)(
		'exposes beverage report chart semantics and graphicFrame chart relationships',
		() => {
			const workbook = readCorpusWorkbook(bevReport)
			const chart1 = chartPart(workbook, 'xl/charts/chart1.xml')
			const chart3 = chartPart(workbook, 'xl/charts/chart3.xml')
			const chart8 = chartPart(workbook, 'xl/charts/chart8.xml')
			expect(workbook.chartParts).toHaveLength(8)
			expect(chart1).toMatchObject({
				sheetName: 'At A Glance',
				chartType: 'barChart',
				title: 'Beverage Revenue Comparison',
			})
			expect(chart1.series[0]).toMatchObject({
				nameText: 'Liquor',
				categoryRef: "('At A Glance'!$AA$2,'At A Glance'!$AE$2,'At A Glance'!$AI$2)",
				valueRef: "('At A Glance'!$AA$3,'At A Glance'!$AE$3,'At A Glance'!$AI$3)",
			})
			expect(chart1.series[1]).toMatchObject({
				nameText: 'Wine',
				valueRef: "('At A Glance'!$AA$4,'At A Glance'!$AE$4,'At A Glance'!$AI$4)",
			})
			expect(chart3).toMatchObject({
				sheetName: 'At A Glance',
				chartType: 'doughnutChart',
				title: 'Current Daily Gain in Beverage Revenue',
			})
			expect(chart3.series[0]).toMatchObject({
				nameText: 'Guage:',
				valueRef: "'At A Glance'!$AA$29:$AC$29",
			})
			expect(chart3.series[1]).toMatchObject({
				nameText: 'Needle:',
				categoryRef: "'At A Glance'!$Z$29:$Z$32",
				valueRef: "'At A Glance'!$AA$30:$AD$30",
			})
			expect(chart8).toMatchObject({
				sheetName: 'LY & Forecast',
				chartType: 'barChart',
			})
			expect(chart8.series[0]).toMatchObject({
				nameText: 'Actual Revenue',
				nameRef: "'YTD Revenue'!$D$61:$D$64",
				categoryRef: "'YTD Revenue'!$D$61:$D$64",
				valueRef: "'LY & Forecast'!$AO$52:$AO$55",
			})
			expectChartRelationshipLinks(workbook, 'At A Glance', [
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 3',
					relId: 'rId1',
					target: 'xl/charts/chart1.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 13',
					relId: 'rId2',
					target: 'xl/charts/chart2.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 6',
					relId: 'rId3',
					target: 'xl/charts/chart3.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 8',
					relId: 'rId4',
					target: 'xl/charts/chart4.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 2',
					relId: 'rId5',
					target: 'xl/charts/chart5.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 5',
					relId: 'rId6',
					target: 'xl/charts/chart6.xml',
				},
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 1',
					relId: 'rId7',
					target: 'xl/charts/chart7.xml',
				},
			])
			expectChartRelationshipLinks(workbook, 'LY & Forecast', [
				{
					drawingPartPath: 'xl/drawings/drawing4.xml',
					name: 'Chart 4',
					relId: 'rId1',
					target: 'xl/charts/chart8.xml',
				},
			])
		},
	)

	it.skipIf(!formulasAndPivots)(
		'exposes pivot workbook charts with same-sheet pivot table context',
		() => {
			const workbook = readCorpusWorkbook(formulasAndPivots)
			const chart1 = chartPart(workbook, 'xl/charts/chart1.xml')
			const chart3 = chartPart(workbook, 'xl/charts/chart3.xml')
			expect(workbook.chartParts).toHaveLength(5)
			expect(chart1).toMatchObject({
				sheetName: 'Pivot 1',
				chartType: 'barChart',
			})
			expect(chart1.series[0]).toMatchObject({
				nameText: 'successful',
				categoryRef: "'Pivot 1'!$A$5:$A$14",
				valueRef: "'Pivot 1'!$B$5:$B$14",
			})
			expect(chart3).toMatchObject({
				sheetName: 'Pivot 3',
				chartType: 'lineChart',
			})
			expect(chart3.series[0]).toMatchObject({
				nameText: 'successful',
				categoryRef: "'Pivot 3'!$A$6:$A$18",
				valueRef: "'Pivot 3'!$B$6:$B$18",
			})
			expect(workbook.pivotTables.find((pivot) => pivot.name === 'PivotTable1')).toMatchObject({
				sheetName: 'Pivot 1',
				cacheId: 0,
				locationRef: 'A3:F14',
				dataFields: [{ name: 'Count of outcome' }],
			})
			expect(workbook.pivotTables.find((pivot) => pivot.name === 'PivotTable7')).toMatchObject({
				sheetName: 'Pivot 3',
				cacheId: 0,
				locationRef: 'A4:E18',
				dataFields: [{ name: 'Count of outcome' }],
			})
			expectChartRelationshipLinks(workbook, 'Pivot 1', [
				{
					drawingPartPath: 'xl/drawings/drawing1.xml',
					name: 'Chart 1',
					relId: 'rId1',
					target: 'xl/charts/chart1.xml',
				},
			])
			expectChartRelationshipLinks(workbook, 'Pivot 3', [
				{
					drawingPartPath: 'xl/drawings/drawing3.xml',
					name: 'Chart 2',
					relId: 'rId1',
					target: 'xl/charts/chart3.xml',
				},
			])
		},
	)
})

describe('corpus: pivot formatting metadata', () => {
	const formulasAndPivots = loadCorpusFile('ms-excel-formulas-and-pivot-tables.xlsx')
	const dashboard = loadCorpusFile('excel-dashboard-v2.xlsx')

	it.skipIf(!formulasAndPivots)(
		'exposes pivot format areas from ms-excel-formulas-and-pivot-tables.xlsx',
		() => {
			const result = readCorpusFile(formulasAndPivots)
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
		const result = readCorpusFile(dashboard)
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
		const result = readCorpusFile(formulasAndPivots)
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
		const result = readCorpusFile(formulasAndPivots)
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
			const result = readCorpusFile(formulasAndPivots)
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
