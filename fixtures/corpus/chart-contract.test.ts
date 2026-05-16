import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { auditXlsxPackageGraphSafeEditIntegrity, inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import { AscendWorkbook, type ChartPartInfo } from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function chartContract(chart: ChartPartInfo): unknown {
	return {
		partPath: chart.partPath,
		sheetName: chart.sheetName,
		chartType: chart.chartType,
		title: chart.title,
		series: chart.series.map((series) => ({
			nameRef: series.nameRef,
			nameText: series.nameText,
			categoryRef: series.categoryRef,
			valueRef: series.valueRef,
		})),
	}
}

function chartByPart(workbook: AscendWorkbook, partPath: string): ChartPartInfo {
	const chart = workbook.visualInventory().charts.find((entry) => entry.partPath === partPath)
	expect(chart).toBeDefined()
	if (!chart) throw new Error(`Missing chart ${partPath}`)
	return chart
}

async function openChartContract(bytes: Uint8Array, partPath: string): Promise<unknown> {
	const workbook = await AscendWorkbook.open(bytes)
	return chartContract(chartByPart(workbook, partPath))
}

describe('chart corpus contract', () => {
	test('edits a public ClosedXML worksheet chart series source with exact journal and clean reopen', async () => {
		const source = loadFixture('../xlsx/closedxml/Other_Charts_PreserveCharts_inputfile.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const before = chartContract(chartByPart(workbook, 'xl/charts/chart1.xml'))
		expect(before).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet2',
			chartType: 'lineChart',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					nameText: 'Value',
					categoryRef: 'Sheet1!$A$2:$A$8',
					valueRef: 'Sheet1!$B$2:$B$8',
				},
			],
		})

		const changed = workbook.apply(
			[
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					valueRef: 'Sheet1!$C$2:$C$8',
				},
			],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.journal).toMatchObject({
			supported: true,
			exact: true,
			issues: [],
			inverseOps: [
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					valueRef: 'Sheet1!$B$2:$B$8',
				},
			],
		})

		const edited = workbook.toBytes()
		expect(await openChartContract(edited, 'xl/charts/chart1.xml')).toMatchObject({
			...(before as Record<string, unknown>),
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					nameText: 'Value',
					categoryRef: 'Sheet1!$A$2:$A$8',
					valueRef: 'Sheet1!$C$2:$C$8',
				},
			],
		})
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
	})

	test('edits a public ExcelJS chart-sheet series source without losing chart-sheet context', async () => {
		const source = loadFixture('../xlsx/exceljs/chart-sheet.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const before = chartContract(chartByPart(workbook, 'xl/charts/chart1.xml'))
		expect(before).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Chart1',
			chartType: 'barChart',
			title: 'Wildlife Population',
			series: [
				expect.objectContaining({
					nameText: 'Bears',
					valueRef: 'Sheet1!$B$2:$B$7',
				}),
				expect.objectContaining({
					nameText: 'Dolphins',
					valueRef: 'Sheet1!$C$2:$C$7',
				}),
				expect.objectContaining({
					nameText: 'Whales',
					valueRef: 'Sheet1!$D$2:$D$7',
				}),
			],
		})

		const changed = workbook.apply(
			[
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 1,
					valueRef: 'Sheet1!$D$2:$D$7',
				},
			],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.journal).toMatchObject({
			supported: true,
			exact: true,
			issues: [],
			inverseOps: [
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 1,
					valueRef: 'Sheet1!$C$2:$C$7',
				},
			],
		})

		const edited = workbook.toBytes()
		expect(await openChartContract(edited, 'xl/charts/chart1.xml')).toMatchObject({
			...(before as Record<string, unknown>),
			series: [
				expect.objectContaining({
					nameText: 'Bears',
					valueRef: 'Sheet1!$B$2:$B$7',
				}),
				expect.objectContaining({
					nameText: 'Dolphins',
					valueRef: 'Sheet1!$D$2:$D$7',
				}),
				expect.objectContaining({
					nameText: 'Whales',
					valueRef: 'Sheet1!$D$2:$D$7',
				}),
			],
		})
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
	})

	test('fails closed when a public chart edit would add an unserializable series source field', async () => {
		const source = loadFixture('../xlsx/poi/WithChart.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const before = chartContract(chartByPart(workbook, 'xl/charts/chart1.xml'))
		expect(before).toEqual(
			expect.objectContaining({
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet2',
				series: expect.arrayContaining([
					expect.objectContaining({
						nameText: '1st Column',
						valueRef: 'Sheet1!$A$1:$A$6',
					}),
				]),
			}),
		)

		const blocked = workbook.apply(
			[
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					nameRef: 'Sheet1!$B$1',
				},
			],
			{ journal: true },
		)

		expect(blocked.errors.map((error) => error.message).join('\n')).toContain('cannot add nameRef')
		expect(blocked.journal).toMatchObject({
			supported: false,
			exact: false,
			inverseOps: [],
		})
		expect(chartContract(chartByPart(workbook, 'xl/charts/chart1.xml'))).toEqual(before)

		const saved = workbook.toBytes()
		expect(await openChartContract(saved, 'xl/charts/chart1.xml')).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(saved),
			),
		).toEqual([])
	})
})
