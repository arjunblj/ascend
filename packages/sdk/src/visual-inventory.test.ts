import { describe, expect, test } from 'bun:test'
import { createWorkbook } from '@ascend/core'
import type { CompatibilityReport } from '@ascend/schema'
import { WorkbookReadView } from './read-view.ts'
import type { WorkbookLoadInfo } from './types.ts'

describe('visual inventory', () => {
	test('summarizes package visual features and sheet image anchors', () => {
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		workbook.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'barChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
			],
		})
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: false }
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			name: 'Logo',
			description: 'Company logo',
			anchor: {
				kind: 'twoCell',
				from: { row: 0, col: 0 },
				to: { row: 3, col: 2 },
			},
		})
		const view = new WorkbookReadView(workbook, visualCompatibilityReport(), fullLoadInfo())

		const inventory = view.visualInventory()

		expect(inventory.sheetImageCount).toBe(1)
		expect(inventory.structuredChartCount).toBe(1)
		expect(inventory.packageChartFeatureCount).toBe(2)
		expect(inventory.packageDrawingFeatureCount).toBe(2)
		expect(inventory.packageMediaFeatureCount).toBe(1)
		expect(inventory.hasPreservedCharts).toBe(true)
		expect(inventory.hasPreservedDrawings).toBe(true)
		expect(inventory.hasPreservedMedia).toBe(true)
		expect(inventory.sheets[0]).toMatchObject({
			sheet: 'Sheet1',
			hasDrawing: true,
			hasLegacyDrawing: false,
			imageCount: 1,
		})
		expect(inventory.sheets[0]?.imageRefs?.[0]?.anchor?.kind).toBe('twoCell')
		expect(inventory.charts[0]).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'barChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
			],
		})
		expect(inventory.packageFeatures.map((feature) => feature.category)).toEqual([
			'chart',
			'drawing',
			'image',
			'shape-or-control',
		])
	})

	test('marks visual sheet details as unknown when metadata is not hydrated', () => {
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Sheet1')
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: true }
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: 'xl/media/image1.png',
		})
		const view = new WorkbookReadView(workbook, visualCompatibilityReport(), metadataLoadInfo())

		const inventory = view.visualInventory()

		expect(inventory.sheetImageCount).toBeNull()
		expect(inventory.sheets[0]?.drawingRefs).toBeNull()
		expect(inventory.sheets[0]?.imageRefs).toBeNull()
		expect(inventory.notes).toContain('Drawing references require full sheet hydration.')
		expect(inventory.notes).toContain('Image references require rich sheet metadata hydration.')
	})
})

function visualCompatibilityReport(): CompatibilityReport {
	return {
		status: 'has-preserved',
		sourceFormat: 'xlsx',
		summary: { exact: 0, normalized: 0, preserved: 4, unsupported: 0 },
		features: [
			{
				feature: 'preservedChart',
				tier: 'preserved',
				count: 2,
				locations: ['xl/charts/chart1.xml', 'xl/charts/chart2.xml'],
			},
			{
				feature: 'preservedDrawing',
				tier: 'preserved',
				count: 1,
				locations: ['xl/drawings/drawing1.xml'],
			},
			{
				feature: 'preservedMedia',
				tier: 'preserved',
				count: 1,
				locations: ['xl/media/image1.png'],
			},
			{
				feature: 'preservedActiveX',
				tier: 'preserved',
				count: 1,
				locations: ['xl/activeX/activeX1.xml'],
			},
		],
	}
}

function fullLoadInfo(): WorkbookLoadInfo {
	return {
		mode: 'full',
		isPartial: false,
		cellsHydrated: true,
		richSheetMetadataHydrated: true,
		hasAllSheets: true,
		sourceSheets: ['Sheet1'],
		loadedSheets: ['Sheet1'],
	}
}

function metadataLoadInfo(): WorkbookLoadInfo {
	return {
		mode: 'metadata-only',
		isPartial: false,
		cellsHydrated: false,
		richSheetMetadataHydrated: false,
		hasAllSheets: true,
		sourceSheets: ['Sheet1'],
		loadedSheets: ['Sheet1'],
	}
}
