import { describe, expect, test } from 'bun:test'
import { createWorkbook } from '@ascend/core'
import type { CompatibilityReport } from '@ascend/schema'
import { makeEmbeddedChartXlsx, makeXlsx } from '../../io-xlsx/test/helpers.ts'
import { AscendWorkbook } from './index.ts'
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
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 2,
			name: 'Callout',
			text: 'Revenue up',
			relIds: ['rId2'],
			relationshipRefs: [
				{
					id: 'rId2',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
					target: 'https://example.com/report',
					targetMode: 'External',
				},
			],
			anchor: {
				kind: 'twoCell',
				from: { row: 4, col: 1 },
				to: { row: 6, col: 4 },
			},
		})
		const view = new WorkbookReadView(workbook, visualCompatibilityReport(), fullLoadInfo())

		const inventory = view.visualInventory()

		expect(inventory.sheetImageCount).toBe(1)
		expect(inventory.sheetDrawingObjectCount).toBe(1)
		expect(inventory.structuredChartCount).toBe(1)
		expect(inventory.packageChartFeatureCount).toBe(4)
		expect(inventory.packageChartSidecarFeatureCount).toBe(2)
		expect(inventory.packageDrawingFeatureCount).toBe(2)
		expect(inventory.packageMediaFeatureCount).toBe(1)
		expect(inventory.hasPreservedCharts).toBe(true)
		expect(inventory.hasPreservedChartSidecars).toBe(true)
		expect(inventory.hasPreservedDrawings).toBe(true)
		expect(inventory.hasPreservedMedia).toBe(true)
		expect(inventory.notes).toContain(
			'Chart style/color sidecars are preserved separately from chart definitions.',
		)
		expect(inventory.sheets[0]).toMatchObject({
			sheet: 'Sheet1',
			hasDrawing: true,
			hasLegacyDrawing: false,
			imageCount: 1,
			drawingObjectCount: 1,
		})
		expect(inventory.sheets[0]?.imageRefs?.[0]?.anchor?.kind).toBe('twoCell')
		expect(inventory.sheets[0]?.drawingObjectRefs?.[0]).toMatchObject({
			kind: 'textBox',
			name: 'Callout',
			text: 'Revenue up',
			relationshipRefs: [
				expect.objectContaining({
					id: 'rId2',
					target: 'https://example.com/report',
					targetMode: 'External',
				}),
			],
		})
		const returnedObject = inventory.sheets[0]?.drawingObjectRefs?.[0]
		if (returnedObject?.relationshipRefs) {
			;(returnedObject.relationshipRefs as Array<{ target: string }>)[0].target = 'mutated'
		}
		expect(sheet.drawingObjectRefs[0]?.relationshipRefs?.[0]?.target).toBe(
			'https://example.com/report',
		)
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
			'chart',
			'chart',
			'drawing',
			'image',
			'shape-or-control',
		])
	})

	test('exposes sheet-owned embedded charts without chart style package parts', async () => {
		const wb = await AscendWorkbook.open(makeEmbeddedChartXlsx({ chartType: 'lineChart' }))

		const inventory = wb.visualInventory()

		expect(inventory.structuredChartCount).toBe(1)
		expect(inventory.charts).toEqual([
			expect.objectContaining({
				partPath: 'xl/charts/chart1.xml',
				sheetName: 'Sheet1',
				chartType: 'lineChart',
			}),
		])
		expect(inventory.charts.map((chart) => chart.partPath)).not.toContain('xl/charts/style1.xml')
		expect(inventory.charts.map((chart) => chart.partPath)).not.toContain('xl/charts/colors1.xml')
	})

	test('setDrawingText edits existing text boxes through SDK save and reopen', async () => {
		const wb = await AscendWorkbook.open(textBoxWorkbook())
		expect(wb.visualInventory().sheets[0]?.drawingObjectRefs?.[0]).toMatchObject({
			kind: 'textBox',
			id: 10,
			name: 'Callout',
			text: 'Revenue up',
		})

		const result = wb.apply([
			{
				op: 'setDrawingText',
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 10,
				text: 'Revenue flat',
			},
		])
		expect(result.errors).toEqual([])

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.visualInventory().sheets[0]?.drawingObjectRefs?.[0]).toMatchObject({
			kind: 'textBox',
			id: 10,
			name: 'Callout',
			text: 'Revenue flat',
		})
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
		expect(inventory.sheetDrawingObjectCount).toBeNull()
		expect(inventory.sheets[0]?.drawingRefs).toBeNull()
		expect(inventory.sheets[0]?.imageRefs).toBeNull()
		expect(inventory.sheets[0]?.drawingObjectRefs).toBeNull()
		expect(inventory.notes).toContain('Drawing references require full sheet hydration.')
		expect(inventory.notes).toContain(
			'Image and drawing-object references require rich sheet metadata hydration.',
		)
	})
})

function visualCompatibilityReport(): CompatibilityReport {
	return {
		status: 'has-preserved',
		sourceFormat: 'xlsx',
		summary: { exact: 0, normalized: 0, preserved: 6, unsupported: 0 },
		features: [
			{
				feature: 'preservedChart',
				tier: 'preserved',
				count: 2,
				locations: ['xl/charts/chart1.xml', 'xl/charts/chart2.xml'],
			},
			{
				feature: 'preservedChartStyle',
				tier: 'preserved',
				count: 1,
				locations: ['xl/charts/style1.xml'],
			},
			{
				feature: 'preservedChartColor',
				tier: 'preserved',
				count: 1,
				locations: ['xl/charts/colors1.xml'],
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

function textBoxWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDrawing"/>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
		'xl/drawings/drawing1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="10" name="Callout" descr="Revenue note"/></xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Revenue up</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
	})
}

function fullLoadInfo(): WorkbookLoadInfo {
	return {
		mode: 'full',
		isPartial: false,
		cellsHydrated: true,
		richSheetMetadataHydrated: true,
		hasAllSheets: true,
		partialReasons: [],
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
		partialReasons: [],
		sourceSheets: ['Sheet1'],
		loadedSheets: ['Sheet1'],
	}
}
