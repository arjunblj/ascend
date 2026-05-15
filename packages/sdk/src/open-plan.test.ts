import { describe, expect, test } from 'bun:test'
import { inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import { makeXlsx } from '../../io-xlsx/test/helpers.ts'
import { inspectWorkbookOpenPlan, planInteractiveOpen, planWorkbookOpen } from './index.ts'

describe('workbook open planner', () => {
	test('recommends values mode for simple value reads', () => {
		const plan = inspectWorkbookOpenPlan(simpleWorkbook(), { intent: 'read-values' })

		expect(plan.recommendedLoadOptions).toEqual({ mode: 'values' })
		expect(plan.reviewBeforeHydration).toBe(false)
		expect(plan.worksheetPartCount).toBe(1)
		expect(plan.riskFeatures).toEqual([])
	})

	test('routes active content to metadata-only review before edit planning', () => {
		const plan = inspectWorkbookOpenPlan(activeContentWorkbook(), { intent: 'edit-plan' })

		expect(plan.recommendedLoadOptions).toEqual({ mode: 'metadata-only' })
		expect(plan.reviewBeforeHydration).toBe(true)
		expect(plan.riskFeatures).toContainEqual(
			expect.objectContaining({
				featureFamily: 'preservedMacro',
				category: 'active-content',
				sampleParts: ['xl/vbaProject.bin'],
			}),
		)
		expect(plan.reasons.join('\n')).toContain('Risk families: preservedMacro')
	})

	test('plans interactive open with trust summary and editable promotion guidance', async () => {
		const plan = await planInteractiveOpen(activeContentWorkbook(), { intent: 'edit-plan' })

		expect(plan.recommendedLoadOptions).toEqual({ mode: 'metadata-only' })
		expect(plan.previewLoadOptions).toEqual({ mode: 'metadata-only' })
		expect(plan.editableLoadOptions).toEqual({ mode: 'full' })
		expect(plan.reviewBeforeEdit).toBe(true)
		expect(plan.trustSummary).toMatchObject({
			severity: 'blocked',
			title: 'Review required',
			findingCount: expect.any(Number),
		})
		expect(plan.trustSummary.recommendedAction).toContain('metadata-only')
		expect(plan.trustSummary.topFindings).toContainEqual(
			expect.objectContaining({
				category: 'active-content',
				code: 'workbook.vbaProject',
			}),
		)
		expect(plan.steps).toEqual([
			{ id: 'plan-open', title: 'Plan open', recommended: true },
			{ id: 'open-preview', title: 'Open preview', recommended: true },
			{ id: 'review-trust', title: 'Review trust findings', recommended: true },
			{ id: 'promote-editable', title: 'Promote editable', recommended: false },
		])
	})

	test('recommends formula mode with rich metadata for dashboard planning', () => {
		const graph = inspectXlsxPackageGraph(dashboardWorkbook())
		const plan = planWorkbookOpen(graph, { intent: 'edit-plan' })

		expect(plan.recommendedLoadOptions).toEqual({ mode: 'formula', richMetadata: true })
		expect(plan.richMetadataRecommended).toBe(true)
		expect(plan.formulaSignal).toBe(true)
		expect(plan.featureSignals).toContainEqual(
			expect.objectContaining({
				featureFamily: 'preservedPivot',
				category: 'analytics',
			}),
		)
		expect(plan.featureSignals).toContainEqual(
			expect.objectContaining({
				featureFamily: 'preservedChart',
				category: 'visual',
			}),
		)
	})
})

function simpleWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
`),
		'_rels/.rels': relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
		'xl/_rels/workbook.xml.rels': relationships(`
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
`),
		'xl/workbook.xml': workbookXml('Sheet1'),
		'xl/worksheets/sheet1.xml': worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>'),
	})
}

function activeContentWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
`),
		'_rels/.rels': relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
		'xl/_rels/workbook.xml.rels': relationships(`
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
`),
		'xl/workbook.xml': workbookXml('Macro'),
		'xl/worksheets/sheet1.xml': worksheetXml(''),
		'xl/vbaProject.bin': 'macro-bytes',
	})
}

function dashboardWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
`),
		'_rels/.rels': relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
		'xl/_rels/workbook.xml.rels': relationships(`
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
`),
		'xl/workbook.xml': workbookXml('Dashboard'),
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><f>SUM(B1:B3)</f><v>6</v></c></row></sheetData>
  <drawing r:id="rIdDrawing"/>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': relationships(`
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
`),
		'xl/drawings/drawing1.xml':
			'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
		'xl/drawings/_rels/drawing1.xml.rels': relationships(`
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
`),
		'xl/charts/chart1.xml':
			'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
		'xl/pivotCache/pivotCacheDefinition1.xml': '<pivotCacheDefinition/>',
		'xl/calcChain.xml':
			'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
	})
}

function contentTypes(extra: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${extra}
</Types>`
}

function relationships(extra: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${extra}
</Relationships>`
}

function workbookXml(sheetName: string): string {
	return `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${sheetName}" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`
}

function worksheetXml(rows: string): string {
	return `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${rows}</sheetData>
</worksheet>`
}
