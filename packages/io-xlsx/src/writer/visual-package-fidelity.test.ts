import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { makeXlsx } from '../../test/helpers.ts'
import { inspectXlsxPackageGraph } from '../package-graph.ts'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
} from '../package-graph-fidelity.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('visual package fidelity', () => {
	test('preserves DrawingML, VML, media, chart sidecars, and chart ownership after a dirty cell edit', () => {
		const sourceBytes = visualOwnershipWorkbook()
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		expect(auditXlsxPackageGraphReadIntegrity(beforeGraph)).toEqual([])

		const opened = readXlsx(sourceBytes)
		expectOk(opened)

		const visualSheet = opened.value.workbook.sheets.find((sheet) => sheet.name === 'Visuals')
		expect(visualSheet?.imageRefs).toEqual([
			expect.objectContaining({
				drawingPartPath: 'xl/drawings/drawing1.xml',
				relId: 'rIdImageMain',
				targetPath: 'xl/media/image 1.png',
				name: 'Anchored image',
				anchor: expect.objectContaining({
					kind: 'twoCell',
					editAs: 'oneCell',
					from: expect.objectContaining({ row: 1, col: 1 }),
					to: expect.objectContaining({ row: 4, col: 3 }),
				}),
			}),
		])
		expect(visualSheet?.drawingObjectRefs).toContainEqual(
			expect.objectContaining({
				source: 'drawingml',
				kind: 'graphicFrame',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				name: 'Embedded chart frame',
				relIds: ['rIdChartEmbedded'],
				relationshipRefs: [
					expect.objectContaining({
						id: 'rIdChartEmbedded',
						target: 'xl/charts/chart1.xml',
					}),
				],
			}),
		)
		expect(visualSheet?.drawingObjectRefs).toContainEqual(
			expect.objectContaining({
				source: 'vml',
				kind: 'shape',
				drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
				vmlObjectType: 'Button',
				anchor: expect.objectContaining({ kind: 'twoCell' }),
				relationshipRefs: [
					expect.objectContaining({
						id: 'rIdVmlImage',
						target: 'xl/media/vmlImage.emf',
					}),
				],
			}),
		)

		expect(
			opened.value.workbook.chartParts
				.map((chart) => ({
					partPath: chart.partPath,
					sheetName: chart.sheetName,
					chartType: chart.chartType,
				}))
				.sort((left, right) => left.partPath.localeCompare(right.partPath)),
		).toEqual([
			{ partPath: 'xl/charts/chart1.xml', sheetName: 'Visuals', chartType: 'barChart' },
			{ partPath: 'xl/charts/chart2.xml', sheetName: 'Chart Sheet', chartType: 'lineChart' },
		])
		expect(opened.value.workbook.chartSheets).toEqual([
			expect.objectContaining({
				name: 'Chart Sheet',
				partPath: 'xl/chartsheets/sheet1.xml',
				chartPartPaths: ['xl/charts/chart2.xml'],
			}),
		])

		expect(beforeGraph.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/drawings/drawing1.xml',
				ownerScope: 'drawing',
				featureFamily: 'preservedDrawing',
			}),
		)
		expect(beforeGraph.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/drawings/vmlDrawing1.vml',
				ownerScope: 'drawing',
				featureFamily: 'preservedVml',
			}),
		)
		expect(beforeGraph.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/charts/style1.xml',
				ownerScope: 'chart',
				featureFamily: 'preservedChartStyle',
			}),
		)
		expect(beforeGraph.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/charts/colors1.xml',
				ownerScope: 'chart',
				featureFamily: 'preservedChartColor',
			}),
		)

		if (!visualSheet) throw new Error('Visuals sheet was not parsed')
		visualSheet.cells.set(0, 0, {
			value: stringValue('Edited without moving visuals'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: ['Visuals'],
		})
		expectOk(written)

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph)).toEqual([])
		expect(auditXlsxPackageGraphBytePreservation(beforeGraph, sourceBytes, written.value)).toEqual(
			[],
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedVisualSheet = reopened.value.workbook.sheets.find(
			(sheet) => sheet.name === 'Visuals',
		)
		expect(reopenedVisualSheet?.imageRefs).toEqual(visualSheet.imageRefs)
		expect(reopenedVisualSheet?.drawingObjectRefs).toEqual(visualSheet.drawingObjectRefs)
		expect(reopened.value.workbook.chartSheets).toEqual(opened.value.workbook.chartSheets)
		expect(
			reopened.value.workbook.chartParts
				.map((chart) => ({
					partPath: chart.partPath,
					sheetName: chart.sheetName,
					chartType: chart.chartType,
				}))
				.sort((left, right) => left.partPath.localeCompare(right.partPath)),
		).toEqual([
			{ partPath: 'xl/charts/chart1.xml', sheetName: 'Visuals', chartType: 'barChart' },
			{ partPath: 'xl/charts/chart2.xml', sheetName: 'Chart Sheet', chartType: 'lineChart' },
		])
	})

	test('keeps chart style and color overrides when default xml would otherwise cover them', () => {
		const sourceBytes = defaultCoveredChartSidecarWorkbook()
		const opened = readXlsx(sourceBytes)
		expectOk(opened)

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: ['Visuals'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		expect(contentTypes).toContain('PartName="/xl/charts/style1.xml" ContentType="application/xml"')
		expect(contentTypes).toContain(
			'PartName="/xl/charts/colors1.xml" ContentType="application/xml"',
		)

		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		const afterGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph)).toEqual([])
		expect(auditXlsxPackageGraphBytePreservation(beforeGraph, sourceBytes, written.value)).toEqual(
			[],
		)
	})
})

function visualOwnershipWorkbook(
	options: { readonly defaultCoveredChartSidecars?: boolean } = {},
): Uint8Array {
	const chartSidecarContentType = options.defaultCoveredChartSidecars
		? 'application/xml'
		: 'application/vnd.ms-office.chartstyle+xml'
	const chartColorContentType = options.defaultCoveredChartSidecars
		? 'application/xml'
		: 'application/vnd.ms-office.chartcolorstyle+xml'
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/vmlDrawing1.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/style1.xml" ContentType="${chartSidecarContentType}"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="${chartColorContentType}"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdChartSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Visuals" sheetId="1" r:id="rIdData"/>
    <sheet name="Chart Sheet" sheetId="2" r:id="rIdChartSheet"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1" t="str"><v>Original</v></c></row></sheetData>
  <drawing r:id="rIdDrawingMain"/>
  <legacyDrawing r:id="rIdVmlMain"/>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawingMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rIdVmlMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`,
		'xl/chartsheets/sheet1.xml': `<?xml version="1.0"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <drawing r:id="rIdChartSheetChart"/>
</chartsheet>`,
		'xl/chartsheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChartSheetChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`,
		'xl/drawings/drawing1.xml': `<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Anchored image" descr="Image with escaped target"/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rIdImageMain"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>4</xdr:col><xdr:row>1</xdr:row></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:row>12</xdr:row></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Embedded chart frame"/></xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChartEmbedded"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
		'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImageMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%201.png"/>
  <Relationship Id="rIdChartEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
		'xl/drawings/vmlDrawing1.vml': `<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape id="_x0000_s1025" o:spid="_x0000_s1025" style="position:absolute;visibility:visible">
    <v:imagedata r:id="rIdVmlImage"/>
    <x:ClientData ObjectType="Button"><x:Anchor>1, 15, 2, 3, 4, 5, 6, 7</x:Anchor><x:Visible/></x:ClientData>
  </v:shape>
</xml>`,
		'xl/drawings/_rels/vmlDrawing1.vml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVmlImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/vmlImage.emf"/>
</Relationships>`,
		'xl/charts/chart1.xml': chartXml(
			'barChart',
			'Embedded visual chart',
			'Visuals!$B$1',
			'Visuals!$A$2:$A$4',
			'Visuals!$B$2:$B$4',
		),
		'xl/charts/_rels/chart1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
  <Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
</Relationships>`,
		'xl/charts/chart2.xml': chartXml(
			'lineChart',
			'Chartsheet chart',
			'Visuals!$C$1',
			'Visuals!$A$2:$A$4',
			'Visuals!$C$2:$C$4',
		),
		'xl/charts/style1.xml':
			'<?xml version="1.0"?><cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="301"/>',
		'xl/charts/colors1.xml':
			'<?xml version="1.0"?><cs:colors xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle"/>',
		'xl/media/image 1.png': 'png-main',
		'xl/media/vmlImage.emf': 'emf-vml',
	})
}

function defaultCoveredChartSidecarWorkbook(): Uint8Array {
	return visualOwnershipWorkbook({ defaultCoveredChartSidecars: true })
}

function chartXml(
	chartType: 'barChart' | 'lineChart',
	title: string,
	nameRef: string,
	categoryRef: string,
	valueRef: string,
): string {
	return `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:${chartType}>
        <c:ser>
          <c:tx><c:strRef><c:f>${nameRef}</c:f><c:strCache><c:pt idx="0"><c:v>Header</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>${categoryRef}</c:f><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>${valueRef}</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:${chartType}>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`
}
