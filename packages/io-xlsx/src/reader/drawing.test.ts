import { describe, expect, test } from 'bun:test'
import {
	parseDrawingImageRefs,
	parseDrawingObjectRefs,
	parseVmlDrawingObjectRefs,
} from './drawing.ts'

describe('drawing inventory', () => {
	test('parses shape, text box, connector, and graphic frame anchors', () => {
		const refs = parseDrawingObjectRefs(
			`<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="10" name="Callout" descr="Revenue note"/></xdr:nvSpPr>
      <xdr:txBody><a:p><a:r><a:t>Revenue up</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>5</xdr:col><xdr:row>6</xdr:row></xdr:from>
    <xdr:cxnSp>
      <xdr:nvCxnSpPr><xdr:cNvPr id="11" name="Connector 1"/></xdr:nvCxnSpPr>
    </xdr:cxnSp>
  </xdr:oneCellAnchor>
  <xdr:absoluteAnchor>
    <xdr:pos x="7" y="8"/><xdr:ext cx="9" cy="10"/>
    <xdr:graphicFrame macro="Book.xlsm!RefreshChart">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="12" name="Chart Frame"/></xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
  </xdr:absoluteAnchor>
</xdr:wsDr>`,
			'xl/drawings/drawing1.xml',
		)

		expect(refs).toEqual([
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				source: 'drawingml',
				kind: 'connector',
				id: 11,
				name: 'Connector 1',
				anchor: { kind: 'oneCell', from: { col: 5, row: 6 } },
			},
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				source: 'drawingml',
				kind: 'textBox',
				id: 10,
				name: 'Callout',
				description: 'Revenue note',
				text: 'Revenue up',
				anchor: {
					kind: 'twoCell',
					editAs: 'oneCell',
					from: { col: 1, row: 2 },
					to: { col: 3, row: 4 },
				},
			},
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				source: 'drawingml',
				kind: 'graphicFrame',
				id: 12,
				name: 'Chart Frame',
				macro: 'Book.xlsm!RefreshChart',
				relIds: ['rIdChart'],
				anchor: { kind: 'absolute', x: 7, y: 8, cx: 9, cy: 10 },
			},
		])
	})

	test('parses DrawingML images and objects when the spreadsheet drawing namespace is prefixed differently', () => {
		const drawingXml = `<?xml version="1.0"?>
<sd:wsDr xmlns:sd="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sd:twoCellAnchor editAs="oneCell">
    <sd:from><sd:col>1</sd:col><sd:row>2</sd:row></sd:from>
    <sd:to><sd:col>3</sd:col><sd:row>4</sd:row></sd:to>
    <sd:pic>
      <sd:nvPicPr><sd:cNvPr id="1" name="Image 1" descr="Hero"/></sd:nvPicPr>
      <sd:blipFill><a:blip r:embed="rIdImg"/></sd:blipFill>
    </sd:pic>
  </sd:twoCellAnchor>
  <sd:absoluteAnchor>
    <sd:pos x="7" y="8"/><sd:ext cx="9" cy="10"/>
    <sd:graphicFrame macro="Book.xlsm!RefreshChart">
      <sd:nvGraphicFramePr><sd:cNvPr id="12" name="Chart Frame"/></sd:nvGraphicFramePr>
      <a:graphic><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic>
    </sd:graphicFrame>
  </sd:absoluteAnchor>
</sd:wsDr>`
		const relationships = [
			{
				id: 'rIdImg',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
				target: '../media/image1.png',
			},
			{
				id: 'rIdChart',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
				target: '../charts/chart1.xml',
			},
		]

		expect(parseDrawingImageRefs(drawingXml, 'xl/drawings/drawing1.xml', relationships)).toEqual([
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				relId: 'rIdImg',
				targetPath: 'xl/media/image1.png',
				name: 'Image 1',
				description: 'Hero',
				anchor: {
					kind: 'twoCell',
					editAs: 'oneCell',
					from: { col: 1, row: 2 },
					to: { col: 3, row: 4 },
				},
			},
		])
		expect(parseDrawingObjectRefs(drawingXml, 'xl/drawings/drawing1.xml', relationships)).toEqual([
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				source: 'drawingml',
				kind: 'graphicFrame',
				id: 12,
				name: 'Chart Frame',
				macro: 'Book.xlsm!RefreshChart',
				relIds: ['rIdChart'],
				relationshipRefs: [
					{
						id: 'rIdChart',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
						target: 'xl/charts/chart1.xml',
					},
				],
				anchor: { kind: 'absolute', x: 7, y: 8, cx: 9, cy: 10 },
			},
		])
	})

	test('parses VML drawing objects with anchors, notes, text, and relationships', () => {
		const refs = parseVmlDrawingObjectRefs(
			`<xml xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t201" style='position:absolute;
    margin-left:106.5pt;margin-top:49.5pt;width:87.75pt;height:70.5pt;
    visibility:visible' o:button="t">
    <v:textbox><div><font><b>Multi<br/>Line<br/>Text</b></font></div></v:textbox>
    <v:imagedata o:relid="rIdImage"/>
    <x:ClientData ObjectType="Button">
      <x:Anchor>2, 14, 3, 6, 4, 3, 8, 0</x:Anchor>
      <x:FmlaMacro>[0]!Button1_Click</x:FmlaMacro>
    </x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s1026">
    <x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData>
  </v:shape>
</xml>`,
			'xl/drawings/vmlDrawing1.vml',
			[
				{
					id: 'rIdImage',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
					target: '../media/image1.emf',
				},
			],
		)

		expect(refs).toEqual([
			{
				drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
				source: 'vml',
				kind: 'textBox',
				id: 1025,
				name: '_x0000_s1025',
				text: 'Multi Line Text',
				macro: '[0]!Button1_Click',
				style: expect.stringContaining('position:absolute'),
				vmlShapeId: '_x0000_s1025',
				vmlObjectType: 'Button',
				visible: true,
				relIds: ['rIdImage'],
				relationshipRefs: [
					{
						id: 'rIdImage',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						target: 'xl/media/image1.emf',
					},
				],
				anchor: {
					kind: 'twoCell',
					from: { col: 2, colOff: 14, row: 3, rowOff: 6 },
					to: { col: 4, colOff: 3, row: 8, rowOff: 0 },
				},
			},
			{
				drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
				source: 'vml',
				kind: 'shape',
				id: 1026,
				name: '_x0000_s1026',
				vmlShapeId: '_x0000_s1026',
				vmlObjectType: 'Note',
			},
		])
	})
})
