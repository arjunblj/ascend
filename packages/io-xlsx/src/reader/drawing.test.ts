import { describe, expect, test } from 'bun:test'
import { parseDrawingObjectRefs, parseVmlDrawingObjectRefs } from './drawing.ts'

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
    <xdr:graphicFrame>
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
				relIds: ['rIdChart'],
				anchor: { kind: 'absolute', x: 7, y: 8, cx: 9, cy: 10 },
			},
		])
	})

	test('parses non-comment VML drawing objects with anchors, text, and relationships', () => {
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
		])
	})
})
