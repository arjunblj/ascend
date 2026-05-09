import { describe, expect, test } from 'bun:test'
import { parseDrawingObjectRefs } from './drawing.ts'

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
				kind: 'connector',
				id: 11,
				name: 'Connector 1',
				anchor: { kind: 'oneCell', from: { col: 5, row: 6 } },
			},
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
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
				kind: 'graphicFrame',
				id: 12,
				name: 'Chart Frame',
				relIds: ['rIdChart'],
				anchor: { kind: 'absolute', x: 7, y: 8, cx: 9, cy: 10 },
			},
		])
	})
})
