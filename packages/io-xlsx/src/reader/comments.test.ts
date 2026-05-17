import { describe, expect, test } from 'bun:test'
import { parseCommentVmlXml } from './comments.ts'

describe('comment VML inventory', () => {
	test('parses note layout shapes when the VML namespace uses a non-v prefix', () => {
		const layouts = parseCommentVmlXml(`<xml xmlns:vm="urn:schemas-microsoft-com:vml"
  xmlns:x="urn:schemas-microsoft-com:office:excel">
  <vm:shape id="_x0000_s2048" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Anchor>1, 15, 0, 2, 3, 20, 4, 8</x:Anchor>
      <x:Visible/>
      <x:Row>0</x:Row>
      <x:Column>1</x:Column>
    </x:ClientData>
  </vm:shape>
</xml>`)

		expect(layouts.get('B1')).toEqual({
			shapeId: '_x0000_s2048',
			style: 'position:absolute;visibility:visible',
			anchor: [1, 15, 0, 2, 3, 20, 4, 8],
			row: 0,
			column: 1,
			visible: true,
			moveWithCells: true,
			sizeWithCells: true,
		})
	})
})
