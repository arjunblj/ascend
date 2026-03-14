import { parseA1Safe, type Sheet } from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const COMMENTS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const VML_NS = 'urn:schemas-microsoft-com:vml'
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office'
const EXCEL_NS = 'urn:schemas-microsoft-com:office:excel'

/**
 * Builds comments XML for a sheet.
 * Note: SheetComment stores only plain text (`text: string`). Rich text in comments
 * is flattened on read (see parseCommentsXml) and is not preserved on write.
 */
export function buildCommentsXml(sheet: Sheet): string {
	const authors = uniqueAuthors(sheet)
	const authorIds = new Map(authors.map((author, index) => [author, index] as const))
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<comments xmlns="${COMMENTS_NS}">`)
	out.push('<authors>')
	for (const author of authors) {
		out.push(`<author>${escapeXml(author)}</author>`)
	}
	out.push('</authors>')
	out.push('<commentList>')
	for (const [ref, comment] of sheet.comments) {
		const attrs = [`ref="${escapeXml(ref)}"`]
		const authorId = authorIds.get(comment.author ?? '')
		if (authorId !== undefined) attrs.push(`authorId="${authorId}"`)
		out.push(`<comment ${attrs.join(' ')}><text><t>${escapeXml(comment.text)}</t></text></comment>`)
	}
	out.push('</commentList>')
	out.push('</comments>')
	return out.toString()
}

export function buildCommentsVml(sheet: Sheet): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<xml xmlns:v="${VML_NS}" xmlns:o="${OFFICE_NS}" xmlns:x="${EXCEL_NS}">`)
	out.push('<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>')
	out.push(
		'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
	)
	out.push('<v:stroke joinstyle="miter"/>')
	out.push('<v:path gradientshapeok="t" o:connecttype="rect"/>')
	out.push('</v:shapetype>')
	let shapeId = 1024
	for (const [ref] of sheet.comments) {
		const pos = parseA1Safe(ref)
		if (!pos) continue
		out.push(
			`<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`,
		)
		out.push('<v:fill color2="#ffffe1"/>')
		out.push('<v:shadow on="t" color="black" obscured="t"/>')
		out.push('<v:path o:connecttype="none"/>')
		out.push(
			'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>',
		)
		out.push('<x:ClientData ObjectType="Note">')
		out.push('<x:MoveWithCells/>')
		out.push('<x:SizeWithCells/>')
		out.push(`<x:Row>${pos.row}</x:Row>`)
		out.push(`<x:Column>${pos.col}</x:Column>`)
		out.push('</x:ClientData>')
		out.push('</v:shape>')
		shapeId++
	}
	out.push('</xml>')
	return out.toString()
}

function uniqueAuthors(sheet: Sheet): string[] {
	const authors = new Set<string>()
	for (const comment of sheet.comments.values()) {
		authors.add(comment.author ?? '')
	}
	return [...authors]
}
