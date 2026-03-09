import type { Sheet } from '@ascend/core'
import { escapeXml } from '../xml.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const COMMENTS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const VML_NS = 'urn:schemas-microsoft-com:vml'
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office'
const EXCEL_NS = 'urn:schemas-microsoft-com:office:excel'

export function buildCommentsXml(sheet: Sheet): string {
	const authors = uniqueAuthors(sheet)
	const authorIds = new Map(authors.map((author, index) => [author, index] as const))
	const parts: string[] = [XML_HEADER, `<comments xmlns="${COMMENTS_NS}">`]
	parts.push('<authors>')
	for (const author of authors) {
		parts.push(`<author>${escapeXml(author)}</author>`)
	}
	parts.push('</authors>')
	parts.push('<commentList>')
	for (const [ref, comment] of sheet.comments) {
		const attrs = [`ref="${escapeXml(ref)}"`]
		const authorId = authorIds.get(comment.author ?? '')
		if (authorId !== undefined) attrs.push(`authorId="${authorId}"`)
		parts.push(
			`<comment ${attrs.join(' ')}><text><t>${escapeXml(comment.text)}</t></text></comment>`,
		)
	}
	parts.push('</commentList>')
	parts.push('</comments>')
	return parts.join('')
}

export function buildCommentsVml(sheet: Sheet): string {
	const parts: string[] = [
		XML_HEADER,
		`<xml xmlns:v="${VML_NS}" xmlns:o="${OFFICE_NS}" xmlns:x="${EXCEL_NS}">`,
		'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
		'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
		'<v:stroke joinstyle="miter"/>',
		'<v:path gradientshapeok="t" o:connecttype="rect"/>',
		'</v:shapetype>',
	]
	let shapeId = 1024
	for (const [ref] of sheet.comments) {
		const pos = parseA1Ref(ref)
		if (!pos) continue
		parts.push(
			`<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`,
		)
		parts.push('<v:fill color2="#ffffe1"/>')
		parts.push('<v:shadow on="t" color="black" obscured="t"/>')
		parts.push('<v:path o:connecttype="none"/>')
		parts.push(
			'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>',
		)
		parts.push('<x:ClientData ObjectType="Note">')
		parts.push('<x:MoveWithCells/>')
		parts.push('<x:SizeWithCells/>')
		parts.push(`<x:Row>${pos.row}</x:Row>`)
		parts.push(`<x:Column>${pos.col}</x:Column>`)
		parts.push('</x:ClientData>')
		parts.push('</v:shape>')
		shapeId++
	}
	parts.push('</xml>')
	return parts.join('')
}

function uniqueAuthors(sheet: Sheet): string[] {
	const authors = new Set<string>()
	for (const comment of sheet.comments.values()) {
		authors.add(comment.author ?? '')
	}
	return [...authors]
}

function parseA1Ref(ref: string): { row: number; col: number } | null {
	const match = /^([A-Za-z]+)(\d+)$/.exec(ref)
	if (!match) return null
	const colLabel = match[1]
	const rowText = match[2]
	if (!colLabel || !rowText) return null
	let col = 0
	for (const ch of colLabel.toUpperCase()) {
		col = col * 26 + (ch.charCodeAt(0) - 64)
	}
	return { row: Number.parseInt(rowText, 10) - 1, col: col - 1 }
}
