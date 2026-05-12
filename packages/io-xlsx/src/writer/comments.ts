import { parseA1Safe, type Sheet } from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'
import { readXmlAttr } from './xml-attrs.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const COMMENTS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const VML_NS = 'urn:schemas-microsoft-com:vml'
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office'
const EXCEL_NS = 'urn:schemas-microsoft-com:office:excel'
const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const COMMENT_RE = new RegExp(String.raw`<(${PREFIXED_TAG}comment)\b([^>]*)>([\s\S]*?)<\/\1>`, 'g')
const COMMENT_TEXT_RE = new RegExp(String.raw`<(${PREFIXED_TAG}text)\b([^>]*)>([\s\S]*?)<\/\1>`)

export interface LegacyCommentTextUpdate {
	readonly ref: string
	readonly text: string
}

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

export function updateCommentsXml(
	xml: string,
	updates: readonly LegacyCommentTextUpdate[],
): string {
	if (updates.length === 0) return xml
	return xml.replace(COMMENT_RE, (commentXml, tag: string, attrs: string, body: string) => {
		const ref = readXmlAttr(attrs, 'ref')
		const update = ref
			? updates.find((entry) => entry.ref.toUpperCase() === ref.toUpperCase())
			: undefined
		if (!update) return commentXml
		return replaceCommentText(commentXml, tag, attrs, body, update.text)
	})
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
		const comment = sheet.comments.get(ref)
		const drawing = comment?.legacyDrawing
		const shapeIdValue = drawing?.shapeId ?? `_x0000_s${shapeId}`
		const visible = drawing?.visible ?? false
		const style = normalizeCommentVmlStyle(drawing?.style, visible)
		out.push(
			`<v:shape id="${escapeXml(shapeIdValue)}" type="#_x0000_t202" style="${escapeXml(style)}" fillcolor="#ffffe1" o:insetmode="auto">`,
		)
		out.push('<v:fill color2="#ffffe1"/>')
		out.push('<v:shadow on="t" color="black" obscured="t"/>')
		out.push('<v:path o:connecttype="none"/>')
		out.push(
			'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>',
		)
		out.push('<x:ClientData ObjectType="Note">')
		if (drawing?.moveWithCells !== false) out.push('<x:MoveWithCells/>')
		if (drawing?.sizeWithCells !== false) out.push('<x:SizeWithCells/>')
		if (drawing?.anchor) out.push(`<x:Anchor>${drawing.anchor.join(', ')}</x:Anchor>`)
		if (drawing?.autoFill !== undefined) out.push(`<x:AutoFill>${drawing.autoFill}</x:AutoFill>`)
		if (visible) out.push('<x:Visible/>')
		else if (drawing?.visible === false) out.push('<x:Visible>false</x:Visible>')
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

function normalizeCommentVmlStyle(style: string | undefined, visible: boolean): string {
	const fallback =
		'position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden'
	const nextStyle = style && style.trim().length > 0 ? style : fallback
	if (/visibility\s*:/i.test(nextStyle)) {
		return nextStyle.replace(
			/visibility\s*:\s*(visible|hidden)/i,
			`visibility:${visible ? 'visible' : 'hidden'}`,
		)
	}
	return `${nextStyle};visibility:${visible ? 'visible' : 'hidden'}`
}

function replaceCommentText(
	commentXml: string,
	tag: string,
	attrs: string,
	body: string,
	text: string,
): string {
	const replacement = body.match(COMMENT_TEXT_RE)
	const prefix = tag.includes(':') ? `${tag.slice(0, tag.indexOf(':'))}:` : ''
	if (!replacement)
		return `<${tag}${attrs}><${prefix}text><${prefix}t>${escapeXml(text)}</${prefix}t></${prefix}text>${body}</${tag}>`
	return commentXml.replace(COMMENT_TEXT_RE, (_match, textTag: string, textAttrs: string) => {
		const textPrefix = textTag.includes(':') ? `${textTag.slice(0, textTag.indexOf(':'))}:` : ''
		return `<${textTag}${textAttrs}><${textPrefix}t>${escapeXml(text)}</${textPrefix}t></${textTag}>`
	})
}
