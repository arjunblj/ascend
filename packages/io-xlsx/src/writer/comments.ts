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
const AUTHOR_RE = new RegExp(String.raw`<(${PREFIXED_TAG}author)\b[^>]*>([\s\S]*?)<\/\1>`, 'g')

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

export function syncCommentsXml(xml: string, sheet: Sheet): string {
	const sourceComments = readSourceComments(xml)
	if (sourceComments.length === 0) return buildCommentsXml(sheet)
	const { xml: xmlWithAuthors, authorIds } = syncCommentAuthorsXml(xml, uniqueAuthors(sheet))
	const sourceByRef = new Map(sourceComments.map((comment) => [comment.ref.toUpperCase(), comment]))
	const fallbackTag = sourceComments[0]?.tag ?? 'comment'
	const renderedComments = [...sheet.comments]
		.map(([ref, comment]) => {
			const source = sourceByRef.get(ref.toUpperCase())
			const authorId = authorIds.get(comment.author ?? '')
			if (!source) return buildCommentElement(fallbackTag, ref, comment.text, authorId)
			const attrs = setXmlAttr(source.attrs, 'authorId', authorId?.toString())
			const withAttrs = `<${source.tag}${attrs}>${source.body}</${source.tag}>`
			if (source.text === comment.text) return withAttrs
			return replaceCommentText(withAttrs, source.tag, attrs, source.body, comment.text)
		})
		.join('\n')
	const stripped = xmlWithAuthors.replace(COMMENT_RE, '').replace(/\s+$/, '')
	const listClose = /<\/((?:[A-Za-z_][\w.-]*:)?commentList)>/u
	const match = stripped.match(listClose)
	if (!match) return buildCommentsXml(sheet)
	const closingTag = match[0].trim()
	return stripped.replace(
		listClose,
		`${renderedComments ? `\n${renderedComments}\n` : '\n'}${closingTag}`,
	)
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
	const usedShapeIds = new Set<string>()
	let nextShapeId = nextGeneratedShapeIdStart(sheet)
	for (const [ref] of sheet.comments) {
		const pos = parseA1Safe(ref)
		if (!pos) continue
		const comment = sheet.comments.get(ref)
		const drawing = comment?.legacyDrawing
		const shapeIdValue = uniqueCommentShapeId(drawing?.shapeId, usedShapeIds, () => {
			let candidate = `_x0000_s${nextShapeId++}`
			while (usedShapeIds.has(candidate)) candidate = `_x0000_s${nextShapeId++}`
			return candidate
		})
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
	}
	out.push('</xml>')
	return out.toString()
}

function nextGeneratedShapeIdStart(sheet: Sheet): number {
	let maxShapeId = 1023
	for (const comment of sheet.comments.values()) {
		const match = comment.legacyDrawing?.shapeId?.match(/^_x0000_s(\d+)$/)
		if (!match) continue
		maxShapeId = Math.max(maxShapeId, Number(match[1]))
	}
	return Math.max(1024, maxShapeId + 1)
}

function uniqueCommentShapeId(
	preferred: string | undefined,
	usedShapeIds: Set<string>,
	nextGenerated: () => string,
): string {
	if (preferred && !usedShapeIds.has(preferred)) {
		usedShapeIds.add(preferred)
		return preferred
	}
	const generated = nextGenerated()
	usedShapeIds.add(generated)
	return generated
}

interface SourceComment {
	readonly xml: string
	readonly tag: string
	readonly attrs: string
	readonly body: string
	readonly ref: string
	readonly text: string
}

function readSourceComments(xml: string): readonly SourceComment[] {
	const comments: SourceComment[] = []
	for (const match of xml.matchAll(COMMENT_RE)) {
		const attrs = match[2] ?? ''
		const ref = readXmlAttr(attrs, 'ref')
		if (!ref) continue
		const body = match[3] ?? ''
		comments.push({
			xml: match[0],
			tag: match[1] ?? 'comment',
			attrs,
			body,
			ref,
			text: readCommentText(body),
		})
	}
	return comments
}

function syncCommentAuthorsXml(
	xml: string,
	requiredAuthors: readonly string[],
): { readonly xml: string; readonly authorIds: ReadonlyMap<string, number> } {
	const authors = readCommentAuthors(xml)
	const authorIds = new Map(authors.map((author, index) => [author, index] as const))
	const missingAuthors = requiredAuthors.filter((author) => !authorIds.has(author))
	if (missingAuthors.length === 0) return { xml, authorIds }
	for (const author of missingAuthors) {
		authorIds.set(author, authors.length)
		authors.push(author)
	}
	const authorTag = inferAuthorTag(xml)
	const renderedAuthors = missingAuthors
		.map((author) => `<${authorTag}>${escapeXml(author)}</${authorTag}>`)
		.join('')
	const authorsClose = /<\/((?:[A-Za-z_][\w.-]*:)?authors)>/u
	if (authorsClose.test(xml)) {
		return { xml: xml.replace(authorsClose, `${renderedAuthors}</$1>`), authorIds }
	}
	const commentsOpen = /<((?:[A-Za-z_][\w.-]*:)?comments)\b[^>]*>/u
	return {
		xml: xml.replace(commentsOpen, (open) => `${open}<authors>${renderedAuthors}</authors>`),
		authorIds,
	}
}

function readCommentAuthors(xml: string): string[] {
	return [...xml.matchAll(AUTHOR_RE)].map((match) => decodeXmlText(stripXmlTags(match[2] ?? '')))
}

function inferAuthorTag(xml: string): string {
	const match = xml.match(new RegExp(String.raw`<(${PREFIXED_TAG}author)\b`, 'u'))
	return match?.[1] ?? 'author'
}

function buildCommentElement(
	tag: string,
	ref: string,
	text: string,
	authorId: number | undefined,
): string {
	const attrs = [`ref="${escapeXml(ref)}"`]
	if (authorId !== undefined) attrs.push(`authorId="${authorId}"`)
	const prefix = tag.includes(':') ? `${tag.slice(0, tag.indexOf(':'))}:` : ''
	return `<${tag} ${attrs.join(' ')}><${prefix}text><${prefix}t>${escapeXml(text)}</${prefix}t></${prefix}text></${tag}>`
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

function readCommentText(body: string): string {
	const match = body.match(COMMENT_TEXT_RE)
	if (!match) return ''
	return decodeXmlText(stripXmlTags(match[3] ?? ''))
}

function stripXmlTags(value: string): string {
	return value.replace(/<[^>]*>/g, '')
}

function decodeXmlText(value: string): string {
	return value
		.replaceAll('&lt;', '<')
		.replaceAll('&gt;', '>')
		.replaceAll('&quot;', '"')
		.replaceAll('&apos;', "'")
		.replaceAll('&amp;', '&')
}

function setXmlAttr(attrs: string, name: string, value: string | undefined): string {
	const attrRe = new RegExp(String.raw`\s${name}\s*=\s*(?:"[^"]*"|'[^']*')`, 'u')
	if (value === undefined) return attrs.replace(attrRe, '')
	const escaped = escapeXml(value)
	if (attrRe.test(attrs)) return attrs.replace(attrRe, ` ${name}="${escaped}"`)
	return `${attrs} ${name}="${escaped}"`
}
