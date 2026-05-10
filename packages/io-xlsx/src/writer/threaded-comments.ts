const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const THREADED_COMMENT_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}threadedComment)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)
const TEXT_RE = new RegExp(String.raw`<(${PREFIXED_TAG}text)\b([^>]*)>([\s\S]*?)<\/\1>`)

export interface ThreadedCommentTextRef {
	readonly id?: string
	readonly ref?: string
	readonly text: string
}

export function updateThreadedCommentsXml(
	xml: string,
	updates: readonly ThreadedCommentTextRef[],
): string {
	if (updates.length === 0) return xml
	return xml.replace(
		THREADED_COMMENT_RE,
		(commentXml, tag: string, attrs: string, body: string) => {
			const update = findThreadedCommentUpdate(attrs, updates)
			return update
				? replaceThreadedCommentText(commentXml, tag, attrs, body, update.text)
				: commentXml
		},
	)
}

export function readThreadedCommentTextRefs(xml: string): readonly ThreadedCommentTextRef[] {
	const refs: ThreadedCommentTextRef[] = []
	for (const match of xml.matchAll(THREADED_COMMENT_RE)) {
		const attrs = match[2] ?? ''
		const body = match[3] ?? ''
		const id = readXmlAttr(attrs, 'id')
		const ref = readXmlAttr(attrs, 'ref')
		refs.push({
			...(id !== undefined ? { id } : {}),
			...(ref !== undefined ? { ref } : {}),
			text: readThreadedCommentText(body),
		})
	}
	return refs
}

function findThreadedCommentUpdate(
	attrs: string,
	updates: readonly ThreadedCommentTextRef[],
): ThreadedCommentTextRef | undefined {
	const id = readXmlAttr(attrs, 'id')
	const ref = readXmlAttr(attrs, 'ref')
	return updates.find((update) => {
		if (update.id !== undefined) return update.id === id
		if (update.ref !== undefined) return update.ref === ref
		return false
	})
}

function replaceThreadedCommentText(
	commentXml: string,
	tag: string,
	attrs: string,
	body: string,
	text: string,
): string {
	const replacement = body.match(TEXT_RE)
	const prefix = tag.includes(':') ? `${tag.slice(0, tag.indexOf(':'))}:` : ''
	if (!replacement)
		return `<${tag}${attrs}><${prefix}text>${escapeXml(text)}</${prefix}text>${body}</${tag}>`
	return commentXml.replace(TEXT_RE, (_match, textTag: string, textAttrs: string) => {
		return `<${textTag}${textAttrs}>${escapeXml(text)}</${textTag}>`
	})
}

function readThreadedCommentText(body: string): string {
	const match = body.match(TEXT_RE)
	if (!match) return ''
	return decodeXmlText(match[3] ?? '')
}

function readXmlAttr(attrs: string, name: string): string | undefined {
	const match = attrs.match(new RegExp(String.raw`\s${name}="([^"]*)"`))
	return match?.[1]
}

function decodeXmlText(value: string): string {
	return value
		.replaceAll('&lt;', '<')
		.replaceAll('&gt;', '>')
		.replaceAll('&quot;', '"')
		.replaceAll('&apos;', "'")
		.replaceAll('&amp;', '&')
}

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&apos;')
}
