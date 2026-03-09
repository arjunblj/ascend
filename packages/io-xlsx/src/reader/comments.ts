import type { SheetComment } from '@ascend/core'
import { asArray, attr, parseXml, type XmlNode } from '../xml.ts'

export function parseCommentsXml(xml: string): Map<string, SheetComment> {
	const doc = parseXml(xml)
	const comments = doc.comments as XmlNode | undefined
	if (!comments) return new Map()

	const authorsNode = comments.authors as XmlNode | undefined
	const authors = asArray<XmlNode | string>(
		authorsNode?.author as XmlNode | XmlNode[] | string[] | undefined,
	).map((author) => readNodeText(author) ?? '')
	const entries = new Map<string, SheetComment>()
	const commentList = comments.commentList as XmlNode | undefined
	if (!commentList) return entries

	for (const comment of asArray<XmlNode>(commentList.comment as XmlNode | XmlNode[])) {
		const ref = attr(comment, 'ref')
		if (!ref) continue
		const authorId = Number(attr(comment, 'authorId') ?? '-1')
		const textNode = comment.text as XmlNode | undefined
		const text = extractCommentText(textNode)
		const author = authorId >= 0 ? authors[authorId] : undefined
		entries.set(ref, author ? { text, author } : { text })
	}
	return entries
}

function extractCommentText(textNode: XmlNode | undefined): string {
	if (!textNode) return ''
	if (textNode.t !== undefined) return String(textNode.t)
	const runs = asArray<XmlNode>(textNode.r as XmlNode | XmlNode[]).map(
		(run) => readNodeText(run.t) ?? '',
	)
	return runs.join('')
}

function readNodeText(node: unknown): string | undefined {
	if (node === undefined || node === null) return undefined
	if (typeof node === 'string' || typeof node === 'number' || typeof node === 'boolean') {
		return String(node)
	}
	if (typeof node === 'object') {
		const text = (node as XmlNode)['#text']
		return text !== undefined && text !== null ? String(text) : undefined
	}
	return undefined
}
