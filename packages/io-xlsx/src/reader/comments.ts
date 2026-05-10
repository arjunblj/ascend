import type { SheetComment, SheetThreadedComment } from '@ascend/core'
import { asArray, attr, boolAttr, parseXml, type XmlNode } from '../xml.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

export function parseCommentsXml(xml: string): Map<string, SheetComment> {
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
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

export function parseThreadedCommentsXml(
	xml: string,
	partPath: string,
	people: ReadonlyMap<string, string> = new Map(),
): SheetThreadedComment[] {
	const doc = parseXml(xml)
	const comments =
		(doc.ThreadedComments as XmlNode | undefined) ?? (doc.threadedComments as XmlNode | undefined)
	if (!comments) return []

	return asArray<XmlNode>(comments.threadedComment as XmlNode | XmlNode[] | undefined).flatMap(
		(comment) => {
			const ref = attr(comment, 'ref')
			if (!ref) return []
			const id = attr(comment, 'id')
			const parentId = attr(comment, 'parentId')
			const personId = attr(comment, 'personId')
			const author = personId ? people.get(personId) : undefined
			const dateTime = attr(comment, 'dT')
			const done = boolAttr(comment, 'done')
			const text = readNodeText(comment.text) ?? ''
			const parsed: SheetThreadedComment = {
				ref,
				text,
				partPath,
				...(id !== undefined ? { id } : {}),
				...(parentId !== undefined ? { parentId } : {}),
				...(personId !== undefined ? { personId } : {}),
				...(author !== undefined ? { author } : {}),
				...(dateTime !== undefined ? { dateTime } : {}),
				...(done !== undefined ? { done } : {}),
			}
			return [parsed]
		},
	)
}

export function parseThreadedCommentPersonsXml(xml: string): Map<string, string> {
	const doc = parseXml(xml)
	const root =
		(doc.personList as XmlNode | undefined) ??
		(doc.persons as XmlNode | undefined) ??
		(doc.PersonList as XmlNode | undefined)
	const people = new Map<string, string>()
	if (!root) return people
	for (const person of asArray<XmlNode>(root.person as XmlNode | XmlNode[] | undefined)) {
		const id = attr(person, 'id') ?? attr(person, 'personId')
		const displayName =
			attr(person, 'displayName') ?? attr(person, 'name') ?? attr(person, 'userId')
		if (id && displayName) people.set(id, displayName)
	}
	return people
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
