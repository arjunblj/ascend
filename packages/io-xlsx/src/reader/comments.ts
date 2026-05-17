import {
	type SheetComment,
	type SheetCommentLegacyDrawing,
	type SheetThreadedComment,
	toA1,
} from '@ascend/core'
import { asArray, attr, boolAttr, parseXml, type XmlNode } from '../xml.ts'
import { normalizeMainSpreadsheetNamespacePrefix } from './xml-utils.ts'

const VML_SHAPE_RE = /<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/gi
const ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const NOTE_CLIENT_DATA_RE =
	/<(?:[A-Za-z_][\w.-]*:)?ClientData\b([^>]*)ObjectType\s*=\s*(?:"Note"|'Note')([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?ClientData>/i

export function parseCommentsXml(xml: string): Map<string, SheetComment> {
	const doc = parseXml(normalizeMainSpreadsheetNamespacePrefix(xml))
	const comments = firstElement(doc, 'comments')
	if (!comments) return new Map()

	const authorsNode = childNode(comments, 'authors')
	const authors = asArray<XmlNode | string>(
		childValues(authorsNode, 'author') as XmlNode | XmlNode[] | string[] | undefined,
	).map((author) => readNodeText(author) ?? '')
	const entries = new Map<string, SheetComment>()
	const commentList = childNode(comments, 'commentList')
	if (!commentList) return entries

	for (const comment of childNodes(commentList, 'comment')) {
		const ref = attr(comment, 'ref')
		if (!ref) continue
		const authorId = Number(attr(comment, 'authorId') ?? '-1')
		const textNode = childNode(comment, 'text')
		const text = extractCommentText(textNode)
		const author = authorId >= 0 ? authors[authorId] : undefined
		entries.set(ref, author ? { text, author } : { text })
	}
	return entries
}

export function parseCommentVmlXml(xml: string): Map<string, SheetCommentLegacyDrawing> {
	const layouts = new Map<string, SheetCommentLegacyDrawing>()
	for (const match of xml.matchAll(VML_SHAPE_RE)) {
		const shapeAttrs = match[1] ?? ''
		const shapeBody = match[2] ?? ''
		const clientData = NOTE_CLIENT_DATA_RE.exec(shapeBody)
		if (!clientData) continue
		const clientBody = clientData[3] ?? ''
		const row = readIntElement(clientBody, 'Row')
		const column = readIntElement(clientBody, 'Column')
		if (row === undefined || column === undefined) continue
		const ref = toA1({ row, col: column })
		const anchor = parseAnchor(readElementText(clientBody, 'Anchor'))
		const visible = readBooleanElement(clientBody, 'Visible')
		const moveWithCells = readBooleanElement(clientBody, 'MoveWithCells')
		const sizeWithCells = readBooleanElement(clientBody, 'SizeWithCells')
		const shapeId = readAttr(shapeAttrs, 'id')
		const style = readAttr(shapeAttrs, 'style')
		const autoFill = readBooleanElement(clientBody, 'AutoFill')
		const layout: SheetCommentLegacyDrawing = {
			...(shapeId !== undefined ? { shapeId } : {}),
			...(style !== undefined ? { style } : {}),
			...(anchor ? { anchor } : {}),
			row,
			column,
			visible:
				visible ??
				(hasSelfClosingElement(clientBody, 'Visible') ||
					/visibility\s*:\s*visible/i.test(style ?? '')),
			moveWithCells: moveWithCells ?? hasSelfClosingElement(clientBody, 'MoveWithCells'),
			sizeWithCells: sizeWithCells ?? hasSelfClosingElement(clientBody, 'SizeWithCells'),
			...(autoFill !== undefined ? { autoFill } : {}),
		}
		layouts.set(ref, layout)
	}
	return layouts
}

export function parseThreadedCommentsXml(
	xml: string,
	partPath: string,
	people: ReadonlyMap<string, string> = new Map(),
): SheetThreadedComment[] {
	const doc = parseXml(xml)
	const comments = firstElement(doc, 'ThreadedComments') ?? firstElement(doc, 'threadedComments')
	if (!comments) return []

	return childNodes(comments, 'threadedComment').flatMap((comment) => {
		const ref = attr(comment, 'ref')
		if (!ref) return []
		const id = attr(comment, 'id')
		const parentId = attr(comment, 'parentId')
		const personId = attr(comment, 'personId')
		const author = personId ? people.get(personId) : undefined
		const dateTime = attr(comment, 'dT')
		const done = boolAttr(comment, 'done')
		const text = readNodeText(childValue(comment, 'text')) ?? ''
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
	})
}

export function parseThreadedCommentPersonsXml(xml: string): Map<string, string> {
	const entries = parseThreadedCommentPersonEntriesXml(xml)
	const people = new Map<string, string>()
	for (const entry of entries) {
		if (entry.displayName) people.set(entry.id, entry.displayName)
	}
	return people
}

export interface ThreadedCommentPersonEntry {
	readonly id: string
	readonly displayName?: string
	readonly index: number
}

export function parseThreadedCommentPersonEntriesXml(
	xml: string,
): readonly ThreadedCommentPersonEntry[] {
	const doc = parseXml(xml)
	const root =
		firstElement(doc, 'personList') ??
		firstElement(doc, 'persons') ??
		firstElement(doc, 'PersonList')
	const entries: ThreadedCommentPersonEntry[] = []
	if (!root) return entries
	const persons = childNodes(root, 'person')
	for (let index = 0; index < persons.length; index++) {
		const person = persons[index]
		if (!person) continue
		const id = attr(person, 'id') ?? attr(person, 'personId')
		const displayName =
			attr(person, 'displayName') ?? attr(person, 'name') ?? attr(person, 'userId')
		if (id) entries.push({ id, ...(displayName ? { displayName } : {}), index })
	}
	return entries
}

function extractCommentText(textNode: XmlNode | undefined): string {
	if (!textNode) return ''
	const directText = childValue(textNode, 't')
	if (directText !== undefined) return readNodeText(directText) ?? ''
	const runs = childNodes(textNode, 'r').map((run) => readNodeText(childValue(run, 't')) ?? '')
	return runs.join('')
}

function firstElement(doc: XmlNode, localName: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(doc)) {
		if (key.startsWith('@_') || localPart(key) !== localName || !isXmlNode(value)) continue
		return value
	}
	return undefined
}

function childNode(node: XmlNode | undefined, localName: string): XmlNode | undefined {
	return childNodes(node, localName)[0]
}

function childValue(node: XmlNode | undefined, localName: string): unknown {
	return childValues(node, localName)[0]
}

function childValues(node: XmlNode | undefined, localName: string): unknown[] {
	if (!node) return []
	const values: unknown[] = []
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || localPart(key) !== localName) continue
		if (Array.isArray(value)) values.push(...value)
		else values.push(value)
	}
	return values
}

function childNodes(node: XmlNode | undefined, localName: string): XmlNode[] {
	return childValues(node, localName).filter(isXmlNode)
}

function isXmlNode(value: unknown): value is XmlNode {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function localPart(name: string): string {
	return name.includes(':') ? (name.split(':').pop() ?? name) : name
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

function readAttr(rawAttrs: string, name: string): string | undefined {
	ATTR_RE.lastIndex = 0
	for (const match of rawAttrs.matchAll(ATTR_RE)) {
		if (match[1] === name) return match[2] ?? match[3] ?? ''
	}
	return undefined
}

function readElementText(xml: string, localName: string): string | undefined {
	const name = `(?:[A-Za-z_][\\w.-]*:)?${localName}`
	const re = new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)<\\/${name}>`, 'i')
	return re.exec(xml)?.[1]?.trim()
}

function readIntElement(xml: string, localName: string): number | undefined {
	const text = readElementText(xml, localName)
	if (text === undefined) return undefined
	const value = Number(text)
	return Number.isInteger(value) ? value : undefined
}

function readBooleanElement(xml: string, localName: string): boolean | undefined {
	const text = readElementText(xml, localName)
	if (text === undefined) return undefined
	if (/^(true|1)$/i.test(text)) return true
	if (/^(false|0)$/i.test(text)) return false
	return undefined
}

function hasSelfClosingElement(xml: string, localName: string): boolean {
	return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*/>`, 'i').test(xml)
}

function parseAnchor(text: string | undefined): SheetCommentLegacyDrawing['anchor'] | undefined {
	if (!text) return undefined
	const values = text
		.split(',')
		.map((part) => Number(part.trim()))
		.filter((value) => Number.isInteger(value))
	if (values.length !== 8) return undefined
	return values as unknown as SheetCommentLegacyDrawing['anchor']
}
