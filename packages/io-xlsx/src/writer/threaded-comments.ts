import type { SheetThreadedComment } from '@ascend/core'
import { readXmlAttr } from './xml-attrs.ts'

const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const THREADED_COMMENT_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}threadedComment)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)
const TEXT_RE = new RegExp(String.raw`<(${PREFIXED_TAG}text)\b([^>]*)>([\s\S]*?)<\/\1>`)
const PERSON_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}person)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)`,
	'g',
)

export interface ThreadedCommentTextRef {
	readonly id?: string
	readonly ref?: string
	readonly parentId?: string
	readonly personId?: string
	readonly dateTime?: string
	readonly done?: boolean
	readonly text: string
}

interface SourceThreadedComment extends ThreadedCommentTextRef {
	readonly xml: string
	readonly tag: string
	readonly attrs: string
	readonly body: string
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
	return readSourceThreadedComments(xml).map(
		({ id, ref, parentId, personId, dateTime, done, text }) => ({
			...(id !== undefined ? { id } : {}),
			...(ref !== undefined ? { ref } : {}),
			...(parentId !== undefined ? { parentId } : {}),
			...(personId !== undefined ? { personId } : {}),
			...(dateTime !== undefined ? { dateTime } : {}),
			...(done !== undefined ? { done } : {}),
			text,
		}),
	)
}

export function threadedCommentsMatchModel(
	xml: string,
	comments: readonly SheetThreadedComment[],
): boolean {
	const sourceComments = readThreadedCommentTextRefs(xml)
	if (sourceComments.length !== comments.length) return false
	const seen = new Set<string>()
	for (const comment of comments) {
		const source = findMatchingThreadedComment(sourceComments, comment, seen)
		if (!source || !threadedCommentMatches(source, comment)) return false
	}
	return true
}

export function syncThreadedCommentsXml(
	xml: string,
	comments: readonly SheetThreadedComment[],
): string {
	const sourceComments = readSourceThreadedComments(xml)
	const sourceByKey = new Map<string, SourceThreadedComment>()
	for (const source of sourceComments) {
		const key =
			source.id !== undefined ? `id:${source.id}` : source.ref ? `ref:${source.ref}` : null
		if (key && !sourceByKey.has(key)) sourceByKey.set(key, source)
	}
	const fallbackTag = sourceComments[0]?.tag ?? 'threadedComment'
	const renderedComments = comments
		.map((comment) => {
			const source = findSourceThreadedComment(sourceByKey, comment)
			return source
				? updateThreadedCommentElement(source, comment)
				: buildThreadedCommentElement(fallbackTag, comment)
		})
		.join('\n')
	const stripped = xml.replace(THREADED_COMMENT_RE, '').replace(/\s+$/, '')
	const rootClose = /<\/((?:[A-Za-z_][\w.-]*:)?(?:ThreadedComments|threadedComments))>\s*$/u
	const match = stripped.match(rootClose)
	if (!match) return `${stripped}\n${renderedComments}`
	const closingTag = match[0].trim()
	return stripped.replace(
		rootClose,
		`${renderedComments ? `\n${renderedComments}\n` : '\n'}${closingTag}`,
	)
}

export function threadedCommentPersonsMatchModel(
	xml: string,
	comments: readonly SheetThreadedComment[],
): boolean {
	const requiredPeople = collectRequiredThreadedCommentPeople(comments)
	if (requiredPeople.size === 0) return true
	const sourcePeople = readThreadedCommentPersons(xml)
	for (const [personId, author] of requiredPeople) {
		if (sourcePeople.get(personId) !== author) return false
	}
	return true
}

export function syncThreadedCommentPersonsXml(
	xml: string,
	comments: readonly SheetThreadedComment[],
): string {
	const requiredPeople = collectRequiredThreadedCommentPeople(comments)
	if (requiredPeople.size === 0) return xml

	const seen = new Set<string>()
	const updated = xml.replace(
		PERSON_RE,
		(personXml, tag: string, attrs: string, closeOrBody: string) => {
			const personId = readXmlAttr(attrs, 'id') ?? readXmlAttr(attrs, 'personId')
			if (!personId) return personXml
			const author = requiredPeople.get(personId)
			if (author === undefined) return personXml
			seen.add(personId)
			const nextAttrs = setXmlAttr(attrs, 'displayName', author)
			return `<${tag}${nextAttrs}${closeOrBody}`
		},
	)
	const missingPeople = [...requiredPeople].filter(([personId]) => !seen.has(personId))
	if (missingPeople.length === 0) return updated

	const rootClose = /<\/((?:[A-Za-z_][\w.-]*:)?(?:personList|persons|PersonList))>\s*$/u
	const match = updated.match(rootClose)
	const personPrefix = inferPersonTagPrefix(updated)
	const renderedPeople = missingPeople
		.map(
			([personId, author]) =>
				`<${personPrefix}person id="${escapeXml(personId)}" displayName="${escapeXml(author)}"/>`,
		)
		.join('\n')
	if (!match) return `${updated}\n${renderedPeople}`
	const closingTag = match[0].trim()
	return updated.replace(
		rootClose,
		`${renderedPeople ? `\n${renderedPeople}\n` : '\n'}${closingTag}`,
	)
}

export function buildThreadedCommentPersonsXml(comments: readonly SheetThreadedComment[]): string {
	const people = collectRequiredThreadedCommentPeople(comments)
	const persons = [...people]
		.map(
			([personId, author]) =>
				`  <person id="${escapeXml(personId)}" displayName="${escapeXml(author)}"/>`,
		)
		.join('\n')
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
${persons}
</personList>`
}

export function hasThreadedCommentPersonAuthors(
	comments: readonly SheetThreadedComment[],
): boolean {
	return collectRequiredThreadedCommentPeople(comments).size > 0
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

function findMatchingThreadedComment(
	sourceComments: readonly ThreadedCommentTextRef[],
	comment: SheetThreadedComment,
	seen: Set<string>,
): ThreadedCommentTextRef | undefined {
	const candidates = comment.id
		? sourceComments.filter((source) => source.id === comment.id)
		: sourceComments.filter((source) => source.ref === comment.ref && source.id === undefined)
	if (candidates.length !== 1) return undefined
	const key = comment.id ? `id:${comment.id}` : `ref:${comment.ref}`
	if (seen.has(key)) return undefined
	seen.add(key)
	return candidates[0]
}

function findSourceThreadedComment(
	sourceByKey: ReadonlyMap<string, SourceThreadedComment>,
	comment: SheetThreadedComment,
): SourceThreadedComment | undefined {
	if (comment.id) {
		const exact = sourceByKey.get(`id:${comment.id}`)
		if (exact) return exact
		const copiedFromId = sourceThreadedCommentIdForCopiedComment(comment.id)
		if (copiedFromId) {
			const source = sourceByKey.get(`id:${copiedFromId}`)
			if (source) return source
		}
	}
	return sourceByKey.get(`ref:${comment.ref}`)
}

function sourceThreadedCommentIdForCopiedComment(id: string): string | undefined {
	const match = id.match(/^(.+)-copy(?:-\d+)?$/u)
	return match?.[1]
}

function threadedCommentMatches(
	source: ThreadedCommentTextRef,
	comment: SheetThreadedComment,
): boolean {
	return (
		source.ref === comment.ref &&
		source.text === comment.text &&
		source.id === comment.id &&
		source.parentId === comment.parentId &&
		source.personId === comment.personId &&
		source.dateTime === comment.dateTime &&
		source.done === comment.done
	)
}

function readSourceThreadedComments(xml: string): readonly SourceThreadedComment[] {
	const refs: SourceThreadedComment[] = []
	for (const match of xml.matchAll(THREADED_COMMENT_RE)) {
		const attrs = match[2] ?? ''
		const body = match[3] ?? ''
		const id = readXmlAttr(attrs, 'id')
		const ref = readXmlAttr(attrs, 'ref')
		const parentId = readXmlAttr(attrs, 'parentId')
		const personId = readXmlAttr(attrs, 'personId')
		const dateTime = readXmlAttr(attrs, 'dT')
		const done = readBoolXmlAttr(attrs, 'done')
		refs.push({
			xml: match[0],
			tag: match[1] ?? 'threadedComment',
			attrs,
			body,
			...(id !== undefined ? { id } : {}),
			...(ref !== undefined ? { ref } : {}),
			...(parentId !== undefined ? { parentId } : {}),
			...(personId !== undefined ? { personId } : {}),
			...(dateTime !== undefined ? { dateTime } : {}),
			...(done !== undefined ? { done } : {}),
			text: readThreadedCommentText(body),
		})
	}
	return refs
}

function readThreadedCommentPersons(xml: string): Map<string, string> {
	const people = new Map<string, string>()
	for (const match of xml.matchAll(PERSON_RE)) {
		const attrs = match[2] ?? ''
		const personId = readXmlAttr(attrs, 'id') ?? readXmlAttr(attrs, 'personId')
		const displayName =
			readXmlAttr(attrs, 'displayName') ??
			readXmlAttr(attrs, 'name') ??
			readXmlAttr(attrs, 'userId')
		if (personId && displayName) people.set(personId, displayName)
	}
	return people
}

function collectRequiredThreadedCommentPeople(
	comments: readonly SheetThreadedComment[],
): Map<string, string> {
	const people = new Map<string, string>()
	for (const comment of comments) {
		if (!comment.personId || !comment.author || people.has(comment.personId)) continue
		people.set(comment.personId, comment.author)
	}
	return people
}

function inferPersonTagPrefix(xml: string): string {
	const match = xml.match(new RegExp(String.raw`<(${PREFIXED_TAG}person)\b`, 'u'))
	const tag = match?.[1]
	return tag?.includes(':') ? `${tag.slice(0, tag.indexOf(':'))}:` : ''
}

function updateThreadedCommentElement(
	source: SourceThreadedComment,
	comment: SheetThreadedComment,
): string {
	const attrs = threadedCommentAttrs(source.attrs, comment)
	const withAttrs = `<${source.tag}${attrs}>${source.body}</${source.tag}>`
	return replaceThreadedCommentText(withAttrs, source.tag, attrs, source.body, comment.text)
}

function buildThreadedCommentElement(tag: string, comment: SheetThreadedComment): string {
	const prefix = tag.includes(':') ? `${tag.slice(0, tag.indexOf(':'))}:` : ''
	const attrs = threadedCommentAttrs('', comment)
	return `<${tag}${attrs}><${prefix}text>${escapeXml(comment.text)}</${prefix}text></${tag}>`
}

function threadedCommentAttrs(attrs: string, comment: SheetThreadedComment): string {
	let next = attrs
	next = setXmlAttr(next, 'ref', comment.ref)
	next = setXmlAttr(next, 'personId', comment.personId)
	next = setXmlAttr(next, 'id', comment.id)
	next = setXmlAttr(next, 'parentId', comment.parentId)
	next = setXmlAttr(next, 'dT', comment.dateTime)
	next = setXmlAttr(next, 'done', comment.done === undefined ? undefined : comment.done ? '1' : '0')
	return next
}

function setXmlAttr(attrs: string, name: string, value: string | undefined): string {
	const attrRe = new RegExp(String.raw`\s${name}\s*=\s*(?:"[^"]*"|'[^']*')`, 'u')
	if (value === undefined) return attrs.replace(attrRe, '')
	const escaped = escapeXml(value)
	if (attrRe.test(attrs)) return attrs.replace(attrRe, ` ${name}="${escaped}"`)
	return `${attrs} ${name}="${escaped}"`
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

function decodeXmlText(value: string): string {
	return value
		.replaceAll('&lt;', '<')
		.replaceAll('&gt;', '>')
		.replaceAll('&quot;', '"')
		.replaceAll('&apos;', "'")
		.replaceAll('&amp;', '&')
}

function readBoolXmlAttr(attrs: string, name: string): boolean | undefined {
	const value = readXmlAttr(attrs, name)
	if (value === undefined) return undefined
	const normalized = value.toLowerCase()
	if (normalized === '1' || normalized === 'true') return true
	if (normalized === '0' || normalized === 'false') return false
	return undefined
}

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&apos;')
}
