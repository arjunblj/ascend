import type { SlicerCacheInfo, SlicerCacheItemInfo } from '@ascend/core'
import { readNumberXmlAttr, removeXmlAttr, setXmlAttr } from './xml-attrs.ts'

const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const TABULAR_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}tabular)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)`,
)
const ITEMS_RE = new RegExp(String.raw`<(${PREFIXED_TAG}items)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)`)
const ITEM_RE = new RegExp(String.raw`<(${PREFIXED_TAG}i)\b([^>]*?)(\/>|>\s*<\/\1>)`, 'g')

export function updateSlicerCacheDefinitionXml(xml: string, cache: SlicerCacheInfo): string {
	if (!cache.items || cache.items.length === 0) return xml
	const items = cache.items
	return xml.replace(TABULAR_RE, (_match, tag: string, attrs: string, tail: string, body = '') => {
		const itemsTag = deriveChildTag(tag, 'items')
		if (tail === '/>') {
			return `<${tag}${attrs}>${buildItemsXml(itemsTag, deriveChildTag(tag, 'i'), items)}</${tag}>`
		}
		const updatedBody = ITEMS_RE.test(body)
			? body.replace(
					ITEMS_RE,
					(
						itemsMatch: string,
						childTag: string,
						itemAttrs: string,
						itemTail: string,
						itemBody = '',
					) => updateItemsXml(itemsMatch, childTag, itemAttrs, itemTail, itemBody, items),
				)
			: `${body}${buildItemsXml(itemsTag, deriveChildTag(tag, 'i'), items)}`
		return `<${tag}${attrs}>${updatedBody}</${tag}>`
	})
}

function updateItemsXml(
	_match: string,
	tag: string,
	attrs: string,
	tail: string,
	body: string,
	items: readonly SlicerCacheItemInfo[],
): string {
	const byIndex = new Map(items.map((item) => [item.index, item]))
	const seen = new Set<number>()
	const itemTag = deriveChildTag(tag, 'i')
	const originalBody = tail === '/>' ? '' : body
	let updatedBody = originalBody.replace(ITEM_RE, (node, childTag: string, itemAttrs: string) => {
		const index = readNumberAttr(itemAttrs, 'x')
		if (index === undefined) return node
		seen.add(index)
		const item = byIndex.get(index)
		return item ? buildItemXml(childTag, item, itemAttrs) : node
	})
	const missing = [...items].filter((item) => !seen.has(item.index))
	if (missing.length > 0) {
		updatedBody += missing.map((item) => buildItemXml(itemTag, item)).join('')
	}
	const count = countItems(updatedBody)
	const updatedAttrs = setXmlAttr(attrs, 'count', String(count))
	return `<${tag}${updatedAttrs}>${updatedBody}</${tag}>`
}

function buildItemsXml(
	itemsTag: string,
	itemTag: string,
	items: readonly SlicerCacheItemInfo[],
): string {
	return `<${itemsTag} count="${items.length}">${items.map((item) => buildItemXml(itemTag, item)).join('')}</${itemsTag}>`
}

function buildItemXml(tag: string, item: SlicerCacheItemInfo, attrs = ''): string {
	let updated = setXmlAttr(attrs, 'x', String(item.index))
	updated = setOptionalBoolAttr(updated, 's', item.selected)
	updated = setOptionalBoolAttr(updated, 'nd', item.noData)
	return `<${tag}${updated}/>`
}

function deriveChildTag(parentTag: string, localName: string): string {
	const colon = parentTag.indexOf(':')
	return colon === -1 ? localName : `${parentTag.slice(0, colon + 1)}${localName}`
}

function readNumberAttr(attrs: string, name: string): number | undefined {
	return readNumberXmlAttr(attrs, name)
}

function setOptionalBoolAttr(attrs: string, name: string, value: boolean | undefined): string {
	if (value === undefined) return removeXmlAttr(attrs, name)
	return setXmlAttr(attrs, name, value ? '1' : '0')
}

function countItems(body: string): number {
	return [...body.matchAll(ITEM_RE)].length
}
