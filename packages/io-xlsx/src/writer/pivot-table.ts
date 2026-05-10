import type { PivotFieldItemInfo, PivotTableInfo } from '@ascend/core'
import { escapeXml } from '../xml.ts'

export function updatePivotTableDefinitionXml(xml: string, pivot: PivotTableInfo): string {
	let out = updatePivotFieldItems(xml, pivot)
	out = updatePageFieldItems(out, pivot)
	return out
}

function updatePivotFieldItems(xml: string, pivot: PivotTableInfo): string {
	let out = xml
	for (const field of pivot.fields) {
		if (!field.items || field.items.length === 0) continue
		const fieldRange = findNthPivotFieldRange(out, field.index)
		if (!fieldRange) continue
		const fieldXml = out.slice(fieldRange.start, fieldRange.end)
		const updatedFieldXml = updateItemElements(fieldXml, field.items)
		out = `${out.slice(0, fieldRange.start)}${updatedFieldXml}${out.slice(fieldRange.end)}`
	}
	return out
}

function updateItemElements(xml: string, items: readonly PivotFieldItemInfo[]): string {
	const fieldRange = findFirstElementRange(xml, 'pivotField')
	if (!fieldRange) return xml
	const itemsRange = directChildElementRanges(xml, fieldRange, 'items')[0]
	if (!itemsRange) return xml
	const itemRanges = directChildElementRanges(xml, itemsRange, 'item')
	let out = xml
	for (let index = itemRanges.length - 1; index >= 0; index--) {
		const range = itemRanges[index]
		const item = items[index]
		if (!range || !item) continue
		const updatedAttrs = serializePivotItemAttributes(range.attrs, item)
		const updated = range.selfClosing
			? `<${range.qualifiedName}${updatedAttrs}/>`
			: `<${range.qualifiedName}${updatedAttrs}></${range.qualifiedName}>`
		out = `${out.slice(0, range.start)}${updated}${out.slice(range.end)}`
	}
	return out
}

function serializePivotItemAttributes(attrs: string, item: PivotFieldItemInfo): string {
	let updated = attrs as string
	updated = syncXmlAttr(updated, 'x', item.cacheIndex)
	updated = syncXmlAttr(updated, 't', item.itemType)
	updated = syncXmlAttr(updated, 'n', item.caption)
	updated = syncXmlAttr(updated, 'h', item.hidden)
	updated = syncXmlAttr(updated, 's', item.manualFilter)
	updated = syncXmlAttr(updated, 'sd', item.showDetails)
	updated = syncXmlAttr(updated, 'f', item.calculated)
	updated = syncXmlAttr(updated, 'm', item.missing)
	updated = syncXmlAttr(updated, 'c', item.childItems)
	updated = syncXmlAttr(updated, 'd', item.expanded)
	updated = syncXmlAttr(updated, 'e', item.drillAcrossAttributes)
	return updated
}

function updatePageFieldItems(xml: string, pivot: PivotTableInfo): string {
	if (pivot.pageFields.length === 0) return xml
	const pageFieldsRange = findFirstElementRange(xml, 'pageFields')
	if (!pageFieldsRange) return xml
	const pageFieldRanges = directChildElementRanges(xml, pageFieldsRange, 'pageField')
	let out = xml
	for (let index = pageFieldRanges.length - 1; index >= 0; index--) {
		const range = pageFieldRanges[index]
		if (!range) continue
		const fieldIndex = readNumberAttr(range.attrs, 'fld')
		if (fieldIndex === undefined) continue
		const pageField = pivot.pageFields.find((entry) => entry.index === fieldIndex)
		if (!pageField) continue
		const updatedAttrs = syncXmlAttr(range.attrs, 'item', pageField.item)
		const updated = range.selfClosing
			? `<${range.qualifiedName}${updatedAttrs}/>`
			: `<${range.qualifiedName}${updatedAttrs}></${range.qualifiedName}>`
		out = `${out.slice(0, range.start)}${updated}${out.slice(range.end)}`
	}
	return out
}

interface ElementRange {
	readonly start: number
	readonly end: number
	readonly openEnd: number
	readonly closeStart: number
	readonly qualifiedName: string
	readonly attrs: string
	readonly selfClosing: boolean
}

function findNthPivotFieldRange(xml: string, index: number): ElementRange | null {
	const pivotFieldsRange = findFirstElementRange(xml, 'pivotFields')
	if (!pivotFieldsRange) return null
	return directChildElementRanges(xml, pivotFieldsRange, 'pivotField')[index] ?? null
}

function findFirstElementRange(xml: string, localName: string): ElementRange | null {
	for (const match of xml.matchAll(elementTagRe())) {
		const tag = parseOpeningTag(match)
		if (!tag || localPart(tag.qualifiedName) !== localName) continue
		return elementRangeFromTag(xml, tag)
	}
	return null
}

function directChildElementRanges(
	xml: string,
	parent: ElementRange,
	localName: string,
): ElementRange[] {
	const ranges: ElementRange[] = []
	const body = xml.slice(parent.openEnd, parent.closeStart)
	const re = elementTagRe()
	let depth = 0
	for (const match of body.matchAll(re)) {
		const raw = match[0]
		if (raw.startsWith('</')) {
			depth = Math.max(0, depth - 1)
			continue
		}
		const tag = parseOpeningTag(match, parent.openEnd)
		if (!tag) continue
		if (depth === 0 && localPart(tag.qualifiedName) === localName) {
			const range = elementRangeFromTag(xml, tag)
			if (range) ranges.push(range)
		}
		if (!tag.selfClosing) depth++
	}
	return ranges
}

function elementTagRe(): RegExp {
	return /<\/?([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)([^>]*?)(\/?)>/g
}

function parseOpeningTag(
	match: RegExpMatchArray,
	offset = 0,
): {
	readonly start: number
	readonly openEnd: number
	readonly qualifiedName: string
	readonly attrs: string
	readonly selfClosing: boolean
} | null {
	const raw = match[0]
	const index = match.index
	const qualifiedName = match[1]
	if (index === undefined || !qualifiedName || raw.startsWith('</')) return null
	return {
		start: offset + index,
		openEnd: offset + index + raw.length,
		qualifiedName,
		attrs: match[2] ?? '',
		selfClosing: raw.endsWith('/>'),
	}
}

function elementRangeFromTag(
	xml: string,
	tag: {
		readonly start: number
		readonly openEnd: number
		readonly qualifiedName: string
		readonly attrs: string
		readonly selfClosing: boolean
	},
): ElementRange | null {
	if (tag.selfClosing) {
		return {
			...tag,
			end: tag.openEnd,
			closeStart: tag.openEnd,
		}
	}
	const close = findMatchingCloseTag(xml, tag.qualifiedName, tag.openEnd)
	if (!close) return null
	return {
		...tag,
		end: close.end,
		closeStart: close.start,
	}
}

function findMatchingCloseTag(
	xml: string,
	qualifiedName: string,
	fromIndex: number,
): { start: number; end: number } | null {
	const re = new RegExp(`<(/?)${escapeRegExp(qualifiedName)}\\b[^>]*(/?)>`, 'g')
	re.lastIndex = fromIndex
	let depth = 1
	for (;;) {
		const match = re.exec(xml)
		if (!match) return null
		const isClosing = match[1] === '/'
		const isSelfClosing = !isClosing && match[2] === '/'
		if (isSelfClosing) continue
		depth += isClosing ? -1 : 1
		if (depth === 0) return { start: match.index, end: match.index + match[0].length }
	}
}

function localPart(qualifiedName: string): string {
	const colon = qualifiedName.indexOf(':')
	return colon === -1 ? qualifiedName : qualifiedName.slice(colon + 1)
}

function escapeRegExp(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function syncXmlAttr(
	attrs: string,
	name: string,
	value: string | number | boolean | undefined,
): string {
	const attrPattern = new RegExp(`\\s${name}="[^"]*"`)
	if (value === undefined) return attrs.replace(attrPattern, '')
	const serialized = typeof value === 'boolean' ? (value ? '1' : '0') : escapeXml(String(value))
	const attrText = `${name}="${serialized}"`
	if (attrPattern.test(attrs)) return attrs.replace(attrPattern, ` ${attrText}`)
	return `${attrs} ${attrText}`
}

function readNumberAttr(attrs: string, name: string): number | undefined {
	const match = new RegExp(`\\s${name}="([^"]*)"`).exec(attrs)
	if (!match?.[1]) return undefined
	const parsed = Number(match[1])
	return Number.isNaN(parsed) ? undefined : parsed
}
