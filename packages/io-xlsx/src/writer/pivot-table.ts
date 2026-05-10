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
		const fieldRange = findNthElementRange(out, 'pivotField', field.index)
		if (!fieldRange) continue
		const fieldXml = out.slice(fieldRange.start, fieldRange.end)
		const updatedFieldXml = updateItemElements(fieldXml, field.items)
		out = `${out.slice(0, fieldRange.start)}${updatedFieldXml}${out.slice(fieldRange.end)}`
	}
	return out
}

function updateItemElements(xml: string, items: readonly PivotFieldItemInfo[]): string {
	let index = 0
	return xml.replace(
		/<(([A-Za-z_][\w.-]*:)?item)\b([^>]*)(\/>|>\s*<\/\1>)/g,
		(match, qualifiedName, _prefix, attrs, tail) => {
			const item = items[index]
			index++
			if (!item) return match
			const updatedAttrs = serializePivotItemAttributes(attrs, item)
			return tail.startsWith('/>')
				? `<${qualifiedName}${updatedAttrs}/>`
				: `<${qualifiedName}${updatedAttrs}></${qualifiedName}>`
		},
	)
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
	return xml.replace(
		/<(([A-Za-z_][\w.-]*:)?pageField)\b([^>]*)(\/>|>\s*<\/\1>)/g,
		(match, qualifiedName, _prefix, attrs, tail) => {
			const fieldIndex = readNumberAttr(attrs, 'fld')
			if (fieldIndex === undefined) return match
			const pageField = pivot.pageFields.find((entry) => entry.index === fieldIndex)
			if (!pageField) return match
			const updatedAttrs = syncXmlAttr(attrs, 'item', pageField.item)
			return tail.startsWith('/>')
				? `<${qualifiedName}${updatedAttrs}/>`
				: `<${qualifiedName}${updatedAttrs}></${qualifiedName}>`
		},
	)
}

function findNthElementRange(
	xml: string,
	localName: string,
	index: number,
): { start: number; end: number } | null {
	const tag = `(?:[A-Za-z_][\\w.-]*:)?${localName}`
	const openRe = new RegExp(`<${tag}\\b[^>]*>`, 'g')
	let seen = 0
	for (const match of xml.matchAll(openRe)) {
		if (seen !== index) {
			seen++
			continue
		}
		const start = match.index
		if (start === undefined) return null
		const openTag = match[0]
		if (openTag.endsWith('/>')) return { start, end: start + openTag.length }
		const closeRe = new RegExp(`</${tag}>`, 'g')
		closeRe.lastIndex = start + openTag.length
		const close = closeRe.exec(xml)
		if (!close) return null
		return { start, end: close.index + close[0].length }
	}
	return null
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
