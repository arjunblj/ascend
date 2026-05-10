import type { SheetAdvancedFilterInfo } from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { buildAutoFilterXml } from './filtering.ts'

const CUSTOM_SHEET_VIEW_RE =
	/<(?:(?<prefix>[A-Za-z_][\w.-]*):)?customSheetView\b(?<attrs>[^>]*)>(?<body>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?customSheetView>/g
const AUTO_FILTER_RE =
	/<(?:(?<prefix>[A-Za-z_][\w.-]*):)?autoFilter\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?autoFilter>|<(?:(?<prefixEmpty>[A-Za-z_][\w.-]*):)?autoFilter\b[^/]*\/>/g

export function updateCustomSheetViewsXml(
	preservedXml: string,
	filters: readonly SheetAdvancedFilterInfo[],
): string {
	if (filters.length === 0) return preservedXml
	let index = 0
	return preservedXml.replace(CUSTOM_SHEET_VIEW_RE, (match, ...args: unknown[]) => {
		const groups = args.at(-1) as
			| { readonly prefix?: string; readonly attrs?: string; readonly body?: string }
			| undefined
		const filter = filters[index++]
		if (!filter?.autoFilter) return match
		const prefix = groups?.prefix
		const tagName = prefix ? `${prefix}:customSheetView` : 'customSheetView'
		const body = groups?.body ?? ''
		const autoFilterXml = buildAutoFilterXml(
			filter.autoFilter,
			prefix ? { tagPrefix: prefix } : undefined,
		)
		const updatedBody = replaceOrInsertAutoFilter(body, autoFilterXml)
		return `<${tagName}${groups?.attrs ?? ''}>${updatedBody}</${tagName}>`
	})
}

export function buildCustomSheetViewsXml(filters: readonly SheetAdvancedFilterInfo[]): string {
	const parts = ['<customSheetViews>']
	let count = 0
	for (const filter of filters) {
		const autoFilter = filter.autoFilter
		if (!autoFilter) continue
		count++
		const attrs: string[] = []
		if (filter.viewName) attrs.push(`name="${escapeXml(filter.viewName)}"`)
		if (filter.guid) attrs.push(`guid="${escapeXml(filter.guid)}"`)
		parts.push(`<customSheetView${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}>`)
		parts.push(buildAutoFilterXml(autoFilter))
		parts.push('</customSheetView>')
	}
	parts.push('</customSheetViews>')
	return count === 0 ? '' : parts.join('')
}

function replaceOrInsertAutoFilter(body: string, autoFilterXml: string): string {
	let replaced = false
	const updated = body.replace(AUTO_FILTER_RE, () => {
		if (replaced) return ''
		replaced = true
		return autoFilterXml
	})
	if (replaced) return updated
	return `${autoFilterXml}${body}`
}
