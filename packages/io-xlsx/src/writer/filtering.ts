import type {
	AutoFilter,
	FilterColumn,
	FilterDateGroupItem,
	SortCondition,
	SortState,
} from '@ascend/core'
import { escapeXml } from '../xml.ts'

interface XmlSink {
	push(s: string): void
}

interface FilterXmlOptions {
	readonly tagPrefix?: string
}

export function buildAutoFilterXml(autoFilter: AutoFilter, options?: FilterXmlOptions): string {
	const parts: string[] = []
	pushAutoFilterXml({ push: (s) => parts.push(s) }, autoFilter, options)
	return parts.join('')
}

export function pushAutoFilterXml(
	out: XmlSink,
	autoFilter: AutoFilter,
	options?: FilterXmlOptions,
): void {
	const tag = tagBuilder(options?.tagPrefix)
	out.push(`<${tag('autoFilter')} ref="${escapeXml(autoFilter.ref)}">`)
	for (const column of autoFilter.columns) {
		pushFilterColumnXml(out, column, tag)
	}
	if (autoFilter.sortState) pushSortStateXml(out, autoFilter.sortState, options)
	out.push(`</${tag('autoFilter')}>`)
}

function pushFilterColumnXml(out: XmlSink, column: FilterColumn, tag: TagBuilder): void {
	const attrs = [`colId="${column.colId}"`]
	if (column.hiddenButton !== undefined)
		attrs.push(`hiddenButton="${column.hiddenButton ? '1' : '0'}"`)
	if (column.showButton !== undefined) attrs.push(`showButton="${column.showButton ? '1' : '0'}"`)

	switch (column.kind) {
		case 'filters': {
			out.push(`<${tag('filterColumn')} ${attrs.join(' ')}>`)
			const filterAttrs: string[] = []
			if (column.blank !== undefined) filterAttrs.push(`blank="${column.blank ? '1' : '0'}"`)
			if (column.calendarType) filterAttrs.push(`calendarType="${escapeXml(column.calendarType)}"`)
			out.push(`<${tag('filters')}${filterAttrs.length > 0 ? ` ${filterAttrs.join(' ')}` : ''}>`)
			for (const value of column.values ?? []) {
				out.push(`<${tag('filter')} val="${escapeXml(value)}"/>`)
			}
			for (const item of column.dateGroupItems ?? []) {
				pushDateGroupItemXml(out, item, tag)
			}
			out.push(`</${tag('filters')}></${tag('filterColumn')}>`)
			return
		}
		case 'customFilters': {
			const innerAttrs: string[] = []
			if (column.and !== undefined) innerAttrs.push(`and="${column.and ? '1' : '0'}"`)
			out.push(
				`<${tag('filterColumn')} ${attrs.join(' ')}><${tag('customFilters')}${innerAttrs.length > 0 ? ` ${innerAttrs.join(' ')}` : ''}>`,
			)
			for (const filter of column.customFilters ?? []) {
				const customAttrs = [`val="${escapeXml(filter.val)}"`]
				if (filter.operator) customAttrs.push(`operator="${escapeXml(filter.operator)}"`)
				out.push(`<${tag('customFilter')} ${customAttrs.join(' ')}/>`)
			}
			out.push(`</${tag('customFilters')}></${tag('filterColumn')}>`)
			return
		}
		case 'dynamicFilter': {
			const innerAttrs: string[] = []
			if (column.dynamicFilterType) innerAttrs.push(`type="${escapeXml(column.dynamicFilterType)}"`)
			if (column.dynamicFilterVal !== undefined) innerAttrs.push(`val="${column.dynamicFilterVal}"`)
			if (column.dynamicFilterMaxVal !== undefined) {
				innerAttrs.push(`maxVal="${column.dynamicFilterMaxVal}"`)
			}
			if (column.dynamicFilterValIso) {
				innerAttrs.push(`valIso="${escapeXml(column.dynamicFilterValIso)}"`)
			}
			if (column.dynamicFilterMaxValIso) {
				innerAttrs.push(`maxValIso="${escapeXml(column.dynamicFilterMaxValIso)}"`)
			}
			out.push(
				`<${tag('filterColumn')} ${attrs.join(' ')}><${tag('dynamicFilter')} ${innerAttrs.join(' ')}/></${tag('filterColumn')}>`,
			)
			return
		}
		case 'top10': {
			const innerAttrs: string[] = []
			if (column.top !== undefined) innerAttrs.push(`top="${column.top ? '1' : '0'}"`)
			if (column.percent !== undefined) innerAttrs.push(`percent="${column.percent ? '1' : '0'}"`)
			if (column.val !== undefined) innerAttrs.push(`val="${column.val}"`)
			if (column.filterVal !== undefined) innerAttrs.push(`filterVal="${column.filterVal}"`)
			out.push(
				`<${tag('filterColumn')} ${attrs.join(' ')}><${tag('top10')} ${innerAttrs.join(' ')}/></${tag('filterColumn')}>`,
			)
			return
		}
		case 'colorFilter': {
			const innerAttrs: string[] = []
			if (column.dxfId !== undefined) innerAttrs.push(`dxfId="${column.dxfId}"`)
			if (column.cellColor !== undefined)
				innerAttrs.push(`cellColor="${column.cellColor ? '1' : '0'}"`)
			out.push(
				`<${tag('filterColumn')} ${attrs.join(' ')}><${tag('colorFilter')} ${innerAttrs.join(' ')}/></${tag('filterColumn')}>`,
			)
			return
		}
		case 'iconFilter': {
			const innerAttrs: string[] = []
			if (column.iconSet) innerAttrs.push(`iconSet="${escapeXml(column.iconSet)}"`)
			if (column.iconId !== undefined) innerAttrs.push(`iconId="${column.iconId}"`)
			out.push(
				`<${tag('filterColumn')} ${attrs.join(' ')}><${tag('iconFilter')} ${innerAttrs.join(' ')}/></${tag('filterColumn')}>`,
			)
			return
		}
	}
}

function pushDateGroupItemXml(out: XmlSink, item: FilterDateGroupItem, tag: TagBuilder): void {
	const attrs: string[] = []
	if (item.year !== undefined) attrs.push(`year="${item.year}"`)
	if (item.month !== undefined) attrs.push(`month="${item.month}"`)
	if (item.day !== undefined) attrs.push(`day="${item.day}"`)
	if (item.hour !== undefined) attrs.push(`hour="${item.hour}"`)
	if (item.minute !== undefined) attrs.push(`minute="${item.minute}"`)
	if (item.second !== undefined) attrs.push(`second="${item.second}"`)
	if (item.dateTimeGrouping) attrs.push(`dateTimeGrouping="${escapeXml(item.dateTimeGrouping)}"`)
	out.push(`<${tag('dateGroupItem')} ${attrs.join(' ')}/>`)
}

export function pushSortStateXml(
	out: XmlSink,
	sortState: SortState,
	options?: FilterXmlOptions,
): void {
	const tag = tagBuilder(options?.tagPrefix)
	const attrs = [`ref="${escapeXml(sortState.ref)}"`]
	if (sortState.caseSensitive !== undefined) {
		attrs.push(`caseSensitive="${sortState.caseSensitive ? '1' : '0'}"`)
	}
	if (sortState.columnSort !== undefined) {
		attrs.push(`columnSort="${sortState.columnSort ? '1' : '0'}"`)
	}
	if (sortState.sortMethod) attrs.push(`sortMethod="${escapeXml(sortState.sortMethod)}"`)
	out.push(`<${tag('sortState')} ${attrs.join(' ')}>`)
	for (const condition of sortState.conditions) {
		pushSortConditionXml(out, condition, tag)
	}
	out.push(`</${tag('sortState')}>`)
}

function pushSortConditionXml(out: XmlSink, condition: SortCondition, tag: TagBuilder): void {
	const attrs = [`ref="${escapeXml(condition.ref)}"`]
	if (condition.descending !== undefined)
		attrs.push(`descending="${condition.descending ? '1' : '0'}"`)
	if (condition.sortBy) attrs.push(`sortBy="${escapeXml(condition.sortBy)}"`)
	if (condition.customList) attrs.push(`customList="${escapeXml(condition.customList)}"`)
	if (condition.dxfId !== undefined) attrs.push(`dxfId="${condition.dxfId}"`)
	if (condition.iconSet) attrs.push(`iconSet="${escapeXml(condition.iconSet)}"`)
	if (condition.iconId !== undefined) attrs.push(`iconId="${condition.iconId}"`)
	out.push(`<${tag('sortCondition')} ${attrs.join(' ')}/>`)
}

type TagBuilder = (name: string) => string

function tagBuilder(prefix: string | undefined): TagBuilder {
	return prefix ? (name) => `${prefix}:${name}` : (name) => name
}
