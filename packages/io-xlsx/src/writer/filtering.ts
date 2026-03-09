import type {
	AutoFilter,
	FilterColumn,
	FilterDateGroupItem,
	SortCondition,
	SortState,
} from '@ascend/core'
import { escapeXml } from '../xml.ts'

export function autoFilterXml(autoFilter: AutoFilter): string {
	const parts: string[] = [`<autoFilter ref="${escapeXml(autoFilter.ref)}">`]
	for (const column of autoFilter.columns) {
		parts.push(filterColumnXml(column))
	}
	if (autoFilter.sortState) parts.push(sortStateXml(autoFilter.sortState))
	parts.push('</autoFilter>')
	return parts.join('')
}

function filterColumnXml(column: FilterColumn): string {
	const attrs = [`colId="${column.colId}"`]
	if (column.hiddenButton !== undefined)
		attrs.push(`hiddenButton="${column.hiddenButton ? '1' : '0'}"`)
	if (column.showButton !== undefined) attrs.push(`showButton="${column.showButton ? '1' : '0'}"`)

	switch (column.kind) {
		case 'filters': {
			const parts: string[] = [`<filterColumn ${attrs.join(' ')}>`]
			const filterAttrs: string[] = []
			if (column.blank !== undefined) filterAttrs.push(`blank="${column.blank ? '1' : '0'}"`)
			if (column.calendarType) filterAttrs.push(`calendarType="${escapeXml(column.calendarType)}"`)
			parts.push(`<filters${filterAttrs.length > 0 ? ` ${filterAttrs.join(' ')}` : ''}>`)
			for (const value of column.values ?? []) {
				parts.push(`<filter val="${escapeXml(value)}"/>`)
			}
			for (const item of column.dateGroupItems ?? []) {
				parts.push(dateGroupItemXml(item))
			}
			parts.push('</filters></filterColumn>')
			return parts.join('')
		}
		case 'customFilters': {
			const parts: string[] = [`<filterColumn ${attrs.join(' ')}><customFilters`]
			const innerAttrs: string[] = []
			if (column.and !== undefined) innerAttrs.push(`and="${column.and ? '1' : '0'}"`)
			parts[parts.length - 1] += innerAttrs.length > 0 ? ` ${innerAttrs.join(' ')}>` : '>'
			for (const filter of column.customFilters ?? []) {
				const customAttrs = [`val="${escapeXml(filter.val)}"`]
				if (filter.operator) customAttrs.push(`operator="${escapeXml(filter.operator)}"`)
				parts.push(`<customFilter ${customAttrs.join(' ')}/>`)
			}
			parts.push('</customFilters></filterColumn>')
			return parts.join('')
		}
		case 'dynamicFilter': {
			const innerAttrs: string[] = []
			if (column.dynamicFilterType) innerAttrs.push(`type="${escapeXml(column.dynamicFilterType)}"`)
			if (column.dynamicFilterVal !== undefined) innerAttrs.push(`val="${column.dynamicFilterVal}"`)
			if (column.dynamicFilterMaxVal !== undefined) {
				innerAttrs.push(`maxVal="${column.dynamicFilterMaxVal}"`)
			}
			return `<filterColumn ${attrs.join(' ')}><dynamicFilter ${innerAttrs.join(' ')}/></filterColumn>`
		}
		case 'top10': {
			const innerAttrs: string[] = []
			if (column.top !== undefined) innerAttrs.push(`top="${column.top ? '1' : '0'}"`)
			if (column.percent !== undefined) innerAttrs.push(`percent="${column.percent ? '1' : '0'}"`)
			if (column.val !== undefined) innerAttrs.push(`val="${column.val}"`)
			if (column.filterVal !== undefined) innerAttrs.push(`filterVal="${column.filterVal}"`)
			return `<filterColumn ${attrs.join(' ')}><top10 ${innerAttrs.join(' ')}/></filterColumn>`
		}
		case 'colorFilter': {
			const innerAttrs: string[] = []
			if (column.dxfId !== undefined) innerAttrs.push(`dxfId="${column.dxfId}"`)
			if (column.cellColor !== undefined)
				innerAttrs.push(`cellColor="${column.cellColor ? '1' : '0'}"`)
			return `<filterColumn ${attrs.join(' ')}><colorFilter ${innerAttrs.join(' ')}/></filterColumn>`
		}
		case 'iconFilter': {
			const innerAttrs: string[] = []
			if (column.iconSet) innerAttrs.push(`iconSet="${escapeXml(column.iconSet)}"`)
			if (column.iconId !== undefined) innerAttrs.push(`iconId="${column.iconId}"`)
			return `<filterColumn ${attrs.join(' ')}><iconFilter ${innerAttrs.join(' ')}/></filterColumn>`
		}
	}
}

function dateGroupItemXml(item: FilterDateGroupItem): string {
	const attrs: string[] = []
	if (item.year !== undefined) attrs.push(`year="${item.year}"`)
	if (item.month !== undefined) attrs.push(`month="${item.month}"`)
	if (item.day !== undefined) attrs.push(`day="${item.day}"`)
	if (item.hour !== undefined) attrs.push(`hour="${item.hour}"`)
	if (item.minute !== undefined) attrs.push(`minute="${item.minute}"`)
	if (item.second !== undefined) attrs.push(`second="${item.second}"`)
	if (item.dateTimeGrouping) attrs.push(`dateTimeGrouping="${escapeXml(item.dateTimeGrouping)}"`)
	return `<dateGroupItem ${attrs.join(' ')}/>`
}

export function sortStateXml(sortState: SortState): string {
	const attrs = [`ref="${escapeXml(sortState.ref)}"`]
	if (sortState.caseSensitive !== undefined) {
		attrs.push(`caseSensitive="${sortState.caseSensitive ? '1' : '0'}"`)
	}
	if (sortState.columnSort !== undefined) {
		attrs.push(`columnSort="${sortState.columnSort ? '1' : '0'}"`)
	}
	if (sortState.sortMethod) attrs.push(`sortMethod="${escapeXml(sortState.sortMethod)}"`)
	const parts: string[] = [`<sortState ${attrs.join(' ')}>`]
	for (const condition of sortState.conditions) {
		parts.push(sortConditionXml(condition))
	}
	parts.push('</sortState>')
	return parts.join('')
}

function sortConditionXml(condition: SortCondition): string {
	const attrs = [`ref="${escapeXml(condition.ref)}"`]
	if (condition.descending !== undefined)
		attrs.push(`descending="${condition.descending ? '1' : '0'}"`)
	if (condition.sortBy) attrs.push(`sortBy="${escapeXml(condition.sortBy)}"`)
	if (condition.customList) attrs.push(`customList="${escapeXml(condition.customList)}"`)
	if (condition.dxfId !== undefined) attrs.push(`dxfId="${condition.dxfId}"`)
	if (condition.iconSet) attrs.push(`iconSet="${escapeXml(condition.iconSet)}"`)
	if (condition.iconId !== undefined) attrs.push(`iconId="${condition.iconId}"`)
	return `<sortCondition ${attrs.join(' ')}/>`
}
