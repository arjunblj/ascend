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

export function pushAutoFilterXml(out: XmlSink, autoFilter: AutoFilter): void {
	out.push(`<autoFilter ref="${escapeXml(autoFilter.ref)}">`)
	for (const column of autoFilter.columns) {
		pushFilterColumnXml(out, column)
	}
	if (autoFilter.sortState) pushSortStateXml(out, autoFilter.sortState)
	out.push('</autoFilter>')
}

function pushFilterColumnXml(out: XmlSink, column: FilterColumn): void {
	const attrs = [`colId="${column.colId}"`]
	if (column.hiddenButton !== undefined)
		attrs.push(`hiddenButton="${column.hiddenButton ? '1' : '0'}"`)
	if (column.showButton !== undefined) attrs.push(`showButton="${column.showButton ? '1' : '0'}"`)

	switch (column.kind) {
		case 'filters': {
			out.push(`<filterColumn ${attrs.join(' ')}>`)
			const filterAttrs: string[] = []
			if (column.blank !== undefined) filterAttrs.push(`blank="${column.blank ? '1' : '0'}"`)
			if (column.calendarType) filterAttrs.push(`calendarType="${escapeXml(column.calendarType)}"`)
			out.push(`<filters${filterAttrs.length > 0 ? ` ${filterAttrs.join(' ')}` : ''}>`)
			for (const value of column.values ?? []) {
				out.push(`<filter val="${escapeXml(value)}"/>`)
			}
			for (const item of column.dateGroupItems ?? []) {
				pushDateGroupItemXml(out, item)
			}
			out.push('</filters></filterColumn>')
			return
		}
		case 'customFilters': {
			const innerAttrs: string[] = []
			if (column.and !== undefined) innerAttrs.push(`and="${column.and ? '1' : '0'}"`)
			out.push(
				`<filterColumn ${attrs.join(' ')}><customFilters${innerAttrs.length > 0 ? ` ${innerAttrs.join(' ')}` : ''}>`,
			)
			for (const filter of column.customFilters ?? []) {
				const customAttrs = [`val="${escapeXml(filter.val)}"`]
				if (filter.operator) customAttrs.push(`operator="${escapeXml(filter.operator)}"`)
				out.push(`<customFilter ${customAttrs.join(' ')}/>`)
			}
			out.push('</customFilters></filterColumn>')
			return
		}
		case 'dynamicFilter': {
			const innerAttrs: string[] = []
			if (column.dynamicFilterType) innerAttrs.push(`type="${escapeXml(column.dynamicFilterType)}"`)
			if (column.dynamicFilterVal !== undefined) innerAttrs.push(`val="${column.dynamicFilterVal}"`)
			if (column.dynamicFilterMaxVal !== undefined) {
				innerAttrs.push(`maxVal="${column.dynamicFilterMaxVal}"`)
			}
			out.push(
				`<filterColumn ${attrs.join(' ')}><dynamicFilter ${innerAttrs.join(' ')}/></filterColumn>`,
			)
			return
		}
		case 'top10': {
			const innerAttrs: string[] = []
			if (column.top !== undefined) innerAttrs.push(`top="${column.top ? '1' : '0'}"`)
			if (column.percent !== undefined) innerAttrs.push(`percent="${column.percent ? '1' : '0'}"`)
			if (column.val !== undefined) innerAttrs.push(`val="${column.val}"`)
			if (column.filterVal !== undefined) innerAttrs.push(`filterVal="${column.filterVal}"`)
			out.push(`<filterColumn ${attrs.join(' ')}><top10 ${innerAttrs.join(' ')}/></filterColumn>`)
			return
		}
		case 'colorFilter': {
			const innerAttrs: string[] = []
			if (column.dxfId !== undefined) innerAttrs.push(`dxfId="${column.dxfId}"`)
			if (column.cellColor !== undefined)
				innerAttrs.push(`cellColor="${column.cellColor ? '1' : '0'}"`)
			out.push(
				`<filterColumn ${attrs.join(' ')}><colorFilter ${innerAttrs.join(' ')}/></filterColumn>`,
			)
			return
		}
		case 'iconFilter': {
			const innerAttrs: string[] = []
			if (column.iconSet) innerAttrs.push(`iconSet="${escapeXml(column.iconSet)}"`)
			if (column.iconId !== undefined) innerAttrs.push(`iconId="${column.iconId}"`)
			out.push(
				`<filterColumn ${attrs.join(' ')}><iconFilter ${innerAttrs.join(' ')}/></filterColumn>`,
			)
			return
		}
	}
}

function pushDateGroupItemXml(out: XmlSink, item: FilterDateGroupItem): void {
	const attrs: string[] = []
	if (item.year !== undefined) attrs.push(`year="${item.year}"`)
	if (item.month !== undefined) attrs.push(`month="${item.month}"`)
	if (item.day !== undefined) attrs.push(`day="${item.day}"`)
	if (item.hour !== undefined) attrs.push(`hour="${item.hour}"`)
	if (item.minute !== undefined) attrs.push(`minute="${item.minute}"`)
	if (item.second !== undefined) attrs.push(`second="${item.second}"`)
	if (item.dateTimeGrouping) attrs.push(`dateTimeGrouping="${escapeXml(item.dateTimeGrouping)}"`)
	out.push(`<dateGroupItem ${attrs.join(' ')}/>`)
}

export function pushSortStateXml(out: XmlSink, sortState: SortState): void {
	const attrs = [`ref="${escapeXml(sortState.ref)}"`]
	if (sortState.caseSensitive !== undefined) {
		attrs.push(`caseSensitive="${sortState.caseSensitive ? '1' : '0'}"`)
	}
	if (sortState.columnSort !== undefined) {
		attrs.push(`columnSort="${sortState.columnSort ? '1' : '0'}"`)
	}
	if (sortState.sortMethod) attrs.push(`sortMethod="${escapeXml(sortState.sortMethod)}"`)
	out.push(`<sortState ${attrs.join(' ')}>`)
	for (const condition of sortState.conditions) {
		pushSortConditionXml(out, condition)
	}
	out.push('</sortState>')
}

function pushSortConditionXml(out: XmlSink, condition: SortCondition): void {
	const attrs = [`ref="${escapeXml(condition.ref)}"`]
	if (condition.descending !== undefined)
		attrs.push(`descending="${condition.descending ? '1' : '0'}"`)
	if (condition.sortBy) attrs.push(`sortBy="${escapeXml(condition.sortBy)}"`)
	if (condition.customList) attrs.push(`customList="${escapeXml(condition.customList)}"`)
	if (condition.dxfId !== undefined) attrs.push(`dxfId="${condition.dxfId}"`)
	if (condition.iconSet) attrs.push(`iconSet="${escapeXml(condition.iconSet)}"`)
	if (condition.iconId !== undefined) attrs.push(`iconId="${condition.iconId}"`)
	out.push(`<sortCondition ${attrs.join(' ')}/>`)
}
