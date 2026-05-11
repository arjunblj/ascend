import type {
	AutoFilter,
	CustomFilter,
	FilterColumn,
	FilterDateGroupItem,
	SortCondition,
	SortState,
} from '@ascend/core'
import { asArray, attr, boolAttr, numAttr, type XmlNode } from '../xml.ts'

interface ParseSortStateOptions {
	readonly preserveAttributes?: boolean
}

export function parseAutoFilterNode(node: XmlNode | undefined): AutoFilter | null {
	if (!node) return null
	const ref = attr(node, 'ref')
	if (!ref) return null

	const columns = asArray<XmlNode>(node.filterColumn as XmlNode | XmlNode[]).map(parseFilterColumn)
	const parsed: {
		ref: string
		uid?: string
		columns: readonly FilterColumn[]
		sortState?: SortState
	} = {
		ref,
		columns,
	}
	const uid = attr(node, 'xr:uid')
	if (uid) parsed.uid = uid
	const sortState = parseSortStateNode(node.sortState as XmlNode | undefined)
	if (sortState) parsed.sortState = sortState
	return parsed as AutoFilter
}

function parseFilterColumn(node: XmlNode): FilterColumn {
	const colId = numAttr(node, 'colId') ?? 0
	const parsed: {
		colId: number
		hiddenButton?: boolean
		showButton?: boolean
		kind: 'filters' | 'customFilters' | 'dynamicFilter' | 'top10' | 'colorFilter' | 'iconFilter'
		values?: readonly string[]
		blank?: boolean
		calendarType?: string
		dateGroupItems?: readonly FilterDateGroupItem[]
		customFilters?: readonly CustomFilter[]
		and?: boolean
		dynamicFilterType?: string
		dynamicFilterVal?: number
		dynamicFilterMaxVal?: number
		dynamicFilterValIso?: string
		dynamicFilterMaxValIso?: string
		top?: boolean
		percent?: boolean
		val?: number
		filterVal?: number
		dxfId?: number
		cellColor?: boolean
		iconSet?: string
		iconId?: number
	} = {
		colId,
		kind: 'filters',
	}
	const hiddenButton = boolAttr(node, 'hiddenButton')
	if (hiddenButton !== undefined) parsed.hiddenButton = hiddenButton
	const showButton = boolAttr(node, 'showButton')
	if (showButton !== undefined) parsed.showButton = showButton

	const filters = node.filters as XmlNode | undefined
	if (filters) {
		parsed.kind = 'filters'
		const values = asArray<XmlNode>(filters.filter as XmlNode | XmlNode[]).map(
			(filter) => attr(filter, 'val') ?? '',
		)
		if (values.length > 0) parsed.values = values
		const blank = boolAttr(filters, 'blank')
		if (blank !== undefined) parsed.blank = blank
		const calendarType = attr(filters, 'calendarType')
		if (calendarType) parsed.calendarType = calendarType
		const dateGroupItems = asArray<XmlNode>(filters.dateGroupItem as XmlNode | XmlNode[]).map(
			parseDateGroupItem,
		)
		if (dateGroupItems.length > 0) parsed.dateGroupItems = dateGroupItems
		return parsed as FilterColumn
	}

	const customFiltersNode = node.customFilters as XmlNode | undefined
	if (customFiltersNode) {
		parsed.kind = 'customFilters'
		const and = boolAttr(customFiltersNode, 'and')
		if (and !== undefined) parsed.and = and
		parsed.customFilters = asArray<XmlNode>(
			customFiltersNode.customFilter as XmlNode | XmlNode[],
		).map((customFilter) => {
			const entry: { operator?: string; val: string } = {
				val: attr(customFilter, 'val') ?? '',
			}
			const operator = attr(customFilter, 'operator')
			if (operator) entry.operator = operator
			return entry as CustomFilter
		})
		return parsed as FilterColumn
	}

	const dynamicFilter = node.dynamicFilter as XmlNode | undefined
	if (dynamicFilter) {
		parsed.kind = 'dynamicFilter'
		const type = attr(dynamicFilter, 'type')
		if (type) parsed.dynamicFilterType = type
		const val = numAttr(dynamicFilter, 'val')
		if (val !== undefined) parsed.dynamicFilterVal = val
		const maxVal = numAttr(dynamicFilter, 'maxVal')
		if (maxVal !== undefined) parsed.dynamicFilterMaxVal = maxVal
		const valIso = attr(dynamicFilter, 'valIso')
		if (valIso) parsed.dynamicFilterValIso = valIso
		const maxValIso = attr(dynamicFilter, 'maxValIso')
		if (maxValIso) parsed.dynamicFilterMaxValIso = maxValIso
		return parsed as FilterColumn
	}

	const top10 = node.top10 as XmlNode | undefined
	if (top10) {
		parsed.kind = 'top10'
		const top = boolAttr(top10, 'top')
		if (top !== undefined) parsed.top = top
		const percent = boolAttr(top10, 'percent')
		if (percent !== undefined) parsed.percent = percent
		const val = numAttr(top10, 'val')
		if (val !== undefined) parsed.val = val
		const filterVal = numAttr(top10, 'filterVal')
		if (filterVal !== undefined) parsed.filterVal = filterVal
		return parsed as FilterColumn
	}

	const colorFilter = node.colorFilter as XmlNode | undefined
	if (colorFilter) {
		parsed.kind = 'colorFilter'
		const dxfId = numAttr(colorFilter, 'dxfId')
		if (dxfId !== undefined) parsed.dxfId = dxfId
		const cellColor = boolAttr(colorFilter, 'cellColor')
		if (cellColor !== undefined) parsed.cellColor = cellColor
		return parsed as FilterColumn
	}

	const iconFilter = node.iconFilter as XmlNode | undefined
	if (iconFilter) {
		parsed.kind = 'iconFilter'
		const iconSet = attr(iconFilter, 'iconSet')
		if (iconSet) parsed.iconSet = iconSet
		const iconId = numAttr(iconFilter, 'iconId')
		if (iconId !== undefined) parsed.iconId = iconId
		return parsed as FilterColumn
	}

	return parsed as FilterColumn
}

function parseDateGroupItem(node: XmlNode): FilterDateGroupItem {
	const parsed: {
		year?: number
		month?: number
		day?: number
		hour?: number
		minute?: number
		second?: number
		dateTimeGrouping?: string
	} = {}
	const year = numAttr(node, 'year')
	if (year !== undefined) parsed.year = year
	const month = numAttr(node, 'month')
	if (month !== undefined) parsed.month = month
	const day = numAttr(node, 'day')
	if (day !== undefined) parsed.day = day
	const hour = numAttr(node, 'hour')
	if (hour !== undefined) parsed.hour = hour
	const minute = numAttr(node, 'minute')
	if (minute !== undefined) parsed.minute = minute
	const second = numAttr(node, 'second')
	if (second !== undefined) parsed.second = second
	const dateTimeGrouping = attr(node, 'dateTimeGrouping')
	if (dateTimeGrouping) parsed.dateTimeGrouping = dateTimeGrouping
	return parsed as FilterDateGroupItem
}

export function parseSortStateNode(
	node: XmlNode | undefined,
	options: ParseSortStateOptions = {},
): SortState | null {
	if (!node) return null
	const ref = attr(node, 'ref')
	if (!ref) return null
	const parsed: {
		ref: string
		caseSensitive?: boolean
		columnSort?: boolean
		sortMethod?: string
		preservedAttributes?: Readonly<Record<string, string>>
		conditions: readonly SortCondition[]
	} = {
		ref,
		conditions: asArray<XmlNode>(node.sortCondition as XmlNode | XmlNode[]).map((condition) => {
			const entry: {
				ref: string
				descending?: boolean
				sortBy?: string
				customList?: string
				dxfId?: number
				iconSet?: string
				iconId?: number
			} = {
				ref: attr(condition, 'ref') ?? '',
			}
			const descending = boolAttr(condition, 'descending')
			if (descending !== undefined) entry.descending = descending
			const sortBy = attr(condition, 'sortBy')
			if (sortBy) entry.sortBy = sortBy
			const customList = attr(condition, 'customList')
			if (customList) entry.customList = customList
			const dxfId = numAttr(condition, 'dxfId')
			if (dxfId !== undefined) entry.dxfId = dxfId
			const iconSet = attr(condition, 'iconSet')
			if (iconSet) entry.iconSet = iconSet
			const iconId = numAttr(condition, 'iconId')
			if (iconId !== undefined) entry.iconId = iconId
			return entry as SortCondition
		}),
	}
	const caseSensitive = boolAttr(node, 'caseSensitive')
	if (caseSensitive !== undefined) parsed.caseSensitive = caseSensitive
	const columnSort = boolAttr(node, 'columnSort')
	if (columnSort !== undefined) parsed.columnSort = columnSort
	const sortMethod = attr(node, 'sortMethod')
	if (sortMethod) parsed.sortMethod = sortMethod
	if (options.preserveAttributes) {
		const preservedAttributes = sortStatePreservedAttributes(node)
		if (Object.keys(preservedAttributes).length > 0) {
			parsed.preservedAttributes = preservedAttributes
		}
	}
	return parsed as SortState
}

function sortStatePreservedAttributes(node: XmlNode): Record<string, string> {
	const attrs: Record<string, string> = {}
	for (const [key, value] of Object.entries(node)) {
		if (!key.startsWith('@_') || value === undefined || value === null) continue
		const name = key.slice(2)
		if (
			name === 'ref' ||
			name === 'caseSensitive' ||
			name === 'columnSort' ||
			name === 'sortMethod'
		) {
			continue
		}
		attrs[name] = String(value)
	}
	return attrs
}
