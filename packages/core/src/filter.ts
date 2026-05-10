export interface FilterDateGroupItem {
	readonly year?: number
	readonly month?: number
	readonly day?: number
	readonly hour?: number
	readonly minute?: number
	readonly second?: number
	readonly dateTimeGrouping?: string
}

export interface CustomFilter {
	readonly operator?: string
	readonly val: string
}

export interface FilterColumn {
	readonly colId: number
	readonly hiddenButton?: boolean
	readonly showButton?: boolean
	readonly kind:
		| 'filters'
		| 'customFilters'
		| 'dynamicFilter'
		| 'top10'
		| 'colorFilter'
		| 'iconFilter'
	readonly values?: readonly string[]
	readonly blank?: boolean
	readonly calendarType?: string
	readonly dateGroupItems?: readonly FilterDateGroupItem[]
	readonly customFilters?: readonly CustomFilter[]
	readonly and?: boolean
	readonly dynamicFilterType?: string
	readonly dynamicFilterVal?: number
	readonly dynamicFilterMaxVal?: number
	readonly dynamicFilterValIso?: string
	readonly dynamicFilterMaxValIso?: string
	readonly top?: boolean
	readonly percent?: boolean
	readonly val?: number
	readonly filterVal?: number
	readonly dxfId?: number
	readonly cellColor?: boolean
	readonly iconSet?: string
	readonly iconId?: number
}

export interface SortCondition {
	readonly ref: string
	readonly descending?: boolean
	readonly sortBy?: string
	readonly customList?: string
	readonly dxfId?: number
	readonly iconSet?: string
	readonly iconId?: number
}

export interface SortState {
	readonly ref: string
	readonly caseSensitive?: boolean
	readonly columnSort?: boolean
	readonly sortMethod?: string
	readonly conditions: readonly SortCondition[]
}

export interface AutoFilter {
	readonly ref: string
	readonly columns: readonly FilterColumn[]
	readonly sortState?: SortState
}
