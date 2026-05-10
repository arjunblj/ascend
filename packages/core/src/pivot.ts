export interface PivotTableInfo {
	readonly partPath: string
	readonly sheetName: string
	readonly name?: string
	readonly cacheId?: number
	readonly locationRef?: string
	readonly fields: readonly PivotFieldInfo[]
	readonly rowFields: readonly PivotFieldReference[]
	readonly columnFields: readonly PivotFieldReference[]
	readonly pageFields: readonly PivotFieldReference[]
	readonly dataFields: readonly PivotDataFieldInfo[]
}

export interface PivotCacheInfo {
	readonly partPath: string
	readonly cacheId?: number
	readonly relId?: string
	readonly recordCount?: number
	readonly refreshedVersion?: number
	readonly minRefreshableVersion?: number
	readonly createdVersion?: number
	readonly refreshedBy?: string
	readonly refreshedDate?: number
	readonly refreshOnLoad?: boolean
	readonly enableRefresh?: boolean
	readonly invalid?: boolean
	readonly saveData?: boolean
	readonly optimizeMemory?: boolean
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly recordsPartPath?: string
	readonly fields: readonly PivotCacheFieldInfo[]
}

export interface PivotCacheFieldInfo {
	readonly index: number
	readonly name?: string
	readonly databaseField?: boolean
	readonly numFmtId?: number
	readonly formula?: string
	readonly sharedItemsInfo?: PivotCacheSharedItemsInfo
	readonly sharedItems?: readonly PivotCacheSharedItemInfo[]
	readonly fieldGroup?: PivotCacheFieldGroupInfo
}

export interface PivotCacheSharedItemInfo {
	readonly index: number
	readonly kind: 'string' | 'number' | 'date' | 'boolean' | 'error' | 'missing'
	readonly value?: string
}

export interface PivotCacheSharedItemsInfo {
	readonly count?: number
	readonly containsBlank?: boolean
	readonly containsDate?: boolean
	readonly containsNonDate?: boolean
	readonly containsNumber?: boolean
	readonly containsInteger?: boolean
	readonly containsString?: boolean
	readonly containsMixedTypes?: boolean
	readonly containsSemiMixedTypes?: boolean
	readonly minValue?: number
	readonly maxValue?: number
	readonly minDate?: string
	readonly maxDate?: string
}

export interface PivotCacheFieldGroupInfo {
	readonly base?: number
	readonly parent?: number
	readonly discreteItems?: readonly PivotCacheFieldGroupDiscreteItemInfo[]
	readonly groupItems?: readonly PivotCacheSharedItemInfo[]
}

export interface PivotCacheFieldGroupDiscreteItemInfo {
	readonly index: number
	readonly value?: number
}

export interface PivotFieldInfo {
	readonly index: number
	readonly axis?: string
	readonly name?: string
	readonly numFmtId?: number
	readonly hidden?: boolean
	readonly dataField?: boolean
	readonly defaultSubtotal?: boolean
	readonly showAll?: boolean
	readonly multipleItemSelectionAllowed?: boolean
	readonly items?: readonly PivotFieldItemInfo[]
}

export interface PivotFieldItemInfo {
	readonly index: number
	readonly cacheIndex?: number
	readonly itemType?: string
	readonly caption?: string
	readonly hidden?: boolean
	readonly manualFilter?: boolean
	readonly showDetails?: boolean
	readonly calculated?: boolean
	readonly missing?: boolean
	readonly childItems?: boolean
	readonly expanded?: boolean
	readonly drillAcrossAttributes?: boolean
}

export interface PivotFieldReference {
	readonly index: number
	readonly name?: string
	readonly item?: number
	readonly hierarchy?: number
	readonly caption?: string
}

export interface PivotDataFieldInfo {
	readonly fieldIndex: number
	readonly name?: string
	readonly subtotal?: string
	readonly numFmtId?: number
}

export function clonePivotCacheInfo(entry: PivotCacheInfo): PivotCacheInfo {
	return {
		...entry,
		fields: entry.fields.map((field) => ({
			...field,
			...(field.sharedItemsInfo ? { sharedItemsInfo: { ...field.sharedItemsInfo } } : {}),
			...(field.sharedItems ? { sharedItems: field.sharedItems.map((item) => ({ ...item })) } : {}),
			...(field.fieldGroup
				? {
						fieldGroup: {
							...field.fieldGroup,
							...(field.fieldGroup.discreteItems
								? { discreteItems: field.fieldGroup.discreteItems.map((item) => ({ ...item })) }
								: {}),
							...(field.fieldGroup.groupItems
								? { groupItems: field.fieldGroup.groupItems.map((item) => ({ ...item })) }
								: {}),
						},
					}
				: {}),
		})),
	}
}

export function clonePivotTableInfo(entry: PivotTableInfo): PivotTableInfo {
	return {
		...entry,
		fields: entry.fields.map((field) => ({
			...field,
			...(field.items ? { items: field.items.map((item) => ({ ...item })) } : {}),
		})),
		rowFields: entry.rowFields.map((field) => ({ ...field })),
		columnFields: entry.columnFields.map((field) => ({ ...field })),
		pageFields: entry.pageFields.map((field) => ({ ...field })),
		dataFields: entry.dataFields.map((field) => ({ ...field })),
	}
}

export interface SlicerCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly items?: readonly SlicerCacheItemInfo[]
}

export interface SlicerCacheItemInfo {
	readonly index: number
	readonly selected?: boolean
	readonly noData?: boolean
}

export interface SlicerInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface TimelineCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly state?: TimelineStateInfo
}

export interface TimelineInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface TimelineRangeInfo {
	readonly startDate: string
	readonly endDate: string
}

export interface TimelineStateInfo {
	readonly filterType?: string
	readonly filterId?: number
	readonly filterPivotName?: string
	readonly filterTabId?: number
	readonly lastRefreshVersion?: number
	readonly minimalRefreshVersion?: number
	readonly pivotCacheId?: number
	readonly singleRangeFilterState?: boolean
	readonly selection?: TimelineRangeInfo
	readonly bounds?: TimelineRangeInfo
}
