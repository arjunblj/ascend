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
	readonly sharedItems?: readonly PivotCacheSharedItemInfo[]
}

export interface PivotCacheSharedItemInfo {
	readonly index: number
	readonly kind: 'string' | 'number' | 'date' | 'boolean' | 'error' | 'missing'
	readonly value?: string
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
}

export interface TimelineInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}
