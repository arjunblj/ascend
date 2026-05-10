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
}

export interface PivotFieldInfo {
	readonly index: number
	readonly axis?: string
	readonly name?: string
	readonly hidden?: boolean
	readonly dataField?: boolean
	readonly defaultSubtotal?: boolean
	readonly showAll?: boolean
}

export interface PivotFieldReference {
	readonly index: number
	readonly name?: string
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
