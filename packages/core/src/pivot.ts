export interface PivotTableInfo {
	readonly partPath: string
	readonly sheetName: string
	readonly name?: string
	readonly cacheId?: number
	readonly locationRef?: string
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
}

export interface SlicerCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
}

export interface SlicerInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}
