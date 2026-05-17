export type WorkbookConnectionPartKind = 'queryTable' | 'connection' | 'powerQueryMashup'

export interface WorkbookConnectionPartInfo {
	readonly kind: WorkbookConnectionPartKind
	readonly partPath: string
	readonly contentType: string
	readonly relType?: string
	readonly sheetName?: string
	readonly relationshipCount: number
	readonly name?: string
	readonly connectionId?: number
	readonly connectionType?: number
	readonly description?: string
	readonly deleted?: boolean
	readonly backgroundRefresh?: boolean
	readonly keepAlive?: boolean
	readonly refreshInterval?: number
	readonly refreshOnLoad?: boolean
	readonly saveData?: boolean
	readonly savePassword?: boolean
	readonly refreshedVersion?: number
	readonly refreshedDateIso?: string
	readonly minRefreshableVersion?: number
	readonly credentials?: string
	readonly singleSignOnId?: string
	readonly sourceFile?: string
	readonly odcFile?: string
	readonly onlyUseConnectionFile?: boolean
	readonly command?: string
	readonly commandType?: number
	readonly serverCommand?: boolean
	readonly hasConnectionString?: boolean
}
