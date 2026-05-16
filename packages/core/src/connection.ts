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
	readonly sourceFile?: string
	readonly odcFile?: string
	readonly onlyUseConnectionFile?: boolean
	readonly command?: string
	readonly hasConnectionString?: boolean
}
