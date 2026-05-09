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
	readonly refreshOnLoad?: boolean
	readonly saveData?: boolean
	readonly refreshedVersion?: number
}
