export type ActiveContentKind =
	| 'vbaProject'
	| 'activeX'
	| 'formControl'
	| 'vbaSignature'
	| 'digitalSignature'
	| 'customUi'
	| 'unknownActiveContent'

export interface ActiveContentInfo {
	readonly kind: ActiveContentKind
	readonly partPath: string
	readonly contentType: string
	readonly anchor: 'workbook' | 'sheet'
	readonly sheetName?: string
	readonly relType?: string
	readonly relationshipCount: number
	readonly byteSize?: number
	readonly opaque?: boolean
	readonly executionPolicy?: 'blocked'
}
