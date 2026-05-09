export type WorkbookDataModelPartKind =
	| 'modelData'
	| 'modelTable'
	| 'modelRelationship'
	| 'modelMetadata'
	| 'unknownDataModel'

export interface WorkbookDataModelPartInfo {
	readonly kind: WorkbookDataModelPartKind
	readonly partPath: string
	readonly contentType: string
	readonly relType?: string
	readonly relationshipCount: number
}
