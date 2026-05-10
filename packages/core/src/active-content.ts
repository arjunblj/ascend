export type ActiveContentKind =
	| 'vbaProject'
	| 'activeX'
	| 'formControl'
	| 'macroSheet'
	| 'vbaSignature'
	| 'digitalSignature'
	| 'customUi'
	| 'unknownActiveContent'

export type VbaModuleKind = 'document' | 'standard' | 'class' | 'designer'

export interface VbaModuleInfo {
	readonly name: string
	readonly kind: VbaModuleKind
}

export interface VbaProjectInfo {
	readonly moduleCount: number
	readonly modules: readonly VbaModuleInfo[]
	readonly projectStreamPresent: boolean
	readonly cfbDirectoryEntryCount?: number
}

export interface ActiveXControlInfo {
	readonly classId?: string
	readonly persistence?: string
	readonly relationshipId?: string
	readonly binaryRelationshipId?: string
	readonly binaryTarget?: string
}

export interface FormControlInfo {
	readonly objectType?: string
	readonly macro?: string
	readonly linkedCell?: string
	readonly listFillRange?: string
	readonly checked?: string
	readonly dropLines?: number
}

export interface ActiveContentInfo {
	readonly kind: ActiveContentKind
	readonly partPath: string
	readonly contentType: string
	readonly anchor: 'workbook' | 'sheet'
	readonly sheetName?: string
	readonly relType?: string
	readonly sourceRelationshipId?: string
	readonly relationshipCount: number
	readonly byteSize?: number
	readonly opaque?: boolean
	readonly executionPolicy?: 'blocked'
	readonly invalidationPolicy?: 'invalidatedByPackageEdit'
	readonly resigningPolicy?: 'notSupported'
	readonly vbaProject?: VbaProjectInfo
	readonly activeX?: ActiveXControlInfo
	readonly formControl?: FormControlInfo
}

export function cloneActiveContentInfo(entry: ActiveContentInfo): ActiveContentInfo {
	return {
		...entry,
		...(entry.activeX ? { activeX: { ...entry.activeX } } : {}),
		...(entry.formControl ? { formControl: { ...entry.formControl } } : {}),
		...(entry.vbaProject
			? {
					vbaProject: {
						...entry.vbaProject,
						modules: entry.vbaProject.modules.map((module) => ({ ...module })),
					},
				}
			: {}),
	}
}
