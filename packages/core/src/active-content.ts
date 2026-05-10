export type ActiveContentKind =
	| 'vbaProject'
	| 'activeX'
	| 'formControl'
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
	readonly vbaProject?: VbaProjectInfo
}

export function cloneActiveContentInfo(entry: ActiveContentInfo): ActiveContentInfo {
	return {
		...entry,
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
