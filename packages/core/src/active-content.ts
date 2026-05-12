import type { SheetImageAnchor } from './sheet.ts'

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

export interface CustomUiCallbackInfo {
	readonly attribute: string
	readonly macro: string
}

export interface CustomUiInfo {
	readonly namespaceUri?: string
	readonly callbackCount: number
	readonly callbacks: readonly CustomUiCallbackInfo[]
}

export interface WorksheetControlInfo {
	readonly shapeId?: number
	readonly name?: string
	readonly relationshipId?: string
	readonly controlPrRelationshipId?: string
	readonly controlPrRelationshipType?: string
	readonly controlPrTarget?: string
	readonly anchor?: SheetImageAnchor
	readonly vmlShapeId?: string
	readonly vmlShapeSpid?: string
	readonly vmlObjectType?: string
	readonly vmlMapOcx?: boolean
	readonly vmlImageRelationshipId?: string
	readonly vmlImageRelationshipType?: string
	readonly vmlImageTarget?: string
}

export interface ActiveContentInfo {
	readonly kind: ActiveContentKind
	readonly partPath: string
	readonly contentType: string
	readonly anchor: 'workbook' | 'sheet'
	readonly sheetName?: string
	readonly sourcePartPath?: string
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
	readonly customUi?: CustomUiInfo
	readonly worksheetControl?: WorksheetControlInfo
}

export function cloneActiveContentInfo(entry: ActiveContentInfo): ActiveContentInfo {
	return {
		...entry,
		...(entry.activeX ? { activeX: { ...entry.activeX } } : {}),
		...(entry.formControl ? { formControl: { ...entry.formControl } } : {}),
		...(entry.customUi
			? {
					customUi: {
						...entry.customUi,
						callbacks: entry.customUi.callbacks.map((callback) => ({ ...callback })),
					},
				}
			: {}),
		...(entry.worksheetControl
			? {
					worksheetControl: {
						...entry.worksheetControl,
						...(entry.worksheetControl.anchor
							? { anchor: cloneControlAnchor(entry.worksheetControl.anchor) }
							: {}),
					},
				}
			: {}),
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

function cloneControlAnchor(anchor: SheetImageAnchor): SheetImageAnchor {
	switch (anchor.kind) {
		case 'oneCell':
			return { ...anchor, from: { ...anchor.from } }
		case 'twoCell':
			return { ...anchor, from: { ...anchor.from }, to: { ...anchor.to } }
		case 'absolute':
			return { ...anchor }
	}
}
