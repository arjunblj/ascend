import type { WorkbookDataModelPartInfo, WorkbookDataModelPartKind } from '@ascend/core'
import type { PreservationCapsule } from '../preserve.ts'

export function parseDataModelPartInfo(
	capsule: PreservationCapsule,
): WorkbookDataModelPartInfo | null {
	const kind = classifyDataModelPart(capsule)
	if (!kind) return null
	return {
		kind,
		partPath: capsule.partPath,
		contentType: capsule.contentType,
		...(capsule.relType ? { relType: capsule.relType } : {}),
		relationshipCount: capsule.relationships.length,
	}
}

function classifyDataModelPart(capsule: PreservationCapsule): WorkbookDataModelPartKind | null {
	const path = capsule.partPath.toLowerCase()
	const contentType = capsule.contentType.toLowerCase()
	const relType = capsule.relType?.toLowerCase() ?? ''
	const looksLikeDataModel =
		path.includes('/model/') ||
		contentType.includes('datamodel') ||
		contentType.includes('model+') ||
		relType.includes('datamodel') ||
		relType.includes('/model')
	if (!looksLikeDataModel) return null
	if (path.includes('/model/tables/')) return 'modelTable'
	if (path.includes('/model/relationships/') || contentType.includes('modelrelationship')) {
		return 'modelRelationship'
	}
	if (path.endsWith('.data') || contentType.includes('model+data')) return 'modelData'
	if (path.endsWith('.xml') || contentType.includes('model')) return 'modelMetadata'
	return 'unknownDataModel'
}
