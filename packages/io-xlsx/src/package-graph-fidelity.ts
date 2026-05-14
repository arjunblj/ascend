import {
	packageFeatureLossPolicy,
	type XlsxPackageGraph,
	type XlsxPackageGraphRelationship,
	type XlsxPackageLossPolicy,
	type XlsxPackagePreservationMode,
	xlsxPackagePreservationModeForPolicy,
} from './package-graph.ts'
import { extractZip } from './reader/zip.ts'

export type XlsxPackageGraphFidelityIssueCode =
	| 'package_feature_classification'
	| 'package_relationship_duplicate_id'
	| 'package_relationship_source'
	| 'package_relationship_target'
	| 'package_content_type_override_target'
	| 'package_content_type_override_mismatch'
	| 'package_content_type_default_set'
	| 'package_content_type_override'
	| 'package_signature_invalidation'
	| 'package_preserved_part'
	| 'package_preserved_part_identity'
	| 'package_preserved_part_bytes'
	| 'package_preserved_relationship'
	| 'package_preserved_relationship_identity'

export type XlsxPackageGraphFidelityIssueSeverity = 'warning' | 'error'

export interface XlsxPackageGraphFidelityIssue {
	readonly code: XlsxPackageGraphFidelityIssueCode
	readonly severity: XlsxPackageGraphFidelityIssueSeverity
	readonly message: string
	readonly partPath?: string
	readonly sourcePartPath?: string
	readonly relationshipPartPath?: string
	readonly relationshipId?: string
	readonly contentType?: string
	readonly featureFamily?: string
	readonly preservationPolicy?: XlsxPackageLossPolicy
	readonly preservationMode?: XlsxPackagePreservationMode
	readonly ownerScope?: string
	readonly suggestedAction?: string
	readonly expected?: unknown
	readonly actual?: unknown
}

export interface XlsxPackageGraphReadIntegrityOptions {
	readonly allowPreservedOtherPart?: (partPath: string) => boolean
}

export interface XlsxPackageGraphSafeEditIntegrityOptions {
	readonly ignoredContentTypeOverrideFamilies?: readonly string[]
	readonly ignoredRelationshipFamilies?: readonly string[]
}

const DEFAULT_IGNORED_CONTENT_TYPE_OVERRIDE_FAMILIES = new Set(['preservedSignature'])
const DEFAULT_IGNORED_RELATIONSHIP_FAMILIES = new Set(['preservedCalcChain', 'preservedSignature'])

export function auditXlsxPackageGraphReadIntegrity(
	graph: XlsxPackageGraph,
	options: XlsxPackageGraphReadIntegrityOptions = {},
): readonly XlsxPackageGraphFidelityIssue[] {
	const issues: XlsxPackageGraphFidelityIssue[] = []
	const partPaths = new Set(graph.parts.map((part) => part.path))
	const partByPath = new Map(graph.parts.map((part) => [part.path, part]))
	const reportedMissingSourceSidecars = new Set<string>()
	const allowPreservedOtherPart = options.allowPreservedOtherPart ?? (() => false)
	for (const duplicate of duplicateRelationshipIds(graph.relationships)) {
		const first = duplicate[0] as XlsxPackageGraphRelationship
		issues.push({
			code: 'package_relationship_duplicate_id',
			severity: 'error',
			message: `relationship part ${first.relationshipPartPath} contains duplicate relationship id ${first.id}`,
			sourcePartPath: first.sourcePartPath,
			relationshipPartPath: first.relationshipPartPath,
			relationshipId: first.id,
			featureFamily: first.featureFamily,
			...preservationFieldsForFeatureFamily(first.featureFamily),
			suggestedAction:
				'Repair duplicate OPC relationship ids before resolving or writing the package; relationship ids must be unique within each .rels part.',
			expected: 'unique relationship id within the relationship part',
			actual: duplicate.map((relationship) => ({
				type: relationship.rawType ?? relationship.type,
				target: relationship.rawTarget,
				resolvedTarget: relationship.resolvedTarget,
				targetMode: relationship.targetMode,
			})),
		})
	}
	for (const part of graph.parts) {
		if (part.featureFamily !== 'preservedOther' || allowPreservedOtherPart(part.path)) continue
		issues.push({
			code: 'package_feature_classification',
			severity: 'warning',
			message: `package graph has unclassified preservedOther part: ${part.path}`,
			partPath: part.path,
			ownerScope: part.ownerScope,
			featureFamily: part.featureFamily,
			...preservationFieldsForPart(part),
			suggestedAction:
				'Classify this OOXML part family or add a narrow fixture-backed allowlist entry before treating the package as fully audited.',
		})
	}
	for (const relationship of graph.relationships) {
		if (relationship.sourcePartPath !== '' && !partPaths.has(relationship.sourcePartPath)) {
			issues.push({
				code: 'package_relationship_source',
				severity: 'error',
				message: `relationship sidecar ${relationship.relationshipPartPath} belongs to missing source part ${relationship.sourcePartPath}`,
				sourcePartPath: relationship.sourcePartPath,
				relationshipPartPath: relationship.relationshipPartPath,
				relationshipId: relationship.id,
				featureFamily: relationship.featureFamily,
				...preservationFieldsForFeatureFamily(relationship.featureFamily),
				suggestedAction:
					'Remove the orphan relationship sidecar or restore the source package part before writing.',
				expected: relationship.sourcePartPath,
				actual: undefined,
			})
			reportedMissingSourceSidecars.add(relationship.relationshipPartPath)
		}
		if (relationship.targetMode?.toLowerCase() === 'external') continue
		if (relationship.resolvedTarget !== undefined && partPaths.has(relationship.resolvedTarget)) {
			continue
		}
		issues.push({
			code: 'package_relationship_target',
			severity: 'error',
			message: `relationship ${relationship.relationshipPartPath}#${relationship.id} resolves to missing target ${relationship.resolvedTarget ?? relationship.rawTarget}`,
			sourcePartPath: relationship.sourcePartPath,
			relationshipPartPath: relationship.relationshipPartPath,
			relationshipId: relationship.id,
			featureFamily: relationship.featureFamily,
			...preservationFieldsForFeatureFamily(relationship.featureFamily),
			suggestedAction:
				'Inspect the relationship target and preserve or repair the referenced package part before writing.',
			expected: relationship.rawTarget,
			actual: relationship.resolvedTarget,
		})
	}
	for (const part of graph.parts) {
		if (part.ownerScope !== 'relationship-part') continue
		if (reportedMissingSourceSidecars.has(part.path)) continue
		const sourcePartPath = sourcePartFromRelationshipPartPath(part.path)
		if (sourcePartPath === null || sourcePartPath === '' || partPaths.has(sourcePartPath)) continue
		issues.push({
			code: 'package_relationship_source',
			severity: 'error',
			message: `relationship sidecar ${part.path} belongs to missing source part ${sourcePartPath}`,
			sourcePartPath,
			relationshipPartPath: part.path,
			featureFamily: part.featureFamily,
			ownerScope: part.ownerScope,
			...preservationFieldsForPart(part),
			suggestedAction:
				'Remove the orphan relationship sidecar or restore the source package part before writing.',
			expected: sourcePartPath,
			actual: undefined,
		})
	}
	for (const override of graph.contentTypeOverrides) {
		const part = partByPath.get(override.partPath)
		if (!part) {
			issues.push({
				code: 'package_content_type_override_target',
				severity: 'error',
				message: `content type override points to missing package part: ${override.partPath}`,
				partPath: override.partPath,
				contentType: override.contentType,
				suggestedAction:
					'Remove the stale content type override or restore the referenced package part before writing.',
				expected: override.partPath,
				actual: undefined,
			})
			continue
		}
		if (part.contentType !== override.contentType) {
			issues.push({
				code: 'package_content_type_override_mismatch',
				severity: 'error',
				message: `content type override for ${override.partPath} declares ${override.contentType} but package graph resolved ${part.contentType}`,
				partPath: override.partPath,
				featureFamily: part.featureFamily,
				ownerScope: part.ownerScope,
				...preservationFieldsForPart(part),
				suggestedAction:
					'Make the content type override agree with the resolved package part type before writing.',
				expected: override.contentType,
				actual: part.contentType,
			})
		}
	}
	return issues
}

function duplicateRelationshipIds(
	relationships: readonly XlsxPackageGraphRelationship[],
): XlsxPackageGraphRelationship[][] {
	const byRelationshipId = new Map<string, XlsxPackageGraphRelationship[]>()
	for (const relationship of relationships) {
		const key = `${relationship.relationshipPartPath}\u0000${relationship.id}`
		const group = byRelationshipId.get(key)
		if (group) group.push(relationship)
		else byRelationshipId.set(key, [relationship])
	}
	return [...byRelationshipId.values()].filter((group) => group.length > 1)
}

function sourcePartFromRelationshipPartPath(path: string): string | null {
	if (path === '_rels/.rels') return ''
	const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(path)
	if (!match) return null
	const fileName = match[2]
	if (!fileName) return null
	return match[1] ? `${match[1]}/${fileName}` : fileName
}

function preservationFieldsForPart(part: XlsxPackageGraph['parts'][number]): {
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly preservationMode: XlsxPackagePreservationMode
} {
	return {
		preservationPolicy: part.preservationPolicy,
		preservationMode: xlsxPackagePreservationModeForPolicy(part.preservationPolicy),
	}
}

function preservationFieldsForFeatureFamily(featureFamily: string): {
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly preservationMode: XlsxPackagePreservationMode
} {
	const preservationPolicy = packageFeatureLossPolicy(featureFamily)
	return {
		preservationPolicy,
		preservationMode: xlsxPackagePreservationModeForPolicy(preservationPolicy),
	}
}

export function auditXlsxPackageGraphSafeEditIntegrity(
	before: XlsxPackageGraph,
	after: XlsxPackageGraph,
	options: XlsxPackageGraphSafeEditIntegrityOptions = {},
): readonly XlsxPackageGraphFidelityIssue[] {
	const issues: XlsxPackageGraphFidelityIssue[] = []
	if (!sameJson(after.contentTypeDefaults, before.contentTypeDefaults)) {
		issues.push({
			code: 'package_content_type_default_set',
			severity: 'error',
			message: 'content type default set changed after safe edit',
			partPath: '[Content_Types].xml',
			ownerScope: 'package',
			featureFamily: 'packageContentTypes',
			...preservationFieldsForFeatureFamily('packageContentTypes'),
			suggestedAction:
				'Preserve the package-level default content type table unless an explicit generated-part rewrite owns the change.',
			expected: before.contentTypeDefaults,
			actual: after.contentTypeDefaults,
		})
	}

	const beforeOverrides = preservationRelevantContentTypeOverrides(before, options)
	const afterOverrides = preservationRelevantContentTypeOverrides(after, options)
	const beforeOverridesByPart = new Map(
		beforeOverrides.map((override) => [override.partPath, override]),
	)
	const afterOverridesByPart = new Map(
		afterOverrides.map((override) => [override.partPath, override]),
	)
	const beforeOverrideKeys = new Set(beforeOverrides.map(contentTypeOverrideKey))
	for (const override of beforeOverrides) {
		const afterOverride = afterOverridesByPart.get(override.partPath)
		if (afterOverride?.contentType === override.contentType) continue
		const beforePart = before.parts.find((part) => part.path === override.partPath)
		issues.push({
			code: 'package_content_type_override',
			severity: 'error',
			message: afterOverride
				? `content type override changed after safe edit: ${override.partPath}`
				: `content type override disappeared after safe edit: ${override.partPath}`,
			partPath: override.partPath,
			...(beforePart?.ownerScope ? { ownerScope: beforePart.ownerScope } : {}),
			...(beforePart?.featureFamily ? { featureFamily: beforePart.featureFamily } : {}),
			...(beforePart ? preservationFieldsForPart(beforePart) : {}),
			suggestedAction:
				'Preserve the exact override for this package part or prove that the generated replacement owns the content type change.',
			expected: override,
			actual: afterOverride,
		})
	}
	for (const override of afterOverrides) {
		if (beforeOverridesByPart.has(override.partPath)) continue
		if (beforeOverrideKeys.has(contentTypeOverrideKey(override))) continue
		const afterPart = after.parts.find((part) => part.path === override.partPath)
		if (!afterPart || afterPart.preservationPolicy === 'generated') continue
		issues.push({
			code: 'package_content_type_override',
			severity: 'error',
			message: `content type override appeared after safe edit: ${override.partPath}`,
			partPath: override.partPath,
			ownerScope: afterPart.ownerScope,
			featureFamily: afterPart.featureFamily,
			...preservationFieldsForPart(afterPart),
			suggestedAction:
				'Inspect why a preserved package part gained a new explicit content type override after a safe edit.',
			expected: undefined,
			actual: override,
		})
	}

	const afterParts = new Map(after.parts.map((part) => [part.path, part]))
	for (const beforePart of before.parts) {
		if (beforePart.preservationPolicy === 'discard-on-recalc') continue
		if (beforePart.preservationPolicy === 'generated') continue
		if (beforePart.preservationPolicy === 'invalidate-on-edit') {
			const afterPart = afterParts.get(beforePart.path)
			if (afterPart === undefined) continue
			issues.push({
				code: 'package_signature_invalidation',
				severity: 'error',
				message: `signature part ${beforePart.path} was retained after a generated workbook mutation`,
				partPath: beforePart.path,
				ownerScope: beforePart.ownerScope,
				featureFamily: beforePart.featureFamily,
				...preservationFieldsForPart(beforePart),
				suggestedAction:
					'Drop invalidated package signatures after dirty writes or block the write until signature handling is explicit.',
				expected: undefined,
				actual: packagePartIdentity(afterPart),
			})
			continue
		}
		const afterPart = afterParts.get(beforePart.path)
		if (!afterPart) {
			issues.push({
				code: 'package_preserved_part',
				severity: 'error',
				message: `preserved package part disappeared after safe edit: ${beforePart.path}`,
				partPath: beforePart.path,
				ownerScope: beforePart.ownerScope,
				featureFamily: beforePart.featureFamily,
				...preservationFieldsForPart(beforePart),
				suggestedAction:
					'Copy this preserved package part through the write path or classify an intentional loss policy.',
				expected: packagePartIdentity(beforePart),
			})
			continue
		}
		const beforeIdentity = packagePartIdentity(beforePart)
		const afterIdentity = packagePartIdentity(afterPart)
		if (sameJson(afterIdentity, beforeIdentity)) continue
		issues.push({
			code: 'package_preserved_part_identity',
			severity: 'error',
			message: `preserved package part identity changed after safe edit: ${beforePart.path}`,
			partPath: beforePart.path,
			ownerScope: beforePart.ownerScope,
			featureFamily: beforePart.featureFamily,
			...preservationFieldsForPart(beforePart),
			suggestedAction:
				'Preserve content type, owner scope, relationship provenance, feature family, and loss policy for this copied-through part.',
			expected: beforeIdentity,
			actual: afterIdentity,
		})
	}

	const afterRels = new Map(
		after.relationships.map((relationship) => [
			packageRelationshipIdentityKey(relationship),
			relationship,
		]),
	)
	for (const beforeRel of before.relationships) {
		if (!isSafeEditIdentityRelationship(beforeRel, options)) continue
		const afterRel = afterRels.get(packageRelationshipIdentityKey(beforeRel))
		if (!afterRel) {
			issues.push({
				code: 'package_preserved_relationship',
				severity: 'error',
				message: `preserved relationship disappeared after safe edit: ${beforeRel.relationshipPartPath}#${beforeRel.id}`,
				sourcePartPath: beforeRel.sourcePartPath,
				relationshipPartPath: beforeRel.relationshipPartPath,
				relationshipId: beforeRel.id,
				featureFamily: beforeRel.featureFamily,
				...preservationFieldsForFeatureFamily(beforeRel.featureFamily),
				suggestedAction:
					'Preserve this relationship id, source relationship part, type dialect, target, and target mode across safe edits.',
				expected: packageRelationshipIdentity(beforeRel),
			})
			continue
		}
		const beforeIdentity = packageRelationshipIdentity(beforeRel)
		const afterIdentity = packageRelationshipIdentity(afterRel)
		if (sameJson(afterIdentity, beforeIdentity)) continue
		issues.push({
			code: 'package_preserved_relationship_identity',
			severity: 'error',
			message: `preserved relationship identity changed after safe edit: ${beforeRel.relationshipPartPath}#${beforeRel.id}`,
			sourcePartPath: beforeRel.sourcePartPath,
			relationshipPartPath: beforeRel.relationshipPartPath,
			relationshipId: beforeRel.id,
			featureFamily: beforeRel.featureFamily,
			...preservationFieldsForFeatureFamily(beforeRel.featureFamily),
			suggestedAction:
				'Preserve relationship id, source part, relationship part path, raw type, normalized type, raw target, resolved target, and target mode.',
			expected: beforeIdentity,
			actual: afterIdentity,
		})
	}
	return issues
}

export function auditXlsxPackageGraphBytePreservation(
	before: XlsxPackageGraph,
	beforeBytes: Uint8Array,
	afterBytes: Uint8Array,
): readonly XlsxPackageGraphFidelityIssue[] {
	const issues: XlsxPackageGraphFidelityIssue[] = []
	const beforeArchive = extractZip(beforeBytes)
	const afterArchive = extractZip(afterBytes)
	for (const part of before.parts) {
		if (!part.bytePreservationExpected) continue
		const beforePartBytes = beforeArchive.readBytes(part.path)
		const afterPartBytes = afterArchive.readBytes(part.path)
		if (beforePartBytes === undefined || afterPartBytes === undefined) {
			issues.push({
				code: 'package_preserved_part_bytes',
				severity: 'error',
				message: `preserved package part bytes are unavailable after safe edit: ${part.path}`,
				partPath: part.path,
				ownerScope: part.ownerScope,
				featureFamily: part.featureFamily,
				...preservationFieldsForPart(part),
				suggestedAction:
					'Keep exact bytes available for copied-through OOXML sidecars during safe-edit writes.',
				expected: beforePartBytes?.byteLength,
				actual: afterPartBytes?.byteLength,
			})
			continue
		}
		if (bytesEqual(beforePartBytes, afterPartBytes)) continue
		issues.push({
			code: 'package_preserved_part_bytes',
			severity: 'error',
			message: `preserved package part bytes changed after safe edit: ${part.path}`,
			partPath: part.path,
			ownerScope: part.ownerScope,
			featureFamily: part.featureFamily,
			...preservationFieldsForPart(part),
			suggestedAction:
				'Copy preserve-exact sidecars byte-for-byte unless the feature has a semantic writer with its own fidelity tests.',
			expected: beforePartBytes.byteLength,
			actual: afterPartBytes.byteLength,
		})
	}
	return issues
}

function preservationRelevantContentTypeOverrides(
	graph: XlsxPackageGraph,
	options: XlsxPackageGraphSafeEditIntegrityOptions,
): readonly XlsxPackageGraph['contentTypeOverrides'][number][] {
	const ignoredFamilies = new Set(
		options.ignoredContentTypeOverrideFamilies ?? DEFAULT_IGNORED_CONTENT_TYPE_OVERRIDE_FAMILIES,
	)
	return graph.contentTypeOverrides.filter(
		(override) =>
			!ignoredFamilies.has(classifyPackageGraphOverrideFamily(graph, override.partPath)),
	)
}

function classifyPackageGraphOverrideFamily(graph: XlsxPackageGraph, partPath: string): string {
	return graph.parts.find((part) => part.path === partPath)?.featureFamily ?? 'preservedOther'
}

function contentTypeOverrideKey(
	override: XlsxPackageGraph['contentTypeOverrides'][number],
): string {
	return `${override.partPath}\u0000${override.contentType}`
}

function isSafeEditIdentityRelationship(
	relationship: XlsxPackageGraphRelationship,
	options: XlsxPackageGraphSafeEditIntegrityOptions,
): boolean {
	const ignoredFamilies = new Set(
		options.ignoredRelationshipFamilies ?? DEFAULT_IGNORED_RELATIONSHIP_FAMILIES,
	)
	if (ignoredFamilies.has(relationship.featureFamily)) return false
	return (
		relationship.sourcePartPath === '' ||
		relationship.featureFamily === 'workbook' ||
		relationship.featureFamily === 'worksheet' ||
		relationship.featureFamily === 'preservedChartSheet' ||
		relationship.featureFamily === 'preservedMacroSheet' ||
		relationship.featureFamily.startsWith('preserved')
	)
}

function packagePartIdentity(part: XlsxPackageGraph['parts'][number]): Record<string, unknown> {
	return {
		contentType: part.contentType,
		contentTypeSource: part.contentTypeSource,
		ownerScope: part.ownerScope,
		sourceRelationshipPart: part.sourceRelationshipPart,
		sourceRelationshipId: part.sourceRelationshipId,
		sourceRelationshipType: part.sourceRelationshipType,
		sourceRelationshipRawType: part.sourceRelationshipRawType,
		sourceRelationshipRawTarget: part.sourceRelationshipRawTarget,
		sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget,
		sourceRelationshipTargetMode: part.sourceRelationshipTargetMode,
		featureFamily: part.featureFamily,
		preservationPolicy: part.preservationPolicy,
		bytePreservationExpected: part.bytePreservationExpected,
	}
}

function packageRelationshipIdentity(
	relationship: XlsxPackageGraphRelationship,
): Record<string, unknown> {
	return {
		sourcePartPath: relationship.sourcePartPath,
		relationshipPartPath: relationship.relationshipPartPath,
		id: relationship.id,
		type: relationship.type,
		rawType: relationship.rawType,
		rawTarget: relationship.rawTarget,
		resolvedTarget: relationship.resolvedTarget,
		targetMode: relationship.targetMode,
		featureFamily: relationship.featureFamily,
	}
}

function packageRelationshipIdentityKey(relationship: XlsxPackageGraphRelationship): string {
	return `${relationship.relationshipPartPath}\u0000${relationship.id}`
}

function sameJson(left: unknown, right: unknown): boolean {
	return JSON.stringify(left) === JSON.stringify(right)
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
	if (left.byteLength !== right.byteLength) return false
	for (let index = 0; index < left.byteLength; index++) {
		if (left[index] !== right[index]) return false
	}
	return true
}
