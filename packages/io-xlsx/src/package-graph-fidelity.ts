import type { XlsxPackageGraph, XlsxPackageGraphRelationship } from './package-graph.ts'
import { extractZip } from './reader/zip.ts'

export type XlsxPackageGraphFidelityIssueCode =
	| 'package_feature_classification'
	| 'package_relationship_target'
	| 'package_content_type_default_set'
	| 'package_content_type_override'
	| 'package_signature_invalidation'
	| 'package_preserved_part'
	| 'package_preserved_part_identity'
	| 'package_preserved_part_bytes'
	| 'package_preserved_relationship'
	| 'package_preserved_relationship_identity'

export interface XlsxPackageGraphFidelityIssue {
	readonly code: XlsxPackageGraphFidelityIssueCode
	readonly message: string
	readonly partPath?: string
	readonly relationshipPartPath?: string
	readonly relationshipId?: string
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
	const allowPreservedOtherPart = options.allowPreservedOtherPart ?? (() => false)
	for (const part of graph.parts) {
		if (part.featureFamily !== 'preservedOther' || allowPreservedOtherPart(part.path)) continue
		issues.push({
			code: 'package_feature_classification',
			message: `package graph has unclassified preservedOther part: ${part.path}`,
			partPath: part.path,
		})
	}
	for (const relationship of graph.relationships) {
		if (relationship.targetMode?.toLowerCase() === 'external') continue
		if (relationship.resolvedTarget !== undefined && partPaths.has(relationship.resolvedTarget)) {
			continue
		}
		issues.push({
			code: 'package_relationship_target',
			message: `relationship ${relationship.relationshipPartPath}#${relationship.id} resolves to missing target ${relationship.resolvedTarget ?? relationship.rawTarget}`,
			relationshipPartPath: relationship.relationshipPartPath,
			relationshipId: relationship.id,
			expected: relationship.rawTarget,
			actual: relationship.resolvedTarget,
		})
	}
	return issues
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
			message: 'content type default set changed after safe edit',
			expected: before.contentTypeDefaults,
			actual: after.contentTypeDefaults,
		})
	}

	const afterOverrides = new Set(
		preservationRelevantContentTypeOverrides(after, options).map(contentTypeOverrideKey),
	)
	for (const override of preservationRelevantContentTypeOverrides(before, options)) {
		if (afterOverrides.has(contentTypeOverrideKey(override))) continue
		issues.push({
			code: 'package_content_type_override',
			message: `content type override disappeared after safe edit: ${override.partPath}`,
			partPath: override.partPath,
			expected: override,
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
				message: `signature part ${beforePart.path} was retained after a generated workbook mutation`,
				partPath: beforePart.path,
				expected: undefined,
				actual: packagePartIdentity(afterPart),
			})
			continue
		}
		const afterPart = afterParts.get(beforePart.path)
		if (!afterPart) {
			issues.push({
				code: 'package_preserved_part',
				message: `preserved package part disappeared after safe edit: ${beforePart.path}`,
				partPath: beforePart.path,
				expected: packagePartIdentity(beforePart),
			})
			continue
		}
		const beforeIdentity = packagePartIdentity(beforePart)
		const afterIdentity = packagePartIdentity(afterPart)
		if (sameJson(afterIdentity, beforeIdentity)) continue
		issues.push({
			code: 'package_preserved_part_identity',
			message: `preserved package part identity changed after safe edit: ${beforePart.path}`,
			partPath: beforePart.path,
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
				message: `preserved relationship disappeared after safe edit: ${beforeRel.relationshipPartPath}#${beforeRel.id}`,
				relationshipPartPath: beforeRel.relationshipPartPath,
				relationshipId: beforeRel.id,
				expected: packageRelationshipIdentity(beforeRel),
			})
			continue
		}
		const beforeIdentity = packageRelationshipIdentity(beforeRel)
		const afterIdentity = packageRelationshipIdentity(afterRel)
		if (sameJson(afterIdentity, beforeIdentity)) continue
		issues.push({
			code: 'package_preserved_relationship_identity',
			message: `preserved relationship identity changed after safe edit: ${beforeRel.relationshipPartPath}#${beforeRel.id}`,
			relationshipPartPath: beforeRel.relationshipPartPath,
			relationshipId: beforeRel.id,
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
				message: `preserved package part bytes are unavailable after safe edit: ${part.path}`,
				partPath: part.path,
				expected: beforePartBytes?.byteLength,
				actual: afterPartBytes?.byteLength,
			})
			continue
		}
		if (bytesEqual(beforePartBytes, afterPartBytes)) continue
		issues.push({
			code: 'package_preserved_part_bytes',
			message: `preserved package part bytes changed after safe edit: ${part.path}`,
			partPath: part.path,
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
