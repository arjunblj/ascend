import type { Workbook, WorkbookProperties } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type RewriteExternalLinkOp = Extract<Operation, { op: 'rewriteExternalLink' }>
type SetWorkbookPropertiesOp = Extract<Operation, { op: 'setWorkbookProperties' }>

const WORKBOOK_PROPERTY_KEYS = [
	'codeName',
	'defaultThemeVersion',
	'filterPrivacy',
	'date1904',
] as const

export function handleSetWorkbookProperties(
	workbook: Workbook,
	op: SetWorkbookPropertiesOp,
): Result<PatchResult> {
	const mode = op.mode ?? 'merge'
	if (mode !== 'merge' && mode !== 'replace') {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookProperties mode must be merge or replace', {
				suggestedFix: 'Use mode="merge" to update selected properties or mode="replace".',
			}),
		)
	}

	const validated = validateWorkbookProperties(op.properties)
	if (!validated.ok) return validated

	const next: Record<string, string | number | boolean> =
		mode === 'replace' ? {} : { ...workbook.workbookProperties }
	for (const key of WORKBOOK_PROPERTY_KEYS) {
		if (!(key in op.properties)) continue
		const value = op.properties[key]
		if (value === null || value === undefined) {
			delete next[key]
		} else {
			next[key] = value
		}
	}

	const oldDateSystem = workbook.calcSettings.dateSystem
	const date1904 =
		'date1904' in op.properties
			? op.properties.date1904 === true
			: mode === 'replace'
				? false
				: workbook.workbookProperties.date1904 === true
	if (mode === 'replace' || 'date1904' in op.properties) {
		workbook.calcSettings = {
			...workbook.calcSettings,
			dateSystem: date1904 ? '1904' : '1900',
		}
	}
	workbook.workbookProperties = next as WorkbookProperties

	return ok(patch([], [], oldDateSystem !== workbook.calcSettings.dateSystem))
}

export function handleRewriteExternalLink(
	workbook: Workbook,
	op: RewriteExternalLinkOp,
): Result<PatchResult> {
	if (
		op.partPath === undefined &&
		op.relId === undefined &&
		op.linkRelId === undefined &&
		op.target === undefined
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'rewriteExternalLink requires partPath, relId, linkRelId, or target',
				{
					suggestedFix:
						'Use inspect --detail external-refs and select a stable external link identifier.',
				},
			),
		)
	}
	if (op.newTarget.trim() === '') {
		return err(
			ascendError('VALIDATION_ERROR', 'newTarget must be a non-empty external workbook target', {
				suggestedFix: 'Provide the replacement workbook path or URL for the external link target.',
			}),
		)
	}

	const matches = workbook.externalReferenceDetails
		.map((entry, index) => ({ entry, index }))
		.filter(({ entry }) => {
			if (op.partPath !== undefined && entry.partPath !== op.partPath) return false
			if (op.relId !== undefined && entry.relId !== op.relId) return false
			if (op.linkRelId !== undefined && entry.linkRelId !== op.linkRelId) return false
			if (op.target !== undefined && entry.target !== op.target) return false
			return true
		})

	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching external link found for rewriteExternalLink', {
				suggestedFix:
					'Inspect workbook externalReferenceDetails and provide a matching partPath, relId, linkRelId, or target.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `rewriteExternalLink matched ${matches.length} links`, {
				suggestedFix:
					'Provide a more specific selector such as partPath plus linkRelId before committing.',
			}),
		)
	}

	const match = matches[0]
	if (!match) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching external link found for rewriteExternalLink'),
		)
	}
	workbook.externalReferenceDetails[match.index] = {
		...match.entry,
		target: op.newTarget,
		...(op.targetMode !== undefined ? { targetMode: op.targetMode } : {}),
	}

	return ok(patch([], [], false))
}

function validateWorkbookProperties(
	properties: SetWorkbookPropertiesOp['properties'],
): Result<undefined> {
	if (!properties || typeof properties !== 'object' || Array.isArray(properties)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookProperties requires a properties object', {
				suggestedFix: 'Provide properties such as { "codeName": "Model", "filterPrivacy": true }.',
			}),
		)
	}
	if (
		properties.codeName !== undefined &&
		properties.codeName !== null &&
		properties.codeName.trim() === ''
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'workbook property codeName must be non-empty', {
				suggestedFix: 'Use null to clear codeName, or provide a non-empty workbook code name.',
			}),
		)
	}
	if (
		properties.defaultThemeVersion !== undefined &&
		properties.defaultThemeVersion !== null &&
		(!Number.isInteger(properties.defaultThemeVersion) || properties.defaultThemeVersion < 0)
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'workbook property defaultThemeVersion must be a non-negative integer',
				{
					suggestedFix: 'Use a non-negative integer theme version, or null to clear it.',
				},
			),
		)
	}
	return ok(undefined)
}
