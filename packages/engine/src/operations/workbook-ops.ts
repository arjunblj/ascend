import type { Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type RewriteExternalLinkOp = Extract<Operation, { op: 'rewriteExternalLink' }>

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
