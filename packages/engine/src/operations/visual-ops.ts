import { Buffer } from 'node:buffer'
import type { Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { getSheet, type PatchResult, patch } from './helpers.ts'

type ReplaceImageOp = Extract<Operation, { op: 'replaceImage' }>

export function handleReplaceImage(workbook: Workbook, op: ReplaceImageOp): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value

	const content = decodeBase64(op.contentBase64)
	if (!content.ok) return content

	const matches = sheet.imageRefs
		.map((image, index) => ({ image, index }))
		.filter(({ image, index }) => {
			if (op.imageIndex !== undefined && index !== op.imageIndex) return false
			if (op.targetPath !== undefined && image.targetPath !== op.targetPath) return false
			if (op.relId !== undefined && image.relId !== op.relId) return false
			if (op.name !== undefined && image.name !== op.name) return false
			return true
		})

	if (
		op.imageIndex === undefined &&
		op.targetPath === undefined &&
		op.relId === undefined &&
		op.name === undefined
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'replaceImage requires targetPath, relId, name, or imageIndex',
				{
					suggestedFix: 'Use ascend inspect --detail images to find image identity fields.',
				},
			),
		)
	}

	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching image found for replaceImage', {
				suggestedFix:
					'Inspect sheet imageRefs and provide a matching targetPath, relId, name, or imageIndex.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `replaceImage matched ${matches.length} images`, {
				suggestedFix:
					'Provide a more specific selector, such as targetPath or imageIndex, before committing.',
			}),
		)
	}

	const match = matches[0]
	if (!match) {
		return err(ascendError('VALIDATION_ERROR', 'No matching image found for replaceImage'))
	}
	sheet.imageRefs[match.index] = {
		...match.image,
		content: content.value,
		contentType: op.contentType,
	}
	sheet.drawingRefs = { ...sheet.drawingRefs, hasDrawing: true }

	return ok(patch([], [sheet.name], false))
}

function decodeBase64(input: string): Result<Uint8Array> {
	try {
		const normalized = input.replace(/\s+/g, '')
		if (!normalized || normalized.length % 4 !== 0 || !/^[A-Za-z0-9+/]+={0,2}$/.test(normalized)) {
			throw new Error('invalid base64')
		}
		return ok(new Uint8Array(Buffer.from(normalized, 'base64')))
	} catch {
		return err(
			ascendError('VALIDATION_ERROR', 'contentBase64 must be valid base64 image bytes', {
				suggestedFix: 'Encode the replacement image as base64 and keep contentType aligned.',
			}),
		)
	}
}
