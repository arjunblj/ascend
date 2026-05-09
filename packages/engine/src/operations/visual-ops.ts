import { Buffer } from 'node:buffer'
import type { Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { getSheet, type PatchResult, patch } from './helpers.ts'

type ReplaceImageOp = Extract<Operation, { op: 'replaceImage' }>
type SetChartSeriesSourceOp = Extract<Operation, { op: 'setChartSeriesSource' }>

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

export function handleSetChartSeriesSource(
	workbook: Workbook,
	op: SetChartSeriesSourceOp,
): Result<PatchResult> {
	if (op.nameRef === undefined && op.categoryRef === undefined && op.valueRef === undefined) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'setChartSeriesSource requires nameRef, categoryRef, or valueRef',
				{
					suggestedFix:
						'Inspect workbook charts and provide at least one updated series source reference.',
				},
			),
		)
	}
	if (op.seriesIndex < 0 || !Number.isInteger(op.seriesIndex)) {
		return err(ascendError('VALIDATION_ERROR', 'seriesIndex must be a non-negative integer'))
	}

	let candidates = workbook.chartParts.map((chart, index) => ({ chart, index }))
	if (op.partPath !== undefined) {
		candidates = candidates.filter(({ chart }) => chart.partPath === op.partPath)
	}
	if (op.sheet !== undefined) {
		candidates = candidates.filter(({ chart }) => chart.sheetName === op.sheet)
	}
	if (op.chartIndex !== undefined) {
		if (op.chartIndex < 0 || !Number.isInteger(op.chartIndex)) {
			return err(ascendError('VALIDATION_ERROR', 'chartIndex must be a non-negative integer'))
		}
		const chart = candidates[op.chartIndex]
		candidates = chart ? [chart] : []
	}

	if (candidates.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching chart found for setChartSeriesSource', {
				suggestedFix:
					'Use inspect --detail visuals or visualInventory to choose partPath, sheet, or chartIndex.',
			}),
		)
	}
	if (candidates.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setChartSeriesSource matched ${candidates.length} charts`, {
				suggestedFix:
					'Provide partPath or chartIndex to select exactly one chart before committing.',
			}),
		)
	}

	const match = candidates[0]
	if (!match) {
		return err(ascendError('VALIDATION_ERROR', 'No matching chart found for setChartSeriesSource'))
	}
	const series = match.chart.series[op.seriesIndex]
	if (!series) {
		return err(
			ascendError('VALIDATION_ERROR', `Chart has no series at index ${op.seriesIndex}`, {
				suggestedFix: 'Inspect chart series and choose an existing zero-based seriesIndex.',
			}),
		)
	}

	workbook.chartParts[match.index] = {
		...match.chart,
		series: match.chart.series.map((entry, index) =>
			index === op.seriesIndex
				? {
						...entry,
						...(op.nameRef !== undefined ? { nameRef: op.nameRef } : {}),
						...(op.categoryRef !== undefined ? { categoryRef: op.categoryRef } : {}),
						...(op.valueRef !== undefined ? { valueRef: op.valueRef } : {}),
					}
				: entry,
		),
	}

	return ok(patch([], match.chart.sheetName ? [match.chart.sheetName] : [], false))
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
