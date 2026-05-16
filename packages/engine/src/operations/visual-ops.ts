import { Buffer } from 'node:buffer'
import type { SheetSparklineGroupInfo, Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { getSheet, type PatchResult, patch, safeParseRange } from './helpers.ts'

type ReplaceImageOp = Extract<Operation, { op: 'replaceImage' }>
type InsertImageOp = Extract<Operation, { op: 'insertImage' }>
type DeleteImageOp = Extract<Operation, { op: 'deleteImage' }>
type SetDrawingTextOp = Extract<Operation, { op: 'setDrawingText' }>
type SetChartSeriesSourceOp = Extract<Operation, { op: 'setChartSeriesSource' }>
type SetSparklineGroupOp = Extract<Operation, { op: 'setSparklineGroup' }>

export function handleReplaceImage(workbook: Workbook, op: ReplaceImageOp): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value

	const content = decodeBase64(op.contentBase64)
	if (!content.ok) return content
	const contentTypeResult = validateImageContentType(op.contentType)
	if (!contentTypeResult.ok) return contentTypeResult
	const imageIndexResult = validateOptionalIndex(op.imageIndex, 'imageIndex')
	if (!imageIndexResult.ok) return imageIndexResult

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
	sheet.ensureWritable()
	sheet.imageRefs[match.index] = {
		...match.image,
		content: content.value,
		contentType: op.contentType,
	}
	sheet.drawingRefs = { ...sheet.drawingRefs, hasDrawing: true }

	return ok(patch([], [sheet.name], false))
}

export function handleInsertImage(workbook: Workbook, op: InsertImageOp): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value

	const content = decodeBase64(op.contentBase64)
	if (!content.ok) return content
	const contentTypeResult = validateImageContentType(op.contentType)
	if (!contentTypeResult.ok) return contentTypeResult

	const drawingPartPath =
		op.drawingPartPath ??
		sheet.imageRefs[0]?.drawingPartPath ??
		sheet.drawingObjectRefs.find((object) => object.source !== 'vml')?.drawingPartPath ??
		'xl/drawings/drawing1.xml'
	const targetPath =
		op.targetPath ?? nextImageTargetPath(workbook, imageExtensionForContentType(op.contentType))
	const relId = op.relId ?? nextImageRelId(sheet, drawingPartPath)
	if (
		workbook.sheets.some((entry) =>
			entry.imageRefs.some((image) => image.targetPath === targetPath),
		)
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'insertImage targetPath already exists in workbook', {
				suggestedFix: 'Provide a unique targetPath or omit it so Ascend can allocate one.',
			}),
		)
	}
	if (drawingPartRelIds(sheet, drawingPartPath).has(relId)) {
		return err(
			ascendError('VALIDATION_ERROR', 'insertImage relId already exists in drawing part', {
				suggestedFix: 'Provide a unique relId or omit it so Ascend can allocate one.',
			}),
		)
	}

	sheet.ensureWritable()
	sheet.imageRefs.push({
		drawingPartPath,
		relId,
		targetPath,
		contentType: op.contentType,
		content: content.value,
		...(op.anchor ? { anchor: op.anchor } : {}),
		...(op.name ? { name: op.name } : {}),
		...(op.description ? { description: op.description } : {}),
	})
	sheet.drawingRefs = { ...sheet.drawingRefs, hasDrawing: true }

	return ok(patch([], [sheet.name], false))
}

export function handleSetSparklineGroup(
	workbook: Workbook,
	op: SetSparklineGroupOp,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value
	if (!Number.isInteger(op.groupIndex) || op.groupIndex < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'setSparklineGroup groupIndex must be non-negative', {
				suggestedFix: 'Use the zero-based groupIndex from inspectSheet().sparklineGroups.',
			}),
		)
	}
	if (!hasSparklineGroupUpdate(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setSparklineGroup requires source or style metadata', {
				suggestedFix:
					'Provide range, locationRange, type, or a marker/display flag to update the sparkline group.',
			}),
		)
	}
	const valueValidation = validateSparklineGroupUpdateValues(op)
	if (!valueValidation.ok) return valueValidation
	const index = sheet.sparklineGroups.findIndex((group) => group.groupIndex === op.groupIndex)
	if (index < 0) {
		return missingSparklineGroupError()
	}

	const group = sheet.sparklineGroups[index]
	if (!group) {
		return missingSparklineGroupError()
	}
	const updated = applySparklineGroupUpdate(group, op)
	sheet.sparklineGroups.splice(index, 1, updated)
	return ok(patch([], [sheet.name], false))
}

export function handleDeleteImage(workbook: Workbook, op: DeleteImageOp): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value

	const imageIndexResult = validateOptionalIndex(op.imageIndex, 'imageIndex')
	if (!imageIndexResult.ok) return imageIndexResult
	if (
		op.imageIndex === undefined &&
		op.targetPath === undefined &&
		op.relId === undefined &&
		op.name === undefined
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'deleteImage requires targetPath, relId, name, or imageIndex',
				{
					suggestedFix: 'Use inspect --detail images to find image identity fields.',
				},
			),
		)
	}

	const matches = sheet.imageRefs
		.map((image, index) => ({ image, index }))
		.filter(({ image, index }) => {
			if (op.imageIndex !== undefined && index !== op.imageIndex) return false
			if (op.targetPath !== undefined && image.targetPath !== op.targetPath) return false
			if (op.relId !== undefined && image.relId !== op.relId) return false
			if (op.name !== undefined && image.name !== op.name) return false
			return true
		})

	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching image found for deleteImage', {
				suggestedFix:
					'Inspect sheet imageRefs and provide a matching targetPath, relId, name, or imageIndex.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `deleteImage matched ${matches.length} images`, {
				suggestedFix: 'Provide a more specific selector, such as targetPath or imageIndex.',
			}),
		)
	}

	const match = matches[0]
	if (!match) return err(ascendError('VALIDATION_ERROR', 'No matching image found for deleteImage'))
	sheet.ensureWritable()
	sheet.imageRefs.splice(match.index, 1)
	if (sheet.imageRefs.length === 0 && sheet.drawingObjectRefs.length === 0) {
		sheet.drawingRefs = { ...sheet.drawingRefs, hasDrawing: false }
	}

	return ok(patch([], [sheet.name], false))
}

export function handleSetDrawingText(
	workbook: Workbook,
	op: SetDrawingTextOp,
): Result<PatchResult> {
	const sheetResult = getSheet(workbook, op.sheet)
	if (!sheetResult.ok) return sheetResult
	const sheet = sheetResult.value

	if (
		op.drawingPartPath === undefined &&
		op.id === undefined &&
		op.name === undefined &&
		op.drawingObjectIndex === undefined
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'setDrawingText requires drawingPartPath, id, name, or drawingObjectIndex',
				{
					suggestedFix:
						'Use inspect --detail visuals or visualInventory to select a drawing object.',
				},
			),
		)
	}
	if (
		op.drawingObjectIndex !== undefined &&
		(op.drawingObjectIndex < 0 || !Number.isInteger(op.drawingObjectIndex))
	) {
		return err(ascendError('VALIDATION_ERROR', 'drawingObjectIndex must be a non-negative integer'))
	}

	const matches = sheet.drawingObjectRefs
		.map((object, index) => ({ object, index }))
		.filter(({ object, index }) => {
			if (op.drawingObjectIndex !== undefined && index !== op.drawingObjectIndex) return false
			if (op.drawingPartPath !== undefined && object.drawingPartPath !== op.drawingPartPath) {
				return false
			}
			if (op.id !== undefined && object.id !== op.id) return false
			if (op.name !== undefined && object.name !== op.name) return false
			return true
		})

	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching drawing object found for setDrawingText', {
				suggestedFix:
					'Inspect sheet drawingObjectRefs and provide a matching drawingPartPath, id, name, or drawingObjectIndex.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setDrawingText matched ${matches.length} drawing objects`, {
				suggestedFix:
					'Provide a more specific selector, such as drawingObjectIndex or id, before committing.',
			}),
		)
	}

	const match = matches[0]
	if (!match) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching drawing object found for setDrawingText'),
		)
	}
	if (match.object.text === undefined) {
		return err(
			ascendError('VALIDATION_ERROR', 'Selected drawing object has no editable text body', {
				suggestedFix: 'Choose a textBox or text-bearing shape from visualInventory.',
			}),
		)
	}

	sheet.ensureWritable()
	sheet.drawingObjectRefs[match.index] = { ...match.object, text: op.text }
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
		const sheetName = op.sheet
		candidates = candidates.filter(
			({ chart }) => chart.sheetName !== undefined && sameSheetName(chart.sheetName, sheetName),
		)
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
	const unsupportedInsertedFields = unsupportedChartSeriesSourceInsertions(op, series)
	if (unsupportedInsertedFields.length > 0) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`setChartSeriesSource cannot add ${formatFieldList(unsupportedInsertedFields)} because the selected chart series has no existing source formula for ${unsupportedInsertedFields.length === 1 ? 'that field' : 'those fields'}`,
				{
					suggestedFix:
						'Only update chart source fields already shown by visualInventory, or preserve the chart unchanged until chart formula insertion is supported.',
				},
			),
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

function unsupportedChartSeriesSourceInsertions(
	op: SetChartSeriesSourceOp,
	series: { readonly nameRef?: string; readonly categoryRef?: string; readonly valueRef?: string },
): string[] {
	const fields: string[] = []
	if (op.nameRef !== undefined && series.nameRef === undefined) fields.push('nameRef')
	if (op.categoryRef !== undefined && series.categoryRef === undefined) fields.push('categoryRef')
	if (op.valueRef !== undefined && series.valueRef === undefined) fields.push('valueRef')
	return fields
}

function formatFieldList(fields: readonly string[]): string {
	if (fields.length <= 1) return fields[0] ?? 'field'
	if (fields.length === 2) return `${fields[0]} or ${fields[1]}`
	return `${fields.slice(0, -1).join(', ')}, or ${fields[fields.length - 1]}`
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

function validateImageContentType(contentType: string): Result<undefined> {
	if (contentType.startsWith('image/')) return ok(undefined)
	return err(
		ascendError('VALIDATION_ERROR', 'contentType must be an image MIME type', {
			suggestedFix: 'Use image/png, image/jpeg, image/gif, or another valid image/* content type.',
		}),
	)
}

function validateOptionalIndex(index: number | undefined, field: string): Result<undefined> {
	if (index === undefined) return ok(undefined)
	if (index >= 0 && Number.isInteger(index)) return ok(undefined)
	return err(ascendError('VALIDATION_ERROR', `${field} must be a non-negative integer`))
}

function hasSparklineGroupUpdate(op: SetSparklineGroupOp): boolean {
	return (
		op.range !== undefined ||
		op.locationRange !== undefined ||
		op.type !== undefined ||
		op.markers !== undefined ||
		op.highPoint !== undefined ||
		op.lowPoint !== undefined ||
		op.firstPoint !== undefined ||
		op.lastPoint !== undefined ||
		op.negative !== undefined ||
		op.displayXAxis !== undefined
	)
}

function validateSparklineGroupUpdateValues(op: SetSparklineGroupOp): Result<undefined> {
	for (const [field, value] of [
		['range', op.range],
		['locationRange', op.locationRange],
	] as const) {
		if (value === undefined) continue
		const parsed = safeParseRange(value)
		if (!parsed.ok) {
			return err(
				ascendError('VALIDATION_ERROR', `setSparklineGroup ${field} must be a valid A1 range`, {
					suggestedFix: `Provide ${field} as a valid A1 range such as A1:A5.`,
				}),
			)
		}
	}
	if (op.type !== undefined && op.type.trim() === '') {
		return err(
			ascendError('VALIDATION_ERROR', 'setSparklineGroup type must be non-empty', {
				suggestedFix: 'Omit type to preserve it or provide a non-empty sparkline type.',
			}),
		)
	}
	for (const [field, value] of [
		['markers', op.markers],
		['highPoint', op.highPoint],
		['lowPoint', op.lowPoint],
		['firstPoint', op.firstPoint],
		['lastPoint', op.lastPoint],
		['negative', op.negative],
		['displayXAxis', op.displayXAxis],
	] as const) {
		if (value !== undefined && typeof value !== 'boolean') {
			return err(
				ascendError('VALIDATION_ERROR', `setSparklineGroup ${field} must be a boolean`, {
					suggestedFix: `Use true or false for ${field}.`,
				}),
			)
		}
	}
	return ok(undefined)
}

function sameSheetName(left: string, right: string): boolean {
	return left.toLowerCase() === right.toLowerCase()
}

function missingSparklineGroupError(): Result<never> {
	return err(
		ascendError('VALIDATION_ERROR', 'No matching sparkline group found', {
			suggestedFix: 'Inspect sheet.sparklineGroups and provide an existing groupIndex.',
		}),
	)
}

function applySparklineGroupUpdate(
	group: SheetSparklineGroupInfo,
	op: SetSparklineGroupOp,
): SheetSparklineGroupInfo {
	return {
		...group,
		...(op.range !== undefined ? { range: op.range } : {}),
		...(op.locationRange !== undefined ? { locationRange: op.locationRange } : {}),
		...(op.type !== undefined ? { type: op.type } : {}),
		...(op.markers !== undefined ? { markers: op.markers } : {}),
		...(op.highPoint !== undefined ? { highPoint: op.highPoint } : {}),
		...(op.lowPoint !== undefined ? { lowPoint: op.lowPoint } : {}),
		...(op.firstPoint !== undefined ? { firstPoint: op.firstPoint } : {}),
		...(op.lastPoint !== undefined ? { lastPoint: op.lastPoint } : {}),
		...(op.negative !== undefined ? { negative: op.negative } : {}),
		...(op.displayXAxis !== undefined ? { displayXAxis: op.displayXAxis } : {}),
	}
}

function nextImageRelId(sheet: Workbook['sheets'][number], drawingPartPath: string): string {
	const used = drawingPartRelIds(sheet, drawingPartPath)
	let index = imageCountForDrawingPart(sheet, drawingPartPath) + 1
	while (used.has(`rIdImage${index}`)) index++
	return `rIdImage${index}`
}

function imageCountForDrawingPart(
	sheet: Workbook['sheets'][number],
	drawingPartPath: string,
): number {
	return sheet.imageRefs.filter((image) => image.drawingPartPath === drawingPartPath).length
}

function drawingPartRelIds(
	sheet: Workbook['sheets'][number],
	drawingPartPath: string,
): Set<string> {
	const used = new Set<string>()
	for (const image of sheet.imageRefs) {
		if (image.drawingPartPath === drawingPartPath) used.add(image.relId)
	}
	for (const object of sheet.drawingObjectRefs) {
		if (object.drawingPartPath !== drawingPartPath) continue
		for (const relId of object.relIds ?? []) used.add(relId)
		for (const relationship of object.relationshipRefs ?? []) used.add(relationship.id)
	}
	return used
}

function nextImageTargetPath(workbook: Workbook, extension: string): string {
	const used = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const image of sheet.imageRefs) used.add(image.targetPath)
	}
	let index = used.size + 1
	while (used.has(`xl/media/image${index}.${extension}`)) index++
	return `xl/media/image${index}.${extension}`
}

function imageExtensionForContentType(contentType: string): string {
	switch (contentType) {
		case 'image/jpeg':
		case 'image/jpg':
			return 'jpg'
		case 'image/gif':
			return 'gif'
		case 'image/bmp':
			return 'bmp'
		case 'image/webp':
			return 'webp'
		default:
			return 'png'
	}
}
