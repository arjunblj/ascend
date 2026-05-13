import {
	type PivotCacheInfo,
	type PivotFieldItemInfo,
	type PivotTableInfo,
	parseRange,
	type Workbook,
} from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type SetPivotCacheOp = Extract<Operation, { op: 'setPivotCache' }>
type SetPivotFieldItemOp = Extract<Operation, { op: 'setPivotFieldItem' }>

export function handleSetPivotCache(workbook: Workbook, op: SetPivotCacheOp): Result<PatchResult> {
	if (!hasPivotSelector(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotCache requires cacheId, partPath, or pivotTable', {
				suggestedFix: 'Use inspect --detail pivots to find pivot cache identity fields.',
			}),
		)
	}
	if (!hasPivotUpdate(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotCache requires at least one editable field', {
				suggestedFix:
					'Set sourceSheet/sourceRef, refreshOnLoad, enableRefresh, invalid, or saveData.',
			}),
		)
	}
	const sourceRefValidation = validatePivotCacheSourceRef(op.sourceRef)
	if (sourceRefValidation) return err(sourceRefValidation)
	const pivotTableSelectorValidation = validatePivotCacheTableSelector(workbook, op)
	if (pivotTableSelectorValidation) return err(pivotTableSelectorValidation)

	const matches = resolvePivotCacheMatches(workbook, op)
	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching pivot cache found', {
				suggestedFix:
					'Inspect pivotCaches and provide a matching cacheId, partPath, or pivotTable.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setPivotCache matched ${matches.length} caches`, {
				suggestedFix: 'Provide a more specific selector such as partPath.',
			}),
		)
	}

	const cache = matches[0]
	if (!cache) return err(ascendError('VALIDATION_ERROR', 'No matching pivot cache found'))
	Object.assign(cache, {
		...(op.sourceSheet !== undefined ? { sourceSheet: op.sourceSheet } : {}),
		...(op.sourceRef !== undefined ? { sourceRef: op.sourceRef } : {}),
		...(op.refreshOnLoad !== undefined ? { refreshOnLoad: op.refreshOnLoad } : {}),
		...(op.enableRefresh !== undefined ? { enableRefresh: op.enableRefresh } : {}),
		...(op.invalid !== undefined ? { invalid: op.invalid } : {}),
		...(op.saveData !== undefined ? { saveData: op.saveData } : {}),
	})

	const sheetsModified = workbook.pivotTables
		.filter((pivot) => pivot.cacheId !== undefined && pivot.cacheId === cache.cacheId)
		.map((pivot) => pivot.sheetName)
	const sourceChanged = op.sourceSheet !== undefined || op.sourceRef !== undefined
	const warnings = sourceChanged
		? [
				ascendError(
					'VALIDATION_ERROR',
					'Pivot cache source changed; pivot table output is stale until Excel refreshes the cache.',
					{
						details: {
							cacheId: cache.cacheId,
							partPath: cache.partPath,
							pivotSheets: sheetsModified,
							refreshOnLoad: cache.refreshOnLoad === true,
							invalid: cache.invalid === true,
						},
						suggestedFix:
							'Set refreshOnLoad=true and invalid=true, then open the workbook in Excel or another pivot-aware engine to refresh output cells.',
					},
				),
			]
		: undefined
	return ok(patch([], sheetsModified, false, warnings))
}

export function handleSetPivotFieldItem(
	workbook: Workbook,
	op: SetPivotFieldItemOp,
): Result<PatchResult> {
	if (!hasPivotTableSelector(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotFieldItem requires pivotTable, partPath, or sheet', {
				suggestedFix: 'Use inspect --detail pivots to find pivot table names and part paths.',
			}),
		)
	}
	if (!Number.isInteger(op.fieldIndex) || op.fieldIndex < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotFieldItem fieldIndex must be non-negative', {
				suggestedFix: 'Use the zero-based pivot field index from inspect --detail pivots.',
			}),
		)
	}
	if (!Number.isInteger(op.itemIndex) || op.itemIndex < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotFieldItem itemIndex must be non-negative', {
				suggestedFix: 'Use the zero-based field item index from inspect --detail pivots.',
			}),
		)
	}
	if (!hasPivotFieldItemUpdate(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotFieldItem requires an item or page filter update', {
				suggestedFix:
					'Set hidden, showDetails, manualFilter, or selectedPageItem; use null to clear a flag.',
			}),
		)
	}

	const matches = resolvePivotTableMatches(workbook, op)
	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching pivot table found', {
				suggestedFix: 'Inspect pivotTables and provide a matching pivotTable, partPath, or sheet.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setPivotFieldItem matched ${matches.length} pivots`, {
				suggestedFix: 'Provide the pivot table partPath or name to disambiguate.',
			}),
		)
	}

	const pivot = matches[0]
	if (!pivot) return err(ascendError('VALIDATION_ERROR', 'No matching pivot table found'))
	const fieldMatches = pivot.fields.filter((entry) => entry.index === op.fieldIndex)
	if (fieldMatches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', `Pivot field ${op.fieldIndex} was not found`, {
				suggestedFix: 'Use a fieldIndex from the pivot table fields inventory.',
			}),
		)
	}
	if (fieldMatches.length > 1) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`Pivot field ${op.fieldIndex} matched ${fieldMatches.length} fields`,
				{
					suggestedFix: 'Repair duplicate pivot field indexes before editing item state.',
				},
			),
		)
	}
	const field = fieldMatches[0]
	if (!field)
		return err(ascendError('VALIDATION_ERROR', `Pivot field ${op.fieldIndex} was not found`))
	const itemMatches = (field.items ?? []).filter((entry) => entry.index === op.itemIndex)
	if (itemMatches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', `Pivot field item ${op.itemIndex} was not found`, {
				suggestedFix: 'Use an itemIndex from the selected pivot field items inventory.',
			}),
		)
	}
	if (itemMatches.length > 1) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`Pivot field item ${op.itemIndex} matched ${itemMatches.length} items`,
				{
					suggestedFix: 'Repair duplicate pivot field item indexes before editing item state.',
				},
			),
		)
	}
	if (
		op.selectedPageItem !== undefined &&
		op.selectedPageItem !== null &&
		(!Number.isInteger(op.selectedPageItem) || op.selectedPageItem < 0)
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'setPivotFieldItem selectedPageItem must be non-negative', {
				suggestedFix:
					'Use selectedPageItem from the selected pivot field items inventory, or null to clear it.',
			}),
		)
	}
	const pageFieldValidation = validateSelectedPageItem(pivot, field, op)
	if (pageFieldValidation) return err(pageFieldValidation)

	const selectedPageItem = op.selectedPageItem
	Object.assign(pivot, {
		fields: pivot.fields.map((entry) =>
			entry.index === op.fieldIndex
				? { ...entry, items: updatePivotFieldItems(entry.items ?? [], op) }
				: entry,
		),
		pageFields:
			selectedPageItem === undefined
				? pivot.pageFields
				: pivot.pageFields.map((entry) =>
						entry.index === op.fieldIndex
							? applyNullableNumber({ ...entry }, selectedPageItem)
							: entry,
					),
	})

	const cache = resolvePivotCacheForTable(workbook, pivot)
	if (cache) Object.assign(cache, { invalid: true, refreshOnLoad: true })

	const warnings = [
		ascendError(
			'VALIDATION_ERROR',
			'Pivot field item state changed; pivot table output is stale until Excel refreshes the cache.',
			{
				details: {
					pivotTable: pivot.name,
					partPath: pivot.partPath,
					sheetName: pivot.sheetName,
					fieldIndex: op.fieldIndex,
					itemIndex: op.itemIndex,
					cacheId: pivot.cacheId,
					refreshOnLoad: cache?.refreshOnLoad === true,
					invalid: cache?.invalid === true,
				},
				suggestedFix:
					'Open the workbook in Excel or another pivot-aware engine to refresh pivot output cells.',
			},
		),
	]
	return ok(patch([], [pivot.sheetName], false, warnings))
}

function hasPivotSelector(op: SetPivotCacheOp): boolean {
	return op.cacheId !== undefined || op.partPath !== undefined || op.pivotTable !== undefined
}

function hasPivotUpdate(op: SetPivotCacheOp): boolean {
	return (
		op.sourceSheet !== undefined ||
		op.sourceRef !== undefined ||
		op.refreshOnLoad !== undefined ||
		op.enableRefresh !== undefined ||
		op.invalid !== undefined ||
		op.saveData !== undefined
	)
}

function validatePivotCacheSourceRef(sourceRef: string | undefined) {
	if (sourceRef === undefined) return null
	try {
		const body = sourceRef.includes('!') ? sourceRef.slice(sourceRef.indexOf('!') + 1) : sourceRef
		if (body.split(':').length > 2) throw new Error('Invalid range reference')
		const range = parseRange(sourceRef)
		if (range.start.row > range.end.row || range.start.col > range.end.col) {
			return ascendError(
				'VALIDATION_ERROR',
				'setPivotCache sourceRef must be an ordered A1 range',
				{
					suggestedFix: 'Use an A1 range whose top-left cell is before its bottom-right cell.',
				},
			)
		}
	} catch {
		return ascendError('VALIDATION_ERROR', 'setPivotCache sourceRef must be a valid A1 range', {
			suggestedFix: 'Use a range such as A1:D100 or Sheet1!A1:D100.',
		})
	}
	return null
}

function validatePivotCacheTableSelector(workbook: Workbook, op: SetPivotCacheOp) {
	if (op.pivotTable === undefined) return null
	const pivots = workbook.pivotTables.filter((entry) => entry.name === op.pivotTable)
	if (pivots.length === 0) {
		return ascendError('VALIDATION_ERROR', `Pivot table "${op.pivotTable}" was not found`, {
			suggestedFix: 'Inspect pivotTables and provide a matching pivotTable name.',
		})
	}
	if (pivots.length > 1) {
		return ascendError(
			'VALIDATION_ERROR',
			`setPivotCache pivotTable matched ${pivots.length} pivots`,
			{
				suggestedFix: 'Provide cacheId or partPath without the ambiguous pivotTable selector.',
			},
		)
	}
	const pivot = pivots[0]
	if (pivot?.cacheId === undefined) {
		return ascendError('VALIDATION_ERROR', `Pivot table "${op.pivotTable}" has no cacheId`, {
			suggestedFix: 'Select the pivot cache by partPath or repair the pivot table cache binding.',
		})
	}
	if (op.cacheId !== undefined && op.cacheId !== pivot.cacheId) {
		return ascendError('VALIDATION_ERROR', 'setPivotCache cacheId does not match pivotTable', {
			suggestedFix: 'Use selectors that identify the same pivot cache.',
		})
	}
	return null
}

function resolvePivotCacheMatches(workbook: Workbook, op: SetPivotCacheOp): PivotCacheInfo[] {
	let expectedCacheId = op.cacheId
	if (op.pivotTable !== undefined) {
		const pivot = workbook.pivotTables.find((entry) => entry.name === op.pivotTable)
		if (!pivot) return []
		expectedCacheId = pivot.cacheId
	}
	return workbook.pivotCaches.filter((cache) => {
		if (op.partPath !== undefined && cache.partPath !== op.partPath) return false
		if (expectedCacheId !== undefined && cache.cacheId !== expectedCacheId) return false
		return true
	})
}

function hasPivotTableSelector(op: SetPivotFieldItemOp): boolean {
	return op.pivotTable !== undefined || op.partPath !== undefined || op.sheet !== undefined
}

function hasPivotFieldItemUpdate(op: SetPivotFieldItemOp): boolean {
	return (
		op.hidden !== undefined ||
		op.showDetails !== undefined ||
		op.manualFilter !== undefined ||
		op.selectedPageItem !== undefined
	)
}

function resolvePivotTableMatches(workbook: Workbook, op: SetPivotFieldItemOp): PivotTableInfo[] {
	return workbook.pivotTables.filter((pivot) => {
		if (op.partPath !== undefined && pivot.partPath !== op.partPath) return false
		if (op.pivotTable !== undefined && pivot.name !== op.pivotTable) return false
		if (op.sheet !== undefined && pivot.sheetName !== op.sheet) return false
		return true
	})
}

function validateSelectedPageItem(
	pivot: PivotTableInfo,
	field: PivotTableInfo['fields'][number],
	op: SetPivotFieldItemOp,
) {
	if (op.selectedPageItem === undefined) return null
	const pageField = pivot.pageFields.find((entry) => entry.index === op.fieldIndex)
	if (!pageField) {
		return ascendError('VALIDATION_ERROR', `Pivot field ${op.fieldIndex} is not a page field`, {
			suggestedFix:
				'Use selectedPageItem only with a fieldIndex from the pivot pageFields inventory.',
		})
	}
	if (op.selectedPageItem === null) return null
	const itemExists = field.items?.some((entry) => entry.index === op.selectedPageItem) === true
	if (!itemExists) {
		return ascendError(
			'VALIDATION_ERROR',
			`Pivot page-field item ${op.selectedPageItem} was not found`,
			{
				suggestedFix:
					'Use selectedPageItem from the selected pivot field items inventory, or null to clear it.',
			},
		)
	}
	return null
}

function updatePivotFieldItems(
	items: readonly PivotFieldItemInfo[],
	op: SetPivotFieldItemOp,
): PivotFieldItemInfo[] {
	return items.map((item) =>
		item.index === op.itemIndex
			? applyPivotItemUpdates({ ...item }, op)
			: {
					...item,
				},
	)
}

function applyPivotItemUpdates(
	item: PivotFieldItemInfo,
	op: SetPivotFieldItemOp,
): PivotFieldItemInfo {
	const next = { ...item } as {
		index: number
		cacheIndex?: number
		itemType?: string
		caption?: string
		hidden?: boolean
		showDetails?: boolean
		manualFilter?: boolean
	}
	applyNullableBool(next, 'hidden', op.hidden)
	applyNullableBool(next, 'showDetails', op.showDetails)
	applyNullableBool(next, 'manualFilter', op.manualFilter)
	return next
}

function applyNullableBool(
	target: { hidden?: boolean; showDetails?: boolean; manualFilter?: boolean },
	field: 'hidden' | 'showDetails' | 'manualFilter',
	value: boolean | null | undefined,
): void {
	if (value === undefined) return
	if (value === null) {
		delete target[field]
		return
	}
	target[field] = value
}

function applyNullableNumber<T extends { item?: number }>(target: T, value: number | null): T {
	if (value === null) delete target.item
	else target.item = value
	return target
}

function resolvePivotCacheForTable(
	workbook: Workbook,
	pivot: PivotTableInfo,
): PivotCacheInfo | undefined {
	if (pivot.cacheId === undefined) return undefined
	return workbook.pivotCaches.find((cache) => cache.cacheId === pivot.cacheId)
}
