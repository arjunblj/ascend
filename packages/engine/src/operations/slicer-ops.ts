import type { SlicerCacheInfo, SlicerCacheItemInfo, Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'
import { markLinkedPivotCachesStale } from './linked-pivot-refresh.ts'

type SetSlicerCacheItemOp = Extract<Operation, { op: 'setSlicerCacheItem' }>

export function handleSetSlicerCacheItem(
	workbook: Workbook,
	op: SetSlicerCacheItemOp,
): Result<PatchResult> {
	if (!hasSlicerCacheSelector(op)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setSlicerCacheItem requires slicerCache or partPath', {
				suggestedFix: 'Use inspect --detail slicers to find slicer cache names and part paths.',
			}),
		)
	}
	if (!Number.isInteger(op.item) || op.item < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'setSlicerCacheItem item must be a non-negative integer', {
				suggestedFix: 'Use the zero-based x item index from the slicer cache inventory.',
			}),
		)
	}
	if (op.selected === undefined && op.noData === undefined) {
		return err(
			ascendError('VALIDATION_ERROR', 'setSlicerCacheItem requires selected or noData', {
				suggestedFix:
					'Set selected/noData to true or false, or set them to null to clear the flag.',
			}),
		)
	}
	const updateValidation = validateSlicerCacheItemUpdateValues(op)
	if (updateValidation) return err(updateValidation)

	const matches = resolveSlicerCacheMatches(workbook, op)
	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching slicer cache found', {
				suggestedFix: 'Inspect slicerCaches and provide a matching slicerCache name or partPath.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setSlicerCacheItem matched ${matches.length} caches`, {
				suggestedFix: 'Provide the slicer cache partPath to disambiguate duplicate cache names.',
			}),
		)
	}

	const cache = matches[0]
	if (!cache) return err(ascendError('VALIDATION_ERROR', 'No matching slicer cache found'))
	const itemMatches = (cache.items ?? []).filter((item) => item.index === op.item)
	if (itemMatches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', `Slicer cache item ${op.item} was not found`, {
				suggestedFix: 'Use an item index from the selected slicer cache inventory.',
			}),
		)
	}
	if (itemMatches.length > 1) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`Slicer cache item ${op.item} matched ${itemMatches.length} items`,
				{
					suggestedFix: 'Repair duplicate slicer cache item indexes before editing item state.',
				},
			),
		)
	}
	const items = updateSlicerCacheItem(cache.items ?? [], op)
	Object.assign(cache, { items })

	const impact = markLinkedPivotCachesStale(workbook, cache.pivotTableNames, cache.pivotCacheId)
	const warnings = [
		ascendError(
			'VALIDATION_ERROR',
			'Slicer cache item state changed; pivot table output is stale until Excel refreshes the cache.',
			{
				details: {
					slicerCache: cache.name,
					partPath: cache.partPath,
					item: op.item,
					pivotTables: impact.pivotTables,
					pivotSheets: impact.pivotSheets,
					cacheIds: impact.cacheIds,
					cachePartPaths: impact.cachePartPaths,
				},
				suggestedFix:
					'Open the workbook in Excel or another pivot-aware engine to refresh slicer-linked pivot output cells.',
			},
		),
	]
	return ok(patch([], [...impact.pivotSheets], false, warnings))
}

function hasSlicerCacheSelector(op: SetSlicerCacheItemOp): boolean {
	return op.slicerCache !== undefined || op.partPath !== undefined
}

function resolveSlicerCacheMatches(
	workbook: Workbook,
	op: SetSlicerCacheItemOp,
): SlicerCacheInfo[] {
	return workbook.slicerCaches.filter((cache) => {
		if (op.partPath !== undefined && cache.partPath !== op.partPath) return false
		if (
			op.slicerCache !== undefined &&
			cache.name !== op.slicerCache &&
			cache.partPath !== op.slicerCache
		) {
			return false
		}
		return true
	})
}

function validateSlicerCacheItemUpdateValues(op: SetSlicerCacheItemOp) {
	for (const field of ['selected', 'noData'] as const) {
		const value = op[field]
		if (value !== undefined && value !== null && typeof value !== 'boolean') {
			return ascendError(
				'VALIDATION_ERROR',
				`setSlicerCacheItem ${field} must be boolean or null`,
				{
					suggestedFix: `Set ${field}=true, ${field}=false, or ${field}=null to clear it.`,
				},
			)
		}
	}
	return null
}

function updateSlicerCacheItem(
	items: readonly SlicerCacheItemInfo[],
	op: SetSlicerCacheItemOp,
): SlicerCacheItemInfo[] {
	return items.map((item) => {
		const next = { ...item }
		if (item.index !== op.item) return next
		applyNullableBool(next, 'selected', op.selected)
		applyNullableBool(next, 'noData', op.noData)
		return next
	})
}

function applyNullableBool(
	target: { selected?: boolean; noData?: boolean },
	field: 'selected' | 'noData',
	value: boolean | null | undefined,
): void {
	if (value === undefined) return
	if (value === null) {
		delete target[field]
		return
	}
	target[field] = value
}
