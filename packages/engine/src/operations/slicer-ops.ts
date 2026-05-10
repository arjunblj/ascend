import type { SlicerCacheInfo, SlicerCacheItemInfo, Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

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
	const items = upsertSlicerCacheItem(cache.items ?? [], op)
	Object.assign(cache, { items })

	const sheetsModified = workbook.pivotTables
		.filter((pivot) => pivot.name !== undefined && cache.pivotTableNames.includes(pivot.name))
		.map((pivot) => pivot.sheetName)
	const warnings = [
		ascendError(
			'VALIDATION_ERROR',
			'Slicer cache item state changed; pivot table output is stale until Excel refreshes the cache.',
			{
				details: {
					slicerCache: cache.name,
					partPath: cache.partPath,
					item: op.item,
					pivotTables: [...cache.pivotTableNames],
					pivotSheets: sheetsModified,
				},
				suggestedFix:
					'Open the workbook in Excel or another pivot-aware engine to refresh slicer-linked pivot output cells.',
			},
		),
	]
	return ok(patch([], sheetsModified, false, warnings))
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

function upsertSlicerCacheItem(
	items: readonly SlicerCacheItemInfo[],
	op: SetSlicerCacheItemOp,
): SlicerCacheItemInfo[] {
	const next = items.map((item) => ({ ...item }))
	const existing = next.find((item) => item.index === op.item)
	const target = existing ?? { index: op.item }
	applyNullableBool(target, 'selected', op.selected)
	applyNullableBool(target, 'noData', op.noData)
	if (!existing) next.push(target)
	next.sort((left, right) => left.index - right.index)
	return next
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
