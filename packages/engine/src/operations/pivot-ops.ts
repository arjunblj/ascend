import type { PivotCacheInfo, Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type SetPivotCacheOp = Extract<Operation, { op: 'setPivotCache' }>

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
	return ok(patch([], sheetsModified, false))
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
