import type { TimelineCacheInfo, Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type SetTimelineRangeOp = Extract<Operation, { op: 'setTimelineRange' }>

export function handleSetTimelineRange(
	workbook: Workbook,
	op: SetTimelineRangeOp,
): Result<PatchResult> {
	if (!op.timelineCache && !op.partPath) {
		return err(
			ascendError('VALIDATION_ERROR', 'setTimelineRange requires timelineCache or partPath', {
				suggestedFix: 'Use inspect --detail timelines to find timeline cache names and part paths.',
			}),
		)
	}
	const startTime = Date.parse(op.startDate)
	const endTime = Date.parse(op.endDate)
	if (!isDateTimeLike(startTime) || !isDateTimeLike(endTime)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setTimelineRange requires ISO-like startDate and endDate', {
				suggestedFix: 'Use dateTime strings such as 2024-01-01T00:00:00.',
			}),
		)
	}
	if (startTime > endTime) {
		return err(
			ascendError('VALIDATION_ERROR', 'setTimelineRange startDate must be <= endDate', {
				suggestedFix: 'Swap the timeline range bounds or choose a later endDate.',
			}),
		)
	}

	const matches = resolveTimelineCacheMatches(workbook, op)
	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching timeline cache found', {
				suggestedFix:
					'Inspect timelineCaches and provide a matching timelineCache name or partPath.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `setTimelineRange matched ${matches.length} caches`, {
				suggestedFix: 'Provide the timeline cache partPath to disambiguate duplicate cache names.',
			}),
		)
	}

	const cache = matches[0]
	if (!cache) return err(ascendError('VALIDATION_ERROR', 'No matching timeline cache found'))
	const state = cache.state ?? {}
	Object.assign(cache, {
		state: {
			...state,
			singleRangeFilterState: true,
			selection: { startDate: op.startDate, endDate: op.endDate },
		},
	})

	const sheetsModified = workbook.pivotTables
		.filter((pivot) => pivot.name !== undefined && cache.pivotTableNames.includes(pivot.name))
		.map((pivot) => pivot.sheetName)
	const warnings = [
		ascendError(
			'VALIDATION_ERROR',
			'Timeline range changed; pivot table output is stale until Excel refreshes the cache.',
			{
				details: {
					timelineCache: cache.name,
					partPath: cache.partPath,
					startDate: op.startDate,
					endDate: op.endDate,
					pivotTables: [...cache.pivotTableNames],
					pivotSheets: sheetsModified,
				},
				suggestedFix:
					'Open the workbook in Excel or another pivot-aware engine to refresh timeline-linked pivot output cells.',
			},
		),
	]
	return ok(patch([], sheetsModified, false, warnings))
}

function resolveTimelineCacheMatches(
	workbook: Workbook,
	op: SetTimelineRangeOp,
): TimelineCacheInfo[] {
	return workbook.timelineCaches.filter((cache) => {
		if (op.partPath !== undefined && cache.partPath !== op.partPath) return false
		if (
			op.timelineCache !== undefined &&
			cache.name !== op.timelineCache &&
			cache.partPath !== op.timelineCache
		) {
			return false
		}
		return true
	})
}

function isDateTimeLike(value: number): boolean {
	return Number.isFinite(value)
}
