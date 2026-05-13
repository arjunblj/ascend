import type { TimelineCacheInfo, Workbook } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'
import { markLinkedPivotCachesStale } from './linked-pivot-refresh.ts'

type SetTimelineRangeOp = Extract<Operation, { op: 'setTimelineRange' }>

export function handleSetTimelineRange(
	workbook: Workbook,
	op: SetTimelineRangeOp,
): Result<PatchResult> {
	if (op.timelineCache === undefined && op.partPath === undefined) {
		return err(
			ascendError('VALIDATION_ERROR', 'setTimelineRange requires timelineCache or partPath', {
				suggestedFix: 'Use inspect --detail timelines to find timeline cache names and part paths.',
			}),
		)
	}
	const startTime = parseTimelineDateTime(op.startDate)
	const endTime = parseTimelineDateTime(op.endDate)
	if (startTime === null || endTime === null) {
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

	const impact = markLinkedPivotCachesStale(workbook, cache.pivotTableNames, cache.pivotCacheId)
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
					pivotTables: impact.pivotTables,
					pivotSheets: impact.pivotSheets,
					cacheIds: impact.cacheIds,
					cachePartPaths: impact.cachePartPaths,
				},
				suggestedFix:
					'Open the workbook in Excel or another pivot-aware engine to refresh timeline-linked pivot output cells.',
			},
		),
	]
	return ok(patch([], [...impact.pivotSheets], false, warnings))
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

const ISO_LIKE_DATE_TIME_RE =
	/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d{1,9})?)?(?:Z|[+-]\d{2}:\d{2})?)?$/

function parseTimelineDateTime(value: string): number | null {
	const match = ISO_LIKE_DATE_TIME_RE.exec(value)
	if (!match) return null
	const year = Number.parseInt(match[1] ?? '', 10)
	const month = Number.parseInt(match[2] ?? '', 10)
	const day = Number.parseInt(match[3] ?? '', 10)
	const hour = match[4] === undefined ? 0 : Number.parseInt(match[4], 10)
	const minute = match[5] === undefined ? 0 : Number.parseInt(match[5], 10)
	const second = match[6] === undefined ? 0 : Number.parseInt(match[6], 10)
	if (hour > 23 || minute > 59 || second > 59) return null
	const calendarDate = new Date(Date.UTC(year, month - 1, day))
	if (
		calendarDate.getUTCFullYear() !== year ||
		calendarDate.getUTCMonth() !== month - 1 ||
		calendarDate.getUTCDate() !== day
	) {
		return null
	}
	const parsed = Date.parse(value)
	return Number.isFinite(parsed) ? parsed : null
}
