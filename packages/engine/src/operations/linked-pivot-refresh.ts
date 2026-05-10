import type { Workbook } from '@ascend/core'

export interface LinkedPivotRefreshImpact {
	readonly pivotTables: readonly string[]
	readonly pivotSheets: readonly string[]
	readonly cacheIds: readonly number[]
	readonly cachePartPaths: readonly string[]
}

export function markLinkedPivotCachesStale(
	workbook: Workbook,
	pivotTableNames: readonly string[],
	pivotCacheId?: number,
): LinkedPivotRefreshImpact {
	const linkedNames = new Set(pivotTableNames)
	const linkedTables = workbook.pivotTables.filter(
		(pivot) =>
			(pivot.name !== undefined && linkedNames.has(pivot.name)) ||
			(pivotCacheId !== undefined && pivot.cacheId === pivotCacheId),
	)
	const cacheIds = uniqueNumbers([
		...(pivotCacheId !== undefined ? [pivotCacheId] : []),
		...linkedTables.flatMap((pivot) => (pivot.cacheId !== undefined ? [pivot.cacheId] : [])),
	])
	const markedCacheIds: number[] = []
	const cachePartPaths: string[] = []
	for (const cache of workbook.pivotCaches) {
		if (cache.cacheId === undefined || !cacheIds.includes(cache.cacheId)) continue
		Object.assign(cache, { invalid: true, refreshOnLoad: true })
		markedCacheIds.push(cache.cacheId)
		cachePartPaths.push(cache.partPath)
	}
	return {
		pivotTables: uniqueStrings(linkedTables.flatMap((pivot) => (pivot.name ? [pivot.name] : []))),
		pivotSheets: uniqueStrings(linkedTables.map((pivot) => pivot.sheetName)),
		cacheIds: uniqueNumbers(markedCacheIds),
		cachePartPaths: uniqueStrings(cachePartPaths),
	}
}

function uniqueNumbers(values: readonly number[]): readonly number[] {
	return [...new Set(values)]
}

function uniqueStrings(values: readonly string[]): readonly string[] {
	return [...new Set(values)]
}
