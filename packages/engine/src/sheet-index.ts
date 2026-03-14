import type { Workbook } from '@ascend/core'

const sheetIndexCache = new WeakMap<Workbook, Map<string, number>>()

export function resolveSheetIndexByMap(
	sheetNameIndex: ReadonlyMap<string, number>,
	sheetName: string | undefined,
	currentSheet: number,
): number {
	if (sheetName === undefined) return currentSheet
	return sheetNameIndex.get(sheetName.toLowerCase()) ?? -1
}

export function resolveSheetIndexInWorkbook(
	workbook: Workbook,
	sheetName: string | undefined,
	currentSheet: number,
): number {
	if (sheetName === undefined) return currentSheet
	let cache = sheetIndexCache.get(workbook)
	if (!cache) {
		cache = new Map()
		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			if (sheet) cache.set(sheet.name.toLowerCase(), i)
		}
		sheetIndexCache.set(workbook, cache)
	}
	return cache.get(sheetName.toLowerCase()) ?? -1
}

export function invalidateSheetIndexCache(workbook: Workbook): void {
	sheetIndexCache.delete(workbook)
}
