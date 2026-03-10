import type { Sheet } from '@ascend/core'
import { rewriteSheetMetadataFormulasForShift } from './formula-rewrite.ts'
import {
	expandSqrefRows,
	shiftA1RangeOrCell,
	shiftA1Ref,
	shiftIndex,
	shiftRangeRef,
	shiftSqref,
} from './ref-shift.ts'

export function shiftSheetCellMetadata(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	shiftMappedRefs(sheet.comments, axis, at, delta)
	shiftMappedRefs(sheet.hyperlinks, axis, at, delta)
	rewriteHyperlinkLocationsForShift(sheet, axis, at, delta)
	shiftRowOrColMap(sheet.rowHeights, axis === 'row', at, delta)
	shiftRowOrColMap(sheet.colWidths, axis === 'col', at, delta)
	shiftSqrefEntries(sheet.dataValidations, axis, at, delta)
	shiftConditionalFormats(sheet.conditionalFormats, axis, at, delta)
	shiftIgnoredErrors(sheet.ignoredErrors, axis, at, delta)
	shiftSheetAutoFilter(sheet, axis, at, delta)
	shiftSheetTables(sheet, axis, at, delta)
	rewriteSheetMetadataFormulasForShift(sheet, axis, at, delta)
}

export function renameHyperlinkLocation(
	location: string | undefined,
	oldName: string,
	newName: string,
): string | undefined {
	if (!location) return location
	const split = splitSheetQualifiedRef(location)
	if (!split || split.sheet !== oldName) return location
	return `${newName}!${split.ref}`
}

function rewriteHyperlinkLocationsForShift(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (const [ref, hyperlink] of sheet.hyperlinks) {
		const location = shiftHyperlinkLocation(hyperlink.location, sheet.name, axis, at, delta)
		if (location === hyperlink.location) continue
		sheet.hyperlinks.set(ref, { ...hyperlink, ...(location !== undefined ? { location } : {}) })
	}
}

function shiftHyperlinkLocation(
	location: string | undefined,
	sheetName: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): string | undefined {
	if (!location) return location
	const split = splitSheetQualifiedRef(location)
	if (!split || split.sheet !== sheetName) return location
	const shifted = shiftA1RangeOrCell(split.ref, axis, at, delta)
	return shifted ? `${split.sheet}!${shifted}` : location
}

function splitSheetQualifiedRef(input: string): { sheet: string; ref: string } | null {
	const bang = input.lastIndexOf('!')
	if (bang === -1) return null
	const sheet = input.slice(0, bang).replace(/^'|'$/g, '')
	const ref = input.slice(bang + 1)
	return sheet && ref ? { sheet, ref } : null
}

function shiftMappedRefs<T>(
	map: Map<string, T>,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	const entries = [...map.entries()]
	map.clear()
	for (const [ref, value] of entries) {
		const next = shiftA1Ref(ref, axis, at, delta)
		if (next) map.set(next, value)
	}
}

function shiftRowOrColMap(
	map: Map<number, number>,
	active: boolean,
	at: number,
	delta: number,
): void {
	if (!active || map.size === 0) return
	const entries = [...map.entries()]
	map.clear()
	for (const [index, value] of entries) {
		const next = shiftIndex(index, at, delta)
		if (next !== null) map.set(next, value)
	}
}

function shiftSqrefEntries(
	entries: Array<{ sqref: string }>,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (let i = entries.length - 1; i >= 0; i--) {
		const entry = entries[i]
		if (!entry) continue
		const next = shiftSqref(entry.sqref, axis, at, delta)
		if (!next) entries.splice(i, 1)
		else entries[i] = { ...entry, sqref: next }
	}
}

function shiftConditionalFormats(
	formats: Array<{ sqref: string }>,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	shiftSqrefEntries(formats, axis, at, delta)
}

function shiftIgnoredErrors(
	ignoredErrors: Array<{ sqref: string }>,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	shiftSqrefEntries(ignoredErrors, axis, at, delta)
}

function shiftSheetAutoFilter(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	if (!sheet.autoFilter) return
	const ref = shiftSqref(sheet.autoFilter.ref, axis, at, delta)
	if (!ref) {
		sheet.autoFilter = null
		return
	}
	sheet.autoFilter = {
		...sheet.autoFilter,
		ref,
		...(sheet.autoFilter.sortState
			? {
					sortState: {
						...sheet.autoFilter.sortState,
						ref:
							shiftSqref(sheet.autoFilter.sortState.ref, axis, at, delta) ??
							sheet.autoFilter.sortState.ref,
						conditions: sheet.autoFilter.sortState.conditions.map((condition) => ({
							...condition,
							ref: shiftSqref(condition.ref, axis, at, delta) ?? condition.ref,
						})),
					},
				}
			: {}),
	}
}

function shiftSheetTables(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	for (let index = 0; index < sheet.tables.length; index++) {
		const table = sheet.tables[index]
		if (!table) continue
		const ref = shiftRangeRef(table.ref, axis, at, delta)
		if (!ref) continue
		sheet.tables[index] = {
			...table,
			ref,
			...(table.autoFilter
				? {
						autoFilter: {
							...table.autoFilter,
							ref: shiftSqref(table.autoFilter.ref, axis, at, delta) ?? table.autoFilter.ref,
							...(table.autoFilter.sortState
								? {
										sortState: {
											...table.autoFilter.sortState,
											ref:
												shiftSqref(table.autoFilter.sortState.ref, axis, at, delta) ??
												table.autoFilter.sortState.ref,
											conditions: table.autoFilter.sortState.conditions.map((condition) => ({
												...condition,
												ref: shiftSqref(condition.ref, axis, at, delta) ?? condition.ref,
											})),
										},
									}
								: {}),
						},
					}
				: {}),
			...(table.sortState
				? {
						sortState: {
							...table.sortState,
							ref: shiftSqref(table.sortState.ref, axis, at, delta) ?? table.sortState.ref,
							conditions: table.sortState.conditions.map((condition) => ({
								...condition,
								ref: shiftSqref(condition.ref, axis, at, delta) ?? condition.ref,
							})),
						},
					}
				: {}),
		}
	}
}

export function expandTableRefRows(ref: string, count: number): string {
	return expandSqrefRows(ref, count)
}
