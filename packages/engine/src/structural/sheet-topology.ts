import type {
	AutoFilter,
	FilterColumn,
	Sheet,
	SheetAdvancedFilterInfo,
	SortState,
	Table,
	TableColumn,
} from '@ascend/core'
import { parseRange } from '@ascend/core'
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
	shiftRowDefs(sheet, axis, at, delta)
	shiftColDefs(sheet, axis, at, delta)
	shiftSqrefEntries(sheet.dataValidations, axis, at, delta)
	shiftConditionalFormats(sheet.conditionalFormats, axis, at, delta)
	shiftX14SqrefEntries(sheet.x14DataValidations, axis, at, delta)
	shiftX14SqrefEntries(sheet.x14ConditionalFormats, axis, at, delta)
	shiftIgnoredErrors(sheet.ignoredErrors, axis, at, delta)
	shiftSheetAutoFilter(sheet, axis, at, delta)
	shiftSheetSortState(sheet, axis, at, delta)
	shiftAdvancedFilters(sheet, axis, at, delta)
	shiftSheetTables(sheet, axis, at, delta)
	rewriteSheetMetadataFormulasForShift(sheet, axis, at, delta)
}

function shiftRowDefs(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	if (axis !== 'row' || sheet.rowDefs.size === 0) return
	const entries = [...sheet.rowDefs.entries()]
	sheet.rowDefs.clear()
	for (const [index, value] of entries) {
		const next = shiftIndex(index, at, delta)
		if (next !== null) sheet.rowDefs.set(next, value)
	}
}

function shiftColDefs(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	if (axis !== 'col' || sheet.colDefs.length === 0) return
	const shifted = sheet.colDefs
		.map((def) => {
			const ref = shiftRangeRef(
				{ start: { row: 0, col: def.min }, end: { row: 0, col: def.max } },
				'col',
				at,
				delta,
			)
			return ref
				? {
						...def,
						min: ref.start.col,
						max: ref.end.col,
					}
				: null
		})
		.filter((def): def is NonNullable<typeof def> => def !== null)
	sheet.colDefs = shifted
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

function shiftX14SqrefEntries(
	entries: Array<{ sqref: string; deleted?: boolean }>,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (let i = 0; i < entries.length; i++) {
		const entry = entries[i]
		if (!entry || entry.deleted) continue
		const next = shiftSqref(entry.sqref, axis, at, delta)
		entries[i] = next ? { ...entry, sqref: next } : { ...entry, sqref: '', deleted: true }
	}
}

function shiftSheetAutoFilter(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	if (!sheet.autoFilter) return
	sheet.autoFilter = shiftAutoFilter(sheet.autoFilter, axis, at, delta) ?? null
	if (!sheet.autoFilter?.sortState) sheet.preservedAutoFilterSortStateAttributes = null
}

function shiftSheetSortState(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	if (!sheet.sortState) return
	sheet.sortState = shiftSortState(sheet.sortState, axis, at, delta) ?? null
	if (!sheet.sortState) sheet.preservedSortStateAttributes = null
}

function shiftAdvancedFilters(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	if (sheet.advancedFilters.length === 0) return
	sheet.advancedFilters = sheet.advancedFilters.map((filter) =>
		shiftAdvancedFilter(filter, axis, at, delta),
	)
}

function shiftAdvancedFilter(
	filter: SheetAdvancedFilterInfo,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): SheetAdvancedFilterInfo {
	if (!filter.autoFilter) return filter
	const autoFilter = shiftAutoFilter(filter.autoFilter, axis, at, delta)
	const { autoFilter: _autoFilter, ref: _ref, ...rest } = filter
	if (!autoFilter) {
		return {
			...rest,
			filterColumnCount: 0,
			sortConditionCount: 0,
		}
	}
	return {
		...rest,
		ref: autoFilter.ref,
		autoFilter,
		filterColumnCount: autoFilter.columns.length,
		sortConditionCount: autoFilter.sortState?.conditions.length ?? 0,
	}
}

function shiftSheetTables(sheet: Sheet, axis: 'row' | 'col', at: number, delta: number): void {
	for (let index = 0; index < sheet.tables.length; index++) {
		const table = sheet.tables[index]
		if (!table) continue
		const ref = shiftRangeRef(table.ref, axis, at, delta)
		if (!ref) continue
		const autoFilter = shiftAutoFilter(table.autoFilter, axis, at, delta)
		const sortState = shiftSortState(table.sortState, axis, at, delta)
		const columns = shiftTableColumns(sheet, table, ref, axis, at, delta)
		const { autoFilter: _autoFilter, sortState: _tableSortState, ...shiftedTable } = table
		sheet.tables[index] = {
			...shiftedTable,
			ref,
			columns,
			...(autoFilter ? { autoFilter } : {}),
			...(sortState ? { sortState } : {}),
		}
	}
}

function shiftTableColumns(
	sheet: Sheet,
	table: Table,
	nextRef: Table['ref'],
	axis: 'row' | 'col',
	at: number,
	delta: number,
): readonly TableColumn[] {
	if (axis !== 'col') return table.columns.map((column) => ({ ...column }))
	const nextStart = nextRef.start.col
	const nextEnd = nextRef.end.col
	const nextWidth = nextEnd - nextStart + 1
	const shiftedByOffset = new Map<number, TableColumn>()
	for (const [index, column] of table.columns.entries()) {
		const oldColumn = table.ref.start.col + index
		const nextColumn = shiftIndex(oldColumn, at, delta)
		if (nextColumn === null || nextColumn < nextStart || nextColumn > nextEnd) continue
		shiftedByOffset.set(nextColumn - nextStart, { ...column })
	}
	const hasExplicitIds = table.columns.some((column) => column.id !== undefined)
	let nextId = Math.max(0, ...table.columns.map((column) => column.id ?? 0)) + 1
	const usedNames = new Set(
		[...shiftedByOffset.values()].map((column) => column.name.toLowerCase()),
	)
	const columns: TableColumn[] = []
	for (let offset = 0; offset < nextWidth; offset++) {
		const shifted = shiftedByOffset.get(offset)
		if (shifted) {
			columns.push(shifted)
			continue
		}
		const name = nextTableColumnName(sheet, nextRef, offset, usedNames)
		columns.push({
			name,
			...(hasExplicitIds ? { id: nextId++ } : {}),
		})
	}
	return columns
}

function nextTableColumnName(
	sheet: Sheet,
	ref: Table['ref'],
	offset: number,
	usedNames: Set<string>,
): string {
	const cellValue = sheet.cells.get(ref.start.row, ref.start.col + offset)?.value
	const base =
		cellValue?.kind === 'string' && cellValue.value.trim() !== ''
			? cellValue.value.trim()
			: `Column${offset + 1}`
	let candidate = base
	let suffix = 2
	while (usedNames.has(candidate.toLowerCase())) {
		candidate = `${base}_${suffix}`
		suffix++
	}
	usedNames.add(candidate.toLowerCase())
	return candidate
}

function shiftAutoFilter(
	autoFilter: AutoFilter | null | undefined,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): AutoFilter | undefined {
	if (!autoFilter) return undefined
	const ref = shiftSqref(autoFilter.ref, axis, at, delta)
	if (!ref) return undefined
	const sortState = shiftSortState(autoFilter.sortState, axis, at, delta)
	const columns = shiftFilterColumns(autoFilter.columns, autoFilter.ref, ref, axis, at, delta)
	const { sortState: _sortState, ref: _ref, columns: _columns, ...shiftedAutoFilter } = autoFilter
	return {
		...shiftedAutoFilter,
		ref,
		columns,
		...(sortState ? { sortState } : {}),
	}
}

function shiftFilterColumns(
	columns: readonly FilterColumn[],
	oldRef: string,
	nextRef: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): readonly FilterColumn[] {
	if (axis !== 'col' || columns.length === 0) return columns.map((column) => ({ ...column }))
	const oldRange = parseSingleRange(oldRef)
	const nextRange = parseSingleRange(nextRef)
	if (!oldRange || !nextRange) return columns.map((column) => ({ ...column }))
	return columns
		.map((column) => {
			const oldColumn = oldRange.start.col + column.colId
			const nextColumn = shiftIndex(oldColumn, at, delta)
			if (
				nextColumn === null ||
				nextColumn < nextRange.start.col ||
				nextColumn > nextRange.end.col
			) {
				return null
			}
			return { ...column, colId: nextColumn - nextRange.start.col }
		})
		.filter((column): column is FilterColumn => column !== null)
		.sort((left, right) => left.colId - right.colId)
}

function parseSingleRange(ref: string) {
	if (ref.trim().includes(' ')) return null
	try {
		return parseRange(ref)
	} catch {
		return null
	}
}

function shiftSortState(
	sortState: SortState | null | undefined,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): SortState | undefined {
	if (!sortState) return undefined
	const ref = shiftSqref(sortState.ref, axis, at, delta)
	if (!ref) return undefined
	const conditions = sortState.conditions
		.map((condition) => {
			const conditionRef = shiftSqref(condition.ref, axis, at, delta)
			return conditionRef ? { ...condition, ref: conditionRef } : null
		})
		.filter((condition): condition is SortState['conditions'][number] => condition !== null)
	if (conditions.length === 0) return undefined
	return { ...sortState, ref, conditions }
}

export function expandTableRefRows(ref: string, count: number): string {
	return expandSqrefRows(ref, count)
}
