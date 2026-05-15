import type { RangeRef, SheetInspectInfo } from '@ascend/sdk'
import { indexToColumn, parseRange } from '@ascend/sdk'
import type {
	GridSemanticCell,
	GridSemanticFlag,
	GridSemanticModel,
	ViewportState,
} from '../runtime/types.ts'

interface Bounds {
	readonly startRow: number
	readonly endRow: number
	readonly startCol: number
	readonly endCol: number
}

export function buildGridSemanticModel(input: {
	readonly sheet: SheetInspectInfo | undefined
	readonly viewport: ViewportState
}): GridSemanticModel {
	const cells = new Map<string, GridSemanticCell>()
	const sheet = input.sheet
	const viewport = viewportBounds(input.viewport)
	if (!sheet) {
		return { cells, frozenRows: 0, frozenCols: 0, protected: false, activeFilterRanges: [] }
	}

	for (const table of sheet.tables ?? []) {
		const bounds = rangeRefBounds(table.ref)
		if (!intersects(bounds, viewport)) continue
		addRangeFlag(cells, bounds, viewport, 'table')
		if (table.hasHeaders) addRowFlag(cells, bounds.startRow, bounds, viewport, 'tableHeader')
		if (table.hasTotals) addRowFlag(cells, bounds.endRow, bounds, viewport, 'tableTotal')
		if (table.autoFilter) {
			addRowFlag(cells, bounds.startRow, bounds, viewport, 'filterAvailable')
			if (hasActiveFilter(table.autoFilter)) {
				addRowFlag(cells, bounds.startRow, bounds, viewport, 'filterActive')
			}
		}
		if (table.sortState) addSortFlags(cells, table.sortState.conditions, viewport)
	}

	if (sheet.autoFilter) {
		const bounds = parseBounds(sheet.autoFilter.ref)
		if (bounds && intersects(bounds, viewport)) {
			addRowFlag(cells, bounds.startRow, bounds, viewport, 'filterAvailable')
			if (hasActiveFilter(sheet.autoFilter))
				addRowFlag(cells, bounds.startRow, bounds, viewport, 'filterActive')
			if (sheet.autoFilter.sortState)
				addSortFlags(cells, sheet.autoFilter.sortState.conditions, viewport)
		}
	}

	for (const comment of sheet.comments ?? []) addRefFlag(cells, comment.ref, viewport, 'comment')
	for (const link of sheet.hyperlinks ?? []) addRefFlag(cells, link.ref, viewport, 'hyperlink')
	for (const validation of sheet.dataValidations ?? []) {
		for (const bounds of sqrefBounds(validation.sqref)) {
			addRangeFlag(cells, bounds, viewport, 'validationDropdown')
		}
	}
	for (const format of sheet.conditionalFormats ?? []) {
		for (const bounds of sqrefBounds(format.sqref))
			addRangeFlag(cells, bounds, viewport, 'conditionalFormat')
	}
	for (const ignored of sheet.ignoredErrors ?? []) {
		for (const bounds of sqrefBounds(ignored.sqref))
			addRangeFlag(cells, bounds, viewport, 'validationInvalid')
	}
	if (sheet.protection) addRangeFlag(cells, viewport, viewport, 'protected')

	return {
		cells,
		frozenRows: sheet.hasFrozenPanes ? 1 : 0,
		frozenCols: sheet.hasFrozenPanes ? 1 : 0,
		protected: sheet.protection !== null,
		activeFilterRanges: activeFilterRanges(sheet),
	}
}

function addSortFlags(
	cells: Map<string, GridSemanticCell>,
	conditions: readonly { readonly ref?: string; readonly descending?: boolean }[],
	viewport: Bounds,
): void {
	for (const condition of conditions) {
		if (!condition.ref) continue
		const bounds = parseBounds(condition.ref)
		if (!bounds) continue
		addRangeFlag(cells, bounds, viewport, condition.descending ? 'sortDesc' : 'sortAsc')
	}
}

function addRefFlag(
	cells: Map<string, GridSemanticCell>,
	ref: string,
	viewport: Bounds,
	flag: GridSemanticFlag,
): void {
	const bounds = parseBounds(ref)
	if (!bounds) return
	addRangeFlag(cells, bounds, viewport, flag)
}

function addRowFlag(
	cells: Map<string, GridSemanticCell>,
	row: number,
	range: Bounds,
	viewport: Bounds,
	flag: GridSemanticFlag,
): void {
	addRangeFlag(cells, { ...range, startRow: row, endRow: row }, viewport, flag)
}

function addRangeFlag(
	cells: Map<string, GridSemanticCell>,
	range: Bounds,
	viewport: Bounds,
	flag: GridSemanticFlag,
): void {
	if (!intersects(range, viewport)) return
	const startRow = Math.max(range.startRow, viewport.startRow)
	const endRow = Math.min(range.endRow, viewport.endRow)
	const startCol = Math.max(range.startCol, viewport.startCol)
	const endCol = Math.min(range.endCol, viewport.endCol)
	for (let row = startRow; row <= endRow; row++) {
		for (let col = startCol; col <= endCol; col++) {
			const ref = `${indexToColumn(col)}${row + 1}`
			const existing = cells.get(ref)
			const flags = existing?.flags.includes(flag)
				? existing.flags
				: [...(existing?.flags ?? []), flag]
			cells.set(ref, { ref, row, col, flags })
		}
	}
}

function sqrefBounds(sqref: string): readonly Bounds[] {
	return sqref
		.split(/\s+/)
		.map(parseBounds)
		.filter((bounds): bounds is Bounds => bounds !== null)
}

function parseBounds(ref: string): Bounds | null {
	try {
		return rangeRefBounds(parseRange(ref))
	} catch {
		return null
	}
}

function rangeRefBounds(ref: RangeRef): Bounds {
	return {
		startRow: Math.min(ref.start.row, ref.end.row),
		endRow: Math.max(ref.start.row, ref.end.row),
		startCol: Math.min(ref.start.col, ref.end.col),
		endCol: Math.max(ref.start.col, ref.end.col),
	}
}

function viewportBounds(viewport: ViewportState): Bounds {
	return {
		startRow: viewport.topRow,
		endRow: viewport.topRow + viewport.visibleRows - 1,
		startCol: viewport.leftCol,
		endCol: viewport.leftCol + viewport.visibleCols - 1,
	}
}

function intersects(left: Bounds, right: Bounds): boolean {
	return (
		left.startRow <= right.endRow &&
		left.endRow >= right.startRow &&
		left.startCol <= right.endCol &&
		left.endCol >= right.startCol
	)
}

function hasActiveFilter(filter: { readonly columns?: readonly unknown[] }): boolean {
	return (filter.columns?.length ?? 0) > 0
}

function activeFilterRanges(sheet: SheetInspectInfo): readonly string[] {
	const ranges: string[] = []
	if (sheet.autoFilter && hasActiveFilter(sheet.autoFilter)) ranges.push(sheet.autoFilter.ref)
	for (const table of sheet.tables ?? []) {
		if (table.autoFilter && hasActiveFilter(table.autoFilter)) ranges.push(table.autoFilter.ref)
	}
	return ranges
}
