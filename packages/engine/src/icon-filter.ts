import type {
	FilterColumn,
	Sheet,
	SheetConditionalFormatRule,
	SheetConditionalFormatValueObject,
} from '@ascend/core'
import { parseRange } from '@ascend/core'
import type { CellValue } from '@ascend/schema'

export interface IconFilterRange {
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
}

interface IconFilterContext {
	readonly priority: number
	readonly rule: SheetConditionalFormatRule
	readonly sqref: string
	readonly thresholds: readonly (number | undefined)[]
}

interface IconValueStats {
	readonly values: readonly number[]
	readonly min?: number
	readonly max?: number
}

const DEFAULT_ICON_SET = '3TrafficLights1'

export function computeIconFilterRows(
	sheet: Sheet,
	range: IconFilterRange,
	column: FilterColumn,
): ReadonlySet<number> | undefined {
	if (column.iconId === undefined || !Number.isInteger(column.iconId)) return undefined
	const targetIconSet = column.iconSet ?? DEFAULT_ICON_SET
	const col = range.startCol + column.colId
	const contexts = collectIconFilterContexts(sheet, targetIconSet, col)
	if (contexts.length === 0) return undefined
	const rows = new Set<number>()
	for (let row = range.startRow + 1; row <= range.endRow; row++) {
		const value = iconComparableNumber(sheet.cells.readValue(row, col))
		if (value === null) continue
		for (const context of contexts) {
			if (!cellInSqref(context.sqref, row, col)) continue
			const iconId = iconIdForValue(value, context)
			if (iconId === column.iconId) rows.add(row)
			break
		}
	}
	return rows
}

function collectIconFilterContexts(
	sheet: Sheet,
	targetIconSet: string,
	col: number,
): readonly IconFilterContext[] {
	const contexts: IconFilterContext[] = []
	let fallbackPriority = 1
	for (const cf of sheet.conditionalFormats) {
		if (!sqrefTouchesColumn(cf.sqref, col)) continue
		for (const rule of cf.rules) {
			const priority = rule.priority ?? fallbackPriority++
			const iconSet = rule.iconSet
			if (!iconSet) continue
			const ruleIconSet = iconSet.iconSet ?? DEFAULT_ICON_SET
			if (ruleIconSet !== targetIconSet) continue
			const valueStats = collectIconRuleValues(sheet, cf.sqref)
			const thresholds = iconSet.cfvo.map((cfvo) =>
				resolveIconThreshold(cfvo, iconSet.percent, valueStats),
			)
			if (
				thresholds.length < 2 ||
				thresholds.slice(1).some((threshold) => threshold === undefined)
			) {
				continue
			}
			contexts.push({ priority, rule, sqref: cf.sqref, thresholds })
		}
	}
	contexts.sort((a, b) => a.priority - b.priority)
	return contexts
}

function iconIdForValue(value: number, context: IconFilterContext): number | null {
	const iconCount = context.thresholds.length
	if (iconCount < 2) return null
	let bucket = 0
	for (let index = 1; index < iconCount; index++) {
		const threshold = context.thresholds[index]
		if (threshold === undefined) return null
		const gte = context.rule.iconSet?.cfvo[index]?.gte !== false
		const passes = gte ? value >= threshold : value > threshold
		if (passes) bucket = index
	}
	return context.rule.iconSet?.reverse === true ? iconCount - 1 - bucket : bucket
}

function resolveIconThreshold(
	cfvo: SheetConditionalFormatValueObject,
	iconSetPercent: boolean | undefined,
	stats: IconValueStats,
): number | undefined {
	const type = cfvo.type ?? (iconSetPercent === false ? 'num' : 'percent')
	switch (type) {
		case 'min':
			return stats.min
		case 'max':
			return stats.max
		case 'num':
			return parseFiniteNumber(cfvo.value)
		case 'percent': {
			const pct = parseFiniteNumber(cfvo.value)
			if (pct === undefined || stats.min === undefined || stats.max === undefined) return undefined
			return stats.min + ((stats.max - stats.min) * pct) / 100
		}
		case 'percentile': {
			const percentile = parseFiniteNumber(cfvo.value)
			return percentile === undefined ? undefined : percentileInclusive(stats.values, percentile)
		}
		default:
			return undefined
	}
}

function collectIconRuleValues(sheet: Sheet, sqref: string): IconValueStats {
	const values: number[] = []
	forEachSqrefCell(sqref, (row, col) => {
		const value = iconComparableNumber(sheet.cells.readValue(row, col))
		if (value !== null) values.push(value)
	})
	values.sort((a, b) => a - b)
	const min = values[0]
	const max = values[values.length - 1]
	return {
		values,
		...(min !== undefined ? { min } : {}),
		...(max !== undefined ? { max } : {}),
	}
}

function iconComparableNumber(value: CellValue): number | null {
	if (value.kind === 'number') return Number.isFinite(value.value) ? value.value : null
	if (value.kind === 'date') return Number.isFinite(value.serial) ? value.serial : null
	return null
}

function percentileInclusive(values: readonly number[], percentile: number): number | undefined {
	if (values.length === 0 || !Number.isFinite(percentile)) return undefined
	if (values.length === 1) return values[0]
	const clamped = Math.min(100, Math.max(0, percentile))
	const rank = (clamped / 100) * (values.length - 1)
	const lower = Math.floor(rank)
	const upper = Math.ceil(rank)
	const lowerValue = values[lower]
	const upperValue = values[upper]
	if (lowerValue === undefined || upperValue === undefined) return undefined
	return lower === upper ? lowerValue : lowerValue + (upperValue - lowerValue) * (rank - lower)
}

function parseFiniteNumber(value: string | undefined): number | undefined {
	if (value === undefined) return undefined
	const parsed = Number(value)
	return Number.isFinite(parsed) ? parsed : undefined
}

function sqrefTouchesColumn(sqref: string, col: number): boolean {
	let touches = false
	forEachSqrefRange(sqref, (startRow, startCol, endRow, endCol) => {
		if (startRow <= endRow && col >= startCol && col <= endCol) touches = true
	})
	return touches
}

function cellInSqref(sqref: string, row: number, col: number): boolean {
	let contains = false
	forEachSqrefRange(sqref, (startRow, startCol, endRow, endCol) => {
		if (row >= startRow && row <= endRow && col >= startCol && col <= endCol) contains = true
	})
	return contains
}

function forEachSqrefCell(sqref: string, fn: (row: number, col: number) => void): void {
	forEachSqrefRange(sqref, (startRow, startCol, endRow, endCol) => {
		for (let row = startRow; row <= endRow; row++) {
			for (let col = startCol; col <= endCol; col++) fn(row, col)
		}
	})
}

function forEachSqrefRange(
	sqref: string,
	fn: (startRow: number, startCol: number, endRow: number, endCol: number) => void,
): void {
	for (const part of sqref.split(/\s+/)) {
		if (part.length === 0) continue
		try {
			const parsed = parseRange(part)
			fn(parsed.start.row, parsed.start.col, parsed.end.row, parsed.end.col)
		} catch {}
	}
}
