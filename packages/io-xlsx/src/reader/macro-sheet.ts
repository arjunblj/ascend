import type { SheetState, WorkbookMacroSheetInfo } from '@ascend/core'
import type { Relationship } from './relationships.ts'

const DIMENSION_RE = /<(?:[A-Za-z_][\w.-]*:)?dimension\b[^>]*\bref="([^"]+)"/
const CELL_RE = /<(?:[A-Za-z_][\w.-]*:)?c\b/g
const FORMULA_RE = /<(?:[A-Za-z_][\w.-]*:)?f\b/g

export function parseMacroSheetInfo(
	xml: string | undefined,
	input: {
		readonly name: string
		readonly sheetId: string
		readonly relId: string
		readonly partPath: string
		readonly state: SheetState
		readonly relationships: readonly Relationship[]
	},
): WorkbookMacroSheetInfo {
	const dimensionRef = xml ? DIMENSION_RE.exec(xml)?.[1] : undefined
	const cellCount = xml ? countMatches(xml, CELL_RE) : undefined
	const formulaCount = xml ? countMatches(xml, FORMULA_RE) : undefined
	return {
		name: input.name,
		sheetId: input.sheetId,
		relId: input.relId,
		partPath: input.partPath,
		state: input.state,
		relationshipCount: input.relationships.length,
		...(dimensionRef ? { dimensionRef } : {}),
		...(cellCount !== undefined ? { cellCount } : {}),
		...(formulaCount !== undefined ? { formulaCount } : {}),
	}
}

function countMatches(xml: string, pattern: RegExp): number {
	pattern.lastIndex = 0
	let count = 0
	while (pattern.exec(xml)) count++
	return count
}
