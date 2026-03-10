import type { Cell, Sheet, SheetColDef } from '@ascend/core'
import { indexToColumn } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import { topLeftScalar } from '@ascend/schema'
import { escapeXml } from '../xml.ts'
import { autoFilterXml } from './filtering.ts'
import type { SharedStringTable } from './shared-strings.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface SheetXmlOptions {
	readonly tableRelIds?: readonly string[]
	readonly drawingRelId?: string
	readonly hyperlinks?: readonly {
		ref: string
		relId?: string
		location?: string
		display?: string
		tooltip?: string
	}[]
	readonly legacyDrawingRelId?: string
}

export function buildSheetXml(
	sheet: Sheet,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
	options: SheetXmlOptions = {},
): string {
	const tableRelIds = options.tableRelIds ?? []
	const hyperlinks = options.hyperlinks ?? []
	const drawingRelId = options.drawingRelId
	const legacyDrawingRelId = options.legacyDrawingRelId
	const worksheetAttrs = [`xmlns="${NS}"`]
	if (
		tableRelIds.length > 0 ||
		hyperlinks.some((link) => link.relId) ||
		drawingRelId ||
		legacyDrawingRelId
	) {
		worksheetAttrs.push(`xmlns:r="${NS_R}"`)
	}
	const parts: string[] = [XML_HEADER, `<worksheet ${worksheetAttrs.join(' ')}>`]

	if (sheet.tabColor) {
		parts.push('<sheetPr>')
		const tcAttrs: string[] = []
		if (sheet.tabColor.rgb) tcAttrs.push(`rgb="${sheet.tabColor.rgb}"`)
		if (sheet.tabColor.theme !== undefined) tcAttrs.push(`theme="${sheet.tabColor.theme}"`)
		if (sheet.tabColor.tint !== undefined) tcAttrs.push(`tint="${sheet.tabColor.tint}"`)
		if (sheet.tabColor.indexed !== undefined) tcAttrs.push(`indexed="${sheet.tabColor.indexed}"`)
		parts.push(`<tabColor ${tcAttrs.join(' ')}/>`)
		parts.push('</sheetPr>')
	}

	if (sheet.sheetFormatPr) {
		const fmtAttrs: string[] = []
		if (sheet.sheetFormatPr.defaultRowHeight !== undefined)
			fmtAttrs.push(`defaultRowHeight="${sheet.sheetFormatPr.defaultRowHeight}"`)
		if (sheet.sheetFormatPr.defaultColWidth !== undefined)
			fmtAttrs.push(`defaultColWidth="${sheet.sheetFormatPr.defaultColWidth}"`)
		if (sheet.sheetFormatPr.outlineLevelRow !== undefined)
			fmtAttrs.push(`outlineLevelRow="${sheet.sheetFormatPr.outlineLevelRow}"`)
		if (sheet.sheetFormatPr.outlineLevelCol !== undefined)
			fmtAttrs.push(`outlineLevelCol="${sheet.sheetFormatPr.outlineLevelCol}"`)
		if (sheet.sheetFormatPr.customHeight) fmtAttrs.push('customHeight="1"')
		if (fmtAttrs.length > 0) parts.push(`<sheetFormatPr ${fmtAttrs.join(' ')}/>`)
	}

	const usedRange = sheet.cells.usedRange()
	if (usedRange) {
		const s = `${indexToColumn(usedRange.start.col)}${usedRange.start.row + 1}`
		const e = `${indexToColumn(usedRange.end.col)}${usedRange.end.row + 1}`
		parts.push(`<dimension ref="${s}:${e}"/>`)
	}

	if (sheet.frozenRows > 0 || sheet.frozenCols > 0) {
		const paneAttrs: string[] = ['state="frozen"']
		if (sheet.frozenCols > 0) paneAttrs.push(`xSplit="${sheet.frozenCols}"`)
		if (sheet.frozenRows > 0) paneAttrs.push(`ySplit="${sheet.frozenRows}"`)
		parts.push('<sheetViews>')
		parts.push('<sheetView workbookViewId="0">')
		parts.push(`<pane ${paneAttrs.join(' ')}/>`)
		parts.push('</sheetView>')
		parts.push('</sheetViews>')
	}

	if (sheet.colWidths.size > 0 || sheet.colDefs.length > 0) {
		parts.push('<cols>')
		const colDefs: readonly SheetColDef[] =
			sheet.colDefs.length > 0 ? sheet.colDefs : groupColumnWidths(sheet)
		for (const group of colDefs) {
			const attrs = [`min="${group.min + 1}"`, `max="${group.max + 1}"`]
			if (group.width !== undefined) attrs.push(`width="${group.width}"`)
			if (group.style !== undefined) attrs.push(`style="${group.style}"`)
			if (group.hidden) attrs.push('hidden="1"')
			if (group.bestFit) attrs.push('bestFit="1"')
			if (group.collapsed) attrs.push('collapsed="1"')
			if (group.outlineLevel !== undefined) attrs.push(`outlineLevel="${group.outlineLevel}"`)
			if (group.customWidth ?? group.width !== undefined) attrs.push('customWidth="1"')
			parts.push(`<col ${attrs.join(' ')}/>`)
		}
		parts.push('</cols>')
	}

	parts.push('<sheetData>')
	const rowHeights = [...sheet.rowHeights.entries()].sort((a, b) => a[0] - b[0])
	const rowIterator = sheet.cells.iterateRows()
	let nextRow = rowIterator.next()
	let rowHeightIndex = 0
	while (!nextRow.done || rowHeightIndex < rowHeights.length) {
		const populatedRow = nextRow.done ? undefined : nextRow.value
		const heightEntry = rowHeights[rowHeightIndex]
		const heightRow = heightEntry?.[0]
		const row =
			populatedRow === undefined
				? (heightRow as number)
				: heightRow === undefined
					? populatedRow[0]
					: Math.min(populatedRow[0], heightRow)
		const cells = populatedRow && populatedRow[0] === row ? populatedRow[1] : []
		const rowAttrs = [`r="${row + 1}"`]
		const rowHeight = heightEntry && heightEntry[0] === row ? heightEntry[1] : undefined
		if (rowHeight !== undefined) {
			rowAttrs.push(`ht="${rowHeight}"`)
			rowAttrs.push('customHeight="1"')
		}
		parts.push(`<row ${rowAttrs.join(' ')}>`)
		for (const [col, cell] of cells) {
			const ref = `${indexToColumn(col)}${row + 1}`
			parts.push(cellXml(ref, cell, ssTable, xfMap))
		}
		parts.push('</row>')
		if (populatedRow && populatedRow[0] === row) nextRow = rowIterator.next()
		if (heightEntry && heightEntry[0] === row) rowHeightIndex++
	}

	parts.push('</sheetData>')

	if (sheet.merges.length > 0) {
		parts.push(`<mergeCells count="${sheet.merges.length}">`)
		for (const merge of sheet.merges) {
			const s = `${indexToColumn(merge.start.col)}${merge.start.row + 1}`
			const e = `${indexToColumn(merge.end.col)}${merge.end.row + 1}`
			parts.push(`<mergeCell ref="${s}:${e}"/>`)
		}
		parts.push('</mergeCells>')
	}

	if (sheet.protection) {
		const attrs = collectProtectionAttrs(sheet.protection)
		if (attrs.length > 0) parts.push(`<sheetProtection ${attrs.join(' ')}/>`)
	}

	if (sheet.autoFilter) {
		parts.push(autoFilterXml(sheet.autoFilter))
	}

	if (sheet.conditionalFormats.length > 0) {
		for (const conditionalFormat of sheet.conditionalFormats) {
			parts.push(`<conditionalFormatting sqref="${escapeXml(conditionalFormat.sqref)}">`)
			for (const rule of conditionalFormat.rules) {
				const attrs = [`type="${escapeXml(rule.type)}"`]
				if (rule.operator) attrs.push(`operator="${escapeXml(rule.operator)}"`)
				if (rule.priority !== undefined) attrs.push(`priority="${rule.priority}"`)
				if (rule.dxfId !== undefined) attrs.push(`dxfId="${rule.dxfId}"`)
				if (rule.stopIfTrue) attrs.push('stopIfTrue="1"')
				parts.push(`<cfRule ${attrs.join(' ')}>`)
				for (const formula of rule.formulas) {
					parts.push(`<formula>${escapeXml(formula)}</formula>`)
				}
				parts.push('</cfRule>')
			}
			parts.push('</conditionalFormatting>')
		}
	}

	if (sheet.dataValidations.length > 0) {
		parts.push(`<dataValidations count="${sheet.dataValidations.length}">`)
		for (const validation of sheet.dataValidations) {
			const attrs = [`sqref="${escapeXml(validation.sqref)}"`]
			if (validation.type) attrs.push(`type="${escapeXml(validation.type)}"`)
			if (validation.operator) attrs.push(`operator="${escapeXml(validation.operator)}"`)
			if (validation.errorStyle) attrs.push(`errorStyle="${escapeXml(validation.errorStyle)}"`)
			if (validation.allowBlank !== undefined) {
				attrs.push(`allowBlank="${validation.allowBlank ? '1' : '0'}"`)
			}
			if (validation.showInputMessage !== undefined) {
				attrs.push(`showInputMessage="${validation.showInputMessage ? '1' : '0'}"`)
			}
			if (validation.showErrorMessage !== undefined) {
				attrs.push(`showErrorMessage="${validation.showErrorMessage ? '1' : '0'}"`)
			}
			if (validation.showDropDown !== undefined) {
				attrs.push(`showDropDown="${validation.showDropDown ? '1' : '0'}"`)
			}
			if (validation.promptTitle) attrs.push(`promptTitle="${escapeXml(validation.promptTitle)}"`)
			if (validation.prompt) attrs.push(`prompt="${escapeXml(validation.prompt)}"`)
			if (validation.errorTitle) attrs.push(`errorTitle="${escapeXml(validation.errorTitle)}"`)
			if (validation.error) attrs.push(`error="${escapeXml(validation.error)}"`)
			parts.push(`<dataValidation ${attrs.join(' ')}>`)
			if (validation.formula1) parts.push(`<formula1>${escapeXml(validation.formula1)}</formula1>`)
			if (validation.formula2) parts.push(`<formula2>${escapeXml(validation.formula2)}</formula2>`)
			parts.push('</dataValidation>')
		}
		parts.push('</dataValidations>')
	}

	if (hyperlinks.length > 0) {
		parts.push('<hyperlinks>')
		for (const hyperlink of hyperlinks) {
			const attrs = [`ref="${escapeXml(hyperlink.ref)}"`]
			if (hyperlink.relId) attrs.push(`r:id="${hyperlink.relId}"`)
			if (hyperlink.location) attrs.push(`location="${escapeXml(hyperlink.location)}"`)
			if (hyperlink.display) attrs.push(`display="${escapeXml(hyperlink.display)}"`)
			if (hyperlink.tooltip) attrs.push(`tooltip="${escapeXml(hyperlink.tooltip)}"`)
			parts.push(`<hyperlink ${attrs.join(' ')}/>`)
		}
		parts.push('</hyperlinks>')
	}

	if (sheet.printOptions) {
		const attrs = collectMixedAttrs(sheet.printOptions)
		if (attrs.length > 0) parts.push(`<printOptions ${attrs.join(' ')}/>`)
	}

	if (sheet.pageMargins) {
		const attrs = collectNumericAttrs(sheet.pageMargins)
		if (attrs.length > 0) parts.push(`<pageMargins ${attrs.join(' ')}/>`)
	}

	if (sheet.pageSetup) {
		const attrs = collectMixedAttrs(sheet.pageSetup)
		if (attrs.length > 0) parts.push(`<pageSetup ${attrs.join(' ')}/>`)
	}

	if (sheet.headerFooter) {
		parts.push('<headerFooter>')
		if (sheet.headerFooter.oddHeader) {
			parts.push(`<oddHeader>${escapeXml(sheet.headerFooter.oddHeader)}</oddHeader>`)
		}
		if (sheet.headerFooter.oddFooter) {
			parts.push(`<oddFooter>${escapeXml(sheet.headerFooter.oddFooter)}</oddFooter>`)
		}
		if (sheet.headerFooter.evenHeader) {
			parts.push(`<evenHeader>${escapeXml(sheet.headerFooter.evenHeader)}</evenHeader>`)
		}
		if (sheet.headerFooter.evenFooter) {
			parts.push(`<evenFooter>${escapeXml(sheet.headerFooter.evenFooter)}</evenFooter>`)
		}
		if (sheet.headerFooter.firstHeader) {
			parts.push(`<firstHeader>${escapeXml(sheet.headerFooter.firstHeader)}</firstHeader>`)
		}
		if (sheet.headerFooter.firstFooter) {
			parts.push(`<firstFooter>${escapeXml(sheet.headerFooter.firstFooter)}</firstFooter>`)
		}
		parts.push('</headerFooter>')
	}

	if (sheet.ignoredErrors.length > 0) {
		parts.push('<ignoredErrors>')
		for (const ie of sheet.ignoredErrors) {
			const attrs = [`sqref="${escapeXml(ie.sqref)}"`]
			if (ie.numberStoredAsText) attrs.push('numberStoredAsText="1"')
			if (ie.formula) attrs.push('formula="1"')
			if (ie.formulaRange) attrs.push('formulaRange="1"')
			if (ie.evalError) attrs.push('evalError="1"')
			if (ie.twoDigitTextYear) attrs.push('twoDigitTextYear="1"')
			if (ie.unlockedFormula) attrs.push('unlockedFormula="1"')
			if (ie.emptyCellReference) attrs.push('emptyCellReference="1"')
			if (ie.listDataValidation) attrs.push('listDataValidation="1"')
			if (ie.calculatedColumn) attrs.push('calculatedColumn="1"')
			parts.push(`<ignoredError ${attrs.join(' ')}/>`)
		}
		parts.push('</ignoredErrors>')
	}

	if (drawingRelId) {
		parts.push(`<drawing r:id="${drawingRelId}"/>`)
	}

	if (legacyDrawingRelId) {
		parts.push(`<legacyDrawing r:id="${legacyDrawingRelId}"/>`)
	}

	if (tableRelIds.length > 0) {
		parts.push(`<tableParts count="${tableRelIds.length}">`)
		for (const relId of tableRelIds) {
			parts.push(`<tablePart r:id="${relId}"/>`)
		}
		parts.push('</tableParts>')
	}

	if (sheet.preservedExtLst) {
		parts.push(sheet.preservedExtLst)
	}

	parts.push('</worksheet>')
	return parts.join('')
}

function groupColumnWidths(sheet: Sheet): SheetColDef[] {
	const cols = [...sheet.colWidths.entries()].sort((a, b) => a[0] - b[0])
	if (cols.length === 0) return []

	const groups: Array<{ min: number; max: number; width: number }> = []
	let [startCol, width] = cols[0] ?? [0, 0]
	let endCol = startCol

	for (let i = 1; i < cols.length; i++) {
		const entry = cols[i]
		if (!entry) continue
		const [col, colWidth] = entry
		if (col === endCol + 1 && colWidth === width) {
			endCol = col
			continue
		}
		groups.push({ min: startCol, max: endCol, width })
		startCol = col
		endCol = col
		width = colWidth
	}

	groups.push({ min: startCol, max: endCol, width })
	return groups
}

function collectNumericAttrs(values: object): string[] {
	const attrs: string[] = []
	for (const [key, value] of Object.entries(values as Record<string, number | undefined>)) {
		if (value !== undefined) attrs.push(`${key}="${value}"`)
	}
	return attrs
}

function collectMixedAttrs(values: object): string[] {
	const attrs: string[] = []
	for (const [key, value] of Object.entries(
		values as Record<string, string | number | boolean | undefined>,
	)) {
		if (value === undefined) continue
		attrs.push(
			`${key}="${typeof value === 'boolean' ? (value ? '1' : '0') : escapeXml(String(value))}"`,
		)
	}
	return attrs
}

function collectProtectionAttrs(protection: NonNullable<Sheet['protection']>): string[] {
	const attrs: string[] = []
	for (const [key, value] of Object.entries(protection)) {
		if (value === undefined) continue
		if (typeof value === 'boolean') {
			attrs.push(`${key}="${value ? '1' : '0'}"`)
			continue
		}
		attrs.push(`${key}="${escapeXml(String(value))}"`)
	}
	return attrs
}

function cellXml(
	ref: string,
	cell: Cell,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
): string {
	const xfIdx = xfMap.get(cell.styleId as number) ?? 0

	if (cell.formula || cell.formulaInfo?.kind === 'shared' || cell.formulaInfo?.kind === 'array') {
		return formulaCellXml(ref, cell, xfIdx)
	}

	const v = cell.value
	if (v.kind === 'number') {
		return xfIdx === 0
			? `<c r="${ref}"><v>${v.value}</v></c>`
			: `<c r="${ref}" s="${xfIdx}"><v>${v.value}</v></c>`
	}
	if (v.kind === 'date') {
		return xfIdx === 0
			? `<c r="${ref}"><v>${v.serial}</v></c>`
			: `<c r="${ref}" s="${xfIdx}"><v>${v.serial}</v></c>`
	}

	return regularCellXml(ref, cell, ssTable, xfIdx)
}

function formulaCellXml(ref: string, cell: Cell, xfIdx: number): string {
	const sAttr = xfIdx !== 0 ? ` s="${xfIdx}"` : ''
	const { typeAttr, valueStr } = formulaValueAttrs(cell.value)
	const tAttr = typeAttr ? ` t="${typeAttr}"` : ''
	const vPart = valueStr !== undefined ? `<v>${valueStr}</v>` : ''
	if (cell.formulaInfo?.kind === 'shared') {
		const sharedAttrs = [
			't="shared"',
			`si="${escapeXml(cell.formulaInfo.sharedIndex)}"`,
			...(cell.formulaInfo.isMaster && cell.formulaInfo.ref
				? [`ref="${escapeXml(cell.formulaInfo.ref)}"`]
				: []),
		]
		const formulaXml = cell.formulaInfo.isMaster
			? `<f ${sharedAttrs.join(' ')}>${escapeXml(cell.formula ?? '')}</f>`
			: `<f ${sharedAttrs.join(' ')}/>`
		return `<c r="${ref}"${sAttr}${tAttr}>${formulaXml}${vPart}</c>`
	}
	if (cell.formulaInfo?.kind === 'array') {
		const arrayAttrs = [
			't="array"',
			...(cell.formulaInfo.ref ? [`ref="${escapeXml(cell.formulaInfo.ref)}"`] : []),
		]
		return `<c r="${ref}"${sAttr}${tAttr}><f ${arrayAttrs.join(' ')}>${escapeXml(cell.formula ?? '')}</f>${vPart}</c>`
	}
	return `<c r="${ref}"${sAttr}${tAttr}><f>${escapeXml(cell.formula ?? '')}</f>${vPart}</c>`
}

function regularCellXml(
	ref: string,
	cell: Cell,
	ssTable: SharedStringTable,
	xfIdx: number,
): string {
	const sAttr = xfIdx !== 0 ? ` s="${xfIdx}"` : ''
	const { typeAttr, valueStr } = regularValueAttrs(cell.value, ssTable)
	const tAttr = typeAttr ? ` t="${typeAttr}"` : ''

	if (valueStr === undefined) return `<c r="${ref}"${sAttr}${tAttr}/>`
	return `<c r="${ref}"${sAttr}${tAttr}><v>${valueStr}</v></c>`
}

function formulaValueAttrs(value: CellValue): {
	typeAttr: string | undefined
	valueStr: string | undefined
} {
	value = topLeftScalar(value)
	switch (value.kind) {
		case 'string':
			return { typeAttr: 'str', valueStr: escapeXml(value.value) }
		case 'number':
			return { typeAttr: undefined, valueStr: String(value.value) }
		case 'boolean':
			return { typeAttr: 'b', valueStr: value.value ? '1' : '0' }
		case 'error':
			return { typeAttr: 'e', valueStr: escapeXml(value.value) }
		case 'date':
			return { typeAttr: undefined, valueStr: String(value.serial) }
		case 'empty':
			return { typeAttr: undefined, valueStr: undefined }
		case 'richText':
			return {
				typeAttr: 'str',
				valueStr: escapeXml(value.runs.map((r) => r.text).join('')),
			}
	}
}

function regularValueAttrs(
	value: CellValue,
	ssTable: SharedStringTable,
): { typeAttr: string | undefined; valueStr: string | undefined } {
	value = topLeftScalar(value)
	switch (value.kind) {
		case 'string':
		case 'richText': {
			const idx = ssTable.getIndex(value)
			return { typeAttr: 's', valueStr: idx !== undefined ? String(idx) : '0' }
		}
		case 'number':
			return { typeAttr: undefined, valueStr: String(value.value) }
		case 'boolean':
			return { typeAttr: 'b', valueStr: value.value ? '1' : '0' }
		case 'error':
			return { typeAttr: 'e', valueStr: escapeXml(value.value) }
		case 'date':
			return { typeAttr: undefined, valueStr: String(value.serial) }
		case 'empty':
			return { typeAttr: undefined, valueStr: undefined }
	}
}
