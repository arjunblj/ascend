import type { Cell, Sheet, SheetColDef } from '@ascend/core'
import { indexToColumn } from '@ascend/core'
import { type FormulaNode, parseFormula, printFormulaWithOffset } from '@ascend/formulas'
import type { CellValue, RichTextRun } from '@ascend/schema'
import { topLeftScalar } from '@ascend/schema'
import { toStoredFormulaText } from '../formula-storage.ts'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'
import { buildColorScaleXml, buildDataBarXml, buildIconSetXml } from './conditional-format.ts'
import { pushAutoFilterXml } from './filtering.ts'
import type { SharedStringTable } from './shared-strings.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MAX_SHARED_FORMULA_LENGTH = 8192

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
	readonly useInlineStrings?: boolean
	readonly usePlainStrings?: boolean
	readonly batchRows?: boolean
	readonly omitDenseCellRefs?: boolean
	/** Map of "cfIdx:ruleIdx" -> dxfId for rules with style but no dxfId */
	readonly cfDxfIdOverrides?: ReadonlyMap<string, number>
}

interface SheetXmlSink {
	push(s: string): void
}

function buildSheetXmlToSink(
	sheet: Sheet,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
	options: SheetXmlOptions,
	out: SheetXmlSink,
): void {
	const sharedFormulaExpansions =
		sheet.cells.formulaCellCount() > 0 ? buildSharedFormulaExpansions(sheet) : new Map()
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
	out.push(XML_HEADER)
	out.push(`<worksheet ${worksheetAttrs.join(' ')}>`)

	if (sheet.tabColor || sheet.outlinePr) {
		out.push('<sheetPr>')
		if (sheet.tabColor) {
			const tcAttrs: string[] = []
			if (sheet.tabColor.rgb) tcAttrs.push(`rgb="${sheet.tabColor.rgb}"`)
			if (sheet.tabColor.theme !== undefined) tcAttrs.push(`theme="${sheet.tabColor.theme}"`)
			if (sheet.tabColor.tint !== undefined) tcAttrs.push(`tint="${sheet.tabColor.tint}"`)
			if (sheet.tabColor.indexed !== undefined) tcAttrs.push(`indexed="${sheet.tabColor.indexed}"`)
			out.push(`<tabColor ${tcAttrs.join(' ')}/>`)
		}
		if (sheet.outlinePr) {
			const outlineAttrs = collectMixedAttrs(sheet.outlinePr)
			if (outlineAttrs.length > 0) out.push(`<outlinePr ${outlineAttrs.join(' ')}/>`)
		}
		out.push('</sheetPr>')
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
		if (fmtAttrs.length > 0) out.push(`<sheetFormatPr ${fmtAttrs.join(' ')}/>`)
	}

	const usedRange = sheet.cells.usedRange()
	if (usedRange) {
		const s = `${indexToColumn(usedRange.start.col)}${usedRange.start.row + 1}`
		const e = `${indexToColumn(usedRange.end.col)}${usedRange.end.row + 1}`
		out.push(`<dimension ref="${s}:${e}"/>`)
	}

	const hasFrozenPanes = sheet.frozenRows > 0 || sheet.frozenCols > 0
	if (hasFrozenPanes || sheet.sheetView) {
		const viewAttrs: string[] = ['workbookViewId="0"']
		if (sheet.sheetView) {
			if (sheet.sheetView.zoomScale !== undefined)
				viewAttrs.push(`zoomScale="${sheet.sheetView.zoomScale}"`)
			if (sheet.sheetView.zoomScaleNormal !== undefined)
				viewAttrs.push(`zoomScaleNormal="${sheet.sheetView.zoomScaleNormal}"`)
			if (sheet.sheetView.showGridLines === false) viewAttrs.push('showGridLines="0"')
			if (sheet.sheetView.showFormulas) viewAttrs.push('showFormulas="1"')
			if (sheet.sheetView.rightToLeft) viewAttrs.push('rightToLeft="1"')
			if (sheet.sheetView.tabSelected) viewAttrs.push('tabSelected="1"')
			if (sheet.sheetView.view) viewAttrs.push(`view="${sheet.sheetView.view}"`)
		}
		out.push('<sheetViews>')
		out.push(`<sheetView ${viewAttrs.join(' ')}>`)
		if (hasFrozenPanes) {
			const paneAttrs: string[] = ['state="frozen"']
			if (sheet.frozenCols > 0) paneAttrs.push(`xSplit="${sheet.frozenCols}"`)
			if (sheet.frozenRows > 0) paneAttrs.push(`ySplit="${sheet.frozenRows}"`)
			out.push(`<pane ${paneAttrs.join(' ')}/>`)
		}
		out.push('</sheetView>')
		out.push('</sheetViews>')
	}

	if (sheet.colWidths.size > 0 || sheet.colDefs.length > 0) {
		out.push('<cols>')
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
			out.push(`<col ${attrs.join(' ')}/>`)
		}
		out.push('</cols>')
	}

	out.push('<sheetData>')
	const rowHeights = [...sheet.rowHeights.entries()].sort((a, b) => a[0] - b[0])
	const rowDefs = [...sheet.rowDefs.entries()].sort((a, b) => a[0] - b[0])
	const rowIterator = sheet.cells.iterateRows()
	const columnNameCache: string[] = []
	let nextRow = rowIterator.next()
	let rowHeightIndex = 0
	let rowDefIndex = 0
	while (!nextRow.done || rowHeightIndex < rowHeights.length || rowDefIndex < rowDefs.length) {
		const populatedRow = nextRow.done ? undefined : nextRow.value
		const heightEntry = rowHeights[rowHeightIndex]
		const heightRow = heightEntry?.[0]
		const rowDefEntry = rowDefs[rowDefIndex]
		const rowDefRow = rowDefEntry?.[0]
		const row =
			populatedRow === undefined
				? heightRow === undefined
					? (rowDefRow as number)
					: rowDefRow === undefined
						? (heightRow as number)
						: Math.min(heightRow, rowDefRow)
				: heightRow === undefined
					? rowDefRow === undefined
						? populatedRow[0]
						: Math.min(populatedRow[0], rowDefRow)
					: rowDefRow === undefined
						? Math.min(populatedRow[0], heightRow)
						: Math.min(populatedRow[0], heightRow, rowDefRow)
		const cells = populatedRow && populatedRow[0] === row ? populatedRow[1] : []
		const rowAttrs = [`r="${row + 1}"`]
		const rowHeight = heightEntry && heightEntry[0] === row ? heightEntry[1] : undefined
		const rowDef = rowDefEntry && rowDefEntry[0] === row ? rowDefEntry[1] : undefined
		if (rowHeight !== undefined) {
			rowAttrs.push(`ht="${rowHeight}"`)
			rowAttrs.push('customHeight="1"')
		}
		if (rowDef?.hidden) rowAttrs.push('hidden="1"')
		if (rowDef?.collapsed) rowAttrs.push('collapsed="1"')
		if (rowDef?.outlineLevel !== undefined) rowAttrs.push(`outlineLevel="${rowDef.outlineLevel}"`)
		const rowStart = `<row ${rowAttrs.join(' ')}>`
		const rowParts = options.batchRows ? [rowStart] : undefined
		const rowOut: SheetXmlSink =
			rowParts === undefined ? out : { push: (chunk) => rowParts.push(chunk) }
		if (rowParts === undefined) out.push(rowStart)
		const rowNumber = row + 1
		const omitCellRefs =
			options.omitDenseCellRefs === true &&
			canOmitDenseCellRefs(cells, options.useInlineStrings, options.usePlainStrings)
		for (const [col, cell] of cells) {
			const ref = omitCellRefs ? undefined : `${cachedColumnName(columnNameCache, col)}${rowNumber}`
			if (
				pushDefaultStyleScalarCellXml(
					rowOut,
					ref,
					cell,
					options.useInlineStrings,
					options.usePlainStrings,
				)
			) {
				continue
			}
			const resolvedRef = ref ?? `${cachedColumnName(columnNameCache, col)}${rowNumber}`
			pushCellXml(
				rowOut,
				resolvedRef,
				cell,
				ssTable,
				xfMap,
				sharedFormulaExpansions,
				options.useInlineStrings,
				options.usePlainStrings,
			)
		}
		if (rowParts === undefined) {
			out.push('</row>')
		} else {
			rowParts.push('</row>')
			out.push(rowParts.join(''))
		}
		if (populatedRow && populatedRow[0] === row) nextRow = rowIterator.next()
		if (heightEntry && heightEntry[0] === row) rowHeightIndex++
		if (rowDefEntry && rowDefEntry[0] === row) rowDefIndex++
	}

	out.push('</sheetData>')

	if (sheet.merges.length > 0) {
		out.push(`<mergeCells count="${sheet.merges.length}">`)
		for (const merge of sheet.merges) {
			const s = `${indexToColumn(merge.start.col)}${merge.start.row + 1}`
			const e = `${indexToColumn(merge.end.col)}${merge.end.row + 1}`
			out.push(`<mergeCell ref="${s}:${e}"/>`)
		}
		out.push('</mergeCells>')
	}

	if (sheet.protection) {
		const attrs = collectProtectionAttrs(sheet.protection)
		if (attrs.length > 0) out.push(`<sheetProtection ${attrs.join(' ')}/>`)
	}

	if (sheet.autoFilter) {
		pushAutoFilterXml(out, sheet.autoFilter)
	}

	if (sheet.conditionalFormats.length > 0) {
		const cfDxfIdOverrides = options.cfDxfIdOverrides
		for (let cfIdx = 0; cfIdx < sheet.conditionalFormats.length; cfIdx++) {
			const conditionalFormat = sheet.conditionalFormats[cfIdx]
			if (!conditionalFormat) continue
			out.push(`<conditionalFormatting sqref="${escapeXml(conditionalFormat.sqref)}">`)
			for (let ruleIdx = 0; ruleIdx < conditionalFormat.rules.length; ruleIdx++) {
				const rule = conditionalFormat.rules[ruleIdx]
				if (!rule) continue
				const effectiveDxfId = rule.dxfId ?? cfDxfIdOverrides?.get(`${cfIdx}:${ruleIdx}`)
				const attrs = [`type="${escapeXml(rule.type)}"`]
				if (rule.operator) attrs.push(`operator="${escapeXml(rule.operator)}"`)
				if (rule.priority !== undefined) attrs.push(`priority="${rule.priority}"`)
				if (effectiveDxfId !== undefined) attrs.push(`dxfId="${effectiveDxfId}"`)
				if (rule.stopIfTrue) attrs.push('stopIfTrue="1"')
				if (rule.rank !== undefined) attrs.push(`rank="${rule.rank}"`)
				if (rule.percent !== undefined) attrs.push(`percent="${rule.percent ? '1' : '0'}"`)
				if (rule.bottom !== undefined) attrs.push(`bottom="${rule.bottom ? '1' : '0'}"`)
				if (rule.aboveAverage !== undefined)
					attrs.push(`aboveAverage="${rule.aboveAverage ? '1' : '0'}"`)
				if (rule.equalAverage !== undefined)
					attrs.push(`equalAverage="${rule.equalAverage ? '1' : '0'}"`)
				if (rule.timePeriod) attrs.push(`timePeriod="${escapeXml(rule.timePeriod)}"`)

				out.push(`<cfRule ${attrs.join(' ')}>`)
				for (const formula of rule.formulas) {
					out.push(`<formula>${escapeXml(formula)}</formula>`)
				}
				if (rule.colorScale) out.push(buildColorScaleXml(rule.colorScale))
				if (rule.dataBar) out.push(buildDataBarXml(rule.dataBar))
				if (rule.iconSet) out.push(buildIconSetXml(rule.iconSet))
				out.push('</cfRule>')
			}
			out.push('</conditionalFormatting>')
		}
	}

	const legacyDataValidations = sheet.dataValidations.filter(
		(validation) => validation.source !== 'x14' || !sheet.preservedExtLst,
	)
	if (legacyDataValidations.length > 0) {
		out.push(`<dataValidations count="${legacyDataValidations.length}">`)
		for (const validation of legacyDataValidations) {
			const attrs = [`sqref="${escapeXml(validation.sqref)}"`]
			if (validation.type) attrs.push(`type="${escapeXml(validation.type)}"`)
			if (validation.operator) attrs.push(`operator="${escapeXml(validation.operator)}"`)
			if (validation.errorStyle) attrs.push(`errorStyle="${escapeXml(validation.errorStyle)}"`)
			if (validation.imeMode) attrs.push(`imeMode="${escapeXml(validation.imeMode)}"`)
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
			out.push(`<dataValidation ${attrs.join(' ')}>`)
			if (validation.formula1) out.push(`<formula1>${escapeXml(validation.formula1)}</formula1>`)
			if (validation.formula2) out.push(`<formula2>${escapeXml(validation.formula2)}</formula2>`)
			out.push('</dataValidation>')
		}
		out.push('</dataValidations>')
	}

	if (hyperlinks.length > 0) {
		out.push('<hyperlinks>')
		for (const hyperlink of hyperlinks) {
			const attrs = [`ref="${escapeXml(hyperlink.ref)}"`]
			if (hyperlink.relId) attrs.push(`r:id="${hyperlink.relId}"`)
			if (hyperlink.location) attrs.push(`location="${escapeXml(hyperlink.location)}"`)
			if (hyperlink.display) attrs.push(`display="${escapeXml(hyperlink.display)}"`)
			if (hyperlink.tooltip) attrs.push(`tooltip="${escapeXml(hyperlink.tooltip)}"`)
			out.push(`<hyperlink ${attrs.join(' ')}/>`)
		}
		out.push('</hyperlinks>')
	}

	if (sheet.printOptions) {
		const attrs = collectMixedAttrs(sheet.printOptions)
		if (attrs.length > 0) out.push(`<printOptions ${attrs.join(' ')}/>`)
	}

	if (sheet.pageMargins) {
		const attrs = collectNumericAttrs(sheet.pageMargins)
		if (attrs.length > 0) out.push(`<pageMargins ${attrs.join(' ')}/>`)
	}

	if (sheet.pageSetup) {
		const attrs = collectMixedAttrs(sheet.pageSetup)
		if (attrs.length > 0) out.push(`<pageSetup ${attrs.join(' ')}/>`)
	}

	if (sheet.headerFooter) {
		out.push('<headerFooter>')
		if (sheet.headerFooter.oddHeader) {
			out.push(`<oddHeader>${escapeXml(sheet.headerFooter.oddHeader)}</oddHeader>`)
		}
		if (sheet.headerFooter.oddFooter) {
			out.push(`<oddFooter>${escapeXml(sheet.headerFooter.oddFooter)}</oddFooter>`)
		}
		if (sheet.headerFooter.evenHeader) {
			out.push(`<evenHeader>${escapeXml(sheet.headerFooter.evenHeader)}</evenHeader>`)
		}
		if (sheet.headerFooter.evenFooter) {
			out.push(`<evenFooter>${escapeXml(sheet.headerFooter.evenFooter)}</evenFooter>`)
		}
		if (sheet.headerFooter.firstHeader) {
			out.push(`<firstHeader>${escapeXml(sheet.headerFooter.firstHeader)}</firstHeader>`)
		}
		if (sheet.headerFooter.firstFooter) {
			out.push(`<firstFooter>${escapeXml(sheet.headerFooter.firstFooter)}</firstFooter>`)
		}
		out.push('</headerFooter>')
	}

	appendBreaks(out, 'rowBreaks', sheet.rowBreaks)
	appendBreaks(out, 'colBreaks', sheet.colBreaks)

	if (sheet.ignoredErrors.length > 0) {
		out.push('<ignoredErrors>')
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
			out.push(`<ignoredError ${attrs.join(' ')}/>`)
		}
		out.push('</ignoredErrors>')
	}

	if (drawingRelId) {
		out.push(`<drawing r:id="${drawingRelId}"/>`)
	}

	if (legacyDrawingRelId) {
		out.push(`<legacyDrawing r:id="${legacyDrawingRelId}"/>`)
	}

	if (tableRelIds.length > 0) {
		out.push(`<tableParts count="${tableRelIds.length}">`)
		for (const relId of tableRelIds) {
			out.push(`<tablePart r:id="${relId}"/>`)
		}
		out.push('</tableParts>')
	}

	if (sheet.preservedExtLst) {
		out.push(sheet.preservedExtLst)
	}

	out.push('</worksheet>')
}

function cachedColumnName(cache: string[], col: number): string {
	let name = cache[col]
	if (name === undefined) {
		name = indexToColumn(col)
		cache[col] = name
	}
	return name
}

function canOmitDenseCellRefs(
	cells: readonly (readonly [number, Cell])[],
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): boolean {
	if (cells.length === 0) return false
	for (let index = 0; index < cells.length; index++) {
		const entry = cells[index]
		if (!entry || entry[0] !== index) return false
		const cell = entry[1]
		if (
			(cell.styleId as number) !== 0 ||
			cell.formula ||
			cell.formulaInfo ||
			!canOmitDefaultStyleScalarCellRef(cell, useInlineStrings, usePlainStrings)
		) {
			return false
		}
	}
	return true
}

function canOmitDefaultStyleScalarCellRef(
	cell: Cell,
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): boolean {
	const value = cell.value
	if (
		value.kind === 'number' ||
		value.kind === 'date' ||
		value.kind === 'boolean' ||
		value.kind === 'error' ||
		value.kind === 'empty'
	) {
		return true
	}
	if (value.kind === 'string') return useInlineStrings === true || usePlainStrings === true
	if (value.kind === 'richText') return useInlineStrings === true || usePlainStrings === true
	return false
}

function appendBreaks(
	out: SheetXmlSink,
	tagName: 'rowBreaks' | 'colBreaks',
	breaks: readonly { id: number; min?: number; max?: number; man?: boolean; pt?: boolean }[],
): void {
	if (breaks.length === 0) return
	const manualBreakCount = breaks.reduce((count, brk) => count + (brk.man ? 1 : 0), 0)
	out.push(`<${tagName} count="${breaks.length}" manualBreakCount="${manualBreakCount}">`)
	for (const brk of breaks) {
		const attrs = [`id="${brk.id}"`]
		if (brk.min !== undefined) attrs.push(`min="${brk.min}"`)
		if (brk.max !== undefined) attrs.push(`max="${brk.max}"`)
		if (brk.man !== undefined) attrs.push(`man="${brk.man ? '1' : '0'}"`)
		if (brk.pt !== undefined) attrs.push(`pt="${brk.pt ? '1' : '0'}"`)
		out.push(`<brk ${attrs.join(' ')}/>`)
	}
	out.push(`</${tagName}>`)
}

export function buildSheetXml(
	sheet: Sheet,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
	options: SheetXmlOptions = {},
): string {
	const out = new ChunkedStringBuilder()
	buildSheetXmlToSink(sheet, ssTable, xfMap, options, out)
	return out.toString()
}

export function buildSheetXmlStreaming(
	sheet: Sheet,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
	options: SheetXmlOptions,
	onChunk: (chunk: string) => void,
): void {
	buildSheetXmlToSink(sheet, ssTable, xfMap, options, { push: onChunk })
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

function pushDefaultStyleScalarCellXml(
	out: SheetXmlSink,
	ref: string | undefined,
	cell: Cell,
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): boolean {
	if (
		(cell.styleId as number) !== 0 ||
		cell.formula ||
		cell.formulaInfo?.kind === 'shared' ||
		cell.formulaInfo?.kind === 'array'
	) {
		return false
	}
	const v = cell.value
	const r = ref === undefined ? '' : ` r="${ref}"`
	if (v.kind === 'number') {
		out.push(`<c${r}><v>${v.value}</v></c>`)
		return true
	}
	if (v.kind === 'date') {
		out.push(`<c${r}><v>${v.serial}</v></c>`)
		return true
	}
	if (v.kind === 'boolean') {
		out.push(`<c${r} t="b"><v>${v.value ? '1' : '0'}</v></c>`)
		return true
	}
	if (v.kind === 'error') {
		out.push(`<c${r} t="e"><v>${escapeXml(v.value)}</v></c>`)
		return true
	}
	if (v.kind === 'empty') {
		out.push(`<c${r}/>`)
		return true
	}
	if (v.kind === 'string') {
		if (usePlainStrings) {
			out.push(`<c${r} t="str"><v>${escapeXml(v.value)}</v></c>`)
			return true
		}
		if (useInlineStrings) {
			out.push(`<c${r} t="inlineStr"><is><t>${escapeXml(v.value)}</t></is></c>`)
			return true
		}
	}
	if ((usePlainStrings || useInlineStrings) && v.kind === 'richText') {
		const runsXml = v.runs.map((r) => inlineStrRunXml(r)).join('')
		out.push(`<c${r} t="inlineStr"><is>${runsXml}</is></c>`)
		return true
	}
	return false
}

function pushCellXml(
	out: SheetXmlSink,
	ref: string,
	cell: Cell,
	ssTable: SharedStringTable,
	xfMap: Map<number, number>,
	sharedFormulaExpansions: ReadonlyMap<string, SharedFormulaExpansion>,
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): void {
	const styleId = cell.styleId as number
	const xfIdx = styleId === 0 ? 0 : (xfMap.get(styleId) ?? 0)

	if (cell.formula || cell.formulaInfo?.kind === 'shared' || cell.formulaInfo?.kind === 'array') {
		out.push(formulaCellXml(ref, cell, xfIdx, sharedFormulaExpansions))
		return
	}

	const v = cell.value
	if (v.kind === 'number') {
		out.push(
			xfIdx === 0
				? `<c r="${ref}"><v>${v.value}</v></c>`
				: `<c r="${ref}" s="${xfIdx}"><v>${v.value}</v></c>`,
		)
		return
	}
	if (v.kind === 'date') {
		out.push(
			xfIdx === 0
				? `<c r="${ref}"><v>${v.serial}</v></c>`
				: `<c r="${ref}" s="${xfIdx}"><v>${v.serial}</v></c>`,
		)
		return
	}
	if (v.kind === 'string') {
		if (usePlainStrings) {
			out.push(
				xfIdx === 0
					? `<c r="${ref}" t="str"><v>${escapeXml(v.value)}</v></c>`
					: `<c r="${ref}" s="${xfIdx}" t="str"><v>${escapeXml(v.value)}</v></c>`,
			)
			return
		}
		if (useInlineStrings) {
			out.push(
				xfIdx === 0
					? `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(v.value)}</t></is></c>`
					: `<c r="${ref}" s="${xfIdx}" t="inlineStr"><is><t>${escapeXml(v.value)}</t></is></c>`,
			)
			return
		}
	}
	if ((usePlainStrings || useInlineStrings) && v.kind === 'richText') {
		const runsXml = v.runs.map((r) => inlineStrRunXml(r)).join('')
		out.push(
			xfIdx === 0
				? `<c r="${ref}" t="inlineStr"><is>${runsXml}</is></c>`
				: `<c r="${ref}" s="${xfIdx}" t="inlineStr"><is>${runsXml}</is></c>`,
		)
		return
	}

	out.push(regularCellXml(ref, cell, ssTable, xfIdx, useInlineStrings, usePlainStrings))
}

function formulaCellXml(
	ref: string,
	cell: Cell,
	xfIdx: number,
	sharedFormulaExpansions: ReadonlyMap<string, SharedFormulaExpansion>,
): string {
	const formulaText = cell.formula ? toStoredFormulaText(cell.formula) : ''
	const sAttr = xfIdx !== 0 ? ` s="${xfIdx}"` : ''
	const dynamicArrayMetadataIndex = dynamicArrayCellMetadataIndex(cell.formulaInfo)
	const cmAttr = dynamicArrayMetadataIndex !== undefined ? ` cm="${dynamicArrayMetadataIndex}"` : ''
	const { typeAttr, valueStr } = formulaValueAttrs(cell.value)
	const tAttr = typeAttr ? ` t="${typeAttr}"` : ''
	const vPart = valueStr !== undefined ? `<v>${valueStr}</v>` : ''
	if (cell.formulaInfo?.kind === 'shared') {
		const expanded = sharedFormulaExpansions.get(ref)
		if (expanded) {
			return `<c r="${ref}"${cmAttr}${sAttr}${tAttr}><f>${escapeXml(expanded.formulaText)}</f>${vPart}</c>`
		}
		let fAttrs = `t="shared" si="${escapeXml(cell.formulaInfo.sharedIndex)}"`
		if (cell.formulaInfo.isMaster && cell.formulaInfo.ref) {
			fAttrs += ` ref="${escapeXml(cell.formulaInfo.ref)}"`
		}
		const formulaXml = cell.formulaInfo.isMaster
			? `<f ${fAttrs}>${escapeXml(formulaText)}</f>`
			: `<f ${fAttrs}/>`
		return `<c r="${ref}"${cmAttr}${sAttr}${tAttr}>${formulaXml}${vPart}</c>`
	}
	if (cell.formulaInfo?.kind === 'array') {
		let fAttrs = 't="array"'
		if (cell.formulaInfo.ref) fAttrs += ` ref="${escapeXml(cell.formulaInfo.ref)}"`
		return `<c r="${ref}"${cmAttr}${sAttr}${tAttr}><f ${fAttrs}>${escapeXml(formulaText)}</f>${vPart}</c>`
	}
	return `<c r="${ref}"${cmAttr}${sAttr}${tAttr}><f>${escapeXml(formulaText)}</f>${vPart}</c>`
}

function dynamicArrayCellMetadataIndex(
	binding: Cell['formulaInfo'] | undefined,
): number | undefined {
	if (binding?.kind === 'dynamicArray') return binding.metadataIndex
	if (binding?.kind === 'spill' && binding.isAnchor) return 1
	return undefined
}

interface SharedFormulaExpansion {
	readonly formulaText: string
}

interface SharedFormulaMaster {
	readonly ast: FormulaNode
	readonly row: number
	readonly col: number
	readonly formulaText: string
}

function buildSharedFormulaExpansions(sheet: Sheet): ReadonlyMap<string, SharedFormulaExpansion> {
	const masters = new Map<string, SharedFormulaMaster>()
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (cell.formulaInfo?.kind !== 'shared' || !cell.formulaInfo.isMaster || !cell.formula) continue
		const formulaText = toStoredFormulaText(cell.formula)
		if (formulaText.length <= MAX_SHARED_FORMULA_LENGTH) continue
		const parsed = parseFormula(formulaText)
		if (!parsed.ok) continue
		const masterRef = cell.formulaInfo.masterRef ?? `${indexToColumn(col)}${row + 1}`
		masters.set(masterRef, {
			ast: parsed.value,
			row,
			col,
			formulaText,
		})
	}

	if (masters.size === 0) return new Map()

	const expansions = new Map<string, SharedFormulaExpansion>()
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (cell.formulaInfo?.kind !== 'shared') continue
		const ref = `${indexToColumn(col)}${row + 1}`
		const masterRef = cell.formulaInfo.masterRef ?? ref
		const master = masters.get(masterRef)
		if (!master) continue
		if (cell.formulaInfo.isMaster) {
			expansions.set(ref, { formulaText: master.formulaText })
			continue
		}
		const rowDelta = row - master.row
		const colDelta = col - master.col
		expansions.set(ref, {
			formulaText: printFormulaWithOffset(master.ast, rowDelta, colDelta),
		})
	}
	return expansions
}

function regularCellXml(
	ref: string,
	cell: Cell,
	ssTable: SharedStringTable,
	xfIdx: number,
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): string {
	const sAttr = xfIdx !== 0 ? ` s="${xfIdx}"` : ''
	if (usePlainStrings) {
		const value = cell.value
		const v = value.kind === 'array' ? topLeftScalar(value) : value
		if (v.kind === 'string') {
			return `<c r="${ref}"${sAttr} t="str"><v>${escapeXml(v.value)}</v></c>`
		}
		if (v.kind === 'richText') {
			const runsXml = v.runs.map((r) => inlineStrRunXml(r)).join('')
			return `<c r="${ref}"${sAttr} t="inlineStr"><is>${runsXml}</is></c>`
		}
	}
	if (useInlineStrings) {
		const value = cell.value
		const v = value.kind === 'array' ? topLeftScalar(value) : value
		if (v.kind === 'string') {
			return `<c r="${ref}"${sAttr} t="inlineStr"><is><t>${escapeXml(v.value)}</t></is></c>`
		}
		if (v.kind === 'richText') {
			const runsXml = v.runs.map((r) => inlineStrRunXml(r)).join('')
			return `<c r="${ref}"${sAttr} t="inlineStr"><is>${runsXml}</is></c>`
		}
	}
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

function inlineStrRunXml(run: RichTextRun): string {
	const parts: string[] = []
	if (run.bold) parts.push('<b/>')
	if (run.italic) parts.push('<i/>')
	if (run.underline) parts.push('<u/>')
	if (run.strikethrough) parts.push('<strike/>')
	if (run.fontSize !== undefined) parts.push(`<sz val="${run.fontSize}"/>`)
	if (run.color) {
		if (typeof run.color === 'string') {
			parts.push(`<color rgb="${escapeXml(run.color)}"/>`)
		} else if (run.color.kind === 'rgb') {
			parts.push(`<color rgb="${escapeXml(run.color.rgb)}"/>`)
		} else if (run.color.kind === 'theme') {
			parts.push(
				run.color.tint !== undefined
					? `<color theme="${run.color.theme}" tint="${run.color.tint}"/>`
					: `<color theme="${run.color.theme}"/>`,
			)
		} else if (run.color.kind === 'indexed') {
			parts.push(`<color indexed="${run.color.index}"/>`)
		}
	}
	if (run.fontName) parts.push(`<rFont val="${escapeXml(run.fontName)}"/>`)
	const rPrEl = parts.length > 0 ? `<rPr>${parts.join('')}</rPr>` : ''
	return `<r>${rPrEl}<t>${escapeXml(run.text)}</t></r>`
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
