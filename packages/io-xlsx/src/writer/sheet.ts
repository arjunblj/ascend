import type {
	Cell,
	RangeRef,
	Sheet,
	SheetCellMetadataAttrs,
	SheetColDef,
	SheetConditionalFormatRule,
} from '@ascend/core'
import { indexToColumn, parseRange } from '@ascend/core'
import { type FormulaNode, parseFormula, printFormulaWithOffset } from '@ascend/formulas'
import type { CellValue, RichTextRun } from '@ascend/schema'
import { topLeftScalar } from '@ascend/schema'
import { normalizeStoredFormulaText, toStoredFormulaText } from '../formula-storage.ts'
import { escapeXml } from '../xml.ts'
import { buildCustomSheetViewsXml, updateCustomSheetViewsXml } from './advanced-filter.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'
import { buildColorScaleXml, buildDataBarXml, buildIconSetXml } from './conditional-format.ts'
import { pushAutoFilterXml, pushSortStateXml } from './filtering.ts'
import type { SharedStringTable } from './shared-strings.ts'
import { buildWorksheetExtLstXml, updateWorksheetExtLstXml } from './sparkline.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const NS_X14AC = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main'
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
		legacyDrawingRelId ||
		sheet.pageSetup?.printerSettingsRelId ||
		sheet.preservedControlsXml
	) {
		worksheetAttrs.push(`xmlns:r="${NS_R}"`)
	}
	if (
		sheet.dataValidations.some((validation) => validation.uid) ||
		sheet.autoFilter?.uid ||
		sheetHasPreservedOfficeExtensionXml(sheet, 'xr:')
	) {
		worksheetAttrs.push(`xmlns:xr="${NS_XR}"`)
	}
	if (sheetHasPreservedOfficeExtensionXml(sheet, 'x14:')) {
		worksheetAttrs.push(`xmlns:x14="${NS_X14}"`)
	}
	if (sheetHasPreservedOfficeExtensionXml(sheet, 'xm:')) {
		worksheetAttrs.push(`xmlns:xm="${NS_XM}"`)
	}
	if (
		sheet.sheetFormatPr?.dyDescent !== undefined ||
		[...sheet.rowDefs.values()].some((rowDef) => rowDef.dyDescent !== undefined) ||
		sheetHasPreservedOfficeExtensionXml(sheet, 'x14ac:')
	) {
		worksheetAttrs.push(`xmlns:x14ac="${NS_X14AC}"`)
	}
	out.push(XML_HEADER)
	out.push(`<worksheet ${worksheetAttrs.join(' ')}>`)

	if (
		sheet.codeName ||
		sheet.filterMode !== null ||
		sheet.enableFormatConditionsCalculation !== null ||
		sheet.tabColor ||
		sheet.outlinePr ||
		sheet.pageSetupPr
	) {
		const sheetPrAttrs: string[] = []
		if (sheet.codeName) sheetPrAttrs.push(`codeName="${escapeXml(sheet.codeName)}"`)
		if (sheet.filterMode !== null) sheetPrAttrs.push(`filterMode="${sheet.filterMode ? '1' : '0'}"`)
		if (sheet.enableFormatConditionsCalculation !== null) {
			sheetPrAttrs.push(
				`enableFormatConditionsCalculation="${sheet.enableFormatConditionsCalculation ? '1' : '0'}"`,
			)
		}
		out.push(`<sheetPr${sheetPrAttrs.length > 0 ? ` ${sheetPrAttrs.join(' ')}` : ''}>`)
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
		if (sheet.pageSetupPr) {
			const pageSetupPrAttrs = collectMixedAttrs(sheet.pageSetupPr)
			if (pageSetupPrAttrs.length > 0) out.push(`<pageSetUpPr ${pageSetupPrAttrs.join(' ')}/>`)
		}
		out.push('</sheetPr>')
	}

	if (sheet.sheetFormatPr) {
		const fmtAttrs: string[] = []
		if (sheet.sheetFormatPr.baseColWidth !== undefined)
			fmtAttrs.push(`baseColWidth="${sheet.sheetFormatPr.baseColWidth}"`)
		if (sheet.sheetFormatPr.defaultRowHeight !== undefined)
			fmtAttrs.push(`defaultRowHeight="${sheet.sheetFormatPr.defaultRowHeight}"`)
		if (sheet.sheetFormatPr.defaultColWidth !== undefined)
			fmtAttrs.push(`defaultColWidth="${sheet.sheetFormatPr.defaultColWidth}"`)
		if (sheet.sheetFormatPr.outlineLevelRow !== undefined)
			fmtAttrs.push(`outlineLevelRow="${sheet.sheetFormatPr.outlineLevelRow}"`)
		if (sheet.sheetFormatPr.outlineLevelCol !== undefined)
			fmtAttrs.push(`outlineLevelCol="${sheet.sheetFormatPr.outlineLevelCol}"`)
		if (sheet.sheetFormatPr.customHeight) fmtAttrs.push('customHeight="1"')
		if (sheet.sheetFormatPr.zeroHeight !== undefined)
			fmtAttrs.push(`zeroHeight="${sheet.sheetFormatPr.zeroHeight ? '1' : '0'}"`)
		if (sheet.sheetFormatPr.dyDescent !== undefined)
			fmtAttrs.push(`x14ac:dyDescent="${sheet.sheetFormatPr.dyDescent}"`)
		if (fmtAttrs.length > 0) out.push(`<sheetFormatPr ${fmtAttrs.join(' ')}/>`)
	}

	const usedRange = combinedDimensionRange(sheet)
	if (usedRange) {
		const s = `${indexToColumn(usedRange.start.col)}${usedRange.start.row + 1}`
		const e = `${indexToColumn(usedRange.end.col)}${usedRange.end.row + 1}`
		out.push(`<dimension ref="${s}:${e}"/>`)
	}

	const hasFrozenPanes = sheet.frozenRows > 0 || sheet.frozenCols > 0
	const preservedPaneAttributes = sheet.preservedPaneAttributes ?? {}
	const hasPreservedPaneAttributes = Object.keys(preservedPaneAttributes).length > 0
	const preservedSheetViewAttributes = sheet.preservedSheetViewAttributes ?? {}
	const hasPreservedSheetViewAttributes = Object.keys(preservedSheetViewAttributes).length > 0
	const preservedSheetViewSelections = sheet.preservedSheetViewSelections ?? []
	const hasPreservedSheetViewSelections = preservedSheetViewSelections.length > 0
	if (
		hasFrozenPanes ||
		hasPreservedPaneAttributes ||
		hasPreservedSheetViewSelections ||
		sheet.sheetView ||
		hasPreservedSheetViewAttributes
	) {
		const viewAttrs = new Map<string, string>([['workbookViewId', '0']])
		for (const [name, value] of Object.entries(preservedSheetViewAttributes)) {
			viewAttrs.set(name, value)
		}
		if (sheet.sheetView) {
			if (sheet.sheetView.zoomScale !== undefined)
				setSheetViewAttr(viewAttrs, 'zoomScale', String(sheet.sheetView.zoomScale))
			if (sheet.sheetView.zoomScaleNormal !== undefined)
				setSheetViewAttr(viewAttrs, 'zoomScaleNormal', String(sheet.sheetView.zoomScaleNormal))
			if (sheet.sheetView.zoomScaleSheetLayoutView !== undefined)
				setSheetViewAttr(
					viewAttrs,
					'zoomScaleSheetLayoutView',
					String(sheet.sheetView.zoomScaleSheetLayoutView),
				)
			if (sheet.sheetView.showGridLines === false) setSheetViewAttr(viewAttrs, 'showGridLines', '0')
			if (sheet.sheetView.showFormulas) setSheetViewAttr(viewAttrs, 'showFormulas', '1')
			if (sheet.sheetView.rightToLeft) setSheetViewAttr(viewAttrs, 'rightToLeft', '1')
			if (sheet.sheetView.tabSelected) setSheetViewAttr(viewAttrs, 'tabSelected', '1')
			if (sheet.sheetView.view) setSheetViewAttr(viewAttrs, 'view', sheet.sheetView.view)
			if (sheet.sheetView.topLeftCell)
				setSheetViewAttr(viewAttrs, 'topLeftCell', sheet.sheetView.topLeftCell)
		}
		out.push('<sheetViews>')
		out.push(`<sheetView ${sheetViewAttrsXml(viewAttrs)}>`)
		if (hasFrozenPanes || hasPreservedPaneAttributes) {
			const paneAttrs = new Map<string, string>(Object.entries(preservedPaneAttributes))
			if (hasFrozenPanes) {
				paneAttrs.set('state', paneAttrs.get('state') === 'frozenSplit' ? 'frozenSplit' : 'frozen')
				if (sheet.frozenCols > 0) paneAttrs.set('xSplit', String(sheet.frozenCols))
				else paneAttrs.delete('xSplit')
				if (sheet.frozenRows > 0) paneAttrs.set('ySplit', String(sheet.frozenRows))
				else paneAttrs.delete('ySplit')
			}
			out.push(`<pane ${sheetViewAttrsXml(paneAttrs)}/>`)
		}
		for (const selection of preservedSheetViewSelections) {
			out.push(`<selection ${sheetViewAttrsXml(new Map(Object.entries(selection)))}/>`)
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
	const blankRows = [...sheet.preservedBlankCells.entries()]
		.filter(([, cells]) => cells.size > 0)
		.sort((a, b) => a[0] - b[0])
	const rowIterator = sheet.cells.iterateRows()
	const columnNameCache: string[] = []
	let nextRow = rowIterator.next()
	let rowHeightIndex = 0
	let rowDefIndex = 0
	let blankRowIndex = 0
	while (
		!nextRow.done ||
		rowHeightIndex < rowHeights.length ||
		rowDefIndex < rowDefs.length ||
		blankRowIndex < blankRows.length
	) {
		const populatedRow = nextRow.done ? undefined : nextRow.value
		const heightEntry = rowHeights[rowHeightIndex]
		const heightRow = heightEntry?.[0]
		const rowDefEntry = rowDefs[rowDefIndex]
		const rowDefRow = rowDefEntry?.[0]
		const blankRowEntry = blankRows[blankRowIndex]
		const blankRow = blankRowEntry?.[0]
		const row = Math.min(
			populatedRow?.[0] ?? Number.POSITIVE_INFINITY,
			heightRow ?? Number.POSITIVE_INFINITY,
			rowDefRow ?? Number.POSITIVE_INFINITY,
			blankRow ?? Number.POSITIVE_INFINITY,
		)
		const cells = populatedRow && populatedRow[0] === row ? populatedRow[1] : []
		const blankCells =
			blankRowEntry && blankRowEntry[0] === row
				? [...blankRowEntry[1].entries()].sort((a, b) => a[0] - b[0])
				: []
		const rowAttrs = [`r="${row + 1}"`]
		const rowHeight = heightEntry && heightEntry[0] === row ? heightEntry[1] : undefined
		const rowDef = rowDefEntry && rowDefEntry[0] === row ? rowDefEntry[1] : undefined
		if (rowDef?.spans !== undefined) rowAttrs.push(`spans="${escapeXml(rowDef.spans)}"`)
		if (rowHeight !== undefined) {
			rowAttrs.push(`ht="${rowHeight}"`)
			rowAttrs.push(`customHeight="${rowDef?.customHeight === false ? '0' : '1'}"`)
		} else if (rowDef?.customHeight !== undefined) {
			rowAttrs.push(`customHeight="${rowDef.customHeight ? '1' : '0'}"`)
		}
		if (rowDef?.style !== undefined) rowAttrs.push(`s="${rowDef.style}"`)
		if (rowDef?.customFormat !== undefined)
			rowAttrs.push(`customFormat="${rowDef.customFormat ? '1' : '0'}"`)
		if (rowDef?.hidden) rowAttrs.push('hidden="1"')
		if (rowDef?.collapsed) rowAttrs.push('collapsed="1"')
		if (rowDef?.outlineLevel !== undefined) rowAttrs.push(`outlineLevel="${rowDef.outlineLevel}"`)
		if (rowDef?.thickTop !== undefined) rowAttrs.push(`thickTop="${rowDef.thickTop ? '1' : '0'}"`)
		if (rowDef?.thickBot !== undefined) rowAttrs.push(`thickBot="${rowDef.thickBot ? '1' : '0'}"`)
		if (rowDef?.dyDescent !== undefined) rowAttrs.push(`x14ac:dyDescent="${rowDef.dyDescent}"`)
		const rowStart = `<row ${rowAttrs.join(' ')}>`
		const denseCellsWithoutRefsXml =
			options.batchRows && options.omitDenseCellRefs === true && blankCells.length === 0
				? defaultStyleScalarCellsWithoutRefsXml(
						cells,
						ssTable,
						options.useInlineStrings,
						options.usePlainStrings,
					)
				: false
		if (denseCellsWithoutRefsXml !== false) out.push(`${rowStart}${denseCellsWithoutRefsXml}</row>`)
		const rowParts =
			options.batchRows && denseCellsWithoutRefsXml === false ? [rowStart] : undefined
		const rowOut: SheetXmlSink =
			rowParts === undefined ? out : { push: (chunk) => rowParts.push(chunk) }
		if (rowParts === undefined && denseCellsWithoutRefsXml === false) out.push(rowStart)
		const rowNumber = row + 1
		const omitCellRefs =
			denseCellsWithoutRefsXml === false &&
			options.omitDenseCellRefs === true &&
			blankCells.length === 0 &&
			canOmitDenseCellRefs(cells)
		let cellIndex = 0
		let blankIndex = 0
		while (
			denseCellsWithoutRefsXml === false &&
			(cellIndex < cells.length || blankIndex < blankCells.length)
		) {
			const cellEntry = cells[cellIndex]
			const blankEntry = blankCells[blankIndex]
			if (blankEntry && (!cellEntry || blankEntry[0] < cellEntry[0])) {
				const col = blankEntry[0]
				const ref = `${cachedColumnName(columnNameCache, col)}${rowNumber}`
				rowOut.push(`<c ${blankCellAttrsXml(blankEntry[1], ref)}/>`)
				blankIndex++
				continue
			}
			if (!cellEntry) break
			if (blankEntry && blankEntry[0] === cellEntry[0]) blankIndex++
			const [col, cell] = cellEntry
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
				cellIndex++
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
				sheet.storedFormulaText.get(formulaStorageKey(row, col)),
				sheet.preservedCellMetadata.get(formulaStorageKey(row, col)),
				options.useInlineStrings,
				options.usePlainStrings,
			)
			cellIndex++
		}
		if (denseCellsWithoutRefsXml === false) {
			if (rowParts === undefined) {
				out.push('</row>')
			} else {
				rowParts.push('</row>')
				out.push(rowParts.join(''))
			}
		}
		if (populatedRow && populatedRow[0] === row) nextRow = rowIterator.next()
		if (heightEntry && heightEntry[0] === row) rowHeightIndex++
		if (rowDefEntry && rowDefEntry[0] === row) rowDefIndex++
		if (blankRowEntry && blankRowEntry[0] === row) blankRowIndex++
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

	if (sheet.protectedRanges.length > 0) {
		out.push(`<protectedRanges count="${sheet.protectedRanges.length}">`)
		for (const range of sheet.protectedRanges) {
			out.push(`<protectedRange ${collectProtectedRangeAttrs(range).join(' ')}/>`)
		}
		out.push('</protectedRanges>')
	}

	if (sheet.phoneticPr) {
		const attrs = collectMixedAttrs(sheet.phoneticPr)
		if (attrs.length > 0) out.push(`<phoneticPr ${attrs.join(' ')}/>`)
	}

	if (sheet.autoFilter) {
		pushAutoFilterXml(out, sheet.autoFilter, {
			includeUid: true,
			...(sheet.preservedAutoFilterSortStateAttributes
				? { sortStateAttributes: sheet.preservedAutoFilterSortStateAttributes }
				: {}),
		})
	}

	if (sheet.sortState) {
		pushSortStateXml(out, sheet.sortState, {
			...(sheet.preservedSortStateAttributes
				? { sortStateAttributes: sheet.preservedSortStateAttributes }
				: {}),
		})
	}

	if (sheet.preservedCustomSheetViews) {
		out.push(updateCustomSheetViewsXml(sheet.preservedCustomSheetViews, sheet.advancedFilters))
	} else if (sheet.advancedFilters.length > 0) {
		out.push(buildCustomSheetViewsXml(sheet.advancedFilters))
	}

	if (sheet.conditionalFormats.length > 0) {
		const cfDxfIdOverrides = options.cfDxfIdOverrides
		for (let cfIdx = 0; cfIdx < sheet.conditionalFormats.length; cfIdx++) {
			const conditionalFormat = sheet.conditionalFormats[cfIdx]
			if (!conditionalFormat) continue
			const attrs = [`sqref="${escapeXml(conditionalFormat.sqref)}"`]
			if (conditionalFormat.pivot !== undefined) {
				attrs.push(`pivot="${conditionalFormat.pivot ? '1' : '0'}"`)
			}
			out.push(`<conditionalFormatting ${attrs.join(' ')}>`)
			for (let ruleIdx = 0; ruleIdx < conditionalFormat.rules.length; ruleIdx++) {
				const rule = conditionalFormat.rules[ruleIdx]
				if (!rule) continue
				const effectiveDxfId = rule.dxfId ?? cfDxfIdOverrides?.get(`${cfIdx}:${ruleIdx}`)
				const attrs = conditionalFormatRuleAttrs(rule, effectiveDxfId)

				out.push(`<cfRule ${attrs.join(' ')}>`)
				for (const formula of rule.formulas) {
					out.push(`<formula>${escapeXml(formula)}</formula>`)
				}
				if (rule.colorScale) out.push(buildColorScaleXml(rule.colorScale))
				if (rule.dataBar) out.push(buildDataBarXml(rule.dataBar))
				if (rule.iconSet) out.push(buildIconSetXml(rule.iconSet))
				for (const childXml of rule.preservedRuleChildXml ?? []) out.push(childXml)
				out.push('</cfRule>')
			}
			out.push('</conditionalFormatting>')
		}
	}

	const legacyDataValidations = sheet.dataValidations.filter(
		(validation) => validation.source !== 'x14' || !sheet.preservedExtLst,
	)
	if (legacyDataValidations.length > 0) {
		const dataValidationAttrs = [`count="${legacyDataValidations.length}"`]
		const dataValidationSettings = sheet.dataValidationSettings
		if (dataValidationSettings?.disablePrompts !== undefined) {
			dataValidationAttrs.push(
				`disablePrompts="${dataValidationSettings.disablePrompts ? '1' : '0'}"`,
			)
		}
		if (dataValidationSettings?.xWindow !== undefined) {
			dataValidationAttrs.push(`xWindow="${dataValidationSettings.xWindow}"`)
		}
		if (dataValidationSettings?.yWindow !== undefined) {
			dataValidationAttrs.push(`yWindow="${dataValidationSettings.yWindow}"`)
		}
		out.push(`<dataValidations ${dataValidationAttrs.join(' ')}>`)
		for (const validation of legacyDataValidations) {
			const attrs = [`sqref="${escapeXml(validation.sqref)}"`]
			if (validation.uid) attrs.push(`xr:uid="${escapeXml(validation.uid)}"`)
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
		const attrs = collectPageSetupAttrs(sheet.pageSetup)
		if (attrs.length > 0) out.push(`<pageSetup ${attrs.join(' ')}/>`)
	}

	if (sheet.headerFooter) {
		const headerFooterAttrs = collectMixedAttrs({
			differentOddEven: sheet.headerFooter.differentOddEven,
			differentFirst: sheet.headerFooter.differentFirst,
			scaleWithDoc: sheet.headerFooter.scaleWithDoc,
			alignWithMargins: sheet.headerFooter.alignWithMargins,
		})
		out.push(
			`<headerFooter${headerFooterAttrs.length > 0 ? ` ${headerFooterAttrs.join(' ')}` : ''}>`,
		)
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

	if (sheet.preservedControlsXml) {
		out.push(sheet.preservedControlsXml)
	}

	if (tableRelIds.length > 0) {
		out.push(`<tableParts count="${tableRelIds.length}">`)
		for (const relId of tableRelIds) {
			out.push(`<tablePart r:id="${relId}"/>`)
		}
		out.push('</tableParts>')
	}

	if (sheet.preservedExtLst) {
		out.push(
			updateWorksheetExtLstXml(sheet.preservedExtLst, {
				sparklineGroups: sheet.sparklineGroups,
				x14ConditionalFormats: sheet.x14ConditionalFormats,
				x14DataValidations: sheet.x14DataValidations,
			}),
		)
	} else {
		const extLst = buildWorksheetExtLstXml({
			x14ConditionalFormats: sheet.x14ConditionalFormats,
			x14DataValidations: sheet.x14DataValidations,
		})
		if (extLst) out.push(extLst)
	}

	out.push('</worksheet>')
}

function sheetHasPreservedOfficeExtensionXml(sheet: Sheet, needle: string): boolean {
	if (sheet.preservedExtLst?.includes(needle) || sheet.preservedControlsXml?.includes(needle)) {
		return true
	}
	for (const validation of sheet.x14DataValidations) {
		if (recordHasNamePrefix(validation.preservedAttributes, needle)) return true
		if (validation.preservedChildXml?.some((xml) => xml.includes(needle))) return true
	}
	for (const format of sheet.x14ConditionalFormats) {
		if (recordHasNamePrefix(format.preservedRuleAttributes, needle)) return true
		if (format.preservedRuleChildXml?.some((xml) => xml.includes(needle))) return true
		if (recordHasNamePrefix(format.colorScale?.preservedAttributes, needle)) return true
		if (format.colorScale?.preservedChildXml?.some((xml) => xml.includes(needle))) return true
	}
	for (const format of sheet.conditionalFormats) {
		for (const rule of format.rules) {
			if (recordHasNamePrefix(rule.preservedRuleAttributes, needle)) return true
			if (rule.preservedRuleChildXml?.some((xml) => xml.includes(needle))) return true
			if (recordHasNamePrefix(rule.colorScale?.preservedAttributes, needle)) return true
			if (rule.colorScale?.preservedChildXml?.some((xml) => xml.includes(needle))) return true
		}
	}
	return false
}

function recordHasNamePrefix(
	record: Readonly<Record<string, string>> | undefined,
	prefix: string,
): boolean {
	return Object.keys(record ?? {}).some((name) => name.startsWith(prefix))
}

function conditionalFormatRuleAttrs(
	rule: SheetConditionalFormatRule,
	effectiveDxfId?: number,
): string[] {
	const attrs = new Map<string, string>()
	for (const [name, value] of Object.entries(rule.preservedRuleAttributes ?? {})) {
		if (canEmitPreservedConditionalFormatRuleAttr(name)) attrs.set(name, value)
	}
	attrs.set('type', rule.type)
	if (rule.operator) attrs.set('operator', rule.operator)
	if (rule.priority !== undefined) attrs.set('priority', String(rule.priority))
	if (effectiveDxfId !== undefined) attrs.set('dxfId', String(effectiveDxfId))
	if (rule.stopIfTrue) attrs.set('stopIfTrue', '1')
	if (rule.rank !== undefined) attrs.set('rank', String(rule.rank))
	if (rule.percent !== undefined) attrs.set('percent', rule.percent ? '1' : '0')
	if (rule.bottom !== undefined) attrs.set('bottom', rule.bottom ? '1' : '0')
	if (rule.aboveAverage !== undefined) attrs.set('aboveAverage', rule.aboveAverage ? '1' : '0')
	if (rule.equalAverage !== undefined) attrs.set('equalAverage', rule.equalAverage ? '1' : '0')
	if (rule.stdDev !== undefined) attrs.set('stdDev', String(rule.stdDev))
	if (rule.text !== undefined) attrs.set('text', rule.text)
	if (rule.timePeriod) attrs.set('timePeriod', rule.timePeriod)
	return [...attrs.entries()].map(([name, value]) => `${name}="${escapeXml(value)}"`)
}

function canEmitPreservedConditionalFormatRuleAttr(name: string): boolean {
	return name !== 'xmlns' && !name.startsWith('xmlns:')
}

function setSheetViewAttr(attrs: Map<string, string>, name: string, value: string): void {
	if (!attrs.has(name)) attrs.set(name, value)
}

function sheetViewAttrsXml(attrs: ReadonlyMap<string, string>): string {
	return [...attrs].map(([name, value]) => `${name}="${escapeXml(value)}"`).join(' ')
}

function combinedDimensionRange(sheet: Sheet): RangeRef | null {
	let range = sheet.cells.usedRange()
	const blankRange = blankCellsRange(sheet.preservedBlankCells)
	range = mergeRanges(range, blankRange)
	if (sheet.preservedBlankCells.size > 0 && sheet.preservedDimensionRef) {
		try {
			range = mergeRanges(range, parseRange(sheet.preservedDimensionRef))
		} catch {
			// Invalid source dimensions are ignored; concrete cells still define the emitted range.
		}
	}
	return range
}

function blankCellsRange(rows: ReadonlyMap<number, ReadonlyMap<number, string>>): RangeRef | null {
	let minRow = Number.POSITIVE_INFINITY
	let minCol = Number.POSITIVE_INFINITY
	let maxRow = -1
	let maxCol = -1
	for (const [row, cells] of rows) {
		for (const col of cells.keys()) {
			minRow = Math.min(minRow, row)
			minCol = Math.min(minCol, col)
			maxRow = Math.max(maxRow, row)
			maxCol = Math.max(maxCol, col)
		}
	}
	return maxRow >= 0
		? { start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } }
		: null
}

function mergeRanges(left: RangeRef | null, right: RangeRef | null): RangeRef | null {
	if (!left) return right
	if (!right) return left
	return {
		start: {
			row: Math.min(left.start.row, right.start.row),
			col: Math.min(left.start.col, right.start.col),
		},
		end: {
			row: Math.max(left.end.row, right.end.row),
			col: Math.max(left.end.col, right.end.col),
		},
	}
}

function blankCellAttrsXml(rawAttrs: string, ref: string): string {
	const attrs = rawAttrs.trim()
	if (/(?:^|\s)r\s*=/.test(attrs)) return attrs
	return attrs ? `r="${ref}" ${attrs}` : `r="${ref}"`
}

function cachedColumnName(cache: string[], col: number): string {
	let name = cache[col]
	if (name === undefined) {
		name = indexToColumn(col)
		cache[col] = name
	}
	return name
}

function canOmitDenseCellRefs(cells: readonly (readonly [number, Cell])[]): boolean {
	if (cells.length === 0) return false
	for (let index = 0; index < cells.length; index++) {
		const entry = cells[index]
		if (!entry || entry[0] !== index) return false
		const cell = entry[1]
		if (
			(cell.styleId as number) !== 0 ||
			cell.formula ||
			cell.formulaInfo ||
			!canOmitDefaultStyleScalarCellRef(cell)
		) {
			return false
		}
	}
	return true
}

function defaultStyleScalarCellsWithoutRefsXml(
	cells: readonly (readonly [number, Cell])[],
	ssTable: SharedStringTable,
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): string | false {
	if (cells.length === 0) return false
	const first = cells[0]
	if (first?.[1].value.kind === 'number') {
		const numberBody = defaultStyleNumberCellsWithoutRefsXml(cells)
		if (numberBody !== false) return numberBody
	}
	if (!canOmitDenseCellRefs(cells)) return false
	let body = ''
	for (let index = 0; index < cells.length; index++) {
		const entry = cells[index]
		if (!entry) return false
		const cell = entry[1]
		const value = cell.value
		if (value.kind === 'number') {
			body += `<c><v>${value.value}</v></c>`
		} else if (value.kind === 'date') {
			body += `<c><v>${value.serial}</v></c>`
		} else if (value.kind === 'boolean') {
			body += value.value ? '<c t="b"><v>1</v></c>' : '<c t="b"><v>0</v></c>'
		} else if (value.kind === 'error') {
			body += `<c t="e"><v>${escapeXml(value.value)}</v></c>`
		} else if (value.kind === 'empty') {
			body += '<c/>'
		} else if (value.kind === 'string') {
			if (usePlainStrings) {
				body += `<c t="str"><v>${escapeXml(value.value)}</v></c>`
			} else if (useInlineStrings) {
				body += `<c t="inlineStr"><is><t>${escapeXml(value.value)}</t></is></c>`
			} else {
				const index = ssTable.getIndex(value)
				if (index === undefined) return false
				body += `<c t="s"><v>${index}</v></c>`
			}
		} else if (value.kind === 'richText') {
			if (usePlainStrings || useInlineStrings) {
				const runsXml = value.runs.map((run) => inlineStrRunXml(run)).join('')
				body += `<c t="inlineStr"><is>${runsXml}</is></c>`
			} else {
				const index = ssTable.getIndex(value)
				if (index === undefined) return false
				body += `<c t="s"><v>${index}</v></c>`
			}
		} else {
			return false
		}
	}
	return body
}

function defaultStyleNumberCellsWithoutRefsXml(
	cells: readonly (readonly [number, Cell])[],
): string | false {
	let body = ''
	for (let index = 0; index < cells.length; index++) {
		const entry = cells[index]
		if (!entry || entry[0] !== index) return false
		const cell = entry[1]
		if ((cell.styleId as number) !== 0 || cell.formula || cell.formulaInfo) return false
		const value = cell.value
		if (value.kind !== 'number') return false
		body += `<c><v>${value.value}</v></c>`
	}
	return body
}

function canOmitDefaultStyleScalarCellRef(cell: Cell): boolean {
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
	if (value.kind === 'string') return true
	if (value.kind === 'richText') return true
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

function collectPageSetupAttrs(pageSetup: NonNullable<Sheet['pageSetup']>): string[] {
	const attrs = collectMixedAttrs({
		orientation: pageSetup.orientation,
		paperSize: pageSetup.paperSize,
		scale: pageSetup.scale,
		fitToWidth: pageSetup.fitToWidth,
		fitToHeight: pageSetup.fitToHeight,
		firstPageNumber: pageSetup.firstPageNumber,
		copies: pageSetup.copies,
		horizontalDpi: pageSetup.horizontalDpi,
		verticalDpi: pageSetup.verticalDpi,
		pageOrder: pageSetup.pageOrder,
		cellComments: pageSetup.cellComments,
		errors: pageSetup.errors,
		blackAndWhite: pageSetup.blackAndWhite,
		draft: pageSetup.draft,
		useFirstPageNumber: pageSetup.useFirstPageNumber,
		usePrinterDefaults: pageSetup.usePrinterDefaults,
	})
	if (pageSetup.printerSettingsRelId) {
		attrs.push(`r:id="${escapeXml(pageSetup.printerSettingsRelId)}"`)
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

function collectProtectedRangeAttrs(range: Sheet['protectedRanges'][number]): string[] {
	const attrs: string[] = []
	if (range.name !== undefined) attrs.push(`name="${escapeXml(range.name)}"`)
	attrs.push(`sqref="${escapeXml(range.sqref)}"`)
	if (range.password !== undefined) attrs.push(`password="${escapeXml(range.password)}"`)
	if (range.algorithmName !== undefined) {
		attrs.push(`algorithmName="${escapeXml(range.algorithmName)}"`)
	}
	if (range.hashValue !== undefined) attrs.push(`hashValue="${escapeXml(range.hashValue)}"`)
	if (range.saltValue !== undefined) attrs.push(`saltValue="${escapeXml(range.saltValue)}"`)
	if (range.spinCount !== undefined) attrs.push(`spinCount="${range.spinCount}"`)
	if (range.securityDescriptor !== undefined) {
		attrs.push(`securityDescriptor="${escapeXml(range.securityDescriptor)}"`)
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
		cell.formulaInfo?.kind === 'array' ||
		cell.formulaInfo?.kind === 'dataTable'
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
	storedFormulaText?: string,
	preservedCellMetadata?: SheetCellMetadataAttrs,
	useInlineStrings?: boolean,
	usePlainStrings?: boolean,
): void {
	const styleId = cell.styleId as number
	const xfIdx = styleId === 0 ? 0 : (xfMap.get(styleId) ?? 0)

	if (
		cell.formula ||
		cell.formulaInfo?.kind === 'shared' ||
		cell.formulaInfo?.kind === 'array' ||
		cell.formulaInfo?.kind === 'dataTable'
	) {
		out.push(
			formulaCellXml(
				ref,
				cell,
				xfIdx,
				sharedFormulaExpansions,
				storedFormulaText,
				preservedCellMetadata,
			),
		)
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
	storedFormulaText?: string,
	preservedCellMetadata?: SheetCellMetadataAttrs,
): string {
	const formulaText = cell.formula
		? effectiveStoredFormulaText(cell.formula, storedFormulaText)
		: ''
	const sAttr = xfIdx !== 0 ? ` s="${xfIdx}"` : ''
	const dynamicArrayMetadataIndex = dynamicArrayCellMetadataIndex(cell.formulaInfo)
	const metadataAttrs = formulaCellMetadataAttrs(dynamicArrayMetadataIndex, preservedCellMetadata)
	const { typeAttr, valueStr } = formulaValueAttrs(cell.value)
	const tAttr = typeAttr ? ` t="${typeAttr}"` : ''
	const vPart = valueStr !== undefined ? `<v>${valueStr}</v>` : ''
	if (cell.formulaInfo?.kind === 'shared') {
		const expanded = sharedFormulaExpansions.get(ref)
		if (expanded) {
			return `<c r="${ref}"${metadataAttrs}${sAttr}${tAttr}><f>${escapeXml(expanded.formulaText)}</f>${vPart}</c>`
		}
		let fAttrs = `t="shared" si="${escapeXml(cell.formulaInfo.sharedIndex)}"`
		if (cell.formulaInfo.isMaster && cell.formulaInfo.ref) {
			fAttrs += ` ref="${escapeXml(cell.formulaInfo.ref)}"`
		}
		const formulaXml = cell.formulaInfo.isMaster
			? `<f ${fAttrs}>${escapeXml(formulaText)}</f>`
			: `<f ${fAttrs}/>`
		return `<c r="${ref}"${metadataAttrs}${sAttr}${tAttr}>${formulaXml}${vPart}</c>`
	}
	if (cell.formulaInfo?.kind === 'array') {
		let fAttrs = 't="array"'
		if (cell.formulaInfo.ref) fAttrs += ` ref="${escapeXml(cell.formulaInfo.ref)}"`
		return `<c r="${ref}"${metadataAttrs}${sAttr}${tAttr}><f ${fAttrs}>${escapeXml(formulaText)}</f>${vPart}</c>`
	}
	if (cell.formulaInfo?.kind === 'dataTable') {
		const fAttrs = dataTableFormulaAttrs(cell.formulaInfo)
		return `<c r="${ref}"${metadataAttrs}${sAttr}${tAttr}><f ${fAttrs}/>${vPart}</c>`
	}
	return `<c r="${ref}"${metadataAttrs}${sAttr}${tAttr}><f>${escapeXml(formulaText)}</f>${vPart}</c>`
}

function formulaCellMetadataAttrs(
	generatedCm: number | undefined,
	preserved: SheetCellMetadataAttrs | undefined,
): string {
	const attrs: string[] = []
	if (generatedCm !== undefined) {
		attrs.push(`cm="${generatedCm}"`)
	} else if (preserved?.cm !== undefined) {
		attrs.push(`cm="${preserved.cm}"`)
	}
	if (preserved?.vm !== undefined) attrs.push(`vm="${preserved.vm}"`)
	if (preserved?.ph !== undefined) attrs.push(`ph="${preserved.ph ? '1' : '0'}"`)
	return attrs.length > 0 ? ` ${attrs.join(' ')}` : ''
}

function dataTableFormulaAttrs(info: Extract<Cell['formulaInfo'], { kind: 'dataTable' }>): string {
	const attrs = ['t="dataTable"']
	if (info.ref) attrs.push(`ref="${escapeXml(info.ref)}"`)
	if (info.dt2D !== undefined) attrs.push(`dt2D="${info.dt2D ? '1' : '0'}"`)
	if (info.dtr !== undefined) attrs.push(`dtr="${info.dtr ? '1' : '0'}"`)
	if (info.r1) attrs.push(`r1="${escapeXml(info.r1)}"`)
	if (info.r2) attrs.push(`r2="${escapeXml(info.r2)}"`)
	if (info.del1 !== undefined) attrs.push(`del1="${info.del1 ? '1' : '0'}"`)
	if (info.del2 !== undefined) attrs.push(`del2="${info.del2 ? '1' : '0'}"`)
	return attrs.join(' ')
}

function effectiveStoredFormulaText(
	formula: string,
	storedFormulaText: string | undefined,
): string {
	if (
		storedFormulaText !== undefined &&
		normalizeStoredFormulaText(storedFormulaText) === formula
	) {
		return storedFormulaText
	}
	return toStoredFormulaText(formula)
}

function formulaStorageKey(row: number, col: number): string {
	return `${row}:${col}`
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
