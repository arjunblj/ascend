import type { Workbook } from '@ascend/core'
import { toStoredFormulaText } from '../formula-storage.ts'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface WorkbookXmlOptions {
	readonly externalReferenceRelIds?: readonly string[]
	readonly pivotCacheRelIds?: readonly string[]
	readonly chartSheetRelIds?: readonly string[]
	readonly calcStateDirty?: boolean
}

export function buildWorkbookXml(workbook: Workbook, options: WorkbookXmlOptions = {}): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`)

	const workbookPrAttrs: string[] = []
	if (workbook.calcSettings.dateSystem === '1904' || workbook.workbookProperties.date1904) {
		workbookPrAttrs.push('date1904="1"')
	}
	if (workbook.workbookProperties.codeName) {
		workbookPrAttrs.push(`codeName="${escapeXml(workbook.workbookProperties.codeName)}"`)
	}
	if (workbook.workbookProperties.defaultThemeVersion !== undefined) {
		workbookPrAttrs.push(`defaultThemeVersion="${workbook.workbookProperties.defaultThemeVersion}"`)
	}
	if (workbook.workbookProperties.filterPrivacy !== undefined) {
		workbookPrAttrs.push(`filterPrivacy="${workbook.workbookProperties.filterPrivacy ? '1' : '0'}"`)
	}
	if (workbookPrAttrs.length > 0) {
		out.push(`<workbookPr ${workbookPrAttrs.join(' ')}/>`)
	}

	if (workbook.workbookViews.length > 0) {
		out.push('<bookViews>')
		for (const view of workbook.workbookViews) {
			const attrs: string[] = []
			if (view.activeTab !== undefined) attrs.push(`activeTab="${view.activeTab}"`)
			if (view.firstSheet !== undefined) attrs.push(`firstSheet="${view.firstSheet}"`)
			if (view.visibility) attrs.push(`visibility="${escapeXml(view.visibility)}"`)
			if (view.tabRatio !== undefined) attrs.push(`tabRatio="${view.tabRatio}"`)
			out.push(`<workbookView ${attrs.join(' ')}/>`)
		}
		out.push('</bookViews>')
	}

	if (workbook.workbookProtection) {
		const attrs = collectWorkbookProtectionAttrs(workbook.workbookProtection)
		if (attrs.length > 0) out.push(`<workbookProtection ${attrs.join(' ')}/>`)
	}

	out.push('<sheets>')
	for (let i = 0; i < workbook.sheets.length; i++) {
		const sheet = workbook.sheets[i]
		if (!sheet) continue
		const attrs = [`name="${escapeXml(sheet.name)}"`, `sheetId="${i + 1}"`, `r:id="rId${i + 1}"`]
		if (sheet.state !== 'visible') {
			attrs.push(`state="${sheet.state}"`)
		}
		out.push(`<sheet ${attrs.join(' ')}/>`)
	}
	for (let i = 0; i < workbook.chartSheets.length; i++) {
		const chartSheet = workbook.chartSheets[i]
		const relId = options.chartSheetRelIds?.[i]
		if (!chartSheet || !relId) continue
		const attrs = [
			`name="${escapeXml(chartSheet.name)}"`,
			`sheetId="${escapeXml(chartSheet.sheetId)}"`,
			`r:id="${escapeXml(relId)}"`,
		]
		if (chartSheet.state !== 'visible') {
			attrs.push(`state="${chartSheet.state}"`)
		}
		out.push(`<sheet ${attrs.join(' ')}/>`)
	}
	out.push('</sheets>')

	if (workbook.definedNames.size > 0) {
		out.push('<definedNames>')
		for (const definedName of workbook.definedNames.list()) {
			const attrs = [`name="${escapeXml(definedName.name)}"`]
			if (definedName.scope.kind === 'sheet') {
				const scope = definedName.scope
				const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === scope.sheetId)
				if (sheetIndex >= 0) attrs.push(`localSheetId="${sheetIndex}"`)
			}
			out.push(
				`<definedName ${attrs.join(' ')}>${escapeXml(toStoredFormulaText(definedName.formula))}</definedName>`,
			)
		}
		out.push('</definedNames>')
	}

	if (
		workbook.pivotCaches.length > 0 &&
		options.pivotCacheRelIds &&
		options.pivotCacheRelIds.length === workbook.pivotCaches.length
	) {
		out.push('<pivotCaches>')
		for (let i = 0; i < workbook.pivotCaches.length; i++) {
			const cache = workbook.pivotCaches[i]
			const relId = options.pivotCacheRelIds[i]
			if (!cache || !relId) continue
			const attrs: string[] = []
			if (cache.cacheId !== undefined) attrs.push(`cacheId="${cache.cacheId}"`)
			attrs.push(`r:id="${escapeXml(relId)}"`)
			out.push(`<pivotCache ${attrs.join(' ')}/>`)
		}
		out.push('</pivotCaches>')
	}

	if (
		workbook.externalReferences.length > 0 &&
		options.externalReferenceRelIds &&
		options.externalReferenceRelIds.length > 0
	) {
		out.push('<externalReferences>')
		for (const relId of options.externalReferenceRelIds) {
			out.push(`<externalReference r:id="${relId}"/>`)
		}
		out.push('</externalReferences>')
	}

	const cs = workbook.calcSettings
	const calcAttrs: string[] = []
	const calcMode = options.calcStateDirty ? 'auto' : cs.calcMode
	if (calcMode !== 'auto') calcAttrs.push(`calcMode="${calcMode}"`)
	calcAttrs.push('fullCalcOnLoad="1"')
	if (options.calcStateDirty || cs.calcCompleted === false) calcAttrs.push('calcCompleted="0"')
	else if (cs.calcCompleted === true) calcAttrs.push('calcCompleted="1"')
	if (options.calcStateDirty || cs.calcOnSave === true) calcAttrs.push('calcOnSave="1"')
	else if (cs.calcOnSave === false) calcAttrs.push('calcOnSave="0"')
	if (options.calcStateDirty || cs.forceFullCalc) calcAttrs.push('forceFullCalc="1"')
	else if (cs.forceFullCalc === false) calcAttrs.push('forceFullCalc="0"')
	if (cs.calcId !== undefined) calcAttrs.push(`calcId="${cs.calcId}"`)
	if (cs.iterativeCalc.enabled) {
		calcAttrs.push('iterate="1"')
		calcAttrs.push(`iterateCount="${cs.iterativeCalc.maxIterations}"`)
		calcAttrs.push(`iterateDelta="${cs.iterativeCalc.maxChange}"`)
	}
	if (calcAttrs.length > 0) {
		out.push(`<calcPr ${calcAttrs.join(' ')}/>`)
	} else {
		out.push('<calcPr/>')
	}

	out.push('</workbook>')
	return out.toString()
}

function collectWorkbookProtectionAttrs(
	protection: NonNullable<Workbook['workbookProtection']>,
): string[] {
	const attrs: string[] = []
	if (protection.lockStructure !== undefined) {
		attrs.push(`lockStructure="${protection.lockStructure ? '1' : '0'}"`)
	}
	if (protection.lockWindows !== undefined) {
		attrs.push(`lockWindows="${protection.lockWindows ? '1' : '0'}"`)
	}
	if (protection.lockRevision !== undefined) {
		attrs.push(`lockRevision="${protection.lockRevision ? '1' : '0'}"`)
	}
	if (protection.workbookPassword)
		attrs.push(`workbookPassword="${escapeXml(protection.workbookPassword)}"`)
	if (protection.revisionsPassword)
		attrs.push(`revisionsPassword="${escapeXml(protection.revisionsPassword)}"`)
	if (protection.workbookAlgorithmName) {
		attrs.push(`workbookAlgorithmName="${escapeXml(protection.workbookAlgorithmName)}"`)
	}
	if (protection.workbookHashValue) {
		attrs.push(`workbookHashValue="${escapeXml(protection.workbookHashValue)}"`)
	}
	if (protection.workbookSaltValue) {
		attrs.push(`workbookSaltValue="${escapeXml(protection.workbookSaltValue)}"`)
	}
	if (protection.workbookSpinCount !== undefined) {
		attrs.push(`workbookSpinCount="${protection.workbookSpinCount}"`)
	}
	if (protection.revisionsAlgorithmName) {
		attrs.push(`revisionsAlgorithmName="${escapeXml(protection.revisionsAlgorithmName)}"`)
	}
	if (protection.revisionsHashValue) {
		attrs.push(`revisionsHashValue="${escapeXml(protection.revisionsHashValue)}"`)
	}
	if (protection.revisionsSaltValue) {
		attrs.push(`revisionsSaltValue="${escapeXml(protection.revisionsSaltValue)}"`)
	}
	if (protection.revisionsSpinCount !== undefined) {
		attrs.push(`revisionsSpinCount="${protection.revisionsSpinCount}"`)
	}
	return attrs
}
