import type { SheetState, WorkbookProperties, WorkbookProtection, WorkbookView } from '@ascend/core'
import type { CalcSettings } from '@ascend/schema'
import { DEFAULT_CALC_SETTINGS } from '@ascend/schema'
import { asArray, attr, boolAttr, numAttr, parseXml, type XmlNode } from '../xml.ts'

export interface SheetEntry {
	readonly name: string
	readonly sheetId: string
	readonly rId: string
	readonly state: SheetState
}

export interface DefinedNameEntry {
	readonly name: string
	readonly formula: string
	readonly localSheetId?: number
}

export interface WorkbookInfo {
	readonly sheets: SheetEntry[]
	readonly definedNames: DefinedNameEntry[]
	readonly calcSettings: CalcSettings
	readonly workbookProperties: WorkbookProperties
	readonly workbookProtection: WorkbookProtection | null
	readonly workbookViews: readonly WorkbookView[]
	readonly externalReferenceRelIds: readonly string[]
	readonly pivotCacheEntries: readonly { cacheId: number; relId: string }[]
}

export function parseWorkbookXml(xml: string): WorkbookInfo {
	const doc = parseXml(xml)
	const wb = doc.workbook as XmlNode | undefined
	if (!wb) {
		return {
			sheets: [],
			definedNames: [],
			calcSettings: DEFAULT_CALC_SETTINGS,
			workbookProperties: {},
			workbookProtection: null,
			workbookViews: [],
			externalReferenceRelIds: [],
			pivotCacheEntries: [],
		}
	}

	const sheets = parseSheets(wb)
	const definedNames = parseDefinedNames(wb)
	const calcSettings = parseCalcSettings(wb)
	const workbookProperties = parseWorkbookProperties(wb)
	const workbookProtection = parseWorkbookProtection(wb)
	const workbookViews = parseWorkbookViews(wb)
	const externalReferenceRelIds = parseExternalReferenceRelIds(wb)
	const pivotCacheEntries = parsePivotCacheEntries(wb)

	return {
		sheets,
		definedNames,
		calcSettings,
		workbookProperties,
		workbookProtection,
		workbookViews,
		externalReferenceRelIds,
		pivotCacheEntries,
	}
}

function parsePivotCacheEntries(wb: XmlNode): readonly { cacheId: number; relId: string }[] {
	const pivotCachesNode = wb.pivotCaches as XmlNode | undefined
	if (!pivotCachesNode) return []
	const caches: Array<{ cacheId: number; relId: string }> = []
	for (const pivotCache of asArray<XmlNode>(pivotCachesNode.pivotCache as XmlNode | XmlNode[])) {
		const cacheId = numAttr(pivotCache, 'cacheId')
		const relId = attr(pivotCache, 'r:id') ?? attr(pivotCache, 'id')
		if (cacheId === undefined || !relId) continue
		caches.push({ cacheId, relId })
	}
	return caches
}

function parseWorkbookProtection(wb: XmlNode): WorkbookProtection | null {
	const protection = wb.workbookProtection as XmlNode | undefined
	if (!protection) return null

	const parsed: Record<string, string | number | boolean> = {}
	setBoolAttr(parsed, 'lockStructure', protection, 'lockStructure')
	setBoolAttr(parsed, 'lockWindows', protection, 'lockWindows')
	setBoolAttr(parsed, 'lockRevision', protection, 'lockRevision')
	setStringAttr(parsed, 'workbookPassword', protection, 'workbookPassword')
	setStringAttr(parsed, 'revisionsPassword', protection, 'revisionsPassword')
	setStringAttr(parsed, 'workbookAlgorithmName', protection, 'workbookAlgorithmName')
	setStringAttr(parsed, 'workbookHashValue', protection, 'workbookHashValue')
	setStringAttr(parsed, 'workbookSaltValue', protection, 'workbookSaltValue')
	setNumberAttr(parsed, 'workbookSpinCount', protection, 'workbookSpinCount')
	setStringAttr(parsed, 'revisionsAlgorithmName', protection, 'revisionsAlgorithmName')
	setStringAttr(parsed, 'revisionsHashValue', protection, 'revisionsHashValue')
	setStringAttr(parsed, 'revisionsSaltValue', protection, 'revisionsSaltValue')
	setNumberAttr(parsed, 'revisionsSpinCount', protection, 'revisionsSpinCount')
	return parsed as WorkbookProtection
}

function setBoolAttr(
	target: Record<string, string | number | boolean>,
	key: string,
	node: XmlNode,
	attrName: string,
): void {
	const value = boolAttr(node, attrName)
	if (value !== undefined) target[key] = value
}

function setStringAttr(
	target: Record<string, string | number | boolean>,
	key: string,
	node: XmlNode,
	attrName: string,
): void {
	const value = attr(node, attrName)
	if (value) target[key] = value
}

function setNumberAttr(
	target: Record<string, string | number | boolean>,
	key: string,
	node: XmlNode,
	attrName: string,
): void {
	const value = numAttr(node, attrName)
	if (value !== undefined) target[key] = value
}

function parseSheets(wb: XmlNode): SheetEntry[] {
	const sheetsNode = wb.sheets as XmlNode | undefined
	if (!sheetsNode) return []

	const entries: SheetEntry[] = []
	for (const s of asArray<XmlNode>(sheetsNode.sheet as XmlNode | XmlNode[])) {
		const name = attr(s, 'name')
		const sheetId = attr(s, 'sheetId')
		const rId = attr(s, 'r:id') ?? attr(s, 'id')
		if (!name || !sheetId || !rId) continue

		const stateStr = attr(s, 'state')
		let state: SheetState = 'visible'
		if (stateStr === 'hidden') state = 'hidden'
		else if (stateStr === 'veryHidden') state = 'veryHidden'

		entries.push({ name, sheetId, rId, state })
	}
	return entries
}

function parseDefinedNames(wb: XmlNode): DefinedNameEntry[] {
	const dnNode = wb.definedNames as XmlNode | undefined
	if (!dnNode) return []

	const entries: DefinedNameEntry[] = []
	for (const dn of asArray<XmlNode>(dnNode.definedName as XmlNode | XmlNode[])) {
		const name = attr(dn, 'name')
		const formula = dn['#text'] !== undefined ? String(dn['#text']) : undefined
		if (!name || !formula) continue

		const localId = numAttr(dn, 'localSheetId')
		entries.push(
			localId !== undefined ? { name, formula, localSheetId: localId } : { name, formula },
		)
	}
	return entries
}

function parseCalcSettings(wb: XmlNode): CalcSettings {
	const wbPr = wb.workbookPr as XmlNode | undefined
	const calcPr = wb.calcPr as XmlNode | undefined

	const date1904 = wbPr ? boolAttr(wbPr, 'date1904') : undefined
	const dateSystem: '1900' | '1904' = date1904 ? '1904' : '1900'

	if (!calcPr) {
		return { ...DEFAULT_CALC_SETTINGS, dateSystem }
	}

	const modeStr = attr(calcPr, 'calcMode')
	const calcMode: 'auto' | 'manual' | 'autoNoTable' =
		modeStr === 'manual' ? 'manual' : modeStr === 'autoNoTable' ? 'autoNoTable' : 'auto'
	const fullCalcOnLoad = boolAttr(calcPr, 'fullCalcOnLoad') ?? false
	const calcCompleted = boolAttr(calcPr, 'calcCompleted')
	const calcOnSave = boolAttr(calcPr, 'calcOnSave')
	const forceFullCalc = boolAttr(calcPr, 'forceFullCalc')
	const calcId = numAttr(calcPr, 'calcId')
	const iterate = boolAttr(calcPr, 'iterate') ?? false
	const iterateCount = numAttr(calcPr, 'iterateCount') ?? 100
	const iterateDelta = numAttr(calcPr, 'iterateDelta') ?? 0.001

	return {
		calcMode,
		fullCalcOnLoad,
		...(calcCompleted !== undefined ? { calcCompleted } : {}),
		...(calcOnSave !== undefined ? { calcOnSave } : {}),
		...(forceFullCalc !== undefined ? { forceFullCalc } : {}),
		...(calcId !== undefined ? { calcId } : {}),
		dateSystem,
		iterativeCalc: {
			enabled: iterate,
			maxIterations: iterateCount,
			maxChange: iterateDelta,
		},
	}
}

function parseWorkbookProperties(wb: XmlNode): WorkbookProperties {
	const wbPr = wb.workbookPr as XmlNode | undefined
	if (!wbPr) return {}

	const props: Record<string, unknown> = {}
	const codeName = attr(wbPr, 'codeName')
	if (codeName) props.codeName = codeName
	const defaultThemeVersion = numAttr(wbPr, 'defaultThemeVersion')
	if (defaultThemeVersion !== undefined) props.defaultThemeVersion = defaultThemeVersion
	const filterPrivacy = boolAttr(wbPr, 'filterPrivacy')
	if (filterPrivacy !== undefined) props.filterPrivacy = filterPrivacy
	const date1904 = boolAttr(wbPr, 'date1904')
	if (date1904 !== undefined) props.date1904 = date1904
	return props as WorkbookProperties
}

function parseWorkbookViews(wb: XmlNode): readonly WorkbookView[] {
	const viewsNode = wb.bookViews as XmlNode | undefined
	if (!viewsNode) return []

	return asArray<XmlNode>(viewsNode.workbookView as XmlNode | XmlNode[]).map((view) => {
		const parsed: Record<string, unknown> = {}
		const activeTab = numAttr(view, 'activeTab')
		if (activeTab !== undefined) parsed.activeTab = activeTab
		const firstSheet = numAttr(view, 'firstSheet')
		if (firstSheet !== undefined) parsed.firstSheet = firstSheet
		const visibility = attr(view, 'visibility')
		if (visibility) parsed.visibility = visibility
		const tabRatio = numAttr(view, 'tabRatio')
		if (tabRatio !== undefined) parsed.tabRatio = tabRatio
		return parsed as WorkbookView
	})
}

function parseExternalReferenceRelIds(wb: XmlNode): readonly string[] {
	const refsNode = wb.externalReferences as XmlNode | undefined
	if (!refsNode) return []

	const relIds: string[] = []
	for (const ref of asArray<XmlNode>(refsNode.externalReference as XmlNode | XmlNode[])) {
		const relId = attr(ref, 'r:id') ?? attr(ref, 'id')
		if (relId) relIds.push(relId)
	}
	return relIds
}
