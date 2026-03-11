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

const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g
const SHEET_RE = /<sheet\b([^>]*)\/?>/g
const DEFINED_NAME_RE = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g
const WORKBOOK_PR_RE = /<workbookPr\b([^>]*)\/?>/
const CALC_PR_RE = /<calcPr\b([^>]*)\/?>/
const WORKBOOK_PROTECTION_RE = /<workbookProtection\b([^>]*)\/?>/
const WORKBOOK_VIEW_RE = /<workbookView\b([^>]*)\/?>/g
const EXTERNAL_REFERENCE_RE = /<externalReference\b([^>]*)\/?>/g
const PIVOT_CACHE_RE = /<pivotCache\b([^>]*)\/?>/g

export function parseWorkbookXml(xml: string): WorkbookInfo {
	const scanned = scanWorkbookXml(xml)
	if (scanned) return scanned
	return parseWorkbookXmlWithDom(xml)
}

function scanWorkbookXml(xml: string): WorkbookInfo | null {
	if (!xml.includes('<workbook') || xml.includes('<mc:AlternateContent')) return null
	return {
		sheets: scanSheets(xml),
		definedNames: scanDefinedNames(xml),
		calcSettings: scanCalcSettings(xml),
		workbookProperties: scanWorkbookProperties(xml),
		workbookProtection: scanWorkbookProtection(xml),
		workbookViews: scanWorkbookViews(xml),
		externalReferenceRelIds: scanExternalReferenceRelIds(xml),
		pivotCacheEntries: scanPivotCacheEntries(xml),
	}
}

function parseWorkbookXmlWithDom(xml: string): WorkbookInfo {
	const doc = parseXml(xml)
	const wb = doc.workbook as XmlNode | undefined
	if (!wb) {
		return emptyWorkbookInfo()
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

function emptyWorkbookInfo(): WorkbookInfo {
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

function scanPivotCacheEntries(xml: string): readonly { cacheId: number; relId: string }[] {
	const caches: Array<{ cacheId: number; relId: string }> = []
	collectAttributes(xml, PIVOT_CACHE_RE, (attrs) => {
		const cacheId = numberAttr(attrs, 'cacheId')
		const relId = attrs.get('r:id') ?? attrs.get('id')
		if (cacheId === undefined || !relId) return
		caches.push({ cacheId, relId })
	})
	return caches
}

function scanWorkbookProtection(xml: string): WorkbookProtection | null {
	const attrs = scanSingleTagAttributes(xml, WORKBOOK_PROTECTION_RE)
	if (!attrs) return null

	const parsed: Record<string, string | number | boolean> = {}
	setBoolAttrFromMap(parsed, 'lockStructure', attrs, 'lockStructure')
	setBoolAttrFromMap(parsed, 'lockWindows', attrs, 'lockWindows')
	setBoolAttrFromMap(parsed, 'lockRevision', attrs, 'lockRevision')
	setStringAttrFromMap(parsed, 'workbookPassword', attrs, 'workbookPassword')
	setStringAttrFromMap(parsed, 'revisionsPassword', attrs, 'revisionsPassword')
	setStringAttrFromMap(parsed, 'workbookAlgorithmName', attrs, 'workbookAlgorithmName')
	setStringAttrFromMap(parsed, 'workbookHashValue', attrs, 'workbookHashValue')
	setStringAttrFromMap(parsed, 'workbookSaltValue', attrs, 'workbookSaltValue')
	setNumberAttrFromMap(parsed, 'workbookSpinCount', attrs, 'workbookSpinCount')
	setStringAttrFromMap(parsed, 'revisionsAlgorithmName', attrs, 'revisionsAlgorithmName')
	setStringAttrFromMap(parsed, 'revisionsHashValue', attrs, 'revisionsHashValue')
	setStringAttrFromMap(parsed, 'revisionsSaltValue', attrs, 'revisionsSaltValue')
	setNumberAttrFromMap(parsed, 'revisionsSpinCount', attrs, 'revisionsSpinCount')
	return parsed as WorkbookProtection
}

function setBoolAttrFromMap(
	target: Record<string, string | number | boolean>,
	key: string,
	attrs: Map<string, string>,
	attrName: string,
): void {
	const value = booleanAttr(attrs, attrName)
	if (value !== undefined) target[key] = value
}

function setStringAttrFromMap(
	target: Record<string, string | number | boolean>,
	key: string,
	attrs: Map<string, string>,
	attrName: string,
): void {
	const value = attrs.get(attrName)
	if (value) target[key] = value
}

function setNumberAttrFromMap(
	target: Record<string, string | number | boolean>,
	key: string,
	attrs: Map<string, string>,
	attrName: string,
): void {
	const value = numberAttr(attrs, attrName)
	if (value !== undefined) target[key] = value
}

function scanSheets(xml: string): SheetEntry[] {
	const entries: SheetEntry[] = []
	collectAttributes(xml, SHEET_RE, (attrs) => {
		const name = attrs.get('name')
		const sheetId = attrs.get('sheetId')
		const rId = attrs.get('r:id') ?? attrs.get('id')
		if (!name || !sheetId || !rId) return

		const stateStr = attrs.get('state')
		let state: SheetState = 'visible'
		if (stateStr === 'hidden') state = 'hidden'
		else if (stateStr === 'veryHidden') state = 'veryHidden'

		entries.push({ name, sheetId, rId, state })
	})
	return entries
}

function scanDefinedNames(xml: string): DefinedNameEntry[] {
	const entries: DefinedNameEntry[] = []
	for (const match of xml.matchAll(DEFINED_NAME_RE)) {
		const rawAttrs = match[1]
		const rawFormula = match[2]
		if (!rawAttrs || rawFormula === undefined) continue
		const attrs = parseAttributes(rawAttrs)
		const name = attrs.get('name')
		const formula = decodeXmlText(rawFormula).trim()
		if (!name || !formula) continue

		const localId = numberAttr(attrs, 'localSheetId')
		entries.push(
			localId !== undefined ? { name, formula, localSheetId: localId } : { name, formula },
		)
	}
	return entries
}

function scanCalcSettings(xml: string): CalcSettings {
	const wbPrAttrs = scanSingleTagAttributes(xml, WORKBOOK_PR_RE)
	const calcPrAttrs = scanSingleTagAttributes(xml, CALC_PR_RE)

	const date1904 = wbPrAttrs ? booleanAttr(wbPrAttrs, 'date1904') : undefined
	const dateSystem: '1900' | '1904' = date1904 ? '1904' : '1900'

	if (!calcPrAttrs) {
		return { ...DEFAULT_CALC_SETTINGS, dateSystem }
	}

	const modeStr = calcPrAttrs.get('calcMode')
	const calcMode: 'auto' | 'manual' | 'autoNoTable' =
		modeStr === 'manual' ? 'manual' : modeStr === 'autoNoTable' ? 'autoNoTable' : 'auto'
	const fullCalcOnLoad = booleanAttr(calcPrAttrs, 'fullCalcOnLoad') ?? false
	const calcCompleted = booleanAttr(calcPrAttrs, 'calcCompleted')
	const calcOnSave = booleanAttr(calcPrAttrs, 'calcOnSave')
	const forceFullCalc = booleanAttr(calcPrAttrs, 'forceFullCalc')
	const calcId = numberAttr(calcPrAttrs, 'calcId')
	const iterate = booleanAttr(calcPrAttrs, 'iterate') ?? false
	const iterateCount = numberAttr(calcPrAttrs, 'iterateCount') ?? 100
	const iterateDelta = numberAttr(calcPrAttrs, 'iterateDelta') ?? 0.001

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

function scanWorkbookProperties(xml: string): WorkbookProperties {
	const attrs = scanSingleTagAttributes(xml, WORKBOOK_PR_RE)
	if (!attrs) return {}

	const props: Record<string, unknown> = {}
	const codeName = attrs.get('codeName')
	if (codeName) props.codeName = codeName
	const defaultThemeVersion = numberAttr(attrs, 'defaultThemeVersion')
	if (defaultThemeVersion !== undefined) props.defaultThemeVersion = defaultThemeVersion
	const filterPrivacy = booleanAttr(attrs, 'filterPrivacy')
	if (filterPrivacy !== undefined) props.filterPrivacy = filterPrivacy
	const date1904 = booleanAttr(attrs, 'date1904')
	if (date1904 !== undefined) props.date1904 = date1904
	return props as WorkbookProperties
}

function scanWorkbookViews(xml: string): readonly WorkbookView[] {
	const views: WorkbookView[] = []
	collectAttributes(xml, WORKBOOK_VIEW_RE, (attrs) => {
		const parsed: Record<string, unknown> = {}
		const activeTab = numberAttr(attrs, 'activeTab')
		if (activeTab !== undefined) parsed.activeTab = activeTab
		const firstSheet = numberAttr(attrs, 'firstSheet')
		if (firstSheet !== undefined) parsed.firstSheet = firstSheet
		const visibility = attrs.get('visibility')
		if (visibility) parsed.visibility = visibility
		const tabRatio = numberAttr(attrs, 'tabRatio')
		if (tabRatio !== undefined) parsed.tabRatio = tabRatio
		views.push(parsed as WorkbookView)
	})
	return views
}

function scanExternalReferenceRelIds(xml: string): readonly string[] {
	const relIds: string[] = []
	collectAttributes(xml, EXTERNAL_REFERENCE_RE, (attrs) => {
		const relId = attrs.get('r:id') ?? attrs.get('id')
		if (relId) relIds.push(relId)
	})
	return relIds
}

function scanSingleTagAttributes(xml: string, pattern: RegExp): Map<string, string> | undefined {
	const match = pattern.exec(xml)
	if (!match?.[1]) return undefined
	return parseAttributes(match[1])
}

function collectAttributes(
	xml: string,
	pattern: RegExp,
	visit: (attrs: Map<string, string>) => void,
): void {
	for (const match of xml.matchAll(pattern)) {
		const rawAttrs = match[1]
		if (!rawAttrs) continue
		visit(parseAttributes(rawAttrs))
	}
}

function parseAttributes(rawAttrs: string): Map<string, string> {
	const attrs = new Map<string, string>()
	for (const attrMatch of rawAttrs.matchAll(ATTR_RE)) {
		const key = attrMatch[1]
		const value = attrMatch[2]
		if (!key || value === undefined) continue
		attrs.set(key, decodeXmlText(value))
	}
	return attrs
}

function numberAttr(attrs: Map<string, string>, name: string): number | undefined {
	const value = attrs.get(name)
	if (value === undefined) return undefined
	const parsed = Number(value)
	return Number.isNaN(parsed) ? undefined : parsed
}

function booleanAttr(attrs: Map<string, string>, name: string): boolean | undefined {
	const value = attrs.get(name)
	if (value === undefined) return undefined
	return value === '1' || value === 'true'
}

function decodeXmlText(text: string): string {
	if (!text.includes('&')) return text
	return text
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&')
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
