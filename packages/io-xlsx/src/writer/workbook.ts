import type { SheetState, Workbook, WorkbookPreservedSheetEntry } from '@ascend/core'
import { toStoredFormulaText } from '../formula-storage.ts'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const NS_X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
const X14_SLICER_CACHES_EXT_URI = '{BBE1A952-AA13-448e-AADC-164F8A28A991}'
const X15_TIMELINE_CACHES_EXT_URI = '{7E03D99C-DC04-49d9-9315-930204A7B6E9}'
const OWNED_WORKBOOK_EXT_URIS = new Set([X14_SLICER_CACHES_EXT_URI, X15_TIMELINE_CACHES_EXT_URI])
const XML_TAG_RE = /<[^>]+>/g
const XML_ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

export interface WorkbookXmlOptions {
	readonly worksheetRelIds?: readonly string[]
	readonly externalReferenceRelIds?: readonly string[]
	readonly pivotCacheRelIds?: readonly string[]
	readonly slicerCacheRelIds?: readonly string[]
	readonly timelineCacheRelIds?: readonly string[]
	readonly chartSheetRelIds?: readonly string[]
	readonly macroSheetRelIds?: readonly string[]
	readonly calcStateDirty?: boolean
	readonly preservedWorkbookXml?: string
}

export function buildWorkbookXml(workbook: Workbook, options: WorkbookXmlOptions = {}): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	const workbookRootAttrs = new Map<string, string>([
		['xmlns', NS_MAIN],
		['xmlns:r', NS_R],
	])
	for (const [name, value] of extractSourceWorkbookNamespaceAttrs(options.preservedWorkbookXml)) {
		if (!workbookRootAttrs.has(name)) workbookRootAttrs.set(name, value)
	}
	for (const [name, value] of extractSourceWorkbookRootAttrs(options.preservedWorkbookXml)) {
		if (!workbookRootAttrs.has(name)) workbookRootAttrs.set(name, value)
	}
	if ((options.slicerCacheRelIds?.length ?? 0) > 0) workbookRootAttrs.set('xmlns:x14', NS_X14)
	if ((options.timelineCacheRelIds?.length ?? 0) > 0) {
		workbookRootAttrs.set('xmlns:x15', NS_X15)
	}
	out.push(
		`<workbook ${Array.from(workbookRootAttrs, ([name, value]) => `${name}="${escapeXml(value)}"`).join(' ')}>`,
	)

	if (workbook.workbookFileVersion) {
		const attrs = collectWorkbookFileVersionAttrs(workbook.workbookFileVersion)
		if (attrs.length > 0) out.push(`<fileVersion ${attrs.join(' ')}/>`)
	}

	if (workbook.workbookFileSharing) {
		const attrs = collectWorkbookFileSharingAttrs(workbook.workbookFileSharing)
		if (attrs.length > 0) out.push(`<fileSharing ${attrs.join(' ')}/>`)
	}

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
	for (const extra of workbook.workbookProperties.extraAttributes ?? []) {
		workbookPrAttrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
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
			for (const extra of view.extraAttributes ?? []) {
				attrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
			}
			out.push(`<workbookView ${attrs.join(' ')}/>`)
		}
		out.push('</bookViews>')
	}

	if (workbook.workbookProtection) {
		const attrs = collectWorkbookProtectionAttrs(workbook.workbookProtection)
		if (attrs.length > 0) out.push(`<workbookProtection ${attrs.join(' ')}/>`)
	}

	out.push('<sheets>')
	const sheetEntries: WorkbookSheetXmlEntry[] = []
	for (let i = 0; i < workbook.sheets.length; i++) {
		const sheet = workbook.sheets[i]
		if (!sheet) continue
		sheetEntries.push({
			kind: 'worksheet',
			name: sheet.name,
			sheetId: storedWorksheetSheetId(i, sheet.id as string),
			relId: options.worksheetRelIds?.[i] ?? `rId${i + 1}`,
			state: sheet.state,
		})
	}
	for (let i = 0; i < workbook.chartSheets.length; i++) {
		const chartSheet = workbook.chartSheets[i]
		const relId = options.chartSheetRelIds?.[i]
		if (!chartSheet || !relId) continue
		sheetEntries.push({
			kind: 'chartsheet',
			name: chartSheet.name,
			sheetId: chartSheet.sheetId,
			relId,
			state: chartSheet.state,
		})
	}
	for (let i = 0; i < workbook.macroSheets.length; i++) {
		const macroSheet = workbook.macroSheets[i]
		const relId = options.macroSheetRelIds?.[i]
		if (!macroSheet || !relId) continue
		sheetEntries.push({
			kind: 'macrosheet',
			name: macroSheet.name,
			sheetId: macroSheet.sheetId,
			relId,
			state: macroSheet.state,
		})
	}
	for (const sheetEntry of orderWorkbookSheetEntries(
		sheetEntries,
		workbook.preservedXml?.sheetEntries,
	)) {
		const attrs = [
			`name="${escapeXml(sheetEntry.name)}"`,
			`sheetId="${escapeXml(sheetEntry.sheetId)}"`,
			`r:id="${escapeXml(sheetEntry.relId)}"`,
		]
		if (sheetEntry.state !== 'visible') attrs.push(`state="${sheetEntry.state}"`)
		out.push(`<sheet ${attrs.join(' ')}/>`)
	}
	out.push('</sheets>')

	if (workbook.definedNames.size > 0) {
		out.push('<definedNames>')
		for (const definedName of workbook.definedNames.list()) {
			const attrs = [`name="${escapeXml(definedName.name)}"`]
			for (const extra of definedName.extraAttributes ?? []) {
				if (isCoreDefinedNameAttribute(extra.name) || !isXmlAttributeName(extra.name)) continue
				attrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
			}
			if (definedName.scope.kind === 'sheet') {
				const scope = definedName.scope
				const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === scope.sheetId)
				if (sheetIndex >= 0) attrs.push(`localSheetId="${sheetIndex}"`)
			}
			if (definedName.hidden !== undefined) {
				attrs.push(`hidden="${definedName.hidden ? '1' : '0'}"`)
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
	if (options.calcStateDirty || cs.fullCalcOnLoad) calcAttrs.push('fullCalcOnLoad="1"')
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
	for (const extra of cs.extraAttributes ?? []) {
		if (isCoreCalcSettingAttribute(extra.name) || !isXmlAttributeName(extra.name)) continue
		calcAttrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
	}
	if (calcAttrs.length > 0) {
		out.push(`<calcPr ${calcAttrs.join(' ')}/>`)
	} else {
		out.push('<calcPr/>')
	}

	const preservedChildrenXml = buildPreservedWorkbookChildrenXml(options)
	if (preservedChildrenXml) out.push(preservedChildrenXml)

	const extLstXml = buildWorkbookExtLstXml(options)
	if (extLstXml) out.push(extLstXml)

	out.push('</workbook>')
	return out.toString()
}

function buildWorkbookExtLstXml(options: WorkbookXmlOptions): string {
	const extXml = extractPreservedWorkbookExtXml(options.preservedWorkbookXml)
	if ((options.slicerCacheRelIds?.length ?? 0) > 0) {
		extXml.push(
			`<ext uri="${X14_SLICER_CACHES_EXT_URI}"><x14:slicerCaches>${options.slicerCacheRelIds
				?.map((relId) => `<x14:slicerCache r:id="${escapeXml(relId)}"/>`)
				.join('')}</x14:slicerCaches></ext>`,
		)
	}
	if ((options.timelineCacheRelIds?.length ?? 0) > 0) {
		extXml.push(
			`<ext uri="${X15_TIMELINE_CACHES_EXT_URI}"><x15:timelineCaches>${options.timelineCacheRelIds
				?.map((relId) => `<x15:timelineCacheRef r:id="${escapeXml(relId)}"/>`)
				.join('')}</x15:timelineCaches></ext>`,
		)
	}
	return extXml.length > 0 ? `<extLst>${extXml.join('')}</extLst>` : ''
}

function buildPreservedWorkbookChildrenXml(options: WorkbookXmlOptions): string {
	if (!options.preservedWorkbookXml) return ''
	return extractDirectChildElements(options.preservedWorkbookXml, 'workbook')
		.filter((child) => {
			const openTagEnd = child.xml.indexOf('>')
			const localName = localNameFromTag(child.xml.slice(1, openTagEnd))
			return !isGeneratedWorkbookChild(localName)
		})
		.map((child) => child.xml)
		.join('')
}

function isGeneratedWorkbookChild(localName: string): boolean {
	return (
		localName === 'fileVersion' ||
		localName === 'fileSharing' ||
		localName === 'workbookPr' ||
		localName === 'workbookProtection' ||
		localName === 'bookViews' ||
		localName === 'sheets' ||
		localName === 'externalReferences' ||
		localName === 'definedNames' ||
		localName === 'calcPr' ||
		localName === 'pivotCaches' ||
		localName === 'extLst'
	)
}

function extractSourceWorkbookNamespaceAttrs(
	sourceXml: string | undefined,
): readonly [string, string][] {
	if (!sourceXml) return []
	const root = findFirstElement(sourceXml, 'workbook', 0, sourceXml.length)
	if (!root) return []
	const attrs: [string, string][] = []
	for (const element of [root, ...extractDirectChildElements(sourceXml, 'workbook', 'extLst')]) {
		for (const [name, value] of parseRawXmlAttributes(element.attrs)) {
			if (name.startsWith('xmlns:')) attrs.push([name, value])
		}
	}
	return attrs
}

function extractSourceWorkbookRootAttrs(
	sourceXml: string | undefined,
): readonly [string, string][] {
	if (!sourceXml) return []
	const root = findFirstElement(sourceXml, 'workbook', 0, sourceXml.length)
	if (!root) return []
	const attrs: [string, string][] = []
	for (const [name, value] of parseRawXmlAttributes(root.attrs)) {
		if (name === 'xmlns' || name.startsWith('xmlns:') || !isXmlAttributeName(name)) continue
		attrs.push([name, value])
	}
	return attrs
}

function extractPreservedWorkbookExtXml(sourceXml: string | undefined): string[] {
	if (!sourceXml) return []
	return extractDirectChildElements(sourceXml, 'workbook', 'extLst')
		.flatMap((extLst) => extractDirectChildElements(extLst.xml, 'extLst', 'ext'))
		.filter((ext) => {
			const uri = parseRawXmlAttributes(ext.attrs).get('uri')
			return !uri || !OWNED_WORKBOOK_EXT_URIS.has(uri)
		})
		.map((ext) => ext.xml)
}

interface XmlElementSlice {
	readonly xml: string
	readonly attrs: string
	readonly start: number
	readonly openEnd: number
	readonly end: number
	readonly selfClosing: boolean
}

function extractDirectChildElements(
	xml: string,
	parentLocalName: string,
	childLocalName?: string,
): XmlElementSlice[] {
	const parent = findFirstElement(xml, parentLocalName, 0, xml.length)
	if (!parent || parent.selfClosing) return []
	const children: XmlElementSlice[] = []
	let cursor = parent.openEnd
	while (cursor < parent.end) {
		const child = findFirstElement(xml, undefined, cursor, parent.end)
		if (!child || child.start >= parent.end) break
		if (
			childLocalName === undefined ||
			localNameFromTag(xml.slice(child.start + 1, child.openEnd)) === childLocalName
		) {
			children.push(child)
		}
		cursor = child.end
	}
	return children
}

function findFirstElement(
	xml: string,
	localName: string | undefined,
	start: number,
	end: number,
): XmlElementSlice | undefined {
	XML_TAG_RE.lastIndex = start
	for (let match = XML_TAG_RE.exec(xml); match && match.index < end; match = XML_TAG_RE.exec(xml)) {
		const raw = match[0]
		const tag = parseTag(raw)
		if (!tag || tag.closing) continue
		if (localName && tag.localName !== localName) continue
		const closeEnd = tag.selfClosing ? match.index + raw.length : findElementEnd(xml, tag.localName)
		if (closeEnd === undefined || closeEnd > end) return undefined
		return {
			xml: xml.slice(match.index, closeEnd),
			attrs: tag.attrs,
			start: match.index,
			openEnd: match.index + raw.length,
			end: closeEnd,
			selfClosing: tag.selfClosing,
		}
	}
	return undefined
}

function findElementEnd(xml: string, localName: string): number | undefined {
	let depth = 1
	for (let match = XML_TAG_RE.exec(xml); match; match = XML_TAG_RE.exec(xml)) {
		const raw = match[0]
		const tag = parseTag(raw)
		if (!tag || tag.localName !== localName) continue
		if (tag.closing) {
			depth--
			if (depth === 0) return match.index + raw.length
		} else if (!tag.selfClosing) {
			depth++
		}
	}
	return undefined
}

function parseTag(rawTag: string):
	| {
			readonly localName: string
			readonly attrs: string
			readonly closing: boolean
			readonly selfClosing: boolean
	  }
	| undefined {
	if (rawTag.startsWith('<?') || rawTag.startsWith('<!')) return undefined
	const closing = rawTag.startsWith('</')
	const bodyStart = closing ? 2 : 1
	const body = rawTag.slice(bodyStart, -1).trim()
	const name = body.match(/^[^\s/>]+/)?.[0]
	if (!name) return undefined
	return {
		localName: name.includes(':') ? (name.split(':').pop() ?? name) : name,
		attrs: closing ? '' : body.slice(name.length).replace(/\/\s*$/, ''),
		closing,
		selfClosing: !closing && /\/\s*>$/.test(rawTag),
	}
}

function localNameFromTag(rawTagBody: string): string {
	const name = rawTagBody.trim().match(/^[^\s/>]+/)?.[0] ?? ''
	return name.includes(':') ? (name.split(':').pop() ?? name) : name
}

function parseRawXmlAttributes(rawAttrs: string): Map<string, string> {
	const attrs = new Map<string, string>()
	XML_ATTR_RE.lastIndex = 0
	for (const match of rawAttrs.matchAll(XML_ATTR_RE)) {
		const name = match[1]
		const value = match[2] ?? match[3]
		if (name && value !== undefined) attrs.set(name, value)
	}
	return attrs
}

function isCoreDefinedNameAttribute(name: string): boolean {
	return name === 'name' || name === 'localSheetId' || name === 'hidden'
}

function isCoreCalcSettingAttribute(name: string): boolean {
	return (
		name === 'calcMode' ||
		name === 'fullCalcOnLoad' ||
		name === 'calcCompleted' ||
		name === 'calcOnSave' ||
		name === 'forceFullCalc' ||
		name === 'calcId' ||
		name === 'iterate' ||
		name === 'iterateCount' ||
		name === 'iterateDelta'
	)
}

function collectWorkbookFileVersionAttrs(
	fileVersion: NonNullable<Workbook['workbookFileVersion']>,
): string[] {
	const attrs: string[] = []
	if (fileVersion.appName) attrs.push(`appName="${escapeXml(fileVersion.appName)}"`)
	if (fileVersion.lastEdited) attrs.push(`lastEdited="${escapeXml(fileVersion.lastEdited)}"`)
	if (fileVersion.lowestEdited) {
		attrs.push(`lowestEdited="${escapeXml(fileVersion.lowestEdited)}"`)
	}
	if (fileVersion.rupBuild) attrs.push(`rupBuild="${escapeXml(fileVersion.rupBuild)}"`)
	if (fileVersion.codeName) attrs.push(`codeName="${escapeXml(fileVersion.codeName)}"`)
	for (const extra of fileVersion.extraAttributes ?? []) {
		if (isCoreWorkbookFileVersionAttribute(extra.name) || !isXmlAttributeName(extra.name)) {
			continue
		}
		attrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
	}
	return attrs
}

function isCoreWorkbookFileVersionAttribute(name: string): boolean {
	return (
		name === 'appName' ||
		name === 'lastEdited' ||
		name === 'lowestEdited' ||
		name === 'rupBuild' ||
		name === 'codeName'
	)
}

function collectWorkbookFileSharingAttrs(
	fileSharing: NonNullable<Workbook['workbookFileSharing']>,
): string[] {
	const attrs: string[] = []
	if (fileSharing.readOnlyRecommended !== undefined) {
		attrs.push(`readOnlyRecommended="${fileSharing.readOnlyRecommended ? '1' : '0'}"`)
	}
	if (fileSharing.userName) attrs.push(`userName="${escapeXml(fileSharing.userName)}"`)
	if (fileSharing.reservationPassword) {
		attrs.push(`reservationPassword="${escapeXml(fileSharing.reservationPassword)}"`)
	}
	if (fileSharing.algorithmName) {
		attrs.push(`algorithmName="${escapeXml(fileSharing.algorithmName)}"`)
	}
	if (fileSharing.hashValue) attrs.push(`hashValue="${escapeXml(fileSharing.hashValue)}"`)
	if (fileSharing.saltValue) attrs.push(`saltValue="${escapeXml(fileSharing.saltValue)}"`)
	if (fileSharing.spinCount !== undefined) attrs.push(`spinCount="${fileSharing.spinCount}"`)
	for (const extra of fileSharing.extraAttributes ?? []) {
		if (isCoreWorkbookFileSharingAttribute(extra.name) || !isXmlAttributeName(extra.name)) {
			continue
		}
		attrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
	}
	return attrs
}

function isCoreWorkbookFileSharingAttribute(name: string): boolean {
	return (
		name === 'readOnlyRecommended' ||
		name === 'userName' ||
		name === 'reservationPassword' ||
		name === 'algorithmName' ||
		name === 'hashValue' ||
		name === 'saltValue' ||
		name === 'spinCount'
	)
}

function isXmlAttributeName(name: string): boolean {
	return /^[A-Za-z_][\w:.-]*$/.test(name)
}

interface WorkbookSheetXmlEntry {
	readonly kind: 'worksheet' | 'chartsheet' | 'macrosheet'
	readonly name: string
	readonly sheetId: string
	readonly relId: string
	readonly state: SheetState
}

function storedWorksheetSheetId(index: number, sheetId: string): string {
	return /^\d+$/.test(sheetId) ? sheetId : String(index + 1)
}

function orderWorkbookSheetEntries(
	entries: readonly WorkbookSheetXmlEntry[],
	preservedEntries: readonly WorkbookPreservedSheetEntry[] | undefined,
): readonly WorkbookSheetXmlEntry[] {
	if (!preservedEntries || preservedEntries.length === 0) return entries
	const byKey = new Map(entries.map((entry) => [sheetOrderKey(entry), entry] as const))
	const ordered: WorkbookSheetXmlEntry[] = []
	const used = new Set<string>()
	for (const preservedEntry of preservedEntries) {
		const key = sheetOrderKey(preservedEntry)
		const entry = byKey.get(key)
		if (!entry || used.has(key)) continue
		ordered.push(entry)
		used.add(key)
	}
	for (const entry of entries) {
		const key = sheetOrderKey(entry)
		if (!used.has(key)) ordered.push(entry)
	}
	return ordered
}

function sheetOrderKey(entry: {
	readonly kind: 'worksheet' | 'chartsheet' | 'macrosheet'
	readonly sheetId: string
}): string {
	return `${entry.kind}:${entry.sheetId}`
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
	for (const extra of protection.extraAttributes ?? []) {
		if (isCoreWorkbookProtectionAttribute(extra.name) || !isXmlAttributeName(extra.name)) continue
		attrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
	}
	return attrs
}

function isCoreWorkbookProtectionAttribute(name: string): boolean {
	return (
		name === 'lockStructure' ||
		name === 'lockWindows' ||
		name === 'lockRevision' ||
		name === 'workbookPassword' ||
		name === 'revisionsPassword' ||
		name === 'workbookAlgorithmName' ||
		name === 'workbookHashValue' ||
		name === 'workbookSaltValue' ||
		name === 'workbookSpinCount' ||
		name === 'revisionsAlgorithmName' ||
		name === 'revisionsHashValue' ||
		name === 'revisionsSaltValue' ||
		name === 'revisionsSpinCount'
	)
}
