import type { SheetState } from '@ascend/core'
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
}

export function parseWorkbookXml(xml: string): WorkbookInfo {
	const doc = parseXml(xml)
	const wb = doc.workbook as XmlNode | undefined
	if (!wb) {
		return {
			sheets: [],
			definedNames: [],
			calcSettings: DEFAULT_CALC_SETTINGS,
		}
	}

	const sheets = parseSheets(wb)
	const definedNames = parseDefinedNames(wb)
	const calcSettings = parseCalcSettings(wb)

	return { sheets, definedNames, calcSettings }
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
	const calcMode: 'auto' | 'manual' = modeStr === 'manual' ? 'manual' : 'auto'
	const fullCalcOnLoad = boolAttr(calcPr, 'fullCalcOnLoad') ?? false
	const iterate = boolAttr(calcPr, 'iterate') ?? false
	const iterateCount = numAttr(calcPr, 'iterateCount') ?? 100
	const iterateDelta = numAttr(calcPr, 'iterateDelta') ?? 0.001

	return {
		calcMode,
		fullCalcOnLoad,
		dateSystem,
		iterativeCalc: {
			enabled: iterate,
			maxIterations: iterateCount,
			maxChange: iterateDelta,
		},
	}
}
