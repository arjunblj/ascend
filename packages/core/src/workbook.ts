import { type CalcSettings, DEFAULT_CALC_SETTINGS } from '@ascend/schema'
import { DefinedNameCollection } from './defined-name.ts'
import { createWorkbookId, type SheetId, type WorkbookId } from './ids.ts'
import type { PivotCacheInfo, PivotTableInfo, SlicerCacheInfo, SlicerInfo } from './pivot.ts'
import { createSheet, type Sheet } from './sheet.ts'
import type { CellStyle } from './style.ts'
import { StyleRegistry } from './style-registry.ts'

export interface WorkbookView {
	readonly activeTab?: number
	readonly firstSheet?: number
	readonly visibility?: string
	readonly tabRatio?: number
}

export interface WorkbookProperties {
	readonly codeName?: string
	readonly defaultThemeVersion?: number
	readonly filterPrivacy?: boolean
	readonly date1904?: boolean
}

export interface WorkbookStyleMetadata {
	readonly numFmtCount: number
	readonly fontCount: number
	readonly fillCount: number
	readonly borderCount: number
	readonly cellXfCount: number
	readonly dxfCount: number
	readonly tableStyleCount: number
}

export interface WorkbookProtection {
	readonly lockStructure?: boolean
	readonly lockWindows?: boolean
	readonly lockRevision?: boolean
	readonly workbookPassword?: string
	readonly revisionsPassword?: string
	readonly workbookAlgorithmName?: string
	readonly workbookHashValue?: string
	readonly workbookSaltValue?: string
	readonly workbookSpinCount?: number
	readonly revisionsAlgorithmName?: string
	readonly revisionsHashValue?: string
	readonly revisionsSaltValue?: string
	readonly revisionsSpinCount?: number
}

export interface WorkbookThemeMetadata {
	readonly name?: string
	readonly colorSchemeName?: string
	readonly colorCount: number
	readonly majorFontLatin?: string
	readonly minorFontLatin?: string
}

export interface WorkbookPreservedStyles {
	readonly xml: string
	readonly xfByStyleId: Readonly<Record<number, number>>
}

export interface WorkbookPreservedTheme {
	readonly path: string
	readonly contentType: string
	readonly xml: string
}

export interface WorkbookPreservedXml {
	readonly workbookXml: string
	readonly workbookRelsXml?: string
}

export class Workbook {
	readonly id: WorkbookId
	readonly sheets: Sheet[] = []
	readonly definedNames = new DefinedNameCollection()
	readonly styles = new StyleRegistry()
	readonly differentialStyles: CellStyle[] = []
	readonly pivotCaches: PivotCacheInfo[] = []
	readonly pivotTables: PivotTableInfo[] = []
	readonly slicerCaches: SlicerCacheInfo[] = []
	readonly slicers: SlicerInfo[] = []
	readonly workbookViews: WorkbookView[] = []
	readonly externalReferences: string[] = []
	workbookProperties: WorkbookProperties = {}
	workbookProtection: WorkbookProtection | null = null
	styleMetadata: WorkbookStyleMetadata = {
		numFmtCount: 0,
		fontCount: 0,
		fillCount: 0,
		borderCount: 0,
		cellXfCount: 0,
		dxfCount: 0,
		tableStyleCount: 0,
	}
	themeMetadata: WorkbookThemeMetadata = {
		colorCount: 0,
	}
	preservedStyles: WorkbookPreservedStyles | null = null
	preservedTheme: WorkbookPreservedTheme | null = null
	preservedXml: WorkbookPreservedXml | null = null
	calcSettings: CalcSettings

	constructor(id?: WorkbookId) {
		this.id = id ?? createWorkbookId()
		this.calcSettings = DEFAULT_CALC_SETTINGS
	}

	addSheet(name: string, id?: SheetId): Sheet {
		const sheet = createSheet(name, id)
		this.sheets.push(sheet)
		return sheet
	}

	getSheet(nameOrId: string): Sheet | undefined {
		return this.sheets.find((s) => s.name === nameOrId || s.id === nameOrId)
	}

	removeSheet(nameOrId: string): boolean {
		const index = this.sheets.findIndex((s) => s.name === nameOrId || s.id === nameOrId)
		if (index === -1) return false
		this.sheets.splice(index, 1)
		return true
	}
}

export function createWorkbook(id?: WorkbookId): Workbook {
	return new Workbook(id)
}
