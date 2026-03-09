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
	readonly path?: string
	readonly xml?: string
	readonly xfByStyleId: Readonly<Record<number, number>>
}

export interface WorkbookPreservedTheme {
	readonly path: string
	readonly contentType: string
	readonly xml?: string
}

export interface WorkbookPreservedSharedStrings {
	readonly path: string
	readonly xml?: string
}

export interface WorkbookPreservedXml {
	readonly workbookPath?: string
	readonly workbookXml?: string
	readonly workbookRelsPath?: string
	readonly workbookRelsXml?: string
	readonly contentType?: string
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
	preservedSharedStrings: WorkbookPreservedSharedStrings | null = null
	preservedXml: WorkbookPreservedXml | null = null
	sourceArchiveBytes: Uint8Array | null = null
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

	clone(): Workbook {
		const clone = new Workbook(this.id)
		clone.calcSettings = structuredClone(this.calcSettings)
		clone.workbookProperties = structuredClone(this.workbookProperties)
		clone.workbookProtection = this.workbookProtection
			? structuredClone(this.workbookProtection)
			: null
		clone.styleMetadata = structuredClone(this.styleMetadata)
		clone.themeMetadata = structuredClone(this.themeMetadata)
		clone.preservedStyles = this.preservedStyles ? structuredClone(this.preservedStyles) : null
		clone.preservedTheme = this.preservedTheme ? structuredClone(this.preservedTheme) : null
		clone.preservedSharedStrings = this.preservedSharedStrings
			? structuredClone(this.preservedSharedStrings)
			: null
		clone.preservedXml = this.preservedXml ? structuredClone(this.preservedXml) : null
		clone.sourceArchiveBytes = this.sourceArchiveBytes
		for (const sheet of this.sheets) clone.sheets.push(sheet.clone())
		for (const definedName of this.definedNames.list()) {
			clone.definedNames.set(
				definedName.name,
				definedName.formula,
				structuredClone(definedName.scope),
			)
		}
		for (let styleId = 0 as import('./ids.ts').StyleId; styleId < this.styles.size; styleId++) {
			const style = this.styles.get(styleId)
			if (style) clone.styles.register(structuredClone(style))
		}
		clone.differentialStyles.push(...this.differentialStyles.map((style) => structuredClone(style)))
		clone.pivotCaches.push(...this.pivotCaches.map((entry) => structuredClone(entry)))
		clone.pivotTables.push(...this.pivotTables.map((entry) => structuredClone(entry)))
		clone.slicerCaches.push(...this.slicerCaches.map((entry) => structuredClone(entry)))
		clone.slicers.push(...this.slicers.map((entry) => structuredClone(entry)))
		clone.workbookViews.push(...this.workbookViews.map((view) => structuredClone(view)))
		clone.externalReferences.push(
			...this.externalReferences.map((reference) => structuredClone(reference)),
		)
		return clone
	}
}

export function createWorkbook(id?: WorkbookId): Workbook {
	return new Workbook(id)
}
