import { type CalcSettings, DEFAULT_CALC_SETTINGS } from '@ascend/schema'
import { type ActiveContentInfo, cloneActiveContentInfo } from './active-content.ts'
import type { ChartPartInfo, ChartSheetInfo } from './chart.ts'
import type { WorkbookConnectionPartInfo } from './connection.ts'
import type { WorkbookDataModelPartInfo } from './data-model.ts'
import { DefinedNameCollection } from './defined-name.ts'
import { createWorkbookId, type SheetId, type WorkbookId } from './ids.ts'
import type {
	PivotCacheInfo,
	PivotTableInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TimelineCacheInfo,
	TimelineInfo,
} from './pivot.ts'
import { createSheet, type Sheet } from './sheet.ts'
import type { CellStyle } from './style.ts'
import { cloneCellStyle } from './style-clone.ts'
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

export interface NamedStyleInfo {
	readonly name: string
	readonly builtinId?: number
	readonly hidden?: boolean
}

export interface WorkbookStyleMetadata {
	readonly numFmtCount: number
	readonly fontCount: number
	readonly fillCount: number
	readonly borderCount: number
	readonly cellXfCount: number
	readonly dxfCount: number
	readonly tableStyleCount: number
	readonly namedStyles?: readonly NamedStyleInfo[]
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

export interface WorkbookThemeColor {
	readonly slot: string
	readonly rgb?: string
	readonly systemColor?: string
	readonly lastColor?: string
}

export interface WorkbookPreservedStyles {
	readonly path?: string
	readonly xml?: string
	readonly xfByStyleId: Readonly<Record<number, number>>
	readonly baseStyleIdByStyleId?: Readonly<Record<number, number>>
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

export interface WorkbookPreservedMetadata {
	readonly path: string
	readonly contentType: string
	readonly xml?: string
	readonly dynamicArrayMetadata?: readonly {
		readonly metadataIndex: number
		readonly collapsed?: boolean
	}[]
}

export interface WorkbookPreservedXml {
	readonly workbookPath?: string
	readonly workbookXml?: string
	readonly workbookRelsPath?: string
	readonly workbookRelsXml?: string
	readonly contentType?: string
	readonly contentTypeDefaults?: readonly WorkbookPreservedContentTypeDefault[]
	readonly sheetEntries?: readonly WorkbookPreservedSheetEntry[]
}

export interface WorkbookPreservedContentTypeDefault {
	readonly extension: string
	readonly contentType: string
}

export interface WorkbookPreservedSheetEntry {
	readonly kind: 'worksheet' | 'chartsheet' | 'macrosheet'
	readonly sheetId: string
	readonly name: string
}

export interface WorkbookMacroSheetInfo {
	readonly name: string
	readonly sheetId: string
	readonly relId: string
	readonly partPath: string
	readonly state: 'visible' | 'hidden' | 'veryHidden'
	readonly relationshipCount: number
	readonly dimensionRef?: string
	readonly cellCount?: number
	readonly formulaCount?: number
}

export interface ExternalReferenceInfo {
	readonly partPath: string
	readonly relId?: string
	readonly linkRelId?: string
	readonly target?: string
	readonly targetMode?: string
}

export class Workbook {
	readonly id: WorkbookId
	readonly sheets: Sheet[] = []
	private _sheetIndex: Map<string, number> | null = null
	readonly definedNames = new DefinedNameCollection()
	readonly styles = new StyleRegistry()
	readonly differentialStyles: CellStyle[] = []
	readonly pivotCaches: PivotCacheInfo[] = []
	readonly pivotTables: PivotTableInfo[] = []
	readonly slicerCaches: SlicerCacheInfo[] = []
	readonly slicers: SlicerInfo[] = []
	readonly timelineCaches: TimelineCacheInfo[] = []
	readonly timelines: TimelineInfo[] = []
	readonly chartParts: ChartPartInfo[] = []
	readonly chartSheets: ChartSheetInfo[] = []
	readonly macroSheets: WorkbookMacroSheetInfo[] = []
	readonly connectionParts: WorkbookConnectionPartInfo[] = []
	readonly dataModelParts: WorkbookDataModelPartInfo[] = []
	readonly activeContent: ActiveContentInfo[] = []
	readonly workbookViews: WorkbookView[] = []
	readonly externalReferences: string[] = []
	readonly externalReferenceDetails: ExternalReferenceInfo[] = []
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
	readonly themeColors: WorkbookThemeColor[] = []
	preservedStyles: WorkbookPreservedStyles | null = null
	preservedTheme: WorkbookPreservedTheme | null = null
	preservedSharedStrings: WorkbookPreservedSharedStrings | null = null
	preservedMetadata: WorkbookPreservedMetadata | null = null
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
		this._sheetIndex = null
		return sheet
	}

	private _getSheetIndex(): Map<string, number> {
		if (!this._sheetIndex) {
			this._sheetIndex = new Map()
			for (const [i, s] of this.sheets.entries()) {
				this._sheetIndex.set(s.name, i)
				this._sheetIndex.set(s.id, i)
			}
		}
		return this._sheetIndex
	}

	getSheet(nameOrId: string): Sheet | undefined {
		const idx = this._getSheetIndex().get(nameOrId)
		return idx !== undefined ? this.sheets[idx] : undefined
	}

	/** Call after modifying sheets array directly (e.g. from apply operations). */
	invalidateSheetCache(): void {
		this._sheetIndex = null
	}

	removeSheet(nameOrId: string): boolean {
		const index = this.sheets.findIndex((s) => s.name === nameOrId || s.id === nameOrId)
		if (index === -1) return false
		this.sheets.splice(index, 1)
		this._sheetIndex = null
		return true
	}

	clone(): Workbook {
		const clone = new Workbook(this.id)
		clone.calcSettings = {
			...this.calcSettings,
			iterativeCalc: { ...this.calcSettings.iterativeCalc },
		}
		clone.workbookProperties = { ...this.workbookProperties }
		clone.workbookProtection = this.workbookProtection ? { ...this.workbookProtection } : null
		clone.styleMetadata = { ...this.styleMetadata }
		clone.themeMetadata = { ...this.themeMetadata }
		clone.themeColors.push(...this.themeColors.map((color) => ({ ...color })))
		clone.preservedStyles = this.preservedStyles
			? {
					...this.preservedStyles,
					xfByStyleId: { ...this.preservedStyles.xfByStyleId },
					...(this.preservedStyles.baseStyleIdByStyleId
						? { baseStyleIdByStyleId: { ...this.preservedStyles.baseStyleIdByStyleId } }
						: {}),
				}
			: null
		clone.preservedTheme = this.preservedTheme ? { ...this.preservedTheme } : null
		clone.preservedSharedStrings = this.preservedSharedStrings
			? { ...this.preservedSharedStrings }
			: null
		clone.preservedMetadata = this.preservedMetadata ? { ...this.preservedMetadata } : null
		clone.preservedXml = this.preservedXml ? { ...this.preservedXml } : null
		clone.sourceArchiveBytes = this.sourceArchiveBytes
		for (const sheet of this.sheets) clone.sheets.push(sheet.clone())
		clone.definedNames.copyFrom(this.definedNames)
		clone.styles.copyFrom(this.styles)
		clone.differentialStyles.push(...this.differentialStyles.map(cloneCellStyle))
		clone.pivotCaches.push(...this.pivotCaches.map(clonePivotCacheInfo))
		clone.pivotTables.push(...this.pivotTables.map(clonePivotTableInfo))
		clone.slicerCaches.push(
			...this.slicerCaches.map((entry) => {
				const clonedItems = entry.items?.map((item) => ({ ...item }))
				return {
					...entry,
					pivotTableNames: [...entry.pivotTableNames],
					...(clonedItems ? { items: clonedItems } : {}),
				}
			}),
		)
		clone.slicers.push(...this.slicers.map((entry) => ({ ...entry })))
		clone.timelineCaches.push(
			...this.timelineCaches.map((entry) => ({
				...entry,
				pivotTableNames: [...entry.pivotTableNames],
			})),
		)
		clone.timelines.push(...this.timelines.map((entry) => ({ ...entry })))
		clone.chartParts.push(
			...this.chartParts.map((entry) => ({
				...entry,
				series: entry.series.map((series) => ({ ...series })),
			})),
		)
		clone.chartSheets.push(
			...this.chartSheets.map((entry) => ({
				...entry,
				chartPartPaths: [...entry.chartPartPaths],
			})),
		)
		clone.macroSheets.push(...this.macroSheets.map((entry) => ({ ...entry })))
		clone.connectionParts.push(...this.connectionParts.map((entry) => ({ ...entry })))
		clone.dataModelParts.push(...this.dataModelParts.map((entry) => ({ ...entry })))
		clone.activeContent.push(...this.activeContent.map(cloneActiveContentInfo))
		clone.workbookViews.push(...this.workbookViews.map((view) => ({ ...view })))
		clone.externalReferences.push(...this.externalReferences)
		clone.externalReferenceDetails.push(
			...this.externalReferenceDetails.map((entry) => ({ ...entry })),
		)
		return clone
	}
}

export function createWorkbook(id?: WorkbookId): Workbook {
	return new Workbook(id)
}

function clonePivotCacheInfo(entry: PivotCacheInfo): PivotCacheInfo {
	return {
		...entry,
		fields: entry.fields.map((field) => ({ ...field })),
	}
}

function clonePivotTableInfo(entry: PivotTableInfo): PivotTableInfo {
	return {
		...entry,
		fields: entry.fields.map((field) => ({ ...field })),
		rowFields: entry.rowFields.map((field) => ({ ...field })),
		columnFields: entry.columnFields.map((field) => ({ ...field })),
		pageFields: entry.pageFields.map((field) => ({ ...field })),
		dataFields: entry.dataFields.map((field) => ({ ...field })),
	}
}
