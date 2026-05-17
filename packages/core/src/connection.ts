export type WorkbookConnectionPartKind = 'queryTable' | 'connection' | 'powerQueryMashup'

export interface WorkbookConnectionPartInfo {
	readonly kind: WorkbookConnectionPartKind
	readonly partPath: string
	readonly contentType: string
	readonly relType?: string
	readonly sheetName?: string
	readonly relationshipCount: number
	readonly name?: string
	readonly connectionId?: number
	readonly connectionType?: number
	readonly description?: string
	readonly deleted?: boolean
	readonly newConnection?: boolean
	readonly backgroundRefresh?: boolean
	readonly firstBackgroundRefresh?: boolean
	readonly keepAlive?: boolean
	readonly refreshInterval?: number
	readonly refreshOnLoad?: boolean
	readonly reconnectionMethod?: number
	readonly saveData?: boolean
	readonly preserveFormatting?: boolean
	readonly adjustColumnWidth?: boolean
	readonly fillFormulas?: boolean
	readonly disableEdit?: boolean
	readonly disableRefresh?: boolean
	readonly headers?: boolean
	readonly rowNumbers?: boolean
	readonly autoFormatId?: number
	readonly applyNumberFormats?: boolean
	readonly applyBorderFormats?: boolean
	readonly applyFontFormats?: boolean
	readonly applyPatternFormats?: boolean
	readonly applyAlignmentFormats?: boolean
	readonly applyWidthHeightFormats?: boolean
	readonly queryTableRefreshNextId?: number
	readonly queryTableFields?: readonly QueryTableFieldInfo[]
	readonly savePassword?: boolean
	readonly refreshedVersion?: number
	readonly refreshedDateIso?: string
	readonly minRefreshableVersion?: number
	readonly credentials?: string
	readonly singleSignOnId?: string
	readonly sourceFile?: string
	readonly odcFile?: string
	readonly onlyUseConnectionFile?: boolean
	readonly command?: string
	readonly commandType?: number
	readonly serverCommand?: boolean
	readonly webUrl?: string
	readonly webHtmlTables?: boolean
	readonly webXml?: boolean
	readonly webSourceData?: boolean
	readonly hasConnectionString?: boolean
}

export interface QueryTableFieldInfo {
	readonly id: number
	readonly name?: string
	readonly tableColumnId?: number
}
