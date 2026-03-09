import type { RangeRef } from '@ascend/core'
import type { CellChange, SheetDiff, WorkbookDiff } from '@ascend/engine'
import type { FormulaNode, Token } from '@ascend/formulas'
import type { AscendError, CellValue, CompatibilityReport } from '@ascend/schema'

export interface WorkbookInfo {
	readonly sheetCount: number
	readonly loadedSheetCount: number
	readonly sheets: readonly SheetInfo[]
	readonly definedNames: readonly string[]
	readonly cellCount: number | null
	readonly commentCount: number | null
	readonly conditionalFormatCount: number | null
	readonly dataValidationCount: number | null
	readonly imageCount: number | null
	readonly pivotTableCount: number
	readonly pivotCacheCount: number
	readonly slicerCount: number
	readonly slicerCacheCount: number
	readonly sourceFormat: string
	readonly workbookViewCount: number
	readonly externalReferenceCount: number
	readonly hasWorkbookProtection: boolean
	readonly pivotTables: readonly {
		readonly partPath: string
		readonly sheetName: string
		readonly name?: string
		readonly cacheId?: number
		readonly locationRef?: string
	}[]
	readonly pivotCaches: readonly {
		readonly partPath: string
		readonly cacheId?: number
		readonly relId?: string
		readonly recordCount?: number
		readonly sourceSheet?: string
		readonly sourceRef?: string
		readonly recordsPartPath?: string
	}[]
	readonly slicerCaches: readonly {
		readonly partPath: string
		readonly name?: string
		readonly sourceName?: string
		readonly pivotCacheId?: number
		readonly pivotTableNames: readonly string[]
	}[]
	readonly slicers: readonly {
		readonly partPath: string
		readonly name?: string
		readonly cacheName?: string
		readonly caption?: string
	}[]
	readonly styleSummary: {
		readonly numFmtCount: number
		readonly fontCount: number
		readonly fillCount: number
		readonly borderCount: number
		readonly cellXfCount: number
		readonly dxfCount: number
		readonly tableStyleCount: number
	}
	readonly themeSummary: {
		readonly hasThemePart: boolean
		readonly name?: string
		readonly colorSchemeName?: string
		readonly colorCount: number
		readonly majorFontLatin?: string
		readonly minorFontLatin?: string
	}
	readonly compatibility: CompatibilityReport
	readonly load: WorkbookLoadInfo
}

export interface SheetInfo {
	readonly name: string
	readonly rowCount: number | null
	readonly colCount: number | null
	readonly cellCount: number | null
	readonly tableCount: number | null
	readonly commentCount: number | null
	readonly conditionalFormatCount: number | null
	readonly dataValidationCount: number | null
	readonly hasFrozenPanes: boolean | null
	readonly colWidthCount: number | null
	readonly imageCount: number | null
	readonly rowHeightCount: number | null
	readonly hyperlinkCount: number | null
	readonly ignoredErrorCount: number | null
	readonly hasAutoFilter: boolean | null
	readonly hasDrawingRefs: boolean | null
	readonly hasPageMetadata: boolean | null
	readonly hasProtection: boolean | null
	readonly cellDataLoaded: boolean
}

export interface WorkbookLoadInfo {
	readonly mode: 'full' | 'metadata-only' | 'selective'
	readonly isPartial: boolean
	readonly cellsHydrated: boolean
	readonly hasAllSheets: boolean
	readonly sourceSheets: readonly string[]
	readonly loadedSheets: readonly string[]
}

export interface DefinedNameInfo {
	readonly name: string
	readonly formula: string
	readonly scope: 'workbook' | 'sheet'
	readonly sheet?: string
}

export interface CellInfo {
	readonly ref: string
	readonly value: CellValue
	readonly formula: string | null
	readonly row: number
	readonly col: number
}

export interface RangeInfo {
	readonly ref: RangeRef
	readonly cells: readonly CellInfo[]
	readonly rowCount: number
	readonly colCount: number
}

export interface RangeWindowInfo extends RangeInfo {
	readonly requestedRef: RangeRef
	readonly rowOffset: number
	readonly rowLimit: number
	readonly hasMore: boolean
	readonly nextRowOffset?: number
}

export interface PreviewResult {
	readonly diff: WorkbookDiff
	readonly sheetDiffs: readonly SheetDiff[]
	readonly cellChanges: readonly CellChange[]
	readonly errors: readonly AscendError[]
}

export interface ApplyResult {
	readonly affectedCells: readonly string[]
	readonly sheetsModified: readonly string[]
	readonly recalcRequired: boolean
	readonly errors: readonly AscendError[]
}

export interface RecalcResult {
	readonly changed: readonly string[]
	readonly errors: ReadonlyArray<{ ref: string; error: AscendError }>
	readonly duration: number
}

export interface CheckResult {
	readonly valid: boolean
	readonly issues: readonly CheckIssue[]
}

export interface CheckIssue {
	readonly severity: 'error' | 'warning'
	readonly message: string
	readonly ref?: string
}

export interface LintResult {
	readonly clean: boolean
	readonly warnings: readonly LintWarning[]
}

export interface LintWarning {
	readonly rule: string
	readonly message: string
	readonly ref?: string
}

export interface TraceResult {
	readonly ref: string
	readonly formula: string | null
	readonly dependsOn: readonly string[]
	readonly feedsInto: readonly string[]
}

export interface FormulaInfo {
	readonly ref: string
	readonly formula: string
	readonly normalizedFormula: string
	readonly value: CellValue
	readonly refs: readonly string[]
	readonly functions: readonly string[]
	readonly volatile: boolean
	readonly tokens: readonly Token[]
	readonly ast?: FormulaNode
	readonly parseError?: string
}
