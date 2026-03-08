import type { RangeRef } from '@ascend/core'
import type { CellChange, SheetDiff, WorkbookDiff } from '@ascend/engine'
import type { AscendError, CellValue } from '@ascend/schema'

export interface WorkbookInfo {
	readonly sheetCount: number
	readonly sheets: readonly SheetInfo[]
	readonly definedNames: readonly string[]
	readonly cellCount: number
	readonly sourceFormat: string
}

export interface SheetInfo {
	readonly name: string
	readonly rowCount: number
	readonly colCount: number
	readonly cellCount: number
	readonly tableCount: number
	readonly hasFrozenPanes: boolean
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
