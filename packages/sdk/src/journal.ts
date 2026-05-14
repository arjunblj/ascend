import {
	type AutoFilter,
	type Cell,
	type CellStyle,
	type ChartPartInfo,
	type ChartSeriesInfo,
	cloneCellStyle,
	clonePivotCacheInfo,
	DEFAULT_STYLE_ID,
	type DefinedName,
	type DefinedNameScope,
	indexToColumn,
	type PivotCacheInfo,
	parseA1,
	parseRange,
	type RangeRef,
	type Sheet,
	type SheetColDef,
	type SheetComment,
	type SheetConditionalFormat,
	type SheetConditionalFormatRule,
	type SheetConditionalFormatValueObject,
	type SheetDataValidation,
	type SheetDrawingObjectRef,
	type SheetFormatPr,
	type SheetHyperlink,
	type SheetImageAnchor,
	type SheetOutlinePr,
	type SheetPageMargins,
	type SheetPageSetup,
	type SheetProtection,
	type SheetRowDef,
	type SheetState,
	type SheetTabColor,
	type SheetThreadedComment,
	type Table,
	type TableColumn,
	type TableStyleInfo,
	toA1,
	toRangeString,
	type Workbook,
	type WorkbookDocumentProperties,
	type WorkbookProperties,
	type WorkbookProtection,
	type WorkbookThemeColor,
	type WorkbookThemeMetadata,
	type WorkbookView,
} from '@ascend/core'
import { applyOperation } from '@ascend/engine'
import { cachedParseFormula, extractRefs, type FormulaRef } from '@ascend/formulas'
import type {
	CalcSettings,
	CellValue,
	ConditionalFormatRule,
	DataValidationRule,
	InputValue,
	Operation,
	ScalarCellValue,
	StyleInput,
} from '@ascend/schema'
import { EMPTY } from '@ascend/schema'

export interface MutationJournalIssue {
	readonly code: 'UNSUPPORTED_OPERATION' | 'LOSSY_INVERSE' | 'UNSUPPORTED_VALUE'
	readonly message: string
	readonly refs?: readonly string[]
}

export interface MutationJournalCellPreimage {
	readonly sheet: string
	readonly ref: string
	readonly existed: boolean
	readonly value: CellValue
	readonly formula: string | null
	readonly formulaInfo?: Cell['formulaInfo']
	readonly styleId: number
	readonly style: CellStyle
}

export interface MutationJournalCommentPreimage {
	readonly sheet: string
	readonly ref: string
	readonly comment: SheetComment | null
}

export interface MutationJournalHyperlinkPreimage {
	readonly sheet: string
	readonly ref: string
	readonly hyperlink: SheetHyperlink | null
}

export interface MutationJournalThreadedCommentPreimage {
	readonly sheet: string
	readonly commentIndex: number | null
	readonly threadedComment: SheetThreadedComment | null
}

export interface MutationJournalDrawingTextPreimage {
	readonly sheet: string
	readonly drawingObjectIndex: number | null
	readonly drawingObject: SheetDrawingObjectRef | null
}

export interface MutationJournalChartSeriesPreimage {
	readonly chartIndex: number | null
	readonly seriesIndex: number
	readonly chart: ChartPartInfo | null
	readonly series: ChartSeriesInfo | null
}

export interface MutationJournalPivotCachePreimage {
	readonly cacheIndex: number | null
	readonly cache: PivotCacheInfo | null
}

export interface MutationJournalPanePreimage {
	readonly sheet: string
	readonly frozenRows: number
	readonly frozenCols: number
}

export interface MutationJournalMergePreimage {
	readonly sheet: string
	readonly range: string
	readonly existed: boolean
}

export interface MutationJournalAutoFilterPreimage {
	readonly sheet: string
	readonly autoFilter: AutoFilter | null
}

export interface MutationJournalDataValidationPreimage {
	readonly sheet: string
	readonly range: string
	readonly validation: SheetDataValidation | null
}

export interface MutationJournalConditionalFormatPreimage {
	readonly sheet: string
	readonly range?: string
	readonly formats: readonly SheetConditionalFormat[]
}

export interface MutationJournalDefinedNamePreimage {
	readonly name: string
	readonly scope?: string
	readonly definedName: DefinedName | null
}

export interface MutationJournalPageSetupPreimage {
	readonly sheet: string
	readonly pageSetup: SheetPageSetup | null
	readonly pageMargins: SheetPageMargins | null
}

export interface MutationJournalTableRenamePreimage {
	readonly table: string
	readonly existed: boolean
}

export interface MutationJournalTableColumnPreimage {
	readonly table: string
	readonly column: string | number
	readonly columnIndex: number
	readonly columnState: TableColumn | null
}

export interface MutationJournalTablePreimage {
	readonly table: Table | null
	readonly sheet: string | null
	readonly ref: string | null
}

export interface MutationJournalTableStylePreimage {
	readonly table: string
	readonly style: TableStyleInfo | null
}

export interface MutationJournalSheetMovePreimage {
	readonly sheet: string
	readonly position: number | null
}

export interface MutationJournalSheetDeletePreimage {
	readonly sheet: string
	readonly position: number | null
	readonly existed: boolean
}

export interface MutationJournalSheetLayoutPreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly index: number
	readonly value: number | null
}

export interface MutationJournalSheetTabColorPreimage {
	readonly sheet: string
	readonly tabColor: SheetTabColor | null
}

export interface MutationJournalSheetProtectionPreimage {
	readonly sheet: string
	readonly protection: SheetProtection | null
}

export interface MutationJournalSheetVisibilityPreimage {
	readonly sheet: string
	readonly state: SheetState | null
}

export interface MutationJournalRowsHiddenPreimage {
	readonly sheet: string
	readonly rows: readonly {
		readonly row: number
		readonly height: number | null
	}[]
}

export interface MutationJournalColsHiddenPreimage {
	readonly sheet: string
	readonly cols: readonly {
		readonly col: number
		readonly colDef: SheetColDef | null
	}[]
}

export interface MutationJournalOutlinePreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly from: number
	readonly to: number
	readonly outlinePr: SheetOutlinePr | null
	readonly sheetFormatPr: SheetFormatPr | null
	readonly rowDefs?: readonly {
		readonly row: number
		readonly rowDef: SheetRowDef | null
	}[]
	readonly colDefs?: readonly {
		readonly col: number
		readonly colDef: SheetColDef | null
	}[]
}

export interface MutationJournalStructuralPreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly at: number
	readonly count: number
	readonly deletedCells: readonly MutationJournalCellPreimage[]
}

export interface MutationJournalWorkbookPropertiesPreimage {
	readonly properties: WorkbookProperties
}

export interface MutationJournalDocumentPropertiesPreimage {
	readonly properties: WorkbookDocumentProperties
}

export interface MutationJournalWorkbookViewPreimage {
	readonly index: number
	readonly view: WorkbookView | null
}

export interface MutationJournalCalcSettingsPreimage {
	readonly settings: CalcSettings
	readonly workbookProperties: WorkbookProperties
}

export interface MutationJournalWorkbookProtectionPreimage {
	readonly protection: WorkbookProtection | null
}

export interface MutationJournalThemePreimage {
	readonly metadata: WorkbookThemeMetadata
	readonly colors: readonly WorkbookThemeColor[]
}

export type MutationJournalPreimage =
	| { readonly kind: 'cells'; readonly cells: readonly MutationJournalCellPreimage[] }
	| { readonly kind: 'comment'; readonly comment: MutationJournalCommentPreimage }
	| { readonly kind: 'hyperlink'; readonly hyperlink: MutationJournalHyperlinkPreimage }
	| {
			readonly kind: 'threaded-comment'
			readonly threadedComment: MutationJournalThreadedCommentPreimage
	  }
	| { readonly kind: 'drawing-text'; readonly drawingText: MutationJournalDrawingTextPreimage }
	| { readonly kind: 'chart-series'; readonly chartSeries: MutationJournalChartSeriesPreimage }
	| { readonly kind: 'pivot-cache'; readonly pivotCache: MutationJournalPivotCachePreimage }
	| { readonly kind: 'pane'; readonly pane: MutationJournalPanePreimage }
	| { readonly kind: 'merge'; readonly merge: MutationJournalMergePreimage }
	| { readonly kind: 'auto-filter'; readonly autoFilter: MutationJournalAutoFilterPreimage }
	| {
			readonly kind: 'data-validations'
			readonly validations: readonly MutationJournalDataValidationPreimage[]
	  }
	| {
			readonly kind: 'conditional-formats'
			readonly conditionalFormats: MutationJournalConditionalFormatPreimage
	  }
	| { readonly kind: 'page-setup'; readonly pageSetup: MutationJournalPageSetupPreimage }
	| { readonly kind: 'defined-name'; readonly definedName: MutationJournalDefinedNamePreimage }
	| { readonly kind: 'table-rename'; readonly tableRename: MutationJournalTableRenamePreimage }
	| { readonly kind: 'table-column'; readonly tableColumn: MutationJournalTableColumnPreimage }
	| { readonly kind: 'table'; readonly table: MutationJournalTablePreimage }
	| { readonly kind: 'table-style'; readonly tableStyle: MutationJournalTableStylePreimage }
	| { readonly kind: 'sheet-move'; readonly sheetMove: MutationJournalSheetMovePreimage }
	| { readonly kind: 'sheet-delete'; readonly sheetDelete: MutationJournalSheetDeletePreimage }
	| { readonly kind: 'sheet-layout'; readonly sheetLayout: MutationJournalSheetLayoutPreimage }
	| {
			readonly kind: 'sheet-tab-color'
			readonly sheetTabColor: MutationJournalSheetTabColorPreimage
	  }
	| {
			readonly kind: 'sheet-protection'
			readonly sheetProtection: MutationJournalSheetProtectionPreimage
	  }
	| {
			readonly kind: 'sheet-visibility'
			readonly sheetVisibility: MutationJournalSheetVisibilityPreimage
	  }
	| { readonly kind: 'rows-hidden'; readonly rowsHidden: MutationJournalRowsHiddenPreimage }
	| { readonly kind: 'cols-hidden'; readonly colsHidden: MutationJournalColsHiddenPreimage }
	| { readonly kind: 'outline'; readonly outline: MutationJournalOutlinePreimage }
	| { readonly kind: 'structural'; readonly structural: MutationJournalStructuralPreimage }
	| {
			readonly kind: 'workbook-properties'
			readonly workbookProperties: MutationJournalWorkbookPropertiesPreimage
	  }
	| {
			readonly kind: 'document-properties'
			readonly documentProperties: MutationJournalDocumentPropertiesPreimage
	  }
	| { readonly kind: 'workbook-view'; readonly workbookView: MutationJournalWorkbookViewPreimage }
	| { readonly kind: 'calc-settings'; readonly calcSettings: MutationJournalCalcSettingsPreimage }
	| {
			readonly kind: 'workbook-protection'
			readonly workbookProtection: MutationJournalWorkbookProtectionPreimage
	  }
	| { readonly kind: 'theme'; readonly theme: MutationJournalThemePreimage }

export interface MutationJournalEntry {
	readonly opIndex: number
	readonly op: Operation
	readonly supported: boolean
	readonly exact: boolean
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
}

export interface MutationJournal {
	readonly entries: readonly MutationJournalEntry[]
	readonly inverseOps: readonly Operation[]
	readonly supported: boolean
	readonly exact: boolean
	readonly issues: readonly MutationJournalIssue[]
}

interface DraftJournalEntry {
	readonly opIndex: number
	readonly op: Operation
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
}

export function buildMutationJournal(
	workbook: Workbook,
	ops: readonly Operation[],
): MutationJournal {
	if (ops.length === 1) {
		const op = ops[0]
		if (!op) return emptyMutationJournal()
		return mutationJournalFromEntries([buildJournalEntry(workbook, op, 0)])
	}
	const journalWorkbook = workbook.clone()
	const entries: MutationJournalEntry[] = []
	for (let opIndex = 0; opIndex < ops.length; opIndex++) {
		const op = ops[opIndex]
		if (!op) continue
		entries.push(buildJournalEntry(journalWorkbook, op, opIndex))
		const result = applyOperation(journalWorkbook, op)
		if (!result.ok) break
	}
	return mutationJournalFromEntries(entries)
}

function emptyMutationJournal(): MutationJournal {
	return {
		entries: [],
		inverseOps: [],
		supported: true,
		exact: true,
		issues: [],
	}
}

function mutationJournalFromEntries(entries: readonly MutationJournalEntry[]): MutationJournal {
	const inverseOps = [...entries].reverse().flatMap((entry) => entry.inverseOps)
	const issues = entries.flatMap((entry) => entry.issues)
	return {
		entries,
		inverseOps,
		supported: entries.every((entry) => entry.supported),
		exact: entries.every((entry) => entry.exact),
		issues,
	}
}

function buildJournalEntry(
	workbook: Workbook,
	op: Operation,
	opIndex: number,
): MutationJournalEntry {
	const draft = buildSupportedJournalEntry(workbook, op, opIndex)
	if (!draft) {
		return {
			opIndex,
			op,
			supported: false,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues: [
				{
					code: 'UNSUPPORTED_OPERATION',
					message: `No reversible journal support for ${op.op}`,
				},
			],
		}
	}
	return {
		...draft,
		supported: draft.issues.every((issue) => issue.code !== 'UNSUPPORTED_OPERATION'),
		exact: draft.issues.length === 0,
	}
}

function buildSupportedJournalEntry(
	workbook: Workbook,
	op: Operation,
	opIndex: number,
): DraftJournalEntry | null {
	switch (op.op) {
		case 'setCells':
			return journalSetCells(workbook, op, opIndex)
		case 'setFormula':
			return journalSetFormula(workbook, op, opIndex)
		case 'fillFormula':
			return journalFillFormula(workbook, op, opIndex)
		case 'setRichText':
			return journalSetRichText(workbook, op, opIndex)
		case 'clearRange':
			return journalClearRange(workbook, op, opIndex)
		case 'insertRows':
			return journalInsertAxis(op, opIndex, 'row')
		case 'insertCols':
			return journalInsertAxis(op, opIndex, 'col')
		case 'deleteRows':
			return journalDeleteAxis(workbook, op, opIndex, 'row')
		case 'deleteCols':
			return journalDeleteAxis(workbook, op, opIndex, 'col')
		case 'setNumberFormat':
		case 'setStyle':
			return journalStyleRange(workbook, op, opIndex)
		case 'mergeCells':
			return journalMergeCells(workbook, op, opIndex)
		case 'unmergeCells':
			return journalUnmergeCells(workbook, op, opIndex)
		case 'setDataValidation':
			return journalSetDataValidation(workbook, op, opIndex)
		case 'deleteDataValidation':
			return journalDeleteDataValidation(workbook, op, opIndex)
		case 'setAutoFilter':
			return journalSetAutoFilter(workbook, op, opIndex)
		case 'clearAutoFilter':
			return journalClearAutoFilter(workbook, op, opIndex)
		case 'setConditionalFormat':
			return journalSetConditionalFormat(workbook, op, opIndex)
		case 'deleteConditionalFormat':
			return journalDeleteConditionalFormat(workbook, op, opIndex)
		case 'setPageSetup':
			return journalSetPageSetup(workbook, op, opIndex)
		case 'setPrintArea':
			return journalSetPrintArea(workbook, op, opIndex)
		case 'setDefinedName':
			return journalSetDefinedName(workbook, op, opIndex)
		case 'deleteDefinedName':
			return journalDeleteDefinedName(workbook, op, opIndex)
		case 'renameTable':
			return journalRenameTable(workbook, op, opIndex)
		case 'createTable':
			return journalCreateTable(op, opIndex)
		case 'deleteTable':
			return journalDeleteTable(workbook, op, opIndex)
		case 'resizeTable':
			return journalResizeTable(workbook, op, opIndex)
		case 'setTableColumn':
			return journalSetTableColumn(workbook, op, opIndex)
		case 'setTableStyle':
			return journalSetTableStyle(workbook, op, opIndex)
		case 'setComment':
			return journalSetComment(workbook, op, opIndex)
		case 'deleteComment':
			return journalDeleteComment(workbook, op, opIndex)
		case 'setHyperlink':
			return journalSetHyperlink(workbook, op, opIndex)
		case 'deleteHyperlink':
			return journalDeleteHyperlink(workbook, op, opIndex)
		case 'setThreadedComment':
			return journalSetThreadedComment(workbook, op, opIndex)
		case 'setDrawingText':
			return journalSetDrawingText(workbook, op, opIndex)
		case 'setChartSeriesSource':
			return journalSetChartSeriesSource(workbook, op, opIndex)
		case 'setPivotCache':
			return journalSetPivotCache(workbook, op, opIndex)
		case 'freezePane':
			return journalFreezePane(workbook, op, opIndex)
		case 'setWorkbookProperties':
			return journalSetWorkbookProperties(workbook, op, opIndex)
		case 'setDocumentProperties':
			return journalSetDocumentProperties(workbook, op, opIndex)
		case 'setWorkbookView':
			return journalSetWorkbookView(workbook, op, opIndex)
		case 'setCalcSettings':
			return journalSetCalcSettings(workbook, op, opIndex)
		case 'setWorkbookProtection':
			return journalSetWorkbookProtection(workbook, op, opIndex)
		case 'setTheme':
			return journalSetTheme(workbook, op, opIndex)
		case 'renameSheet':
			return journalRenameSheet(workbook, op, opIndex)
		case 'moveSheet':
			return journalMoveSheet(workbook, op, opIndex)
		case 'addSheet':
			return {
				opIndex,
				op,
				inverseOps: [{ op: 'deleteSheet', sheet: op.name }],
				preimages: [],
				issues: [],
			}
		case 'deleteSheet':
			return journalDeleteSheet(workbook, op, opIndex)
		case 'copySheet':
			return {
				opIndex,
				op,
				inverseOps: [{ op: 'deleteSheet', sheet: op.newName }],
				preimages: [],
				issues: [],
			}
		case 'setRowHeight':
			return journalSetSheetLayout(workbook, op, opIndex, 'row')
		case 'setColWidth':
			return journalSetSheetLayout(workbook, op, opIndex, 'col')
		case 'setTabColor':
			return journalSetTabColor(workbook, op, opIndex)
		case 'setSheetProtection':
			return journalSetSheetProtection(workbook, op, opIndex)
		case 'hideSheet':
			return journalHideSheet(workbook, op, opIndex)
		case 'hideRows':
			return journalHideRows(workbook, op, opIndex)
		case 'hideCols':
			return journalHideCols(workbook, op, opIndex)
		case 'groupRows':
			return journalGroupOutline(workbook, op, opIndex, 'row')
		case 'groupCols':
			return journalGroupOutline(workbook, op, opIndex, 'col')
		default:
			return null
	}
}

function journalSetTabColor(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTabColor' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const tabColor = sheet?.tabColor ? cloneSheetTabColor(sheet.tabColor) : null
	const preimage: MutationJournalSheetTabColorPreimage = {
		sheet: op.sheet,
		tabColor,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-tab-color', sheetTabColor: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore tab color for ${op.sheet} because the sheet was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const inverseOps: Operation[] = tabColor?.rgb
		? [{ op: 'setTabColor', sheet: op.sheet, color: tabColor.rgb }]
		: []
	const issues = tabColorIsPublicRgb(tabColor)
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: tabColor
						? `Sheet tab color for ${op.sheet} uses unsupported color metadata and cannot be fully restored with public operations`
						: `Sheet tab color absence for ${op.sheet} cannot be restored with public operations`,
					refs: [`sheet:${op.sheet}:tabColor`],
				} satisfies MutationJournalIssue,
			]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'sheet-tab-color', sheetTabColor: preimage }],
		issues,
	}
}

function journalSetSheetProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setSheetProtection' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const protection = sheet?.protection ? cloneSheetProtection(sheet.protection) : null
	const preimage: MutationJournalSheetProtectionPreimage = {
		sheet: op.sheet,
		protection,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-protection', sheetProtection: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore sheet protection for ${op.sheet} because the sheet was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const inverseOp = protection ? sheetProtectionInverseOp(op.sheet, protection) : null
	const unsupportedKeys = protection ? unsupportedSheetProtectionKeys(protection) : []
	const issues: MutationJournalIssue[] = []
	if (!protection) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sheet protection absence for ${op.sheet} cannot be restored with public operations`,
			refs: [`sheet:${op.sheet}:protection`],
		})
	} else if (protection.sheet !== true || unsupportedKeys.length > 0) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sheet protection for ${op.sheet} contains metadata that cannot be fully restored with public operations`,
			refs: unsupportedKeys.map((key) => `sheet:${op.sheet}:protection:${key}`),
		})
	}
	return {
		opIndex,
		op,
		inverseOps: inverseOp ? [inverseOp] : [],
		preimages: [{ kind: 'sheet-protection', sheetProtection: preimage }],
		issues,
	}
}

function journalDeleteSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const position = workbook.sheets.findIndex((sheet) => sheet.name === op.sheet)
	const sheet = position >= 0 ? workbook.sheets[position] : undefined
	const preimage: MutationJournalSheetDeletePreimage = {
		sheet: op.sheet,
		position: position >= 0 ? position : null,
		existed: sheet !== undefined,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-delete', sheetDelete: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore deleted sheet ${op.sheet} because it was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const refs = sheetDeleteLossRefs(workbook, sheet)
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'addSheet', name: op.sheet, position }],
		preimages: [{ kind: 'sheet-delete', sheetDelete: preimage }],
		issues:
			refs.length === 0
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message: `Deleted sheet ${op.sheet} cannot be fully restored with public operations`,
							refs,
						},
					],
	}
}

function journalMoveSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'moveSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const position = workbook.sheets.findIndex((sheet) => sheet.name === op.sheet)
	const preimage: MutationJournalSheetMovePreimage = {
		sheet: op.sheet,
		position: position >= 0 ? position : null,
	}
	const issues: MutationJournalIssue[] =
		position >= 0
			? []
			: [
					{
						code: 'UNSUPPORTED_VALUE',
						message: `Cannot restore sheet move for ${op.sheet} because the sheet was not found`,
					},
				]
	return {
		opIndex,
		op,
		inverseOps: position >= 0 ? [{ op: 'moveSheet', sheet: op.sheet, position }] : [],
		preimages: [{ kind: 'sheet-move', sheetMove: preimage }],
		issues,
	}
}

function journalSetSheetLayout(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRowHeight' | 'setColWidth' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const index =
		axis === 'row' && op.op === 'setRowHeight' ? op.row : op.op === 'setColWidth' ? op.col : -1
	const value = sheet
		? axis === 'row'
			? sheet.rowHeights.get(index)
			: sheet.colWidths.get(index)
		: undefined
	const preimage: MutationJournalSheetLayoutPreimage = {
		sheet: op.sheet,
		axis,
		index,
		value: value ?? null,
	}
	const ref = layoutRef(op.sheet, axis, index)
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-layout', sheetLayout: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore ${axis} layout at ${ref} because the sheet was not found`,
					refs: [ref],
				},
			],
		}
	}
	if (value === undefined) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-layout', sheetLayout: preimage }],
			issues: [
				{
					code: 'LOSSY_INVERSE',
					message: `Created ${axis} layout at ${ref} cannot be cleared with public operations`,
					refs: [ref],
				},
			],
		}
	}
	const inverseOps: Operation[] =
		axis === 'row'
			? [{ op: 'setRowHeight', sheet: op.sheet, row: index, height: value }]
			: [{ op: 'setColWidth', sheet: op.sheet, col: index, width: value }]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'sheet-layout', sheetLayout: preimage }],
		issues: [],
	}
}

function journalHideSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const preimage: MutationJournalSheetVisibilityPreimage = {
		sheet: op.sheet,
		state: sheet?.state ?? null,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-visibility', sheetVisibility: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore sheet visibility for ${op.sheet} because the sheet was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const issues: MutationJournalIssue[] =
		sheet.state === 'veryHidden'
			? [
					{
						code: 'LOSSY_INVERSE',
						message: `Sheet visibility for ${op.sheet} was veryHidden and cannot be restored with public operations`,
						refs: [`sheet:${op.sheet}:state:veryHidden`],
					},
				]
			: []
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'hideSheet', sheet: op.sheet, hidden: sheet.state !== 'visible' }],
		preimages: [{ kind: 'sheet-visibility', sheetVisibility: preimage }],
		issues,
	}
}

function journalHideRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideRows' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const rows = Array.from({ length: op.count }, (_, offset) => {
		const row = op.at + offset
		const height = sheet?.rowHeights.get(row)
		return { row, height: height ?? null }
	})
	const preimage: MutationJournalRowsHiddenPreimage = { sheet: op.sheet, rows }
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'rows-hidden', rowsHidden: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore row visibility for ${op.sheet} because the sheet was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	if (op.hidden === false) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'rows-hidden', rowsHidden: preimage }],
			issues: [],
		}
	}
	const inverseOps = rows.flatMap((row): Operation[] =>
		row.height === null
			? []
			: [{ op: 'setRowHeight', sheet: op.sheet, row: row.row, height: row.height }],
	)
	const refs = rows
		.filter((row) => row.height === null)
		.map((row) => layoutRef(op.sheet, 'row', row.row))
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'rows-hidden', rowsHidden: preimage }],
		issues:
			refs.length === 0
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message: `Created row hide metadata cannot be cleared with public operations`,
							refs,
						},
					],
	}
}

function journalHideCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideCols' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const cols = Array.from({ length: op.count }, (_, offset) => {
		const col = op.at + offset
		const colDef = sheet?.colDefs.find((def) => def.min === col + 1 && def.max === col + 1)
		return { col, colDef: colDef ? cloneSheetColDef(colDef) : null }
	})
	const preimage: MutationJournalColsHiddenPreimage = { sheet: op.sheet, cols }
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'cols-hidden', colsHidden: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore column visibility for ${op.sheet} because the sheet was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const inverseOps = cols.flatMap((col): Operation[] =>
		col.colDef?.hidden === undefined
			? []
			: [{ op: 'hideCols', sheet: op.sheet, at: col.col, count: 1, hidden: col.colDef.hidden }],
	)
	const refs = cols
		.filter((col) => col.colDef?.hidden === undefined)
		.map((col) => layoutRef(op.sheet, 'col', col.col))
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cols-hidden', colsHidden: preimage }],
		issues:
			refs.length === 0
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message: `Created or unkeyed column hide metadata cannot be cleared with public operations`,
							refs,
						},
					],
	}
}

function journalGroupOutline(
	workbook: Workbook,
	op: Extract<Operation, { op: 'groupRows' | 'groupCols' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const preimage: MutationJournalOutlinePreimage = {
		sheet: op.sheet,
		axis,
		from: op.from,
		to: op.to,
		outlinePr: sheet?.outlinePr ? cloneSheetOutlinePr(sheet.outlinePr) : null,
		sheetFormatPr: sheet?.sheetFormatPr ? cloneSheetFormatPr(sheet.sheetFormatPr) : null,
		...(axis === 'row'
			? {
					rowDefs: outlineIndexes(op).map((row) => {
						const rowDef = sheet?.rowDefs.get(row)
						return {
							row,
							rowDef: rowDef ? cloneSheetRowDef(rowDef) : null,
						}
					}),
				}
			: {
					colDefs: outlineIndexes(op).map((col) => ({
						col,
						colDef: cloneMatchingSheetColDef(sheet, col),
					})),
				}),
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'outline', outline: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore ${axis} outline grouping for ${op.sheet} because the sheet was not found`,
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const outlineRef = axis === 'row' ? 'summaryBelow' : 'summaryRight'
	const formatRef = axis === 'row' ? 'outlineLevelRow' : 'outlineLevelCol'
	const refs = [
		...outlineIndexes(op).map((index) => layoutRef(op.sheet, axis, index)),
		`sheet:${op.sheet}:outlinePr:${outlineRef}`,
		`sheet:${op.sheet}:sheetFormatPr:${formatRef}`,
	]
	return {
		opIndex,
		op,
		inverseOps: [],
		preimages: [{ kind: 'outline', outline: preimage }],
		issues: [
			{
				code: 'LOSSY_INVERSE',
				message: `Grouped ${axis === 'row' ? 'rows' : 'columns'} for ${op.sheet} cannot be restored with public operations`,
				refs: [...new Set(refs)],
			},
		],
	}
}

function cloneSheetColDef(def: SheetColDef): SheetColDef {
	return { ...def }
}

function cloneMatchingSheetColDef(sheet: Sheet | undefined, col: number): SheetColDef | null {
	const colDef = sheet?.colDefs.find((def) => def.min === col && def.max === col)
	return colDef ? cloneSheetColDef(colDef) : null
}

function cloneSheetRowDef(def: SheetRowDef): SheetRowDef {
	return { ...def }
}

function cloneSheetOutlinePr(outlinePr: SheetOutlinePr): SheetOutlinePr {
	return { ...outlinePr }
}

function cloneSheetFormatPr(sheetFormatPr: SheetFormatPr): SheetFormatPr {
	return { ...sheetFormatPr }
}

function cloneSheetTabColor(tabColor: SheetTabColor): SheetTabColor {
	return { ...tabColor }
}

function cloneSheetProtection(protection: SheetProtection): SheetProtection {
	return { ...protection }
}

function tabColorIsPublicRgb(tabColor: SheetTabColor | null): boolean {
	if (!tabColor?.rgb) return false
	return (
		tabColor.theme === undefined && tabColor.tint === undefined && tabColor.indexed === undefined
	)
}

function sheetProtectionInverseOp(
	sheet: string,
	protection: SheetProtection,
): Extract<Operation, { op: 'setSheetProtection' }> {
	type SheetProtectionInverseOptions = {
		-readonly [Key in keyof NonNullable<
			Extract<Operation, { op: 'setSheetProtection' }>['options']
		>]?: boolean
	}
	const options: SheetProtectionInverseOptions = {}
	for (const key of SHEET_PROTECTION_OPTION_KEYS) {
		const value = protection[key]
		if (value !== undefined) options[key] = value
	}
	return {
		op: 'setSheetProtection',
		sheet,
		...(protection.password ? { password: protection.password } : {}),
		...(Object.keys(options).length > 0 ? { options } : {}),
	}
}

const SHEET_PROTECTION_OPTION_KEYS = [
	'formatCells',
	'formatColumns',
	'formatRows',
	'insertColumns',
	'insertRows',
	'deleteColumns',
	'deleteRows',
	'sort',
	'autoFilter',
] as const satisfies readonly (keyof SheetProtection)[]

function unsupportedSheetProtectionKeys(protection: SheetProtection): string[] {
	const supported = new Set<string>(['sheet', 'password', ...SHEET_PROTECTION_OPTION_KEYS])
	return Object.keys(protection).filter((key) => !supported.has(key))
}

function outlineIndexes(op: Extract<Operation, { op: 'groupRows' | 'groupCols' }>): number[] {
	const indexes: number[] = []
	if (op.from <= op.to) {
		for (let index = op.from; index <= op.to; index++) indexes.push(index)
	}
	if (op.collapsed) {
		const summaryAfter =
			op.op === 'groupRows' ? (op.summaryBelow ?? true) : (op.summaryRight ?? true)
		const boundary = summaryAfter ? op.to + 1 : op.from - 1
		if (boundary >= 0) indexes.push(boundary)
	}
	return indexes
}

function layoutRef(sheet: string, axis: 'row' | 'col', index: number): string {
	return axis === 'row' ? `${sheet}!${index + 1}` : `${sheet}!${indexToColumn(index)}`
}

function sheetDeleteLossRefs(workbook: Workbook, sheet: Sheet): string[] {
	const refs: string[] = []
	const cellCount = sheet.cells.storageStats().cellCount
	if (cellCount > 0) refs.push(`${sheet.name}!cells:${cellCount}`)
	if (sheet.merges.length > 0) refs.push(`${sheet.name}!merges:${sheet.merges.length}`)
	if (sheet.tables.length > 0) refs.push(`${sheet.name}!tables:${sheet.tables.length}`)
	if (sheet.comments.size > 0) refs.push(`${sheet.name}!comments:${sheet.comments.size}`)
	if (sheet.threadedComments.length > 0) {
		refs.push(`${sheet.name}!threadedComments:${sheet.threadedComments.length}`)
	}
	if (sheet.hyperlinks.size > 0) refs.push(`${sheet.name}!hyperlinks:${sheet.hyperlinks.size}`)
	if (sheet.dataValidations.length > 0) {
		refs.push(`${sheet.name}!validations:${sheet.dataValidations.length}`)
	}
	if (sheet.conditionalFormats.length > 0) {
		refs.push(`${sheet.name}!conditionalFormats:${sheet.conditionalFormats.length}`)
	}
	if (sheet.autoFilter) refs.push(`${sheet.name}!autoFilter:${sheet.autoFilter.ref}`)
	if (sheet.rowHeights.size > 0) refs.push(`${sheet.name}!rowHeights:${sheet.rowHeights.size}`)
	if (sheet.colWidths.size > 0) refs.push(`${sheet.name}!colWidths:${sheet.colWidths.size}`)
	if (sheet.frozenRows > 0 || sheet.frozenCols > 0) refs.push(`${sheet.name}!pane`)
	if (sheet.state !== 'visible') refs.push(`${sheet.name}!state:${sheet.state}`)
	if (sheet.imageRefs.length > 0) refs.push(`${sheet.name}!images:${sheet.imageRefs.length}`)
	if (sheet.drawingObjectRefs.length > 0) {
		refs.push(`${sheet.name}!drawings:${sheet.drawingObjectRefs.length}`)
	}
	if (sheet.x14DataValidations.length > 0) {
		refs.push(`${sheet.name}!x14Validations:${sheet.x14DataValidations.length}`)
	}
	if (sheet.x14ConditionalFormats.length > 0) {
		refs.push(`${sheet.name}!x14ConditionalFormats:${sheet.x14ConditionalFormats.length}`)
	}
	if (sheet.advancedFilters.length > 0) {
		refs.push(`${sheet.name}!advancedFilters:${sheet.advancedFilters.length}`)
	}
	for (const entry of workbook.definedNames.list()) {
		const scopeSheet =
			entry.scope.kind === 'sheet' ? sheetNameForId(workbook, entry.scope.sheetId) : undefined
		if (scopeSheet === sheet.name) refs.push(definedNameJournalKey(workbook, entry))
		if (formulaReferencesSheet(workbook, entry.formula, scopeSheet ?? sheet.name, sheet.name)) {
			refs.push(definedNameJournalKey(workbook, entry))
		}
	}
	for (const workbookSheet of workbook.sheets) {
		if (workbookSheet.name === sheet.name) continue
		for (const [row, col, cell] of workbookSheet.cells.iterate()) {
			if (
				cell.formula &&
				formulaReferencesSheet(workbook, cell.formula, workbookSheet.name, sheet.name)
			) {
				refs.push(`${workbookSheet.name}!${toA1({ row, col })}`)
			}
		}
	}
	for (const chart of workbook.chartParts) {
		if (chart.sheetName === sheet.name) refs.push(`chart:${chart.partPath}`)
		chart.series.forEach((series, seriesIndex) => {
			for (const [field, formula] of [
				['nameRef', series.nameRef],
				['categoryRef', series.categoryRef],
				['valueRef', series.valueRef],
			] as const) {
				if (
					formula &&
					formulaReferencesSheet(workbook, formula, chart.sheetName ?? sheet.name, sheet.name)
				) {
					refs.push(`chart:${chart.partPath}:series:${seriesIndex}:${field}`)
				}
			}
		})
	}
	for (const pivot of workbook.pivotTables) {
		if (pivot.sheetName === sheet.name) refs.push(`pivotTable:${pivot.partPath}`)
	}
	for (const pivotCache of workbook.pivotCaches) {
		if (pivotCache.sourceSheet === sheet.name) refs.push(`pivotCache:${pivotCache.partPath}`)
	}
	return [...new Set(refs)]
}

function formulaReferencesSheet(
	workbook: Workbook,
	formula: string,
	ownerSheet: string,
	sheetName: string,
): boolean {
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return false
	return extractRefs(parsed.value).some((ref) =>
		formulaRefReferencesSheet(workbook, ref, ownerSheet, sheetName),
	)
}

function formulaRefReferencesSheet(
	workbook: Workbook,
	ref: FormulaRef,
	ownerSheet: string,
	sheetName: string,
): boolean {
	if (ref.kind === 'sheetSpan') {
		return (
			sheetSpanIncludes(workbook, ref.startSheet, ref.endSheet, sheetName) ||
			formulaRefReferencesSheet(workbook, ref.target, ownerSheet, sheetName)
		)
	}
	const refSheet = 'sheet' in ref && ref.sheet !== undefined ? ref.sheet : ownerSheet
	return refSheet === sheetName
}

function journalSetCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(
		workbook,
		op.sheet,
		op.updates.map((update) => update.ref),
	)
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalSetFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setFormula' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(workbook, op.sheet, [op.ref])
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalFillFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'fillFormula' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(workbook, op.sheet, refsInRange(op.range))
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalSetRichText(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRichText' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(workbook, op.sheet, [op.ref])
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalClearRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'clearRange' }>,
	opIndex: number,
): DraftJournalEntry {
	const refs = refsInRange(op.range)
	if (op.what === 'styles') {
		const cells = cellPreimages(workbook, op.sheet, refs)
		return {
			opIndex,
			op,
			inverseOps: styleInverseOps(cells),
			preimages: [{ kind: 'cells', cells }],
			issues: [],
		}
	}
	const cells = cellEditPreimages(workbook, op.sheet, refs)
	const { inverseOps: cellInverseOps, issues } = inverseCellOps(cells)
	const inverseOps =
		op.what === 'all' ? [...cellInverseOps, ...styleInverseOps(cells)] : cellInverseOps
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalStyleRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setNumberFormat' | 'setStyle' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellPreimages(workbook, op.sheet, refsInRange(op.range))
	return {
		opIndex,
		op,
		inverseOps: styleInverseOps(cells),
		preimages: [{ kind: 'cells', cells }],
		issues: [],
	}
}

function journalInsertAxis(
	op: Extract<Operation, { op: 'insertRows' | 'insertCols' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const inverseOps: Operation[] =
		axis === 'row'
			? [{ op: 'deleteRows', sheet: op.sheet, at: op.at, count: op.count }]
			: [{ op: 'deleteCols', sheet: op.sheet, at: op.at, count: op.count }]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [
			{
				kind: 'structural',
				structural: { sheet: op.sheet, axis, at: op.at, count: op.count, deletedCells: [] },
			},
		],
		issues: [],
	}
}

function journalDeleteAxis(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteRows' | 'deleteCols' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const preimage = structuralDeletePreimage(workbook, op.sheet, axis, op.at, op.count)
	const { inverseOps: cellInverseOps, issues: cellIssues } = inverseCellOps(preimage.deletedCells)
	const inverseOps: Operation[] =
		axis === 'row'
			? [{ op: 'insertRows', sheet: op.sheet, at: op.at, count: op.count }]
			: [{ op: 'insertCols', sheet: op.sheet, at: op.at, count: op.count }]
	inverseOps.push(...cellInverseOps, ...styleInverseOps(preimage.deletedCells))
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'structural', structural: preimage }],
		issues: [...cellIssues, ...structuralDeleteIssues(workbook, preimage)],
	}
}

function journalMergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'mergeCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const merge = mergePreimage(workbook, op.sheet, op.range)
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'unmergeCells', sheet: op.sheet, range: merge.range }],
		preimages: [{ kind: 'merge', merge }],
		issues: [],
	}
}

function journalUnmergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'unmergeCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const merge = mergePreimage(workbook, op.sheet, op.range)
	return {
		opIndex,
		op,
		inverseOps: merge.existed ? [{ op: 'mergeCells', sheet: op.sheet, range: merge.range }] : [],
		preimages: [{ kind: 'merge', merge }],
		issues: [],
	}
}

function journalSetDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDataValidation' }>,
	opIndex: number,
): DraftJournalEntry {
	const validation = dataValidationPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } = restoreDataValidationOps(validation)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'data-validations', validations: [validation] }],
		issues,
	}
}

function journalDeleteDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDataValidation' }>,
	opIndex: number,
): DraftJournalEntry {
	const validation = dataValidationPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } = validation.validation
		? restoreDataValidationOps(validation)
		: { inverseOps: [], issues: [] }
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'data-validations', validations: [validation] }],
		issues,
	}
}

function journalSetAutoFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setAutoFilter' }>,
	opIndex: number,
): DraftJournalEntry {
	const autoFilter = autoFilterPreimage(workbook, op.sheet)
	const { inverseOps, issues } = restoreAutoFilterOps(autoFilter)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'auto-filter', autoFilter }],
		issues,
	}
}

function journalClearAutoFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'clearAutoFilter' }>,
	opIndex: number,
): DraftJournalEntry {
	const autoFilter = autoFilterPreimage(workbook, op.sheet)
	const { inverseOps, issues } = autoFilter.autoFilter
		? restoreAutoFilterOps(autoFilter)
		: { inverseOps: [], issues: [] }
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'auto-filter', autoFilter }],
		issues,
	}
}

function journalSetConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setConditionalFormat' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = conditionalFormatPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } =
		preimage.formats.length > 0
			? restoreConditionalFormatOps(op.sheet, preimage.formats)
			: {
					inverseOps: [
						{ op: 'deleteConditionalFormat', sheet: op.sheet, range: op.range },
					] satisfies Operation[],
					issues: [],
				}
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'conditional-formats', conditionalFormats: preimage }],
		issues,
	}
}

function journalDeleteConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteConditionalFormat' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = conditionalFormatPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } = restoreConditionalFormatOps(op.sheet, preimage.formats)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'conditional-formats', conditionalFormats: preimage }],
		issues: [
			...issues,
			...(op.range === undefined
				? [
						{
							code: 'LOSSY_INVERSE' as const,
							message:
								'Conditional-format deletion without a range may not restore original rule ordering exactly',
						},
					]
				: []),
		],
	}
}

function journalSetPageSetup(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPageSetup' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = pageSetupPreimage(workbook, op.sheet)
	const { inverseOps, issues } = restorePageSetupOps(preimage, op.setup.margins !== undefined)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'page-setup', pageSetup: preimage }],
		issues,
	}
}

function journalSetPrintArea(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPrintArea' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = definedNamePreimage(workbook, '_xlnm.Print_Area', op.sheet)
	const { inverseOps, issues } = restoreDefinedNameOps(workbook, preimage)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'defined-name', definedName: preimage }],
		issues,
	}
}

function journalSetDefinedName(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDefinedName' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = definedNamePreimage(workbook, op.name, op.scope)
	const { inverseOps, issues } = restoreDefinedNameOps(workbook, preimage)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'defined-name', definedName: preimage }],
		issues,
	}
}

function journalDeleteDefinedName(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDefinedName' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = definedNamePreimage(workbook, op.name, op.scope)
	const { inverseOps, issues } = preimage.definedName
		? restoreDefinedNameOps(workbook, preimage)
		: { inverseOps: [], issues: [] }
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'defined-name', definedName: preimage }],
		issues,
	}
}

function journalRenameSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'renameSheet', sheet: op.newName, newName: op.sheet }],
		preimages: formulaBindings.preimages,
		issues: formulaBindings.issues,
	}
}

function journalRenameTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const tableRename = { table: op.table, existed: true }
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'renameTable', table: op.newName, newName: op.table }],
		preimages: [{ kind: 'table-rename', tableRename }, ...formulaBindings.preimages],
		issues: formulaBindings.issues,
	}
}

function journalCreateTable(
	op: Extract<Operation, { op: 'createTable' }>,
	opIndex: number,
): DraftJournalEntry {
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'deleteTable', table: op.name }],
		preimages: [
			{
				kind: 'table',
				table: { table: null, sheet: op.sheet, ref: op.ref },
			},
		],
		issues: [],
	}
}

function journalDeleteTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tablePreimage(workbook, op.table)
	const { inverseOps, issues } = restoreTableOps(preimage)
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table', table: preimage }, ...formulaBindings.preimages],
		issues: [...issues, ...formulaBindings.issues],
	}
}

function journalResizeTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'resizeTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tablePreimage(workbook, op.table)
	const { inverseOps, issues } = preimage.table
		? restoreExistingTableOps(preimage.table, preimage.sheet ?? undefined, {
				includeCreate: false,
				tableName: op.table,
			})
		: { inverseOps: [], issues: missingTableIssues(op.table) }
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table', table: preimage }, ...formulaBindings.preimages],
		issues: [...issues, ...formulaBindings.issues],
	}
}

function journalSetTableColumn(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableColumn' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tableColumnPreimage(workbook, op.table, op.column)
	const { inverseOps, issues } = restoreTableColumnOps(op, preimage)
	const formulaBindings = formulaBindingJournalAddendum(
		tableColumnFormulaBindingCellPreimages(workbook, op),
	)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table-column', tableColumn: preimage }, ...formulaBindings.preimages],
		issues: [...issues, ...formulaBindings.issues],
	}
}

function journalSetTableStyle(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableStyle' }>,
	opIndex: number,
): DraftJournalEntry {
	const style = tableStylePreimage(workbook, op.table)
	const issues = tableStyleRestoreIssues(op, style.style)
	return {
		opIndex,
		op,
		inverseOps: [tableStyleSetOperation(op.table, style.style)],
		preimages: [{ kind: 'table-style', tableStyle: style }],
		issues,
	}
}

function journalSetComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setComment' }>,
	opIndex: number,
): DraftJournalEntry {
	const comment = commentPreimage(workbook, op.sheet, op.ref)
	const inverseOps: Operation[] = comment.comment
		? [
				{
					op: 'setComment',
					sheet: op.sheet,
					ref: comment.ref,
					text: comment.comment.text,
					...(comment.comment.author !== undefined ? { author: comment.comment.author } : {}),
				},
			]
		: [{ op: 'deleteComment', sheet: op.sheet, ref: comment.ref }]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'comment', comment }],
		issues: [],
	}
}

function journalDeleteComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteComment' }>,
	opIndex: number,
): DraftJournalEntry {
	const comment = commentPreimage(workbook, op.sheet, op.ref)
	const inverseOps: Operation[] = comment.comment
		? [
				{
					op: 'setComment',
					sheet: op.sheet,
					ref: comment.ref,
					text: comment.comment.text,
					...(comment.comment.author !== undefined ? { author: comment.comment.author } : {}),
				},
			]
		: []
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'comment', comment }],
		issues: [],
	}
}

function journalSetHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setHyperlink' }>,
	opIndex: number,
): DraftJournalEntry {
	const hyperlink = hyperlinkPreimage(workbook, op.sheet, op.ref)
	const inverseOps: Operation[] = hyperlink.hyperlink
		? [setHyperlinkInverse(op.sheet, hyperlink.ref, hyperlink.hyperlink)]
		: [{ op: 'deleteHyperlink', sheet: op.sheet, ref: hyperlink.ref }]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'hyperlink', hyperlink }],
		issues: [],
	}
}

function journalDeleteHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteHyperlink' }>,
	opIndex: number,
): DraftJournalEntry {
	const hyperlink = hyperlinkPreimage(workbook, op.sheet, op.ref)
	const inverseOps: Operation[] = hyperlink.hyperlink
		? [setHyperlinkInverse(op.sheet, hyperlink.ref, hyperlink.hyperlink)]
		: []
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'hyperlink', hyperlink }],
		issues: [],
	}
}

function journalSetThreadedComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setThreadedComment' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = threadedCommentPreimage(workbook, op)
	const inverseOps: Operation[] = preimage.threadedComment
		? [
				{
					op: 'setThreadedComment',
					sheet: preimage.sheet,
					text: preimage.threadedComment.text,
					...threadedCommentStableSelector(preimage),
				},
			]
		: []
	const issues: MutationJournalIssue[] = preimage.threadedComment
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: `Threaded comment selector on ${op.sheet} cannot be resolved exactly`,
				},
			]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'threaded-comment', threadedComment: preimage }],
		issues,
	}
}

function journalSetDrawingText(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDrawingText' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = drawingTextPreimage(workbook, op)
	const inverseOps: Operation[] =
		preimage.drawingObject?.text !== undefined
			? [
					{
						op: 'setDrawingText',
						sheet: preimage.sheet,
						text: preimage.drawingObject.text,
						...drawingObjectStableSelector(preimage),
					},
				]
			: []
	const issues: MutationJournalIssue[] =
		preimage.drawingObject?.text !== undefined
			? []
			: [
					{
						code: 'LOSSY_INVERSE',
						message: `Drawing object selector on ${op.sheet} cannot be resolved to editable text exactly`,
					},
				]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'drawing-text', drawingText: preimage }],
		issues,
	}
}

function journalSetChartSeriesSource(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setChartSeriesSource' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = chartSeriesPreimage(workbook, op)
	const inverse = chartSeriesInverseOperation(op, preimage)
	const issues: MutationJournalIssue[] = inverse.exact
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: `Chart series selector cannot be restored exactly for series ${op.seriesIndex}`,
				},
			]
	return {
		opIndex,
		op,
		inverseOps: inverse.op ? [inverse.op] : [],
		preimages: [{ kind: 'chart-series', chartSeries: preimage }],
		issues,
	}
}

function journalSetPivotCache(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPivotCache' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = pivotCachePreimage(workbook, op)
	const inverse = pivotCacheInverseOperation(op, preimage)
	const issues: MutationJournalIssue[] = inverse.exact
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: 'Pivot cache selector cannot be restored exactly',
				},
			]
	return {
		opIndex,
		op,
		inverseOps: inverse.op ? [inverse.op] : [],
		preimages: [{ kind: 'pivot-cache', pivotCache: preimage }],
		issues,
	}
}

function journalFreezePane(
	workbook: Workbook,
	op: Extract<Operation, { op: 'freezePane' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const pane = {
		sheet: op.sheet,
		frozenRows: sheet?.frozenRows ?? 0,
		frozenCols: sheet?.frozenCols ?? 0,
	}
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'freezePane', sheet: op.sheet, row: pane.frozenRows, col: pane.frozenCols }],
		preimages: [{ kind: 'pane', pane }],
		issues: [],
	}
}

function journalSetWorkbookProperties(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookProperties' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = { properties: { ...workbook.workbookProperties } }
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'setWorkbookProperties', properties: preimage.properties, mode: 'replace' }],
		preimages: [{ kind: 'workbook-properties', workbookProperties: preimage }],
		issues: [],
	}
}

function journalSetDocumentProperties(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDocumentProperties' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		properties: clonePlain(workbook.documentProperties) as WorkbookDocumentProperties,
	}
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'setDocumentProperties', properties: preimage.properties, mode: 'replace' }],
		preimages: [{ kind: 'document-properties', documentProperties: preimage }],
		issues: [],
	}
}

function journalSetWorkbookView(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookView' }>,
	opIndex: number,
): DraftJournalEntry {
	const index = op.index ?? 0
	const view = workbook.workbookViews[index]
	const preimage = { index, view: view ? { ...view } : null }
	return {
		opIndex,
		op,
		inverseOps: [
			preimage.view
				? { op: 'setWorkbookView', index, view: preimage.view, mode: 'replace' }
				: { op: 'setWorkbookView', index, view: null },
		],
		preimages: [{ kind: 'workbook-view', workbookView: preimage }],
		issues: [],
	}
}

function journalSetCalcSettings(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCalcSettings' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		settings: clonePlain(workbook.calcSettings) as CalcSettings,
		workbookProperties: { ...workbook.workbookProperties },
	}
	return {
		opIndex,
		op,
		inverseOps: [
			{ op: 'setCalcSettings', settings: preimage.settings },
			{ op: 'setWorkbookProperties', properties: preimage.workbookProperties, mode: 'replace' },
		],
		preimages: [{ kind: 'calc-settings', calcSettings: preimage }],
		issues: [],
	}
}

function journalSetWorkbookProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookProtection' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		protection: workbook.workbookProtection ? { ...workbook.workbookProtection } : null,
	}
	const issues: MutationJournalIssue[] = preimage.protection
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: 'Workbook protection absence cannot be restored exactly with public operations',
				},
			]
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'setWorkbookProtection', protection: preimage.protection ?? {} }],
		preimages: [{ kind: 'workbook-protection', workbookProtection: preimage }],
		issues,
	}
}

function journalSetTheme(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTheme' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		metadata: { ...workbook.themeMetadata },
		colors: workbook.themeColors.map((color) => ({ ...color })),
	}
	const inverse = themeInverseOperation(op, preimage)
	return {
		opIndex,
		op,
		inverseOps: inverse.op ? [inverse.op] : [],
		preimages: [{ kind: 'theme', theme: preimage }],
		issues: inverse.issues,
	}
}

type SetThemeOperation = Extract<Operation, { op: 'setTheme' }>
type MutableSetThemeOperation = {
	-readonly [K in keyof SetThemeOperation]: SetThemeOperation[K]
}

function themeInverseOperation(
	op: SetThemeOperation,
	preimage: MutationJournalThemePreimage,
): {
	readonly op: SetThemeOperation | null
	readonly issues: readonly MutationJournalIssue[]
} {
	const inverse: Partial<MutableSetThemeOperation> = { op: 'setTheme' }
	const issues: MutationJournalIssue[] = []
	for (const [opKey, metadataKey] of [
		['themeName', 'name'],
		['colorSchemeName', 'colorSchemeName'],
		['majorFontLatin', 'majorFontLatin'],
		['minorFontLatin', 'minorFontLatin'],
	] as const) {
		if (op[opKey] === undefined) continue
		const previous = preimage.metadata[metadataKey]
		if (previous !== undefined) {
			inverse[opKey] = previous
			continue
		}
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Theme metadata field ${opKey} cannot be removed with public operations`,
		})
	}
	if (op.themeColors !== undefined) {
		const colorsBySlot = new Map(preimage.colors.map((color) => [color.slot, color]))
		const inverseColors: WorkbookThemeColor[] = []
		for (const color of op.themeColors) {
			const previous = colorsBySlot.get(color.slot)
			if (!previous) {
				issues.push({
					code: 'LOSSY_INVERSE',
					message: `Theme color slot ${color.slot} cannot be removed with public operations`,
				})
				continue
			}
			if (previous.rgb === undefined && previous.systemColor === undefined) {
				issues.push({
					code: 'LOSSY_INVERSE',
					message: `Theme color slot ${color.slot} cannot be restored with public operations`,
				})
				continue
			}
			inverseColors.push({ ...previous })
		}
		if (inverseColors.length > 0) inverse.themeColors = inverseColors
	}
	const hasInverse =
		inverse.themeName !== undefined ||
		inverse.colorSchemeName !== undefined ||
		inverse.majorFontLatin !== undefined ||
		inverse.minorFontLatin !== undefined ||
		inverse.themeColors !== undefined
	return {
		op: hasInverse ? (inverse as SetThemeOperation) : null,
		issues,
	}
}

function mergePreimage(
	workbook: Workbook,
	sheetName: string,
	rangeText: string,
): MutationJournalMergePreimage {
	const target = parseRange(rangeText)
	const existed =
		workbook.getSheet(sheetName)?.merges.some((merge) => sameRange(merge, target)) ?? false
	return {
		sheet: sheetName,
		range: toRangeString(target),
		existed,
	}
}

function dataValidationPreimage(
	workbook: Workbook,
	sheetName: string,
	range: string,
): MutationJournalDataValidationPreimage {
	const validation = workbook.getSheet(sheetName)?.dataValidations.find((dv) => dv.sqref === range)
	return {
		sheet: sheetName,
		range,
		validation: validation ? { ...validation } : null,
	}
}

function restoreDataValidationOps(validation: MutationJournalDataValidationPreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (!validation.validation) {
		return {
			inverseOps: [
				{ op: 'deleteDataValidation', sheet: validation.sheet, range: validation.range },
			],
			issues: [],
		}
	}
	const { rule, issues } = dataValidationRuleFromSheet(validation)
	return {
		inverseOps: rule
			? [{ op: 'setDataValidation', sheet: validation.sheet, range: validation.range, rule }]
			: [],
		issues,
	}
}

function dataValidationRuleFromSheet(validation: MutationJournalDataValidationPreimage): {
	readonly rule: DataValidationRule | null
	readonly issues: readonly MutationJournalIssue[]
} {
	const source = validation.validation
	if (!source?.type || !isDataValidationType(source.type)) {
		return {
			rule: null,
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore data validation at ${validation.sheet}!${validation.range} with unsupported type ${source?.type ?? '<missing>'}`,
					refs: [`${validation.sheet}!${validation.range}`],
				},
			],
		}
	}
	const issues: MutationJournalIssue[] = []
	if (source.source === 'x14' || source.uid !== undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Data validation extension metadata at ${validation.sheet}!${validation.range} cannot be restored with public operations`,
			refs: [`${validation.sheet}!${validation.range}`],
		})
	}
	return {
		rule: {
			type: source.type,
			...(isDataValidationOperator(source.operator) ? { operator: source.operator } : {}),
			...(source.formula1 !== undefined ? { formula1: source.formula1 } : {}),
			...(source.formula2 !== undefined ? { formula2: source.formula2 } : {}),
			...(source.allowBlank !== undefined ? { allowBlank: source.allowBlank } : {}),
			...(source.showDropDown !== undefined ? { showDropDown: source.showDropDown } : {}),
			...(source.showErrorMessage !== undefined
				? { showErrorMessage: source.showErrorMessage }
				: {}),
			...(source.errorTitle !== undefined ? { errorTitle: source.errorTitle } : {}),
			...(source.error !== undefined ? { errorMessage: source.error } : {}),
			...(source.errorStyle !== undefined ? { errorStyle: source.errorStyle } : {}),
			...(source.imeMode !== undefined ? { imeMode: source.imeMode } : {}),
			...(source.showInputMessage !== undefined
				? { showInputMessage: source.showInputMessage }
				: {}),
			...(source.promptTitle !== undefined ? { promptTitle: source.promptTitle } : {}),
			...(source.prompt !== undefined ? { prompt: source.prompt } : {}),
		},
		issues,
	}
}

function autoFilterPreimage(
	workbook: Workbook,
	sheetName: string,
): MutationJournalAutoFilterPreimage {
	const autoFilter = workbook.getSheet(sheetName)?.autoFilter
	return {
		sheet: sheetName,
		autoFilter: autoFilter ? cloneAutoFilter(autoFilter) : null,
	}
}

function restoreAutoFilterOps(preimage: MutationJournalAutoFilterPreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const autoFilter = preimage.autoFilter
	if (!autoFilter) {
		return {
			inverseOps: [{ op: 'clearAutoFilter', sheet: preimage.sheet }],
			issues: [],
		}
	}
	const inverseOps: Operation[] = [
		{ op: 'setAutoFilter', sheet: preimage.sheet, range: autoFilter.ref },
	]
	const issues: MutationJournalIssue[] = []
	for (const column of autoFilter.columns) {
		if (
			column.kind === 'filters' &&
			column.values !== undefined &&
			column.blank === undefined &&
			column.calendarType === undefined &&
			column.dateGroupItems === undefined &&
			column.hiddenButton === undefined &&
			column.showButton === undefined
		) {
			inverseOps.push({
				op: 'setAutoFilter',
				sheet: preimage.sheet,
				range: autoFilter.ref,
				column: column.colId,
				values: [...column.values],
			})
			continue
		}
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `AutoFilter column ${column.colId} on ${preimage.sheet}!${autoFilter.ref} cannot be fully restored with public operations`,
			refs: [`${preimage.sheet}!${autoFilter.ref}`],
		})
	}
	if (autoFilter.uid !== undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `AutoFilter extension metadata on ${preimage.sheet}!${autoFilter.ref} cannot be restored with public operations`,
			refs: [`${preimage.sheet}!${autoFilter.ref}`],
		})
	}
	if (autoFilter.sortState) {
		const [firstCondition, ...extraConditions] = autoFilter.sortState.conditions
		if (
			autoFilter.sortState.caseSensitive !== undefined ||
			autoFilter.sortState.columnSort !== undefined ||
			autoFilter.sortState.sortMethod !== undefined ||
			autoFilter.sortState.preservedAttributes !== undefined ||
			extraConditions.length > 0 ||
			firstCondition?.customList !== undefined ||
			firstCondition?.dxfId !== undefined ||
			firstCondition?.iconSet !== undefined ||
			firstCondition?.iconId !== undefined
		) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `AutoFilter sort metadata on ${preimage.sheet}!${autoFilter.ref} cannot be fully restored with public operations`,
				refs: [`${preimage.sheet}!${autoFilter.ref}`],
			})
		}
		if (firstCondition) {
			inverseOps.push({
				op: 'setAutoFilter',
				sheet: preimage.sheet,
				range: autoFilter.ref,
				sortRef: autoFilter.sortState.ref,
				sortBy: firstCondition.ref,
				...(firstCondition.descending !== undefined
					? { descending: firstCondition.descending }
					: {}),
			})
		}
	}
	return { inverseOps, issues }
}

function conditionalFormatPreimage(
	workbook: Workbook,
	sheetName: string,
	range: string | undefined,
): MutationJournalConditionalFormatPreimage {
	const formats = workbook
		.getSheet(sheetName)
		?.conditionalFormats.filter((cf) => range === undefined || cf.sqref === range)
	return {
		sheet: sheetName,
		...(range !== undefined ? { range } : {}),
		formats: (formats ?? []).map(cloneConditionalFormat),
	}
}

function restoreConditionalFormatOps(
	sheet: string,
	formats: readonly SheetConditionalFormat[],
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const inverseOps: Operation[] = []
	const issues: MutationJournalIssue[] = []
	for (const format of formats) {
		inverseOps.push({ op: 'deleteConditionalFormat', sheet, range: format.sqref })
		format.rules.forEach((rule, index) => {
			const converted = conditionalFormatRuleFromSheet(sheet, format.sqref, rule)
			issues.push(...converted.issues)
			if (!converted.rule) return
			inverseOps.push({
				op: 'setConditionalFormat',
				sheet,
				range: format.sqref,
				rule: converted.rule,
				mode: index === 0 ? 'replace' : 'append',
			})
		})
	}
	return { inverseOps, issues }
}

function conditionalFormatRuleFromSheet(
	sheet: string,
	range: string,
	rule: SheetConditionalFormatRule,
): {
	readonly rule: ConditionalFormatRule | null
	readonly issues: readonly MutationJournalIssue[]
} {
	const issues: MutationJournalIssue[] = []
	if (!isConditionalFormatType(rule.type)) {
		return {
			rule: null,
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore conditional format at ${sheet}!${range} with unsupported type ${rule.type}`,
					refs: [`${sheet}!${range}`],
				},
			],
		}
	}
	if (
		rule.dxfId !== undefined ||
		rule.rank !== undefined ||
		rule.percent !== undefined ||
		rule.bottom !== undefined ||
		rule.aboveAverage !== undefined ||
		rule.equalAverage !== undefined ||
		rule.stdDev !== undefined ||
		rule.text !== undefined ||
		rule.timePeriod !== undefined ||
		rule.preservedRuleAttributes !== undefined ||
		rule.preservedRuleChildXml !== undefined
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Conditional format metadata at ${sheet}!${range} cannot be fully restored with public operations`,
			refs: [`${sheet}!${range}`],
		})
	}
	return {
		rule: {
			type: rule.type,
			...(isConditionalFormatOperator(rule.operator) ? { operator: rule.operator } : {}),
			...(rule.formulas[0] !== undefined ? { formula: rule.formulas[0] } : {}),
			...(rule.formulas[1] !== undefined ? { formula2: rule.formulas[1] } : {}),
			...(rule.priority !== undefined ? { priority: rule.priority } : {}),
			...(rule.stopIfTrue !== undefined ? { stopIfTrue: rule.stopIfTrue } : {}),
			...(rule.style ? { style: cloneCellStyle(rule.style) as StyleInput } : {}),
			...(rule.colorScale ? { colorScale: clonePlain(rule.colorScale) } : {}),
			...(rule.dataBar ? { dataBar: clonePlain(rule.dataBar) } : {}),
			...(rule.iconSet ? { iconSet: clonePlain(rule.iconSet) } : {}),
		},
		issues,
	}
}

function definedNamePreimage(
	workbook: Workbook,
	name: string,
	scopeSheetName: string | undefined,
): MutationJournalDefinedNamePreimage {
	const scope = scopeSheetName
		? sheetScopeForName(workbook, scopeSheetName)
		: ({ kind: 'workbook' } as const)
	const definedName = scope ? workbook.definedNames.getEntry(name, scope) : undefined
	return {
		name,
		...(scopeSheetName !== undefined ? { scope: scopeSheetName } : {}),
		definedName: definedName ? cloneDefinedName(definedName) : null,
	}
}

function pageSetupPreimage(
	workbook: Workbook,
	sheetName: string,
): MutationJournalPageSetupPreimage {
	const sheet = workbook.getSheet(sheetName)
	return {
		sheet: sheetName,
		pageSetup: sheet?.pageSetup ? clonePageSetup(sheet.pageSetup) : null,
		pageMargins: sheet?.pageMargins ? { ...sheet.pageMargins } : null,
	}
}

function restorePageSetupOps(
	preimage: MutationJournalPageSetupPreimage,
	marginsTouched: boolean,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const issues: MutationJournalIssue[] = []
	const setup = preimage.pageSetup ? pageSetupToInput(preimage.pageSetup) : null
	if (preimage.pageSetup && !setup) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Page setup for ${preimage.sheet} contains metadata that cannot be restored with public operations`,
			refs: [preimage.sheet],
		})
	}
	if (!preimage.pageSetup) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Page setup for ${preimage.sheet} cannot be removed with public operations`,
			refs: [preimage.sheet],
		})
	}
	if (marginsTouched && !preimage.pageMargins) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Page margins for ${preimage.sheet} cannot be removed with public operations`,
			refs: [preimage.sheet],
		})
	}
	if (!setup) return { inverseOps: [], issues }
	return {
		inverseOps: [
			{
				op: 'setPageSetup',
				sheet: preimage.sheet,
				setup: {
					...setup,
					...(preimage.pageMargins ? { margins: { ...preimage.pageMargins } } : {}),
				},
			},
		],
		issues,
	}
}

function pageSetupToInput(
	setup: SheetPageSetup,
): Omit<Extract<Operation, { op: 'setPageSetup' }>['setup'], 'margins'> | null {
	const unsupported = [
		setup.firstPageNumber,
		setup.copies,
		setup.horizontalDpi,
		setup.verticalDpi,
		setup.pageOrder,
		setup.cellComments,
		setup.errors,
		setup.blackAndWhite,
		setup.draft,
		setup.useFirstPageNumber,
		setup.usePrinterDefaults,
		setup.printerSettingsRelId,
	]
	if (unsupported.some((value) => value !== undefined)) return null
	if (
		setup.orientation !== undefined &&
		setup.orientation !== 'portrait' &&
		setup.orientation !== 'landscape'
	) {
		return null
	}
	return {
		...(setup.orientation !== undefined ? { orientation: setup.orientation } : {}),
		...(setup.paperSize !== undefined ? { paperSize: setup.paperSize } : {}),
		...(setup.scale !== undefined ? { scale: setup.scale } : {}),
		...(setup.fitToWidth !== undefined ? { fitToWidth: setup.fitToWidth } : {}),
		...(setup.fitToHeight !== undefined ? { fitToHeight: setup.fitToHeight } : {}),
	}
}

function restoreDefinedNameOps(
	workbook: Workbook,
	preimage: MutationJournalDefinedNamePreimage,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const definedName = preimage.definedName
	if (!definedName) {
		return {
			inverseOps: [
				{
					op: 'deleteDefinedName',
					name: preimage.name,
					...(preimage.scope !== undefined ? { scope: preimage.scope } : {}),
				},
			],
			issues: [],
		}
	}
	const scope =
		definedName.scope.kind === 'sheet'
			? sheetNameForId(workbook, definedName.scope.sheetId)
			: undefined
	const refs = [`${scope ? `${scope}!` : ''}${definedName.name}`]
	const issues: MutationJournalIssue[] = []
	if (definedName.hidden !== undefined || definedName.extraAttributes !== undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Defined name ${refs[0]} has metadata that cannot be restored with public operations`,
			refs,
		})
	}
	if (definedName.scope.kind === 'sheet' && scope === undefined) {
		return {
			inverseOps: [],
			issues: [
				...issues,
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore sheet-scoped defined name ${definedName.name} because its sheet scope is missing`,
				},
			],
		}
	}
	return {
		inverseOps: [
			{
				op: 'setDefinedName',
				name: definedName.name,
				ref: definedName.formula,
				...(scope !== undefined ? { scope } : {}),
			},
		],
		issues,
	}
}

function tableColumnPreimage(
	workbook: Workbook,
	tableName: string,
	column: string | number,
): MutationJournalTableColumnPreimage {
	const located = findTableColumn(workbook, tableName, column)
	return {
		table: tableName,
		column,
		columnIndex: located?.columnIndex ?? -1,
		columnState: located?.column ? cloneTableColumn(located.column) : null,
	}
}

function restoreTableColumnOps(
	op: Extract<Operation, { op: 'setTableColumn' }>,
	preimage: MutationJournalTableColumnPreimage,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const column = preimage.columnState
	if (!column) {
		return {
			inverseOps: [],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore table column ${op.table}[${String(op.column)}] because it was not found before the edit`,
				},
			],
		}
	}
	const inverse: Extract<Operation, { op: 'setTableColumn' }> = {
		op: 'setTableColumn',
		table: op.table,
		column: op.newName ?? op.column,
		...(op.newName !== undefined ? { newName: column.name } : {}),
		...(op.formula !== undefined ? { formula: column.formula ?? null } : {}),
		...(op.totalsRowFunction !== undefined
			? { totalsRowFunction: column.totalsRowFunction ?? null }
			: {}),
		...(op.totalsRowFormula !== undefined
			? { totalsRowFormula: column.totalsRowFormula ?? null }
			: {}),
		...(op.totalsRowLabel !== undefined ? { totalsRowLabel: column.totalsRowLabel ?? null } : {}),
	}
	return { inverseOps: [inverse], issues: [] }
}

function tablePreimage(workbook: Workbook, tableName: string): MutationJournalTablePreimage {
	const located = findTable(workbook, tableName)
	return {
		table: located ? cloneTable(located.table) : null,
		sheet: located?.sheet.name ?? null,
		ref: located ? toRangeString(located.table.ref) : null,
	}
}

function restoreTableOps(preimage: MutationJournalTablePreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (!preimage.table) {
		return {
			inverseOps: [],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: 'Cannot restore deleted table because it was not found before the edit',
				},
			],
		}
	}
	return restoreExistingTableOps(preimage.table, preimage.sheet ?? undefined, {
		includeCreate: true,
	})
}

function restoreExistingTableOps(
	table: Table,
	sheet: string | undefined,
	options: { readonly includeCreate: boolean; readonly tableName?: string },
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const issues = tableRestoreIssues(table, sheet, options.includeCreate)
	if (!sheet) {
		return {
			inverseOps: [],
			issues: [
				...issues,
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore table ${table.name} because its sheet is missing`,
				},
			],
		}
	}
	const inverseOps: Operation[] = options.includeCreate
		? [
				{
					op: 'createTable',
					sheet,
					ref: toRangeString(table.ref),
					name: table.name,
					hasHeaders: table.hasHeaders,
				},
			]
		: [
				{
					op: 'resizeTable',
					table: options.tableName ?? table.name,
					ref: toRangeString(table.ref),
				},
			]
	if (table.tableStyleInfo) {
		inverseOps.push(tableStyleSetOperation(table.name, table.tableStyleInfo))
	}
	table.columns.forEach((column, index) => {
		const op = tableColumnRestoreOperation(table.name, index, column, {
			restoreName: options.includeCreate,
		})
		if (op) inverseOps.push(op)
	})
	return { inverseOps, issues }
}

function tableStylePreimage(
	workbook: Workbook,
	tableName: string,
): MutationJournalTableStylePreimage {
	const style = findTable(workbook, tableName)?.table.tableStyleInfo
	return {
		table: tableName,
		style: style ? { ...style } : null,
	}
}

function tableStyleSetOperation(
	table: string,
	style: TableStyleInfo | null,
): Extract<Operation, { op: 'setTableStyle' }> {
	return {
		op: 'setTableStyle',
		table,
		styleName: style?.name ?? null,
		...(style?.showFirstColumn !== undefined ? { showFirstColumn: style.showFirstColumn } : {}),
		...(style?.showLastColumn !== undefined ? { showLastColumn: style.showLastColumn } : {}),
		...(style?.showRowStripes !== undefined ? { showRowStripes: style.showRowStripes } : {}),
		...(style?.showColumnStripes !== undefined
			? { showColumnStripes: style.showColumnStripes }
			: {}),
	}
}

function tableStyleRestoreIssues(
	op: Extract<Operation, { op: 'setTableStyle' }>,
	style: TableStyleInfo | null,
): MutationJournalIssue[] {
	const touchedBooleanWithNoPreimage =
		(op.showFirstColumn !== undefined && style?.showFirstColumn === undefined) ||
		(op.showLastColumn !== undefined && style?.showLastColumn === undefined) ||
		(op.showRowStripes !== undefined && style?.showRowStripes === undefined) ||
		(op.showColumnStripes !== undefined && style?.showColumnStripes === undefined)
	return touchedBooleanWithNoPreimage
		? [
				{
					code: 'LOSSY_INVERSE',
					message: `Table style flags for ${op.table} cannot be cleared with public operations`,
				},
			]
		: []
}

function tableColumnRestoreOperation(
	table: string,
	columnIndex: number,
	column: TableColumn,
	options: { readonly restoreName: boolean },
): Extract<Operation, { op: 'setTableColumn' }> | null {
	const op: Extract<Operation, { op: 'setTableColumn' }> = {
		op: 'setTableColumn',
		table,
		column: columnIndex,
		...(options.restoreName ? { newName: column.name } : {}),
		...(column.formula !== undefined ? { formula: column.formula } : {}),
		...(column.totalsRowFunction !== undefined
			? { totalsRowFunction: column.totalsRowFunction }
			: {}),
		...(column.totalsRowFormula !== undefined ? { totalsRowFormula: column.totalsRowFormula } : {}),
		...(column.totalsRowLabel !== undefined ? { totalsRowLabel: column.totalsRowLabel } : {}),
	}
	return Object.keys(op).length > 3 ? op : null
}

function tableRestoreIssues(
	table: Table,
	sheet: string | undefined,
	includeCreate: boolean,
): MutationJournalIssue[] {
	const refs = [sheet ? `${sheet}!${toRangeString(table.ref)}` : table.name]
	const issues: MutationJournalIssue[] = []
	if (
		table.partPath !== undefined ||
		table.contentType !== undefined ||
		table.contentTypeSource !== undefined ||
		table.sourcePartPath !== undefined ||
		table.sourceRelationshipPart !== undefined ||
		table.sourceRelationshipId !== undefined ||
		table.sourceRelationshipType !== undefined ||
		table.sourceRelationshipRawType !== undefined ||
		table.sourceRelationshipRawTarget !== undefined ||
		table.sourceRelationshipResolvedTarget !== undefined ||
		table.sourceRelationshipTargetMode !== undefined ||
		table.uid !== undefined ||
		table.tableType !== undefined ||
		table.insertRow !== undefined ||
		table.insertRowShift !== undefined ||
		table.altText !== undefined ||
		table.altTextSummary !== undefined ||
		table.autoFilter !== undefined ||
		table.sortState !== undefined ||
		table.dxfId !== undefined ||
		table.dataCellStyle !== undefined ||
		table.headerRowDxfId !== undefined ||
		table.headerRowCellStyle !== undefined ||
		table.dataDxfId !== undefined ||
		table.totalsRowDxfId !== undefined ||
		table.headerRowBorderDxfId !== undefined ||
		table.tableBorderDxfId !== undefined ||
		table.queryTable !== undefined
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Table ${table.name} has metadata that cannot be restored with public operations`,
			refs,
		})
	}
	if (includeCreate && table.hasTotals) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Table ${table.name} totals-row state cannot be recreated with public operations`,
			refs,
		})
	}
	for (const column of table.columns) {
		if (
			column.id !== undefined ||
			column.uid !== undefined ||
			column.uniqueName !== undefined ||
			column.formulaIsArray !== undefined ||
			column.xmlColumnPr !== undefined ||
			column.queryTableFieldId !== undefined ||
			column.dataCellStyle !== undefined ||
			column.dataDxfId !== undefined ||
			column.headerRowDxfId !== undefined ||
			column.totalsRowDxfId !== undefined
		) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Table column ${table.name}[${column.name}] has metadata that cannot be restored with public operations`,
				refs,
			})
		}
	}
	return issues
}

function missingTableIssues(table: string): readonly MutationJournalIssue[] {
	return [
		{
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot restore table ${table} because it was not found before the edit`,
		},
	]
}

function structuralDeletePreimage(
	workbook: Workbook,
	sheetName: string,
	axis: 'row' | 'col',
	at: number,
	count: number,
): MutationJournalStructuralPreimage {
	const refs = structuralDeletedCellRefs(workbook.getSheet(sheetName), axis, at, count)
	return {
		sheet: sheetName,
		axis,
		at,
		count,
		deletedCells: cellPreimages(workbook, sheetName, refs),
	}
}

function structuralDeletedCellRefs(
	sheet: Sheet | undefined,
	axis: 'row' | 'col',
	at: number,
	count: number,
): string[] {
	const usedRange = sheet?.cells.usedRange()
	if (!sheet || !usedRange) return []
	const deleted =
		axis === 'row'
			? {
					start: { row: at, col: usedRange.start.col },
					end: { row: at + count - 1, col: usedRange.end.col },
				}
			: {
					start: { row: usedRange.start.row, col: at },
					end: { row: usedRange.end.row, col: at + count - 1 },
				}
	const refs: string[] = []
	for (const [row, col] of sheet.cells.getRange(deleted)) {
		refs.push(toA1({ row, col }))
	}
	return refs
}

function structuralDeleteIssues(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): readonly MutationJournalIssue[] {
	const sheet = workbook.getSheet(preimage.sheet)
	if (!sheet) return []
	const refs = [
		`${preimage.sheet}!${preimage.axis === 'row' ? preimage.at + 1 : toA1({ row: 0, col: preimage.at }).replace(/\d+$/, '')}`,
	]
	const issues: MutationJournalIssue[] = []
	const affected = structuralAffectedRange(preimage.axis, preimage.at, preimage.count)
	if (
		hasDeletedAxisDimensions(sheet, preimage.axis, preimage.at, preimage.count) ||
		hasMapRefInRange(sheet.comments, affected) ||
		hasMapRefInRange(sheet.hyperlinks, affected) ||
		sheet.threadedComments.some((comment) => refInRange(comment.ref, affected)) ||
		sheet.merges.some((merge) => rangesOverlap(merge, affected)) ||
		sheet.dataValidations.some((validation) => sqrefOverlaps(validation.sqref, affected)) ||
		sheet.conditionalFormats.some((format) => sqrefOverlaps(format.sqref, affected)) ||
		sheet.tables.some((table) => rangesOverlap(table.ref, affected)) ||
		(sheet.autoFilter?.ref ? sqrefOverlaps(sheet.autoFilter.ref, affected) : false)
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} metadata on ${preimage.sheet} cannot be fully restored with public operations`,
			refs,
		})
	}
	issues.push(...structuralFormulaReferenceIssues(workbook, preimage))
	issues.push(...structuralX14MetadataIssues(sheet, preimage, affected))
	issues.push(...structuralRepresentedMetadataIssues(workbook, sheet, preimage, affected))
	return issues
}

function structuralX14MetadataIssues(
	sheet: Sheet,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): readonly MutationJournalIssue[] {
	const refs = [
		...sheet.x14DataValidations
			.filter((validation) => !validation.deleted && sqrefOverlaps(validation.sqref, affected))
			.map((validation) => `${sheet.name}!x14Validation:${validation.sqref}:${validation.index}`),
		...sheet.x14ConditionalFormats
			.filter((format) => !format.deleted && sqrefOverlaps(format.sqref, affected))
			.map((format) => `${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}`),
	]
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} x14 metadata on ${preimage.sheet} cannot be fully restored with public operations`,
			refs,
		},
	]
}

function structuralFormulaReferenceIssues(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): readonly MutationJournalIssue[] {
	const refs = structuralFormulaReferenceLocations(workbook, preimage)
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} formula references on ${preimage.sheet} cannot be restored with public operations`,
			refs,
		},
	]
}

function structuralRepresentedMetadataIssues(
	workbook: Workbook,
	sheet: Sheet,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): readonly MutationJournalIssue[] {
	const refs = [
		...sheet.ignoredErrors
			.filter((entry) => sqrefOverlaps(entry.sqref, affected))
			.map((entry) => `${sheet.name}!ignoredError:${entry.sqref}`),
		...sortStateStructuralRefs(sheet.name, sheet.sortState, affected, 'sortState'),
		...sheet.advancedFilters.flatMap((filter, index) => [
			...(filter.ref && sqrefOverlaps(filter.ref, affected)
				? [`${sheet.name}!advancedFilter:${index}:${filter.ref}`]
				: []),
			...autoFilterStructuralRefs(
				sheet.name,
				filter.autoFilter,
				affected,
				`advancedFilter:${index}:autoFilter`,
			),
		]),
		...sheet.imageRefs
			.filter((image) => anchorOverlapsAffected(image.anchor, affected))
			.map((image, index) => `${sheet.name}!image:${image.drawingPartPath}:${index}`),
		...sheet.drawingObjectRefs
			.filter((object) => anchorOverlapsAffected(object.anchor, affected))
			.map((object, index) => `${sheet.name}!drawing:${object.drawingPartPath}:${index}`),
		...chartStructuralFormulaRefs(workbook, preimage),
		...pivotStructuralRefs(workbook, preimage, affected),
	]
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} represented metadata on ${preimage.sheet} cannot be fully restored with public operations`,
			refs,
		},
	]
}

function autoFilterStructuralRefs(
	sheetName: string,
	autoFilter: AutoFilter | null | undefined,
	affected: RangeRef,
	prefix: string,
): string[] {
	if (!autoFilter) return []
	return [
		...(sqrefOverlaps(autoFilter.ref, affected)
			? [`${sheetName}!${prefix}:${autoFilter.ref}`]
			: []),
		...sortStateStructuralRefs(sheetName, autoFilter.sortState, affected, `${prefix}:sortState`),
	]
}

function sortStateStructuralRefs(
	sheetName: string,
	sortState: AutoFilter['sortState'] | null,
	affected: RangeRef,
	prefix: string,
): string[] {
	if (!sortState) return []
	return [
		...(sqrefOverlaps(sortState.ref, affected) ? [`${sheetName}!${prefix}:${sortState.ref}`] : []),
		...sortState.conditions
			.filter((condition) => sqrefOverlaps(condition.ref, affected))
			.map((condition, index) => `${sheetName}!${prefix}:condition:${index}:${condition.ref}`),
	]
}

function structuralFormulaReferenceLocations(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): string[] {
	const refs: string[] = []
	for (const sheet of workbook.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			if (formulaReferencesDeletedAxis(workbook, cell.formula, sheet.name, preimage)) {
				refs.push(`${sheet.name}!${toA1({ row, col })}`)
			}
		}
		for (const validation of sheet.dataValidations) {
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula1,
				sheet.name,
				preimage,
				`${sheet.name}!validation:${validation.sqref}:formula1`,
			)
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula2,
				sheet.name,
				preimage,
				`${sheet.name}!validation:${validation.sqref}:formula2`,
			)
		}
		for (const validation of sheet.x14DataValidations) {
			if (validation.deleted) continue
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula1,
				sheet.name,
				preimage,
				`${sheet.name}!x14Validation:${validation.sqref}:formula1`,
			)
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula2,
				sheet.name,
				preimage,
				`${sheet.name}!x14Validation:${validation.sqref}:formula2`,
			)
		}
		sheet.conditionalFormats.forEach((format, formatIndex) => {
			format.rules.forEach((rule, ruleIndex) => {
				rule.formulas.forEach((formula, formulaIndex) => {
					pushMetadataFormulaReferenceLocation(
						refs,
						workbook,
						formula,
						sheet.name,
						preimage,
						`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:${formulaIndex}`,
					)
				})
				pushConditionalFormatValueObjectReferenceLocations(
					refs,
					workbook,
					rule.colorScale?.cfvo,
					sheet.name,
					preimage,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:colorScale.cfvo`,
				)
				pushConditionalFormatValueObjectReferenceLocations(
					refs,
					workbook,
					rule.dataBar?.cfvo,
					sheet.name,
					preimage,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:dataBar.cfvo`,
				)
				pushConditionalFormatValueObjectReferenceLocations(
					refs,
					workbook,
					rule.iconSet?.cfvo,
					sheet.name,
					preimage,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:iconSet.cfvo`,
				)
			})
		})
		sheet.x14ConditionalFormats.forEach((format) => {
			if (format.deleted) return
			format.formulas.forEach((formula, formulaIndex) => {
				pushMetadataFormulaReferenceLocation(
					refs,
					workbook,
					formula,
					sheet.name,
					preimage,
					`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:${formulaIndex}`,
				)
			})
			pushConditionalFormatValueObjectReferenceLocations(
				refs,
				workbook,
				format.colorScale?.cfvo,
				sheet.name,
				preimage,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:colorScale.cfvo`,
			)
			pushConditionalFormatValueObjectReferenceLocations(
				refs,
				workbook,
				format.dataBar?.cfvo,
				sheet.name,
				preimage,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:dataBar.cfvo`,
			)
			pushConditionalFormatValueObjectReferenceLocations(
				refs,
				workbook,
				format.iconSet?.cfvo,
				sheet.name,
				preimage,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:iconSet.cfvo`,
			)
		})
		for (const table of sheet.tables) {
			for (const column of table.columns) {
				pushMetadataFormulaReferenceLocation(
					refs,
					workbook,
					column.formula,
					sheet.name,
					preimage,
					`${sheet.name}!table:${table.name}:${column.name}:formula`,
				)
				pushMetadataFormulaReferenceLocation(
					refs,
					workbook,
					column.totalsRowFormula,
					sheet.name,
					preimage,
					`${sheet.name}!table:${table.name}:${column.name}:totalsRowFormula`,
				)
			}
		}
	}
	for (const name of workbook.definedNames.list()) {
		const scopeSheet =
			name.scope.kind === 'sheet' ? sheetNameForId(workbook, name.scope.sheetId) : undefined
		if (
			formulaReferencesDeletedAxis(workbook, name.formula, scopeSheet ?? preimage.sheet, preimage)
		) {
			refs.push(definedNameJournalKey(workbook, name))
		}
	}
	return refs
}

function pushConditionalFormatValueObjectReferenceLocations(
	refs: string[],
	workbook: Workbook,
	entries: readonly SheetConditionalFormatValueObject[] | undefined,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
	location: string,
): void {
	entries?.forEach((entry, index) => {
		pushMetadataFormulaReferenceLocation(
			refs,
			workbook,
			entry.value,
			ownerSheet,
			preimage,
			`${location}:${index}`,
		)
	})
}

function chartStructuralFormulaRefs(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): string[] {
	const refs: string[] = []
	for (const chart of workbook.chartParts) {
		chart.series.forEach((series, seriesIndex) => {
			for (const field of ['nameRef', 'categoryRef', 'valueRef'] as const) {
				const formula = series[field]
				if (
					formula !== undefined &&
					formulaReferencesDeletedAxis(
						workbook,
						formula,
						chart.sheetName ?? preimage.sheet,
						preimage,
					)
				) {
					refs.push(`chart:${chart.partPath}:series:${seriesIndex}:${field}`)
				}
			}
		})
	}
	return refs
}

function pivotStructuralRefs(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): string[] {
	const refs: string[] = []
	for (const cache of workbook.pivotCaches) {
		if (
			cache.sourceRef !== undefined &&
			refTextOverlapsAffected(cache.sourceRef, cache.sourceSheet, preimage, affected)
		) {
			refs.push(`pivotCache:${cache.partPath}:sourceRef`)
		}
	}
	for (const pivot of workbook.pivotTables) {
		if (
			pivot.locationRef &&
			refTextOverlapsAffected(pivot.locationRef, pivot.sheetName, preimage, affected)
		) {
			refs.push(`pivotTable:${pivot.partPath}:locationRef`)
		}
		if (
			pivot.location?.ref &&
			refTextOverlapsAffected(pivot.location.ref, pivot.sheetName, preimage, affected)
		) {
			refs.push(`pivotTable:${pivot.partPath}:location.ref`)
		}
	}
	return refs
}

function refTextOverlapsAffected(
	refText: string,
	ownerSheet: string | undefined,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): boolean {
	const split = splitSheetQualifiedRefText(refText)
	const sheetName = split?.sheet ?? ownerSheet
	if (sheetName !== preimage.sheet) return false
	const ref = refTextToRange(split?.ref ?? refText)
	return ref ? rangesOverlap(ref, affected) : false
}

function splitSheetQualifiedRefText(
	input: string,
): { readonly sheet: string; readonly ref: string } | null {
	const bang = input.lastIndexOf('!')
	if (bang < 0) return null
	const sheet = input.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
	const ref = input.slice(bang + 1)
	return sheet && ref ? { sheet, ref } : null
}

function refTextToRange(input: string): RangeRef | null {
	const normalized = input.replace(/\$/g, '')
	try {
		return parseRange(normalized)
	} catch {
		try {
			const ref = parseA1(normalized)
			return { start: ref, end: ref }
		} catch {
			return null
		}
	}
}

function anchorOverlapsAffected(anchor: SheetImageAnchor | undefined, affected: RangeRef): boolean {
	if (!anchor || anchor.kind === 'absolute') return false
	const range =
		anchor.kind === 'oneCell'
			? {
					start: { row: anchor.from.row, col: anchor.from.col },
					end: { row: anchor.from.row, col: anchor.from.col },
				}
			: {
					start: {
						row: Math.min(anchor.from.row, anchor.to.row),
						col: Math.min(anchor.from.col, anchor.to.col),
					},
					end: {
						row: Math.max(anchor.from.row, anchor.to.row),
						col: Math.max(anchor.from.col, anchor.to.col),
					},
				}
	return rangesOverlap(range, affected)
}

function pushMetadataFormulaReferenceLocation(
	refs: string[],
	workbook: Workbook,
	formula: string | undefined,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
	location: string,
): void {
	if (formula === undefined) return
	if (formulaReferencesDeletedAxis(workbook, formula, ownerSheet, preimage)) refs.push(location)
}

function formulaReferencesDeletedAxis(
	workbook: Workbook,
	formula: string,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
): boolean {
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return false
	return extractRefs(parsed.value).some((ref) =>
		formulaRefOverlapsDeletedAxis(workbook, ref, ownerSheet, preimage),
	)
}

function formulaRefOverlapsDeletedAxis(
	workbook: Workbook,
	ref: FormulaRef,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
): boolean {
	if (ref.kind === 'sheetSpan') {
		return (
			sheetSpanIncludes(workbook, ref.startSheet, ref.endSheet, preimage.sheet) &&
			formulaRefOverlapsDeletedAxis(workbook, ref.target, preimage.sheet, preimage)
		)
	}
	const sheet = 'sheet' in ref && ref.sheet !== undefined ? ref.sheet : ownerSheet
	if (sheet !== preimage.sheet) return false
	const end = preimage.at + preimage.count - 1
	switch (ref.kind) {
		case 'cell':
			return preimage.axis === 'row'
				? ref.ref.row >= preimage.at && ref.ref.row <= end
				: ref.ref.col >= preimage.at && ref.ref.col <= end
		case 'range':
			return preimage.axis === 'row'
				? ref.start.row <= end && ref.end.row >= preimage.at
				: ref.start.col <= end && ref.end.col >= preimage.at
		case 'wholeRowRange':
			return preimage.axis === 'row' && ref.startRow <= end && ref.endRow >= preimage.at
		case 'wholeColumnRange':
			return preimage.axis === 'col' && ref.startCol <= end && ref.endCol >= preimage.at
		default:
			return false
	}
}

function sheetSpanIncludes(
	workbook: Workbook,
	startSheet: string,
	endSheet: string,
	targetSheet: string,
): boolean {
	const start = workbook.sheets.findIndex((sheet) => sheet.name === startSheet)
	const end = workbook.sheets.findIndex((sheet) => sheet.name === endSheet)
	const target = workbook.sheets.findIndex((sheet) => sheet.name === targetSheet)
	if (start < 0 || end < 0 || target < 0) return false
	return target >= Math.min(start, end) && target <= Math.max(start, end)
}

function definedNameJournalKey(workbook: Workbook, entry: DefinedName): string {
	if (entry.scope.kind === 'workbook') return `name:${entry.name}`
	const sheetName = sheetNameForId(workbook, entry.scope.sheetId) ?? 'Sheet'
	return `name:${sheetName}!${entry.name}`
}

function structuralAffectedRange(axis: 'row' | 'col', at: number, count: number): RangeRef {
	const maxSpreadsheetIndex = Number.MAX_SAFE_INTEGER
	return axis === 'row'
		? {
				start: { row: at, col: 0 },
				end: { row: at + count - 1, col: maxSpreadsheetIndex },
			}
		: {
				start: { row: 0, col: at },
				end: { row: maxSpreadsheetIndex, col: at + count - 1 },
			}
}

function hasDeletedAxisDimensions(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	count: number,
): boolean {
	const end = at + count
	if (axis === 'row') {
		for (const row of sheet.rowHeights.keys()) if (row >= at && row < end) return true
		for (const row of sheet.rowDefs.keys()) if (row >= at && row < end) return true
		return false
	}
	for (const col of sheet.colWidths.keys()) if (col >= at && col < end) return true
	return sheet.colDefs.some((def) => def.max >= at && def.min < end)
}

function hasMapRefInRange<T>(map: ReadonlyMap<string, T>, range: RangeRef): boolean {
	for (const ref of map.keys()) {
		if (refInRange(ref, range)) return true
	}
	return false
}

function refInRange(ref: string, range: RangeRef): boolean {
	try {
		const parsed = parseA1(ref)
		return (
			parsed.row >= range.start.row &&
			parsed.row <= range.end.row &&
			parsed.col >= range.start.col &&
			parsed.col <= range.end.col
		)
	} catch {
		return false
	}
}

function sqrefOverlaps(sqref: string, range: RangeRef): boolean {
	return sqref
		.split(/\s+/)
		.filter(Boolean)
		.some((part) => {
			try {
				return rangesOverlap(parseRange(part), range)
			} catch {
				return false
			}
		})
}

function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
	return (
		a.start.row <= b.end.row &&
		a.end.row >= b.start.row &&
		a.start.col <= b.end.col &&
		a.end.col >= b.start.col
	)
}

function cellEditPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalCellPreimage[] {
	return cellPreimages(workbook, sheetName, formulaBindingEditRefs(workbook, sheetName, refs))
}

function formulaBindingJournalAddendum(cells: readonly MutationJournalCellPreimage[]): {
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (cells.length === 0) return { preimages: [], issues: [] }
	const { issues } = inverseCellOps(cells)
	return { preimages: [{ kind: 'cells', cells }], issues }
}

function workbookFormulaBindingCellPreimages(workbook: Workbook): MutationJournalCellPreimage[] {
	const cells: MutationJournalCellPreimage[] = []
	for (const sheet of workbook.sheets) {
		if (sheet.cells.formulaInfoCellCount() === 0) continue
		const refs: string[] = []
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (isMaterializedFormulaBindingInfo(cell.formulaInfo)) refs.push(toA1({ row, col }))
		}
		cells.push(...cellPreimages(workbook, sheet.name, refs))
	}
	return cells
}

function tableColumnFormulaBindingCellPreimages(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableColumn' }>,
): MutationJournalCellPreimage[] {
	if (op.newName !== undefined) return workbookFormulaBindingCellPreimages(workbook)
	const located = findTable(workbook, op.table)
	if (!located) return []
	const columnIndex = tableColumnIndex(located.table, op.column)
	if (columnIndex < 0) return []
	const col = located.table.ref.start.col + columnIndex
	const refs = new Set<string>()
	if (op.formula !== undefined) {
		const bodyStart = located.table.ref.start.row + (located.table.hasHeaders ? 1 : 0)
		const bodyEnd = located.table.ref.end.row - (located.table.hasTotals ? 1 : 0)
		for (let row = bodyStart; row <= bodyEnd; row++) refs.add(toA1({ row, col }))
	}
	if (
		located.table.hasTotals &&
		(op.totalsRowFunction !== undefined ||
			op.totalsRowFormula !== undefined ||
			op.totalsRowLabel !== undefined)
	) {
		refs.add(toA1({ row: located.table.ref.end.row, col }))
	}
	if (refs.size === 0) return []
	return formulaBindingOnlyPreimages(workbook, located.sheet.name, [...refs])
}

function tableColumnIndex(table: Table, columnSelector: string | number): number {
	return typeof columnSelector === 'number'
		? columnSelector
		: table.columns.findIndex(
				(column) => column.name.toLowerCase() === columnSelector.toLowerCase(),
			)
}

function formulaBindingOnlyPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalCellPreimage[] {
	return cellEditPreimages(workbook, sheetName, refs).filter((cell) =>
		isMaterializedFormulaBindingInfo(cell.formulaInfo),
	)
}

function isMaterializedFormulaBindingInfo(
	formulaInfo: Cell['formulaInfo'],
): formulaInfo is Exclude<NonNullable<Cell['formulaInfo']>, { kind: 'array' }> {
	return formulaInfo !== undefined && formulaInfo.kind !== 'array'
}

function formulaBindingEditRefs(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): string[] {
	const sheet = workbook.getSheet(sheetName)
	const parsedRefs = refs.map((ref) => parseA1(ref))
	const expanded = new Map<string, { readonly row: number; readonly col: number }>()
	const push = (row: number, col: number) => expanded.set(toA1({ row, col }), { row, col })
	for (const ref of parsedRefs) push(ref.row, ref.col)
	if (!sheet || sheet.cells.formulaInfoCellCount() === 0) return [...expanded.keys()]

	for (const ref of parsedRefs) {
		const binding = sheet.cells.get(ref.row, ref.col)?.formulaInfo
		if (!binding) continue
		if (binding.kind === 'shared') {
			for (const [row, col, cell] of sheet.cells.iterate()) {
				if (sameSharedFormulaBinding(binding, cell.formulaInfo)) push(row, col)
			}
			continue
		}
		if (isSpillFormulaBinding(binding)) {
			for (const [row, col, cell] of sheet.cells.iterate()) {
				if (sameSpillFormulaBinding(binding, cell.formulaInfo)) push(row, col)
			}
		}
	}

	for (const [row, col, cell] of sheet.cells.iterate()) {
		const binding = cell.formulaInfo
		if (binding?.kind !== 'dataTable') continue
		const tableRange = dataTableFormulaRange(binding, row, col)
		if (parsedRefs.some((ref) => rangeContainsCell(tableRange, ref))) push(row, col)
	}
	return [...expanded.keys()]
}

function sameSharedFormulaBinding(
	binding: Extract<NonNullable<Cell['formulaInfo']>, { kind: 'shared' }>,
	candidate: Cell['formulaInfo'],
): boolean {
	if (candidate?.kind !== 'shared') return false
	if (binding.sharedIndex !== undefined) return candidate.sharedIndex === binding.sharedIndex
	return candidate.masterRef === binding.masterRef
}

function isSpillFormulaBinding(
	binding: NonNullable<Cell['formulaInfo']>,
): binding is Extract<
	NonNullable<Cell['formulaInfo']>,
	{ kind: 'dynamicArray' | 'spill' | 'blockedSpill' }
> {
	return (
		binding.kind === 'dynamicArray' || binding.kind === 'spill' || binding.kind === 'blockedSpill'
	)
}

function sameSpillFormulaBinding(
	binding: Extract<
		NonNullable<Cell['formulaInfo']>,
		{ kind: 'dynamicArray' | 'spill' | 'blockedSpill' }
	>,
	candidate: Cell['formulaInfo'],
): boolean {
	if (!candidate) return false
	if (binding.kind === 'dynamicArray') {
		return candidate.kind === 'dynamicArray' && candidate.metadataIndex === binding.metadataIndex
	}
	if (candidate.kind !== 'spill' && candidate.kind !== 'blockedSpill') return false
	return candidate.anchorRef === binding.anchorRef
}

function dataTableFormulaRange(
	binding: Extract<NonNullable<Cell['formulaInfo']>, { kind: 'dataTable' }>,
	row: number,
	col: number,
): RangeRef {
	if (binding.ref) {
		try {
			return parseRange(binding.ref)
		} catch {
			// Fall back to the formula cell when imported metadata carries a malformed range.
		}
	}
	return { start: { row, col }, end: { row, col } }
}

function rangeContainsCell(
	range: RangeRef,
	ref: { readonly row: number; readonly col: number },
): boolean {
	return (
		ref.row >= range.start.row &&
		ref.row <= range.end.row &&
		ref.col >= range.start.col &&
		ref.col <= range.end.col
	)
}

function cellPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalCellPreimage[] {
	const sheet = workbook.getSheet(sheetName)
	return refs.map((ref) => {
		const parsed = parseA1(ref)
		const existing = sheet?.cells.get(parsed.row, parsed.col)
		const styleId = existing?.styleId ?? DEFAULT_STYLE_ID
		const style = cloneCellStyle(workbook.styles.get(styleId) ?? {})
		return {
			sheet: sheetName,
			ref: toA1(parsed),
			existed: existing !== undefined,
			value: cloneCellValue(existing?.value ?? EMPTY),
			formula: existing?.formula ?? null,
			...(existing?.formulaInfo ? { formulaInfo: cloneFormulaInfo(existing.formulaInfo) } : {}),
			styleId,
			style,
		}
	})
}

function inverseCellOps(cells: readonly MutationJournalCellPreimage[]): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const inverseOps: Operation[] = []
	const issues: MutationJournalIssue[] = []
	const scalarUpdatesBySheet = new Map<string, Array<{ ref: string; value: InputValue }>>()
	for (const cell of cells) {
		if (cell.formulaInfo) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Formula binding metadata for ${cell.sheet}!${cell.ref} cannot be restored with public operations`,
				refs: [`${cell.sheet}!${cell.ref}`],
			})
		}
		if (!cell.existed) {
			inverseOps.push({ op: 'clearRange', sheet: cell.sheet, range: cell.ref, what: 'all' })
			continue
		}
		if (cell.formula) {
			inverseOps.push({
				op: 'setFormula',
				sheet: cell.sheet,
				ref: cell.ref,
				formula: cell.formula,
			})
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Formula cache for ${cell.sheet}!${cell.ref} cannot be restored with public operations`,
				refs: [`${cell.sheet}!${cell.ref}`],
			})
			continue
		}
		if (cell.value.kind === 'richText') {
			const runs = richTextRunsToOperationRuns(cell.value.runs)
			if (runs.supported) {
				inverseOps.push({ op: 'setRichText', sheet: cell.sheet, ref: cell.ref, runs: runs.runs })
				continue
			}
			issues.push({
				code: 'UNSUPPORTED_VALUE',
				message: `Cannot restore richText at ${cell.sheet}!${cell.ref} with setRichText`,
				refs: [`${cell.sheet}!${cell.ref}`],
			})
			continue
		}
		const input = cellValueToInput(cell.value)
		if (input.supported) {
			const updates = scalarUpdatesBySheet.get(cell.sheet) ?? []
			updates.push({ ref: cell.ref, value: input.value })
			scalarUpdatesBySheet.set(cell.sheet, updates)
			continue
		}
		issues.push({
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot restore ${cell.value.kind} at ${cell.sheet}!${cell.ref} with setCells`,
			refs: [`${cell.sheet}!${cell.ref}`],
		})
	}
	for (const [sheet, updates] of scalarUpdatesBySheet) {
		inverseOps.push({ op: 'setCells', sheet, updates })
	}
	return { inverseOps, issues }
}

function richTextRunsToOperationRuns(runs: Extract<CellValue, { kind: 'richText' }>['runs']):
	| {
			readonly supported: true
			readonly runs: Extract<Operation, { op: 'setRichText' }>['runs']
	  }
	| { readonly supported: false } {
	const operationRuns: Array<Extract<Operation, { op: 'setRichText' }>['runs'][number]> = []
	for (const run of runs) {
		if (
			run.strikethrough !== undefined ||
			run.fontName !== undefined ||
			run.fontSize !== undefined ||
			(run.color !== undefined && typeof run.color !== 'string')
		) {
			return { supported: false }
		}
		operationRuns.push({
			text: run.text,
			...(run.bold !== undefined ? { bold: run.bold } : {}),
			...(run.italic !== undefined ? { italic: run.italic } : {}),
			...(run.underline !== undefined ? { underline: run.underline } : {}),
			...(run.color !== undefined ? { color: run.color } : {}),
		})
	}
	return { supported: true, runs: operationRuns }
}

function styleInverseOps(cells: readonly MutationJournalCellPreimage[]): Operation[] {
	const inverseOps: Operation[] = []
	for (const cell of cells) {
		if (!cell.existed) {
			inverseOps.push({ op: 'clearRange', sheet: cell.sheet, range: cell.ref, what: 'all' })
			continue
		}
		inverseOps.push({ op: 'clearRange', sheet: cell.sheet, range: cell.ref, what: 'styles' })
		if (Object.keys(cell.style).length > 0) {
			inverseOps.push({
				op: 'setStyle',
				sheet: cell.sheet,
				range: cell.ref,
				style: cell.style,
			})
		}
	}
	return inverseOps
}

function cellValueToInput(
	value: CellValue,
): { readonly supported: true; readonly value: InputValue } | { readonly supported: false } {
	switch (value.kind) {
		case 'empty':
			return { supported: true, value: null }
		case 'number':
		case 'string':
		case 'boolean':
			return { supported: true, value: value.value }
		default:
			return { supported: false }
	}
}

function refsInRange(rangeText: string): string[] {
	const range = parseRange(rangeText)
	const refs: string[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			refs.push(toA1({ row, col }))
		}
	}
	return refs
}

function sameRange(a: RangeRef, b: RangeRef): boolean {
	return (
		a.start.row === b.start.row &&
		a.start.col === b.start.col &&
		a.end.row === b.end.row &&
		a.end.col === b.end.col
	)
}

function sheetScopeForName(
	workbook: Workbook,
	sheetName: string,
): Extract<DefinedNameScope, { kind: 'sheet' }> | null {
	const sheet = workbook.getSheet(sheetName)
	return sheet ? { kind: 'sheet', sheetId: sheet.id } : null
}

function sheetNameForId(workbook: Workbook, sheetId: string): string | undefined {
	return workbook.sheets.find((sheet) => sheet.id === sheetId)?.name
}

function findTable(
	workbook: Workbook,
	tableName: string,
): { readonly sheet: Sheet; readonly table: Table } | null {
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.name.toLowerCase() === tableName.toLowerCase()) return { sheet, table }
		}
	}
	return null
}

function findTableColumn(
	workbook: Workbook,
	tableName: string,
	columnSelector: string | number,
): { readonly columnIndex: number; readonly column: TableColumn } | null {
	const located = findTable(workbook, tableName)
	if (!located) return null
	const columnIndex =
		typeof columnSelector === 'number'
			? columnSelector
			: located.table.columns.findIndex(
					(column) => column.name.toLowerCase() === columnSelector.toLowerCase(),
				)
	const column = located.table.columns[columnIndex]
	return column ? { columnIndex, column } : null
}

function commentPreimage(
	workbook: Workbook,
	sheetName: string,
	refText: string,
): MutationJournalCommentPreimage {
	const ref = refText.toUpperCase()
	const comment = findComment(workbook.getSheet(sheetName), ref)
	return {
		sheet: sheetName,
		ref,
		comment: comment ? { ...comment } : null,
	}
}

function findComment(sheet: Sheet | undefined, ref: string): SheetComment | null {
	if (!sheet) return null
	for (const [commentRef, comment] of sheet.comments) {
		if (commentRef.toUpperCase() === ref) return comment
	}
	return null
}

function hyperlinkPreimage(
	workbook: Workbook,
	sheetName: string,
	refText: string,
): MutationJournalHyperlinkPreimage {
	const ref = refText.toUpperCase()
	const hyperlink = findHyperlink(workbook.getSheet(sheetName), ref)
	return {
		sheet: sheetName,
		ref,
		hyperlink: hyperlink ? { ...hyperlink } : null,
	}
}

function findHyperlink(sheet: Sheet | undefined, ref: string): SheetHyperlink | null {
	if (!sheet) return null
	for (const [linkRef, hyperlink] of sheet.hyperlinks) {
		if (linkRef.toUpperCase() === ref) return hyperlink
	}
	return null
}

function threadedCommentPreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setThreadedComment' }>,
): MutationJournalThreadedCommentPreimage {
	const sheet = workbook.getSheet(op.sheet)
	if (
		!sheet ||
		(op.commentIndex !== undefined && (op.commentIndex < 0 || !Number.isInteger(op.commentIndex)))
	) {
		return { sheet: op.sheet, commentIndex: null, threadedComment: null }
	}
	const matches = sheet.threadedComments
		.map((comment, index) => ({ comment, index }))
		.filter(({ comment, index }) => {
			if (op.commentIndex !== undefined && index !== op.commentIndex) return false
			if (op.partPath !== undefined && comment.partPath !== op.partPath) return false
			if (op.threadedCommentId !== undefined && comment.id !== op.threadedCommentId) return false
			if (op.ref !== undefined && comment.ref.toUpperCase() !== op.ref.toUpperCase()) return false
			return true
		})
	const match = matches.length === 1 ? matches[0] : undefined
	return {
		sheet: op.sheet,
		commentIndex: match?.index ?? null,
		threadedComment: match ? cloneThreadedComment(match.comment) : null,
	}
}

function threadedCommentStableSelector(
	preimage: MutationJournalThreadedCommentPreimage,
): Omit<Extract<Operation, { op: 'setThreadedComment' }>, 'op' | 'sheet' | 'text'> {
	const comment = preimage.threadedComment
	if (!comment) return {}
	if (comment.partPath !== undefined && comment.id !== undefined) {
		return { partPath: comment.partPath, threadedCommentId: comment.id }
	}
	if (comment.id !== undefined) return { threadedCommentId: comment.id }
	if (comment.partPath !== undefined) return { partPath: comment.partPath }
	if (preimage.commentIndex !== null) return { commentIndex: preimage.commentIndex }
	return { ref: comment.ref }
}

function drawingTextPreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDrawingText' }>,
): MutationJournalDrawingTextPreimage {
	const sheet = workbook.getSheet(op.sheet)
	if (
		!sheet ||
		(op.drawingObjectIndex !== undefined &&
			(op.drawingObjectIndex < 0 || !Number.isInteger(op.drawingObjectIndex)))
	) {
		return { sheet: op.sheet, drawingObjectIndex: null, drawingObject: null }
	}
	const matches = sheet.drawingObjectRefs
		.map((object, index) => ({ object, index }))
		.filter(({ object, index }) => {
			if (op.drawingObjectIndex !== undefined && index !== op.drawingObjectIndex) return false
			if (op.drawingPartPath !== undefined && object.drawingPartPath !== op.drawingPartPath) {
				return false
			}
			if (op.id !== undefined && object.id !== op.id) return false
			if (op.name !== undefined && object.name !== op.name) return false
			return true
		})
	const match = matches.length === 1 ? matches[0] : undefined
	return {
		sheet: op.sheet,
		drawingObjectIndex: match?.index ?? null,
		drawingObject: match ? cloneDrawingObjectRef(match.object) : null,
	}
}

function drawingObjectStableSelector(
	preimage: MutationJournalDrawingTextPreimage,
): Omit<Extract<Operation, { op: 'setDrawingText' }>, 'op' | 'sheet' | 'text'> {
	const object = preimage.drawingObject
	if (!object) return {}
	if (object.drawingPartPath !== undefined && object.id !== undefined) {
		return { drawingPartPath: object.drawingPartPath, id: object.id }
	}
	if (object.id !== undefined) return { id: object.id }
	if (object.drawingPartPath !== undefined && object.name !== undefined) {
		return { drawingPartPath: object.drawingPartPath, name: object.name }
	}
	if (object.name !== undefined) return { name: object.name }
	if (preimage.drawingObjectIndex !== null) {
		return { drawingObjectIndex: preimage.drawingObjectIndex }
	}
	return { drawingPartPath: object.drawingPartPath }
}

function chartSeriesPreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setChartSeriesSource' }>,
): MutationJournalChartSeriesPreimage {
	if (
		op.seriesIndex < 0 ||
		!Number.isInteger(op.seriesIndex) ||
		(op.chartIndex !== undefined && (op.chartIndex < 0 || !Number.isInteger(op.chartIndex)))
	) {
		return { chartIndex: null, seriesIndex: op.seriesIndex, chart: null, series: null }
	}
	let candidates = workbook.chartParts.map((chart, index) => ({ chart, index }))
	if (op.partPath !== undefined) {
		candidates = candidates.filter(({ chart }) => chart.partPath === op.partPath)
	}
	if (op.sheet !== undefined) {
		candidates = candidates.filter(({ chart }) => chart.sheetName === op.sheet)
	}
	if (op.chartIndex !== undefined) {
		const chart = candidates[op.chartIndex]
		candidates = chart ? [chart] : []
	}
	const match = candidates.length === 1 ? candidates[0] : undefined
	const series = match?.chart.series[op.seriesIndex]
	return {
		chartIndex: match?.index ?? null,
		seriesIndex: op.seriesIndex,
		chart: match ? cloneChartPart(match.chart) : null,
		series: series ? cloneChartSeries(series) : null,
	}
}

function chartSeriesInverseOperation(
	op: Extract<Operation, { op: 'setChartSeriesSource' }>,
	preimage: MutationJournalChartSeriesPreimage,
): { readonly exact: boolean; readonly op?: Operation } {
	if (!preimage.chart || !preimage.series) return { exact: false }
	const refs: { nameRef?: string; categoryRef?: string; valueRef?: string } = {}
	let exact = true
	for (const field of ['nameRef', 'categoryRef', 'valueRef'] as const) {
		if (op[field] === undefined) continue
		const value = preimage.series[field]
		if (value === undefined) {
			exact = false
			continue
		}
		refs[field] = value
	}
	const inverse: Extract<Operation, { op: 'setChartSeriesSource' }> = {
		op: 'setChartSeriesSource',
		seriesIndex: preimage.seriesIndex,
		...chartSeriesStableSelector(preimage),
		...refs,
	}
	return hasChartSeriesRefUpdate(inverse) ? { exact, op: inverse } : { exact: false }
}

function chartSeriesStableSelector(
	preimage: MutationJournalChartSeriesPreimage,
): Omit<
	Extract<Operation, { op: 'setChartSeriesSource' }>,
	'op' | 'seriesIndex' | 'nameRef' | 'categoryRef' | 'valueRef'
> {
	const chart = preimage.chart
	if (!chart) return {}
	if (chart.partPath !== undefined) return { partPath: chart.partPath }
	if (chart.sheetName !== undefined && preimage.chartIndex !== null) {
		return { sheet: chart.sheetName, chartIndex: preimage.chartIndex }
	}
	if (chart.sheetName !== undefined) return { sheet: chart.sheetName }
	if (preimage.chartIndex !== null) return { chartIndex: preimage.chartIndex }
	return {}
}

function hasChartSeriesRefUpdate(op: Extract<Operation, { op: 'setChartSeriesSource' }>): boolean {
	return op.nameRef !== undefined || op.categoryRef !== undefined || op.valueRef !== undefined
}

function pivotCachePreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPivotCache' }>,
): MutationJournalPivotCachePreimage {
	let expectedCacheId = op.cacheId
	if (op.pivotTable !== undefined) {
		const pivots = workbook.pivotTables.filter((pivot) => pivot.name === op.pivotTable)
		if (pivots.length !== 1 || pivots[0]?.cacheId === undefined) {
			return { cacheIndex: null, cache: null }
		}
		if (expectedCacheId !== undefined && expectedCacheId !== pivots[0].cacheId) {
			return { cacheIndex: null, cache: null }
		}
		expectedCacheId = pivots[0].cacheId
	}
	const matches = workbook.pivotCaches
		.map((cache, index) => ({ cache, index }))
		.filter(({ cache }) => {
			if (op.partPath !== undefined && cache.partPath !== op.partPath) return false
			if (expectedCacheId !== undefined && cache.cacheId !== expectedCacheId) return false
			return true
		})
	const match = matches.length === 1 ? matches[0] : undefined
	return {
		cacheIndex: match?.index ?? null,
		cache: match ? clonePivotCacheInfo(match.cache) : null,
	}
}

function pivotCacheInverseOperation(
	op: Extract<Operation, { op: 'setPivotCache' }>,
	preimage: MutationJournalPivotCachePreimage,
): { readonly exact: boolean; readonly op?: Operation } {
	if (!preimage.cache) return { exact: false }
	const fields: {
		sourceSheet?: string
		sourceRef?: string
		refreshOnLoad?: boolean
		enableRefresh?: boolean
		invalid?: boolean
		saveData?: boolean
	} = {}
	let exact = true
	for (const field of [
		'sourceSheet',
		'sourceRef',
		'refreshOnLoad',
		'enableRefresh',
		'invalid',
		'saveData',
	] as const) {
		if (op[field] === undefined) continue
		switch (field) {
			case 'sourceSheet': {
				const value = preimage.cache.sourceSheet
				if (value === undefined) {
					exact = false
					continue
				}
				fields.sourceSheet = value
				break
			}
			case 'sourceRef': {
				const value = preimage.cache.sourceRef
				if (value === undefined) {
					exact = false
					continue
				}
				fields.sourceRef = value
				break
			}
			case 'refreshOnLoad': {
				const value = preimage.cache.refreshOnLoad
				if (value === undefined) {
					exact = false
					continue
				}
				fields.refreshOnLoad = value
				break
			}
			case 'enableRefresh': {
				const value = preimage.cache.enableRefresh
				if (value === undefined) {
					exact = false
					continue
				}
				fields.enableRefresh = value
				break
			}
			case 'invalid': {
				const value = preimage.cache.invalid
				if (value === undefined) {
					exact = false
					continue
				}
				fields.invalid = value
				break
			}
			case 'saveData': {
				const value = preimage.cache.saveData
				if (value === undefined) {
					exact = false
					continue
				}
				fields.saveData = value
				break
			}
		}
	}
	const inverse: Extract<Operation, { op: 'setPivotCache' }> = {
		op: 'setPivotCache',
		partPath: preimage.cache.partPath,
		...fields,
	}
	return hasPivotCacheUpdate(inverse) ? { exact, op: inverse } : { exact: false }
}

function hasPivotCacheUpdate(op: Extract<Operation, { op: 'setPivotCache' }>): boolean {
	return (
		op.sourceSheet !== undefined ||
		op.sourceRef !== undefined ||
		op.refreshOnLoad !== undefined ||
		op.enableRefresh !== undefined ||
		op.invalid !== undefined ||
		op.saveData !== undefined
	)
}

function setHyperlinkInverse(
	sheet: string,
	ref: string,
	hyperlink: SheetHyperlink,
): Extract<Operation, { op: 'setHyperlink' }> {
	return {
		op: 'setHyperlink',
		sheet,
		ref,
		...(hyperlink.target !== undefined ? { url: hyperlink.target } : {}),
		...(hyperlink.location !== undefined ? { location: hyperlink.location } : {}),
		...(hyperlink.display !== undefined ? { display: hyperlink.display } : {}),
		...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
	}
}

function cloneFormulaInfo(formulaInfo: NonNullable<Cell['formulaInfo']>): Cell['formulaInfo'] {
	if (formulaInfo.kind === 'blockedSpill') {
		return { ...formulaInfo, blockingRefs: [...formulaInfo.blockingRefs] }
	}
	return { ...formulaInfo }
}

function cloneCellValue(value: CellValue): CellValue {
	switch (value.kind) {
		case 'richText':
			return { kind: 'richText', runs: value.runs.map((run) => ({ ...run })) }
		case 'array':
			return {
				kind: 'array',
				rows: value.rows.map((row) => row.map(cloneScalarCellValue)),
			}
		default:
			return { ...value }
	}
}

function cloneScalarCellValue(value: ScalarCellValue): ScalarCellValue {
	return cloneCellValue(value) as ScalarCellValue
}

function cloneDefinedName(definedName: DefinedName): DefinedName {
	return {
		...definedName,
		scope: { ...definedName.scope },
		...(definedName.extraAttributes
			? { extraAttributes: definedName.extraAttributes.map((attribute) => ({ ...attribute })) }
			: {}),
	}
}

function clonePageSetup(setup: SheetPageSetup): SheetPageSetup {
	return { ...setup }
}

function cloneTableColumn(column: TableColumn): TableColumn {
	return {
		...column,
		...(column.xmlColumnPr ? { xmlColumnPr: { ...column.xmlColumnPr } } : {}),
	}
}

function cloneTable(table: Table): Table {
	return {
		...table,
		ref: { start: { ...table.ref.start }, end: { ...table.ref.end } },
		columns: table.columns.map(cloneTableColumn),
		...(table.autoFilter ? { autoFilter: cloneAutoFilter(table.autoFilter) } : {}),
		...(table.sortState
			? {
					sortState: {
						...table.sortState,
						...(table.sortState.preservedAttributes
							? { preservedAttributes: { ...table.sortState.preservedAttributes } }
							: {}),
						conditions: table.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
		...(table.tableStyleInfo ? { tableStyleInfo: { ...table.tableStyleInfo } } : {}),
		...(table.queryTable ? { queryTable: { ...table.queryTable } } : {}),
	}
}

function cloneThreadedComment(comment: SheetThreadedComment): SheetThreadedComment {
	return { ...comment }
}

function cloneDrawingObjectRef(object: SheetDrawingObjectRef): SheetDrawingObjectRef {
	return {
		...object,
		...(object.anchor ? { anchor: clonePlain(object.anchor) } : {}),
		...(object.relIds ? { relIds: [...object.relIds] } : {}),
		...(object.relationshipRefs
			? { relationshipRefs: object.relationshipRefs.map((ref) => ({ ...ref })) }
			: {}),
	}
}

function cloneChartPart(chart: ChartPartInfo): ChartPartInfo {
	return {
		...chart,
		series: chart.series.map(cloneChartSeries),
	}
}

function cloneChartSeries(series: ChartSeriesInfo): ChartSeriesInfo {
	return { ...series }
}

function cloneConditionalFormat(format: SheetConditionalFormat): SheetConditionalFormat {
	return {
		...format,
		rules: format.rules.map((rule) => ({
			...rule,
			formulas: [...rule.formulas],
			...(rule.style ? { style: cloneCellStyle(rule.style) } : {}),
			...(rule.preservedRuleAttributes
				? { preservedRuleAttributes: { ...rule.preservedRuleAttributes } }
				: {}),
			...(rule.preservedRuleChildXml
				? { preservedRuleChildXml: [...rule.preservedRuleChildXml] }
				: {}),
			...(rule.colorScale ? { colorScale: clonePlain(rule.colorScale) } : {}),
			...(rule.dataBar ? { dataBar: clonePlain(rule.dataBar) } : {}),
			...(rule.iconSet ? { iconSet: clonePlain(rule.iconSet) } : {}),
		})),
	}
}

function cloneAutoFilter(autoFilter: AutoFilter): AutoFilter {
	return {
		...autoFilter,
		columns: autoFilter.columns.map((column) => ({
			...column,
			...(column.values ? { values: [...column.values] } : {}),
			...(column.dateGroupItems
				? { dateGroupItems: column.dateGroupItems.map((item) => ({ ...item })) }
				: {}),
			...(column.customFilters
				? { customFilters: column.customFilters.map((filter) => ({ ...filter })) }
				: {}),
		})),
		...(autoFilter.sortState
			? {
					sortState: {
						...autoFilter.sortState,
						...(autoFilter.sortState.preservedAttributes
							? { preservedAttributes: { ...autoFilter.sortState.preservedAttributes } }
							: {}),
						conditions: autoFilter.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
	}
}

function clonePlain<T>(value: T): T {
	return JSON.parse(JSON.stringify(value)) as T
}

function isDataValidationType(value: string): value is DataValidationRule['type'] {
	return DATA_VALIDATION_TYPES.has(value as DataValidationRule['type'])
}

function isDataValidationOperator(
	value: string | undefined,
): value is NonNullable<DataValidationRule['operator']> {
	return (
		value !== undefined &&
		DATA_VALIDATION_OPERATORS.has(value as NonNullable<DataValidationRule['operator']>)
	)
}

function isConditionalFormatType(value: string): value is ConditionalFormatRule['type'] {
	return CONDITIONAL_FORMAT_TYPES.has(value as ConditionalFormatRule['type'])
}

function isConditionalFormatOperator(
	value: string | undefined,
): value is NonNullable<ConditionalFormatRule['operator']> {
	return (
		value !== undefined &&
		CONDITIONAL_FORMAT_OPERATORS.has(value as NonNullable<ConditionalFormatRule['operator']>)
	)
}

const DATA_VALIDATION_TYPES = new Set<DataValidationRule['type']>([
	'list',
	'whole',
	'decimal',
	'date',
	'time',
	'textLength',
	'custom',
])

const DATA_VALIDATION_OPERATORS = new Set<NonNullable<DataValidationRule['operator']>>([
	'between',
	'notBetween',
	'equal',
	'notEqual',
	'greaterThan',
	'lessThan',
	'greaterThanOrEqual',
	'lessThanOrEqual',
])

const CONDITIONAL_FORMAT_TYPES = new Set<ConditionalFormatRule['type']>([
	'cellIs',
	'expression',
	'colorScale',
	'dataBar',
	'iconSet',
	'top10',
	'aboveAverage',
	'duplicateValues',
	'containsText',
])

const CONDITIONAL_FORMAT_OPERATORS = new Set<NonNullable<ConditionalFormatRule['operator']>>([
	'greaterThan',
	'lessThan',
	'equal',
	'between',
	'greaterThanOrEqual',
	'lessThanOrEqual',
	'notEqual',
	'notBetween',
])
